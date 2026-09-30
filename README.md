# o365-to-mailcow

Migrate Microsoft 365 **mail, calendars and contacts** into **mailcow** (Dovecot + SOGo),
reading everything through Microsoft Graph and writing through IMAP, CalDAV and CardDAV.

- One container, one config file, four commands: `plan`, `migrate`, `verify`, `cleanup`.
- Read-only against Microsoft 365: it never deletes, moves or flags anything at the source.
- Re-runnable: every write is keyed on a stable identifier, so a second run only copies
  what is new or changed. Interrupted runs resume where they stopped.
- Proves the result: `verify` compares per folder, per calendar and per address book, and
  exits non-zero while anything is failed, skipped or mismatched.
- No imapsync, no EWS, no subprocesses. Outbound hosts are exactly
  `login.microsoftonline.com`, `graph.microsoft.com` and your mailcow host.

## How it works

| Data | Source (Graph v1.0) | Destination |
|------|--------------------|-------------|
| Mail | `/mailFolders`, `/messages/{id}/$value` (MIME, byte-for-byte) | IMAP `APPEND` over TLS |
| Calendars | `/calendars`, `/calendars/{id}/events`, `/events/{id}/instances` | SOGo CalDAV `PUT` (`.ics`) |
| Contacts | `/contactFolders` (recursive), `/contacts`, contact photos | SOGo CardDAV `PUT` (`.vcf`) |

Destination credentials are temporary mailcow **app passwords** (one per mailbox and run,
named `o365-migration-<id>`, protocols `imap_access` and `dav_access` only), created
through the mailcow API at the start of a run and deleted at the end, also when the run
fails or is interrupted with Ctrl-C. A hard kill (`docker stop`, power loss) can leave one
behind; it is recorded in the state database and removed by the next `migrate` or by
`cleanup`.

## 1. Register the app in Microsoft Entra

Pick one of two modes. **App-only (`auth_mode = "app"`) is the default and recommended**:
no user signs in and nobody needs Full Access to the mailboxes.

### Option A: app-only (client credentials)

1. Entra admin center > **Identity > Applications > App registrations > New registration**.
   Name `o365-to-mailcow`, *Accounts in this organizational directory only*, no redirect URI.
2. Note the **Application (client) ID** and **Directory (tenant) ID** from *Overview*.
3. **API permissions > Add a permission > Microsoft Graph > Application permissions**, add:
   - `Mail.Read`
   - `Calendars.Read`
   - `Contacts.Read`
   - only for the optional [web UI](#web-ui-optional): `User.Read.All` (lists the mailboxes)

   Then click **Grant admin consent for <tenant>**. No write permission is needed or used.
4. **Certificates & secrets > Client secrets > New client secret**. Copy the *Value*
   into `.env` as `O365MIG_CLIENT_SECRET` (it is shown only once).

#### Optional but recommended: restrict the app to the migrated mailboxes

Application permissions apply to every mailbox in the tenant. Scope them with an Exchange
Online **Application Access Policy** (Exchange Online PowerShell):

```powershell
Connect-ExchangeOnline
# mail-enabled security group containing exactly the mailboxes to migrate
New-DistributionGroup -Name "o365-migration-scope" -Type Security `
  -Members alice@contoso.com,bob@contoso.com
New-ApplicationAccessPolicy -AppId <client-id> `
  -PolicyScopeGroupId o365-migration-scope@contoso.com `
  -AccessRight RestrictAccess `
  -Description "o365-to-mailcow may read only the migrated mailboxes"
# check: should say AccessCheckResult Granted / Denied as expected
Test-ApplicationAccessPolicy -Identity alice@contoso.com -AppId <client-id>
Test-ApplicationAccessPolicy -Identity ceo@contoso.com -AppId <client-id>
```

Policy changes can take up to an hour to apply. (Microsoft is moving this feature to
*RBAC for Applications*; either mechanism works with this tool.)

### Option B: delegated device-code sign-in (fallback)

Use this only if your tenant forbids application permissions.

1. Register the app as in option A, steps 1 and 2.
2. **Authentication > Advanced settings > Allow public client flows: Yes**.
3. **API permissions > Add a permission > Microsoft Graph > Delegated permissions**, add:
   - `Mail.Read.Shared`
   - `Calendars.Read.Shared`
   - `Contacts.Read.Shared`
   - `User.Read`
   - `offline_access` (MSAL requests this automatically; it keeps the run signed in)
   - only for the optional [web UI](#web-ui-optional): `User.Read.All`

   Grant admin consent.
4. The signing-in administrator needs **Full Access** to every migrated mailbox:
   ```powershell
   Add-MailboxPermission -Identity alice@contoso.com -User admin@contoso.com `
     -AccessRights FullAccess -AutoMapping $false
   ```
5. Set `auth_mode = "delegated"`. On the first run the tool prints a URL and a code; sign
   in on any device. The token cache is stored in `state/msal_cache.bin` (mode 0600) so
   later runs do not prompt again.

If sign-in fails with a policy error, a Conditional Access policy is probably blocking
the **device code flow** (Conditions > *Authentication flows*). Exclude the administrator
for the duration of the migration or switch to app-only.

## 2. Create the mailcow API key

1. mailcow UI > **System > Configuration > Access > API**.
2. Expand **Read-Write Access**, tick **Activate API**, and in **Allow API access from
   these IPs/networks** enter the address the container's traffic comes from as seen by
   mailcow (the Docker host's public IP, or its LAN IP if it is on the same network).
3. Save and copy the key into `.env` as `O365MIG_MAILCOW_API_KEY`.

The key could create or delete mailboxes, so the tool refuses every endpoint except
`get/mailbox`, `get/app-passwd`, `add/app-passwd` and `delete/app-passwd`. `migrate` never
creates mailboxes; either create them in the mailcow UI first or use the separate
`provision` command below. Consider deactivating the API again after the migration.

## 3. Configure and run

```sh
mkdir -p state config
cp config.example.toml config/config.toml
cp .env.example .env && chmod 600 .env                  # fill in both secrets
$EDITOR config/config.toml                               # tenant, client id, host, mailboxes
sudo chown -R 10001:10001 state config                   # the container runs as uid 10001
chmod 600 config/config.toml
docker compose build
```

Secrets given through `.env` are visible to anyone who can run `docker inspect` on the
host; on a shared host, prefer Docker secrets or a root-only `.env`.

Every option in `config.example.toml` is commented. Secrets belong in `.env`, not in the
config file; the tool warns when the config file is readable by other users.

Mailboxes come from `mailboxes = [...]` in the config and/or a CSV passed with
`--mailboxes` (rows `source[,destination]`; the destination may differ from the source).

### Provisioning mailboxes (optional)

`o365mig provision` creates every destination mailbox that does not exist yet, and only
those. It is a separate command, never part of `migrate`, and it is the only command that
may call mailcow's `get/domain` and `add/mailbox` endpoints. Rules:

- The mailbox's **domain must already exist in mailcow** (adding a domain is a DNS
  decision; the tool refuses and tells you).
- Display name and quota come from the mailbox list: CSV rows are
  `source,destination,name,quota_mib`, TOML entries accept `name` and `quota_mib`; the
  defaults are the local part of the address and `provision_quota_mib` (3072).
- Each mailbox gets a generated 32-character password with **"change password at first
  login"** set (TLS enforcement only with `provision_tls_enforce = true`, because it
  rejects mail from senders without TLS). The passwords are written to
  `state/provisioned-<timestamp>.csv` (mode 0600) *before* each mailbox is created and
  the row is updated afterwards (`pending` → `created`, or `failed`), so an interrupted
  run never loses a password it already set. The file is never logged. Distribute the
  passwords, then delete the file. Users can enable two-factor authentication themselves.
- `--dry-run` shows what would be created and creates nothing.

```sh
docker compose run --rm o365mig --mailboxes /config/mailboxes.csv provision --dry-run
docker compose run --rm o365mig --mailboxes /config/mailboxes.csv provision
```

### The four commands

```sh
docker compose run --rm o365mig plan      # what would move; creates nothing
docker compose run --rm o365mig migrate   # move it (safe to re-run)
docker compose run --rm o365mig verify    # prove it; exit 1 if anything is off
docker compose run --rm o365mig cleanup   # delete all o365-migration app passwords
```

| Command | Does | Writes |
|---------|------|--------|
| `plan` | Lists folders with message counts and sizes, calendars, contact folders, skipped items; checks that each destination mailbox exists | nothing (Graph reads, mailcow `get` calls) |
| `migrate` | Creates a temporary app password per mailbox, migrates, deletes the app password | IMAP APPEND, DAV MKCALENDAR/MKCOL/PUT, app password add/delete |
| `verify` | Per folder: Graph total, skipped, failed, expected, IMAP count. Per calendar/address book: Graph count vs DAV count. Prints every skipped and failed category | app password add/delete only |
| `cleanup` | Deletes every app password named `o365-migration-*` for the configured mailboxes plus any recorded in state | app password delete |

Options (before or after the command):

- `--config PATH` (default `$O365MIG_CONFIG`, in the image `/config/config.toml`)
- `--mailboxes CSV` extra mailbox list
- `--mailbox ADDRESS` only this mailbox (source or destination address)
- `--only mail|calendar|contacts` only this kind of data
- `--dry-run` list and count only: no APPEND, PUT, MKCALENDAR, MKCOL, no app passwords
- `--keep-app-passwords` do not delete the temporary app passwords (debugging)
- `verify --sample N` re-download N random migrated messages per mailbox and compare
  them with the destination copy. Exchange re-renders the MIME of some items (mail it
  stores natively, not as received bytes); those show up as "re-rendered", not as
  mismatches, when Message-ID, Date, From, Subject and size class agree
- `-v` debug logging

Exit codes: **0** everything succeeded, **1** something failed (or, for `verify`, anything
was skipped, failed or mismatched), **2** configuration or sign-in error.

Every run writes `state/reports/<UTC timestamp>.json` (per mailbox, per folder, per
calendar and address book: counts, failures, skips, durations) and a log in
`state/logs/`. Progress (done/total and items per minute per mailbox) is printed at
least every 30 seconds.

### Web UI (optional)

`o365mig web` serves a local page for the work around the commands: list the tenant's
mailboxes, pick the ones to migrate, set destination address, display name and quota per
mailbox, check which destinations exist in mailcow, save the list, start `plan`,
`provision`, `migrate`, `verify` and `cleanup`, watch their output and read the latest
report.

```sh
docker compose up -d web
docker compose logs web    # web UI: http://0.0.0.0:8080/#token=<token>
```

Open `http://127.0.0.1:8080/#token=<token>` on the Docker host. Without Docker,
`o365mig --config config.toml web` prints `http://127.0.0.1:8080/#token=<token>`
(options `--bind ADDRESS`, `--port N`).

- **Connections panel.** The tenant id, client id, sign-in mode, client secret, mailcow
  host and API key can be entered in the page instead of the files. They are saved to
  `state/settings.toml` (mode 0600) and take precedence over `config.toml` and `.env`;
  secrets are never shown again, only "set"/"not set". "Test connections" tries a
  Microsoft sign-in, a Graph user listing (`User.Read.All`) and a mailcow API call.
  Two rules protect the credentials: a saved mailcow host is only ever used with the API
  key saved *with* it (changing the host means entering the key again), and saved tenant
  or client ids only with the client secret saved with them; the page can never pair a
  new host with a key from `.env` or the config file. Every change is logged with the
  field names. `o365mig web --lock-settings` makes the panel read-only for hardened
  setups; `--allow-host NAME` accepts an extra `Host` header value (loopback and the bind
  address are always accepted; anything else is refused to defeat DNS rebinding).
- **Token.** Every API call needs the token from that URL. It is new at every start, or
  fixed with `O365MIG_WEB_TOKEN` in `.env` (at least 16 characters, for example from
  `python3 -c 'import secrets; print(secrets.token_urlsafe(32))'`). It travels in the URL
  fragment, which browsers never send to a server, but it is printed in the container log.
  Anyone holding it can start migrations: treat it like the API key.
- **Keep it on localhost.** The server speaks plain HTTP and the token is its only
  protection. The compose file publishes it on `127.0.0.1:8080` only (inside the container
  it listens on all interfaces, hence the "reachable from the network" warning in its
  log). From another machine use an SSH tunnel (`ssh -L 8080:127.0.0.1:8080 docker-host`)
  or your own reverse proxy with TLS and authentication in front of it; never publish the
  port on a public interface.
- **Extra Graph permission.** Listing the tenant's mailboxes needs **`User.Read.All`**:
  as an *application* permission for app-only sign-in, as a *delegated* permission for
  device-code sign-in, with admin consent either way. The migration itself does not use
  it; without it only the listing fails, with a message saying so.
- **Mailbox list.** The page saves `state/mailboxes.csv` (mode 0600, the same
  `source,destination,name,quota_mib` format as `--mailboxes`) and every command it starts
  runs with `--mailboxes` pointing at that file **and `--mailboxes-only`**, so a job
  started from the page acts on the saved selection alone, never on the config's own
  `mailboxes = [...]`. The confirmation names the number of saved mailboxes, and a job is
  refused if the saved selection changed after the page loaded it.
- **One command at a time.** Commands run inside the web container and take the same
  state lock as `docker compose run`, so a page-started run and a CLI run never overlap.
  Stopping the web container stops a running command; start it again to resume, exactly as
  after Ctrl-C. In delegated mode the device-code prompt appears in the command's output
  (for the mailbox listing: in the container log).

## What is migrated, and how

**Mail.** Every folder, recursively. Well-known folders map to the Dovecot names:
Inbox > `INBOX`, Sent Items > `Sent`, Drafts > `Drafts`, Deleted Items > `Trash`,
Junk Email > `Junk`, Archive > `Archive`. Other folders keep their names and hierarchy;
control characters are removed, a hierarchy delimiter inside a name becomes `_`, and a
name that collides with a mapped folder gets ` (2)`. Messages are appended as the
MIME Graph returns (line endings normalised to CRLF by IMAP), with `\Seen` (read), `\Flagged` (flagged), `\Draft` (draft), Outlook
categories as IMAP keywords (spaces and special characters become `_`), and the received
date as the internal date. Messages above `max_message_bytes` (default 150 MiB) are skipped
and reported.

Re-runs: messages are keyed on (mailbox, source folder, immutable Graph id). Before
appending a message into a folder that already had content, the tool searches the
destination for its `Message-ID` and compares the content, so mail that is already there
(for example from an earlier manual copy) is not duplicated, while a different message
that merely carries the same Message-ID is still copied. After the first full pass each
folder uses a Graph delta link, so later runs only look at new messages; if Microsoft has
expired the link, the folder is listed fully again. Messages deleted at the source are
counted and reported, **never deleted at the destination**. If a destination folder's
UIDVALIDITY changes, the tool re-checks every message of that folder by Message-ID.

Read and flag changes made at the source *after* a message was copied are not carried
over by later runs (the tool never modifies existing messages at the destination). Plan
the final run for a moment when users have stopped working in Microsoft 365.

**Calendars.** The default calendar goes to SOGo's `personal` calendar; every other
calendar the mailbox owns is created with MKCALENDAR. Calendars shared *to* the mailbox by
others are skipped and listed. Series masters carry their recurrence, exceptions and
cancelled occurrences; attendees and organizer are kept with `SCHEDULE-AGENT=CLIENT` so
SOGo sends no invitations (`calendar_attendees = "strip"` removes them entirely).
Events are only re-uploaded when Graph's `lastModifiedDateTime` changed.

**Contacts.** The default contact folder goes to SOGo's `personal` address book; every
other folder (recursively) becomes its own address book. If SOGo refuses to create an
address book, those contacts go to `personal` with the folder name as a category, and
the report says so. Photos are copied when `contacts_photos = true`.

## What is not migrated

- Teams chats, OneDrive, SharePoint, To Do tasks, OneNote, Outlook rules, signatures,
  category colours, retention and litigation holds.
- Online archive mailboxes (separate Graph resources; not in this version).
- Sharing permissions, delegates, calendar ACLs: recreate them in SOGo by hand.
- Personal distribution lists (contact groups): Graph does not expose them.
- Calendars shared to the user by someone else (they belong to the other mailbox).
- Well-known system folders: Conversation History, Outbox, Sync Issues (and its
  Conflicts / Local Failures / Server Failures), Recoverable Items. `verify` lists them.
- Deletions after the first run (by design; the tool never deletes at the destination).
- Exchange on-premises.

## Test first

Before migrating real users, run against one test mailbox and check these three things
by hand; they depend on your SOGo and tenant configuration:

1. **No invitations are sent.** Create a meeting in the test mailbox with an *external
   attendee whose mailbox you control*, migrate only that mailbox
   (`--mailbox test@contoso.com --only calendar migrate`), and confirm the external
   address received nothing.
2. **MKCALENDAR works.** Give the test mailbox a second calendar and confirm it appears
   in SOGo after `migrate` (and that `verify` shows matching counts for it).
3. **App-password DAV login works on a 2FA mailbox.** If you use two-factor
   authentication in mailcow, enable it on the test mailbox and confirm calendars and
   contacts still migrate (SOGo must accept the app password for DAV).

Then run `plan` for everyone, `migrate` a few mailboxes, `verify`, and only then the rest.

## First live run (pilot protocol)

The test suite checks this tool against its own model of Graph, mailcow and SOGo, not
against the real services. Treat the first run as the test that matters:

1. **Back up mailcow** (or snapshot the VM). Use a throwaway destination mailbox.
2. **Seed one pilot mailbox in Microsoft 365** with awkward content: a large attachment
   (over 25 MB), a `.msg` attached inside a message, a folder whose name has non-ASCII
   characters and a `/`, a weekly series with one moved and one cancelled occurrence, an
   event created in a non-UTC time zone, a private event, a meeting organised by someone
   else, a contact with a photo, and a few thousand filler messages so paging happens.
3. **`plan`** and compare folder names and counts with what Outlook shows.
4. **`migrate`**, kill it part-way (Ctrl-C), run it again: nothing duplicated, nothing
   missing. Run it a third time: it must make zero writes.
5. **`verify`** (with `--sample 20`), then look by eye: dates and flags in a mail client,
   the series and time zones in SOGo and on a phone over CalDAV, the contact on the phone.
6. **`cleanup`**, then confirm in the mailcow UI that no `o365-migration-*` app password
   remains, and check `docker compose logs netfilter-mailcow` for the container's IP.
7. Only then one real, low-stakes mailbox; only then batches.

**fail2ban:** mailcow's netfilter bans an IP after repeated failed logins. The tool never
retries a refused login, but a wrong API key or a mailbox with the protocols disabled can
still produce a few failures per run. If the container's address gets banned, unban it in
**System > Configuration > Options > Fail2ban parameters** or whitelist it there first.

## Troubleshooting

| Symptom | Cause and fix |
|---|---|
| "Test connections": `microsoft ✓` but `graph_users ✗ … HTTP 403 … Authorization_RequestDenied` | The app registration lacks the mailbox-listing permission. Entra admin center → your app → **API permissions** → **Add a permission** → **Microsoft Graph** → **Application permissions** → `User.Read.All` → Add → **Grant admin consent for \<tenant\>**. Wait a minute or two, test again. Only the web page's listing needs it; the migration does not. |
| "Test connections": `microsoft ✗ … AADSTS7000215` (invalid client secret) | The secret *Value* was not copied, or the secret expired. Create a new one under **Certificates & secrets** and enter it in the Connections panel together with the tenant and client ids. |
| "Test connections": `microsoft ✗ … AADSTS700016` (application not found) | Wrong tenant id or client id: both come from the app's **Overview** page. |
| "Test connections": `mailcow ✗ … HTTP 401` | API key wrong, API not activated, or this host's IP is not in mailcow's allowed API networks (System → Configuration → Access → API). Enter the key together with the host. |
| `Load from Microsoft 365` shows no shared mailboxes | Shared mailboxes are users without a licence; they are listed as kind `shared`. If they are missing entirely they have no `mail` attribute or are guests. |
| A job fails immediately with `missing required configuration` | Save the Connections panel first (or fill `config.toml`/`.env`); status shows `configured: false` until then. |
| Phones or mail programs cannot log in to a *provisioned* mailbox | The initial password must be changed at first login in the mailcow UI (`force_pw_update`); mail programs cannot do that. |
| The container's IP gets banned by mailcow | See **fail2ban** under "First live run"; unban or whitelist it under System → Configuration → Options → Fail2ban parameters. |

## State, security and operations

- `state/state.db` (mode 0600) is the idempotency ledger: identifiers, statuses, error
  summaries, app-password IDs. It never contains message content or credentials. Keep it
  between runs; deleting it makes the next run fall back to Message-ID checks.
- In delegated mode, `state/msal_cache.bin` holds a refresh token with Full Access to
  every migrated mailbox. **Delete it when the migration is finished.**
- App-password IDs are recorded *before* use; a crashed run's leftovers are deleted by
  the next `migrate` or by `cleanup`. Do not run `cleanup` while a `migrate` is running.
- Secrets are read from the environment, never logged, never written to state or reports.
- TLS verification is always on and cannot be switched off; redirects are never followed
  (a redirect from any endpoint is an error, so credentials cannot be sent elsewhere).
  Graph requests carry a timeout, retry 429/503/504 with `Retry-After`, and are capped at
  four in flight per mailbox (Graph's per-mailbox limit). Downloads are streamed and
  abandoned past `max_message_bytes`; the compose file caps memory at 3 GiB.
- The container runs as uid 10001, read-only root filesystem in `compose.yaml`.

## Development

```sh
python3.12 -m venv .venv && . .venv/bin/activate
pip install -e '.[dev]'
ruff check src tests && bandit -r src -ll && pip-audit -r requirements.txt && pytest
docker build -t o365-to-mailcow . && docker run --rm o365-to-mailcow --help
```

Dependencies are locked with hashes in `requirements.txt`
(`uv pip compile pyproject.toml -o requirements.txt --generate-hashes`); the build
backend is locked the same way in `requirements-build.txt`.
