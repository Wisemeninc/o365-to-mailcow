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

Destination credentials are temporary mailcow **app passwords** (one per mailbox, named
`o365-migration`, protocols `imap_access` and `dav_access` only), created through the
mailcow API at the start of a run and deleted at the end, also when the run fails.

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
`get/mailbox`, `get/app-passwd`, `add/app-passwd` and `delete/app-passwd`. It never creates
mailboxes: **create the destination mailboxes in mailcow first**. Consider deactivating
the API again after the migration.

## 3. Configure and run

```sh
mkdir -p state config && sudo chown 10001:10001 state   # container runs as uid 10001
cp config.example.toml config/config.toml && chmod 600 config/config.toml
cp .env.example .env && chmod 600 .env                  # fill in both secrets
$EDITOR config/config.toml                               # tenant, client id, host, mailboxes
docker compose build
```

Every option in `config.example.toml` is commented. Secrets belong in `.env`, not in the
config file; the tool warns when the config file is readable by other users.

Mailboxes come from `mailboxes = [...]` in the config and/or a CSV passed with
`--mailboxes` (rows `source[,destination]`; the destination may differ from the source).

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
| `cleanup` | Deletes every app password named `o365-migration` for the configured mailboxes plus any recorded in state | app password delete |

Options (before or after the command):

- `--config PATH` (default `$O365MIG_CONFIG`, in the image `/config/config.toml`)
- `--mailboxes CSV` extra mailbox list
- `--mailbox ADDRESS` only this mailbox (source or destination address)
- `--only mail|calendar|contacts` only this kind of data
- `--dry-run` list and count only: no APPEND, PUT, MKCALENDAR, MKCOL, no app passwords
- `--keep-app-passwords` do not delete the temporary app passwords (debugging)
- `verify --sample N` re-download N random migrated messages per mailbox and compare
  their SHA-256 with the destination copy
- `-v` debug logging

Exit codes: **0** everything succeeded, **1** something failed (or, for `verify`, anything
was skipped, failed or mismatched), **2** configuration or sign-in error.

Every run writes `state/reports/<UTC timestamp>.json` (per mailbox, per folder, per
calendar and address book: counts, failures, skips, durations) and a log in
`state/logs/`. Progress (done/total and items per minute per mailbox) is printed at
least every 30 seconds.

## What is migrated, and how

**Mail.** Every folder, recursively. Well-known folders map to the Dovecot names:
Inbox > `INBOX`, Sent Items > `Sent`, Drafts > `Drafts`, Deleted Items > `Trash`,
Junk Email > `Junk`, Archive > `Archive`. Other folders keep their names and hierarchy;
control characters are removed, a hierarchy delimiter inside a name becomes `_`, and a
name that collides with a mapped folder gets ` (2)`. Messages are appended as the exact
MIME Graph returns, with `\Seen` (read), `\Flagged` (flagged), `\Draft` (draft), Outlook
categories as IMAP keywords (spaces and special characters become `_`), and the received
date as the internal date. Messages above `max_message_bytes` (default 150 MiB) are skipped
and reported.

Re-runs: messages are keyed on (mailbox, source folder, immutable Graph id). Before
appending a message the tool searches the destination folder for its `Message-ID`, so
mail that is already there (for example from an earlier manual copy) is not duplicated.
After the first full pass each folder uses a Graph delta link, so later runs only look at
new messages. Messages deleted at the source are counted and reported, **never deleted at
the destination**. If a destination folder's UIDVALIDITY changes, the tool re-checks
every message of that folder by Message-ID.

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

## State, security and operations

- `state/state.db` (mode 0600) is the idempotency ledger: identifiers, statuses, error
  summaries, app-password IDs. It never contains message content or credentials. Keep it
  between runs; deleting it makes the next run fall back to Message-ID checks.
- App-password IDs are recorded *before* use; a crashed run's leftovers are deleted by
  the next `migrate` or by `cleanup`. Do not run `cleanup` while a `migrate` is running.
- Secrets are read from the environment, never logged, never written to state or reports.
- TLS verification is always on. Graph requests carry a timeout, retry 429/503/504 with
  `Retry-After`, and are capped at four in flight per mailbox (Graph's per-mailbox limit).
- The container runs as uid 10001, read-only root filesystem in `compose.yaml`.

## Development

```sh
python3.12 -m venv .venv && . .venv/bin/activate
pip install -e '.[dev]'
ruff check src tests && bandit -r src -ll && pip-audit -r requirements.txt && pytest
docker build -t o365-to-mailcow . && docker run --rm o365-to-mailcow --help
```

Dependencies are locked with hashes in `requirements.txt`
(`uv pip compile pyproject.toml -o requirements.txt --generate-hashes`).
