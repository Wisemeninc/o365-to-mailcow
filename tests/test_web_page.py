"""Static checks on the single-file web page (ISC-162, ISC-164)."""

from __future__ import annotations

import re
from collections import Counter
from html.parser import HTMLParser
from pathlib import Path

PAGE = (
    Path(__file__).resolve().parents[1]
    / "src"
    / "o365_to_mailcow"
    / "web_static"
    / "index.html"
)


class PageParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=False)
        self.start_tags: list[tuple[str, dict[str, str]]] = []
        self._in_script = False
        self._script_chunks: list[str] = []

    def _record(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        clean: dict[str, str] = {}
        for key, value in attrs:
            clean[key] = "" if value is None else value
        self.start_tags.append((tag, clean))
        if tag == "script":
            self._in_script = True

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        self._record(tag, attrs)

    def handle_startendtag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        self._record(tag, attrs)

    def handle_endtag(self, tag: str) -> None:
        if tag == "script":
            self._in_script = False

    def handle_data(self, data: str) -> None:
        if self._in_script:
            self._script_chunks.append(data)

    @property
    def script_text(self) -> str:
        return "".join(self._script_chunks)


PAGE_BYTES = PAGE.read_bytes()
PAGE_TEXT = PAGE_BYTES.decode("utf-8")
PARSER = PageParser()
PARSER.feed(PAGE_TEXT)
TAGS = PARSER.start_tags
SCRIPT_TEXT = PARSER.script_text


def test_one_inline_script_and_one_style_both_with_nonce() -> None:
    scripts = [attrs for tag, attrs in TAGS if tag == "script"]
    styles = [attrs for tag, attrs in TAGS if tag == "style"]
    assert len(scripts) == 1, f"Expected one script tag, found {len(scripts)}"
    assert len(styles) == 1, f"Expected one style tag, found {len(styles)}"

    script = scripts[0]
    style = styles[0]
    assert script.get("nonce") == "__CSP_NONCE__", "Script nonce must be __CSP_NONCE__"
    assert style.get("nonce") == "__CSP_NONCE__", "Style nonce must be __CSP_NONCE__"
    assert "src" not in script, "Inline script must not have src"

    nonce_count = PAGE_TEXT.count("__CSP_NONCE__")
    assert nonce_count == 2, f"Expected __CSP_NONCE__ twice, found {nonce_count}"


def test_no_inline_event_handlers_or_style_attributes() -> None:
    offenders: list[str] = []
    for tag, attrs in TAGS:
        for name in attrs:
            if name.startswith("on"):
                offenders.append(f"<{tag}> has inline handler {name}")
            if name == "style":
                offenders.append(f"<{tag}> has style attribute")
    assert not offenders, "Inline handlers/style attrs are forbidden: " + "; ".join(offenders)

    pattern = re.compile(r"\.on[a-z]+\s*=(?!=)")
    assert pattern.search(SCRIPT_TEXT) is None, "Script uses .on* property assignment"


def test_no_external_resources() -> None:
    src_tags = [tag for tag, attrs in TAGS if "src" in attrs]
    assert not src_tags, f"No src attrs allowed, found on tags: {src_tags}"

    hrefs = [attrs["href"] for _, attrs in TAGS if "href" in attrs]
    bad_hrefs = [href for href in hrefs if not (href == "data:," or href.startswith("#"))]
    assert not bad_hrefs, f"Only data:, or hash hrefs allowed, found: {bad_hrefs}"

    banned_tags = {"img", "iframe", "object", "embed", "svg", "frame"}
    present_banned = [tag for tag, _ in TAGS if tag in banned_tags]
    assert not present_banned, f"Forbidden tags present: {present_banned}"

    links = [attrs for tag, attrs in TAGS if tag == "link"]
    bad_links = [attrs for attrs in links if attrs.get("rel") != "icon"]
    assert not bad_links, "Every link tag must have rel=icon"

    forbidden_fragments = ["http://", "https://", "@import", "url("]
    for fragment in forbidden_fragments:
        assert fragment not in PAGE_TEXT, f"Forbidden content found: {fragment}"


def test_no_html_injection_sinks_or_dynamic_code() -> None:
    banned = [
        "innerHTML",
        "outerHTML",
        "insertAdjacentHTML",
        "document.write",
        "eval(",
        "new Function",
    ]
    for needle in banned:
        assert needle not in PAGE_TEXT, f"Forbidden sink/code primitive found: {needle}"

    quoted_timer = re.compile(r"set(?:Timeout|Interval)\(\s*[\"'`]")
    assert quoted_timer.search(PAGE_TEXT) is None, "setTimeout/setInterval must not use string code"


def test_single_authenticated_fetch_wrapper() -> None:
    count = PAGE_TEXT.count("fetch(")
    assert count == 1, f"Expected exactly one fetch( occurrence, found {count}"
    assert "Authorization" in SCRIPT_TEXT, "Script must send Authorization header"
    assert '"Bearer "' in SCRIPT_TEXT, 'Script must use "Bearer " token prefix'


def test_uses_every_contract_route() -> None:
    required = [
        "/api/status",
        "/api/tenant/mailboxes",
        "refresh=1",
        "/api/mailcow/check",
        "/api/selection",
        "/api/jobs",
        "/output",
        "/api/reports/latest",
        "/api/overview",
    ]
    missing = [route for route in required if route not in SCRIPT_TEXT]
    assert not missing, "Script is missing required route markers: " + ", ".join(missing)


def test_page_behaviour_markers() -> None:
    markers = [
        "textContent",
        "sessionStorage",
        "history.replaceState",
        "showModal(",
        "beforeunload",
        '"PUT"',
        '"POST"',
    ]
    missing = [marker for marker in markers if marker not in SCRIPT_TEXT]
    assert not missing, "Script is missing behavior markers: " + ", ".join(missing)
    assert "confirm(" not in SCRIPT_TEXT, "Script must use the run sheet, never confirm()"

    assert "prefers-color-scheme" in PAGE_TEXT, "CSS must include prefers-color-scheme"

    has_viewport_meta = any(
        tag == "meta" and attrs.get("name") == "viewport"
        for tag, attrs in TAGS
    )
    assert has_viewport_meta, "Document must include <meta name=viewport>"


def test_expected_controls_exist() -> None:
    ids = [attrs["id"] for _, attrs in TAGS if "id" in attrs]
    counts = Counter(ids)
    dupes = sorted(name for name, count in counts.items() if count > 1)
    assert not dupes, "IDs must be unique; duplicates found: " + ", ".join(dupes)

    required_ids = {
        "token-form",
        "tenant-body",
        "chk-all",
        "f-search",
        "f-kind",
        "f-enabled",
        "bulk-domain",
        "btn-apply-domain",
        "btn-check",
        "btn-save",
        "run-buttons",
        "opt-only",
        "opt-sample",
        "job-output",
        "btn-full-output",
        "steps",
        "outcome",
        "mailboxes-panel",
        "detail-view",
        "job-view",
        "job-progress",
        "run-sheet",
        "run-form",
        "sheet-start",
        "sheet-cancel",
        "opt-mailbox",
        "opt-since",
        "opt-dry",
        "btn-conn-toggle",
        "hdr-job",
        "settings-section",
        "tenant-section",
    }
    missing_ids = sorted(required_ids.difference(ids))
    assert not missing_ids, "Missing required ids: " + ", ".join(missing_ids)

    commands = {
        attrs["data-command"] for tag, attrs in TAGS if tag == "button" and "data-command" in attrs
    }
    expected_commands = {"plan", "provision", "migrate", "verify", "cleanup"}
    assert commands == expected_commands, (
        "Run button commands mismatch; "
        f"got {sorted(commands)}, expected {sorted(expected_commands)}"
    )
    dry_buttons = [attrs for tag, attrs in TAGS if tag == "button" and "data-dry" in attrs]
    assert not dry_buttons, "Run buttons must not carry data-dry; the run sheet chooses dry runs"

    def tags_with_id(element_id: str) -> list[str]:
        return [tag for tag, attrs in TAGS if attrs.get("id") == element_id]

    assert tags_with_id("job-output") == ["pre"], "Element #job-output must be a <pre>"
    assert tags_with_id("run-sheet") == ["dialog"], "Element #run-sheet must be a <dialog>"
    assert tags_with_id("job-progress") == ["progress"], "#job-progress must be a <progress>"

    run_form = [attrs for _, attrs in TAGS if attrs.get("id") == "run-form"]
    assert run_form and run_form[0].get("method") == "dialog", '#run-form must have method="dialog"'

    hdr_job = [attrs for _, attrs in TAGS if attrs.get("id") == "hdr-job"]
    assert hdr_job and hdr_job[0].get("role") == "status", '#hdr-job must have role="status"'

    live = [attrs for _, attrs in TAGS if "aria-live" in attrs]
    assert len(live) >= 5, f"Expected at least five aria-live regions, found {len(live)}"


def test_cancel_is_the_run_sheets_default_button() -> None:
    """Enter in a field of the run sheet submits with the form's first submit button, so that
    button must be Cancel: a run starts only from an explicit press of Start."""
    start = next(i for i, (_, attrs) in enumerate(TAGS) if attrs.get("id") == "run-form")
    submits = [
        attrs
        for tag, attrs in TAGS[start:]
        if (tag == "button" and attrs.get("type", "submit") == "submit")
        or (tag == "input" and attrs.get("type") in ("submit", "image"))
    ]
    assert [attrs.get("id") for attrs in submits] == ["sheet-cancel", "sheet-start"], (
        "The run sheet must have exactly Cancel then Start as its submit buttons"
    )
    assert submits[0].get("value") == "cancel", "The first submit button must cancel"
    assert submits[1].get("value") == "start", "Only #sheet-start may carry the start value"
    assert 'returnValue !== "start"' in SCRIPT_TEXT, "A run must require the start value"


def test_size_encoding_and_no_markers() -> None:
    size = len(PAGE_BYTES)
    assert size <= 120 * 1024, f"Page too large: {size} bytes (max 122880)"

    try:
        PAGE_BYTES.decode("utf-8")
    except UnicodeDecodeError as exc:  # pragma: no cover
        raise AssertionError(f"Page must decode as UTF-8: {exc}") from exc

    for marker in ["TODO", "FIXME", "XXX"]:
        assert marker not in PAGE_TEXT, f"Forbidden marker found: {marker}"
