#!/usr/bin/env python3
"""Apple Mail CLI — fast listing via Mail's SQLite index, body/attachments via AppleScript.

Usage:
    python3 apple_mail.py accounts
    python3 apple_mail.py list --days 3
    python3 apple_mail.py list --has-attachments --since 2026-06-01
    python3 apple_mail.py search "invoice|rechnung" --days 30
    python3 apple_mail.py search "flight|booking" --body --since 2025-07-01 --until 2025-09-30
    python3 apple_mail.py show 41030
    python3 apple_mail.py context --days 3
    python3 apple_mail.py attachments 40674
    python3 apple_mail.py save 40674 --index 1 --out ~/Downloads/event.ics

Fast fields come from Mail's SQLite index (subject, date, read, mailbox,
attachments, body preview). Sender + full body require Mail.app (AppleScript).
"""

from __future__ import annotations

import argparse
import json
import re
import sqlite3
import subprocess
import sys
from dataclasses import asdict, dataclass
from datetime import datetime, timedelta
from pathlib import Path
from urllib.parse import unquote

LOCAL_TZ = datetime.now().astimezone().tzinfo

# Separator for group_concat / AppleScript batching: ASCII Unit Separator,
# safe against commas (and anything else printable) in attachment names.
FIELD_SEP = "\x1f"

# Word-boundary matching so "Bin" doesn't exclude "Binder", "Draft" doesn't
# exclude a "Draftjs" folder, etc. Sent is separate so --include-sent can
# re-include it (e.g. to verify what actually went out).
SENT_MAILBOX_RE = re.compile(r"(?<!\w)(Sent(\s+Messages)?|Gesendet(e)?)(?!\w)", re.I)
EXCLUDED_MAILBOX_RE = re.compile(
    r"(?<!\w)(Trash|Bin|Papierkorb|Spam|Drafts?|Entw(ü|u)rfe?|Entwurf|Deleted"
    r"|Junk|Chat Segmentation)(?!\w)",
    re.I,
)

CALENDAR_HINT_RE = re.compile(
    r"(?i)(termin|einladung|invitation|calendar|kalender|meeting|kolloquium|"
    r"deadline|frist|rsvp|webinar|zoom|teams|buchungsbestätigung|boarding pass|"
    r"check-in|flug|bahn|reise|trip|reservierung|appointment|schedule)"
)

ACTION_HINT_RE = re.compile(
    r"(?i)(action required|handlung|bitte|dringend|urgent|frist|deadline|"
    r"bestätig|confirm|sign|unterschreib|zahlung|payment|rechnung|invoice|"
    r"verifiz|verify|wichtig|important|response required|antwort)"
)

ACCOUNT_CACHE: dict[str, str] | None = None


class MailError(RuntimeError):
    pass


@dataclass
class MessageRow:
    id: int
    date: str
    read: bool
    subject: str
    mailbox: str
    account: str
    sender: str = ""
    attachments: list[str] | None = None
    snippet: str | None = None

    @property
    def unread(self) -> bool:
        return not self.read


def find_envelope_index() -> Path:
    mail_root = Path.home() / "Library/Mail"
    if not mail_root.is_dir():
        raise MailError(f"Mail library not found: {mail_root}")
    candidates = sorted(mail_root.glob("V*/MailData/Envelope Index"), reverse=True)
    for path in candidates:
        if path.is_file():
            return path
    raise MailError("Mail Envelope Index not found (is Mail.app configured?)")


def run_osascript(source: str) -> str:
    proc = subprocess.run(
        ["osascript", "-e", source],
        capture_output=True,
        text=True,
    )
    if proc.returncode != 0:
        detail = (proc.stderr or proc.stdout or "osascript failed").strip()
        raise MailError(detail)
    return proc.stdout.strip()


def load_account_map() -> dict[str, str]:
    global ACCOUNT_CACHE
    if ACCOUNT_CACHE is not None:
        return ACCOUNT_CACHE

    raw = run_osascript(
        'tell application "Mail" to get {name, id} of every account'
    )
    parts = [p.strip() for p in raw.split(",") if p.strip()]
    if len(parts) % 2:
        raise MailError(f"Unexpected account list from Mail.app: {raw!r}")

    half = len(parts) // 2
    names = parts[:half]
    uuids = parts[half:]
    mapping = {uuid: name for name, uuid in zip(names, uuids, strict=True)}
    ACCOUNT_CACHE = mapping
    return mapping


def account_for_mailbox_url(url: str, account_map: dict[str, str]) -> str:
    match = re.match(r"^[a-z]+://([^/]+)/", url or "")
    if not match:
        return "unknown"
    return account_map.get(match.group(1), match.group(1))


def mailbox_label(url: str) -> str:
    if not url:
        return "unknown"
    path = url.split("/", 3)[-1] if "://" in url else url
    return unquote(path)


def connect_db(readonly: bool = True) -> sqlite3.Connection:
    db_path = find_envelope_index()
    uri = f"file:{db_path}?mode={'ro' if readonly else 'rw'}"
    conn = sqlite3.connect(uri, uri=True)
    conn.row_factory = sqlite3.Row
    return conn


def cutoff_epoch(days: int) -> int:
    start = datetime.now(tz=LOCAL_TZ) - timedelta(days=max(days, 0))
    return int(start.timestamp())


def parse_date(value: str) -> int:
    """Parse YYYY-MM-DD (local midnight) into an epoch int."""
    try:
        dt = datetime.strptime(value, "%Y-%m-%d").replace(tzinfo=LOCAL_TZ)
    except ValueError as exc:
        raise MailError(f"Invalid date {value!r}; expected YYYY-MM-DD") from exc
    return int(dt.timestamp())


def split_query_terms(query: str | None) -> list[str]:
    """Split an OR-query on '|' into trimmed, non-empty terms."""
    if not query:
        return []
    return [term.strip() for term in query.split("|") if term.strip()]


def like_variants(term: str) -> list[str]:
    """Case variants for the SQLite LIKE pre-pass.

    SQLite LIKE is only case-insensitive for ASCII, so a term like "prüfung"
    would miss "Prüfung" in the narrowing step before the (properly Unicode
    case-insensitive) Python post-filter even runs. Generate the common case
    variants when the term contains non-ASCII characters.
    """
    if term.isascii():
        return [term]
    variants = {term, term.lower(), term.upper(), term.capitalize(), term.title()}
    return sorted(variants)


# --- Matching ---------------------------------------------------------------
# Match modes for `search`:
#   "prefix"    (DEFAULT) word-start boundary: term must begin at a word
#               boundary, but may be followed by more letters. Avoids the
#               "asics" -> "basics" false positive while still matching German
#               inflections ("rechnung" -> "Rechnungen"). Unicode-aware, so
#               umlauts count as word characters.
#   "word"      (--word) strict whole-word: boundary on BOTH sides.
#   "substring" (--substring) the old behaviour: plain substring, anywhere.
#
# SQLite LIKE cannot express word boundaries, so we use LIKE only to *narrow*
# candidate rows, then post-filter in Python with a compiled regex.

def compile_term_regex(term: str, mode: str) -> re.Pattern[str]:
    """Compile a term into a case-insensitive, Unicode-aware matcher."""
    core = re.escape(term)
    if mode == "substring":
        pattern = core
    elif mode == "word":
        pattern = r"(?<!\w)" + core + r"(?!\w)"
    else:  # "prefix" (default)
        pattern = r"(?<!\w)" + core
    return re.compile(pattern, re.IGNORECASE | re.UNICODE)


def text_matches(haystack: str, patterns: list[re.Pattern[str]]) -> bool:
    return any(p.search(haystack) for p in patterns)


def effective_window(
    days: int, since: str | None, until: str | None
) -> tuple[int, int | None, str, str]:
    """Resolve the date window into (since_epoch, until_epoch, start_label, end_label)."""
    since_epoch = parse_date(since) if since else cutoff_epoch(days)
    until_epoch = parse_date(until) + 86400 if until else None  # inclusive day
    start_label = datetime.fromtimestamp(since_epoch, tz=LOCAL_TZ).strftime("%Y-%m-%d")
    end_label = until if until else "now"
    return since_epoch, until_epoch, start_label, end_label


def resolve_account_uuid(account: str | None, account_map: dict[str, str]) -> str | None:
    if not account:
        return None
    lowered = account.casefold()
    for uuid, name in account_map.items():
        if name.casefold() == lowered or uuid.casefold() == lowered:
            return uuid
    known = ", ".join(sorted(account_map.values()))
    raise MailError(f"Unknown account {account!r}. Known accounts: {known}")


def base_message_sql(
    *,
    since_epoch: int,
    until_epoch: int | None,
    account: str | None,
    unread_only: bool,
    query_terms: list[str],
    search_body: bool,
    has_attachments: bool,
    limit: int,
) -> tuple[str, list]:
    account_map = load_account_map()
    account_uuid = resolve_account_uuid(account, account_map)

    sql = [
        "SELECT m.ROWID AS id, m.date_received, m.read, s.subject, mb.url AS mailbox_url,",
        "       su.summary AS snippet,",
        # char(31) separator: attachment names may contain commas. (DISTINCT
        # can't be combined with a custom separator; deduped in Python.)
        "       group_concat(att.name, char(31)) AS attachment_names",
        "FROM messages m",
        "JOIN subjects s ON m.subject = s.ROWID",
        "JOIN mailboxes mb ON m.mailbox = mb.ROWID",
        "LEFT JOIN summaries su ON m.summary = su.ROWID",
        "LEFT JOIN attachments att ON att.message = m.ROWID",
        "WHERE m.date_received >= ?",
    ]
    params: list = [since_epoch]

    if until_epoch is not None:
        sql.append("AND m.date_received < ?")
        params.append(until_epoch)
    if unread_only:
        sql.append("AND m.read = 0")
    if account_uuid:
        sql.append("AND mb.url LIKE ?")
        params.append(f"imap://{account_uuid}/%")
    if query_terms:
        clauses = []
        for term in query_terms:
            for variant in like_variants(term):
                if search_body:
                    clauses.append("(s.subject LIKE ? OR su.summary LIKE ?)")
                    params.extend([f"%{variant}%", f"%{variant}%"])
                else:
                    clauses.append("s.subject LIKE ?")
                    params.append(f"%{variant}%")
        sql.append("AND (" + " OR ".join(clauses) + ")")
    if has_attachments:
        sql.append("AND EXISTS (SELECT 1 FROM attachments x WHERE x.message = m.ROWID)")

    sql.extend(
        [
            "GROUP BY m.ROWID",
            "ORDER BY m.date_received DESC",
            "LIMIT ?",
        ]
    )
    params.append(limit)
    return "\n".join(sql), params


def row_to_message(
    row: sqlite3.Row,
    account_map: dict[str, str],
    *,
    include_sent: bool = False,
    skip_exclusion: bool = False,
) -> MessageRow | None:
    mailbox_url = row["mailbox_url"] or ""
    label = mailbox_label(mailbox_url)
    if not skip_exclusion:
        if EXCLUDED_MAILBOX_RE.search(label):
            return None
        if not include_sent and SENT_MAILBOX_RE.search(label):
            return None

    attachments = []
    if row["attachment_names"]:
        seen: set[str] = set()
        for part in row["attachment_names"].split(FIELD_SEP):
            if part and part not in seen:
                seen.add(part)
                attachments.append(part)

    ts = int(row["date_received"] or 0)
    dt = datetime.fromtimestamp(ts, tz=LOCAL_TZ).strftime("%Y-%m-%d %H:%M")

    snippet = None
    if "snippet" in row.keys() and row["snippet"]:
        snippet = clean_snippet(row["snippet"])

    return MessageRow(
        id=int(row["id"]),
        date=dt,
        read=bool(row["read"]),
        subject=row["subject"] or "",
        mailbox=mailbox_label(mailbox_url),
        account=account_for_mailbox_url(mailbox_url, account_map),
        sender="",
        attachments=attachments or None,
        snippet=snippet,
    )


def clean_snippet(text: str, limit: int = 160) -> str:
    """Collapse the noisy Mail preview text into a short single-line snippet."""
    # Mail pads previews with zero-width / soft-hyphen filler; strip it.
    cleaned = text.replace("\u00ad", " ").replace("\u200c", " ").replace("\u200b", " ")
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    if len(cleaned) > limit:
        cleaned = cleaned[:limit].rstrip() + "…"
    return cleaned


def fetch_messages(
    *,
    days: int,
    account: str | None,
    unread_only: bool,
    query: str | None,
    limit: int,
    since: str | None = None,
    until: str | None = None,
    search_body: bool = False,
    has_attachments: bool = False,
    with_sender: bool = False,
    match_mode: str = "substring",
    include_sent: bool = False,
) -> list[MessageRow]:
    account_map = load_account_map()
    since_epoch, until_epoch, _start, _end = effective_window(days, since, until)
    terms = split_query_terms(query)

    # Word/prefix modes need a Python post-filter, so the SQL LIKE acts only as a
    # coarse narrowing step. Fetch a generous candidate cap (within the window)
    # so a buried match isn't dropped before the regex filter runs.
    post_filter = bool(terms) and match_mode != "substring"
    sql_limit = max(limit * 50, 2000) if post_filter else limit * 3

    sql, params = base_message_sql(
        since_epoch=since_epoch,
        until_epoch=until_epoch,
        account=account,
        unread_only=unread_only,
        query_terms=terms,
        search_body=search_body,
        has_attachments=has_attachments,
        limit=sql_limit,
    )
    with connect_db() as conn:
        rows = conn.execute(sql, params).fetchall()

    if post_filter:
        patterns = [compile_term_regex(t, match_mode) for t in terms]
        kept = []
        for row in rows:
            # Match against the FULL stored summary, not the truncated snippet.
            haystack = row["subject"] or ""
            if search_body and "snippet" in row.keys() and row["snippet"]:
                haystack += "\n" + row["snippet"]
            if text_matches(haystack, patterns):
                kept.append(row)
        rows = kept

    messages: list[MessageRow] = []
    for row in rows:
        msg = row_to_message(row, account_map, include_sent=include_sent)
        if msg is None:
            continue
        messages.append(msg)
        if len(messages) >= limit:
            break

    if with_sender:
        enrich_senders(messages)
    return messages


def full_body_search(
    *,
    days: int,
    account: str | None,
    unread_only: bool,
    query: str,
    limit: int,
    since: str | None,
    until: str | None,
    has_attachments: bool,
    with_sender: bool,
    match_mode: str,
    body_cap: int,
    body_chars: int,
    include_sent: bool = False,
) -> tuple[list[MessageRow], bool, int]:
    """Slow opt-in search over real message bodies (via AppleScript).

    Subject + stored preview are checked first (free). Only candidates that do
    NOT already match get their full body fetched, capped at ``body_cap`` body
    loads. Returns (messages, truncated, bodies_fetched); ``truncated`` is True
    if the cap stopped us before all candidates were body-checked.
    """
    account_map = load_account_map()
    since_epoch, until_epoch, _start, _end = effective_window(days, since, until)
    terms = split_query_terms(query)
    patterns = [compile_term_regex(t, match_mode) for t in terms]

    # No subject/summary LIKE here: a body-only match has neither.
    sql, params = base_message_sql(
        since_epoch=since_epoch,
        until_epoch=until_epoch,
        account=account,
        unread_only=unread_only,
        query_terms=[],
        search_body=False,
        has_attachments=has_attachments,
        limit=max(body_cap * 10, 1000),
    )
    with connect_db() as conn:
        rows = conn.execute(sql, params).fetchall()

    messages: list[MessageRow] = []
    bodies_fetched = 0
    truncated = False
    for row in rows:
        msg = row_to_message(row, account_map, include_sent=include_sent)
        if msg is None:
            continue
        haystack = row["subject"] or ""
        if "snippet" in row.keys() and row["snippet"]:
            haystack += "\n" + row["snippet"]
        if text_matches(haystack, patterns):
            messages.append(msg)
        elif bodies_fetched < body_cap:
            bodies_fetched += 1
            try:
                body = fetch_message_body(msg, body_chars)
            except MailError:
                body = ""
            if body and text_matches(body, patterns):
                messages.append(msg)
        else:
            truncated = True
        if len(messages) >= limit:
            break

    if with_sender:
        enrich_senders(messages)
    return messages, truncated, bodies_fetched


def enrich_senders(messages: list[MessageRow]) -> None:
    """Fill in senders via one batched AppleScript round-trip.

    Messages are grouped by (account, mailbox) so Mail runs ONE `whose id is
    A or id is B ...` query per mailbox instead of one per message (each
    `whose` query costs ~1s; Mail does not support `id is in {...}`). Returns
    id/sender record pairs; anything Mail can't resolve degrades to "". Falls
    back to per-message lookups only if the batch call itself errors out.
    """
    if not messages:
        return

    groups: dict[tuple[str, str], list[int]] = {}
    for msg in messages:
        groups.setdefault((msg.account, msg.mailbox), []).append(msg.id)

    blocks = []
    for (account, mailbox), ids in groups.items():
        acc = escape_applescript(account)
        mbx = escape_applescript(mailbox)
        id_clause = " or ".join(f"id is {i}" for i in ids)
        blocks.append(
            f'''  try
    set msgs to (every message of mailbox "{mbx}" of account "{acc}" whose {id_clause})
    repeat with m in msgs
      set end of out to ((id of m) as text) & (ASCII character 30) & (sender of m)
    end repeat
  end try'''
        )
    source = (
        'tell application "Mail"\n'
        "  set out to {}\n"
        + "\n".join(blocks)
        + "\n  set AppleScript's text item delimiters to (ASCII character 31)\n"
        "  return out as text\n"
        "end tell"
    )
    try:
        raw = run_osascript(source)
        by_id: dict[int, str] = {}
        for record in raw.split(FIELD_SEP):
            if "\x1e" not in record:
                continue
            msg_id, _, sender = record.partition("\x1e")
            try:
                by_id[int(msg_id.strip())] = sender.strip()
            except ValueError:
                continue
        for msg in messages:
            msg.sender = by_id.get(msg.id, "")
    except MailError:
        for msg in messages:
            try:
                msg.sender = fetch_sender(msg)
            except MailError:
                msg.sender = ""


def fetch_message_meta(message_id: int) -> MessageRow:
    account_map = load_account_map()
    with connect_db() as conn:
        row = conn.execute(
            """
            SELECT m.ROWID AS id, m.date_received, m.read, s.subject, mb.url AS mailbox_url,
                   NULL AS snippet,
                   group_concat(att.name, char(31)) AS attachment_names
            FROM messages m
            JOIN subjects s ON m.subject = s.ROWID
            JOIN mailboxes mb ON m.mailbox = mb.ROWID
            LEFT JOIN attachments att ON att.message = m.ROWID
            WHERE m.ROWID = ?
            GROUP BY m.ROWID
            """,
            (message_id,),
        ).fetchone()
    if row is None:
        raise MailError(f"No message with id {message_id}")
    # Explicit ID = deliberate: no mailbox exclusion (works for Sent etc.).
    msg = row_to_message(row, account_map, skip_exclusion=True)
    assert msg is not None
    return msg


def escape_applescript(value: str) -> str:
    return value.replace("\\", "\\\\").replace('"', '\\"')


def fetch_message_body(message: MessageRow, max_chars: int) -> str:
    account = escape_applescript(message.account)
    mailbox = escape_applescript(message.mailbox)
    source = f'''
tell application "Mail"
  set targetAccount to account "{account}"
  set targetMailbox to mailbox "{mailbox}" of targetAccount
  set msgRef to first message of targetMailbox whose id is {message.id}
  set bodyText to content of msgRef
  if length of bodyText > {max_chars} then
    set bodyText to text 1 thru {max_chars} of bodyText
  end if
  return bodyText
end tell
'''
    try:
        return run_osascript(source)
    except MailError:
        fallback = f'''
tell application "Mail"
  repeat with targetAccount in accounts
    repeat with targetMailbox in mailboxes of targetAccount
      try
        set msgRef to first message of targetMailbox whose id is {message.id}
        set bodyText to content of msgRef
        if length of bodyText > {max_chars} then
          set bodyText to text 1 thru {max_chars} of bodyText
        end if
        return bodyText
      end try
    end repeat
  end repeat
end tell
'''
        return run_osascript(fallback)


def fetch_sender(message: MessageRow) -> str:
    if message.sender:
        return message.sender
    account = escape_applescript(message.account)
    mailbox = escape_applescript(message.mailbox)
    source = f'''
tell application "Mail"
  set msgRef to first message of mailbox "{mailbox}" of account "{account}" whose id is {message.id}
  return sender of msgRef
end tell
'''
    try:
        return run_osascript(source)
    except MailError:
        return ""


def save_attachment(message: MessageRow, index: int, out_path: Path) -> Path:
    out_path = out_path.expanduser().resolve()
    out_path.parent.mkdir(parents=True, exist_ok=True)
    account = escape_applescript(message.account)
    mailbox = escape_applescript(message.mailbox)
    posix_out = escape_applescript(str(out_path))
    source = f'''
tell application "Mail"
  set msgRef to first message of mailbox "{mailbox}" of account "{account}" whose id is {message.id}
  set attList to mail attachments of msgRef
  if (count of attList) < {index} then error "Attachment index out of range"
  set att to item {index} of attList
  save att in POSIX file "{posix_out}"
  return name of att
end tell
'''
    name = run_osascript(source)
    if not out_path.exists():
        raise MailError(f"Attachment save reported success but file missing: {out_path}")
    print(f"saved {name} -> {out_path}")
    return out_path


def has_calendar_hint(message: MessageRow) -> bool:
    haystack = f"{message.subject} {message.snippet or ''}"
    if CALENDAR_HINT_RE.search(haystack):
        return True
    if message.attachments and any(name.lower().endswith(".ics") for name in message.attachments):
        return True
    return False


def has_action_hint(message: MessageRow) -> bool:
    haystack = f"{message.subject} {message.snippet or ''}"
    return bool(ACTION_HINT_RE.search(haystack))


def format_attachment_list(names: list[str], shown: int = 3) -> str:
    head = ", ".join(names[:shown])
    extra = len(names) - shown
    return f"{head} (+{extra} more)" if extra > 0 else head


def print_messages(messages: list[MessageRow]) -> None:
    if not messages:
        print("No messages matched.")
        return
    for msg in messages:
        unread = "unread" if msg.unread else "read"
        sender = f" | {msg.sender}" if msg.sender else ""
        print(f"[{msg.id}] {msg.date} | {msg.account} | {unread}{sender}")
        print(f"  {msg.subject}")
        if msg.snippet:
            print(f"  > {msg.snippet}")
        if msg.attachments:
            print(f"  attachments: {format_attachment_list(msg.attachments)}")
    print(f"\n{len(messages)} message(s).")


def cmd_accounts(args: argparse.Namespace) -> int:
    account_map = load_account_map()
    rows = [{"uuid": uuid, "name": name} for uuid, name in account_map.items()]
    rows.sort(key=lambda row: row["name"].casefold())
    if args.json:
        print(json.dumps(rows, indent=2, ensure_ascii=False))
        return 0
    for row in rows:
        print(f"{row['name']}\t{row['uuid']}")
    return 0


def cmd_mailboxes(args: argparse.Namespace) -> int:
    account_map = load_account_map()
    account_filter = args.account
    uuid = None
    if account_filter:
        for key, name in account_map.items():
            if name.casefold() == account_filter.casefold() or key.casefold() == account_filter.casefold():
                uuid = key
                break
        if not uuid:
            raise MailError(f"Unknown account {account_filter!r}")

    sql = "SELECT url, total_count, unread_count FROM mailboxes"
    params: list = []
    if uuid:
        sql += " WHERE url LIKE ?"
        params.append(f"imap://{uuid}/%")
    sql += " ORDER BY url"

    with connect_db() as conn:
        rows = conn.execute(sql, params).fetchall()

    output = []
    for row in rows:
        url = row["url"] or ""
        output.append(
            {
                "account": account_for_mailbox_url(url, account_map),
                "mailbox": mailbox_label(url),
                "url": url,
                "total": int(row["total_count"] or 0),
                "unread": int(row["unread_count"] or 0),
            }
        )

    if args.json:
        print(json.dumps(output, indent=2, ensure_ascii=False))
        return 0

    for row in output:
        print(
            f"{row['account']}\t{row['mailbox']}\t"
            f"total={row['total']}\tunread={row['unread']}"
        )
    return 0


def cmd_list(args: argparse.Namespace) -> int:
    messages = fetch_messages(
        days=args.days,
        account=args.account,
        unread_only=args.unread,
        query=None,
        limit=args.limit,
        since=args.since,
        until=args.until,
        has_attachments=args.has_attachments,
        with_sender=args.with_sender,
        include_sent=args.include_sent,
    )
    if args.json:
        print(json.dumps([asdict(m) for m in messages], indent=2, ensure_ascii=False))
        return 0
    print_messages(messages)
    return 0


def cmd_search(args: argparse.Namespace) -> int:
    match_mode = "substring" if args.substring else ("word" if args.word else "prefix")
    since_epoch, until_epoch, start_label, end_label = effective_window(
        args.days, args.since, args.until
    )

    if args.full_body:
        scope = "full-body"
    elif args.body:
        scope = "subject+preview"
    else:
        scope = "subject"

    header = (
        f"search '{args.query}' | window {start_label}..{end_label} "
        f"| scope={scope} | match={match_mode}"
    )

    truncated = False
    if args.full_body:
        messages, truncated, _fetched = full_body_search(
            days=args.days,
            account=args.account,
            unread_only=args.unread,
            query=args.query,
            limit=args.limit,
            since=args.since,
            until=args.until,
            has_attachments=args.has_attachments,
            with_sender=args.with_sender,
            match_mode=match_mode,
            body_cap=args.full_body_cap,
            body_chars=args.full_body_chars,
            include_sent=args.include_sent,
        )
    else:
        messages = fetch_messages(
            days=args.days,
            account=args.account,
            unread_only=args.unread,
            query=args.query,
            limit=args.limit,
            since=args.since,
            until=args.until,
            search_body=args.body,
            has_attachments=args.has_attachments,
            with_sender=args.with_sender,
            match_mode=match_mode,
            include_sent=args.include_sent,
        )

    trunc_note = (
        f"note: --full-body cap of {args.full_body_cap} body loads reached; "
        f"some candidates were not full-text checked. Narrow the date range or "
        f"raise --full-body-cap."
    ) if truncated else None

    empty_note = None
    if not messages:
        hints = ["widen with --since YYYY-MM-DD"]
        if not args.body and not args.full_body:
            hints.append("add --body for preview text or --full-body for full text")
        if match_mode != "substring":
            hints.append("try --substring for partial-word matches")
        empty_note = (
            f"0 results in window {start_label}..{end_label} "
            f"(scope={scope}, match={match_mode}). Try: " + "; ".join(hints) + "."
        )

    if args.json:
        # Keep stdout pure JSON: scope/hints go to stderr.
        print(header, file=sys.stderr)
        if trunc_note:
            print(trunc_note, file=sys.stderr)
        if empty_note:
            print(empty_note, file=sys.stderr)
        print(json.dumps([asdict(m) for m in messages], indent=2, ensure_ascii=False))
        return 0

    print(header)
    if trunc_note:
        print(trunc_note)
    print_messages(messages)
    if empty_note:
        print(empty_note)
    return 0


def cmd_show(args: argparse.Namespace) -> int:
    message = fetch_message_meta(args.id)
    sender = fetch_sender(message)
    body = fetch_message_body(message, args.body_chars)

    payload = {
        **asdict(message),
        "sender": sender or message.sender,
        "body": body,
    }
    if args.json:
        print(json.dumps(payload, indent=2, ensure_ascii=False))
        return 0

    print(f"id: {message.id}")
    print(f"date: {message.date}")
    print(f"account: {message.account}")
    print(f"mailbox: {message.mailbox}")
    print(f"read: {'yes' if message.read else 'no'}")
    if sender:
        print(f"from: {sender}")
    print(f"subject: {message.subject}")
    if message.attachments:
        print(f"attachments: {', '.join(message.attachments)}")
    print("\n--- body ---")
    print(body)
    return 0


def cmd_attachments(args: argparse.Namespace) -> int:
    message = fetch_message_meta(args.id)
    names = message.attachments or []
    if args.json:
        print(json.dumps({"id": message.id, "attachments": names}, indent=2, ensure_ascii=False))
        return 0
    if not names:
        print("No attachments.")
        return 0
    for idx, name in enumerate(names, start=1):
        print(f"{idx}\t{name}")
    return 0


def cmd_save(args: argparse.Namespace) -> int:
    message = fetch_message_meta(args.id)
    out = Path(args.out)
    save_attachment(message, args.index, out)
    return 0


def cmd_context(args: argparse.Namespace) -> int:
    # Fetch generously: the three sections filter this pool independently, so a
    # mail-heavy window must not push unread/action mails out of the candidates.
    messages = fetch_messages(
        days=args.days,
        account=args.account,
        unread_only=False,
        query=None,
        limit=max(args.limit * 5, 200),
        since=args.since,
        until=args.until,
        include_sent=args.include_sent,
    )
    unread = [m for m in messages if m.unread][: args.limit]
    calendar = [m for m in messages if has_calendar_hint(m)][: args.limit]
    action = [m for m in messages if has_action_hint(m)][: args.limit]

    payload = {
        "days": args.days,
        "generated_at": datetime.now(tz=LOCAL_TZ).isoformat(timespec="seconds"),
        "unread": [asdict(m) for m in unread],
        "calendar_hints": [asdict(m) for m in calendar],
        "action_hints": [asdict(m) for m in action],
    }
    if args.json:
        print(json.dumps(payload, indent=2, ensure_ascii=False))
        return 0

    print(f"Apple Mail context (last {args.days} days)")
    print()

    def section(title: str, items: list[MessageRow]) -> None:
        print(title)
        if not items:
            print("  (none)")
            print()
            return
        for msg in items:
            flag = "UNREAD" if msg.unread else "read"
            attach = (
                f" + {format_attachment_list(msg.attachments, shown=2)}"
                if msg.attachments
                else ""
            )
            print(f"  [{msg.id}] {msg.date} | {msg.account} | {flag} | {msg.subject}{attach}")
        print()

    section(f"Unread ({len(unread)})", unread)
    section(f"Calendar / travel hints ({len(calendar)})", calendar)
    section(f"Action / reply hints ({len(action)})", action)
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Read Apple Mail locally (SQLite index + AppleScript for bodies).",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python3 apple_mail.py accounts
  python3 apple_mail.py list --days 3 --unread
  python3 apple_mail.py list --has-attachments --since 2026-06-01
  python3 apple_mail.py search "Invoice|rechnung" --days 60
  python3 apple_mail.py search "flight|booking" --body --since 2025-07-01 --until 2025-09-30
  python3 apple_mail.py show 41030 --body-chars 800
  python3 apple_mail.py context --days 5
  python3 apple_mail.py save 40674 --index 1 --out ~/Downloads/event.ics

Notes:
  - Fast fields (SQLite index): subject, date, read state, mailbox, attachments, body preview.
  - Sender + full body need Mail.app (AppleScript): automatic in `show`, opt-in via --with-sender.
""",
    )
    sub = parser.add_subparsers(dest="command", required=True)

    common = argparse.ArgumentParser(add_help=False)
    common.add_argument("--json", action="store_true", help="Print JSON output")
    common.add_argument(
        "--plain",
        action="store_true",
        help="Plain text output (default; kept for compatibility)",
    )

    p_accounts = sub.add_parser("accounts", parents=[common], help="List Mail accounts")
    p_accounts.set_defaults(func=cmd_accounts)

    p_mailboxes = sub.add_parser("mailboxes", parents=[common], help="List mailboxes")
    p_mailboxes.add_argument("--account", help="Filter by account name or UUID")
    p_mailboxes.set_defaults(func=cmd_mailboxes)

    list_parent = argparse.ArgumentParser(add_help=False)
    list_parent.add_argument("--days", type=int, default=7, help="Lookback window in days (default: 7)")
    list_parent.add_argument("--since", help="Start date YYYY-MM-DD (overrides --days)")
    list_parent.add_argument("--until", help="End date YYYY-MM-DD inclusive (use with --since)")
    list_parent.add_argument("--account", help="Filter by account name or UUID")
    list_parent.add_argument("--limit", type=int, default=20, help="Max rows (default: 20)")
    list_parent.add_argument("--unread", action="store_true", help="Unread only")
    list_parent.add_argument(
        "--has-attachments",
        dest="has_attachments",
        action="store_true",
        help="Only messages that have attachments",
    )
    list_parent.add_argument(
        "--with-sender",
        dest="with_sender",
        action="store_true",
        help="Resolve senders via one batched AppleScript call (SQLite has no sender)",
    )
    list_parent.add_argument(
        "--include-sent",
        dest="include_sent",
        action="store_true",
        help="Also include Sent/Gesendet mailboxes (e.g. to verify outgoing mail); "
        "Trash/Spam/Drafts stay excluded",
    )

    p_list = sub.add_parser("list", parents=[common, list_parent], help="List recent messages")
    p_list.set_defaults(func=cmd_list)

    p_search = sub.add_parser(
        "search",
        parents=[common, list_parent],
        help="Search subjects (and optionally body previews / full body)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
The query matches the subject by default. Use '|' for OR terms.

Match modes (case-insensitive, Unicode/umlaut-aware):
  default    word-START boundary: term must begin at a word boundary but may be
             followed by more letters. So "asics" does NOT match "basics", but
             "rechnung" still matches "Rechnungen".
  --word     strict whole word: boundary on both sides ("asics" != "asics2").
  --substring  old behaviour: plain substring anywhere (matches "basics").

Search scope:
  (default)    subject only.
  --body       ALSO match Mail's stored body PREVIEW (~1000 chars/mail, fast).
               WARNING: this is a truncated summary, NOT the full body — text
               deeper in long emails is invisible to it.
  --full-body  slow opt-in: load each candidate's real body via Mail.app and
               grep the full text (capped by --full-body-cap). Use a tight
               date range; subject/preview matches are still found for free.

Default lookback is 365 days (override with --days / --since / --until).

  search "invoice|rechnung|payment" --days 90
  search "flight|booking|reservation" --body --since 2025-07-01 --until 2025-09-30
  search "asics" --body --since 2023-01-01
  search "vertragsnummer" --full-body --since 2025-01-01 --full-body-cap 60
""",
    )
    p_search.add_argument(
        "query",
        help="Search pattern; split OR terms with '|' (e.g. \"invoice|rechnung\")",
    )
    p_search.add_argument(
        "--body",
        action="store_true",
        help="Also match the body PREVIEW (Mail's stored ~1000-char summary, fast; truncated)",
    )
    p_search.add_argument(
        "--full-body",
        dest="full_body",
        action="store_true",
        help="Slow: load each candidate's real body via Mail.app and search full text",
    )
    p_search.add_argument(
        "--full-body-cap",
        dest="full_body_cap",
        type=int,
        default=40,
        help="Max number of bodies to fetch in --full-body mode (default: 40)",
    )
    p_search.add_argument(
        "--full-body-chars",
        dest="full_body_chars",
        type=int,
        default=8000,
        help="Max characters of each body to fetch in --full-body mode (default: 8000)",
    )
    mode_group = p_search.add_mutually_exclusive_group()
    mode_group.add_argument(
        "--word",
        action="store_true",
        help="Strict whole-word match (boundary on both sides)",
    )
    mode_group.add_argument(
        "--substring",
        action="store_true",
        help="Plain substring match anywhere (old behaviour; may over-match)",
    )
    p_search.set_defaults(func=cmd_search, days=365)

    p_show = sub.add_parser("show", parents=[common], help="Show one message")
    p_show.add_argument("id", type=int, help="Message id (ROWID from list/search)")
    p_show.add_argument(
        "--body-chars",
        type=int,
        default=1200,
        help="Max body characters to fetch via AppleScript (default: 1200)",
    )
    p_show.set_defaults(func=cmd_show)

    p_attach = sub.add_parser("attachments", parents=[common], help="List attachments")
    p_attach.add_argument("id", type=int)
    p_attach.set_defaults(func=cmd_attachments)

    p_save = sub.add_parser("save", help="Save one attachment via Mail.app")
    p_save.add_argument("id", type=int)
    p_save.add_argument("--index", type=int, default=1, help="Attachment index (default: 1)")
    p_save.add_argument("--out", required=True, help="Output file path")
    p_save.set_defaults(func=cmd_save)

    p_context = sub.add_parser("context", parents=[common, list_parent], help="Agent-oriented briefing")
    p_context.set_defaults(func=cmd_context)

    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()
    try:
        return args.func(args)
    except MailError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
