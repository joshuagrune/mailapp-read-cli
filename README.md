# mailapp-read-cli

Read-only macOS helper for **Mail.app** — list, search, and inspect mail locally via SQLite + AppleScript.

Companion to [mailapp-send-cli](https://github.com/joshuagrune/mailapp-send-cli) (outgoing mail).

## Requirements

- macOS with Mail.app configured
- Python 3.10+
- Automation permission for Terminal/Cursor when using body, sender, or attachment commands

## How it works

| Layer | Role |
| ----- | ---- |
| **SQLite** (`~/Library/Mail/V*/MailData/Envelope Index`) | Fast: subject, date, read state, mailbox, attachment names, body preview |
| **AppleScript** | Full message body, sender, save attachments |

On many Mail versions the `sender_addresses` table is empty — **From** comes from AppleScript only. Fast search covers **subject + body preview**, not sender.

## Usage

```bash
python3 apple_mail.py accounts
python3 apple_mail.py mailboxes [--account NAME]
python3 apple_mail.py list [--days 7] [--unread] [--has-attachments] [--with-sender]
python3 apple_mail.py search "term1|term2" [--body] [--full-body] [--since YYYY-MM-DD]
python3 apple_mail.py show MESSAGE_ID [--body-chars 1200]
python3 apple_mail.py context --days 3
python3 apple_mail.py attachments MESSAGE_ID
python3 apple_mail.py save MESSAGE_ID --index 1 --out ~/Downloads/file.ics
```

Add `--json` on read commands for structured output.

### Search highlights

- **OR terms:** `"invoice|rechnung"`
- **Default match:** word-start boundary (`asics` does not match "basics")
- **`--body`:** also match Mail's ~1000-char body preview (fast, truncated)
- **`--full-body`:** load real bodies via Mail.app (slow; use tight date windows)
- **`--include-sent`:** re-include Sent/Gesendet mailboxes (Gmail sent copies often live in All Mail instead)

Listing excludes Trash/Spam/Drafts by default. Message **id** = SQLite `ROWID` from `list`/`search`.

## Agent / automation

`context --days 3` returns unread mail plus calendar/travel and action hints — useful for daily briefings. Read-only by default; only `save` writes a file.

For **sending** mail via Mail.app, use [mailapp-send-cli](https://github.com/joshuagrune/mailapp-send-cli) instead.

## License

MIT
