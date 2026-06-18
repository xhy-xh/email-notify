---
name: email-notify
description: Monitor 163.com NetEase email for project progress updates, auto-save to organized folders with attachments, and send Windows desktop notifications. Use when the user asks to check email, 查邮件, 看邮件, 邮箱, 订单进度, 基因合成进展, 物流, or wants email-to-desktop-notification automation.
version: 1.0.0
---

# Email Notify — 163邮箱进展监控 + Windows弹窗

Checks 163.com NetEase email for progress-related emails, saves them to organized project folders, downloads attachments, and sends Windows desktop popups with summaries.

## Architecture

```
Claude Code (WSL) --> email_notify.py --> IMAP (163.com)
                                      --> Save to project folder
                                      --> Windows MsgBox popup
```

## When to use

Trigger when the user says any of:
- "查邮件" / "看邮件" / "邮箱"
- "检查邮箱" / "有没有新邮件"
- "订单进度" / "基因合成进展" / "物流"
- "邮件提醒" / "收邮件"

## How to use

### Quick check

```bash
NETEASE_USER="your@163.com" NETEASE_PASS="<auth_code>" python3 email_notify.py
```

### Catch-up mode (process recent N emails)

```bash
NETEASE_USER="..." NETEASE_PASS="..." python3 email_notify.py --catchup 30
```

### Dry run (preview only)

```bash
NETEASE_USER="..." NETEASE_PASS="..." python3 email_notify.py --dry
```

### Set up auto-check

Use Claude Code cron or /loop:

```
/loop 12h 检查163邮箱进展邮件
```

## File organization

Emails are saved to the configured directory with this structure:

```
{EMAIL_SAVE_DIR}/
├── ORDER0000_GENE-A_基因合成/
│   ├── README.md              <- structured email summary
│   ├── summary.txt             <- brief summary
│   ├── Quote-ORDER0000.pdf    <- attachment (original name)
│   ├── OptimizationResult.zip
│   └── sequence.dna
├── ORDER1111_GENE-B_基因合成/
├── Logistics_12345678/
└── Sequencing/
```

Directory naming convention:
- Orders: `{OrderNumber}_{GeneName}_{Type}`
- Same order emails append to existing README
- Logistics/other: `{Sender}_{Subject}`

## Configuration

Set env vars to customize save paths:

| Variable | Description |
|----------|-------------|
| `EMAIL_SAVE_DIR_WIN` | Windows path for saving emails |
| `EMAIL_POP_DIR_WIN` | Temp dir for VBS popup scripts |
| `NETEASE_USER` | 163.com email address |
| `NETEASE_PASS` | 163.com IMAP auth code |

## Monitored senders

Customize `WATCH_SENDERS` in the script. Default watches for:
- Gene synthesis order confirmations
- Sequencing results
- Logistics/delivery notifications
- Order quotes and status updates

## Filtered out (auto-mark read)

Customize `SKIP_SENDERS` in the script. Default skips:
- NetEase member ads
- Security notice emails
- Other known junk senders

## Popup format

When progress emails are found, a Windows MsgBox pops up with:

```
Email - 2 new
------------------------------
Sender: Company Support
Subject: Quote ORDER0000 gene synthesis

Attachments:
  + Quote-ORDER0000.pdf
  + Order-ORDER0000.txt

Saved to:
  C:\Users\...\email_archive\ORDER0000_GENE-A_基因合成
```

No popup if no progress emails found.

## Dependencies

- Python 3
- WSL with `/mnt/c/Windows/System32/wscript.exe` accessible
- IMAP access to `imap.163.com:993`
- Email credentials in env vars or memory

## Notes

- 163/Coremail requires IMAP ID command before LOGIN
- VBScript popups must be GBK-encoded for Chinese text
- Same-size attachments are not re-downloaded
- Email index is relative (oldest=1), changes as emails arrive
