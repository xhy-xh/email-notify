# Email Notify — 163邮箱进展监控

Monitor 163.com NetEase email for project progress updates, auto-save to organized folders, and send Windows desktop notifications from WSL/Ubuntu Claude Code.

## Features

- **Auto-filter**: Watches order confirmations, logistics, gene synthesis, sequencing
- **Auto-save**: Downloads attachments and saves emails to organized project folders
- **Auto-notify**: Windows desktop popup with email summary + attachment list + save path
- **Auto-merge**: Same order emails append to existing README, no duplicates
- **Auto-clean**: Junk mail (ads, security notices) auto-marked as read

## Quick Start

```bash
# Set your email credentials
export NETEASE_USER="your@163.com"
export NETEASE_PASS="your_auth_code"

# Check unread emails
python3 email_notify.py

# Preview only (no save, no popup)
python3 email_notify.py --dry

# Process recent emails (catch-up mode)
python3 email_notify.py --catchup 30
```

## Configuration

Set these env vars to customize save locations:

| Variable | Default | Description |
|----------|---------|-------------|
| `EMAIL_SAVE_DIR_WIN` | `C:\Users\<user>\email_archive` | Where emails are saved |
| `EMAIL_POP_DIR_WIN` | `C:\Users\<user>\temp` | Temp dir for popup scripts |

## File Organization

```
{EMAIL_SAVE_DIR}/
├── C1234560_GENE-A_基因合成/
│   ├── README.md           <- structured email summary
│   ├── summary.txt          <- brief summary
│   ├── Quote-C1234560.pdf   <- attachment (original name)
│   ├── OptimizationResult.zip
│   └── sequence.dna
├── Logistics_12345678/
│   └── ...
└── Sequencing/
    └── ...
```

Directory naming convention:
- Orders: `{OrderNumber}_{GeneName}_{Type}` (auto-detected from email)
- Same order emails append to existing README
- Logistics/other: `{Sender}_{Subject}`

## Requirements

- Python 3
- WSL with `/mnt/c/Windows/System32/wscript.exe`
- 163.com email with IMAP enabled

## Setup

1. Clone this repo
2. Symlink to Claude Code skills: `ln -s $(pwd) ~/.claude/skills/email-notify`
3. Set credentials as env vars or in Claude Code memory
4. Optional: set up cron for periodic checking

## License

MIT
