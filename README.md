# Email Notify — 163邮箱进展监控

Monitor 163.com NetEase email for project progress updates, auto-save to organized folders, and send Windows desktop notifications from WSL/Ubuntu Claude Code.

## Features

- **Auto-filter**: Only watches Genscript orders, logistics, gene synthesis, sequencing results
- **Auto-save**: Downloads attachments and saves emails to organized project folders
- **Auto-notify**: Windows desktop popup with email summary + attachment list + save path
- **Auto-merge**: Same order emails append to existing README instead of creating duplicates
- **Auto-clean**: Junk mail (ads, security notices) auto-marked as read

## Quick Start

```bash
# Check unread emails
NETEASE_USER="your@163.com" NETEASE_PASS="auth_code" python3 email_notify.py

# Preview only (no save, no popup)
NETEASE_USER="your@163.com" NETEASE_PASS="auth_code" python3 email_notify.py --dry

# Process recent emails (catch-up mode)
NETEASE_USER="your@163.com" NETEASE_PASS="auth_code" python3 email_notify.py --catchup 30
```

## File Organization

```
C:\Users\xh\OneDrive\文档\工作\beelab\邮箱\
├── C7449KFMG0_CSFV-PRV_基因合成/
│   ├── README.md
│   ├── 说明.txt
│   ├── Quote-C7449KFMG0.pdf
│   └── CSFV-PRV-pUC57-Kan.dna
├── C555NCDJG0_PIV5-F_基因合成/
└── 生生物流_90519130/
```

## Requirements

- Python 3
- WSL with `/mnt/c/Windows/System32/wscript.exe`
- 163.com email with IMAP enabled
- Depends on [wsl-win-notify](https://github.com/xhy-xh/wsl-win-notify) for Windows popups

## Setup

1. Clone this repo
2. Symlink to Claude Code skills: `ln -s $(pwd) ~/.claude/skills/email-notify`
3. Set credentials as env vars or in Claude Code memory
4. Optional: set up cron for periodic checking

## License

MIT
