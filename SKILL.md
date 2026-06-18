---
name: email-notify
description: Monitor 163.com NetEase email for project progress updates, auto-save to organized folders with attachments, and send Windows desktop notifications. Use when the user asks to check email, 查邮件, 看邮件, 邮箱, 订单进度, 基因合成进展, 物流, or wants email-to-desktop-notification automation.
version: 1.0.0
author: Xiang Haiyuan
---

# Email Notify — 163邮箱进展监控 + Windows弹窗

Checks 163.com NetEase email for progress-related emails (金斯瑞/金唯智/生生物流), saves them to organized project folders, downloads attachments, and sends Windows desktop popups with summaries.

## Architecture

```
Claude Code (WSL) ──> email_notify.py ──> IMAP (163.com)
                                      ──> Save to OneDrive/beelab/邮箱/
                                      ──> notify.py ──> Windows MsgBox
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
NETEASE_USER="13056545170@163.com" NETEASE_PASS="<from memory>" python3 /home/xh/.email_notify.py
```

### Catch-up mode (process recent N emails)

```bash
NETEASE_USER="..." NETEASE_PASS="..." python3 /home/xh/.email_notify.py --catchup 30
```

### Dry run (preview only)

```bash
NETEASE_USER="..." NETEASE_PASS="..." python3 /home/xh/.email_notify.py --dry
```

### Set up auto-check

Use Claude Code cron or /loop:

```
/loop 12h 检查163邮箱进展邮件
```

Or schedule via cron:

```bash
# Every 12 hours at :07 past the hour
*/12 * * * * NETEASE_USER="..." NETEASE_PASS="..." python3 /home/xh/.email_notify.py
```

## File organization

Emails are saved to `C:\Users\xh\OneDrive\文档\工作\beelab\邮箱\` with this structure:

```
邮箱/
├── C7449KFMG0_CSFV-PRV_基因合成/
│   ├── README.md              ← structured email summary
│   ├── 说明.txt                ← Chinese summary
│   ├── Quote-C7449KFMG0.pdf   ← attachment (original name)
│   ├── GB_Final_C7449KFMG0.zip
│   └── CSFV-PRV-pUC57-Kan.dna
├── C555NCDJG0_PIV5-F_基因合成/
├── C563PGKGG0_CSFV-E2-PRV-gD_基因合成/
├── 生生物流_90519130/
└── 金唯智测序/
```

Directory naming convention:
- Orders: `{订单号}_{基因名}_{业务类型}` (e.g. `C7449KFMG0_CSFV-PRV_基因合成`)
- Same order emails append to existing README
- Logistics/other: `{发件人}_{主题关键词}`

## Monitored senders

| Type | Sender | Example |
|------|--------|---------|
| 金斯瑞技术支持 | tech-caoyang@genscript.com.cn | 报价、订单安排 |
| 金斯瑞订单状态 | yori.yu@genscript.com | 进度更新 |
| 金斯瑞验收 | jinhe1.wang@genscript.com | 订单验收确认 |
| 生生物流 | noreply@ashsh.cn | 物流/温度记录 |
| 金唯智测序 | genewiz/azenta | 测序结果 |

## Filtered out (auto-mark read)

- 网易邮箱会员广告 (`member@service.netease.com`)
- 网易安全提醒 (`safe@service.netease.com`)
- giffgaff 手机卡邮件

## Popup format

When progress emails are found, a Windows MsgBox pops up with:

```
邮件提醒 - X封新邮件
─────────────────
发件人: Caoyang Genscript
主题: 报价：C7449KFMG0 密码子优化质粒构建

附件:
  + Quote-C7449KFMG0.pdf
  + Order-C7449KFMG0.txt

已保存至:
  C:\Users\xh\OneDrive\文档\工作\beelab\邮箱\C7449KFMG0_CSFV-PRV_基因合成
```

No popup if no progress emails found.

## Dependencies

- Python 3 with `openpyxl` (not strictly required, used for reading primer inventory)
- WSL with `/mnt/c/Windows/System32/wscript.exe` accessible
- IMAP access to `imap.163.com:993`
- Email credentials in memory (`memory/netemail_credentials.md`)

## Notes

- 163/Coremail requires IMAP ID command before LOGIN
- VBScript popups must be GBK-encoded for Chinese text
- Same-size attachments are not re-downloaded
- Email index is relative (oldest=1), changes as emails arrive
