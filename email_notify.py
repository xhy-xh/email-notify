#!/usr/bin/env python3
r"""
Check 163 email for progress-related unread -> save to project folder -> Windows popup.

For each relevant email:
  - If order number found (e.g. C1234ABCD0): dirname = order_gene_type
  - Otherwise: dirname = sender_subject
  - Saves README.md (structured summary) + summary.txt
  - Downloads all attachments with original readable filenames
  - If directory already exists, appends new email to existing README
  - Sends Windows popup summarizing what was saved and where

Usage:
  python3 email_notify.py              # check unread only
  python3 email_notify.py --dry        # preview only
  python3 email_notify.py --catchup N  # process last N emails (even if read)

Configuration via env vars:
  EMAIL_SAVE_DIR_WIN  - Windows path for saving emails (default: C:\Users\<user>\email_archive)
  EMAIL_POP_DIR_WIN   - temp dir for VBS popup scripts (default: C:\Users\<user>\temp)
"""

import email
import hashlib
import imaplib
import os
import re
import ssl
import subprocess
import sys
from datetime import datetime
from email.header import decode_header, make_header
from email.policy import default

IMAP_HOST = "imap.163.com"
IMAP_PORT = 993

# Override via env vars
_DEFAULT_SAVE = os.path.join(os.environ.get("USERPROFILE", "C:/Users/default"), "email_archive")
SAVE_ROOT_WIN = os.environ.get("EMAIL_SAVE_DIR_WIN", _DEFAULT_SAVE)
SAVE_ROOT_WSL = SAVE_ROOT_WIN.replace("C:\\", "/mnt/c/").replace("\\", "/")

_DEFAULT_POP = os.path.join(os.environ.get("USERPROFILE", "C:/Users/default"), "temp")
POP_DIR_WIN = os.environ.get("EMAIL_POP_DIR_WIN", _DEFAULT_POP)
POP_DIR_WSL = POP_DIR_WIN.replace("C:\\", "/mnt/c/").replace("\\", "/")

WSCRIPT = "/mnt/c/Windows/System32/wscript.exe"

# Senders to watch (customize for your needs)
WATCH_SENDERS = [
    "genscript.com", "genscript.com.cn",
    "ashsh.cn",
    "order", "合成", "基因", "引物", "测序",
    "genewiz", "azenta",
]
# Senders to always skip (junk)
SKIP_SENDERS = [
    "member@service.netease.com", "safe@service.netease.com",
    "noreply", "no_reply",
]


def connect():
    ctx = ssl.create_default_context()
    conn = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT, ssl_context=ctx, timeout=30)
    imaplib.Commands['ID'] = ('NONAUTH', 'AUTH', 'SELECTED')
    conn._simple_command('ID', '("name" "Mac Mail" "version" "16.0" "os" "macOS" '
                         '"os-version" "14.0" "vendor" "Apple Inc.")')
    conn.login(os.environ["NETEASE_USER"], os.environ["NETEASE_PASS"])
    return conn


def decode_text(s):
    if s is None:
        return ""
    try:
        return str(make_header(decode_header(s)))
    except Exception:
        return str(s)


def get_attachment_filename(part):
    filename = part.get_filename()
    if filename is None:
        ct = part.get("Content-Type", "")
        m = re.search(r'name\*?=["\']?([^"\';\r\n]+)["\']?', ct)
        if m:
            filename = m.group(1)
    if filename is None:
        return None
    return decode_text(filename)


def parse_email(msg):
    """Parse raw email -> (body_text, [(filename, data, content_type), ...])."""
    body_parts = []
    attachments = []

    if msg.is_multipart():
        for part in msg.walk():
            ct = part.get_content_type()
            cd = str(part.get("Content-Disposition", ""))
            is_attachment = "attachment" in cd

            if is_attachment:
                fname = get_attachment_filename(part)
                payload = part.get_payload(decode=True)
                if payload and fname:
                    attachments.append((fname, payload, ct))
            elif ct == "text/plain":
                payload = part.get_payload(decode=True)
                if payload:
                    charset = part.get_content_charset() or "utf-8"
                    try:
                        body_parts.append(("plain", payload.decode(charset, errors="replace")))
                    except Exception:
                        body_parts.append(("plain", payload.decode("utf-8", errors="replace")))
            elif ct == "text/html":
                payload = part.get_payload(decode=True)
                if payload:
                    charset = part.get_content_charset() or "utf-8"
                    try:
                        body_parts.append(("html", payload.decode(charset, errors="replace")))
                    except Exception:
                        body_parts.append(("html", payload.decode("utf-8", errors="replace")))
    else:
        payload = msg.get_payload(decode=True)
        ct = msg.get_content_type()
        if payload:
            charset = msg.get_content_charset() or "utf-8"
            body_parts.append(("text", payload.decode(charset, errors="replace")))

    body = ""
    for btype, btext in body_parts:
        if btype == "plain":
            body = btext
            break
    if not body:
        for btype, btext in body_parts:
            if btype == "html":
                import html as _html
                body = re.sub(r'<[^>]+>', ' ', btext)
                body = _html.unescape(body)
                body = re.sub(r'\s+', ' ', body).strip()
                break
    if not body and body_parts:
        body = body_parts[0][1]

    return body.strip(), attachments


def sanitize_filename(s, max_len=100):
    s = re.sub(r'[\\/:*?"<>|\r\n\t]', '', s)
    s = s.strip().strip('.')
    if not s:
        s = "unnamed"
    if len(s) > max_len:
        base, ext = os.path.splitext(s)
        s = base[:max_len - len(ext)] + ext
    return s


def extract_order_number(subject, body):
    """Extract order number from subject/body."""
    text = subject + " " + body[:500]
    # Format: C + digits + letters
    m = re.search(r'\b(C\d{3,4}[A-Z]{2,6}\d{0,3})\b', text)
    if m:
        return m.group(1)
    # Format: 80-xxxxxxxx
    m = re.search(r'\b(80-\d{7,10})\b', text)
    if m:
        return m.group(1)
    # Format: 8+ digits (logistics tracking)
    m = re.search(r'(?<!CRM:)(?<!CRM)\b(\d{8,10})\b', subject)
    if m:
        return m.group(1)
    return None


def extract_gene_name(subject, body):
    """Extract gene/product name from subject/body."""
    text = subject + " " + body[:1000]
    patterns = [
        r'([A-Z]{2,5}[\w-]*[A-Z]\d[\w-]*)',
        r'基因名[称]?[：:]\s*(\S{3,20})',
        r'Gene[：:]\s*(\S{3,20})',
    ]
    found = []
    for pat in patterns:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            name = m.group(1).upper()
            if len(name) <= 20 or not re.match(r'^[A-Z]+$', name):
                found.append(name[:20])
    return "-".join(found[:3]) if found else None


def extract_biz_type(subject, body):
    """Extract business type from subject."""
    text = (subject + " " + body[:500]).lower()
    types = {
        "基因合成": ["基因合成", "gene synthesis", "gene order", "codon optimization",
                   "密码子优化"],
        "测序": ["测序", "sequencing", "测序结果"],
        "引物": ["引物", "primer", "oligo"],
        "物流": ["物流", "送达", "温度记录", "delivery", "shipment"],
        "报价": ["报价", "quote", "quotation"],
    }
    for label, keywords in types.items():
        for kw in keywords:
            if kw in text:
                return label
    return None


def build_dirname(from_addr, subject, body):
    """Build directory name."""
    order = extract_order_number(subject, body)
    gene = extract_gene_name(subject, body)
    biz = extract_biz_type(subject, body)

    if order and gene and biz:
        return sanitize_filename(f"{order}_{gene}_{biz}", max_len=80)
    elif order and gene:
        return sanitize_filename(f"{order}_{gene}", max_len=80)
    elif order and biz:
        return sanitize_filename(f"{order}_{biz}", max_len=80)
    elif order:
        return sanitize_filename(f"{order}", max_len=80)
    else:
        sender = short_sender(from_addr)
        subj = sanitize_filename(subject, max_len=50)
        return sanitize_filename(f"{sender}_{subj}", max_len=80)


def short_sender(from_addr):
    """Extract readable short sender name."""
    m = re.match(r'"?([^"<@]+)"?\s*<', from_addr)
    if m:
        name = m.group(1).strip()
        if name and not re.match(r'^[\d.@]+$', name):
            return sanitize_filename(name, max_len=20)
    m = re.search(r'([\w.-]+)@', from_addr)
    if m:
        return m.group(1)
    return sanitize_filename(from_addr, max_len=20)


def find_existing_dir(dirname):
    """Check if a directory matching this order already exists (by order number prefix)."""
    if not os.path.isdir(SAVE_ROOT_WSL):
        return None
    exact = os.path.join(SAVE_ROOT_WSL, dirname)
    if os.path.isdir(exact):
        return exact
    prefix = dirname.split('_')[0]
    if re.match(r'^(C[A-Za-z0-9]{4,}|80-\d+|\d{8,})$', prefix):
        for entry in os.listdir(SAVE_ROOT_WSL):
            if entry.startswith(prefix + '_') or entry == prefix:
                return os.path.join(SAVE_ROOT_WSL, entry)
    return None


def build_readme(from_addr, subject, body, attachments, date_str):
    """Build README.md content for a single email."""
    lines = []
    m = re.match(r'"?([^"<]+)"?\s*<(.+?)>', from_addr)
    sender_name = m.group(1).strip() if m else from_addr

    lines.append(f"# {subject}")
    lines.append("")
    lines.append(f"- **时间**: {date_str}")
    lines.append(f"- **发件人**: {sender_name}")
    if attachments:
        att_names = [a[0] for a in attachments]
        lines.append(f"- **附件**: {', '.join(att_names)}")
    else:
        lines.append(f"- **附件**: 无")
    lines.append("")
    lines.append("---")
    lines.append("")
    lines.append("## 邮件内容")
    lines.append("")
    clean_body = body[:3000]
    lines.append(clean_body)
    if len(body) > 3000:
        lines.append("")
        lines.append("... (内容较长，完整内容见邮件原文)")
    return "\n".join(lines)


def build_shuoming(from_addr, subject, attachments, date_str, savedir_win):
    """Build summary.txt content."""
    lines = []
    lines.append(f"邮件: {subject}")
    lines.append(f"发件人: {from_addr}")
    lines.append(f"时间: {date_str}")
    lines.append(f"保存位置: {savedir_win}")
    if attachments:
        lines.append(f"附件 ({len(attachments)}个):")
        for fname, _, _ in attachments:
            lines.append(f"  - {fname}")
    else:
        lines.append("附件: 无")
    return "\n".join(lines)


def send_popup(title, message_lines):
    """Send Windows MsgBox popup."""
    def esc(s):
        return str(s).replace('"', '""')
    parts = [f'"{esc(line)}"' for line in message_lines]
    msg_expr = " & vbCrLf & ".join(parts)
    vbs = f'MsgBox {msg_expr}, 64, "{esc(title)}"\r\n'
    os.makedirs(POP_DIR_WSL, exist_ok=True)
    ts = datetime.now().strftime('%Y%m%d_%H%M%S')
    filename = f"_email_{ts}.vbs"
    path = os.path.join(POP_DIR_WSL, filename)
    with open(path, 'w', encoding='gbk') as f:
        f.write(vbs)
    subprocess.Popen(
        [WSCRIPT, POP_DIR_WIN + "\\" + filename],
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
    )


def should_watch(from_addr, subject):
    combined = (from_addr + " " + subject).lower()
    for kw in WATCH_SENDERS:
        if kw.lower() in combined:
            return True
    return False


def should_skip(from_addr):
    from_lower = from_addr.lower()
    for kw in SKIP_SENDERS:
        if kw.lower() in from_lower:
            return True
    return False


def main():
    dry_run = "--dry" in sys.argv
    catchup = None
    for i, arg in enumerate(sys.argv):
        if arg == "--catchup" and i + 1 < len(sys.argv):
            catchup = int(sys.argv[i + 1])
            break

    if "NETEASE_USER" not in os.environ or "NETEASE_PASS" not in os.environ:
        print("Set NETEASE_USER and NETEASE_PASS env vars")
        sys.exit(1)

    conn = connect()
    try:
        conn.select("INBOX")

        if catchup:
            status, data = conn.search(None, "ALL")
            all_ids = data[0].split()
            ids = all_ids[-catchup:]
            is_unseen = set()
            status, unseen_data = conn.search(None, "UNSEEN")
            for mid in unseen_data[0].split():
                is_unseen.add(mid)
            print(f"Scanning last {len(ids)} of {len(all_ids)} total emails...")
        else:
            status, data = conn.search(None, "UNSEEN")
            ids = data[0].split()
            is_unseen = set(ids)
            if not ids:
                print("No unread mail.")
                return
            print(f"Found {len(ids)} unread, filtering...")

        relevant = []
        for mid in ids:
            status, data = conn.fetch(mid, "(BODY.PEEK[HEADER.FIELDS (FROM SUBJECT DATE)])")
            raw = data[0][1].decode("utf-8", errors="replace")
            hdr = email.message_from_string(raw, policy=default)
            from_addr = decode_text(hdr.get('From', ''))
            subject = decode_text(hdr.get('Subject', '(no subject)'))
            date_str = decode_text(hdr.get('Date', ''))

            if should_skip(from_addr):
                conn.store(mid, '+FLAGS', '\\Seen')
                continue
            if should_watch(from_addr, subject):
                status, data = conn.fetch(mid, "(RFC822)")
                msg = email.message_from_bytes(data[0][1], policy=default)
                body, attachments = parse_email(msg)
                relevant.append({
                    'mid': mid,
                    'from': from_addr,
                    'subject': subject,
                    'date': date_str,
                    'body': body,
                    'attachments': attachments,
                    'is_unseen': mid in is_unseen,
                })

        if not relevant:
            print("No progress-related emails found.")
            return

        print(f"Found {len(relevant)} progress-related\n")

        if dry_run:
            for i, r in enumerate(relevant):
                dirname = build_dirname(r['from'], r['subject'], r['body'])
                existing = find_existing_dir(dirname)
                tag = "[NEW]" if not existing else "[EXISTING]"
                unread_tag = "[UNREAD]" if r['is_unseen'] else "[READ]"
                print(f"[{i+1}] {unread_tag} {tag} {r['from'][:40]} | {r['subject'][:60]}")
                print(f"    Dir: {existing or dirname}")
                print(f"    Attachments: {len(r['attachments'])}")
                for fname, data, ct in r['attachments']:
                    print(f"      - {fname} ({len(data)} bytes)")
                print(f"    Body: {r['body'][:200]}...")
                print()
            print(f"[Dry run — {len(relevant)} emails would be processed]")
            return

        popup_lines = []

        for i, r in enumerate(relevant):
            dirname = build_dirname(r['from'], r['subject'], r['body'])
            existing = find_existing_dir(dirname)
            if existing:
                savedir_wsl = existing
                savedir_win = SAVE_ROOT_WIN + "\\" + os.path.basename(existing)
                is_new_dir = False
            else:
                savedir_wsl = os.path.join(SAVE_ROOT_WSL, dirname)
                savedir_win = f"{SAVE_ROOT_WIN}\\{dirname}"
                os.makedirs(savedir_wsl, exist_ok=True)
                is_new_dir = True

            if not r['is_unseen']:
                readme_path = os.path.join(savedir_wsl, "README.md")
                if os.path.exists(readme_path):
                    with open(readme_path, 'r', encoding='utf-8', errors='replace') as rf:
                        if r['subject'][:60] in rf.read():
                            print(f"[{i+1}] SKIP (already saved): {savedir_win}")
                            continue

            email_hash = hashlib.md5(f"{r['date']}{r['subject']}".encode()).hexdigest()[:8]

            readme_path = os.path.join(savedir_wsl, "README.md")
            if is_new_dir:
                readme_content = build_readme(r['from'], r['subject'], r['body'],
                                              r['attachments'], r['date'])
                with open(readme_path, 'w', encoding='utf-8') as f:
                    f.write(readme_content)
            else:
                existing_readme = ""
                if os.path.exists(readme_path):
                    with open(readme_path, 'r', encoding='utf-8', errors='replace') as f:
                        existing_readme = f.read()
                if email_hash not in existing_readme and r['subject'] not in existing_readme:
                    append_block = f"\n\n---\n\n### {r['subject']}\n"
                    append_block += f"- **日期**: {r['date']}\n"
                    append_block += f"- **发件人**: {r['from']}\n"
                    if r['attachments']:
                        att_names = [a[0] for a in r['attachments']]
                        append_block += f"- **附件**: {', '.join(att_names)}\n"
                    else:
                        append_block += f"- **附件**: 无\n"
                    with open(readme_path, 'a', encoding='utf-8') as f:
                        f.write(append_block)

            shuoming = build_shuoming(r['from'], r['subject'], r['attachments'],
                                       r['date'], savedir_win)
            with open(os.path.join(savedir_wsl, "summary.txt"), 'w', encoding='utf-8') as f:
                f.write(shuoming)

            saved_files = []
            for fname, fdata, fct in r['attachments']:
                clean_name = sanitize_filename(fname)
                att_path = os.path.join(savedir_wsl, clean_name)
                if os.path.exists(att_path) and os.path.getsize(att_path) == len(fdata):
                    saved_files.append(f"{clean_name} (exists)")
                    continue
                with open(att_path, 'wb') as f:
                    f.write(fdata)
                saved_files.append(clean_name)

            print(f"[{i+1}] -> {savedir_win}")
            print(f"      attachments: {saved_files}")

            if len(popup_lines) < 20:
                if popup_lines:
                    popup_lines.append("-" * 30)
                popup_lines.append(f"Sender: {short_sender(r['from'])}")
                popup_lines.append(f"Subject: {r['subject']}")
                if r['attachments']:
                    popup_lines.append("Attachments:")
                    for fname, _, _ in r['attachments']:
                        popup_lines.append(f"  + {sanitize_filename(fname)}")
                else:
                    popup_lines.append("Attachments: none")
                popup_lines.append(f"Saved to:")
                popup_lines.append(f"  {savedir_win}")

            if r['is_unseen']:
                conn.store(r['mid'], '+FLAGS', '\\Seen')

        if popup_lines:
            send_popup(f"Email - {len(relevant)} new", popup_lines[:25])
            print(f"\nDone: {len(relevant)} email(s) saved + notified.")
        else:
            print("\nNo new emails to save.")

    finally:
        conn.logout()


if __name__ == "__main__":
    main()
