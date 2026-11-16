import os
import sys
import re
import json
import argparse
from datetime import datetime, date
import win32com.client
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

DEFAULT_EXCEL_PATH = r"C:\Users\Yehya\Desktop\Bonga\MailS\Shows_Tracker.xlsx"
DEFAULT_SHEET_NAME = "Sheet1"
DEFAULT_OUTLOOK_FOLDER = "Inbox"

def clean_whitespace(text: str) -> str:
    """Normalize non-breaking spaces, en-spaces, and line endings."""
    if not text:
        return ""
    text = re.sub(r"[\u00a0\u2000-\u200b\u202f\u205f\u3000]", " ", text)
    return text.replace("\r\n", "\n").replace("\r", "\n")

def extract_urls_and_comment(line: str):
    """
    Extract canonical full URLs (prioritizing bracketed <URL> from Outlook)
    and any Locating comment (which begins with '*').
    """
    bracket_urls = re.findall(r'<((?:https?://|www\.)[^\s>]+)>', line, re.IGNORECASE)
    bare_urls = re.findall(r'(?:https?://|www\.)[^\s<>]+', line, re.IGNORECASE)

    all_urls = []
    for u in bracket_urls:
        u_clean = u.strip()
        if not u_clean.lower().startswith('http'):
            u_clean = 'https://' + u_clean
        if u_clean not in all_urls:
            all_urls.append(u_clean)

    for u in bare_urls:
        u_clean = u.strip()
        if not u_clean.lower().startswith('http'):
            u_clean = 'https://' + u_clean
        # Only add if not already part of an existing bracket URL
        if not any(u_clean in b for b in all_urls):
            all_urls.append(u_clean)

    # Extract comment: Locating comments always start with '*' (e.g. '**** Not Available On WebSite')
    comment = ""
    if '*' in line:
        star_idx = line.find('*')
        raw_com = line[star_idx:]
        comment = re.sub(r'\*+', ' ', raw_com)
        comment = re.sub(r'<https?://[^>]+>', ' ', comment)
        comment = re.sub(r'https?://\S+', ' ', comment)
        comment = re.sub(r'\s+', ' ', comment).strip(' -:\t\n')

    return all_urls, comment

def extract_show_name(subject: str) -> str:
    """
    Strip Fwd:/FW:/Re: and dynamic sender prefixes (e.g. 'PANDA - ', 'Semary - ', 'S automated - ')
    while keeping 'Xshow - ', 'XZShow - ', 'ZSHOW - ', 'Show - ', etc.
    """
    s = clean_whitespace(subject).strip()
    s = re.sub(r"^(?:(?:fwd|fw|re)\s*:\s*)+", "", s, flags=re.IGNORECASE).strip()
    s = re.sub(
        r"^(?!x?z?show\b)[A-Za-z0-9_ ]+\s*-\s*",
        "",
        s,
        count=1,
        flags=re.IGNORECASE,
    ).strip()
    return s

def parse_email_body(body: str) -> dict:
    """
    Extract exhibitor list URLs, inline/follow-up comments, and bottom list range.
    Handles hyperlinked 'List <url> : range' without dropping the range.
    """
    if not body:
        return {"links": [], "list_range": "", "range_comment": ""}

    text = clean_whitespace(body)
    lines = [line.strip() for line in text.split("\n")]

    links = []  # list of tuples: (url, comment)
    in_list_section = False

    stop_section_re = re.compile(
        r"^(?:housing|venue|show\s*offer|we\s*have|hotel|location)\b\s*:?",
        re.IGNORECASE,
    )

    for line in lines:
        if not line:
            continue

        # Stop if we hit non-list sections like Housing or Venue
        if in_list_section and stop_section_re.match(line):
            in_list_section = False
            continue

        # Check if line is a bottom range line (even if it has an embedded URL like 'List <url> : 100 - 150')
        line_no_urls = re.sub(r"<https?://[^>]+>", " ", line)
        line_no_urls = re.sub(r"https?://\S+", " ", line_no_urls)
        is_range_line = bool(re.search(r"l[a-z]{0,6}\s*:\s*\d+\s*(?:[-–—>|/]|to)\s*\d+", line_no_urls, re.IGNORECASE))

        if is_range_line:
            # Range line, handled in bottom-up pass
            continue

        is_list_line = bool(re.match(r"^l[a-z]{0,6}\s*[:<]", line, re.IGNORECASE))
        if is_list_line:
            in_list_section = True

        if in_list_section:
            found_urls, inline_com = extract_urls_and_comment(line)
            if found_urls:
                for idx, u in enumerate(found_urls):
                    # Attach the inline comment to the first URL on this line
                    com = inline_com if idx == 0 else ""
                    links.append((u, com))
            elif links and (line.startswith("*") or line.startswith("-")):
                # Follow-up comment line below the link!
                extra_comment = re.sub(r"[\*\-]+", " ", line)
                extra_comment = re.sub(r"\s+", " ", extra_comment).strip(" -:\t\n")
                if extra_comment:
                    prev_url, prev_com = links[-1]
                    combined = f"{prev_com} | {extra_comment}" if prev_com else extra_comment
                    links[-1] = (prev_url, combined)

    # Bottom range extraction (scans upwards from bottom of email)
    list_range = ""
    range_comment = ""
    for line in reversed(lines):
        line_no_urls = re.sub(r"<https?://[^>]+>", " ", line)
        line_no_urls = re.sub(r"https?://\S+", " ", line_no_urls)
        m_range = re.search(r"l[a-z]{0,6}\s*:\s*(\d+\s*(?:[-–—>|/]|to)\s*\d+)(.*)", line_no_urls, re.IGNORECASE)
        if m_range:
            nums = re.sub(r"\s+", " ", m_range.group(1)).strip()
            nums = re.sub(r"\s*(?:[-–—>|/]|to)\s*", " - ", nums)
            list_range = f"List: {nums}"
            raw_com = m_range.group(2)
            if '*' in raw_com:
                star_idx = raw_com.find('*')
                raw_com = raw_com[star_idx:]
            range_comment = re.sub(r"[\*\-]+", " ", raw_com)
            range_comment = re.sub(r"\s+", " ", range_comment).strip(" -:\t\n")
            break

    return {
        "links": links,
        "list_range": list_range,
        "range_comment": range_comment,
    }

def get_or_create_workbook(path: str, sheet_name: str):
    """Load existing tracker workbook or initialize a new one with correct styling."""
    if os.path.exists(path):
        wb = openpyxl.load_workbook(path)
        ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.active
        ws.sheet_view.zoomScale = 85
        return wb, ws

    os.makedirs(os.path.dirname(path), exist_ok=True)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_name
    ws.sheet_view.zoomScale = 85

    headers = ["Show Name", "Exhbitor list", "Num of Exhibitor", "Comment", "Agent Name"]
    ws.append(headers)

    header_fill = PatternFill(start_color="D96B27", end_color="D96B27", fill_type="solid")
    header_font = Font(name="Arial", size=10, bold=True, color="1F4E78")
    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for col_idx in range(1, 6):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border

    ws.row_dimensions[1].height = 15
    ws.column_dimensions["A"].width = 88
    ws.column_dimensions["B"].width = 62
    ws.column_dimensions["C"].width = 22
    ws.column_dimensions["D"].width = 28
    ws.column_dimensions["E"].width = 18

    return wb, ws

def get_existing_shows(ws) -> set:
    """Collect already-logged show names to prevent duplicate entries."""
    existing = set()
    for r in range(2, ws.max_row + 1):
        val = str(ws.cell(row=r, column=1).value or "").strip()
        if val:
            existing.add(val.lower())
    return existing

def ensure_date_banner(ws, target_dt: datetime):
    """Insert a merged grey/orange date banner row if target date isn't already active."""
    target_str = target_dt.strftime("%A, %B %d, %Y").replace(" 0", " ")

    for r in range(ws.max_row, 1, -1):
        val = str(ws.cell(row=r, column=1).value or "").strip()
        if val == target_str:
            return
        if val and any(
            k in val.lower() for k in ["2025", "2026", "2027", "2028", "xshow", "zshow", "usa"]
        ):
            break

    banner_row = ws.max_row + 1
    ws.merge_cells(start_row=banner_row, start_column=1, end_row=banner_row, end_column=5)

    banner_cell = ws.cell(row=banner_row, column=1, value=target_str)
    banner_cell.fill = PatternFill(start_color="808080", end_color="808080", fill_type="solid")
    banner_cell.font = Font(name="Arial", size=10, bold=True, color="F4B084")
    banner_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[banner_row].height = 15

    thin = Side(style="thin", color="000000")
    for c in range(1, 6):
        ws.cell(row=banner_row, column=c).border = Border(left=thin, right=thin, top=thin, bottom=thin)

def append_show_to_sheet(ws, show_name: str, parsed: dict, agent_name: str = ""):
    """Write one or more rows for a show."""
    links = parsed["links"] or [("", "")]
    list_range = parsed["list_range"]
    range_comment = parsed["range_comment"]

    thin = Side(style="thin", color="000000")
    cell_border = Border(left=thin, right=thin, top=thin, bottom=thin)
    bold_font = Font(name="Arial", size=10, bold=True, color="000000")
    link_font = Font(name="Arial", size=10, underline="single", color="0563C1")

    for idx, (url, inline_comment) in enumerate(links):
        row_num = ws.max_row + 1
        is_first = idx == 0

        # Each link row gets its own corresponding comment
        comments = [c for c in ([inline_comment, range_comment] if is_first else [inline_comment]) if c]
        final_comment = " | ".join(comments)

        c_name = ws.cell(row=row_num, column=1, value=show_name if is_first else "")
        c_url = ws.cell(row=row_num, column=2, value=url)
        c_range = ws.cell(row=row_num, column=3, value=list_range if is_first else "")
        c_comment = ws.cell(row=row_num, column=4, value=final_comment)
        c_agent = ws.cell(row=row_num, column=5, value=agent_name if is_first else "")

        ws.row_dimensions[row_num].height = 15

        c_name.font = bold_font
        c_range.font = bold_font
        c_comment.font = bold_font
        c_agent.font = bold_font
        c_agent.alignment = Alignment(horizontal="center", vertical="center")

        if url:
            c_url.hyperlink = url
            c_url.font = link_font

        for cell in (c_name, c_url, c_range, c_comment, c_agent):
            cell.border = cell_border

def get_outlook_folder(namespace, folder_path: str):
    """Resolve 'Inbox' or nested paths like 'Inbox/Shows'."""
    parts = [p.strip() for p in folder_path.split("/") if p.strip()]
    folder = namespace.GetDefaultFolder(6)  # 6 = olFolderInbox
    start_idx = 1 if parts and parts[0].lower() == "inbox" else 0
    for part in parts[start_idx:]:
        folder = folder.Folders[part]
    return folder

def main():
    parser = argparse.ArgumentParser(description="Zorg-Nexus Outlook Shows Scanner")
    parser.add_argument("--date", type=str, default=None, help="Target date in YYYY-MM-DD format (default: today)")
    parser.add_argument("--excel-path", type=str, default=DEFAULT_EXCEL_PATH, help="Path to Shows_Tracker.xlsx")
    parser.add_argument("--folder", type=str, default=DEFAULT_OUTLOOK_FOLDER, help="Outlook folder to scan")
    parser.add_argument("--only-unread", action="store_true", help="Only scan unread emails")
    parser.add_argument("--mark-read", action="store_true", help="Mark scraped emails as read")
    parser.add_argument("--dry-run", action="store_true", help="Do not save changes to Excel")

    args = parser.parse_args()

    if args.date:
        try:
            target_date = datetime.strptime(args.date.strip(), "%Y-%m-%d").date()
        except ValueError:
            sys.stderr.write(f"Invalid date format: {args.date}. Expected YYYY-MM-DD.\n")
            sys.exit(1)
    else:
        target_date = datetime.now().date()

    target_dt = datetime.combine(target_date, datetime.min.time())

    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        folder = get_outlook_folder(outlook, args.folder)
    except Exception as e:
        sys.stderr.write(f"Outlook MAPI Connection Error: {str(e)}\n")
        print(json.dumps({"success": False, "error": f"Outlook Connection Error: {str(e)}"}))
        sys.exit(1)

    messages = folder.Items
    messages.Sort("[ReceivedTime]", True)

    wb, ws = get_or_create_workbook(args.excel_path, DEFAULT_SHEET_NAME)
    existing_shows = get_existing_shows(ws)

    raw_shows = []

    for msg in messages:
        if getattr(msg, "Class", None) != 43:
            continue

        try:
            msg_date = msg.ReceivedTime.date()
        except Exception:
            continue

        # Skip future emails if any
        if msg_date > target_date:
            continue
        # Stop scanning once we reach dates earlier than the target
        if msg_date < target_date:
            break

        if args.only_unread and not msg.UnRead:
            continue

        subject = msg.Subject or ""
        body = msg.Body or ""

        # Must contain list indicator
        if not re.search(r"(?:^|\n)\s*l[a-z]{0,6}\s*[:<]", clean_whitespace(body), re.IGNORECASE):
            continue

        show_name = extract_show_name(subject)
        if not show_name:
            continue

        parsed = parse_email_body(body)
        is_duplicate = show_name.lower() in existing_shows

        # Build clean combined comment string
        link_comments = [c for _, c in parsed["links"] if c]
        all_comments = []
        for c in link_comments:
            if c and c not in all_comments:
                all_comments.append(c)
        if parsed["range_comment"] and parsed["range_comment"] not in all_comments:
            all_comments.append(parsed["range_comment"])

        raw_shows.append({
            "showName": show_name,
            "parsed": parsed,
            "links": [u for u, _ in parsed["links"] if u],
            "linkComments": link_comments,
            "listRange": parsed["list_range"],
            "rangeComment": parsed["range_comment"],
            "comment": " | ".join(all_comments),
            "receivedTime": str(msg.ReceivedTime),
            "isDuplicate": is_duplicate,
            "msg": msg
        })

    # 1. Reverse collected shows so they are chronological:
    # First email sent in the day -> First item; Last email sent -> Last item
    raw_shows.reverse()

    # 2. Partition: "New" shows at the top, followed by "XShows"
    def is_x_show(name: str) -> bool:
        return bool(re.search(r'\bx\s*[-_]?\s*z?\s*show\b', name, re.IGNORECASE))

    new_shows = [s for s in raw_shows if not is_x_show(s["showName"])]
    x_shows = [s for s in raw_shows if is_x_show(s["showName"])]
    ordered_shows = new_shows + x_shows

    banner_ensured = False

    # 3. Write to Excel in this exact ordered sequence
    for item in ordered_shows:
        if not item["isDuplicate"] and not args.dry_run:
            if not banner_ensured:
                ensure_date_banner(ws, target_dt)
                banner_ensured = True
            append_show_to_sheet(ws, item["showName"], item["parsed"])
            existing_shows.add(item["showName"].lower())

            if args.mark_read and item["msg"].UnRead:
                try:
                    item["msg"].UnRead = False
                    item["msg"].Save()
                except Exception:
                    pass

    if not args.dry_run and len(ordered_shows) > 0 and banner_ensured:
        try:
            wb.save(args.excel_path)
        except Exception as e:
            sys.stderr.write(f"Error saving Excel file: {str(e)}\n")

    # 4. Prepare JSON output (remove msg COM object before JSON serialization)
    final_shows = []
    for s in ordered_shows:
        final_shows.append({
            "showName": s["showName"],
            "links": s["links"],
            "linkComments": s["linkComments"],
            "listRange": s["listRange"],
            "rangeComment": s["rangeComment"],
            "comment": s["comment"],
            "receivedTime": s["receivedTime"],
            "isDuplicate": s["isDuplicate"],
            "isXShow": is_x_show(s["showName"])
        })

    output_payload = {
        "success": True,
        "date": str(target_date),
        "scrapedCount": len(final_shows),
        "newCount": len([s for s in final_shows if not s["isDuplicate"]]),
        "shows": final_shows
    }

    print(json.dumps(output_payload))

if __name__ == "__main__":
    main()
