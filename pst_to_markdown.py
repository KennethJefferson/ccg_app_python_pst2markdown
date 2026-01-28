#!/usr/bin/env python3
"""
PST to Markdown Converter

Extracts emails from Outlook PST files and converts them to Markdown format.
Uses Outlook COM interface (requires Outlook installed on Windows).
"""

import argparse
import os
import re
import sys
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path
from typing import Optional

try:
    import win32com.client
    from win32com.client import CDispatch
except ImportError:
    print("Error: pywin32 is required. Install with: pip install pywin32")
    sys.exit(1)

try:
    from markdownify import markdownify as md
except ImportError:
    print("Error: markdownify is required. Install with: pip install markdownify")
    sys.exit(1)

try:
    from tqdm import tqdm
except ImportError:
    print("Error: tqdm is required. Install with: pip install tqdm")
    sys.exit(1)


def sanitize_filename(name: str, max_length: int = 100) -> str:
    """Remove/replace characters invalid in Windows filenames."""
    invalid_chars = r'[<>:"/\\|?*\x00-\x1f]'
    sanitized = re.sub(invalid_chars, '_', name)
    sanitized = re.sub(r'_+', '_', sanitized)
    sanitized = sanitized.strip('. _')
    if len(sanitized) > max_length:
        sanitized = sanitized[:max_length].rstrip('. _')
    return sanitized or "unnamed"


def extract_sender_name(sender: str) -> str:
    """
    Extract sender name from formats like:
    - 'The Neuron <theneuron@newsletter.theneurondaily.com>' -> 'The Neuron'
    - 'john.doe@example.com' -> 'john.doe'
    - 'John Doe' -> 'John Doe'
    """
    if not sender:
        return "unknown"

    match = re.match(r'^([^<]+)\s*<[^>]+>$', sender.strip())
    if match:
        return match.group(1).strip()

    if '@' in sender and '<' not in sender:
        return sender.split('@')[0]

    return sender.strip()


def generate_unique_filename(base_path: Path, filename: str, extension: str = "") -> Path:
    """Generate a unique filename by adding suffix if file exists."""
    full_path = base_path / f"{filename}{extension}"
    if not full_path.exists():
        return full_path

    counter = 1
    while True:
        new_path = base_path / f"{filename}_{counter}{extension}"
        if not new_path.exists():
            return new_path
        counter += 1


def format_email_date(received_time) -> tuple[str, str]:
    """
    Format email received time for filename and display.
    Returns (filename_date, display_date).
    """
    if received_time is None:
        now = datetime.now()
        return now.strftime("%Y-%m-%d"), now.strftime("%Y-%m-%d %H:%M:%S")

    try:
        if hasattr(received_time, 'strftime'):
            return received_time.strftime("%Y-%m-%d"), received_time.strftime("%Y-%m-%d %H:%M:%S")
        dt = datetime.fromisoformat(str(received_time))
        return dt.strftime("%Y-%m-%d"), dt.strftime("%Y-%m-%d %H:%M:%S")
    except Exception:
        return "unknown-date", "Unknown"


def html_to_markdown(html_body: str, plain_body: str) -> str:
    """Convert HTML body to Markdown, falling back to plain text if needed."""
    if html_body:
        try:
            return md(html_body, heading_style="ATX", bullets="-")
        except Exception:
            pass
    return plain_body or ""


def format_recipients(recipients) -> str:
    """Format recipients collection to string."""
    if not recipients:
        return ""

    try:
        recipient_list = []
        for i in range(1, recipients.Count + 1):
            recip = recipients.Item(i)
            name = recip.Name or recip.Address or "Unknown"
            address = recip.Address or ""
            if address and address != name:
                recipient_list.append(f"{name} <{address}>")
            else:
                recipient_list.append(name)
        return ", ".join(recipient_list)
    except Exception:
        return ""


def create_email_markdown(mail_item) -> tuple[str, str, list[tuple[str, bytes]]]:
    """
    Convert an Outlook mail item to Markdown content.
    Returns (filename_base, markdown_content, attachments_list).
    Attachments list contains tuples of (filename, data).
    """
    try:
        subject = mail_item.Subject or "No Subject"
        sender = mail_item.SenderName or mail_item.SenderEmailAddress or "Unknown"
        received_time = mail_item.ReceivedTime

        to_recipients = format_recipients(mail_item.Recipients)
        cc_recipients = ""
        try:
            cc_recipients = mail_item.CC or ""
        except Exception:
            pass

        html_body = ""
        plain_body = ""
        try:
            html_body = mail_item.HTMLBody or ""
        except Exception:
            pass
        try:
            plain_body = mail_item.Body or ""
        except Exception:
            pass

        date_filename, date_display = format_email_date(received_time)
        sender_name = extract_sender_name(sender)

        filename_base = sanitize_filename(f"{date_filename}_{sender_name}_{subject}")

        body_md = html_to_markdown(html_body, plain_body)

        attachments: list[tuple[str, bytes]] = []
        attachment_links: list[str] = []

        try:
            for i in range(1, mail_item.Attachments.Count + 1):
                att = mail_item.Attachments.Item(i)
                att_filename = sanitize_filename(att.FileName or f"attachment_{i}")

                temp_path = Path(os.environ.get('TEMP', '.')) / f"pst_extract_{att_filename}"
                try:
                    att.SaveAsFile(str(temp_path))
                    with open(temp_path, 'rb') as f:
                        att_data = f.read()
                    attachments.append((att_filename, att_data))
                    attachment_links.append(f"- [{att_filename}]({att_filename})")
                finally:
                    if temp_path.exists():
                        temp_path.unlink()
        except Exception:
            pass

        md_content = f"""# {subject}

| Field | Value |
|-------|-------|
| **From** | {sender} |
| **To** | {to_recipients} |
| **CC** | {cc_recipients} |
| **Date** | {date_display} |

"""

        if attachment_links:
            md_content += "## Attachments\n\n"
            md_content += "\n".join(attachment_links)
            md_content += "\n\n"

        md_content += "## Content\n\n"
        md_content += body_md

        return filename_base, md_content, attachments

    except Exception as e:
        return "error_email", f"# Error Processing Email\n\nError: {e}", []


def count_emails_in_folder(folder) -> int:
    """Recursively count all emails in a folder and its subfolders."""
    count = 0
    try:
        count = folder.Items.Count
    except Exception:
        pass

    try:
        for i in range(1, folder.Folders.Count + 1):
            subfolder = folder.Folders.Item(i)
            count += count_emails_in_folder(subfolder)
    except Exception:
        pass

    return count


def process_folder(folder, output_dir: Path, pbar: tqdm) -> int:
    """
    Recursively process all emails in a folder and its subfolders.
    Returns count of processed emails.
    """
    processed = 0

    try:
        items = folder.Items
        for i in range(1, items.Count + 1):
            try:
                item = items.Item(i)
                if item.Class == 43:  # olMail
                    filename_base, md_content, attachments = create_email_markdown(item)

                    if attachments:
                        email_folder = generate_unique_filename(output_dir, filename_base)
                        email_folder.mkdir(parents=True, exist_ok=True)

                        md_path = email_folder / f"{filename_base}.md"
                        md_path.write_text(md_content, encoding='utf-8')

                        for att_name, att_data in attachments:
                            att_path = email_folder / att_name
                            att_path.write_bytes(att_data)
                    else:
                        md_path = generate_unique_filename(output_dir, filename_base, ".md")
                        md_path.write_text(md_content, encoding='utf-8')

                    processed += 1
                    pbar.update(1)
            except Exception as e:
                pbar.write(f"Warning: Failed to process item: {e}")
                pbar.update(1)
    except Exception as e:
        pbar.write(f"Warning: Error accessing folder items: {e}")

    try:
        for i in range(1, folder.Folders.Count + 1):
            subfolder = folder.Folders.Item(i)
            processed += process_folder(subfolder, output_dir, pbar)
    except Exception:
        pass

    return processed


def process_pst_file(pst_path: Path, output_dir: Optional[Path], worker_id: int) -> tuple[str, int, Optional[str]]:
    """
    Process a single PST file.
    Returns (pst_name, processed_count, error_message).
    """
    pst_name = pst_path.name

    if not pst_path.exists():
        return pst_name, 0, f"File not found: {pst_path}"

    if output_dir is None:
        output_dir = pst_path.parent

    output_dir.mkdir(parents=True, exist_ok=True)

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")

        namespace.AddStore(str(pst_path))

        pst_folder = None
        for i in range(1, namespace.Folders.Count + 1):
            folder = namespace.Folders.Item(i)
            try:
                store_path = folder.Store.FilePath
                if Path(store_path).resolve() == pst_path.resolve():
                    pst_folder = folder
                    break
            except Exception:
                continue

        if pst_folder is None:
            for i in range(namespace.Folders.Count, 0, -1):
                folder = namespace.Folders.Item(i)
                pst_folder = folder
                break

        if pst_folder is None:
            return pst_name, 0, "Could not locate PST folder in Outlook"

        total_emails = count_emails_in_folder(pst_folder)

        with tqdm(total=total_emails, desc=f"[Worker {worker_id}] {pst_name}",
                  unit="email", position=worker_id, leave=True) as pbar:
            processed = process_folder(pst_folder, output_dir, pbar)

        try:
            namespace.RemoveStore(pst_folder)
        except Exception:
            pass

        return pst_name, processed, None

    except Exception as e:
        return pst_name, 0, str(e)


def main():
    parser = argparse.ArgumentParser(
        description="Extract emails from Outlook PST files to Markdown format",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s -i emails.pst
  %(prog)s -i file1.pst file2.pst -o ./output -w 2
        """
    )

    parser.add_argument(
        '-i', '--input',
        nargs='+',
        required=True,
        type=Path,
        metavar='PST',
        help='Input PST file(s)'
    )

    parser.add_argument(
        '-o', '--output',
        type=Path,
        metavar='DIR',
        help='Output directory (default: same directory as input PST)'
    )

    parser.add_argument(
        '-w', '--workers',
        type=int,
        default=1,
        metavar='N',
        help='Number of parallel workers (default: 1)'
    )

    args = parser.parse_args()

    pst_files = [p.resolve() for p in args.input]
    for pst in pst_files:
        if not pst.exists():
            print(f"Error: File not found: {pst}")
            sys.exit(1)
        if pst.suffix.lower() != '.pst':
            print(f"Warning: {pst} may not be a PST file")

    output_dir = args.output.resolve() if args.output else None

    num_workers = min(args.workers, len(pst_files))

    print(f"Processing {len(pst_files)} PST file(s) with {num_workers} worker(s)")
    print()

    results = []

    if num_workers == 1:
        for pst_path in pst_files:
            result = process_pst_file(pst_path, output_dir, 0)
            results.append(result)
    else:
        with ThreadPoolExecutor(max_workers=num_workers) as executor:
            futures = {
                executor.submit(process_pst_file, pst_path, output_dir, idx): pst_path
                for idx, pst_path in enumerate(pst_files)
            }

            for future in as_completed(futures):
                results.append(future.result())

    print()
    print("=" * 50)
    print("Summary")
    print("=" * 50)

    total_processed = 0
    for pst_name, count, error in results:
        if error:
            print(f"  {pst_name}: FAILED - {error}")
        else:
            print(f"  {pst_name}: {count} emails processed")
            total_processed += count

    print(f"\nTotal: {total_processed} emails converted to Markdown")


if __name__ == "__main__":
    main()
