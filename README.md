# PST to Markdown Converter

A Python CLI tool that extracts emails from Outlook PST files and converts them to Markdown format.

## Features

- Extract emails from one or more PST files
- Convert HTML email bodies to Markdown
- Preserve email metadata (From, To, CC, Date)
- Save attachments alongside emails
- Progress bars for large mailboxes
- Parallel processing support (one worker per PST file)

## Requirements

- Windows OS
- Microsoft Outlook installed
- Python 3.10+

## Installation

```bash
pip install -r requirements.txt
```

## Quick Start

```bash
# Convert a single PST file (output to same directory)
python pst_to_markdown.py -i emails.pst

# Convert multiple PST files with custom output directory
python pst_to_markdown.py -i file1.pst file2.pst -o ./output

# Use parallel workers (one per PST file)
python pst_to_markdown.py -i file1.pst file2.pst -w 2
```

## Output Structure

```
output_directory/
  2024-01-15_John Smith_Meeting Notes.md          # Email without attachments
  2024-01-16_Jane Doe_Report/                     # Email with attachments
    2024-01-16_Jane Doe_Report.md
    quarterly_report.pdf
    data.xlsx
```

## Markdown Format

Each email is converted to Markdown with the following structure:

```markdown
# Subject Line

| Field | Value |
|-------|-------|
| **From** | Sender Name <email@example.com> |
| **To** | recipient@example.com |
| **CC** | cc@example.com |
| **Date** | 2024-01-15 10:30:00 |

## Attachments

- [attachment.pdf](attachment.pdf)

## Content

Email body converted to Markdown...
```

## License

MIT
