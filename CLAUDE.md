# CLAUDE.md

## Project Overview
PST to Markdown converter - extracts emails from Outlook PST files and converts them to Markdown format.

## Tech Stack
- Python 3.10+
- pywin32 (Outlook COM interface)
- markdownify (HTML to Markdown)
- tqdm (progress bars)

## Key Files
- `pst_to_markdown.py` - Main CLI tool
- `requirements.txt` - Python dependencies

## CLI Arguments
- `-i, --input` - Input PST file(s) (required)
- `-o, --output` - Output directory (optional, defaults to input directory)
- `-w, --workers` - Parallel workers (default: 1)

## Development Notes
- Requires Outlook installed on Windows (uses COM interface)
- Emails with attachments get their own folder
- Emails without attachments are flat .md files
- Filename format: `{date}_{senderName}_{subject}.md`

## Testing
```bash
python pst_to_markdown.py -i path/to/file.pst
```
