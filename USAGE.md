# Usage Guide

## Command Line Interface

```
pst_to_markdown.py [-h] -i PST [PST ...] [-o DIR] [-w N]
```

## Arguments

| Argument | Required | Description |
|----------|----------|-------------|
| `-i, --input` | Yes | One or more input PST files |
| `-o, --output` | No | Output directory (default: same as input PST) |
| `-w, --workers` | No | Number of parallel workers (default: 1) |
| `-h, --help` | No | Show help message |

## Examples

### Basic Usage

Convert a single PST file. Output files are placed in the same directory as the PST:

```bash
python pst_to_markdown.py -i "C:\Users\Me\emails.pst"
```

### Custom Output Directory

Specify where to save the Markdown files:

```bash
python pst_to_markdown.py -i emails.pst -o "C:\Users\Me\converted"
```

### Multiple PST Files

Process multiple PST files in one command:

```bash
python pst_to_markdown.py -i work.pst personal.pst archive.pst
```

### Parallel Processing

Use multiple workers to process PST files concurrently (one worker per PST):

```bash
python pst_to_markdown.py -i file1.pst file2.pst file3.pst -w 3
```

## Output Behavior

### Emails Without Attachments

Saved as flat `.md` files in the output directory:

```
2024-01-15_John Smith_Meeting Notes.md
```

### Emails With Attachments

Saved in a dedicated folder containing the `.md` file and all attachments:

```
2024-01-15_John Smith_Meeting Notes/
  2024-01-15_John Smith_Meeting Notes.md
  attachment1.pdf
  attachment2.docx
```

### Filename Format

```
{YYYY-MM-DD}_{SenderName}_{Subject}.md
```

- **Date**: Received date in ISO format
- **SenderName**: Extracted from "Name <email>" format
- **Subject**: Email subject line

### Duplicate Handling

If a filename already exists, a numeric suffix is added:

```
2024-01-15_Newsletter_Weekly Update.md
2024-01-15_Newsletter_Weekly Update_1.md
2024-01-15_Newsletter_Weekly Update_2.md
```

### Invalid Characters

Characters not allowed in Windows filenames (`< > : " / \ | ? *`) are replaced with underscores.

## Troubleshooting

### "Outlook not found" Error

Ensure Microsoft Outlook is installed and has been opened at least once to complete initial setup.

### PST File Not Loading

- Verify the PST file path is correct
- Ensure the PST file is not open in Outlook
- Check that you have read permissions on the file

### Slow Performance

- Large PST files with thousands of emails will take time
- Use `-w` flag with multiple PST files to parallelize
- Progress bars show estimated completion time
