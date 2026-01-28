# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.0.0] - 2025-01-28

### Added

- Initial release
- Extract emails from Outlook PST files via COM interface
- Convert HTML email bodies to Markdown using markdownify
- Preserve email metadata (From, To, CC, Date) in table format
- Save attachments in dedicated folders per email
- Progress bars with tqdm for large mailboxes
- Parallel processing support with configurable workers
- CLI with `-i` (input), `-o` (output), `-w` (workers) arguments
- Automatic sender name extraction from "Name <email>" format
- Filename sanitization for Windows compatibility
- Duplicate filename handling with numeric suffixes
- Recursive folder traversal with flattened output
