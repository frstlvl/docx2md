# docx2md

[![Tests](https://github.com/frstlvl/docx2md/workflows/Tests/badge.svg)](https://github.com/frstlvl/docx2md/actions)

Convert Microsoft Word `.docx` files to Obsidian-friendly Markdown with YAML front matter extracted from document properties.

## Features

- **Multiple Input Types**: Convert single files or entire directories
- **YAML Front Matter**: Automatically extract title, author, creation/modification dates from DOCX properties
- **Media Extraction**: Extract and organize embedded images with proper relative paths
- **Dual Conversion Engines**:
  - Primary: Pandoc (for best results)
  - Fallback: Mammoth + Markdownify (pure Python)
- **Obsidian Compatible**: Generated Markdown works seamlessly in Obsidian
- **Flexible Output**: Preserve directory structure or flatten to single directory
- **Smart Filename Handling**: Spaces → underscores, case preservation

## Installation

### Requirements

- Python 3.13+
- Pandoc (recommended, optional)

### Install with uv

```bash
# Clone and install
git clone https://github.com/frstlvl/docx2md.git
cd docx2md
uv sync
```

### Install Pandoc (Recommended)

For best conversion quality, install Pandoc:

- **Windows**: Download from [pandoc.org](https://pandoc.org/installing.html)
- **macOS**: `brew install pandoc`
- **Linux**: `sudo apt install pandoc` or equivalent

## Usage

### Basic Examples

```bash
# Convert single file (output in same directory)
docx2md document.docx

# Convert to specific output directory
docx2md document.docx -o output/

# Convert all .docx files in a directory
docx2md input_folder/ -o output/

# Recursive conversion preserving structure
docx2md input_folder/ -r -o output/

# Force pure Python (skip Pandoc)
docx2md document.docx --strict-pure-python
```

### Command Line Options

```python
Usage: docx2md [OPTIONS] INPUTS...

  Convert DOCX files to Obsidian-friendly Markdown.

Options:
  -o, --output-dir PATH        Output directory for .md files
  -r, --recursive              Recursively search directories for .docx files
  --no-preserve-structure      Do not preserve directory structure in output
  --overwrite                  Overwrite existing .md files
  --media-dir TEXT             Media directory name (default: media)
  --pandoc-path PATH           Path to pandoc executable
  --strict-pure-python         Use only Python libraries (skip Pandoc)
  --no-front-matter            Disable YAML front matter generation
  -v, --verbose                Enable verbose logging
  --help                       Show this message and exit.
```

## Output Format

### YAML Front Matter

When document properties are available, docx2md adds YAML front matter:

```yaml
---
title: Document Title
author: Author Name
created: 2024-01-15T10:30:00Z
modified: 2024-01-16T14:20:00Z
source_file: original_document.docx
---
```

### Media Organization

Images are extracted to organized directories:

```text
output/
├── document_name.md
└── media/
    └── document_name/
        ├── image1.png
        ├── image2.jpg
        └── ...
```

### Filename Sanitization

- Spaces replaced with underscores
- Invalid characters removed
- Case preserved
- Example: `"My Document.docx"` → `"My_Document.md"`

## Conversion Engines

### Pandoc (Primary)

- Best quality conversion
- GitHub Flavored Markdown output
- Advanced formatting support
- Automatic media extraction

### Mammoth + Markdownify (Fallback)

- Pure Python implementation
- No external dependencies
- Good for basic documents
- Automatic fallback when Pandoc unavailable

## Exit Codes

- `0`: All conversions succeeded
- `1`: At least one conversion failed
- `2`: No .docx files found

## Development

### Setup Development Environment

```bash
git clone https://github.com/frstlvl/docx2md.git
cd docx2md
uv sync --dev
```

### Running Tests

```bash
uv run pytest
```

### Code Standards

- PEP 8 compliance
- Type hints (PEP 585)
- pathlib for path handling
- Comprehensive error handling

## Supported Formats

- ✅ `.docx` (Office Open XML)
- ❌ `.doc` (legacy binary format)
- ❌ `.docm` (macro-enabled documents)

## Troubleshooting

### Common Issues

1. **Pandoc not found**: Install Pandoc or use `--strict-pure-python`
2. **Permission errors**: Ensure write access to output directory
3. **Corrupt DOCX**: Check file integrity, try with different converter

### Verbose Output

Use `-v` flag for detailed logging:

```bash
docx2md document.docx -v
```

## License

MIT License - see LICENSE file for details.

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests
5. Submit a pull request

## Acknowledgments

- [Pandoc](https://pandoc.org/) - Universal document converter
- [Mammoth](https://github.com/mwilliamson/python-mammoth) - DOCX to HTML converter
- [Markdownify](https://github.com/matthewwithanm/python-markdownify) - HTML to Markdown converter
