# Agent Instructions

This project uses **bd** (beads) for issue tracking. Run `bd onboard` to get started.

## Quick Reference

```bash
bd ready              # Find available work
bd show <id>          # View issue details
bd update <id> --status in_progress  # Claim work
bd close <id>         # Complete work
bd sync               # Sync with git
```

## Beads Structure & Conventions

Beads MUST follow an **epic / task / subtask** hierarchy:

- **Epic**: High-level feature or initiative (e.g., "Output format improvements")
- **Task**: A concrete, shippable unit of work within an epic (e.g., "Add --format flag for obsidian/gfm/standard")
- **Subtask**: A granular implementation step within a task (e.g., "Add Pandoc target mapping for format flag")

### Rules

1. **All tasks and subtasks MUST have a parent set** — orphan tasks are not allowed
2. **All beads (epics, tasks, subtasks) MUST have full descriptions** — title alone is insufficient
3. **Descriptions must be self-contained** — another AI agent in a fresh session with NO other context must be able to understand and complete the work from the bead description alone
4. **Include in every description**:
   - **What**: What needs to be done
   - **Why**: The motivation or user need
   - **Where**: Which files/methods/lines are affected
   - **How**: Implementation approach and key decisions
   - **Acceptance criteria**: How to verify the work is complete

### Example

```bash
# Create epic
bd create "Output format improvements" --type epic --label enhancement

# Create task under epic
bd create "Replace smart typographic characters with ASCII equivalents" \
  --type task --parent docx2md-1 --label enhancement \
  --description "Add a post-processing step _normalize_typography() in docx2md.py that replaces Unicode smart characters (curly quotes ""'', en/em dashes –—, ellipsis …, guillemets «», primes ′″) with ASCII equivalents. Always-on by default with --keep-smart-chars opt-out flag. Insert after _fix_sequential_numbering() in apply_markdown_linting_rules(). Add tests in tests/test_typography_normalization.py."

# Create subtask under task
bd create "Implement _normalize_typography method" \
  --type subtask --parent docx2md-2 \
  --description "Create _normalize_typography(self, content: str) -> str in DocxConverter class after _fix_sequential_numbering (~line 545). Use str.replace() for multi-char mappings (— -> --) then str.translate() for single-char ones. Full mapping table: \" \" -> \", ' ' -> ', – -> -, — -> --, … -> ..., « » -> \", ′ -> ', ″ -> \". Gate call in apply_markdown_linting_rules with self.normalize_typography."
```

## Landing the Plane (Session Completion)

**When ending a work session**, you MUST complete ALL steps below. Work is NOT complete until `git push` succeeds.

**MANDATORY WORKFLOW:**

1. **File issues for remaining work** - Create issues for anything that needs follow-up
2. **Run quality gates** (if code changed) - Tests, linters, builds
3. **Update issue status** - Close finished work, update in-progress items
4. **PUSH TO REMOTE** - This is MANDATORY:

   ```bash
   git pull --rebase
   bd sync
   git push
   git status  # MUST show "up to date with origin"
   ```

5. **Clean up** - Clear stashes, prune remote branches
6. **Verify** - All changes committed AND pushed
7. **Hand off** - Provide context for next session

**CRITICAL RULES:**

- Work is NOT complete until `git push` succeeds
- NEVER stop before pushing - that leaves work stranded locally
- NEVER say "ready to push when you are" - YOU must push
- If push fails, resolve and retry until it succeeds

## Project Context

### Overview

`docx2md` is a sophisticated Python CLI tool that converts Microsoft Word `.docx` files to Obsidian-friendly Markdown with YAML front matter. The project demonstrates modern Python development practices with comprehensive testing, CI/CD, and professional documentation.

### Project Structure

```text
docx2md/
├── docx2md.py              # Main module with DocxConverter class
├── tests/                  # Comprehensive test suite (43 tests, 59% coverage)
│   ├── test_docx2md.py     # Core converter tests
│   ├── test_file_filtering.py
│   ├── test_markdown_linting.py
│   ├── test_sequential_numbering.py
│   ├── test_title_extraction.py
│   └── test_toc_links.py
├── pyproject.toml          # Modern Python packaging configuration
├── README.md               # Comprehensive user documentation
├── .github/workflows/test.yml  # CI/CD pipeline
└── AGENTS.md               # This context file
```

### Technical Architecture

#### Core Components

1. **DocxConverter Class** (`docx2md.py`)
   - Main conversion engine with dual fallback strategy
   - Configurable front matter fields (new feature)
   - Rich progress display and professional CLI output
   - Comprehensive error handling and statistics tracking

2. **Conversion Engines**
   - **Primary**: Pandoc (external binary, best quality)
   - **Fallback**: Mammoth + Markdownify (pure Python)

3. **Feature Set**
   - YAML front matter extraction from DOCX core properties
   - Media extraction and organization
   - Markdown linting (MD012, MD022, MD032, MD047)
   - TOC link fixing with proper anchors
   - Sequential numbering correction
   - Title extraction with smart fallbacks
   - Temporary file filtering

#### Recent Enhancement: Configurable Front Matter

**Implementation Details:**

- Added `front_matter_fields` parameter to `DocxConverter.__init__`
- Modified `create_yaml_front_matter()` to filter properties based on configuration
- New CLI option: `--front-matter-fields` with comma-separated field list
- Available fields: `title`, `author`, `created`, `modified`, `source_file`
- Backward compatible with default `['title', 'source_file']`

**Usage Examples:**

```bash
# Default behavior
docx2md document.docx

# All fields
docx2md document.docx --front-matter-fields title,author,created,modified,source_file

# Custom subset
docx2md document.docx --front-matter-fields title,created
```

### Dependencies & Environment

#### Core Dependencies

- **Python 3.13+** (modern features, uv package manager)
- **click** (CLI framework)
- **mammoth** (DOCX→HTML conversion)
- **markdownify** (HTML→Markdown conversion)
- **rich** (terminal formatting, progress bars)

#### Development Dependencies

- **pytest** (testing framework)
- **pytest-cov** (coverage reporting)

#### External Tools

- **Pandoc** (optional, recommended for best conversion quality)
- **uv** (Python package manager, replaces pip)

### Key Implementation Patterns

#### 1. Dual Engine Strategy

```python
# Try Pandoc first (unless strict pure Python)
success = False
if not self.strict_pure_python:
   success = self.convert_with_pandoc(docx_path, output_path, media_base)

# Fall back to Mammoth if Pandoc failed
if not success:
   success = self.convert_with_mammoth(docx_path, output_path, media_base)
```

#### 2. Rich Progress Display

```python
with Progress(
   SpinnerColumn(),
   TextColumn("[progress.description]{task.description}"),
   BarColumn(),
   TaskProgressColumn(),
   console=console,
) as progress:
   # Conversion logic with progress updates
```

#### 3. Configurable Front Matter

```python
def create_yaml_front_matter(self, properties: Dict[str, Any]) -> str:
   # Filter properties based on configured fields
   filtered_properties = {}
   for field in self.front_matter_fields:
      if field in properties and properties[field]:
         filtered_properties[field] = properties[field]
```

### Testing Strategy

#### Test Coverage

- 43 tests, 59% coverage
- **Unit Tests**: Core functionality, edge cases, error handling
- **Integration Tests**: Real DOCX conversion with extracted properties
- **Feature Tests**: Each markdown linting rule, TOC fixing, numbering
- **File Filtering Tests**: Temporary file detection, format validation

#### Key Test Files

- `tests/test_docx2md.py`: Core converter functionality
- `tests/test_markdown_linting.py`: MD012, MD022, MD032, MD047 rules
- `tests/test_toc_links.py`: TOC anchor generation and link fixing
- `tests/test_title_extraction.py`: Smart title detection algorithms
- `tests/test_sequential_numbering.py`: Header numbering correction

#### Test Execution

```bash
# Run all tests
uv run pytest

# Run with coverage
uv run pytest --cov=docx2md

# Verbose output
uv run pytest -v
```

### CI/CD Pipeline

#### GitHub Actions (`.github/workflows/test.yml`)

- **Matrix Testing**: Ubuntu, Windows, macOS
- **Python 3.13**: Latest stable version
- **Pandoc Installation**: Platform-specific setup
- **Dependency Management**: `uv sync --group test`
- **Test Execution**: Full test suite on all platforms

#### Pipeline Features

- Automated testing on pull requests
- Cross-platform compatibility validation
- Dependency caching for faster builds
- No external service dependencies (removed Codecov)

### File Processing Logic

#### 1. Discovery Phase

```python
def discover_docx_files(self, inputs: List[Path], recursive: bool = False):
   # Scan for .docx files
   # Filter temporary files (~$, .~lock.)
   # Skip unsupported formats (.doc, .docm)
   # Return (file_path, input_root) tuples
```

#### 2. Conversion Phase

```python
def convert_single_file(self, docx_path: Path, input_root: Optional[Path] = None):
   # Extract document properties
   # Try Pandoc → Mammoth conversion
   # Add YAML front matter
   # Apply markdown linting rules
   # Clean up empty media directories
```

#### 3. Post-Processing

- **Markdown Linting**: Fix common formatting issues
- **TOC Link Repair**: Convert Word TOC links to markdown anchors
- **Sequential Numbering**: Fix flattened numbered sections
- **Media Cleanup**: Remove empty directories

### Error Handling & User Experience

#### Graceful Degradation

- Pandoc unavailable → Mammoth fallback
- Properties extraction fails → Continue without front matter
- Individual file failures → Continue batch processing

#### Rich Terminal Output

- Colored status messages (success/skip/error)
- Progress bars for batch operations
- Formatted summary tables
- Professional error reporting

#### Exit Codes

- `0`: All conversions successful
- `1`: Some conversions failed
- `2`: No valid input files found

### Common Development Tasks

#### Adding New Features

1. Implement functionality in `DocxConverter` class
2. Add CLI option in `main()` function
3. Write comprehensive tests
4. Update documentation
5. Test across platforms

#### Adding New Tests

1. Create test file in `tests/` directory
2. Use descriptive test class/method names
3. Include edge cases and error conditions
4. Maintain good test coverage

#### Debugging Conversion Issues

1. Use `--verbose` flag for detailed logging
2. Test with `--strict-pure-python` to isolate Pandoc issues
3. Check DOCX properties with test documents
4. Verify media extraction in output directories

### Code Quality Standards

#### Modern Python Practices

- **Type Hints**: Full PEP 585 compliance
- **pathlib**: Used throughout for path handling
- **f-strings**: Modern string formatting
- **Context Managers**: Proper resource management
- **Dataclasses**: Where appropriate for structured data

#### Error Handling Philosophy

- **Fail Gracefully**: Continue processing on individual errors
- **User-Friendly Messages**: Clear error descriptions
- **Logging**: Appropriate levels (DEBUG, INFO, WARNING, ERROR)
- **Exit Codes**: Standard conventions

### Integration Points

#### Obsidian Compatibility

- YAML front matter format compatible with Obsidian
- Relative media paths work in Obsidian vaults
- Markdown follows Obsidian conventions
- Link syntax matches Obsidian expectations

#### External Tool Integration

- **Pandoc**: Primary conversion engine
- **uv**: Modern Python dependency management
- **GitHub Actions**: Automated testing and validation

### Performance Considerations

#### Batch Processing

- Progress reporting for user feedback
- Memory-efficient file-by-file processing
- Parallel-safe design (though currently sequential)

#### Media Handling

- Organized directory structure prevents conflicts
- Empty directory cleanup reduces clutter
- Relative paths ensure portability

### Future Enhancement Opportunities

#### Potential Features

1. **Parallel Processing**: Batch conversion with multiprocessing
2. **Custom Templates**: User-defined front matter templates
3. **Format Extensions**: Support for .doc, .rtf files
4. **Plugin System**: Extensible post-processing pipeline
5. **Configuration Files**: `.docx2md.toml` for project settings

#### Architecture Improvements

1. **Plugin Architecture**: Modular post-processing
2. **Async Processing**: For large batch operations
3. **Caching**: Avoid re-converting unchanged files
4. **Preview Mode**: Dry-run capabilities

### Troubleshooting Guide

#### Common Issues

1. **Pandoc Not Found**: Install Pandoc or use `--strict-pure-python`
2. **Permission Errors**: Check output directory write permissions
3. **Encoding Issues**: Ensure UTF-8 support in terminal
4. **Empty Output**: Check input file validity

#### Debug Commands

```bash
# Verbose output
docx2md document.docx -v

# Pure Python mode
docx2md document.docx --strict-pure-python

# Test with single file
docx2md single_file.docx -o test_output/
```

This project represents a mature, well-tested CLI tool with modern Python development practices, comprehensive testing, and production-ready CI/CD pipeline.
