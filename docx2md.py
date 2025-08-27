#!/usr/bin/env python3
"""
docx2md: Convert DOCX files to Obsidian-friendly Markdown with YAML front matter.

This tool converts Microsoft Word .docx files to Markdown format, extracting
embedded media and optionally adding YAML front matter from document properties.
Prefers Pandoc for conversion but falls back to Mammoth+Markdownify if unavailable.
"""

import logging
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
from xml.etree import ElementTree as ET

import click
import mammoth
from markdownify import markdownify
from rich import print as rprint
from rich.console import Console
from rich.panel import Panel
from rich.progress import (
    BarColumn,
    Progress,
    SpinnerColumn,
    TaskProgressColumn,
    TextColumn,
)
from rich.table import Table
from rich.text import Text

# Initialize Rich console
console = Console()

# Configure logging to work with Rich
logging.basicConfig(
    level=logging.INFO,
    format="%(levelname)s: %(message)s",
    handlers=[logging.StreamHandler()],
)
logger = logging.getLogger(__name__)


class DocxConverter:
    """Main converter class handling DOCX to Markdown conversion."""

    def __init__(
        self,
        output_dir: Optional[Path] = None,
        preserve_structure: bool = True,
        overwrite: bool = False,
        media_dir: str = "media",
        pandoc_path: Optional[Path] = None,
        strict_pure_python: bool = False,
        enable_front_matter: bool = True,
    ):
        self.output_dir = output_dir
        self.preserve_structure = preserve_structure
        self.overwrite = overwrite
        self.media_dir = media_dir
        self.pandoc_path = pandoc_path
        self.strict_pure_python = strict_pure_python
        self.enable_front_matter = enable_front_matter

        # Track conversion statistics
        self.stats = {"success": 0, "skipped": 0, "failed": 0}

    def sanitize_filename(self, name: str) -> str:
        """Sanitize filename: spaces to underscores, preserve case."""
        # Replace spaces with underscores
        sanitized = name.replace(" ", "_")

        # Remove or replace invalid filename characters
        sanitized = re.sub(r'[<>:"/\\|?*]', "", sanitized)

        return sanitized

    def find_pandoc(self) -> Optional[Path]:
        """Find Pandoc executable."""
        if self.pandoc_path:
            if self.pandoc_path.exists():
                return self.pandoc_path
            else:
                logger.warning(f"Specified pandoc path not found: {self.pandoc_path}")
                return None

        # Try to find pandoc in PATH
        pandoc_cmd = shutil.which("pandoc")
        if pandoc_cmd:
            return Path(pandoc_cmd)

        return None

    def extract_core_properties(self, docx_path: Path) -> Dict[str, Any]:
        """Extract Dublin Core properties from DOCX file."""
        properties = {}

        try:
            with zipfile.ZipFile(docx_path, "r") as docx_zip:
                # Try to read core properties
                try:
                    core_xml = docx_zip.read("docProps/core.xml")
                    root = ET.fromstring(core_xml)

                    # Define namespace mappings
                    namespaces = {
                        "dc": "http://purl.org/dc/elements/1.1/",
                        "dcterms": "http://purl.org/dc/terms/",
                        "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
                    }

                    # Extract properties
                    title_elem = root.find(".//dc:title", namespaces)
                    if title_elem is not None and title_elem.text:
                        properties["title"] = title_elem.text

                    creator_elem = root.find(".//dc:creator", namespaces)
                    if creator_elem is not None and creator_elem.text:
                        properties["author"] = creator_elem.text

                    created_elem = root.find(".//dcterms:created", namespaces)
                    if created_elem is not None and created_elem.text:
                        properties["created"] = created_elem.text

                    modified_elem = root.find(".//dcterms:modified", namespaces)
                    if modified_elem is not None and modified_elem.text:
                        properties["modified"] = modified_elem.text

                except KeyError:
                    logger.debug(f"No core properties found in {docx_path}")

        except (zipfile.BadZipFile, ET.ParseError) as e:
            logger.warning(f"Could not extract properties from {docx_path}: {e}")

        # Always add source file
        properties["source_file"] = docx_path.name

        return properties

    def extract_title_from_markdown(self, md_content: str) -> Optional[str]:
        """Extract title from the first heading in markdown content."""
        lines = md_content.split("\n")

        for line in lines:
            line = line.strip()
            # Look for H1 headings (# Title or **Title**)
            if line.startswith("# "):
                return line[2:].strip()
            # Look for bold text that might be a title (common in converted docs)
            elif line.startswith("**") and line.endswith("**") and len(line) > 4:
                potential_title = line[2:-2].strip()
                # Only consider it a title if it's not too long and doesn't contain common body text indicators
                if len(potential_title) < 100 and not any(
                    word in potential_title.lower()
                    for word in [
                        "innehållsförteckning",
                        "table of contents",
                        "inledning",
                        "introduction",
                    ]
                ):
                    return potential_title

        return None

    def is_generic_title(self, title: str) -> bool:
        """Check if a title appears to be generic or placeholder text."""
        if not title:
            return True

        generic_patterns = [
            r"^report\s*v?\d*\.?\d*$",
            r"^document\s*v?\d*\.?\d*$",
            r"^untitled",
            r"^new\s+document",
            r"^draft",
            r"^\s*$",
        ]

        return any(
            re.match(pattern, title.strip().lower()) for pattern in generic_patterns
        )

    def create_yaml_front_matter(self, properties: Dict[str, Any]) -> str:
        """Create YAML front matter from properties."""
        if not properties:
            return ""

        # Only include title and source_file in front matter
        filtered_properties = {}

        if "title" in properties and properties["title"]:
            filtered_properties["title"] = properties["title"]

        if "source_file" in properties and properties["source_file"]:
            filtered_properties["source_file"] = properties["source_file"]

        if not filtered_properties:
            return ""

        lines = ["---"]
        for key, value in filtered_properties.items():
            if value:
                # Escape quotes in values
                if isinstance(value, str) and ('"' in value or "'" in value):
                    value = repr(value)
                lines.append(f"{key}: {value}")
        lines.append("---")
        lines.append("")  # Empty line after front matter

        return "\n".join(lines)

    def convert_with_pandoc(
        self, docx_path: Path, output_path: Path, media_base: Path
    ) -> bool:
        """Convert DOCX to Markdown using Pandoc."""
        pandoc = self.find_pandoc()
        if not pandoc:
            return False

        try:
            # Create media directory for this document
            doc_stem = self.sanitize_filename(docx_path.stem)
            media_dir = media_base / doc_stem
            media_dir.mkdir(parents=True, exist_ok=True)

            # Pandoc command
            cmd = [
                str(pandoc),
                str(docx_path),
                "-f",
                "docx",
                "-t",
                "gfm",
                "--wrap=auto",
                f"--extract-media={media_dir}",
                "-o",
                str(output_path),
            ]

            result = subprocess.run(cmd, capture_output=True, text=True, check=True)

            logger.debug(f"Pandoc output: {result.stdout}")
            return True

        except subprocess.CalledProcessError as e:
            logger.error(f"Pandoc failed for {docx_path}: {e.stderr}")
            return False
        except Exception as e:
            logger.error(f"Error running Pandoc for {docx_path}: {e}")
            return False

    def convert_with_mammoth(
        self, docx_path: Path, output_path: Path, media_base: Path
    ) -> bool:
        """Convert DOCX to Markdown using Mammoth + Markdownify."""
        try:
            # Create media directory for this document
            doc_stem = self.sanitize_filename(docx_path.stem)
            media_dir = media_base / doc_stem
            media_dir.mkdir(parents=True, exist_ok=True)

            # Convert DOCX to HTML with Mammoth
            with open(docx_path, "rb") as docx_file:
                result = mammoth.convert_to_html(docx_file)
                html = result.value

                if result.messages:
                    for message in result.messages:
                        logger.debug(f"Mammoth message: {message}")

            # Convert HTML to Markdown
            markdown = markdownify(html, heading_style="ATX")

            # Save to file
            with open(output_path, "w", encoding="utf-8") as md_file:
                md_file.write(markdown)

            return True

        except Exception as e:
            logger.error(f"Mammoth conversion failed for {docx_path}: {e}")
            return False

    def add_front_matter_to_file(self, md_path: Path, properties: Dict[str, Any]):
        """Add YAML front matter to an existing Markdown file."""
        if not self.enable_front_matter or not properties:
            return

        try:
            # Read existing content
            with open(md_path, "r", encoding="utf-8") as f:
                content = f.read()

            # Check if the title from metadata is generic, and if so, try to extract from content
            metadata_title = properties.get("title", "")
            if self.is_generic_title(metadata_title):
                content_title = self.extract_title_from_markdown(content)
                if content_title:
                    logger.debug(
                        f"Using content title '{content_title}' instead of generic metadata title '{metadata_title}'"
                    )
                    properties["title"] = content_title

            # Create front matter
            front_matter = self.create_yaml_front_matter(properties)

            # Write back with front matter
            with open(md_path, "w", encoding="utf-8") as f:
                f.write(front_matter + content)

        except Exception as e:
            logger.warning(f"Could not add front matter to {md_path}: {e}")

    def cleanup_empty_media_dirs(self, media_base: Path, doc_stem: str):
        """Remove empty media directories after conversion."""
        try:
            media_dir = media_base / doc_stem
            if media_dir.exists() and media_dir.is_dir():
                # Check if directory is empty
                if not any(media_dir.iterdir()):
                    media_dir.rmdir()
                    logger.debug(f"Removed empty media directory: {media_dir}")

                    # Also try to remove parent media directory if it becomes empty
                    if media_base.exists() and media_base.is_dir():
                        if not any(media_base.iterdir()):
                            media_base.rmdir()
                            logger.debug(
                                f"Removed empty parent media directory: {media_base}"
                            )
        except Exception as e:
            logger.debug(f"Could not cleanup media directories: {e}")

    def apply_markdown_linting_rules(self, md_path: Path):
        """Apply markdown linting rules as specified in copilot.md."""
        try:
            with open(md_path, "r", encoding="utf-8") as f:
                content = f.read()

            # Apply the required markdown rules
            cleaned_content = self._clean_markdown_content(content)

            # Fix TOC links
            cleaned_content = self._fix_toc_links(cleaned_content)

            # Fix sequential numbering
            cleaned_content = self._fix_sequential_numbering(cleaned_content)

            # Write back the cleaned content
            with open(md_path, "w", encoding="utf-8") as f:
                f.write(cleaned_content)

        except Exception as e:
            logger.debug(f"Could not apply markdown linting rules to {md_path}: {e}")

    def _clean_markdown_content(self, content: str) -> str:
        """Clean markdown content according to linting rules."""
        lines = content.split("\n")
        cleaned_lines = []
        i = 0

        while i < len(lines):
            line = lines[i]

            # MD022: Surround headings with blank lines
            if line.strip().startswith("#"):
                # Add blank line before heading (if not already there and not at start)
                if (
                    cleaned_lines
                    and cleaned_lines[-1].strip() != ""
                    and not cleaned_lines[-1].strip().startswith("#")
                ):
                    cleaned_lines.append("")

                cleaned_lines.append(line)

                # Add blank line after heading (if next line isn't blank and exists)
                if (
                    i + 1 < len(lines)
                    and lines[i + 1].strip() != ""
                    and not lines[i + 1].strip().startswith("#")
                ):
                    cleaned_lines.append("")

            # MD032: Surround lists with blank lines
            elif self._is_list_item(line):
                # Add blank line before list (if not already there)
                if (
                    cleaned_lines
                    and cleaned_lines[-1].strip() != ""
                    and not self._is_list_item(cleaned_lines[-1])
                ):
                    cleaned_lines.append("")

                # Add all consecutive list items
                while i < len(lines) and (
                    self._is_list_item(lines[i]) or lines[i].strip() == ""
                ):
                    cleaned_lines.append(lines[i])
                    i += 1
                i -= 1  # Adjust for the increment at end of loop

                # Add blank line after list (if next line exists and isn't blank)
                if (
                    i + 1 < len(lines)
                    and lines[i + 1].strip() != ""
                    and not self._is_list_item(lines[i + 1])
                ):
                    cleaned_lines.append("")

            else:
                cleaned_lines.append(line)

            i += 1

        # MD047: End file with single newline character
        result = "\n".join(cleaned_lines)

        # MD012: Remove multiple consecutive blank lines
        # Replace 2 or more consecutive blank lines with exactly 1 blank line
        result = re.sub(r"\n\s*\n\s*\n+", "\n\n", result)

        # Remove any trailing whitespace/newlines and add single newline
        result = result.rstrip() + "\n"

        return result

    def _is_list_item(self, line: str) -> bool:
        """Check if a line is a list item."""
        stripped = line.strip()
        if not stripped:
            return False

        # Unordered lists: -, *, +
        if stripped.startswith(("- ", "* ", "+ ")):
            return True

        # Ordered lists: number followed by . or )
        import re

        if re.match(r"^\d+[\.\)] ", stripped):
            return True

        return False

    def _create_heading_anchor(self, heading_text: str) -> str:
        """Create a proper markdown anchor from heading text."""
        # Remove markdown formatting and extra whitespace
        clean_text = re.sub(r"[#*_`]", "", heading_text).strip()

        # Convert to lowercase
        anchor = clean_text.lower()

        # Replace spaces and special characters with hyphens
        anchor = re.sub(r"[^\w\s-]", "", anchor)
        anchor = re.sub(r"[-\s]+", "-", anchor)

        # Remove leading/trailing hyphens
        anchor = anchor.strip("-")

        return anchor

    def _fix_toc_links(self, content: str) -> str:
        """Fix table of contents links to use proper markdown anchors."""
        lines = content.split("\n")

        # First pass: collect all headings and their anchors
        headings = {}
        for line in lines:
            if line.strip().startswith("#"):
                heading_text = re.sub(r"^#+\s*", "", line.strip())
                anchor = self._create_heading_anchor(heading_text)
                # Store both the original heading text and variations for matching
                headings[heading_text] = anchor
                # Also store simplified versions for better matching
                simplified = re.sub(
                    r"\s+\d+$", "", heading_text
                )  # Remove trailing numbers
                headings[simplified] = anchor

        # Second pass: fix TOC links
        fixed_lines = []
        for line in lines:
            # Look for TOC-style links: [text](#_Toc123456)
            toc_pattern = r"\[([^\]]+)\]\(#_Toc\d+\)"

            def replace_toc_link(match):
                link_text = match.group(1)

                # Extract the heading text from the link text
                # Remove numbering at the start and page numbers at the end
                clean_heading = re.sub(
                    r"^\d+\.?\s*", "", link_text
                )  # Remove leading numbers
                clean_heading = re.sub(
                    r"\s+\d+$", "", clean_heading
                )  # Remove trailing page numbers
                clean_heading = clean_heading.strip()

                # Find matching heading
                if clean_heading in headings:
                    anchor = headings[clean_heading]
                    return f"[{link_text}](#{anchor})"

                # Try partial matching
                for heading_text, anchor in headings.items():
                    if (
                        clean_heading.lower() in heading_text.lower()
                        or heading_text.lower() in clean_heading.lower()
                    ):
                        return f"[{link_text}](#{anchor})"

                # If no match found, remove the link but keep the text
                return link_text

            fixed_line = re.sub(toc_pattern, replace_toc_link, line)
            fixed_lines.append(fixed_line)

        return "\n".join(fixed_lines)

    def _fix_sequential_numbering(self, content: str) -> str:
        """Fix sequential numbering of headers that got flattened during conversion."""
        lines = content.split("\n")
        section_counter = 0

        for i, line in enumerate(lines):
            # Look for lines that start with "1. **" (indicating a numbered section)
            if re.match(r"^\s*1\.\s*\*\*", line.strip()):
                section_counter += 1
                # Replace "1." with the correct sequential number
                lines[i] = re.sub(r"^\s*1\.", f"{section_counter}.", line)

        return "\n".join(lines)

    def convert_single_file(
        self, docx_path: Path, input_root: Optional[Path] = None
    ) -> bool:
        """Convert a single DOCX file to Markdown."""
        try:
            # Determine output path
            if self.output_dir:
                if self.preserve_structure and input_root:
                    # Preserve directory structure
                    rel_path = docx_path.relative_to(input_root)
                    output_path = self.output_dir / rel_path.with_suffix(".md")
                else:
                    # Flat structure
                    output_path = (
                        self.output_dir / f"{self.sanitize_filename(docx_path.stem)}.md"
                    )
            else:
                # Same directory as input
                output_path = docx_path.with_suffix(".md")
                output_path = output_path.with_name(
                    self.sanitize_filename(output_path.name)
                )

            # Check if output already exists
            if output_path.exists() and not self.overwrite:
                # Use console.print instead of logger for better formatting
                self.stats["skipped"] += 1
                return True

            # Create output directory
            output_path.parent.mkdir(parents=True, exist_ok=True)

            # Determine media base directory
            if self.output_dir:
                media_base = self.output_dir / self.media_dir
            else:
                media_base = docx_path.parent / self.media_dir

            # Extract document properties for front matter
            properties = {}
            if self.enable_front_matter:
                properties = self.extract_core_properties(docx_path)

            # Try conversion with Pandoc first (unless strict pure Python)
            success = False
            if not self.strict_pure_python:
                success = self.convert_with_pandoc(docx_path, output_path, media_base)

            # Fall back to Mammoth if Pandoc failed or unavailable
            if not success:
                success = self.convert_with_mammoth(docx_path, output_path, media_base)

            if success:
                # Add front matter if enabled
                if self.enable_front_matter and properties:
                    self.add_front_matter_to_file(output_path, properties)

                # Apply markdown linting rules
                self.apply_markdown_linting_rules(output_path)

                # Clean up empty media directories
                doc_stem = self.sanitize_filename(docx_path.stem)
                self.cleanup_empty_media_dirs(media_base, doc_stem)

                self.stats["success"] += 1
                return True
            else:
                # Clean up empty media directories even on failure
                doc_stem = self.sanitize_filename(docx_path.stem)
                self.cleanup_empty_media_dirs(media_base, doc_stem)

                self.stats["failed"] += 1
                return False

        except Exception as e:
            self.stats["failed"] += 1
            return False

    def is_temporary_file(self, file_path: Path) -> bool:
        """Check if a file is a temporary Word file that should be skipped."""
        filename = file_path.name

        # Word temporary files start with ~$
        if filename.startswith("~$"):
            return True

        # Word lock files start with .~lock.
        if filename.startswith(".~lock."):
            return True

        # Hidden files (starting with .)
        if filename.startswith(".") and filename != ".":
            return True

        return False

    def get_file_skip_reason(self, file_path: Path) -> Optional[str]:
        """Get the reason why a file should be skipped, or None if it shouldn't be."""
        if self.is_temporary_file(file_path):
            if file_path.name.startswith("~$"):
                return "Word temporary/lock file"
            elif file_path.name.startswith(".~lock."):
                return "LibreOffice lock file"
            elif file_path.name.startswith("."):
                return "hidden file"

        if file_path.suffix.lower() in [".doc", ".docm"]:
            return "unsupported format (.doc/.docm)"

        if file_path.suffix.lower() != ".docx":
            return "not a .docx file"

        return None

    def discover_docx_files(
        self, inputs: List[Path], recursive: bool = False
    ) -> List[Tuple[Path, Optional[Path]]]:
        """Discover DOCX files from input paths. Returns (file_path, input_root) tuples."""
        files = []
        skipped_count = 0

        for input_path in inputs:
            if input_path.is_file():
                skip_reason = self.get_file_skip_reason(input_path)
                if skip_reason:
                    console.print(
                        f"[yellow]SKIP[/yellow]: {input_path.name} ([dim]{skip_reason}[/dim])"
                    )
                    skipped_count += 1
                else:
                    files.append((input_path, input_path.parent))
            elif input_path.is_dir():
                pattern = "**/*.docx" if recursive else "*.docx"
                docx_files = list(input_path.glob(pattern))
                for docx_file in docx_files:
                    skip_reason = self.get_file_skip_reason(docx_file)
                    if skip_reason:
                        console.print(
                            f"[yellow]SKIP[/yellow]: {docx_file.name} ([dim]{skip_reason}[/dim])"
                        )
                        skipped_count += 1
                    else:
                        files.append((docx_file, input_path))
            else:
                console.print(f"[red]ERROR[/red]: {input_path} (path not found)")

        if skipped_count > 0:
            console.print(
                f"[dim]Skipped {skipped_count} non-DOCX or temporary file(s)[/dim]"
            )

        return files

    def convert_files(self, inputs: List[Path], recursive: bool = False) -> int:
        """Convert multiple DOCX files. Returns exit code."""

        # Print header
        console.print()
        console.print(
            Panel.fit(
                "[bold cyan]DOCX to Markdown Converter[/bold cyan]\n"
                "Converting Word documents to Obsidian-friendly Markdown",
                border_style="cyan",
            )
        )

        # Discover all DOCX files
        console.print(f"\n[bold]Scanning for .docx files...[/bold]")
        files = self.discover_docx_files(inputs, recursive)

        if not files:
            console.print("[red bold]✗ No valid .docx files found[/red bold]")
            return 2

        console.print(
            f"\n[green]✓ Found {len(files)} valid .docx file(s) to convert[/green]"
        )

        # Convert each file with progress
        with Progress(
            SpinnerColumn(),
            TextColumn("[progress.description]{task.description}"),
            BarColumn(),
            TaskProgressColumn(),
            console=console,
        ) as progress:
            task = progress.add_task("Converting files...", total=len(files))

            for docx_path, input_root in files:
                progress.update(task, description=f"Converting {docx_path.name}")
                self.convert_single_file(docx_path, input_root)
                progress.advance(task)

        # Print summary table
        self._print_summary_table()

        # Return appropriate exit code
        if self.stats["failed"] > 0:
            return 1
        return 0

    def _print_summary_table(self):
        """Print a formatted summary table using Rich."""
        total = self.stats["success"] + self.stats["skipped"] + self.stats["failed"]

        table = Table(
            title="Conversion Summary", show_header=True, header_style="bold magenta"
        )
        table.add_column("Status", style="bold", justify="center")
        table.add_column("Count", justify="right")
        table.add_column("Percentage", justify="right")

        if self.stats["success"] > 0:
            pct = (self.stats["success"] / total) * 100
            table.add_row(
                "[green]✓ Succeeded[/green]", str(self.stats["success"]), f"{pct:.1f}%"
            )

        if self.stats["skipped"] > 0:
            pct = (self.stats["skipped"] / total) * 100
            table.add_row(
                "[yellow]⊘ Skipped[/yellow]", str(self.stats["skipped"]), f"{pct:.1f}%"
            )

        if self.stats["failed"] > 0:
            pct = (self.stats["failed"] / total) * 100
            table.add_row(
                "[red]✗ Failed[/red]", str(self.stats["failed"]), f"{pct:.1f}%"
            )

        table.add_row(
            "[bold]Total[/bold]", f"[bold]{total}[/bold]", "[bold]100.0%[/bold]"
        )

        console.print()
        console.print(table)
        console.print()


@click.command()
@click.argument(
    "inputs", nargs=-1, required=True, type=click.Path(exists=True, path_type=Path)
)
@click.option(
    "--output-dir",
    "-o",
    type=click.Path(path_type=Path),
    help="Output directory for .md files",
)
@click.option(
    "--recursive",
    "-r",
    is_flag=True,
    help="Recursively search directories for .docx files",
)
@click.option(
    "--no-preserve-structure",
    is_flag=True,
    help="Do not preserve directory structure in output",
)
@click.option("--overwrite", is_flag=True, help="Overwrite existing .md files")
@click.option(
    "--media-dir", default="media", help="Media directory name (default: media)"
)
@click.option(
    "--pandoc-path", type=click.Path(path_type=Path), help="Path to pandoc executable"
)
@click.option(
    "--strict-pure-python", is_flag=True, help="Use only Python libraries (skip Pandoc)"
)
@click.option(
    "--no-front-matter", is_flag=True, help="Disable YAML front matter generation"
)
@click.option("--verbose", "-v", is_flag=True, help="Enable verbose logging")
def main(
    inputs: Tuple[Path, ...],
    output_dir: Optional[Path],
    recursive: bool,
    no_preserve_structure: bool,
    overwrite: bool,
    media_dir: str,
    pandoc_path: Optional[Path],
    strict_pure_python: bool,
    no_front_matter: bool,
    verbose: bool,
):
    """Convert DOCX files to Obsidian-friendly Markdown.

    INPUTS can be a mix of .docx files and directories containing .docx files.

    Examples:

        # Convert single file to same directory
        docx2md document.docx

        # Convert file to specific output directory
        docx2md document.docx -o output/

        # Convert all .docx files in directory recursively
        docx2md input_folder/ -r -o output/

        # Force pure Python conversion (skip Pandoc)
        docx2md document.docx --strict-pure-python
    """

    # Set logging level
    if verbose:
        logging.getLogger().setLevel(logging.DEBUG)

    # Create converter
    converter = DocxConverter(
        output_dir=output_dir,
        preserve_structure=not no_preserve_structure,
        overwrite=overwrite,
        media_dir=media_dir,
        pandoc_path=pandoc_path,
        strict_pure_python=strict_pure_python,
        enable_front_matter=not no_front_matter,
    )

    # Convert files
    exit_code = converter.convert_files(list(inputs), recursive)
    sys.exit(exit_code)


if __name__ == "__main__":
    main()
