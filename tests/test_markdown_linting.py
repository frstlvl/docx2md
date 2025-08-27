"""Tests for markdown linting functionality."""

import pytest

from docx2md import DocxConverter


class TestMarkdownLinting:
    """Test markdown linting rules."""

    def test_fix_multiple_blank_lines_md012(self):
        """Test MD012: Fix multiple consecutive blank lines."""
        converter = DocxConverter()

        content = """# Header 1



Some content


Another line



# Header 2"""

        result = converter._clean_markdown_content(content)

        # Should have exactly one blank line between sections
        assert "\n\n\n" not in result
        assert "Header 1\n\nSome content\n\nAnother line\n\n# Header 2" in result

    def test_surround_headers_with_blank_lines_md022(self):
        """Test MD022: Headers should be surrounded by blank lines."""
        converter = DocxConverter()

        content = """Some text
# Header 1
More text
## Header 2
Even more text"""

        result = converter._clean_markdown_content(content)

        # Headers should have blank lines before and after
        assert (
            "Some text\n\n# Header 1\n\nMore text\n\n## Header 2\n\nEven more text"
            in result
        )

    def test_surround_lists_with_blank_lines_md032(self):
        """Test MD032: Lists should be surrounded by blank lines."""
        converter = DocxConverter()

        content = """Some text
* Item 1
* Item 2
More text"""

        result = converter._clean_markdown_content(content)

        # Lists should have blank lines before and after
        assert "Some text\n\n* Item 1\n* Item 2\n\nMore text" in result

    def test_file_ends_with_newline_md047(self):
        """Test MD047: Files should end with a single newline."""
        converter = DocxConverter()

        content = "Some content without newline"
        result = converter._clean_markdown_content(content)
        assert result.endswith("\n")
        assert not result.endswith("\n\n")

    def test_is_list_item_detection(self):
        """Test list item detection."""
        converter = DocxConverter()

        assert converter._is_list_item("* Item")
        assert converter._is_list_item("- Item")
        assert converter._is_list_item("+ Item")
        assert converter._is_list_item("1. Numbered item")
        assert converter._is_list_item("10. Double digit")

        assert not converter._is_list_item("Not a list")
        assert not converter._is_list_item("# Header")
        assert not converter._is_list_item("")

    def test_complex_markdown_cleaning(self):
        """Test comprehensive markdown cleaning."""
        converter = DocxConverter()

        content = """# Title


Some intro text
## Section 1
* Point 1
* Point 2
Some text after list
### Subsection
1. First
2. Second
Final text"""

        result = converter._clean_markdown_content(content)

        # Should have proper spacing throughout
        lines = result.split("\n")

        # Find headers and check they have proper spacing
        for i, line in enumerate(lines):
            if line.startswith("#"):
                # Check blank line before (except first line)
                if i > 0:
                    assert (
                        lines[i - 1].strip() == ""
                    ), f"Header '{line}' should have blank line before"
                # Check blank line after (except if next is also header or end)
                if i < len(lines) - 1 and not lines[i + 1].startswith("#"):
                    assert (
                        lines[i + 1].strip() == ""
                    ), f"Header '{line}' should have blank line after"
                    assert (
                        lines[i + 1].strip() == ""
                    ), f"Header '{line}' should have blank line after"
