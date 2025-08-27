"""Tests for sequential numbering fix functionality."""

import pytest

from docx2md import DocxConverter


class TestSequentialNumbering:
    """Test sequential numbering fix functionality."""

    def test_fix_sequential_numbering_basic(self):
        """Test basic sequential numbering fix."""
        converter = DocxConverter()

        content = """1. **Introduction**

Some content.

1. **Main Section**

More content.

1. **Conclusion**

Final content."""

        result = converter._fix_sequential_numbering(content)

        assert "1. **Introduction**" in result
        assert "2. **Main Section**" in result
        assert "3. **Conclusion**" in result

    def test_fix_sequential_numbering_with_spacing(self):
        """Test sequential numbering with various spacing."""
        converter = DocxConverter()

        content = """1. **First Section**

Content here.

  1. **Second Section**

More content.

    1. **Third Section**

Final content."""

        result = converter._fix_sequential_numbering(content)

        lines = result.split("\n")
        numbered_lines = [line for line in lines if "**" in line and ". **" in line]

        assert "1. **First Section**" in numbered_lines[0]
        assert "2. **Second Section**" in numbered_lines[1]
        assert "3. **Third Section**" in numbered_lines[2]

    def test_fix_sequential_numbering_mixed_content(self):
        """Test sequential numbering with mixed content."""
        converter = DocxConverter()

        content = """# Title

1. **Section One**

Some regular text here.

* A bullet point
* Another bullet point

1. **Section Two**

More text.

1. **Section Three**

Final text."""

        result = converter._fix_sequential_numbering(content)

        assert "1. **Section One**" in result
        assert "2. **Section Two**" in result
        assert "3. **Section Three**" in result

        # Should not affect other numbering
        assert "* A bullet point" in result
        assert "* Another bullet point" in result

    def test_fix_sequential_numbering_no_numbered_sections(self):
        """Test sequential numbering with no numbered sections."""
        converter = DocxConverter()

        content = """# Title

Some content without numbered sections.

## Subsection

More content.

* List item
* Another item"""

        result = converter._fix_sequential_numbering(content)

        # Should return unchanged
        assert result == content

    def test_fix_sequential_numbering_already_correct(self):
        """Test sequential numbering when already correct."""
        converter = DocxConverter()

        content = """1. **First**

Content.

2. **Second**

More content.

3. **Third**

Final content."""

        result = converter._fix_sequential_numbering(content)

        # Should only fix the "1." patterns, so this should be unchanged since it's already correct
        # But our function specifically looks for "1." patterns to fix
        assert "1. **First**" in result
        assert "2. **Second**" in result
        assert "3. **Third**" in result
        assert "3. **Third**" in result
