"""Tests for title extraction functionality."""

import pytest

from docx2md import DocxConverter


class TestTitleExtraction:
    """Test smart title extraction from content."""

    def test_extract_title_from_markdown_with_h1(self):
        """Test title extraction from H1 header."""
        converter = DocxConverter()

        content = """# Document Title

Some content here.

## Section 1

More content."""

        title = converter.extract_title_from_markdown(content)
        assert title == "Document Title"

    def test_extract_title_from_markdown_with_bold_first_line(self):
        """Test title extraction from bold first line."""
        converter = DocxConverter()

        content = """**Important Document**

This is the content of the document.

## Section 1

More content."""

        title = converter.extract_title_from_markdown(content)
        assert title == "Important Document"

    def test_extract_title_from_markdown_no_clear_title(self):
        """Test title extraction when no clear title exists."""
        converter = DocxConverter()

        content = """This is just regular text.

## Section 1

Some content here.

## Section 2

More content."""

        title = converter.extract_title_from_markdown(content)
        assert title is None

    def test_extract_title_from_markdown_multiple_h1(self):
        """Test title extraction with multiple H1 headers (uses first)."""
        converter = DocxConverter()

        content = """# First Title

Some content.

# Second Title

More content."""

        title = converter.extract_title_from_markdown(content)
        assert title == "First Title"

    def test_extract_title_from_markdown_empty_content(self):
        """Test title extraction from empty content."""
        converter = DocxConverter()

        content = ""
        title = converter.extract_title_from_markdown(content)
        assert title is None

    def test_extract_title_from_markdown_whitespace_only(self):
        """Test title extraction from whitespace-only content."""
        converter = DocxConverter()

        content = "   \n\n   \n   "
        title = converter.extract_title_from_markdown(content)
        assert title is None

    def test_extract_title_from_markdown_bold_with_formatting(self):
        """Test title extraction from bold text with additional formatting."""
        converter = DocxConverter()

        content = """**Policy Document - Version 2.1**

This document describes...

## Overview

Content here."""

        title = converter.extract_title_from_markdown(content)
        assert title == "Policy Document - Version 2.1"

    def test_extract_title_from_markdown_h1_with_whitespace(self):
        """Test title extraction from H1 with extra whitespace."""
        converter = DocxConverter()

        content = """#    Document Title   

Some content here."""

        title = converter.extract_title_from_markdown(content)
        assert title == "Document Title"
        assert title == "Document Title"
