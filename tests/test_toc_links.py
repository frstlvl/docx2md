"""Tests for TOC link fixing functionality."""

import pytest

from docx2md import DocxConverter


class TestTOCLinkFixing:
    """Test TOC link fixing functionality."""

    def test_create_heading_anchor_basic(self):
        """Test basic heading anchor creation."""
        converter = DocxConverter()

        assert converter._create_heading_anchor("Introduction") == "introduction"
        assert converter._create_heading_anchor("Main Section") == "main-section"
        assert (
            converter._create_heading_anchor("Multiple   Spaces") == "multiple-spaces"
        )

    def test_create_heading_anchor_swedish_chars(self):
        """Test anchor creation with Swedish characters."""
        converter = DocxConverter()

        assert converter._create_heading_anchor("Inledning") == "inledning"
        assert (
            converter._create_heading_anchor("Förvaring och lagring")
            == "förvaring-och-lagring"
        )
        assert converter._create_heading_anchor("Säkerhet") == "säkerhet"

    def test_create_heading_anchor_special_chars(self):
        """Test anchor creation with special characters."""
        converter = DocxConverter()

        assert converter._create_heading_anchor("Section 1.1") == "section-11"
        assert converter._create_heading_anchor("Q&A") == "qa"
        assert converter._create_heading_anchor("Cost/Benefit") == "costbenefit"

    def test_create_heading_anchor_formatting(self):
        """Test anchor creation with markdown formatting."""
        converter = DocxConverter()

        assert converter._create_heading_anchor("**Bold Text**") == "bold-text"
        assert converter._create_heading_anchor("*Italic Text*") == "italic-text"
        assert converter._create_heading_anchor("`Code Text`") == "code-text"

    def test_fix_toc_links_basic(self):
        """Test basic TOC link fixing."""
        converter = DocxConverter()

        content = """# Introduction

[1. Introduction](#_Toc123456789)
[2. Main Section](#_Toc987654321)

# Main Section

Some content here."""

        result = converter._fix_toc_links(content)

        assert "[1. Introduction](#introduction)" in result
        assert "[2. Main Section](#main-section)" in result
        assert "#_Toc" not in result

    def test_fix_toc_links_with_page_numbers(self):
        """Test TOC link fixing with page numbers."""
        converter = DocxConverter()

        content = """# Introduction

[1. Introduction 5](#_Toc123456789)
[2. Main Section 12](#_Toc987654321)

# Introduction

# Main Section"""

        result = converter._fix_toc_links(content)

        assert "[1. Introduction 5](#introduction)" in result
        assert "[2. Main Section 12](#main-section)" in result

    def test_fix_toc_links_complex_headings(self):
        """Test TOC link fixing with complex heading text."""
        converter = DocxConverter()

        content = """# Förvaring och lagring

[3.1.3 Förvaring och lagring 4](#_Toc180757069)

# Förvaring och lagring

Some content."""

        result = converter._fix_toc_links(content)

        assert "[3.1.3 Förvaring och lagring 4](#förvaring-och-lagring)" in result

    def test_fix_toc_links_no_match(self):
        """Test TOC link fixing when no heading match is found."""
        converter = DocxConverter()

        content = """# Introduction

[Nonexistent Section](#_Toc123456789)

# Introduction"""

        result = converter._fix_toc_links(content)

        # Should remove the link but keep the text
        assert "Nonexistent Section" in result
        assert "#_Toc" not in result
        assert "[Nonexistent Section]" not in result
        assert "[Nonexistent Section]" not in result
