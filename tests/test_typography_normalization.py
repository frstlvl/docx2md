"""Tests for smart typographic character normalization."""

import pytest
from pathlib import Path
from docx2md import DocxConverter


@pytest.fixture
def converter():
    """Create a DocxConverter with default settings (normalization on)."""
    return DocxConverter()


@pytest.fixture
def converter_keep_smart():
    """Create a DocxConverter with normalization disabled."""
    return DocxConverter(normalize_typography=False)


class TestNormalizeTypography:
    """Test _normalize_typography replaces smart chars with ASCII."""

    def test_left_double_quote(self, converter):
        assert converter._normalize_typography("\u201cHello\u201d") == '"Hello"'

    def test_right_double_quote(self, converter):
        assert converter._normalize_typography("said \u201cyes\u201d") == 'said "yes"'

    def test_left_single_quote(self, converter):
        assert converter._normalize_typography("\u2018word\u2019") == "'word'"

    def test_right_single_quote_apostrophe(self, converter):
        assert converter._normalize_typography("don\u2019t") == "don't"

    def test_en_dash(self, converter):
        assert converter._normalize_typography("pages 1\u20135") == "pages 1-5"

    def test_em_dash(self, converter):
        assert converter._normalize_typography("word\u2014another") == "word--another"

    def test_ellipsis(self, converter):
        assert converter._normalize_typography("wait\u2026") == "wait..."

    def test_left_guillemet(self, converter):
        assert converter._normalize_typography("\u00abquote\u00bb") == '"quote"'

    def test_prime(self, converter):
        assert converter._normalize_typography("5\u2032") == "5'"

    def test_double_prime(self, converter):
        assert converter._normalize_typography('5\u2033') == '5"'

    def test_modifier_letter_apostrophe(self, converter):
        assert converter._normalize_typography("rock\u02bcn\u02bcroll") == "rock'n'roll"

    def test_mixed_smart_chars(self, converter):
        text = "\u201cHe said, \u2018don\u2019t worry\u2019\u201d \u2014 it\u2019s fine\u2026"
        expected = "\"He said, 'don't worry'\" -- it's fine..."
        assert converter._normalize_typography(text) == expected

    def test_normal_ascii_unchanged(self, converter):
        text = 'Normal "quoted" text with \'apostrophes\' and - dashes.'
        assert converter._normalize_typography(text) == text

    def test_empty_string(self, converter):
        assert converter._normalize_typography("") == ""

    def test_no_smart_chars(self, converter):
        text = "# Heading\n\nSome regular markdown content.\n\n- List item\n"
        assert converter._normalize_typography(text) == text

    def test_smart_chars_in_front_matter(self, converter):
        text = '---\ntitle: \u201cMy Document\u201d\n---\n\nContent here.\n'
        expected = '---\ntitle: "My Document"\n---\n\nContent here.\n'
        assert converter._normalize_typography(text) == expected

    def test_multiple_occurrences(self, converter):
        text = "\u201cone\u201d and \u201ctwo\u201d and \u201cthree\u201d"
        expected = '"one" and "two" and "three"'
        assert converter._normalize_typography(text) == expected


class TestNormalizeTypographyOptOut:
    """Test that --keep-smart-chars preserves smart characters."""

    def test_smart_chars_preserved_when_disabled(self, converter_keep_smart):
        """When normalize_typography=False, the method is not called by the pipeline.
        Verify the flag is correctly set so apply_markdown_linting_rules skips it."""
        assert converter_keep_smart.normalize_typography is False

    def test_normalize_typography_flag_default_true(self):
        converter = DocxConverter()
        assert converter.normalize_typography is True

    def test_normalize_typography_flag_false(self):
        converter = DocxConverter(normalize_typography=False)
        assert converter.normalize_typography is False


class TestNormalizeTypographyIntegration:
    """Test that normalization is wired into the linting pipeline."""

    def test_linting_applies_normalization(self, converter, tmp_path):
        """Verify apply_markdown_linting_rules normalizes typography."""
        md_file = tmp_path / "test.md"
        md_file.write_text(
            "\u201cSmart quotes\u201d and \u2014 dashes\n",
            encoding="utf-8",
        )
        converter.apply_markdown_linting_rules(md_file)
        result = md_file.read_text(encoding="utf-8")
        assert "\u201c" not in result
        assert "\u201d" not in result
        assert "\u2014" not in result
        assert '"Smart quotes"' in result
        assert "-- dashes" in result

    def test_linting_skips_normalization_when_disabled(
        self, converter_keep_smart, tmp_path
    ):
        """Verify normalization is skipped when normalize_typography=False."""
        md_file = tmp_path / "test.md"
        original = "\u201cSmart quotes\u201d and \u2014 dashes\n"
        md_file.write_text(original, encoding="utf-8")
        converter_keep_smart.apply_markdown_linting_rules(md_file)
        result = md_file.read_text(encoding="utf-8")
        assert "\u201c" in result
        assert "\u201d" in result
        assert "\u2014" in result
