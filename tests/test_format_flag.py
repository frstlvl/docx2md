"""Tests for the --format flag (obsidian/gfm/standard markdown dialect)."""

import pytest
from unittest.mock import patch, MagicMock
from docx2md import DocxConverter


class TestOutputFormatParam:
    """Test output_format parameter initialization and defaults."""

    def test_default_format_is_obsidian(self):
        converter = DocxConverter()
        assert converter.output_format == "obsidian"

    def test_format_gfm(self):
        converter = DocxConverter(output_format="gfm")
        assert converter.output_format == "gfm"

    def test_format_standard(self):
        converter = DocxConverter(output_format="standard")
        assert converter.output_format == "standard"

    def test_format_obsidian_explicit(self):
        converter = DocxConverter(output_format="obsidian")
        assert converter.output_format == "obsidian"


class TestPandocTargetMapping:
    """Test that output_format maps to correct Pandoc -t argument."""

    def test_obsidian_uses_gfm_target(self):
        converter = DocxConverter(output_format="obsidian")
        assert converter.PANDOC_TARGETS["obsidian"] == "gfm"

    def test_gfm_uses_gfm_target(self):
        converter = DocxConverter(output_format="gfm")
        assert converter.PANDOC_TARGETS["gfm"] == "gfm"

    def test_standard_uses_commonmark_target(self):
        converter = DocxConverter(output_format="standard")
        assert converter.PANDOC_TARGETS["standard"] == "commonmark"

    @patch("shutil.which", return_value="/usr/bin/pandoc")
    @patch("subprocess.run")
    def test_pandoc_called_with_gfm_for_obsidian(self, mock_run, mock_which, tmp_path):
        """Verify Pandoc is called with -t gfm when format is obsidian."""
        converter = DocxConverter(output_format="obsidian")
        mock_run.return_value = MagicMock(stdout="", returncode=0)

        fake_docx = tmp_path / "test.docx"
        fake_docx.touch()
        output = tmp_path / "test.md"
        media = tmp_path / "media"

        converter.convert_with_pandoc(fake_docx, output, media)

        cmd = mock_run.call_args[0][0]
        assert "-t" in cmd
        t_index = cmd.index("-t")
        assert cmd[t_index + 1] == "gfm"

    @patch("shutil.which", return_value="/usr/bin/pandoc")
    @patch("subprocess.run")
    def test_pandoc_called_with_gfm_for_gfm(self, mock_run, mock_which, tmp_path):
        """Verify Pandoc is called with -t gfm when format is gfm."""
        converter = DocxConverter(output_format="gfm")
        mock_run.return_value = MagicMock(stdout="", returncode=0)

        fake_docx = tmp_path / "test.docx"
        fake_docx.touch()
        output = tmp_path / "test.md"
        media = tmp_path / "media"

        converter.convert_with_pandoc(fake_docx, output, media)

        cmd = mock_run.call_args[0][0]
        t_index = cmd.index("-t")
        assert cmd[t_index + 1] == "gfm"

    @patch("shutil.which", return_value="/usr/bin/pandoc")
    @patch("subprocess.run")
    def test_pandoc_called_with_commonmark_for_standard(
        self, mock_run, mock_which, tmp_path
    ):
        """Verify Pandoc is called with -t commonmark when format is standard."""
        converter = DocxConverter(output_format="standard")
        mock_run.return_value = MagicMock(stdout="", returncode=0)

        fake_docx = tmp_path / "test.docx"
        fake_docx.touch()
        output = tmp_path / "test.md"
        media = tmp_path / "media"

        converter.convert_with_pandoc(fake_docx, output, media)

        cmd = mock_run.call_args[0][0]
        t_index = cmd.index("-t")
        assert cmd[t_index + 1] == "commonmark"


class TestFormatNames:
    """Test human-readable format name mapping."""

    def test_obsidian_format_name(self):
        assert DocxConverter.FORMAT_NAMES["obsidian"] == "Obsidian-flavored"

    def test_gfm_format_name(self):
        assert DocxConverter.FORMAT_NAMES["gfm"] == "GitHub Flavored (GFM)"

    def test_standard_format_name(self):
        assert DocxConverter.FORMAT_NAMES["standard"] == "standard (CommonMark)"


class TestFormatSpecificPostProcessing:
    """Test format-specific post-processing hook behavior."""

    def test_obsidian_calls_obsidian_hook(self, tmp_path):
        converter = DocxConverter(output_format="obsidian")
        md_path = tmp_path / "test.md"
        md_path.write_text("Some content\n", encoding="utf-8")

        with patch.object(
            converter,
            "_apply_obsidian_format_postprocessing",
            return_value="Hooked output\n",
        ) as mock_hook:
            converter.apply_markdown_linting_rules(md_path)
            mock_hook.assert_called_once()

        assert md_path.read_text(encoding="utf-8") == "Hooked output\n"

    @pytest.mark.parametrize("output_format", ["gfm", "standard"])
    def test_non_obsidian_skips_obsidian_hook(self, tmp_path, output_format):
        converter = DocxConverter(output_format=output_format)
        md_path = tmp_path / f"test-{output_format}.md"
        md_path.write_text("Some content\n", encoding="utf-8")

        with patch.object(
            converter,
            "_apply_obsidian_format_postprocessing",
            return_value="Should not be used\n",
        ) as mock_hook:
            converter.apply_markdown_linting_rules(md_path)
            mock_hook.assert_not_called()

        assert md_path.read_text(encoding="utf-8") == "Some content\n"
