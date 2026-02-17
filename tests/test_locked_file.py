"""Tests for locked file detection and warning."""

import pytest
from pathlib import Path
from unittest.mock import patch, MagicMock
from docx2md import DocxConverter


@pytest.fixture
def converter(tmp_path):
    """Create a DocxConverter with output to tmp_path."""
    return DocxConverter(output_dir=tmp_path, overwrite=True)


class TestLockedFileDetection:
    """Test that locked/inaccessible DOCX files produce clear warnings."""

    def test_locked_file_skipped_with_warning(self, converter, tmp_path):
        """A file that raises PermissionError should be skipped with clear error."""
        fake_docx = tmp_path / "locked.docx"
        fake_docx.touch()

        # Only mock the specific open call for the file check, not all open calls
        original_open = open

        def mock_open_locked(*args, **kwargs):
            # Only raise PermissionError for rb mode (the lock check)
            if len(args) >= 2 and args[1] == "rb":
                raise PermissionError("Access denied")
            return original_open(*args, **kwargs)

        with patch("builtins.open", side_effect=mock_open_locked):
            result = converter.convert_single_file(fake_docx)

        assert result is False
        assert converter.stats["failed"] == 1
        assert len(converter.stats["failed_files"]) == 1
        assert "open in another application" in converter.stats["failed_files"][0]["error"]

    def test_locked_file_error_message_mentions_word(self, converter, tmp_path):
        """Error message should mention Word as a common cause."""
        fake_docx = tmp_path / "locked.docx"
        fake_docx.touch()

        original_open = open

        def mock_open_locked(*args, **kwargs):
            if len(args) >= 2 and args[1] == "rb":
                raise PermissionError("Access denied")
            return original_open(*args, **kwargs)

        with patch("builtins.open", side_effect=mock_open_locked):
            converter.convert_single_file(fake_docx)

        error_msg = converter.stats["failed_files"][0]["error"]
        assert "Word" in error_msg

    def test_locked_file_filename_recorded(self, converter, tmp_path):
        """The failed file's name should be recorded in stats."""
        fake_docx = tmp_path / "important_document.docx"
        fake_docx.touch()

        original_open = open

        def mock_open_locked(*args, **kwargs):
            if len(args) >= 2 and args[1] == "rb":
                raise PermissionError("Access denied")
            return original_open(*args, **kwargs)

        with patch("builtins.open", side_effect=mock_open_locked):
            converter.convert_single_file(fake_docx)

        assert converter.stats["failed_files"][0]["file"] == "important_document.docx"

    def test_accessible_file_not_blocked(self, converter, tmp_path):
        """A normal accessible file should not be blocked by the permission check."""
        import zipfile
        fake_docx = tmp_path / "accessible.docx"
        with zipfile.ZipFile(fake_docx, "w") as zf:
            zf.writestr("word/document.xml", "<w:document/>")

        # The conversion will fail (not a valid docx) but should NOT fail
        # at the permission check stage
        result = converter.convert_single_file(fake_docx)
        if not result and converter.stats["failed_files"]:
            assert "open in another application" not in converter.stats["failed_files"][0]["error"]

    def test_stats_initialized_with_failed_files(self):
        """Stats dict should include failed_files list from initialization."""
        converter = DocxConverter()
        assert "failed_files" in converter.stats
        assert converter.stats["failed_files"] == []


class TestPermissionErrorInSubMethods:
    """Test PermissionError handling in specific conversion methods."""

    def test_extract_core_properties_permission_error(self, tmp_path):
        """extract_core_properties should handle PermissionError gracefully."""
        converter = DocxConverter()
        fake_docx = tmp_path / "locked.docx"
        fake_docx.touch()

        with patch("zipfile.ZipFile", side_effect=PermissionError("Access denied")):
            props = converter.extract_core_properties(fake_docx)

        # Should return partial properties (at least source_file)
        assert "source_file" in props
        assert props["source_file"] == "locked.docx"

    def test_mammoth_permission_error(self, tmp_path):
        """convert_with_mammoth should return False on PermissionError."""
        converter = DocxConverter()
        fake_docx = tmp_path / "locked.docx"
        fake_docx.touch()
        output = tmp_path / "output.md"
        media = tmp_path / "media"

        with patch("builtins.open", side_effect=PermissionError("Access denied")):
            result = converter.convert_with_mammoth(fake_docx, output, media)

        assert result is False
