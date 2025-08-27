"""Tests for file filtering functionality."""

from pathlib import Path

from docx2md import DocxConverter


class TestFileFiltering:
    """Test file filtering for temporary and lock files."""

    def test_is_temporary_file_tilde_prefix(self):
        """Test detection of temporary files with tilde prefix."""
        converter = DocxConverter()

        # Temporary files start with ~$
        assert converter.is_temporary_file(Path("~$Document.docx"))
        assert converter.is_temporary_file(Path("~$My File.docx"))
        assert converter.is_temporary_file(Path("folder/~$Document.docx"))

    def test_is_temporary_file_lock_files(self):
        """Test detection of lock files."""
        converter = DocxConverter()

        # Lock files start with .~lock
        assert converter.is_temporary_file(Path(".~lock.Document.docx#"))
        assert converter.is_temporary_file(Path("folder/.~lock.My File.docx#"))

    def test_is_temporary_file_normal_files(self):
        """Test that normal files are not flagged as temporary."""
        converter = DocxConverter()

        # Normal files should not be flagged
        assert not converter.is_temporary_file(Path("Document.docx"))
        assert not converter.is_temporary_file(Path("My Important File.docx"))
        assert not converter.is_temporary_file(Path("folder/subfolder/Report.docx"))
        assert not converter.is_temporary_file(Path("~Document.txt"))  # Wrong extension

    def test_is_temporary_file_edge_cases(self):
        """Test edge cases for temporary file detection."""
        converter = DocxConverter()

        # Files that look similar but aren't temporary
        assert not converter.is_temporary_file(Path("File~.docx"))  # Tilde at end
        assert not converter.is_temporary_file(Path("~File.docx"))  # Single tilde
        assert not converter.is_temporary_file(Path("lock.docx"))  # No prefix

    def test_discover_docx_files_filters_temporary(self, tmp_path):
        """Test that file discovery filters out temporary files."""
        converter = DocxConverter()

        # Create various files
        (tmp_path / "normal.docx").touch()
        (tmp_path / "~$temp.docx").touch()  # Temporary
        (tmp_path / ".~lock.document.docx#").touch()  # Lock file
        (tmp_path / "another.docx").touch()

        files = converter.discover_docx_files([tmp_path])

        # Should only find normal files
        file_names = {f[0].name for f in files}
        assert file_names == {"normal.docx", "another.docx"}
        assert "~$temp.docx" not in file_names
        assert ".~lock.document.docx#" not in file_names
        assert ".~lock.document.docx#" not in file_names
