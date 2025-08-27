"""Tests for docx2md converter."""

import tempfile
import zipfile
from pathlib import Path
from unittest.mock import Mock, patch

import pytest

from docx2md import DocxConverter


class TestDocxConverter:
    """Test cases for DocxConverter class."""

    def test_sanitize_filename(self):
        """Test filename sanitization."""
        converter = DocxConverter()

        # Test space replacement
        assert converter.sanitize_filename("My Document") == "My_Document"

        # Test invalid character removal
        assert converter.sanitize_filename("File<>Name") == "FileName"

        # Test case preservation
        assert converter.sanitize_filename("CamelCase") == "CamelCase"

    def test_create_yaml_front_matter(self):
        """Test YAML front matter generation."""
        converter = DocxConverter()

        properties = {
            "title": "Test Document",
            "source_file": "test.docx",
        }

        yaml = converter.create_yaml_front_matter(properties)

        assert yaml.startswith("---")
        assert yaml.endswith("---\n")
        assert "title: Test Document" in yaml
        assert "source_file: test.docx" in yaml

    def test_create_yaml_front_matter_empty(self):
        """Test YAML front matter with empty properties."""
        converter = DocxConverter()
        yaml = converter.create_yaml_front_matter({})
        assert yaml == ""

    @patch("shutil.which")
    def test_find_pandoc_in_path(self, mock_which):
        """Test finding pandoc in PATH."""
        mock_which.return_value = "/usr/bin/pandoc"
        converter = DocxConverter()

        pandoc_path = converter.find_pandoc()
        assert pandoc_path == Path("/usr/bin/pandoc")
        mock_which.assert_called_once_with("pandoc")

    @patch("shutil.which")
    def test_find_pandoc_not_found(self, mock_which):
        """Test pandoc not found."""
        mock_which.return_value = None
        converter = DocxConverter()

        pandoc_path = converter.find_pandoc()
        assert pandoc_path is None

    def test_discover_docx_files_single_file(self, tmp_path):
        """Test discovering single DOCX file."""
        # Create test file
        test_file = tmp_path / "test.docx"
        test_file.touch()

        converter = DocxConverter()
        files = converter.discover_docx_files([test_file])

        assert len(files) == 1
        assert files[0][0] == test_file
        assert files[0][1] == test_file.parent

    def test_discover_docx_files_directory(self, tmp_path):
        """Test discovering DOCX files in directory."""
        # Create test files
        (tmp_path / "test1.docx").touch()
        (tmp_path / "test2.docx").touch()
        (tmp_path / "other.txt").touch()

        converter = DocxConverter()
        files = converter.discover_docx_files([tmp_path])

        assert len(files) == 2
        assert all(f[0].suffix == ".docx" for f in files)

    def test_discover_docx_files_recursive(self, tmp_path):
        """Test recursive discovery."""
        # Create nested structure
        subdir = tmp_path / "subdir"
        subdir.mkdir()
        (tmp_path / "root.docx").touch()
        (subdir / "nested.docx").touch()

        converter = DocxConverter()
        files = converter.discover_docx_files([tmp_path], recursive=True)

        assert len(files) == 2
        docx_names = {f[0].name for f in files}
        assert docx_names == {"root.docx", "nested.docx"}

    def test_skip_unsupported_formats(self, tmp_path):
        """Test skipping unsupported file formats."""
        (tmp_path / "test.doc").touch()
        (tmp_path / "test.docm").touch()

        converter = DocxConverter()
        files = converter.discover_docx_files([tmp_path])

        assert len(files) == 0


@pytest.fixture
def sample_docx(tmp_path):
    """Create a minimal valid DOCX file for testing."""
    docx_path = tmp_path / "sample.docx"

    # Create minimal DOCX structure
    with zipfile.ZipFile(docx_path, "w") as docx_zip:
        # Add minimal content
        docx_zip.writestr("[Content_Types].xml", '<?xml version="1.0"?><Types/>')
        docx_zip.writestr("_rels/.rels", '<?xml version="1.0"?><Relationships/>')
        docx_zip.writestr(
            "word/document.xml",
            '<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Test content</w:t></w:r></w:p></w:body></w:document>',
        )

        # Add core properties
        core_xml = """<?xml version="1.0"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" 
                   xmlns:dc="http://purl.org/dc/elements/1.1/" 
                   xmlns:dcterms="http://purl.org/dc/terms/">
  <dc:title>Test Document</dc:title>
  <dc:creator>Test Author</dc:creator>
  <dcterms:created>2024-01-01T00:00:00Z</dcterms:created>
  <dcterms:modified>2024-01-02T00:00:00Z</dcterms:modified>
</cp:coreProperties>"""
        docx_zip.writestr("docProps/core.xml", core_xml)

    return docx_path


class TestIntegration:
    """Integration tests."""

    def test_extract_core_properties(self, sample_docx):
        """Test extracting core properties from DOCX."""
        converter = DocxConverter()
        properties = converter.extract_core_properties(sample_docx)

        assert properties["title"] == "Test Document"
        assert properties["author"] == "Test Author"
        assert properties["created"] == "2024-01-01T00:00:00Z"
        assert properties["modified"] == "2024-01-02T00:00:00Z"
        assert properties["source_file"] == "sample.docx"

    @patch("mammoth.convert_to_html")
    def test_convert_with_mammoth(self, mock_mammoth, sample_docx, tmp_path):
        """Test conversion using Mammoth."""
        # Mock mammoth conversion
        mock_result = Mock()
        mock_result.value = "<h1>Test</h1><p>Content</p>"
        mock_result.messages = []
        mock_mammoth.return_value = mock_result

        converter = DocxConverter()
        output_path = tmp_path / "output.md"
        media_base = tmp_path / "media"

        success = converter.convert_with_mammoth(sample_docx, output_path, media_base)

        assert success
        assert output_path.exists()

        # Check content was converted
        content = output_path.read_text()
        assert "Test" in content
        assert "Content" in content


if __name__ == "__main__":
    pytest.main([__file__])
