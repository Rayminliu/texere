"""Validate CLI tests — 9-item validation coverage.

pytest -q tests/test_validate.py
"""

import json
import os
import shutil
import subprocess
import sys

KIT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _build_sample_docx(tmp_path):
    """Build a sample docx for testing."""
    # Copy sample.md and config
    src = tmp_path / "src"
    src.mkdir()
    shutil.copy(os.path.join(KIT, "assets", "sample.md"), src / "01_sample.md")

    # Render to docx
    out = str(tmp_path / "test.docx")
    subprocess.run(
        [
            sys.executable,
            os.path.join(KIT, "scripts", "render.py"),
            "--src",
            str(src),
            "--out",
            out,
            "--config",
            os.path.join(KIT, "assets", "sample_config.json"),
            "--pdf",
        ],
        check=True,
        capture_output=True,
    )
    return out


class TestPackageIntegrity:
    """Test package integrity checks."""

    def test_valid_package(self, tmp_path):
        """Valid DOCX should pass."""
        docx = _build_sample_docx(tmp_path)

        result = subprocess.run(
            [
                sys.executable,
                os.path.join(KIT, "scripts", "validate.py"),
                docx,
            ],  # Don't use --quiet for this test
            capture_output=True,
            text=True,
        )

        assert result.returncode == 0
        assert "package_integrity" in result.stdout or "package_integrity" in result.stderr

    def test_corrupted_package_exits(self, tmp_path):
        """Corrupted ZIP should fail."""
        # Create a fake ZIP file
        bad_docx = tmp_path / "bad.docx"
        bad_docx.write_bytes(b"PK\x03\x04invalid")

        result = subprocess.run(
            [
                sys.executable,
                os.path.join(KIT, "scripts", "validate.py"),
                str(bad_docx),
            ],
            capture_output=True,
            text=True,
        )

        assert result.returncode != 0
        assert "FAIL" in result.stdout or "error" in result.stderr.lower()


class TestImageEmbedding:
    """Test image embedding checks."""

    def test_images_embedded(self, tmp_path):
        """Sample has no images but should report correctly."""
        docx = _build_sample_docx(tmp_path)

        result = subprocess.run(
            [
                sys.executable,
                os.path.join(KIT, "scripts", "validate.py"),
                docx,
                "--quiet",
            ],
            capture_output=True,
            text=True,
        )

        assert "image_embedding" in result.stdout
        assert "PASS" in result.stdout


class TestSectionCount:
    """Test section count validation."""

    def test_reasonable_sections(self, tmp_path):
        """Sample document should have reasonable sections."""
        docx = _build_sample_docx(tmp_path)

        result = subprocess.run(
            [
                sys.executable,
                os.path.join(KIT, "scripts", "validate.py"),
                docx,
                "--quiet",
            ],
            capture_output=True,
            text=True,
        )

        assert "section_count" in result.stdout
        assert "PASS" in result.stdout


class TestTOCField:
    """Test TOC field validation."""

    def test_toc_exists(self, tmp_path):
        """Sample should have TOC."""
        docx = _build_sample_docx(tmp_path)

        result = subprocess.run(
            [
                sys.executable,
                os.path.join(KIT, "scripts", "validate.py"),
                docx,
                "--quiet",
            ],
            capture_output=True,
            text=True,
        )

        # Should either pass (TOC exists) or skip (no TOC detected)
        assert "toc_field" in result.stdout


class TestPageNumbering:
    """Test page numbering validation."""

    def test_continuous_pages(self, tmp_path):
        """Pages should be continuous."""
        docx = _build_sample_docx(tmp_path)

        result = subprocess.run(
            [
                sys.executable,
                os.path.join(KIT, "scripts", "validate.py"),
                docx,
                "--quiet",
            ],
            capture_output=True,
            text=True,
        )

        assert "page_numbering" in result.stdout
        assert "PASS" in result.stdout


class TestBlankPages:
    """Test blank page detection."""

    def test_no_blank_pages(self, tmp_path):
        """Sample should have no blank pages."""
        docx = _build_sample_docx(tmp_path)

        result = subprocess.run(
            [
                sys.executable,
                os.path.join(KIT, "scripts", "validate.py"),
                docx,
                "--quiet",
            ],
            capture_output=True,
            text=True,
        )

        assert "blank_pages" in result.stdout
        # Sample has 4 pages, should pass with max_empty=0
        assert "PASS" in result.stdout

    def test_max_empty_relaxes_threshold(self, tmp_path):
        """--max-empty should relax the threshold."""
        docx = _build_sample_docx(tmp_path)

        # This should work even if there are blank pages
        result = subprocess.run(
            [
                sys.executable,
                os.path.join(KIT, "scripts", "validate.py"),
                docx,
                "--max-empty",
                "5",
                "--quiet",
            ],
            capture_output=True,
            text=True,
        )

        assert result.returncode == 0


class TestWordAcceptance:
    """Test Word acceptance validation."""

    def test_word_opens_and_exports(self, tmp_path):
        """Document should open in Word and export PDF."""
        docx = _build_sample_docx(tmp_path)

        result = subprocess.run(
            [
                sys.executable,
                os.path.join(KIT, "scripts", "validate.py"),
                docx,
                "--quiet",
            ],
            capture_output=True,
            text=True,
        )

        assert "word_acceptance" in result.stdout
        assert "PASS" in result.stdout


class TestVisualDrift:
    """Test visual drift detection."""

    def test_drift_skipped_without_baseline(self, tmp_path):
        """Should skip drift check without baseline."""
        docx = _build_sample_docx(tmp_path)

        result = subprocess.run(
            [
                sys.executable,
                os.path.join(KIT, "scripts", "validate.py"),
                docx,
                "--quiet",
            ],
            capture_output=True,
            text=True,
        )

        assert "visual_drift" in result.stdout
        # Should PASS with skip message
        assert "PASS" in result.stdout or "跳过" in result.stdout

    def test_drift_with_profile(self, tmp_path):
        """Should use profile if provided."""
        docx = _build_sample_docx(tmp_path)
        profile = os.path.join(KIT, "profiles", "formal-cn-v1.json")

        result = subprocess.run(
            [
                sys.executable,
                os.path.join(KIT, "scripts", "validate.py"),
                docx,
                "--profile",
                profile,
                "--quiet",
            ],
            capture_output=True,
            text=True,
        )

        assert result.returncode == 0


class TestEvidencePackage:
    """Test evidence package generation."""

    def test_evidence_directory_created(self, tmp_path):
        """Evidence directory should be created."""
        docx = _build_sample_docx(tmp_path)
        evidence_dir = tmp_path / "evidence"

        result = subprocess.run(
            [
                sys.executable,
                os.path.join(KIT, "scripts", "validate.py"),
                docx,
                "--out",
                str(evidence_dir),
                "--quiet",
            ],
            capture_output=True,
            text=True,
        )

        assert result.returncode == 0
        assert evidence_dir.exists()
        assert any(evidence_dir.iterdir())

    def test_report_json_created(self, tmp_path):
        """report.json should be created."""
        docx = _build_sample_docx(tmp_path)
        evidence_dir = tmp_path / "evidence"

        subprocess.run(
            [
                sys.executable,
                os.path.join(KIT, "scripts", "validate.py"),
                docx,
                "--out",
                str(evidence_dir),
                "--quiet",
            ],
            check=True,
        )

        report_path = evidence_dir / "report.json"
        assert report_path.exists()

        # Verify JSON structure
        report = json.load(open(report_path, encoding="utf-8"))
        assert "metadata" in report
        assert "checks" in report
        assert "summary" in report
        assert "total" in report["summary"]
        assert "passed" in report["summary"]
        assert "failed" in report["summary"]

    def test_signature_created(self, tmp_path):
        """Signature file should be created."""
        docx = _build_sample_docx(tmp_path)
        evidence_dir = tmp_path / "evidence"

        subprocess.run(
            [
                sys.executable,
                os.path.join(KIT, "scripts", "validate.py"),
                docx,
                "--out",
                str(evidence_dir),
                "--quiet",
            ],
            check=True,
        )

        sig_path = evidence_dir / "signature"
        assert sig_path.exists()

        # Verify signature content
        content = sig_path.read_text(encoding="utf-8")
        assert "document_hash" in content
        assert "checks_passed" in content
        # report.json 的 hash 也要进证据清单：否则改报告结论不会破坏证据
        assert "report_hash" in content
        assert "cryptographic" in content  # 明说是 checksum manifest，不是签名

    def test_screenshot_created(self, tmp_path):
        """Screenshot PNGs should be created."""
        docx = _build_sample_docx(tmp_path)
        evidence_dir = tmp_path / "evidence"

        subprocess.run(
            [
                sys.executable,
                os.path.join(KIT, "scripts", "validate.py"),
                docx,
                "--out",
                str(evidence_dir),
                "--quiet",
            ],
            check=True,
        )

        # Should have at least one screenshot
        png_files = list(evidence_dir.glob("page-*.png"))
        assert len(png_files) >= 1


class TestExitCodes:
    """Test exit code behavior."""

    def test_success_exit_code(self, tmp_path):
        """Successful validation should exit 0."""
        docx = _build_sample_docx(tmp_path)

        result = subprocess.run(
            [
                sys.executable,
                os.path.join(KIT, "scripts", "validate.py"),
                docx,
                "--quiet",
            ],
            capture_output=True,
        )

        assert result.returncode == 0

    def test_failure_exit_code(self, tmp_path):
        """Failed validation should exit 1."""
        # Use invalid file
        bad_file = tmp_path / "notexist.docx"

        result = subprocess.run(
            [
                sys.executable,
                os.path.join(KIT, "scripts", "validate.py"),
                str(bad_file),
            ],
            capture_output=True,
        )

        assert result.returncode != 0


class TestHelpAndVersion:
    """Test CLI help and version."""

    def test_help_flag(self):
        """--help should show usage."""
        result = subprocess.run(
            [sys.executable, os.path.join(KIT, "scripts", "validate.py"), "--help"],
            capture_output=True,
            text=True,
        )

        assert result.returncode == 0
        assert "usage:" in result.stdout.lower()
        assert "validate.py" in result.stdout

    def test_positional_argument_required(self, tmp_path):
        """Missing positional argument should exit with error."""
        result = subprocess.run(
            [sys.executable, os.path.join(KIT, "scripts", "validate.py")],
            capture_output=True,
            text=True,
        )

        assert result.returncode != 0


class TestProfileValidation:
    """Test profile integration."""

    def test_formal_cn_profile(self, tmp_path):
        """Default formal-cn profile should work."""
        docx = _build_sample_docx(tmp_path)
        profile = os.path.join(KIT, "profiles", "formal-cn-v1.json")

        result = subprocess.run(
            [
                sys.executable,
                os.path.join(KIT, "scripts", "validate.py"),
                docx,
                "--profile",
                profile,
                "--quiet",
            ],
            capture_output=True,
            text=True,
        )

        assert result.returncode == 0

    def test_tender_profile(self, tmp_path):
        """Tender profile should work."""
        docx = _build_sample_docx(tmp_path)
        profile = os.path.join(KIT, "profiles", "tender-v1.json")

        result = subprocess.run(
            [
                sys.executable,
                os.path.join(KIT, "scripts", "validate.py"),
                docx,
                "--profile",
                profile,
                "--quiet",
            ],
            capture_output=True,
            text=True,
        )

        assert result.returncode == 0

    def test_gongwen_profile(self, tmp_path):
        """Gongwen profile should work."""
        docx = _build_sample_docx(tmp_path)
        profile = os.path.join(KIT, "profiles", "gongwen-v1.json")

        result = subprocess.run(
            [
                sys.executable,
                os.path.join(KIT, "scripts", "validate.py"),
                docx,
                "--profile",
                profile,
                "--quiet",
            ],
            capture_output=True,
            text=True,
        )

        assert result.returncode == 0

    def test_application_profile(self, tmp_path):
        """Application profile should work."""
        docx = _build_sample_docx(tmp_path)
        profile = os.path.join(KIT, "profiles", "application-v1.json")

        result = subprocess.run(
            [
                sys.executable,
                os.path.join(KIT, "scripts", "validate.py"),
                docx,
                "--profile",
                profile,
                "--quiet",
            ],
            capture_output=True,
            text=True,
        )

        assert result.returncode == 0


class TestReportStructure:
    """Test report JSON structure."""

    def test_report_has_all_checks(self, tmp_path):
        """Report should contain all 9 checks."""
        docx = _build_sample_docx(tmp_path)
        evidence_dir = tmp_path / "evidence"

        subprocess.run(
            [
                sys.executable,
                os.path.join(KIT, "scripts", "validate.py"),
                docx,
                "--out",
                str(evidence_dir),
                "--quiet",
            ],
            check=True,
        )

        report = json.load(open(evidence_dir / "report.json", encoding="utf-8"))

        expected_checks = [
            "package_integrity",
            "source_content",
            "image_embedding",
            "section_count",
            "toc_field",
            "page_numbering",
            "blank_pages",
            "word_acceptance",
            "visual_drift",
        ]

        for check_name in expected_checks:
            assert check_name in report["checks"], f"Missing check: {check_name}"

    def test_report_timestamp(self, tmp_path):
        """Report should have timestamp."""
        docx = _build_sample_docx(tmp_path)
        evidence_dir = tmp_path / "evidence"

        subprocess.run(
            [
                sys.executable,
                os.path.join(KIT, "scripts", "validate.py"),
                docx,
                "--out",
                str(evidence_dir),
                "--quiet",
            ],
            check=True,
        )

        report = json.load(open(evidence_dir / "report.json", encoding="utf-8"))

        assert "timestamp" in report["metadata"]
        # Should be ISO format
        assert "T" in report["metadata"]["timestamp"]

    def test_report_tool_version(self, tmp_path):
        """Report should have tool version."""
        docx = _build_sample_docx(tmp_path)
        evidence_dir = tmp_path / "evidence"

        subprocess.run(
            [
                sys.executable,
                os.path.join(KIT, "scripts", "validate.py"),
                docx,
                "--out",
                str(evidence_dir),
                "--quiet",
            ],
            check=True,
        )

        report = json.load(open(evidence_dir / "report.json", encoding="utf-8"))

        assert "tool_version" in report["metadata"]
        # 版本从 scripts/_version.py 单一来源读取，发版升号时不必再改本测试
        import re

        with open(os.path.join(KIT, "scripts", "_version.py"), encoding="utf-8") as f:
            m = re.search(r'__version__\s*=\s*"([^"]+)"', f.read())
        assert m, "scripts/_version.py 里找不到 __version__"
        assert report["metadata"]["tool_version"] == m.group(1)
