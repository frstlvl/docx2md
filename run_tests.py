#!/usr/bin/env python3
"""
Test runner script for docx2md project.

This script installs test dependencies and runs the full test suite.
"""

import subprocess
import sys
from pathlib import Path


def run_command(cmd: list[str]) -> bool:
    """Run a command and return True if successful."""
    try:
        result = subprocess.run(cmd, check=True, capture_output=True, text=True)
        print(f"✓ {' '.join(cmd)}")
        return True
    except subprocess.CalledProcessError as e:
        print(f"✗ {' '.join(cmd)}")
        print(f"Error: {e.stderr}")
        return False


def main():
    """Main test runner."""
    print("🧪 DocX2MD Test Suite")
    print("=" * 50)

    # Change to project root
    project_root = Path(__file__).parent
    subprocess.run(["cd", str(project_root)], shell=True)

    # Install test dependencies
    print("\n📦 Installing test dependencies...")
    if not run_command(["uv", "pip", "install", "-e", ".[test]"]):
        print("❌ Failed to install test dependencies")
        return 1

    # Run tests
    print("\n🔍 Running tests...")
    if not run_command(["uv", "run", "python", "-m", "pytest", "tests/", "-v"]):
        print("❌ Tests failed")
        return 1

    # Run tests with coverage
    print("\n📊 Running tests with coverage...")
    if not run_command(
        [
            "uv",
            "run",
            "python",
            "-m",
            "pytest",
            "tests/",
            "--cov=docx2md",
            "--cov-report=term-missing",
        ]
    ):
        print("❌ Coverage analysis failed")
        return 1

    print("\n✅ All tests passed!")
    return 0


if __name__ == "__main__":
    sys.exit(main())
