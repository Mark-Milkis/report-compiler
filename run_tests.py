#!/usr/bin/env python3
"""
Test runner script for the report compiler.
"""

import sys
import unittest
from pathlib import Path

# Add the project root to the Python path
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

from tests.test_config import discover_and_run_tests


def main():
    """Main test runner function."""
    print("Running Report Compiler Test Suite")
    print("=" * 50)
    
    success = discover_and_run_tests()
    
    if success:
        print("\n" + "=" * 50)
        print("✅ All tests passed!")
        sys.exit(0)
    else:
        print("\n" + "=" * 50)
        print("❌ Some tests failed!")
        sys.exit(1)


if __name__ == '__main__':
    main()
