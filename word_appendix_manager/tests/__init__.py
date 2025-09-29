# tests/__init__.py
"""
Test suite for Word Appendix Manager
"""

import sys
import os
from pathlib import Path

# Add src directory to Python path for imports
PROJECT_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(PROJECT_ROOT / "src"))

# Test configuration
TEST_DATA_DIR = Path(__file__).parent / "fixtures"
TEST_OUTPUT_DIR = Path(__file__).parent / "output"

# Ensure test directories exist
TEST_OUTPUT_DIR.mkdir(exist_ok=True) 