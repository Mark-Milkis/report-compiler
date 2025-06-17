#!/usr/bin/env python3
"""Direct test of compiler module."""

import sys
import traceback

try:
    print("Testing direct import of compiler module...")
    
    # Add current directory to path
    sys.path.insert(0, '.')
    
    # Import the module directly
    exec(open('report_compiler/core/compiler.py').read())
    
    print("✓ Module executed successfully")
    print("Available names:", [name for name in locals() if not name.startswith('_')])
    
except Exception as e:
    print(f"❌ Error executing compiler module: {e}")
    traceback.print_exc()
