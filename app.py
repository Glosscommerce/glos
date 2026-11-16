"""CORRECTGLOSS V3 entry point for the isolated V3 test branch.

This file delegates to v3_app.py. The production V2 main branch is untouched.
"""

import runpy
from pathlib import Path

runpy.run_path(str(Path(__file__).with_name("v3_app.py")), run_name="__main__")
