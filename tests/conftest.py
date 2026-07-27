import os
import sys

# Make the repo root AND the scripts/ dir importable so both
# `import scripts...` and `import analyze_fe_declarations` / `import build_model` work.
_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, _ROOT)
sys.path.insert(0, os.path.join(_ROOT, "scripts"))
