"""Build training export from BSC.xlsx + Datamart. Does not modify inputs."""

import sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parents[3]
_SHARED = _ROOT / "shared" / "python"
if str(_SHARED) not in sys.path:
    sys.path.insert(0, str(_SHARED))

from training_processor import main_for_source

if __name__ == "__main__":
    main_for_source("BSC.xlsx", "BSC_training_output.xlsx")
