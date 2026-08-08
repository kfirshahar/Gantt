"""Generate Gantt_Template.xlsx."""

import sys
from pathlib import Path

from gantt.build import build

if __name__ == "__main__":
    out = Path(sys.argv[1] if len(sys.argv) > 1 else "Gantt_Template.xlsx")
    print(f"wrote {build(out)}")
