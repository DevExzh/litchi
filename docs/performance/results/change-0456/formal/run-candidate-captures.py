#!/usr/bin/env python3
"""Complete the frozen matrix serially after the initial six baseline lanes."""
import subprocess,sys
from pathlib import Path
ROOT=Path(__file__).resolve().parent
for lane in [*range(4,16),*range(18,24)]:
    subprocess.run([sys.executable,'-B',str(ROOT/'capture.py'),str(lane)],check=True)
