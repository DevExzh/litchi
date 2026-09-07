#!/usr/bin/env python3
import json,subprocess,sys
from pathlib import Path
ROOT=Path(__file__).resolve().parent
protocol=json.loads((ROOT/'protocol.json').read_text())
for pilot,indices in [(True,protocol['pilot_lanes']),(False,range(24))]:
 for i in indices:
  args=[sys.executable,'-B',str(ROOT/'capture.py'),'--lane',str(i)]+(['--pilot'] if pilot else [])
  tag=('pilot-' if pilot else 'formal-')+str(i)
  subprocess.run([sys.executable,'-B',str(ROOT/'check.py'),'--tag',tag,'--',*args],check=True)
