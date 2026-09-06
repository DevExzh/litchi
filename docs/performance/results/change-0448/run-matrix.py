#!/usr/bin/env python3
"""Run every frozen process serially and stop on the first failure."""
import json,subprocess,sys
from pathlib import Path
ROOT=Path(__file__).resolve().parent
for index in range(8):
    subprocess.run([sys.executable,'-B',str(ROOT/'check.py'),'--tag','formal-'+str(index),'--',sys.executable,'-B',str(ROOT/'capture.py'),'--lane',str(index)],check=True)
for row in json.loads((ROOT/'protocol.json').read_text())['profile_order']:
    subprocess.run([sys.executable,'-B',str(ROOT/'check.py'),'--tag',f"profile-{row['lane']}-{row['kind']}",'--',sys.executable,'-B',str(ROOT/'capture.py'),'--lane',str(row['lane']),'--profile',row['kind']],check=True)
