#!/usr/bin/env python3
"""All processes run serially; stop on any workload or independent-oracle failure."""
import argparse,json,subprocess,sys
from pathlib import Path
ROOT=Path(__file__).resolve().parent
p=argparse.ArgumentParser();p.add_argument('--phase',choices=['pilots','formal'],required=True);a=p.parse_args()
for lane in (range(8) if a.phase=='pilots' else []):
    subprocess.run([sys.executable,'-B',str(ROOT/'check.py'),'--tag','pilot-'+str(lane),'--',sys.executable,'-B',str(ROOT/'capture.py'),'--lane',str(lane),'--pilot'],check=True)
if a.phase=='pilots':raise SystemExit(0)
for lane in range(8):assert json.loads((ROOT/f'pilots/{lane}/receipt.json').read_text())['status']=='pass'
for lane in range(16):
    subprocess.run([sys.executable,'-B',str(ROOT/'check.py'),'--tag','formal-'+str(lane),'--',sys.executable,'-B',str(ROOT/'capture.py'),'--lane',str(lane)],check=True)
for row in json.loads((ROOT/'protocol.json').read_text())['profile_order']:
    subprocess.run([sys.executable,'-B',str(ROOT/'check.py'),'--tag',f"profile-{row['lane']}-{row['kind']}",'--',sys.executable,'-B',str(ROOT/'capture.py'),'--lane',str(row['lane']),'--profile',row['kind']],check=True)
