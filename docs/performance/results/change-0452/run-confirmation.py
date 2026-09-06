#!/usr/bin/env python3
import subprocess,sys
from pathlib import Path
ROOT=Path(__file__).resolve().parent
for lane in range(4):
    subprocess.run([sys.executable,'-B',str(ROOT/'check.py'),'--tag','confirmation-'+str(lane),'--',sys.executable,'-B',str(ROOT/'confirmation-capture.py'),'--lane',str(lane)],check=True)
