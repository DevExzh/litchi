#!/usr/bin/env python3
import subprocess,sys
from pathlib import Path
ROOT=Path(__file__).resolve().parent
for i in range(8):
 subprocess.run([sys.executable,'-B',str(ROOT/'check.py'),'--tag','confirmation-'+str(i),'--',sys.executable,'-B',str(ROOT/'confirmation-capture.py'),'--lane',str(i)],check=True)
