"""Run explicitly selected source-bound checks serially."""
import argparse
import subprocess
import sys

from run import HERE, REPO

parser = argparse.ArgumentParser()
parser.add_argument('epoch')
parser.add_argument('lanes', nargs='+')
args = parser.parse_args()
for lane in args.lanes:
    subprocess.run([sys.executable, '-B', str(HERE / 'check.py'), '--epoch', args.epoch, lane],
                   cwd=REPO, check=True)
