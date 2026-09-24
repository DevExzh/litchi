#!/bin/bash
# The 0769 correctness lanes, one arm after the other, pinned to core 30.
# Run from the scratch root. Each lane's outputs depend only on inputs and
# seeds, so the two arms' outputs compare line for line.
set -u
cd /home/zhuhe/code/litchi-worktrees/scratch/0769
for arm in base cand; do
  mkdir -p census verdicts/$arm sweep
  # 1. The reader-agreement census over every OLE2 file found (1,830) and the
  #    normalized real-producer file.
  taskset -c 30 bin/$arm/probe --mode census --input census-paths.txt > census/$arm.jsonl 2> census/$arm.err
  echo "census $arm exit=$?"
  # 2. Record 0767's fault lane: same inputs, cases and seed.
  python3 - "$arm" <<'PY'
import json, subprocess, sys
arm = sys.argv[1]
for item in json.load(open('lane-inputs.json')):
    with open(f"verdicts/{arm}/{item['file']}", 'w') as out:
        code = subprocess.run(['taskset', '-c', '30', f'bin/{arm}/probe', '--mode', 'verdicts', '--input', item['input'], '--cases', str(item['cases']), '--seed', '767'], stdout=out).returncode
    if code != 0:
        print('verdicts', arm, item['file'], 'exit', code)
PY
  echo "verdicts $arm done"
  # 3. Every root size within 128 bytes of each file's own.
  taskset -c 30 bin/$arm/probe --mode root-sweep --input sweep-paths.txt > sweep/$arm.jsonl 2> sweep/$arm.err
  echo "sweep $arm exit=$?"
done
echo lanes-finished
