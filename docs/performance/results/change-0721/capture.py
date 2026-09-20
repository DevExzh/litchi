#!/usr/bin/env python3
"""Run the frozen native/control ABBA cycles, then the allocator ABBA cycle."""
import json,subprocess,sys
from custody import P,ROOT,census,sha
plan=json.loads((P/'plan.json').read_text())
freeze=json.loads((P/'capture-freeze.json').read_text())
def check():
    for name,expected in freeze['files'].items():
        assert sha(ROOT/name)==expected, name
    assert census()==json.loads((P/'source-candidate.json').read_text())
rows=json.loads((P/'capture.json').read_text()) if (P/'capture.json').exists() else []
assert all(row['exit_code']==0 for row in rows)
position=0
for lane,stages in [('native',[s['label'] for s in plan['stages']]),('allocator',plan['allocator_stages'])]:
    for stage in stages:
        jobs=[['pilot.py',stage,lane]]
        if lane=='native': jobs.append(['read-controls.py','capture',stage])
        for args in jobs:
            check()
            command=[sys.executable,str(P/args[0]),*args[1:]]
            if position < len(rows):
                assert rows[position]["command"] == command
                position += 1
                continue
            position += 1
            result=subprocess.run(command,cwd=ROOT)
            check()
            rows.append({'command':command,'exit_code':result.returncode})
            (P/'capture.json').write_text(json.dumps(rows,indent=2)+'\n')
            assert result.returncode==0
