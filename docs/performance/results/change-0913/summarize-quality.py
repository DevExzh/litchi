"""Summarize the final passing quality commands without treating failed attempts as passes."""
import json
from pathlib import Path
import re

P=Path(__file__).resolve().parent

def main():
    attempts=sorted((d for d in P.glob('quality-v*') if d.is_dir()),key=lambda d:int(d.name.removeprefix('quality-v')))
    final=attempts[-1]
    result={'status':'pass','final_attempt':final.name,'commands':{},'earlier_attempts':[]}
    for kind in ('fmt','tests','clippy','docs'):
        row=json.loads((final/kind/'command.json').read_text());assert row['exit_code']==0
        result['commands'][kind]={'argv':row['argv'],'exit_code':0}
    text=(final/'tests/stdout.txt').read_text()
    rows=[tuple(map(int,m)) for m in re.findall(r'test result: ok\. (\d+) passed; (\d+) failed; (\d+) ignored; (\d+) measured; (\d+) filtered out',text)]
    assert rows and all(r[1]==0 for r in rows)
    result['test_suites']=len(rows)
    result['tests']={k:sum(r[i] for r in rows) for i,k in enumerate(('passed','failed','ignored','measured','filtered_out'))}
    for attempt in [P/'quality',*attempts[:-1]]:
        result['earlier_attempts'].append({'path':attempt.name,'commands':[{'kind':f.parent.name,'exit_code':json.loads(f.read_text())['exit_code']} for f in sorted(attempt.glob('*/command.json'))]})
    with (P/'quality-summary.json').open('x') as out:json.dump(result,out,indent=2);out.write('\n')
    print(json.dumps(result['tests']))

if __name__=='__main__':main()
