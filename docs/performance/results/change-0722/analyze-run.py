#!/usr/bin/env python3
"""Record the initial gate decisions or the finalized source-bound analyses."""
import json,subprocess,sys
sys.dont_write_bytecode = True
from custody import P,ROOT,sha
mode=sys.argv[1];assert mode in ['initial','final']
rows=json.loads((P/'capture.json').read_text());assert len(rows)==20 and all(r['exit_code']==0 for r in rows)
receipt=P/f'analysis-{mode}-commands.json';assert not receipt.exists()
rows=[]
for script,stem in [('read-controls-analyze.py','read-controls'),('analyze.py','primary')]:
 name=(f'{stem}-initial-analysis.json' if mode=='initial' else ('read-controls-analysis.json' if stem=='read-controls' else 'analysis.json'))
 output=P/name;assert not output.exists();log=P/f'analysis-{mode}-{stem}.log';assert not log.exists()
 command=[sys.executable,'-B',str(P/script),*(['analyze'] if stem=='read-controls' else []),'--output',str(output)]
 with log.open('w') as f:r=subprocess.run(command,cwd=ROOT,stdout=f,stderr=subprocess.STDOUT)
 row={'command':command,'exit_code':r.returncode,'script_sha256':sha(P/script),'log':log.name,'log_sha256':sha(log)}
 if output.exists():row['output_sha256']=sha(output)
 rows.append(row);receipt.write_text(json.dumps(rows,indent=2)+'\n')
 assert output.exists(),log.read_text()
 value=json.loads(output.read_text());decision=value['decision'];assert r.returncode==0 or (mode=='initial' and stem=='read-controls' and r.returncode==1 and decision['accepted'] is False),log.read_text()
 print(stem,'accepted:',decision['accepted'],'failed gates:',sum(not x['pass'] for x in decision['hard_gates']),flush=True)
