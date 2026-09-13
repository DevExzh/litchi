#!/usr/bin/env python3
from __future__ import annotations
import csv, hashlib, json, re, shlex, subprocess, sys, time
from pathlib import Path
FIELDS=("group","workload","case","input_bytes","repeat","warmups","iterations","expected_success","mean_ns","p50_ns","p95_ns","p99_ns","alloc_calls_p50","alloc_calls_max","dealloc_calls_p50","dealloc_calls_max","requested_bytes_p50","requested_bytes_max","released_bytes_p50","released_bytes_max","live_before_p50","live_after_p50","live_after_max","peak_live_delta_p50","peak_live_delta_max","successes_p50","successes_max","checksum_p50","checksum_max","max_rss_kib","status")

def sha(p):
 h=hashlib.sha256()
 with p.open('rb') as f:
  for b in iter(lambda:f.read(1<<20),b''): h.update(b)
 return h.hexdigest()
def kv(line): return dict(x.split('=',1) for x in line.split()[1:])
def rss(p):
 m=re.search(r'^\s*Maximum resident set size \(kbytes\):\s*(\d+)\s*$',p.read_text(),re.M)
 return m.group(1) if m else ''
def capture(binary,out):
 out.mkdir(parents=True,exist_ok=True); case='parse-local-refs-256'; repeat=128; stem='parse-'+case
 so=out/f'{stem}.stdout'; se=out/f'{stem}.stderr'; tm=out/f'{stem}.time'; st=out/f'{stem}.status'
 cmd=['taskset','-c','2','/usr/bin/time','-v','-o',str(tm),str(binary),'--workload','parse','--case',case,'--warmups','3','--iterations','15','--repeat',str(repeat)]
 (out/'commands.txt').write_text(shlex.join(cmd)+'\n')
 with so.open('w') as fo,se.open('w') as fe: p=subprocess.run(cmd,stdout=fo,stderr=fe)
 st.write_text(str(p.returncode)+'\n'); lines=so.read_text().splitlines(); cfg=kv(next(x for x in lines if x.startswith('config '))); res=kv(next(x for x in lines if x.startswith('result ')))
 row={'group':'local-refs-abab','workload':'parse','case':case,'input_bytes':cfg['input_bytes'],'repeat':cfg['repeat'],'warmups':cfg['warmups'],'iterations':cfg['iterations'],'expected_success':cfg['expected_success'],**res,'max_rss_kib':rss(tm),'status':str(p.returncode)}
 with (out/'raw.csv').open('w',newline='') as f:
  writer=csv.DictWriter(f,fieldnames=FIELDS,lineterminator='\n')
  writer.writeheader()
  writer.writerow({x:row.get(x,'') for x in FIELDS})
 return p.returncode
if __name__=='__main__':
 if len(sys.argv)!=4: raise SystemExit(f'usage: {sys.argv[0]} LABEL BINARY OUTPUT')
 label,binary_text,out_text=sys.argv[1:]; binary=Path(binary_text); out=Path(out_text); start=time.time(); bh=sha(binary)
 (out/'run-start.json').write_text(json.dumps({'label':label,'binary':str(binary),'binary_sha256':bh,'case':'parse-local-refs-256','repeat':128,'warmups':3,'iterations':15,'started_at':time.strftime('%Y-%m-%dT%H:%M:%SZ',time.gmtime(start))},indent=2)+'\n')
 status=capture(binary,out); raw=sha(out/'raw.csv'); end=time.time(); (out/'run-end.json').write_text(json.dumps({'label':label,'binary_sha256':bh,'status':status,'raw_sha256':raw,'elapsed_seconds':end-start,'finished_at':time.strftime('%Y-%m-%dT%H:%M:%SZ',time.gmtime(end))},indent=2)+'\n'); raise SystemExit(status)
