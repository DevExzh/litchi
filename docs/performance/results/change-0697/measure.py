#!/usr/bin/env python3
"""Four ordered native legs, buffered samples, no allocator instrumentation."""
import hashlib,json,statistics,subprocess,math
from pathlib import Path
P=Path(__file__).resolve().parent
ROOT=P.parents[3]
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def stats(values):
 s=sorted(values)
 return dict(p50_ns=statistics.median(s),mean_ns=statistics.mean(s),p95_ns=s[math.ceil(.95*len(s))-1],p99_ns=s[math.ceil(.99*len(s))-1],min_ns=s[0],max_ns=s[-1])
def main():
 build=json.loads((P/'build.json').read_text());binary=Path(build['binary'])
 assert sha(binary)==build['binary_sha256']
 sequences=sorted((P/'sequences').glob('*.txt'))
 assert sequences
 out=P/'native';out.mkdir(exist_ok=True)
 rows=[]
 for leg in range(4):
  for seq in (sequences if leg%2==0 else list(reversed(sequences))):
   command=['taskset','-c','12',str(binary),str(seq),'10','200','4']
   r=subprocess.run(command,capture_output=True,text=True,check=True)
   path=out/f'{seq.stem}-{leg}.tsv';path.write_text(r.stdout)
   error=path.with_suffix('.stderr');error.write_text(r.stderr)
   samples=[int(line.rsplit('=',1)[1])/4 for line in r.stdout.splitlines() if line.startswith('SAMPLE\t')]
   identity=[line for line in r.stdout.splitlines() if line.startswith('IDENTITY\t')]
   assert len(samples)==200
   rows.append(dict(case=seq.stem,leg=leg,command=command,exit_code=r.returncode,binary_sha256=sha(binary),sequence=str(seq.relative_to(P)),sequence_sha256=sha(seq),output=str(path.relative_to(P)),output_sha256=sha(path),stderr_sha256=sha(error),identity=identity,batch=4,samples=200,stats=stats(samples)))
   (P/'runs.json').write_text(json.dumps(rows,indent=2)+'\n')
   print(seq.stem,leg,flush=True)
 summary=[]
 for name in sorted({r['case'] for r in rows}):
  group=[r for r in rows if r['case']==name]
  assert len({tuple(r['identity']) for r in group})==1
  medians=[r['stats']['p50_ns'] for r in group]
  summary.append(dict(case=name,leg_median_min_ns=min(medians),leg_median_max_ns=max(medians),median_of_leg_medians_ns=statistics.median(medians),leg_median_spread_pct=(max(medians)/min(medians)-1)*100,legs=[r['stats'] for r in group]))
 (P/'summary.json').write_text(json.dumps(summary,indent=2)+'\n')
if __name__=='__main__':main()
