#!/usr/bin/env python3
"""Recompute diagnostic sampled periods and inclusive ancestry by API marker."""
import collections
import json
import re
from pathlib import Path
ROOT=Path(__file__).resolve().parent
HEADER=re.compile(r"^\S.*:\s+(\d+)\s+cycles:u:\s*$")
MARKER=re.compile(r"odp_append_attribution::phase_(snapshot_open|transaction|add|commit|publication)(?:\W|$)")
def classify(path):
 groups=collections.defaultdict(lambda:{"samples":0,"period":0,"leaves":collections.Counter(),"inclusive":collections.Counter()})
 frames=[]; period=None; malformed=0
 def symbol(frame): return re.sub(r"\+0x[0-9a-f]+$","",re.sub(r"^[0-9a-f]+\s+","",frame).split(" (")[0])
 def flush():
  nonlocal frames,period,malformed
  if period is None:return
  if not frames:malformed+=1
  else:
   markers={m.group(1) for frame in frames for m in MARKER.finditer(frame)}
   group=next(iter(markers)) if len(markers)==1 else ("unattributed" if not markers else "ambiguous")
   row=groups[group];row["samples"]+=1;row["period"]+=period
   row["leaves"][symbol(frames[0])]+=period
   for name in {symbol(frame) for frame in frames}:row["inclusive"][name]+=period
  frames=[];period=None
 for line in path.read_text(errors="replace").splitlines():
  match=HEADER.match(line)
  if match:flush();period=int(match.group(1))
  elif not line.strip():flush()
  elif line[:1].isspace() and period is not None:frames.append(line.strip())
 flush()
 total=sum(v["period"] for v in groups.values());assert total>0
 result={}
 for name,row in sorted(groups.items()):
  value={"samples":row["samples"],"period":row["period"],"share_total_period_pct":row["period"]/total*100}
  for key in ["leaves","inclusive"]:
   value[key]=[{"symbol":symbol,"period":weight,"share_phase_period_pct":weight/row["period"]*100} for symbol,weight in sorted(row[key].items(), key=lambda item: (-item[1], item[0]))[:40]]
  result[name]=value
 return {"samples":sum(v["samples"] for v in groups.values()),"malformed":malformed,"total_period":total,"groups":result}
def derive():
 return {"schema":"litchi-0459-diagnostic-summary-v1","scope":"whole process, including setup and warmups; sampled period estimates, not per-phase hardware counters; inclusive symbol shares overlap","modes":{mode:classify(ROOT/"diagnostic"/mode/"perf-script.stdout") for mode in ["fp","dwarf"]}}
if __name__=="__main__":print(json.dumps(derive(),indent=2,sort_keys=True))
