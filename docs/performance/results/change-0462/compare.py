#!/usr/bin/env python3
"""Recompute matched lifecycle results and the predeclared keep gate."""
import csv
import hashlib
import json
import math
from pathlib import Path
import random
import re
import statistics
import oracle
ROOT=Path(__file__).resolve().parent
def stats(values):
 values=sorted(values)
 return {"count":len(values),"min":values[0],"p50":statistics.median(values),"p95":values[math.ceil(.95*len(values))-1],"p99":values[math.ceil(.99*len(values))-1],"max":values[-1],"mean":statistics.mean(values)}
def bootstrap(before,after,seed):
 rng=random.Random(seed);values=[]
 for _ in range(10000):values.append((statistics.median(rng.choices(after,k=len(after)))/statistics.median(rng.choices(before,k=len(before)))-1)*100)
 values.sort()
 return {"low":values[249],"high":values[9749],"seed":seed,"resamples":10000,"method":"independent median ratio bootstrap; nearest-rank 2.5/97.5 percentiles; no multiple-comparison correction"}
def counters():
 result={}
 for variant in ["baseline","candidate"]:
  values={}
  with (ROOT/"counters"/variant/"counters.csv").open() as stream:
   for row in csv.reader(stream):
    if len(row)>=5 and row[2] in ["cycles","instructions","branches","branch-misses","cache-misses","page-faults","context-switches"]:
     values[row[2]]={"value":row[0],"unit":row[1],"event_runtime":row[3],"running_pct":row[4]}
  assert len(values)==7
  result[variant]=values
 delta={k:(int(result["candidate"][k]["value"])/int(result["baseline"][k]["value"])-1)*100 for k in result["baseline"] if result["baseline"][k]["value"].isdigit() and int(result["baseline"][k]["value"])>0 and result["candidate"][k]["value"].isdigit()}
 return {"scope":"separate whole-process runs including setup, warmups and checks; perf scaling retained; not operation-only counters","values":result,"delta_pct":delta}
def derive():
 protocol=json.loads((ROOT/"protocol.json").read_text());lanes=[];reports={}
 for variant in ["baseline","candidate"]:
  for lane in protocol["order"]:
   directory=ROOT/"runs"/variant/lane["id"];path=directory/"report.json"
   oracle.validate(path,lane,protocol);report=json.loads(path.read_text());rows=report["rows"]
   rss=re.search(r"Maximum resident set size \(kbytes\):\s*(\d+)",(directory/"resource.log").read_text());assert rss
   result={"variant":variant,"lane":lane,"report_sha256":hashlib.sha256(path.read_bytes()).hexdigest(),"lifecycle_ns":stats([r["lifecycle_ns"] for r in rows]),"operations_per_second":stats([1e9/r["lifecycle_ns"] for r in rows]),"process_lifetime_maxrss_kib":int(rss.group(1))}
   if lane["instrumentation"]=="allocator":
    allocation=[r["lifecycle_allocation_metrics"] for r in rows]
    result["allocation"]={key:stats([r[key] for r in allocation]) for key in ["allocated_bytes","allocation_calls","deallocation_calls","reallocation_calls"]}
    result["allocation"]["peak_above_entry_bytes"]=stats([r["region_peak_live_bytes"]-r["live_bytes_before"] for r in allocation])
    result["allocation"]["retained_live_delta_bytes"]=stats([r["live_bytes_after"]-r["live_bytes_before"] for r in allocation])
   lanes.append(result);reports[(variant,lane["id"])]=(result,rows)
 comparisons=[];flags=[];allocation_flags=[]
 for lane in protocol["order"]:
  before,br=reports[("baseline",lane["id"])];after,ar=reports[("candidate",lane["id"])]
  timing={q:(after["lifecycle_ns"][q]/before["lifecycle_ns"][q]-1)*100 for q in ["p50","p95","p99"]}
  rss=(after["process_lifetime_maxrss_kib"]/before["process_lifetime_maxrss_kib"]-1)*100
  row={"lane":lane,"elapsed_delta_pct":timing,"process_lifetime_maxrss_delta_pct":rss,"p50_delta_95pct_interval":bootstrap([r["lifecycle_ns"] for r in br],[r["lifecycle_ns"] for r in ar],462+len(comparisons))}
  for metric,delta in list(timing.items())+[("maxrss",rss)]:
   if delta>5:flags.append({"lane":lane["id"],"metric":metric,"delta_pct":delta})
  if "allocation" in before:
   row["allocation_equal"]=before["allocation"]==after["allocation"]
   row["allocation_p50_delta"]={k:after["allocation"][k]["p50"]-before["allocation"][k]["p50"] for k in before["allocation"]}
   row["allocation_p50_delta_pct"]={k:(after["allocation"][k]["p50"]/before["allocation"][k]["p50"]-1)*100 if before["allocation"][k]["p50"] else None for k in before["allocation"]}
   for metric,delta in row["allocation_p50_delta"].items():
    if delta>0:allocation_flags.append({"lane":lane["id"],"metric":metric,"delta":delta})
  comparisons.append(row)
 targets=[row for row in comparisons if row["lane"]["instrumentation"]=="normal" and row["lane"]["shape"] in ["medium","large"]]
 gate=all(row["elapsed_delta_pct"]["p50"]<=-3 and row["p50_delta_95pct_interval"]["high"] < 0 for row in targets)
 return {"schema":"litchi-0462-comparison-v1","reports":len(lanes),"samples":sum(row["lifecycle_ns"]["count"] for row in lanes),"scope":"unsegmented owned ODP append lifecycle; same flags, corpus and instrumentation; API phases only in separate diagnostic profiles","quantiles":"p50 midpoint median; p95/p99 nearest rank","lanes":lanes,"comparisons":comparisons,"adverse_5pct_flags":flags,"predeclared_latency_gate_pass":gate,"all_allocation_metrics_equal":all(row.get("allocation_equal",True) for row in comparisons),"counters":counters(),"allocation_increase_review_flags":allocation_flags}
if __name__=="__main__":print(json.dumps(derive(),indent=2,sort_keys=True))
