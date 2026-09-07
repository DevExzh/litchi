#!/usr/bin/env python3
"""Recompute supplementary public API phase timings."""
import json
from pathlib import Path
import oracle
from compare import stats
ROOT=Path(__file__).resolve().parent
def derive():
 protocol=json.loads((ROOT/"protocol.json").read_text());reports={};result=[]
 for repeat in ["R1","R2"]:
  for variant in ["baseline","candidate"]:
   lane={"id":variant+"-"+repeat,"repeat":repeat,"instrumentation":"normal","shape":"large","scope":"phases"}
   path=ROOT/"phase-diagnostic"/lane["id"]/"report.json";oracle.validate(path,lane,protocol)
   rows=json.loads(path.read_text())["rows"]
   row={"variant":variant,"repeat":repeat,"lifecycle_ns":stats([r["lifecycle_ns"] for r in rows]),"phases":{name:stats([r["phases"][i]["elapsed_ns"] for r in rows]) for i,name in enumerate(oracle.PHASE_NAMES)}}
   reports[(variant,repeat)]=row;result.append(row)
 deltas=[]
 for repeat in ["R1","R2"]:
  before=reports[("baseline",repeat)];after=reports[("candidate",repeat)]
  deltas.append({"repeat":repeat,"p50_delta_pct":{name:(after["phases"][name]["p50"]/before["phases"][name]["p50"]-1)*100 for name in oracle.PHASE_NAMES}})
 return {"schema":"litchi-0462-phase-summary-v1","reports":4,"samples":120,"scope":"supplementary large normal public-API phase clocks; separate from unsegmented keep-decision matrix","results":result,"comparisons":deltas}
if __name__=="__main__":print(json.dumps(derive(),indent=2,sort_keys=True))
