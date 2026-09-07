#!/usr/bin/env python3
"""Supplement the lifecycle matrix with matched large public-API phase clocks."""
import json
import os
from pathlib import Path
import subprocess
import capture
import oracle
ROOT=Path(__file__).resolve().parent
protocol=json.loads((ROOT/"protocol.json").read_text())
base=ROOT/"phase-diagnostic";base.mkdir()
for repeat,variants in [("R1",["baseline","candidate"]),("R2",["candidate","baseline"])]:
 for variant in variants:
  binding_path=ROOT/(variant+"-binding.json");binding=json.loads(binding_path.read_text());binary=binding["binaries"]["normal"]
  assert capture.sha(Path(binary["path"]))==binary["sha256"]
  directory=base/(variant+"-"+repeat);directory.mkdir();report=directory/"report.json"
  lane={"id":variant+"-"+repeat,"repeat":repeat,"instrumentation":"normal","shape":"large","scope":"phases"}
  workload=[part.format_map(dict(lane,binary=binary["path"],report=str(report))) for part in protocol["argv"]]
  argv=["/usr/bin/time","-v","-o",str(directory/"resource.log"),"taskset","-c","2",*workload]
  row={"schema":"litchi-0462-phase-diagnostic-v1","lane":lane,"variant":variant,"binding_sha256":capture.sha(binding_path),"driver_sha256":capture.sha(Path(__file__)),"protocol_sha256":capture.sha(ROOT/"protocol.json"),"started_utc":capture.now(),"argv":argv,"status":"failed"}
  with (directory/"stdout.log").open("xb") as out,(directory/"stderr.log").open("xb") as err:
   result=subprocess.run(argv,cwd=capture.REPO,env=os.environ|protocol["environment"],stdout=out,stderr=err)
  row["exit_code"]=result.returncode
  try:
   assert result.returncode==0
   row["oracle"]=oracle.validate(report,lane,protocol)
   assert capture.sha(Path(binary["path"]))==binary["sha256"]
   row["status"]="pass"
  except Exception as error:row["error"]=repr(error)
  row["finished_utc"]=capture.now();row["artifacts"]=[capture.artifact(p) for p in sorted(directory.iterdir()) if p.is_file()]
  capture.write(directory/"receipt.json",row);print(variant,repeat,row["status"],flush=True)
  if row["status"]!="pass":raise SystemExit(1)
