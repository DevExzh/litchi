#!/usr/bin/env python3
"""Capture separate matched whole-process counters from retained binaries."""
import json
import os
from pathlib import Path
import subprocess
import capture
import oracle
ROOT=Path(__file__).resolve().parent
protocol=json.loads((ROOT/"protocol.json").read_text()) | {"samples":100}
env=os.environ | protocol["environment"]
base=ROOT/"counters";base.mkdir()
for variant in ["baseline","candidate"]:
 binding_path=ROOT/(variant+"-binding.json");binding=json.loads(binding_path.read_text());binary=binding["binaries"]["normal"]
 assert capture.sha(Path(binary["path"]))==binary["sha256"]
 directory=base/variant;directory.mkdir();report=directory/"report.json"
 lane={"id":variant,"repeat":"diagnostic","instrumentation":"normal","shape":"large","scope":"lifecycle"}
 workload=[binary["path"],"odp-append-attribution","--mode","lifecycle","--shape","large","--warmup","3","--samples","100","--repeat","diagnostic","--output",str(report)]
 argv=["/usr/bin/time","-v","-o",str(directory/"resource.log"),"taskset","-c","2","perf","stat","-x,","-o",str(directory/"counters.csv"),"-e","cycles,instructions,branches,branch-misses,cache-misses,page-faults,context-switches","--",*workload]
 row={"schema":"litchi-0459-counters-v1","variant":variant,"binding_sha256":capture.sha(binding_path),"driver_sha256":capture.sha(Path(__file__)),"started_utc":capture.now(),"argv":argv,"scope":"whole process including setup, warmups, checks and reporting","status":"failed"}
 with (directory/"stdout.log").open("xb") as out,(directory/"stderr.log").open("xb") as err:
  result=subprocess.run(argv,cwd=capture.REPO,env=env,stdout=out,stderr=err)
 row["exit_code"]=result.returncode
 try:
  assert result.returncode==0
  row["oracle"]=oracle.validate(report,lane,protocol)
  assert capture.sha(Path(binary["path"]))==binary["sha256"]
  row["status"]="pass"
 except Exception as error:row["error"]=repr(error)
 row["finished_utc"]=capture.now();row["artifacts"]=[capture.artifact(p) for p in sorted(directory.iterdir()) if p.is_file()]
 capture.write(directory/"receipt.json",row);print(variant,row["status"],flush=True)
 if row["status"]!="pass":raise SystemExit(1)
