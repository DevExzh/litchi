#!/usr/bin/env python3
"""Bind and profile the unchanged source with explicit frame-pointer flags."""
import hashlib
import json
import os
from pathlib import Path
import subprocess
import datetime
import oracle

ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[3]
BINARY = Path("/tmp/litchi-goal-0459/diagnostic-target/release/litchi-perf-baseline")
def sha(path): return hashlib.sha256(path.read_bytes()).hexdigest()
def now(): return datetime.datetime.now(datetime.timezone.utc).isoformat()
def artifact(path): return {"path": str(path.relative_to(ROOT)), "bytes": path.stat().st_size, "sha256": sha(path)}
def write(path,value):
 with path.open("x") as f: json.dump(value,f,indent=2); f.write("\n")
def run(directory,name,argv,env):
 with (directory/(name+".stdout")).open("xb") as out, (directory/(name+".stderr")).open("xb") as err:
  result = subprocess.run(argv,cwd=REPO,env=env,stdout=out,stderr=err)
 return {"argv":argv,"exit_code":result.returncode}
def main():
 build=ROOT/"checks/diagnostic-build.json"
 receipt=json.loads(build.read_text()); assert receipt["status"]=="pass" and receipt["source_unchanged"]
 binary_sha=sha(BINARY)
 binding={"binary":str(BINARY),"sha256":binary_sha,"bytes":BINARY.stat().st_size,"build_receipt":artifact(build),"source":receipt["source_after"],"revision":receipt["revision"],"driver_sha256":sha(Path(__file__)),"oracle_sha256":sha(ROOT/"oracle.py"),"scope":"diagnostic build with force-frame-pointers and force-unwind-tables; distinct from ordinary timing binaries"}
 write(ROOT/"diagnostic-binding.json",binding)
 environment=os.environ | {"RUSTUP_TOOLCHAIN":"1.98.1","DEBUGINFOD_URLS":"","PYTHONDONTWRITEBYTECODE":"1"}
 protocol=json.loads((ROOT/"reference-protocol.json").read_text()) | {"samples":100}
 base=ROOT/"diagnostic";base.mkdir()
 for mode in ["fp","dwarf"]:
  directory=base/mode;directory.mkdir()
  lane={"id":mode,"scope":"phases","shape":"large","instrumentation":"normal","repeat":"diagnostic"}
  report=directory/"report.json"
  workload=[str(BINARY),"odp-append-attribution","--mode","phases","--shape","large","--warmup","3","--samples","100","--repeat","diagnostic","--output",str(report)]
  argv=["/usr/bin/time","-v","-o",str(directory/"resource.log"),"taskset","-c","2","perf","record","--no-buildid-cache","-F","99","-e","cycles:u","--call-graph", "fp" if mode=="fp" else "dwarf,32768", "-o",str(directory/"perf.data"),"--",*workload]
  row={"started_utc":now(),"binding_sha256":sha(ROOT/"diagnostic-binding.json"),"mode":mode,"status":"failed"}
  row["workload"]=run(directory,"workload",argv,environment)
  try:
   assert row["workload"]["exit_code"]==0
   row["oracle"]=oracle.validate(report,lane,protocol)
   row["script"]=run(directory,"perf-script",["perf","script","--no-inline","-i",str(directory/"perf.data")],environment)
   row["symbols"]=run(directory,"symbols",["nm","-C",str(BINARY)],environment)
   assert row["script"]["exit_code"]==row["symbols"]["exit_code"]==0
   assert sha(BINARY)==binary_sha
   row["status"]="pass"
  except Exception as error: row["error"]=repr(error)
  row["finished_utc"]=now(); row["artifacts"]=[artifact(p) for p in sorted(directory.iterdir()) if p.is_file()]
  write(directory/"receipt.json",row)
  print(mode,row["status"],flush=True)
  if row["status"]!="pass": return 1
 return 0
if __name__=="__main__": raise SystemExit(main())
