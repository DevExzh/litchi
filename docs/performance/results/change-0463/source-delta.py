#!/usr/bin/env python3
"""Bind exact changed source bytes to both compiled epochs without new Rust inputs."""
import difflib
import hashlib
import json
from pathlib import Path
import subprocess
ROOT=Path(__file__).resolve().parent
REPO=ROOT.parents[3]
def sha(raw): return hashlib.sha256(raw).hexdigest()
def load(name): return json.loads((ROOT/name).read_text())
bindings={v:load(v+"-binding.json") for v in ["baseline","candidate"]}
manifests={v:load(b["source_manifest"]["path"]) for v,b in bindings.items()}
changed=sorted(k for k in set(manifests["baseline"])|set(manifests["candidate"]) if manifests["baseline"].get(k)!=manifests["candidate"].get(k))
rows=[];patch=[]
for name in changed:
 row={"path":name};texts={}
 for variant in ["baseline","candidate"]:
  if name not in manifests[variant]: row[variant]=None;texts[variant]="";continue
  raw=subprocess.check_output(["git","show",bindings[variant]["revision"]+":"+name],cwd=REPO) if variant=="baseline" else (REPO/name).read_bytes()
  assert sha(raw)==manifests[variant][name],(variant,name)
  path=ROOT/"source-artifacts"/variant/(name+".txt")
  path.parent.mkdir(parents=True,exist_ok=True)
  with path.open("xb") as stream: stream.write(raw)
  row[variant]={"path":str(path.relative_to(ROOT)),"bytes":len(raw),"sha256":sha(raw)}
  texts[variant]=raw.decode()
 patch.extend(difflib.unified_diff(texts["baseline"].splitlines(True),texts["candidate"].splitlines(True),fromfile="a/"+name,tofile="b/"+name))
 rows.append(row)
with (ROOT/"source-delta.json").open("x") as stream: json.dump({"schema":"litchi-0463-source-delta-v1","change":463,"files":rows},stream,indent=2);stream.write("\n")
with (ROOT/"experiment.patch").open("x") as stream: stream.write("".join(patch))
print(json.dumps({"changed_files":len(rows),"status":"pass"}))
