#!/usr/bin/env python3
"""Recheck retained public corpus output and complete generated visit digests."""
import json
from pathlib import Path
P=Path(__file__).resolve().parent

def scrub(value):
    if isinstance(value, dict):
        return {k:scrub(v) for k,v in value.items() if not k.endswith("elapsed_ns")}
    if isinstance(value, list): return [scrub(v) for v in value]
    return value

rows=[]
for name in ["owned.json","file.json","synthetic-70000-default-owned.json","synthetic-70000-default-file.json"]:
    a=json.loads((P/"corpus"/"baseline"/name).read_text())
    b=json.loads((P/"corpus"/"candidate"/name).read_text())
    assert scrub(a)==scrub(b), name
    if name.startswith("synthetic"):
        v=b["records"][0]["visit"]
        assert v["actual_callbacks"]==v["oracle_callbacks"]==70001
        assert v["outcome"]["status"]=="ok"
        rows.append({"file":name,"callbacks":70001,"digest":v["oracle_digest"],"equal":True})
    else:
        assert b["query_mismatches"]==0
        rows.append({"file":name,"files_seen":b["files_seen"],"query_mismatches":0,"equal":True})
print(json.dumps(rows,indent=2))
