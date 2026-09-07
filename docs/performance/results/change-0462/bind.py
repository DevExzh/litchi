#!/usr/bin/env python3
"""Retain exact successful build executables before source or target changes."""
import json
from pathlib import Path
import shutil
import sys
import capture
ROOT=Path(__file__).resolve().parent
variant=sys.argv[1];assert variant in ["baseline","candidate"]
build_path=ROOT/"checks"/(variant+"-build.json")
build=json.loads(build_path.read_text());assert build["status"]=="pass" and build["source_unchanged"]
directory=Path("/tmp/litchi-goal-0462")/variant;directory.mkdir(parents=True)
binaries={}
for mode,name in [("normal","litchi-perf-baseline"),("allocator","litchi-perf-baseline-alloc")]:
 source=capture.REPO/"tools/perf-baseline/target/release"/name;target=directory/name
 shutil.copy2(source,target);assert capture.sha(source)==capture.sha(target)
 binaries[mode]={"path":str(target),"source":str(source.relative_to(capture.REPO)),"sha256":capture.sha(target),"bytes":target.stat().st_size}
capture.write(ROOT/(variant+"-binding.json"),{"schema":"litchi-0462-binary-binding-v1","change":462,"variant":variant,"revision":build["revision"],"build_receipt":capture.artifact(build_path),"source_manifest":build["source_after"],"binaries":binaries})
print(variant,"bound")
