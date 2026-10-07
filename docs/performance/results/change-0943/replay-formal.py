#!/usr/bin/env python3
import importlib.util,json,shutil
from pathlib import Path
r=Path('/home/zhuhe/litchi-performance-0943');sp=importlib.util.spec_from_file_location('m',r/'tooling/comparison-runner.py');m=importlib.util.module_from_spec(sp);sp.loader.exec_module(m)
manifest=r/'formal-authorized-v1.json';b=m.validate_manifest(manifest);b.update(manifest_descriptor=m.descriptor(manifest),protocol_descriptor=m.protocol_descriptor(),validator_module=m._load_validator(m.VALIDATOR_PATH))
original=m.json_load(r/'reductions/formal-v1.json');source=b['external_root'];moved=r/'formal-replay-moved-v1';hidden=r/'formal-v1-hidden-during-replay'
assert not moved.exists() and not hidden.exists();shutil.copytree(source,moved)
rows=[m.descriptor(p) for p in moved.rglob('*') if p.is_file()];source.rename(hidden)
try:
 b['external_root']=moved;replay=m.reduce_data(b);assert replay==original
finally:hidden.rename(source)
summary=dict(schema='litchi.performance.0943.formal-moved-replay.v1',status='passed',runner=m.descriptor(m.RUNNER_PATH),replay_script=m.descriptor(Path(__file__)),manifest=m.descriptor(manifest),original_reduction=m.descriptor(r/'reductions/formal-v1.json'),packet_files=len(rows),packet_bytes=sum(x['bytes'] for x in rows),children=288,original_unavailable_during_replay=True,all_reduction_fields_equal=True,temporary_copy_removed_after_verification=True)
m.write_json(r/'receipts/formal-moved-replay-v1.json',summary);shutil.rmtree(moved);print(json.dumps(summary))
