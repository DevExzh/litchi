#!/usr/bin/env python3
import copy, importlib.util, json, shutil
from pathlib import Path
ROOT=Path('/home/zhuhe/litchi-performance-0943')
spec=importlib.util.spec_from_file_location('runner943',ROOT/'tooling/comparison-runner.py')
m=importlib.util.module_from_spec(spec);spec.loader.exec_module(m)
packet=ROOT/'synthetic-full-root-v3'; out=ROOT/'root-controls-v3';out.mkdir(exist_ok=False)
rows=m.all_schedule();source=ROOT/'baseline-source-v1';p=m.protocol()
binding=dict(external_root=packet,protocol=p,protocol_descriptor=m.protocol_descriptor(),manifest_descriptor=m.descriptor(packet/'synthetic-manifest.json'),validator_module=m._load_validator(m.VALIDATOR_PATH),source_bindings={},build_bindings={})
for version in m.VERSIONS:
 binding['source_bindings'][version]=dict(source_root=source,source_revision=m.SOURCE_REVISION)
 binding['build_bindings'][version]={}
 for lane in ['native','allocator']:
  binary=Path(p['baseline_builds'][lane]['binary']['path'])
  binding['build_bindings'][version][lane]=dict(binary=m.descriptor(binary),binary_path=binary)
original=m.reduce_data(binding)
moved=out/'moved';shutil.copytree(packet,moved)
hidden=packet.with_name(packet.name+'-hidden-during-replay')
packet.rename(hidden)
try:
 moved_binding={**binding,'external_root':moved}
 replay=m.reduce_data(moved_binding)
 assert replay==original
finally: hidden.rename(packet)
controls=[];first=rows[0]
def mutant(name,change):
 root=out/name;root.mkdir()
 child=Path('children')/first['label'];shutil.copytree(moved/child,root/child)
 recpath=m.receipt_path(root,first);recpath.parent.mkdir()
 rec=m.json_load(m.receipt_path(moved,first));terminal=m.json_load(root/rec['terminal']['path'])
 change(rec,terminal)
 (root/rec['terminal']['path']).write_text(json.dumps(terminal)+'\n')
 rec['terminal']=m.packet_descriptor(root/rec['terminal']['path'],root)
 recpath.write_text(json.dumps(rec)+'\n')
 try:m.load_child_receipt({**binding,'external_root':root},first,rows)
 except m.RunnerError as e:controls.append(dict(name=name,status='rejected',error=str(e)))
 else:raise AssertionError('accepted '+name)
mutant('exit-status',lambda r,t:t.update(exit_code=1))
mutant('deadline',lambda r,t:t.update(deadline_exceeded=True))
mutant('runner-error',lambda r,t:t.update(runner_error='synthetic-error'))
mutant('unclosed-group',lambda r,t:t.update(process_group_closed=False))
mutant('forced-cleanup',lambda r,t:t.update(process_group_forced_cleanup=True))
def environment(r,t):r['environment']['LANG']='bad';t['environment']['LANG']='bad'
mutant('environment',environment)
def source_change(r,t):r['source']['revision']='0'*40;t['source_revision']='0'*40
mutant('source',source_change)
def binary_change(r,t):r['binary']['sha256']='0'*64;t['binary']['sha256']='0'*64
mutant('binary',binary_change)
def command_change(r,t):
 r['command'][r['command'].index('--edits')+1]='2'
 r['packet_command']=m.packet_command_view(r['command'],Path(r['capture_origin']))
 t['argv']=r['command'];t['packet_command']=r['packet_command']
mutant('scheduled-argv',command_change)
# Isolate the final gate conjunction using measured-shape synthetic reductions.
# This is a control, never evidence for actual adoption.
base_native=copy.deepcopy(original['native']);base_alloc=copy.deepcopy(original['allocator'])
for cell in base_alloc['cells']:
 if cell['edits']==32 and cell['retention']=='default':
  cell['candidate']['allocated_bytes']=int(cell['baseline']['allocated_bytes']*0.94)
m.reduce_native=lambda b:copy.deepcopy(base_native)
m.reduce_allocator=lambda b:copy.deepcopy(base_alloc)
positive=m.reduce_data(binding);assert positive['gates']['adoption'] is True
base_native['cells'][0]['latency_gate']=False
negative=m.reduce_data(binding);assert negative['gates']['adoption'] is False
controls.extend([dict(name='all-gates-positive',status='passed'),dict(name='one-latency-gate-false',status='rejected')])
duplicate=out/'duplicate-custody.json'
duplicate.write_text('{"exit_code":1,"exit_code":0}\n')
try:m.json_load(duplicate)
except m.RunnerError:controls.append(dict(name='duplicate-custody-json',status='rejected'))
else:raise AssertionError('duplicate custody key accepted')
summary=dict(schema='litchi.performance.0943.root-controls.v1',status='passed',runner=m.descriptor(m.RUNNER_PATH),protocol=m.protocol_descriptor(),controls=controls,moved_full_children=288,original_packet_unavailable_during_replay=True,full_reduction_equal=True,claim_scope='synthetic controls only; no Office workload or performance claim')
m.write_json(out/'summary.json',summary)
print(json.dumps(summary))
