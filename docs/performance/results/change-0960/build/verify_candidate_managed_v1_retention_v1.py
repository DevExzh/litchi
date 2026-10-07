#!/usr/bin/env python3
from __future__ import annotations
import hashlib, json, os, stat
from datetime import datetime, timezone
from pathlib import Path

ROOT=Path('/home/zhuhe/litchi-performance-0960')
RECEIPTS=ROOT/'receipts'
TARGET=ROOT/'target-candidate-managed-v1'
EXEC=ROOT/'executables'
BUILD_BIN= {
 'xlsx_managed_compression_alloc': TARGET/'release/xlsx_managed_compression_alloc',
 'xlsx_managed_compression_benchmark': TARGET/'release/examples/xlsx_managed_compression_benchmark',
}
RETAINED= {
 'xlsx_managed_compression_alloc': EXEC/'xlsx_managed_compression_alloc.v1',
 'xlsx_managed_compression_benchmark': EXEC/'xlsx_managed_compression_benchmark.v1',
}
ORIGINAL=RECEIPTS/'candidate-managed-v1-retention.json'
OUT=RECEIPTS/'candidate-managed-v1-retention-correction-v1.json'
RESULT=RECEIPTS/'candidate-managed-v1-build-result.json'

def sha(p):
 h=hashlib.sha256()
 with p.open('rb') as f:
  for b in iter(lambda:f.read(1024*1024),b''): h.update(b)
 return h.hexdigest()
def desc(p):
 st=p.stat(); return {'path':str(p),'exists':True,'bytes':st.st_size,'mode':stat.S_IMODE(st.st_mode),'sha256':sha(p)}
def helper_desc():
 p=Path(__file__); return desc(p)

if OUT.exists(): raise SystemExit(f'existing output: {OUT}')
rows={}; passed=True
for role in BUILD_BIN:
 a=desc(BUILD_BIN[role]); b=desc(RETAINED[role])
 match=all(a[k]==b[k] for k in ('bytes','mode','sha256'))
 rows[role]={'build_output':a,'retained':b,'byte_hash_mode_match':match}
 passed &= match
if not passed: raise SystemExit(json.dumps(rows,sort_keys=True))
orig=json.loads(ORIGINAL.read_text()); result=json.loads(RESULT.read_text())
receipt={
 'schema':'litchi.performance.0960.candidate-managed-v1.retention-correction.v1',
 'created_utc':datetime.now(timezone.utc).isoformat().replace('+00:00','Z'),
 'status':'validated',
 'original_retention_receipt':{'path':str(ORIGINAL),'bytes':ORIGINAL.stat().st_size,'sha256':sha(ORIGINAL)},
 'original_build_result':{'path':str(RESULT),'bytes':RESULT.stat().st_size,'sha256':sha(RESULT),'cargo_returncode':result.get('returncode'),'source_unchanged':result.get('source_unchanged')},
 'correction_reason':'Original wrapper compared full descriptors including path, so distinct build-output and retained paths made byte-identical copies report false. This correction compares bytes, mode, and SHA-256 while retaining both path descriptors.',
 'helper':helper_desc(),
 'binaries':rows,
 'no_rebuild':True,
 'no_workload':True,
}
OUT.write_text(json.dumps(receipt,indent=2,sort_keys=True)+'\n')
corrected={
 'schema':'litchi.performance.0960.candidate-managed-v1.build-result-correction.v1',
 'created_utc':receipt['created_utc'],
 'status':'validated',
 'original_result':{'path':str(RESULT),'bytes':RESULT.stat().st_size,'sha256':sha(RESULT)},
 'retention_correction':{'path':str(OUT),'bytes':OUT.stat().st_size,'sha256':sha(OUT)},
 'cargo_returncode':result.get('returncode'),
 'source_unchanged':result.get('source_unchanged'),
 'source_verification_start_pass':result.get('source_verification_start_pass'),
 'source_verification_end_pass':result.get('source_verification_end_pass'),
 'binary_parse_after_terminal':result.get('binary_parse_after_terminal'),
 'all_binaries_present':result.get('all_binaries_present'),
 'retained_executables':True,
 'wrapper_false_negative_corrected':True,
 'no_rebuild':True,
}
cp=RECEIPTS/'candidate-managed-v1-build-result-correction-v1.json'
if cp.exists(): raise SystemExit(f'existing output: {cp}')
cp.write_text(json.dumps(corrected,indent=2,sort_keys=True)+'\n')
print(json.dumps({'retention_receipt':str(OUT),'retention_sha256':sha(OUT),'result_correction':str(cp),'result_correction_sha256':sha(cp),'helper_sha256':helper_desc()['sha256'],'validated':True,'no_rebuild':True},sort_keys=True))
