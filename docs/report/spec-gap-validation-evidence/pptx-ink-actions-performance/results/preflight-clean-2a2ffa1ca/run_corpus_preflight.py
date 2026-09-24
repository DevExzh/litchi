import hashlib
import importlib.util
import json
import os
from pathlib import Path

root = Path('/var/tmp/litchi-pptx-ink-actions-profiler-2a2ffa1c')
out = Path('/var/tmp/litchi-pptx-ink-actions-profiler-2a2ffa1c-preflight')
verify_path = root / 'docs/report/spec-gap-validation-evidence/pptx-ink-actions-performance/verify.py'
spec = importlib.util.spec_from_file_location('pptx_profile_verify', verify_path)
if spec is None or spec.loader is None:
    raise SystemExit('cannot load verifier')
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)
corpus_path = verify_path.parent / 'corpus-manifest.json'
recipes, lanes, generator = module.verify_corpus(corpus_path)
lock = verify_path.parent / 'harness/Cargo.lock'
receipt = {
    'schema': 'pptx-ink-actions-preflight-v1',
    'worktree': str(root),
    'head': os.popen(f'git -C {root} rev-parse HEAD').read().strip(),
    'owner_commit': module.SOURCE_COMMIT,
    'corpus': str(corpus_path),
    'recipe_count': len(recipes),
    'lane_count': len(lanes),
    'generator_sha256': generator,
    'lock_sha256': hashlib.sha256(lock.read_bytes()).hexdigest(),
    'timing_run': False,
    'matrix_run': False,
    'build_run': False,
}
(out / 'corpus-preflight.json').write_text(json.dumps(receipt, indent=2, sort_keys=True) + '\n')
print(json.dumps(receipt, sort_keys=True))
