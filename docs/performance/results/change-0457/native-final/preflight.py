#!/usr/bin/env python3
"""Authenticate and replay native records before the final bundle seal exists.

This preflight does not replace verify.py's final sealed-bundle check.
"""
import json
from pathlib import Path
import runpy

ROOT = Path(__file__).resolve().parent
verifier = runpy.run_path(str(ROOT / 'verify.py'))
bindings = verifier['authenticate_bindings'](True)
fixtures = verifier['authenticate_inventory'](bindings['inventory_path'], True)
records, validated, probe_failures, oracle_failures = verifier['verify_records'](
    fixtures, bindings, True)
assert validated == len(fixtures) == 10
assert probe_failures == oracle_failures == 0
replayed = verifier['replay_oracles'](
    records, fixtures, ROOT / 'runs', bindings['oracle_path'])
assert replayed == 10
print(json.dumps({'status': 'pass', 'records': len(records),
                  'validated_records': validated, 'oracle_replays': replayed,
                  'original_sources_and_binary_checked': True,
                  'final_sealed_bundle_verification': 'pending'}))
