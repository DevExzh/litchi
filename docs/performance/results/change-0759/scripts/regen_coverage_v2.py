"""Rebind crud-coverage-index-v2.json to the merged registry and 0508 catalog."""
import json, sys, hashlib, re
from pathlib import Path
ROOT = Path.cwd()
sys.path.insert(0, str(ROOT))
from tools.validate_crud_coverage_index import (_selector_registry_names, _selector_names_digest,
    FORBIDDEN_IWORK, OUT_OF_SCOPE_REASON)

v2p = ROOT / 'docs/performance/crud-coverage-index-v2.json'
v1 = json.loads((ROOT / 'docs/performance/crud-coverage-index-v1.json').read_text(encoding='utf-8'))
v2 = json.loads(v2p.read_text(encoding='utf-8'))

# 1. HEAD's 0508 content changes (the only differences between the merged v1 and incoming v2).
v2['checked_catalog'] = v1['checked_catalog']
v1_categories = {c['id']: c for c in v1['categories']}
v2['categories'] = [v1_categories['conversion-export'] if c['id'] == 'conversion-export' else c
                    for c in v2['categories']]
for c in v2['categories']:
    assert json.dumps(c, sort_keys=True) == json.dumps(v1_categories[c['id']], sort_keys=True), c['id']

# 2. Registry identity from the merged Case::name order.
names = _selector_registry_names((ROOT / 'tools/perf-baseline/src/lib.rs').read_text(encoding='utf-8'))
registry = v2['selector_registry']
old_names = registry['selector_names']
registry['selector_names'] = names
registry['selector_count'] = len(names)
registry['minimum_selectable_cases'] = len(names)
registry['selector_names_sha256'] = _selector_names_digest(names)

# 3. Mapped bindings follow the category scenarios; every other name is excluded.
mapped = {}
for category in v2['categories']:
    for scenario in category['scenarios']:
        if 'selector' in scenario:
            mapped[scenario['selector']] = (category['id'], scenario['status'])
old_excluded = {e['selector']: e['reason'] for e in registry['coverage']['excluded']}
registry['coverage']['mapped'] = [
    {'selector': s, 'category_id': mapped[s][0], 'status': mapped[s][1]} for s in sorted(mapped)]
excluded = []
for s in sorted(set(names) - set(mapped)):
    reason = OUT_OF_SCOPE_REASON if FORBIDDEN_IWORK.search(s) else 'not-selected-in-representative-matrix'
    if s in old_excluded:
        assert old_excluded[s] == reason, (s, old_excluded[s], reason)
    excluded.append({'selector': s, 'reason': reason})
registry['coverage']['excluded'] = excluded

v2p.write_text(json.dumps(v2, indent=2, ensure_ascii=False) + '\n', encoding='utf-8')
added = sorted(set(names) - set(old_names))
removed = sorted(set(old_names) - set(names))
print('selectors', len(names), 'digest', registry['selector_names_sha256'])
print('mapped', len(mapped), 'excluded', len(excluded),
      'iwork-excluded', sum(1 for e in excluded if e['reason'] == OUT_OF_SCOPE_REASON))
print('added', len(added), 'removed', len(removed), removed)
Path('/home/zhuhe/code/litchi-worktrees/merge-scratch/coverage-v2-added-selectors.txt').write_text('\n'.join(added) + '\n')
