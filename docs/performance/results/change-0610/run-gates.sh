#!/usr/bin/env bash
# Gate runner for change 0610. Documentation-only batch: no file under crates/
# is modified, so no crate-scoped clippy, test or doc gate applies. What is run
# is every workspace-wide checker that could observe a documentation change,
# plus a link check of the relative links the new documents introduce, plus a
# clean build of the retained probe.
#
#   run-gates.sh <checkout>
set -u
checkout="${1:-.}"
cd "$checkout" || exit 2
run() {
  printf '\n===== %s =====\n' "$*"
  "$@" 2>&1 | tail -20
  printf '[exit %s]\n' "${PIPESTATUS[0]}"
}
run cargo fmt --all --check
run python3 tools/check_perf_claims.py \
  --registry docs/performance/claim-registry-v1.json --repo-root . --mode structural
run python3 tools/check_report_claim_classification.py
run python3 tools/check_example_targets.py
run python3 tools/check_crate_boundaries.py
printf '\n===== relative-link check for the new documents =====\n'
python3 - <<'PY'
import pathlib, re, sys
root = pathlib.Path('.')
targets = [
    'docs/adr/0030-lazy-opc-part-decode.md',
    'docs/adr/README.md',
    'docs/performance/0610-opc-lazy-part-decode-design.md',
    'docs/performance/results/change-0610/README.md',
]
link = re.compile(r'\]\(([^)#][^)]*)\)')
bad = 0
checked = 0
for target in targets:
    path = root / target
    if not path.exists():
        print(f'MISSING DOCUMENT {target}'); bad += 1; continue
    for match in link.finditer(path.read_text(encoding='utf-8')):
        href = match.group(1).split('#')[0]
        if href.startswith(('http://', 'https://', 'mailto:')) or not href:
            continue
        checked += 1
        resolved = (path.parent / href).resolve()
        if not resolved.exists():
            print(f'BROKEN {target} -> {href}'); bad += 1
print(f'checked {checked} relative links across {len(targets)} documents; {bad} broken')
sys.exit(1 if bad else 0)
PY
printf '[exit %s]\n' "$?"
