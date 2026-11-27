#!/bin/bash
set -euo pipefail
# Compare with reviewed snapshots; regeneration is a separate reviewed action.
unset LITCHI_BLESS
cargo test -q -p litchi-ops --all-features --test capabilities --test registry
features=$(python3 - <<'PY'
import tomllib
with open('crates/litchi-ops/Cargo.toml', 'rb') as f:
    features=tomllib.load(f)['features']
print(','.join(k for k in features if k not in ('default', 'encryption')))
PY
)
cargo test -q -p litchi-ops --features "$features" --test capabilities --test registry
cargo test -q -p litchi-cli --all-features --test officecli_corpus
python3 tools/officecli_capability_delta.py --check
python3 tools/officecli_corpus.py --check
