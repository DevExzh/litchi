#!/usr/bin/env bash
set -euo pipefail

if [[ "${PROFILE_FROZEN:-}" != 1 ]]; then
    echo "refusing full profile before the DOCX lifecycle owner is frozen; set PROFILE_FROZEN=1" >&2
    exit 2
fi
for variable in RUSTFLAGS CARGO_ENCODED_RUSTFLAGS RUSTC_BOOTSTRAP; do
    if [[ -n "${!variable:-}" ]]; then
        echo "refusing profile with ${variable} set" >&2
        exit 2
    fi
done
SCRIPT_DIR=$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)
PROFILE_MODE=full exec "$SCRIPT_DIR/run_smoke.sh"

