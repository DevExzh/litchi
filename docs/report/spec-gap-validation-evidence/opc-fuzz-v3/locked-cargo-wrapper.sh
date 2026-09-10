#!/usr/bin/env bash
set -euo pipefail
args=()
inserted=0
for arg in "$@"; do
    if [[ "$arg" == "--" && "$inserted" == 0 ]]; then
        args+=(--locked)
        inserted=1
    fi
    args+=("$arg")
done
if [[ "$inserted" == 0 ]]; then
    args+=(--locked)
fi
exec /tmp/litchi-spec-gap-rustup/toolchains/nightly-x86_64-unknown-linux-gnu/bin/cargo "${args[@]}"
