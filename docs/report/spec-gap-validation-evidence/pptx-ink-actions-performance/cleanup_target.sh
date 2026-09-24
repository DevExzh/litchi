#!/usr/bin/env bash
set -euo pipefail

if [[ "$#" -ne 4 ]]; then
    echo "usage: cleanup_target.sh STATUS ROOT TARGET SENTINEL" >&2
    exit 2
fi

status=$1
root=$2
target=$3
sentinel=$4
[[ "$status" =~ ^[0-9]+$ ]] || {
    echo "cleanup status is not a nonnegative integer" >&2
    exit 2
}

expected_sentinel=$(printf 'pptx-ink-actions-profile-target-v1\nroot=%s\ntarget=%s' "$root" "$target")
actual_sentinel=''
if [[ -f "$sentinel" && ! -L "$sentinel" ]]; then
    actual_sentinel=$(cat -- "$sentinel" 2>/dev/null || true)
fi

success_sentinel="$target/.pptx-ink-actions-profile-success.sentinel"
expected_success=$(printf 'pptx-ink-actions-profile-success-v1\nroot=%s\ntarget=%s' "$root" "$target")
actual_success=''
if [[ -f "$success_sentinel" && ! -L "$success_sentinel" ]]; then
    actual_success=$(cat -- "$success_sentinel" 2>/dev/null || true)
fi

target_real=''
if [[ -d "$target" && ! -L "$target" ]]; then
    target_real=$(realpath -e -- "$target" 2>/dev/null || true)
fi

if [[ "$actual_sentinel" == "$expected_sentinel" \
    && "$actual_success" == "$expected_success" \
    && "$target_real" == "$target" ]]; then
    if [[ "$status" -eq 0 ]]; then
        find "$target" -depth -delete
        if [[ -e "$target" || -L "$target" ]]; then
            echo "verified profile-target cleanup did not remove the owned target" >&2
            exit 1
        fi
    else
        echo "retaining failed profile target for diagnosis: $target" >&2
    fi
    exit "$status"
fi

echo "refusing profile-target cleanup: sentinel or canonical path check failed" >&2
if [[ "$status" -eq 0 ]]; then
    exit 1
fi
exit "$status"
