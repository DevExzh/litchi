#!/usr/bin/env bash
set -euo pipefail

script_dir="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)"
if [[ $# -gt 1 ]]; then
    echo "usage: replay.sh [result-directory]" >&2
    exit 2
fi
if [[ $# -eq 1 ]]; then
    prior_result="$(realpath -m -- "$1")"
    if [[ -f "$prior_result/source-root.txt" ]]; then
        prior_source="$(sed -n '1p' "$prior_result/source-root.txt")"
        if [[ -e "$prior_source/.git" ]] && \
            [[ "$(git --no-replace-objects -C "$prior_source" rev-parse HEAD 2>/dev/null || true)" == "16102fe751d7c5492042330f1bd1f49c304495f0" ]]; then
            export XLSB_CUSTOM_DATA_SOURCE_ROOT="$prior_source"
        else
            echo "retained source checkout is unavailable; reconstructing commit 16102fe751d7c5492042330f1bd1f49c304495f0" >&2
            unset XLSB_CUSTOM_DATA_SOURCE_ROOT
        fi
    else
        echo "prior receipt has no source-root.txt; reconstructing pinned source checkout" >&2
        unset XLSB_CUSTOM_DATA_SOURCE_ROOT
    fi
fi
export XLSB_CUSTOM_DATA_RESULT_DIR="${XLSB_CUSTOM_DATA_REPLAY_RESULT_DIR:-$script_dir/replay/run-$(date -u +%Y%m%dT%H%M%SZ)-$$}"
"$script_dir/run_profile.sh"
