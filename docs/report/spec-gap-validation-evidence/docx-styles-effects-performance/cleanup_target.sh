#!/usr/bin/env bash

# Shared by run_profile.sh and its failure-policy regression test. The caller
# supplies RESULTS, TARGET, and SUCCESS_SENTINEL as absolute paths.
cleanup_profile_target() {
    if [[ -f "$SUCCESS_SENTINEL" && -f "$RESULTS/verification.json" ]] \
        && [[ "$(grep -c '^verification_sha256=' "$SUCCESS_SENTINEL")" == "1" ]]; then
        local expected actual
        expected=$(awk -F= '$1 == "verification_sha256" {print $2}' "$SUCCESS_SENTINEL")
        actual=$(sha256sum "$RESULTS/verification.json" | awk '{print $1}')
        if [[ -n "$expected" && "$expected" == "$actual" ]]; then
            find "$TARGET" -depth -delete 2>/dev/null || true
            return
        fi
    fi
    echo "profile target retained after unsuccessful or unverifiable run: $TARGET" >&2
}
