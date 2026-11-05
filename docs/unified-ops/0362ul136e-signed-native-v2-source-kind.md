# 0362ul136e — signed-native v2 source-kind boundary

Source baseline: 87324b92781120f372858423b30c5d0e94fea03d (fix: bind
signed-native v2 candidate fail-closed).

This source-only follow-up removes an ambiguous identity label from the v2
endpoint receipts. The independently verified f2 endpoint remains
source_kind=f2; the eventual after endpoint is now consistently
source_kind=candidate in the policy, endpoint identity, source closure,
consumer receipts, and test fixtures. The retained 472b source is only
historical harness provenance and remains explicitly labelled
source_kind=472b under historical_harness; it cannot select or stand in
for the after endpoint.

The checked-in candidate binding is still pending, so no source tree,
archive, Cargo process, metadata command, native process, profiler, or timing
workload was run. The focused source-only suite, including the no-upstream
reference check, remains the required validation:

    PYTHONDONTWRITEBYTECODE=1 python3 -m unittest +      tools.test_signed_native_policy_v2 +      tools.test_signed_native_v2_consumer +      tools.test_no_3rdparty_references

The commit for this follow-up must preserve the repository evidence guards;
no binary or large evidence artifact is added.
