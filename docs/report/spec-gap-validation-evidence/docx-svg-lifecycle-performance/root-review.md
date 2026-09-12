# Root review record

## Historical scaffold checkpoint

At the scaffold checkpoint, independent review approved smoke wiring and
evidence integrity only. The retained smoke run contained 21 lanes, one
process, and one measured sample per lane. The full profile had not yet run
and remained gated. Those smoke samples did not establish throughput,
scaling, or a before/after performance improvement.

That review replayed the verifier, checked 25 manifest inputs, and passed the
five verifier regressions. Typed refusal paths exercised the production API
and verified physical republish, metadata, and reopened semantics. Cleanup
simulation retained full-mode and unrelated evidence. Allocation metrics and
whole-process RSS remained separate.

After that checkpoint, the README invocation examples were corrected from
`sh` to `bash`, matching the runners' Bash syntax. The smoke manifest and
provenance receipts were refreshed for that prose-only change. Production
source, harness source, and recorded binary hashes did not change, and the
smoke measurements were not rerun for the documentation correction.

## Postmeasurement full review

The approved full run then used the exact committed baseline
`892441d95db29da4351390716ef5c65b4c7c97de` with three fresh processes, two
warmups, and twenty measured samples per lane: 21 lanes and 1,260 samples.
The full verifier and an independent replay passed all 26 manifest inputs.
Source manifests, Cargo metadata, and the profile binary hash were stable;
all 63 RSS sidecars had one valid marker, exit status zero, and empty
stderr. The full report records absolute scenario-scoped observations and
makes no before/after claim.

Both expected 64-owner refusal lanes returned the typed
`SourceBackedOverlayUnavailable` error in all 60 samples per lane, preserved
the source package's physical bytes and reopened metadata, and emitted no
output bytes. The dominant-phase report records owner scaling and the costs
of staging, commit, publication, validation, and reopen separately.

This postmeasurement note is a prose refresh only. The full source manifest,
provenance receipt, and verification receipt were regenerated after this
file changed; measurements, production source, harness source, and recorded
binary did not change. The isolated checkout, target, and staging paths were
cleaned after the run. `baseline-source/` remains a retained copy for replay,
bound to the committed baseline named in the manifests.
