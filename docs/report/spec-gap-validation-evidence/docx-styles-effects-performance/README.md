# `stylesWithEffects` bounded smoke

This directory contains the correctness scaffold for the public DOCX
`stylesWithEffects` owner API. It exercises native captures, source snapshot
no-op, projection reads, replacement, removal, absent-owner addition, exact
inverse, stale and signed refusals, main/glossary independence, exact OPC cap
boundaries, malformed relationship topologies, malformed opaque XML grammar
(including QName, namespace, delimiter, control, character-reference, and
XML-version cases), and caller XML event/depth limits. The adapter calls the
committed public API and records typed errors; it does not synthesize timing or
refusal results.

The historical `run_smoke.sh` pins production inputs to commit
`d1f299d00e0dd5cc5cd8ddf9811c4b1ad21d1119`. The runner requires a clean,
committed descendant checkout containing the harness and four hashed native
fixtures, checks the transitive local Cargo source closure against Git blobs,
and uses only fresh external results and target directories. The standalone
`harness/Cargo.lock` is authoritative for this isolated workspace; the
repository-root ignored lock is not part of the evidence closure.
The separate `profile-harness/Cargo.lock` is likewise authoritative for the
gated performance scaffold.

Run the source-capture and replay-verifier regression tests from this directory:

```bash
python3 -B -m unittest test_source_snapshot test_verify
```

Both test modules are required committed inputs and appear in the smoke's
before/after source manifests.

To reproduce that historical smoke, restore its frozen historical scaffold
checkout outside the active worktree and choose two absent external paths.
Do not combine `run_smoke.sh` with the current harness, whose receipt source
pin is different. The current-source gate uses `run_current_smoke.sh` as
described below. Historical invocation:

```bash
git worktree add --detach /tmp/litchi-docx-styles-effects-smoke <scaffold-commit>
DOCX_STYLES_EFFECTS_RESULTS=/var/tmp/litchi-docx-styles-effects-results-unique \
DOCX_STYLES_EFFECTS_TARGET_DIR=/var/tmp/litchi-docx-styles-effects-target-unique \
bash /tmp/litchi-docx-styles-effects-smoke/docs/report/spec-gap-validation-evidence/docx-styles-effects-performance/run_smoke.sh
```

Raw JSON receipts, `/usr/bin/time -v` RSS sidecars, build/source provenance,
and the fail-closed verification receipt are retained in the fresh external
results directory. The current scaffold runs one
fresh process, zero warmups, and one correctness sample per lane. Full timing,
scaling, and speedup claims remain gated until this smoke is independently
reviewed and a measurement harness is approved.

The first frozen-source run has a durable failure slice in
[`historical-first-failed-20260912/`](historical-first-failed-20260912/).
The preserved capture contained 43 lane receipts and was rejected by the
fail-closed verifier because existing-owner replacement in
`cap_total_part_bytes` grew aggregate part bytes from 110,027 to 110,151 while
the one-unit-under limit was 110,150 and the transaction commit-stage refusal
was not observed. The exact cap receipt, source manifest, build/provenance
receipts, commands, and hashes are retained in that directory. The complete
byte-exact raw bundle remains outside the checkout at
`/var/tmp/litchi-docx-styles-effects-first-failed-raw-20260912-b`; its per-file
SHA manifest is retained as `historical-first-failed-20260912/raw-bundle-sha256.txt`.

Production commit `d1f299d00e0dd5cc5cd8ddf9811c4b1ad21d1119` projects replacement
aggregate bytes at commit and adds the focused regression. The fresh
[clean 46c456848 capture](results/clean-46c456848/) passes all 52 lanes,
including the existing-owner commit-time cap and opaque-member checks. The
runner and root verifier replays agree. This closes the correctness-smoke
gate for that recorded source; the smoke itself is not performance evidence.

The current attribution baseline has a separate, unrun 52-lane gate in
[`run_current_smoke.sh`](run_current_smoke.sh). It binds the clean descendant
to `8702fd4db8723acceb7deb51bcb40ff66604bf10`, verifies the same native
package and member hashes through `current-corpus-manifest.json`, and invokes
`verify_current_smoke.py`. It writes `smoke-*` receipts only into a fresh
external results directory supplied by `DOCX_STYLES_EFFECTS_RESULTS`; it does
not read or rewrite the retained `clean-46c456848` receipts. Run it only after
the current-source preflight is reviewed. The repaired frozen harness passed
the [current-source 52-lane correctness smoke](results/smoke-current-637082e31/)
with independent review and root verifier replay. No current-source timing
capture is included.

The separate performance scaffold passed independent review and uses a
refusal-by-default runner. It requires a reviewed committed descendant, fresh
external results and Cargo target paths, and the explicit `PROFILE_FROZEN=1` plus
`DOCX_STYLES_EFFECTS_PROFILE_API_WIRED=1` gates. Its reviewed invocation is
`bash run_profile.sh`; it uses three fresh processes, two warmups, and twenty
samples over the bounded representative success matrix. The current scaffold
checks the approved current smoke's exact file set and bytes against
`fa927a8a9` before building, writes generated
64 KiB/1 MiB XML and DOCX bytes once to a fresh external manifest, and records
actual source status, toolchain, target, linker, environment, and process-time
receipts. A failed build, process, or verifier retains the disposable Cargo
target; cleanup happens only after a matching post-verification sentinel.
The generated fixture paths and result/target paths are checked for exact
hashes and external disjointness during replay. The current scaffold
does not authorize a timing run; those environment gates are set only for a
separately authorized clean run. It does not produce a before/after or speedup
claim.

The initial authorized [profile capture](results/profile-clean-96958498f/)
is retained in `470bbf288`: 31 lane/scale rows, 93 fresh processes, and 1,860
measured samples at native, 64 KiB, and 1 MiB sizes. Independent review and
root replay passed. Commit `ecc2ce802` additionally verifies restoration of
all 307 raw files from Git, replay at their recorded paths, and cleanup of
external duplicates. A verified source bundle preserves the isolated capture
commit. These are bounded absolute observations of the pinned source; the
report records the shared-host evidence limits. They establish neither a
before/after improvement nor performance of later production changes.
