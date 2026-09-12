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

The smoke pins production inputs to commit
`d1f299d00e0dd5cc5cd8ddf9811c4b1ad21d1119`. The runner requires a clean,
committed descendant checkout containing the harness and four hashed native
fixtures, checks the transitive local Cargo source closure against Git blobs,
and uses only fresh external results and target directories. The standalone
`harness/Cargo.lock` is authoritative for this isolated workspace; the
repository-root ignored lock is not part of the evidence closure.

Run the source-capture and replay-verifier regression tests from this directory:

```bash
python3 -B -m unittest test_source_snapshot test_verify
```

Both test modules are required committed inputs and appear in the smoke's
before/after source manifests.

Prepare an isolated checkout at the committed scaffold descendant outside the
active worktree, choose two absent external paths, then run:

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
byte-exact raw bundle and its per-file SHA manifest remain outside the checkout
at `/var/tmp/litchi-docx-styles-effects-first-failed-raw-20260912-b`.

Production commit `d1f299d00e0dd5cc5cd8ddf9811c4b1ad21d1119` projects replacement
aggregate bytes at commit and adds the focused regression. A fresh clean 52-lane
smoke must validate that sealed source pin before any timing run.
