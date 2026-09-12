# `stylesWithEffects` bounded smoke

This directory contains the correctness scaffold for the public DOCX
`stylesWithEffects` owner API. It exercises native captures, source snapshot
no-op, projection reads, replacement, removal, absent-owner addition, exact
inverse, stale and signed refusals, main/glossary independence, exact OPC cap
boundaries, malformed relationship topologies, malformed opaque XML, and
caller XML event/depth limits. The adapter calls the committed public API and
records typed errors; it does not synthesize timing or refusal results.

The smoke pins production inputs to commit
`d687e38349e4348506a56dc4ae996298844d4091`. The runner requires a clean,
committed descendant checkout containing the harness and four hashed native
fixtures, checks the transitive local Cargo source closure against Git blobs,
and uses only fresh external results and target directories. The standalone
`harness/Cargo.lock` is authoritative for this isolated workspace; the
repository-root ignored lock is not part of the evidence closure.

Prepare an isolated checkout at the committed scaffold descendant outside the
active worktree, choose two absent external paths, then run:

```bash
git worktree add --detach /tmp/litchi-docx-styles-effects-smoke <scaffold-commit>
DOCX_STYLES_EFFECTS_RESULTS=/var/tmp/litchi-docx-styles-effects-results-unique \
DOCX_STYLES_EFFECTS_TARGET_DIR=/var/tmp/litchi-docx-styles-effects-target-unique \
bash docs/report/spec-gap-validation-evidence/docx-styles-effects-performance/run_smoke.sh
```

Raw JSON receipts, `/usr/bin/time -v` RSS sidecars, build/source provenance,
and the fail-closed verification receipt are retained in the fresh external
results directory. The current scaffold runs one
fresh process, zero warmups, and one correctness sample per lane. Full timing,
scaling, and speedup claims remain gated until this smoke is independently
reviewed and a measurement harness is approved.
