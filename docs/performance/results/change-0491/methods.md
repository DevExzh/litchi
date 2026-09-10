# 0491 DOCX source-provider and cache-state methods

Previous turn classification: progress. The latest cleanup removed 96 stale
compiler intermediate directories, reclaiming about 8.5 GiB, and rebased onto
the fetched remote branch while preserving all 228 local modified/untracked files.
The earlier cleanup reclaimed about 46 GiB and removed one clean merged worktree. Change 0490
is committed. The full non-iWork goal remains active.

This batch addresses missing provider and cache-state evidence, not another
candidate-XML optimization. Accepted ADR hashes were refreshed against the
previously read set and are unchanged. The protected spec-gap worktree and
associated validation/build directories are excluded.

## Filesystem lifecycle

Reuse `docx_file_source_open_full_text_lifecycle`. Its timer includes source
open, full-text extraction, and document destruction. The returned text remains
live for verification after the timer. The existing
`docx_file_source_full_text` prepared-query selector is an ineligibility
control, never a verified-cold performance result.

The synthetic corpus is the existing pinned DOCX media corpus: 200 paragraphs
and eight 2 MiB media entries. Bind the actual full manifest/archive hash in
every captured report. The verifier-owned aligned copy uses EOCD comment
padding and must retain the same semantic oracle.

Warm, cold-requested, and cold-verified are separate cache states. Warm uses a
fresh priming child. Cold-requested records only an accepted eviction request.
Verified cold requires sync and DONTNEED on the aligned copy, strict fincore
zero resident/dirty/writeback observation, then a positive read_bytes delta
inside the open/read lifecycle. No source hashing or preparation may occur
between the final residency observation and the timer. This proves the stated
page-cache and process-accounting condition, not physical device temperature.
No host-wide caches are dropped.

The planned formal matrix has two repeats per normal/allocator role, three
warmups and 30 samples per admitted cache state; reverse role order in the
second repeat. Explicit ineligible outcomes must remain visible. Freeze the
protocol and validators after pilot/review and before formal capture. Pilot
captures are retained separately and cannot substitute for formal samples.

## Provider lifecycle

A benchmark-only provider entry point keeps one identical text oracle
across owned, instrumented, bounded-short, bounded-delayed and positional-file
sources. Source construction is outside its timer; API open, text extraction
and document destruction are inside; returned text verification, hashing and
destruction are outside. Its file source preparation scope
must remain explicit and must not be compared as an equal-scope alternative
to filesystem path-open timing. Adapter counters describe logical reads,
not network or device I/O. All simulated delay/range/bandwidth parameters
must be finite, explicit and retained.

A copied slice or static source does not satisfy non-static borrowed input.
The independent ownership audit will identify the actual implementation gap.
Read-only evidence cannot establish sequential publication, atomic save,
parallel scaling, native producer acceptance, or complete CRUD coverage.

## Custody and cleanup

Use isolated `~/.cache/litchi-build-0491` compiler output and
`~/.cache/litchi-goal-0491` scratch, with the existing shared CPU lock and
CPU 2 for serial captures. Bind actual source manifests, build commands,
toolchain/flags, executable hashes, environment and retained process output.
Do not reuse an old executable as if built from current sources. Remove
completed compiler/scratch output after validation, retaining only executable
copies required by the final evidence and the raw reports/recipes.

## Pre-freeze findings

The initial current-source release build passed. Its first filesystem pilot
failed before timing because the historical DOCX archive pin no longer matches
the current deterministic producer. Expected SHA-256 is
`a4a2e4921235a6da6b38e31d26ddcca1301909885e37330ab4f83ecc0c4e04f4`;
observed SHA-256 is
`cc3f3f836b0f1568caf24c35a011563eafda546021eba86a833e91b84560491d`.
The failure is retained in `validation/cold-pilot1.*`. An unpinned edit-selector
diagnostic emitted the current catalog and passed its existing semantics;
it is not a replacement cold result or sufficient justification to repin.
The generator-scoped restoration now reproduces the original archive exactly;
`corpus-restoration-proof.json` retains the member-level proof. Further producer
drift fails closed.

Source review also found that the existing full-text lifecycle discards the
actual timed text and checks a fresh read afterward. Formal collection will
require the actual measured text to survive until post-timer comparison against
an independent oracle. This changes the timing boundary: document destruction
remains inside, text verification/hash/destruction outside. Retain an explicit
versioned scope and do not compare these times as identical to old lifecycle
captures. No production library behavior changes are implied.

The original 64 KiB range pilot had no short reads: the largest request was
only 1,424 bytes. The formal short-read arm therefore caps returns at 64 bytes
and requires observed short reads. The delayed 64 KiB arm measures the explicit
latency/bandwidth policy without claiming that its cap is exercised.

The first aligned cold pilot failed its old zero-payload-overlap replay rule.
An independent trace shows the padded EOCD triggers a single 64 KiB metadata
tail search. That search overlaps compressed media bytes without materializing
the media. The final replay evidence must retain those raw overlaps and prove
the exact aligned identity, EOF window and sole probe; ordinary warm opening
retains the original zero-overlap rule. Main-document preparation and the
zero-I/O query remain separate phases.


The aligned verifier priming pass receives the structural tail proof without
inventing a cold-residency observation. The cache count used by that proof is
captured at open: zero successful payload cache loads. The subsequent document
query materializes the main part once; the final replay reports one load.
These are different phase observations and must not be interchanged.

Pilot validation exposed historical report conventions: replay semantic hashes
are SHA-256 of the raw finalized text SHA-256 bytes, while the new timed-text
field is SHA-256 of the actual UTF-8 text. The validator checks both against
the independent fixed text oracle. Elapsed result vectors are sorted; their
sample-order permutation reconciles them with chronological child evidence.
Failed development builds and pilots are retained as diagnostics, superseded
by explicitly named successful final gates and formal captures.

CPU affinity does not reserve a CPU. The shared lock serializes this lane's
captures/checks, not unrelated agents or host activity. Other work on the shared
machine was not stopped. Cold and delayed tails vary substantially across
repeats; the report makes no low-noise or causal regression claim.

All four profiles2 whole-child diagnostics and their canonical verifier pass.
The earlier profiles1 capture succeeded but its verifier used the wrong private
TMPDIR path. Its original helper is retained under development/; profiles2 is
the final recapture after correcting the header/cleanup path reconstruction.

Cleanup is confined to the current user's two 0491 cache roots. Its process
audit checks current-user benchmark/gate processes and cwd/executable/fd
references. Other users and protected systemd/PAM/SSH session daemons are
explicitly listed as excluded where procfs denies inspection; inspectable
descendants remain audited. This does not claim visibility into every host fd.
Every retained 0491 gate has an authoritative terminal receipt before deletion.
