# Change 0480: shared DOCX publication XML

`performance_claim: none; scoped control/candidate ownership comparison`

`claim_authorized: false`

The existing DOCX paragraph-copy publisher now passes the target's immutable
`Arc<Vec<u8>>` to the existing shared OPC overlay method. It previously cloned
the complete XML and allocated a new Arc for the overlay. Both entry points
use the same OPC validation/writer implementation; source recapture, exact
patch application, fingerprints, limits, XML validation, untouched member
preservation, no-op/inverse and partial-sink behavior remain authoritative.

The production delta is the two-line `source-change.patch`. `before-source.txt`
and `source-change.json` retain its exact source identity. No harness, public
API, cache policy, parallelism or iWork source is changed. All 30 previously
read ADR/README hashes remain unchanged.

## Protocol

Fresh control and candidate normal/allocator builds use the same repository
and Cargo paths, Rust 1.98.1, four build jobs, no incremental compilation,
release debug level 1 and frame pointers/unwind tables. Control source is the
0479 measured manifest; candidate differs in one DOCX source file. Each arm's
normal/allocator source manifests agree, and all eight embedded DOCX templates
are copied and bound. All commands run serially under the batch lock.

The frozen sequence is A1/B1/B2/A2: control forward, candidate forward,
candidate reversed and control reversed. Each group crosses normal/allocator,
64/8,192/131,072 source paragraphs and total/six-phase modes. All 48 processes
have 30 measured operations after three warmups, pinned to CPU 2. Eight pilots
have one sample at each size (24 excluded operations). Two separate normal
large-process perf-stat runs have 30 samples each (60 excluded operations).
The copied 0479 corpus manifest and frozen report-schema helper establish the
same independent XML/semantic/physical-member/patch/inverse oracle.

The lifecycle includes source adapter/package construction, snapshot, staging,
commit, sequential publication, digest finalization and owner destruction.
Corpus and oracle storage preexists that region. Phases are separate executions;
phase peaks cannot be summed into a total. Normal and allocator timing remain
separate. Allocated bytes include full reallocation requests and are not
physical memory-copy traffic. `/usr/bin/time` RSS and perf-stat cover the whole
process, including setup, oracles and report teardown.

The full non-iWork goal remains open. This materialized transaction still owns
complete XML and paragraph indexes; removing one clone does not supply the
required explicit-window append capability. Logical repeated append, section
properties, native/cold/range sources and broader CRUD/scaling remain separate.

## Results and review

Incremental operation peak falls by 3,376 / 401,648 / 6,422,768 bytes at the
three sizes, exactly the candidate XML payload plus 40 bytes on this target.
Requested allocation bytes fall by the same amount and allocation callbacks
fall by two. Every allocator lifecycle returns to its entry live bytes.
Output, corpus, source reads and sink accounting agree across all 48 reports.
No total-lifecycle latency or whole-process RSS pair has a positive change
above 5%. Individual phase statistics and repeat drift are retained in
`measurement-review.md`; these observations do not establish a speedup.

`source-review.md` explains why the existing shared OPC route preserves the
publication contract. `measurement-review.md` independently recomputes all
formal samples. The production change is committed as `b3a9c499d`.

Scoped DOCX formatting, tests, all-feature Clippy, rustdoc, crate boundaries
and the strict registry check pass. The earlier workspace-format and
default-feature Clippy failures remain in `validation/`; the per-change
report explains the existing Keynote formatting and feature-gated lint issues.

The completed bundle can be copied independently and verified with
`python3 -B verify.py` from its directory. The verifier reads retained local
artifacts; recorded repository and binary paths are provenance, not runtime
dependencies. `python3 -B verify.py --data-only` checks measurements without
requiring final runtime-cleanup and gate-ledger closure.

All 22 required final gates pass; `rust-validation.json` retains 28 attempts,
including the two disclosed failures and successful development gates. All
14 final evidence tests pass. Full sealed verification and a fresh copied
baseline pass after the authenticated runtime binaries were removed. Eight
independently resealed corruptions are rejected: source reads, output digest,
allocator exit retention, instrumentation, summary arithmetic, PMU arithmetic,
missing required gate and source exclusion. `portable.json` retains each
result and `validation-seal.txt` retains its input seal. Both shared Cargo
caches and both user-owned untracked documents were preserved.
