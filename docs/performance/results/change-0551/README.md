# 0551: XLSX layout handoff feasibility

`performance_claim: none; retained-profile reanalysis and source audit`

This batch narrows the 0550 layout-scan opportunity to a compact, private
per-cell source proof. It does not implement or admit that proof. A full
retained `Layout` would still construct tags and primary-span allocations for
every cell during planning, including plans that become exact no-ops. A
row-only proof followed by edited-row reparsing would revisit almost the
entire dense/sparse corpus in the registered 1% workload.

The immediately preceding turn only acknowledged the user's OLE2/OOXML
priority and checked clean state, so it made no progress. The earlier completed
0550 batch (`41b55e8e8`) was progress and supplies the retained measurements. This
batch adds reproducible scanner partitions and corpus coverage calculations,
and records the residual proof obligations before implementation. ODF remains
deferred; iWork is excluded and the broad goal remains active.

## Evidence and scope

`inputs.json` binds the source revision, analyzer, eight measured raw dumps,
0550 profile report, eight native reports, prior source/ADR inventories,
plan and seal. `analysis.json` is deterministic; replay it from the repository:

```bash
python3 -B docs/performance/results/change-0551/analyze.py
```

The replay hashes all 8,593 existing source inventory entries and 30 accepted
ADR/index documents, classifies each dump from its positive incoming owner
edge, and verifies scanner self plus direct-child costs against its unique
positive incoming rewrite edge. It excludes lifecycle calls. The corpus
calculation ports the hash-bound harness inventory and evenly spaced update
algorithm; retained native corpus counts independently check its totals.
It is not a run of a proposed row-reparse implementation.

| Shape | Scan Ir, R1 / R2 | Named reader + namespace Ir, R1 / R2 | `start_cell` Ir, R1 / R2 |
| --- | ---: | ---: | ---: |
| medium | 70,977,241 / 70,990,409 | 24,406,523 / 24,406,376 | 20,976,367 / 20,984,980 |
| dense-sparse | 137,507,562 / 137,536,341 | 47,067,631 / 47,067,245 | 42,199,101 / 42,215,736 |
| noncompact | 91,137,124 / 91,137,400 | 26,469,145 / 26,466,528 | 38,892,525 / 38,892,935 |
| vendor-extension | 71,003,018 / 70,977,289 | 24,406,713 / 24,406,818 | 20,991,658 / 20,992,603 |

The named reader plus namespace category includes `read_event_impl`,
`resolve_event`, and `process_event` only. It accounts for 29.04–34.39% of
scanner Ir, or 15.95–18.31% of exact commit Ir. Other reader/resolver operations
remain explicitly under other children. Inlined scanner work also remains in
scanner self. These are attribution partitions, not removable fractions,
elapsed shares, allocation counts, or a forecast of fusion speedup.

| 1% shape | Edited cells | Edited rows | Cells in edited rows / all planned cells |
| --- | ---: | ---: | ---: |
| medium | 93 | 93 | 4,464 / 9,216 |
| dense-sparse | 178 | 142 | 16,769 / 17,792 |
| noncompact | 93 | 93 | 4,464 / 9,216 |
| vendor-extension | 93 | 93 | 4,464 / 9,216 |

All 128 rows of the dense sheet are edited. A row-only design that reparses
edited rows would therefore inspect rows containing 94.25% of the dense/sparse
corpus's cells. One-cell controls are also retained: 48/2,304 planned cells for
medium/noncompact/vendor and 128/16,384 for dense/sparse. This rejects row-only
reparsing as the selected approach to eliminating whole-sheet cell work; it
does not establish a measured latency regression for such an implementation.

## Selected implementation direction

Produce a compact per-cell offset proof during the existing eligible source
traversal; materialize ordinary tag details only for cells that actually
change. Share the raw parser's already resolved address when feasible, rather
than parsing A1 references again. Keep row envelopes for byte copying and
unchanged-row readback provenance, with cell offsets for sparse edits within
rows. Global structure/ordering checks must still match the ordinary scanner.
The design must avoid retaining the complete `Layout` or reviving the rejected
0527 row arena. A precise representation and byte cap remain to be designed
and measured; no retained-memory estimate is claimed here.

The source-backed validator is more restrictive than the generic raw parser:
its element and parent allowlists exclude protection, validation, merge,
extension and foreign-element cases. A successfully completed validation can
prove their absence. However, attribute allowlisting is not equivalent to
`wire::tag`/`cell_tag` decoding and normalization: namespace declarations are
skipped by the allowlist and otherwise permitted values are not all decoded
there. A proof must cover latent scanner errors on unchanged tags as well as
changed cells. Failure to do so would silently broaden accepted commits.

The implementation must satisfy the field and refusal inventory in
`design-review.md`. That inventory is a source audit, not a completed
differential equivalence proof. The next batch must first implement and test
the proof against the unchanged scanner, including deliberately malformed
attribute values, before any performance admission campaign.

## Admission requirements and limits

- Carry original input offsets only on the existing byte-identical UTF-8,
  marker-free eligible route, within its input and event caps. Expose decoder,
  namespace resolver and start/end positions privately if needed; do not
  reuse transformed MCE-buffer positions as source positions.
- A provisional builder failure drops its scratch and declines the fast path.
  It must not change the historical planning error or surface a commit-only
  scanner error during planning. Keep the complete scanner as fallback at
  commit for missing/incomplete proof, unsupported structures/actions,
  source-identity/version mismatch, new rows and shared formulas.
- Check ordering, bounds, inferred addresses and global dimension expansion;
  checking only edited cells is insufficient. Bind the finished proof to the
  immutable source owner. Preserve exact no-op sharing, patch inverse and
  clone/recommit behavior without hidden mutation of snapshot state.
- Bound metadata before growth with fallible reservations; preserve execution
  and cancellation checks. Measure planning/commit overlap and post-commit
  retention, including before snapshots held by patches. Do not assume an
  event-count cap is a metadata-byte cap.
- Preserve full emitted-output validation, independent changed-cell readback,
  unchanged-byte provenance and atomic multi-sheet commit construction.
- Freeze fresh matched baseline/candidate native and allocator measurements
  covering planning + staging/commit + publication, one-cell/1%, all four
  shapes, managed execution, exact no-op, refusals and cap boundaries. Include
  retained-state memory, not only incremental commit-region peaks. Keep all
  adverse rows and require representative workflow gains.

No Rust or harness source changed. No fresh timing, allocation, hardware,
native Office, fuzz, concurrency or full Rust test run is claimed. The
runtime baseline remains 0550. If the compact proof cannot avoid equivalent
work or pass memory/error-order gates, close that candidate and return to the
measured XML-validation branch rather than weakening its contract.

## Verification and cleanup

Deterministic analysis replay and Python syntax checks pass. Fresh crate
boundary and strict performance-claim checks pass; exact commands and outputs
are retained in `metadata-checks.json` and its stdout/stderr files. The prior
0550 bundle's complete seal also passes. Independent profiler and source
audits are recorded in `profile-review.md` and `design-review.md`.

No build target, scratch checkout, temporary capture directory or Python
bytecode cache was created for this batch. The retained bundle contains
evidence only. `documentation-manifest.json` binds the final narrative files;
`SHA256SUMS` binds the bundle. Run the final read-only verification with:

```bash
python3 -B docs/performance/results/change-0551/verify.py
```

The verifier replays the analysis and checks both seals, source/ADR custody,
documentation, metadata receipts, change scope and bundle immutability. It
does not rerun Rust tests or reinterpret historical timing as fresh evidence.
