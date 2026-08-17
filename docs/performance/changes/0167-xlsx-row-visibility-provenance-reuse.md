# Change 0167: XLSX row-visibility publication provenance reuse

Date: 2026-08-17

Status: production work elimination retained; release latency claim withheld.
The clean release comparison observed a large publication-phase reduction in
both paired directions, but same-implementation temporal drift exceeded the
predeclared 5% gate. The distributions therefore remain descriptive evidence,
not an acceptance-grade end-to-end speedup claim.

## Mechanism and semantic boundary

The source-backed row-visibility publisher previously loaded the selected
worksheet semantically a second time after the edit had already produced a
source-bound cell-values patch and row snapshot. That publication-only reload
reparsed the worksheet cell store and rescanned all row tags before entering
the mandatory OPC overlay publisher.

The row publisher now delegates its embedded cell-values patch to the existing
crate-private tri-state provenance boundary. A matched lineage and version can
reuse the authenticated before/after snapshots; a mismatched lineage still
returns `PatchConflict`; unavailable provenance retains the conservative
semantic reload and exact-source comparison. Publication still enters the
ordinary OPC overlay path, including cancellation, source freshness, signature
policy, selected-Part validation, bounded sequential output, and raw copying of
untouched members. No public API or dependency edge changed.

An integration regression uses a worksheet larger than the default 8 MiB
payload cache and permits exactly one small selected-member read during
publication. That read is required by OPC overlay publication. A second read,
which the former redundant semantic reload required, fails the test. The gate
also reopens the output, checks the changed row state, and verifies an opaque
member byte-for-byte. Existing tests retain exact no-op, inverse, foreign and
stale source, cancellation, managed-Budget, signature, protection, formula,
MCE, macro, relationship, and partial-output behavior.

## Verification

The frozen candidate passed:

- 765 `litchi-xlsx` library tests;
- 16 source-backed row-visibility integration tests;
- 30 source-backed cell-values integration tests;
- strict focused Clippy with `-D warnings -D deprecated`;
- formatting and diff checks;
- independent read-only semantic review.

`RUSTDOCFLAGS=-D warnings cargo doc -p litchi-xlsx --no-deps` remains blocked
by existing unrelated private/broken intra-doc links in conditional formatting,
drawing, named sheet views, page setup, pivot charts, styles, volatile
dependencies, workbook, and XML maps. None is in this change's production or
test paths.

## Clean release protocol

Control `7dc05de69d9e5e70f827be739dd578bc400ef23a` and candidate
`4b156db82baa58988a04699fb026cba5b8e8a04c` were built from clean detached
worktrees. Binary SHA-256 values are respectively
`1996336bdc8b0d29d6262b0c39c017709bb4d996f33ef7192f501bbf0afba54e`
and
`b4b0e7e034523b3218c01acc0e9aaf82f19af4ad0e39127db056b86e48d6e4a2`.
Every raw record reports `git_worktree_dirty: false`, CPU affinity `2`, Rust
1.95.0, the Rust system allocator, and the AMD EPYC 9575F host.

Fresh processes ran strictly `A1 control, B1 candidate, B2 candidate, A2
control`. Each leg used 20 warmups and 500 retained samples for source-backed
hide-one and exact-256 unhide over the deterministic medium and large
one-sheet media-rich corpora. The four raw records therefore retain 8,000
top-level samples plus the same-cardinality phase vectors. p50 uses the integer
midpoint of the two central sorted values; p95 and p99 use nearest rank.

## Descriptive result and rejection boundary

Positive numbers below mean that the candidate was lower than the control.
Publication is the separately timed sequential publication interval.

| Shape / operation | Pair | publication p50 | mean | p95 | p99 |
|---|---|---:|---:|---:|---:|
| medium / hide one | A1 -> B1 | 56.54% | 56.06% | 52.55% | 50.42% |
| medium / hide one | B2 -> A2 | 58.10% | 57.67% | 56.17% | 57.13% |
| medium / unhide 256 | A1 -> B1 | 57.45% | 56.97% | 53.83% | 53.43% |
| medium / unhide 256 | B2 -> A2 | 59.67% | 59.65% | 59.29% | 59.64% |
| large / hide one | A1 -> B1 | 62.70% | 62.22% | 59.02% | 57.83% |
| large / hide one | B2 -> A2 | 66.91% | 66.62% | 66.15% | 66.55% |
| large / unhide 256 | A1 -> B1 | 60.90% | 60.77% | 56.87% | 54.78% |
| large / unhide 256 | B2 -> A2 | 67.39% | 67.52% | 66.97% | 68.23% |

The publication direction agrees for every reported case and statistic. The
complete timed workflow does not satisfy the same rule: the first medium
hide/unhide p99 comparisons are 6.95% and 2.69% slower. More importantly, the
5% same-implementation gate fails: maximum absolute drift is 34.80% for
control large/unhide publication p99 and 10.23% for candidate medium/hide
complete-workflow p50. The complete operation therefore
has no accepted latency or tail-latency result. The publication percentages
are retained to show that the intended work moved in both pair directions,
but they are not promoted to an acceptance-grade speedup claim until a stable
repeat closes the drift gate.

Logical source topology is unchanged: medium records retain 204 logical
`ReadAt` calls and one selected-worksheet overlap; large records retain 209
calls and six overlaps. Removing the semantic reload also removes 13
source-version calls per sample (572 -> 559 medium and 574 -> 561 large).
These are owned-source logical counters, not physical I/O.
`/usr/bin/time -v` whole-process maximum RSS is 244,916 / 231,516 /
229,548 / 236,868 KiB in A1/B1/B2/A2 order. The sidecars include corpus setup,
warmups, all cases, and untimed gates, and A2 records additional filesystem
input. No allocation, RSS, physical-I/O, decompression, recompression, copied-
byte, cold-cache, or real-producer improvement is accepted.

## Artifacts and reproduction

The [summary](../results/xlsx-row-visibility-provenance-0167-summary.json),
[primary statistics](../results/xlsx-row-visibility-provenance-0167-primary-stats.tsv),
[comparisons](../results/xlsx-row-visibility-provenance-0167-comparisons.tsv),
and [manifest](../results/xlsx-row-visibility-provenance-0167-manifest.json)
bind the four compressed raw schema-1 JSONs and four `/usr/bin/time -v`
sidecars. The raw vectors, revisions, corpus/output/semantic hashes, source
counters, and all refusal/preservation gates remain independently
recomputable.

```sh
taskset -c 2 litchi-perf-baseline \
  --case xlsx_source_backed_row_visibility_edit_save,\
xlsx_source_backed_row_visibility_batch_edit_save \
  --xlsx-row-visibility-shape medium,large \
  --warmup 20 --samples 500 --json RESULT.json
```

This change does not alter the four selectors or the selectable/default matrix
counts from change 0166. It does not cover eager row publication, formulas,
new row owners, structural row insertion/removal, multiple worksheets,
third-party producers, filesystem cold state, or remote sources.
