# 0513 follow-up: next XLSX production candidate

Change 0513 is a harness-only enabler at revision
`53a7c4a523e8518a7150d259ccf3435cb134e4a4`. It does not change XLSX
production behavior. The diff adds operation allocation observations to the
existing commit and commit/save cases, adds the `#[inline(never)]`
`xlsx_commit_save_operation` helper around `Edit::commit` plus sequential
`write_to`, and adds focused tests for metric alignment, unavailable normal
allocator status, sink promotion, exact output, semantic commit contents and
short-sink errors.

The plan now measures four existing XLSX cases across tiny, medium and
dense-wide shapes with 100-sample/3-warmup serial ABBA native captures. A
separate allocator binary collects 10-sample/1-warmup operation regions, and a
three-sample Callgrind lane toggles only the exact save helper. Setup, expected
output construction, sink reservation, reopen/oracle work and caller drops
remain outside the operation boundaries as specified by `plan.json`. Normal
binaries publish allocator status as unavailable; they must not fabricate zero
vectors.

The current 0512 operation-scoped profile selects a larger parser/Layout
candidate. Its disjoint direct commit children are source Store parsing
(25.80%), worksheet rewrite (27.18%), changed-worksheet validation parsing
(25.77%) and changed XML compaction (20.38%). Nested eager-parser and snapshot
rows overlap those children. The shared-formula loop and plain-cell tag work
are too small to stand in for this larger traversal work. These percentages
are current attribution evidence from 0512, not latency or memory claims;
they must not be added together with nested rows.

The next production candidate should be a private, conditional fusion of the
source eager worksheet parse and the lossless snapshot Layout scan for the
ordinary MCE-free path, followed by reuse of that ephemeral Layout by the
worksheet rewrite. The fused path should use one namespace-aware XML event
reader to feed the semantic parser and snapshot scanner, then pass the Layout
to a private `rewrite_with_layout`-style entry point. It should be restricted
initially to ordinary grid edits on a source for which `process_ooxml` returns
the original bytes. MCE-transformed input, merge-container rewrites and other
cases that cannot prove original-byte span identity should retain the existing
separate passes.

The fusion must defer any snapshot error until the point at which the current
transaction would have run `rewrite`:

* The source parser, shared-formula resolution, Store materialization and
  style validation must retain precedence over snapshot errors.
* Action projection must still be able to reject an invalid or ineffective
  edit before a snapshot error is exposed. A no-op or metadata-only edit must
  discard the temporary Layout and preserve its current behavior.
* If the scanner finds an error while the parser continues, retain that typed
  scanner error only as deferred state; a parser, x14ac or style failure must
  win. Do not publish a partial Layout or a partially initialized Store.
* The `x14ac::may_contain_descent`/capture ordering and the rejected-parser
  fallback in `raw/worksheet/mod.rs:45-50` must remain unchanged. The MCE
  marker branch in `mce/codec.rs:530-637` must continue to use the existing
  transformed-input path, because its offsets do not identify source bytes.

The snapshot scanner’s lexical and error contracts remain part of the fused
boundary. In `raw/worksheet/edit/codec/snapshot/scan.rs:952-1002`, address
checking runs before `wire::cell_tag`: the checked raw attribute pass rejects
malformed syntax and duplicates, decodes unqualified `r` as encountered,
then coordinate parsing, row agreement, inferred-column limits and
`row.last_column` mutation occur before tag-phase decoding of unknown or
qualified values. Invalid element/attribute UTF-8 remains a tag-phase error;
qualified `x:r` remains a lossless attribute rather than an address. Preserve
attribute order and normalization, unknown attributes, exact untouched source
spans, regenerated coordinates for changed cells, and 0472’s `Option<Tag>`
elision for plain cells. Existing formula, merge, namespace, MCE,
malformed-tail, duplicate and source-preservation tests must exercise both the
fused and fallback paths.

The candidate must keep all source and publication boundaries from ADRs
0001/0003/0005/0006/0011/0018: typed fallible errors, immutable atomic
publication, source provenance and lossless unsupported content, checked
limits and reservations, and XLSX/OPC ownership. Do not add a persistent
Layout cache, unsafe code, public API changes or a larger Store handoff. The
4,096-cell/1 MiB validated-Store handoff and post-compaction grid/readback
validation remain in force. Exact output bytes, patch/inverse behavior and
no-op identity need differential coverage against the unfused implementation.

The new allocator fields require a precise memory interpretation. For every
aligned allocator sample:

| Field | Meaning | Use in the production decision |
| --- | --- | --- |
| `region_peak_live_bytes` | Absolute process live-byte maximum observed during the operation region, including entry live bytes and callback-ordered activity | Report separately; it is not operation-owned or incremental memory |
| `live_bytes_before` | Absolute process live bytes at region entry | Subtract from the region peak per sample |
| `region_peak_live_bytes - live_bytes_before` | Incremental peak live demand above the operation entry baseline | Primary operation-overlap comparison for the fused temporary Layout |
| `peak_live_bytes_before` / `peak_live_bytes_after` | Absolute process-lifetime high-water snapshots | Do not label either value as the operation peak |
| `live_bytes_after` | Absolute live bytes at region exit | Check endpoint bounds; it is not a peak |

The region begins before the operation timer and ends after elapsed capture.
For commit/save, the exact helper returns `Commit` before caller teardown, so
that retained result is intentionally inside the measured region while
reopen, verification and drops remain outside. The allocator observer does
not reset process counters, does not establish RSS or physical-memory use,
and does not expose hidden realloc-copy overlap. A candidate memory result
must therefore report both the absolute region vector and the per-sample
incremental difference, with aligned sample identity; neither is a document
peak or a general RSS claim.

Before retaining the fusion, compare the fresh control/candidate allocator
vectors and the native ABBA rows, while retaining all adverse rows and
instrumentation caveats. The three-sample profile is useful for direct-call
and traversal attribution only. If the Layout/RawCell overlap is materially
larger or the correctness/error differential fails, fall back to a separately
measured snapshot cell address/tag pass reduction; do not combine both ideas
in one attribution batch. 0471’s early rewrite-buffer release remains
rejected, and 0472’s plain-tag elision remains retained rather than being
recounted as new work.

This follow-up does not complete the broader performance goal. OLE2 and OOXML
remain the active priority until their optimization goal is complete; ODF work
is deferred, and iWork remains excluded. Small queue or harness-only changes
must not be reported as full OLE2/OOXML completion.
