# 0922 candidate selection: owner-level borrowed FIB lookup

The selected 0922 candidate is a private borrowed FIB metadata helper in
`crates/litchi-doc/src/parts/fib.rs`, consumed by the saved-selection paths in
`crates/litchi-doc/src/tracked_revision/package.rs`. It is a bounded source
change; this note does not claim a latency benefit.

## Evidence and attribution boundary

The retained 0919 `docfloat`, 32-edit Float32 top-byte rows identify full
`WordDocument` suffix copies in the edit path:

| trace | requested bytes | relevant stack boundary |
| ---: | ---: | --- |
| 84105 | 4,256,256 | `FileInformationBlock::parse_at` → `apply_saved_selection_splice` |
| 83930 | 4,069,239 | `FileInformationBlock::parse_at` → `saved_selection_splice` → preparation |
| 83974 | 4,069,239 | `FileInformationBlock::parse_at` → `saved_selection_splice` → prepared splice |

These three FIB rows sum to **12,394,734 requested bytes**. Trace 84108 has a
same-sized 4,256,256-byte row, but its stack is `RevisionEditor::commit`
handoff and it is excluded as an unrelated attribution trap. The rows are
allocation evidence from the retained Heaptrack profile, not a latency,
physical-I/O, RSS, peak-live, lifetime, or speedup result.

`FileInformationBlock::parse_at` currently copies the whole suffix with
`data.to_vec()` before constructing its shared representation. The saved
selection consumers need the validated WSS table pointer and range, not
`raw_data()`. The candidate may add an owner-level borrowed view/helper in
`fib.rs` and use it only from the relevant `RevisionEditor` saved-selection
preflight/splice paths. Public `parse`, `parse_at`, `raw_data`, and the prior
0596 compatibility behavior remain unchanged. The helper must preserve the
existing magic, minimum-size, table-count, pointer-range, truncation, and
typed-error behavior.

The 0914–0916 DOCX measurements remain the next alternative: query and
publication accounted for about 90% of the 0914 clock, and 0915 attributed
about 111.6M IR to fingerprint/SHA work in each route. That work needs a
source-freshness/provenance proof before it can become a candidate. The 0913
PPTX owned-source read remains rejected because its controlled timing gate
failed and the 0912 pilot did not establish native attribution.

## Deferred-emission future work

The 0919 ledger measured 44,759,310 requested render bytes and 32,756,816
requested package-capture bytes for the 32-edit Float32 case. A future private
text-edit batch could investigate deferring physical CFB emission, but this
packet records no feasibility or speedup claim. Every operation would still
need its typed target/dependency/capacity/tracked-range guards, clone-first
failure atomicity, package stream/count/total-limit checks, and current refusal
errors. It would have to flush, render, reopen, recapture, and rediscover
objects before resource operations or any `finish_candidate` path, then repeat
the full public-reader, protection/signature, source-identity, and patch
validation at final finish. `retained_render_bytes()` and per-operation
publication errors are observable, so simply disabling per-edit validation is
not an acceptable design. Direct `RevisionEditor` mutators also commit per
call; any first implementation would need a deliberately scoped higher-level
edit transaction.

No benchmark was rerun and no production source was changed for this scouting
record.
