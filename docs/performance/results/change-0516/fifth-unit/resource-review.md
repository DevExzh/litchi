# Final resource review

This is a static review of the candidate fifth-unit snapshot supplied by the
coordinator:

`a08915252157449f2aba4c4fc491a0c59d0849aa82df5657a7180bd0d32db3c6`

The review is limited to the provisional parser resource proof. It does not
claim a process-wide or worksheet-wide 128 MiB limit. The `SpeculativeBudget`
accounts for compact-output `Vec` capacity plus the newly constructed
`EventParser` and its proof/finalization allowance. The existing compactor
reader/writer/probe, `preserve` and normalized-attribute storage, and other
baseline allocations remain outside this logical bound, as stated in
`raw/compact.rs`.

## Verdict

**Conditional: one feed-time inline-text overlap still needs an explicit
precharge before this snapshot can be called sound.** The finalization
allowance, merge allowance, seen-row allowance, dependency exclusions, and
logical scope are otherwise acceptable for the current design. A final
approval should be issued after the End-`c` inline decode is charged while the
parser is still provisional.

## Findings

The retained-payload charge in `raw/worksheet/codec.rs:1152-1158` is two
payload units and the EOF estimate at `:1274-1303` adds two more. At the final
guard in `raw/compact.rs:183-188`, this gives a coarse four-unit allowance for
parser-retained text and the short-lived materialization/finalizer copies.
That covers the ordinary `t="str"` path where `semantic.rs:124-132` decodes
the value and constructs `Text(Arc<str>)`; the parser-owned raw value remains
live while `finish_store` materializes cells.

There is an earlier overlap that the final guard cannot cover. For an inline
string, `raw/worksheet/codec.rs:1046-1059` calls
`decode_spreadsheet_text(&cell.inline)` from `finish_cell`. The right-hand side
allocates the decoded `String` while the encoded `cell.inline` allocation is
still live, and only then replaces the field. This runs from
`Parser::consume_event`'s End branch (`:387-392`) during the feed callback
(`raw/compact.rs:165-171`). For `Event::End`, `try_charge_event` sees only the
closing element name (`codec.rs:1190`) and therefore charges only the fixed
512-byte event scratch plus that short name. The two-unit finalization term is
not installed until after compaction returns (`compact.rs:183-188`).

The existing two-unit retained charge cannot by itself prove the simultaneous
old-capacity-plus-new-decoded allocation: a `String` grown through repeated
`try_reserve` calls can retain capacity above its current length, and the
decoded output can require another payload-sized allocation. The minimal fix
is to add an End-`c` precharge, before `feed.consume`, for the pending inline
encoded length/capacity plus the decoded-output bound (or an equivalent
payload-sized extra term), and to run the budget check before the decode. Moving
inline decoding to finalization or decoding in place would also close this
specific gap. No broader policy is required.

The dependency exclusion is correctly based on decoded parser state in
`codec.rs:1326-1347`: `RawFormulaKind::Shared` and cell type `"s"` are detected
after attribute/text decoding, including escaped forms such as
`t="shar&#x65;d"` and `t="&#x73;"`. The callback drops the provisional parser
before the next event (`compact.rs:165-179`), so shared-string package data and
shared-formula translation growth are handled by the exact parser. The
one-shot callback in `WorksheetOutput::take_store` (`compact.rs:73-85`) does
not retry the finalizer and invoke the package owner twice.

The row charge is now explicit: `SEEN_ROW_BYTES` is 16 `u32` units per row,
with a 1 KiB first-table base (`codec.rs:1162-1171` and `:1199-1209`). That is
ample for the `HashSet<u32>` control/bucket storage and allocator metadata on
the current target, in addition to the row record itself. There is no seen-row
resource blocker.

The merge charge is materially stronger than the former approximately
128-byte estimate. `MERGE_INDEX_BASE_BYTES = 64 KiB` and
`MERGE_INDEX_BYTES = 1 KiB` (`codec.rs:1275-1281`) cover the fixed container
startup and, per range, the retained range/index/node/span arrays,
`build`'s temporary lower/spanning/upper partition vectors, and the live
`BTreeMap`/`BinaryHeap` used by `merge.rs:107-134`, `:176-273`, and
`:294-334`. The 1 KiB/range term is a conservative practical allowance for
the current standard-library target; the old 128-byte term was not enough to
cover the map/heap and temporary partition allocations. This review does not
turn it into an allocator-independent guarantee for arbitrary Rust targets.

The coordinator's dense 256x256 numeric worksheet eligibility result is
consistent with these gates: it has no shared-string callback, shared-formula
translation, or merge-index dependency and qualifies under the logical
128 MiB ceiling. The result does not extend that claim to worksheets containing
the excluded dependencies or to baseline process memory.

## File bindings

The following SHA-256 values bind the source inspected in the candidate
worktree. The snapshot identifier above is the coordinator's fifth-unit
binding.

| File | SHA-256 |
| --- | --- |
| `crates/litchi-xlsx/src/raw/compact.rs` | `0ac488dad6162f79b5d38713a730dbf6c4269f20f323505009366de7008f39d3` |
| `crates/litchi-xlsx/src/raw/worksheet/codec.rs` | `eebb281737e008b8cbd79bb8acdc564ac81b57558586f9a3b11677f1f1340481` |
| `crates/litchi-xlsx/src/raw/worksheet/model.rs` | `439d6cc4a645891d51b34f7ed6c55098e853ff8662daba7bec2cdd870fc474f5` |
| `crates/litchi-xlsx/src/raw/worksheet/semantic.rs` | `1cd0f76748070b78049033b984eaa3a37ed86fddb9c8d92f9890c467e764cd93` |
| `crates/litchi-xlsx/src/raw/formula.rs` | `5b68bfcd9afaff3ac945336d81eb694adb59ff33ad093145c85e46a0a101658b` |
| `crates/litchi-xlsx/src/merge.rs` | `88d005b2a2540bf33fa69bd550c8cb3b7b7d41ae05c2597f55104992f2e0da1e` |
| `crates/litchi-xlsx/src/cell.rs` | `26bd09c58eda4996e67b0999a7430bcd462fb51430cf2c8601cc0ecac044bd06` |
| `crates/litchi-xlsx/src/column.rs` | `c78029bd2601b4458c138efefd3ff0fdcf7004e50954410d85ad26d5cc2a08ea` |
| `crates/litchi-xlsx/src/raw/worksheet/mod.rs` | `da769ff4964042367945bc827a3c3067b718e8714696cf4bc87dc2777d26722b` |

No build, benchmark, or production-code edit was performed for this review;
the report itself is the requested review artifact.
