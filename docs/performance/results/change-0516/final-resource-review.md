# Final resource review

**Verdict: approved for the narrow logical scope.** This review is bound to
the post-fifth-unit candidate manifest
`38007ba3313a420bf1f5823fa8aafb2232d9f3f74bc0874ab416e302263b1df6`; the
preceding fifth-unit snapshot was
`a08915252157449f2aba4c4fc491a0c59d0849aa82df5657a7180bd0d32db3c6`.

The 128 MiB `SpeculativeBudget` claim is limited to compact-output `Vec`
capacity plus newly constructed provisional parser/proof/finalization state.
Existing compactor `NsReader`/writer/probe, `preserve` and normalized-tag
storage, and other baseline allocations remain outside this bound. It is not a
process-wide or whole-worksheet memory limit.

The final fix changes `EventParser::try_charge_event` to retain four payload
units throughout feed (`raw/worksheet/codec.rs:1152-1159`) and removes the
redundant final payload pair. This covers the previously identified End-`c`
overlap: `Parser::consume_event` reaches `finish_cell` at
`:387-392`, where `decode_spreadsheet_text` at `:1053-1059` allocates the
decoded inline `String` while the encoded `String` is still live. The budget
check at `raw/compact.rs:150-168` now already holds the four-unit charge before
that decode. The same four units cover final `t="str"` materialization's raw
capacity, decoded value, and `Text(Arc<str>)` while `Store` is built; no extra
final payload term is needed.

The other reviewed resource terms remain sufficient for this scope:

- `SEEN_ROW_BYTES = 16 * size_of::<u32>()` with a 1 KiB first-table base
  (`codec.rs:1167-1176`, `:1204-1214`) covers `HashSet<u32>` bucket/control
  storage and row records.
- `MERGE_INDEX_BASE_BYTES = 64 KiB` plus `MERGE_INDEX_BYTES = 1 KiB` per
  range (`codec.rs:1273-1301`) covers the current `merge::Index::new` arrays,
  recursive partition vectors, and validation `BTreeMap`/`BinaryHeap` in
  `merge.rs:107-134`, `:176-273`, and `:294-334`. The former approximately
  128-byte term was insufficient; 1 KiB/range is a conservative practical
  allowance for the current standard-library target.
- Decoded parser fields exclude shared-string (`t="s"`) and shared-formula
  records (`codec.rs:1320-1341`), including escaped attribute forms. The
  provisional parser is dropped before another event, and `take_store` calls
  the package callback once (`compact.rs:73-85`), so callback-owned shared
  strings and shared-formula translation growth are left to the exact path.

The reported dense 256x256 numeric worksheet remains eligible under this
logical ceiling because it has none of those excluded dependencies or merge
index state. This is a static resource review; no build, benchmark, or
production-code edit was performed.

## File SHA bindings

| File | SHA-256 |
| --- | --- |
| `crates/litchi-xlsx/src/raw/compact.rs` | `0ac488dad6162f79b5d38713a730dbf6c4269f20f323505009366de7008f39d3` |
| `crates/litchi-xlsx/src/raw/worksheet/codec.rs` | `0515166610208cd8b811aaff03ac880ac85f7bce689fd119f5a2fcbe6fa2f71c` |
| `crates/litchi-xlsx/src/raw/worksheet/model.rs` | `439d6cc4a645891d51b34f7ed6c55098e853ff8662daba7bec2cdd870fc474f5` |
| `crates/litchi-xlsx/src/raw/worksheet/semantic.rs` | `1cd0f76748070b78049033b984eaa3a37ed86fddb9c8d92f9890c467e764cd93` |
| `crates/litchi-xlsx/src/raw/formula.rs` | `5b68bfcd9afaff3ac945336d81eb694adb59ff33ad093145c85e46a0a101658b` |
| `crates/litchi-xlsx/src/merge.rs` | `88d005b2a2540bf33fa69bd550c8cb3b7b7d41ae05c2597f55104992f2e0da1e` |
| `crates/litchi-xlsx/src/cell.rs` | `26bd09c58eda4996e67b0999a7430bcd462fb51430cf2c8601cc0ecac044bd06` |
| `crates/litchi-xlsx/src/column.rs` | `c78029bd2601b4458c138efefd3ff0fdcf7004e50954410d85ad26d5cc2a08ea` |
| `crates/litchi-xlsx/src/raw/worksheet/mod.rs` | `da769ff4964042367945bc827a3c3067b718e8714696cf4bc87dc2777d26722b` |

