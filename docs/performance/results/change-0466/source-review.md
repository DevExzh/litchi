# Ordinary dense XLSX commit/save path

The measured Rust/TOML/lock source is unchanged from 0465, as checked against
its 7,032-file manifest. This is investigation of the ordinary eager API,
not the source-backed selected-cell scenario profiled in 0409/0410.

`tools/perf-baseline/src/lib.rs:40699` prepares the source workbook, stages
updates, and reserves the bounded counting sink before timing. The timer
contains `Edit::commit` followed by `Workbook::write_to`. Exact comparison
with deterministic expected output, reopen, complete cell verification and
teardown are outside. The expected output is itself constructed before the
iteration loop. Whole-process profilers include these surrounding operations;
even a commit ancestor can occur in expected-output construction. Ancestor
fractions therefore do not identify retained-sample-only phase times.

Dense-wide has two 256-by-256 sheets, 131,072 numeric cells, 1,311 staged
updates, seven ZIP members and a 384,525-byte source archive. The successful
output has 388,095 bytes, 37 writes, and a 65,536-byte largest write. The old
corpus `uncompressed_payload_bytes` field describes the logical integer grid
(524,288 bytes), not actual inflated ZIP/XML bytes; no instructions-per-XML-byte
or decompression rate is derived from it.

The production path is:

1. `workbook/edit/semantic/transaction.rs:1023`: ordinary commit retains the
   unsigned check and atomic edit validation. `prepare_xlsx_updates` stages
   actions without materializing worksheet stores.
2. `transaction.rs:1363` and `workbook/model.rs:1442`: each edited source
   worksheet's `store()` materializes a complete semantic store.
3. `transaction.rs:1587`, `raw/worksheet/edit/package.rs:16`, and the snapshot
   scan/write owners rebuild the edited worksheet XML.
4. `transaction.rs:1669`: `raw::compact::changed` compacts changed XML.
   `transaction.rs:1671` reparses that result and validates style references
   and staged effects. The web-binding reader also verifies the whole sheet.
5. Candidate publication retains its catalog and invariant checks. The
   existing bounded validated-store handoff deliberately excludes 65,536-cell
   dense-wide sheets: record 0025 rejected unrestricted retention for an RSS
   regression. Raising this bound is not an established optimization.
6. `workbook/model.rs:937` delegates sequential output to the format writer
   and OPC publication writer, including authored-XML audit and compression.

Thus four complete store parses per ordinary timed commit are expected from
the two source and two rewritten sheets. Further parses in the whole-process
profile come from untimed reopen/verification. This count is a source-derived
expectation, not a newly instrumented parse counter.

One concrete allocation lead is `raw/worksheet/codec.rs:667`: `start_cell`
looks up `r`, `s`, `cm`, `vm` and `t` through separate checked attribute scans.
The shared `unqualified_attribute_value` helper scans all attributes and
materializes selected decoded strings. Heaptrack identifies repeated
quick-xml attribute duplicate-check growth, but its totals include fixture,
warmup, oracle and teardown work. A fused cell-attribute view may remove
repeat work; it must preserve duplicate/malformed error ordering, numeric
bounds, entity normalization, qualified extension handling and source byte
preservation. Disabling duplicate checks is not an admissible shortcut.

No production code, parser contract, public API, retention limit, execution
context, native-producer gate or default corpus identity changes in this batch.
