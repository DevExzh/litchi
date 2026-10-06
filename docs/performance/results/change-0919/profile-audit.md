# 0919 Heaptrack allocation profile audit

This is a read-only audit of the six retained Heaptrack recordings. It does
not change production or probe code and makes no native timing, cache, RSS,
peak-live, lifetime, or physical-copy claim. The raw recordings and the
3.5--5.3 MiB folded-stack reports remain outside the repository at
`/tmp/litchi-performance-0919/`.

## Independent conservation check

The profiled executable is the retained 0917 binary
`/home/zhuhe/litchi-performance-0917/target/release/doc_retention_probe`,
55,146,832 bytes, SHA-256
`2ae48b4f7548c1468c8e7691d56d89b231c4a63f52ed48383f07fe3609c59a13`.
It is an ELF with DWARF `.debug_info`/`.debug_line` and `.symtab`; the symbols
resolve against `/home/zhuhe/litchi-performance-0917/source-base.tar`,
164,352,000 bytes, SHA-256
`61df35b45519c95b43fb4194b0f679c69ce17aa8b39e5c112c0aea9b02f95863`.
Heaptrack 1.5.0 writes the
interpreted `v 10500 3` format. A strict decoder of that format was run on all
six `.zst` files. It validates string, instruction-pointer, parent-trace and
allocation-descriptor references, counts `+` events, and sums each descriptor's
requested size. `-` events are checked but excluded from requested-allocation
totals.

| run pair | whole `+` calls / requested bytes | exact `measure_one` calls / bytes | folded-stack rows / calls |
| --- | ---: | ---: | ---: |
| FloatingPictures, 2 edits | 583,873 / 534,703,235 | 18,014 / 17,369,691 | 1,433 / 18,014 |
| FloatingPictures, 32 edits (Float32) | 900,674 / 1,281,109,930 | 95,057 / 164,233,879 | 1,436 / 95,057 |
| NoHeadFoot, 32 edits (NoHF32) | 150,755 / 123,903,928 | 19,024 / 16,781,690 | 957 / 19,024 |

The second repetition of every row reproduces the same numbers. Each raw
record has exactly one fewer `-` than `+`, matching the one leaked process
allocation reported by Heaptrack. There are zero duplicate exact owner frames
and zero exact owner matches in another module in all six decodes. The folded
reports were produced with the retained command
`heaptrack_print --filter-bt-function measure_one`; every nonempty stack row
contains `doc_retention_probe_0917::measure_one`, and its final decimal count
sums to the filtered call count. The two folded files for each workload are
byte-identical. Raw `.zst` bytes differ between repetitions because recording
metadata and trace numbering are local to each process.

`heaptrack_print -H` is a whole-trace histogram even when
`--filter-bt-function` is supplied. For the Float32 recording it independently
reports 900,674 calls and 1,281,109,930 requested bytes with and without the
filter. Therefore scoped requested bytes must come from the raw `a`/`+` event
decode; the filtered `heaptrack_print` stack report is a call-count
cross-check, not a scoped-byte report.

The exact lifecycle totals reconcile with the independently instrumented 0917
System allocator for Float2, Float32, and NoHF32. This establishes operation
scope and requested-allocation accounting. It does not establish an
allocation-speed or memory-residency result.

## Ownership boundary and classification

The accepted boundary is one exact frame for
`doc_retention_probe_0917::measure_one::h2b5462a44ecd6077` in the retained
binary. Preparation, reference routes, correctness oracles, and post-sample
JSON work are outside the scoped set. Stack membership is inclusive, so a
trace can contain several phase frames; inclusive rows must never be added.
The exclusive ledger applies these predicates in order for each event:

* If `Edit::replace_paragraph` is present, choose `capture_container`, then
  `render_with_layout`, then `commit_candidate_with_rendered`, then
  `RevisionEditor::commit`, then `replace_plain_text`, then
  `target_and_guards`.
* Otherwise choose `Edit::commit`, then `Snapshot::open_bounded`, then
  `Snapshot::edit`, then `probe_boundary_and_other`. A final commit owns its
  nested reopen, so that reopen is not counted as initial open.

This separates the two ownership questions needed for Float32 versus NoHF32:

| stack evidence | interpretation | source boundary |
| --- | --- | --- |
| `RevisionEditor::clone`, with `FcRun`, `PapxRun`, or `Stream` clone descendants under `replace_plain_text` | clone-first mutable candidate staging; the candidate is made before mutation | `package.rs:912`, called at `replace_plain_text:1505` |
| `replace_plain_text`, `rewrite_chpx`, `append_clx_and_cp_tables` | text/CP/CHPX mutation staging and serialization into the isolated candidate | `package.rs:1469`, `2886`, `2988` |
| `source_paragraphs`, `target_spans`, `revisions`, `parse_sprms` | target resolution and tracked-format/dependency guards required before mutation | `body_text.rs:4276`; `sprm.rs:172` |
| `RevisionEditor::commit`, `Arc::from(self.word/table/data.clone())`, and candidate `Editor::clone` | stream handoff and failure-atomic candidate isolation | `package.rs:3020`; `editor.rs:530`, candidate clone at `548` |
| `commit_candidate_with_rendered`, `render_with_layout`, `OleWriter::write_to`, `ReusePlan::emit` | required package check, CFB layout and output publication | `editor.rs:977`; `codec.rs:639`; CFB writer `1631`/`2263` |
| `Package::capture`, `capture_container`, FAT/minifat reads | required reopen, directory capture, stream readback and discovery | `codec.rs:28`, `1157` |
| `Edit::commit`, `Snapshot::open_bounded_with_policy` | final public snapshot reopen and semantic validation | `body_text.rs:2086`, `671` |

Thus a direct `RevisionEditor::clone` allocation is a staging-copy owner even
when its caller stack also contains the public replacement frame. Rendering,
CFB recapture, and final reopen remain validation/publication owners. The
phase names describe allocation ownership, not elapsed-time phases.

For the 32-edit cases, the exclusive requested-byte ledger is:

| exclusive owner | Float32 calls / bytes | NoHF32 calls / bytes |
| --- | ---: | ---: |
| initial snapshot open | 5,307 / 4,005,029 | 521 / 255,802 |
| edit construction | 2,355 / 2,268,807 | 226 / 153,967 |
| staged target and guards | 21,009 / 13,173,117 | 1,585 / 1,082,146 |
| clone/mutation staging (`replace_plain`) | 33,102 / 42,275,408 | 4,660 / 4,477,272 |
| stream handoff (`RevisionEditor::commit`) | 3,328 / 15,523,744 | 1,088 / 1,978,960 |
| candidate commit, other | 2,816 / 802,310 | 1,693 / 296,171 |
| required render | 17,779 / 44,759,310 | 7,227 / 4,965,983 |
| required package capture | 3,872 / 32,756,816 | 1,332 / 2,802,864 |
| final commit and nested reopen | 5,487 / 7,647,386 | 690 / 678,925 |
| probe boundary and other | 2 / 1,021,952 | 2 / 89,600 |
| **total** | **95,057 / 164,233,879** | **19,024 / 16,781,690** |

The largest inclusive subtrees are useful for ranking but are not disjoint:
Float32 `parse_sprms` is 21,213 calls / 8,263,944 bytes,
`rewrite_chpx` is 10,046 / 8,133,040, and `source_paragraphs` is 4,189 /
2,370,318. The largest individual traces are required CFB emission and
recapture stacks; they include output growth and validation storage, so their
requested bytes are not an avoidable-copy estimate.

## Tool feasibility and replay

`/usr/bin/heaptrack` and `heaptrack_print` are 1.5.0. Valgrind 3.26.0 is also
installed, with Massif and DHAT available as `valgrind --tool=massif` and
`valgrind --tool=dhat`; Massif is useful for process heap snapshots and DHAT
for allocation lifetimes/contexts, but neither was needed after the retained
Heaptrack captures and neither was run here. The binary's DWARF symbols are
sufficient for `addr2line` source mapping. Callgrind/perf remain instruction or
sampling tools and cannot turn this report into native timing.

Replay uses the frozen rows in
[`profile-protocol.json`](profile-protocol.json), the recorded capture and
`heaptrack_print` commands in [`profile-results.json`](profile-results.json),
and the raw event method in [`profile-method.md`](profile-method.md). The
read-only decoder source used for the strict check is
`docs/performance/results/change-0830/heaptrack_reader.py` (21,711 bytes,
SHA-256 `fd3e9b795246cce859195ad8946ac0669665782ca219a4e26a874e155782f1ea`).
Its six-line output receipt is retained externally as
`/tmp/litchi-performance-0919-audit/raw-decode.tsv` (300 bytes, SHA-256
`4f6e9e8405645491a245c6bf8cc4699deac9e705ee5e6612fe2ad9b1fe7b229e`). The
independent Float32 whole-trace histogram is
`/tmp/litchi-performance-0919-audit/float32-hist.tsv` (5,271 bytes, SHA-256
`aa6f9081a208e21dc5e7d7d66bc898ec1e38d44e1a78946d0074b13ab28c39bc`); its
`awk` sum is 900,674 calls and 1,281,109,930 bytes. Recreate the raw receipt
with this exact offline command (it reads only retained external captures):

```sh
for trace in /tmp/litchi-performance-0919/r*/heaptrack.zst; do
  python3 -B - "$trace" <<'PY'
import importlib.util, pathlib, subprocess, sys
spec = importlib.util.spec_from_file_location(
    "heaptrack_reader", "docs/performance/results/change-0830/heaptrack_reader.py"
)
reader = importlib.util.module_from_spec(spec)
sys.modules[spec.name] = reader
spec.loader.exec_module(reader)
owner = "doc_retention_probe_0917::measure_one::h2b5462a44ecd6077"
binary = "/home/zhuhe/litchi-performance-0917/target/release/doc_retention_probe"
data = subprocess.check_output(["zstd", "-q", "-d", "-c", sys.argv[1]])
result = reader.parse(data, owner, binary)
print(pathlib.Path(sys.argv[1]).parent.name,
      result["records"]["allocation_events"],
      result["records"]["deallocation_events"],
      result["owner_calls"], result["owner_requested_bytes"],
      result["diagnostics"]["duplicate_owner_frames"]["allocation_calls"],
      result["diagnostics"]["owner_symbol_in_wrong_dso"]["allocation_calls"])
PY
done
```

Recreate the independent histogram receipt with:

```sh
heaptrack_print -f /tmp/litchi-performance-0919/r01-docfloat-e32/heaptrack.zst \
  -H /tmp/litchi-performance-0919-audit/float32-hist.tsv -p 0 -T 0 -a 0 \
  >/tmp/litchi-performance-0919-audit/float32-hist.stdout \
  2>/tmp/litchi-performance-0919-audit/float32-hist.stderr
awk '{calls += $2; bytes += $1 * $2} END {print calls, bytes}' \
  /tmp/litchi-performance-0919-audit/float32-hist.tsv
```

Keep raw `.zst` and folded-stack files external; this summary and the small
external receipts above are the only audit outputs. No production candidate is
authorized by this profile, and the rejected 0036 handoff and 0918 Arc
candidate are not revived.
