# 0928 binary callsite mapping (read-only)

Binary: `/tmp/litchi-performance-0925/measurement-target-v4-final/release/litchi-perf-baseline`

- SHA-256: `7138c6c458dc13f1677de3ba3f413120ffc1af155d07848a243a43ebcf335feb7`
- frozen harness source: `/tmp/litchi-performance-0925/measurement-source-v4/tools/perf-baseline/src/lib.rs`
- harness source SHA-256: `7904a6161e7b40a675d61acd9867b31b81a47f049b1ebb5a9979778bf14a373a`
- native record: `/tmp/litchi-performance-0928/perf.data`, SHA-256 `af3121cf9499534ee0d0581e5e9bf0e8a9d5f02855d4936b8738912c27796ff6`
- period symbol/offset export: `/tmp/litchi-performance-0928/perf-script-period-symbol-offset.txt`, SHA-256 `a6bc4005f335bb5cf459d311539542424745030f8937a986319c741f9d3e1da8`

The binary retains `.symtab`, `.strtab`, `.eh_frame` and `.eh_frame_hdr`, but no `.debug_*` sections. The PIE load bias observed by perf is `0x5e06dda9f000`; static addresses below are ELF virtual addresses.

## Static source-bound callsites

Source line numbers refer to the frozen v4 `lib.rs`.

| phase/source | source | static instruction | target / return boundary | interpretation |
|---|---:|---:|---|---|
| expected snapshot | 50757 | `0x105511c` | GOT `0x38b23b0` -> `Package::section_layout_snapshot` (`0x31fc180`); wrapper calls `main_section_layout_snapshot` through GOT `0x38b5e18` -> `0x31fc750` | expected oracle, before sample loop |
| expected publication | 50774–50775 | `0x10558d6` | direct `Package::publish_section_layout_commit_to_stream` clone `0x17e4610`; return `0x10558db` | expected oracle, outside timed loop |
| preparation timer end | 50969 | `0x10574f1` | indirect `Instant::elapsed` via `0x38b0f40` -> `0x37f6cc0` | closes preparation |
| open timer start/end | 50971 / 51006 | `0x105754f` / `0x1057b13` | `Instant::now` via `0x38b0f38` -> `0x37f6cb0`; elapsed via `0x38b0f40` | open bracket; route-specific `from_read_at` calls use GOT `0x38b1458` -> `0x31fb720` |
| query timer start/end | 51008 / 51011 | `0x1057b59` / `0x1057c8b` | now/elapsed; query body call at `0x1057b79` -> GOT `0x38b23b0` -> `section_layout_snapshot`; `0x1057b71` is the preceding `lea` argument setup, not the call | timed query bracket; call return is `0x1057b7f`; underlying main snapshot is `0x31fc750` |
| edit timer start/end | 51013 / 51025 | `0x1057cd1` / `0x1057e5c` | now/elapsed; `set_columns` at `0x1057e2c` -> GOT `0x38b25d0` -> `Edit::set_columns` | timed edit bracket |
| commit timer start/end | 51027 / 51030 | `0x1057ea2` / `0x1057fb5` | now/elapsed; commit at `0x1057eda` -> GOT `0x38b25c8` -> `Edit::commit` | timed commit bracket |
| publication timer start/end | 51033 / 51036 | `0x1058011` / `0x1058131` | now/elapsed; publication at `0x1058059` -> clone `0x17e5310`; return `0x105805e` | timed publication bracket |
| timed publication snapshot | 51034 through publication clone | `0x17e536f` | GOT `0x38b5e18` -> `main_section_layout_snapshot` `0x31fc750` | internal publication recomputation, inside timed publication |
| operation boundary | 51045 onward | immediately after `0x1058131` | source/provider snapshots before line 51096 verification | provider observation is outside publication timer |
| output verification | 51123 | `0x10588c4` | direct `verify_docx_section_layout_output`; return `0x10588c9` | post-operation correctness, outside all phase timers |
| inverse publication | 51128 | `0x1058a74` | direct inverse clone `0x17e6010`; return `0x1058a79` | post-operation inverse/reopen, outside timers |

`0x17e3910` has no direct call in the runner's static disassembly; `0x17e4610` is the expected oracle clone and `0x17e5310` is the timed clone. All three publication clones call the same `main_section_layout_snapshot` GOT target at clone-relative offsets (expected `0x17e466f`, timed `0x17e536f`; unassigned clone `0x17e396f`).

The `CountingSink::write` implementation is `0x10b89d0`; the output vtable's relocated entry at `0x38b2858` points to it. Publication uses indirect trait dispatch through the generic `write_single_part_overlay_to_stream` (timed clone's call `0x17e5832` -> `0x126bc90`), so there is no unique direct `call CountingSink::write` instruction in the runner. A self sample at `CountingSink::write+0x1ef` identifies sink implementation work, but needs a valid caller chain to assign it to a phase.

## What the 0928 record actually contains

`perf.data` was recorded with `cycles:u`, `--call-graph dwarf,16384`, 997 Hz, generic route, 200 samples, CPU 12. Lost samples are reported as zero, but the libdw export emits 2,046 `address range overlaps an existing module` failures and many unwind diagnostics.

The symbol/offset export contains exactly three runner frames:

- `run_docx_source_backed_existing_section_layout_edit_save+0x3dcd` (2 events, 8,352,268 cycles), static `0x10588ad`, the `test` immediately after the pre-verification `bcmp` and before the call at `0x10588c4`;
- `...+0x3de8` (1 event, 3,799,490 cycles), static `0x10588c8`, the final byte of the call instruction to `verify_docx_section_layout_output` at `0x10588c4`;
- `CountingSink::write+0x1ef` (3 events, 12,301,811 cycles), with no valid parent runner/publication frame.

Thus the only directly observed runner IPs are in post-publication output verification, not the timed query/publication callsites. The sink self samples are consistent with publication sink work from the source structure, but the broken/missing ancestors prevent robust timed-vs-untimed attribution. The suspicious chains that show `verify_docx_section_layout_output -> run+0x3de8` should not be treated as control-flow proof: the static address `0x10588c8` is inside the verification call instruction, and the export's unwind failures make parent relationships unreliable.

Flat `perf script -G` output has no self samples in the runner, query, publication clone, `main_section_layout_snapshot`, or sink implementation; these symbols only appear as callchain frames in the DWARF export.

## Limitation and next feasible remedy

The static map is sufficient to label future samples by instruction address: a sample/call return in the interval after `0x1058011` and before `0x1058131`, or a valid chain through `0x105805e`/`0x17e5310`, is timed publication; query requires the `0x1057b59..0x1057c8b` bracket and `0x1057b7f` return from the call at `0x1057b79`. The current record cannot make that attribution because it has no valid phase ancestor chains or self IPs.

A diagnostic-only rebuild should use `-C force-frame-pointers=yes` and frame-pointer call graphs (after checking the resulting assembly/source/binary identity), or benchmark-only non-inlined marker wrappers that actually surround the existing timed query/publication work. Keep the normal runner and timing semantics unchanged; do not infer phase shares from symbol names alone. No rebuild or capture was performed here.
