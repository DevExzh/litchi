# 0808 reader schema review

This is a read-only comparison of the developing 0808 readers with the sealed
0806 qualification reports and probe source. The inputs inspected were all 18
files under `change-0806/qualification/`, their `seal.json`, and all six files
under `change-0806/probe-src/`. No replay, build, Cargo command, or workload was
run.

## Compatibility confirmed

- The retained reports use `litchi.pptx.public-workflow-probe-0806.v1`,
  `public-pptx-probe-0806`, and marker
  `litchi-perf-0780-static-mce-capabilities`. The constants in both 0808
  readers match; there is no schema or tool-name mismatch.
- The three modes are `capture`, `commit`, and `lifecycle`. Their timing scopes
  are respectively `Package::opened_presentation only`, `Transaction::commit
  only; package capture and one set_shape_text staging are outside the clock`,
  and `Package::opened_presentation, edit, set_shape_text, commit,
  apply_opened_presentation_commit, and Package::to_bytes`.
- Shape dimensions are tiny `(3, 4)`, medium `(12, 8)`, large `(100, 100)`,
  vendor `(12, 8)`, unicode-vendor `(12, 8)`, and valid-4attr `(12, 8)`. The
  developing `analysis.py` checks the report `shape`; `root_audit.py` currently
  checks only the dimensions.
- Every retained sample has the fields `index`, `elapsed_ns`, `metrics`,
  `source_sha256`, `output`, `allocation`, and `verification`. Capture metrics
  additionally contain `captured_slides` and `captured_shapes_per_slide`.
- The valid-4attr fixture and all nine extension verification fields match the
  exact 0806 values in both readers. Non-valid shapes have those extension
  fields set to `null`.
- The allocation object has these eleven raw fields in addition to `status` and
  `scope`: `allocation_calls`, `deallocation_calls`, `reallocation_calls`,
  `failed_allocation_calls`, `allocated_bytes`, `deallocated_bytes`,
  `live_bytes_before`, `live_bytes_after`, `peak_live_bytes_before`,
  `peak_live_bytes_after`, and `region_peak_live_bytes`. `net_live` and
  `peak_above_entry` are derived values, not raw report fields.
- The six 0808 probe-source files (`Cargo.lock`, `Cargo.toml`,
  `Cargo.toml.template`, `src/allocation_metrics.rs`,
  `src/counting_allocator.rs`, and `src/main.rs`) compare byte-for-byte equal
  to the retained 0806 copies.

## Concrete reader mismatches and gaps

1. **The independent audit misclassifies qualification allocation data.**
   `capture.py` deliberately selects the allocation binary for the
   qualification lane (`capture.py:18-22`), and all 18 retained 0806
   qualification samples contain a complete measured allocation object. Yet
   `root_audit.py:audit()` calls `rows("qualification", 1, 1, 0, False)` and
   `raw_sample()`'s non-allocation branch requires `sample.get("allocation") is
   None` (`root_audit.py:193-194`). This will reject the actual qualification
   report shape; it is not merely an omitted allocation check. Parse
   qualification with full allocation validation, while keeping its allocation
   values out of the decision/guard comparisons if that is the intended policy.

2. **`root_audit.py` does not require report shape identity.**
   `fixture()` validates only `(slides, shapes_per_slide)` (`root_audit.py:98-99`).
   Medium, vendor, unicode-vendor, and valid-4attr all share `(12, 8)`, so a
   report can carry the wrong shape and still pass this check. Match
   `analysis.py:364-367` with `report.get("shape") == shape`.

3. **The root allocation check omits two raw fields and weakens typing.**
   Its integer loop (`root_audit.py:180-183`) omits
   `reallocation_calls` and `failed_allocation_calls`; it only compares the
   latter to zero (`root_audit.py:186`). The retained raw schema includes both,
   and `analysis.py:54-57,345-347` validates both as integers. Keep the four
   `ALLOC_FIELDS` used for guard comparisons as the policy subset, but validate
   all eleven raw allocation counters and peak fields structurally. The raw
   probe also guarantees `peak_live_bytes_after >= region_peak_live_bytes`;
   the root check currently does not assert that ordering.

4. **Vendor fixture identity is under-checked.**
   The exact 0806 vendor fixture has URI list
   `Xttp://schemas.openxmlformats.org/presentationml/2006/main`,
   `Xttp://purl.oclc.org/ooxml/presentationml/main`,
   `Xttp://schemas.openxmlformats.org/drawingml/2006/main`,
   `Xttp://purl.oclc.org/ooxml/drawingml/main`,
   `Xttp://schemas.openxmlformats.org/officeDocument/2006/relationships`,
   `Xttp://purl.oclc.org/ooxml/officeDocument/relationships` and names
   `vP:probeP,vPS:probePS,vA:probeA,vAS:probeAS,vR:probeR,vRS:probeRS`.
   `analysis.py` checks only six non-empty unique strings for vendor and
   unicode-vendor (`analysis.py:281-286`); `root_audit.py` checks neither list.
   The ordinary fixtures also have exact empty URI/name lists, which only
   `analysis.py` currently checks (`analysis.py:287-289`). Either bind the
   exact fixture arrays, or explicitly document why only cardinality is part of
   the oracle.

5. **Capture-only metrics are not checked.**
   The raw capture reports contain `captured_slides` and
   `captured_shapes_per_slide` in `metrics`; neither reader validates them.
   Add a mode-conditional check if the raw schema is intended to be fail-closed.
   Commit and lifecycle reports correctly contain only the elapsed/slides/shape
   metric subset.

6. **`root_audit.py` under-validates top-level and output identity.**
   The 0806 raw contract has a `source` object with `sha256` and positive
   `bytes`, and an `allocator` object. `analysis.py` validates source and
   allocator identity (`analysis.py:371-385`), while `root_audit.py` indexes
   `report["source"]` but does not validate its type, hash, or byte count and
   does not check allocator identity. It also accepts any truthy output hash
   (`root_audit.py:172-174`) without checking hash format or output byte type.
   These are raw-schema omissions, not observed 0806 value differences.

7. **Historical qualification comparison omits allocation and report metadata.**
   `analysis.py:668-685` compares source, fixture, output, and selected semantic
   and extension fields, intentionally excluding timings. The retained 0806
   qualification sample also has the full measured allocation object and the
   top-level mode/shape/timing/allocator fields. If “exact retained raw
   qualification” includes non-timing allocation and metadata, compare those
   fields too; otherwise record the exclusion explicitly so it is not mistaken
   for full raw custody.

The schema/tool constants should remain as they are. The highest-priority fix
is the qualification allocation branch in the independent audit, followed by
the shape assertion and complete raw allocation-field validation.

## Terminal disposition

The unexecuted full-trial reader drafts reviewed above were removed after the
quality stop. They did not analyze or qualify measurements. The retained
`validate_early_stop.py` independently validates the actual eighteen allocation-
instrumented qualification reports, fixture/semantic oracles, partial quality
results, restoration, and three-binary cleanup. `analysis.py --qualification
--check` also replays the accepted before-only audit after cleanup.
