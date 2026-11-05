# 0362ul137 — DrawingML v2 consumer wiring

This source-only follow-up completes the current-source consumer boundary for
the DrawingML raw policy.  The v1 paired runner, validator, compile receipts,
manifests, and historical source remain unchanged.  The new versioned runner
uses `prepare_drawingml_raw_inputs_v2`, validates its v2 preparation receipt,
and materializes the reviewed generic process manifest with all eight probe
copy paths and all three role-specific lock paths.  The v2 validator keeps the
v1 child-probe schema but writes a v2 summary only after checking the bound
receipt; the v2 compile adapter applies the same policy gate to compile
receipts without relabeling an existing v1 receipt.

The checked-in manifest records the five raw lanes, the separate family route,
four detached-owner routes, the reviewed `baseline,candidate,candidate,baseline`
order, warmups, samples, block count, minimum/target durations, separate
allocator route, fixture handoff, and source-fence metadata.  The trusted
`472b...` `lib.rs` blob, role locks, candidate manifests, vendor roots, and
candidate binding remain owned by the v2 preparer.  `CANDIDATE_BINDING` stays
`None`; the runner and compile adapter refuse before resolving execution
paths until the coordinator commits the final source pin.

Focused source-only checks:

```text
PYTHONDONTWRITEBYTECODE=1 python3 -m unittest \
  tools.test_drawingml_raw_policy_v2 \
  tools.test_drawingml_raw_qualification_v2 \
  tools.test_drawingml_raw_qualification \
  tools.test_drawingml_compile_driver
```

No Cargo, native process, allocator capture, timing, or performance result is
claimed by this record.
