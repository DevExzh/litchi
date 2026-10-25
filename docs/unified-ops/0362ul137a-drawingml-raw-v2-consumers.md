# 0362ul137a — DrawingML v2 execution-boundary repair

This source-only follow-up closes the independent review findings against
`6387be98d375b6e1c625f6522432df392be19cf0`.  It keeps the candidate binding
pending and does not run Cargo, a native probe, an allocator capture, or a
performance workload.

The v2 runner now requires the exact reviewed route and owner profile sets;
omission or reordering cannot silently reduce the qualification scope.  The
generic runner rechecks the v2 preparation receipt at its direct execution
boundary, including the exact probe, dirty-path, and role-lock routes.  All
three versioned entry points bootstrap the repository `tools` import path, so
absolute-path execution has the same policy behavior as module execution.

The compile adapter validates the complete v2 preparation receipt before
calling the reviewed legacy Cargo driver.  Its v2 compile receipt binds the
role source identity and preparation artifact, and records the fixed
historical `472b...` source, all three role locks, eight probe-copy paths,
candidate manifests, vendor roots, and toolchain paths.  Receipt consumption
rejects a relabelled v1 receipt or any changed policy projection.  The v2
preparation path also materializes and records the selected toolchain files.

Focused source-only validation:

```text
PYTHONDONTWRITEBYTECODE=1 python3 -m unittest \
  tools.test_drawingml_raw_policy_v2 \
  tools.test_drawingml_raw_qualification_v2 \
  tools.test_drawingml_raw_qualification \
  tools.test_drawingml_compile_driver
```

The run passed 43 tests.  No historical measurements or source pins were
changed.
