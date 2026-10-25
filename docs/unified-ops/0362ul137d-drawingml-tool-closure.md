# 0362ul137d — DrawingML v2 imported-tool closure

This source-only follow-up preserves `c7925ddf260c3003bc8f560dc82fc41a782b2908`
and closes the last protocol-tool identity gap.  The v2 protocol imports the
reviewed legacy DrawingML compile helper, so
`tools/native_qualification/drawingml_compile_driver.py` is now part of the
trusted protocol-tool identity set and is rehashed at every generic-runner
execution boundary.  A mutation of that imported helper is refused before
policy receipt validation or workload execution.

The source-only suite passes 44 tests, including the imported-helper mutation
regression.  No candidate/source pin is self-referenced, and no Cargo, native,
allocator, or performance command was run.
