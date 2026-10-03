# 0883 Report-aware benchmark metadata oracles

The performance harness builds against the current leniency APIs again. At
`ca2b1f5fdb`, the removed `properties::read` API left two XLSX metadata oracles
uncompilable. Both now call one harness-local helper using `read_with_report`.
A reported core-properties repair is a validation error; absent properties
still produce default metadata. Production code and measurement intervals are
unchanged. This repair establishes no speedup or new performance baseline.

## Validation boundary

The eager OPC metadata oracle remains outside the timed source/facade operation
for XLSX bytes-root and filesystem-root cases. The helper checks the report
returned by core-properties reading. It does not introduce a separate policy
for package-ingress tolerance reports.

The new regression test checks absent properties, a clean title-bearing part,
and rejection of an orphaned core-properties part that the lenient reader can
otherwise adopt. The existing filesystem integration test checks both open
and open-lifecycle selectors, two samples each in four distinct fresh children,
against the unchanged source and semantic SHA-256 pins.

## Checks and retained evidence

All five [final command receipts](results/change-0883/final/) passed:

- Harness formatting.
- Offline, locked all-feature/all-target compilation.
- The focused metadata oracle test: one passed.
- The existing filesystem integration test: one passed, four child samples.
- Harness library Clippy with warnings denied.

The [summary](results/change-0883/summary.json) binds the tested source hashes.
The [initial observation](results/change-0883/initial-observation.json) records
the removed-API failure as a tool-output summary. The retained initial
`check.log` records a development-test compile error: `Metadata` does not
implement `PartialEq`. The final test compares the existing serialized metadata
digests instead. No failure was converted into a passing receipt.

To repeat the checks into a new directory:

```sh
python3 -B docs/performance/results/change-0883/check.py --output /tmp/litchi-0883-recheck
```

This is scoped harness qualification, not a full-workspace quality gate.
Current-source semantic workload measurements remain necessary before further
optimization. The broader non-iWork performance goal remains open.
