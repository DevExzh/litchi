# 0361j — Pages registry error fidelity

- Base: `d0e5542d0e`.
- Contract: 0351 typed resource errors and content-free diagnostics.
- Status: accepted after targeted validation.

The section text-style family previously mapped every package-metadata
resource failure to `RegistryBytes` with fabricated zero observations and
ceilings. It now preserves the codec's actual figures and distinguishes
registry bytes, fields, work, nesting, components, references and additions.
Allocation errors remain typed, and malformed metadata remains `InvalidSource`.
Unknown future resource variants do not acquire invented measurements.

The File consumer already classifies this family's `Error::Limit` as
`limit_exceeded`, independent of its resource kind. A consumer test verifies
the new registry-field kind and preserves malformed-input classification.
The owner regression uses actual metadata inspection: exact input/field
ceilings succeed, one-under ceilings keep their measured counts, work
exhaustion stays distinct, and malformed wire remains invalid input.

Only failure translation, resource labels and tests change. The complete
read/edit/commit/replay implementation following `Run` is byte-identical to
the base; this slice introduces no success-path buffer, scan or validation.
All 209 Pages library tests, the File classification regression and
all-target/all-feature Pages/ops Clippy pass. The [validation receipt](results/0361j/validation.json)
records the implementation comparison, commands and log hashes. The initial
test incorrectly counted one codec pass; its failed receipt is retained, and
the corrected boundary uses the codec's aggregate report.

This closes the queued zero-count registry mapping only. Remaining owner
Snappy propagation, other shared resource conversions, wave integration and
the full workspace gate remain required.
