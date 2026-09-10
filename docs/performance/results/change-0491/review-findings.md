# 0491 review findings and disposition

The Rust read-only review found no blocker for the source-backed formal matrix.
It checked timer/drop scope, actual returned-text verification, aligned EOCD
identity and raw overlap accounting, and open-phase cache observation. The root
also ran the complete benchmark library suite, allocator tests, and lint gates.

Driver review led to exact build command/schema/revision binding, executable
identity and corpus-catalog checks, chronological/sorted sample reconciliation,
allocator presence/live-byte reconciliation, fincore executable/version and
filesystem binding, range delay/pacing and counter invariants, and separate
inner-adapter short-read and allocation-increment summaries. Successful pilot
capture alone was insufficient: failures in analysis/header custody were kept
as diagnostics and corrected before formal collection. Protocol/helper copies
for superseded attempts are retained under development/.

Known limitations outside this matrix remain explicit:

- The provider report's allocation-scope prose says normal rows
  contain null; serde actually omits the field. Validators require omission.
- Provider staging can leave private partial scratch if creation/write/sync
  fails before its drop guard is returned. The capture driver refuses a pass
  when its private TMPDIR is not empty. Final cleanup audits these roots.
- Existing eager DOCX lifecycle preparation extracts text before a second
  extraction; that selector is outside this source-backed matrix.
- Existing generic operation-metric attachment accepts unique noncontiguous
  sample indices and can index observations unchecked. Its OPC setter can
  replace the sample order without checking prior identities. These predate
  this batch; current matrices use contiguous identities, and cold validation
  requires the exact permutation. Add focused fail-closed attachment tests and
  bounds/order checks in the next harness correctness batch.

This evidence does not implement borrowed lifetimes, remote transport,
coalesced range access, parallel scaling, publication, native Office round trips,
or the remaining CRUD intersections. Those remain part of the active goal.
