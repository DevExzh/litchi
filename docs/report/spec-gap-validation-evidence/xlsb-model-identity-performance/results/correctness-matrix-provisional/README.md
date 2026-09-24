# Synthetic XLDM 140 correctness matrix (provisional)

This receipt is correctness evidence only. It contains no wall-clock,
allocator, RSS, or native XLSB acceptance claim.

Command:

```text
cargo run --manifest-path docs/report/spec-gap-validation-evidence/xlsb-model-identity-performance/harness/Cargo.toml --locked --offline -- --matrix-correctness
```

The run covers 12 table/relationship cases, two endpoint layouts, and five
name profiles (120 coordinates). 110 coordinates pass generation, source
identity, public stage/commit, save/reopen semantic vectors, exact inverse, and
exact no-op gates. The ten `T=64,R=256` coordinates are retained as typed
`limit_exceeded` graph-work refusals under the public default proof ceiling;
they are not counted as successful renames and explicitly report that complete
source proof is unavailable at the graph-work limit. Successful points include
complete source/candidate hashes for OPC parts, relationship owners,
content-types, and XLDM inner members; the verifier recomputes renamed vectors
before accepting their equality fields. The receipt records the separate
`OlapProofLimits::DEFAULT` object, including `max_work=4_000_000`.
The current provisional receipt SHA-256 is
`f427ddd13a8e928faed90c3e6b07ff2158dfb526397b189762d1926c749a798c`.

This receipt was generated before the scaffold commit and is retained as
historical provisional correctness evidence; it is not a sealed source/binary
provenance record. The smoke runner now regenerates the matrix receipt inside
the before/after source and binary capture, retaining
`matrix-correctness.argv.json`, `matrix-correctness.stdout.json`,
`matrix-correctness.stderr.log`, and `matrix-correctness.exit.txt`. There is no
native XLSB Data Model acceptance evidence here.
