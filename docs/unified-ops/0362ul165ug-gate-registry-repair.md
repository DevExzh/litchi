# 0362ul165ug — gate registry and benchmark coverage repair

The current 7c remaining-checks diagnostic exposed three real registry
omissions and five benchmark-matrix coverage omissions. This leaf repairs the
checked-in policy and its tests without changing production code or weakening
any gate.

## Boundary registry

`litchi-ooxml-common` has a test-only `soapberry-zip` dependency. It is now a
canonical workspace edge with an explicit `dev_only_edges` annotation. The
`litchi-ops` feature graph already declares `litchi-vba` as an optional normal
workspace dependency; the canonical edge is now present in the policy. The
boundary test's exact dev-edge inventory is migrated with the policy.

The policy remains strict: unclassified edges still fail, a dev-only edge used
outside dev still fails, and annotations must refer to a canonical edge.

## Benchmark coverage

Records `0362ul165e`, `0362ul165g`, `0362ul165j`, `0362ul165k`, and
`0362ul79f` contain protocol, compiler, diagnostic, or tooling validation
receipts. They do not contain a leaf-versus-unified timing or allocation pair.
Each is therefore listed in `evidence_without_rows` with a record-specific
reason. No performance claim is introduced, and the native qualification rows
remain required.

## Checks

The policy and curated JSON parse successfully. `tools/unified_bench_matrix.py
--check` passes, as does `tools.test_unified_bench_matrix`. Focused boundary
policy tests for exact dev-edge registration, promotion rejection,
annotation rejection, and canonical-edge requirements pass. Cargo, rustc,
native qualification, and performance commands were not run while the
coordinator's bounded diagnostic was active.
