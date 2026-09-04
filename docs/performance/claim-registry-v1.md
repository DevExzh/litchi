# Performance claim registry v1 boundary

[`claim-registry-v1.schema.json`](claim-registry-v1.schema.json) defines the
JSON shape and closed vocabulary of the registry. It is a structural schema;
it does not express relationships between independent fields that JSON
Schema cannot safely derive here.

[`tools/check_perf_claims.py`](../../tools/check_perf_claims.py) is the policy
authority. Structural validation accepts the shape, while strict validation
checks accepted/adverse counts and disjointness, status/code-state
consistency, scope cell sets, and recomputed evidence identities.

The optional `latency_evidence.metric_profile` defaults to `elapsed_ns` for
legacy entries. `publication_ns` binds the custom 0402 summary producer and
the four normal ABBA reports. Allocator reports remain outside this registry
claim boundary and are validated by
`tools/validate_opc_overlay_allocator_abba.py`. The full self-excluding
evidence inventory also remains outside the boundary; its identity is checked
through the retained `evidence-manifest.json` hashes and documented bundle
integrity/audit checks.
