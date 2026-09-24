# Rejected scanner V6 diagnostic evidence

This archive preserves the perf, allocator, and heaptrack diagnostics for the
rejected scanner V6 experiment at historical source `cf4f383c4c203381880d4c2e8445d2616a53d283`.
The candidate diff is `de5495e515335e08336a2f3e0588907bc6b2dcc533edc3907e267351d89c1cf2`.
Use the exact patch hash in the embedded verifier and source manifest as authority.

Extract and verify without a checkout, Cargo, network, or workload:

```sh
tar -xzf evidence.tar.gz
python3 evidence/tools/verify_v6_diagnostic.py --check
```

V3 allocator deltas cover the measured operation. V4 heaptrack counts cover the
whole process, including one setup commit/save and one measured commit/save:
19,014,852 candidate calls and 11,402,262 control calls. V3 heaptrack traced the
`env` wrapper and is retained as invalid historical attribution.

Perf diagnostics include seven lost samples, six lost chunks, and unavailable
kernel symbols. Instrumented runtime and RSS are not normal performance claims.
Release binaries are omitted; their identities, source patch, depfiles, raw reports,
grants, and receipts are retained. This evidence makes no performance gain claim.
The historical source predates the separately accepted cell-tag optimization.

The standalone verifier binds file hashes, source and binary identities, the
single expected scenario per report, allocation scopes, and the recorded caveats.
The archive includes mutation probes and the independent review receipt alongside
this wrapper records the final publication review. Root verification used a fresh
extraction and did not execute a benchmark.
