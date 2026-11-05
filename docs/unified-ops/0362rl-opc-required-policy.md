# 0362rl — OPC required source policy and consumer migration

Status: source migration complete on `uo2/opc-required-policy`, refreshed onto
`d98a99f2dd`. This lane closes the Office-specific OPC raw ingress and alias
debt identified by 0362rd. It does not rewrite generic CFB or ZIP APIs and it
does not change the separate DOCX raw helper boundary.

## Canonical public surface

Every public OPC ingress now takes a required `SourceLimits` argument at its
canonical name. `OwnedPhysPkgReader::{open,from_bytes,from_reader}` and
`PhysPkgReader::new` no longer have limit-less or `_with_limits` twins.
`OpcPackage::{open,from_reader,from_vec,from_shared_vec,from_bytes}` follow
the same rule. Content-type and relationship source/replace/fragment APIs,
catalog probing, relationship parsing, and `validate_read_at` likewise take
the required policy at their canonical names.

`SourceBackedPackage::from_path` requires `SourceLimits` and keeps the exact
read, default-cache, unmanaged behavior. The owned constructors are the
single canonical policy-preserving entries:

```text
from_reader(reader, limits, cache_limits,
            execution_context: Option<ExecutionContext>)
from_vec(data, limits, cache_limits,
         execution_context: Option<ExecutionContext>)
```

Their `SourceLimits` bound ingestion, while the explicit cache and execution
arguments remain attached to deferred reads. The positional constructor is:

```text
from_read_at(source, limits, cache_limits, source_read_policy,
             execution_context: Option<ExecutionContext>)
```

Default callers pass `SourceCacheLimits::default()` and `None`; policy-aware
callers retain their selected values. Reader and vector paths check the
execution context before ingestion or moving the owned bytes, then retain it
through catalog admission and deferred payload reads. Advanced positional
callers still supply cache, read-ahead and execution policy in the one
constructor, so no capability or source-freshness check is lost. Internal
`_with_limits` helpers remain private implementation details only.

The six `VerifiedDecodedReaderError` accessor aliases and the three
`OpcOperationAccounting` accessor aliases are removed; all consumers use the
canonical accessors. The two OPC glossary profiles now register the
`_with_limits` suffix ban.

Native XLSX packages also expose their retained OPC source policy through
`Package::source_limits()`. The unified XLSX owner adopts that policy for both
fresh-workbook and decrypted-package routes, so default facade options cannot
widen a supplied package's bounds. A tight exact-input regression covers the
decrypted-package route and the resulting `File` raw-admission policy.

## Consumers and boundary

The same source change migrates the litchi-opc examples, fuzz targets, tests
and internal modules; DOCX, PPTX, XLSX and XLSB, OOXML common, litchi-ops and
facade callers; and the tracked perf-baseline and perf-execution sources.
ODF's independent `litchi_odf_common::core::SourceBackedPackage` calls remain
unchanged. Foreign causal payloads in crate-owned errors remain typed and
preserve progress/classification as allowed by ADR0034.

## Static validation

Direct `rustfmt --edition 2024 --check`, `git diff --check`, and
`python3 -m unittest tools.test_check_api_glossary` passed. Cargo, rustdoc,
and build caches were intentionally not run in this source-only lane; the
coordinator owns the isolated build and ABBA qualification after integration.
