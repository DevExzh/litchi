# Change 0428 source-cache diagnostic forwarding

This note records the PPTX owner seam for change 0428. The implementation
adds an additive fallible diagnostic method to both
`SourceBackedPresentation` and `SourceBackedPresentationEditor`:

```rust,ignore
pub fn try_cache_diagnostics(
    &self,
) -> std::result::Result<
    litchi_opc::SourceCacheDiagnostics,
    litchi_opc::SourceCacheDiagnosticsError,
>
```

Each method forwards the result from its owning
`litchi_opc::SourceBackedPackage` unchanged. The existing infallible
`cache_diagnostics` methods remain available for compatibility callers. The
fallible methods are intended for instrumentation and telemetry that must
reject a poisoned cache-state mutex or checked counter overflow rather than
recording a recovered snapshot as valid evidence. Neither method loads a
source payload while observing the cache state.

Focused in-file tests cover both public owners. They compare each forwarder
with the private owning package's fallible result, observe a managed package
before and after selecting a slide, and use both the total positional-source
read-call counter and the second-slide payload counter to prove that
diagnostic observation adds no source I/O or payload read. The tests also
retain a selected presentation handle or slide snapshot while the owner is
dropped, then verify that managed budget memory and object usage are zero
after the selected handle is dropped. The editor test does not assume any
particular intermediate object count while its snapshot is retained.

The PPTX tests exercise healthy forwarding and do not inject poisoned mutexes
or counter overflow. Those typed error paths remain owned by the existing OPC
tests; this change makes no claim that the PPTX forwarder independently
manufactures or injects those failures.

The seam follows the accepted ownership and API contracts: typed fallible
public errors and a concise format owner (ADRs 0001 and 0002), immutable
source-backed handles whose ownership remains explicit (ADR 0003), bounded
positional reads and fail-closed telemetry (ADR 0005), and preservation of
the OPC validation/error boundary (ADRs 0006, 0011, and 0024). This is an API
and correctness change, not a latency, allocation, cache-retention, or RSS
claim.

No Cargo, build, test, profiler, or CPU workload command was run while
preparing this source change; the parent validation run remains authoritative.
