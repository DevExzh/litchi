# Change 0305: XLSB semantic cache single-flight

Status: implemented

`SourceBackedWorkbook` now coordinates shared-string and style semantic cache
initialization with private `OnceCell<Arc<T>>` values. Each workbook resource
has its own cache, so a successful concurrent initialization runs once and all
callers retain the same immutable `Arc`. Failed initializers are not cached;
the next caller retries the bounded parse. OPC payload caching and worksheet
materialization behavior are unchanged.

Execution and source-version checks remain before initialization, throughout
the loader, immediately before publication, and after obtaining the retained
Arc. The loader does not hold an explicit user mutex guard across OPC reads or
semantic parsing; `OnceCell` owns initializer coordination. The initializer is
structurally non-reentrant, so custom source callbacks must not recursively
enter the same owner while reads are in progress. Raw per-Part read limits
bound the source payload, not decoded semantic allocation; one retained Arc
means one decoded cache allocation plus cloned handles.

## Performance claim

`performance_claim: none`

The only claim is one concurrent successful semantic initialization and one
retained Arc per workbook shared-string or style resource. This note excludes
failed retry counts, worksheet parses, total semantic memory, RSS, OOM
behavior, latency, I/O volume, and OPC cache diagnostics.
