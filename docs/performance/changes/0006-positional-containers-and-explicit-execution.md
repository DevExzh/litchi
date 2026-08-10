# Positional containers and explicit execution

Status: implemented foundation; broad performance acceptance pending

CFB now has positional `SharedOleFile` access and explicit bounded bulk reads.
soapberry-zip has one indexed positional ZIP representation, opaque `EntryId`
handles, and reusable local `ParallelReadSession`s. `litchi-core` supplies the
runtime-neutral `ExecutionContext`; OPC's public `OpenSession` maps its worker,
task, byte, threshold, affinity, cancellation, and hierarchical budget policy
to a local ZIP session. Ordinary/legacy opens remain serial, and hidden global
Rayon paths were removed.

The contract is deliberately additive: public document CRUD facades do not
expose executors or physical IDs. Cancellation, budget, and session failures
remain typed; in-flight work is bounded. Focused container, core, OPC, and
format tests plus changed-crate formatter and warning-denied Clippy checks
passed. Performance matrices for concurrent CFB and explicit OPC sessions are
not yet sufficient for a throughput or latency claim.
