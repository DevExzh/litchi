# 0512 ADR and validation scope

All 29 accepted ADRs and the README were previously read and freshly verified
unchanged against [adr-manifest.json](adr-manifest.json). Production and
harness source inventories are identical to the committed 0511 endpoint;
no API, dependency, owner, unsafe-code, memory limit or parser behavior changes.
No ADR exception is required.

ADRs 0001/0005/0008 govern evidence: operation timers, scoped simulated
instructions and whole-child hardware/RSS remain separate. Source, binary,
compile fixtures, raw samples, corpus catalogs and tooling hashes bind replay.
The native observations have 30 samples per row and support current descriptive
context, not a registered optimization claim. Missing allocation and physical
I/O metrics are not replaced with zeroes or inferred process deltas.

ADRs 0003/0006/0011 constrain the follow-up: parser/snapshot fusion must preserve
immutable atomic publication, original-source spans, MCE preprocessing/error
precedence, unknown content, typed failures and XLSX/OPC ownership. The other
ADRs are unaffected; ODF is deferred and iWork excluded by user priority.

The fresh unchanged build and all captured benchmark correctness oracles pass.
Metadata gates and verifier negative-vector checks validate this evidence-only
change. No new Rust test matrix, fuzz campaign, native Office compatibility,
physical-cold/range-source or scaling result is claimed. Prior source tests
remain historical evidence, not falsely counted as current executions.
