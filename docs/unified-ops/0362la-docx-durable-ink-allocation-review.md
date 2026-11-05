# 0362la — Return durable Ink for allocation repair

- Candidate: `5559d9eaef`; baseline: `d0406be017`.
- Status: returned for repair; no performance acceptance.

The matched release binaries pass all six exact-output comparisons, including
the forward, inverse and no-op durable forms. All four counting legs agree.
The [verified receipt](results/0362l/evidence.json) records the protected
binaries, source manifests, dependency lock, probe and commands. Raw evidence
remains outside the checkout under `.coord/durable-ink-0362l/matched-v1`.

Per 100 operations, forward and inverse replay have unchanged allocation
counts and requested bytes; their peaks fall by 40 and 272 bytes respectively.
Their publication routes have unchanged counts, requested bytes and peaks.
No-op replay adds 200 allocations and 8,000 requested bytes, with peak live
bytes rising from 3,309 to 4,667. No-op publication adds the same counts and
bytes, with peak rising from 3,592 to 4,667. All routes retain zero bytes after
the operation. No native timing run has been attempted.

The new no-op commit retains its admitted graph binding through fingerprint
verification and allocates two empty patch buffers. The binding duplicates
typed graph metadata already held in its immutable OPC package. The next
repair will share that ownership while preserving exact lexical tokens,
relationship-member presence, complete typed graph comparison, internal
target validation and all existing admission limits. Fresh correctness and
matched allocation checks are required before native qualification.
