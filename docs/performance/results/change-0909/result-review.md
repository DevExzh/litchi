# 0909 owning-input PPTX result review

**Result:** qualified_generated_fixture. The v3 owning-input control passes its bounded correctness, preservation, transaction, refusal, sink, limit, durable patch, and replay gates for the frozen 0907 generated corpus. This review was an offline source and artifact review; it did not run a build, workload, profiler, test, or replay. The retained replay and quality records are reported as evidence, rather than rerun here.

This review is bound to the final [0909 qualification report](../../0909-pptx-owning-input-preservation-control.md), which records the same source custody, four-census reduction, cleanup, and scope decision.

## Source custody and final records

The v3 packet is bound to parent commit 57de5495f899a43574110f8c572823a7d9907e5f and source manifest [source-v3.json](source-v3.json), which records 9,352 files. Before the temporary build checkout was cleaned, I matched every manifest entry against that checkout: zero files were missing and zero byte or SHA-256 mismatches were found. The retained [replay-before-cleanup.json](replay-before-cleanup.json) independently records reconstruction of all 9,352 files, eight capture artifacts, and four byte-exact archive replays.

The v3 overlay patch [source-overlay-v3.patch](source-overlay-v3.patch) is 29,255 bytes with SHA-256 e7b6ea69348e14febdf7ca507eccf9eb997a832c6f1c552f0d87effe596ee840. It adds only the performance harness paths tools/perf-baseline/src/lib.rs, tools/perf-baseline/src/pptx_owned_control.rs, and tools/perf-baseline/examples/pptx_owned_control.rs. No production implementation path is in the overlay. The v3 build and capture receipts both have exit code 0. The canonical [analysis.json](analysis.json) binds these records and reports qualified_generated_fixture.

The capture report intentionally retains the intermediate status probe_passed_pending_independent_raw_census. That status is resolved by the separate four-run census packet and the final analysis record; it is not being treated as a qualified result by itself.

## Corpus and ownership qualification

The source is the frozen media-rich 0907 corpus: 200 slides, eight text boxes per slide, eight 2 MiB media entries, 229 OPC entries, 445 ZIP members, 17,017,139 archive bytes, and source SHA-256 61b2b99083ca27ebd37955db600955e3f41289b93dba71951983164239eff757. The edit selects slide positions 0, 28, 57, 85, 114, 142, 171, and 199 and replaces 64 text boxes.

The ownership record is dynamic evidence from the package state:

| Observation | Recorded value |
| --- | --- |
| Owning ingress is unmodified | true |
| Ingress source length and SHA | 17,017,139 bytes; 61b2b99083ca27ebd37955db600955e3f41289b93dba71951983164239eff757 |
| After edit, exact unmodified owner remains | false |
| After edit, source admission limits remain | true |
| After edit, exact source length and SHA | absent |
| After no-op commit, unmodified owner remains | true |

These values are obtained from the owning OPC package and then checked by error-return guards. The serialized true flags are proof-state markers reached only after those guards; they are not presented as independent instrumentation. The independent archive hashes and census provide the separate output evidence. Input ownership also remains distinct from a claim that no copy occurred: this packet does not measure input-copy instructions or allocations.

The transaction path is exercised through an empty commit, a changed commit, inverse application, forward replay, serialized patch round trips, and fresh reopened packages. The exact no-op equals the source. The changed output passes the semantic slide and OPC topology checks, and the untouched logical-part check. Both durable forward and inverse patches are 60,953 bytes; their recorded SHA-256 values are 9997855655db832e443636f37b6b118b1c42a25b030096f169a4f26a49784af6 and 52498fa4f0c025bac28f8a5c05012b22518d81e695457bd2f007d07472ca5a2a.

## Runtime gates

The final v3 gate record has all scalar correctness gates true, with the following material witnesses:

| Gate | Evidence |
| --- | --- |
| Source identity and shape | Frozen source bytes, SHA, 229 entries, 445 members, and generated 200-slide shape match |
| Exact no-op | No-op archive is 17,017,139 bytes with the source SHA |
| Forward edit | Output is 17,017,145 bytes with SHA 23d6d7b8dd433ff453307ee56485efadc34d6b18b455a6858d9f565c4b1b6cd9 and passes semantic/output checks |
| Inverse and replay | Inverse equals the source bytes; forward replay equals the output bytes |
| Durable reopen | Reopened source plus durable forward patch equals output; reopened output plus durable inverse patch equals source |
| Stale and foreign patches | Both are refused with typed UnsafeEdit debug records and leave output bytes unchanged |
| Input limit | Actual 17,017,139 versus maximum 17,017,138, returned as the typed InputBytes ReadLimit boundary |
| Source owner | The source byte owner remains unchanged after the run |

The stale fixture mutates an affected selected slide before applying the forward patch, so this is an affected-resource freshness check. It does not certify unrelated-resource or every possible mutation scope.

The partial sink accepts exactly 8,508,572 bytes and returns OpcError::IncompleteOutput. I independently hashed the retained output prefix through that byte boundary and obtained 6286ed5e43e5a7d443392494736fc7b08887923cf0156392dfc3dddf21c26909, matching the gate record. The zero sink accepts zero bytes, leaves the package unchanged, and returns the existing preservation-writer mapping OpcError::ZipError containing the injected sink cause. The result makes no claim that the underlying std::io::ErrorKind is publicly retained.

## Independent raw ZIP census

The canonical [census-runs.json](census-runs.json) uses [zip-census.py](zip-census.py), SHA-256 0cdee4682959e9a448d81bc598a810c7c16f780bf378bcd8faae294f1cf1efd3, for four independent runs. Every run exited zero with status pass, 445 members, zero violations, equal filename order, equal archive comments, data descriptors observed in both archives, and ZIP64 false in this generated fixture.

The output census finds exactly 437 unchanged members and exactly these eight changed members:

- ppt/slides/slide1.xml
- ppt/slides/slide29.xml
- ppt/slides/slide58.xml
- ppt/slides/slide86.xml
- ppt/slides/slide115.xml
- ppt/slides/slide143.xml
- ppt/slides/slide172.xml
- ppt/slides/slide200.xml

For those eight members, the changed components are the central metadata, decoded payload, local data descriptor, and raw compressed payload. Local header metadata, local and central names, local extra fields, central extra fields, and central comments remain equal. The no-op, inverse, and forward-replay censuses each report 445 unchanged members and zero changed members.

I independently inspected all five retained archives with the standard ZIP reader. Each has 445 unique members, and the local header offsets are physically sorted in source, no-op, output, inverse, and forward-replay. The first and last member names and central filename order agree across each comparison. The generated archives use data descriptors and do not use ZIP64; this records the fixture properties and is not adversarial ZIP64 or descriptor coverage.

The source, no-op, and inverse artifact hashes are identical. The output and forward-replay hashes are identical. Full inverse byte identity, including the durable reopened inverse, therefore implies resource-byte, logical, and semantic identity for this known admitted source. It does not establish byte identity for arbitrary producers or every possible Patch input. The external census establishes output preservation properties; it does not reveal an internal Copy or Regenerate decision and does not prove avoided compression instructions.

## Quality, replay, and retained history

The final scoped [quality-v3-v2/receipt.json](quality-v3-v2/receipt.json) passes rustfmt using the retained rustfmt configuration and Clippy with warnings denied for the owning-control example. Its scope is the changed harness and example; it is not a full-workspace test result. The earlier quality-v3 receipt failed because it used the default formatting policy without the source snapshot's rustfmt configuration. That failure remains retained history and is superseded for this scoped quality check.

The retained v1 and v2 records also remain useful history: v1 failed to compile on collection type inference and the ReadLimit field spelling; v2 corrected those errors but stopped at an overly specific zero-sink IoError expectation. The qualification amendment records the actual ZipError mapping, and v3 captures all runtime gates with that existing mapping.

## Qualification boundary

This result qualifies the existing route for the same generated 0907 corpus and its recorded edits. It makes no performance, allocation, input-copy, compression-instruction, or production provenance-bridge claim. No public production Copy/Regenerate observer exists in this packet.

The remaining coverage is explicit: unknown-member variants, signed input, encrypted input, malformed input, and source mutation during publication are not covered; the source-mutation lane uses an immutable Vec owner. The independent census receives the primary generated corpus only, and generated data descriptors/ZIP64 values do not stand in for adversarial archive coverage. The local Rust raw inventory omits replay, which is covered by the separate external replay census. A broader producer claim would require new inputs and the same source, ownership, semantic, refusal, sink, limit, inverse, and raw archive witnesses.

No correctness blocker remains for this bounded qualification. A production adoption or performance decision still requires a comparable owning-input benchmark with explicit input-copy and publication boundaries.
