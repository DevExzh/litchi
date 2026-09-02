# Change 0372: ODP source-backed catalog fused parse

## Scope

The source-backed ODP catalog open path previously validated `content.xml`
and then scanned it again to build the presentation catalog. Change 0372
delivers each borrowing quick-xml event first to the shared ODF content
validator and then to the ODP catalog scanner in one plain-`Reader` pass.
Catalog-only errors are deferred while validation continues to verified EOF,
so tokenizer, validation, and validation-finish errors retain precedence.

The shared namespace tracker is now fallible for all input-dependent buffer,
binding, unknown-prefix, and allocation-error-context growth. A materialized
size preflight rejects `content.xml` over 256 MiB before allocation, including
the encrypted-plaintext metadata case. The 4,096 content-depth limit,
256 namespace declarations per element, 65,536-page ceiling, 1 MiB page-name
limit, source freshness fences, ZIP verification, MIME checks, publication
fences, and media locality remain in force.

The implementation stays in the existing ODF-common and ODP owner layers. It
adds no public CRUD API, dependency edge, package identifier, archive handle,
runtime handle, lock, raw type, or unsafe code.

## Correctness evidence

Locked/offline release tests passed for every executed `litchi-odf-common`,
`litchi-odp`, and `litchi-odt` target. Library totals were 282, 162, and 557
tests respectively. Two pre-existing tests in unmodified ODF writer code were
excluded by exact name: `encryption_authoring_uses_no_unsafe_code` and
`metadata_is_validated_and_bounded_before_member_output`.

Scoped Clippy passed with only the pre-existing
`clippy::large_enum_variant` diagnostic in
`litchi-odf-common/src/package/model.rs` allowed. The crate-boundary gate
passed for 64 packages, 240 dependency declarations, and 14 explicit debt
entries. Independent architecture, semantics, tests, resource-safety, and
final static reviews accepted the batch.

## Clean ABBA evidence

The clean CPU-2 ABBA used one worker, 30 warmups, and 500 samples per selector
per leg in canonical `A1-B1-B2-A2` order. The deterministic media-rich ODP
corpus contains 12 slides and eight media payloads; its 16,785,912-byte
archive has SHA-256
`661ae80396d4eda673d35e45d208443cc359052e4b9b27fed0ba6681602a913a`.

Control was clean revision `32290f7cef837d4eb8377a7085b9607d21651d4d`
with tree `cd76a9c5bc40cd1711337089b40d748645d60867` and binary SHA-256
`7291b2b99859f7e1a48a7e8ed143818cf64aa75b797c08978fb78eefa8cbf4c5`.
Candidate was clean revision
`922eb5e2c56f2a65c8b755514da251f682a902e6` with tree
`52feb9fed1558b98ee8bbf199796ea7234763ba9` and binary SHA-256
`9a26cd1997d36ce18bdcf724e083984edab722b23222be0e3e0e7a579677c847`.

| ABBA legs | open p50 | open mean | open p95 | open p99 |
| --- | ---: | ---: | ---: | ---: |
| A1/B1 reduction | 15.627% | 11.226% | -1.378% | -4.989% |
| A2/B2 reduction | 17.327% | 17.038% | 14.361% | 17.307% |

All semantic, topology, media-locality, source-replay, freshness,
content-read, zero-post-preparation-list-read, and zero-picture-read gates
passed in every leg. The control observed 51 source versions and the
candidate 53 because the candidate performs two additional bounded
preflight/version observations; each binary was internally stable, so
identical source-observation counts are not claimed.

The retained machine-readable evidence is [the clean ABBA result](../results/odp-source-catalog-0372-abba.json).

## Claim boundary and decision

The accepted claim is limited to p50 latency for a fresh
`SourceBackedPresentationCatalog::from_read_at` open on this deterministic
media-rich ODP corpus. The two comparisons improved by 15.627% and 17.327%,
while control and candidate p50 same-side drift stayed below 5%.

Open mean is withheld because control mean drift was 6.189%. Open p95 and p99
are withheld because control tail drift exceeded the applicable gates and the
first comparison moved adversely. List is rejected because its
tens-of-nanoseconds measurements were unstable and changed direction. Query
is rejected because both comparisons regressed.

No broad open-latency, list, query, all-ODP, RSS, allocation, fixed-memory,
physical-I/O, cold-cache, throughput, or OOM-prevention claim is made.
