# 0911 PPTX cold opened-snapshot source audit

This audit binds the owner-level attribution to the frozen 0910 source and
records the setup call discovered in the first raw Callgrind capture. It is a
source and call-graph audit; it makes no production change and does not turn
the initial capture into a qualified pairwise result.

## Custody

The source receipt is
`docs/performance/results/change-0910/source-v1.json`, 1,756,112 bytes,
SHA-256
`7a6d1310fb97dde1f275ec8702efabbef2cfb643c9ee4c1178169fed53aa4586`,
with parent `468fe2931d6532c32f4c6608590816b8eaa88e31`. The retained binary
used by the raw capture is 9,301,456 bytes,
SHA-256
`f4ad7ee158061cc80354c3c8674c9fe6a8100ed76a0c8f743653efcb06bb2e8d`.
The relevant source files match the source manifest exactly, including
`tools/perf-baseline/src/lib.rs`
(`8542f4e8e580bef2b12b3782b03bcfabc8e0d60cefc80f27c55042c57baff860`),
`tools/perf-baseline/src/pptx_owned_benchmark.rs`
(`121a4b8817da5410222d4bd35252327cea75fc73866c3b7f8ed9f71b15a3aa46`),
`crates/litchi-pptx/src/package/model.rs`
(`f81617c03d7ea2f6d579d9c4c66f4be6f795dd78af3e4259cdc8728198d38985`),
`crates/litchi-pptx/src/opened/model.rs`
(`9f36d2bb66516616be1f5a4850d7b2e8aeb9d885558bd61ffafd6d40d994cce0`),
and `crates/litchi-pptx/src/opened/transaction.rs`
(`01352c03f17c8cdbed0902ac582a9d94368ddce6ae3b6f79c157c89f6b7cb310`).

The first raw capture is
`capture/materialized-a`. Its command receipt binds the exact binary,
`warmup=0`, `samples=1`, CPU 12, and the owner toggle. Its output passes the
existing source, semantic, output, sink, and immutability gates: source
17,017,139 bytes with SHA-256
`61b2b99083ca27ebd37955db600955e3f41289b93dba71951983164239eff757`, and
output 17,017,145 bytes with SHA-256
`23d6d7b8dd433ff453307ee56485efadc34d6b18b455a6858d9f565c4b1b6cd9`.

## Timed owner and harness boundary

The benchmark calls the selected owner exactly once in each timed iteration:
`timed_iteration` at
`tools/perf-baseline/src/pptx_owned_benchmark.rs:294` calls
`Package::opened_presentation_transaction()`. The `snapshot_ns` timer starts
immediately before that call and ends immediately after it. The selected
method returns before shape edits, commit, publication, serialization, sink
writes, semantic verification, and artifact writes.

The normal output verifier does not call the owner. The path
`verify_output` → `verify_pptx_multi_slide_edit_output` →
`verify_pptx_source_edit_semantics_for` reopens the output through borrowed
`Package::from_bytes` and reads the presentation, slides, and OPC parts.

There is one setup exception. `run` constructs the corpus before the warmup
and measured loops. `build_pptx_source_edit_corpus` calls
`pptx_source_edit_bytes` (the helper is inlined in the release binary), and
that fixture-authoring helper performs one setup transaction at
`tools/perf-baseline/src/lib.rs:16978` to add the eight media parts. This call
is outside the benchmark timer, but the Callgrind owner toggle still sees it.
The corpus builder's later borrowed semantic and OPC checks do not call the
owner.

## Cold owner path

`Package::opened_presentation_transaction` (package/model.rs:296) calls
`opened_presentation().edit()`. `opened_presentation` forwards to
`opened_presentation_with_limits` with default limits. On a fresh
`Package::from_vec` package, `part_digest_memo()` is empty, so the branch at
package/model.rs:281 selects `capture_with_provenance`, which invokes
`capture_internal` with `Revision::Cold` and no parent MCE or slide-root memo.

`capture_internal` performs the following work before it constructs the
transaction:

1. `PresentationPart::from_package` resolves and parses the main
   `presentation.xml`; `slide_references_strict` parses and validates the
   ordered slide catalog and its relationship targets.
2. `capture_slides_with_mce` walks all 200 slide references. Each slide is
   accessed through `get_part`, which forces an owned deferred ZIP payload on
   first access, then validates the slide root, processes markup-compatibility
   XML, extracts the producer name, and performs the notes-root conformance
   scan. There is no parent MCE/root memo on the cold branch.
3. The notes snapshot loader rechecks the presentation and slide-root
   classifications through the proofs from the slide pass. This generated
   fixture has a notes master part but no `notesMasterId`/notes-slide graph;
   the loader therefore performs metadata/oracle checks and returns no notes
   graph rather than materializing notes payloads.
4. It builds the 200-entry slide name index, clones the OPC graph into the
   owned immutable snapshot, and computes the complete package fingerprint.
   The cold fingerprint sorts all 229 parts, feeds package relationships and
   non-part members, forces/decodes every deferred part payload, hashes every
   payload with SHA-256, then hashes each part's name, content type, payload
   digest, and sorted relationships. The resulting `PartDigests` retains the
   payload allocations.
5. `Transaction::new` then clones the snapshot package graph and the slide
   vectors (`opened/transaction.rs:42–49`). This clone shares payload Arcs; it
   does not re-decompress the archive. The owner includes this construction
   and the temporary snapshot drop.

Owned ingress work precedes this boundary. `Package::from_vec` moves the
`Vec` into the OPC retained-source owner, builds the physical ZIP/catalog and
deferred part descriptors, reads structural content types and relationship
members, and establishes preservation provenance. It does not represent the
snapshot owner's cold fingerprint or slide capture, and it is outside the
selected owner interval.

## Raw attribution result and admission rule

The initial materialized-a capture has two positive Callgrind dumps. The
termination dump is zero. The positive dumps are attributable by their exact
dynamic caller, not by file order or magnitude:

| raw part | exact incoming caller | owner calls | owner edge Ir | interpretation |
| --- | --- | ---: | ---: | --- |
| `callgrind.out.1` | `litchi_perf_baseline::build_pptx_source_edit_corpus` | 1 | 126,617,815 | setup fixture authoring, untimed |
| `callgrind.out.2` | `litchi_perf_baseline::pptx_owned_benchmark::timed_iteration` | 1 | 1,000,613,578 | measured snapshot owner |

The raw graph contains no third positive `cfn` edge to the owner. The first
protocol's single-positive-interval gate therefore fails, because setup is
inside the owner toggle even though it is outside `snapshot_ns`. A bounded
revised admission may retain and report both intervals, classify them by the
exact incoming caller above, and use only the `timed_iteration` edge for the
snapshot attribution. Selecting the largest dump, assuming dump order, or
subtracting setup after the fact is not admissible.

The measured edge is still only an instruction count for the selected owner
on this generated 0910 fixture. It is not elapsed time, allocation count,
RSS, cache behavior, or a causal explanation of the 0910 wall-clock ratio.
The source map identifies candidate work for a later profiled change—slide
MCE/root parsing, complete payload fingerprint/decompression/hash, and graph
cloning—but does not authorize an optimization from this capture alone.
