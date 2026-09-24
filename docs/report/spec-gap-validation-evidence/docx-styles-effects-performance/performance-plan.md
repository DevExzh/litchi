# `stylesWithEffects` performance scaffold plan

Status: **reviewed, unrun attribution split**. The profile has no new timing,
allocation, RSS, scaling, native Office acceptance, or optimization result.
The committed correctness smoke and its receipts remain the gate for the
profile; they are not part of a timed sample. The profile harness now has the
reviewable `apply_ns`/`serialize_ns` split and matching subphase allocator
peaks described below. The bounded scaffold review and timing-free checks are
complete; a later clean capture requires its separate gate. The receipt
schema advances to `docx-styles-effects-profile-scaffold-v2`; retained v1
historical receipts remain byte unchanged and are replayed with their captured
source bundle.

The profile measures the public `litchi_docx::styles::effects` package owner
through the `Package` facade. It reports named end-to-end operations on a
bounded corpus and makes the setup, API, publication, reopen, and verification
boundaries visible. It does not model Word rendering or the complete visual
effects vocabulary.

## Source and evidence boundary

The historical correctness smoke is bound to
`d1f299d00e0dd5cc5cd8ddf9811c4b1ad21d1119`. The reviewed correctness smoke
descendant is `8444e88baaaca128eed50a2eed26e0ecd25063c9`; its committed result
tree at `results/clean-46c456848/` is retained unchanged. The current
attribution baseline is the separately approved full commit
`8702fd4db8723acceb7deb51bcb40ff66604bf10`, which includes the current OPC
publication changes. Its fresh 52-lane correctness smoke is retained at
`results/smoke-current-637082e31/` in commit
`fa927a8a9de94891858a6bb3d44c21d5fa63697d`. The current profile prerequisite
checks that exact approved file set and its Git blob hashes; a committed
descendant rewrite cannot replace the approved evidence. A future profile
must use a new clean descendant of both the production baseline and the
current smoke retention commit. The historical smoke source identity stays
separate; neither the active worktree nor silent repinning is admissible.

The performance harness is separate from `run_smoke.sh` and must not overwrite
the smoke harness, corpus manifest, or frozen receipts. Its source manifest
will bind every production local Cargo input to the approved production Git
blob and every harness, generator, verifier, and fixture input under this
evidence subtree to the captured scaffold `HEAD`. A path that already existed
at the production pin still binds to the captured evidence `HEAD`; a new or
changed production input fails closed. The manifest records the exact evidence
root, source commit, captured `HEAD`, metadata digest, every package file, and
every extra input.

The future runner will retain, before any build or timed process:

* the detached checkout `HEAD`, complete Git status, source pin, evidence-root
  path, and source-manifest digest;
* `cargo metadata --locked --offline` before and after the build, the exact
  `Cargo.lock`, all local path dependencies, and their Git blob hashes;
* `rustc -vV`, `cargo -V`, the selected rustup toolchain, target triple,
  compiler profile, linker, allocator observer, and all relevant environment
  variables;
* host/kernel identity, CPU model, logical and physical core counts when
  available, memory size, OS release, filesystem/mount identity, and the
  `/usr/bin/time -v` version; and
* the build log, binary hash before and after the profile, exact commands,
  process exit status, stdout/stderr sidecars, and source/build provenance.

The profile build uses the committed standalone lockfile with
`--release --locked --offline`, `CARGO_INCREMENTAL=0`, `LC_ALL=C`, and no
`RUSTFLAGS`, `CARGO_ENCODED_RUSTFLAGS`, `RUSTC_BOOTSTRAP`, or `RUSTDOCFLAGS`.
The runner refuses a dirty checkout, missing or moving Git inputs, pre-existing
result or target paths, and a missing source/build receipt. Results and target
are outside the checkout and disjoint. A target created by the runner is
removed only after all receipts are verified; raw results are never removed by
the runner.

## Inputs and scale classes

The four committed native copies and their member hashes come from
`corpus-manifest.json` and are reused without copying a submodule:

| class | fixture | owner state used by the profile |
| --- | --- | --- |
| native | `Bug54849.docx` | main and glossary present, independent XML |
| native | `ms-office-2010-signed.docx` | signed main present, glossary absent |
| native | `ComplexNumberedLists.docx` | main present, glossary absent |
| native | `testGlossary.docx` | main and glossary present, independent XML |

Before a fixture enters setup, the harness checks package bytes, effects-member
bytes, SHA-256, conformance, content type, relationship target and count,
owner presence, signature state, and the comparison with ordinary
`word/styles.xml` and `word/glossary/styles.xml` where present. Every sample
receipt repeats the package and selected-member hashes and records the owner.

The profile also uses deterministic source-controlled synthetic resources. The
generator creates a valid `w:styles` root in the fixture's conformance, a
bounded typed style projection, and an inert unknown extension marker. It does
not claim to exercise unmodeled visual effects. Generated XML and packages are
created in a fresh external fixture directory; they are not silently treated
as committed source inputs. The generator source, input, seed, exact generated
bytes, event count, maximum depth, style count, opaque-marker size, resource
hash, and complete package hash are captured in a replayable manifest before
the fixture is used. Generation, ZIP replacement, XML event/depth counting,
and package hashing occur before the operation clock.

The scale matrix is:

* **native/small:** the committed native effects member, approximately
  15--20 KiB;
* **initial scaled:** deterministic resources targeting 64 KiB and 1 MiB;
* **follow-up scaled:** 8 MiB, only after the initial profile is reviewed;
* **near-limit preflight:** one valid resource below the 32 MiB XML ceiling,
  with event and depth counts below their independent limits. It is generated
  and preflighted for a later profile only; it is not part of the initial
  timing matrix.

The near-limit preflight is admitted only after the correctness smoke proves
the source and XML limits. It is a bounded resource observation, not an OOM,
scratch-budget, or hard process-memory-cap test. A later near-limit timing run
requires a separate review of generated-fixture retention, process duration,
and memory interpretation. Cap lanes that depend on package topology use the
smallest fixture that changes the selected metric and record why a scale is
inapplicable.

## Timed lane matrix

The 52-lane smoke matrix remains the complete correctness matrix. Every smoke
success, cap boundary, typed refusal, malformed input, source readback, exact
inverse, and opaque-member check must continue to pass, but those 52 lanes are
not all initial timing rows. The first timed pass is deliberately bounded to a
representative public workflow matrix at native, 64 KiB, and 1 MiB resource
classes. It must not grow into a large timing framework before the first
receipts are reviewed.

The profile keeps stable lane names with fixture, owner, and scale fields in
each receipt. The initial timed rows are:

| family | initial rows | scale |
| --- | --- | --- |
| capture | Bug54849 main/glossary; signed main; Complex main; testGlossary main/glossary | native |
| capture | main and glossary on the Bug54849 package shape | 64 KiB, 1 MiB |
| projection | main and glossary on the Bug54849 package shape | native, 64 KiB, 1 MiB |
| exact source no-op | Bug54849 main and glossary | native |
| replace | Bug54849 main and glossary | native, 64 KiB, 1 MiB |
| remove | Bug54849 main and glossary | native |
| add absent owner | Complex main owner after prepared main-effects removal | native |
| inverse | Bug54849 main replace and remove/add | native |
| owner independence | Bug54849 main and glossary | native |

The signed/absent native captures remain in the initial capture rows where
they are representative; absent-owner, signed-change, stale-patch, all cap,
and all malformed/topology lanes remain correctness-only in this first timing
pass. A later refusal profile may add them without changing this baseline.
This keeps the first profile interpretable while retaining the complete
correctness evidence required by the owner contract.

### Read and source paths

* `native_capture_{fixture}_{owner}` covers each present owner listed in the
  initial table. `native_absent_{fixture}_{owner}` remains a correctness-only
  smoke row until a later timing extension is approved.
* `capture_{owner}_{scale}` opens a prepared synthetic package and loads the
  selected owner snapshot. `projection_{owner}_{scale}` reads the projection
  length, one ID/name lookup, one default-style lookup, and one bounded
  numbering lookup. Projection work is reported separately from package open.
* `source_noop_{fixture}_{owner}_{scale}` loads the source snapshot, creates
  an empty edit/commit, and publishes the exact no-op through the public
  source-checked path. A retained-resource `put` no-op is a separate row when
  the public operation is exercised. Pointer identity is diagnostic only;
  exact bytes and source tokens are the correctness gates.

### Changed resource and inverse paths

* `replace_{fixture}_{owner}_{scale}` replaces one present resource with a
  deterministic valid resource whose inert extension marker differs, commits,
  publishes, reopens, and checks the selected owner and every unrelated member.
* `remove_{fixture}_{owner}_{scale}` removes one present owner and verifies
  that the other owner, relationship graph, and opaque members remain intact.
* `add_{fixture}_{owner}_{scale}` adds a resource to an absent owner on a
  fixture whose package permits that owner. The missing-glossary case is a
  separate refusal and never seeds a glossary implicitly.
* `inverse_replace_main` and `inverse_remove_main` publish a changed patch,
  construct and apply the source-checked inverse and serialize its output under
  `inverse_ns`, then start `inverse_reopen_ns` immediately before
  `Package::from_reader` on the restored bytes. Reopening and loading the owner
  must recover the original package/member/owner hashes.
* `independent_{fixture}_{owner}_{scale}` changes one owner while hashing the
  other owner before and after. Main and glossary observations never combine
  into one anonymous “effects resource” row.

### Limits and refusals: correctness matrix and later timing extension

All cap and refusal lanes are retained as correctness-only rows in the initial
profile. The seven effect-owned cap families are separate smoke lanes:

`Parts`, `TotalPartBytes`, `TotalRelationships`,
`TotalRelationshipXmlEvents`, `TotalRelationshipXmlBytes`,
`RelationshipParts`, and `RelationshipGraphNodes`.

For each applicable family, the correctness setup derives the source and
projected value from the actual candidate package. It executes an exact-fit
success and a one-unit-under refusal, including the source-bound
`Transaction::commit` refusal where applicable. The receipt records the full
`ReadLimits` policy and the selected typed `ReadResource`, observed value,
maximum, source metrics, projected metrics, exact-fit output, and unchanged
source readback. A count that cannot grow for a replacement is marked
inapplicable and is exercised with the source-less add case instead of being
mislabeled as an edit boundary.

The full `ReadLimits` policy fields retained in every cap setup are
`input_bytes`, `archive_members`, `archive_total_entries`,
`archive_member_name_bytes`, `archive_metadata_bytes`,
`archive_compressed_bytes`, `archive_entry_bytes`, `archive_total_bytes`,
`parts`, `part_bytes`, `total_part_bytes`, `content_types_bytes`,
`content_type_mappings`, `relationship_parts`, `relationship_xml_bytes`,
`total_relationship_xml_bytes`, `relationships_per_part`,
`total_relationships`, `relationship_graph_nodes`, `xml_events`,
`total_relationship_xml_events`, `xml_depth`, `xml_attribute_bytes`, and
`relationship_target_bytes`. A cap receipt must carry the exact field name,
`ReadResource`, observed value, and maximum from the actual typed error; a
message substring or a copied limit number is insufficient.

If a later refusal timing pass is approved, setup will derive the same fields
from the candidate and label each row as one of two boundaries:

* **ingress refusal:** `Package::from_reader_with_limits` or owner capture
  rejects before a usable source package exists. The timed call starts before
  that ingress API, and no post-ingress edit cost is attributed to it.
* **post-ingress refusal:** a source package was admitted and a public
  `put`, `remove`, patch application, or `Transaction::commit` rejects. The
  timed call starts before that public operation and records the typed refusal,
  source readback, and no-output check separately.

The receipt records `refusal_boundary=ingress|post_ingress`, selected
`ReadResource`, and whether readback was possible. These boundaries must not be
combined into one “cap/refusal” timing number.

The correctness refusal matrix covers the public call for:

* missing glossary owner addition, stale/source-mismatched patch, and changed
  signed publication;
* duplicate, third/orphan, external, wrong-content-type, outbound,
  shared-inbound, root, namespace, and opaque XML topology/grammar cases; and
* unbound descendant, invalid QName, raw attribute/text, control,
  invalid-character-reference, empty-prefix, reserved-XML-URI, XML-version,
  XML-event, and XML-depth cases.

Malformed and signed inputs are constructed and hashed outside any future
refusal timing operation. An ingress refusal must retain the typed error and
input/source hash, emit no output, and leave the input unchanged; no usable
package exists for readback. A post-ingress refusal must additionally pass
source readback and physical, metadata, and opaque-member equality checks.
An unexpected success, generic error, partial output, or missing applicable
readback is a failed run, never a passing refusal sample.

## End-to-end operation and phase boundaries

The fixture bytes, synthetic XML, replacement `Resource`, malformed package,
cap limit, expected hashes, and derived metric are prepared before the elapsed
clock. The harness keeps package setup outside a mutation clock when the lane
is explicitly a prepared-source mutation; capture lanes include package open so
their end-to-end read cost is not confused with a prepared snapshot lookup.
The receipt names the boundary rather than combining unlike operations.

Each measured sample has one outer `elapsed_ns` clock and disjoint phase clocks.
Every phase starts and stops around the named public work; phase sums must be
no greater than `elapsed_ns`. The phase contract is:

| phase | timed work | excluded or separately named work |
| --- | --- | --- |
| `capture_ns` | `Package::from_reader` plus owner snapshot load for capture lanes | fixture I/O and synthetic generation |
| `snapshot_ns` | owner snapshot capture on a prepared package for mutation lanes | package setup and typed projection reads |
| `stage_ns` | edit creation and resource replacement/clear only | `Transaction::commit` and candidate publication |
| `commit_ns` | exactly `Transaction::commit` source/graph validation and commit formation | edit construction and patch publication |
| `publish_ns` | exactly the public patch/commit application and candidate publication | post-publication reopen |
| `reopen_ns` | ordinary reopen of published bytes and selected owner load | semantic/opaque checks |
| `inverse_ns` | inverse patch construction, source validation, and inverse publication | restored package reopen |
| `inverse_reopen_ns` | restored `Package::from_reader` and owner load, timed from before the call | inverse setup and post-checks |
| `projection_ns` | typed projection and selected scalar lookups when the lane names them | owner capture |
| `opaque_ns` | complete unchanged-member checks, including ordinary `Package` validation cost | graph metrics |
| `graph_ns` | package/relationship/owner metrics and closure checks | opaque member comparison |
| `readback_ns` | package readback after a refusal or failed publication | typed-error classification |
| `validation_ns` | typed error classification and semantic checks not assigned above | report serialization |

The harness takes each phase clock around a non-nested public call. It must not
start `commit_ns` inside `stage_ns`, include publication in `commit_ns`, or
include reopen/opaque checks in `publish_ns`. If a public call internally
overlaps a listed phase, the receipt records the overlap and excludes that
work from the disjoint phase sum. No opaque or graph work is hidden in an
unnamed validation bucket. The operation clock includes the designated
readback and API-result checks; report serialization, fixture creation, and
receipt writing are outside it.

The unrun attribution split refines `publish_ns` without replacing it. The
current bounded scaffold emits the split for successful prepared publication
lanes: `noop_main`, `noop_glossary`, `replace_main` and `replace_glossary` at
all three resource scales, `remove_main`, `remove_glossary`,
`add_main_absent`, `inverse_replace_main`, `inverse_remove_main`,
`independent_main`, and `independent_glossary`. Capture, projection, refusal,
and signed lanes retain null split fields; this does not expand the timing
matrix. For a source no-op, `apply_ns` includes both
`apply_styles_with_effects_patch` and the resulting `put_styles_with_effects`
call; `serialize_ns` remains the subsequent package serialization.

`apply_ns` covers the public mutation call, and `serialize_ns` covers only
`Package::to_stream` into the in-memory output cursor. Their sums must fit
within `publish_ns`; they are child clocks and are excluded from the outer
phase sum so work is not double-counted. Inverse mutation carries
`inverse_apply_ns` and `inverse_serialize_ns` under `inverse_ns`; inverse patch
construction is inside `inverse_apply_ns`. The complete observed inverse cost
is `inverse_ns + inverse_reopen_ns`: the former contains inverse construction,
public publication, and serialization, while the latter contains the restored
`Package::from_reader` and owner load. These clocks are disjoint. The harness
uses a separate allocator peak baseline for each child while preserving the
outer allocation counters and peak. Hashing, opaque checks, graph metrics,
and receipt construction remain outside these child clocks.

The `inverse_reopen_ns` clock begins before the restored `Package::from_reader`.
Starting it after parsing would under-report the inverse operation and fails
profile verification. Refusal timing separates typed classification from
`readback_ns`, and source readback reads the package after the actual failed
public call rather than comparing a captured snapshot with an immutable fixture.

## Allocation, live memory, and RSS accounting

The executable allocator regressions share process-global counters. Run the
profile harness tests serially from the repository root:

```sh
cargo test --locked --offline --manifest-path \
  docs/report/spec-gap-validation-evidence/docx-styles-effects-performance/profile-harness/Cargo.toml \
  -- --test-threads=1
```

Concurrent test execution is not a supported allocator verification mode.

The profile binary uses the existing process-local `CountingAllocator` observer
and emits raw per-sample counters. Before each measured operation, the harness
records `live_before` and resets the operation-local peak baseline. Each sample
records independently:

* `direct_allocated_bytes`, `realloc_old_bytes`, `realloc_new_bytes`,
  `requested_alloc_bytes`, allocation/reallocation/deallocation calls, and
  allocator failures;
* `live_before`, `live_after`, `peak_live_delta`, underflow/invalid flags, and
  the checked live-byte equation; and
* `elapsed_ns`, all phase values, output bytes, input/package/member hashes,
  and process/sample identity.

Publication lanes additionally carry `apply_allocation` and
`serialize_allocation` objects with the same checked counters and the
subphase-specific `peak_live_delta`. Inverse lanes carry the corresponding
`inverse_apply_allocation` and `inverse_serialize_allocation` objects. A
non-publication lane records these objects as null. These child counters are
allocator traffic for the named child only; process RSS remains the separate
fresh-process maximum from `/usr/bin/time -v`.

The verifier rejects negative, boolean, fractional, string, overflowing, or
impossible numeric values. It checks unsigned phase types, disjoint phase
bounds, `peak_live_delta >= max(0, live_after - live_before)`, allocator
equations, and balanced counters. `requested_alloc_bytes` is allocator traffic;
it is not retained memory.

The outer runner invokes each fresh process under `/usr/bin/time -v` and
retains the complete sidecar. The one `Maximum resident set size` marker is
the **process maximum** across that process's two warmups and twenty measured
samples; it is not a per-sample RSS, a phase RSS, an average, or a quantile
derived from allocator counters. The receipt names it `rss_max_kib` and also
records user/system CPU time, elapsed wall time, process exit status, and
stderr. RSS is a whole-process observation that includes runtime state,
mappings, allocator arenas, and page effects. `peak_live_delta` is a
process-local allocator observation. The report never substitutes one for the
other or calls either a hard memory cap.

## Fresh-process sampling and uncertainty

For every lane/fixture/owner/scale tuple, the runner executes **three fresh
processes sequentially**. Each process performs **two warmups followed by 20
measured samples** for that one tuple. The processes use the same prepared
fixture hash, command, environment, binary hash, and source manifest. No other
agent's timing workload may run concurrently.

Raw JSON and `/usr/bin/time -v` receipts are retained per process. The summary
reports each process's p50, p95, p99, allocation quantiles, peak-live
quantiles, and `rss_max_kib` (the process-level RSS maximum), then reports the median and min/max range of the three
process-level p50 values. Per-process tails remain visible; pooled samples are
labelled descriptive and are not treated as 60 independent machine runs.
The report records sample count, warmups, process IDs, exit status, and any
host contention indication. A missing process, short sample set, changed
binary/source hash, or nonzero sidecar status fails the profile.

The host-hygiene receipt is captured immediately before and after the process
set. It records UTC time, load average, CPU model and online CPU count, memory
availability, kernel/OS identity, and requested and observed affinity when a
runner uses one. `host_probe.sh` does not inspect concurrent processes. The
operator must separately retain a visible-process census before and after
the run, coordinate the prohibition on other agents' timing workloads, and
report any observed build/profile contention. A census is an observation,
not proof that unobserved host activity was absent. The runner runs the three
processes sequentially and does not itself run a concurrent build
or benchmark, and does not claim CPU isolation, frequency locking, cache
flushing, thermal stability, or idle-host status unless the receipt proves
that property. Unknown host fields are recorded as unavailable rather than
silently inferred.

## Fail-closed publication gate

Before a future timing run, the scaffold must pass the existing source and
receipt regression tests plus a profile-specific verifier review. The runner
must reject a dirty or non-descendant source tree, changed production blobs,
uncommitted harness/generator/fixture files, stale metadata, altered toolchain
or flags, missing native hashes, missing synthetic hashes, pre-existing output,
or missing build/binary/host receipts. It must preserve the frozen smoke
receipts and write profile results to a new mode-specific path.

After all samples, verification must pass every semantic lane, typed refusal,
exact inverse, physical/metadata/opaque comparison, source-manifest before /
after equality, binary before/after equality, phase bound, allocator equation,
RSS sidecar, process-status, and sample-count gate. The report is generated
only from verified raw receipts.

The first profile report will therefore be an absolute, scenario-scoped
baseline. It may identify dominant end-to-end and phase costs and process
uncertainty. It will make no before/after speedup claim, no native Word
acceptance claim, and no claim about the wider DOCX library until a separately
matched candidate and control run are reviewed.
