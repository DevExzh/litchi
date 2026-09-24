# PPTX existing InkAction evidence requirements

This is the acceptance contract for the bounded adapter. It is not a measurement
result. The initial profile contains exactly 23 recipes and 42 lanes; the
complete names and recipe mapping are in `corpus-manifest.json`.

## Public API binding and setup

The adapter uses only the public existing-target routes:

| scope | required call | setup excluded from the timer |
|---|---|---|
| package inventory | `Package::ink_actions()` or `Package::ink_actions_with_limits(owner_limits)` | deterministic `OpcPackage`, mutable `Package::from_opc_package(opc)`, source hashes, and expected values |
| presentation inventory | `Package::presentation()` followed by `Presentation::ink_actions()` or `_with_limits(owner_limits)` | package construction and borrowed presentation handle creation |
| existing-target edit | `Snapshot::edit()`, `Edit::edit_profile`, `Edit::commit` | inventory and post-commit semantic assertions |
| publication | `Package::apply_ink_actions_patch` on a fresh mutable package | fixture/package setup and commit construction |
| save | `Package::to_bytes()` on a fresh mutable package | fixture/setup and output checks |
| retained OPC reopen | `Package::from_vec_with_limits(bytes, retained_opc_limits)` followed by a public read | source bytes and expected semantic values; the call must retain the supplied OPC limits |

`Package::from_opc_package` is the setup route for mutable package lanes. A
future adapter must not clone a package that a prior sample mutated. The
retained OPC `ReadLimits` and owner `ink_actions::Limits` are separate receipt
objects. OPC limits cover retained package parts/bytes and relationship XML;
owner limits cover anchors, one target, aggregate unique targets, and selected
target relationships. One must not be substituted for the other.

Stale and signed recipes use two raw `OpcPackage` values in setup. The
baseline graph is wrapped with `Package::from_opc_package` to create the source
snapshot and patch. For stale lanes, a separate raw graph is then mutated in
exactly one owner XML, owner `.rels`, target blob, or target content type
before it is wrapped with a fresh `Package::from_opc_package`. Signed lanes
instead place the same signature edge on both raw graphs before wrapping, then
apply a changed or exact no-op patch from the signed baseline. The timed stale
or signature lane receives an already-prepared public package; no raw OPC
mutation is performed inside the timed public operation.

The adapter may not call private codecs, use a neutral `iact:actions` fragment
as a host substitute, pass relationship IDs/package URIs as semantic
selectors, infer a fresh action path/MIME, or change production code.

## Exact recipes and lane count

The initial matrix has exactly 23 recipes and 42 lanes. The concrete recipe
IDs are `r01_tiny_shared`, `r02_small_shared`, `r03_small_distinct`,
`r04_medium_shared`, `r05_medium_distinct`, `r06_large_shared`,
`r07_large_distinct`, `r08_near_shared`, `r09_near_distinct`,
`r10_multislide_shared`, `r11_multislide_distinct`,
`r12_case_equivalent_shared`, `r13_strict_shared`,
`r14_opaque_unknown_external`, `r15_signed`, `r16_stale_owner`,
`r17_stale_owner_rels`, `r18_stale_target`, `r19_stale_content_type`,
`r20_limit_anchor`, `r21_limit_target`, `r22_limit_aggregate`, and
`r23_limit_graph`. The lane IDs are the exact 42 entries in the manifest;
there is no implicit Cartesian scale matrix or generator-selected point.

The nonzero boundary triples are fixed:

| resource | one-under | exact | one-over | source |
|---|---:|---:|---:|---:|
| anchors | 7 | 8 | 9 | `r20_limit_anchor` has 8 anchors |
| one target bytes | 16383 | 16384 | 16385 | `r21_limit_target` has 16,384 bytes |
| aggregate target bytes | 32767 | 32768 | 32769 | `r22_limit_aggregate` has two 16,384-byte targets |
| selected graph edges | 7 | 8 | 9 | `r23_limit_graph` has 8 retained edges |

Exact values must succeed when the source fits. One-under values must refuse
before the relevant typed result/candidate is retained. One-over values are a
positive control and must remain nonzero. The actual public error and resource
are authoritative if a boundary is reached by an outer retained OPC limit.

## Correctness and graph gates

Each successful sample must prove, outside the timer:

- semantic slide/anchor selector and source fingerprint resolution, with no
  physical relationship ID as public identity;
- exact MCE namespace/`Requires`, package-dialect custom-XML relationship,
  internal target mode, declared `text/xml`, action root, and graph closure;
- shared target one-physical-replacement/all-inbound-owner behavior and
  distinct target isolation;
- case-equivalent target resolution with both original lexical target refs
  preserved;
- Strict PML and Strict relationship dialect with the fixed ECMA MCE URI;
- inactive/unknown MCE choices, fallback `p:pic`, arbitrary namespace
  spelling, and admitted definitions/transform/trace bytes preserved exactly;
- valid unknown internal outbound relationship retained as a diagnostic and
  valid unknown external outbound relationship retained without traversal,
  fetch, or drop;
- scalar readback, action/group counts, target bytes, source/semantic ordinals,
  and inbound/outbound edge vectors;
- exact no-op empty patch and source allocation/byte sharing where public
  slices make pointer comparison possible;
- changed publication atomicity, save/reopen typed readback, and exact inverse
  restoration; and
- stale source, signature policy, malformed/limit refusal, and source
  reopenability without partial publication.

Unknown or external outbound edges are valid only for read/scalar-edit lanes;
they are retained and diagnosed, never followed. A graph-changing operation
that would affect an unmodeled outbound closure must refuse with its typed
error. A signed changed patch must return
`Error::Opc(OpcError::SignedSourceRequiresExplicitPolicy)`; a signed no-op must
preserve the signature edge and serialized source. Stale recipes mutate one
owner XML, owner `.rels`, or target and must return `Error::StaleSource` (or the
exact documented stale closure error). The content-type recipe mutates the
target content type and must return the actual typed `Error::ContentType`
refusal. A retained OPC read-limit mismatch is not in this initial 42-lane
matrix; it remains a separately recorded input to `from_vec_with_limits`.

Refusal samples record the actual public error variant, resource, and numeric
limit structurally; bounded debug/display text is retained only as diagnostic
context. They are valid only if all source members remain unchanged and
reopenable.

## Disjoint timing and memory phases

The initial measurement scope is absolute elapsed latency, requested
allocation bytes, incremental allocator peak-live bytes, and process maximum
RSS from `/usr/bin/time -v`. Every sample is retained. p50, p95, and p99 are
descriptive process-uncertainty summaries, not confidence intervals or a
general performance claim. Hardware counters, flamegraphs, I/O tracing, and
parallel scaling are outside the initial scope.

The harness snapshots these disjoint boundaries:

1. `setup`: fixture construction, `from_opc_package`, `from_vec_with_limits`
   source preparation, source hashes, retained limits, snapshots, patches,
   and expected values;
2. `operation`: only the named public call, including its own bounded
   validation/output allocation; `apply` includes the current owner’s source
   revalidation, package-wide relationship index/XML validation, staged clone,
   candidate inventory/reopen, and publication;
3. `validation`: semantic readback, source/member hashes, pointer observations,
   and expected typed-error checks, all outside the operation timer;
4. `drop`: operation result/error and validation temporaries dropped while the
   prepared source/package baseline remains retained; and
5. `postdrop`: the baseline is reopened with `Package::from_vec_with_limits`
   using retained OPC limits, read through public `Package::ink_actions()`,
   checked for the expected anchor count, and manifest-reopened before the
   retained baseline is released.

For each phase `P`, the process-local allocator must record direct allocations
`A_P`, successful realloc-new bytes `R+_P`, realloc-old bytes `R-_P`, and
deallocations `D_P`, then verify:

```text
live_after_P = live_before_P + A_P + R+_P - R-_P - D_P
requested_alloc_P = A_P + R+_P
peak_live_delta_P = peak_live_during_P - live_before_P
```

`live_before_operation` is the retained prepared set. For an apply lane, setup
keeps one baseline `Package` and builds a second fresh
`Package::from_opc_package` working package after the baseline live snapshot;
the working package is setup-only, and is dropped in the named `drop` phase.
The operation timer therefore contains only patch publication, while its
`live_before` includes the retained working package that will be released in
`drop`. After drop, the retained baseline must be accounted for explicitly;
allocator cache
behavior means equality is a bounded accounting check, not a language-level
leak proof. Failed allocations, underflow, and invalid counters are separate
receipt fields. The harness expects
`tracked_live_after_drop = tracked_live_before_operation` after operation and
validation temporaries are released, then records
`tracked_live_after_postdrop = tracked_live_after_drop - baseline_release` as
the separate baseline-release boundary. Pointer sharing describes only the
observed retained snapshot, not a general zero-copy guarantee. Process RSS is
the per-sample maximum from `/usr/bin/time -v`; it is not added to allocator
bytes.

## Receipt and provenance gate

Each sample records at least:

```text
schema, lane, recipe_id, source_commit, semantic_owner_commit,
production_source_baseline_commit, capture_head, helper_sha256, generator_sha256,
semantic_owner_design_path, semantic_owner_design_sha256,
semantic_owner_design_git_blob,
fixture_bytes, fixture_fnv1a64, fixture_sha256,
retained_opc_limits, owner_limits,
source_bytes, owner_xml_bytes, unique_target_bytes, output_bytes,
retained_owner_xml_bytes, retained_target_bytes, retained_profile_bytes,
shared_pointer_observation,
anchors, unique_targets, inbound_edges, outbound_edges,
outbound_diagnostic_modes, unknown_internal_outbound_preserved,
unknown_external_outbound_preserved,
action_count, action_group_count, semantic_ok, preservation_ok,
inverse_ok, expected_error, actual_error_type, actual_error_resource,
actual_error_limit,
elapsed_ns, requested_alloc_bytes, peak_live_delta_bytes,
rss_max_kib, allocation_equation_ok, process_id, warmup
```

`generator_sha256` is the SHA-256 of the retained bounded OPC generator source
(`harness/adapter.rs`); `helper_sha256` must match the committed helper hash in
the manifest. Unavailable fields are `null`, never measured zero.

The later sealed run requires three fresh processes per lane, two warm-ups,
and twenty measured samples per process: 126 process launches, 252 warm-ups,
2,520 measured calls, and 2,772 operation calls at most. It also requires a
clean isolated checkout,
authoritative isolated `harness/Cargo.lock`, full source/toolchain/host/
binary/fixture hashes, commands, stderr/exit status, and before/after source
manifests. It rejects dirty/untracked production inputs, source drift,
nonempty `RUSTFLAGS`/bootstrap overrides, missing/duplicate receipts,
allocator equation failures, unexpected typed errors, and binary hash changes.
No build or timing command is authorized for this scaffold before commit and
review.

Before any later build, the implementation must materialize a tracked
`harness/Cargo.lock` from the isolated harness manifest and record its
SHA-256. A missing, dirty, or repository-root lockfile substitution fails
preflight. The isolated harness manifest and lockfile are materialized in this scaffold;
the later sealed runner still refuses any root-lockfile substitution.

The matrix preflight must assert 23 unique recipe IDs, 42 unique lane IDs,
every recipe is referenced by at least one lane, and every lane references an
existing recipe. A duplicate, orphaned, or unresolved entry is a manifest
failure before any build or timing command.
