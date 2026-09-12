# PPTX existing InkAction owner performance plan

Status: **implementation scaffold present; no release timing run has been performed**.

This profile is a bounded absolute observation of the committed
source-backed existing-target `iact:actions` owner at semantic commit
`cf6fdb8e91dd232d7d762596763d2e9d8a5b9dbd`. Production transitive source and
workspace/build inputs are pinned to baseline
`2a2ffa1cae4e6b7070082768ce84483e5d411dc8`; the later capture HEAD is recorded
at runtime. Its initial set is exactly 23
named recipes and 42 named lanes in `corpus-manifest.json`. There is no
Cartesian expansion, generator-selected scale, or hidden lane. The profile
does not make a blanket `docs/GOAL.md` performance claim; it may report scoped
Litchi host observations for these recipes after the later run gate passes.

Planning consulted the workspace `docs/GOAL.md` and follows accepted ADRs
0001, 0003, 0005, and 0006. The ADRs are current capture context and are
hashed against the runtime capture HEAD. Because that GOAL file is untracked,
it is recorded for context but excluded from the reproducible source input
set; the committed `source-contract.json` is the profile contract. The
semantic-owner contract is the explicit
[`pptx-ink-actions-design.md`](../pptx-ink-actions-design.md) input at Git blob
`597400950b1027c47cd6e4cbbedd23915bc0980e` and SHA-256
`30b78cca84c4ca24ae44f3d3694c3097f54b5e5a1f2004af9ce5007bcaf4173d`. The owner is
the slide MCE anchor plus its relationship, content-type, target, and inbound
graph closure. The shared DrawingML profile is used only after that closure
has been validated.

## Source and fixture authority

The later run must use a clean isolated checkout whose exact capture HEAD
descends from both explicit pins above. The synthetic OPC helper retained in
the semantic owner commit is
`crates/litchi-pptx/tests/pptx_ink_actions.rs`; its committed SHA-256 is
`bec6baafcf735d778216fb54fe6299312707e912dcaed5f89de988f7112bb58e` and its
Git blob ID is `ad7e43e8c2c362c1b9e1938806f59b9fb3ab1dea`. That helper is the
fixture authority and must be hashed again from the isolated checkout before
any future build. The adapter may retain a separate OPC recipe generator only
inside this profile directory; once present, its path and SHA-256 become
required source inputs. The implementation scaffold also includes the
executable harness and its independently locked Cargo manifest.

The helper builds a complete synthetic owner closure with a slide, a
PresentationML presentation relationship, owner `.rels`, target part, and
`text/xml` content type. It deliberately has no native PowerPoint provenance.
The checked corpus has no native `inkAction` package. The profile therefore
excludes native PowerPoint acceptance, rendering, playback, recognition,
producer-specific MIME/path policy, and interoperability claims. It may state
what the Litchi package/presentation host did for the named synthetic OPC
graphs.

## Public API and setup boundary

Every positive read/edit/publication lane starts from a fresh mutable package.
Fixture construction and all expected values are prepared before the timed
operation:

1. Build the deterministic `OpcPackage` recipe from the retained helper
   shape. For package/edit/apply/save lanes, construct the public mutable
   package with `Package::from_opc_package(opc)` before timing. Do not clone or
   reuse a package that a previous sample mutated.
2. Capture source snapshots and prepare selectors/commits outside the timer
   when the lane names a prebuilt patch. `Snapshot::edit()`, scalar mutation,
   and `Edit::commit()` are timed only in the edit/no-op lanes.
3. For an archive-reopen lane, serialize a prepared source or candidate once
   during setup, then time `Package::from_vec_with_limits(bytes, retained_opc)`
   followed by the selected public read. `from_vec_with_limits`'s retained
   OPC `ReadLimits` are recorded separately from
   `ink_actions::Limits`, which bound anchors, target bytes, aggregate target
   bytes, and target relationship edges.
4. For save lanes, use a newly prepared mutable package and time
   `Package::to_bytes()` only. Save output hashing, package-member manifests,
   and semantic reopen are validation work after the timer unless the lane is
   explicitly `save_reopen`.
5. Stale and signed lanes prepare two raw OPC graphs before public wrapping.
   For a stale lane, build the baseline `OpcPackage`, wrap it with
   `Package::from_opc_package`, create the source snapshot/patch, and
   independently mutate the second raw `OpcPackage`'s owner XML, owner
   `.rels`, target blob, or target content type. Only then wrap the mutated
   graph with a fresh `Package::from_opc_package`. For a signed lane, put the
   same signature edge on both raw graphs before wrapping, then prepare the
   changed or exact no-op patch from the signed baseline. The timed operation
   receives an already-prepared public package and only applies the patch; no
   stale mutation or raw OPC edit is hidden inside the timer.

Apply and inverse lanes retain a fresh baseline `Package` for the source-preservation
receipt and construct a second fresh `Package::from_opc_package` working
package during setup. The retained-baseline live snapshot is taken after the
baseline package and manifest are ready but before that setup-only working
package; the working package is released in the named drop phase. This keeps
the timed apply/inverse calls limited to patch publication while making the
post-drop baseline equation explicit. Inverse publication may change internal
capacities even when it restores exact source bytes; those working allocations
belong to the operation package, not the retained baseline.

The public routes under test are `Package::{ink_actions,
ink_actions_with_limits, apply_ink_actions_patch, to_bytes,
from_opc_package, from_vec_with_limits}` and
`Presentation::{ink_actions, ink_actions_with_limits}`. A presentation handle
is prepared before a `Presentation`-route timer. The `Package` and
`Presentation` inventories are reported as separate Litchi host routes; a
neutral fragment parser is never substituted for either.

## Exact bounded recipes

The adapter must implement exactly the recipes listed here. `A` is active
anchor count in the selected slide set, `T` is bytes in one target, `U` is
unique target count, `E` is selected-target inbound owner-edge count, and `S`
is slide count. Outbound diagnostics are recorded separately; for example,
`r14_opaque_unknown_external` has one inbound owner edge plus one retained
unknown internal outbound edge and one retained unknown external outbound edge.
All byte and limit values are positive.
`T` is the actual target-part blob length after deterministic padding, not a
requested nominal size; setup records the resulting member length and rejects
the recipe if it does not equal the manifest value.

| recipe | shape | purpose/required source variant |
|---|---|---|
| `r01_tiny_shared` | `A=1,T=1024,U=1,E=1,S=1` | baseline Transitional existing target |
| `r02_small_shared` | `A=8,T=16384,U=1,E=8,S=1` | shared target and scalar edit |
| `r03_small_distinct` | `A=8,T=1024,U=8,E=8,S=1` | distinct target retained bytes |
| `r04_medium_shared` | `A=64,T=262144,U=1,E=64,S=1` | shared target and package apply |
| `r05_medium_distinct` | `A=64,T=16384,U=8,E=64,S=1` | distinct aggregate target bytes |
| `r06_large_shared` | `A=256,T=1048576,U=1,E=256,S=1` | larger owner scan/target parse |
| `r07_large_distinct` | `A=256,T=16384,U=32,E=256,S=1` | larger distinct graph and target cache |
| `r08_near_shared` | `A=1024,T=8388608,U=1,E=1024,S=1` | bounded near-limit shared source |
| `r09_near_distinct` | `A=1024,T=65536,U=128,E=1024,S=1` | bounded near-limit distinct aggregate |
| `r10_multislide_shared` | `A=2,T=16384,U=1,E=2,S=2` | two slide owners share one target |
| `r11_multislide_distinct` | `A=2,T=16384,U=2,E=2,S=2` | two slide owners use distinct targets |
| `r12_case_equivalent_shared` | `A=2,T=16384,U=1,E=2,S=1` | two lexical target refs resolve one physical target; preserve both tokens |
| `r13_strict_shared` | `A=8,T=16384,U=1,E=8,S=1` | Strict PML/relationship dialect with fixed MCE URI |
| `r14_opaque_unknown_external` | `A=1,T=16384,U=1,E=1,S=1` | inactive/unknown MCE material, opaque action descendants, one valid unknown internal outbound relationship, and one valid unknown external outbound relationship |
| `r15_signed` | `A=1,T=1024,U=1,E=1,S=1` | package-root signature edge; no-op allowed, changed edit refused |
| `r16_stale_owner` | `A=1,T=1024,U=1,E=1,S=1` | mutate owner XML after commit |
| `r17_stale_owner_rels` | `A=1,T=1024,U=1,E=1,S=1` | mutate owner `.rels` source after commit |
| `r18_stale_target` | `A=1,T=1024,U=1,E=1,S=1` | mutate target bytes after commit |
| `r19_stale_content_type` | `A=1,T=1024,U=1,E=1,S=1` | mutate target content-type source; expect typed `ContentType` refusal |
| `r20_limit_anchor` | `A=8,T=16384,U=1,E=8,S=1` | anchor limits `7/8/9` for one-under/exact/one-over |
| `r21_limit_target` | `A=1,T=16384,U=1,E=1,S=1` | target-byte limits `16383/16384/16385` |
| `r22_limit_aggregate` | `A=2,T=16384,U=2,E=2,S=1` | aggregate-byte limits `32767/32768/32769` |
| `r23_limit_graph` | `A=8,T=16384,U=1,E=8,S=1` | edge limits `7/8/9` |

The owner hard ceilings remain 4,096 anchors, 16 MiB per target, 256 MiB
aggregate unique target bytes, and 4,096 selected target relationships. The
shared profile ceilings remain 16 MiB source, 100,000 nodes, 65,536 actions,
and 16,384 groups. `r08` and `r09` are deliberately below every ceiling.
The retained OPC `ReadLimits` are an independent finite policy and are never
replaced by owner limits.

## Exact 42-lane initial matrix

The initial lane set is exactly the following 42 IDs; each maps to one recipe
above. A future implementation may fix a correctness defect in a recipe but
may not add a Cartesian scale point without a reviewed plan update.

| lanes | recipe | timed public operation |
|---|---|---|
| `package_read_tiny_shared`, `presentation_read_small_shared` | `r01_tiny_shared`, `r02_small_shared` | package or prepared presentation inventory |
| `package_read_medium_shared`, `package_read_small_distinct` | `r04_medium_shared`, `r03_small_distinct` | package inventory |
| `package_read_large_shared`, `package_read_large_distinct` | `r06_large_shared`, `r07_large_distinct` | package inventory |
| `package_read_near_shared`, `package_read_near_distinct` | `r08_near_shared`, `r09_near_distinct` | package inventory |
| `package_read_multislide_shared`, `package_read_multislide_distinct` | `r10_multislide_shared`, `r11_multislide_distinct` | package inventory across slide owners |
| `package_scalar_edit_small_shared`, `presentation_scalar_edit_small_shared` | `r02_small_shared` | snapshot scalar edit plus `commit()` |
| `package_noop_small_shared`, `presentation_noop_small_shared` | `r02_small_shared` | exact scalar no-op plus `commit()` |
| `package_apply_medium_shared`, `package_apply_medium_distinct` | `r04_medium_shared`, `r05_medium_distinct` | `Package::apply_ink_actions_patch()` |
| `package_inverse_small_shared`, `package_inverse_case_equivalent` | `r02_small_shared`, `r12_case_equivalent_shared` | forward and inverse patch application |
| `package_save_medium_shared`, `package_save_reopen_medium_shared` | `r04_medium_shared` | `to_bytes()`; save then `from_vec_with_limits` and read |
| `stale_owner`, `stale_owner_rels`, `stale_target`, `stale_content_type` | `r16`, `r17`, `r18`, `r19` | changed patch application with one changed source input |
| `signed_noop`, `signed_changed_refusal` | `r15_signed` | exact no-op or changed patch on signed package |
| `opaque_mce_scalar_edit`, `opaque_mce_save_reopen`, `unknown_outbound_read_edit` | `r14_opaque_unknown_external` | Package patch publication, save/reopen, and read/edit with retained MCE and internal/external outbound diagnostics |
| `strict_shared_edit` | `r13_strict_shared` | strict read and existing-target scalar edit |
| `limit_anchor_one_under`, `limit_anchor_exact`, `limit_anchor_one_over` | `r20_limit_anchor` | owner anchor limits `7`, `8`, `9` |
| `limit_target_one_under`, `limit_target_exact`, `limit_target_one_over` | `r21_limit_target` | owner target-byte limits `16383`, `16384`, `16385` |
| `limit_aggregate_one_under`, `limit_aggregate_exact`, `limit_aggregate_one_over` | `r22_limit_aggregate` | owner aggregate-byte limits `32767`, `32768`, `32769` |
| `limit_graph_one_under`, `limit_graph_exact`, `limit_graph_one_over` | `r23_limit_graph` | owner edge limits `7`, `8`, `9` |

The table contains 42 concrete IDs even where one row groups related IDs. No
lane uses a zero limit: one-under values are nonzero and exact values are
expected to succeed when the source fits exactly; one-over values are a
positive control for the same source.

## Valid graph and preservation variants

The recipe variants have these precise meanings:

- **Shared:** at least two inbound internal relationships resolve to one
  physical target. Retain every inbound edge; a scalar target change is one
  physical replacement visible from every owner. Charge target bytes once.
- **Distinct:** each named target has its own physical part and retained source
  allocation. Charge each unique target against aggregate bytes.
- **Case-equivalent shared:** original relationship target strings differ in
  case or lexical relative spelling but OPC resolution maps both to one
  physical `PackURI`. Preserve the original strings and update one target.
- **Strict:** use Strict PresentationML and Strict relationship namespaces,
  the Strict `customXml` relationship URI, and the fixed ECMA MCE URI. The
  Transitional relationship URI is a typed mismatch.
- **Opaque MCE:** retain inactive/unknown choices, `mc:Fallback` `p:pic`,
  arbitrary prefixes/default bindings, and unknown source bytes. Only the
  selected target scalar may change.
- **Unknown outbound:** an action target may have an unrecognized outbound
  internal relationship to a retained package part. It is valid for typed read
  and scalar profile edit when the edge is retained as a diagnostic and is not
  traversed. A graph-changing operation or deletion would refuse until its
  closure is modeled. `r14` includes one such internal edge and records its
  retained target/member identity without parsing that target as an action
  closure.
- **External outbound:** an unknown external outbound relationship is retained
  as inert metadata; no network or target fetch occurs. Its presence is not a
  malformed target and must not make a scalar edit drop the edge. `r14` also
  includes one external edge and records the two modes separately.
- **Signed:** a package-root digital-signature-origin relationship remains in
  the source. An exact no-op preserves it; a changed publication returns the
  actual `SignedSourceRequiresExplicitPolicy` error before mutation.
- **Stale:** mutate exactly one owner XML, owner `.rels`, or target bytes after
  creating the patch. Apply must return `StaleSource` (or the exact
  owner-documented typed stale error), leave every source member unchanged, and
  require reload before a new edit. The content-type lane is a separate
  refusal: changing the target content type must return the actual typed
  `Error::ContentType` result before publication. A retained OPC read-limit
  mismatch is outside this initial lane set; it remains a source field that a
  future reviewed lane may exercise.

## Bounded execution budget

The later sealed run has 42 lanes, three fresh processes per lane, two warm-up
calls per process, and twenty measured calls per process. Its fixed maximum is
126 process launches, 252 warm-up calls, 2,520 measured calls, and 2,772
operation calls in total. A lane therefore has exactly six warm-up and sixty
measured receipts across its three processes. The runner does not add samples
for a large recipe, retry failed correctness, or expand a recipe into a
Cartesian product. A correctness failure stops that lane and is reported as a
failed receipt rather than silently spending an unbounded retry budget.

The implementation has materialized a tracked `harness/Cargo.lock` from the
isolated harness manifest and records its SHA-256 in the source/build-input
manifest. A missing, dirty, or root-lockfile substitution fails preflight.
Compile and host-probe checks validate this scaffold; release timing remains
gated on review and explicit authorization.

## Timed scope and disjoint phases

The initial profile records only absolute elapsed latency, requested allocator
bytes, incremental allocator peak-live bytes, and whole-process maximum RSS.
It retains every sample and reports p50, p95, and p99 as descriptive process
uncertainty summaries. It does not require a large hardware-counter,
flamegraph, I/O-tracing, or parallel-scaling framework. Source/member hashes,
semantic checks, and allocation shape fields are correctness/provenance data,
not additional performance claims.

Each sample has disjoint phases:

1. **Setup:** fixture generation, package serialization, helper hash checks,
   `Package::from_opc_package`, retained `ReadLimits`, owner `Limits`, source
   snapshot/patch preparation, and expected-value construction. Setup is not
   attributed to the operation.
2. **Operation:** only the named public call in the lane table. For apply and inverse,
   this includes the current Litchi owner’s source revalidation, package-wide
   relationship index/XML validation, staged package clone, target install,
   candidate inventory/reopen, and atomic publication. Inverse lanes execute
   that publication scope for both the forward and inverse calls. The profile may report
   this observed host scope; it must not infer a parser-pass count by
   subtracting unrelated timers.
3. **Validation:** typed readback, member/source manifests, semantic
   preservation, pointer identity observations, and expected-error comparison.
   Validation allocations are recorded outside the operation timer.
4. **Drop:** drop operation results/errors and validation temporaries while the
   prepared source/package retained baseline remains live. Record this phase
   separately and do not subtract its deallocations from operation totals.
5. **Post-drop:** reopen the retained baseline with
   `Package::from_vec_with_limits` under retained OPC limits, perform the
   public `Package::ink_actions()` read and expected-anchor check, verify the
   source manifest, then release the baseline after the receipt has been
   serialized. For stale fixtures, use the untouched generated source bytes
   and source manifest for this check. The separately mutated candidate bytes
   and manifest remain the evidence for refusal and candidate preservation;
   an expected typed refusal is not a successful source-baseline reopen.

The process allocator records cumulative counters at every boundary. For any
phase `P`, with direct allocations `A_P`, realloc-new bytes `R+_P`,
realloc-old bytes `R-_P`, and deallocations `D_P`, the verifier checks:

```text
live_after_P = live_before_P + A_P + R+_P - R-_P - D_P
peak_live_delta_P = peak_live_during_P - live_before_P
requested_alloc_P = A_P + R+_P
```

`live_before_operation` is the retained prepared baseline, not process zero.
The expected post-drop relation is checked against the same retained baseline
after operation results and validation temporaries are dropped. Allocator
caches and process teardown mean this is not a language-level leak proof. The
receipt must expose failed allocations and invalid/underflow flags separately.
The receipt also exposes retained owner XML, unique target bytes, typed profile
payload, and shared-target pointer observations so duplication is visible in
the baseline. Apply lanes retain the current package-wide relationship index,
source revalidation, staged clone, and candidate graph reopen in the timed
operation; the report may describe that observed graph-parse/allocation cost,
but it must not infer a parser-pass count or a zero-copy guarantee from
unrelated timers.

For the retained-baseline check, the harness records
`tracked_live_before_operation = retained_baseline_live` and expects
`tracked_live_after_drop = retained_baseline_live` after operation results and
validation temporaries are released. The later `postdrop` boundary records
`tracked_live_after_postdrop = tracked_live_after_drop - baseline_release`;
baseline release is separate from operation/drop totals. Whole-process RSS is
`rss_max_kib = max(/usr/bin/time -v Max RSS)` for the sample and is never
added to phase allocations or treated as a leak proof.

## Provenance and later run gate

No release timing capture may run before this scaffold is committed and
reviewed and the parent records the production-freeze decision. The harness
owns an isolated `harness/Cargo.lock`; it is the authoritative profile
lockfile and is hashed as a build input. The repository root lockfile is not
silently reused as the profile lock. The isolated run must capture the full
source/build-input manifest, owner/helper/generator hashes, lockfile hash,
Rust/Cargo/toolchain,
target and flags, allocator, CPU/host/OS/load, binary SHA-256, and all recipe
package hashes before and after execution.

The later authorized capture uses three fresh processes per lane, two warm-up
samples, and twenty measured samples per process. It rejects dirty or
untracked production inputs, source drift, missing/duplicate receipts,
nonempty `RUSTFLAGS`/bootstrap overrides, allocator equation failures,
unexpected typed errors, and executable hash changes. It may delete only a
new isolated target that it created. This scaffold has no release binary, host
receipt, or timing result; the harness lock is present and hashed above.

The sealed verification report retains the host/load metadata, every setup,
build, preflight, lane, postflight, and verifier command's start/exit status,
argv, environment, and stdout/stderr paths, together with each `/usr/bin/time
-v` user/system/elapsed field. A status-zero run writes an exact
verification-success sentinel only after that report passes; cleanup deletes a
target only when both the ownership and verification-success sentinels match
the canonical fresh target path.
