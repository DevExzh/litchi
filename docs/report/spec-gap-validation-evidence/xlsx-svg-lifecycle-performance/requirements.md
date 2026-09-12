# XLSX ordinary worksheet SVG lifecycle profile requirements

This is a requirements and runner scaffold. It is not a completed profile.
No receipt, table row, or narrative in this directory may be presented as a
measurement until the source-backed XLSX owner, its focused correctness tests,
and the public selector / attach / detach API are frozen.

## Scope and evidence boundary

Most inputs are OPC XLSX packages with one ordinary worksheet drawing part. The
`multisheet_attach_detach` input is the deliberate two-worksheet exception,
with one independent drawing part per worksheet. The profile admits direct
`xdr:pic` owners beneath `twoCellAnchor`, `oneCellAnchor`, or
`absoluteAnchor`. Every picture has an existing internal PNG fallback.
Attach adds one embedded internal SVG owner; detach removes only the selected
SVG owner and removes a media leaf only after package-wide incoming-edge
reachability proves it is unreferenced. The selector is semantic and
source-bound. It does not accept an `rId`, package URI, XML offset, generated
media name, or extension prefix from the caller.

The native fixture is a producer-shape and shared-target read case, not an
acceptance or speed fixture:

| item | value |
|---|---|
| retained path | `fixtures/tdf169496_hidden_graphic.xlsx` (unchanged LibreOffice corpus input; see `fixtures/README.md`) |
| archive bytes | 12,470 |
| archive SHA-256 | `0b647da300a085f39914fdfae961463ae9e54ffe772b2e0eb9860a841ab93f72` |
| drawing member | `xl/drawings/drawing1.xml` |
| drawing SHA-256 | `c9d4149c14d847d4979239fc9771de0f34383c5e01ee505e270aea1bce05ffe8` |
| relationship member | `xl/drawings/_rels/drawing1.xml.rels` |
| relationship SHA-256 | `308917dc6427cacfe0ef7a7cc4b74fc058b7b31e21289de1055af09e59e675ce` |
| SVG member | `xl/media/image2.svg` |
| SVG SHA-256 | `05769bd518f6ce504ee4f465a878eb8bd848c906e23c1adee6491e1653167cbf` |

The profile does not establish native Excel or LibreOffice acceptance,
rendering behavior, SVG security semantics, image conversion cost, external
fetching, or general XLSX performance. It also does not infer an algorithmic
complexity claim from an inventory lane. A lane is an absolute observation of
the named bounded workload until a matched baseline exists.

## Matrix

The lane list is duplicated in `harness/adapter.rs`, `run_profile.sh`, and
`summarize.py`; `verify.py` rejects drift between receipts and that list.
The `{anchor}` values are `two_cell`, `one_cell`, and `absolute` and map to
the corresponding source XML element. The `{size}` values are `small` and
`large`, defined in `corpus-manifest.json`.

| lane family | required cases | timed scope |
|---|---|---|
| `capture_native_fixture` | LibreOffice producer fixture | open the retained native package, scan its ordinary drawing, and check the two shared-owner pictures |
| `capture_raster_{anchor}_{size}` | all 3 anchor forms × 2 payload sizes | open package, source-backed picture capture, and the requested read projection |
| `capture_attached_{anchor}_{size}` | all 3 anchor forms × 2 payload sizes | open package, source-backed SVG capture, fallback and owner assertions |
| `clone_raster_{size}`, `clone_attached_{size}` | small and large; adapter must exercise representative all-3-anchor fixtures | open/capture plus the named immutable snapshot clones; not a clone-only claim |
| `clone_captured_owner_{size}` | small and large; all-3-anchor source present | pre-capture and validate one `PictureSource` before the timer; time only the clone of that captured owner, without workbook open/save or package publication |
| `inventory_shared_{256,1024}` | many direct pictures sharing one raster and one SVG target | worksheet/drawing picture inventory and every descriptor check |
| `inventory_distinct_{256,1024}` | many direct pictures with distinct SVG targets | same inventory checks with distinct target metadata |
| `namespace_heavy` | inherited bindings and opaque descendants below configured limits | owner capture on a namespace-heavy direct picture |
| `namespace_limit_refusal` | syntactically complete source at the configured active-binding limit | typed pre-allocation refusal; earlier package/XML rejection is a failed lane |
| `attach_end_to_end_{anchor}_{size}` | all 3 anchor forms × 2 payload sizes | capture, fallible commit/closure validation, publish, reopen, and semantic package checks |
| `strict_attach_end_to_end_two_cell_small` | strict worksheet/drawing dialect | attach, commit, publish, reopen, and exact strict worksheet/drawing relationship checks |
| `inverse_attach_detach_{anchor}_{size}` | all 3 anchor forms × 2 payload sizes | attach and publish, reopen the forward result, apply the source-bound in-memory inverse, require exact source restoration, and refuse replay on the already-published snapshot |
| `detach_end_to_end_shared_first_{anchor}` | all 3 anchor forms | detach one of two owners; shared SVG edge and part must remain |
| `detach_end_to_end_shared_final_{anchor}` | all 3 anchor forms | detach final owner; reachability and content-type checks may remove SVG |
| `strict_detach_end_to_end_shared_final_two_cell` | strict worksheet/drawing dialect | detach final owner, preserve strict host relationships, publish, reopen, and verify SVG cleanup |
| `incoming_edge_shared_final_two_cell` | one selected owner plus a separate incoming SVG relationship | remove the selected drawing-to-SVG relationship while a package-wide incoming edge retains the SVG leaf |
| `detach_end_to_end_distinct_{anchor}_{size}` | all 3 anchor forms × 2 payload sizes | detach a distinct target and reopen the changed closure |
| `same_picture_attach_detach_{anchor}` | all 3 anchor forms, small payload | attach one selected picture and publish; reopen, detach that same semantic picture and publish again; restore source drawing/content-types and close the graph |
| `multisheet_attach_detach` | Sheet1/two-cell picture and Sheet2/one-cell picture in one package | attach one owner on each worksheet in one public transaction, reopen, detach the same owners in a second transaction, and verify independent drawings/media and final graph cleanup |
| `noop_detach_{anchor}` | all 3 anchor forms | exact no-op path; source bytes and package snapshot must be shared |
| `limit_{small,large}` | caller output/edit budget just below required staging | refusal before oversized temporary output allocation or publication |
| `mixed_caps_rejection` | shared attached fixture under each retained aggregate cap | run bounded detach+attach attempts at exact part, byte, relationship, relationship-XML, relationship-event, and content-type ceilings; require atomic refusal and unchanged source for every cap |
| `malformed_duplicate_owner` | two admitted direct owners in one picture | safe read refusal for edit; no output publication |
| `malformed_mce_owner` | owner below `mc:AlternateContent` | safe read preservation and edit refusal |
| `malformed_linked_owner` | `r:link` SVG owner | linked/inert read and first-profile edit refusal |
| `malformed_unknown_uri` | unknown extension URI or foreign SVG namespace | inert preservation, no inferred resource, and exact detach no-op |
| `multi_picture_same_drawing_{16,64,256}` | 16, 64, and 256 attach intents in one ordinary drawing transaction | composed attach, commit, publish, reopen, and all-owner checks; exposes per-intent rescan and cumulative staging cost without a scaling claim |

The freeze-gated acceptance set has 69 lanes. The adapter also accepts six
exploratory-only lanes,
`multi_picture_same_drawing_detach_{shared,distinct}_{16,64,256}`. Each lane
detaches every direct picture in one transaction and checks the final raster
fallback, SVG reachability cleanup, commit, and reopen. They remain outside
the freeze-gated 69-lane acceptance set and carry no scaling claim.

The adapter also accepts the exploratory-only
`inventory_shared_root_namespace_32` lane. It inventories 32 direct pictures
with distinct SVG relationship IDs beneath one drawing root carrying 128
large inherited namespace values. The lane checks the source-backed owner
shape and retained source visibility for all 32 owners, including one
raw-source and one namespace-context projection per picture; it is paired with
the retained scope-probe receipt in `../xlsx-svg-source/` and carries no
timing, peak-memory, or complexity claim.

`capture`, `clone`, and `inventory` are isolated operations. `attach`, inverse,
and `detach` end-to-end lanes include the public fallible commit path, complete
dependency-closure validation, publication, and candidate reopen. A future
adapter may add explicit stage receipts for source scan, splice, graph plan,
and reopen; it must not call those stages measured when the API only exposes a
single opaque operation. The multi-picture lanes stage several selected
pictures on one drawing in one transaction. They are the planned cost lanes
for the current composed planner, whose implementation still rescans the
source per intent; they must remain bounded workload observations and carry no
linear-scaling claim. The inverse lanes also include an in-memory inverse
publication and replay refusal in their timed operation; the mixed-cap lane is
a composite refusal check and is not a per-cap timing comparison.

## Deterministic fixture contract

The adapter generates packages beneath its own temporary directory from the
recipes in `corpus-manifest.json`. The package recipe must be deterministic
for a lane and must report byte length, FNV-1a-64 hash, and a SHA-256 digest of
the exact input identity in every raw receipt. Fixture setup, payload
construction, and expected-result derivation are outside an isolated operation
timer. Payload bytes are opaque; the host
must not parse or render the SVG during this profile.

Each single-worksheet synthetic package contains only the parts needed for the
lane: workbook, worksheet, worksheet relationship, one drawing, drawing
relationships, a PNG fallback, optional SVG leaves, and content types. The
multisheet fixture adds a second worksheet, worksheet relationship, drawing,
and PNG fallback with the same bounded shape. Anchor geometry must remain
observable in the source-backed readback:

- two-cell fixtures retain validated `from`, `to`, and `editAs`;
- one-cell fixtures retain validated `from` and EMU `ext`; and
- absolute fixtures retain validated EMU `pos` and `ext`.

The shared-target fixtures use two or more direct pictures referencing the same
SVG relationship/part. The final-owner lane must perform a package-wide
incoming edge scan before deleting that part. The many-picture fixtures use
256 and 1,024 direct pictures, keep the PNG relationship shared, and vary only
the SVG-target topology. Namespace-heavy fixtures use 252 inherited bindings
and 1,024 opaque descendants as a bounded workload, matching the
resolver-pressure shape used by the existing source-backed profiles. The
namespace refusal fixture is derived by probing the admitted host policy and
binary-searching the first refused active-binding count; the receipt records
`namespace_generated_bindings`, `namespace_active_bindings`, and
`namespace_active_limit`. The generated count covers only declarations added
by the pressure fragment; the active count includes the seven fixed drawing
root bindings. The first refused fixture must satisfy
`active = generated + 7 = active_limit + 1`, with the preceding generated
count accepted.

Shared-final fixtures are constructed directly with one selected picture
carrying the shared SVG owner. They do not call the production detach API to
prepare the input, so the detach lane measures only the requested lifecycle
operation. The separate incoming-edge fixture adds an opaque package part that
retains an internal relationship to the SVG leaf after the selected drawing
owner is removed.

Opaque extension and namespace-descendant fixtures assert the exact opaque
fragment bytes and expected occurrence count after every read, attach, inverse,
and detach operation; a `source() == None` observation alone is not treated as
memory or preservation evidence.

The native fixture is never mutated by the generator. It is copied only into
the adapter's private temporary workspace when a capture lane needs it. No
confidential or user documents may enter the retained evidence tree.

## Receipt and allocator contract

Every process receipt has schema
`xlsx-svg-lifecycle-profile-v1` and contains:

```text
lane, input_bytes, input_hash_fnv1a64, input_sha256, warmup, sample_count,
namespace_generated_bindings, namespace_active_bindings, namespace_active_limit,
caller_limits,
expected_success, samples[]
```

Each sample records:

```text
elapsed_ns,
requested_alloc_bytes, direct_allocated_bytes, realloc_new_bytes,
realloc_old_bytes, deallocated_bytes,
live_before, live_after, peak_live_delta,
alloc_balance_ok, alloc_invalid, alloc_failed,
actual_success, semantic_ok, output_exact, error,
stage_ns: {capture, clone, splice, graph_plan, commit_validate,
           publish, reopen},
source_bytes_read, range_read_count, range_request_bytes,
bytes_decompressed, bytes_recompressed, bytes_copied,
source_copy_bytes, staged_bytes, output_bytes,
write_call_count, write_call_bytes,
hardware_counters: {cycles, instructions, ipc, branches,
                    branch_misses, l1d_misses, llc_misses, page_faults}
```

The top-level receipt also records the caller-limit profile used by a lane:
`caller_limits` is either `null` for the default source-backed package, an
object with the borrowed-input ceiling for `limit_{small,large}`, or an object
whose `ceilings` array records each actual requested `ReadLimits` value for
`mixed_caps_rejection`. The composite names are `parts`, aggregate part
bytes, relationships, relationship XML bytes/events, and content-type
mappings; the verifier requires the complete ordered set and positive integer
values, and requires those values to agree across fresh processes. A future
range-source adapter must add the exact request-size list to
`range_request_bytes`; a null field means that the public source abstraction
did not expose that metric and must not be interpreted as zero.

`stage_ns` is optional for an operation that cannot expose a stage boundary,
but absent stages must be documented as unavailable rather than inferred from
the end-to-end timer. The I/O, copy, write, staged-byte, and
hardware-counter fields may be `null` when the platform or source abstraction
cannot observe them, but the adapter must distinguish unavailable counters
from measured zero. `source_copy_bytes` covers source snapshot copies and
`staged_bytes` covers cumulative XML, relationship, content-type, and media
staging when the production owner exposes those counters; neither may be
inferred from allocator totals. For range sources, `range_request_bytes`
retains the request-size distribution rather than only its sum.
`requested_alloc_bytes` is direct allocations plus new reallocation bytes. The
allocator equation is:

```text
live_after = live_before + direct_allocated_bytes
             + realloc_new_bytes - realloc_old_bytes
             - deallocated_bytes
```

The process-local counting allocator must report failed allocations and
underflow separately. Successful samples return to their starting live-byte
balance after fixture and assertion values are dropped. A refusal may retain
its typed error at the observation boundary; the adapter must drop it after
the receipt snapshot or record the retained bytes explicitly. An expected
refusal is a distinct adapter outcome: it is emitted only after the public
operation returns a classified refusal and the source bytes remain unchanged.
Unexpected acceptance, source mutation, setup failure, or an unrelated API
error is a hard lane failure and must not be relabeled as the requested
refusal class.

The captured-owner clone lanes start the timer immediately before cloning the
already prepared `PictureSource` and take the allocator snapshot before
post-timer byte, owner, and relationship assertions. The cloned owner remains
live through that snapshot, so `live_after` intentionally includes the clone;
the adapter drops it before receipt encoding and the next sample. This keeps
the clone itself in the measured allocation scope while keeping validation and
receipt construction outside the timed operation.

Three fresh processes run each lane, with two warm-up samples and twenty
measured samples by default. The runner may increase those counts, but never
decrease them for sealed evidence. Report p50, p95, and p99 elapsed values,
p50/p95 allocation and peak-live values, and min/max `/usr/bin/time -v` RSS.
Use confidence intervals or a stated robust uncertainty summary in the final
report. One noisy run is not evidence of a speedup or regression.

The reproducible `uncertainty.py` postprocessor is one permitted uncertainty
summary. It consumes only already verified raw receipts and `/usr/bin/time -v`
sidecars, requires exactly three process identities with the sealed warmup and
sample minimums, and checks the semantic, allocator, and typed-refusal fields
before deriving any values. For every lane it reports each process median,
within-process min/max range, between-process median range, and RSS values.
It writes derived Markdown and optional JSON files without changing raw
receipts. With n=3, those ranges are descriptive rather than confidence
intervals; this summary must not be used to claim a speedup, regression,
causal effect, or scaling law.

The optional exploratory phase decomposition is a separate binary and receipt
schema. For `multi_picture_same_drawing_{16,64,256}`, it measures only the
named public phases `open`, `stages`, `commit`, `firstsave`,
`reopen_secondsave`, and complete `validation`. Every phase records elapsed
time, phase-local requested allocation, allocator equations, and the live
bytes retained at its boundary. The phase harness must preserve the complete
semantic predicate and must expose retained live sets so phase allocation
values are not summed or subtracted across unlike lifetimes. It is outside the
69-lane acceptance matrix and cannot be used to make a baseline comparison or
speedup, regression, causal, or scaling claim.

## Runner and source-integrity requirements

`run_profile.sh` must:

1. require `PROFILE_FROZEN=1` and `XLSX_SVG_PROFILE_API_WIRED=1`;
2. refuse nonempty `RUSTFLAGS`, `CARGO_ENCODED_RUSTFLAGS`, and
   `RUSTC_BOOTSTRAP`;
3. build with a locked offline Cargo invocation and capture metadata before and
   after the run;
4. hash all Cargo package source files and scaffold inputs before and after;
5. use only a new, explicitly isolated Cargo target unless reuse is opted into;
6. run every lane in fresh processes under `/usr/bin/time -v`;
7. retain raw receipts and commands, but remove a target it created itself;
8. run `summarize.py` and `verify.py` only after all receipts are present; and
9. require `XLSX_SVG_PROFILE_RESULTS` to name a fresh nonexistent directory
   outside the checkout and disjoint from the Cargo target; and
10. fail if the source manifest, package input hash, allocator equations,
   semantic readback, output exactness, or expected refusal class changes.

The runner and verifier never delete the repository's main target, historical
evidence, or a pre-existing output directory. A future sealed run must record
CPU model, core count, memory, storage, OS, Rust/Cargo versions, compiler
flags, allocator, and relevant environment settings in
`$XLSX_SVG_PROFILE_RESULTS/build-provenance.txt`.

The default runner measures fresh processes and warm-up samples. If a future
adapter adds file-cache or thread-count lanes, it must identify cold versus
warm cache state and record 1, 2, 4, 8, and available-core settings where
parallel work is actually exposed. The ordinary lifecycle path is expected to
be serial until a production API demonstrates otherwise; no scaling or Amdahl
claim may be inferred from this scaffold.

## Required semantic gates

Before any receipt is considered valid, the adapter must assert:

- native fixture read exposes two pictures, unchanged PNG fallback, and the
  admitted embedded SVG owner;
- attach preserves the selected anchor geometry and raster fallback;
- attach adds one valid SVG relationship, media leaf, and content type;
- detach removes only the selected owner;
- shared SVG targets remain after first-owner detach and are removed only after
  final-owner reachability validation;
- strict/transitional relationship dialect and lexical source preservation are
  respected;
- strict worksheet and drawing relationship types remain intact through attach
  and final-owner detach;
- package-wide incoming SVG relationships keep a shared leaf alive after the
  selected drawing owner is detached;
- unknown, duplicate, MCE-wrapped, linked, and malformed owners are inert or
  refused according to the design;
- exact no-op publication shares source bytes;
- inverse restoration returns exact accepted source bytes and graph topology;
- stale source and replay are refused atomically; and
- same-picture and multisheet attach/detach candidates reopen byte-stably,
  preserve every selected and unselected anchor geometry, and retain exact
  opaque fragments after attach and restore;
- composite caller-limit receipts enumerate the complete ordered set of actual
  requested ceilings, and those values agree across fresh processes; and
- candidate reopen finds no dangling relationship, duplicate content-type
  override, or orphaned media owner.

The profile does not run until those checks are implemented in the production
owner's focused test suite and the adapter can call the resulting public API.
The native producer lane remains read/capture only; it never contributes to
authoring or Office-acceptance claims. A source pin and clean committed-input
guard are also required before any timing receipt is sealed.
