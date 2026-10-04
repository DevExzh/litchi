# 0908 — eager PPTX opportunity audit

This is a finite, read-only source audit of the eager multi-slide route. It
uses source commit `3d3650ac03be2cd8ba923a5a6e940e84b14f5eab`, the frozen
0907 native binary
(`e0a7707dcd849c92e01c8dc9115d61414fba10584bb8415e442f4c225828af84`), and
the first eager Callgrind artifact
`capture/eager-a/callgrind.out.1`
(`ab88e84ba3c227d78eda3f754149d94a39974480153496249df54847958e9ee4`). The
artifact records `Ir` only: `4,577,995,784` guest instructions under the
helper's exact-owner toggle. These are attribution counts, not native phase
fractions, latency, or a speedup claim. No build, workload, profiler run, code
change, or worktree was made for this review.

The corpus is the media-rich PPTX fixture with 229 OPC Parts, 445 physical ZIP
members, 17,568,429 uncompressed payload bytes, and Deflate compression. The
semantic receipt passed; its `229` eager materializations is a correctness
guard and does not establish the cost of any phase.

## The dominant owners

The raw first profile attributes the following owner edges:

| Owner | Ir | Share of profile | Source interpretation |
| --- | ---: | ---: | --- |
| `Package::to_bytes` | 3,189,476,327 | 69.67% | The final eager publication at `tools/perf-baseline/src/lib.rs:51115`. |
| `Package::opened_presentation_with_limits` | 994,945,242 | 21.73% | Opened snapshot and complete-revision work reached by the timed helper; the owner edge also covers work whose call context includes replay. |
| `Transaction::commit` | 42,871,786 | 0.94% | The eager transaction commit. |
| `Transaction::set_shape_texts` | 10,835,948 | 0.24% | Eight selected-slide staging calls. |
| `Package::apply_opened_presentation_commit` | 7,522,298 | 0.16% | The production eager commit publication wrapper. |

The `to_bytes` subtree is overwhelmingly ZIP Deflate. The raw call tree shows
`ReusableDeflateState::compress_once` at 2,927,199,198 Ir,
`flate2::mem::Compress::compress` at 3,060,606,421 Ir, and
`zlib_rs::deflate::algorithm::medium::deflate_medium` at 3,059,982,924 Ir;
these are nested inclusive rows and must not be added. The first eager
exclusive view further identifies `deflate_medium` and `longest_match` as the
large codec bodies. This is enough to reject a search for tiny name or
relationship comparisons as the next eager opportunity.

The opened owner is different. `opened_presentation_with_limits` calls
`capture_with_provenance` or `capture_with_parent_digests`
(`crates/litchi-pptx/src/package/model.rs:263-287`). Its capture validates the
complete slide reference set, captures every slide and its roots, checks
relationship/name uniqueness, loads notes, clones the package, and computes a
complete revision (`crates/litchi-pptx/src/opened/model.rs:730-878`).
`package_fingerprint_with_memo` sorts all Parts and feeds root relationships,
opaque non-Part members, each Part name/content type, every payload digest,
and every relationship (`:912-991`). A cold eager package has no parent
digest memo, so this exact revision legitimately hashes its payloads. A parent
memo reuses a digest only when the same payload allocation is proven; it does
not make a semantic materialization count into a cost estimate.

Slide capture and fingerprinting are therefore required validation for the
current generic eager transaction: they bind stale patch application and the
complete semantic graph plus opaque-member inventory to a complete revision.
Skipping them
because only eight slides are edited would weaken the snapshot contract. A
future digest-transfer change would need its own proof that source reads and
the owned package allocations are exactly the bytes covered by the memo. It is
not selected as the one next candidate here.

The separate `source-a` profile is only a route control: its owner total is
108,165,330 Ir, with 55,654,082 in publication, 38,232,988 in open, and
13,962,259 in text staging. Those counts are not cross-route native phase
fractions. They do show why a direct same-source comparison tweak is a poor
eager priority beside the Deflate and fingerprint owners above.

## Why eager `to_bytes` Deflates the package

The measured helper deliberately takes a different ownership route from a
source-preserving package:

```text
SourceBackedPackage::from_read_at(source)
    -> into_opc_package()
    -> Package::from_opc_package(opc)
```

The calls are at `tools/perf-baseline/src/lib.rs:51084-51088`. The consuming
conversion in `crates/litchi-opc/src/source_backed.rs:7874-7990` starts an
`OpcPackage::new()`, reads every source Part, and passes each decoded payload
to `PartFactory::load_shared`. It moves non-Part descriptors, but it does not
move the retained source archive, the physical member index, or verified
compressed spans into that new package. `PartFactory::load_shared` creates a
ready decoded payload; it does not create a `PartPayload::Transferred` token.

The following `Package::from_opc_package` call explicitly selects
`from_opc_package_with_provenance(opc, false)`
(`crates/litchi-pptx/src/package/codec.rs:259-302`). That boolean is not a
substitute for OPC preservation provenance, but it confirms that this eager
package cannot claim the physical source authorization used by source-bound
operations.

`Package::to_bytes` delegates to `PackageWriter::to_bytes`
(`crates/litchi-pptx/src/package/codec.rs:355-370`). The writer builds a
publication plan and first calls `try_write_preserved`
(`crates/litchi-opc/src/pkgwriter.rs:236-240,416-446`). That function can
only plan `PreservationAction::Copy` when `OpcPackage::preservation_source()`
exists and the source archive, provenance member list, current relationships,
and current Part payloads still agree. The eager `OpcPackage::new()` has none
of that, so the writer falls back to `materialize_pristine` and the physical
writer Deflates its Parts. The existing `PartPayload::Transferred` path and
`part_entry` compressed-transfer check
(`crates/litchi-opc/src/payload.rs:379-400` and
`crates/litchi-opc/src/pkgwriter.rs:1104-1126`) cannot help unless a verified
source token was retained.

This makes the large eager serialization owner avoidable only at the route
boundary: an owning package that has no exact source provenance must encode a
new archive. A generic compressor change, compression-level change, or
blindly marking the materialized package as source-backed would either leave
the dominant work in place or authorize stale/opaque bytes. The source
preserving writer already has the safer shape: it copies untouched entries
and regenerates only changed entries. The source-backed overlay writer uses
that shape at
`crates/litchi-opc/src/source_backed.rs:10505-10508,11936-11953`.

## Timer and diagnostic boundary

The helper toggled for this profile is
`publish_pptx_multi_slide_batch_edit`. Its timed body includes the initial
eager open/materialization, transaction staging and commit, two eager inverse
and forward patch applications at `tools/perf-baseline/src/lib.rs:51103-51110`,
and the final `to_bytes` plus sink writes at `:51115-51118`. The two replay
applications are harness-only qualification work. They are useful gates for
reversibility, but they are not production publication phases.

The `to_bytes` owner has one production source call at `:51115`; nested
Callgrind function-view call counts (for example, the `PackageWriter` edge)
must not be read as repeated serializations. The opened owner is harder to
assign: initial transaction capture and replay-owned patch capture share the
same global `capture_internal` and fingerprint functions. Thus the
994,945,242-Ir opened owner cannot be claimed entirely by the production
initial snapshot or entirely by replay.

The existing observer does not repair this attribution. OPC
`SourceCacheDiagnostics` reports content-free cache/flight counters and lock
boundaries; it carries no publication phase or member identity. The existing
PPTX opened-transaction phase vectors describe other selectors and are not
wired to this multi-slide helper, whose current operation metrics are absent.
The post-helper source summary likewise reports aggregate reads and
materializations, not whether a read or hash occurred during production
publication or a replay. The next profiling batch therefore needs explicit
phase boundaries around eager open/snapshot, transaction commit, production
publication, and each replay gate. It must retain the replays as separately
labelled qualification work and must not infer cost from the `229` count.

## Prior context that rules out smaller or unsafe detours

The accepted 0845 EntryId iterator work is a bounded immutable topology lookup
optimization. It does not address the 3.189-billion-Ir Deflate owner or the
complete revision fingerprint, so it is not a reason to broaden the current
eager change into more source-name loops. The 0845 priority review also records
that a naive fresh-compression change failed durable inverse identity. The
existing verified precompressed APIs (`authorize_precompressed` and
`AuthorizedPrecompressedPart`) are the relevant safe seam; they validate
source lineage, member layout, decoded-byte equality, CRC/sizes, limits,
signature/encryption/trailing-byte policy, and source freshness before a
compressed token can be published. They cannot be replaced by a caller-made
compressed buffer.

ADR 0030 keeps direct routing of format crates onto `SourceBackedPackage` as
unauthorized without a general source-backed save contract. ADR 0005 requires
preserve-mode raw copying where possible, finite resource checks, and real
profiles before an optimization claim. ADR 0006 makes unknown markup and
physical preservation a preserve-or-refuse boundary. ADR 0003 requires a
complete validated candidate and reversible source-bound publication. Those
constraints rule out both a generic codec shortcut and weakening the eager
snapshot checks.

## One bounded next step: an owned-source preservation proof

Before changing the eager serializer, qualify one existing owned-ingress
control using `Package::from_vec` (or the equivalent `from_reader` route),
with the same corpus and semantic multi-slide edit. Do not use borrowed
`from_bytes`: it intentionally enters the non-owned path. Do not route the
format facade directly through `SourceBackedPackage`; this control only asks
whether the existing owning `OpcPackage` preservation machinery can be the
reference for a future eager bridge.

The proof has one question: after the same multi-slide commit, does an owned
package that retains its source archive and preservation provenance reach the
existing `PackageWriter::try_write_preserved` plan, copy unchanged physical
members, and regenerate only the changed dependency closure? The control must
record copy/regenerate decisions and compressed bytes by member class, while
keeping the production publication and the inverse/forward replays in
separate labels. This is a route qualification, not a performance result.

Admission requires all of the following before any eager bridge is proposed:

* The archive remains the exact retained source: lineage/version, archive end,
  no trailing opaque bytes, source freshness, cancellation, read limits,
  reservations, and sink preflight all remain enforced.
* The changed set is complete. Slide XML, presentation relationships,
  content types, package relationships, notes, and any other dependency that
  changes must be regenerated; every unchanged known Part, relationship,
  ordering record, compression record, and unknown/non-Part member must remain
  raw-copy eligible or cause a typed refusal.
* Signed, encrypted, malformed, duplicate/case-folded, unsupported-topology,
  and source-mutated inputs retain their existing refusal identity. No raw
  compressed token may survive a payload, content-type, relationship, or
  source-lineage change.
* The output reopens under the retained limits and passes the existing slide,
  relationship, media, untouched-Part, unknown-member, route-specific output
  oracle, and durable inverse/partial-sink checks. The source archive remains
  immutable.

If this control passes, the only bounded production candidate justified by
this audit is to retain equivalent owned-source provenance through the eager
materialization boundary for this same-topology edit and reuse the existing
preservation writer. That candidate must fall back or refuse when the proof
does not hold; it does not authorize a general source-backed facade, a
compression-policy change, unsafe access, or skipped snapshot fingerprinting.
If the control cannot meet those gates, treat the full eager Deflate as the
required cost of the current provenance-free owning route and profile the
snapshot phase separately before proposing another change.

The source blobs reviewed at the frozen commit are identified by these hashes:

```text
tools/perf-baseline/src/lib.rs              f712da842169399d687eb452c7042deed61682a80efdd40b2c3c178af2c036ea
crates/litchi-pptx/src/package/codec.rs     bd76dc84510be99758949f65ebd569fe193db9a19d6fe733e2840ccd18ad8103
crates/litchi-pptx/src/package/model.rs     f81617c03d7ea2f6d579d9c4c66f4be6f795dd78af3e4259cdc8728198d38985
crates/litchi-pptx/src/opened/model.rs      9f36d2bb66516616be1f5a4850d7b2e8aeb9d885558bd61ffafd6d40d994cce0
crates/litchi-opc/src/source_backed.rs      0de2778325c86d81a6a4d5c209d94d0c9b3452489c8c777e9528c79ee87ce8ad
crates/litchi-opc/src/pkgwriter.rs          067eb6ef070ce00e1d5dc1033ad65c227a7931cbb45d71469e55e1f2f3df601f
crates/litchi-opc/src/package.rs            03e01110a41bc65171e75fd857dfd9785872aedce7451561b7ddac605a48e81a
```
