# 0477 retained ownership audit: fresh PPTX streaming

This audit covers the public fresh PPTX streaming writer through the OPC and
ZIP layers. It is a source audit for the next implementation batch; it does
not claim that the end-to-end streaming memory requirement is complete. The
coordinator's unchanged ADR hash refresh is the authority for the accepted
record set. The relevant contract applied here is that a sequential
non-seekable output may use explicit scratch storage, while scratch must be an
explicit caller-selected capability with typed limits and failures.

## What the PPTX layer retains

The semantic writer already has a constant-size authoring state. The parent
retains `slide_count`, `next_slide`, options, limits, and aggregate text bytes;
the active slide retains one part writer and scalar counters
([`streaming.rs:217-238`](../../../../crates/litchi-pptx/src/writer/streaming.rs)).
The module itself documents the remaining gap: its semantic window is bounded,
but the ZIP transport retains central-directory and member-name metadata that
grows with part count (`streaming.rs:1-11`).

Slide IDs and relationship IDs are not retained in a slide vector. The checked
slide ID is `256 + next_slide` (`streaming.rs:316-324`), and the presentation
manifest writes each `p:sldId` from that formula (`streaming.rs:1117-1158`).
Presentation relationships use `rId(4 + index)` and `slides/slide(index + 1).xml`
(`streaming.rs:1161-1195`). Content-type overrides and relationship records are
written directly to their members in indexed loops (`streaming.rs:999-1047`),
so the loop's `format!` values are temporary. The slide's own relationship
part is constant (`rId1` to layout 1), and shape IDs are scalar counters
(`streaming.rs:407-474`).

The constructor requires the total slide count before output and emits the
content-types member, package relationships, static parts, presentation
manifest, and presentation relationships before the first slide
(`streaming.rs:267-288`). It therefore avoids a late manifest patch or an
unknown relationship allocator. The topology is 15 fixed members, two members
per layout for the fixed layout set, and two members per slide; the checked
ZIP32 entry bound is 32,748 slides (`streaming.rs:41-50`, `:680-761`).

## Growing owners below PPTX

| Owner | Current state | Why it grows | Requirement for a constant-memory route |
|---|---|---|---|
| `StreamingArchiveWriter.names` | `HashSet<String>` of every normalized ZIP member name (`soapberry-zip/src/office.rs:5326-5337`) | `validate_entry_name` must reject normalized duplicates; `record_streaming_entry` inserts after each member (`office.rs:5876-5925`, `:6018-6023`) | Generic arbitrary-name APIs need an explicit bounded index or external duplicate index. The PPTX generated route can use a checked sequential name plan.
| `ZipArchiveWriter.files` | `Vec<FileHeader>` (`soapberry-zip/src/writer.rs:185-190`) | ZIP central records are emitted only at finalization; each member's offset, sizes, CRC, flags, ZIP64 state, and central extras are retained (`writer.rs:1456-1485`, `:2310-2323`) | Append each finalized central record to explicit replayable scratch and retain only scalar count/size/ZIP64 state.
| `ZipArchiveWriter.file_names` | Concatenated name bytes (`writer.rs:185-190`, `:846-858`) | Central records need their names at finalization | Write name bytes into the same central-record spool; do not keep a second concatenated in-memory name store.
| `PhysPkgWriter.part_names.names` | `HashMap<String, PackURI>` of folded full names (`litchi-opc/src/phys_pkg.rs:946-955`) | OPC rejects exact and ASCII-equivalent duplicate parts (`phys_pkg.rs:986-1013`) | Generic arbitrary OPC names require this index or an external bounded index. PPTX can prove names from a finite generated topology and consume opaque next-name tokens.
| `PhysPkgWriter.part_names.descendants` | `HashMap<String, String>` of folded ancestor to one descendant (`phys_pkg.rs:946-955`) | OPC rejects ancestor/descendant conflicts; `prepare` clones a folded descendant for every ancestor (`phys_pkg.rs:964-983`) | A generated plan must validate the entire fixed topology and path-prefix relation before output, then enforce the expected sequence. Arbitrary names need the existing index or external storage.
| `PackURI` and prepared names | Owned URI plus folded full and ancestor/descendant strings (`packuri.rs:16-20`; `phys_pkg.rs:958-983`, `:1430-1433`) | The same physical name is represented repeatedly across active and retained indexes | Generated tokens should own no independently retained name index; central spool records remain the publication representation.

`StreamingArchiveLimits.max_metadata_bytes` charges normalized names and
generated ZIP64 extra fields, and bounds entry count, but it is a cumulative
retained-metadata ceiling, not a fixed-memory window. `PhysPkgWriter` has no
corresponding name-index byte budget. Thus a directory spool alone cannot
close the requirement while the ZIP and OPC indexes remain in memory.

The default 0476 evidence is consistent with this ownership map. The sealed
large PPTX allocator lane reduced requested allocation work from
6,809,604,013 to 31,428,173 bytes (99.538473%), while incremental peak heap
changed only from 8,875,092 to 8,875,252 bytes (`change-0476/README.md:43-50`).
The retained directory/name structures are therefore a distinct peak-memory
owner from repeated Deflate allocation work. The 0476 record explicitly makes
no constant-memory claim (`README.md:29-34`, `:53-58`).

## Checked generated-name capability

A retention-free PPTX name path is safe only as a checked capability with a
closed input domain. It must be constructed from the already validated
`slide_count` and fixed resource topology, before the sink is touched, and
must prove all of the following:

1. Every fixed and indexed member name is canonical under the ZIP normalizer
   and is distinct under OPC's ASCII case-equivalence rule.
2. No generated member is an ancestor or descendant of another member.
3. The slide ID, relationship ID, member-count, and name-length arithmetic is
   checked before publication.
4. Each subsequent request consumes the next opaque plan token. The token
   yields the one expected canonical name; it does not accept an arbitrary
   caller string.
5. The sequence includes every static member, slide XML member, and slide
   relationship member exactly once, so a skipped, repeated, or reordered
   token is a typed refusal.

This capability can remove the PPTX route's `StreamingArchiveWriter.names`
and `PhysPkgWriter::PartNameSet` retention without weakening duplicate or
derived-name validation. It must remain private to the PPTX streaming route,
or be exposed only as a narrow low-level capability that cannot be combined
with arbitrary `&str` names. The existing generic low-level APIs must retain
their indexes for callers that can supply arbitrary names, case variants,
ancestor paths, or producer duplicates.

The distinction matters: a `skip_name_validation` flag would turn malformed
or conflicting arbitrary input into silent publication and would violate the
OPC preservation/safety contract. A checked finite plan is a proof of one
specific generated namespace and sequence; it is not a general duplicate
index replacement.

## Central-directory spool requirement

The ZIP writer must gain an explicit central-directory replay capability for
the sequential sink route. When a member is finalized, it already knows the
CRC, compressed and uncompressed sizes, local-header offset, flags, ZIP64
fields, and name. It can serialize the complete central record in emission
order into caller-provided scratch. The in-memory writer then retains only:

- finalized entry count;
- central-directory byte count;
- current output offset;
- whether any member or offset requires ZIP64; and
- bounded working buffers for one record and compressor state.

At `finish`, the writer replays scratch to the output, computes the final
central-directory offset/size, emits ZIP64 structures when required, and emits
EOCD. The spool interface must make append, replay, maximum bytes, short
reads/writes, cancellation, and cleanup explicit. Memory scratch may preserve
the current small/in-memory behavior; large streaming creation must be able to
select a caller-owned or encrypted temporary store through an explicit
capability. No ambient filesystem or hidden temporary path is acceptable.

Central-record spooling is required even with generated names: CRCs, sizes,
and local offsets are learned only after each payload is emitted. Generated
names remove the name-index tables, but they cannot reconstruct those
per-member descriptor fields from a non-seekable output later.

## Next implementation sequence

The next batch should complete the central spool substrate first, as selected
by the coordinator:

1. Add an internal/low-level replayable central-record store to
   `soapberry-zip`. Preserve the existing in-memory writer as one store, and
   add an explicit bounded scratch-backed store. Serialize records at entry
   finalization; remove `files` and `file_names` retention in the scratch mode.
2. Track central count, size, and ZIP64 requirements incrementally. Preserve
   central-record ordering, data-descriptor behavior, ZIP64 extras, sink
   partial-failure progress, deterministic output, and typed allocation/limit
   failures. Add exact small archives, ZIP64, short-scratch, scratch-failure,
   and sequential-sink tests.
3. Thread the explicit store through `PhysPkgWriter` without leaking ZIP
   implementation types through ordinary PPTX CRUD signatures. Keep the
   default generic/arbitrary-name route's current duplicate and derived-name
   validation.
4. Add the checked finite generated-name capability and use it only from the
   fresh PPTX writer. Validate the complete 15-fixed/layout/slide topology and
   sequence before output; then bypass only the already-proven retained name
   indexes, never the topology proof.
5. Add PPTX scaling tests at several slide counts that verify deterministic
   bytes, full reopen/semantic round-trip, duplicate/sequence refusal through
   the generic route, generated-plan refusal before sink mutation, scratch
   limits, and partial output errors.
6. Re-measure complete operation allocation lifetime separately from scratch
   reservation, requested allocation work, and process RSS. Report whether
   incremental peak is independent of slide/member count; retain any residual
   central-spool or output-sink costs as separate scopes.

This sequence addresses fresh streaming creation. Logical append to an
existing structure, adding a new package Part, and arbitrary modification
followed by repackaging remain separate scenarios and remain open.
