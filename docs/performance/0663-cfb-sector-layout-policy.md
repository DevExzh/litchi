# 0663: source-backed OLE2 saves keep their sector layout, with a bounded fallback

Status: **retained, implemented in `litchi-cfb`, `litchi-ole-common`,
`litchi-doc` and `litchi-ppt`.** `performance_claim: none` — the evidence
measures layout retention, validated same-length overlays and the cost of the
policy decision; it does not claim the 13–30% speedup priced by 0617.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred
until that goal completes; iWork is excluded.

This is the implementation of decision 10 in [0652](0652-owner-decisions-for-the-third-wave.md),
which [0651](0651-queue-refresh-after-the-second-wave.md) authorized from
[0617](0617-cfb-copy-through-writer-design.md). The writer now has a public
`SectorLayoutPolicy`, defaulting to `Reuse`, and a source-adoption entry point.
An adopted source is parsed through the ordinary bounded CFB parser. When its
sector geometry and directory shape still match, unchanged stream allocations
are kept, released sectors are reclaimed in deterministic order, and growth
uses those sectors before appending. A shape, geometry, changed nonzero
class-ID, DIFAT or planner gate declines the reuse plan and records a typed
fallback before the existing from-scratch writer runs; an explicit zero
class-ID is a supported clear.

The source-backed DOC route reaches this policy by default. `litchi-doc`'s
`RevisionEditor` and embedded-object transactions open an
`litchi-ole-common::object::Editor`, which retains the original CFB bytes and
calls its source-backed `finish`; the common editor defaults to `Reuse`. For a
same-length stream edit with unchanged topology and metadata, that route first
builds a bounded physical overlay over the captured source, reopens the
composed view through the ordinary CFB reader, and only then materializes the
output. Length changes and metadata edits use the source-layout writer instead.
The PowerPoint embedded-object finish path explicitly adopts its original
source before writing. The public `litchi-doc::Writer` constructs a new
document and has no source artifact to adopt, so it remains the from-scratch
route. This is the meaningful scope boundary: the policy is active on an
opened DOC/OLE edit, while a newly authored DOC cannot reuse a layout it never
received.

The layout decisions follow the local [MS-CFB specification](../../3rdparty/specs/[MS-CFB]/ToC.md): the header's cutoff, FAT count and DIFAT fields are
defined in [§2.2](../../3rdparty/specs/[MS-CFB]/2%20Structures/2.2%20Compound%20File%20Header.md),
MiniFAT and the root mini-stream chain in [§2.4](../../3rdparty/specs/[MS-CFB]/2%20Structures/2.4%20Compound%20File%20Mini%20FAT%20Sectors.md),
and directory entry start/size/CLSID preservation in [§2.6.1–§2.6.3](../../3rdparty/specs/[MS-CFB]/2%20Structures/2.6%20Compound%20File%20Directory%20Sectors.md).
The planner zero-fills unused stream-sector tails as recommended by [§2.7](../../3rdparty/specs/[MS-CFB]/2%20Structures/2.7%20Compound%20File%20User-Defined%20Data%20Sectors.md)
and declines DIFAT layouts covered by [§2.5](../../3rdparty/specs/[MS-CFB]/2%20Structures/2.5%20Compound%20File%20DIFAT%20Sectors.md) until a bounded witness is available.

For version 3 files, the high DWORD of a stream size follows the compatibility
rule in §2.6.1: the parser accepts an older nonzero reserved word, while the
writer canonicalizes it to zero. The common-editor overlay therefore declines
such a source so that the layout writer can normalize the directory before
publication.

The selected allocation policy is **B, reuse then append**. The planner puts
released sectors and source `FREESECT` sectors into one ordered pool, consumes
that pool in ascending physical order for growth, and appends only after it is
empty. Decision 10 explicitly allows reclaim where it is cheap; this keeps
the source geometry and avoids unnecessary file growth without promising
physical offsets as a public preservation contract. Policy C, compaction to a
new packed layout, remains the opt-in `Rewrite` control. The corpus and smoke
evidence records reclaimed, appended and free-sector counts alongside stream,
directory and reopen checks.

## Preservation and determinism

The reuse planner retains the source directory image and class identifiers,
then changes only the stream start/size fields, explicit class-ID clears and
root mini-stream fields it owns. A requested nonzero class-ID change declines
to the from-scratch writer. It handles both sides of the 4096-byte
MiniFAT cutoff. The corpus test
reopens every output, compares every stream byte with the rewritten model,
checks the reused directory entries against the source including CLSIDs, and
runs the CFB validation walk. The 214-file OLE2 corpus yielded 211 examined
fixtures and 3 parser skips; 209 of 211 admitted no-op, same-length and
length-changing reuse, and 195 length-changing cases appended sectors. The
same-length and growth cases also exercise the exact source-backed route's
logical model, while the smoke corpus covers shrink/reclaim and both
mini-to-regular and regular-to-mini migrations.

The planner uses ordered maps and sets for allocation and directory matching.
The corpus repeats the same serialization in one process and launches the
same integration-test binary three times with a fresh process hash seed; the
layout digest is identical in all three children. A mutation sweep over a DOC
fixture declined 37 tampered sources and admitted 5 mutations only when the
ordinary parser still proved a valid layout; admitted outputs retained all
streams, while declined outputs matched the from-scratch bytes.

## What the measurement means

The representative release measurement uses `picture.doc` (1,448,448 bytes),
24 ABBA samples, and a same-window A/A floor. One window reports a reuse p50
of 350,291 ns and rewrite p50 of 248,411 ns, a +41.01% difference, with paired
floors of 1.74% and 2.15%. Two repeat windows report reuse p50s of 344,796 ns
and 348,376 ns against rewrite p50s of 246,556 ns and 249,481 ns, or +39.93%
and +39.64%; their paired floors are 1.80%/2.66% and 2.64%/2.21%.
The observed Reuse difference exceeds the corresponding Reuse floor in all
three windows. It is an engineering regression signal for this
`picture.doc` layout benchmark, so no latency claim is registered.

`kept_sectors` is a physical placement-retention count. For length-changing
or metadata edits, the source-anchored emitter still writes each output
sector from the current payload model. Equal-length stream edits through the
common editor additionally use validated source-backed overlay spans: the
unchanged source sectors are read through the composed view instead of being
re-laid out by `OleWriter`, and the changed spans are supplied by the edited
streams. The editor still materializes the final `Vec<u8>` after validation,
so this is an avoided re-layout/materialization path rather than a claim that
a generic sink receives no output bytes. On `picture.doc`, the no-op and
same-length outputs retain 2,795 of 2,828 body sectors and keep the source
length; the length-changing case keeps 2,795, reclaims one, and appends three.
The 0617 13–30% estimate remains withheld because this record does not isolate
that end-to-end cost.

A separate attribution probe compares the opened-DOC route from the clean
pre-0663 checkout with the final Reuse route on `FloatingPictures.doc`: the
median process times are 2.4811 ms before and 2.5383 ms after. Without a
paired A/A floor this is recorded as a baseline check, not a DOC latency claim.
The same current probe reports 2.9724 ms for the container-only Reuse leg and
3.9763 ms for a Rewrite control; that policy control is distinct from the
whole-save timing above and does not establish a speedup.

## Breaking changes

`litchi-cfb` adds the public `SectorLayoutPolicy`, `SectorLayoutFallback` and
`SectorLayoutReport` types, `OleWriter::adopt_source_layout`, policy accessors,
and owned/shared stream creation methods. `litchi-ole-common::object::Editor`
adds the policy setter/getter. The crates are 0.0.x and 0652 explicitly accepts
this breaking API movement. No unsafe code, dependency, archive type, raw
lock or executor is introduced.

The retained evidence packet is [`results/change-0663/README.md`](results/change-0663/README.md).
