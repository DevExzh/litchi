# 0477 checked generated-name capability design

This is a proposed next-batch design for the fresh PPTX streaming route. It
is not an implementation or a new performance result. The coordinator has
selected the central-directory spool as the first implementation batch; the
end-to-end constant-memory result remains open. The parent branch's ADR hash
refresh is unchanged, so this note applies the accepted explicit-scratch,
typed-failure, deterministic-output, and public-semantics constraints without
claiming a fresh ADR reread.

## Boundary and source facts

The current PPTX writer already has a fixed semantic state. Its parent keeps
the declared slide count, next-slide counter, options, limits, and aggregate
text bytes; an active slide keeps one part writer and scalar counters
([`streaming.rs:217-238`](../../../../crates/litchi-pptx/src/writer/streaming.rs)).
The module documents the remaining transport gap: central-directory and
member-name metadata still grow with the number of parts
([`streaming.rs:1-11`](../../../../crates/litchi-pptx/src/writer/streaming.rs)).

The identifiers do not require a retained slide table. The slide ID is
computed as `256 + next_slide` ([`streaming.rs:316-324`](../../../../crates/litchi-pptx/src/writer/streaming.rs));
the presentation manifest emits that formula directly
([`streaming.rs:1117-1158`](../../../../crates/litchi-pptx/src/writer/streaming.rs));
and presentation relationships use `rId(4 + index)` and
`slides/slide(index + 1).xml`
([`streaming.rs:1161-1195`](../../../../crates/litchi-pptx/src/writer/streaming.rs)).
The slide relationship part uses the fixed `rId1` to layout 1. The constructor
knows the complete slide count before touching the output and writes the
fixed members before the first slide
([`streaming.rs:267-288`](../../../../crates/litchi-pptx/src/writer/streaming.rs)).
That is the useful invariant for a generated-name capability: PPTX already
has a finite topology and a forward sequence.

The retained owners below that layer are different:

* `StreamingArchiveWriter.names` is a `HashSet<String>` used by
  `validate_entry_name` and `record_streaming_entry` to reject normalized ZIP
  duplicates ([`office.rs:5326-5337`](../../../../crates/soapberry-zip/src/office.rs),
  [`office.rs:5876-5925`](../../../../crates/soapberry-zip/src/office.rs),
  [`office.rs:6018-6023`](../../../../crates/soapberry-zip/src/office.rs)).
* The low-level ZIP writer's `files` and `file_names` retain central-record
  fields and names until finalization ([`writer.rs:185-190`](../../../../crates/soapberry-zip/src/writer.rs)).
  The selected central spool removes those vectors in spool mode, but it does
  not answer duplicate-name queries.
* `PartNameSet.names` retains folded full OPC names and
  `PartNameSet.descendants` retains folded ancestor-to-descendant relations
  ([`phys_pkg.rs:946-955`](../../../../crates/litchi-opc/src/phys_pkg.rs)).
  `prepare` also constructs a folded candidate and one descendant string for
  each ancestor ([`phys_pkg.rs:964-983`](../../../../crates/litchi-opc/src/phys_pkg.rs)).
  `validate` enforces exact duplicates, ASCII-equivalent duplicates, and
  ancestor/descendant conflicts ([`phys_pkg.rs:986-1013`](../../../../crates/litchi-opc/src/phys_pkg.rs)).

Central-record spooling and generated names solve separate problems. A
central record contains CRC, sizes, flags, local offset, ZIP64 state, and the
name learned after a member is emitted. A replayable caller-owned store is
therefore still required for a non-seekable output. A generated plan can
remove the two name indexes only after it has proved the complete namespace;
it cannot reconstruct central records or make arbitrary names safe.

The sealed 0476 default evidence separates these owners from compressor
allocation work: the large lane's requested allocation work fell from
6,809,604,013 to 31,428,173 bytes (99.538473%), while incremental peak heap
changed from 8,875,092 to 8,875,252 bytes (+160 bytes)
([`change-0476/README.md:43-50`](../change-0476/README.md)). That record makes
no constant-memory claim; the retained ZIP/OPC indexes and central metadata
remain a separate design obligation.

## Proposed capability shape

The low-level proof language belongs in the ZIP layer (or a small dependency
below it), not in OPC. OPC should consume a checked generic sequence of
canonical member names and enforce token order; it must not know that a name
is a PPTX slide. PPTX owns the constants, relationship formulas, layout
count, and its 15-fixed/two-per-layout/two-per-slide topology. The current
ZIP32 topology bound of 65,534 entries and the resulting 32,748-slide bound
remain PPTX validation facts ([`streaming.rs:41-50`](../../../../crates/litchi-pptx/src/writer/streaming.rs)).

The next API should expose an explicit, checked capability rather than a
boolean switch on ordinary writers. Names below are API-shape names and can
be adjusted during implementation:

```text
GeneratedNamePlan::builder()
    .literal(canonical_member_name)
    .indexed_family(prefix, first, count, suffix)
    .repeat_pair(family_a, family_b)
    .finish() -> Result<GeneratedNamePlan, NamePlanError>

StreamingArchiveWriter::with_generated_names(
    output, limits, central_spool, plan
) -> Result<GeneratedArchiveWriter, Error>

GeneratedArchiveWriter::start_next_entry(method)
    -> Result<GeneratedEntry<'_>, Error>

PhysPkgWriter::with_generated_names(
    output, metadata_spool, plan
) -> Result<GeneratedPhysPkgWriter, OpcError>

GeneratedPhysPkgWriter::start_next_part()
    -> Result<GeneratedPartWriter<'_>, OpcError>
```

The actual public spelling may instead use a `GeneratedArchiveWriter` and
`GeneratedPartWriter` capability returned by the ordinary constructors. The
important properties are that the plan is supplied before publication, the
generated writers have no method accepting an arbitrary name, and the
ordinary `start_entry(name, ...)` and `start_part(&PackURI)` paths remain
unchanged. The capability is not an unchecked constructor and does not expose
a caller-created proof token.

### Proof representation

`GeneratedNamePlan` should contain only a fixed number of descriptors, not one
descriptor per emitted member. A useful initial representation is deliberately
narrow enough to make the proof mechanical:

* `Literal`: one canonical member path, with UTF-8 bytes and component
  boundaries validated by the same ZIP path normalizer used by ordinary
  entries;
* `IndexedFamily`: a literal prefix, an inclusive `first` value, a checked
  `count`, and a literal suffix. Initially the numeric slot is one decimal
  run in the final path component; all parent components are fixed, the
  decimal encoding has no leading zeros (except zero itself), and the suffix
  begins with a non-digit. The descriptor rejects overflow before
  publication. This covers `slide{n}.xml` and `slide{n}.xml.rels` while
  avoiding a general string-equation solver;
* `SequenceStep`: a fixed literal, one family range, or a fixed pair of
  family ranges. A pair step is useful for `slideN.xml` followed by
  `slideN.xml.rels` without storing N tokens; and
* a checked family/descriptor count and a plan identity held privately by the
  generated writer.

The representation is intentionally narrow. It should initially accept only
the path spellings needed by the PPTX route and reject any family whose
canonical or OPC-folded form cannot be proved. An implementation may store
literal bytes in an arena-like plan allocation, but that allocation is
proportional to the fixed topology and name-pattern lengths, not to the
number of slides. The generated cursor stores only a step index, a family
index, the next numeric value, and an in-flight flag.

`GeneratedNamePlan::validate` runs before the output sink is touched and uses
the production normalizers. It must prove:

1. Every literal and family expansion is a valid canonical ZIP member name.
2. Decimal encoding is injective over each finite range and all arithmetic is
   checked.
3. Every pair of literal/family descriptors is disjoint under ZIP duplicate
   normalization and OPC ASCII-case folding.
4. No generated name is an ancestor or descendant of another generated name.
   A path-prefix proof compares complete component boundaries, so `a/b` is
   related to `a/b/c` but not to `a/bc`.
5. Each sequence step has a finite count and the sum agrees with the expected
   member count and any caller-provided output limits.

The first proof engine can then use a finite relation table. For two indexed
families, first compare their fixed parent components and the bytes before
the numeric slot under ZIP normalization and OPC folding. If those fixed
parts differ at a position, the families are disjoint. If they are equal,
compare the fixed suffixes: equal suffixes are disjoint exactly when the
checked numeric intervals do not intersect; different non-digit suffixes are
disjoint because decimal digits cannot consume them. Any prefix relation
between static parts, a differing slot position, or an ambiguous folded
spelling is `UnprovableNamePlan` rather than an optimistic proof. For a
literal and a family, compare parent components and test whether the literal's
final component has exactly the family's prefix, one canonical decimal run in
the checked range, and the family's suffix; otherwise the pair is disjoint.
For ancestor checks, compare complete component prefixes. A literal that is a
proper component prefix of a family is a conflict; equal-depth names and
different final components are not. The restricted family form means two
families cannot become ancestors merely by varying the final filename. Never
expand a large range. An `UnprovableNamePlan` error is safer than silently
treating an unproved plan as trusted.

With F family descriptors, L literals, and S fixed sequence steps, plan
validation is O((F + L)^2) descriptor work plus path-pattern length. The
cursor is O(1) per active writer. For PPTX, F, L, and S are fixed by the
resource topology, so this is O(1) with respect to slide/member count.

The PPTX plan can then be constructed as follows, without putting PPTX
grammar in OPC:

```text
fixed PPTX members (15 literals)
fixed layout members (2 literals per built-in layout)
slide XML family:       /ppt/slides/slide{1..N}.xml
slide relationship family:
                         /ppt/slides/_rels/slide{1..N}.xml.rels
sequence: fixed members, then for each index the two slide families
```

The exact order must match the current writer's byte output and tests. The
plan builder receives the already checked `slide_count` and layout set from
PPTX; OPC sees only the canonical generic names. The plan must include every
static member currently emitted by `with_options`, so a missing static name is
a plan-construction error rather than a late ZIP failure.

### Cursor and state machine

The generated writer owns the validated plan and exposes only
`start_next_*`. Each call performs this sequence:

1. Check that the writer is not poisoned and that the cursor has a remaining
   step.
2. Derive the next canonical name into one bounded working buffer and verify
   its expected ordinal and canonical form against the plan. The writer keeps
   a scalar ordinal, not a set of previous names.
3. Mark one opaque, non-`Clone` generated lease as in flight and pass that
   lease to the ZIP/OPC entry writer. The lease contains no public arbitrary
   name input; its private plan identity and ordinal prevent mixing plans.
4. On successful entry/part finalization, advance the cursor exactly once.
   A repeated, skipped, reordered, or foreign lease returns a typed refusal.
5. If payload output has started and finalization fails, poison the writer as
   the existing forward-only APIs do; no rollback is promised for a
   non-seekable sink. If validation fails before the first output write, leave
   the sink untouched.

The name buffer may be reused between entries. It must not be retained in the
plan or copied into a growing index. The central spool still receives the
serialized central record when the entry closes. At finalization the cursor
must be exhausted; otherwise the generated writer returns a typed incomplete
plan error and does not emit a misleading complete archive.

An implementation can make the lease entirely private to the generated
writer and call a low-level internal `start_checked_entry` method. If a public
cross-layer token is needed, use an opaque token tied to a private plan
identity and expected ordinal. Do not expose `unsafe`, `trusted`, or
`skip_validation` parameters, and do not let callers manufacture a token from
an arbitrary path.

## What remains exact for ordinary callers

The existing ordinary APIs must continue to perform their current checks and
retain their current semantics:

* `StreamingArchiveWriter::start_entry(name, method)` keeps normalized-name
  duplicate detection and entry-limit accounting.
* `PhysPkgWriter::start_part(&PackURI)` keeps exact duplicate,
  ASCII-equivalent duplicate, and ancestor/descendant checks through
  `PartNameSet`.
* ZIP and OPC errors retain the existing typed allocation, limit, invalid-name,
  and incomplete-output behavior. Central spooling is an explicit storage
  capability and does not silently convert an ordinary writer into a
  filesystem-backed writer.

Arbitrary names cannot be replaced by a generated proof. A caller may submit
two equal names in different spellings, a later descendant of an earlier
part, or a name whose conflict is not known until the later call. Exact
validation therefore needs the current in-memory indexes, a separately
specified external exact name-index capability with lookup/insert/failure
semantics, or a preflight pass that knows the complete namespace. A central
directory spool alone supplies none of those queries. The external index is a
separate future design and must not be smuggled in as a mode that weakens
validation.

## Proven design versus unresolved trade-offs

The source-backed facts are established: PPTX IDs and relationship IDs are
formula-generated scalars; the PPTX topology is known up front; ZIP central
vectors and ZIP/OPC name indexes are separate retained owners; and the
current OPC validator checks both duplicate equivalence and path topology.
The proposed proof is also structurally sufficient: finite descriptors can
be checked before output, and a scalar cursor can enforce one deterministic
sequence without retaining prior names.

The following choices remain implementation work rather than proven API
decisions:

* whether the proof types live directly in `soapberry-zip` or in a tiny
  neutral lower-level crate; they must remain below OPC and contain no PPTX
  grammar;
* whether canonical generated names are represented as ZIP member paths
  without a leading slash, OPC `/`-prefixed URIs, or a pair of views over one
  canonical byte sequence. The implementation must avoid retaining two
  independently owned copies and must run both production normalizers;
* how broad the initial family language is. A narrow accepted language with
  explicit `UnprovableNamePlan` errors is preferable to a symbolic solver
  whose corner cases are not tested;
* the exact lease lifetime on a payload write error, including whether the
  failed lease poisons immediately or is consumed by `finish`;
* whether plan identity is represented by a private capability object or by
  keeping all low-level operations inside the generated writer; and
* the final public spelling and provider bounds for the already-selected
  central spool. Its replay, short-read/write, limit, cancellation, and
  cleanup behavior must stay explicit.

These trade-offs must be resolved with semantic, determinism, invalid-plan,
partial-output, and failure-injection tests before the capability is used by
PPTX. A post-change allocation capture must distinguish plan construction,
central-spool reservation, compressor working memory, and retained operation
heap; a constant-memory claim is not justified by the proof shape alone.

## Next implementation sequence

The high-impact sequence for the next batches is:

1. Complete and review the selected explicit central-directory spool in
   `soapberry-zip`. Keep the existing in-memory path, add replayable bounded
   scratch, preserve central order/ZIP64/data-descriptor bytes, and test short
   scratch reads/writes, limits, failures, deterministic output, and
   sequential sinks.
2. Add the generic `GeneratedNamePlan` descriptors and symbolic validator in
   the ZIP layer. Test canonicalization, ASCII folding, decimal ranges,
   duplicate families, component-prefix relations, ambiguous/unprovable
   patterns, overflow, and pre-publication sink immutability. This stage must
   not alter ordinary arbitrary-name APIs.
3. Add the generated ZIP cursor and opaque lease state machine. Pair it with
   central spooling, keep only scalar cursor/metadata state, and test skipped,
   repeated, reordered, foreign, and incomplete sequences plus post-output
   poisoning.
4. Add the OPC generated writer as an explicit capability. It should pass the
   generic checked token/name sequence downward and omit `PartNameSet` only in
   this mode; its ordinary `start_part` path remains unchanged. Test exact
   OPC conflict behavior through the ordinary API and invalid-plan refusal in
   generated mode.
5. Have PPTX build the fixed topology plan from its existing checked slide
   count and resource set, then route fresh streaming creation through the
   generated OPC/ZIP capability and central spool. Verify exact member order,
   deterministic bytes, full reopen/semantic round-trip, and the existing
   slide/relationship ID formulas.
6. Run failure, security, resource-limit, and scaling evidence at tiny,
   medium, and large slide counts. Measure complete operation lifetime and
   report any remaining storage or sink costs separately. Only after that
   decide whether arbitrary-name external indexing warrants a separate API
   batch.

The coordinator's selected first batch is step 1. Steps 2 through 6 are
therefore specified work, not a claim that the fresh PPTX end-to-end route is
complete in this change.
