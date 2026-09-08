# 0479 DOCX plain-paragraph append contract audit

This audit fixes the workload definition for the first 0479 DOCX append
measurement. It describes the public source-backed
`litchi_docx::source_backed::paragraph_copy` path as it exists at commit
`45de59730a81dd465dff195226b2fc7bc47748c5`. It does not authorize a source-tail
publisher or change any production or harness code. iWork is outside this
audit.

The repository-level requirements remain those in [`docs/GOAL.md`](../../../GOAL.md).
The accepted ADR/README set is the 30-file manifest in
[`change-0478/adr-refresh.json`](../change-0478/adr-refresh.json). Its recorded
SHA-256 values were checked against the current tree during this audit; all 30
match, including the README. The two user-owned untracked documents were also
checked and left untouched (`docs/GOAL.md` SHA-256
`bed4058bb76330daab8ce9d4bceff639ab3fbd7ea06634158bef41b133c4d1f1` and
`docs/report/spec-gap-audit.md` SHA-256
`f71288dddc0827d5f5a16bf3936985606ec1b341d6e5c06fced4e612c7972ddf`).

## Contract finding

The existing public operation is one exact copy of one source paragraph. A
faithful tail-append baseline therefore has this shape:

```text
N source paragraphs
source = Position::new(0)       # or another fixed source index
before = Position::new(N)       # the tail insertion slot
one Edit::copy_plain_paragraph
commit
publish to a caller-owned sequential Write sink
```

`Edit::copy_plain_paragraph` documents that `source` is in `0..N`, `before` is
in `0..=N`, and `before == N` inserts immediately before `</w:body>`
(`crates/litchi-docx/src/source_backed/paragraph_copy.rs:318-324`). The
operation slot is explicitly single-use: after one successful copy, a second
call returns `Error::Limit { resource: "operations", max: 1, actual: 2 }`
(`paragraph_copy.rs:324-339`). Positions always refer to the immutable source
order, so the projected paragraph count does not create another valid slot for
the same edit.

Consequently, the primary baseline matrix should use one appended paragraph at
each source size (64, 8,192, and 131,072). The 131,072 case cannot use the
convenience `Package::edit_plain_paragraph_copy()`, because that constructor
captures `Limits::default()`, whose `max_paragraphs` is 65,536
(`paragraph_copy.rs:205-215` and `941-983`). It must instead call
`plain_paragraph_copy_snapshot_with_limits`, set a finite
`max_paragraphs >= 131,073` for the candidate, and set finite XML/event/output
ceilings large enough for the selected corpus. The candidate has one more
paragraph than the source, so a limit of exactly 131,072 would reject the
successful edit during the candidate scan.

An experiment labelled append-64 or append-256 is a different lifecycle, not a
larger request to this edit. It would have to reopen the previous emitted
artifact, capture a new snapshot, perform one copy, commit, and publish again
for each iteration. Publication consumes the package (`publish_*` takes
`self`), and a changed output is a new source artifact. Such a repeated
reopen/edit/commit/publish workload may be useful as a separate scenario, but
its total must be reported as repeated one-operation lifecycles. It must not be
presented as a 64- or 256-operation transaction, and it must not be combined
with the single-append baseline row.

The current parser has no supported final section property. Its accepted
topology is `document -> body -> p -> r -> t`; `next_scope` rejects all other
children (`paragraph_copy.rs:1554-1789`). In particular, a body-level
`w:sectPr` is a `ComplexDocument` refusal and a paragraph-property
`w:sectPr` is a `ComplexParagraph` refusal. The existing test includes both
forms (`crates/litchi-docx/tests/source_backed_paragraph_copy.rs:429-473`). A
current-path correctness oracle must therefore assert that the corpus has no
section element and that the copied paragraph is before `</w:body>`. A
“final `sectPr` placement” oracle belongs to a future, explicitly broadened
capability; adding it to this baseline would measure a different contract.

## ADR obligations

The relevant accepted decisions impose the following obligations on the
baseline and on any later bounded candidate:

| ADR | Obligation for this workload |
| --- | --- |
| [0001](../../../adr/0001-priorities-and-api-layers.md:18) | Correctness and safety outrank measured speed; this specialized path must remain explicit, typed, validated, panic-free, and must not expose raw archive types through ordinary DOCX signatures. Unsupported editing remains a typed refusal (`:23-52`). |
| [0002](../../../adr/0002-crate-topology.md:6) and [0024](../../../adr/0024-current-topology.md:15) | DOCX owns the semantic narrow capability, while `litchi-opc` owns physical OPC packaging; the current workspace has standalone `litchi-docx` and `litchi-opc` owners (`0024:17-42`). A performance seam cannot move ZIP grammar or physical IDs into the DOCX public contract. |
| [0003](../../../adr/0003-snapshots-edits-and-patches.md:6) | The source remains immutable; an edit is isolated; commit validates atomically and returns a new snapshot plus reversible patch without mutating the source (`:8-25`). Exact identity, lineage/version, and deterministic conflict behavior remain publication gates. |
| [0005](../../../adr/0005-io-memory-and-performance.md:6) | Use immutable positional `ReadAt` with stable source identity/version (`:8-12`), finite resource limits and explicit execution/scratch policy (`:19-28`), sequential sink semantics with incomplete-output progress (`:30-46`), and measured allocation/I/O/copy evidence rather than intuition (`:54-61`). Existing-document append is explicitly a restricted tail-only transaction, so this audit's `before == N` case is the append control. |
| [0006](../../../adr/0006-validation-security-and-compatibility.md:6) | Preserve untouched bytes, ordering, compression, unknown markup, namespaces, and lexical details where possible; validation does not mutate; normal save does not repair (`:8-23`). External links, VBA, signatures, protection, and other unsafe edits stay inert or refused (`:87-111`). |
| [0008](../../../adr/0008-migration-and-verification.md:1) | The baseline is evidence for a buildable migration slice, not a support claim. Any optimization must retain focused compile/test evidence and the required scenario gates; no broader DOCX append support can be advertised before its checklist and verification evidence pass (`:8-28`). |
| [0010](../../../adr/0010-facade-archive-ownership.md:14) and [0011](../../../adr/0011-ooxml-physical-package-ownership.md:19) | Format and OPC owners may use archive implementation details internally, but facade/format APIs must not expose concrete ZIP reader/writer types. `litchi-opc` is the sole OOXML physical package owner, and unchanged members must remain raw-copied through its boundary (`0011:21-37,44-52`). |

The audit therefore treats the current materialized path as the measurement
control. It does not infer that the existing implementation already meets the
eventual explicit-window memory target in `docs/GOAL.md`; that target requires
separate source-tail design review and evidence.

## Public lifecycle and phase boundaries

The operation has four externally meaningful phases. A benchmark may time them
separately, but the total operation must include all four and the retention
owned by the current implementation.

| Phase | Public/source path | Current work and retained data |
| --- | --- | --- |
| Open/index | `source_backed::Package::from_read_at` (`crates/litchi-docx/src/source_backed.rs:308-314`) -> `SourceBackedPackage::from_read_at` (`crates/litchi-opc/src/source_backed.rs:5364-5371`) | Captures source version and length, checks input limits, indexes ZIP metadata, content types, package/part relationships, and validates freshness. Ordinary part payloads stay deferred. The immutable positional source is retained in `SourceBackedPackage`; catalog vectors/maps and the ZIP index are retained for the package lifetime. |
| Snapshot/materialization | `plain_paragraph_copy_snapshot_with_limits` (`paragraph_copy.rs:941-978`) | Revalidates the narrow package topology, loads the main `word/document.xml` payload with `main.data()`, converts it to an owned `Arc<Vec<u8>>` on this unmanaged path, scans the complete XML, builds an `Arc<[Range]>` paragraph index, records `body_end`, and fingerprints the complete source artifact. It checks source version again before exposing the snapshot. |
| Edit/commit | `Snapshot::edit`, `Edit::copy_plain_paragraph`, `Edit::commit` (`paragraph_copy.rs:240-357`) | `Snapshot::edit` clones handles and initially shares the source XML/layout. The one copy computes `source.xml.len() + fragment.len()`, reserves one complete candidate XML vector, copies prefix + selected paragraph + suffix, scans the complete candidate again, and checks every projected paragraph with `validate_copy_readback` (`paragraph_copy.rs:1319-1369,1432-1458`). The commit retains projected bytes and a patch with shared `before` and `after` byte handles, the source fingerprint/version, limits, and the one operation. |
| Publish | `publish_plain_paragraph_copy_patch_to_stream` (`paragraph_copy.rs:1000-1050`) -> `SourceBackedPackage::write_part_overlay_to_stream` (`crates/litchi-opc/src/source_backed.rs:7678-7802`) | Reopens the current source closure by repeating topology validation, main payload materialization, full scan, and source-artifact fingerprinting. `Patch::apply` clones the full `after` XML, scans it, and checks the exact source identity (`paragraph_copy.rs:413-424`). The changed path clones the target XML for the replacement payload, reads the original selected Part for framing/CRC and no-op comparison, validates XML, builds a preservation index and plan, raw-copies all untouched ZIP members, regenerates only `word/document.xml`, and streams the fresh archive. A successful `Publication` retains the target/current snapshots, an O(1) source-artifact handle, the published fingerprint, and an inverse patch (`paragraph_copy.rs:603-628,1019-1049`). |

The exact no-op path is separate: when `Patch::is_noop()` is true, the
publisher copies the complete `SourceArtifact` byte-for-byte rather than
building a changed overlay (`paragraph_copy.rs:1028-1037`; OPC
`source_backed.rs:7789-7798`). The append baseline is a real change, so it
uses the overlay path.

The source-backed OPC owner keeps physical archive concerns below DOCX. Its
`SourceReader` forwards archive reads to the caller's `ReadAt` adapter
(`crates/litchi-opc/src/source_backed.rs:2623-2671`). Each read is bounded by
the selected buffer and, on managed paths, an input-byte reservation; the
source version is checked before and after the read
(`source_backed.rs:10943-11015`). The public paragraph-copy path uses the
unmanaged compatibility constructor, so the benchmark must still count
`ReadAt` calls/bytes with its source wrapper, while not labelling those logical
reads as physical disk I/O.

The main payload load is a deferred ZIP read/decompression. The later source
fingerprint reads the entire source artifact in 64 KiB chunks
(`source_backed.rs:2720-2758`, with
`SOURCE_PUBLICATION_CHUNK_BYTES = 64 * 1024` at line 45). Publication then
replays the physical ZIP preservation plan and writes the output sequentially.
The existing `OpcOperationAccounting` type exposes cold compressed/stored
payload reads, decoded bytes, generated payload bytes, raw unchanged archive
bytes, and accepted output bytes (`crates/litchi-opc/src/accounting.rs:1-122`),
but `publish_plain_paragraph_copy_patch_to_stream` currently calls the
non-accounting public overlay method. A 0479 harness must therefore either
instrument the `ReadAt`/sink boundaries and clearly report only those logical
counters, or add a separately reviewed accounting seam; it must not infer
decompression or physical-disk bytes from elapsed time or total output size.

## Allocation, copy, and retention inventory

The current operation is intentionally materialized. The significant byte
owners are:

1. The caller's `ReadAt` owner and the immutable source artifact remain live
   from open. With an `OwnedSource`, the input archive bytes are already owned
   by that source; with a file or remote/range source, the owner has a distinct
   storage contract. The benchmark must record source ownership at operation
   entry separately from allocations performed by the operation.
2. Opening retains the bounded ZIP index and OPC catalog. It does not retain
   ordinary decompressed Part payloads merely because the package was opened.
3. Snapshot capture owns the complete decompressed main XML in `Arc<Vec<u8>>`,
   the paragraph range slice, namespace/parser scratch during the scan, and an
   O(1) source-artifact handle while the fingerprint is computed.
4. The edit candidate owns a second complete main XML allocation. The source
   XML and candidate XML coexist in `Edit`/`Commit`; the range slices and the
   patch share their `Arc` handles rather than copying those byte vectors again.
5. Publication re-captures a current snapshot and `Patch::apply` clones the
   complete target XML. The replacement passed to OPC is another owned clone
   of that target XML. The preservation index/plan, ZIP scratch, compressor
   state, and sink-owned output are additional operation or caller storage.
   The publication result retains the snapshots, original artifact handle, and
   inverse material after successful output.

The edit's full-vector behavior is visible in `copy_fragment`: it reserves
`output_len` exactly, extends the prefix, selected fragment, and suffix, then
wraps that vector in a new snapshot (`paragraph_copy.rs:1344-1368`). The
candidate scan reserves the paragraph-range vector up to the configured bound
and grows it as needed (`paragraph_copy.rs:1563-1581`). The source and
candidate vectors, patch handles, source owner, and publication result must
not be collapsed into one peak number without naming which were already live,
which were transient, and which remain retained after the phase.

The operation's meaningful copy categories are therefore:

- compressed ZIP payload bytes read/decompressed for the selected main Part;
- source XML bytes copied into the complete candidate during edit staging;
- candidate bytes cloned for patch application and the OPC replacement;
- unchanged ZIP local/central spans copied by the preservation publisher;
- regenerated main-Part framing/payload bytes and accepted sink bytes.

`Arc::clone` of an existing immutable byte owner is a handle operation, not a
payload-byte copy. It still prolongs retention and must be reflected in live
owner accounting. Conversely, `checked_clone` in patch application and
publication is a real full-byte allocation (`paragraph_copy.rs:1844-1850`).

## Limits and parser/refusal contract

`Limits::new` requires all six limits to be nonzero and rejects a durable-patch
limit above 64 MiB (`paragraph_copy.rs:123-165`). The six dimensions are:

- main XML bytes;
- direct paragraph count;
- XML event count;
- XML nesting depth;
- projected main-document output bytes;
- durable patch bytes.

The default policy is 16 MiB XML, 65,536 paragraphs, 1,000,000 events, depth
16, 32 MiB projected output, and 64 MiB durable patch bytes
(`paragraph_copy.rs:205-215`). Every measurement policy must remain finite and
must be recorded with the result. The 131,072-source case requires an explicit
policy as described above.

The XML scanner (`paragraph_copy.rs:1563-1718`) accepts only:

- one `w:document` root and one `w:body` in either the Transitional or Strict
  Word namespace, with one consistent resolved Word namespace;
- direct body children that are plain `w:p` elements, each containing only
  `w:r` and `w:t` descendants;
- text events only inside `w:t`, plus the five predefined entity references;
- whitespace outside text, an optional declaration in the prolog, namespace
  declarations, and `xml:space="preserve"` or `"default"` on `w:t`.

Everything else is a typed refusal. The scope transition and attribute checks
reject section properties, tables, hyperlinks, bookmarks, fields, drawings,
wrappers, unknown children/namespaces, unsupported attributes, comments,
CDATA, processing instructions, DTDs, and non-predefined references
(`paragraph_copy.rs:1693-1831`). The scanner also refuses an empty document or
missing body closure. An empty `w:p` is accepted; empty body is accepted but
has no source paragraph to copy, so `source = 0` returns `OutOfBounds`.

Before scanning the main XML, `validate_plain_paragraph_copy_topology` requires
`/word/document.xml`, a normal non-macro main content type, matching package
and root dialect, and package-wide relationship/Part safety
(`paragraph_copy.rs:1136-1229`). It refuses external relationships, signature
infrastructure, VBA/macro content, comments/footnotes/endnotes/custom XML and
altChunk dependencies, and protected/write-protected or tracked-revision
settings. These are not optional corpus annotations: malformed or unsupported
cases must remain refusal gates outside the timed success path.

## Publication and failure obligations

The exact source checks are part of the operation contract. `Patch::apply`
requires complete artifact fingerprint equality, exact main XML equality, and
the captured `SourceVersion`; it returns `Error::StaleSource` on any mismatch
(`paragraph_copy.rs:413-424`). The publication path recaptures the closure
before output, monitors the source during preservation, and checks source
freshness at the beginning and during physical replay
(`crates/litchi-opc/src/source_backed.rs:9251-9295,9361-9467`). A source
revision discovered after accepted output takes precedence and is reported as
typed incomplete output with the accepted byte count
(`source_backed.rs:10680-10698`).

The sink is caller-owned and only requires `Write`. A short-writing sink is
handled by the ZIP writer; a zero-progress sink is an error. Once bytes have
been accepted, a non-atomic sink may contain a partial artifact, so the
benchmark must record accepted progress and must not call a failed output a
valid document. These behaviors are exercised by
`partial_sinks_complete_and_write_zero_fails_without_false_progress`
(`crates/litchi-docx/tests/source_backed_paragraph_copy.rs:623-720`). Source
revision refusal before output is covered by
`changed_source_version_is_rejected_before_publication_output`
(`source_backed_paragraph_copy.rs:572-621`).

Successful publication retains an inverse authorized for the exact emitted
artifact. The inverse checks the complete published-artifact fingerprint
before writing the retained original artifact; it does not restore a foreign
or subsequently changed package (`paragraph_copy.rs:1052-1067`). Exact no-op
publication writes the complete source unchanged, and the corresponding test
also checks a stale inverse writes no bytes
(`source_backed_paragraph_copy.rs:239-259`).

## Existing executable evidence

The focused test file already supplies the correctness and refusal seam for a
0479 baseline:

| Test | Contract evidence |
| --- | --- |
| `copies_exact_fragment_at_first_middle_and_last_source_order_slots` (`source_backed_paragraph_copy.rs:114-157`) | Exact source fragment duplication and insertion order at first, middle, and tail slots. |
| `publication_raw_copies_unselected_members_and_inverse_restores_exact_artifact` (`:159-192`) | Untouched ZIP payload bytes remain raw-identical; inverse restores the complete original artifact. |
| `durable_forward_inverse_are_canonical_and_reject_whole_artifact_staleness` (`:194-237`) | Canonical durable encoding, exact source application, inverse, and whole-artifact stale rejection. |
| `empty_edit_is_an_exact_noop_and_stale_inverse_writes_nothing` (`:239-259`) | Exact no-op bytes and stale inverse zero progress. |
| `positions_operation_count_and_resource_limits_are_failure_atomic` (`:261-328`) | Checked source/insertion positions, one-operation ceiling, XML/paragraph/output/durable limits, and unchanged projected state on refusal. |
| `relationships_dependencies_signatures_macros_and_paths_are_refused` (`:330-427`) | External, unsupported dependency, signature, macro, and noncanonical main-path refusal. |
| `structured_paragraphs_sections_and_hostile_namespaces_are_refused` (`:429-473`) | Direct-plain grammar and explicit section/complex-content refusal. |
| `strict_and_ordinary_producer_namespace_declarations_are_supported` (`:475-518`) | Strict and Transitional dialect acceptance and escaped text. |
| `empty_body_and_empty_or_sole_paragraph_slots_are_unambiguous` (`:520-545`) | Empty-body and empty-paragraph slot behavior. |
| `enforced_protection_is_refused` (`:547-570`) | Document protection and tracked-revision refusal. |
| `changed_source_version_is_rejected_before_publication_output` (`:601-621`) | Version change is rejected before output bytes. |
| `partial_sinks_complete_and_write_zero_fails_without_false_progress` (`:672-720`) | Short-write success, zero-progress failure, and partial sink failure. |

The broader source-backed file also verifies that opening/cataloguing leaves
ordinary payloads cold until a selected query and that changed source files
are rejected (`crates/litchi-docx/tests/source_backed_file.rs:139-309`). Those
tests support the open/materialization phase split but do not turn the
plain-paragraph copy path into a streaming append operation.

## Measurement contract for 0479

The first baseline should record, per fresh process and per successful
single-append case:

- source paragraph count, source index, tail insertion slot, explicit
  `Limits`, source archive/XML/output sizes and hashes;
- open/index, snapshot, edit/staging, commit, and publication timings, plus
  the total including destruction/retention boundaries chosen by the harness;
- allocator requested/allocated/deallocated/reallocated bytes, peak live
  bytes, live bytes at entry/exit, and retained source XML, candidate XML,
  patch, publication, catalog, and sink storage separately;
- `ReadAt` call count and requested/returned bytes by phase, logical ZIP
  compressed/stored/decompressed counters when an approved accounting seam is
  available, and sequential sink write calls/accepted sizes;
- inserted paragraph bytes, exact output hash, paragraph order/text,
  no-`sectPr`/tail placement, untouched-member digests, complete reopen/readback,
  and source/output artifact fingerprints;
- untimed mandatory refusal gates for stale source, XML/paragraph/output and
  durable limits, malformed/complex/section sources, short/zero/interrupted
  sinks, cancellation where a managed source is explicitly exercised, and
  output-limit behavior.

The operation must be reported as a materialized source-backed edit/publication
control. It is not evidence of constant memory, constant RSS, zero-copy
publication, or bounded-window append. The control is valuable precisely
because its candidate and publication retention are explicit. A future
bounded source-tail candidate may be compared with it only after a separate ADR
review proves source preservation, stale-source handling, failure atomicity,
limits, and the public publication contract; the future candidate must retain
the same correctness/refusal matrix.
