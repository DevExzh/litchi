# Change 0483 DOCX source checkpoint

Status: bounded source checkpoint for the existing public source-backed DOCX
tail append. This document records what the current implementation admits and
what the 0483 comparison may measure. It makes no performance claim. Root owns
the final gate decision and any later measurement interpretation.

The governing context is [GOAL.md](../../../GOAL.md),
[docs/adr/README.md](../../../adr/README.md), and the accepted ADR set. The
30 accepted ADR hashes were revalidated unchanged for this checkpoint; the
compact obligation matrix is in
[adr-matrix.md](adr-matrix.md). This is a DOCX source route only. iWork remains
outside the scope and user documents are preserved by the existing package
ownership rules.

## Why this source checkpoint is needed

The older public plain_paragraph_copy operation opens a materialized main
document, owns complete XML, builds a paragraph range vector, constructs a
candidate, and then publishes through the ordinary patch path. The 0479 audit
records that contract in
[contract-audit.md](../change-0479/contract-audit.md), while the 0480
publication change and 0481 profile keep the complete source XML and paragraph
index in that route. The 0481 follow-up identifies the missing replay substrate
and makes clear that a local allocation reduction is not an explicit-window
tail API. See [0481 source review](../change-0481/source-review.md),
[0481 window contract](../change-0481/window-contract.md), and the 0482
[OPC/XML audit boundary](../change-0482/xml-audit-review.md).

The 0483 materialized-versus-bounded harness therefore has a narrower purpose:
it supplies the same UTF-8 first-paragraph text to the existing materialized
copy route and to the source-backed route, and checks the two outputs through
independent package, XML, semantic, raw-member, and inverse oracles. Its
source, candidate, paragraph-vector, and output observations motivate this
checkpoint; they do not authorize a speed, allocation, RSS, or constant-memory
claim. The harness contract and timing boundary are recorded in
[harness-contract.md](harness-contract.md).

## Public closure admitted by the current source route

The public operation is
source_backed::Package::tail_append_plain_paragraph, with the explicit
tail_append_noop route separately available. The implementation and public
proof types are in
[tail_append.rs](../../../../crates/litchi-docx/src/source_backed/tail_append.rs#L1)
and the module re-export is in
[source_backed.rs](../../../../crates/litchi-docx/src/source_backed.rs#L1).
The changed operation admits exactly one authored plain paragraph per
Edit/Plan/Commit lifecycle:

* The target is the ordinary main /word/document.xml part reached through
  the single supported main-document relationship. The package and settings
  dialects must agree. Signed, encrypted, macro-enabled, externally related,
  unsupported-dependent, protected, tracked-revision, and otherwise complex
  topologies are typed refusals. The topology and settings checks run before
  output publication.
* The main story is a direct document/body sequence of direct plain p
  paragraphs containing the admitted r/t text grammar. Tables, wrappers,
  bookmarks, fields, drawings, comments, processing instructions, CDATA,
  unknown direct body or paragraph content, and malformed XML are refusal
  boundaries. Authored text is XML-escaped, and invalid XML 1.0 controls,
  malformed names, and unsupported declarations are refused.
* A direct final WordprocessingML w:sectPr is admitted as one opaque lexical
  span. It must be the unique direct final section-properties element; its
  bytes, length, and hash are retained as scalar facts and replayed unchanged.
  The new paragraph is inserted immediately before that span, preserving the
  source whitespace around the anchor. If there is no direct final sectPr,
  insertion is immediately before the body close. Nested, non-final, duplicate,
  wrapped, cross-dialect Word sectPr, or unsupported events inside the opaque
  span are refused. A same-local-name foreign vendor element can remain opaque
  when the namespace policy admits it.
* The generated paragraph is namespace-bound to the source context. Empty
  authored text still produces one real empty paragraph. xml:space and local
  namespace behavior are preserved according to the source scanner's policy.
  The fragment is bounded by max_fragment_bytes; text length and owned text
  capacity are bounded by max_text_bytes.
* Every configured limit is finite and checked before the corresponding owner
  is retained: source XML, text, fragment, candidate XML, parser events and
  depth, paragraph count, settings XML, workspace, output, and token bytes.
  The defaults and validation rules are defined by Limits in
  [tail_append.rs](../../../../crates/litchi-docx/src/source_backed/tail_append.rs#L152).

This is the exact one-paragraph closure. It does not admit a caller-supplied
stream of paragraphs, a paragraph-window selector, arbitrary body CRUD, or a
final-section rewrite. The current API has no batch argument that makes
append-1/64/256 one edit. A repeated append is a sequence of distinct
lifecycles: publish one candidate, establish that candidate as the next source
package, then open or otherwise provide the current package to a new
tail_append_plain_paragraph call. Each call rescans and revalidates its own
source and candidate and receives its own limits, leases, proof, publication,
and inverse authorization. A retained snapshot does not turn that sequence
into an implicit source-tail provider or a durable forward patch. This is why
0481's replayable full-stream and explicit-window objectives remain open; the
0483 route must not be described as satisfying them.

## Read, scan, copy, and publication path

The source-backed package is opened from an explicit positional ReadAt
provider. Catalog and part metadata are validated lazily; the bounded tail
route does not call the ordinary document() materialization path, which owns
complete main XML and a paragraph index. The package source, source version,
execution context, and cache ownership remain authoritative under
[source_backed.rs](../../../../crates/litchi-docx/src/source_backed.rs#L320).

Edit::prepare follows this source-derived phase path (the implementation's
phase order is in
[tail_append.rs](../../../../crates/litchi-docx/src/source_backed/tail_append.rs#L551)):

1. Validate finite limits and source identity/version, admit caller text length
   and owned capacity, and reserve owned text storage when applicable. Validate
   topology, relationships and settings. A stale source is rejected before a
   plan is returned.
2. Scan the decoded source with a guarded BufRead and match its dialect to the
   package. This pass computes the source scalar proof, direct grammar facts,
   event/depth/paragraph counters,
   insertion offset, and opaque final-section span hash without retaining the
   complete source XML.
3. Encode one bounded paragraph fragment using the proved source dialect.
   Release the temporary text owner and its reservation after fragment
   ownership is established.
4. Scan a TailSpliceReader over source prefix, bounded fragment, and source
   suffix. This candidate readback authenticates the exact candidate length,
   generated paragraph count and offset, namespace/grammar facts, and unchanged
   opaque section span before any sink receives output.
5. Hand the bounded fragment and scalar proof to OPC. OPC prepares its source
   splice plan, preservation/replay state, and physical-member checks while
   retaining no complete decoded source or candidate member.
6. Commit produces a publication object. Forward publication replays the
   source archive through a sequential sink, splicing the bounded fragment at
   the proved decoded offset and preserving unselected physical members. The
   publication writer checks the package execution context and cancellation
   before writes and flushes. Sink failure reports incomplete output according
   to the existing OPC contract.

Thus the route has separate source scan, fragment encoding, candidate scan,
OPC preparation, and publication replay work. Source and candidate readback
are proof phases, while publication is a fresh source replay owned by OPC; an
independent reopened readback or oracle is outside the operation's publication
contract. Logical ReadAt behavior, short reads, and one-time interruption are
tested at the source-provider boundary. These phases explain which reads and
copies a measurement must attribute; they do not establish a fixed number of
physical reads for every provider or ZIP implementation.

## Execution authority and memory leases

The package-owned ExecutionContext is the authority for semantic scanning,
fragment storage, settings admission, and OPC publication. Edit::with_options
can add its cooperative cancellation token. Options exposes no execution-context
setter; effective_options obtains that authority from the package, so an edit
cannot bypass the package's hierarchical budget. The relevant authority and guards
are [tail_append.rs](../../../../crates/litchi-docx/src/source_backed/tail_append.rs#L701)
and [source_backed.rs](../../../../crates/litchi-docx/src/source_backed.rs#L405).

Managed work is prepaid and checked with typed Resource::Memory and
Resource::Work leases. The source route charges, in the phase where each owner
exists:

* caller-owned text capacity and the bounded encoded fragment;
* guarded XML token windows, opened-name/index vectors, namespace resolver
  state, attribute duplicate-check scratch, scope stacks, counters, and
  section hashing;
* settings source/cache bytes, MCE output and scratch, post-MCE scanner state,
  the complete settings model, relationship/model work, and retained processed
  output; and
* OPC XML audit state, preservation/replay/compressor state, output writer
  state, and decoded replay work.

[mce_workspace.rs](../../../../crates/litchi-docx/src/source_backed/tail_append/mce_workspace.rs#L1)
and
[settings_workspace.rs](../../../../crates/litchi-docx/src/source_backed/tail_append/settings_workspace.rs#L1)
are checked requested-storage envelopes. They are not process-RSS or allocator
metadata measurements. Settings MCE output and scratch leases are reserved
before the synchronous codec call; scratch is dropped when that call returns.
The processed output remains charged while the post-MCE guard and complete
settings model run. The model lease is admitted after post-MCE facts are
recounted, and the full parser receives the original relationships map without
collecting an unbounded matching vector. Boundary work and cancellation checks
surround codec calls, so a cancellation occurring inside a synchronous codec
is observed at the next documented phase boundary.

## Source and candidate scalar proofs

SourceProof records the source version, decoded source length and SHA-256,
proved insertion offset, direct paragraph/event/depth facts, strict namespace
state, and final opaque sectPr length/hash. CandidateProof records the
candidate length/hash, candidate paragraph/event/depth facts, generated offset
and once-only insertion fact, and the unchanged section length/hash. Their
public definitions are in
[tail_append.rs](../../../../crates/litchi-docx/src/source_backed/tail_append.rs#L283).

The OPC handoff uses SourcePartSpliceProof, whose raw scalar fields bind the
decoded source, bounded fragment, and candidate lengths/hashes and decoded
insertion offsets. OPC rechecks the source artifact and current execution
context before changed publication. Candidate validation requires, among other
conditions, candidate length equal to source plus fragment, exactly one added
plain paragraph, one generated insertion at the shifted offset, unchanged
strictness and opaque section facts, and nondecreasing event/depth facts. The
scanner's candidate checks are in
[validate_candidate](../../../../crates/litchi-docx/src/source_backed/tail_append.rs#L3490);
OPC's scalar owner is documented in
[splice.rs](../../../../crates/litchi-opc/src/source_backed/splice.rs#L187).

These proofs diagnose and bind the borrowed current plan. They are not a
rehydratable patch format. Persisting a proof for later use would require an
explicit DOCX source/artifact binding and a replayable provider.

## Settings admission is the full model path

The settings phase is part of the changed-document safety contract. The source
path locates the settings relationship, checks the declared part and dialect,
then guards the bounded settings bytes. It processes one MCE pass with the
DOCX settings capability set, retains only a bounded owned output when the
codec changes the bytes, recounts the processed XML, and runs the complete
borrowed settings, extension, and mail-merge model with the original
relationships. The implementation sequence and leases are in
[tail_append.rs](../../../../crates/litchi-docx/src/source_backed/tail_append.rs#L1237);
the source-derived owner review is in
[settings-review.md](settings-review.md) and the MCE audit is in
[mce-memory-review.md](mce-memory-review.md).

The full model preserves the existing defaults and refusal behavior for
documentProtection, writeProtection, and trackRevisions; duplicate and
malformed scalar settings, dialect mismatches, attached-template relationship
validation, unknown extensions, and unsupported MustUnderstand input remain
fail-closed. A settings part is bounded and charged under its explicit settings
limit; this does not imply package-wide constant memory or a bounded arbitrary
settings/document producer. The settings helper's concrete owner formulas and
phase-boundary obligations are recorded in
[settings-review.md](settings-review.md), while the scanner envelope and
quick-xml owners are recorded in [memory-review.md](memory-review.md).

## No-op, publication, and inverse semantics

tail_append_noop is an explicit lifecycle and must not be inferred from empty
authored text. It hashes and verifies the current main source, builds equal
source/candidate scalar proofs with an empty fragment, and delegates to the
OPC exact no-op plan. The generic OPC no-op policy can copy a signed or
malformed source byte-for-byte because it is reproducing authenticated source
bytes; the changed DOCX route still refuses that topology. No-op publication
therefore has exact source bytes and an immediate exact inverse under the
existing artifact authorization rules.

For a changed append, publication preserves untouched ZIP members and their raw
records/compressed payloads except for required archive offsets, while the
main member receives the bounded splice. The publication exposes a candidate
artifact fingerprint. The inverse first authenticates that complete candidate
archive and current candidate context, then copies the retained original
source artifact byte-for-byte. Foreign, stale, or tampered candidates fail
authentication before restoration starts. Cancellation or sink failure during
replay can leave partial bytes in a sequential destination and returns the
OPC output-progress error; the caller owns atomic destination replacement.
The inverse uses the current candidate package's execution context; it does not
carry the forward edit token into a later operation. OPC source-splice tests
cover raw-member retention, exact no-op copying, proof rejection, candidate
freshness, and inverse cancellation in
[source_part_splice.rs](../../../../crates/litchi-opc/tests/source_part_splice.rs#L582).

## Existing correctness and production-gate evidence

The focused DOCX integration matrix is in
[source_backed_tail_append.rs](../../../../crates/litchi-docx/tests/source_backed_tail_append.rs#L848)
and summarized in [tests-review.md](tests-review.md). It covers:

* final opaque-section insertion for Store and Deflate archives, strict and
  transitional namespaces, unusual prefixes/attributes/children, and the
  no-section body-close path;
* XML escaping, xml:space, empty authored paragraphs, malformed names and
  declarations, unsupported grammar, nested/non-final/duplicate section
  properties, namespace refusals, and pre-output refusal sinks;
* signed/encrypted/external/protected/unsupported topology refusal, settings
  protection and relationship closure, MCE fallback and MustUnderstand,
  settings output/workspace ceilings, and duplicate/cardinality cases;
* source immutability and stale-source checks, explicit no-op and immediate
  inverse, managed memory/work leases, cancellation before prepare and during
  publication, short sinks, source short/Interrupted reads, and bounded main
  XML cache diagnostics; and
* OPC raw opaque-member preservation and source/candidate proof rechecks.

The accepted production gates below all passed against source manifest
`30d3ad60b0ae90911992b9855c5328913de2741f27090ba6c31b760c2ebcaeaf`:

The repository ignores the root workspace lock. Its observed post-validation
copy is retained as [workspace-Cargo.lock.txt](workspace-Cargo.lock.txt), with
scope and checksum in [workspace-lock.json](workspace-lock.json), so the locked
workspace commands can be reproduced. The benchmark uses its separate tracked
lock; both fuzz targets bind their retained fuzz-workspace lock.

| Receipt | Scope recorded by the gate |
| --- | --- |
| [all features](validation/docx-full-all-features-accepted.json) | litchi-docx release tests with all features; the bundle README records the source-stable 1,378-test run with 31 ignored documentation examples. |
| [no default features](validation/docx-full-no-default-accepted.json) | litchi-docx release tests with default features disabled; the coordinator ledger records the 1,359-test no-default run. |
| [Clippy](validation/docx-clippy-accepted.json) | litchi-docx all-targets release Clippy with warnings denied. |
| [all-feature Clippy](validation/docx-all-features-clippy-accepted.json) | all-features, all-targets release Clippy with warnings denied. |
| [rustdoc](validation/docx-docs-accepted.json) | all-features release rustdoc with warnings denied and dependencies excluded. |
| [formatting](validation/workspace-format-accepted.json) | workspace formatting gate. |
| [crate boundaries](validation/crate-boundaries-accepted.json) | crate-boundary policy check. |
| [non-iWork workspace](validation/non-iwork-workspace-check-accepted.json) | all-features, all-targets workspace check excluding the iWork crates. |

These receipts establish correctness and source custody for the admitted
operation. The [benchmark tests](validation/harness-tests-accepted.json)
passed all five focused cases; its [Clippy gate](validation/harness-clippy-accepted.json)
also checked test targets with warnings denied. The
[public example](validation/docx-example-build-accepted.json) built, and
[LibreOffice readback](consumer/final/result.json) verified both synthetic
plain-body and final-section documents after the append. The consumer probe
uses an explicit UTF-8 export filter. Its earlier locale-dependent export
failure remains in [final-47](validation/docx-consumer-final-47.json).
This certifies that synthetic open/TXT-export scenario, not Microsoft Office
compatibility or broader producer coverage.

Both the new [tail-append sanitizer target](validation/docx-fuzz-smoke-accepted.json)
and the existing [DOCX parser target](validation/parse-docx-fuzz-smoke-accepted.json)
completed 10,000 libFuzzer iterations under AddressSanitizer, using seed 483
and the 62-seed corpus. Their build, input, binary, mutated-corpus and crash
inventories are retained under `fuzz/accepted/`. These bounded smoke runs do
not prove exhaustive malformed-input coverage.

The prior [example-target failure](validation/example-targets-final-40.json)
contains only iWork collisions. The [scoped replacement](validation/non-iwork-examples-accepted.json)
passes while continuing to reject non-iWork and mixed collisions. The earlier
full-workspace iWork example type error is recorded separately in
[workspace-iwork-exclusion.json](workspace-iwork-exclusion.json). Neither
failure is represented as a passing full-workspace result.

Formal measurements followed this source checkpoint. The accepted
[720-sample comparison](measurements.md) and [process profiles](profile-review.md)
record lower large-source operation heap, roughly doubled normal latency and
similar whole-process RSS. This checkpoint describes the source contract;
those separate artifacts are the performance evidence.

## Open boundary and nonclaims

The source checkpoint closes the admitted one-paragraph operation only. It
leaves the following work open:

* a replayable authored paragraph stream or explicit source window API for
  larger tails, including a durable source provider and rehydratable patch
  semantics;
* repeated-append scaling as a separate sequence of source-backed lifecycles,
  batch append, and authored-size memory scaling beyond the measured
  three-source-size, one-paragraph comparison;
* general paragraph CRUD, non-plain story content, broader producer diversity,
  arbitrary section rewriting, native Office save/reopen breadth, and
  cold/warm/concurrency evidence; and
* phase attribution and removal of repeated CPU work without weakening the
  source, semantic and publication proof obligations.

The source checkpoint and the later performance closure are separate commits.
The complete non-iWork objective remains open in [next-work.md](next-work.md).
