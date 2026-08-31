# Change 0342: XLSX source-backed hyperlink overlay

Status: implemented

`performance_claim: none`

## Scope

This change adds a deliberately narrow source-backed hyperlink overlay for
one existing XLSX worksheet. The selected worksheet is identified by the
source-backed package snapshot; the overlay does not create a worksheet, move
it, or search across the workbook. Its relationship topology is frozen for the
life of the operation.

The overlay exposes typed hyperlink observations and reversible edits for
hyperlinks authored in that worksheet's worksheet XML:

- internal hyperlinks may be added, removed, or replaced;
- an existing external hyperlink may have metadata-only changes while its
  target and relationship ID remain byte-for-byte the same;
- adding, removing, retargeting, or otherwise changing any external
  relationship is refused;
- external, internal, file, and other link targets are inert data. The overlay
  never follows, resolves, downloads, executes, or interprets a target.

The shared overlay is localized to the selected worksheet XML. Its patch keeps
that worksheet's `.rels` part unchanged and leaves every unselected worksheet,
relationship part, package member, content type, extension, and unknown XML
fragment untouched. A hyperlink is not silently migrated to a different
relationship or represented by a guessed relationship when its authored
relationship cannot be proven.

This is an overlay over an existing source-backed package, not an eager
workbook model. Snapshotting and listing hyperlink metadata do not inflate
unrelated package members. An edit materializes only the bounded worksheet
part needed to produce its patch; it does not aggregate all workbook sheets or
all package payloads.

Edit planning uses in-place bounded incremental staging. A selected worksheet
may stage at most 65,536 hyperlink entries and 256 mutations; exceeding either
bound returns a typed limit refusal instead of growing an unbounded rewrite
buffer. Baseline/current identity and relationship-count indexes are built by
linear, fallible scans over those bounded records, so allocation and limit
failures remain explicit and deterministic.

## Snapshot, patch, inverse, and freshness contract

The snapshot records the selected worksheet identity, its source generation,
the worksheet XML span/metadata required for the hyperlink overlay, and the
relationship state used to validate every external `r:id`. It records enough
authored state to distinguish an internal target from an existing external
target without relying on a later package lookup.

Each accepted edit produces a typed patch limited to the selected worksheet
XML. The patch includes an inverse patch over the same source spans, so apply,
undo, and redo can be checked against the snapshot rather than reconstructed
from a normalized workbook. Applying an already-effective edit is a no-op;
the no-op does not rewrite the source and preserves exact original bytes.

The source generation and the relevant worksheet/relationship fingerprints
are checked before planning and again before application. A changed source,
stale snapshot, changed relationship ID, altered external target, ambiguous
relationship, or missing worksheet returns a typed stale/refusal error. The
overlay never applies a patch to a different physical member merely because
its canonical name appears equivalent. A failed application has no successful
publication result and does not expose a partially rewritten package.

Remove/restore and composed edits use the same baseline/current identity and
relationship-count checks. In particular, moving an external hyperlink target
or composing a moved-external edit with a set change is refused even when the
resulting XML would otherwise be well-formed; only metadata changes retaining
the original target and `r:id` are eligible for an existing external link.

## Signed, protected, and lossless behavior

- A signed package with a true no-op edit is copied exactly, including
  signature parts and all unselected members. No signature metadata is
  rewritten merely to report that nothing changed.
- Any effective hyperlink edit on a signed package is refused before source
  mutation. The overlay does not silently strip, regenerate, or invalidate a
  signature.
- Effective edits on a protected worksheet or protected workbook are refused
  unless an explicitly owned, validated capability authorizes that exact
  operation. Protection is not bypassed by editing raw XML.
- Markup Compatibility and Extensibility content is preserved. If an affected
  hyperlink is inside unsupported `mc:AlternateContent`, extension markup, or
  another construct whose authored semantics cannot be preserved, the edit is
  refused rather than selecting one branch or dropping the other.
- Encrypted packages remain behind the existing encryption/decryption
  boundary. The overlay neither decrypts by itself nor emits plaintext as a
  shortcut; if the current package state does not provide the required
  validated writable capability, the operation is refused.
- Unsupported, malformed, duplicate, or ambiguous hyperlink metadata is
  surfaced as a typed refusal or retained as inert source content according
  to the owning parser policy. It is never silently normalized or discarded.

External relationship metadata is read only for the selected worksheet's
existing relationship IDs. The target URI and `r:id` are immutable in this
change, and `.rels` bytes remain unchanged even when permitted anchor metadata
changes. External relationship creation, deletion, replacement, retargeting,
or relationship-part normalization is outside the contract and always
refused.

## Limits, cancellation, and safety boundaries

Worksheet and relationship scans use the configured source-backed read limits
for package members, XML materialization, relationship records, names, and
allocation. Every bounded scan, snapshot operation, patch plan, and patch
application checks cancellation and source freshness. A limit, cancellation,
I/O, XML, relationship, signature, encryption, protection, or stale-source
failure is returned as a typed error.

Cancellation is also checked periodically through the rewrite scan and output
serialization, so a long but bounded worksheet cannot defer cancellation until
the final package write. Structural Markup Compatibility constructs and
unknown direct children at an insertion point are refused when preserving the
authored XML would not be provable; the overlay never drops them to make an
insertion succeed.

The overlay does not read or decode hyperlink targets, embedded objects,
macros, controls, formulas, or linked files. It does not follow external
targets or open arbitrary filesystem/network resources. No public API exposes
the internal package handle, relationship graph lock, runtime task, or
unbounded buffer. A caller can request only the selected worksheet's bounded
snapshot and the resulting localized patch.

## Validation status

The focused XLSX source-backed hyperlink suite passed 6 tests with 0 failures.
It covered internal add/remove/replace, same-target/same-`r:id` external
metadata edits, refusal of external relationship topology changes, inverse
patches, remove/restore, moved-external/set composition refusal, stale
snapshots, exact signed no-ops, effective signed refusals, protected-sheet
refusal, inert targets, configured limits and cancellation, structural MCE and
unknown-direct-child insertion refusal, `.rels` byte preservation, and
preservation of unselected members. The exact serialized, disk-backed command
was:

```sh
CARGO_TARGET_DIR=/home/zhuhe/CodeProjects/.cargo-targets/litchi-0342-target \
CARGO_BUILD_JOBS=1 CARGO_INCREMENTAL=0 CARGO_PROFILE_TEST_DEBUG=0 \
cargo test -p litchi-xlsx --test source_backed_hyperlinks -- \
--test-threads=1
```

The result was `6 passed, 0 failed`. This is correctness and source-locality
evidence only; it is not a benchmark or a process-resource measurement. The
target was removed after validation with:

```sh
find /home/zhuhe/CodeProjects/.cargo-targets/litchi-0342-target \
-xdev -depth -delete
```

Encrypted-state refusal/preservation paths remain represented by the typed
refusal boundaries and are deferred for dedicated fixture coverage.

## Locality and performance claims

`performance_claim: none`

Locality is a correctness property for this change, not a benchmark result:
only the selected worksheet XML may be patched; its `.rels` and all
unselected members must remain unchanged. The implementation must not claim
latency, throughput, allocation, RSS, decompression, or I/O improvements.
Any future measurement requires a bounded fixture and a separate recorded
comparison.

## Deferrals

The following remain deferred:

- overlays spanning multiple worksheets or workbook-level hyperlink stores;
- adding, removing, retargeting, or changing external relationships;
- changing an existing external target URI or relationship ID;
- hyperlink editing in drawings, charts, comments, headers, footers, tables,
  threaded comments, or other parts outside the selected worksheet XML;
- recursive target inspection, external-resource access, embedded-object
  decoding, formula/link evaluation, and executable or macro behavior;
- edits requiring relationship-part rewriting, worksheet creation/movement,
  package-wide normalization, encryption-key handling, protection bypass, or
  signature regeneration.

Any future expansion must retain exact source preservation, explicit inverse
patches, freshness fencing, inert-target behavior, configured limits,
cancellation, and typed refusal semantics.
