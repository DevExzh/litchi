# 0728 probe review

Status: independent review of the probe contract.  The probe source was still
incomplete when this review began, so this file records the qualification gates
that must hold before a native timing result is treated as a current baseline.
It is not a timing result and it does not authorize a production change.

The measurement question is narrower than “how long does an edited file take to
save?”  The public DOC/PPT operation and the common OLE editor must use the same
source bytes and the exact replacement stream bytes produced by that operation.
The public operation owns format parsing, semantic staging, and publication. The
common editor prices the container path for those already-produced bytes. The
two paths are related controls, not nested calls, so medians from separate
processes must never be subtracted to manufacture a phase fraction.

## Required timing boundaries

The source file read, argument parsing, fixture selection, target discovery,
replacement-byte construction, stream inventory, semantic oracle, hashing,
JSON formatting, and cleanup belong outside timed regions. The input bytes must
be loaded once and then supplied to every fresh operation; a second read or a
separately serialized “equivalent” source changes the question.

For each public format case, record three separately labelled intervals:

* `open`: the first call into the public owner with the already-owned source
  bytes through the owner’s complete bounded validation and snapshot creation;
* `stage`: creation of the edit/transaction and the one selected semantic
  operation, including target resolution and replacement encoding;
* `finish`: commit/publication, candidate validation, CFB emission, reopen and
  semantic readback that the public owner performs before returning its target.

The interval names must match the actual calls.  In particular,
`body_text::Edit::commit` and `slide_order::Transaction::commit` are publication
operations, not staging-only calls: they render or append the candidate and
reopen/validate it.  A probe that calls those methods inside `stage` and then
calls another `finish` is double-counting publication and must label that
window accordingly.  If a public API cannot expose a true staging-only
boundary, report the available lifecycle boundary instead of inventing one.

For the common `litchi-ole-common::object::Editor`, use fresh editors and the
same source/replacements for each policy and record:

* `common_open`: `Editor::open` through CFB parse, capture and target
  discovery;
* `common_stage`: only the operation that installs the prepared replacement
  streams, with its real validation behavior documented;
* `common_finish`: `Editor::finish` and the final output materialization.

Current `put_stream_shared`/`put_streams_shared` are not inert setters.  They
clone a candidate, render it, reopen it and recapture it before returning.  A
probe that uses one of them must either put that work in `common_stage` or
rename the phase to `replace_and_validate`; it cannot call it “staging” while
claiming that `common_finish` contains the render/recapture cost.  The same
warning applies to `put_stream_shared_with_rendered`.

Run `Reuse` and `Rewrite` as explicit, fresh controls.  Keep policy in every
row, binary identity, and output record.  The public source-backed route and
the common editor’s default policy must not be silently compared with a
Rewrite-only run.  Policy controls may produce different physical CFB bytes,
but their logical stream and semantic oracles must agree.

## Changed-length qualification

“Length-changing” must be proved at the logical stream level.  A physical CFB
file can retain the same byte length after a stream grows if the selected
layout has free sectors, so `source_file_len != output_file_len` is neither
necessary nor sufficient.

For every retained case, record at least:

* source and output package byte lengths;
* the complete source/output stream path sets;
* every changed stream path, source byte length, output byte length, and exact
  source/output bytes or a separately retained digest plus byte comparison;
* the selected semantic old/new value and its encoded length (DOC UTF-16
  units; PPT live slide count/order and persisted-record change);
* the expected changed stream set obtained from the actual public output;
* the source and output SHA-256 values as identity aids, never as the oracle.

The DOC witness must select a nonempty ordinary main-story paragraph and a
replacement whose UTF-16 length differs and whose text differs.  After reopen,
the target paragraph must contain exactly the requested replacement.  The
changed stream length must differ in the direction implied by the encoded
edit.  Do not hard-code `WordDocument`/`1Table` as the only changed streams
without checking the actual fixture and output; the replacement set must be
derived from the real public result.

The PPT witness must remove one existing live slide.  A changed `PowerPoint
Document` stream is expected to contain append-only history, so the removed
slide’s old bytes may remain unreachable.  The proof is that the live slide
directory has exactly one fewer entry and the selected slide identity is absent
from that live order, not that a byte search finds no copy of the old record.

## Complete CFB preservation oracle

The inventory must be a union comparison, not an iteration over output paths
only.  The latter misses a removed source stream.  Recursively enumerate every
stream and storage using the full path and exact stored spelling, including
empty streams, control-character names, nested `ObjectPool` members, and
unusual producer metadata.  Fail on any added or removed path unless the edit
contract explicitly permits it.

For every path, reopen and compare complete stream bytes.  All untouched stream
bytes must equal the source byte-for-byte.  Changed stream bytes must equal the
bytes from the qualified public format output when they are fed to the common
editor; do not regenerate them in a second model.  The common editor output
need not be package-byte-identical to the public output because its layout
policy may serialize physical CFB metadata differently, but its stream set,
stream bytes, and semantic result must satisfy the same expected model.

Stream names and lengths alone are a false positive: a same-length payload
mutation, an omitted empty stream, or a replacement under a different nested
path can pass that check.  A single whole-package hash is also insufficient;
it does not identify a missing path, an accidental semantic change masked by a
different physical layout, or which stream was changed.  Hashes are useful as
captured identity fields after the path/byte checks.

Directory metadata needs a separate check.  At minimum compare, by logical
path, entry kind, exact name, CLSID, logical size and MiniFAT classification.
For a Reuse output, use the 0663 corpus precedent: compare a normalized raw
directory image while blanking only planner-owned allocation fields (stream
start/size and the version-3 reserved high size word where applicable).  That
image keeps names, sibling/child links, node colours, CLSIDs, state bits and
both timestamps.  The public `DirectoryEntry` view does not expose all of
those fields, so a probe that checks only its `name`, `entry_type` and `clsid`
has not proved metadata preservation.  A bounded raw-directory helper in the
probe is appropriate; changing production CFB APIs for this measurement is
not.

For Rewrite, do not require source SIDs or source sibling-tree shape.  Require
the logical hierarchy and all non-owned metadata, and compare the same edited
model under both policies.  For Reuse, classify any metadata difference as
either an explicitly planner-owned field or a qualification failure.

The source identity fence is essential.  Record the hash and length of the
exact source passed to the public owner and the exact source passed to the
common editor, and require byte equality before timing.  Replacement bytes
must be captured from the public commit over that source.  A second fixture with
the same length, or a source that happens to produce the same replacement
lengths, must be rejected by a negative source-swap control.

## DOC semantic oracle

Before the timer, capture the source `body_text::Snapshot` projection and the
selected target.  After the output is reopened through the public DOC reader:

1. `Projection::All` paragraph sequence must equal the source sequence with
   exactly the selected paragraph replaced by the requested text; paragraph
   count and all non-target text must remain unchanged.
2. Where the fixture exposes them, compare `Accepted` and `Rejected`
   projections, every story returned by `story_paragraphs`, simple table-cell
   text, field-result text, revision inventory and author order.  A main-story
   paragraph edit must not silently rewrite another story or tracked range.
3. Reopen through the ordinary package reader and require the expected DOC
   generation/FIB and complete CFB validation to succeed.
4. Compare the exact stream/path and directory oracles above, including
   embedded/object-pool streams.  Semantic text readback by itself does not
   prove lossless preservation.

The target and expected sequence must be captured before editing.  Re-querying
“the first nonempty paragraph” after publication can accidentally select a
different paragraph and turn a wrong edit into a passing oracle.

## PPT slide-removal semantic oracle

Capture the source live slide directory before editing.  For each entry retain
at least `(persist_id, slide_id, flags, text_placeholder_count, list_text,
outline references/interactions)` and, where practical, each surviving slide’s
semantic text and raw live persisted record.  After reopen through
`litchi_ppt::Package`/`Presentation`:

1. The live directory must equal the source ordered list with exactly the
   selected entry removed.  Compare stable slide and persist IDs, flags,
   placeholder counts and outline metadata; `slide_count == source - 1` alone
   is not enough.
2. Each surviving slide must retain its semantic text and its live persisted
   record bytes.  The selected slide must be absent from the live directory,
   while its unreachable incremental-history bytes may remain in the document
   stream.
3. Reopen through the normal package/presentation reader and validate the live
   document/current-user mapping.  Compare all unaffected CFB stream paths and
   bytes, including `Pictures`, property streams, masters and embedded
   storages.
4. If the fixture has notes, comments, media, fonts, or other reachable
   dependencies, retain their inventories and require unchanged semantics.
   Otherwise record the fixture’s absence explicitly rather than treating an
   unobserved dependency as preserved.

An order digest or package hash alone is a false positive: a wrong slide can be
removed while the number of slides and a digest still look plausible.  The
ordered stable identities and survivor readback are the required semantic
proof.

## Allocation-region review

The allocator binary must be a separate process from native timing.  Its
region boundary must be explicit about ownership:

* take the entry snapshot after all setup and warm-up allocations;
* begin before the first operation call and end after the operation’s intended
  returned value is either deliberately retained or deliberately dropped;
* keep output ownership consistent across open, stage and finish rows;
* do not include inventory, hashing, JSON rendering, process setup, or fixture
  generation in a region;
* run each phase on a fresh editor/source so retained allocations from a prior
  phase cannot appear in the next phase’s `retained_bytes` or peak;
* report failed operations separately and do not turn a missing region into a
  zero-cost success.

The current partial `alloc_metrics.rs` has two review hazards:

1. `record_allocation`, `record_deallocation` and `record_reallocation` do not
   consult `ENABLED`, although the module says counters are inert until
   `enable()`.  Either guard all record methods or remove that claim and make
   the report’s `instrumented` meaning precise.
2. `region` snapshots counters before dropping its return value.  That is valid
   only if `retained_bytes` is intentionally “live at operation return”.  A
   closure such as `commit_once(...).map(|value| value.len())` drops the output
   before the boundary, whereas returning the `Vec` keeps it live.  The probe
   must use one convention and document it; otherwise allocation comparisons
   can differ solely because of a hidden drop boundary.  Error paths also need
   a balanced region or an explicit rejected sample record.

The allocation wrapper’s instrumentation overhead is acceptable only because
it is isolated from the native process.  It must not be used to rank native
latency, and the report must not present allocator-region retained bytes as
RSS.

## Qualification and negative controls

No current cost claim is qualified until all of the following pass for both DOC
and PPT and for every measured policy:

* source/probe/input/binary identities are captured and the source identity
  fence passes;
* the public output reopens and passes the DOC/PPT semantic oracle;
* the complete union stream/path/byte comparison passes;
* the directory metadata oracle passes, with Reuse allocation-field exceptions
  explicitly listed and Rewrite differences explicitly labelled;
* logical changed-stream length proof passes even if physical package length is
  unchanged;
* repeated fresh-process output is deterministic for the same source, edit,
  policy and binary;
* allocation and timing regions have the stated ownership boundaries; and
* the independent A/A floor and all raw samples are retained before ranking a
  phase.

The qualification harness should include deliberate failures that prove the
oracle is live: remove an empty stream, mutate one untouched stream without
changing its length, alter a directory CLSID/state/timestamp, swap in a
same-length source, replace the selected DOC text with a different string, and
remove the wrong PPT slide while preserving the count.  Each must fail before
any timing result is admitted.  A passing negative-control set is evidence that
the probe can detect the false positives called out above.

Until these gates and a clean A/A floor are recorded, describe 0728 as a
qualified-probe attempt or harness work.  Do not call its numbers a current
DOC/PPT baseline, do not revive 0617’s fractions, and do not infer a copy-
through opportunity from a reduced-source container.

## Findings in the current probe source

The completed `lib.rs` makes several of the above requirements concrete, but
the following are still open qualification findings:

* `oracle_for_output` compares the common output with the public output’s
  stream bytes and compares root/storage CLSIDs with the public output.  It
  does not compare the public output’s untouched streams with the source, so a
  public implementation that changed every stream could become the expected
  oracle and pass.  Add source-to-expected checks for every unchanged stream,
  root/storage/stream metadata, and the full path union before using the
  expected result as a control.
* The current inventory retains stream paths/bytes and storage CLSIDs but not
  stream-entry CLSIDs, entry kinds for every path, state bits, timestamps,
  sibling/child links, or the raw directory image.  Root and storage CLSIDs
  are necessary and are explicitly checked, but they do not establish the
  0663 metadata contract.  Add the bounded normalized-directory comparison
  described above, with Reuse allocation-field exceptions.
* There is no independent DOC paragraph oracle or PPT live-slide oracle.  The
  public expected output is accepted after structural/path/byte self-comparison
  without checking that DOC paragraph 0 contains the requested replacement or
  that PPT live slide 1 is absent while the surviving ordered identities and
  payloads remain.  Add those checks before deriving replacements and repeat
  them for every output policy.
* `changed_length_proof.length_change_proven` is currently true when the
  physical CFB file length changes even if no logical stream length changes.
  A changed physical file length is useful evidence but cannot prove the
  requested length-changing stream edit.  The admission bit must require an
  actual changed stream length plus the format-specific semantic proof.
* The public format timer is one whole `open → edit → commit` interval;
  `PhaseTimes::format` leaves open/stage/finish unset.  That is valid only if
  the result is named a whole public lifecycle.  It must not be reported as
  open/stage/finish attribution or compared phase-by-phase with the common
  editor.  Conversely, the common `stage` interval contains
  `put_streams_shared`, whose implementation renders, reopens and recaptures
  the candidate.  Its label must say `replace_and_validate` (or equivalent),
  and `finish` must not be presented as the complete common render cost.
* The `--policy` argument is ignored for `format` operations while still being
  printed in the result.  Reject non-default policy for format rows or report
  the policy as `ignored`; do not let a Reuse/Rewrite matrix imply two public
  format variants that do not exist.
* The case selector records an input path and later prints its digest, but the
  probe does not bind `--case` to the manifest’s expected byte count and SHA.
  A wrapper may enforce this, but the qualification record must show that
  guard; otherwise a same-length substitute can be measured under a fixture’s
  name and still pass the self-derived expected oracle.
* The allocation implementation keeps the opened editor across the open
  region, the staged editor across the stage region, and the final `Vec` across
  the finish region.  That is a coherent ownership convention, but the format
  allocation region maps the output to `len()` and drops it before the region
  boundary, so its retained bytes do not describe a returned format artifact.
  State this distinction in the schema and use the same convention for every
  phase.  Also resolve the existing `ENABLED` mismatch in `alloc_metrics.rs`:
  record methods currently update counters even when `enable()` was never
  called, despite the module’s “inert until enable” contract.
* `--oracle-only` runs the same self-derived oracle and has no deliberate
  corruption controls.  Qualification needs negative controls that remove an
  empty stream, mutate an untouched same-length stream, alter a root/storage
  metadata field, swap the source, replace the wrong DOC text, and remove the
  wrong PPT slide.  Each must fail independently of timing; otherwise the
  path/length/hash fields remain vulnerable to the false positives above.

## Source basis reviewed

The review follows the 0728 hypothesis, the retained 0617 probe design, and
0663’s implemented Reuse policy and corpus oracle.  Relevant implementation
boundaries are `litchi-doc::body_text::Snapshot`/`Edit::commit`,
`litchi-ppt::slide_order::Snapshot`/`Transaction::commit`, and
`litchi-ole-common::object::Editor::{open,put_streams_shared,finish}`.  The
0663 precedent is the raw normalized-directory comparison in
`crates/litchi-cfb/tests/sector_layout_corpus.rs`; it is stronger than the
path/length inventory used by the historical 0617 probe.

## Follow-up review of the handed-off probe

The current handoff improves the earlier partial source: the format output is
now used as an explicit expected stream-byte model, the DOC oracle checks the
ordered `Projection::All` paragraph sequence, the PPT oracle checks the live
ordered `(slide_id, persist_id, flags, text_count)` sequence, and the allocator
record methods honor `ENABLED`.  Those checks are useful, but the following
items still prevent a qualified baseline claim.

* **The current test helper does not match the JSON schema.**
  `InventorySummary.directory_entries` is a `Vec<DirectoryEntrySummary>`, while
  `fake_inventory` initializes it with `BTreeMap::new()` (the assignment near
  the test helper at the end of `src/lib.rs`).  This is a compile-time handoff
  blocker and must be corrected before a build receipt can be accepted.

* **Directory preservation is still only a public-view comparison.**
  The inventory records entry type, CLSID, logical size, start sector and the
  MiniFAT classification.  It does not retain exact directory names as raw
  records, node colours, left/right/child links, state bits, or creation and
  modification timestamps.  `semantic_directory_metadata_matches` also
  intentionally omits `start_sector` from its gate, while the differences list
  merely reports it.  Root and storage CLSID checks are necessary but do not
  prove the 0663 metadata contract.  Add the bounded normalized raw-directory
  image (zeroing only policy-owned allocation fields), compare source to the
  public expected result, and then compare each common-policy result to the
  policy-appropriate model.  The 0663 contract makes this a hard gate for
  Reuse: raw state bits, timestamps, node colours and links outside the
  explicitly planner-owned allocation fields must remain source-equal.  Rewrite
  is deliberately allowed to derive a new directory image (including colours
  and timestamps), so report those normalized differences as policy-selected
  rather than failing on physical equality; still gate its logical
  path/name/kind/CLSID hierarchy and stream bytes.  If the probe cannot make
  this Reuse/Rewrite split, leave the rows unqualified.  Comparing expected to
  actual alone can let both outputs carry the same silently lost source
  metadata.

* **The source-to-expected checks are not independent enough.**
  `unchanged_stream_bytes_match` derives `changed_paths` from the same
  source/expected byte comparison and then rechecks the complement, so its
  inner test is tautological.  The exact expected/output comparison does make
  a common result follow the public bytes, but it does not make the public
  edit's allowed changed streams a proof of the intended edit.  Keep an
  explicit source-versus-output comparison for every path outside the captured
  replacement set, and compare source/expected directory paths and storage
  hierarchy as well as stream paths.  A public edit that adds an empty storage
  or mutates an allowed stream outside the target's semantic closure must be
  rejected or recorded as an unqualified format result.

* **DOC semantics cover only one projection.**  The current `All` paragraph
  sequence check is a good ordered witness for the hard-coded main paragraph
  target, but it does not capture the target's old text or UTF-16 length, nor
  `Accepted`/`Rejected` projections, other stories, table cells, field
  results, revisions, authors, or embedded-object reachability.  At minimum,
  emit the selected position, old/new text digests and UTF-16 lengths, and
  explicit absent/unchanged results for the other public collections on each
  fixture.  Do not let a replacement of the wrong paragraph pass merely
  because it has the same count and the first paragraph happens to match.

* **PPT semantics omit ordered payload content.**  The custom projection
  proves that the live identity list is the source list with index 1 removed,
  which catches removal of the wrong identity.  It does not compare
  `SlideDirectoryEntry::list_text`, outline references/interactions, each
  survivor's semantic slide text, or the raw live persisted record.  It also
  does not expose the selected identity and survivor records in the result
  schema.  Use the public package/presentation view for those fields (and
  retain the raw live record where available), then state explicitly whether
  notes, comments, media, masters and embedded dependencies are absent or
  unchanged.  A slide count plus four identity integers is still vulnerable to
  payload corruption in a surviving slide.

* **The length gate is still too broad.**
  `logical_stream_length_change_proven` is only `any_stream_length_changed`.
  It does not require the selected DOC replacement to differ in UTF-16 length,
  the changed stream to be tied to that edit, or the PPT witness to record the
  removed live identity and the expected append/history effect.  A caller can
  supply a same-length `--text`, or a permitted stream can change for an
  unrelated reason, and still satisfy the length bit.  Store and gate on the
  format-specific old/new witness, its encoded lengths and direction, the
  selected PPT identity/order, and the expected changed-path set.  Physical
  package length remains an auxiliary field only.

* **Corruption controls are helper tests rather than live oracle tests.**
  The unit tests cover missing stream paths, a wrong DOC vector, a wrong PPT
  identity vector, and root CLSID preservation.  They do not invoke the full
  `oracle_for_output` against a mutated artifact.  There is no control for a
  same-length untouched-stream mutation, an empty-stream removal, a directory
  state/timestamp mutation, a same-length source swap, or a mutated survivor
  payload.  `--oracle-only` runs the normal self-derived operation; it does not
  run those negative fixtures.  Add bounded oracle-level mutations and assert
  each failure before timing data is admitted, preserving the mutation name
  and rejection reason in a qualification receipt.

* **The source identity fence is only partially represented.**  The format and
  common routes share one in-memory `source_bytes`, which is correct inside a
  process, and the analyzer checks the manifest SHA.  The probe does not bind
  the case to the manifest byte count/path itself, and the analyzer does not
  assert `source_inventory.file_bytes == cases[c].bytes` or the declared input
  path.  Add those fields/checks so a case name cannot be paired with a
  same-length substitute and have its self-derived expected output accepted.

* **Allocation lifetime is coherent but undocumented in the result.**
  The format output is held in an outer `Option<Vec<u8>>` across the whole
  region; the container editor is deliberately retained from `open` through
  `stage`, and the final output is held across `finish`.  Because
  `alloc_metrics::region` snapshots before dropping its closure return value,
  these boundaries measure returned/retained ownership at the intended
  boundary.  The JSON needs an explicit ownership convention and boundary
  label, otherwise `retained_bytes` is easy to misread as a per-phase fresh
  allocation total.  Failed regions still return no region/sample; if failures
  are part of qualification, record them separately rather than converting
  them to zero.

* **The analyzer can accept schema omissions.**  `qualify.py` checks only
  process exit codes and hashes; it does not parse the output JSON or reject a
  failed oracle.  `analyze.py` checks boolean fields and failure reasons, but
  does not require `schema_version`, `scope`, `policy_application_scope`, the
  declared metadata field list, semantic witness fields, source file bytes,
  or a format-policy `policy_applied == false` contract.  It also has no
  receipt for the corruption controls above.  Make qualification parse and
  validate the same schema used for capture, require the explicit format
  lifecycle label and common `replace_and_validate` stage meaning, and fail
  closed when any semantic, metadata, source-preservation, or negative-control
  field is absent.

Until these are fixed and the serialized build/qualification receipts pass,
the new source is a stronger probe candidate but its output remains harness
work, not a current DOC/PPT baseline.
