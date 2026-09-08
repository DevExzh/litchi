# 0468 source review: post-0467 XLSX commit/save candidates

This is a read-only source review for the next ordinary XLSX dense commit/save
candidate. The current source is `HEAD` `933cb6b80eaed36af21f4dce984bd6b896d3543e`
(`perf(xlsx): qualify single-pass cell attribute parsing`), the evidence commit
that follows the 0467 production candidate
`87733cf3b86c5500ee7aee4cf6a4cb0f3cabf6a7`. No 0468 production code is
authorized by this review. The 0468 frame-pointer capture is intended to rank
these hypotheses; historical 0466 profile shares must not be reused as current
claims.

The ordinary timed path remains the eager worksheet store materialization,
snapshot rewrite, changed-XML compaction, and sequential package write. The
source binding for this checkout records 6,992 compiled Rust/TOML/lock files
and two compile-time fixtures. The following hashes bind the source locations
examined here to the current checkout:

| Source | Location used by this review | SHA-256 |
| --- | --- | --- |
| `crates/litchi-xlsx/src/raw/worksheet/codec.rs` | `scan_cell_attributes`, `Parser::start_cell` | `69719bc7a0aa0754ab4745f49baba077303ca7909f4e9adbc5243189e85fb9a7` |
| `crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/scan.rs` | `Scanner::start_cell`, `cell_address` | `0880c6590963951ce8c801b4986ba72d981acaaba901d653cd86c8f6f2fc8a03` |
| `crates/litchi-xlsx/src/raw/worksheet/edit/codec/wire.rs` | `tag` | `6fd04c3e60272d97b9c42cc8602fd0fd2e07434dd38a8eb81eaf1caa78231f6a` |
| `crates/litchi-ooxml-common/src/xml.rs` | `unqualified_attribute_value` | `9b21c474a44fcdccb270880caca15813bcafdab4f9922f3593136bcf14ac6796` |
| `crates/litchi-xlsx/src/raw/worksheet/edit/package.rs` | `rewrite` and `rewrite_merges` | `170352f7ed181837838dbd5d841bc38300499d2336e6d42cc516507d35ec1279` |
| `crates/litchi-xlsx/src/raw/compact.rs` | `changed`, `write_start` | `7697b69475b355d39dd31de7b43651ddc88c070088c8c30dce40e91e33b6dfe0` |
| `crates/litchi-xlsx/src/raw/web.rs` | `read`, `replace`, `scan_extension_spans` | `e31f21c91050bb13c20c1691bfab9314146630bda08632a20a8e8b767cd0705d` |
| `crates/litchi-xlsx/src/cell.rs` | `Store::from_unsorted`, `Store::entry` | `26bd09c58eda4996e67b0999a7430bcd462fb51430cf2c8601cc0ecac044bd06` |
| `crates/litchi-xlsx/src/workbook/edit/semantic/transaction.rs` | worksheet rewrite, compaction, and verification | `03cc3eb5a698dca1c20c3c493d5d1d84fedbcaa39609978f7acafee5342911df` |

## Eager-parser contract already established by 0467

The current eager parser in `raw/worksheet/codec.rs:184-238` uses one
`element.attributes().with_checks(true)` traversal for the five modeled cell
fields. `r` is decoded while encountered; the scan completes before the
coordinate is parsed. `s`, `cm`, `vm`, and `t` remain raw attributes until the
existing semantic order in `Parser::start_cell`:

1. complete checked attribute scan and `r` decode;
2. `parse_a1` and row/column validation;
3. `s` decode, integer parse, and style bound;
4. `cm` decode, integer parse, and metadata bound;
5. `vm` decode, integer parse, and metadata bound;
6. `t` decode while constructing `PendingCell`.

The old 0467 implementation called
`unqualified_attribute_value` separately for `r`, `s`, `cm`, `vm`, and `t`.
Each helper completed a checked raw-attribute iteration before returning. The
new implementation therefore has these exact obligations:

| Input condition | Old observable order | Required current order |
| --- | --- | --- |
| Invalid entity in unqualified `r`, followed by a duplicate or malformed later attribute | `r` decode error from the first lookup returns before the later attribute is visited | Decode `r` at encounter time and return before the later iterator error |
| Valid `r`, followed by a duplicate raw attribute or malformed attribute syntax | The first `r` lookup completes its checked scan before `parse_a1` | Complete `with_checks(true)` scan before `parse_a1` |
| Invalid coordinate with a later invalid `s` or `t` value | `parse_a1` fails before the later field lookup decodes its value | Coordinate parsing fails before `s`, `cm`, `vm`, or `t` decoding |
| Invalid `s` with invalid `cm`, `vm`, or `t` | Style decode/parse/bound fails first | Style checks fail first |
| Valid `s`, invalid `cm` with invalid `vm` or `t` | Cell-metadata checks fail first | Cell-metadata checks fail first |
| Valid `s` and `cm`, invalid `vm` with invalid `t` | Value-metadata checks fail first | Value-metadata checks fail first |
| Duplicate ignored, unknown, or qualified raw names | The first checked helper rejects the duplicate before coordinate parsing | The single checked scan rejects it before coordinate parsing |
| Qualified `x:r`, `x:s`, `x:cm`, `x:vm`, or `x:t` | Qualified fields are ignored semantically; their raw duplicate syntax is still checked | Same behavior |
| Invalid entity in an ignored or qualified value | The shared helper does not decode that value | The scan must not decode it merely because it is present |

The 0467 tests in `raw/worksheet/tests.rs:513-641` cover these obligations
for the eager path. They are a compatibility boundary for any later change;
the source-backed and snapshot edit readers must not be changed by inference
from this eager-parser optimization.

## Candidate A: fuse snapshot cell address and tag scanning

The snapshot edit scanner has a separate repeated attribute walk at
`raw/worksheet/edit/codec/snapshot/scan.rs:951-1000`:

```text
Scanner::start_cell
  -> cell_address
       -> unqualified_attribute_value(element, "r", decoder)
  -> wire::tag(element, decoder)
```

`cell_address` performs a checked attribute iteration to obtain unqualified
`r`; `wire::tag` at `raw/worksheet/edit/codec/wire.rs:59-82` performs another
checked iteration and allocates/normalizes every attribute for the lossless
writer. A private cell-specific scanner could share one checked iteration and
still construct the existing `Tag` representation.

The two snapshot passes have a different contract from the eager parser. The
old order is:

1. `cell_address` iterates the complete attribute list with duplicate checks;
2. an invalid unqualified `r` value returns immediately while that iteration
   is in progress;
3. after a successful scan, `parse_a1`, row-mismatch, or inferred-column
   validation runs;
4. only after the address succeeds, `wire::tag` validates the element name and
   every attribute name/value, including unknown and qualified values, and
   performs its second duplicate-checked iteration.

A fused implementation must preserve the following distinctions:

- A duplicate or malformed attribute-token error found by the first checked
  iteration precedes coordinate parsing.
- An invalid unqualified `r` entity precedes a later duplicate or malformed
  attribute because `r` is decoded at encounter time.
- An invalid coordinate or row mismatch precedes an invalid entity in an
  unknown or qualified attribute value, because the first pass does not decode
  those values and the old `tag` pass runs only after address validation.
- Invalid UTF-8 in the element name or an unknown attribute name is a `tag`
  pass error after coordinate validation. A helper that decodes the tag name or
  all names before `parse_a1` would change this order.
- Qualified `x:r` is not an address. It remains in the lossless `Tag`; its
  value is decoded only at the old `tag` phase. Duplicate `x:r` names must
  still be rejected by the first raw checked iteration.
- The inferred-column and row-mismatch checks must update `row.last_column` at
  the same point as the current implementation, including on later semantic
  errors.

There are two implementation shapes with different risk profiles:

1. Keep two passes but call a cell-only `tag_after_checked_attributes` helper
   with duplicate checks disabled. This is potentially a small, low-risk
   reduction: a successful `cell_address` pass has already checked every raw
   attribute name for duplicates, and the second pass still parses syntax,
   validates UTF-8 names, and decodes values. The duplicate-disabled helper
   must be private to this path; changing `wire::tag` globally would weaken
   callers that do not have a preceding checked pass. Its tests must prove
   that duplicate rejection still occurs in the first pass and that tag-phase
   decode errors remain after coordinate errors.
2. Stage raw attribute spans or borrowed input ranges during the first pass,
   parse the coordinate, and materialize the owned `Tag` afterward. This can
   avoid the second iterator and may reduce repeated parsing, but a
   `quick_xml::Attribute<'_>` cannot be stored in `PendingCell` after the event
   borrow ends. Storing owned raw names/values may erase the allocation benefit;
   storing source ranges requires a stable input-relative representation and
   deferred UTF-8/entity validation. The staging pass must still check raw
   syntax and duplicates immediately, decode only unqualified `r` immediately,
   and defer all other tag name/value decoding until after coordinate parsing.

Focused tests for either shape should include duplicate `r`, duplicate
unknown, duplicate qualified names, an invalid `r` entity before a later
duplicate, a valid `r` with a malformed unknown value, invalid coordinates with
invalid unknown/qualified values, invalid UTF-8 names, entity-normalized
attributes, inferred addresses, row mismatch, and byte-exact preservation of
untouched cells. Existing edit dependency, formula, namespace, malformed-tail,
and reparse tests remain required.

The profile symbols that distinguish this candidate are
`snapshot::scan::scan_with_limit`, `Scanner::start_cell`,
`Scanner::cell_address`, `snapshot::wire::tag`,
`litchi_ooxml_common::xml::unqualified_attribute_value`, quick-xml attribute
iteration, and attribute normalization. The 0468 capture should determine
whether the duplicate-disabled second pass is sufficient before attempting
raw-span staging.

## Candidate B: remove per-event ownership in changed-XML compaction

The lower-risk compaction candidate is at `raw/compact.rs:25`:

```rust
let event = reader.read_event().map_err(xml_error)?.into_owned();
```

The current loop owns every event before processing it. Cargo.lock binds this
crate to quick-xml 0.41.0. Source inspection of that version shows
`NsReader::read_event` returning `Event<'i>` for the slice-backed reader and
`Writer::write_event<'a, E: Into<Event<'a>>>` consuming and synchronously
writing a borrowed event. The loop does not retain an event across the next
read: it stores only `bool` values in `preserve`. `write_start` creates a
borrowed `BytesStart` and writes it before returning; the `End` branch likewise
writes its borrowed name immediately.

Removing `.into_owned()` is therefore a bounded source-level hypothesis for
removing per-event allocations while retaining the same reader, writer,
`xml:space` stack, attribute duplicate checks, normalization, event order, and
error mapping. It still requires a compile/test gate because the borrow
relationships are enforced by the Rust type system and because all event
variants must remain accepted by the writer.

Required tests are the existing compact test plus declaration, PI, comment,
CDATA, general-reference, text, empty-element, namespace-qualified,
`xml:space` inherited/preserved, malformed attribute, malformed end-name, and
malformed-tail cases. The changed output must remain byte-for-byte equal to the
current implementation for the same input. The distinguishing fresh-profile
symbols are `raw::compact::changed`, quick-xml `NsReader::read_event` and
`Writer::write_event`, event allocation functions, and allocator call/free
sites. This candidate does not fuse the snapshot rewrite and compaction passes
and does not change the publication contract.

## Candidate C: compaction-pass fusion remains higher risk

The broader alternative is to fuse snapshot publication with
`raw::compact::changed`, currently called by
`workbook/edit/semantic/transaction.rs:1669-1677` after
`raw/worksheet/edit/package.rs::rewrite`. That could remove a complete parse
and event-writing pass, but it would combine two separate contracts: lossless
span copying and canonical compaction. Any such change must preserve XML
declarations, namespace-qualified names, comments, CDATA, general references,
entity spelling, `xml:space` inheritance, malformed-input rejection, source
provenance, and the existing no-op path. It should be considered only if the
fresh profile makes `raw::compact::changed` material and Candidate B does not
address the cost. Its distinguishing symbols are `raw::compact::changed`,
`raw::compact::write_start`, quick-xml event processing/writing, and the
snapshot scanner/writer closure.

## Store sorting clarification

`cell::Store::from_unsorted` at `cell.rs:736-761` unconditionally calls
`sort_unstable_by_key` for cells and rows, but Rust’s current unstable sort
implementation detects already-sorted input. The eager worksheet parser also
enforces nondecreasing row order at `raw/worksheet/codec.rs:559-579`, while
explicit cell columns may still arrive out of order. Adding a separate
monotonicity scan would duplicate work and is not presumed useful. The 0468
source review therefore does not recommend that change from source inspection
alone.

If a fresh profile shows meaningful `Store::from_unsorted` or
`core::slice::sort::unstable::ipnsort` cost, a specialized parser-proven
`from_sorted` path could be considered. It would need to retain the fallback
for direct unsorted callers, preserve duplicate-cell/row diagnostics, and
avoid adding a redundant pre-scan. Until that symbol is material in the new
profile, this is a deprioritized hypothesis rather than an implementation
target.

## Candidate D: the unconditional final `raw::web::read` pass

The fresh exact-commit profile makes this a material path: `raw::web::read` is
10.33% of the capture. That is a smaller inclusive symbol than the eager
worksheet parser (40.01%), snapshot scanner (24.41%), and compaction
(13.27%), but it is large enough that an ordinary cell-only commit should not
silently discard it from the optimization ranking. The values are from the
fresh capture only; no old profile share is being reused here.

The source locations are `raw/web.rs:27-237` (`read`), `raw/web.rs:240-258`
(`replace`), and `raw/web.rs:392-488` (`scan_extension_spans`). The ordinary
commit caller is `workbook/edit/semantic/transaction.rs:1602-1604` for a
requested web replacement and `:1669-1698` for the final worksheet path. After
all requested byte rewrites, the transaction calls
`raw::compact::changed(&after, ...)` at `:1671`, optionally reparses the grid
store at `:1672-1674`, and then unconditionally calls
`raw::web::read(&after)` at `:1677`. The returned `parsed_web` is compared with
the expected value only for a `Change::Web` at `:1689-1698`; for an ordinary
cell, row, column, or merge edit it is parsed and dropped. This explains why a
dense one-percent edit still pays for the web pass even when it requested no
web action.

`read` is a complete worksheet event traversal. It rejects an oversized
worksheet before creating the reader, then consumes events through `Eof` with
`NsReader::read_event`. For every `Start`, `Empty`, and `End` it resolves the
element name, checks the single SpreadsheetML root and depth state, and rejects
unknown namespace prefixes. DTDs are rejected globally. Text and CDATA are
decoded even outside a binding, so malformed text encoding in an unrelated
worksheet element currently remains an error. The reader's XML event errors,
unbalanced depth, missing root, unterminated root, and unfinished extension
state are also surfaced. This is more than a search for the web extension,
although it does not iterate arbitrary attributes on ordinary worksheet
elements: attribute decoding and duplicate checks are performed when the
recognized `ext`, `webExtensions`, `webExtension`, or `xm:f` grammar requires
them (`attribute` at `:539-560`, `reject_attributes` at `:562-570`, and
`reject_other_attributes` at `:572-583`).

The binding grammar checks that must remain are equally specific. A matching
SpreadsheetML `ext` must contain one x15 collection; the collection must have
at least one binding; each x15 binding must have exactly one `appRef` and one
xm formula; formula text, CDATA, and general references are decoded and
bounded; non-whitespace text or references outside xm:f are rejected; duplicate
collections/extensions and unexpected attributes/elements are rejected; and
the binding count and formula/string limits remain enforced. The transitional
and strict SpreadsheetML root namespaces both remain valid. `replace` first
calls `read` on its input at `:242`, then runs another complete span traversal
to preserve unrelated bytes. A web edit therefore has an input validation pass
before span mutation and the final validation/verification pass in the
transaction.

A blanket omission for cell-only edits is not semantics-preserving. The current
commit phase reports errors in this order for an ordinary worksheet rewrite:

1. the changed-XML compactor at `:1671`;
2. the optional grid-store parse at `:1672-1674` when ordinary/add/merge work
   requires it;
3. the whole-worksheet web validation at `:1677`;
4. style validation at `:1679-1681`; and
5. change-specific verification, including the web comparison when present.

Skipping `read` would allow a malformed web-extension tail, DTD, unknown
namespace prefix, invalid formula/reference, duplicate matching extension, or
missing root state to pass a cell-only publication. It would also move or
remove a web error relative to compact, grid parse, and style errors. For a
web edit, skipping the final read would additionally remove the existing
`parsed_web == expected` verification. Any candidate must preserve this phase
ordering, including the earlier `replace -> read -> scan_extension_spans`
ordering when a web action is requested.

There is one bounded direction worth measuring, but source inspection does not
prove it profitable: preserve the full event/root/depth traversal and make the
binding-specific state machine lazy until a matching direct `extLst`/`ext`/
extension URI is observed. This could avoid `values`/formula work on the common
no-web worksheet, but it cannot skip `resolved_name` for every element, DTD
rejection, text/CDATA decoding, or depth checks without changing the current
error surface. In particular, a literal search for the extension URI is not a
safe replacement because namespace prefixes, XML entities, strict/transitional
namespaces, and malformed tails are all handled by the reader. If the profile
attributes most of the 10.33% to reader/name resolution rather than binding
allocation and formula decoding, this bounded state split will have little
benefit.

The larger but still reviewable hypothesis is to return a web-validation proof
from a pass already required after rewriting, then consume that proof at
`:1677`. It must be tied to the exact post-compaction bytes or to a proof that
the web-extension spans and their semantics were preserved. A proof from the
pre-compaction input is insufficient: `raw::compact::changed` re-reads and
re-emits the worksheet, so byte spans and textual entity/whitespace spelling
can change. Fusing validation into compaction also risks reporting a web error
before a later compaction error, which would violate the current phase order;
deferring such an error while continuing the pass needs an explicit bounded
error representation and must not retain unbounded worksheet data. A full
pass reuse or fusion should therefore remain a measured follow-up, not an
assumed safe fast path.

The profile symbols that distinguish these cases are the inclusive and leaf
costs for `litchi_xlsx::raw::web::read`, quick-xml `NsReader::read_event`,
`NamespaceResolver::resolve_element`/`resolve_event`, this module's
`resolved_name`, `attribute`, `push_formula`, and allocation/deallocation
sites. Compare those with `raw::compact::changed` and its reader/writer symbols
before selecting the web experiment. The lower-risk `.into_owned()` removal in
Candidate B remains the better first code experiment if the profile shows
compact allocation cost; the 10.33% web pass should remain the next ranked
source opportunity rather than being treated as already solved.

Any web-pass change needs focused semantic coverage before adoption: ordinary
cell-only edits with no extension; valid worksheets with an untouched web
extension; web insert/replace/remove and final expected-binding verification;
strict and transitional roots; unknown prefixes and DTDs; duplicate extension
or collection nodes; invalid attributes, formula references, CDATA, text, and
empty elements; depth, binding-count, and formula-size limits; malformed XML
after the edited cell; and paired cases where compact, grid parsing, or style
validation fails as well. Each case should assert both the error class/message
and the existing phase ordering. No source-backed safe skip is selected by
this review pending those profile and test results.

No source-backed selected-cell path, store-retention bound, public API,
unknown-namespace policy, preservation rule, or ADR boundary is changed by
this review. The completed 0468 profile ranks eager Parser context at 40.01%, snapshot
scanning at 24.41%, compaction at 13.27%, and this web pass at 10.33% of
exact commit-context weight. These inclusive rows overlap. The next narrow
experiment is borrowed compaction events, while the larger validated-pass
reuse opportunities remain open.
