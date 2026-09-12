# 0522 pre-candidate review

This is a read-only review of the exact candidate patch and the current test and
harness files. The live production checkout remains the 0522 baseline at
`96fe51a3a70d97dfc49610788bf32963ae58ea2e`; `cell-reference-candidate.patch`
has not been applied. No build, test, capture, or source edit was performed by
this review. The supplied baseline-focused records report exit code zero after
the retained truncation-fixture and harness-counter corrections.

## Input identity

The following hashes were taken from the files reviewed at the time of this
record.

| reviewed input | SHA-256 |
| --- | --- |
| `docs/performance/results/change-0522/cell-reference-candidate.patch` | `e65057e7aedafb6cb9947ba77662fe3050cd95c266d4eedf6cc15cb8ac7b0d40` |
| live `crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/scan.rs` | `88549673fdcda62bdc622b0b398941f585d482edffae3fda18626202dc0a32f9` |
| `crates/litchi-xlsx/src/raw/worksheet/edit/codec/wire.rs` | `86418ecd5c0517d6c5dd000c1e126db51426b3231d0a96d1bf31bc80459c707f` |
| `crates/litchi-ooxml-common/src/xml.rs` | `9b21c474a44fcdccb270880caca15813fcabdaf4b9922f3593136bcf14ac6796` |
| `crates/litchi-xlsx/src/raw/worksheet/codec.rs` | `69719bc7a0aa0754ab4745f49baba077303ca7909f4e9adbc5243189e85fb9a7` |
| `crates/litchi-xlsx/src/cell_values/snapshot.rs` | `69e48ce22ca179edc9cc855a1dae5eeb2e794b702c381640f4e519bb2f176a5f` |
| current `crates/litchi-xlsx/src/raw/worksheet/edit/codec/tests.rs` | `fd6fa40c037533a9546a74816202955adb18063e4b3d9f1c08323c1466f13879` |
| current `tools/perf-baseline/src/lib.rs` | `3f8ea064c82f7a5d5884b2bbc5686df71f4b3e44e164b0078c0521dd9699f29a` |
| `docs/performance/results/change-0522/plan.json` | `d7560f2aa733fdae98621d0fb2a6d85d926acd23751b87a3adb09f3e8a4703d0` |

The retained baseline focused-record hashes are
`baseline-focused.json` = `87c7f74290c88faa41f4dd311c0f7b827b989a15f472e0569f702b3120b9e8b5`
and `baseline-harness.json` =
`5b411eaab8f703a6df893d60618c600cfab5e08822bcbb71e850b80f08948df5`. Their
source bindings match the live scanner, test, and harness hashes above.

## Candidate and retained profile evidence

The patch adds `cell_reference` at candidate lines 124–155 of
`scan.rs` (the hunk is anchored at old lines 116–121), and changes both the
empty-cell branch at candidate lines 689–711 and the start-cell branch at
candidate lines 992–1014. The address owner changes at candidate lines
1017–1052. It returns the checked, decoded unqualified `r` value as a
`Cow<str>` plus a proof bit, so one successful attribute pass can feed both
coordinate parsing and the compact-tag decision.

The retained fourth-dump Callgrind attribution supports this owner choice. The
medium profile reports positive-edge Ir of 20,978,620 for
`Scanner::start_cell`, 13,147,503 for `Scanner::cell_address`, 10,616,903 for
`unqualified_attribute_value`, 7,333,453 for `cell_tag`, and 1,532,544 for
`parse_a1`. The dense profile reports 42,197,278, 26,834,782, 21,695,460,
14,401,728, and 3,219,445 Ir respectively. These are attribution values used
to rank redundant work. Inner call metadata in the raw dumps includes
collection-off setup/readback activity and is not treated as measured timed
event, call, or allocation counts. No 0522 candidate capture is claimed here;
the allocator lane remains the source for allocation measurements.

## Correctness, error order, and resource proof

The old address path is `cell_address` followed by `cell_tag` at live
`scan.rs:653–671` and `scan.rs:952–970`. Its first operation is
`unqualified_attribute_value` at `xml.rs:116–140`. The proposed helper mirrors
the observable contract of that loop:

1. `element.attributes()` keeps quick-xml's checked iterator. Iterator failures
   still become `XmlError::Malformed(error.to_string())`.
2. Attributes remain examined in source order. Unknown attributes are neither
   name-decoded nor value-decoded during the reference pass, so a bad or
   mismatched coordinate still precedes a later tag-attribute error.
3. An unqualified `r` uses `XmlVersion::Explicit1_0` and the same decoder.
   Decode failures retain the same malformed error mapping, while the duplicate
   unqualified-reference branch retains the same invalid message.
4. The existing A1 parse, row comparison, inferred-column bounds, row cursor
   update, and `Address::at` check retain their order. Passing a borrowed or
   owned `Cow` to `parse_a1` does not change its lexical value.

The proof bit starts with the raw element name exactly equal to `c` and remains
true only when every parsed attribute has the raw key exactly `r`. With the
checked iterator this means no attributes or one unqualified `r`, precisely the
cases for which `wire.rs:105–147` returns `None` from `cell_tag`. A prefixed
element, prefixed attribute, namespace declaration, style/type/metadata field,
future field, or second attribute makes the bit false. The fallback
`cell_tag(element, decoder)` call then remains after address parsing, preserving
its UTF-8 checks, value decoding, source order, prefix retention, and owned
attribute order. The optimized branch skips no unknown-key validation because
the only accepted key is the known raw `r` spelling and its value was already
decoded.

`Snapshot::from_rewritten_source` remains unchanged at
`cell_values/snapshot.rs:686–694`: execution checks run before and after
validation, `validation::worksheet_xml` still runs, and the full semantic
`raw::worksheet::parse` still reparses the rewritten bytes. Its cell parser
continues to own the semantic cell-attribute pass and relevant decoding at
`raw/worksheet/codec.rs:172–223` and `:736–810`; cell finalization and address
validation remain at `:989–1045`. The candidate therefore changes only the
edit-layout scanner's temporary tag work and does not bypass semantic parsing,
readback, or publication behavior.

The scanner event and depth limits at `scan.rs:191–240`, allocation guards, XML
namespace routing, source spans, writer behavior, and execution budgets are
untouched. The new `Cow` and proof bit are event-local; an escaped reference
may allocate a short-lived owned value, but no source borrow is retained in a
layout or snapshot. Plain cells retain the existing `tag = None` representation
and fallback cells retain the existing owned `Tag` representation. No 0514 or
0516 fusion is present.

## Guard adequacy

The current codec tests at `tests.rs:275–339` compare addresses, empty/start
classification, compact-tag presence, names, attribute order, and decoded
values for inferred cells, plain `r` cells, a start cell, a prefixed `x:c`, an
escaped reference, and an attribute-rich cell. The differential error tests at
`tests.rs:341–414` compare typed and display errors with the old helper for
malformed trailing values, duplicate future attributes, invalid coordinates,
invalid UTF-8, and a genuinely truncated start. Existing checks at
`tests.rs:417–451` cover duplicate references, row mismatch, and malformed
references preceding later tag errors. This is a strong scanner guard for the
new proof and fallback boundary.

The current harness adds `Noncompact` as an explicitly selectable shape at
`lib.rs:497–515`, rewrites the generated four-sheet medium grid at
`lib.rs:20591–20680`, and applies it before package serialization at
`lib.rs:20682–20711`. Its test at `lib.rs:66587–66657` checks deterministic
archives, both tag spellings, the SpreadsheetML prefix binding, numeric type
attributes, semantic reopen, one source-backed edit/save, output and semantic
evidence presence, and untouched members. The frozen plan explicitly selects
this shape for native guards, including the managed path.

`Noncompact` is intentionally absent from `XlsxCellCrudShape::ALL` at
`lib.rs:504–505`; an unqualified matrix therefore does not exercise it. The
0522 capture commands pass `--xlsx-cell-crud-shape noncompact`, so omission is
an operational risk only if a future guard command relies on `ALL`. The fixture
rewriter is a deterministic byte-pattern helper over generated numeric XML,
not a general XML producer or interoperability corpus. It exercises fallback
behavior (unprefixed `c` with `r+t`, and prefixed `x:c+r`); the primary
medium/dense shapes and codec tests cover the optimized plain path.

One minor coverage gap remains: the differential malformed-attribute cases use
self-closing cells, while the start-cell success path is represented by the
inferred `<c><v>` case. Since both event handlers share the same new
`cell_address` helper and only differ in slot finalization, this is a guard
enhancement rather than a discovered semantic defect; a non-empty start cell
with a valid `r` followed by a malformed future attribute would make that
shared-path assertion explicit.

## Concrete blockers before adoption

No static correctness or resource-budget blocker was found in the patch. Two
adoption blockers remain:

1. The candidate source has not been compiled, formatted, or run through the
   candidate-focused checks in this review. The frozen quality sequence must
   clear the XLSX tests, noncompact and allocator harness guards, workspace
   check, clippy, rustdoc, boundary, and claims gates against the applied
   candidate.
2. There is no candidate native, allocator, or Callgrind comparison yet. The
   candidate must complete the frozen primary and guard matrix, show useful
   repeatable primary improvement, preserve output/semantic/resource oracles,
   and report allocator deltas from the canonical allocator lane. Retained Ir
   attribution alone cannot admit the change, and no allocation conclusion
   should be drawn from inner Callgrind metadata.

The explicit `noncompact` selection is a required capture condition because
the shape is excluded from `ALL`. If the candidate run follows the frozen
plan, this condition is already represented; otherwise the mixed
prefix/multi-attribute guard is absent.
