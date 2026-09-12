# 0522 final source review

The 0522 production candidate is applied in the checkout and is frozen by the
candidate source manifest. This review covers source identity and semantics
only. It does not claim a candidate build, test, native capture, allocator
capture, or Callgrind result; runtime admission remains pending the frozen
campaign and its quality gates.

The checkout revision is `96fe51a3a70d97dfc49610788bf32963ae58ea2e`.

## Hash binding

| input | SHA-256 |
| --- | --- |
| `docs/performance/results/change-0522/cell-reference-candidate.patch` | `e65057e7aedafb6cb9947ba77662fe3050cd95c266d4eedf6cc15cb8ac7b0d40` |
| `docs/performance/results/change-0522/candidate/source.patch` | `c4c79c56f188820c2bc7d800b9ce17994f9c4094933fcd04b9b09b911caba4f0` |
| `docs/performance/results/change-0522/candidate/source-manifest.json` | `d9bac2799c7b75a8bf8829e0e0956ae94927fdda985def2aab58f8496bd773de` |
| `docs/performance/results/change-0522/plan.json` | `d7560f2aa733fdae98621d0fb2a6d85d926acd23751b87a3adb09f3e8a4703d0` |
| applied `crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/scan.rs` | `adc39e78ab5d3fed25e6d9d93259c0b4cdd5df9594f4751b7b661de187c32cc2` |
| applied `crates/litchi-xlsx/src/raw/worksheet/edit/codec/tests.rs` | `fd6fa40c037533a9546a74816202955adb18063e4b3d9f1c08323c1466f13879` |
| applied `tools/perf-baseline/src/lib.rs` | `3f8ea064c82f7a5d5884b2bbc5686df71f4b3e44e164b0078c0521dd9699f29a` |
| reviewed `crates/litchi-xlsx/src/raw/worksheet/edit/codec/wire.rs` | `86418ecd5c0517d6c5dd000c1e126db51426b3231d0a96d1bf31bc80459c707f` |
| reviewed `crates/litchi-ooxml-common/src/xml.rs` | `9b21c474a44fcdccb270880caca15813fcabdaf4b9922f3593136bcf14ac6796` |
| reviewed `crates/litchi-xlsx/src/cell_values/snapshot.rs` | `69e48ce22ca179edc9cc855a1dae5eeb2e794b702c381640f4e519bb2f176a5f` |
| reviewed `crates/litchi-xlsx/src/raw/worksheet/codec.rs` | `69719bc7a0aa0754ab4745f49baba077303ca7909f4e9adbc5243189e85fb9a7` |

The applied `candidate/source.patch` contains the three expected paths: the
production scanner, the codec tests, and the Noncompact harness. Its scanner
hunks match the recorded `cell-reference-candidate.patch`; the three current
file hashes above also match their entries in `candidate/source-manifest.json`.

## Candidate operation

The production change is confined to `scan.rs`. The new `cell_reference`
helper at lines 124–155 performs one checked attribute iteration and returns
the decoded unqualified `r` value as `Cow<str>` together with a raw-name proof.
The empty-cell path at lines 689–711 and start-cell path at lines 992–1014
call `cell_tag` only when that proof is false. `cell_address` at lines
1017–1052 retains the existing row lookup, A1 parse, row check, inferred-column
path, row cursor update, and typed address construction.

The proof begins with `element.name().as_ref() == b"c"` and remains true only
when every successfully parsed attribute has raw key `r`. With quick-xml's
checked iterator, that is exactly an unprefixed `c` with no attributes or one
unqualified `r` attribute. Those are the cases for which the unchanged
`wire.rs:105–147` `cell_tag` implementation returns `None`. A prefix,
namespace declaration, style/type/metadata field, future field, or additional
attribute forces the existing owned-tag path.

## Error and preservation proof

The former path called `unqualified_attribute_value` and then `cell_tag`. The
new helper preserves the first operation's observable contract from
`xml.rs:116–140`:

1. The checked `element.attributes()` iterator still reports malformed and
   duplicate raw attributes, mapped to `XmlError::Malformed` with the same
   string.
2. Attributes are examined in source order. Non-`r` fields are not decoded or
   UTF-8 checked during the reference pass, so coordinate and row errors still
   precede later tag-attribute errors.
3. An unqualified `r` is decoded and normalized with
   `XmlVersion::Explicit1_0` and the same reader decoder. Decode and duplicate
   reference errors retain their previous mappings and messages.
4. `parse_a1`, row comparison, inferred-column bounds, pending-row mutation,
   and `Address::at` retain their previous order. `Cow` changes ownership, not
   the string passed to the parser or used in an error message.

When the proof is false, `cell_tag(element, decoder)?` remains after address
parsing. Its name/value UTF-8 checks, normalized values, source attribute order,
prefix retention, and owned `Tag` result are unchanged. The optimized branch
does not skip validation of an unknown key: the only accepted raw key is the
known `r`, whose value was already decoded by the helper. No lexical tag bytes
were newly discarded; the existing `tag = None` representation and writer
reconstruction remain in force.

The source-backed semantic owner remains unchanged. At
`cell_values/snapshot.rs:686–694`, `Snapshot::from_rewritten_source` checks
execution before and after `validation::worksheet_xml`, reparses the complete
worksheet with `raw::worksheet::parse`, clones the source state, and stores the
rewritten bytes. The semantic parser's cell-attribute pass and relevant decode
paths at `raw/worksheet/codec.rs:172–223` and `:736–810`, followed by cell
finalization at `:989–1045`, are untouched. The candidate therefore does not
bypass validation, semantic parsing, readback, publication, or source/resource
ownership checks.

The scanner event/depth controls at `scan.rs:227–280`, checked reservations,
namespace routing, source spans, and edit limits are unchanged. The only new
per-cell state is a temporary `Cow` and boolean; an escaped reference may own a
short-lived decoded value, while no borrow is retained in the layout or
snapshot. No 0514 or 0516 fusion is present.

## Guard review

The applied codec tests at `tests.rs:260–451` exercise inferred and referenced
plain cells, a non-empty start cell, prefixed `x:c`, escaped references, and
attribute-rich tags. They compare final addresses, empty/start state, tag
presence, names, order, and decoded values. The legacy-pipeline differential
cases compare typed and display errors for malformed trailing values, duplicate
attributes, invalid references, invalid UTF-8, and truncation; the existing
cases also cover duplicate references and row mismatch.

The applied harness additions at `lib.rs:497–515`, `:20591–20711`, and
`:66587–66657` define and verify a deterministic four-sheet Noncompact shape.
The guard checks both unprefixed and prefixed cell tags, the SpreadsheetML
binding, numeric type attributes, semantic reopen, a source-backed edit/save,
output and semantic evidence, and untouched package members. The shape is
deliberately absent from `XlsxCellCrudShape::ALL` at lines 504–505, so the
frozen capture must keep passing `--xlsx-cell-crud-shape noncompact`; the 0522
plan does so for both relevant guard cases. The byte-pattern fixture is scoped
to generated numeric XML and is a guard corpus rather than a general producer
or interoperability claim. Its mixed cells intentionally exercise fallback
prefix/multi-attribute handling; the medium/dense primary shapes and codec
tests exercise the optimized plain path.

## Source verdict and runtime boundary

The source review passes: the applied patch is the recorded candidate, the
proof is equivalent to the existing compact-tag predicate, fallback and error
ordering are preserved, and semantic validation/readback and budgets remain in
place. No source-level blocker was found.

Runtime admission is still pending. It requires the candidate quality gates,
the frozen native primary and guard matrix, canonical allocator measurements,
and the planned Callgrind comparison. Retained positive-edge Ir attribution can
motivate the candidate but cannot establish elapsed-time or allocation gains;
inner Callgrind call metadata is not a timed event or allocation count.
