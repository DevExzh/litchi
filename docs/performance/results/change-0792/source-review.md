# 0792 — source review for the exact-empty attribute-tail candidate

This is a bounded review of the current 0792 working tree and retained
pre-measurement diagnostics. The candidate is private to the PPTX notes codec;
this review does not approve adoption or make a performance claim. Native
paired captures and numerical qualification are a separate later review.

## Reviewed change and source seam

The production candidate adds one branch in
[`inspect_element`](../../../../crates/litchi-pptx/src/notes/codec.rs:447):

```rust
if element.attributes_raw().is_empty() {
    return Ok(());
}
```

It executes only after the existing element namespace resolution, local-name
UTF-8 check, and root namespace/name check. The scanner has already enforced
raw and processed byte ceilings, and `scan_processed_xml` has already charged
the event against depth and node ceilings before calling the inspector. The
branch therefore removes only the checked-attribute iterator for an exact
empty raw tail. It does not skip XML event parsing, namespace resolution,
element-name validation, root validation, or scanner limits.

quick-xml's `BytesStart::attributes_raw()` is the bytes after the element name.
The reader removes the final `/>` before constructing an empty-element
`BytesStart`, so `<p:x/>` has an empty raw tail. A space, tab, newline, slash
syntax, namespace declaration, or ordinary attribute leaves nonempty bytes and
continues through `checked_attributes()`. This preserves malformed attribute
syntax, duplicate detection, attribute-name/value UTF-8 checks, entity
unescaping, relationship projection, and attribute count/decoded-byte checks.

The early return is also after `resolved(reader.resolver().resolve_element(...))`
and the local-name conversion. An undeclared element prefix, invalid element
name, or wrong root remains a refusal even when the element has no attributes.
An undeclared attribute prefix cannot reach the branch because its attribute
tail is nonempty.

## Boundary and limit coverage

The added scanner test compares the candidate with the unchanged independent
`buffered_scan_oracle`, including projected relationship, notes-master, and
slide-ID values and the formatted refusal text. Its cases cover:

- exact-empty `Start` and `Empty` events;
- space, tab, and newline tails before an empty-element close;
- a namespace declaration, malformed attribute syntax, invalid attribute
  names, malformed entities, and duplicate names;
- undeclared element and attribute prefixes, invalid UTF-8 element names, an
  invalid bare root, and an invalid root namespace; and
- the exact-empty path under Strict conformance.

The supplemental direct-inspector tests set the running attribute counters at
their limits for empty children and confirm that the branch leaves them
unchanged. Attribute-bearing `Start` and `Empty` children still increment the
counters and return the existing `notes XML attributes` or `notes XML
attribute bytes` limit error when the counter is already at its limit. A
100,001-node `Start` case and a 100,001-node `Empty` case compare the scanner
with the unchanged oracle and retain the `notes XML nodes` refusal. These
tests exercise the limit checks at their actual existing seams; the candidate
does not move or weaken a limit.

The focused cases supplement, rather than replace, the existing handcrafted,
mutation, and PPTX-corpus differential suites. The oracle body is unchanged,
and the candidate has no public API, dependency, cache, retained state,
unsafe code, I/O, executor, ownership, or publication change.

## Namespace test correction

The second working-tree change is test-only in
[`notes/mod.rs`](../../../../crates/litchi-pptx/src/notes/mod.rs:150). The
production helper is unchanged:

```rust
fn known_namespace(value: &[u8]) -> Option<&'static str>
```

For each exact PresentationML, DrawingML, or relationship URI it returns the
corresponding static constant. `resolved` returns that static result on the
known path. For every other bound value it still calls
`std::str::from_utf8(value)`, so the successful fallback borrows the resolver
input. `Unbound` still returns the empty string and `Unknown` still produces
the same invalid-prefix message.

The original focused test compared `actual.as_ptr()` with the pointer of an
independent equal string constant. The retained Rust reference receipt states
that references to equal constant values are not required to have the same
address. The isolated baseline diagnostic passed this test; after the
candidate's scanner/test layout changed, the initial quality run failed only
on that pointer-identity assertion while preserving value equality. The
failure is recorded in `quality-0/02.log` and
[`test-correction.json`](../test-correction.json).

The correction keeps exact value equality and instead asserts that a known
result does not point into the live input `Vec`. The existing neighboring
tests continue to assert input-pointer borrowing for valid single-byte
fallback substitutions and vendor/Unicode values, while malformed-byte and
unknown-prefix refusal text remains compared with the independent original
helper. This is a principled test of the contract—static known results versus
input-borrowing fallback—and avoids relying on addresses of equal constants.
It does not weaken the namespace implementation or alter its error behavior.

The retained chronology supports the correction: the baseline diagnostic
passed, the first candidate quality run passed formatting and checking but
reported 820 tests passed, one failed, and one ignored, and no paired native
capture had started. The corrected quality receipt records formatting,
checking, tests, Clippy, and rustdoc as passing. These are retained packet
receipts; this review did not rerun them.

## Allowlist and adoption boundary

The original candidate was limited to `crates/litchi-pptx/src/notes/codec.rs`.
The test-only pointer correction was recorded before paired measurement as an
explicit source-allowlist amendment adding
`crates/litchi-pptx/src/notes/mod.rs`; no other production file is in the
0792 allowlist. The amendment does not change the candidate branch or the
scanner's semantics.

The retained adoption policy and analysis plan remain unchanged: the public
capture or lifecycle lane needs at least a 3% paired-p50 improvement with the
bootstrap upper bound below one; a significant regression above 5% rejects;
allocation calls, allocated bytes, net live bytes, and peak-above-entry bytes
may not increase. The policy uses 10,000 draws, seed 792079, and sorted
endpoints 250/9749. No policy threshold was changed by the test correction.

The source evidence supports a paired experiment of this narrow branch. It
does not establish the share of tags with an exact empty tail, a native
benefit, an allocation reduction, or adoption. Those questions belong to the
fresh 0792 capture and resource receipts.
