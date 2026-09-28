# 0802 source review

## Scope

This review covers the frozen source-only candidate in `candidate/after/` and
its `candidate/before/` snapshots, against quick-xml 0.41's attribute lexer and
the corrected production baseline at `334a22f101`. It covers all five helper
copies and the canonical shared test source. No Cargo build, test execution,
native capture, counter run, profiler, or workflow measurement was performed
for this review. The candidate remains unadopted until the root-owned gates
finish.

## Findings

### Lexer and error-boundary parity

`key_at` follows `quick_xml::events::attributes::IterState::next`: it skips
quick-xml's four whitespace bytes, consumes the first non-whitespace byte as
part of the key, stops at `=` or whitespace, and recognizes a separated `=`
only when it is the first following non-whitespace byte. This preserves raw
lexer behavior for leading-`=` keys and for non-ASCII byte sequences. The
corrected `name_at` fallback uses the same consumed-first-byte rule for errors
returned after the unchecked parser has scanned a value. This layer continues
to treat the input as raw bytes; it does not add XML Name validation.

The second, third, linear, and ordered phases all perform their key preflight
before calling the unchecked parser. A repeated key with a recognized equals
sign therefore returns `Duplicated` before an unquoted, missing, or
unterminated value is scanned. A repeated key without a recognized equals sign
is passed through to quick-xml and retains `ExpectedEq`; the fallback remaps
only `UnquotedValue`, `ExpectedValue`, and `ExpectedQuote`. Positions come from
the raw key slice or the same borrowed value offsets as quick-xml.

### Empty path, phases, and no replay

Construction marks the iterator `Done` only when the raw buffer ends exactly at
the name. A whitespace-only or malformed tail still enters quick-xml's normal
unchecked path, so the shortcut cannot suppress a lexical error. The phase
sequence is `First`, `Second`, `ShortTwo`, direct `OwnCheck::Linear`, then
`OwnCheck::Ordered`, or `Done`.

The first two attributes retain 0801's raw-key preflight and the third is
checked before its value. A successful third item with a remaining tail seeds
one boxed fixed array of 32 borrowed `Name` slices. The linear stage stores at
most 32 names and switches only after a unique 33rd item. The ordered stage is
seeded by moving the stored key slices and the current 33rd key; it never
reconstructs a checked quick-xml iterator or reparses any earlier value. Array
initialization, the bounded 32-name scan, and the ordered-map seed remain costs
for the measurement protocol to establish, rather than source-level claims.

### Bounds, slices, and positions

All raw-key recovery starts with `tag.get(from..)` and derives the post-first
byte slice only after a non-whitespace byte has been found. The resulting
`start + 1` and `end` indices are within `tag` by construction; the
`map_or(tag.len() - (start + 1), ...)` path cannot underflow. The tail scan is
also guarded by the derived index. `trailing_whitespace` uses a checked slice
range. `end_of` retains the existing borrowed-value invariant and saturating
offset additions used by the baseline. Stored names are borrowed from the
owning `BytesStart`, and the fixed array and ordered map grow only to the
number of attributes already reached.

`LinearCheck::first_position` compares the occupied array through `Name::cmp`,
so the shared test counter observes the linear stage. The fixed first/second/
third raw-slice comparisons are a constant-size prefix check. Ordered-map
lookups and inserts also use `Name::cmp`; the comparison tests cover the
linear counts, the map handoff, the `O(n log n)` bound, and growth across large
distinct tags without allowing an unbounded linear backend.

### Clone and fusion

`CheckedAttributes`, `Phase`, `OwnCheck`, the linear array and the ordered map
all derive `Clone`, so cloning at the short, linear, and ordered boundaries
copies the parser state and borrowed keys independently. Early duplicate
preflights set `Done` before returning their error. `next_own` sets `Done` on
an error or `None`, and the exact-empty path starts in `Done`; repeated calls
therefore remain fused. The shared candidate tests exercise short, linear,
ordered, malformed, duplicate, clone, and fused transitions, including the
leading-`=` late duplicate.

### Five-copy source path

The canonical OPC helper, the four dependent helper copies, and the shared
test source are byte-identical after the expected crate visibility and test
path substitutions. The candidate archive keeps those six sources together,
with no public API, dependency, unsafe-code, ownership, or call-site change.

One documentation sentence on the trait method still describes the post-third
path as immediately using ordered-map comparisons; the implementation and
module-level description correctly include the bounded linear prefix. The
sentence is retained as a documented wording erratum because the tested source
archive is frozen; it is not a semantic blocker or evidence of a different
implementation.

## Disposition

I found no source-level correctness blocker. The candidate is suitable for the
root-owned five-copy semantic, first-error, clone/fused, comparison-bound,
39-case, native, and counter gates. The review establishes no compilation,
measurement, resource, public-workflow, cross-format, or production-suitability
claim; the frozen 18-case protected policy and the no-adoption disposition
remain in force.
