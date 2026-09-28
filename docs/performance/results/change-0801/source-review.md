# 0801 source review

## Review scope

This review covers the source-only candidate in `candidate/after/` against the
`candidate/before/` snapshots and the quick-xml 0.41 attribute lexer contract.
The archive identifies `54aa0f3fb4967d98c923f58b3944911b9b4525ab` as its base,
which is the corrected 0800 production commit. The five helper snapshots and
the canonical OPC test snapshot in `before/` are byte-identical to the current
workspace sources at that base. The five `after/` helpers share the same
implementation, while retaining each crate's existing visibility and shared
test-module path.

No Cargo build, test run, native capture, or performance measurement was used
for this review. The candidate remains source-only, deferred, and unadopted.

## Findings

### quick-xml key-prefix parity

`key_at` follows `quick_xml::events::attributes::IterState::next` at the
duplicate boundary: it skips the same four whitespace bytes, consumes the
first non-whitespace key byte before scanning, stops at the first `=` or
whitespace, and treats only the first following non-whitespace byte as the
separated equals-sign decision. It returns the same key start position that
quick-xml uses for `AttrError::Duplicated`.

The first-byte rule is preserved for unusual raw keys, including keys beginning
with `=` and non-ASCII bytes. The corrected `name_at` fallback uses the same
first-byte rule. This is a raw lexer parity property; it does not validate XML
Names.

The preflight is placed before each unchecked parser call that can reach a
value: `next_second`, `next_short_two`, and `OwnCheck::next`. Consequently a
repeated key followed by an unquoted, missing, or unterminated value returns
`Duplicated` before the value scan, including at the third item and after the
boxed backend begins. A key without a recognized equals sign still goes to
quick-xml and retains `ExpectedEq`.

### Offsets and borrowed slices

The phase offsets have the intended meanings. `Second` stores the unchecked
parser's next offset after the first successful value. `ShortTwo::first_end`
addresses the second key and `ShortTwo::next_start` addresses the third key.
After the third success, `OwnCheck::next_start` is set to the end of that
value. `from_prefix` recovers the first two names from those raw offsets and
takes the third name from the borrowed `Attribute`; it never reparses an
earlier value.

The `key_at` slices are guarded by `get`, and each derived end is bounded by
the source slice before it is used. `end_of` and `offset_in` rely on the same
borrowed-value invariant as the checked baseline and quick-xml's `Attributes`
iterator. The map stores borrowed key slices tied to the owning
`BytesStart`, so cloning the iterator does not detach or mutate the source.

### Phase transitions, termination, and cloning

The `First`, `Second`, `ShortTwo`, `Own`, and `Done` transitions preserve the
checked iterator's first-error behavior. A successful item with only trailing
quick-xml whitespace enters `Done` without allocating the map. A malformed
third item, end of input, or any preflight/parser error also enters `Done`.
Once an early duplicate is returned, the underlying unchecked iterator is not
advanced, but the `Done` phase makes that internal position unobservable and
ensures fused exhaustion. The same state is cloned at every phase, including
the boxed map, so subsequent results remain independent and identical.

The ordered backend is allocated only after a successful third item has a
non-whitespace tail. It seeds exactly the three names already yielded and
checks all later names with `BTreeMap`, giving bounded `O(log n)` comparisons
per later name and `O(n)` stored key entries. There is no checked-iterator
reconstruction, hidden seed consumption, or replay of an earlier value.

### Test ownership and coverage

The candidate keeps the canonical differential, first-error, recovery,
unusual-key/leading-equals, clone, fused-exhaustion, and bounded-comparison
tests in the shared OPC test source. It adds direct coverage for third and
later transitions, malformed duplicate values after the short prefix, long
prior values, clone advances through short and boxed phases, backend
duplicate preflight, and a late leading-equals duplicate. The four dependent
crates continue to compile that canonical test source through their existing
`#[path]` declarations; no divergent test copy is introduced.

## Disposition

I found no source-level correctness blocker in the rebased candidate. The
archive is suitable for the root-owned semantic gates and direct layout/
oracle checks. This review does not establish compilation, five-copy parity at
build time, the frozen 39-case preflight, native cost ratios, resource limits,
public workflow behavior, or production suitability. Those gates remain
required, with the 0799 18-case protected-error policy unchanged.
