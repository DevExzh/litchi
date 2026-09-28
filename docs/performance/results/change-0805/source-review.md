# 0805 source review

Result: source review passes. I found no semantic or source-identity blocker
for the root-owned quality gates. This is a source review of the archived
production baseline and combined candidate; it does not certify a build,
native timing run, profile capture, or production adoption.

## Scope and identity

The six `candidate/before` files are byte-identical to the current production
files at the recorded 87359b8ef3 base: the five helper copies and the OPC
shared test source. The six `candidate/after` files are byte-identical to the
0804 `candidate/after` archive. The five helper copies in each leg become
identical after the existing test normalization removes the two visibility
spellings and the copied test-module path. The OPC helper keeps its `pub`
visibility; the other four keep `pub(crate)`. The shared test file is the only
test source in the packet, so the mirror quality run should cover the
production 14-test suite in all five crates (70 tests total) and the 0804
20-test suite in all five crates (100 tests total).

The combined source intervention is the archived 0804 implementation as a
whole: it disables quick-xml duplicate checking for the short prefix, checks
raw key prefixes before asking the parser for a value, keeps the first 32
borrowed names in one bounded array, compares those names with byte equality,
and lazily switches to an ordered map after a unique 33rd name. It also makes
the exact-empty decision on the first request. There is no additional source
change in the 0805 after leg.

## Semantic review

The candidate retains the public iterator contract. `checked_attributes` still
returns the same item and error types, and `unchecked_attributes` remains the
direct quick-xml unchecked iterator. The candidate's short states advance from
the end of the last borrowed value, so a yielded value is never reparsed. A
duplicate key is preflighted only when the key prefix has the equals boundary
that makes quick-xml check the name before parsing its value. A repeated key
with an absent equals sign therefore remains quick-xml's lexical error. For a
value error, `duplicate_before_value` re-reads only the next raw key prefix and
preserves duplicate precedence and the original byte positions.

The bounded linear path checks the stored names in arrival order and reports
the first matching name's offset. At the 32-name boundary, the current unique
item is returned once while the ordered map is seeded from the 32 stored
borrowed keys plus that item. The ordered path applies the same raw-key
preflight before parsing and keeps the first position for every name. Its
end/error transition and the linear path's fused `Done` transition preserve
the first-error and post-error behavior. I found no duplicate suppression,
value ownership, offset, or replay issue in the source diff.

The short, linear, and ordered states are all covered by the after tests,
including clone continuation and repeated exhaustion. The 39 frozen fixture
inputs cover 13 distinct-name cases, six ordinary duplicate cases, long quoted
and unterminated duplicate values, unquoted duplicate values, and four cases
each for flag, unique-tail, and equals-value syntax. The equals-value cases
retain the leading-`=` key shape used by the 0800 duplicate-error correction;
the after source also has a dedicated late leading-equals oracle test. The
after source's quick-xml oracle tests and adversarial random generator cover
the first/second/third boundary, malformed tails, duplicate-before-value
precedence, clone/fusion behavior, and the ordered fallback. These are source
coverage observations; the root quality run remains the authority for the
0805 execution result.

## Debug and API surface

`CheckedAttributes` remains the same public type and retains the same trait
methods and visibility in all five copies. The new `Phase`, `OwnCheck`,
`LinearCheck`, `OrderedCheck`, and `linear_name_matches` items are private.
No public method, trait, field, type, or copy visibility is added or removed.

There is an observable derived-`Debug` change that must stay explicit. At
construction the production iterator contains `Phase::QuickXml(0)` and an
`Attributes` value whose nested iterator state has duplicate checks enabled;
the candidate contains `Phase::First` and turns duplicate checks off before
the first request. The derived debug output therefore differs at construction
and at the short/linear/ordered transitions, even though `Phase` and the
iterator fields are private. Cloning preserves each leg's own state and
debug representation. This is a private-state/debug-text compatibility
change, not a new API surface, and it is not a production-adoption approval.

The inherited candidate comments still call the implementation an 0802
candidate and retain the stale sentence that describes an immediate map. The
actual archived code uses the bounded linear array and delayed map. This
documentation mismatch is a provenance caveat to disclose in the packet; it
does not change the reviewed behavior and is not a reason to alter the exact
0804 after archive for this preflight.

## Disposition

The source is suitable for the isolated 0805 quality, fixture, native, and
profile gates. Any timing result compares current production with this
combined candidate and must not be multiplied by earlier diagnostic ratios.
The candidate remains preflight-only: the frozen protected-boundary policy
and fresh workflow/resource/cross-format gates still govern any later
advancement, while production adoption remains false for this packet.
