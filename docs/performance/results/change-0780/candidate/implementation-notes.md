# 0780 candidate: static MCE baseline capabilities

This packet contains a candidate only. It is not an adoption decision and it
does not authorize applying the candidate to the live worktree. The source
change is limited to `crates/litchi-ooxml-common/src/mce/model.rs`.

## Shape

`Capabilities::ooxml_baseline()` now stores a private
`NamespaceSet::Baseline` tag. The 17 exact Transitional/Strict OOXML URI
spellings live in one static reference table, and `understands` performs an
exact membership check against that table. `Capabilities::new()` and custom
profiles retain an owned `HashSet<String>` in `NamespaceSet::Explicit`.

Adding an arbitrary namespace to a baseline profile materializes the same 17
strings into the owned set and then inserts the caller's namespace. Adding a
baseline URI that is already covered by the static profile remains a no-op.
This keeps custom registration, clone isolation, and all public method
signatures unchanged. The extension `HashSet<Name>` is untouched.

The two-variant representation is deliberate. On the measured target,
`NamespaceSet` occupies the same owner size as `HashSet<String>` through the
standard layout niche, so `Capabilities` does not grow by a padding-sized
boolean. A focused layout test records that invariant for the current target;
if another target does not provide that layout, the test calls for a matching
owner-budget update before adoption. The materialized custom path reserves an
18-entry table before inserting the 17 baseline names and the first custom
name, retaining the existing bounded owner model. No unsafe code, global cache,
new dependency, or API method is introduced.

## Differential coverage

The added model tests construct an explicit profile with the old 17-name loop
as their oracle and compare it with the static profile over:

- every strict and transitional baseline URI;
- empty, near-match, VML near-match, and unrelated custom URI negatives;
- `new`, `default`, custom registration, clone isolation, and extension
  preservation;
- legacy byte-buffer MCE output and reports;
- default streaming MCE event/report traces;
- AlternateContent choice selection;
- MustUnderstand/custom registration behavior;
- malformed XML, an unbound ignorable prefix, and an input limit.

The tests intentionally compare the old explicit set with the candidate rather
than only asserting expected output. Existing MCE tests remain the broader
coverage for preservation, limits, errors, and extension semantics.

## Debug and bounds

`Capabilities` still derives `Debug`, but the debug representation of the
private understood set changes: a baseline profile prints the tag
`Baseline` instead of a heap-backed `HashSet` containing 17 URI
strings. An explicit profile prints its owned set within an `Explicit` tag. This is an
intentional diagnostic representation change; public membership, clone,
extension, processing, and error behavior are the compatibility contract.

The static check is a bounded scan of 17 constant references. It does not
consult a global mutable cache and does not weaken any MCE namespace, input,
output, depth, directive, choice, attribute, or stream limits. This candidate
contains no benchmark or Cargo result; root-owned build, quality, differential
and public-workflow measurements must decide whether it is retained.


## Root application follow-through

`model.rs` and `model.patch` retain the original coder draft. The first root
quality attempt found a test-only array-iterator pattern error in that draft:
`for &namespace in OLD_BASELINE_NAMESPACES` was corrected to
`for namespace in OLD_BASELINE_NAMESPACES`. Formatting and this correction
are captured in `applied-model.rs` and `applied-model.patch`, the exact final
source used by the successful quality and after-build receipts. The prior
formatted file is retained as `applied-model-before-test-fix.rs`, and
`quality-0` retains the failed attempt. Production logic did not change in
this correction. Reproduction should apply `applied-model.patch`.
