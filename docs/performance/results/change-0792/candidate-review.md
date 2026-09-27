# 0792 candidate review — exact-empty PPTX notes attribute tails

Status: archive-only candidate. This packet does not change production source,
run a build, run tests, collect native measurements, or claim an optimization.
The candidate patch applies cleanly to `4fa65522a0`; `git apply --check` was the
only validation performed here.

## Evidence and bounded hypothesis

The current 0791 profile keeps the notes scanner as the dominant observed
owner-local path. In the large Callgrind pass, `inspect_element` was entered
181,678 times and its calls to `CheckedAttributes::next` numbered 273,263. The
profile does not say how many of those elements have an exactly empty raw
attribute tail, so the removable fraction and any latency benefit remain
unknown. The candidate therefore tests a small, falsifiable work-removal seam;
it does not reinterpret the rejected 0787 cache experiment or claim a phase
fraction from sampled stacks.

## Candidate source change

`candidate.patch` adds this branch in
`crates/litchi-pptx/src/notes/codec.rs::inspect_element`, immediately after
element namespace resolution, local-name decoding, and the root conformance
check:

```rust
if element.attributes_raw().is_empty() {
    return Ok(());
}
```

The predicate is exact. A raw tail containing a space, tab, newline, namespace
declaration, or any attribute continues through `checked_attributes()`. The
branch therefore cannot bypass attribute syntax, duplicate detection,
namespace binding, UTF-8 or entity decoding, relationship projection, or the
attribute count and byte limits. Element and root validation still precede the
return, so an undeclared prefix, invalid element name, or wrong root remains a
typed refusal even when that element has no attributes.

## Focused semantic boundary test

The patch adds one test that calls the existing independent
`buffered_scan_oracle` and compares its projected values and typed refusal text
with `scan_processed_xml`. The oracle body is unchanged. The cases cover:

- exact-empty `Start` and `Empty` events;
- space, tab, and newline before an empty-element close;
- a namespace declaration;
- malformed attribute syntax, an invalid attribute name, a malformed entity
  value, and duplicate attributes;
- an undeclared element prefix with no attributes and an undeclared attribute
  prefix;
- an invalid UTF-8 element name with no attributes;
- a bare root with no attributes and an invalid root namespace; and
- the exact-empty path under the Strict namespace.

The repository's existing handcrafted, mutation, and PPTX-corpus differential
suites remain the broader semantic gate; their explicit scanner-limit cases
cover raw-byte, processed-byte, and depth refusals. The separate
`candidate-limits.patch` adds bounded direct-inspector checks for the attribute
count and decoded-byte ceilings and one Start/Empty node-ceiling comparison
against the unchanged oracle. The focused tests keep their own expected
refusal bits so a shared accidental acceptance cannot make them vacuous.

## Constraints and adoption gate

The change remains private to the PPTX notes codec and its tests. It adds no
public API, dependency, cache, retained state, unsafe code, ambient I/O,
parallel execution, or budget behavior. It preserves the notes owner and
atomic package graph boundary required by ADR 0013, the crate ownership and
dependency direction in ADRs 0001/0002/0024, the preservation and typed
refusal rules in ADR 0006, and the measured-evidence gate in ADRs 0005/0008.

Applying the patch is not adoption. The frozen 0792 follow-up is the 15-case
0785 probe: fresh paired native and allocation runs over capture, staged commit,
and full lifecycle, including ASCII and Unicode vendor controls, with baseline
qualification kept separate. It uses the 0792 policy's 3% paired-p50 benefit
gate with bootstrap upper bound below one, rejects a significant regression
over 5%, and rejects increases in allocation calls, allocated bytes, net live
bytes, or peak above entry. This is a different lane from 0791's six-block,
30-sample current-wrapper profile; no native instruction or guest-Ir result is
claimed by this archive candidate.

Callgrind or frame-pointer profiling is a deferred diagnostic follow-up after
the frozen native gate establishes a useful candidate; it is not a mandatory
0792 adoption gate. Output and semantic identity, refusal identity, error
order, and the explicit raw/processed/depth plus supplemental attribute/node
limit checks must remain unchanged before any production claim.
