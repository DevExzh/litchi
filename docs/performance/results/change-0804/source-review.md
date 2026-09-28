# 0804 source review

Result: source review passes. I found no source blocker for the root-owned
quality gates. This review covers the current candidate archives only; it does
not claim a build, native timing run, or profile result.

## Scope and identity

The `before` helper in each of the five copies is byte-identical to the
corresponding `change-0803/candidate/after` helper. The OPC copy intentionally
has its public visibility and local test-module path, while the other four
copies use crate visibility and the shared-test path, so raw file digests are
not expected to match across those roles. Applying the exact normalization
used by `every_copy_of_the_module_is_the_canonical_code`—starting at the common
imports, normalizing those two visibility spellings, and removing the shared
test-module path—makes the OPC body and each of the other four bodies match in
both legs. The shared `litchi-opc` test source is byte-identical between
`before` and `after`. Thus the comparison is against the exact 0803
after-control, with the same test corpus and the same helper-copy layout.

The source intervention in every helper copy consists of two hunks in the
bounded linear comparison area: the call-site replacement in
`LinearCheck::first_position` and the added helper:

```text
- .find(|candidate| candidate.cmp(&Name(name)).is_eq())
+ .find(|candidate| linear_name_matches(candidate.0, name))

+#[inline]
+fn linear_name_matches(candidate: &[u8], name: &[u8]) -> bool {
+    #[cfg(test)]
+    tests::count_comparison();
+    candidate == name
+}
```

The added helper increments the existing test-only comparison counter once per
visited linear candidate and then returns `candidate == name`. The old
`Name::cmp` path also incremented that counter once per visited candidate. The
byte equality result is equivalent to `Ordering::is_eq()` for the same byte
slices, including different-length names.

## Semantic review

The intervention is confined to the bounded linear name scan. The ordered
`BTreeMap` backend and `Name` ordering implementation are unchanged. The
following control paths are unchanged from 0803: raw-key preflight, parser
handoff, `key_at`/`name_at`, lexical-error remapping before a value is read,
byte-offset calculation, trailing-whitespace termination, phase transitions,
clone state, and fused end behavior. The five copies remain parity-equivalent.

The unchanged shared tests retain coverage for quick-xml trace equivalence,
duplicate error precedence and positions, malformed tails, the short/linear/
ordered boundaries, clone behavior, fused termination, borrowed values, map
comparison growth, and canonical-copy parity. No test-only source amendment is
present in this diagnostic.

The old module documentation still calls the inherited implementation an
0802 candidate. That wording is part of the frozen control archive and does
not describe an additional source change; it is not a blocker for this
source-only comparison.

## Diagnostic bounds

This isolates the choice between ordering-based equality and byte-slice
equality in the bounded linear prefix. It does not isolate a particular
machine instruction or compiler layout, and the native leg must report those
effects as associated with the source intervention. It says nothing about the
ordered backend, production behavior, public workflow speed, resources, or
cross-format performance. The candidate must remain diagnostic-only: no
workflow advancement and no production adoption may be inferred from it.

Reviewed files:

```text
candidate/before/*.rs
candidate/after/*.rs
plan.json
README.md
```
