# 0806 production lint amendment review

Review status: **source-approved for the protected preflight only**. The
applied 0806 candidate is unchanged in its archived form. The production
quality attempt is correctly retained as failed: `cargo fmt --check` passed,
while the all-features/all-targets check stopped in `litchi-sign` on the
unused `BytesStartExt::unchecked_attributes` method under `-D unused`.
This review covers the proposed call-routing repair without running a build,
test, capture, profile, or preflight.

## Exact proposed change

The applied candidate has five independent copies of the same helper module:

| Production source | Applied candidate SHA-256 |
| --- | --- |
| `crates/litchi-ole-common/src/xml_attributes.rs` | `4d6f09141cbea4ea35ee83a3ac582d996167a4206858a0ed8a4675a9927d7775` |
| `crates/litchi-opc/src/xml_attributes.rs` | `6ad5b582be2e1b5f0b24e6aaafcfb08302e0d81bf1749cf569d231202d4530c5` |
| `crates/litchi-sign/src/xml_attributes.rs` | `4d6f09141cbea4ea35ee83a3ac582d996167a4206858a0ed8a4675a9927d7775` |
| `crates/litchi-xldm/src/xml_attributes.rs` | `4d6f09141cbea4ea35ee83a3ac582d996167a4206858a0ed8a4675a9927d7775` |
| `crates/xml-minifier/src/xml_attributes.rs` | `4d6f09141cbea4ea35ee83a3ac582d996167a4206858a0ed8a4675a9927d7775` |

In each copy, `CheckedAttributes::new` currently constructs the unchecked
quick-xml iterator inline:

```rust
let mut attributes = tag.attributes();
attributes.with_checks(false);
```

The proposed amendment replaces those two statements with the existing
`tag.unchecked_attributes()` helper and removes only the constructor's now
unneeded `clippy::disallowed_methods` allowance. The helper itself remains:

```rust
let mut attributes = self.attributes();
attributes.with_checks(false);
attributes
```

The shared `crates/litchi-opc/src/xml_attributes/tests.rs` file is outside the
amendment. The quality amendment itself changes no public trait signature,
visibility, test, fixture, error path, state type, or production module outside
these five copies. The separate candidate visibility regression is corrected
below; it is not part of the five-helper call-routing rewrite.

## Semantic review

This is call routing, not an algorithm change. The helper and the former
constructor body perform the same operations in the same order and return the
same `Attributes<'a>` value with duplicate checks disabled. The constructor
still stores the same tag borrow, iterator, and `Phase::First`; first-request
empty-tag handling and every subsequent lexical, duplicate, malformed-tail,
fusion, and ordered-backend path remain unchanged. The existing helper is
`#[inline]`, so the optimized implementation has the same intended lowering;
the protected preflight must supply the execution evidence rather than this
source review claiming it.

Reusing the helper makes the existing method a real production call site in
all five copies. That addresses the specific `litchi-sign` warning while
retaining the narrow `clippy::disallowed_methods` allowance at the one helper
that intentionally calls quick-xml's unchecked API. It does not add a broad
dead-code allowance or alter the runtime counter, fixture, or candidate
algorithm.

The applied candidate and its six-file source archive were compared before
this review and match byte-for-byte. The amendment must be archived as a new
five-file source identity while retaining the exact candidate archive as the
pre-amendment witness. The shared test source and all five before/after
candidate archives must remain unchanged by the repair.

## Required next evidence

The repair is eligible for the protected native preflight described by the
coordinator: seed `806082`, 39 cases, six paired blocks, 30 samples, and the
70/100 helper-test gate. That preflight is a source-quality and semantic
invariance check; it does not add a performance or profile claim. The
production quality gate must be rerun against the amended source, and its
result must be retained separately from the earlier failed quality packet.

Until those records pass and bind the amended five-file source identity, the
candidate remains preflight-only and no public workflow capture should use the
amended source.

## Correction to the visibility assessment

The earlier statement that the candidate's public visibility was unchanged was
incorrect relative to the sealed candidate baseline. Comparing every helper's
`candidate/before` and `candidate/after` archive shows that the 0806 candidate
changed `litchi-ole-common`'s `BytesStartExt` trait and `CheckedAttributes`
struct from `pub` to `pub(crate)`. `litchi-opc` remains `pub` in both versions;
`litchi-sign`, `litchi-xldm`, and `xml-minifier` remain `pub(crate)` in both.
The current workspace and the quality-amendment after archive still contain
the accidental OLE `pub(crate)` declarations, which explains the downstream
`litchi-crypto` quality failure.

The visibility-only amendment in
`candidate-visibility-amendment/after/litchi-ole-common-xml_attributes.rs`
restores exactly those two declarations to `pub`. Its before file is
byte-identical to the quality-amendment OLE after file, and its after file is
the exact two-token rewrite; no iterator body, state, oracle, test, or timing
scope changes. This correction is reviewed separately in
`candidate-visibility-amendment/review.md`. It requires the full production
quality gate to be rerun, while a visibility-only repair does not require a
new protected native micro-capture.
