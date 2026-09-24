# PPTX readonlyRecommended v2 delta review

Date: 2026-09-10 UTC
Candidate: `/var/tmp/litchi-pptx-readonly-v2-candidate-20260910`
Frozen v1: `/var/tmp/litchi-pptx-readonly-candidate-20260910`
Candidate base: `f0097deb0e80aa4ca7d4347cbba1aebdeadee764`

## Verdict

CLEAR for the v1 -> v2 delta. The delta is limited to the raw attribute scanner and its focused regressions. It correctly broadens the QName/attribute separator to XML ASCII whitespace while retaining the existing `>` and `/` terminators. The tests cover tab, LF, CR, CRLF, and mixed separators, whitespace around `=`, exact no-op bytes, scalar splice, reopen/readback, and inverse restoration. No candidate or primary files were edited, staged, or committed.

## Delta audit

`diff -qr` between v1 and v2's `readonly_recommended` module reports only:

* `src/presentation_properties/readonly_recommended/codec.rs`
* `src/presentation_properties/readonly_recommended/tests.rs`

The production change at `codec.rs:561-565` is the minimal correction:

```rust
while index < raw.len()
    && !raw[index].is_ascii_whitespace()
    && raw[index] != b'>'
    && raw[index] != b'/'
```

The test at `tests.rs:23-51` constructs valid XML with each separator and `val\t=\n'1'`, verifies `Some(true)`, exact no-op bytes, replacement to `'false'`, reopen/readback, and inverse byte restoration. This resolves the v1 defect reproduced by the external harness without changing URI/owner, boolean projection, MCE, signature, or package lifecycle code.

## Validation

* `cargo test -p litchi-pptx --all-features readonly_recommended -- --nocapture`: 12 passed, including the new whitespace regression. Log: `/var/tmp/litchi-pptx-readonly-v2-review-20260910-lib.log`, SHA-256 `dcbdc569b2569bfeddb4de755c6485323d465b312cce83028fcbd84a48e29061`.
* `cargo test -p litchi-pptx --all-features --test pptx_readonly_recommended -- --nocapture`: 4 passed. Log: `/var/tmp/litchi-pptx-readonly-v2-review-20260910-integration.log`, SHA-256 `dad822f606fb56fd553bb770c3bce0fd13823620df9dcd592e8be0f5f5b4c6e1`.
* `cargo test -p litchi-pptx --all-features --test pptx_modification_verifier -- --nocapture`: 2 passed. Log: `/var/tmp/litchi-pptx-readonly-v2-review-20260910-writer.log`, SHA-256 `4a63f78a12e8414747c70a03f3addcf6fc5a3411a19fc503e02622b0f0bc8731`.
* Separate target: `/var/tmp/litchi-pptx-readonly-v2-review-target-20260910`; Rust/cargo 1.95.0.

The parent-provided v2 full gate receipt reports the broader all-target suite and lint/docs gates green; this review independently checked the delta and focused tests only.
