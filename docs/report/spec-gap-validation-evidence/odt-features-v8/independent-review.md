# Independent ODT V8 mutable-provenance review

Date: 2026-09-10 UTC
Reviewer: `/root/crypto_native_vectors`
Verdict: **CLEAR for the V7 blocker and requested V8 scope**

No candidate, primary worktree, staging area, or commit was modified by this review. Only this isolated harness directory and receipt were written.

## Candidate identity

- Worktree: `/var/tmp/litchi-odt-feature-work-v8-20260910`
- Base: `42be13576f97958c81526d9b239351c72d179fa8`
- Manifest: `/var/tmp/litchi-odt-feature-manifest-v8-final-20260910.json`
- Manifest SHA-256: `6e0fc2afa0345188cbc8bc1d13167fcc05a33a7ca9169d580e4cd8ac245ea9a8`
- V8 delta patch: `/var/tmp/litchi-odt-feature-v8-delta-20260910.patch`
- V8 delta patch SHA-256: `e77fa918e1a2cf2cbaf181a395bbdcdc44ee794f8e2ceb9db7f715b14d9af8ea`
- Relevant source SHA-256:
  - `crates/litchi-odt/src/elements/element.rs`: `2e6b7af29b26c65067befee6f391b00bdd886fac235062b42461e209cccb2c7f`
  - `crates/litchi-odt/src/elements/text.rs`: `712aeee5723d32ed5def21eba08b0dab51028a7ddab466fd1fa5ce7dcc26244b`

## Independent harness

- Directory: `/var/tmp/litchi-odt-review-v8-independent-check-20260910`
- Target: `/var/tmp/litchi-odt-review-v8-independent-target-20260910`
- Source: `src/main.rs`, SHA-256 `ee1bd029416ed2d7359d9f980d5962820a91c666abec555758ae6fff7ffed402`
- Cargo manifest SHA-256: `e916dd31f4fcf5764561024337023b2223b8ea8094e40b513e904fcceb0c31ba`
- Final log: `harness-v8-final2.log`, SHA-256 `d5291e8f52d40af13c9435f4f5af4433412d5d00bb28381a0229c7dccd8e055a`
- Toolchain: Rust `1.95.0`; `RUSTUP_HOME=/tmp/litchi-spec-gap-rustup`; `TMPDIR=/var/tmp`; four Cargo jobs; CPU affinity `8-31`
- Command:
  ```text
  RUSTUP_HOME=/tmp/litchi-spec-gap-rustup CARGO_TARGET_DIR=/var/tmp/litchi-odt-review-v8-independent-target-20260910 TMPDIR=/var/tmp CARGO_BUILD_JOBS=4 taskset -c 8-31 cargo +1.95.0 run --manifest-path /var/tmp/litchi-odt-review-v8-independent-check-20260910/Cargo.toml --offline --quiet
  ```
- Exit status: `0`

## Findings

V8 fixes the V7 regression. For parsed input:

```xml
<e a="&#x31;" b="x&amp;y"><child>before</child></e>
```

The independent probes verified:

- A no-op `attributes_mut()` borrow preserves both raw lexicals and the exact serialized output; the serialized XML reopens and serializes identically.
- Changing only `a` emits `a="&amp;#x33;"` as a logical caller value while preserving `b="x&amp;y"`; the result reopens identically.
- Direct mutable-map remove/reinsert of the exact baseline spelling preserves the source lexical. Reinsert of a different value escapes it.
- Explicit `set_attribute` and named remove-then-set treat the same text as a caller-logical value and emit `&amp;#x31;`, so the named API does not accidentally retain parsed provenance.
- Parsed numbered-paragraph references survive a mutable borrow, typed text mutation, serialization, and reopen.
- A changed mutable-map value that merely looks like a character reference (`text:level = "&#x31;"` replacing a parsed plain `1`) remains caller-logical: typed mutation rejects it and the element stays unchanged.
- Reassigning a value exactly equal to its parsed baseline is intentionally source-equivalent; typed validation uses the raw source path and preserves the original output. This is the only unavoidable distinction caused by returning a raw mutable `HashMap`: equal final values cannot reveal whether the caller performed a transient assignment. The explicit setter remains the documented logical route.
- Seven malformed/reference cases remain rejected atomically by the typed numbered-paragraph mutation path: unknown entity, escaped entity-looking literal, XML control, out-of-range scalar, surrogate, missing numeric digits, and missing semicolon.
- Quotes, ampersands, `<`, and `>` inserted through the mutable map are escaped and cannot inject XML.

Representative final log lines:

```text
noop-borrow: <e a="&#x31;" b="x&amp;y"><child>before</child></e>
partial-change: <e a="&amp;#x33;" b="x&amp;y"><child>before</child></e>
remove-reinsert-same: <e a="&#x31;" b="x&amp;y"><child>before</child></e>
named-same-value: <e a="&amp;#x31;" b="x&amp;y"><child>before</child></e>
logical-same-text: rejected as caller value
raw-same-text: retained as source-equivalent baseline
independent V8 mutable provenance, baseline, reopen, logical-value, injection, and malformed-reference checks passed
```

The candidate manifest's own 1,031-test, strict Clippy, rustdoc, and formatting receipts remain separate evidence; this review did not rerun the full suite.
