# PPTX presence/threading V5 independent review

Date: 2026-09-10 UTC
Verdict: CLEAR

## Source identity

- Frozen candidate: `/var/tmp/litchi-pptx-presence-v5-clean-20260910`
- Base revision: `fdd2b80dfa9efc246a5983a3fec64de25e5c4f63`
- Candidate manifest: `/var/tmp/litchi-pptx-presence-v5-clean-20260910.manifest.json`
- Candidate manifest SHA-256: `797c553dea9c0df2644cd17a790ff8f83cba2f2af6f3802df931364125b8ae15`
- Manifest reports exactly nine presence paths and exactly two paths changed from V4: `collaboration/codec.rs` and `collaboration/tests.rs`.
- No candidate source, staging area, or commit was modified by this review.

## Independent checks

Focused command, run with Rust 1.95, `--offline`, `CARGO_TARGET_DIR=/var/tmp/litchi-pptx-presence-v5-review-target-195`, `TMPDIR=/var/tmp`, and CPUs 8-31:

```text
CARGO_TARGET_DIR=/var/tmp/litchi-pptx-presence-v5-review-target-195 TMPDIR=/var/tmp taskset -c 8-31 cargo +1.95.0 test --offline -p litchi-pptx --all-features 'collaboration::tests::'
```

Result: `15 passed; 0 failed; 0 ignored` for the collaboration tests. The filtered package integration binaries also completed with no selected failures.

Log: `/var/tmp/litchi-pptx-presence-v5-review-20260910.log`
Log SHA-256: `f9b005e6c817a09dd82c6377c11488646d6938730b8a7370574a797e78ea2c3e`

Independent external harness:

- Harness directory: `/var/tmp/litchi-pptx-presence-v5-review-harness-20260910`
- Harness source SHA-256: `754ce9a05767dfffe09c5871ba34b00fc0298ce43f9a5d53240a5b8af07def51`
- Harness manifest SHA-256: `ad0140fa012413f9feb5f3b6961ca0d8bfb1292eca1c37bd25b0375f65c32b59`
- Harness log: `/var/tmp/litchi-pptx-presence-v5-review-harness-20260910.log`
- Harness log SHA-256: `8b37eba300a4cfbd44951c1dfbf7750e6539b2ab0a3adbc520a060b177222dff`
- Harness depended on the V5 `litchi-pptx` path and ran with Rust 1.95, `--offline`, separate target, `TMPDIR=/var/tmp`, CPUs 8-31.

Observed output:

```text
xml_noncharacter_setter=Err(Invalid("presenceInfo userId contains XML 1.0-forbidden character U+FFFE"))
xml_noncharacter_commit=false
inactive_recognized_only=ok:None
```

## Delta assessment

- `validate_string` now applies the XML 1.0 scalar predicate, including rejection of U+FFFE/U+FFFF and C0 controls while allowing tab, LF, and CR. It reports the first forbidden scalar before output serialization; the new focused test covers both noncharacters and unchanged edit state.
- The inactive recognized MCE-only branch now returns an absent typed value while retaining all recognized source extension spans as opaque. The external harness confirms parsing succeeds with `None`; the focused test confirms a later active edit preserves the unrelated active extension, the inactive fallback presence bytes, and exact inverse restoration.
- The focused tests cover the two V5 fixes and all V4 lifecycle, MCE, source-preservation, preflight, and signature-policy checks. I found no remaining blocker in this delta.

V4 proof remains applicable because the manifest identifies only the two reviewed changed files; all other eight V4 source hashes are unchanged.
