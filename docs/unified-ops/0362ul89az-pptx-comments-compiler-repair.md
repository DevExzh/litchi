# 0362ul89az — PPTX comments compiler diagnostic repair

Status: source-ready compiler follow-up based on composed tip
`fc5ef75a9e8de76e950ebe5b1230c89074e5fd57`. Cargo, rustdoc, native fixtures,
and the full gate were not run on this leaf; the coordinator supplied the
keep-going diagnostics and will validate the composed candidate.

## Corrections

The minimal all-targets diagnostics are resolved without changing behavior or
test assertions:

- `parse_slide_comments_inner` no longer marks its forwarded tolerance option
  as mutable when it is only passed through.
- The date-time threading boundary tests pass the generated `String` as
  borrowed bytes to `selected_string_budget` and `validate_threading_replay`,
  matching their raw-source API. Existing `Vec<u8>` tests retain their slice
  coercions.

No lint allowance, compatibility alias, fallback policy, or test weakening was
introduced.

## Source-only checks

```text
rustfmt --edition 2024 --check on changed Rust files  PASS
git diff --check                                      PASS
Cargo/native/full gate/rustdoc/ABBA                 pending coordinator slot
```

No generated evidence, fixture, executable, or `3rdparty/` path is part of
this record.
