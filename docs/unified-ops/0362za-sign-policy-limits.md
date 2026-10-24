# 0362za — bounded signature policy and raw CFB limits

- Date: 2026-10-24.
- Base: `945eb1d6b9ef6c56fa26c82e12758d38c1b9015d`.
- Scope: `litchi-sign`'s two public `with_limits` inventory rows and every
  consumer in the workspace.

`Policy` is configuration, so its consuming setter is now `limits` and its
borrowed accessor is `limits_value`. The old `with_limits` spelling and the
ambiguous getter are removed. The policy's weak-SHA-1, partial-coverage,
cryptographic validation, and all resource ceilings are unchanged.

`cfb::Editor` is a raw editor. Its one constructor is now
`Editor::open(bytes, format, limits)`, with a required caller-selected
`litchi_sign::Limits`; the old default `open(bytes, format)` and
`Editor::with_limits` entries are removed. Existing DOC/PPT/XLS signature
removal paths had no custom signature policy and pass the former standard
policy explicitly. Their format-specific source limits remain enforced by
the owning readers before and after signature removal.

The `sign` API glossary profile bans both the `_with_limits` suffix and the
exact `with_limits` name. Focused signature tests cover policy round-tripping,
custom CFB admission, exact clean finish allocation, signing, verification,
tamper detection, and atomic failure behavior.

Validation performed for this slice:

- `cargo fmt --all --check`
- `cargo test --offline -p litchi-sign --all-features`
- `cargo test --offline -p litchi-opc --all-features --lib`
- `cargo test --offline -p litchi-ops --all-features --lib`
- `cargo test --offline -p litchi-doc --all-features --lib`
- `python3 tools/check_api_glossary.py --profile sign`
- `python3 -m unittest tools.test_check_api_glossary`

This base checkout has no tracked `Cargo.lock`, so the offline Cargo commands
resolved and retained their local dependency graph for diagnostics (`Cargo.lock`
SHA-256 `6d7895e6056c931fc19780ff4d0a031230e2ac79b9b30a9060bc8b6fad6747da`).
The coordinator's composed gate supplies the frozen lock for final validation.

Performance is unchanged: the migration only makes the existing standard
signature policy explicit at raw-editor call sites.
