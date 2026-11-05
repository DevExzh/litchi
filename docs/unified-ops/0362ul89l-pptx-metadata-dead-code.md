# 0362ul89l — PPTX metadata dead-code return

Date: 2026-10-25. This is a source-only compiler repair record; it makes no
runtime, performance, qualification, or acceptance claim. The retained F5
failure remains part of the evidence until the composed candidate is checked
again.

## Retained actual F5 failure

The isolated capture is retained outside Git at
`/home/zhuhe/code/litchi-worktrees/.coord/pptx-combined-0362ul89-f5-actual-237d6801a6`.
It uses source commit
`237d6801a641b85b4477c432a26ac53e531987ed` (tree
`97a156bd9e7c7dec682a004407d20f35cb13b496`) and the approved finite F5
driver. The formatting row passed, then
`cargo check --offline --locked -p litchi-pptx --all-targets --all-features`
exited 101 after 59.157 seconds. The retained compiler output contains 113
dead-code diagnostics only: 112 unused functions or associated items and one
unread field, across 37 source paths. The failed target reached 209,664,179
bytes. No later gate row is inferred to pass.

| Artifact | Bytes | SHA-256 |
| --- | ---: | --- |
| `01-pptx-all-features-check-attempt-actual-f5-01.log` | 49,586 | `e3004cf8f79ec32a2dd2b082f4204dc1efe5a673dab7df97dcb63f9203aa06ea` |
| `diagnostics-index.json` | 16,760 | `775ff0196bcbcf8ec70f1bb66ee969419181664f262944dd3775f308c09f235e` |
| `rows.json` | 31,626 | `f85e2b7676312bb642802f0113e9762eb5334512404780c55bb833e460d2432d` |
| `result.json` | 33,305 | `304601bb309e5fce64a96285b21d625faae327f7681a045c39eba7afdb192074` |

## Source repair

Commit `aaa93fdca174b2125903832943a0bca7de53d42e`, based directly on the
retained `237d6801a641b85b4477c432a26ac53e531987ed` failure, changes only:

- `crates/litchi-pptx/src/presentation_properties/metadata/mod.rs`: remove
  the uncalled private `limits_for` helper and the uncalled
  `BoundedString::with_capacity` constructor. Production selected-limit
  routes continue to derive `Limits::from_source_limits` at their package or
  replay boundary, and bounded writers continue through `new` and
  `reserve_for`.
- `crates/litchi-pptx/src/presentation_properties/metadata/structure/codec.rs`:
  retain the custom-show patch test helper, its schema-order helper, and their
  two layout fields only under `cfg(test)`. Production custom-show edits still
  use `custom_show::rewrite_source`; section edits still use `patch_sections`.
  This removes dead production collection and does not remove a public
  capability.

The external source patch is retained at
`/home/zhuhe/code/litchi-worktrees/.coord/pptx-combined-0362ul89/metadata-dead-code-89-aaa93fdca1.patch`
(4,519 bytes, SHA-256
`76a2566583675925107d9758e0ea26ff349cd9a4dc779d43d0c7cfcefb2f662e`).
Call-site and replacement audits found no production consumer for either
deleted private helper. `git diff --check`, direct source parsing with
rustfmt, repository size guards, and reverse patch validation passed. Cargo,
rustdoc, and runtime tests were not run.

The next composed candidate still requires a fresh isolated F5 check; this
record does not reclassify the retained failure or claim acceptance.
