# 0362ul89v — PPTX remaining test diagnostics

The actual keep-going check on `096422a20a855b138acfdd8113eb4ad1be709e91`
passes formatting across 392 Rust paths, then reports 15 library-test errors.
Twelve are unnecessary qualifications in tests whose imports are now in
scope; three are borrowed `observed` fields that need dereferencing in exact
limit match guards. This commit repairs every reported occurrence in eight
source files, retaining all assertions and production behavior.

Retained actual state:
`(deleted-coord-artifact)`.
The command is
`cargo check --keep-going --offline --locked -p litchi-pptx --all-targets --all-features`.
It exits 101 after 73.362 seconds; target size is 226,067,450 bytes.
Root verifies exact command, source/three-lock fences, logs/traces, isolated
preflight, live target guard and terminal ownership. All traced Cargo PIDs
are absent. The prior 130-diagnostic capture remains retained.

| Artifact | Bytes | SHA-256 |
|---|---:|---|
| `rows.json` | 31,777 | `6b3ebb7f4be29171947c1f560eb1cd349cb62f0f6eb92a463f2d490dfd23b85a` |
| `result.json` | 33,462 | `59f92956fd3f8421768a59b7c8ca517d84115820b290e5f5f10c80b438e80fc2` |
| `01-pptx-all-features-check-attempt-actual-f5-01.log` | 32,480 | `7b614ab458c0397b4351a1bc4ba849132e8576f3459124d65fa17e793488900d` |

Direct rustfmt and repository guards pass. A fresh actual compiler check,
all runtime/Clippy/glossary and performance qualification remain required.
