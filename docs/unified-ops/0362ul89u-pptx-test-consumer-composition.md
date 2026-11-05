# 0362ul89u — PPTX all-target consumer composition

The five independently reviewed source returns 89p–89t are composed into
one private candidate commit, including all 36 changed Rust consumers and
five leaf records. The retained actual `15d8d53a0c` keep-going check exposed
130 diagnostics; no all-target compiler success was inferred from prior
source-only audits. Its evidence is bound in records 89p–89t.

| Scope | Reviewed source / final leaf |
|---|---|
| Integration tests | `b953bbe4610d5145e645b655be018ff4d1c43214` |
| Content, notes, embedded tests | `9568aaabf71d60e47e71d123f865762465d07f15` / `b62777b82316880a584cb518d50e513c96b5daac` |
| Opened/source owners | `4e64e2d889c06da4c9ec22054eaeee3c59aa1933` |
| Shape, laser, tags | `e7ade643faafd3277eabeb1945fd54abb7e7d64d` / `8e281d96a53ff8be7a35242c74547e0891cb97a6` |
| Metadata, transitions | `254d43fbd6e15da6edc036c9065e99df3109f2fb` / `f0025673965ba8d6e3099373f525ae71a3c912a7` |

The changes migrate every identified stale call to required limits or the
current read-only projection, repair typed match guards without weakening
existing assertions, and remove unused test imports/reexports. The
fingerprint log is exposed only within the crate under `cfg(test)` so the
replay-cost assertions remain intact. No production visibility or default
wrapper is restored. All original assertions, resource dimensions, selected
profiles, lexical fidelity checks, and atomicity checks remain.

Root proves composed source bytes equal the reviewed leaf bytes after direct
rustfmt normalization of seven paths; formatting each immutable leaf preimage
produces the exact composed bytes. The only leaf-record correction adds the actual `--keep-going` flag to record 89t's
command. Direct rustfmt and repository guards are source checks; fresh
isolated F5 compiler, runtime/Clippy/glossary and performance validation
remain required. Nothing is accepted onto integration or trunk yet.
