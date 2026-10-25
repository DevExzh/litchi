# 0362ul89ao — PPTX source-related qualification cleanup

Status: source-only terminal repair based on
`e80fcebe6bc2785e3b08bf415d5a545a2cca4e9a`.

The selected `source_related` integration test already imports `Error` and
`OpenOptions`.  This leaf removes the remaining 18 redundant `litchi_pptx::`
qualifiers from those exact uses, preserving every assertion and error shape.

`rustfmt --edition 2024 --check` and `git diff --check` pass.  No Cargo,
native, rustdoc, generated evidence, fixture, executable, or `3rdparty/` path
was used or changed; the coordinator owns the qualification rerun.
