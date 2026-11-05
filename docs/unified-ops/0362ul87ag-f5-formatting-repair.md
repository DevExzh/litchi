# 0362ul87ag — F5 scoped-formatting repair

The first actual 0362ul87a F5 attempt was intentionally stopped at row 0,
before any Cargo, rustdoc, test, native or performance command. It ran from
source commit `e13bd7e843e3dfd5f7e22387c22a881aa6d80d43`, tree
`3a72769c796af23d4f5705afa9a5aebc22360d8d`, on branch
`uo2/ooxml-properties-publication-0362ul87a`. The preserved result and row
receipt are:

```text
(deleted-coord-artifact)
(deleted-coord-artifact)
```

The `fmt-scoped` row exited 1 because `rustfmt --check` reported formatting
diffs in exactly these two paths:

```text
crates/litchi-ooxml-common/src/custom_xml/tests.rs
tools/perf-baseline/src/lib.rs
```

Its preserved log is
`00-fmt-scoped-attempt-actual-f5-00.log`, SHA-256
`394fda5b21fc3c4b8c767522da24b59fc955c96a5249da3b3bb5fbbf0a914a4f`.
The target remained empty and the recorded source and lock identities did not
change.

The repair applies only the exact command used by that row:

```text
rustfmt --edition 2024 --config skip_children=true \
  crates/litchi-ooxml-common/src/custom_xml/tests.rs \
  tools/perf-baseline/src/lib.rs
```

The resulting diff is formatter-only and is restricted to those two paths.
This record does not claim that the replacement F5 attempt has run or passed;
the replacement preflight must bind the new commit and refreshed source
identities before a new `--go` launch.
