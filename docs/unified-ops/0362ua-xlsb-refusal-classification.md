# 0362ua — Preserve XLSB refusal kinds through File

The merged test run stopped at group 87 because two consumers expected the
pre-contract error codes. Source review also found that merged-cell requests
used `UnsupportedContent` for both an invalid value and an unsupported target.
Record 0351 §2.4 requires classification by the actual refusal kind.

The repair changes overlapping merge values, including a decoded overlapping
merge set, to `InvalidValue`. A partial unmerge names a target that cannot take
that operation and uses `UnsupportedTarget`. Partial formula-group removal
already had the correct `UnsupportedTarget` kind; its facade expectation is
migrated without changing the leaf guard. All reasons remain fixed and
content-free. The facade preserves the native typed cause, adds the `merge`
property and guidance, and uses the common kind-to-code mapping.

| Request | Pre-contract facade expectation | Contract result |
| --- | --- | --- |
| Merge overlapping an existing merge | `invalid_argument` | `invalid_value` |
| Unmerge only part of an existing merge | `invalid_argument` | `unsupported_type` |
| Remove part of a formula group | `unsupported_capability` | `unsupported_type` |

The two updated facade suites pass all 20 tests with all features. They assert
the exact kind, retained leaf cause, merge guidance and byte-identical source
after refusal for both eager and retained-source owners. Other merge behavior,
no-ops, replay/inverse, formula guards and signature/encryption tests pass in
those suites. Capability and OfficeCLI corpus checks pass without changing or
blessing goldens. The v1 OfficeCLI corpus does not exercise XLSB; these direct
regressions cover the affected error paths instead.

Only refusal branches and their tests change. Successful merge, replay and
publication paths are unchanged. This is an integration repair, not performance
acceptance or the final workspace gate. Full merged validation remains open.

The leaf merge/merged-cell suites and all-target/all-feature XLSB Clippy also
pass. Combined XLSB/ops Clippy reaches the existing DOCX `TailChange`
`large_enum_variant` failure; that failure is retained, not suppressed, and
its separate structural repair is required before the combined gate passes.
The exact commands, statuses and log hashes are in
[validation.json](results/0362ua/validation.json).
