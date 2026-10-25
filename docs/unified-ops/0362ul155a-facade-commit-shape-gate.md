# 0362ul155a — require the complete facade Commit value shape

This follow-up repairs a real gate gap found during review of
[0362ul155](0362ul155-batchcommit-accessor-repair.md).  The implementation is
`5968adc718e066d72ebbaba5f6fec6eedfacc032`, immediately after the accessor
migration `29107fe739f3ea729b643cf4406b3a8729d9864e` and its provenance record
`22c1ea4df3`.  The earlier rustdoc resolver recognized the unified facade's
owner accessors (`file`/`into_file`) but, unlike the production API, did not
require its value accessors (`value`/`into_value`).  Thus a hypothetical
facade `Commit` or `BatchCommit` with the old result shape could pass the
result-return gate.  This record closes that gap without changing the core or
format-owned Commit rules.

## Gate repair

For every path in `FACADE_COMMIT_PATHS`—the root and module BatchCommit
re-exports and `transaction::Commit`—`ReturnTypes.is_commit` now requires the
complete facade shape:

```text
file(), into_file(), patch(), diagnostics(), value(), into_value(), into_parts()
```

The API-debt inventory uses the same seven-member set when deciding whether a
facade commit nominal is covered and includes `value`/`into_value` in missing-
member diagnostics.  Format-owned and core Commit identities continue to
require their existing `package`/`into_package` shape; no alias or basename
fallback was introduced.

## Negative rustdoc-shaped fixtures

The valid facade fixture now includes both value accessors.  The inventory
test adds a negative case for each of the seven required members, proving that
removing any one member produces a `commit_type` finding and that the resolver
does not identify the type as a facade Commit.  A separate legacy-only fixture
with `results`/`into_results` instead of `value`/`into_value` also fails and
reports both missing canonical members.  These fixtures use the same
rustdoc-shaped graph consumed by the inventory and exercise both the resolver
and inventory paths.

## Source-only validation

```text
PYTHONDONTWRITEBYTECODE=1 python3 -B -m tools.test_api_debt_inventory
Ran 14 tests ... OK

PYTHONDONTWRITEBYTECODE=1 python3 -B -m tools.test_check_api_glossary
Ran 57 tests ... OK

git diff --check
```

The staged artifact guards passed: no staged evidence object at or above 1 MiB
and no staged `binaries/` path.  No Cargo, rustc, rustdoc, native,
benchmark, or full workspace gate command was run in this leaf.  The separate
full-65 C1 run still has its terminal CRUD PPT hyperlink-add failure; this
follow-up neither fixes nor validates that unrelated failure and makes no
full-gate pass claim.  The coordinator must rerun the source-bound glossary,
inventory, compiler, focused tests, native fidelity and ABBA checks after
integration.  Only redaction audit22 remains an allowed deferral.
