# Native LibreOffice value-inspection receipt

This directory retains one local LibreOffice conversion fixture for the sixteen
§6.13 names in this batch: `ERROR.TYPE`, `ISBLANK`, `ISERR`, `ISERROR`,
`ISEVEN`, `ISLOGICAL`, `ISNA`, `ISNONTEXT`, `ISNUMBER`, `ISODD`, `ISTEXT`, `N`,
`NA`, `NUMBERVALUE`, `TYPE`, and `VALUE`. The source is
[`inspection-native.fods`](inspection-native.fods). Its formulas were
recalculated by `/usr/bin/libreoffice` 26.2.5.2 with a fresh headless profile;
the resulting document is retained as [`recalculated.ods`](recalculated.ods).

The fixture has 123 formula rows. It includes raw Empty, empty Text, Number,
Logical, Text, `#N/A`, and `#DIV/0!` inputs; `TYPE` scalar and array probes;
`N` and strict `NA()` arity; truncation and 2^53 parity boundaries;
`NUMBERVALUE` separator, whitespace, percent, exponent, non-finite, and
malformed-input forms; and the required `VALUE` integer, decimal, fraction,
time, ISO date, datetime, and all seven en_US date forms.

[`native-results.json`](native-results.json) records the original formula, the
formula spelling persisted by LibreOffice (large literals may be rewritten in
scientific notation), the independently selected normative profile where one
is fixed, the typed native result, and the comparison disposition. The
receipt is an observation, not a claim that LibreOffice implements the
repository contract. It records seventeen fixed-profile divergences:
LibreOffice exposes native `TYPE` codes for formula references and logicals,
rejects Text and Empty parity conversion, rejects the omitted-separator
`NUMBERVALUE` default, accepts disallowed fraction spellings, applies a
different date/locale profile, and coerces non-Text `VALUE` arguments. Other
error tokens are compared by error kind; `#N/A` is compared exactly.

The normative source was read from the repository-local ODF distribution. Its
archive SHA-256 is
`9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4`; the
formula HTML member SHA-256 is
`ace07938ef54303b57af8472e0b66b289fc6946c32390fc23b8e13fdeeb5ffa1`.
The semantic contract is recorded in the sibling
[`contract.md`](../contract.md), SHA-256
`28f0c2538175efa2000114b9fcfae2a244af73103dde79335648bbc16b3d5f5b`. The
receipt keeps host behavior separate from the contract and does not turn native
results into expected evaluator behavior.

The independent, production-free expected corpus is
[`../inspection-goldens.json`](../inspection-goldens.json), generated and
checked by [`../inspection_oracle.py`](../inspection_oracle.py). It contains 93
observations spanning all sixteen functions, including date serial, fraction,
separator, and parity boundary cases. Run
`python3 ../inspection_oracle.py --check` from this directory to verify its
bytes. The Rust scalar consumer is
[`crates/litchi-ods/tests/ods_formula_inspection_oracle.rs`](../../../../../crates/litchi-ods/tests/ods_formula_inspection_oracle.rs);
resolver-backed Empty and Reference cases remain in the dedicated evaluation
and resource suites.

Run the bounded local check with:

```text
python3 reproduce.py
```

The script performs no network access. It checks the input SHA-256, converts
the fixture with a fresh temporary profile, checks the stable `content.xml`
SHA-256 and the exact typed formula sequence, and removes its temporary
profile/output tree on exit. The retained ODS ZIP hash is a capture hash;
container metadata can vary between conversions.
