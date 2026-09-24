# Theme-family removal closure

The later [MCE refusal follow-up](../theme-family-mce-refusal/README.md) adds
selective-mutation guards. Receipts here describe the closure revision recorded
in their source hashes; use the follow-up for current MCE validation.

This follow-up supersedes the wrapper-retention policy recorded in commit
`9facb6057` and its `xlsb-theme-family` evidence. It implements the later explicit
requirement to close semantically empty owned extension containers.

Removal first validates the recognized family, then widens a single deletion
span from the family to its selected `ext` and finally to the root `extLst` when
the corresponding container has only XML whitespace left. Literal whitespace,
whitespace character references, and whitespace CDATA count as empty. Comments,
unrelated elements/payload, or attributes other than namespace declarations and
the selected `ext@uri` retain the affected container. A retained empty native
extension continues to own its native URI; when closure removes that extension,
a later default add creates the normative profile.

The operation preserves all bytes outside the final deletion span. Existing
XLSB patch publication and inverse restoration preserve exact source across
save/reopen. No package relationships or physical Theme-part lifecycle rules
change. Closure uses borrowed source ranges and checks the final output bound;
it does not construct intermediate complete Theme copies.

The previous source-bound validation and performance receipts remain historical
records for their recorded revision. They are not validation of this changed
removal behavior. This directory records fresh compile-first checks and tests;
no new performance comparison or native Office acceptance is claimed.

## Validation

Independent review approved the new production diff. The compile-first gate
passes **1,136 tests including doctests**, with **12 pre-existing skips**, strict
all-target/all-feature Clippy, warning-denied rustdoc, formatting, and whitespace
checks. The **45 focused cases** comprise 22 complete-part, seven ownership,
14 XLSB family, and two new XLSB removal tests. The verifier checks **868 source
inputs**, the three reviewed production hashes, and all gate log hashes.
The reported unused `rewrite_source` issue is absent: that wrapper is already
restricted to test builds. No warning-suppression flags were added.

Reproduce from the repository root:

```sh
python3 docs/report/spec-gap-validation-evidence/theme-family-removal-closure/run_checks.py
python3 docs/report/spec-gap-validation-evidence/theme-family-removal-closure/verify.py
```

No dependencies or crate boundaries changed. Tests keep generated workbooks in
memory and use the existing workspace build target; no owned temporary build
or probe directories remain from this follow-up.
