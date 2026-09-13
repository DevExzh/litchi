# Final native whole-document schema probe

The local OpenDocument 1.4 OASIS RNG was applied to `content.xml` extracted
from the API-mutated native package and both LibreOffice outputs. This is
separate from the valid minimal owner-level fixture in
[odf-v14-owner-rng-validation.md](odf-v14-owner-rng-validation.md).

Validator: `lxml.etree.RelaxNG` 6.0.2 on Python 3.14.4. RNG SHA-256:
`4034ec6be29205d5fc1ee5f42468ac6ef824287b3aba6d9289032af4fafbda7f`.

| Package | RNG result | First diagnostics |
|---|---:|---|
| `ods_metadata_native_input.ods` | `False` (12 errors) | invalid interleave/root attributes and native `table-cell` `value-type` extensions |
| `save1/ods_metadata_native_input.ods` | `False` (10 errors) | native `table-cell` `value-type` extensions; expected `table-header-rows` before `table-row` |
| `save2/ods_metadata_native_input.ods` | `False` (10 errors) | same retained native extensions/order as save1 |

Package and `content.xml` hashes are recorded by the final native run; the
packages are retained under
[native-run-final-freeze-0202-N9FJ](native-run-final-freeze-0202-N9FJ/). This
result establishes no blanket whole-package normative-validity claim. It
records the producer markup limitation while the owner-level address and
metadata grammar is validated independently.
