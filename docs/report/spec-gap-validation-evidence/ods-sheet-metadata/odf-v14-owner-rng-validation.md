# ODF 1.4 metadata-owner RNG validation

This receipt validates a minimal `office:document-content` document containing
the four sheet-metadata owners against the local OpenDocument 1.4 OASIS schema.
The fixture uses fully qualified cell endpoints (`Data.A1:Data.A2`) required by
the normative `cellRangeAddress` pattern. It includes consolidation, label
ranges, a cell-range source, and detective highlighted-range/operation values.

The validator was `lxml.etree.RelaxNG` 6.0.2 (Python 3.14.4), using
`/home/zhuhe/code/litchi/3rdparty/specs/OpenDocument-v1.4-os.zip`, extracted
entry `schemas/OpenDocument-v1.4-schema.rng`. The extracted schema and the
validation script were kept under `/var/tmp/odf-schema-v14` so this receipt does
not add a tool or fixture to the product tree.

```
schema zip sha256: 9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4
RNG sha256:        4034ec6be29205d5fc1ee5f42468ac6ef824287b3aba6d9289032af4fafbda7f
script sha256:     eb67230a570b859b45418937763945fcc67647c02f95ce53e90470ac816f966d
result:            valid=True
```

This is owner-level evidence for the address and metadata element/attribute
grammar. It does not establish that the larger native LibreOffice packages are
RNG-valid; those packages retain producer extensions and are reported
separately in `native-schema-validation-limitation.md`.
