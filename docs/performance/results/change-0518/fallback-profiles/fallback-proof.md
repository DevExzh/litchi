# 0518 fallback `source_xml_with_hint` guard proof

These two debug integration-test profiles cover the foreign-lineage and
derived-hint fallback cases. They are safety-path evidence only; they are not
latency, allocation, RSS, or speedup measurements and are excluded from the
managed DOCX performance comparison.

Both receipts report exit code 0, one selected test, zero warmups, one repeat,
verified cleanup, and unchanged source. Each test stdout reports one passed
test. The raw scope analysis in `analysis.json` validates one positive scoped
call to
`litchi_opc::source_backed::PartView::source_xml_with_hint`, with the selected
cost equal to the raw summary and to selected self plus direct-edge cost.
The retained raw hashes are `ed5a584ad6c32e1da66f16b652caa2452e2665b78a997fe32f6e3974f96dcee9`
and `b61dab102dc5266ef81f547239b375483b2eb7f78a504b52e7a276fceb7facde`;
both use binary
`7148673fb1d1d934fa533869050b9abc6ff2ed56e805a972907922d4d2251deb` and
source manifest
`7aa4c9c0a35a9120d7a180d0466b2a1a5b7b4840e40a6462b1919453f68dfaef`.

| Case | Raw selected Ir | Direct `source_xml_part_with_hint` | `from_source_parts` | `SourceXmlPart::new` | `validate_source_xml` |
| --- | ---: | ---: | ---: | ---: | ---: |
| `equal_version_foreign_lineage_falls_back_to_current_original` | 225,021 (1x) | 225,010 (1x) | 144,860 (2x) | 144,771 (2x) | 137,627 (2x) |
| `derived_hint_returns_current_original_and_keeps_derived_bytes_live` | 180,632 (1x) | 180,621 (1x) | 144,973 (2x) | 144,884 (2x) | 137,740 (2x) |

The positive inclusive annotation tree for each case retains the same path:

```
PartView::source_xml_with_hint (1x)
  -> SourceBackedPackage::source_xml_part_with_hint (1x)
    -> SourceXmlPart::from_source_parts (2x)
      -> SourceXmlPart::new (2x)
        -> validate_source_xml (2x)
```

The child `2x` values are retained raw call metadata, which can include
collection-off invocations. They do not establish two validator executions
inside the selected method; the positive instruction costs establish the
fallback path.

Thus a foreign or derived hint reaches the current-source constructor and full
XML validator. The test assertions additionally establish that the returned
bytes are the current original, foreign or derived bytes are not returned as
the target, and the derived token remains live until its explicit drop.

Inclusive and exclusive annotations are retained for both raw profiles in
[`profile-annotations.json`](profile-annotations.json). The raw profiles and
their hashes remain unchanged; [`metadata-correction.json`](metadata-correction.json)
documents the receipt-only descriptor correction and confirms that the first
completed capture was not recaptured.
