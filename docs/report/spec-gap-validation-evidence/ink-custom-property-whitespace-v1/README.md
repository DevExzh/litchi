Ink custom-property XML whitespace
==================================

Custom brush property names, values and units containing TAB, LF or CR previously passed construction but failed authoring readback after XML attribute normalization. The writer now emits character references for these characters so exact logical custom strings survive finish and reopen. Known typed-property emission and schema normalization are unchanged. Shared escaping still handles XML metacharacters; output sizing uses the same sink path.

Root validation passed 39 focused Ink tests, all-target/all-feature strict DrawingML Clippy, warning-denied documentation, and scoped formatting on the source in source.json. Independent ink_custom_review cleared the change, including arbitrary-name support, invalid XML rejection and exact custom-string comparisons; independently ran the custom regression and all 15 authoring tests.

The compressed source patch and raw gate logs retain reproducible hashes in source.json and root-gates.json. This is a focused authoring fix; no whole-workspace or performance claim is made.
