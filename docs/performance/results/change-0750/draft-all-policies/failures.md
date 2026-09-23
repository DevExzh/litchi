# The all-policy draft: dependents' test failures

The first implementation applied every new check under every policy (authored, default and streaming too). These are the test results it produced, kept as the evidence for confining the checks to the source policy. Raw logs are not retained.

## OOXML dependents (xml-minifier, litchi-opc, litchi-ooxml-common, litchi-docx, litchi-xlsx, litchi-pptx, litchi-xlsb)

passed 5581, failed 46

Failures by test target:

- 28 tests/source_backed_managed_document_edit.rs
- 7 tests/source_backed_page_setup.rs
- 5 tests/source_backed_managed_paragraph_batch.rs
- 2 tests/source_backed_glossary_story_text.rs
- 1 unittests
- 1 tests/source_backed_managed.rs
- 1 tests/source_backed_secondary_story_text.rs
- 1 tests/validation.rs

Refusal details in the failure output (count of occurrences):

- 42 `undeclared namespace prefix`
- 2 `reference to an undeclared entity`
- 1 `character reference to a character XML does not allow`

## ODF and OLE2 dependents (litchi-odf-common, litchi-odt, litchi-ods, litchi-odp, litchi-xls, litchi-ppt), read-only

passed 4963, failed 86

Failures by test target:

- 12 tests/streaming_plain_paragraphs.rs
- 11 unittests
- 10 tests/streaming_creation.rs
- 9 tests/streaming_provider.rs
- 6 tests/sheet_copy_transactions.rs
- 5 tests/streaming_text_spans.rs
- 4 tests/encryption_authoring.rs
- 4 tests/packaged_transactions.rs
- 4 tests/paragraph_transfer.rs
- 2 tests/odp_unified_transaction.rs
- 2 tests/logical_row_transactions.rs
- 2 tests/sheet_move_transactions.rs
- 2 tests/plain_paragraph_move.rs
- 1 tests/source_member_reader.rs
- 1 tests/odp_rich_content.rs
- 1 tests/sequential_text.rs
- 1 tests/sheet_transfer_transactions.rs
- 1 tests/source_cell_transactions.rs
- 1 tests/xml_references.rs
- 1 tests/generic_variable_declaration_mutation.rs
- 1 tests/header_footer_properties.rs
- 1 tests/list_label_alignment.rs
- 1 tests/odt_form_controls.rs
- 1 tests/odt_page_sequence_authoring.rs
- 1 tests/paragraph_drop_caps.rs
- 1 tests/style_columns.rs

Refusal details in the failure output (count of occurrences):

- 62 `malformed XML at byte N: undeclared namespace prefix`
- 5 `malformed XML at byte N: an XML declaration is allowed only at the start of the document`
