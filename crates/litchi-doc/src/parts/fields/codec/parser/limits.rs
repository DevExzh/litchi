//! Allocation and switch limits for field-instruction parsers.

pub(in crate::parts::fields) const MAX_MACRO_BUTTON_INSTRUCTION_BYTES: usize = 64 * 1024;

pub(in crate::parts::fields) const MAX_GO_TO_BUTTON_INSTRUCTION_BYTES: usize = 64 * 1024;

pub(in crate::parts::fields) const MAX_MERGE_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;

pub(in crate::parts::fields) const MAX_MERGE_FIELD_SWITCHES: usize = 64;

pub(in crate::parts::fields) const MAX_MAIL_MERGE_DATA_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;

pub(in crate::parts::fields) const MAX_MAIL_MERGE_DATA_FIELD_SWITCHES: usize = 64;

pub(in crate::parts::fields) const MAX_TABLE_OF_CONTENTS_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;

pub(in crate::parts::fields) const MAX_TABLE_OF_CONTENTS_FIELD_SWITCHES: usize = 64;

pub(in crate::parts::fields) const MAX_TABLE_OF_CONTENTS_ENTRY_FIELD_INSTRUCTION_BYTES: usize =
    64 * 1024;

pub(in crate::parts::fields) const MAX_TABLE_OF_CONTENTS_ENTRY_FIELD_SWITCHES: usize = 64;

pub(in crate::parts::fields) const MAX_TABLE_OF_AUTHORITIES_ENTRY_FIELD_INSTRUCTION_BYTES: usize =
    64 * 1024;

pub(in crate::parts::fields) const MAX_TABLE_OF_AUTHORITIES_ENTRY_FIELD_SWITCHES: usize = 64;

pub(in crate::parts::fields) const MAX_INDEX_ENTRY_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;

pub(in crate::parts::fields) const MAX_INDEX_ENTRY_FIELD_SWITCHES: usize = 64;

pub(in crate::parts::fields) const MAX_REFERENCED_DOCUMENT_FIELD_INSTRUCTION_BYTES: usize =
    64 * 1024;

pub(in crate::parts::fields) const MAX_REFERENCED_DOCUMENT_FIELD_SWITCHES: usize = 64;

pub(in crate::parts::fields) const MAX_PRIVATE_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;

pub(in crate::parts::fields) const MAX_TABLE_OF_AUTHORITIES_FIELD_INSTRUCTION_BYTES: usize =
    64 * 1024;

pub(in crate::parts::fields) const MAX_TABLE_OF_AUTHORITIES_FIELD_SWITCHES: usize = 64;

pub(in crate::parts::fields) const MAX_INDEX_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;

pub(in crate::parts::fields) const MAX_INDEX_FIELD_SWITCHES: usize = 64;

pub(in crate::parts::fields) const MAX_REFERENCE_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;

pub(in crate::parts::fields) const MAX_REFERENCE_FIELD_SWITCHES: usize = 64;

pub(in crate::parts::fields) const MAX_SET_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;

pub(in crate::parts::fields) const MAX_FORMULA_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;

pub(in crate::parts::fields) const MAX_EQUATION_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;

pub(in crate::parts::fields) const MAX_HYPERLINK_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;

pub(in crate::parts::fields) const MAX_HYPERLINK_FIELD_SWITCHES: usize = 64;

pub(in crate::parts::fields) const MAX_QUOTE_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;

pub(in crate::parts::fields) const MAX_QUOTE_FIELD_SWITCHES: usize = 64;

pub(in crate::parts::fields) const MAX_PRINT_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;

pub(in crate::parts::fields) const MAX_EMBED_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;

pub(in crate::parts::fields) const MAX_BARCODE_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;

pub(in crate::parts::fields) const MAX_BIDI_OUTLINE_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;

pub(in crate::parts::fields) const MAX_SHAPE_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;

pub(in crate::parts::fields) const MAX_LEGACY_FORM_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;

pub(in crate::parts::fields) const MAX_SYMBOL_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;

pub(in crate::parts::fields) const MAX_SYMBOL_FIELD_SWITCHES: usize = 64;

pub(in crate::parts::fields) const MAX_AUTO_NUMBER_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;

pub(in crate::parts::fields) const MAX_AUTO_NUMBER_FIELD_SWITCHES: usize = 64;

pub(in crate::parts::fields) const MAX_LIST_NUMBER_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;

pub(in crate::parts::fields) const MAX_LIST_NUMBER_FIELD_SWITCHES: usize = 64;

pub(in crate::parts::fields) const MAX_SEQUENCE_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;

pub(in crate::parts::fields) const MAX_STYLE_REFERENCE_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;

pub(in crate::parts::fields) const MAX_STYLE_REFERENCE_FIELD_SWITCHES: usize = 64;

pub(in crate::parts::fields) const MAX_AUTO_TEXT_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;

pub(in crate::parts::fields) const MAX_AUTO_TEXT_FIELD_SWITCHES: usize = 64;

pub(in crate::parts::fields) const MAX_AUTO_TEXT_LIST_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;

pub(in crate::parts::fields) const MAX_AUTO_TEXT_LIST_FIELD_SWITCHES: usize = 64;

pub(in crate::parts::fields) const MAX_DOCUMENT_VARIABLE_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;

pub(in crate::parts::fields) const MAX_DOCUMENT_VARIABLE_FIELD_SWITCHES: usize = 64;

pub(in crate::parts::fields) const MAX_DOCUMENT_PROPERTY_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;

pub(in crate::parts::fields) const MAX_DOCUMENT_PROPERTY_FIELD_SWITCHES: usize = 64;

pub(in crate::parts::fields) const MAX_INFO_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;

pub(in crate::parts::fields) const MAX_INFO_FIELD_SWITCHES: usize = 64;

pub(in crate::parts::fields) const MAX_DOCUMENT_INFORMATION_FIELD_INSTRUCTION_BYTES: usize =
    64 * 1024;

pub(in crate::parts::fields) const MAX_DOCUMENT_INFORMATION_FIELD_SWITCHES: usize = 64;

pub(in crate::parts::fields) const MAX_DOCUMENT_CONTEXT_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;

pub(in crate::parts::fields) const MAX_DOCUMENT_CONTEXT_FIELD_SWITCHES: usize = 64;

pub(in crate::parts::fields) const MAX_DDE_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;

pub(in crate::parts::fields) const MAX_DDE_FIELD_SWITCHES: usize = 64;

pub(in crate::parts::fields) const MAX_LINK_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;

pub(in crate::parts::fields) const MAX_LINK_FIELD_SWITCHES: usize = 64;

pub(in crate::parts::fields) const MAX_EXTERNAL_INCLUDE_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;

pub(in crate::parts::fields) const MAX_EXTERNAL_INCLUDE_FIELD_SWITCHES: usize = 64;

pub(in crate::parts::fields) const MAX_MAIL_MERGE_COUNTER_INSTRUCTION_BYTES: usize = 64 * 1024;

pub(in crate::parts::fields) const MAX_MAIL_MERGE_NEXT_INSTRUCTION_BYTES: usize = 64 * 1024;

pub(in crate::parts::fields) const MAX_MAIL_MERGE_CONDITIONAL_CONTROL_INSTRUCTION_BYTES: usize =
    64 * 1024;

pub(in crate::parts::fields) const MAX_IF_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;

pub(in crate::parts::fields) const MAX_COMPARE_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;

pub(in crate::parts::fields) const MAX_PROMPT_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;

pub(in crate::parts::fields) const MAX_USER_IDENTITY_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;

pub(in crate::parts::fields) const MAX_ADVANCE_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;

pub(in crate::parts::fields) const MAX_ADVANCE_FIELD_ADJUSTMENTS: usize = 64;

pub(in crate::parts::fields) const MAX_MAIL_MERGE_RECIPIENT_FIELD_INSTRUCTION_BYTES: usize =
    64 * 1024;

pub(in crate::parts::fields) const MAX_MAIL_MERGE_RECIPIENT_FIELD_SWITCHES: usize = 64;
