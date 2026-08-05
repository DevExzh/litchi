//! Field instruction grammar and typed parser parts.

use super::super::model::*;
use std::result::Result as ParseResult;
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

pub(in crate::parts::fields) struct DdeParts {
    pub(in crate::parts::fields) kind: DdeFieldKind,
    pub(in crate::parts::fields) application: String,
    pub(in crate::parts::fields) source: String,
    pub(in crate::parts::fields) item: Option<String>,
    pub(in crate::parts::fields) automatic_updates: bool,
    pub(in crate::parts::fields) representation: Option<DdeRepresentation>,
    pub(in crate::parts::fields) omit_graphic_data: bool,
    pub(in crate::parts::fields) unknown_switches: Vec<MergeFieldSwitch>,
}

pub(in crate::parts::fields) struct HyperlinkParts {
    pub(in crate::parts::fields) external_target: Option<String>,
    pub(in crate::parts::fields) bookmark: Option<String>,
    pub(in crate::parts::fields) screen_tip: Option<String>,
    pub(in crate::parts::fields) target_frame: Option<String>,
    pub(in crate::parts::fields) appends_image_map_coordinates: bool,
    pub(in crate::parts::fields) opens_new_window: bool,
    pub(in crate::parts::fields) unknown_switches: Vec<MergeFieldSwitch>,
}

pub(in crate::parts::fields) struct LinkParts {
    pub(in crate::parts::fields) application_type: String,
    pub(in crate::parts::fields) source: String,
    pub(in crate::parts::fields) item: Option<String>,
    pub(in crate::parts::fields) automatic_updates: bool,
    pub(in crate::parts::fields) result_options: Vec<LinkResultOption>,
    pub(in crate::parts::fields) formatting_modes: Vec<LinkFormatting>,
    pub(in crate::parts::fields) switches: Vec<MergeFieldSwitch>,
}

pub(in crate::parts::fields) struct ExternalIncludeParts {
    pub(in crate::parts::fields) kind: IncludeFieldKind,
    pub(in crate::parts::fields) source: String,
    pub(in crate::parts::fields) bookmark: Option<String>,
    pub(in crate::parts::fields) suppress_nested_field_updates: bool,
    pub(in crate::parts::fields) omit_picture_data: bool,
    pub(in crate::parts::fields) options: Vec<ExternalIncludeOption>,
    pub(in crate::parts::fields) unknown_switches: Vec<MergeFieldSwitch>,
}

pub(in crate::parts::fields) struct TableOfContentsParts {
    pub(in crate::parts::fields) options: Vec<TableOfContentsOption>,
    pub(in crate::parts::fields) unknown_switches: Vec<MergeFieldSwitch>,
}

pub(in crate::parts::fields) struct TableOfContentsEntryParts {
    pub(in crate::parts::fields) entry: String,
    pub(in crate::parts::fields) options: Vec<TableOfContentsEntryOption>,
    pub(in crate::parts::fields) unknown_switches: Vec<MergeFieldSwitch>,
}

pub(in crate::parts::fields) struct TableOfAuthoritiesEntryParts {
    pub(in crate::parts::fields) options: Vec<TableOfAuthoritiesEntryOption>,
    pub(in crate::parts::fields) unknown_switches: Vec<MergeFieldSwitch>,
}

pub(in crate::parts::fields) struct IndexEntryParts {
    pub(in crate::parts::fields) entry: String,
    pub(in crate::parts::fields) options: Vec<IndexEntryOption>,
    pub(in crate::parts::fields) unknown_switches: Vec<MergeFieldSwitch>,
}

pub(in crate::parts::fields) struct ReferencedDocumentParts {
    pub(in crate::parts::fields) source: String,
    pub(in crate::parts::fields) relative_path: bool,
    pub(in crate::parts::fields) switches: Vec<MergeFieldSwitch>,
}

pub(in crate::parts::fields) struct TableOfAuthoritiesParts {
    pub(in crate::parts::fields) options: Vec<TableOfAuthoritiesOption>,
    pub(in crate::parts::fields) unknown_switches: Vec<MergeFieldSwitch>,
}

pub(in crate::parts::fields) struct IndexParts {
    pub(in crate::parts::fields) options: Vec<IndexOption>,
    pub(in crate::parts::fields) unknown_switches: Vec<MergeFieldSwitch>,
}

pub(in crate::parts::fields) struct ReferenceParts {
    pub(in crate::parts::fields) bookmark: String,
    pub(in crate::parts::fields) options: Vec<ReferenceFieldOption>,
    pub(in crate::parts::fields) unknown_switches: Vec<MergeFieldSwitch>,
}

pub(in crate::parts::fields) struct StyleReferenceParts {
    pub(in crate::parts::fields) style_name: String,
    pub(in crate::parts::fields) options: Vec<StyleReferenceFieldOption>,
    pub(in crate::parts::fields) unknown_switches: Vec<MergeFieldSwitch>,
}

pub(in crate::parts::fields) struct AutoTextParts {
    pub(in crate::parts::fields) entry_name: String,
    pub(in crate::parts::fields) unknown_switches: Vec<MergeFieldSwitch>,
}

pub(in crate::parts::fields) struct AutoTextListParts {
    pub(in crate::parts::fields) display_text: Option<String>,
    pub(in crate::parts::fields) options: Vec<AutoTextListOption>,
    pub(in crate::parts::fields) unknown_switches: Vec<MergeFieldSwitch>,
}

pub(in crate::parts::fields) fn parse_macro_button_parts(
    instruction: &str,
) -> Option<(String, String)> {
    if instruction.len() > MAX_MACRO_BUTTON_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    if !keyword.eq_ignore_ascii_case("MACROBUTTON") {
        return None;
    }

    let macro_name = next_field_argument(instruction, &mut position).ok()??;
    if macro_name.is_empty() {
        return None;
    }
    let display_text = next_field_argument(instruction, &mut position).ok()??;
    if display_text.is_empty() {
        return None;
    }
    if next_field_argument(instruction, &mut position)
        .ok()?
        .is_some()
    {
        return None;
    }

    Some((macro_name, display_text))
}

pub(in crate::parts::fields) fn parse_go_to_button_parts(
    instruction: &str,
) -> Option<(String, String)> {
    if instruction.len() > MAX_GO_TO_BUTTON_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    if !keyword.eq_ignore_ascii_case("GOTOBUTTON") {
        return None;
    }

    let target = next_field_argument(instruction, &mut position).ok()??;
    if target.is_empty() {
        return None;
    }
    let button_text = next_field_argument(instruction, &mut position).ok()??;
    if button_text.is_empty() {
        return None;
    }
    if next_field_argument(instruction, &mut position)
        .ok()?
        .is_some()
    {
        return None;
    }

    Some((target, button_text))
}

pub(in crate::parts::fields) fn parse_merge_field_parts(
    instruction: &str,
) -> Option<(String, Vec<MergeFieldSwitch>)> {
    if instruction.len() > MAX_MERGE_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    if !keyword.eq_ignore_ascii_case("MERGEFIELD") {
        return None;
    }

    let field_name = next_field_argument(instruction, &mut position).ok()??;
    if field_name.is_empty() {
        return None;
    }

    let mut switches = Vec::new();
    loop {
        skip_field_whitespace(instruction, &mut position);
        let Some(introducer) = next_field_character(instruction, &mut position) else {
            break;
        };
        if introducer != '\\' || switches.len() >= MAX_MERGE_FIELD_SWITCHES {
            return None;
        }

        let name = next_field_character(instruction, &mut position)?;
        if name == '\\' || name.is_whitespace() {
            return None;
        }

        skip_field_whitespace(instruction, &mut position);
        let argument = match peek_field_character(instruction, position) {
            None | Some('\\') => None,
            Some(_) => next_field_argument(instruction, &mut position).ok()?,
        };
        switches.push(MergeFieldSwitch {
            name: name.to_ascii_lowercase(),
            argument,
        });
    }

    Some((field_name, switches))
}

pub(in crate::parts::fields) fn parse_mail_merge_data_field_parts(
    instruction: &str,
) -> Option<(String, Option<String>, Vec<MergeFieldSwitch>)> {
    if instruction.len() > MAX_MAIL_MERGE_DATA_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    if !keyword.eq_ignore_ascii_case("DATA") {
        return None;
    }

    let data_source = next_field_argument(instruction, &mut position).ok()??;
    if data_source.is_empty() {
        return None;
    }

    skip_field_whitespace(instruction, &mut position);
    let header_source = match peek_field_character(instruction, position) {
        None | Some('\\') => None,
        Some(_) => {
            let source = next_field_argument(instruction, &mut position).ok()??;
            if source.is_empty() {
                return None;
            }
            Some(source)
        },
    };

    let mut switches = Vec::new();
    loop {
        skip_field_whitespace(instruction, &mut position);
        let Some(introducer) = next_field_character(instruction, &mut position) else {
            break;
        };
        if introducer != '\\' || switches.len() >= MAX_MAIL_MERGE_DATA_FIELD_SWITCHES {
            return None;
        }

        let name = next_field_character(instruction, &mut position)?;
        if name == '\\' || name.is_whitespace() {
            return None;
        }

        skip_field_whitespace(instruction, &mut position);
        let argument = match peek_field_character(instruction, position) {
            None | Some('\\') => None,
            Some(_) => next_field_argument(instruction, &mut position).ok()?,
        };
        switches.push(MergeFieldSwitch {
            name: name.to_ascii_lowercase(),
            argument,
        });
    }

    Some((data_source, header_source, switches))
}

pub(in crate::parts::fields) fn parse_table_of_contents_field_parts(
    instruction: &str,
) -> Option<TableOfContentsParts> {
    if instruction.len() > MAX_TABLE_OF_CONTENTS_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    if !keyword.eq_ignore_ascii_case("TOC") {
        return None;
    }

    let mut options = Vec::new();
    let mut unknown_switches = Vec::new();
    loop {
        skip_field_whitespace(instruction, &mut position);
        let Some(introducer) = next_field_character(instruction, &mut position) else {
            break;
        };
        if introducer != '\\'
            || options.len() + unknown_switches.len() >= MAX_TABLE_OF_CONTENTS_FIELD_SWITCHES
        {
            return None;
        }

        let name = next_field_character(instruction, &mut position)?;
        if name == '\\' || name.is_whitespace() {
            return None;
        }
        let name = name.to_ascii_lowercase();

        skip_field_whitespace(instruction, &mut position);
        let argument = match peek_field_character(instruction, position) {
            None | Some('\\') => None,
            Some(_) => next_field_argument(instruction, &mut position).ok()?,
        };
        match name {
            'a' => options.push(TableOfContentsOption::CaptionWithoutLabel(
                argument.clone()?,
            )),
            'b' => options.push(TableOfContentsOption::Bookmark(argument.clone()?)),
            'c' => options.push(TableOfContentsOption::CaptionSequence(argument.clone()?)),
            'd' => options.push(TableOfContentsOption::SequencePageSeparator(
                argument.clone()?,
            )),
            'f' => options.push(TableOfContentsOption::TableEntryIdentifier(
                argument.clone()?,
            )),
            'h' => {
                if argument.is_some() {
                    return None;
                }
                options.push(TableOfContentsOption::Hyperlinks);
            },
            'l' => options.push(TableOfContentsOption::TableEntryLevels(argument.clone()?)),
            'n' => options.push(TableOfContentsOption::OmitPageNumbers(argument)),
            'o' => options.push(TableOfContentsOption::HeadingStyleRange(argument)),
            'p' => options.push(TableOfContentsOption::EntryPageNumberSeparator(
                argument.clone()?,
            )),
            's' => options.push(TableOfContentsOption::SequenceIdentifier(argument.clone()?)),
            't' => options.push(TableOfContentsOption::StyleMappings(argument.clone()?)),
            'u' => {
                if argument.is_some() {
                    return None;
                }
                options.push(TableOfContentsOption::OutlineLevels);
            },
            'w' => {
                if argument.is_some() {
                    return None;
                }
                options.push(TableOfContentsOption::PreserveTabs);
            },
            'x' => {
                if argument.is_some() {
                    return None;
                }
                options.push(TableOfContentsOption::PreserveNewlines);
            },
            'z' => {
                if argument.is_some() {
                    return None;
                }
                options.push(TableOfContentsOption::HidePageNumbersInWebLayout);
            },
            _ => unknown_switches.push(MergeFieldSwitch { name, argument }),
        }
    }

    Some(TableOfContentsParts {
        options,
        unknown_switches,
    })
}

pub(in crate::parts::fields) fn parse_table_of_contents_entry_field_parts(
    instruction: &str,
) -> Option<TableOfContentsEntryParts> {
    if instruction.len() > MAX_TABLE_OF_CONTENTS_ENTRY_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    if !keyword.eq_ignore_ascii_case("TC") {
        return None;
    }

    skip_field_whitespace(instruction, &mut position);
    if matches!(
        peek_field_character(instruction, position),
        None | Some('\\')
    ) {
        return None;
    }
    let entry = next_field_argument(instruction, &mut position).ok()??;
    if entry.is_empty() {
        return None;
    }

    let mut options = Vec::new();
    let mut unknown_switches = Vec::new();
    loop {
        skip_field_whitespace(instruction, &mut position);
        let Some(introducer) = next_field_character(instruction, &mut position) else {
            break;
        };
        if introducer != '\\'
            || options.len() + unknown_switches.len() >= MAX_TABLE_OF_CONTENTS_ENTRY_FIELD_SWITCHES
        {
            return None;
        }

        let name = next_field_character(instruction, &mut position)?;
        if name == '\\' || name.is_whitespace() {
            return None;
        }
        let name = name.to_ascii_lowercase();

        skip_field_whitespace(instruction, &mut position);
        let argument = match peek_field_character(instruction, position) {
            None | Some('\\') => None,
            Some(_) => next_field_argument(instruction, &mut position).ok()?,
        };
        match name {
            'f' => options.push(TableOfContentsEntryOption::ListIdentifier(argument?)),
            'l' => options.push(TableOfContentsEntryOption::Level(argument?)),
            'n' => {
                if argument.is_some() {
                    return None;
                }
                options.push(TableOfContentsEntryOption::OmitPageNumber);
            },
            _ => unknown_switches.push(MergeFieldSwitch { name, argument }),
        }
    }

    Some(TableOfContentsEntryParts {
        entry,
        options,
        unknown_switches,
    })
}

pub(in crate::parts::fields) fn parse_table_of_authorities_entry_field_parts(
    instruction: &str,
) -> Option<TableOfAuthoritiesEntryParts> {
    if instruction.len() > MAX_TABLE_OF_AUTHORITIES_ENTRY_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    if !keyword.eq_ignore_ascii_case("TA") {
        return None;
    }

    let mut options = Vec::new();
    let mut unknown_switches = Vec::new();
    loop {
        skip_field_whitespace(instruction, &mut position);
        let Some(introducer) = next_field_character(instruction, &mut position) else {
            break;
        };
        if introducer != '\\'
            || options.len() + unknown_switches.len()
                >= MAX_TABLE_OF_AUTHORITIES_ENTRY_FIELD_SWITCHES
        {
            return None;
        }

        let name = next_field_character(instruction, &mut position)?;
        if name == '\\' || name.is_whitespace() {
            return None;
        }
        let name = name.to_ascii_lowercase();

        skip_field_whitespace(instruction, &mut position);
        let argument = match peek_field_character(instruction, position) {
            None | Some('\\') => None,
            Some(_) => next_field_argument(instruction, &mut position).ok()?,
        };
        match name {
            'b' => {
                if argument.is_some() {
                    return None;
                }
                options.push(TableOfAuthoritiesEntryOption::BoldPageNumber);
            },
            'c' => options.push(TableOfAuthoritiesEntryOption::Category(argument?)),
            'i' => {
                if argument.is_some() {
                    return None;
                }
                options.push(TableOfAuthoritiesEntryOption::ItalicPageNumber);
            },
            'l' => options.push(TableOfAuthoritiesEntryOption::LongCitation(argument?)),
            'r' => options.push(TableOfAuthoritiesEntryOption::PageRangeBookmark(argument?)),
            's' => options.push(TableOfAuthoritiesEntryOption::ShortCitation(argument?)),
            _ => unknown_switches.push(MergeFieldSwitch { name, argument }),
        }
    }

    Some(TableOfAuthoritiesEntryParts {
        options,
        unknown_switches,
    })
}

pub(in crate::parts::fields) fn parse_index_entry_field_parts(
    instruction: &str,
) -> Option<IndexEntryParts> {
    if instruction.len() > MAX_INDEX_ENTRY_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    if !keyword.eq_ignore_ascii_case("XE") {
        return None;
    }

    skip_field_whitespace(instruction, &mut position);
    if matches!(
        peek_field_character(instruction, position),
        None | Some('\\')
    ) {
        return None;
    }
    let entry = next_field_argument(instruction, &mut position).ok()??;
    if entry.is_empty() {
        return None;
    }

    let mut options = Vec::new();
    let mut unknown_switches = Vec::new();
    loop {
        skip_field_whitespace(instruction, &mut position);
        let Some(introducer) = next_field_character(instruction, &mut position) else {
            break;
        };
        if introducer != '\\'
            || options.len() + unknown_switches.len() >= MAX_INDEX_ENTRY_FIELD_SWITCHES
        {
            return None;
        }

        let name = next_field_character(instruction, &mut position)?;
        if name == '\\' || name.is_whitespace() {
            return None;
        }
        let name = name.to_ascii_lowercase();

        skip_field_whitespace(instruction, &mut position);
        let argument = match peek_field_character(instruction, position) {
            None | Some('\\') => None,
            Some(_) => next_field_argument(instruction, &mut position).ok()?,
        };
        match name {
            'b' => {
                if argument.is_some() {
                    return None;
                }
                options.push(IndexEntryOption::BoldPageNumber);
            },
            'f' => options.push(IndexEntryOption::EntryType(argument?)),
            'i' => {
                if argument.is_some() {
                    return None;
                }
                options.push(IndexEntryOption::ItalicPageNumber);
            },
            'r' => options.push(IndexEntryOption::PageRangeBookmark(argument?)),
            't' => options.push(IndexEntryOption::CrossReference(argument?)),
            'y' => options.push(IndexEntryOption::Yomi(argument?)),
            _ => unknown_switches.push(MergeFieldSwitch { name, argument }),
        }
    }

    Some(IndexEntryParts {
        entry,
        options,
        unknown_switches,
    })
}

pub(in crate::parts::fields) fn parse_referenced_document_field_parts(
    instruction: &str,
) -> Option<ReferencedDocumentParts> {
    if instruction.len() > MAX_REFERENCED_DOCUMENT_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    if !keyword.eq_ignore_ascii_case("RD") {
        return None;
    }

    let source = next_field_argument(instruction, &mut position).ok()??;
    if source.is_empty() {
        return None;
    }

    let mut relative_path = false;
    let mut switches = Vec::new();
    loop {
        skip_field_whitespace(instruction, &mut position);
        let Some(introducer) = next_field_character(instruction, &mut position) else {
            break;
        };
        if introducer != '\\' || switches.len() >= MAX_REFERENCED_DOCUMENT_FIELD_SWITCHES {
            return None;
        }

        let name = next_field_character(instruction, &mut position)?;
        if name == '\\' || name.is_whitespace() {
            return None;
        }
        let name = name.to_ascii_lowercase();

        skip_field_whitespace(instruction, &mut position);
        let argument = match peek_field_character(instruction, position) {
            None | Some('\\') => None,
            Some(_) => next_field_argument(instruction, &mut position).ok()?,
        };
        if name == 'f' {
            if relative_path || argument.is_some() {
                return None;
            }
            relative_path = true;
        }
        switches.push(MergeFieldSwitch { name, argument });
    }

    Some(ReferencedDocumentParts {
        source,
        relative_path,
        switches,
    })
}

pub(in crate::parts::fields) fn private_field_opaque_instructions(
    instruction: &str,
) -> Option<String> {
    if instruction.len() > MAX_PRIVATE_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let instruction = instruction.trim_start();
    let keyword = instruction.get(.."PRIVATE".len())?;
    if !keyword.eq_ignore_ascii_case("PRIVATE") {
        return None;
    }
    let remainder = instruction.get("PRIVATE".len()..)?;
    match remainder.chars().next() {
        None | Some('"') | Some('\\') => Some(remainder.trim().to_string()),
        Some(character) if character.is_whitespace() => Some(remainder.trim().to_string()),
        Some(_) => None,
    }
}

pub(in crate::parts::fields) fn parse_table_of_authorities_field_parts(
    instruction: &str,
) -> Option<TableOfAuthoritiesParts> {
    if instruction.len() > MAX_TABLE_OF_AUTHORITIES_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    if !keyword.eq_ignore_ascii_case("TOA") {
        return None;
    }

    let mut options = Vec::new();
    let mut unknown_switches = Vec::new();
    loop {
        skip_field_whitespace(instruction, &mut position);
        let Some(introducer) = next_field_character(instruction, &mut position) else {
            break;
        };
        if introducer != '\\'
            || options.len() + unknown_switches.len() >= MAX_TABLE_OF_AUTHORITIES_FIELD_SWITCHES
        {
            return None;
        }

        let name = next_field_character(instruction, &mut position)?;
        if name == '\\' || name.is_whitespace() {
            return None;
        }
        let name = name.to_ascii_lowercase();

        skip_field_whitespace(instruction, &mut position);
        let argument = match peek_field_character(instruction, position) {
            None | Some('\\') => None,
            Some(_) => next_field_argument(instruction, &mut position).ok()?,
        };
        match name {
            'b' => options.push(TableOfAuthoritiesOption::Bookmark(argument.clone()?)),
            'c' => options.push(TableOfAuthoritiesOption::Category(argument.clone()?)),
            'd' => options.push(TableOfAuthoritiesOption::SequencePageSeparator(
                argument.clone()?,
            )),
            'e' => options.push(TableOfAuthoritiesOption::EntryPageNumberSeparator(
                argument.clone()?,
            )),
            'f' => {
                if argument.is_some() {
                    return None;
                }
                options.push(TableOfAuthoritiesOption::EntryFormatting);
            },
            'g' => options.push(TableOfAuthoritiesOption::PageRangeSeparator(
                argument.clone()?,
            )),
            'h' => {
                if argument.is_some() {
                    return None;
                }
                options.push(TableOfAuthoritiesOption::CategoryHeadings);
            },
            'l' => options.push(TableOfAuthoritiesOption::PageReferenceSeparator(
                argument.clone()?,
            )),
            'p' => {
                if argument.is_some() {
                    return None;
                }
                options.push(TableOfAuthoritiesOption::UsePassim);
            },
            's' => options.push(TableOfAuthoritiesOption::SequenceIdentifier(
                argument.clone()?,
            )),
            _ => unknown_switches.push(MergeFieldSwitch { name, argument }),
        }
    }

    Some(TableOfAuthoritiesParts {
        options,
        unknown_switches,
    })
}

pub(in crate::parts::fields) fn parse_reference_field_parts(
    instruction: &str,
    kind: ReferenceFieldKind,
) -> Option<ReferenceParts> {
    if instruction.len() > MAX_REFERENCE_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let bookmark = if kind == ReferenceFieldKind::ReferenceWithoutKeyword {
        next_field_argument(instruction, &mut position).ok()??
    } else {
        let keyword = next_field_argument(instruction, &mut position).ok()??;
        let keyword_matches = match kind {
            ReferenceFieldKind::Reference => keyword.eq_ignore_ascii_case("REF"),
            ReferenceFieldKind::PageReference => keyword.eq_ignore_ascii_case("PAGEREF"),
            ReferenceFieldKind::FootnoteReference | ReferenceFieldKind::NoteReference => {
                keyword.eq_ignore_ascii_case("FTNREF") || keyword.eq_ignore_ascii_case("NOTEREF")
            },
            ReferenceFieldKind::ReferenceWithoutKeyword => false,
        };
        if !keyword_matches {
            return None;
        }
        next_field_argument(instruction, &mut position).ok()??
    };
    if bookmark.is_empty() {
        return None;
    }

    let mut options = Vec::new();
    let mut unknown_switches = Vec::new();
    let is_ref = matches!(
        kind,
        ReferenceFieldKind::Reference | ReferenceFieldKind::ReferenceWithoutKeyword
    );
    let is_note_reference = matches!(
        kind,
        ReferenceFieldKind::FootnoteReference | ReferenceFieldKind::NoteReference
    );
    loop {
        skip_field_whitespace(instruction, &mut position);
        let Some(introducer) = next_field_character(instruction, &mut position) else {
            break;
        };
        if introducer != '\\'
            || options.len() + unknown_switches.len() >= MAX_REFERENCE_FIELD_SWITCHES
        {
            return None;
        }

        let name = next_field_character(instruction, &mut position)?;
        if name == '\\' || name.is_whitespace() {
            return None;
        }
        let name = name.to_ascii_lowercase();

        skip_field_whitespace(instruction, &mut position);
        let argument = match peek_field_character(instruction, position) {
            None | Some('\\') => None,
            Some(_) => next_field_argument(instruction, &mut position).ok()?,
        };
        match name {
            'd' if is_ref => {
                options.push(ReferenceFieldOption::SequencePageSeparator(
                    argument.clone()?,
                ));
            },
            'f' if is_ref => {
                if argument.is_some() {
                    return None;
                }
                options.push(ReferenceFieldOption::ReferencedNoteContent);
            },
            'f' if is_note_reference => {
                if argument.is_some() {
                    return None;
                }
                options.push(ReferenceFieldOption::NoteMarkFormatting);
            },
            'h' => {
                if argument.is_some() {
                    return None;
                }
                options.push(ReferenceFieldOption::Hyperlink);
            },
            'n' if is_ref => {
                if argument.is_some() {
                    return None;
                }
                options.push(ReferenceFieldOption::ParagraphNumberWithoutContext);
            },
            'p' => {
                if argument.is_some() {
                    return None;
                }
                options.push(ReferenceFieldOption::RelativePosition);
            },
            'r' if is_ref => {
                if argument.is_some() {
                    return None;
                }
                options.push(ReferenceFieldOption::ParagraphNumberRelativeContext);
            },
            't' if is_ref => {
                if argument.is_some() {
                    return None;
                }
                options.push(ReferenceFieldOption::SuppressNonNumberText);
            },
            'w' if is_ref => {
                if argument.is_some() {
                    return None;
                }
                options.push(ReferenceFieldOption::ParagraphNumberFullContext);
            },
            _ => unknown_switches.push(MergeFieldSwitch { name, argument }),
        }
    }

    Some(ReferenceParts {
        bookmark,
        options,
        unknown_switches,
    })
}

pub(in crate::parts::fields) fn parse_set_field_parts(
    instruction: &str,
) -> Option<(String, String)> {
    if instruction.len() > MAX_SET_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    if !keyword.eq_ignore_ascii_case("SET") {
        return None;
    }

    let target_name = next_field_argument(instruction, &mut position).ok()??;
    if target_name.is_empty() {
        return None;
    }

    skip_field_whitespace(instruction, &mut position);
    let expression = instruction.get(position..)?;
    if expression.trim().is_empty() {
        return None;
    }

    Some((target_name, expression.to_string()))
}

pub(in crate::parts::fields) fn parse_formula_field_formula(
    instruction: &str,
) -> Option<Option<String>> {
    if instruction.len() > MAX_FORMULA_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let formula = instruction.trim().strip_prefix('=')?.trim();
    Some((!formula.is_empty()).then_some(formula.to_string()))
}

pub(in crate::parts::fields) fn parse_equation_field_expression(
    instruction: &str,
) -> Option<String> {
    if instruction.len() > MAX_EQUATION_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    if !keyword.eq_ignore_ascii_case("EQ") {
        return None;
    }
    Some(instruction.get(position..)?.trim().to_string())
}

pub(in crate::parts::fields) fn parse_hyperlink_field_parts(
    instruction: &str,
) -> Option<HyperlinkParts> {
    if instruction.len() > MAX_HYPERLINK_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    if !keyword.eq_ignore_ascii_case("HYPERLINK") {
        return None;
    }

    skip_field_whitespace(instruction, &mut position);
    let external_target = match peek_field_character(instruction, position) {
        None | Some('\\') => None,
        Some(_) => {
            let target = next_field_argument(instruction, &mut position).ok()??;
            if target.is_empty() {
                return None;
            }
            Some(target)
        },
    };

    let mut bookmark = None;
    let mut screen_tip = None;
    let mut target_frame = None;
    let mut appends_image_map_coordinates = false;
    let mut opens_new_window = false;
    let mut unknown_switches = Vec::new();
    let mut switch_count = 0;
    loop {
        skip_field_whitespace(instruction, &mut position);
        let Some(introducer) = next_field_character(instruction, &mut position) else {
            break;
        };
        if introducer != '\\' || switch_count >= MAX_HYPERLINK_FIELD_SWITCHES {
            return None;
        }
        switch_count += 1;

        let name = next_field_character(instruction, &mut position)?;
        if name == '\\' || name.is_whitespace() {
            return None;
        }
        let name = name.to_ascii_lowercase();

        skip_field_whitespace(instruction, &mut position);
        let argument = match peek_field_character(instruction, position) {
            None | Some('\\') => None,
            Some(_) => next_field_argument(instruction, &mut position).ok()?,
        };

        let slot = match name {
            'l' => &mut bookmark,
            'o' => &mut screen_tip,
            't' => &mut target_frame,
            'm' => {
                if appends_image_map_coordinates || argument.is_some() {
                    return None;
                }
                appends_image_map_coordinates = true;
                continue;
            },
            'n' => {
                if opens_new_window || argument.is_some() {
                    return None;
                }
                opens_new_window = true;
                continue;
            },
            _ => {
                unknown_switches.push(MergeFieldSwitch { name, argument });
                continue;
            },
        };
        let value = argument?;
        if value.is_empty() || slot.replace(value).is_some() {
            return None;
        }
    }

    if external_target.is_none() && bookmark.is_none() {
        return None;
    }

    Some(HyperlinkParts {
        external_target,
        bookmark,
        screen_tip,
        target_frame,
        appends_image_map_coordinates,
        opens_new_window,
        unknown_switches,
    })
}

pub(in crate::parts::fields) fn parse_print_field_instructions(
    instruction: &str,
) -> Option<String> {
    if instruction.len() > MAX_PRINT_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    if !keyword.eq_ignore_ascii_case("PRINT") {
        return None;
    }
    Some(instruction.get(position..)?.trim().to_string())
}

pub(in crate::parts::fields) fn parse_embed_field_instructions(
    instruction: &str,
) -> Option<String> {
    if instruction.len() > MAX_EMBED_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    if !keyword.eq_ignore_ascii_case("EMBED") {
        return None;
    }
    Some(instruction.get(position..)?.trim().to_string())
}

pub(in crate::parts::fields) fn parse_barcode_field_instructions(
    instruction: &str,
) -> Option<String> {
    if instruction.len() > MAX_BARCODE_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    if !keyword.eq_ignore_ascii_case("BARCODE") {
        return None;
    }
    Some(instruction.get(position..)?.trim().to_string())
}

pub(in crate::parts::fields) fn parse_bidi_outline_field_instructions(
    instruction: &str,
) -> Option<String> {
    if instruction.len() > MAX_BIDI_OUTLINE_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    if !keyword.eq_ignore_ascii_case("BIDIOUTLINE") {
        return None;
    }
    Some(instruction.get(position..)?.trim().to_string())
}

pub(in crate::parts::fields) fn parse_shape_field_instructions(
    instruction: &str,
) -> Option<String> {
    if instruction.len() > MAX_SHAPE_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    if !keyword.eq_ignore_ascii_case("SHAPE") {
        return None;
    }
    Some(instruction.get(position..)?.trim().to_string())
}

pub(in crate::parts::fields) fn parse_legacy_form_field_instructions(
    instruction: &str,
    expected_keyword: &str,
) -> Option<String> {
    if instruction.len() > MAX_LEGACY_FORM_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    if !keyword.eq_ignore_ascii_case(expected_keyword) {
        return None;
    }
    Some(instruction.get(position..)?.trim().to_string())
}

pub(in crate::parts::fields) fn parse_quote_field_parts(
    instruction: &str,
) -> Option<(String, Vec<MergeFieldSwitch>)> {
    if instruction.len() > MAX_QUOTE_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    if !keyword.eq_ignore_ascii_case("QUOTE") {
        return None;
    }

    let text = next_field_argument(instruction, &mut position).ok()??;
    let mut switches = Vec::new();
    loop {
        skip_field_whitespace(instruction, &mut position);
        let Some(introducer) = next_field_character(instruction, &mut position) else {
            break;
        };
        if introducer != '\\' || switches.len() >= MAX_QUOTE_FIELD_SWITCHES {
            return None;
        }

        let name = next_field_character(instruction, &mut position)?;
        if name == '\\' || name.is_whitespace() {
            return None;
        }

        skip_field_whitespace(instruction, &mut position);
        let argument = match peek_field_character(instruction, position) {
            None | Some('\\') => None,
            Some(_) => next_field_argument(instruction, &mut position).ok()?,
        };
        switches.push(MergeFieldSwitch { name, argument });
    }

    Some((text, switches))
}

pub(in crate::parts::fields) fn parse_symbol_field_parts(
    instruction: &str,
) -> Option<(String, Vec<MergeFieldSwitch>)> {
    if instruction.len() > MAX_SYMBOL_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    if !keyword.eq_ignore_ascii_case("SYMBOL") {
        return None;
    }

    let character_argument = next_field_argument(instruction, &mut position).ok()??;
    let mut switches = Vec::new();
    loop {
        skip_field_whitespace(instruction, &mut position);
        let Some(introducer) = next_field_character(instruction, &mut position) else {
            break;
        };
        if introducer != '\\' || switches.len() >= MAX_SYMBOL_FIELD_SWITCHES {
            return None;
        }

        let name = next_field_character(instruction, &mut position)?;
        if name == '\\' || name.is_whitespace() {
            return None;
        }

        skip_field_whitespace(instruction, &mut position);
        let argument = match peek_field_character(instruction, position) {
            None | Some('\\') => None,
            Some(_) => next_field_argument(instruction, &mut position).ok()?,
        };
        switches.push(MergeFieldSwitch { name, argument });
    }

    Some((character_argument, switches))
}

pub(in crate::parts::fields) fn parse_auto_number_field_parts(
    instruction: &str,
) -> Option<(AutoNumberFieldKind, Vec<MergeFieldSwitch>)> {
    if instruction.len() > MAX_AUTO_NUMBER_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    let kind = AutoNumberFieldKind::from_keyword(&keyword)?;
    let mut switches = Vec::new();
    loop {
        skip_field_whitespace(instruction, &mut position);
        let Some(introducer) = next_field_character(instruction, &mut position) else {
            break;
        };
        if introducer != '\\' || switches.len() >= MAX_AUTO_NUMBER_FIELD_SWITCHES {
            return None;
        }

        let name = next_field_character(instruction, &mut position)?;
        if name == '\\' || name.is_whitespace() {
            return None;
        }

        skip_field_whitespace(instruction, &mut position);
        let argument = match peek_field_character(instruction, position) {
            None | Some('\\') => None,
            Some(_) => next_field_argument(instruction, &mut position).ok()?,
        };
        switches.push(MergeFieldSwitch { name, argument });
    }

    Some((kind, switches))
}

pub(in crate::parts::fields) fn parse_list_number_field_parts(
    instruction: &str,
) -> Option<(Option<String>, Vec<MergeFieldSwitch>)> {
    if instruction.len() > MAX_LIST_NUMBER_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    if !keyword.eq_ignore_ascii_case("LISTNUM") {
        return None;
    }

    skip_field_whitespace(instruction, &mut position);
    let list_name = match peek_field_character(instruction, position) {
        None | Some('\\') => None,
        Some(_) => Some(next_field_argument(instruction, &mut position).ok()??),
    };
    let mut switches = Vec::new();
    loop {
        skip_field_whitespace(instruction, &mut position);
        let Some(introducer) = next_field_character(instruction, &mut position) else {
            break;
        };
        if introducer != '\\' || switches.len() >= MAX_LIST_NUMBER_FIELD_SWITCHES {
            return None;
        }

        let name = next_field_character(instruction, &mut position)?;
        if name == '\\' || name.is_whitespace() {
            return None;
        }

        skip_field_whitespace(instruction, &mut position);
        let argument = match peek_field_character(instruction, position) {
            None | Some('\\') => None,
            Some(_) => next_field_argument(instruction, &mut position).ok()?,
        };
        switches.push(MergeFieldSwitch { name, argument });
    }

    Some((list_name, switches))
}

pub(in crate::parts::fields) fn parse_sequence_field_parts(
    instruction: &str,
) -> Option<(String, Option<String>, String)> {
    if instruction.len() > MAX_SEQUENCE_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    if !keyword.eq_ignore_ascii_case("SEQ") {
        return None;
    }

    let identifier = next_field_argument(instruction, &mut position).ok()??;
    if identifier.is_empty() {
        return None;
    }

    skip_field_whitespace(instruction, &mut position);
    let bookmark = match peek_field_character(instruction, position) {
        None | Some('\\') => None,
        Some(_) => {
            let bookmark = next_field_argument(instruction, &mut position).ok()??;
            if bookmark.is_empty() {
                return None;
            }
            Some(bookmark)
        },
    };

    skip_field_whitespace(instruction, &mut position);
    let tail = instruction.get(position..)?.trim().to_string();
    Some((identifier, bookmark, tail))
}

pub(in crate::parts::fields) fn parse_style_reference_field_parts(
    instruction: &str,
) -> Option<StyleReferenceParts> {
    if instruction.len() > MAX_STYLE_REFERENCE_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    if !keyword.eq_ignore_ascii_case("STYLEREF") {
        return None;
    }

    let style_name = next_field_argument(instruction, &mut position).ok()??;
    if style_name.is_empty() {
        return None;
    }

    let mut options = Vec::new();
    let mut unknown_switches = Vec::new();
    loop {
        skip_field_whitespace(instruction, &mut position);
        let Some(introducer) = next_field_character(instruction, &mut position) else {
            break;
        };
        if introducer != '\\'
            || options.len() + unknown_switches.len() >= MAX_STYLE_REFERENCE_FIELD_SWITCHES
        {
            return None;
        }

        let name = next_field_character(instruction, &mut position)?;
        if name == '\\' || name.is_whitespace() {
            return None;
        }
        let name = name.to_ascii_lowercase();

        skip_field_whitespace(instruction, &mut position);
        let argument = match peek_field_character(instruction, position) {
            None | Some('\\') => None,
            Some(_) => next_field_argument(instruction, &mut position).ok()?,
        };
        match name {
            'l' => {
                if argument.is_some() {
                    return None;
                }
                options.push(StyleReferenceFieldOption::FollowingText);
            },
            'n' => {
                if argument.is_some() {
                    return None;
                }
                options.push(StyleReferenceFieldOption::ParagraphNumber);
            },
            'p' => {
                if argument.is_some() {
                    return None;
                }
                options.push(StyleReferenceFieldOption::RelativePosition);
            },
            'r' => {
                if argument.is_some() {
                    return None;
                }
                options.push(StyleReferenceFieldOption::ParagraphNumberRelativeContext);
            },
            't' => {
                if argument.is_some() {
                    return None;
                }
                options.push(StyleReferenceFieldOption::SuppressNonNumberText);
            },
            'w' => {
                if argument.is_some() {
                    return None;
                }
                options.push(StyleReferenceFieldOption::ParagraphNumberFullContext);
            },
            _ => unknown_switches.push(MergeFieldSwitch { name, argument }),
        }
    }

    Some(StyleReferenceParts {
        style_name,
        options,
        unknown_switches,
    })
}

pub(in crate::parts::fields) fn parse_index_field_parts(instruction: &str) -> Option<IndexParts> {
    if instruction.len() > MAX_INDEX_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    if !keyword.eq_ignore_ascii_case("INDEX") {
        return None;
    }

    let mut options = Vec::new();
    let mut unknown_switches = Vec::new();
    loop {
        skip_field_whitespace(instruction, &mut position);
        let Some(introducer) = next_field_character(instruction, &mut position) else {
            break;
        };
        if introducer != '\\' || options.len() + unknown_switches.len() >= MAX_INDEX_FIELD_SWITCHES
        {
            return None;
        }

        let name = next_field_character(instruction, &mut position)?;
        if name == '\\' || name.is_whitespace() {
            return None;
        }
        let name = name.to_ascii_lowercase();

        skip_field_whitespace(instruction, &mut position);
        let argument = match peek_field_character(instruction, position) {
            None | Some('\\') => None,
            Some(_) => next_field_argument(instruction, &mut position).ok()?,
        };
        match name {
            'b' => options.push(IndexOption::Bookmark(argument.clone()?)),
            'c' => options.push(IndexOption::Columns(argument.clone()?)),
            'd' => options.push(IndexOption::SequencePageSeparator(argument.clone()?)),
            'e' => options.push(IndexOption::EntryPageNumberSeparator(argument.clone()?)),
            'f' => options.push(IndexOption::EntryType(argument.clone()?)),
            'g' => options.push(IndexOption::PageRangeSeparator(argument.clone()?)),
            'h' => options.push(IndexOption::Heading(argument.clone()?)),
            'k' => options.push(IndexOption::CrossReferenceSeparator(argument.clone()?)),
            'l' => options.push(IndexOption::PageNumberSeparator(argument.clone()?)),
            'o' => options.push(IndexOption::EastAsianSortOrder(argument.clone()?)),
            'p' => options.push(IndexOption::LetterRange(argument.clone()?)),
            'r' => {
                if argument.is_some() {
                    return None;
                }
                options.push(IndexOption::RunIn);
            },
            's' => options.push(IndexOption::SequenceIdentifier(argument.clone()?)),
            'y' => {
                if argument.is_some() {
                    return None;
                }
                options.push(IndexOption::UseYomi);
            },
            'z' => options.push(IndexOption::LanguageId(argument.clone()?)),
            _ => unknown_switches.push(MergeFieldSwitch { name, argument }),
        }
    }

    Some(IndexParts {
        options,
        unknown_switches,
    })
}

pub(in crate::parts::fields) fn parse_auto_text_field_parts(
    instruction: &str,
) -> Option<AutoTextParts> {
    if instruction.len() > MAX_AUTO_TEXT_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    if !keyword.eq_ignore_ascii_case("GLOSSARY") && !keyword.eq_ignore_ascii_case("AUTOTEXT") {
        return None;
    }
    let entry_name = next_field_argument(instruction, &mut position).ok()??;
    if entry_name.is_empty() {
        return None;
    }

    let mut unknown_switches = Vec::new();
    loop {
        skip_field_whitespace(instruction, &mut position);
        let Some(introducer) = next_field_character(instruction, &mut position) else {
            break;
        };
        if introducer != '\\' || unknown_switches.len() >= MAX_AUTO_TEXT_FIELD_SWITCHES {
            return None;
        }

        let name = next_field_character(instruction, &mut position)?;
        if name == '\\' || name.is_whitespace() {
            return None;
        }
        let name = name.to_ascii_lowercase();

        skip_field_whitespace(instruction, &mut position);
        let argument = match peek_field_character(instruction, position) {
            None | Some('\\') => None,
            Some(_) => next_field_argument(instruction, &mut position).ok()?,
        };
        unknown_switches.push(MergeFieldSwitch { name, argument });
    }

    Some(AutoTextParts {
        entry_name,
        unknown_switches,
    })
}

pub(in crate::parts::fields) fn parse_auto_text_list_field_parts(
    instruction: &str,
) -> Option<AutoTextListParts> {
    if instruction.len() > MAX_AUTO_TEXT_LIST_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    if !keyword.eq_ignore_ascii_case("AUTOTEXTLIST") {
        return None;
    }

    skip_field_whitespace(instruction, &mut position);
    let display_text = match peek_field_character(instruction, position) {
        None | Some('\\') => None,
        Some(_) => Some(next_field_argument(instruction, &mut position).ok()??),
    };

    let mut options = Vec::new();
    let mut unknown_switches = Vec::new();
    loop {
        skip_field_whitespace(instruction, &mut position);
        let Some(introducer) = next_field_character(instruction, &mut position) else {
            break;
        };
        if introducer != '\\'
            || options.len() + unknown_switches.len() >= MAX_AUTO_TEXT_LIST_FIELD_SWITCHES
        {
            return None;
        }

        let name = next_field_character(instruction, &mut position)?;
        if name == '\\' || name.is_whitespace() {
            return None;
        }
        let name = name.to_ascii_lowercase();

        skip_field_whitespace(instruction, &mut position);
        let argument = match peek_field_character(instruction, position) {
            None | Some('\\') => None,
            Some(_) => next_field_argument(instruction, &mut position).ok()?,
        };
        match name {
            's' => options.push(AutoTextListOption::Style(argument.clone()?)),
            't' => options.push(AutoTextListOption::Tip(argument.clone()?)),
            _ => unknown_switches.push(MergeFieldSwitch { name, argument }),
        }
    }

    Some(AutoTextListParts {
        display_text,
        options,
        unknown_switches,
    })
}

pub(in crate::parts::fields) fn parse_document_variable_field_parts(
    instruction: &str,
) -> Option<(String, Vec<MergeFieldSwitch>)> {
    if instruction.len() > MAX_DOCUMENT_VARIABLE_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    if !keyword.eq_ignore_ascii_case("DOCVARIABLE") {
        return None;
    }

    let variable_name = next_field_argument(instruction, &mut position).ok()??;
    if variable_name.is_empty() {
        return None;
    }

    let mut unknown_switches = Vec::new();
    loop {
        skip_field_whitespace(instruction, &mut position);
        let Some(introducer) = next_field_character(instruction, &mut position) else {
            break;
        };
        if introducer != '\\' || unknown_switches.len() >= MAX_DOCUMENT_VARIABLE_FIELD_SWITCHES {
            return None;
        }

        let name = next_field_character(instruction, &mut position)?;
        if name == '\\' || name.is_whitespace() {
            return None;
        }

        skip_field_whitespace(instruction, &mut position);
        let argument = match peek_field_character(instruction, position) {
            None | Some('\\') => None,
            Some(_) => next_field_argument(instruction, &mut position).ok()?,
        };
        unknown_switches.push(MergeFieldSwitch {
            name: name.to_ascii_lowercase(),
            argument,
        });
    }

    Some((variable_name, unknown_switches))
}

pub(in crate::parts::fields) fn parse_document_property_field_parts(
    instruction: &str,
) -> Option<(String, Vec<MergeFieldSwitch>)> {
    if instruction.len() > MAX_DOCUMENT_PROPERTY_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    if !keyword.eq_ignore_ascii_case("DOCPROPERTY") {
        return None;
    }

    let property_name = next_field_argument(instruction, &mut position).ok()??;
    if property_name.is_empty() {
        return None;
    }

    let mut switches = Vec::new();
    loop {
        skip_field_whitespace(instruction, &mut position);
        let Some(introducer) = next_field_character(instruction, &mut position) else {
            break;
        };
        if introducer != '\\' || switches.len() >= MAX_DOCUMENT_PROPERTY_FIELD_SWITCHES {
            return None;
        }

        let name = next_field_character(instruction, &mut position)?;
        if name == '\\' || name.is_whitespace() {
            return None;
        }

        skip_field_whitespace(instruction, &mut position);
        let argument = match peek_field_character(instruction, position) {
            None | Some('\\') => None,
            Some(_) => next_field_argument(instruction, &mut position).ok()?,
        };
        switches.push(MergeFieldSwitch {
            name: name.to_ascii_lowercase(),
            argument,
        });
    }

    Some((property_name, switches))
}

pub(in crate::parts::fields) fn parse_info_field_parts(
    instruction: &str,
) -> Option<(String, Option<String>, Vec<MergeFieldSwitch>)> {
    if instruction.len() > MAX_INFO_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let first_argument = next_field_argument(instruction, &mut position).ok()??;
    let information_type = if first_argument.eq_ignore_ascii_case("INFO") {
        next_field_argument(instruction, &mut position).ok()??
    } else {
        first_argument
    };
    if information_type.is_empty() {
        return None;
    }

    skip_field_whitespace(instruction, &mut position);
    let new_value = match peek_field_character(instruction, position) {
        None | Some('\\') => None,
        Some(_) => Some(next_field_argument(instruction, &mut position).ok()??),
    };

    let mut switches = Vec::new();
    loop {
        skip_field_whitespace(instruction, &mut position);
        let Some(introducer) = next_field_character(instruction, &mut position) else {
            break;
        };
        if introducer != '\\' || switches.len() >= MAX_INFO_FIELD_SWITCHES {
            return None;
        }

        let name = next_field_character(instruction, &mut position)?;
        if name == '\\' || name.is_whitespace() {
            return None;
        }

        skip_field_whitespace(instruction, &mut position);
        let argument = match peek_field_character(instruction, position) {
            None | Some('\\') => None,
            Some(_) => next_field_argument(instruction, &mut position).ok()?,
        };
        switches.push(MergeFieldSwitch {
            name: name.to_ascii_lowercase(),
            argument,
        });
    }

    Some((information_type, new_value, switches))
}

pub(in crate::parts::fields) fn parse_document_information_field_parts(
    instruction: &str,
) -> Option<(DocumentInformationFieldKind, Vec<MergeFieldSwitch>)> {
    if instruction.len() > MAX_DOCUMENT_INFORMATION_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    let kind = DocumentInformationFieldKind::from_keyword(&keyword)?;

    let mut switches = Vec::new();
    loop {
        skip_field_whitespace(instruction, &mut position);
        let Some(introducer) = next_field_character(instruction, &mut position) else {
            break;
        };
        if introducer != '\\' || switches.len() >= MAX_DOCUMENT_INFORMATION_FIELD_SWITCHES {
            return None;
        }

        let name = next_field_character(instruction, &mut position)?;
        if name == '\\' || name.is_whitespace() {
            return None;
        }

        skip_field_whitespace(instruction, &mut position);
        let argument = match peek_field_character(instruction, position) {
            None | Some('\\') => None,
            Some(_) => next_field_argument(instruction, &mut position).ok()?,
        };
        switches.push(MergeFieldSwitch {
            name: name.to_ascii_lowercase(),
            argument,
        });
    }

    Some((kind, switches))
}

pub(in crate::parts::fields) fn parse_document_context_field_parts(
    instruction: &str,
) -> Option<(DocumentContextFieldKind, Vec<MergeFieldSwitch>)> {
    if instruction.len() > MAX_DOCUMENT_CONTEXT_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    let kind = DocumentContextFieldKind::from_keyword(&keyword)?;

    let mut switches = Vec::new();
    loop {
        skip_field_whitespace(instruction, &mut position);
        let Some(introducer) = next_field_character(instruction, &mut position) else {
            break;
        };
        if introducer != '\\' || switches.len() >= MAX_DOCUMENT_CONTEXT_FIELD_SWITCHES {
            return None;
        }

        let name = next_field_character(instruction, &mut position)?;
        if name == '\\' || name.is_whitespace() {
            return None;
        }

        skip_field_whitespace(instruction, &mut position);
        let argument = match peek_field_character(instruction, position) {
            None | Some('\\') => None,
            Some(_) => next_field_argument(instruction, &mut position).ok()?,
        };
        switches.push(MergeFieldSwitch {
            name: name.to_ascii_lowercase(),
            argument,
        });
    }

    Some((kind, switches))
}

pub(in crate::parts::fields) fn parse_dde_field_parts(instruction: &str) -> Option<DdeParts> {
    if instruction.len() > MAX_DDE_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    let kind = if keyword.eq_ignore_ascii_case("DDE") {
        DdeFieldKind::Dde
    } else if keyword.eq_ignore_ascii_case("DDEAUTO") {
        DdeFieldKind::DdeAuto
    } else {
        return None;
    };

    let application = next_field_argument(instruction, &mut position).ok()??;
    if application.is_empty() {
        return None;
    }
    let source = next_field_argument(instruction, &mut position).ok()??;
    if source.is_empty() {
        return None;
    }

    skip_field_whitespace(instruction, &mut position);
    let item = match peek_field_character(instruction, position) {
        None | Some('\\') => None,
        Some(_) => Some(next_field_argument(instruction, &mut position).ok()??),
    };

    let mut automatic_updates = kind == DdeFieldKind::DdeAuto;
    let mut saw_automatic_update = false;
    let mut representation = None;
    let mut omit_graphic_data = false;
    let mut unknown_switches = Vec::new();
    loop {
        skip_field_whitespace(instruction, &mut position);
        let Some(introducer) = next_field_character(instruction, &mut position) else {
            break;
        };
        if introducer != '\\' || unknown_switches.len() >= MAX_DDE_FIELD_SWITCHES {
            return None;
        }

        let name = next_field_character(instruction, &mut position)?;
        if name == '\\' || name.is_whitespace() {
            return None;
        }
        let name = name.to_ascii_lowercase();

        skip_field_whitespace(instruction, &mut position);
        let argument = match peek_field_character(instruction, position) {
            None | Some('\\') => None,
            Some(_) => next_field_argument(instruction, &mut position).ok()?,
        };
        match name {
            'a' if kind == DdeFieldKind::Dde => {
                if saw_automatic_update || argument.is_some() {
                    return None;
                }
                automatic_updates = true;
                saw_automatic_update = true;
            },
            'a' => return None,
            'd' => {
                if representation.is_some() || omit_graphic_data || argument.is_some() {
                    return None;
                }
                omit_graphic_data = true;
            },
            'b' | 'h' | 'p' | 'r' | 't' | 'u' => {
                if representation.is_some() || omit_graphic_data || argument.is_some() {
                    return None;
                }
                representation = Some(match name {
                    'b' => DdeRepresentation::Bitmap,
                    'h' => DdeRepresentation::Html,
                    'p' => DdeRepresentation::Picture,
                    'r' => DdeRepresentation::RichText,
                    't' => DdeRepresentation::Text,
                    'u' => DdeRepresentation::UnicodeText,
                    _ => unreachable!("DDE representation switch was matched above"),
                });
            },
            _ => unknown_switches.push(MergeFieldSwitch { name, argument }),
        }
    }

    Some(DdeParts {
        kind,
        application,
        source,
        item,
        automatic_updates,
        representation,
        omit_graphic_data,
        unknown_switches,
    })
}

pub(in crate::parts::fields) fn parse_link_field_parts(instruction: &str) -> Option<LinkParts> {
    if instruction.len() > MAX_LINK_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    if !keyword.eq_ignore_ascii_case("LINK") {
        return None;
    }

    let application_type = next_field_argument(instruction, &mut position).ok()??;
    if application_type.is_empty() {
        return None;
    }
    let source = next_field_argument(instruction, &mut position).ok()??;
    if source.is_empty() {
        return None;
    }

    skip_field_whitespace(instruction, &mut position);
    let item = match peek_field_character(instruction, position) {
        None | Some('\\') => None,
        Some(_) => Some(next_field_argument(instruction, &mut position).ok()??),
    };

    let mut switches = Vec::new();
    loop {
        skip_field_whitespace(instruction, &mut position);
        let Some(introducer) = next_field_character(instruction, &mut position) else {
            break;
        };
        if introducer != '\\' || switches.len() >= MAX_LINK_FIELD_SWITCHES {
            return None;
        }

        let name = next_field_character(instruction, &mut position)?;
        if name == '\\' || name.is_whitespace() {
            return None;
        }
        let name = name.to_ascii_lowercase();

        skip_field_whitespace(instruction, &mut position);
        let argument = match peek_field_character(instruction, position) {
            None | Some('\\') => None,
            Some(_) => next_field_argument(instruction, &mut position).ok()?,
        };
        switches.push(MergeFieldSwitch { name, argument });
    }

    let mut automatic_updates = false;
    let mut result_options = Vec::new();
    let mut formatting_modes = Vec::new();
    for switch in &switches {
        match switch.name {
            'a' => {
                if switch.argument.is_some() {
                    return None;
                }
                automatic_updates = true;
            },
            'f' => {
                let value = switch.argument.as_deref()?.parse::<i64>().ok()?;
                formatting_modes.push(match value {
                    0 => LinkFormatting::Source,
                    2 => LinkFormatting::Destination,
                    4 => LinkFormatting::SpreadsheetSource,
                    5 => LinkFormatting::SpreadsheetDestination,
                    other => LinkFormatting::Unsupported(other),
                });
            },
            'b' | 'd' | 'h' | 'p' | 'r' | 't' | 'u' => {
                if switch.argument.is_some() {
                    return None;
                }
                result_options.push(match switch.name {
                    'b' => LinkResultOption::Bitmap,
                    'd' => LinkResultOption::OmitGraphicData,
                    'h' => LinkResultOption::Html,
                    'p' => LinkResultOption::Picture,
                    'r' => LinkResultOption::RichText,
                    't' => LinkResultOption::Text,
                    'u' => LinkResultOption::UnicodeText,
                    _ => unreachable!("LINK result switch was matched above"),
                });
            },
            _ => {},
        }
    }

    Some(LinkParts {
        application_type,
        source,
        item,
        automatic_updates,
        result_options,
        formatting_modes,
        switches,
    })
}

pub(in crate::parts::fields) fn parse_external_include_field_parts(
    instruction: &str,
) -> Option<ExternalIncludeParts> {
    if instruction.len() > MAX_EXTERNAL_INCLUDE_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    let kind =
        if keyword.eq_ignore_ascii_case("INCLUDETEXT") || keyword.eq_ignore_ascii_case("INCLUDE") {
            IncludeFieldKind::Text
        } else if keyword.eq_ignore_ascii_case("INCLUDEPICTURE")
            || keyword.eq_ignore_ascii_case("IMPORT")
        {
            IncludeFieldKind::Picture
        } else {
            return None;
        };

    let source = next_field_argument(instruction, &mut position).ok()??;
    if source.is_empty() {
        return None;
    }

    skip_field_whitespace(instruction, &mut position);
    let bookmark = match (kind, peek_field_character(instruction, position)) {
        (IncludeFieldKind::Text, None | Some('\\')) => None,
        (IncludeFieldKind::Text, Some(_)) => {
            Some(next_field_argument(instruction, &mut position).ok()??)
        },
        (IncludeFieldKind::Picture, _) => None,
    };

    let mut suppress_nested_field_updates = false;
    let mut omit_picture_data = false;
    let mut options = Vec::new();
    let mut unknown_switches = Vec::new();
    let mut switch_count = 0;
    loop {
        skip_field_whitespace(instruction, &mut position);
        let Some(introducer) = next_field_character(instruction, &mut position) else {
            break;
        };
        if introducer != '\\' || switch_count >= MAX_EXTERNAL_INCLUDE_FIELD_SWITCHES {
            return None;
        }
        switch_count += 1;

        let name = next_field_character(instruction, &mut position)?;
        if name == '\\' || name.is_whitespace() {
            return None;
        }
        let name = name.to_ascii_lowercase();

        skip_field_whitespace(instruction, &mut position);
        let argument = match peek_field_character(instruction, position) {
            None | Some('\\') => None,
            Some(_) => next_field_argument(instruction, &mut position).ok()?,
        };

        match (kind, name) {
            (_, 'c') => options.push(ExternalIncludeOption::Converter(argument?)),
            (IncludeFieldKind::Text, 'e') => {
                options.push(ExternalIncludeOption::Encoding(argument?));
            },
            (IncludeFieldKind::Text, 'm') => {
                options.push(ExternalIncludeOption::MimeType(argument?));
            },
            (IncludeFieldKind::Text, 'n') => {
                options.push(ExternalIncludeOption::NamespaceMapping(argument?));
            },
            (IncludeFieldKind::Text, 't') => {
                options.push(ExternalIncludeOption::Xslt(argument?));
            },
            (IncludeFieldKind::Text, 'x') => {
                options.push(ExternalIncludeOption::XPath(argument?));
            },
            (IncludeFieldKind::Text, '!') => {
                if suppress_nested_field_updates || argument.is_some() {
                    return None;
                }
                suppress_nested_field_updates = true;
            },
            (IncludeFieldKind::Picture, 'd') => {
                if omit_picture_data || argument.is_some() {
                    return None;
                }
                omit_picture_data = true;
            },
            _ => unknown_switches.push(MergeFieldSwitch { name, argument }),
        }
    }

    Some(ExternalIncludeParts {
        kind,
        source,
        bookmark,
        suppress_nested_field_updates,
        omit_picture_data,
        options,
        unknown_switches,
    })
}

pub(in crate::parts::fields) fn parse_mail_merge_counter_kind(
    instruction: &str,
) -> Option<MailMergeCounterKind> {
    if instruction.len() > MAX_MAIL_MERGE_COUNTER_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    let kind = if keyword.eq_ignore_ascii_case("MERGEREC") {
        MailMergeCounterKind::Record
    } else if keyword.eq_ignore_ascii_case("MERGESEQ") {
        MailMergeCounterKind::Sequence
    } else {
        return None;
    };
    if next_field_argument(instruction, &mut position)
        .ok()?
        .is_some()
    {
        return None;
    }

    Some(kind)
}

pub(in crate::parts::fields) fn is_mail_merge_next_instruction(instruction: &str) -> bool {
    if instruction.len() > MAX_MAIL_MERGE_NEXT_INSTRUCTION_BYTES {
        return false;
    }

    let mut position = 0;
    let Ok(Some(keyword)) = next_field_argument(instruction, &mut position) else {
        return false;
    };
    keyword.eq_ignore_ascii_case("NEXT")
        && matches!(next_field_argument(instruction, &mut position), Ok(None))
}

pub(in crate::parts::fields) fn parse_mail_merge_conditional_control_parts(
    instruction: &str,
) -> Option<(MailMergeConditionalControlKind, String)> {
    if instruction.len() > MAX_MAIL_MERGE_CONDITIONAL_CONTROL_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    let kind = if keyword.eq_ignore_ascii_case("NEXTIF") {
        MailMergeConditionalControlKind::NextIf
    } else if keyword.eq_ignore_ascii_case("SKIPIF") {
        MailMergeConditionalControlKind::SkipIf
    } else {
        return None;
    };
    let comparison = instruction.get(position..)?.trim();
    (!comparison.is_empty()).then_some((kind, comparison.to_string()))
}

pub(in crate::parts::fields) fn parse_if_field_expression(instruction: &str) -> Option<String> {
    if instruction.len() > MAX_IF_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    if !keyword.eq_ignore_ascii_case("IF") {
        return None;
    }
    let expression = instruction.get(position..)?.trim();
    (!expression.is_empty()).then_some(expression.to_string())
}

pub(in crate::parts::fields) fn parse_compare_field_comparison(
    instruction: &str,
) -> Option<String> {
    if instruction.len() > MAX_COMPARE_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    if !keyword.eq_ignore_ascii_case("COMPARE") {
        return None;
    }
    let comparison = instruction.get(position..)?.trim();
    (!comparison.is_empty()).then_some(comparison.to_string())
}

#[allow(clippy::type_complexity)]
pub(in crate::parts::fields) fn parse_prompt_field_parts(
    instruction: &str,
) -> Option<(
    PromptFieldKind,
    Option<String>,
    Option<String>,
    Option<String>,
    bool,
)> {
    if instruction.len() > MAX_PROMPT_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    let kind = if keyword.eq_ignore_ascii_case("ASK") {
        PromptFieldKind::Ask
    } else if keyword.eq_ignore_ascii_case("FILLIN") {
        PromptFieldKind::FillIn
    } else {
        return None;
    };

    let (bookmark, prompt) = match kind {
        PromptFieldKind::Ask => {
            let bookmark = next_field_argument(instruction, &mut position).ok()??;
            if bookmark.is_empty() {
                return None;
            }
            let prompt = next_field_argument(instruction, &mut position).ok()??;
            (Some(bookmark), Some(prompt))
        },
        PromptFieldKind::FillIn => {
            skip_field_whitespace(instruction, &mut position);
            let prompt = match peek_field_character(instruction, position) {
                None | Some('\\') => None,
                Some(_) => next_field_argument(instruction, &mut position).ok()?,
            };
            (None, prompt)
        },
    };

    let mut default_response = None;
    let mut prompts_once_per_mail_merge = false;
    loop {
        skip_field_whitespace(instruction, &mut position);
        let Some(introducer) = next_field_character(instruction, &mut position) else {
            break;
        };
        if introducer != '\\' {
            return None;
        }

        let name = next_field_character(instruction, &mut position)?;
        if name == '\\' || name.is_whitespace() {
            return None;
        }
        match name.to_ascii_lowercase() {
            'd' => {
                if default_response.is_some() {
                    return None;
                }
                default_response = Some(next_field_argument(instruction, &mut position).ok()??);
            },
            'o' => {
                if prompts_once_per_mail_merge {
                    return None;
                }
                skip_field_whitespace(instruction, &mut position);
                if !matches!(
                    peek_field_character(instruction, position),
                    None | Some('\\')
                ) {
                    return None;
                }
                prompts_once_per_mail_merge = true;
            },
            _ => return None,
        }
    }

    Some((
        kind,
        bookmark,
        prompt,
        default_response,
        prompts_once_per_mail_merge,
    ))
}

pub(in crate::parts::fields) fn parse_user_identity_field_parts(
    instruction: &str,
) -> Option<(
    UserIdentityFieldKind,
    Option<String>,
    Option<UserIdentityFormatting>,
)> {
    if instruction.len() > MAX_USER_IDENTITY_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    let kind = if keyword.eq_ignore_ascii_case("USERADDRESS") {
        UserIdentityFieldKind::Address
    } else if keyword.eq_ignore_ascii_case("USERINITIALS") {
        UserIdentityFieldKind::Initials
    } else if keyword.eq_ignore_ascii_case("USERNAME") {
        UserIdentityFieldKind::Name
    } else {
        return None;
    };

    skip_field_whitespace(instruction, &mut position);
    let override_value = match peek_field_character(instruction, position) {
        None | Some('\\') => None,
        Some(_) => Some(next_field_argument(instruction, &mut position).ok()??),
    };

    let mut formatting = None;
    loop {
        skip_field_whitespace(instruction, &mut position);
        let Some(introducer) = next_field_character(instruction, &mut position) else {
            break;
        };
        if introducer != '\\' {
            return None;
        }
        let name = next_field_character(instruction, &mut position)?;
        if name != '*' || formatting.is_some() {
            return None;
        }
        let value = next_field_argument(instruction, &mut position).ok()??;
        formatting = Some(if value.eq_ignore_ascii_case("Caps") {
            UserIdentityFormatting::Caps
        } else if value.eq_ignore_ascii_case("FirstCap") {
            UserIdentityFormatting::FirstCap
        } else if value.eq_ignore_ascii_case("Lower") {
            UserIdentityFormatting::Lower
        } else if value.eq_ignore_ascii_case("Upper") {
            UserIdentityFormatting::Upper
        } else {
            return None;
        });
    }

    Some((kind, override_value, formatting))
}

pub(in crate::parts::fields) fn parse_advance_field_adjustments(
    instruction: &str,
) -> Option<Vec<AdvanceFieldAdjustment>> {
    if instruction.len() > MAX_ADVANCE_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    if !keyword.eq_ignore_ascii_case("ADVANCE") {
        return None;
    }

    let mut adjustments = Vec::new();
    loop {
        skip_field_whitespace(instruction, &mut position);
        let Some(introducer) = next_field_character(instruction, &mut position) else {
            break;
        };
        if introducer != '\\' {
            return None;
        }
        let name = next_field_character(instruction, &mut position)?;
        let operation = match name.to_ascii_lowercase() {
            'd' => AdvanceFieldOperation::Down,
            'l' => AdvanceFieldOperation::Left,
            'r' => AdvanceFieldOperation::Right,
            'u' => AdvanceFieldOperation::Up,
            'x' => AdvanceFieldOperation::HorizontalPosition,
            'y' => AdvanceFieldOperation::VerticalPosition,
            _ => return None,
        };
        if adjustments.len() >= MAX_ADVANCE_FIELD_ADJUSTMENTS {
            return None;
        }
        let points = next_field_argument(instruction, &mut position)
            .ok()??
            .parse::<i64>()
            .ok()?;
        adjustments.push(AdvanceFieldAdjustment { operation, points });
    }

    Some(adjustments)
}

#[allow(clippy::type_complexity)]
pub(in crate::parts::fields) fn parse_mail_merge_recipient_field_parts(
    instruction: &str,
) -> Option<(
    MailMergeRecipientFieldKind,
    Option<AddressBlockCountryInclusion>,
    bool,
    Vec<String>,
    Option<String>,
    Option<String>,
    Option<String>,
    Vec<MergeFieldSwitch>,
)> {
    if instruction.len() > MAX_MAIL_MERGE_RECIPIENT_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    let kind = if keyword.eq_ignore_ascii_case("ADDRESSBLOCK") {
        MailMergeRecipientFieldKind::AddressBlock
    } else if keyword.eq_ignore_ascii_case("GREETINGLINE") {
        MailMergeRecipientFieldKind::GreetingLine
    } else {
        return None;
    };

    let mut country_inclusion = None;
    let mut formats_using_recipient_country = false;
    let mut excluded_countries = Vec::new();
    let mut format_template = None;
    let mut language = None;
    let mut greeting_fallback_text = None;
    let mut unknown_switches = Vec::new();
    let mut switch_count = 0;
    loop {
        skip_field_whitespace(instruction, &mut position);
        let Some(introducer) = next_field_character(instruction, &mut position) else {
            break;
        };
        if introducer != '\\' || switch_count >= MAX_MAIL_MERGE_RECIPIENT_FIELD_SWITCHES {
            return None;
        }
        switch_count += 1;

        let name = next_field_character(instruction, &mut position)?;
        if name == '\\' || name.is_whitespace() {
            return None;
        }
        match (kind, name.to_ascii_lowercase()) {
            (MailMergeRecipientFieldKind::AddressBlock, 'c') => {
                if country_inclusion.is_some() {
                    return None;
                }
                let value = next_field_argument(instruction, &mut position).ok()??;
                country_inclusion = Some(match value.as_str() {
                    "0" => AddressBlockCountryInclusion::Omit,
                    "1" => AddressBlockCountryInclusion::Always,
                    "2" => AddressBlockCountryInclusion::UnlessExcluded,
                    _ => return None,
                });
            },
            (MailMergeRecipientFieldKind::AddressBlock, 'd') => {
                if formats_using_recipient_country {
                    return None;
                }
                skip_field_whitespace(instruction, &mut position);
                if !matches!(
                    peek_field_character(instruction, position),
                    None | Some('\\')
                ) {
                    return None;
                }
                formats_using_recipient_country = true;
            },
            (MailMergeRecipientFieldKind::AddressBlock, 'e') => {
                excluded_countries.push(next_field_argument(instruction, &mut position).ok()??);
            },
            (_, 'f') => {
                if format_template.is_some() {
                    return None;
                }
                format_template = Some(next_field_argument(instruction, &mut position).ok()??);
            },
            (_, 'l') => {
                if language.is_some() {
                    return None;
                }
                language = Some(next_field_argument(instruction, &mut position).ok()??);
            },
            (MailMergeRecipientFieldKind::GreetingLine, 'c' | 'e') => {
                if greeting_fallback_text.is_some() {
                    return None;
                }
                greeting_fallback_text =
                    Some(next_field_argument(instruction, &mut position).ok()??);
            },
            _ => {
                skip_field_whitespace(instruction, &mut position);
                let argument = match peek_field_character(instruction, position) {
                    None | Some('\\') => None,
                    Some(_) => next_field_argument(instruction, &mut position).ok()?,
                };
                unknown_switches.push(MergeFieldSwitch {
                    name: name.to_ascii_lowercase(),
                    argument,
                });
            },
        }
    }

    Some((
        kind,
        country_inclusion,
        formats_using_recipient_country,
        excluded_countries,
        format_template,
        language,
        greeting_fallback_text,
        unknown_switches,
    ))
}

fn next_field_argument(input: &str, position: &mut usize) -> ParseResult<Option<String>, ()> {
    skip_field_whitespace(input, position);
    let Some(first) = next_field_character(input, position) else {
        return Ok(None);
    };

    if first != '"' {
        *position -= first.len_utf8();
        let mut value = String::new();
        while let Some(character) = next_field_character(input, position) {
            if character.is_whitespace() || character == '"' {
                *position -= character.len_utf8();
                break;
            }
            if character == '\\' {
                let escaped = next_field_character(input, position).ok_or(())?;
                if !matches!(escaped, '"' | '\\') {
                    return Err(());
                }
                value.push(escaped);
            } else {
                value.push(character);
            }
        }
        return Ok(Some(value));
    }

    let mut value = String::new();
    loop {
        let character = next_field_character(input, position).ok_or(())?;
        match character {
            '"' => return Ok(Some(value)),
            '\\' => {
                let escaped = next_field_character(input, position).ok_or(())?;
                if !matches!(escaped, '"' | '\\') {
                    return Err(());
                }
                value.push(escaped);
            },
            _ => value.push(character),
        }
    }
}

fn skip_field_whitespace(input: &str, position: &mut usize) {
    while let Some(character) = input.get(*position..).and_then(|rest| rest.chars().next()) {
        if !character.is_whitespace() {
            break;
        }
        *position += character.len_utf8();
    }
}

fn next_field_character(input: &str, position: &mut usize) -> Option<char> {
    let character = input.get(*position..)?.chars().next()?;
    *position += character.len_utf8();
    Some(character)
}

fn peek_field_character(input: &str, position: usize) -> Option<char> {
    input.get(position..)?.chars().next()
}
