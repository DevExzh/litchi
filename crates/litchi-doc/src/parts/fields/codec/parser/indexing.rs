//! Bounded indexing field-instruction parsers.

use super::prelude::*;

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
