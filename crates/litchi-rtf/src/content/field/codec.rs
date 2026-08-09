use std::borrow::Cow;

use super::model::{
    AddressBlockCountryInclusion, AdvanceFieldAdjustment, AdvanceFieldOperation,
    AutoNumberFieldKind, AutoNumberFieldParts, AutoTextFieldKind, AutoTextListOption,
    BarcodeDisplayFieldKind, BibliographyOption, BibliographyParts, CitationOption, CitationParts,
    DdeFieldKind, DdeFieldParts, DdeRepresentation, DocumentContextFieldKind,
    DocumentContextFieldParts, DocumentInformationFieldKind, DocumentInformationFieldParts,
    DocumentPropertyFieldParts, DocumentVariableFieldParts, ExternalIncludeOption,
    ExternalIncludeParts, FieldCodeError, FieldCodeToken, FieldSwitch, HyperlinkCode,
    IncludeFieldKind, IndexEntryOption, IndexEntryParts, IndexOption, IndexParts, InfoFieldParts,
    LinkFieldParts, LinkFormatting, LinkResultOption, ListNumberFieldParts,
    MailMergeConditionalControlKind, MailMergeCounterKind, MailMergeDataFieldParts,
    MailMergeRecipientFieldKind, MailMergeRecipientFieldParts, MergeFieldParts, ParsedFieldCode,
    PromptFieldKind, PromptFieldParts, ReferenceCode, ReferencedDocumentFieldParts,
    StyleReferenceFieldOption, TableOfAuthoritiesEntryOption, TableOfAuthoritiesEntryParts,
    TableOfAuthoritiesOption, TableOfAuthoritiesParts, TableOfContentsEntryOption,
    TableOfContentsEntryParts, TableOfContentsOption, TableOfContentsParts, UserIdentityFieldKind,
    UserIdentityFieldParts, UserIdentityFormatting,
};
use super::{MAX_INSTRUCTION_LEN, MAX_TOKENS};

pub(super) fn equation_expression(raw_instruction: &str) -> Option<&str> {
    let instruction = raw_instruction.trim_start_matches(|value: char| value.is_ascii_whitespace());
    let keyword_len = instruction
        .find(|value: char| value.is_ascii_whitespace())
        .unwrap_or(instruction.len());
    let keyword = instruction.get(..keyword_len)?;
    let remainder = instruction.get(keyword_len..)?;
    keyword
        .eq_ignore_ascii_case("EQ")
        .then(|| remainder.trim_start_matches(|value: char| value.is_ascii_whitespace()))
}

pub(super) fn print_field_instructions(raw_instruction: &str) -> Option<&str> {
    if raw_instruction.len() > MAX_INSTRUCTION_LEN {
        return None;
    }
    let instruction = raw_instruction.trim_start_matches(|value: char| value.is_ascii_whitespace());
    let keyword_len = instruction
        .find(|value: char| value.is_ascii_whitespace())
        .unwrap_or(instruction.len());
    let keyword = instruction.get(..keyword_len)?;
    let remainder = instruction.get(keyword_len..)?;
    keyword
        .eq_ignore_ascii_case("PRINT")
        .then(|| remainder.trim())
}

pub(super) fn embed_field_instructions(raw_instruction: &str) -> Option<&str> {
    if raw_instruction.len() > MAX_INSTRUCTION_LEN {
        return None;
    }
    let instruction = raw_instruction.trim_start_matches(|value: char| value.is_ascii_whitespace());
    let keyword = instruction.get(.."EMBED".len())?;
    if !keyword.eq_ignore_ascii_case("EMBED") {
        return None;
    }
    let remainder = instruction.get("EMBED".len()..)?;
    match remainder.chars().next() {
        None | Some('"' | '\\') => Some(remainder.trim()),
        Some(value) if value.is_ascii_whitespace() => Some(remainder.trim()),
        Some(_) => None,
    }
}

pub(super) fn barcode_field_instructions(raw_instruction: &str) -> Option<&str> {
    if raw_instruction.len() > MAX_INSTRUCTION_LEN {
        return None;
    }
    let instruction = raw_instruction.trim_start_matches(|value: char| value.is_ascii_whitespace());
    let keyword = instruction.get(.."BARCODE".len())?;
    if !keyword.eq_ignore_ascii_case("BARCODE") {
        return None;
    }
    let remainder = instruction.get("BARCODE".len()..)?;
    match remainder.chars().next() {
        None | Some('"' | '\\') => Some(remainder.trim()),
        Some(value) if value.is_ascii_whitespace() => Some(remainder.trim()),
        Some(_) => None,
    }
}

#[allow(
    clippy::type_complexity,
    reason = "the returned tuple mirrors the field-instruction grammar"
)]
pub(super) fn barcode_display_field_parts(
    instruction: &str,
) -> Option<(
    BarcodeDisplayFieldKind,
    Cow<'_, str>,
    Cow<'_, str>,
    Vec<FieldSwitch<'_>>,
)> {
    let tokens = tokenize(instruction).ok()?;
    let keyword = tokens.first()?;
    let kind = if keyword.value.eq_ignore_ascii_case("DISPLAYBARCODE") {
        BarcodeDisplayFieldKind::DisplayBarcode
    } else if keyword.value.eq_ignore_ascii_case("MERGEBARCODE") {
        BarcodeDisplayFieldKind::MergeBarcode
    } else {
        return None;
    };
    let data_argument = tokens.get(1).filter(|token| is_field_operand(token))?;
    let barcode_type = tokens.get(2).filter(|token| is_field_operand(token))?;

    let mut switches = Vec::new();
    let mut index = 3;
    while index < tokens.len() {
        let name = switch_name(tokens.get(index)?)?;
        let value = tokens
            .get(index + 1)
            .filter(|token| is_field_operand(token));
        switches.push(FieldSwitch {
            name: Cow::Owned(name.to_string()),
            value: value.map(|token| token.value.clone()),
        });
        index += 1 + usize::from(value.is_some());
    }

    Some((
        kind,
        data_argument.value.clone(),
        barcode_type.value.clone(),
        switches,
    ))
}

pub(super) fn bidi_outline_field_instructions(raw_instruction: &str) -> Option<&str> {
    if raw_instruction.len() > MAX_INSTRUCTION_LEN {
        return None;
    }
    let instruction = raw_instruction.trim_start_matches(|value: char| value.is_ascii_whitespace());
    let keyword = instruction.get(.."BIDIOUTLINE".len())?;
    if !keyword.eq_ignore_ascii_case("BIDIOUTLINE") {
        return None;
    }
    let remainder = instruction.get("BIDIOUTLINE".len()..)?;
    match remainder.chars().next() {
        None | Some('"' | '\\') => Some(remainder.trim()),
        Some(value) if value.is_ascii_whitespace() => Some(remainder.trim()),
        Some(_) => None,
    }
}

pub(super) fn shape_field_instructions(raw_instruction: &str) -> Option<&str> {
    if raw_instruction.len() > MAX_INSTRUCTION_LEN {
        return None;
    }
    let instruction = raw_instruction.trim_start_matches(|value: char| value.is_ascii_whitespace());
    let keyword = instruction.get(.."SHAPE".len())?;
    if !keyword.eq_ignore_ascii_case("SHAPE") {
        return None;
    }
    let remainder = instruction.get("SHAPE".len()..)?;
    match remainder.chars().next() {
        None | Some('"' | '\\') => Some(remainder.trim()),
        Some(value) if value.is_ascii_whitespace() => Some(remainder.trim()),
        Some(_) => None,
    }
}

pub(super) fn legacy_form_field_instructions<'a>(
    raw_instruction: &'a str,
    expected_keyword: &str,
) -> Option<&'a str> {
    if raw_instruction.len() > MAX_INSTRUCTION_LEN {
        return None;
    }
    let instruction = raw_instruction.trim_start_matches(|value: char| value.is_ascii_whitespace());
    let keyword = instruction.get(..expected_keyword.len())?;
    if !keyword.eq_ignore_ascii_case(expected_keyword) {
        return None;
    }
    let remainder = instruction.get(expected_keyword.len()..)?;
    match remainder.chars().next() {
        None | Some('"' | '\\') => Some(remainder.trim()),
        Some(value) if value.is_ascii_whitespace() => Some(remainder.trim()),
        Some(_) => None,
    }
}

pub(super) fn private_field_instructions(raw_instruction: &str) -> Option<&str> {
    if raw_instruction.len() > MAX_INSTRUCTION_LEN {
        return None;
    }
    let instruction = raw_instruction.trim_start_matches(|value: char| value.is_ascii_whitespace());
    let keyword = instruction.get(.."PRIVATE".len())?;
    if !keyword.eq_ignore_ascii_case("PRIVATE") {
        return None;
    }
    let remainder = instruction.get("PRIVATE".len()..)?;
    match remainder.chars().next() {
        None | Some('"' | '\\') => Some(remainder.trim()),
        Some(value) if value.is_ascii_whitespace() => Some(remainder.trim()),
        Some(_) => None,
    }
}

pub(super) fn macro_button_parts(
    instruction: &str,
) -> Option<(Cow<'_, str>, Option<Cow<'_, str>>)> {
    let tokens = tokenize(instruction).ok()?;
    let (keyword, arguments) = tokens.split_first()?;
    if !keyword.value.eq_ignore_ascii_case("MACROBUTTON") {
        return None;
    }
    let (macro_name, display_tokens) = arguments.split_first()?;
    if macro_name.value.is_empty() {
        return None;
    }
    let display_text = match display_tokens {
        [] => None,
        [display] => Some(display.value.clone()),
        displays => Some(Cow::Owned(
            displays
                .iter()
                .map(|token| token.value.as_ref())
                .collect::<Vec<_>>()
                .join(" "),
        )),
    };
    Some((macro_name.value.clone(), display_text))
}

pub(super) fn go_to_button_parts(instruction: &str) -> Option<(Cow<'_, str>, Cow<'_, str>)> {
    let tokens = tokenize(instruction).ok()?;
    let [keyword, target, button_text] = tokens.as_slice() else {
        return None;
    };
    if !keyword.value.eq_ignore_ascii_case("GOTOBUTTON")
        || target.value.is_empty()
        || button_text.value.is_empty()
        || switch_name(target).is_some()
        || switch_name(button_text).is_some()
    {
        return None;
    }
    Some((target.value.clone(), button_text.value.clone()))
}

pub(super) fn auto_text_field_parts(
    instruction: &str,
) -> Option<(AutoTextFieldKind, Cow<'_, str>, Vec<FieldSwitch<'_>>)> {
    let tokens = tokenize(instruction).ok()?;
    let keyword = tokens.first()?;
    let kind = if keyword.value.eq_ignore_ascii_case("GLOSSARY") {
        AutoTextFieldKind::Glossary
    } else if keyword.value.eq_ignore_ascii_case("AUTOTEXT") {
        AutoTextFieldKind::AutoText
    } else {
        return None;
    };
    let entry_name = tokens.get(1)?.value.clone();
    if entry_name.is_empty() || switch_name(tokens.get(1)?).is_some() {
        return None;
    }

    let mut unknown_switches = Vec::new();
    let mut index = 2;
    while index < tokens.len() {
        let name = switch_name(tokens.get(index)?)?;
        let value = tokens
            .get(index + 1)
            .filter(|next| switch_name(next).is_none());
        unknown_switches.push(FieldSwitch {
            name: Cow::Owned(name.to_string()),
            value: value.map(|token| token.value.clone()),
        });
        index += 1 + usize::from(value.is_some());
    }
    Some((kind, entry_name, unknown_switches))
}

#[allow(
    clippy::type_complexity,
    reason = "the returned tuple mirrors the field-instruction grammar"
)]
pub(super) fn auto_text_list_field_parts(
    instruction: &str,
) -> Option<(
    Option<Cow<'_, str>>,
    Vec<AutoTextListOption<'_>>,
    Vec<FieldSwitch<'_>>,
)> {
    let tokens = tokenize(instruction).ok()?;
    let keyword = tokens.first()?;
    if !keyword.value.eq_ignore_ascii_case("AUTOTEXTLIST") {
        return None;
    }

    let mut index = 1;
    let display_text = tokens
        .get(index)
        .filter(|token| is_field_operand(token))
        .map(|token| {
            index += 1;
            token.value.clone()
        });
    let mut options = Vec::new();
    let mut unknown_switches = Vec::new();
    while index < tokens.len() {
        let name = switch_name(tokens.get(index)?)?;
        let value = tokens.get(index + 1).filter(|next| is_field_operand(next));
        match name.to_ascii_lowercase().as_str() {
            "s" => {
                options.push(AutoTextListOption::Style(value?.value.clone()));
                index += 2;
            },
            "t" => {
                options.push(AutoTextListOption::Tip(value?.value.clone()));
                index += 2;
            },
            _ => {
                unknown_switches.push(FieldSwitch {
                    name: Cow::Owned(name.to_string()),
                    value: value.map(|token| token.value.clone()),
                });
                index += 1 + usize::from(value.is_some());
            },
        }
    }
    Some((display_text, options, unknown_switches))
}

pub(super) fn dde_field_parts(instruction: &str) -> Option<DdeFieldParts<'_>> {
    let mut tokens = tokenize(instruction).ok()?;
    let keyword = tokens.first()?;
    let kind = if keyword.value.eq_ignore_ascii_case("DDE") {
        DdeFieldKind::Dde
    } else if keyword.value.eq_ignore_ascii_case("DDEAUTO") {
        DdeFieldKind::DdeAuto
    } else {
        return None;
    };
    tokens.remove(0);

    let application = tokens.first()?.value.clone();
    if application.is_empty() || switch_name(tokens.first()?).is_some() {
        return None;
    }
    tokens.remove(0);

    let source = tokens.first()?.value.clone();
    if source.is_empty() || switch_name(tokens.first()?).is_some() {
        return None;
    }
    tokens.remove(0);

    let item = if tokens
        .first()
        .is_some_and(|token| switch_name(token).is_none())
    {
        Some(tokens.remove(0).value)
    } else {
        None
    };

    let mut automatic_updates = kind == DdeFieldKind::DdeAuto;
    let mut has_automatic_update_switch = false;
    let mut representation = None;
    let mut omit_graphic_data = false;
    let mut unknown_switches = Vec::new();
    let mut index = 0;
    while index < tokens.len() {
        let name = switch_name(tokens.get(index)?)?;
        let normalized_name = name.to_ascii_lowercase();
        match normalized_name.as_str() {
            "a" if kind == DdeFieldKind::Dde => {
                if has_automatic_update_switch
                    || tokens
                        .get(index + 1)
                        .is_some_and(|token| switch_name(token).is_none())
                {
                    return None;
                }
                automatic_updates = true;
                has_automatic_update_switch = true;
                index += 1;
            },
            "a" => return None,
            "d" => {
                if representation.is_some()
                    || omit_graphic_data
                    || tokens
                        .get(index + 1)
                        .is_some_and(|token| switch_name(token).is_none())
                {
                    return None;
                }
                omit_graphic_data = true;
                index += 1;
            },
            "b" | "h" | "p" | "r" | "t" | "u" => {
                if representation.is_some()
                    || omit_graphic_data
                    || tokens
                        .get(index + 1)
                        .is_some_and(|token| switch_name(token).is_none())
                {
                    return None;
                }
                representation = Some(match normalized_name.as_str() {
                    "b" => DdeRepresentation::Bitmap,
                    "h" => DdeRepresentation::Html,
                    "p" => DdeRepresentation::Picture,
                    "r" => DdeRepresentation::RichText,
                    "t" => DdeRepresentation::Text,
                    "u" => DdeRepresentation::UnicodeText,
                    _ => return None,
                });
                index += 1;
            },
            _ => {
                let value = tokens
                    .get(index + 1)
                    .filter(|token| switch_name(token).is_none());
                unknown_switches.push(FieldSwitch {
                    name: Cow::Owned(name.to_string()),
                    value: value.map(|token| token.value.clone()),
                });
                index += 1 + usize::from(value.is_some());
            },
        }
    }

    Some(DdeFieldParts {
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

pub(super) fn link_field_parts(instruction: &str) -> Option<LinkFieldParts<'_>> {
    let mut tokens = tokenize(instruction).ok()?;
    let keyword = tokens.first()?;
    if !keyword.value.eq_ignore_ascii_case("LINK") {
        return None;
    }
    tokens.remove(0);

    let application_type = tokens.first()?.value.clone();
    if application_type.is_empty() || switch_name(tokens.first()?).is_some() {
        return None;
    }
    tokens.remove(0);

    let source = tokens.first()?.value.clone();
    if source.is_empty() || switch_name(tokens.first()?).is_some() {
        return None;
    }
    tokens.remove(0);

    let item = if tokens
        .first()
        .is_some_and(|token| switch_name(token).is_none())
    {
        Some(tokens.remove(0).value)
    } else {
        None
    };

    let mut automatic_updates = false;
    let mut result_options = Vec::new();
    let mut formatting_modes = Vec::new();
    let mut unknown_switches = Vec::new();
    let mut index = 0;
    while index < tokens.len() {
        let name = switch_name(tokens.get(index)?)?;
        let normalized_name = name.to_ascii_lowercase();
        match normalized_name.as_str() {
            "a" => {
                if tokens
                    .get(index + 1)
                    .is_some_and(|token| switch_name(token).is_none())
                {
                    return None;
                }
                automatic_updates = true;
                index += 1;
            },
            "f" => {
                let value = switch_value(&tokens, index, name).ok()?;
                let parsed = value.parse::<i64>().ok()?;
                formatting_modes.push(match parsed {
                    0 => LinkFormatting::Source,
                    2 => LinkFormatting::Destination,
                    4 => LinkFormatting::SpreadsheetSource,
                    5 => LinkFormatting::SpreadsheetDestination,
                    other => LinkFormatting::Unsupported(other),
                });
                index += 2;
            },
            "b" | "d" | "h" | "p" | "r" | "t" | "u" => {
                if tokens
                    .get(index + 1)
                    .is_some_and(|token| switch_name(token).is_none())
                {
                    return None;
                }
                result_options.push(match normalized_name.as_str() {
                    "b" => LinkResultOption::Bitmap,
                    "d" => LinkResultOption::OmitGraphicData,
                    "h" => LinkResultOption::Html,
                    "p" => LinkResultOption::Picture,
                    "r" => LinkResultOption::RichText,
                    "t" => LinkResultOption::Text,
                    "u" => LinkResultOption::UnicodeText,
                    _ => return None,
                });
                index += 1;
            },
            _ => {
                let value = tokens
                    .get(index + 1)
                    .filter(|token| switch_name(token).is_none());
                unknown_switches.push(FieldSwitch {
                    name: Cow::Owned(name.to_string()),
                    value: value.map(|token| token.value.clone()),
                });
                index += 1 + usize::from(value.is_some());
            },
        }
    }

    Some(LinkFieldParts {
        application_type,
        source,
        item,
        automatic_updates,
        result_options,
        formatting_modes,
        unknown_switches,
    })
}

pub(super) fn external_include_parts(instruction: &str) -> Option<ExternalIncludeParts<'_>> {
    let mut tokens = tokenize(instruction).ok()?;
    let keyword = tokens.first()?;
    let kind = if keyword.value.eq_ignore_ascii_case("INCLUDETEXT")
        || keyword.value.eq_ignore_ascii_case("INCLUDE")
    {
        IncludeFieldKind::Text
    } else if keyword.value.eq_ignore_ascii_case("INCLUDEPICTURE")
        || keyword.value.eq_ignore_ascii_case("IMPORT")
    {
        IncludeFieldKind::Picture
    } else {
        return None;
    };
    tokens.remove(0);

    let source = tokens.first()?.value.clone();
    if source.is_empty() || switch_name(tokens.first()?).is_some() {
        return None;
    }
    tokens.remove(0);

    let bookmark = if kind == IncludeFieldKind::Text
        && tokens
            .first()
            .is_some_and(|token| switch_name(token).is_none())
    {
        Some(tokens.remove(0).value)
    } else {
        None
    };
    if kind == IncludeFieldKind::Picture
        && tokens
            .first()
            .is_some_and(|token| switch_name(token).is_none())
    {
        return None;
    }

    let mut options = Vec::new();
    let mut suppress_nested_field_updates = false;
    let mut omit_picture_data = false;
    let mut unknown_switches = Vec::new();
    let mut index = 0;
    while index < tokens.len() {
        let name = switch_name(tokens.get(index)?)?;
        if name.eq_ignore_ascii_case("c") {
            options.push(ExternalIncludeOption::Converter(
                switch_value(&tokens, index, name).ok()?,
            ));
            index += 2;
        } else if kind == IncludeFieldKind::Text && name.eq_ignore_ascii_case("e") {
            options.push(ExternalIncludeOption::Encoding(
                switch_value(&tokens, index, name).ok()?,
            ));
            index += 2;
        } else if kind == IncludeFieldKind::Text && name.eq_ignore_ascii_case("m") {
            options.push(ExternalIncludeOption::MimeType(
                switch_value(&tokens, index, name).ok()?,
            ));
            index += 2;
        } else if kind == IncludeFieldKind::Text && name.eq_ignore_ascii_case("n") {
            options.push(ExternalIncludeOption::NamespaceMapping(
                switch_value(&tokens, index, name).ok()?,
            ));
            index += 2;
        } else if kind == IncludeFieldKind::Text && name.eq_ignore_ascii_case("t") {
            options.push(ExternalIncludeOption::Xslt(
                switch_value(&tokens, index, name).ok()?,
            ));
            index += 2;
        } else if kind == IncludeFieldKind::Text && name.eq_ignore_ascii_case("x") {
            options.push(ExternalIncludeOption::XPath(
                switch_value(&tokens, index, name).ok()?,
            ));
            index += 2;
        } else if kind == IncludeFieldKind::Text && name == "!" {
            if suppress_nested_field_updates
                || tokens
                    .get(index + 1)
                    .is_some_and(|token| switch_name(token).is_none())
            {
                return None;
            }
            suppress_nested_field_updates = true;
            index += 1;
        } else if kind == IncludeFieldKind::Picture && name.eq_ignore_ascii_case("d") {
            if omit_picture_data
                || tokens
                    .get(index + 1)
                    .is_some_and(|token| switch_name(token).is_none())
            {
                return None;
            }
            omit_picture_data = true;
            index += 1;
        } else {
            let value = tokens
                .get(index + 1)
                .filter(|token| switch_name(token).is_none());
            unknown_switches.push(FieldSwitch {
                name: Cow::Owned(name.to_string()),
                value: value.map(|token| token.value.clone()),
            });
            index += 1 + usize::from(value.is_some());
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

pub(super) fn referenced_document_field_parts(
    instruction: &str,
) -> Option<ReferencedDocumentFieldParts<'_>> {
    let mut tokens = tokenize(instruction).ok()?;
    let keyword = tokens.first()?;
    if !keyword.value.eq_ignore_ascii_case("RD") {
        return None;
    }
    tokens.remove(0);

    let source = tokens.first()?.value.clone();
    if source.is_empty() || !is_field_operand(tokens.first()?) {
        return None;
    }
    tokens.remove(0);

    let mut relative_path = false;
    let mut switches = Vec::new();
    let mut index = 0;
    while index < tokens.len() {
        let name = switch_name(tokens.get(index)?)?;
        let value = tokens
            .get(index + 1)
            .filter(|token| is_field_operand(token));
        if name.eq_ignore_ascii_case("f") {
            if relative_path || value.is_some() {
                return None;
            }
            relative_path = true;
        }
        switches.push(FieldSwitch {
            name: Cow::Owned(name.to_string()),
            value: value.map(|token| token.value.clone()),
        });
        index += 1 + usize::from(value.is_some());
    }

    Some(ReferencedDocumentFieldParts {
        source,
        relative_path,
        switches,
    })
}

pub(super) fn table_of_contents_parts(instruction: &str) -> Option<TableOfContentsParts<'_>> {
    let mut tokens = tokenize(instruction).ok()?;
    let keyword = tokens.first()?;
    if !keyword.value.eq_ignore_ascii_case("TOC") {
        return None;
    }
    tokens.remove(0);

    let mut options = Vec::new();
    let mut unknown_switches = Vec::new();
    let mut index = 0;
    while index < tokens.len() {
        let name = switch_name(tokens.get(index)?)?;
        let normalized_name = name.to_ascii_lowercase();
        match normalized_name.as_str() {
            "a" => {
                options.push(TableOfContentsOption::CaptionWithoutLabel(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "b" => {
                options.push(TableOfContentsOption::Bookmark(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "c" => {
                options.push(TableOfContentsOption::CaptionSequence(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "d" => {
                options.push(TableOfContentsOption::SequencePageSeparator(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "f" => {
                options.push(TableOfContentsOption::TableEntryIdentifier(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "h" => {
                if tokens
                    .get(index + 1)
                    .is_some_and(|token| switch_name(token).is_none())
                {
                    return None;
                }
                options.push(TableOfContentsOption::Hyperlinks);
                index += 1;
            },
            "l" => {
                options.push(TableOfContentsOption::TableEntryLevels(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "n" => {
                let range = tokens
                    .get(index + 1)
                    .filter(|token| switch_name(token).is_none())
                    .map(|token| token.value.clone());
                options.push(TableOfContentsOption::OmitPageNumbers(range));
                index += 1 + usize::from(
                    tokens
                        .get(index + 1)
                        .is_some_and(|token| switch_name(token).is_none()),
                );
            },
            "o" => {
                let range = tokens
                    .get(index + 1)
                    .filter(|token| switch_name(token).is_none())
                    .map(|token| token.value.clone());
                options.push(TableOfContentsOption::HeadingStyleRange(range));
                index += 1 + usize::from(
                    tokens
                        .get(index + 1)
                        .is_some_and(|token| switch_name(token).is_none()),
                );
            },
            "p" => {
                options.push(TableOfContentsOption::EntryPageNumberSeparator(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "s" => {
                options.push(TableOfContentsOption::SequenceIdentifier(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "t" => {
                options.push(TableOfContentsOption::StyleMappings(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "u" => {
                if tokens
                    .get(index + 1)
                    .is_some_and(|token| switch_name(token).is_none())
                {
                    return None;
                }
                options.push(TableOfContentsOption::OutlineLevels);
                index += 1;
            },
            "w" => {
                if tokens
                    .get(index + 1)
                    .is_some_and(|token| switch_name(token).is_none())
                {
                    return None;
                }
                options.push(TableOfContentsOption::PreserveTabs);
                index += 1;
            },
            "x" => {
                if tokens
                    .get(index + 1)
                    .is_some_and(|token| switch_name(token).is_none())
                {
                    return None;
                }
                options.push(TableOfContentsOption::PreserveNewlines);
                index += 1;
            },
            "z" => {
                if tokens
                    .get(index + 1)
                    .is_some_and(|token| switch_name(token).is_none())
                {
                    return None;
                }
                options.push(TableOfContentsOption::HidePageNumbersInWebView);
                index += 1;
            },
            _ => {
                let value = tokens
                    .get(index + 1)
                    .filter(|token| switch_name(token).is_none());
                unknown_switches.push(FieldSwitch {
                    name: Cow::Owned(name.to_string()),
                    value: value.map(|token| token.value.clone()),
                });
                index += 1 + usize::from(value.is_some());
            },
        }
    }

    Some(TableOfContentsParts {
        options,
        unknown_switches,
    })
}

pub(super) fn table_of_contents_entry_parts(
    instruction: &str,
) -> Option<TableOfContentsEntryParts<'_>> {
    let mut tokens = tokenize(instruction).ok()?;
    let keyword = tokens.first()?;
    if !keyword.value.eq_ignore_ascii_case("TC") {
        return None;
    }
    tokens.remove(0);

    let entry = tokens.first()?.value.clone();
    if entry.is_empty() || switch_name(tokens.first()?).is_some() {
        return None;
    }
    tokens.remove(0);

    let mut options = Vec::new();
    let mut unknown_switches = Vec::new();
    let mut index = 0;
    while index < tokens.len() {
        let name = switch_name(tokens.get(index)?)?;
        let normalized_name = name.to_ascii_lowercase();
        match normalized_name.as_str() {
            "f" => {
                options.push(TableOfContentsEntryOption::ListIdentifier(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "l" => {
                options.push(TableOfContentsEntryOption::Level(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "n" => {
                if tokens
                    .get(index + 1)
                    .is_some_and(|token| switch_name(token).is_none())
                {
                    return None;
                }
                options.push(TableOfContentsEntryOption::OmitPageNumber);
                index += 1;
            },
            _ => {
                let value = tokens
                    .get(index + 1)
                    .filter(|token| switch_name(token).is_none());
                unknown_switches.push(FieldSwitch {
                    name: Cow::Owned(name.to_string()),
                    value: value.map(|token| token.value.clone()),
                });
                index += 1 + usize::from(value.is_some());
            },
        }
    }

    Some(TableOfContentsEntryParts {
        entry,
        options,
        unknown_switches,
    })
}

pub(super) fn table_of_authorities_entry_parts(
    instruction: &str,
) -> Option<TableOfAuthoritiesEntryParts<'_>> {
    let mut tokens = tokenize(instruction).ok()?;
    let keyword = tokens.first()?;
    if !keyword.value.eq_ignore_ascii_case("TA") {
        return None;
    }
    tokens.remove(0);

    let mut options = Vec::new();
    let mut unknown_switches = Vec::new();
    let mut index = 0;
    while index < tokens.len() {
        let name = switch_name(tokens.get(index)?)?;
        let normalized_name = name.to_ascii_lowercase();
        match normalized_name.as_str() {
            "b" => {
                if tokens
                    .get(index + 1)
                    .is_some_and(|token| switch_name(token).is_none())
                {
                    return None;
                }
                options.push(TableOfAuthoritiesEntryOption::BoldPageNumber);
                index += 1;
            },
            "c" => {
                options.push(TableOfAuthoritiesEntryOption::Category(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "i" => {
                if tokens
                    .get(index + 1)
                    .is_some_and(|token| switch_name(token).is_none())
                {
                    return None;
                }
                options.push(TableOfAuthoritiesEntryOption::ItalicPageNumber);
                index += 1;
            },
            "l" => {
                options.push(TableOfAuthoritiesEntryOption::LongCitation(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "r" => {
                options.push(TableOfAuthoritiesEntryOption::PageRangeBookmark(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "s" => {
                options.push(TableOfAuthoritiesEntryOption::ShortCitation(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            _ => {
                let value = tokens
                    .get(index + 1)
                    .filter(|token| switch_name(token).is_none());
                unknown_switches.push(FieldSwitch {
                    name: Cow::Owned(name.to_string()),
                    value: value.map(|token| token.value.clone()),
                });
                index += 1 + usize::from(value.is_some());
            },
        }
    }

    Some(TableOfAuthoritiesEntryParts {
        options,
        unknown_switches,
    })
}

pub(super) fn table_of_authorities_parts(instruction: &str) -> Option<TableOfAuthoritiesParts<'_>> {
    let mut tokens = tokenize(instruction).ok()?;
    let keyword = tokens.first()?;
    if !keyword.value.eq_ignore_ascii_case("TOA") {
        return None;
    }
    tokens.remove(0);

    let mut options = Vec::new();
    let mut unknown_switches = Vec::new();
    let mut index = 0;
    while index < tokens.len() {
        let name = switch_name(tokens.get(index)?)?;
        let normalized_name = name.to_ascii_lowercase();
        match normalized_name.as_str() {
            "b" => {
                options.push(TableOfAuthoritiesOption::Bookmark(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "c" => {
                options.push(TableOfAuthoritiesOption::Category(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "d" => {
                options.push(TableOfAuthoritiesOption::SequencePageSeparator(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "e" => {
                options.push(TableOfAuthoritiesOption::EntryPageNumberSeparator(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "f" => {
                if tokens
                    .get(index + 1)
                    .is_some_and(|token| switch_name(token).is_none())
                {
                    return None;
                }
                options.push(TableOfAuthoritiesOption::RemoveEntryFormatting);
                index += 1;
            },
            "g" => {
                options.push(TableOfAuthoritiesOption::PageRangeSeparator(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "h" => {
                if tokens
                    .get(index + 1)
                    .is_some_and(|token| switch_name(token).is_none())
                {
                    return None;
                }
                options.push(TableOfAuthoritiesOption::CategoryHeadings);
                index += 1;
            },
            "l" => {
                options.push(TableOfAuthoritiesOption::PageReferenceSeparator(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "p" => {
                if tokens
                    .get(index + 1)
                    .is_some_and(|token| switch_name(token).is_none())
                {
                    return None;
                }
                options.push(TableOfAuthoritiesOption::UsePassim);
                index += 1;
            },
            "s" => {
                options.push(TableOfAuthoritiesOption::SequenceIdentifier(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            _ => {
                let value = tokens
                    .get(index + 1)
                    .filter(|token| switch_name(token).is_none());
                unknown_switches.push(FieldSwitch {
                    name: Cow::Owned(name.to_string()),
                    value: value.map(|token| token.value.clone()),
                });
                index += 1 + usize::from(value.is_some());
            },
        }
    }

    Some(TableOfAuthoritiesParts {
        options,
        unknown_switches,
    })
}

pub(super) fn index_parts(instruction: &str) -> Option<IndexParts<'_>> {
    let mut tokens = tokenize(instruction).ok()?;
    let keyword = tokens.first()?;
    if !keyword.value.eq_ignore_ascii_case("INDEX") {
        return None;
    }
    tokens.remove(0);

    let mut options = Vec::new();
    let mut unknown_switches = Vec::new();
    let mut index = 0;
    while index < tokens.len() {
        let name = switch_name(tokens.get(index)?)?;
        let normalized_name = name.to_ascii_lowercase();
        match normalized_name.as_str() {
            "b" => {
                options.push(IndexOption::Bookmark(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "c" => {
                options.push(IndexOption::Columns(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "d" => {
                options.push(IndexOption::SequencePageSeparator(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "e" => {
                options.push(IndexOption::EntryPageNumberSeparator(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "f" => {
                options.push(IndexOption::EntryType(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "g" => {
                options.push(IndexOption::PageRangeSeparator(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "h" => {
                options.push(IndexOption::Heading(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "k" => {
                options.push(IndexOption::CrossReferenceSeparator(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "l" => {
                options.push(IndexOption::PageNumberSeparator(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "p" => {
                options.push(IndexOption::LetterRange(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "r" => {
                if tokens
                    .get(index + 1)
                    .is_some_and(|token| switch_name(token).is_none())
                {
                    return None;
                }
                options.push(IndexOption::RunIn);
                index += 1;
            },
            "s" => {
                options.push(IndexOption::SequenceIdentifier(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "y" => {
                if tokens
                    .get(index + 1)
                    .is_some_and(|token| switch_name(token).is_none())
                {
                    return None;
                }
                options.push(IndexOption::UseYomi);
                index += 1;
            },
            "z" => {
                options.push(IndexOption::LanguageId(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            _ => {
                let value = tokens
                    .get(index + 1)
                    .filter(|token| switch_name(token).is_none());
                unknown_switches.push(FieldSwitch {
                    name: Cow::Owned(name.to_string()),
                    value: value.map(|token| token.value.clone()),
                });
                index += 1 + usize::from(value.is_some());
            },
        }
    }

    Some(IndexParts {
        options,
        unknown_switches,
    })
}

pub(super) fn index_entry_parts(instruction: &str) -> Option<IndexEntryParts<'_>> {
    let mut tokens = tokenize(instruction).ok()?;
    let keyword = tokens.first()?;
    if !keyword.value.eq_ignore_ascii_case("XE") {
        return None;
    }
    tokens.remove(0);

    let entry = tokens.first()?.value.clone();
    if entry.is_empty() || switch_name(tokens.first()?).is_some() {
        return None;
    }
    tokens.remove(0);

    let mut options = Vec::new();
    let mut unknown_switches = Vec::new();
    let mut index = 0;
    while index < tokens.len() {
        let name = switch_name(tokens.get(index)?)?;
        let normalized_name = name.to_ascii_lowercase();
        match normalized_name.as_str() {
            "b" => {
                if tokens
                    .get(index + 1)
                    .is_some_and(|token| switch_name(token).is_none())
                {
                    return None;
                }
                options.push(IndexEntryOption::BoldPageNumber);
                index += 1;
            },
            "f" => {
                options.push(IndexEntryOption::EntryType(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "i" => {
                if tokens
                    .get(index + 1)
                    .is_some_and(|token| switch_name(token).is_none())
                {
                    return None;
                }
                options.push(IndexEntryOption::ItalicPageNumber);
                index += 1;
            },
            "r" => {
                options.push(IndexEntryOption::PageRangeBookmark(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "t" => {
                options.push(IndexEntryOption::CrossReference(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "y" => {
                options.push(IndexEntryOption::Yomi(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            _ => {
                let value = tokens
                    .get(index + 1)
                    .filter(|token| switch_name(token).is_none());
                unknown_switches.push(FieldSwitch {
                    name: Cow::Owned(name.to_string()),
                    value: value.map(|token| token.value.clone()),
                });
                index += 1 + usize::from(value.is_some());
            },
        }
    }

    Some(IndexEntryParts {
        entry,
        options,
        unknown_switches,
    })
}

pub(super) fn citation_parts(instruction: &str) -> Option<CitationParts<'_>> {
    let mut tokens = tokenize(instruction).ok()?;
    let keyword = tokens.first()?;
    if !keyword.value.eq_ignore_ascii_case("CITATION") {
        return None;
    }
    tokens.remove(0);

    let source_tag = tokens.first()?.value.clone();
    if source_tag.is_empty() || switch_name(tokens.first()?).is_some() {
        return None;
    }
    tokens.remove(0);

    let mut options = Vec::new();
    let mut unknown_switches = Vec::new();
    let mut index = 0;
    while index < tokens.len() {
        let name = switch_name(tokens.get(index)?)?;
        let normalized_name = name.to_ascii_lowercase();
        match normalized_name.as_str() {
            "l" => {
                options.push(CitationOption::LanguageId(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "f" => {
                options.push(CitationOption::Prefix(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "s" => {
                options.push(CitationOption::Suffix(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "p" => {
                options.push(CitationOption::PageNumber(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "v" => {
                options.push(CitationOption::VolumeNumber(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "n" | "t" | "y" => {
                if tokens
                    .get(index + 1)
                    .is_some_and(|token| switch_name(token).is_none())
                {
                    return None;
                }
                options.push(match normalized_name.as_str() {
                    "n" => CitationOption::SuppressAuthor,
                    "t" => CitationOption::SuppressTitle,
                    "y" => CitationOption::SuppressYear,
                    _ => return None,
                });
                index += 1;
            },
            "m" => {
                options.push(CitationOption::AdditionalSourceTag(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            _ => {
                let value = tokens
                    .get(index + 1)
                    .filter(|token| switch_name(token).is_none());
                unknown_switches.push(FieldSwitch {
                    name: Cow::Owned(name.to_string()),
                    value: value.map(|token| token.value.clone()),
                });
                index += 1 + usize::from(value.is_some());
            },
        }
    }

    Some(CitationParts {
        source_tag,
        options,
        unknown_switches,
    })
}

pub(super) fn bibliography_parts(instruction: &str) -> Option<BibliographyParts<'_>> {
    let mut tokens = tokenize(instruction).ok()?;
    let keyword = tokens.first()?;
    if !keyword.value.eq_ignore_ascii_case("BIBLIOGRAPHY") {
        return None;
    }
    tokens.remove(0);

    let mut options = Vec::new();
    let mut unknown_switches = Vec::new();
    let mut index = 0;
    while index < tokens.len() {
        let name = switch_name(tokens.get(index)?)?;
        let normalized_name = name.to_ascii_lowercase();
        match normalized_name.as_str() {
            "l" => {
                options.push(BibliographyOption::LanguageId(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "f" => {
                options.push(BibliographyOption::FilterLanguageId(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "m" => {
                options.push(BibliographyOption::SourceTag(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            _ => {
                let value = tokens
                    .get(index + 1)
                    .filter(|token| switch_name(token).is_none());
                unknown_switches.push(FieldSwitch {
                    name: Cow::Owned(name.to_string()),
                    value: value.map(|token| token.value.clone()),
                });
                index += 1 + usize::from(value.is_some());
            },
        }
    }

    Some(BibliographyParts {
        options,
        unknown_switches,
    })
}

pub(super) fn document_variable_field_parts(
    instruction: &str,
) -> Option<DocumentVariableFieldParts<'_>> {
    let mut tokens = tokenize(instruction).ok()?;
    let keyword = tokens.first()?;
    if !keyword.value.eq_ignore_ascii_case("DOCVARIABLE") {
        return None;
    }
    tokens.remove(0);

    let variable_name = tokens.first()?.value.clone();
    if variable_name.is_empty() || switch_name(tokens.first()?).is_some() {
        return None;
    }
    tokens.remove(0);

    let mut unknown_switches = Vec::new();
    let mut index = 0;
    while index < tokens.len() {
        let name = switch_name(tokens.get(index)?)?;
        let value = tokens
            .get(index + 1)
            .filter(|token| switch_name(token).is_none());
        unknown_switches.push(FieldSwitch {
            name: Cow::Owned(name.to_string()),
            value: value.map(|token| token.value.clone()),
        });
        index += 1 + usize::from(value.is_some());
    }

    Some(DocumentVariableFieldParts {
        variable_name,
        unknown_switches,
    })
}

pub(super) fn document_property_field_parts(
    instruction: &str,
) -> Option<DocumentPropertyFieldParts<'_>> {
    let mut tokens = tokenize(instruction).ok()?;
    let keyword = tokens.first()?;
    if !keyword.value.eq_ignore_ascii_case("DOCPROPERTY") {
        return None;
    }
    tokens.remove(0);

    let property_name = tokens.first()?.value.clone();
    if property_name.is_empty() || !is_field_operand(tokens.first()?) {
        return None;
    }
    tokens.remove(0);

    let mut switches = Vec::new();
    let mut index = 0;
    while index < tokens.len() {
        let name = switch_name(tokens.get(index)?)?;
        let value = tokens
            .get(index + 1)
            .filter(|token| is_field_operand(token));
        switches.push(FieldSwitch {
            name: Cow::Owned(name.to_string()),
            value: value.map(|token| token.value.clone()),
        });
        index += 1 + usize::from(value.is_some());
    }

    Some(DocumentPropertyFieldParts {
        property_name,
        switches,
    })
}

pub(super) fn info_field_parts(instruction: &str) -> Option<InfoFieldParts<'_>> {
    let mut tokens = tokenize(instruction).ok()?;
    let keyword = tokens.first()?;
    if !keyword.value.eq_ignore_ascii_case("INFO") {
        return None;
    }
    tokens.remove(0);

    let information_type = tokens.first()?.value.clone();
    if information_type.is_empty() || !is_field_operand(tokens.first()?) {
        return None;
    }
    tokens.remove(0);

    let new_value = tokens
        .first()
        .filter(|token| is_field_operand(token))
        .map(|token| token.value.clone());
    if new_value.is_some() {
        tokens.remove(0);
    }

    let mut switches = Vec::new();
    let mut index = 0;
    while index < tokens.len() {
        let name = switch_name(tokens.get(index)?)?;
        let value = tokens
            .get(index + 1)
            .filter(|token| is_field_operand(token));
        switches.push(FieldSwitch {
            name: Cow::Owned(name.to_string()),
            value: value.map(|token| token.value.clone()),
        });
        index += 1 + usize::from(value.is_some());
    }

    Some(InfoFieldParts {
        information_type,
        new_value,
        switches,
    })
}

pub(super) fn document_information_field_parts(
    instruction: &str,
) -> Option<DocumentInformationFieldParts<'_>> {
    let mut tokens = tokenize(instruction).ok()?;
    let keyword = tokens.first()?;
    let kind = DocumentInformationFieldKind::from_keyword(keyword.value.as_ref())?;
    tokens.remove(0);

    let mut switches = Vec::new();
    let mut index = 0;
    while index < tokens.len() {
        let name = switch_name(tokens.get(index)?)?;
        let value = tokens
            .get(index + 1)
            .filter(|token| is_field_operand(token));
        switches.push(FieldSwitch {
            name: Cow::Owned(name.to_string()),
            value: value.map(|token| token.value.clone()),
        });
        index += 1 + usize::from(value.is_some());
    }

    Some(DocumentInformationFieldParts { kind, switches })
}

pub(super) fn document_context_field_parts(
    instruction: &str,
) -> Option<DocumentContextFieldParts<'_>> {
    let mut tokens = tokenize(instruction).ok()?;
    let keyword = tokens.first()?;
    let kind = DocumentContextFieldKind::from_keyword(keyword.value.as_ref())?;
    tokens.remove(0);

    let mut switches = Vec::new();
    let mut index = 0;
    while index < tokens.len() {
        let name = switch_name(tokens.get(index)?)?;
        let value = tokens
            .get(index + 1)
            .filter(|token| is_field_operand(token));
        switches.push(FieldSwitch {
            name: Cow::Owned(name.to_string()),
            value: value.map(|token| token.value.clone()),
        });
        index += 1 + usize::from(value.is_some());
    }

    Some(DocumentContextFieldParts { kind, switches })
}

pub(super) fn merge_field_parts(instruction: &str) -> Option<MergeFieldParts<'_>> {
    let mut tokens = tokenize(instruction).ok()?;
    let keyword = tokens.first()?;
    if !keyword.value.eq_ignore_ascii_case("MERGEFIELD") {
        return None;
    }
    tokens.remove(0);

    let field_name = tokens.first()?.value.clone();
    if field_name.is_empty() || switch_name(tokens.first()?).is_some() {
        return None;
    }
    tokens.remove(0);

    let mut switches = Vec::new();
    let mut index = 0;
    while index < tokens.len() {
        let name = switch_name(tokens.get(index)?)?;
        let value = tokens
            .get(index + 1)
            .filter(|token| switch_name(token).is_none());
        switches.push(FieldSwitch {
            name: Cow::Owned(name.to_string()),
            value: value.map(|token| token.value.clone()),
        });
        index += 1 + usize::from(value.is_some());
    }

    Some(MergeFieldParts {
        field_name,
        switches,
    })
}

pub(super) fn database_field_instructions(raw_instruction: &str) -> Option<&str> {
    if raw_instruction.len() > MAX_INSTRUCTION_LEN {
        return None;
    }
    let instruction = raw_instruction.trim_start_matches(|value: char| value.is_ascii_whitespace());
    let keyword = instruction.get(.."DATABASE".len())?;
    if !keyword.eq_ignore_ascii_case("DATABASE") {
        return None;
    }
    let remainder = instruction.get("DATABASE".len()..)?;
    match remainder.chars().next() {
        None | Some('"' | '\\') => Some(remainder.trim()),
        Some(value) if value.is_ascii_whitespace() => Some(remainder.trim()),
        Some(_) => None,
    }
}

pub(super) fn mail_merge_data_field_parts(
    instruction: &str,
) -> Option<MailMergeDataFieldParts<'_>> {
    let mut tokens = tokenize(instruction).ok()?;
    let keyword = tokens.first()?;
    if !keyword.value.eq_ignore_ascii_case("DATA") {
        return None;
    }
    tokens.remove(0);

    let data_source = tokens.first()?.value.clone();
    if data_source.is_empty() || switch_name(tokens.first()?).is_some() {
        return None;
    }
    tokens.remove(0);

    let header_source = match tokens.first() {
        Some(token) if switch_name(token).is_none() && !token.value.is_empty() => {
            let header_source = token.value.clone();
            tokens.remove(0);
            Some(header_source)
        },
        Some(token) if switch_name(token).is_none() => return None,
        _ => None,
    };

    let mut switches = Vec::new();
    let mut index = 0;
    while index < tokens.len() {
        let name = switch_name(tokens.get(index)?)?;
        let value = tokens
            .get(index + 1)
            .filter(|token| switch_name(token).is_none());
        switches.push(FieldSwitch {
            name: Cow::Owned(name.to_string()),
            value: value.map(|token| token.value.clone()),
        });
        index += 1 + usize::from(value.is_some());
    }

    Some(MailMergeDataFieldParts {
        data_source,
        header_source,
        switches,
    })
}

pub(super) fn prompt_field_parts(instruction: &str) -> Option<PromptFieldParts<'_>> {
    let mut tokens = tokenize(instruction).ok()?;
    let keyword = tokens.first()?;
    let kind = if keyword.value.eq_ignore_ascii_case("ASK") {
        PromptFieldKind::Ask
    } else if keyword.value.eq_ignore_ascii_case("FILLIN") {
        PromptFieldKind::FillIn
    } else {
        return None;
    };
    tokens.remove(0);

    let (bookmark, prompt) = match kind {
        PromptFieldKind::Ask => {
            let bookmark_token = tokens.first()?;
            if bookmark_token.value.is_empty() || switch_name(bookmark_token).is_some() {
                return None;
            }
            let bookmark = tokens.remove(0).value;

            let prompt_token = tokens.first()?;
            if switch_name(prompt_token).is_some() {
                return None;
            }
            let prompt = tokens.remove(0).value;
            (Some(bookmark), Some(prompt))
        },
        PromptFieldKind::FillIn => {
            let prompt = if tokens
                .first()
                .is_some_and(|token| switch_name(token).is_none())
            {
                Some(tokens.remove(0).value)
            } else {
                None
            };
            (None, prompt)
        },
    };

    let mut default_response = None;
    let mut prompts_once_per_mail_merge = false;
    let mut index = 0;
    while index < tokens.len() {
        let name = switch_name(tokens.get(index)?)?;
        match name.to_ascii_lowercase().as_str() {
            "d" => {
                if default_response.is_some() {
                    return None;
                }
                let value = tokens
                    .get(index + 1)
                    .filter(|token| switch_name(token).is_none())?;
                default_response = Some(value.value.clone());
                index += 2;
            },
            "o" => {
                if prompts_once_per_mail_merge
                    || tokens
                        .get(index + 1)
                        .is_some_and(|token| switch_name(token).is_none())
                {
                    return None;
                }
                prompts_once_per_mail_merge = true;
                index += 1;
            },
            _ => return None,
        }
    }

    Some(PromptFieldParts {
        kind,
        bookmark,
        prompt,
        default_response,
        prompts_once_per_mail_merge,
    })
}

pub(super) fn user_identity_field_parts(instruction: &str) -> Option<UserIdentityFieldParts<'_>> {
    let mut tokens = tokenize(instruction).ok()?;
    let keyword = tokens.first()?;
    let kind = if keyword.value.eq_ignore_ascii_case("USERADDRESS") {
        UserIdentityFieldKind::Address
    } else if keyword.value.eq_ignore_ascii_case("USERINITIALS") {
        UserIdentityFieldKind::Initials
    } else if keyword.value.eq_ignore_ascii_case("USERNAME") {
        UserIdentityFieldKind::Name
    } else {
        return None;
    };
    tokens.remove(0);

    let override_value = if tokens
        .first()
        .is_some_and(|token| switch_name(token).is_none())
    {
        Some(tokens.remove(0).value)
    } else {
        None
    };

    let mut formatting = None;
    let mut index = 0;
    while index < tokens.len() {
        let name = switch_name(tokens.get(index)?)?;
        if name != "*" || formatting.is_some() {
            return None;
        }
        let value = tokens
            .get(index + 1)
            .filter(|token| switch_name(token).is_none())?
            .value
            .as_ref();
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
        index += 2;
    }

    Some(UserIdentityFieldParts {
        kind,
        override_value,
        formatting,
    })
}

pub(super) fn advance_field_adjustments(instruction: &str) -> Option<Vec<AdvanceFieldAdjustment>> {
    let mut tokens = tokenize(instruction).ok()?;
    let keyword = tokens.first()?;
    if !keyword.value.eq_ignore_ascii_case("ADVANCE") {
        return None;
    }
    tokens.remove(0);

    let mut adjustments = Vec::with_capacity(tokens.len() / 2);
    let mut index = 0;
    while index < tokens.len() {
        let name = switch_name(tokens.get(index)?)?;
        let operation = match name.to_ascii_lowercase().as_str() {
            "d" => AdvanceFieldOperation::Down,
            "l" => AdvanceFieldOperation::Left,
            "r" => AdvanceFieldOperation::Right,
            "u" => AdvanceFieldOperation::Up,
            "x" => AdvanceFieldOperation::HorizontalPosition,
            "y" => AdvanceFieldOperation::VerticalPosition,
            _ => return None,
        };
        let points = tokens
            .get(index + 1)
            .filter(|token| switch_name(token).is_none())?
            .value
            .parse::<i64>()
            .ok()?;
        adjustments.push(AdvanceFieldAdjustment { operation, points });
        index += 2;
    }

    Some(adjustments)
}

pub(super) fn mail_merge_recipient_field_parts(
    instruction: &str,
) -> Option<MailMergeRecipientFieldParts<'_>> {
    let mut tokens = tokenize(instruction).ok()?;
    let keyword = tokens.first()?;
    let kind = if keyword.value.eq_ignore_ascii_case("ADDRESSBLOCK") {
        MailMergeRecipientFieldKind::AddressBlock
    } else if keyword.value.eq_ignore_ascii_case("GREETINGLINE") {
        MailMergeRecipientFieldKind::GreetingLine
    } else {
        return None;
    };
    tokens.remove(0);

    let mut country_inclusion = None;
    let mut formats_using_recipient_country = false;
    let mut excluded_countries = Vec::new();
    let mut format_template = None;
    let mut language = None;
    let mut greeting_fallback_text = None;
    let mut unknown_switches = Vec::new();
    let mut index = 0;
    while index < tokens.len() {
        let name = switch_name(tokens.get(index)?)?;
        let normalized = name.to_ascii_lowercase();
        let value = tokens
            .get(index + 1)
            .filter(|token| switch_name(token).is_none());

        match (kind, normalized.as_str()) {
            (MailMergeRecipientFieldKind::AddressBlock, "c") => {
                if country_inclusion.is_some() {
                    return None;
                }
                let inclusion_value = value?.value.clone();
                country_inclusion = Some(match inclusion_value.as_ref() {
                    "0" => AddressBlockCountryInclusion::Omit,
                    "1" => AddressBlockCountryInclusion::Always,
                    "2" => AddressBlockCountryInclusion::UnlessExcluded,
                    _ => return None,
                });
                index += 2;
            },
            (MailMergeRecipientFieldKind::AddressBlock, "d") => {
                if formats_using_recipient_country || value.is_some() {
                    return None;
                }
                formats_using_recipient_country = true;
                index += 1;
            },
            (MailMergeRecipientFieldKind::AddressBlock, "e") => {
                excluded_countries.push(value?.value.clone());
                index += 2;
            },
            (_, "f") => {
                if format_template.is_some() {
                    return None;
                }
                format_template = Some(value?.value.clone());
                index += 2;
            },
            (_, "l") => {
                if language.is_some() {
                    return None;
                }
                language = Some(value?.value.clone());
                index += 2;
            },
            (MailMergeRecipientFieldKind::GreetingLine, "c" | "e") => {
                if greeting_fallback_text.is_some() {
                    return None;
                }
                greeting_fallback_text = Some(value?.value.clone());
                index += 2;
            },
            _ => {
                unknown_switches.push(FieldSwitch {
                    name: Cow::Owned(name.to_string()),
                    value: value.map(|token| token.value.clone()),
                });
                index += 1 + usize::from(value.is_some());
            },
        }
    }

    Some(MailMergeRecipientFieldParts {
        kind,
        country_inclusion,
        formats_using_recipient_country,
        excluded_countries,
        format_template,
        language,
        greeting_fallback_text,
        unknown_switches,
    })
}

pub(super) fn mail_merge_counter_kind(instruction: &str) -> Option<MailMergeCounterKind> {
    let tokens = tokenize(instruction).ok()?;
    let [keyword] = tokens.as_slice() else {
        return None;
    };
    if keyword.value.eq_ignore_ascii_case("MERGEREC") {
        Some(MailMergeCounterKind::Record)
    } else if keyword.value.eq_ignore_ascii_case("MERGESEQ") {
        Some(MailMergeCounterKind::Sequence)
    } else {
        None
    }
}

pub(super) fn is_mail_merge_next_instruction(instruction: &str) -> bool {
    let Ok(tokens) = tokenize(instruction) else {
        return false;
    };
    matches!(tokens.as_slice(), [keyword] if keyword.value.eq_ignore_ascii_case("NEXT"))
}

pub(super) fn mail_merge_conditional_control_parts(
    raw_instruction: &str,
) -> Option<(MailMergeConditionalControlKind, &str)> {
    tokenize(raw_instruction).ok()?;
    let instruction =
        raw_instruction.trim_start_matches(|character: char| character.is_ascii_whitespace());
    let (kind, keyword) = if instruction
        .get(..6)
        .is_some_and(|candidate| candidate.eq_ignore_ascii_case("NEXTIF"))
    {
        (MailMergeConditionalControlKind::NextIf, "NEXTIF")
    } else if instruction
        .get(..6)
        .is_some_and(|candidate| candidate.eq_ignore_ascii_case("SKIPIF"))
    {
        (MailMergeConditionalControlKind::SkipIf, "SKIPIF")
    } else {
        return None;
    };
    let remainder = instruction.get(keyword.len()..)?;
    if !remainder
        .chars()
        .next()
        .is_some_and(|character| character.is_ascii_whitespace())
    {
        return None;
    }
    let comparison = remainder.trim_matches(|character: char| character.is_ascii_whitespace());
    (!comparison.is_empty()).then_some((kind, comparison))
}

pub(super) fn comparison_field_expression<'a>(
    raw_instruction: &'a str,
    field_type: &str,
) -> Option<&'a str> {
    tokenize(raw_instruction).ok()?;
    let instruction =
        raw_instruction.trim_start_matches(|character: char| character.is_ascii_whitespace());
    let candidate = instruction.get(..field_type.len())?;
    if !candidate.eq_ignore_ascii_case(field_type) {
        return None;
    }
    let remainder = instruction.get(field_type.len()..)?;
    if !remainder
        .chars()
        .next()
        .is_some_and(|character| character.is_ascii_whitespace())
    {
        return None;
    }
    let comparison = remainder.trim_matches(|character: char| character.is_ascii_whitespace());
    (!comparison.is_empty()).then_some(comparison)
}

pub(super) fn if_field_expression(instruction: &str) -> Option<&str> {
    comparison_field_expression(instruction, "IF")
}

pub(super) fn compare_field_comparison(instruction: &str) -> Option<&str> {
    comparison_field_expression(instruction, "COMPARE")
}

pub(super) fn set_field_parts(raw_instruction: &str) -> Option<(Cow<'_, str>, &str)> {
    let tokens = tokenize(raw_instruction).ok()?;
    let keyword = tokens.first()?;
    if !keyword.value.eq_ignore_ascii_case("SET") {
        return None;
    }
    let target_name = tokens.get(1)?.value.clone();
    if target_name.is_empty() || switch_name(tokens.get(1)?).is_some() {
        return None;
    }

    let instruction =
        raw_instruction.trim_start_matches(|character: char| character.is_ascii_whitespace());
    let remainder = instruction.get("SET".len()..)?;
    if !remainder
        .chars()
        .next()
        .is_some_and(|character| character.is_ascii_whitespace())
    {
        return None;
    }
    let target_start =
        remainder.trim_start_matches(|character: char| character.is_ascii_whitespace());
    let target_end = field_argument_end(target_start)?;
    let expression = target_start.get(target_end..)?;
    if !expression
        .chars()
        .next()
        .is_some_and(|character| character.is_ascii_whitespace())
    {
        return None;
    }
    let trimmed_expression =
        expression.trim_matches(|character: char| character.is_ascii_whitespace());
    (!trimmed_expression.is_empty()).then_some((target_name, trimmed_expression))
}

pub(super) fn field_argument_end(input: &str) -> Option<usize> {
    let bytes = input.as_bytes();
    if bytes.first().copied()? != b'"' {
        return Some(
            bytes
                .iter()
                .position(u8::is_ascii_whitespace)
                .unwrap_or(bytes.len()),
        );
    }

    let mut index = 1;
    while let Some(byte) = bytes.get(index).copied() {
        match byte {
            b'"' => return Some(index + 1),
            b'\\'
                if bytes
                    .get(index + 1)
                    .is_some_and(|next| matches!(*next, b'\\' | b'"')) =>
            {
                index += 2;
            },
            _ => {
                let character = input.get(index..)?.chars().next()?;
                index += character.len_utf8();
            },
        }
    }
    None
}

#[allow(
    clippy::type_complexity,
    reason = "the returned tuple mirrors the field-instruction grammar"
)]
pub(super) fn sequence_field_parts(
    raw_instruction: &str,
) -> Option<(Cow<'_, str>, Option<Cow<'_, str>>, &str)> {
    let tokens = tokenize(raw_instruction).ok()?;
    let keyword = tokens.first()?;
    if !keyword.value.eq_ignore_ascii_case("SEQ") {
        return None;
    }
    let identifier = tokens.get(1)?.value.clone();
    if identifier.is_empty() || switch_name(tokens.get(1)?).is_some() {
        return None;
    }

    let instruction =
        raw_instruction.trim_start_matches(|character: char| character.is_ascii_whitespace());
    let remainder = instruction.get("SEQ".len()..)?;
    if !remainder
        .chars()
        .next()
        .is_some_and(|character| character.is_ascii_whitespace())
    {
        return None;
    }
    let identifier_start =
        remainder.trim_start_matches(|character: char| character.is_ascii_whitespace());
    let identifier_end = field_argument_end(identifier_start)?;
    let after_identifier = identifier_start.get(identifier_end..)?;
    if after_identifier.is_empty() {
        return Some((identifier, None, ""));
    }
    if !after_identifier
        .chars()
        .next()
        .is_some_and(|character| character.is_ascii_whitespace())
    {
        return None;
    }
    let trailing =
        after_identifier.trim_start_matches(|character: char| character.is_ascii_whitespace());
    if trailing.is_empty() || trailing.starts_with('\\') {
        return Some((
            identifier,
            None,
            trailing.trim_matches(|character: char| character.is_ascii_whitespace()),
        ));
    }

    let bookmark = tokens.get(2)?.value.clone();
    if bookmark.is_empty() || switch_name(tokens.get(2)?).is_some() {
        return None;
    }
    let bookmark_end = field_argument_end(trailing)?;
    let after_bookmark = trailing.get(bookmark_end..)?;
    if !after_bookmark.is_empty()
        && !after_bookmark
            .chars()
            .next()
            .is_some_and(|character| character.is_ascii_whitespace())
    {
        return None;
    }
    let tail = after_bookmark.trim_matches(|character: char| character.is_ascii_whitespace());
    Some((identifier, Some(bookmark), tail))
}

pub(super) fn is_formula_field_instruction(instruction: &str) -> bool {
    tokenize(instruction).is_ok() && instruction.trim_start().starts_with('=')
}

pub(super) fn formula_field_formula(instruction: &str) -> Option<&str> {
    tokenize(instruction).ok()?;
    let formula = instruction.trim().strip_prefix('=')?.trim();
    (!formula.is_empty()).then_some(formula)
}

pub(super) fn quote_field_parts(instruction: &str) -> Option<(Cow<'_, str>, Vec<FieldSwitch<'_>>)> {
    let tokens = tokenize(instruction).ok()?;
    let keyword = tokens.first()?;
    if !keyword.value.eq_ignore_ascii_case("QUOTE") {
        return None;
    }
    let text = tokens.get(1)?.value.clone();
    if switch_name(tokens.get(1)?).is_some() {
        return None;
    }

    let mut switches = Vec::new();
    let mut index = 2;
    while index < tokens.len() {
        let name = switch_name(tokens.get(index)?)?;
        let value = tokens
            .get(index + 1)
            .filter(|next| switch_name(next).is_none());
        switches.push(FieldSwitch {
            name: Cow::Owned(name.to_ascii_lowercase()),
            value: value.map(|token| token.value.clone()),
        });
        index += 1 + usize::from(value.is_some());
    }
    Some((text, switches))
}

pub(super) fn symbol_field_parts(
    instruction: &str,
) -> Option<(Cow<'_, str>, Vec<FieldSwitch<'_>>)> {
    let tokens = tokenize(instruction).ok()?;
    let keyword = tokens.first()?;
    if !keyword.value.eq_ignore_ascii_case("SYMBOL") {
        return None;
    }
    let character_argument = tokens.get(1)?.value.clone();
    if switch_name(tokens.get(1)?).is_some() {
        return None;
    }

    let mut switches = Vec::new();
    let mut index = 2;
    while index < tokens.len() {
        let name = switch_name(tokens.get(index)?)?;
        let value = tokens
            .get(index + 1)
            .filter(|next| switch_name(next).is_none());
        switches.push(FieldSwitch {
            name: Cow::Owned(name.to_ascii_lowercase()),
            value: value.map(|token| token.value.clone()),
        });
        index += 1 + usize::from(value.is_some());
    }
    Some((character_argument, switches))
}

pub(super) fn auto_number_field_parts(instruction: &str) -> Option<AutoNumberFieldParts<'_>> {
    let mut tokens = tokenize(instruction).ok()?;
    let keyword = tokens.first()?;
    let kind = AutoNumberFieldKind::from_keyword(keyword.value.as_ref())?;
    tokens.remove(0);

    let mut switches = Vec::new();
    let mut index = 0;
    while index < tokens.len() {
        let name = switch_name(tokens.get(index)?)?;
        let value = tokens
            .get(index + 1)
            .filter(|token| is_field_operand(token));
        switches.push(FieldSwitch {
            name: Cow::Owned(name.to_string()),
            value: value.map(|token| token.value.clone()),
        });
        index += 1 + usize::from(value.is_some());
    }

    Some(AutoNumberFieldParts { kind, switches })
}

pub(super) fn list_number_field_parts(instruction: &str) -> Option<ListNumberFieldParts<'_>> {
    let mut tokens = tokenize(instruction).ok()?;
    let keyword = tokens.first()?;
    if !keyword.value.eq_ignore_ascii_case("LISTNUM") {
        return None;
    }
    tokens.remove(0);

    let list_name = tokens
        .first()
        .filter(|token| is_field_operand(token))
        .map(|token| token.value.clone());
    if list_name.is_some() {
        tokens.remove(0);
    }

    let mut switches = Vec::new();
    let mut index = 0;
    while index < tokens.len() {
        let name = switch_name(tokens.get(index)?)?;
        let value = tokens
            .get(index + 1)
            .filter(|token| is_field_operand(token));
        switches.push(FieldSwitch {
            name: Cow::Owned(name.to_string()),
            value: value.map(|token| token.value.clone()),
        });
        index += 1 + usize::from(value.is_some());
    }

    Some(ListNumberFieldParts {
        list_name,
        switches,
    })
}

pub(super) fn style_reference_field_parts(
    instruction: &str,
) -> Option<(
    Cow<'_, str>,
    Vec<StyleReferenceFieldOption>,
    Vec<FieldSwitch<'_>>,
)> {
    let tokens = tokenize(instruction).ok()?;
    let keyword = tokens.first()?;
    if !keyword.value.eq_ignore_ascii_case("STYLEREF") {
        return None;
    }
    let style_name = tokens.get(1)?.value.clone();
    if style_name.is_empty() || switch_name(tokens.get(1)?).is_some() {
        return None;
    }

    let mut options = Vec::new();
    let mut unknown_switches = Vec::new();
    let mut index = 2;
    while index < tokens.len() {
        let token = tokens.get(index)?;
        let name = switch_name(token)?;
        let option = match name.to_ascii_lowercase().as_str() {
            "l" => Some(StyleReferenceFieldOption::FollowingText),
            "n" => Some(StyleReferenceFieldOption::ParagraphNumber),
            "p" => Some(StyleReferenceFieldOption::RelativePosition),
            "r" => Some(StyleReferenceFieldOption::ParagraphNumberRelativeContext),
            "t" => Some(StyleReferenceFieldOption::SuppressNonNumberText),
            "w" => Some(StyleReferenceFieldOption::ParagraphNumberFullContext),
            _ => None,
        };
        if let Some(field_option) = option {
            if tokens
                .get(index + 1)
                .is_some_and(|next| switch_name(next).is_none())
            {
                return None;
            }
            options.push(field_option);
            index += 1;
            continue;
        }

        let value = tokens
            .get(index + 1)
            .filter(|next| switch_name(next).is_none());
        unknown_switches.push(FieldSwitch {
            name: Cow::Owned(name.to_string()),
            value: value.map(|switch_token| switch_token.value.clone()),
        });
        index += 1 + usize::from(value.is_some());
    }
    Some((style_name, options, unknown_switches))
}

#[must_use]
pub fn parse_field_code(instruction: &str) -> ParsedFieldCode<'_> {
    match parse_field_code_inner(instruction) {
        Ok(parsed) => parsed,
        Err(error) => ParsedFieldCode::Malformed(error),
    }
}

pub(super) fn parse_field_code_inner(
    instruction: &str,
) -> Result<ParsedFieldCode<'_>, FieldCodeError> {
    let mut tokens = tokenize(instruction)?;
    if tokens.is_empty() {
        return Err(FieldCodeError::MissingKeyword);
    }
    let keyword = tokens.remove(0);
    if keyword.value.eq_ignore_ascii_case("HYPERLINK") {
        return parse_hyperlink(&tokens).map(ParsedFieldCode::Hyperlink);
    }
    for (name, constructor) in [("REF", 0u8), ("PAGEREF", 1u8), ("NOTEREF", 2u8)] {
        if keyword.value.eq_ignore_ascii_case(name) {
            let code = parse_reference(&tokens)?;
            return Ok(match constructor {
                0 => ParsedFieldCode::Reference(code),
                1 => ParsedFieldCode::PageReference(code),
                _ => ParsedFieldCode::NoteReference(code),
            });
        }
    }
    Ok(ParsedFieldCode::Other {
        keyword: keyword.value,
        arguments: tokens,
    })
}

pub(super) fn parse_hyperlink<'a>(
    tokens: &[FieldCodeToken<'a>],
) -> Result<HyperlinkCode<'a>, FieldCodeError> {
    let mut code = HyperlinkCode {
        external_target: None,
        bookmark: None,
        screen_tip: None,
        target_frame: None,
        coordinates: None,
        new_window: false,
        unknown_switches: Vec::new(),
    };
    let mut index = 0;
    while let Some(token) = tokens.get(index) {
        if let Some(name) = switch_name(token) {
            let normalized = name.to_ascii_lowercase();
            match normalized.as_str() {
                "n" => {
                    if code.new_window {
                        return Err(FieldCodeError::DuplicateOperand("\\n"));
                    }
                    code.new_window = true;
                    index += 1;
                },
                "l" | "o" | "t" | "m" => {
                    let value = switch_value(tokens, index, name)?;
                    let slot = match normalized.as_str() {
                        "l" => &mut code.bookmark,
                        "o" => &mut code.screen_tip,
                        "t" => &mut code.target_frame,
                        _ => &mut code.coordinates,
                    };
                    if slot.replace(value).is_some() {
                        return Err(FieldCodeError::DuplicateOperand(
                            match normalized.as_str() {
                                "l" => "\\l",
                                "o" => "\\o",
                                "t" => "\\t",
                                _ => "\\m",
                            },
                        ));
                    }
                    index += 2;
                },
                _ => {
                    let value = tokens
                        .get(index + 1)
                        .filter(|next| switch_name(next).is_none());
                    code.unknown_switches.push(FieldSwitch {
                        name: Cow::Owned(name.to_string()),
                        value: value.map(|switch_token| switch_token.value.clone()),
                    });
                    index += 1 + usize::from(value.is_some());
                },
            }
        } else {
            if code.external_target.replace(token.value.clone()).is_some() {
                return Err(FieldCodeError::UnexpectedOperand(token.value.to_string()));
            }
            index += 1;
        }
    }
    if code.external_target.is_none() && code.bookmark.is_none() {
        return Err(FieldCodeError::MissingOperand(
            "hyperlink target or \\l bookmark",
        ));
    }
    Ok(code)
}

pub(super) fn parse_reference<'a>(
    tokens: &[FieldCodeToken<'a>],
) -> Result<ReferenceCode<'a>, FieldCodeError> {
    let Some(first) = tokens.first() else {
        return Err(FieldCodeError::MissingOperand("bookmark"));
    };
    if switch_name(first).is_some() {
        return Err(FieldCodeError::MissingOperand("bookmark"));
    }
    let mut code = ReferenceCode {
        bookmark: first.value.clone(),
        hyperlink: false,
        position: false,
        footnote_mark: false,
        unknown_switches: Vec::new(),
    };
    let mut index = 1;
    while let Some(token) = tokens.get(index) {
        let Some(name) = switch_name(token) else {
            return Err(FieldCodeError::UnexpectedOperand(token.value.to_string()));
        };
        match name.to_ascii_lowercase().as_str() {
            "h" if !code.hyperlink => code.hyperlink = true,
            "p" if !code.position => code.position = true,
            "f" if !code.footnote_mark => code.footnote_mark = true,
            "h" => return Err(FieldCodeError::DuplicateOperand("\\h")),
            "p" => return Err(FieldCodeError::DuplicateOperand("\\p")),
            "f" => return Err(FieldCodeError::DuplicateOperand("\\f")),
            _ => {
                let value = tokens
                    .get(index + 1)
                    .filter(|next| switch_name(next).is_none());
                code.unknown_switches.push(FieldSwitch {
                    name: Cow::Owned(name.to_string()),
                    value: value.map(|switch_token| switch_token.value.clone()),
                });
                if value.is_some() {
                    index += 1;
                }
            },
        }
        index += 1;
    }
    Ok(code)
}

pub(super) fn switch_name<'a>(token: &'a FieldCodeToken<'_>) -> Option<&'a str> {
    if token.quoted {
        return None;
    }
    token
        .value
        .strip_prefix('\\')
        .filter(|name| !name.is_empty())
}

pub(super) fn is_field_operand(token: &FieldCodeToken<'_>) -> bool {
    switch_name(token).is_none() && (token.quoted || !token.value.starts_with('\\'))
}

pub(super) fn switch_value<'a>(
    tokens: &[FieldCodeToken<'a>],
    index: usize,
    name: &str,
) -> Result<Cow<'a, str>, FieldCodeError> {
    let value = tokens
        .get(index + 1)
        .filter(|value| switch_name(value).is_none())
        .ok_or(FieldCodeError::MissingOperand("switch value"))?;
    if name.is_empty() {
        return Err(FieldCodeError::MissingOperand("switch name"));
    }
    Ok(value.value.clone())
}

pub(super) fn tokenize(instruction: &str) -> Result<Vec<FieldCodeToken<'_>>, FieldCodeError> {
    if instruction.len() > MAX_INSTRUCTION_LEN {
        return Err(FieldCodeError::InstructionTooLong);
    }
    let bytes = instruction.as_bytes();
    let mut tokens = Vec::new();
    let mut index = 0;
    while index < bytes.len() {
        while bytes.get(index).is_some_and(u8::is_ascii_whitespace) {
            index += 1;
        }
        if index == bytes.len() {
            break;
        }
        if tokens.len() >= MAX_TOKENS {
            return Err(FieldCodeError::TooManyTokens);
        }
        if bytes.get(index).copied() == Some(b'"') {
            index += 1;
            let mut value = String::new();
            let mut closed = false;
            while let Some(byte) = bytes.get(index).copied() {
                match byte {
                    b'"' => {
                        index += 1;
                        closed = true;
                        break;
                    },
                    b'\\'
                        if bytes
                            .get(index + 1)
                            .is_some_and(|next| matches!(*next, b'\\' | b'"')) =>
                    {
                        let escaped = bytes.get(index + 1).copied().ok_or_else(|| {
                            FieldCodeError::UnexpectedOperand(
                                "missing quoted field escape".to_string(),
                            )
                        })?;
                        value.push(char::from(escaped));
                        index += 2;
                    },
                    _ => {
                        let character = instruction
                            .get(index..)
                            .and_then(|remainder| remainder.chars().next())
                            .ok_or_else(|| {
                                FieldCodeError::UnexpectedOperand(
                                    "invalid quoted field operand boundary".to_string(),
                                )
                            })?;
                        value.push(character);
                        index += character.len_utf8();
                    },
                }
            }
            if !closed {
                return Err(FieldCodeError::UnterminatedQuote);
            }
            tokens.push(FieldCodeToken {
                value: Cow::Owned(value),
                quoted: true,
            });
        } else {
            let start = index;
            while bytes
                .get(index)
                .is_some_and(|byte| !byte.is_ascii_whitespace())
            {
                index += 1;
            }
            tokens.push(FieldCodeToken {
                value: Cow::Borrowed(instruction.get(start..index).ok_or_else(|| {
                    FieldCodeError::UnexpectedOperand("invalid field operand boundary".to_string())
                })?),
                quoted: false,
            });
        }
    }
    Ok(tokens)
}

pub(crate) fn quoted_field_operand(value: &str) -> String {
    let mut output = String::with_capacity(value.len() + 2);
    output.push('"');
    for character in value.chars() {
        match character {
            '\\' => output.push_str("\\\\"),
            '"' => output.push_str("\\\""),
            _ => output.push(character),
        }
    }
    output.push('"');
    output
}
