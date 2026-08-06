//! Bounded metadata field-instruction parsers.

use super::prelude::*;

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
