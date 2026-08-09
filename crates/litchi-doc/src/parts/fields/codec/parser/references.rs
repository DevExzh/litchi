//! Bounded references field-instruction parsers.

use super::prelude::*;

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
        None | Some('"' | '\\') => Some(remainder.trim().to_string()),
        Some(character) if character.is_whitespace() => Some(remainder.trim().to_string()),
        Some(_) => None,
    }
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
