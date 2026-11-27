//! Bounded numbering field-instruction parsers.

use super::prelude::*;

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
