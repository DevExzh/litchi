//! Bounded buttons field-instruction parsers.

use super::prelude::*;

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
