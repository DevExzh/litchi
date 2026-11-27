//! Bounded merge field-instruction parsers.

use super::prelude::*;

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
