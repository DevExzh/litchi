//! Bounded formatting field-instruction parsers.

use super::prelude::*;

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
