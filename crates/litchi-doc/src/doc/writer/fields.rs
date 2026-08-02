//! Strict MS-DOC field-table authoring.
//!
//! `Plcfld` stores balanced field characters and a native type byte on each
//! begin marker. MS-DOC §2.8.25 requires five field kinds (`TC`, `TA`, `XE`,
//! `RD`, and `PRIVATE`) to remain text-only and be omitted from the table.

use super::core::DocWriteError;
use crate::doc::parts::fields::FieldType;

const MAX_WRITER_FIELD_DEPTH: usize = 128;
const MAX_WRITER_FIELD_KEYWORD_BYTES: usize = 32;

#[derive(Clone, Copy)]
struct OpenWriterField {
    begin_marker: usize,
    separator_marker: Option<usize>,
}

fn story_code_unit(story_text: &[u8], cp: u32) -> Option<u16> {
    let offset = usize::try_from(cp).ok()?.checked_mul(2)?;
    let bytes = story_text.get(offset..offset.checked_add(2)?)?;
    Some(u16::from_le_bytes([bytes[0], bytes[1]]))
}

fn stored_field_keyword(
    story_text: &[u8],
    start_cp: u32,
    end_cp: u32,
) -> Result<Option<String>, DocWriteError> {
    if start_cp > end_cp {
        return Err(invalid("DOC field instruction range is reversed"));
    }
    let start = usize::try_from(start_cp)
        .ok()
        .and_then(|value| value.checked_mul(2))
        .ok_or_else(|| invalid("DOC field instruction start overflows"))?;
    let end = usize::try_from(end_cp)
        .ok()
        .and_then(|value| value.checked_mul(2))
        .ok_or_else(|| invalid("DOC field instruction end overflows"))?;
    let bytes = story_text
        .get(start..end)
        .ok_or_else(|| invalid("DOC field instruction exceeds its story"))?;

    let mut keyword = String::new();
    for pair in bytes.chunks_exact(2) {
        let unit = u16::from_le_bytes([pair[0], pair[1]]);
        if keyword.is_empty() && matches!(unit, 0x0009..=0x000D | 0x0020) {
            continue;
        }
        if matches!(unit, 0x0009..=0x000D | 0x0020) {
            break;
        }
        if keyword.is_empty() && unit == u16::from(b'=') {
            keyword.push('=');
            break;
        }
        let byte = u8::try_from(unit).ok().filter(u8::is_ascii);
        let Some(byte) = byte else {
            return Ok(None);
        };
        if keyword.len() >= MAX_WRITER_FIELD_KEYWORD_BYTES {
            return Ok(None);
        }
        keyword.push(char::from(byte));
    }
    Ok((!keyword.is_empty()).then_some(keyword))
}

fn is_non_plcf_field_keyword(keyword: &str) -> bool {
    ["TC", "TA", "XE", "RD", "PRIVATE"]
        .into_iter()
        .any(|excluded| keyword.eq_ignore_ascii_case(excluded))
}

/// Build one complete `Plcfld`, validating its marker graph and native types.
///
/// `field_char_cps` and `text_length` use story-relative UTF-16 character
/// positions. `story_text` is the exact little-endian UTF-16 story payload.
/// An empty vector is returned when every field is one of the five kinds that
/// MS-DOC excludes from `Plcfld`.
pub(super) fn build_plcffld(
    field_char_cps: &[(u32, u16)],
    text_length: u32,
    story_text: &[u8],
) -> Result<Vec<u8>, DocWriteError> {
    let expected_bytes = usize::try_from(text_length)
        .ok()
        .and_then(|value| value.checked_mul(2))
        .ok_or_else(|| invalid("DOC field story size overflows"))?;
    if story_text.len() != expected_bytes {
        return Err(invalid(
            "DOC field story byte count disagrees with its character count",
        ));
    }

    let mut included = vec![true; field_char_cps.len()];
    let mut begin_types = vec![None; field_char_cps.len()];
    let mut stack = Vec::new();
    let mut previous_cp = None;
    for (index, &(cp, field_character)) in field_char_cps.iter().enumerate() {
        if cp >= text_length
            || previous_cp.is_some_and(|previous| cp <= previous)
            || story_code_unit(story_text, cp) != Some(field_character)
        {
            return Err(invalid("DOC field markers do not match their story text"));
        }
        previous_cp = Some(cp);
        match field_character {
            0x13 => {
                if stack.len() >= MAX_WRITER_FIELD_DEPTH {
                    return Err(invalid("DOC field nesting exceeds the writer limit"));
                }
                stack.push(OpenWriterField {
                    begin_marker: index,
                    separator_marker: None,
                });
            },
            0x14 => {
                let open = stack
                    .last_mut()
                    .ok_or_else(|| invalid("DOC field separator has no begin marker"))?;
                if open.separator_marker.replace(index).is_some() {
                    return Err(invalid("DOC field contains duplicate separators"));
                }
            },
            0x15 => {
                let open = stack
                    .pop()
                    .ok_or_else(|| invalid("DOC field end has no begin marker"))?;
                let instruction_end = open
                    .separator_marker
                    .map_or(cp, |separator| field_char_cps[separator].0);
                let instruction_start = field_char_cps[open.begin_marker]
                    .0
                    .checked_add(1)
                    .ok_or_else(|| invalid("DOC field instruction start overflows"))?;
                let keyword = stored_field_keyword(story_text, instruction_start, instruction_end)?;
                if keyword.as_deref().is_some_and(is_non_plcf_field_keyword) {
                    included[open.begin_marker] = false;
                    included[index] = false;
                    if let Some(separator) = open.separator_marker {
                        included[separator] = false;
                    }
                } else {
                    begin_types[open.begin_marker] = Some(
                        keyword
                            .as_deref()
                            .and_then(FieldType::from_keyword)
                            .unwrap_or(FieldType::Unparsed),
                    );
                }
            },
            _ => return Err(invalid("DOC field marker has an invalid character")),
        }
    }
    if !stack.is_empty() {
        return Err(invalid("DOC field begin has no end marker"));
    }

    let marker_count = included.iter().filter(|&&value| value).count();
    if marker_count == 0 {
        return Ok(Vec::new());
    }
    let capacity = marker_count
        .checked_add(1)
        .and_then(|count| count.checked_mul(4))
        .and_then(|cp_bytes| {
            marker_count
                .checked_mul(2)
                .and_then(|field_bytes| cp_bytes.checked_add(field_bytes))
        })
        .ok_or_else(|| invalid("DOC Plcfld size overflows"))?;
    let mut plcffld = Vec::with_capacity(capacity);
    for (index, (cp, _)) in field_char_cps.iter().enumerate() {
        if included[index] {
            plcffld.extend_from_slice(&cp.to_le_bytes());
        }
    }
    plcffld.extend_from_slice(&text_length.to_le_bytes());

    let mut field_stack = Vec::new();
    for (index, (_, field_character)) in field_char_cps.iter().enumerate() {
        if !included[index] {
            continue;
        }
        let (fldch, flt_or_flags) = match *field_character {
            0x13 => {
                field_stack.push(false);
                let field_type = begin_types[index]
                    .ok_or_else(|| invalid("DOC field begin has no classified field type"))?;
                (0x13, field_type.as_u8())
            },
            0x14 => {
                let has_separator = field_stack
                    .last_mut()
                    .ok_or_else(|| invalid("DOC filtered field separator has no begin marker"))?;
                *has_separator = true;
                (0x14, 0x00)
            },
            0x15 => {
                let has_separator = field_stack
                    .pop()
                    .ok_or_else(|| invalid("DOC filtered field end has no begin marker"))?;
                let flags = u8::from(has_separator) << 7 | u8::from(!field_stack.is_empty()) << 6;
                (0x15, flags)
            },
            _ => unreachable!("validated field marker"),
        };
        plcffld.push(fldch);
        plcffld.push(flt_or_flags);
    }
    debug_assert!(field_stack.is_empty());
    Ok(plcffld)
}

fn invalid(message: impl Into<String>) -> DocWriteError {
    DocWriteError::InvalidData(message.into())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn build(text: &str) -> Result<Vec<u8>, DocWriteError> {
        let units: Vec<u16> = text.encode_utf16().collect();
        let markers: Vec<(u32, u16)> = units
            .iter()
            .enumerate()
            .filter(|(_, unit)| matches!(**unit, 0x13..=0x15))
            .map(|(cp, unit)| (u32::try_from(cp).unwrap(), *unit))
            .collect();
        let bytes: Vec<u8> = units.iter().flat_map(|unit| unit.to_le_bytes()).collect();
        build_plcffld(&markers, u32::try_from(units.len()).unwrap(), &bytes)
    }

    #[test]
    fn emits_native_types_nested_flags_and_text_only_exclusions() {
        let top_level = build("\u{0013}HYPERLINK \"x\"\u{0014}link\u{0015}\r").unwrap();
        assert_eq!(&top_level[16..], &[0x13, 0x58, 0x14, 0x00, 0x15, 0x80]);

        let nested =
            build("\u{0013}IF 1 = \u{0013}HYPERLINK \"x\"\u{0014}x\u{0015}\u{0014}yes\u{0015}\r")
                .unwrap();
        assert_eq!(
            &nested[28..],
            &[
                0x13, 0x07, 0x13, 0x58, 0x14, 0x00, 0x15, 0xC0, 0x14, 0x00, 0x15, 0x80,
            ]
        );

        let excluded = build(concat!(
            "\u{0013}TC Entry\u{0015}",
            "\u{0013}HYPERLINK \"x\"\u{0014}x\u{0015}\r",
        ))
        .unwrap();
        assert_eq!(excluded.len(), 22);
        assert_eq!(&excluded[16..], &[0x13, 0x58, 0x14, 0x00, 0x15, 0x80]);
        assert!(
            build("\u{0013}PRIVATE opaque\u{0015}\r")
                .unwrap()
                .is_empty()
        );
    }

    #[test]
    fn rejects_malformed_marker_graphs_and_positions() {
        let bytes = |text: &str| {
            text.encode_utf16()
                .flat_map(u16::to_le_bytes)
                .collect::<Vec<_>>()
        };

        let unmatched = bytes("\u{0013}HYPERLINK\r");
        assert!(build_plcffld(&[(0, 0x13)], 11, &unmatched).is_err());

        let duplicate_separator = bytes("\u{0013}IF\u{0014}x\u{0014}y\u{0015}\r");
        assert!(
            build_plcffld(
                &[(0, 0x13), (3, 0x14), (5, 0x14), (7, 0x15)],
                9,
                &duplicate_separator,
            )
            .is_err()
        );

        let valid = bytes("\u{0013}PAGE\u{0015}\r");
        assert!(build_plcffld(&[(0, 0x13), (4, 0x15)], 7, &valid).is_err());
    }
}
