//! Pure list resolution and counter state for WordprocessingML paragraphs.

use crate::error::{Error, Result};
use crate::numbering::{Collection, Format, Level, Paragraph, Restart, Suffix};
use std::collections::HashMap;

const MAX_LEVEL_TEXT_BYTES: usize = 4096;
const MAX_LABEL_BYTES: usize = 8192;

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ListMarker {
    Text(String),
    PictureBullet { id: u32 },
    UnsupportedFormat { format: String, value: i64 },
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ListItem {
    pub paragraph_index: usize,
    pub numbering: Paragraph,
    pub marker: ListMarker,
    pub suffix: Suffix,
    pub text: String,
}

#[derive(Debug, Clone, Default)]
pub struct ListCounterState {
    counters: HashMap<u32, [Option<i64>; 9]>,
}

/// Resolve an immutable level, applying a concrete instance's level override.
fn resolve_level(numbering: &Collection, num_id: u32, level: u8) -> Result<Level> {
    if level > 8 {
        return Err(invalid(&format!("invalid numbering level '{level}'")));
    }
    let num = numbering
        .get_num(num_id)
        .ok_or_else(|| invalid(&format!("paragraph references missing numId {num_id}")))?;
    let abstract_num = numbering
        .get_abstract_num(num.abstract_num_id())
        .ok_or_else(|| {
            invalid(&format!(
                "numId {num_id} references a missing abstract numbering definition"
            ))
        })?;
    let level_override = num.overrides().iter().find(|value| value.level == level);
    let mut resolved = match level_override.and_then(|value| value.definition.clone()) {
        Some(value) => value,
        None => abstract_num.level(level).cloned().ok_or_else(|| {
            invalid(&format!(
                "numId {num_id} has no definition for level {level}"
            ))
        })?,
    };
    if let Some(start) = level_override.and_then(|value| value.start_override) {
        resolved.start = start;
    }
    Ok(resolved)
}

impl ListCounterState {
    pub fn new() -> Self {
        Self::default()
    }

    /// Advance one numbered paragraph and return its typed marker and suffix.
    pub fn advance(
        &mut self,
        numbering: &Collection,
        properties: Paragraph,
    ) -> Result<(ListMarker, Suffix)> {
        if properties.num_id == 0 {
            return Err(invalid("numId 0 cancels numbering and cannot be advanced"));
        }
        let current_level = resolve_level(numbering, properties.num_id, properties.level)?;
        let counters = self.counters.entry(properties.num_id).or_insert([None; 9]);

        for ancestor in 0..properties.level {
            if counters[usize::from(ancestor)].is_none() {
                counters[usize::from(ancestor)] =
                    Some(resolve_level(numbering, properties.num_id, ancestor)?.start);
            }
        }

        let slot = &mut counters[usize::from(properties.level)];
        *slot = Some(match *slot {
            None => current_level.start,
            Some(value) => value
                .checked_add(1)
                .ok_or_else(|| invalid("numbering counter overflow"))?,
        });

        for deeper in properties.level.saturating_add(1)..=8 {
            let should_reset = match resolve_level(numbering, properties.num_id, deeper) {
                Ok(value) => match value.restart {
                    Restart::Default => true,
                    Restart::Never => false,
                    Restart::After(restart_level) => properties.level <= restart_level,
                },
                Err(_) => false,
            };
            if should_reset {
                counters[usize::from(deeper)] = None;
            }
        }

        let marker = render_marker(numbering, properties, counters, &current_level)?;
        Ok((marker, current_level.suffix))
    }
}

fn render_marker(
    numbering: &Collection,
    properties: Paragraph,
    counters: &[Option<i64>; 9],
    current: &Level,
) -> Result<ListMarker> {
    if let Some(id) = current.picture_bullet_id {
        return Ok(ListMarker::PictureBullet { id });
    }
    let template = current.level_text.as_deref().unwrap_or_default();
    if template.len() > MAX_LEVEL_TEXT_BYTES {
        return Err(invalid("numbering level text exceeds the expansion limit"));
    }
    if current.format == Format::Bullet {
        return Ok(ListMarker::Text(template.to_owned()));
    }

    let mut output = String::with_capacity(template.len());
    let mut chars = template.chars().peekable();
    while let Some(ch) = chars.next() {
        if ch == '%'
            && let Some(next @ '1'..='9') = chars.peek().copied()
        {
            chars.next();
            let referenced = next.to_digit(10).expect("digit checked") as u8 - 1;
            if referenced <= properties.level {
                let level = resolve_level(numbering, properties.num_id, referenced)?;
                let value = counters[usize::from(referenced)].unwrap_or(level.start);
                let format = if current.legal {
                    &Format::Decimal
                } else {
                    &level.format
                };
                match format_number(value, format) {
                    Some(rendered) => output.push_str(&rendered),
                    None => {
                        return Ok(ListMarker::UnsupportedFormat {
                            format: format.as_str().to_owned(),
                            value,
                        });
                    },
                }
            }
            if output.len() > MAX_LABEL_BYTES {
                return Err(invalid("expanded list label exceeds the output limit"));
            }
            continue;
        }
        output.push(ch);
        if output.len() > MAX_LABEL_BYTES {
            return Err(invalid("expanded list label exceeds the output limit"));
        }
    }
    Ok(ListMarker::Text(output))
}

fn format_number(value: i64, format: &Format) -> Option<String> {
    match format {
        Format::Decimal => Some(value.to_string()),
        Format::DecimalHalfWidth => Some(value.to_string()),
        Format::DecimalFullWidth | Format::DecimalFullWidth2 => Some(full_width_decimal(value)),
        Format::DecimalZero => {
            if (0..10).contains(&value) {
                Some(format!("0{value}"))
            } else {
                Some(value.to_string())
            }
        },
        Format::Hex => u64::try_from(value).ok().map(|value| format!("{value:X}")),
        Format::LowerLetter => letters(value, false),
        Format::UpperLetter => letters(value, true),
        Format::LowerRoman => roman(value).map(|value| value.to_ascii_lowercase()),
        Format::UpperRoman => roman(value),
        Format::None => Some(String::new()),
        _ => None,
    }
}

fn full_width_decimal(value: i64) -> String {
    let decimal = value.to_string();
    let mut output = String::with_capacity(decimal.len().saturating_mul(3));
    for character in decimal.chars() {
        match character {
            '0'..='9' => output.push(
                char::from_u32(u32::from(character) - u32::from('0') + u32::from('\u{ff10}'))
                    .expect("ASCII digit always maps to a full-width digit"),
            ),
            other => output.push(other),
        }
    }
    output
}

fn letters(value: i64, uppercase: bool) -> Option<String> {
    let mut value = u64::try_from(value).ok()?;
    if value == 0 {
        return None;
    }
    let mut output = Vec::new();
    while value != 0 {
        value -= 1;
        output.push(if uppercase { b'A' } else { b'a' } + (value % 26) as u8);
        value /= 26;
        if output.len() > 32 {
            return None;
        }
    }
    output.reverse();
    String::from_utf8(output).ok()
}

fn roman(value: i64) -> Option<String> {
    let mut value = u16::try_from(value).ok()?;
    if value == 0 || value > 3999 {
        return None;
    }
    let mut output = String::new();
    for (amount, digits) in [
        (1000, "M"),
        (900, "CM"),
        (500, "D"),
        (400, "CD"),
        (100, "C"),
        (90, "XC"),
        (50, "L"),
        (40, "XL"),
        (10, "X"),
        (9, "IX"),
        (5, "V"),
        (4, "IV"),
        (1, "I"),
    ] {
        while value >= amount {
            output.push_str(digits);
            value -= amount;
        }
    }
    Some(output)
}

fn invalid(message: &str) -> Error {
    Error::InvalidFormat(message.to_owned())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::numbering::{Definition, Instance, Override};

    fn level(index: u8, start: i64, format: Format, text: &str) -> Level {
        Level {
            level: index,
            start,
            format,
            custom_format: None,
            level_text: Some(text.to_owned()),
            suffix: Suffix::Space,
            restart: Restart::Default,
            legal: false,
            paragraph_style: None,
            picture_bullet_id: None,
        }
    }

    fn numbering(levels: Vec<Level>) -> Collection {
        Collection {
            abstract_nums: vec![Definition {
                id: 1,
                num_type: None,
                num_style_link: None,
                style_link: None,
                levels,
            }],
            nums: vec![Instance {
                id: 4,
                abstract_num_id: 1,
                overrides: Vec::new(),
            }],
            picture_bullets: Vec::new(),
        }
    }

    #[test]
    fn expands_multilevel_tokens_and_restarts_deeper_levels() {
        let numbering = numbering(vec![
            level(0, 1, Format::Decimal, "%1."),
            level(1, 1, Format::LowerLetter, "%1.%2)"),
        ]);
        let mut state = ListCounterState::new();
        assert_eq!(
            state
                .advance(
                    &numbering,
                    Paragraph {
                        num_id: 4,
                        level: 0
                    }
                )
                .unwrap()
                .0,
            ListMarker::Text("1.".to_owned())
        );
        assert_eq!(
            state
                .advance(
                    &numbering,
                    Paragraph {
                        num_id: 4,
                        level: 1
                    }
                )
                .unwrap()
                .0,
            ListMarker::Text("1.a)".to_owned())
        );
        state
            .advance(
                &numbering,
                Paragraph {
                    num_id: 4,
                    level: 0,
                },
            )
            .unwrap();
        assert_eq!(
            state
                .advance(
                    &numbering,
                    Paragraph {
                        num_id: 4,
                        level: 1
                    }
                )
                .unwrap()
                .0,
            ListMarker::Text("2.a)".to_owned())
        );
    }

    #[test]
    fn applies_start_override_without_mutating_definition() {
        let mut value = numbering(vec![level(0, 1, Format::Decimal, "%1")]);
        value.nums[0].overrides.push(Override {
            level: 0,
            start_override: Some(7),
            definition: None,
        });
        let mut state = ListCounterState::new();
        assert_eq!(
            state
                .advance(
                    &value,
                    Paragraph {
                        num_id: 4,
                        level: 0
                    }
                )
                .unwrap()
                .0,
            ListMarker::Text("7".to_owned())
        );
        assert_eq!(value.abstract_nums[0].levels[0].start, 1);
    }

    #[test]
    fn returns_typed_markers_for_unsupported_and_picture_formats() {
        let value = numbering(vec![level(0, 2, Format::NumberInDash, "%1")]);
        let mut state = ListCounterState::new();
        assert_eq!(
            state
                .advance(
                    &value,
                    Paragraph {
                        num_id: 4,
                        level: 0
                    }
                )
                .unwrap()
                .0,
            ListMarker::UnsupportedFormat {
                format: "numberInDash".to_owned(),
                value: 2
            }
        );

        let mut picture = level(0, 1, Format::Bullet, "ignored");
        picture.picture_bullet_id = Some(3);
        let value = numbering(vec![picture]);
        assert_eq!(
            ListCounterState::new()
                .advance(
                    &value,
                    Paragraph {
                        num_id: 4,
                        level: 0
                    }
                )
                .unwrap()
                .0,
            ListMarker::PictureBullet { id: 3 }
        );
    }

    #[test]
    fn renders_script_independent_extended_number_formats() {
        for (format, start, expected) in [
            (Format::Hex, 255, "FF"),
            (Format::DecimalHalfWidth, 123, "123"),
            (Format::DecimalFullWidth, 123, "１２３"),
            (Format::DecimalFullWidth2, -42, "-４２"),
        ] {
            let value = numbering(vec![level(0, start, format, "%1")]);
            assert_eq!(
                ListCounterState::new()
                    .advance(
                        &value,
                        Paragraph {
                            num_id: 4,
                            level: 0
                        }
                    )
                    .unwrap()
                    .0,
                ListMarker::Text(expected.to_owned())
            );
        }

        let value = numbering(vec![level(0, -1, Format::Hex, "%1")]);
        assert_eq!(
            ListCounterState::new()
                .advance(
                    &value,
                    Paragraph {
                        num_id: 4,
                        level: 0
                    }
                )
                .unwrap()
                .0,
            ListMarker::UnsupportedFormat {
                format: "hex".to_owned(),
                value: -1
            }
        );
    }

    #[test]
    fn rejects_counter_overflow_and_oversized_templates() {
        let value = numbering(vec![level(0, i64::MAX, Format::Decimal, "%1")]);
        let mut state = ListCounterState::new();
        state
            .advance(
                &value,
                Paragraph {
                    num_id: 4,
                    level: 0,
                },
            )
            .unwrap();
        assert!(
            state
                .advance(
                    &value,
                    Paragraph {
                        num_id: 4,
                        level: 0
                    }
                )
                .is_err()
        );

        let value = numbering(vec![level(
            0,
            1,
            Format::Decimal,
            &"x".repeat(MAX_LEVEL_TEXT_BYTES + 1),
        )]);
        assert!(
            ListCounterState::new()
                .advance(
                    &value,
                    Paragraph {
                        num_id: 4,
                        level: 0
                    }
                )
                .is_err()
        );
    }
}
