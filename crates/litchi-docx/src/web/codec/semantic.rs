//! Semantic validation and bounded resource accounting for web settings.

use super::super::MAX_XML_EVENTS;
use super::super::model::{Child, Div, Frameset, Key, Settings};
use super::super::{MAX_FRAMESET_NESTING, MAX_TEXT_BYTES, MAX_XML_BYTES, invalid};
use crate::{Error, Result};

#[derive(Default)]
struct Budget {
    nodes: usize,
    encoded_bytes: usize,
}

impl Budget {
    fn node(&mut self, description: &str) -> Result<()> {
        self.nodes = self
            .nodes
            .checked_add(1)
            .ok_or_else(|| invalid("web-settings node count overflow"))?;
        if self.nodes > MAX_XML_EVENTS / 2 {
            return Err(invalid(format!(
                "web-settings {description} exceeds the node limit"
            )));
        }
        self.bytes(256, description)
    }

    fn text(&mut self, value: &str, description: &str) -> Result<()> {
        validate_text(value, description, true)?;
        let escaped = value
            .len()
            .checked_mul(6)
            .ok_or_else(|| invalid("web-settings escaped text size overflow"))?;
        self.bytes(escaped, description)
    }

    fn bytes(&mut self, bytes: usize, description: &str) -> Result<()> {
        self.encoded_bytes = self
            .encoded_bytes
            .checked_add(bytes)
            .ok_or_else(|| invalid("web-settings output size overflow"))?;
        if self.encoded_bytes > MAX_XML_BYTES {
            return Err(invalid(format!(
                "web-settings {description} exceeds {MAX_XML_BYTES} output bytes"
            )));
        }
        Ok(())
    }
}

pub(crate) fn validate_value(value: &Settings) -> Result<usize> {
    let mut budget = Budget {
        nodes: 1,
        encoded_bytes: 512,
    };
    if let Some(encoding) = &value.encoding {
        validate_encoding(encoding)?;
        budget.node("encoding")?;
        budget.text(encoding, "encoding")?;
    }
    if let Some(pixels) = value.pixels_per_inch {
        validate_pixels_per_inch(pixels)?;
        budget.node("pixels-per-inch")?;
    }
    if let Some(frameset) = &value.frameset {
        validate_frameset(frameset, 1, &mut budget)?;
    }
    if let Some(divs) = &value.divs {
        if divs.is_empty() {
            return Err(invalid("Word HTML division container must not be empty"));
        }
        budget.node("division container")?;
        validate_div_slice(divs, 1, &mut budget)?;
    }
    for present in [
        value.optimize_for_browser,
        value.rely_on_vml,
        value.allow_png,
        value.do_not_rely_on_css,
        value.do_not_save_as_single_file,
        value.do_not_organize_in_folder,
        value.do_not_use_long_file_names,
        value.target_screen_size.map(|_| true),
        value.save_smart_tags_as_xml,
    ] {
        if present.is_some() {
            budget.node("scalar setting")?;
        }
    }
    Ok(budget.encoded_bytes)
}

fn validate_frameset(value: &Frameset, depth: usize, budget: &mut Budget) -> Result<()> {
    if depth > MAX_FRAMESET_NESTING {
        return Err(invalid("web frameset nesting exceeds the safety limit"));
    }
    budget.node("frameset")?;
    if let Some(size) = &value.size {
        budget.node("frameset size")?;
        budget.text(size, "frameset size")?;
    }
    if let Some(split) = &value.split_bar {
        budget.node("frameset split bar")?;
        if let Some(color) = &split.color {
            validate_color(&color.value, "frameset splitter color")?;
            budget.node("frameset splitter color")?;
            budget.text(&color.value, "frameset splitter color")?;
        }
    }
    if value.layout.is_some() {
        budget.node("frameset layout")?;
    }
    for child in &value.children {
        match child {
            Child::Frameset(nested) => validate_frameset(nested, depth + 1, budget)?,
            Child::Frame(frame) => {
                budget.node("frame")?;
                for (text, description) in [
                    (frame.size.as_deref(), "frame size"),
                    (frame.name.as_deref(), "frame name"),
                    (
                        frame.source_file_relationship_id.as_deref(),
                        "frame relationship ID",
                    ),
                ] {
                    if let Some(text) = text {
                        if description == "frame relationship ID" {
                            validate_relationship_id(text)?;
                        }
                        budget.node(description)?;
                        budget.text(text, description)?;
                    }
                }
            },
        }
    }
    Ok(())
}

pub(crate) fn validate_divs(divs: &[Div], depth: usize) -> Result<()> {
    if divs.is_empty() {
        return Err(invalid("Word HTML division container must not be empty"));
    }
    let mut budget = Budget::default();
    validate_div_slice(divs, depth, &mut budget)
}

fn validate_div_slice(divs: &[Div], depth: usize, budget: &mut Budget) -> Result<()> {
    if depth > MAX_FRAMESET_NESTING {
        return Err(invalid("HTML division nesting exceeds the safety limit"));
    }
    let mut ids = std::collections::HashSet::new();
    ids.try_reserve(divs.len())
        .map_err(|source| Error::Allocation {
            resource: "HTML division identifier index",
            source,
        })?;
    for div in divs {
        if !ids.insert(div.id) {
            return Err(invalid(format!("HTML division '{}' is ambiguous", div.id)));
        }
        budget.node("HTML division")?;
        budget.bytes(decimal_len(div.id.get()), "HTML division ID")?;
        for margin in [div.left, div.right, div.top, div.bottom] {
            budget.node("HTML division margin")?;
            budget.bytes(decimal_len(margin.get()), "HTML division margin")?;
        }
        if let Some(borders) = &div.borders {
            budget.node("HTML division borders")?;
            for border in [&borders.top, &borders.left, &borders.bottom, &borders.right]
                .into_iter()
                .flatten()
            {
                validate_border_style(&border.style)?;
                budget.node("HTML division border")?;
                budget.text(&border.style, "HTML division border style")?;
                if let Some(color) = &border.color {
                    validate_color(color, "HTML division border color")?;
                    budget.text(color, "HTML division border color")?;
                }
            }
        }
        if !div.children.is_empty() {
            validate_div_slice(&div.children, depth + 1, budget)?;
        }
    }
    Ok(())
}

pub(crate) fn div_position(divs: &[Div], key: Key) -> Result<Option<usize>> {
    match key {
        Key::Index(index) => {
            if index >= divs.len() {
                Err(invalid(format!(
                    "HTML division position {index} is outside 0..{}",
                    divs.len()
                )))
            } else {
                Ok(Some(index))
            }
        },
        Key::Id(id) => {
            let mut matches = divs
                .iter()
                .enumerate()
                .filter(|(_, div)| div.id == id)
                .map(|(index, _)| index);
            let first = matches.next();
            if first.is_some() && matches.next().is_some() {
                Err(invalid(format!("HTML division ID '{id}' is ambiguous")))
            } else {
                Ok(first)
            }
        },
    }
}

pub(crate) fn validate_text(value: &str, description: &str, allow_empty: bool) -> Result<()> {
    if (!allow_empty && value.is_empty()) || value.len() > MAX_TEXT_BYTES {
        return Err(invalid(format!("invalid {description} length")));
    }
    if value
        .chars()
        .any(|character| character.is_control() && !matches!(character, '\t' | '\n' | '\r'))
    {
        return Err(invalid(format!(
            "{description} contains a control character"
        )));
    }
    Ok(())
}

pub(crate) fn validate_encoding(value: &str) -> Result<()> {
    if value.is_empty()
        || value.len() > 255
        || !value
            .bytes()
            .all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'-' | b'_' | b'.'))
    {
        return Err(invalid("web encoding is not a bounded character-set name"));
    }
    Ok(())
}

pub(crate) fn validate_relationship_id(value: &str) -> Result<()> {
    if value.is_empty()
        || value.len() > 255
        || !value
            .bytes()
            .all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'_' | b'-' | b'.' | b':'))
    {
        return Err(invalid(
            "frame relationship ID is not a safe XML identifier",
        ));
    }
    Ok(())
}

pub(crate) fn validate_pixels_per_inch(value: u16) -> Result<()> {
    if value <= 1023 {
        Ok(())
    } else {
        Err(invalid("pixels-per-inch must be in the range 0..=1023"))
    }
}

pub(crate) fn parse_i64(value: &str, description: &str) -> Result<i64> {
    let value = value.trim();
    if value.is_empty() || value.len() > MAX_TEXT_BYTES {
        return Err(invalid(format!("invalid {description} length")));
    }
    value
        .parse::<i64>()
        .map_err(|_| invalid(format!("invalid {description} value '{value}'")))
}

fn decimal_len(value: i64) -> usize {
    let magnitude = value.unsigned_abs();
    let digits = if magnitude == 0 {
        1
    } else {
        magnitude.ilog10() as usize + 1
    };
    digits + usize::from(value.is_negative())
}

pub(crate) fn validate_border_style(value: &str) -> Result<()> {
    if value.is_empty()
        || value.len() > 64
        || !value.bytes().all(|byte| byte.is_ascii_alphanumeric())
    {
        return Err(invalid("HTML division border style is not a schema token"));
    }
    Ok(())
}

fn validate_color(value: &str, description: &str) -> Result<()> {
    if value.eq_ignore_ascii_case("auto")
        || (value.len() == 6 && value.bytes().all(|byte| byte.is_ascii_hexdigit()))
    {
        Ok(())
    } else {
        Err(invalid(format!(
            "invalid {description} '{value}'; expected auto or six hexadecimal digits"
        )))
    }
}

pub(crate) fn validate_word_color(value: String, description: &str) -> Result<String> {
    if value != "auto" && (value.len() != 6 || !value.bytes().all(|byte| byte.is_ascii_hexdigit()))
    {
        return Err(Error::Invalid(format!("invalid {description} '{value}'")));
    }
    Ok(value)
}
