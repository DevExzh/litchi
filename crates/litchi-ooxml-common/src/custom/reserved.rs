//! Validation for the reserved custom-property names in MS-OI29500 3.11.
//!
//! The specification only gives an opaque persistence allowance for Microsoft
//! Information Protection SDK metadata. In particular, the `MSIP_Label_*`
//! samples do not define a property-name grammar, value type, or date format,
//! so they deliberately remain ordinary custom properties here.

use std::cmp::Ordering;

use super::model::{Props, Value};
use super::schema::invalid;
use crate::Result;

const HEADER_FONT: &str = "ClassificationContentMarkingHeaderFontProps";
const HEADER_SHAPES: &str = "ClassificationContentMarkingHeaderShapeIds";
const HEADER_TEXT: &str = "ClassificationContentMarkingHeaderText";
const FOOTER_FONT: &str = "ClassificationContentMarkingFooterFontProps";
const FOOTER_SHAPES: &str = "ClassificationContentMarkingFooterShapeIds";
const FOOTER_TEXT: &str = "ClassificationContentMarkingFooterText";
const WATERMARK_FONT: &str = "ClassificationWatermarkFontProps";
const WATERMARK_SHAPES: &str = "ClassificationWatermarkShapeIds";
const WATERMARK_TEXT: &str = "ClassificationWatermarkText";
const HEADER_LOCATIONS: &str = "ClassificationContentMarkingHeaderLocations";
const FOOTER_LOCATIONS: &str = "ClassificationContentMarkingFooterLocations";
const WATERMARK_LOCATIONS: &str = "ClassificationWatermarkLocations";
const SHAPE_PROPERTIES: [&str; 3] = [HEADER_SHAPES, FOOTER_SHAPES, WATERMARK_SHAPES];

/// OOXML host that owns sensitivity-label rendering properties.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Host {
    /// Excel workbooks.
    Excel,
    /// Word documents.
    Word,
    /// PowerPoint presentations.
    PowerPoint,
}

#[derive(Clone, Copy)]
enum Scope {
    Any,
    Word,
    PowerPoint,
    WordOrPowerPoint,
}

pub(super) fn validate_property(name: &str, value: &Value, host: Option<Host>) -> Result<()> {
    if name.eq_ignore_ascii_case("Sensitivity") {
        validate_scope(name, Scope::Any, host)?;
        return validate_guid(name, value);
    }

    if let Some((_, suffix)) = shape_fragment(name) {
        validate_shape_fragment_name(name, suffix)?;
    }

    let Some((scope, format)) = classify(name) else {
        return Ok(());
    };
    validate_scope(name, scope, host)?;
    let text = text_value(name, value)?;
    match format {
        Format::Text => Ok(()),
        Format::Font => validate_font(name, text),
        Format::ShapeIds => validate_shape_ids(name, text),
        Format::Locations => validate_locations(name, text),
    }
}

pub(super) fn validate(properties: &Props, host: Option<Host>) -> Result<()> {
    for (name, property) in &properties.properties {
        validate_property(name, &property.value, host)?;
    }
    for base in SHAPE_PROPERTIES {
        validate_shape_fragments(properties, base)?;
    }
    Ok(())
}

#[derive(Clone, Copy)]
enum Format {
    Text,
    Font,
    ShapeIds,
    Locations,
}

fn classify(name: &str) -> Option<(Scope, Format)> {
    if matches_any(name, &[HEADER_FONT, FOOTER_FONT, WATERMARK_FONT]) {
        Some((Scope::Word, Format::Font))
    } else if matches_any(name, &[HEADER_SHAPES, FOOTER_SHAPES, WATERMARK_SHAPES]) {
        Some((Scope::Word, Format::ShapeIds))
    } else if matches_any(name, &[HEADER_TEXT, FOOTER_TEXT, WATERMARK_TEXT]) {
        Some((Scope::WordOrPowerPoint, Format::Text))
    } else if matches_any(
        name,
        &[HEADER_LOCATIONS, FOOTER_LOCATIONS, WATERMARK_LOCATIONS],
    ) {
        Some((Scope::PowerPoint, Format::Locations))
    } else if shape_fragment(name).is_some() {
        Some((Scope::Word, Format::ShapeIds))
    } else {
        None
    }
}

fn validate_scope(name: &str, scope: Scope, host: Option<Host>) -> Result<()> {
    let Some(host) = host else {
        return Ok(());
    };
    let permitted = match scope {
        Scope::Any => true,
        Scope::Word => host == Host::Word,
        Scope::PowerPoint => host == Host::PowerPoint,
        Scope::WordOrPowerPoint => matches!(host, Host::Word | Host::PowerPoint),
    };
    if permitted {
        Ok(())
    } else {
        Err(invalid(format!(
            "reserved custom property '{name}' is not valid for {host:?}"
        )))
    }
}

fn validate_guid(name: &str, value: &Value) -> Result<()> {
    let text = text_value(name, value)?;
    let valid = text.len() == 36
        && text.bytes().enumerate().all(|(index, byte)| {
            matches!(index, 8 | 13 | 18 | 23) && byte == b'-'
                || !matches!(index, 8 | 13 | 18 | 23) && byte.is_ascii_hexdigit()
        });
    if valid {
        Ok(())
    } else {
        Err(invalid(format!(
            "reserved custom property '{name}' must be a hyphenated GUID"
        )))
    }
}

fn text_value<'a>(name: &str, value: &'a Value) -> Result<&'a str> {
    match value {
        Value::Text(text) => Ok(text),
        _ => Err(invalid(format!(
            "reserved custom property '{name}' must use a text value"
        ))),
    }
}

fn validate_font(name: &str, value: &str) -> Result<()> {
    let mut fields = value.splitn(3, ',');
    let color = fields
        .next()
        .expect("splitn always returns the first field");
    let points = fields.next();
    let face = fields.next();
    let valid_color = color.len() == 7
        && color.starts_with('#')
        && color.as_bytes()[1..].iter().all(u8::is_ascii_hexdigit);
    let valid_points = points
        .and_then(|field| field.parse::<f64>().ok())
        .is_some_and(f64::is_finite);
    if valid_color && valid_points && face.is_some_and(|field| !field.is_empty()) {
        Ok(())
    } else {
        Err(invalid(format!(
            "reserved custom property '{name}' must be '#RRGGBB,points,font face'"
        )))
    }
}

fn validate_shape_ids(name: &str, value: &str) -> Result<()> {
    if !value.is_empty()
        && value.split(',').all(|identifier| {
            !identifier.is_empty() && identifier.bytes().all(|byte| byte.is_ascii_hexdigit())
        })
    {
        Ok(())
    } else {
        Err(invalid(format!(
            "reserved custom property '{name}' must be a comma-separated hexadecimal shape-ID list"
        )))
    }
}

fn validate_shape_fragment_name(name: &str, suffix: &str) -> Result<()> {
    if !suffix.is_empty() && suffix.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        Ok(())
    } else {
        Err(invalid(format!(
            "reserved custom-property fragment '{name}' must have a hexadecimal suffix"
        )))
    }
}

fn validate_locations(name: &str, value: &str) -> Result<()> {
    if value.is_empty() {
        return Err(invalid(format!(
            "reserved custom property '{name}' cannot be empty"
        )));
    }

    let mut characters = value.chars();
    let mut in_shape_id = false;
    let mut name_chars = 0usize;
    let mut shape_id_chars = 0usize;
    while let Some(character) = characters.next() {
        match character {
            ':' if !in_shape_id && name_chars != 0 => in_shape_id = true,
            ':' => {
                return Err(invalid(format!(
                    "reserved custom property '{name}' has an invalid location"
                )));
            },
            '\\' if in_shape_id => {
                if shape_id_chars == 0 {
                    return Err(invalid(format!(
                        "reserved custom property '{name}' has an empty shape ID"
                    )));
                }
                in_shape_id = false;
                name_chars = 0;
                shape_id_chars = 0;
            },
            '\\' => match characters.next() {
                Some(':' | '\\') => name_chars += 1,
                _ => {
                    return Err(invalid(format!(
                        "reserved custom property '{name}' has an invalid location escape"
                    )));
                },
            },
            _ if in_shape_id => shape_id_chars += 1,
            _ => name_chars += 1,
        }
    }
    if in_shape_id && shape_id_chars != 0 {
        Ok(())
    } else {
        Err(invalid(format!(
            "reserved custom property '{name}' has an incomplete location"
        )))
    }
}

fn validate_shape_fragments(properties: &Props, base: &str) -> Result<()> {
    let mut suffixes = properties
        .properties
        .keys()
        .filter_map(|name| shape_fragment_for(name, base))
        .collect::<Vec<_>>();
    if suffixes.is_empty() {
        return Ok(());
    }
    if !properties.contains(base) {
        return Err(invalid(format!(
            "reserved custom-property fragment for '{base}' requires the base property"
        )));
    }
    suffixes.sort_unstable_by(compare_hex);
    for (index, suffix) in suffixes.into_iter().enumerate() {
        let expected = format!("{:x}", index + 1);
        if !suffix.eq_ignore_ascii_case(&expected) {
            return Err(invalid(format!(
                "reserved custom-property fragments for '{base}' must start at 1 and increment by 1"
            )));
        }
    }
    Ok(())
}

fn shape_fragment(name: &str) -> Option<(&'static str, &str)> {
    SHAPE_PROPERTIES
        .into_iter()
        .find_map(|base| shape_fragment_for(name, base).map(|suffix| (base, suffix)))
}

fn shape_fragment_for<'a>(name: &'a str, base: &str) -> Option<&'a str> {
    let prefix = name.get(..base.len())?;
    if prefix.eq_ignore_ascii_case(base) {
        name.get(base.len()..)?.strip_prefix('-')
    } else {
        None
    }
}

fn matches_any(name: &str, candidates: &[&str]) -> bool {
    candidates
        .iter()
        .any(|candidate| name.eq_ignore_ascii_case(candidate))
}

fn compare_hex(left: &&str, right: &&str) -> Ordering {
    left.len().cmp(&right.len()).then_with(|| {
        left.bytes()
            .map(|byte| byte.to_ascii_lowercase())
            .cmp(right.bytes().map(|byte| byte.to_ascii_lowercase()))
    })
}
