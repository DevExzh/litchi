//! Lexical validation for the generated graphic-property specifications.

use litchi_core::{Error, Result};

use super::model::{Kind, Value};
use crate::graphic_properties::MAX_VALUE;

pub(crate) fn bad(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

pub(crate) fn safe(value: &str, name: &str, empty: bool) -> Result<()> {
    if (!empty && value.is_empty())
        || value.len() > MAX_VALUE
        || value.chars().any(
            |c| matches!(c,'\0'..='\u{8}'|'\u{b}'|'\u{c}'|'\u{e}'..='\u{1f}'|'\u{fffe}'|'\u{ffff}'),
        )
    {
        return Err(bad(format!("invalid {name}")));
    }
    Ok(())
}

fn decimal(value: &str, signed: bool) -> bool {
    let value = if signed {
        value.strip_prefix('-').unwrap_or(value)
    } else {
        value
    };
    if value.is_empty() {
        return false;
    }
    let mut parts = value.split('.');
    let left = parts.next().unwrap_or_default();
    let right = parts.next();
    if parts.next().is_some() {
        return false;
    }
    match right {
        None => !left.is_empty() && left.bytes().all(|b| b.is_ascii_digit()),
        Some(right) => {
            (!left.is_empty() || !right.is_empty())
                && left.bytes().all(|b| b.is_ascii_digit())
                && right.bytes().all(|b| b.is_ascii_digit())
        },
    }
}

fn length(value: &str, signed: bool, positive: bool, pixel_only: bool) -> bool {
    let units = if pixel_only {
        &["px"][..]
    } else {
        &["cm", "mm", "in", "pt", "pc", "px"][..]
    };
    units.iter().any(|unit| {
        value.strip_suffix(unit).is_some_and(|number| {
            decimal(number, signed)
                && (!positive || number.bytes().any(|b| b.is_ascii_digit() && b != b'0'))
        })
    })
}

fn percent(value: &str, signed: bool, ranged: bool) -> bool {
    let Some(number) = value.strip_suffix('%') else {
        return false;
    };
    if !decimal(number, signed) {
        return false;
    }
    if !ranged {
        return true;
    }
    number
        .trim_start_matches('-')
        .parse::<f64>()
        .is_ok_and(|number| number <= 100.0)
}

fn integer(value: &str, positive: bool) -> bool {
    let value = value.strip_prefix('+').unwrap_or(value);
    if positive {
        !value.is_empty()
            && value.bytes().all(|b| b.is_ascii_digit())
            && value.bytes().any(|b| b != b'0')
    } else {
        !value.is_empty() && value.bytes().all(|b| b.is_ascii_digit())
    }
}

fn duration(value: &str) -> bool {
    let value = value.strip_prefix('-').unwrap_or(value);
    value.starts_with('P')
        && value.len() > 1
        && value.bytes().any(|b| b.is_ascii_digit())
        && value.bytes().all(|b| {
            b.is_ascii_digit() || matches!(b, b'P' | b'T' | b'Y' | b'M' | b'D' | b'H' | b'S' | b'.')
        })
}

pub(crate) fn ncname(value: &str, empty: bool) -> bool {
    if value.is_empty() {
        return empty;
    }
    let mut chars = value.chars();
    let Some(first) = chars.next() else {
        return false;
    };
    (first == '_' || first.is_alphabetic())
        && chars.all(|c| c == '_' || c == '-' || c == '.' || c.is_alphanumeric())
}

fn clip(value: &str) -> bool {
    let Some(inner) = value
        .strip_prefix("rect(")
        .and_then(|v| v.strip_suffix(')'))
    else {
        return false;
    };
    let values: Vec<_> = inner.split(',').map(str::trim).collect();
    values.len() == 4
        && values.iter().all(|value| {
            *value == "auto"
                || ["cm", "mm", "in", "pt", "pc"].iter().any(|unit| {
                    value
                        .strip_suffix(unit)
                        .is_some_and(|number| decimal(number, true))
                })
        })
}

fn validate_ref(reference: &str, value: &str) -> Option<Value> {
    match reference {
        "boolean" => match value {
            "true" => Some(Value::Boolean(true)),
            "false" => Some(Value::Boolean(false)),
            _ => None,
        },
        "color" => (value.len() == 7
            && value.starts_with('#')
            && value[1..].bytes().all(|b| b.is_ascii_hexdigit()))
        .then(|| Value::Color(value.to_owned())),
        "length" | "coordinate" | "distance" => {
            length(value, true, false, false).then(|| Value::Length(value.to_owned()))
        },
        "nonNegativeLength" => {
            length(value, false, false, false).then(|| Value::NonNegativeLength(value.to_owned()))
        },
        "positiveLength" => {
            length(value, false, true, false).then(|| Value::PositiveLength(value.to_owned()))
        },
        "nonNegativePixelLength" => length(value, false, false, true)
            .then(|| Value::NonNegativePixelLength(value.to_owned())),
        "percent" => percent(value, true, false).then(|| Value::Percent(value.to_owned())),
        "zeroToHundredPercent" => {
            percent(value, false, true).then(|| Value::ZeroToHundredPercent(value.to_owned()))
        },
        "signedZeroToHundredPercent" => {
            percent(value, true, true).then(|| Value::SignedZeroToHundredPercent(value.to_owned()))
        },
        "nonNegativeInteger" => {
            integer(value, false).then(|| Value::NonNegativeInteger(value.to_owned()))
        },
        "positiveInteger" => integer(value, true).then(|| Value::PositiveInteger(value.to_owned())),
        "duration" => duration(value).then(|| Value::Duration(value.to_owned())),
        "clipShape" => clip(value).then(|| Value::Clip(value.to_owned())),
        "styleNameRef" => ncname(value, true).then(|| Value::StyleNameRef(value.to_owned())),
        "styleNameRefs" => {
            let values: Vec<_> = value.split_ascii_whitespace().map(str::to_owned).collect();
            values
                .iter()
                .all(|value| ncname(value, false))
                .then_some(Value::StyleNameRefs(values))
        },
        "horizontal-mirror" => matches!(
            value,
            "horizontal" | "horizontal-on-odd" | "horizontal-on-even"
        )
        .then(|| Value::Keyword(value.to_owned())),
        "borderWidths" => {
            let values: Vec<_> = value.split_ascii_whitespace().collect();
            (values.len() == 3 && values.iter().all(|value| length(value, false, true, false)))
                .then(|| Value::Compound(value.to_owned()))
        },
        "angle" | "string" | "shadowType" => Some(Value::Text(value.to_owned())),
        _ => None,
    }
}

pub(crate) fn validate_spec(
    value: &str,
    keywords: &[&str],
    references: &[&str],
    list: bool,
    kind: Kind,
) -> Result<Value> {
    safe(value, "graphic property value", true)?;
    if kind == Kind::DrawTileRepeatOffset {
        let parts: Vec<_> = value.split_ascii_whitespace().collect();
        if parts.len() == 2
            && percent(parts[0], false, true)
            && matches!(parts[1], "horizontal" | "vertical")
        {
            return Ok(Value::Compound(value.to_owned()));
        }
        return Err(bad("invalid draw:tile-repeat-offset"));
    }
    if list {
        let parts: Vec<_> = value.split_ascii_whitespace().collect();
        if parts.iter().all(|part| {
            keywords.contains(part)
                || references
                    .iter()
                    .any(|reference| validate_ref(reference, part).is_some())
        }) {
            return Ok(Value::Compound(value.to_owned()));
        }
        return Err(bad(format!(
            "invalid {}:{} list",
            kind.namespace().prefix(),
            kind.local_name()
        )));
    }
    if keywords.contains(&value) {
        return Ok(Value::Keyword(value.to_owned()));
    }
    for reference in references {
        if let Some(value) = validate_ref(reference, value) {
            return Ok(value);
        }
    }
    Err(bad(format!(
        "invalid {}:{} value",
        kind.namespace().prefix(),
        kind.local_name()
    )))
}
