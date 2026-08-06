//! Safety and semantic validation for handout masters.

use super::Master;
use litchi_core::{Error, Result};
use litchi_odf_common::style::master::ChildKind;

pub(crate) const MAX_XML_BYTES: usize = 64 * 1_048_576;
pub(crate) const MAX_FRAGMENT_BYTES: usize = 16 * 1_048_576;
pub(crate) const MAX_CHILDREN: usize = 65_536;
pub(crate) const MAX_TOTAL_CHILD_BYTES: usize = 32 * 1_048_576;
pub(crate) const MAX_NAME_BYTES: usize = 4_096;

pub(crate) fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

pub(crate) fn validate(master: &Master) -> Result<()> {
    validate_fields(master)?;
    for child in &master.children {
        super::codec::validate_shape_fragment(&child.xml)?;
    }
    Ok(())
}

pub(crate) fn validate_fields(master: &Master) -> Result<()> {
    validate_name(&master.page_layout_name, "handout page layout name")?;
    for (value, label) in [
        (
            master.presentation_page_layout_name.as_deref(),
            "handout presentation page layout name",
        ),
        (
            master.drawing_style_name.as_deref(),
            "handout drawing style name",
        ),
        (
            master.header_name.as_deref(),
            "handout header declaration name",
        ),
        (
            master.footer_name.as_deref(),
            "handout footer declaration name",
        ),
        (
            master.date_time_name.as_deref(),
            "handout date-time declaration name",
        ),
    ] {
        if let Some(value) = value {
            validate_name(value, label)?;
        }
    }
    if master.children.len() > MAX_CHILDREN {
        return Err(invalid("handout master has too many drawing children"));
    }
    let mut total = 0usize;
    for child in &master.children {
        if child.kind != ChildKind::Shape {
            return Err(invalid("handout master contains a non-drawing child"));
        }
        if child.xml.len() > MAX_FRAGMENT_BYTES {
            return Err(invalid("handout drawing child exceeds 16 MiB"));
        }
        total = total
            .checked_add(child.xml.len())
            .ok_or_else(|| invalid("handout drawing child size overflows"))?;
        if total > MAX_TOTAL_CHILD_BYTES {
            return Err(invalid("handout drawing children exceed 32 MiB"));
        }
    }
    Ok(())
}

pub(crate) fn validate_name(value: &str, context: &str) -> Result<()> {
    if value.is_empty() || value.len() > MAX_NAME_BYTES {
        return Err(invalid(format!(
            "{context} must contain 1..={MAX_NAME_BYTES} bytes"
        )));
    }
    if value.chars().any(
        |character| matches!(character, '\0'..='\u{8}' | '\u{b}' | '\u{c}' | '\u{e}'..='\u{1f}'),
    ) {
        return Err(invalid(format!(
            "{context} contains an XML control character"
        )));
    }
    if !is_ncname(value) {
        return Err(invalid(format!("{context} is not an XML NCName")));
    }
    Ok(())
}

fn is_ncname(value: &str) -> bool {
    let mut chars = value.chars();
    let Some(first) = chars.next() else {
        return false;
    };
    is_ncname_start(first) && chars.all(is_ncname_continue)
}

fn is_ncname_start(value: char) -> bool {
    value == '_'
        || value.is_ascii_alphabetic()
        || matches!(value, '\u{c0}'..='\u{d6}' | '\u{d8}'..='\u{f6}' | '\u{f8}'..='\u{2ff}' | '\u{370}'..='\u{37d}' | '\u{37f}'..='\u{1fff}' | '\u{200c}'..='\u{200d}' | '\u{2070}'..='\u{218f}' | '\u{2c00}'..='\u{2fef}' | '\u{3001}'..='\u{d7ff}' | '\u{f900}'..='\u{fdcf}' | '\u{fdf0}'..='\u{fffd}' | '\u{10000}'..='\u{effff}')
}

fn is_ncname_continue(value: char) -> bool {
    is_ncname_start(value)
        || matches!(value, '-' | '.' | '0'..='9' | '\u{b7}' | '\u{300}'..='\u{36f}' | '\u{203f}'..='\u{2040}')
}
