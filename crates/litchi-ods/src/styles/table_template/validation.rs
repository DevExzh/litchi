//! Structural and resource validation for table-template semantics.

use litchi_core::{Error, Result};

use super::semantic::{Region, Template};

pub(super) const MAX_TEMPLATES: usize = 1_000_000;
pub(super) const MAX_VALUE_BYTES: usize = 65_536;
pub(super) const MAX_AGGREGATE_BYTES: usize = 16 * 1_048_576;
pub(super) const MAX_EXTENSION_DEPTH: usize = 256;
pub(super) const MAX_XML_DEPTH: usize = 1_024;

/// Validate a template's required band structure and style references.
impl Template {
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn validate(&self) -> Result<()> {
        validate_template_value(&self.name, "table template name")?;
        if self.first_row_start_column.is_none()
            || self.first_row_end_column.is_none()
            || self.last_row_start_column.is_none()
            || self.last_row_end_column.is_none()
        {
            return Err(Error::InvalidFormat(
                "table template requires all four row edge selectors".to_string(),
            ));
        }
        if self.body.is_none() {
            return Err(Error::InvalidFormat(
                "table template requires a body style".to_string(),
            ));
        }
        for region in Region::ALL {
            let Some(style) = self.region(region) else {
                continue;
            };
            validate_style_name_ref(style.style_name.as_str(), "table template style name")?;
            if let Some(paragraph) = &style.paragraph_style_name {
                if region == Region::Background {
                    return Err(Error::InvalidFormat(
                        "table:background cannot have a paragraph style".to_string(),
                    ));
                }
                validate_style_name_ref(paragraph, "table template paragraph style name")?;
            }
        }
        Ok(())
    }
}

pub(super) fn validate_template_value(value: &str, name: &str) -> Result<()> {
    if value.len() > MAX_VALUE_BYTES {
        return Err(Error::InvalidFormat(format!("{name} exceeds 64 KiB")));
    }
    if value
        .bytes()
        .any(|byte| byte < 0x20 && !matches!(byte, 0x09 | 0x0A | 0x0D))
    {
        return Err(Error::InvalidFormat(format!(
            "{name} contains an XML-forbidden control character"
        )));
    }
    Ok(())
}

/// Validate the ODF `styleNameRef` datatype: either the empty string or an
/// XML `NCName`.  The schema deliberately permits an empty style reference,
/// while a non-empty reference must reject spaces, colons, and other names
/// outside the XML name character classes.
pub(super) fn validate_style_name_ref(value: &str, name: &str) -> Result<()> {
    validate_template_value(value, name)?;
    if value.is_empty() {
        return Ok(());
    }
    let mut characters = value.chars();
    let Some(first) = characters.next() else {
        return Ok(());
    };
    if !is_nc_name_start(first) || !characters.all(is_nc_name_char) {
        return Err(Error::InvalidFormat(format!(
            "{name} must be an XML NCName or empty"
        )));
    }
    Ok(())
}

fn is_nc_name_start(value: char) -> bool {
    matches!(
        value,
        'A'..='Z'
            | '_'
            | 'a'..='z'
            | '\u{00C0}'..='\u{00D6}'
            | '\u{00D8}'..='\u{00F6}'
            | '\u{00F8}'..='\u{02FF}'
            | '\u{0370}'..='\u{037D}'
            | '\u{037F}'..='\u{1FFF}'
            | '\u{200C}'..='\u{200D}'
            | '\u{2070}'..='\u{218F}'
            | '\u{2C00}'..='\u{2FEF}'
            | '\u{3001}'..='\u{D7FF}'
            | '\u{F900}'..='\u{FDCF}'
            | '\u{FDF0}'..='\u{FFFD}'
            | '\u{10000}'..='\u{EFFFF}'
    )
}

fn is_nc_name_char(value: char) -> bool {
    is_nc_name_start(value)
        || matches!(
            value,
            '-' | '.'
                | '0'..='9'
                | '\u{00B7}'
                | '\u{0300}'..='\u{036F}'
                | '\u{203F}'..='\u{2040}'
        )
}
