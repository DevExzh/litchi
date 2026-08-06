//! Structural and resource validation for table-template semantics.

use litchi_core::{Error, Result};

use super::semantic::{Region, Template};

pub(super) const MAX_TEMPLATES: usize = 1_000_000;
pub(super) const MAX_VALUE_BYTES: usize = 65_536;
pub(super) const MAX_AGGREGATE_BYTES: usize = 16 * 1_048_576;
pub(super) const MAX_EXTENSION_DEPTH: usize = 256;

/// Validate a template's required band structure and style references.
impl Template {
    pub fn validate(&self) -> Result<()> {
        validate_template_value(&self.name, "table template name")?;
        if self.even_rows.is_some() != self.odd_rows.is_some() {
            return Err(Error::InvalidFormat(
                "table template requires both even-rows and odd-rows".to_string(),
            ));
        }
        if self.even_columns.is_some() != self.odd_columns.is_some() {
            return Err(Error::InvalidFormat(
                "table template requires both even-columns and odd-columns".to_string(),
            ));
        }
        if self.body.is_none() && self.even_rows.is_none() && self.even_columns.is_none() {
            return Err(Error::InvalidFormat(
                "table template requires body or a complete row/column band pair".to_string(),
            ));
        }
        for region in Region::ALL {
            let Some(style) = self.region(region) else {
                continue;
            };
            validate_template_value(style.style_name.as_str(), "table template style name")?;
            if let Some(paragraph) = &style.paragraph_style_name {
                if region == Region::Background {
                    return Err(Error::InvalidFormat(
                        "table:background cannot have a paragraph style".to_string(),
                    ));
                }
                validate_template_value(paragraph, "table template paragraph style name")?;
            }
        }
        Ok(())
    }
}

pub(super) fn validate_template_value(value: &str, name: &str) -> Result<()> {
    if value.is_empty() {
        return Err(Error::InvalidFormat(format!("{name} must not be empty")));
    }
    if value.len() > MAX_VALUE_BYTES {
        return Err(Error::InvalidFormat(format!("{name} exceeds 64 KiB")));
    }
    Ok(())
}
