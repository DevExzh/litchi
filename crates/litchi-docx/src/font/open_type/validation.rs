//! Bounded semantic validation for Word 2010 OpenType extensions.

use crate::error::Result;

use super::model::{OpenType, StyleSet};

pub(crate) fn validate(value: &OpenType) -> Result<()> {
    for style_set in &value.stylistic_sets {
        validate_style_set(style_set)?;
    }
    for (index, left) in value.stylistic_sets.iter().enumerate() {
        if value.stylistic_sets[index + 1..]
            .iter()
            .any(|right| right.id == left.id)
        {
            return Err(super::model::invalid(format!(
                "duplicate OpenType stylistic set id {}",
                left.id.get()
            )));
        }
    }
    Ok(())
}

pub(crate) fn validate_style_set(value: &StyleSet) -> Result<()> {
    if !(1..=20).contains(&value.id.get()) {
        return Err(super::model::invalid(format!(
            "stylistic set id {} is outside 1..=20",
            value.id.get()
        )));
    }
    Ok(())
}

pub(crate) const MAX_XML_BYTES: usize = 4 * 1024 * 1024;
pub(crate) const MAX_XML_DEPTH: usize = 128;
pub(crate) const MAX_XML_NODES: usize = 65_536;
pub(crate) const MAX_STYLE_SETS: usize = 20;
