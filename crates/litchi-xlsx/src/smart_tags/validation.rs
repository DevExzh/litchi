use std::collections::HashSet;

use super::model::{Cell, Collection, Property, Tag};
use super::{MAX_CELLS, MAX_PROPERTIES, MAX_TAGS, MAX_TEXT_BYTES};
use crate::error::{Result, invalid};

/// Validate a complete worksheet smart-tag collection.
pub fn collection(value: &Collection) -> Result<()> {
    if value.cells().len() > MAX_CELLS {
        return Err(invalid("worksheet smart tags exceed the cell safety limit"));
    }
    let mut previous = None;
    let mut tag_count = 0usize;
    let mut property_count = 0usize;
    for cell_value in value.cells() {
        if previous.is_some_and(|address| address >= cell_value.address()) {
            return Err(invalid(format!(
                "worksheet smart-tag cells are duplicated or out of order at {}",
                cell_value.address()
            )));
        }
        previous = Some(cell_value.address());
        cell(cell_value)?;
        tag_count = tag_count
            .checked_add(cell_value.tags().len())
            .ok_or_else(|| invalid("smart-tag count overflow"))?;
        for tag_value in cell_value.tags() {
            property_count = property_count
                .checked_add(tag_value.properties().len())
                .ok_or_else(|| invalid("smart-tag property count overflow"))?;
        }
    }
    if tag_count > MAX_TAGS {
        return Err(invalid("worksheet smart tags exceed the tag safety limit"));
    }
    if property_count > MAX_PROPERTIES {
        return Err(invalid(
            "worksheet smart tags exceed the property safety limit",
        ));
    }
    Ok(())
}

pub(crate) fn cell(value: &Cell) -> Result<()> {
    if value.tags().is_empty() {
        return Err(invalid(format!(
            "cell {} requires at least one smart tag",
            value.address()
        )));
    }
    for tag_value in value.tags() {
        tag(tag_value)?;
    }
    Ok(())
}

pub(crate) fn tag(value: &Tag) -> Result<()> {
    type_id(value.type_id())?;
    let mut keys = HashSet::with_capacity(value.properties().len());
    for property_value in value.properties() {
        property(property_value)?;
        if !keys.insert(property_value.key()) {
            return Err(invalid(format!(
                "duplicate smart-tag property key '{}'",
                property_value.key()
            )));
        }
    }
    Ok(())
}

pub(crate) fn type_id(value: u32) -> Result<()> {
    if value > 32_768 {
        return Err(invalid(format!(
            "smart-tag type {value} is outside Office's 0..=32768 domain"
        )));
    }
    Ok(())
}

pub(crate) fn property(value: &Property) -> Result<()> {
    text(value.key(), "property key")?;
    text(value.value(), "property value")
}

fn text(value: &str, field: &str) -> Result<()> {
    if value.is_empty() {
        return Err(invalid(format!("smart-tag {field} cannot be empty")));
    }
    if value.len() > MAX_TEXT_BYTES {
        return Err(invalid(format!(
            "smart-tag {field} exceeds {MAX_TEXT_BYTES} bytes"
        )));
    }
    if value
        .chars()
        .any(|ch| !matches!(ch, '\u{9}' | '\u{A}' | '\u{D}' | '\u{20}'..='\u{D7FF}' | '\u{E000}'..='\u{FFFD}' | '\u{10000}'..='\u{10FFFF}'))
    {
        return Err(invalid(format!(
            "smart-tag {field} contains a character forbidden by XML 1.0"
        )));
    }
    Ok(())
}
