//! Safety checks for the workbook external-link graph.

use std::collections::HashSet;

use super::model::{Conformance, Link};
use super::package::Entry;
use super::{invalid, limit};
use crate::error::Result;

/// Validate one inert external-link model before package publication.
pub fn link(value: &Link, conformance: Conformance) -> Result<()> {
    let xml = value.to_xml_with_conformance(conformance)?;
    if xml.len() > super::model::MAX_CACHE_TEXT_BYTES {
        return Err(limit("serialized XML"));
    }
    Ok(())
}

/// Validate the typed entries and their stable workbook ownership metadata.
pub fn entries(values: &[Entry], conformance: Conformance) -> Result<()> {
    if values.len() > super::model::MAX_LINK_ITEMS {
        return Err(limit("workbook external links"));
    }
    let mut relationship_ids = HashSet::with_capacity(values.len());
    let mut part_names = HashSet::with_capacity(values.len());
    for entry in values {
        if entry.relationship_id.is_empty()
            || entry.relationship_id.len() > 1024
            || entry.relationship_id.chars().any(char::is_control)
        {
            return Err(invalid("external-link workbook relationship ID is invalid"));
        }
        if !relationship_ids.insert(entry.relationship_id.as_str()) {
            return Err(invalid(format!(
                "duplicate external-link workbook relationship ID '{}'",
                entry.relationship_id
            )));
        }
        if !part_names.insert(entry.part_uri.as_str()) {
            return Err(invalid(format!(
                "duplicate external-link part '{}'",
                entry.part_uri
            )));
        }
        link(&entry.link, conformance)?;
    }
    Ok(())
}
