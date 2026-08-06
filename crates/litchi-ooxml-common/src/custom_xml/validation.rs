//! Bounded semantic and package-graph validation for Custom XML edits.

use crate::{Error, Result};

use super::codec::{require_rel_id, validate_content_type, validate_payload, validate_props};
use super::model::{Item, MAX_ITEMS};

pub(super) fn items(values: &[Item]) -> Result<()> {
    if values.len() > MAX_ITEMS {
        return Err(Error::Limit {
            resource: "custom XML items",
            max: MAX_ITEMS,
            actual: values.len(),
        });
    }

    for (index, value) in values.iter().enumerate() {
        item(index, value)?;
        if values[..index]
            .iter()
            .any(|other| other.source() == value.source() && other.rel_id() == value.rel_id())
        {
            return Err(Error::Relationship(format!(
                "duplicate custom XML relationship '{}' on '{}'",
                value.rel_id(),
                value.source().as_str()
            )));
        }
        if let Some(props) = value.props() {
            if values[..index].iter().any(|other| {
                other
                    .props()
                    .is_some_and(|candidate| candidate.id.eq_ignore_ascii_case(&props.id))
            }) {
                return Err(Error::Invalid(format!(
                    "duplicate custom XML itemID '{}'",
                    props.id
                )));
            }
        }
    }
    Ok(())
}

pub(super) fn item(index: usize, value: &Item) -> Result<()> {
    validate_content_type(value.content_type())?;
    validate_payload(value.xml())?;
    require_rel_id(value.rel_id(), "custom XML relationship")?;
    if value.source_relationship().id != value.rel_id() {
        return Err(Error::Relationship(format!(
            "custom XML source relationship ID '{}' disagrees with item ID '{}'",
            value.source_relationship().id,
            value.rel_id()
        )));
    }
    if let Some(props) = value.props() {
        validate_props(props)?;
        let Some(props_xml) = value.props_xml() else {
            return Err(Error::Invalid(format!(
                "custom XML item {index} has typed properties without source XML"
            )));
        };
        let parsed = super::codec::read_props(props_xml)?;
        if &parsed != props {
            return Err(Error::Invalid(format!(
                "custom XML item {index} properties projection is stale"
            )));
        }
    } else if value.props_xml().is_some() {
        return Err(Error::Invalid(format!(
            "custom XML item {index} has properties XML without typed properties"
        )));
    }
    for relationship in value.relationships() {
        require_rel_id(&relationship.id, "custom XML data relationship")?;
    }
    Ok(())
}
