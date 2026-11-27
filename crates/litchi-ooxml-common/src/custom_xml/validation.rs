//! Bounded semantic and package-graph validation for Custom XML edits.

use crate::{Error, Result};
use litchi_opc::SourceLimits;

use super::codec::{require_rel_id, validate_content_type, validate_payload, validate_props};
use super::model::{Item, MAX_ITEMS};

pub(super) fn items(values: &[Item], source_limits: &SourceLimits) -> Result<()> {
    if values.len() > MAX_ITEMS {
        return Err(Error::Limit {
            kind: crate::LimitKind::CustomXml(super::LimitKind::Items),
            maximum: MAX_ITEMS,
            observed: values.len(),
        });
    }

    for (index, value) in values.iter().enumerate() {
        item(index, value, source_limits)?;
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
        if let Some(props) = value.props()
            && values[..index].iter().any(|other| {
                other
                    .props()
                    .is_some_and(|candidate| candidate.id.eq_ignore_ascii_case(&props.id))
            })
        {
            return Err(Error::Invalid(format!(
                "duplicate custom XML itemID '{}'",
                props.id
            )));
        }
    }
    Ok(())
}

pub(super) fn item(_index: usize, value: &Item, source_limits: &SourceLimits) -> Result<()> {
    validate_content_type(value.content_type())?;
    validate_payload(value.xml(), source_limits)?;
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
            return Err(Error::Internal(
                "custom XML item has typed properties without source XML",
            ));
        };
        let parsed = super::codec::read_props(props_xml, source_limits)?;
        if &parsed != props {
            return Err(Error::Internal(
                "custom XML item properties projection is stale",
            ));
        }
    } else if value.props_xml().is_some() {
        return Err(Error::Internal(
            "custom XML item has properties XML without typed properties",
        ));
    }
    for relationship in value.relationships() {
        require_rel_id(&relationship.id, "custom XML data relationship")?;
    }
    Ok(())
}
