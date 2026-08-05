//! OPC/package integration for document settings.

use super::model::DocumentSettings;
use crate::Variables;
use crate::error::{Error, Result};
use crate::mail_merge::validate_mail_merge_relationships;
use litchi_opc::part::Part;

/// Relationship type required by Word for an attached document template.
pub(crate) const ATTACHED_TEMPLATE_RELATIONSHIP: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/attachedTemplate";
/// Strict OOXML alias accepted when reading existing packages.
pub(crate) const STRICT_ATTACHED_TEMPLATE_RELATIONSHIP: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/attachedTemplate";

const MAX_ATTACHED_TEMPLATE_TARGET_LEN: usize = 32 * 1024;

pub(crate) fn is_attached_template_relationship(value: &str) -> bool {
    matches!(
        value,
        ATTACHED_TEMPLATE_RELATIONSHIP | STRICT_ATTACHED_TEMPLATE_RELATIONSHIP
    )
}

/// Decode the canonical DOCX document-variable collection from a settings part.
///
/// Markup-compatibility preprocessing and OPC ownership remain host concerns;
/// the validated model and XML codec belong to `litchi-docx`.
pub(crate) fn extract_document_variables(part: &dyn Part) -> Result<Variables> {
    let limit = crate::variables::MAX_DOCUMENT_VARIABLE_XML_BYTES;
    if part.blob().len() > limit {
        return Err(Error::InvalidFormat(format!(
            "settings XML exceeds the {limit} byte document-variable limit"
        )));
    }
    let xml = litchi_ooxml_common::mce::process_part(part)?;
    Ok(crate::parse_variables(xml.as_ref())?)
}

impl DocumentSettings {
    /// Extract settings from a settings.xml part.
    ///
    /// # Arguments
    ///
    /// * `part` - The settings part
    ///
    /// # Returns
    ///
    /// A Settings object
    pub(crate) fn extract_from_part(part: &dyn Part) -> Result<Self> {
        let xml = litchi_ooxml_common::mce::process_part(part)?;
        let mut settings = Self::extract_from_xml(xml.as_ref())?;
        validate_mail_merge_relationships(part, settings.mail_merge.as_ref())?;
        validate_attached_template_relationship(part, &mut settings)?;
        Ok(settings)
    }
}

pub(crate) fn validate_attached_template_target(target: &str) -> Result<()> {
    if target.is_empty() || target.len() > MAX_ATTACHED_TEMPLATE_TARGET_LEN {
        return Err(Error::InvalidFormat(
            "attached-template target must contain 1 to 32768 bytes".into(),
        ));
    }
    if target
        .chars()
        .any(|character| character.is_control() || character.is_whitespace())
    {
        return Err(Error::InvalidFormat(
            "attached-template target contains an invalid control or whitespace character".into(),
        ));
    }
    Ok(())
}

fn validate_attached_template_relationship(
    part: &dyn Part,
    settings: &mut DocumentSettings,
) -> Result<()> {
    let matching = part
        .rels()
        .iter()
        .filter(|relationship| is_attached_template_relationship(relationship.reltype()))
        .collect::<Vec<_>>();
    if matching.len() > 1 {
        return Err(Error::InvalidFormat(
            "settings part has multiple attached-template relationships".into(),
        ));
    }

    let Some(attached_template) = settings.attached_template.as_mut() else {
        if matching.is_empty() {
            return Ok(());
        }
        return Err(Error::InvalidFormat(
            "settings part has an attached-template relationship without an attachedTemplate element"
                .into(),
        ));
    };
    let relationship = part
        .rels()
        .get(&attached_template.relationship_id)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "attachedTemplate references missing relationship {:?}",
                attached_template.relationship_id
            ))
        })?;
    if !is_attached_template_relationship(relationship.reltype()) {
        return Err(Error::InvalidFormat(
            "attachedTemplate relationship has the wrong type".into(),
        ));
    }
    if !relationship.is_external() {
        return Err(Error::InvalidFormat(
            "attachedTemplate relationship must use external target mode".into(),
        ));
    }
    validate_attached_template_target(relationship.target_ref())?;
    attached_template.target_uri = relationship.target_ref().to_owned();
    Ok(())
}
