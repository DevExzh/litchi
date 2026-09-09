#![expect(
    clippy::arbitrary_source_item_ordering,
    reason = "items remain grouped by OOXML schema family and package lifecycle"
)]
//! OPC/package integration for document settings.

use std::borrow::Cow;

use super::model::DocumentSettings;
use crate::Variables;
use crate::error::{Error, Result};
use crate::mail_merge::{validate_mail_merge_relationships, validate_mail_merge_relationships_in};
use litchi_ooxml_common::mce::Limits as MceLimits;
use litchi_opc::Relationships;
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
    crate::parse_variables(xml.as_ref())
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
        let xml = super::super::extensions::process_part(part)?;
        let mut settings = Self::extract_from_xml(xml.as_ref())?;
        validate_mail_merge_relationships(part, settings.mail_merge.as_ref())?;
        validate_attached_template_relationship(part.rels(), &mut settings)?;
        Ok(settings)
    }

    /// Apply the DOCX MCE capability profile once and return its borrowed or
    /// owned XML result.  The caller keeps that result alive while it performs
    /// the complete settings and mail-merge validation pass.
    pub(crate) fn process_bytes_with_mce_limits<'a>(
        bytes: &'a [u8],
        mce_limits: &MceLimits,
    ) -> Result<Cow<'a, [u8]>> {
        super::super::extensions::process_bytes_with_limits(bytes, mce_limits)
    }

    /// Validate a settings model which has already passed the package MCE
    /// capability profile, against the original relationship map.
    pub(crate) fn extract_from_processed_xml_with_relationships(
        xml: &[u8],
        relationships: &Relationships,
    ) -> Result<Self> {
        let mut settings = Self::extract_from_processed_xml(xml)?;
        validate_mail_merge_relationships_in(relationships, settings.mail_merge.as_ref())?;
        validate_attached_template_relationship(relationships, &mut settings)?;
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
    relationships: &Relationships,
    settings: &mut DocumentSettings,
) -> Result<()> {
    let mut matching = None;
    for relationship in relationships
        .iter()
        .filter(|relationship| is_attached_template_relationship(relationship.reltype()))
    {
        if matching.replace(relationship).is_some() {
            return Err(Error::InvalidFormat(
                "settings part has multiple attached-template relationships".into(),
            ));
        }
    }

    let Some(attached_template) = settings.attached_template.as_mut() else {
        if matching.is_none() {
            return Ok(());
        }
        return Err(Error::InvalidFormat(
            "settings part has an attached-template relationship without an attachedTemplate element"
                .into(),
        ));
    };
    let relationship = relationships
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
