//! DOCX mail-merge package boundary.
//!
//! `litchi_docx::mail_merge` owns the settings, ODSO/field-map, recipient, and
//! XML semantics. This adapter only maps owner errors to the host boundary and
//! validates the settings-part relationship closure; source and recipient
//! targets remain inert package data.

use crate::error::{OoxmlError, Result};
use litchi_opc::part::Part;

pub use litchi_docx::mail_merge::{
    Conformance, DataSourceObject, DataType, Destination, FieldMap, FieldMappingType,
    MainDocumentType, RECIPIENT_CONTENT_TYPE, Recipient, Recipients, Settings, Source, Target,
};

pub(crate) fn map_docx_error(error: litchi_docx::Error) -> OoxmlError {
    match error {
        litchi_docx::Error::Opc(error) => OoxmlError::Opc(error),
        litchi_docx::Error::Xml(message) => OoxmlError::Xml(message),
        litchi_docx::Error::ContentType { expected, actual } => OoxmlError::InvalidContentType {
            expected,
            got: actual,
        },
        litchi_docx::Error::Invalid(message) => OoxmlError::InvalidFormat(message),
        litchi_docx::Error::Uri(message) => OoxmlError::InvalidUri(message),
        litchi_docx::Error::Mce(error) => {
            OoxmlError::Common(litchi_ooxml_common::Error::Mce(error))
        },
        litchi_docx::Error::Allocation { resource, source } => {
            OoxmlError::Allocation { resource, source }
        },
        other => OoxmlError::Docx(other),
    }
}

pub(crate) fn parse_settings_mail_merge(xml: &[u8]) -> Result<Option<Settings>> {
    litchi_docx::mail_merge::parse_settings_mail_merge(xml).map_err(map_docx_error)
}

pub(crate) fn extract_recipients(part: &dyn Part) -> Result<Recipients> {
    Recipients::extract_from_part(part).map_err(map_docx_error)
}

/// Validate only the host package relationship closure around typed settings.
///
/// The owner validates XML and semantic metadata bounds. Relationship
/// cardinality, target mode, and relationship type are package-graph concerns
/// and remain here.
pub(crate) fn validate_mail_merge_relationships(
    part: &dyn Part,
    value: Option<&Settings>,
) -> Result<()> {
    let mut recipient_relationship = None;
    for relationship in part
        .rels()
        .iter()
        .filter(|rel| reltype_is(rel.reltype(), "recipientData"))
    {
        if recipient_relationship.is_some() {
            return Err(invalid(
                "settings part has multiple recipient-data relationships",
            ));
        }
        recipient_relationship = Some(relationship);
    }
    if recipient_relationship.is_some_and(|relationship| relationship.is_external()) {
        return Err(invalid("recipient-data relationship must be internal"));
    }

    let Some(value) = value else {
        if recipient_relationship.is_some() {
            return Err(invalid(
                "recipient-data relationship is not referenced by mailMerge",
            ));
        }
        return Ok(());
    };

    validate_optional_relationship(
        part,
        value.data_source_relationship_id(),
        "mailMergeSource",
        true,
    )?;
    validate_optional_relationship(
        part,
        value.header_source_relationship_id(),
        "mailMergeHeaderSource",
        true,
    )?;

    let recipient_id = if let Some(odso) = value.odso() {
        validate_optional_relationship(
            part,
            odso.source_relationship_id(),
            "mailMergeSource",
            true,
        )?;
        validate_optional_relationship(
            part,
            odso.recipient_data_relationship_id(),
            "recipientData",
            false,
        )?;
        odso.recipient_data_relationship_id()
    } else {
        None
    };

    match (recipient_relationship, recipient_id) {
        (Some(relationship), Some(id)) if relationship.r_id() == id => {},
        (None, None) => {},
        (Some(_), None) => {
            return Err(invalid(
                "recipient-data relationship is not referenced by odso",
            ));
        },
        (None, Some(_)) => return Err(invalid("odso recipientData relationship is missing")),
        (Some(_), Some(_)) => {
            return Err(invalid(
                "odso references the wrong recipient-data relationship",
            ));
        },
    }
    Ok(())
}

fn validate_optional_relationship(
    part: &dyn Part,
    id: Option<&str>,
    suffix: &str,
    allow_external: bool,
) -> Result<()> {
    let Some(id) = id else {
        return Ok(());
    };
    let relationship = part
        .rels()
        .get(id)
        .ok_or_else(|| invalid(format!("mail-merge relationship '{id}' is missing")))?;
    if !reltype_is(relationship.reltype(), suffix) {
        return Err(invalid(format!(
            "mail-merge relationship '{id}' has invalid type '{}'",
            relationship.reltype()
        )));
    }
    if !allow_external && relationship.is_external() {
        return Err(invalid(format!(
            "mail-merge relationship '{id}' must be internal"
        )));
    }
    Ok(())
}

pub(crate) fn is_mail_merge_relationship_type(value: &str) -> bool {
    ["mailMergeSource", "mailMergeHeaderSource", "recipientData"]
        .iter()
        .any(|suffix| reltype_is(value, suffix))
}

fn reltype_is(value: &str, suffix: &str) -> bool {
    const TRANSITIONAL: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
    const STRICT: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/";

    value
        .strip_prefix(TRANSITIONAL)
        .is_some_and(|candidate| candidate == suffix)
        || value
            .strip_prefix(STRICT)
            .is_some_and(|candidate| candidate == suffix)
}

pub(crate) fn is_settings_relationship(value: &str) -> bool {
    reltype_is(value, "settings")
}

fn invalid(message: impl Into<String>) -> OoxmlError {
    OoxmlError::InvalidFormat(message.into())
}
