//! Compatibility adapter for the canonical DOCX mail-merge codec.
//!
//! The typed settings, ODSO/field-map model, bounded XML codec, and inert
//! recipient-data parser live in `litchi_docx::mail_merge`. This module keeps
//! the historical `litchi_ooxml::docx` names and maps canonical failures to
//! the host error boundary. Package relationship/resource orchestration stays
//! in the host and never follows or executes a source.

use crate::error::{OoxmlError, Result};
use litchi_opc::part::Part;

pub use litchi_docx::mail_merge::{
    MailMergeConformance, MailMergeDataSourceObject, MailMergeDataType, MailMergeDestination,
    MailMergeFieldMap, MailMergeFieldMappingType, MailMergeMainDocumentType, MailMergeRecipient,
    MailMergeSource, MailMergeTarget, RECIPIENT_CONTENT_TYPE,
};

fn map_docx_error(error: litchi_docx::Error) -> OoxmlError {
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

/// Historical mail-merge settings facade with the host's error type.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MailMergeSettings {
    inner: litchi_docx::mail_merge::MailMergeSettings,
}

impl Default for MailMergeSettings {
    fn default() -> Self {
        Self::new()
    }
}

impl MailMergeSettings {
    pub fn new() -> Self {
        Self {
            inner: litchi_docx::mail_merge::MailMergeSettings::new(),
        }
    }

    fn from_owner(inner: litchi_docx::mail_merge::MailMergeSettings) -> Self {
        Self { inner }
    }

    pub fn set_main_document_type(&mut self, value: MailMergeMainDocumentType) -> &mut Self {
        self.inner.set_main_document_type(value);
        self
    }

    pub fn set_link_to_query(&mut self, value: bool) -> &mut Self {
        self.inner.set_link_to_query(value);
        self
    }

    pub fn set_data_type(&mut self, value: Option<MailMergeDataType>) -> &mut Self {
        self.inner.set_data_type(value);
        self
    }

    pub fn set_connect_string(&mut self, value: Option<String>) -> &mut Self {
        self.inner.set_connect_string(value);
        self
    }

    pub fn set_query(&mut self, value: Option<String>) -> &mut Self {
        self.inner.set_query(value);
        self
    }

    pub fn set_do_not_suppress_blank_lines(&mut self, value: bool) -> &mut Self {
        self.inner.set_do_not_suppress_blank_lines(value);
        self
    }

    pub fn set_destination(&mut self, value: MailMergeDestination) -> &mut Self {
        self.inner.set_destination(value);
        self
    }

    pub fn set_address_field_name(&mut self, value: Option<String>) -> &mut Self {
        self.inner.set_address_field_name(value);
        self
    }

    pub fn set_mail_subject(&mut self, value: Option<String>) -> &mut Self {
        self.inner.set_mail_subject(value);
        self
    }

    pub fn set_mail_as_attachment(&mut self, value: bool) -> &mut Self {
        self.inner.set_mail_as_attachment(value);
        self
    }

    pub fn set_view_merged_data(&mut self, value: bool) -> &mut Self {
        self.inner.set_view_merged_data(value);
        self
    }

    pub fn set_active_record(&mut self, value: i32) -> &mut Self {
        self.inner.set_active_record(value);
        self
    }

    pub fn set_check_errors(&mut self, value: i32) -> &mut Self {
        self.inner.set_check_errors(value);
        self
    }

    pub fn set_odso(&mut self, value: Option<MailMergeDataSourceObject>) -> &mut Self {
        self.inner.set_odso(value);
        self
    }

    pub(crate) fn assign_package_relationships(
        &mut self,
        data_source: Option<String>,
        header_source: Option<String>,
        recipient_data: Option<String>,
    ) {
        self.inner
            .assign_package_relationships(data_source, header_source, recipient_data);
    }

    pub fn main_document_type(&self) -> MailMergeMainDocumentType {
        self.inner.main_document_type()
    }

    pub fn link_to_query(&self) -> bool {
        self.inner.link_to_query()
    }

    pub fn data_type(&self) -> Option<MailMergeDataType> {
        self.inner.data_type()
    }

    pub fn connect_string(&self) -> Option<&str> {
        self.inner.connect_string()
    }

    pub fn query(&self) -> Option<&str> {
        self.inner.query()
    }

    pub fn data_source_relationship_id(&self) -> Option<&str> {
        self.inner.data_source_relationship_id()
    }

    pub fn header_source_relationship_id(&self) -> Option<&str> {
        self.inner.header_source_relationship_id()
    }

    pub fn do_not_suppress_blank_lines(&self) -> bool {
        self.inner.do_not_suppress_blank_lines()
    }

    pub fn destination(&self) -> MailMergeDestination {
        self.inner.destination()
    }

    pub fn address_field_name(&self) -> Option<&str> {
        self.inner.address_field_name()
    }

    pub fn mail_subject(&self) -> Option<&str> {
        self.inner.mail_subject()
    }

    pub fn mail_as_attachment(&self) -> bool {
        self.inner.mail_as_attachment()
    }

    pub fn view_merged_data(&self) -> bool {
        self.inner.view_merged_data()
    }

    pub fn active_record(&self) -> i32 {
        self.inner.active_record()
    }

    pub fn check_errors(&self) -> i32 {
        self.inner.check_errors()
    }

    pub fn odso(&self) -> Option<&MailMergeDataSourceObject> {
        self.inner.odso()
    }

    /// Serialize a standalone `w:mailMerge` fragment in schema order.
    pub fn to_xml(&self, conformance: MailMergeConformance) -> Result<String> {
        self.inner.to_xml(conformance).map_err(map_docx_error)
    }
}

/// Historical recipient collection facade with the host's error type.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct MailMergeRecipients {
    inner: litchi_docx::mail_merge::MailMergeRecipients,
}

impl MailMergeRecipients {
    pub fn new() -> Self {
        Self::default()
    }

    fn from_owner(inner: litchi_docx::mail_merge::MailMergeRecipients) -> Self {
        Self { inner }
    }

    pub fn recipients(&self) -> &[MailMergeRecipient] {
        self.inner.recipients()
    }

    pub fn recipients_mut(&mut self) -> &mut Vec<MailMergeRecipient> {
        self.inner.recipients_mut()
    }

    pub fn add_recipient(&mut self, recipient: MailMergeRecipient) -> Result<&mut Self> {
        self.inner
            .add_recipient(recipient)
            .map_err(map_docx_error)?;
        Ok(self)
    }

    pub fn set_recipient_active(&mut self, index: usize, active: bool) -> Result<()> {
        self.inner
            .set_recipient_active(index, active)
            .map_err(map_docx_error)
    }

    pub(crate) fn content_type() -> &'static str {
        litchi_docx::mail_merge::RECIPIENT_CONTENT_TYPE
    }

    pub(crate) fn extract_from_part(part: &dyn Part) -> Result<Self> {
        litchi_docx::mail_merge::MailMergeRecipients::extract_from_part(part)
            .map(Self::from_owner)
            .map_err(map_docx_error)
    }

    pub fn to_xml(&self, conformance: MailMergeConformance) -> Result<String> {
        self.inner.to_xml(conformance).map_err(map_docx_error)
    }
}

pub(crate) fn parse_settings_mail_merge(xml: &[u8]) -> Result<Option<MailMergeSettings>> {
    litchi_docx::mail_merge::parse_settings_mail_merge(xml)
        .map(|value| value.map(MailMergeSettings::from_owner))
        .map_err(map_docx_error)
}

/// Validate only the host package relationship closure around typed settings.
///
/// The owner validates XML and metadata bounds. Relationship cardinality,
/// target mode, and type are package-graph concerns and remain here.
pub(crate) fn validate_mail_merge_relationships(
    part: &dyn Part,
    value: Option<&MailMergeSettings>,
) -> Result<()> {
    let recipient_relationships: Vec<_> = part
        .rels()
        .iter()
        .filter(|rel| reltype_is(rel.reltype(), "recipientData"))
        .collect();
    if recipient_relationships.len() > 1 {
        return Err(invalid(
            "settings part has multiple recipient-data relationships",
        ));
    }
    if recipient_relationships.iter().any(|rel| rel.is_external()) {
        return Err(invalid("recipient-data relationship must be internal"));
    }

    let Some(value) = value else {
        if !recipient_relationships.is_empty() {
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

    match (recipient_relationships.first(), recipient_id) {
        (Some(rel), Some(id)) if rel.r_id() == id => {},
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

fn reltype_is(value: &str, suffix: &str) -> bool {
    value == format!("http://schemas.openxmlformats.org/officeDocument/2006/relationships/{suffix}")
        || value == format!("http://purl.oclc.org/ooxml/officeDocument/relationships/{suffix}")
}

pub(crate) fn is_settings_relationship(value: &str) -> bool {
    value == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"
        || value == "http://purl.oclc.org/ooxml/officeDocument/relationships/settings"
}

fn invalid(message: impl Into<String>) -> OoxmlError {
    OoxmlError::InvalidFormat(message.into())
}
