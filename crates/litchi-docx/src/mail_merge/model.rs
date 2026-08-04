//! Package-neutral mail-merge value models.

use crate::{Error, Result};

pub(super) const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
pub(super) const STRICT_W: &str = "http://purl.oclc.org/ooxml/wordprocessingml/main";
pub(super) const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
pub(super) const STRICT_R: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
pub(super) const MAX_DEPTH: usize = 128;
pub(super) const MAX_NODES: usize = 1_000_000;
pub(super) const MAX_STRING_BYTES: usize = 16 * 1024 * 1024;
pub(super) const MAX_RELATIONSHIP_ID_BYTES: usize = 1024;
pub(super) const MAX_FIELD_MAPS: usize = 16_384;
pub(super) const MAX_RECIPIENTS: usize = 1_000_000;
pub(super) const MAX_UNIQUE_TAG_BYTES: usize = 1024 * 1024;
pub(super) const MAX_ATTRIBUTES_PER_NODE: usize = 256;

pub(super) fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

/// Namespace family used by deterministic mail-merge serializers.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Conformance {
    Transitional,
    Strict,
}

impl Conformance {
    pub(super) fn word(self) -> &'static str {
        match self {
            Self::Transitional => W,
            Self::Strict => STRICT_W,
        }
    }

    pub(super) fn relationships(self) -> &'static str {
        match self {
            Self::Transitional => R,
            Self::Strict => STRICT_R,
        }
    }
}

// The explicit implementations keep rustdoc useful and defaults visible.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum MainDocumentType {
    Catalog,
    Envelopes,
    MailingLabels,
    #[default]
    FormLetters,
    Email,
    Fax,
}

impl MainDocumentType {
    pub(super) fn parse(value: &str) -> Result<Self> {
        match value {
            "catalog" => Ok(Self::Catalog),
            "envelopes" => Ok(Self::Envelopes),
            "mailingLabels" => Ok(Self::MailingLabels),
            "formLetters" => Ok(Self::FormLetters),
            "email" => Ok(Self::Email),
            "fax" => Ok(Self::Fax),
            _ => Err(invalid(format!(
                "invalid mail-merge document type '{value}'"
            ))),
        }
    }

    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Catalog => "catalog",
            Self::Envelopes => "envelopes",
            Self::MailingLabels => "mailingLabels",
            Self::FormLetters => "formLetters",
            Self::Email => "email",
            Self::Fax => "fax",
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DataType {
    TextFile,
    Database,
    Spreadsheet,
    Query,
    Odbc,
    Native,
}

impl DataType {
    pub(super) fn parse(value: &str) -> Result<Self> {
        match value {
            "textFile" => Ok(Self::TextFile),
            "database" => Ok(Self::Database),
            "spreadsheet" => Ok(Self::Spreadsheet),
            "query" => Ok(Self::Query),
            "odbc" => Ok(Self::Odbc),
            "native" => Ok(Self::Native),
            _ => Err(invalid(format!("invalid mail-merge data type '{value}'"))),
        }
    }

    pub const fn as_str(self) -> &'static str {
        match self {
            Self::TextFile => "textFile",
            Self::Database => "database",
            Self::Spreadsheet => "spreadsheet",
            Self::Query => "query",
            Self::Odbc => "odbc",
            Self::Native => "native",
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum Destination {
    #[default]
    NewDocument,
    Printer,
    Email,
    Fax,
}

impl Destination {
    pub(super) fn parse(value: &str) -> Result<Self> {
        match value {
            "newDocument" => Ok(Self::NewDocument),
            "printer" => Ok(Self::Printer),
            "email" => Ok(Self::Email),
            "fax" => Ok(Self::Fax),
            _ => Err(invalid(format!("invalid mail-merge destination '{value}'"))),
        }
    }

    pub const fn as_str(self) -> &'static str {
        match self {
            Self::NewDocument => "newDocument",
            Self::Printer => "printer",
            Self::Email => "email",
            Self::Fax => "fax",
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FieldMappingType {
    Null,
    DatabaseColumn,
}

impl FieldMappingType {
    pub(super) fn parse(value: &str) -> Result<Self> {
        match value {
            "null" => Ok(Self::Null),
            "dbColumn" => Ok(Self::DatabaseColumn),
            _ => Err(invalid(format!(
                "invalid mail-merge field mapping type '{value}'"
            ))),
        }
    }

    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Null => "null",
            Self::DatabaseColumn => "dbColumn",
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct FieldMap {
    pub(super) mapping_type: Option<FieldMappingType>,
    pub(super) name: Option<String>,
    pub(super) mapped_name: Option<String>,
    pub(super) column: Option<i32>,
    pub(super) language_id: Option<String>,
    pub(super) dynamic_address: bool,
}

impl FieldMap {
    pub fn new() -> Self {
        Self::default()
    }
    pub fn set_mapping_type(&mut self, value: Option<FieldMappingType>) -> &mut Self {
        self.mapping_type = value;
        self
    }
    pub fn set_name(&mut self, value: Option<String>) -> &mut Self {
        self.name = value;
        self
    }
    pub fn set_mapped_name(&mut self, value: Option<String>) -> &mut Self {
        self.mapped_name = value;
        self
    }
    pub fn set_column(&mut self, value: Option<i32>) -> &mut Self {
        self.column = value;
        self
    }
    pub fn set_language_id(&mut self, value: Option<String>) -> &mut Self {
        self.language_id = value;
        self
    }
    pub fn set_dynamic_address(&mut self, value: bool) -> &mut Self {
        self.dynamic_address = value;
        self
    }
    pub fn mapping_type(&self) -> Option<FieldMappingType> {
        self.mapping_type
    }
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }
    pub fn mapped_name(&self) -> Option<&str> {
        self.mapped_name.as_deref()
    }
    pub fn column(&self) -> Option<i32> {
        self.column
    }
    pub fn language_id(&self) -> Option<&str> {
        self.language_id.as_deref()
    }
    pub fn dynamic_address(&self) -> bool {
        self.dynamic_address
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct DataSourceObject {
    pub(super) udl: Option<String>,
    pub(super) table: Option<String>,
    pub(super) source_relationship_id: Option<String>,
    pub(super) column_delimiter: Option<i32>,
    pub(super) source_type: Option<String>,
    pub(super) first_row_header: bool,
    pub(super) field_maps: Vec<FieldMap>,
    pub(super) recipient_data_relationship_id: Option<String>,
}

impl DataSourceObject {
    pub fn new() -> Self {
        Self::default()
    }
    pub fn set_udl(&mut self, value: Option<String>) -> &mut Self {
        self.udl = value;
        self
    }
    pub fn set_table(&mut self, value: Option<String>) -> &mut Self {
        self.table = value;
        self
    }
    pub fn set_column_delimiter(&mut self, value: Option<i32>) -> &mut Self {
        self.column_delimiter = value;
        self
    }
    pub fn set_source_type(&mut self, value: Option<String>) -> &mut Self {
        self.source_type = value;
        self
    }
    pub fn set_first_row_header(&mut self, value: bool) -> &mut Self {
        self.first_row_header = value;
        self
    }
    pub fn field_maps_mut(&mut self) -> &mut Vec<FieldMap> {
        &mut self.field_maps
    }
    pub fn add_field_map(&mut self, value: FieldMap) -> &mut Self {
        self.field_maps.push(value);
        self
    }
    pub fn udl(&self) -> Option<&str> {
        self.udl.as_deref()
    }
    pub fn table(&self) -> Option<&str> {
        self.table.as_deref()
    }
    pub fn source_relationship_id(&self) -> Option<&str> {
        self.source_relationship_id.as_deref()
    }
    pub fn column_delimiter(&self) -> Option<i32> {
        self.column_delimiter
    }
    pub fn source_type(&self) -> Option<&str> {
        self.source_type.as_deref()
    }
    pub fn first_row_header(&self) -> bool {
        self.first_row_header
    }
    pub fn field_maps(&self) -> &[FieldMap] {
        &self.field_maps
    }
    pub fn recipient_data_relationship_id(&self) -> Option<&str> {
        self.recipient_data_relationship_id.as_deref()
    }
}

/// Complete inert metadata from `w:mailMerge`.
///
/// The Word-specific optional `mainDocumentType` and single-`odso` behavior
/// follows the checked-in `[MS-OE376]` sections 2.1.381 and 2.1.384. The
/// relationship IDs are retained as inert tokens; this owner never resolves
/// their targets.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Settings {
    pub(super) main_document_type: MainDocumentType,
    pub(super) link_to_query: bool,
    pub(super) data_type: Option<DataType>,
    pub(super) connect_string: Option<String>,
    pub(super) query: Option<String>,
    pub(super) data_source_relationship_id: Option<String>,
    pub(super) header_source_relationship_id: Option<String>,
    pub(super) do_not_suppress_blank_lines: bool,
    pub(super) destination: Destination,
    pub(super) address_field_name: Option<String>,
    pub(super) mail_subject: Option<String>,
    pub(super) mail_as_attachment: bool,
    pub(super) view_merged_data: bool,
    pub(super) active_record: i32,
    pub(super) check_errors: i32,
    pub(super) odso: Option<DataSourceObject>,
}

impl Default for Settings {
    fn default() -> Self {
        Self {
            main_document_type: MainDocumentType::FormLetters,
            link_to_query: false,
            data_type: None,
            connect_string: None,
            query: None,
            data_source_relationship_id: None,
            header_source_relationship_id: None,
            do_not_suppress_blank_lines: false,
            destination: Destination::NewDocument,
            address_field_name: None,
            mail_subject: None,
            mail_as_attachment: false,
            view_merged_data: false,
            active_record: 1,
            check_errors: 2,
            odso: None,
        }
    }
}

impl Settings {
    pub fn new() -> Self {
        Self::default()
    }
    pub fn set_main_document_type(&mut self, value: MainDocumentType) -> &mut Self {
        self.main_document_type = value;
        self
    }
    pub fn set_link_to_query(&mut self, value: bool) -> &mut Self {
        self.link_to_query = value;
        self
    }
    pub fn set_data_type(&mut self, value: Option<DataType>) -> &mut Self {
        self.data_type = value;
        self
    }
    pub fn set_connect_string(&mut self, value: Option<String>) -> &mut Self {
        self.connect_string = value;
        self
    }
    pub fn set_query(&mut self, value: Option<String>) -> &mut Self {
        self.query = value;
        self
    }
    pub fn set_do_not_suppress_blank_lines(&mut self, value: bool) -> &mut Self {
        self.do_not_suppress_blank_lines = value;
        self
    }
    pub fn set_destination(&mut self, value: Destination) -> &mut Self {
        self.destination = value;
        self
    }
    pub fn set_address_field_name(&mut self, value: Option<String>) -> &mut Self {
        self.address_field_name = value;
        self
    }
    pub fn set_mail_subject(&mut self, value: Option<String>) -> &mut Self {
        self.mail_subject = value;
        self
    }
    pub fn set_mail_as_attachment(&mut self, value: bool) -> &mut Self {
        self.mail_as_attachment = value;
        self
    }
    pub fn set_view_merged_data(&mut self, value: bool) -> &mut Self {
        self.view_merged_data = value;
        self
    }
    pub fn set_active_record(&mut self, value: i32) -> &mut Self {
        self.active_record = value;
        self
    }
    pub fn set_check_errors(&mut self, value: i32) -> &mut Self {
        self.check_errors = value;
        self
    }
    pub fn set_odso(&mut self, value: Option<DataSourceObject>) -> &mut Self {
        self.odso = value;
        self
    }

    pub fn assign_package_relationships(
        &mut self,
        data_source: Option<String>,
        header_source: Option<String>,
        recipient_data: Option<String>,
    ) {
        self.data_source_relationship_id = data_source.clone();
        self.header_source_relationship_id = header_source;
        if self.odso.is_none() && (data_source.is_some() || recipient_data.is_some()) {
            self.odso = Some(DataSourceObject::default());
        }
        if let Some(odso) = &mut self.odso {
            odso.source_relationship_id = data_source;
            odso.recipient_data_relationship_id = recipient_data;
        }
    }
    pub fn main_document_type(&self) -> MainDocumentType {
        self.main_document_type
    }
    pub fn link_to_query(&self) -> bool {
        self.link_to_query
    }
    pub fn data_type(&self) -> Option<DataType> {
        self.data_type
    }
    pub fn connect_string(&self) -> Option<&str> {
        self.connect_string.as_deref()
    }
    pub fn query(&self) -> Option<&str> {
        self.query.as_deref()
    }
    pub fn data_source_relationship_id(&self) -> Option<&str> {
        self.data_source_relationship_id.as_deref()
    }
    pub fn header_source_relationship_id(&self) -> Option<&str> {
        self.header_source_relationship_id.as_deref()
    }
    pub fn do_not_suppress_blank_lines(&self) -> bool {
        self.do_not_suppress_blank_lines
    }
    pub fn destination(&self) -> Destination {
        self.destination
    }
    pub fn address_field_name(&self) -> Option<&str> {
        self.address_field_name.as_deref()
    }
    pub fn mail_subject(&self) -> Option<&str> {
        self.mail_subject.as_deref()
    }
    pub fn mail_as_attachment(&self) -> bool {
        self.mail_as_attachment
    }
    pub fn view_merged_data(&self) -> bool {
        self.view_merged_data
    }
    pub fn active_record(&self) -> i32 {
        self.active_record
    }
    pub fn check_errors(&self) -> i32 {
        self.check_errors
    }
    pub fn odso(&self) -> Option<&DataSourceObject> {
        self.odso.as_ref()
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Recipient {
    pub(super) active: bool,
    pub(super) column: Option<i32>,
    pub(super) unique_tag: Option<Vec<u8>>,
}

impl Recipient {
    pub fn new(active: bool, column: Option<i32>, unique_tag: Option<Vec<u8>>) -> Self {
        Self {
            active,
            column,
            unique_tag,
        }
    }
    pub fn set_active(&mut self, value: bool) -> &mut Self {
        self.active = value;
        self
    }
    pub fn set_column(&mut self, value: Option<i32>) -> &mut Self {
        self.column = value;
        self
    }
    pub fn set_unique_tag(&mut self, value: Option<Vec<u8>>) -> &mut Self {
        self.unique_tag = value;
        self
    }
    pub fn active(&self) -> bool {
        self.active
    }
    pub fn column(&self) -> Option<i32> {
        self.column
    }
    pub fn unique_tag(&self) -> Option<&[u8]> {
        self.unique_tag.as_deref()
    }
}

/// Bounded inclusion/exclusion metadata from the inert recipient-data part.
///
/// The checked-in `[MS-OE376]` recipient-data part and relationship sections
/// are treated as package metadata only: no source record is opened, fetched,
/// hashed, or executed by this codec.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct Recipients {
    pub(super) recipients: Vec<Recipient>,
}

impl Recipients {
    pub fn new() -> Self {
        Self::default()
    }
    pub fn recipients(&self) -> &[Recipient] {
        &self.recipients
    }

    pub fn recipients_mut(&mut self) -> &mut Vec<Recipient> {
        &mut self.recipients
    }
    pub fn add_recipient(&mut self, recipient: Recipient) -> Result<&mut Self> {
        if self.recipients.len() >= MAX_RECIPIENTS {
            return Err(invalid("too many mail-merge recipients"));
        }
        self.recipients.push(recipient);
        Ok(self)
    }
    pub fn set_recipient_active(&mut self, index: usize, active: bool) -> Result<()> {
        let recipient = self
            .recipients
            .get_mut(index)
            .ok_or_else(|| invalid(format!("recipient index {index} is out of range")))?;
        recipient.active = active;
        Ok(())
    }

    pub fn content_type() -> &'static str {
        super::RECIPIENT_CONTENT_TYPE
    }
}
