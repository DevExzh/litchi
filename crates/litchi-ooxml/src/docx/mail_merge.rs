//! Inert WordprocessingML mail-merge settings and recipient metadata.

use crate::error::{OoxmlError, Result};
use base64::Engine;
use base64::engine::general_purpose::STANDARD as BASE64;
use litchi_opc::part::Part;
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const STRICT_W: &str = "http://purl.oclc.org/ooxml/wordprocessingml/main";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_R: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const RECIPIENT_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.mailMergeRecipientData+xml";
const MAX_DEPTH: usize = 128;
const MAX_NODES: usize = 1_000_000;
const MAX_STRING_BYTES: usize = 16 * 1024 * 1024;
const MAX_RELATIONSHIP_ID_BYTES: usize = 1024;
const MAX_FIELD_MAPS: usize = 16_384;
const MAX_RECIPIENTS: usize = 1_000_000;
const MAX_UNIQUE_TAG_BYTES: usize = 1024 * 1024;

/// Opaque mail-merge source to relate from `settings.xml`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum MailMergeSource {
    /// Bytes are stored as an inert package part and never opened or interpreted.
    Internal {
        bytes: Vec<u8>,
        content_type: String,
        extension: String,
    },
    /// URI is stored as an external relationship and never fetched.
    External(String),
}

/// Owned, inert relationship target returned by package lookup.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum MailMergeTarget {
    Internal {
        part_name: litchi_opc::PackURI,
        bytes: Vec<u8>,
        content_type: String,
    },
    External(String),
}

fn invalid(message: impl Into<String>) -> OoxmlError {
    OoxmlError::InvalidFormat(message.into())
}

/// Namespace family used by deterministic mail-merge serializers.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MailMergeConformance {
    Transitional,
    Strict,
}

impl MailMergeConformance {
    fn word(self) -> &'static str {
        match self {
            Self::Transitional => W,
            Self::Strict => STRICT_W,
        }
    }

    fn relationships(self) -> &'static str {
        match self {
            Self::Transitional => R,
            Self::Strict => STRICT_R,
        }
    }
}

// The explicit implementations keep rustdoc useful and defaults visible.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum MailMergeMainDocumentType {
    Catalog,
    Envelopes,
    MailingLabels,
    #[default]
    FormLetters,
    Email,
    Fax,
}

impl MailMergeMainDocumentType {
    fn parse(value: &str) -> Result<Self> {
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
pub enum MailMergeDataType {
    TextFile,
    Database,
    Spreadsheet,
    Query,
    Odbc,
    Native,
}

impl MailMergeDataType {
    fn parse(value: &str) -> Result<Self> {
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
pub enum MailMergeDestination {
    #[default]
    NewDocument,
    Printer,
    Email,
    Fax,
}

impl MailMergeDestination {
    fn parse(value: &str) -> Result<Self> {
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
pub enum MailMergeFieldMappingType {
    Null,
    DatabaseColumn,
}

impl MailMergeFieldMappingType {
    fn parse(value: &str) -> Result<Self> {
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
pub struct MailMergeFieldMap {
    mapping_type: Option<MailMergeFieldMappingType>,
    name: Option<String>,
    mapped_name: Option<String>,
    column: Option<i32>,
    language_id: Option<String>,
    dynamic_address: bool,
}

impl MailMergeFieldMap {
    pub fn new() -> Self { Self::default() }
    pub fn set_mapping_type(&mut self, value: Option<MailMergeFieldMappingType>) -> &mut Self { self.mapping_type = value; self }
    pub fn set_name(&mut self, value: Option<String>) -> &mut Self { self.name = value; self }
    pub fn set_mapped_name(&mut self, value: Option<String>) -> &mut Self { self.mapped_name = value; self }
    pub fn set_column(&mut self, value: Option<i32>) -> &mut Self { self.column = value; self }
    pub fn set_language_id(&mut self, value: Option<String>) -> &mut Self { self.language_id = value; self }
    pub fn set_dynamic_address(&mut self, value: bool) -> &mut Self { self.dynamic_address = value; self }
    pub fn mapping_type(&self) -> Option<MailMergeFieldMappingType> {
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
pub struct MailMergeDataSourceObject {
    udl: Option<String>,
    table: Option<String>,
    source_relationship_id: Option<String>,
    column_delimiter: Option<i32>,
    source_type: Option<String>,
    first_row_header: bool,
    field_maps: Vec<MailMergeFieldMap>,
    recipient_data_relationship_id: Option<String>,
}

impl MailMergeDataSourceObject {
    pub fn new() -> Self { Self::default() }
    pub fn set_udl(&mut self, value: Option<String>) -> &mut Self { self.udl = value; self }
    pub fn set_table(&mut self, value: Option<String>) -> &mut Self { self.table = value; self }
    pub fn set_column_delimiter(&mut self, value: Option<i32>) -> &mut Self { self.column_delimiter = value; self }
    pub fn set_source_type(&mut self, value: Option<String>) -> &mut Self { self.source_type = value; self }
    pub fn set_first_row_header(&mut self, value: bool) -> &mut Self { self.first_row_header = value; self }
    pub fn field_maps_mut(&mut self) -> &mut Vec<MailMergeFieldMap> { &mut self.field_maps }
    pub fn add_field_map(&mut self, value: MailMergeFieldMap) -> &mut Self { self.field_maps.push(value); self }
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
    pub fn field_maps(&self) -> &[MailMergeFieldMap] {
        &self.field_maps
    }
    pub fn recipient_data_relationship_id(&self) -> Option<&str> {
        self.recipient_data_relationship_id.as_deref()
    }
}

/// Complete inert metadata from `w:mailMerge`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MailMergeSettings {
    main_document_type: MailMergeMainDocumentType,
    link_to_query: bool,
    data_type: Option<MailMergeDataType>,
    connect_string: Option<String>,
    query: Option<String>,
    data_source_relationship_id: Option<String>,
    header_source_relationship_id: Option<String>,
    do_not_suppress_blank_lines: bool,
    destination: MailMergeDestination,
    address_field_name: Option<String>,
    mail_subject: Option<String>,
    mail_as_attachment: bool,
    view_merged_data: bool,
    active_record: i32,
    check_errors: i32,
    odso: Option<MailMergeDataSourceObject>,
}

impl Default for MailMergeSettings {
    fn default() -> Self {
        Self {
            main_document_type: MailMergeMainDocumentType::FormLetters,
            link_to_query: false,
            data_type: None,
            connect_string: None,
            query: None,
            data_source_relationship_id: None,
            header_source_relationship_id: None,
            do_not_suppress_blank_lines: false,
            destination: MailMergeDestination::NewDocument,
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

impl MailMergeSettings {
    pub fn new() -> Self { Self::default() }
    pub fn set_main_document_type(&mut self, value: MailMergeMainDocumentType) -> &mut Self { self.main_document_type = value; self }
    pub fn set_link_to_query(&mut self, value: bool) -> &mut Self { self.link_to_query = value; self }
    pub fn set_data_type(&mut self, value: Option<MailMergeDataType>) -> &mut Self { self.data_type = value; self }
    pub fn set_connect_string(&mut self, value: Option<String>) -> &mut Self { self.connect_string = value; self }
    pub fn set_query(&mut self, value: Option<String>) -> &mut Self { self.query = value; self }
    pub fn set_do_not_suppress_blank_lines(&mut self, value: bool) -> &mut Self { self.do_not_suppress_blank_lines = value; self }
    pub fn set_destination(&mut self, value: MailMergeDestination) -> &mut Self { self.destination = value; self }
    pub fn set_address_field_name(&mut self, value: Option<String>) -> &mut Self { self.address_field_name = value; self }
    pub fn set_mail_subject(&mut self, value: Option<String>) -> &mut Self { self.mail_subject = value; self }
    pub fn set_mail_as_attachment(&mut self, value: bool) -> &mut Self { self.mail_as_attachment = value; self }
    pub fn set_view_merged_data(&mut self, value: bool) -> &mut Self { self.view_merged_data = value; self }
    pub fn set_active_record(&mut self, value: i32) -> &mut Self { self.active_record = value; self }
    pub fn set_check_errors(&mut self, value: i32) -> &mut Self { self.check_errors = value; self }
    pub fn set_odso(&mut self, value: Option<MailMergeDataSourceObject>) -> &mut Self { self.odso = value; self }

    pub(crate) fn assign_package_relationships(
        &mut self,
        data_source: Option<String>,
        header_source: Option<String>,
        recipient_data: Option<String>,
    ) {
        self.data_source_relationship_id = data_source.clone();
        self.header_source_relationship_id = header_source;
        if self.odso.is_none() && (data_source.is_some() || recipient_data.is_some()) {
            self.odso = Some(MailMergeDataSourceObject::default());
        }
        if let Some(odso) = &mut self.odso {
            odso.source_relationship_id = data_source;
            odso.recipient_data_relationship_id = recipient_data;
        }
    }
    pub fn main_document_type(&self) -> MailMergeMainDocumentType {
        self.main_document_type
    }
    pub fn link_to_query(&self) -> bool {
        self.link_to_query
    }
    pub fn data_type(&self) -> Option<MailMergeDataType> {
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
    pub fn destination(&self) -> MailMergeDestination {
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
    pub fn odso(&self) -> Option<&MailMergeDataSourceObject> {
        self.odso.as_ref()
    }

    /// Serialize a standalone `w:mailMerge` fragment in schema order.
    pub fn to_xml(&self, conformance: MailMergeConformance) -> Result<String> {
        validate_model(self)?;
        let mut xml = format!(
            r#"<w:mailMerge xmlns:w="{}" xmlns:r="{}">"#,
            conformance.word(),
            conformance.relationships()
        );
        if self.main_document_type != MailMergeMainDocumentType::FormLetters {
            value_leaf(
                &mut xml,
                "mainDocumentType",
                self.main_document_type.as_str(),
            );
        }
        on_off_leaf(&mut xml, "linkToQuery", self.link_to_query);
        if let Some(value) = self.data_type {
            value_leaf(&mut xml, "dataType", value.as_str());
        }
        optional_string_leaf(&mut xml, "connectString", self.connect_string.as_deref());
        optional_string_leaf(&mut xml, "query", self.query.as_deref());
        relationship_leaf(
            &mut xml,
            "dataSource",
            self.data_source_relationship_id.as_deref(),
        );
        relationship_leaf(
            &mut xml,
            "headerSource",
            self.header_source_relationship_id.as_deref(),
        );
        on_off_leaf(
            &mut xml,
            "doNotSuppressBlankLines",
            self.do_not_suppress_blank_lines,
        );
        if self.destination != MailMergeDestination::NewDocument {
            value_leaf(&mut xml, "destination", self.destination.as_str());
        }
        optional_string_leaf(
            &mut xml,
            "addressFieldName",
            self.address_field_name.as_deref(),
        );
        optional_string_leaf(&mut xml, "mailSubject", self.mail_subject.as_deref());
        on_off_leaf(&mut xml, "mailAsAttachment", self.mail_as_attachment);
        on_off_leaf(&mut xml, "viewMergedData", self.view_merged_data);
        if self.active_record != 1 {
            value_leaf(&mut xml, "activeRecord", &self.active_record.to_string());
        }
        if self.check_errors != 2 {
            value_leaf(&mut xml, "checkErrors", &self.check_errors.to_string());
        }
        if let Some(odso) = &self.odso {
            write_odso(&mut xml, odso);
        }
        xml.push_str("</w:mailMerge>");
        Ok(xml)
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MailMergeRecipient {
    active: bool,
    column: Option<i32>,
    unique_tag: Option<Vec<u8>>,
}

impl MailMergeRecipient {
    pub fn new(active: bool, column: Option<i32>, unique_tag: Option<Vec<u8>>) -> Self {
        Self { active, column, unique_tag }
    }
    pub fn set_active(&mut self, value: bool) -> &mut Self { self.active = value; self }
    pub fn set_column(&mut self, value: Option<i32>) -> &mut Self { self.column = value; self }
    pub fn set_unique_tag(&mut self, value: Option<Vec<u8>>) -> &mut Self { self.unique_tag = value; self }
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

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct MailMergeRecipients {
    recipients: Vec<MailMergeRecipient>,
}

impl MailMergeRecipients {
    pub fn new() -> Self { Self::default() }
    pub fn recipients(&self) -> &[MailMergeRecipient] {
        &self.recipients
    }

    pub fn recipients_mut(&mut self) -> &mut Vec<MailMergeRecipient> { &mut self.recipients }
    pub fn add_recipient(&mut self, recipient: MailMergeRecipient) -> Result<&mut Self> {
        if self.recipients.len() >= MAX_RECIPIENTS {
            return Err(invalid("too many mail-merge recipients"));
        }
        self.recipients.push(recipient);
        Ok(self)
    }
    pub fn set_recipient_active(&mut self, index: usize, active: bool) -> Result<()> {
        let recipient = self.recipients.get_mut(index).ok_or_else(|| invalid(format!("recipient index {index} is out of range")))?;
        recipient.active = active;
        Ok(())
    }

    pub(crate) fn content_type() -> &'static str { RECIPIENT_CONTENT_TYPE }

    pub(crate) fn extract_from_part(part: &dyn Part) -> Result<Self> {
        if part.content_type() != RECIPIENT_CONTENT_TYPE {
            return Err(invalid(format!(
                "invalid mail-merge recipient-data content type '{}'",
                part.content_type()
            )));
        }
        if part.rels().iter().next().is_some() {
            return Err(invalid(
                "mail-merge recipient-data part cannot have relationships",
            ));
        }
        let xml = crate::common::mce::process_part(part)?;
        Self::extract_from_xml(xml.as_ref())
    }

    fn extract_from_xml(xml: &[u8]) -> Result<Self> {
        let root = parse_tree(xml)?;
        require_word_element(&root, "recipients")?;
        ensure_no_schema_attrs(&root)?;
        let mut recipients = Vec::new();
        for child in &root.children {
            if !child.is_word() {
                if child.local == "recipientData" {
                    return Err(invalid("spoofed recipientData element namespace"));
                }
                continue;
            }
            if child.local != "recipientData" {
                return Err(invalid(format!(
                    "unexpected recipients child '{}'",
                    child.local
                )));
            }
            if recipients.len() >= MAX_RECIPIENTS {
                return Err(invalid("too many mail-merge recipients"));
            }
            recipients.push(parse_recipient(child)?);
        }
        Ok(Self { recipients })
    }

    pub fn to_xml(&self, conformance: MailMergeConformance) -> Result<String> {
        if self.recipients.len() > MAX_RECIPIENTS {
            return Err(invalid("too many mail-merge recipients"));
        }
        let mut xml = format!(r#"<w:recipients xmlns:w="{}">"#, conformance.word());
        for recipient in &self.recipients {
            xml.push_str("<w:recipientData>");
            if !recipient.active {
                xml.push_str(r#"<w:active w:val="0"/>"#);
            }
            if let Some(column) = recipient.column {
                if column < 0 {
                    return Err(invalid("mail-merge recipient column cannot be negative"));
                }
                value_leaf(&mut xml, "column", &column.to_string());
            }
            if let Some(tag) = &recipient.unique_tag {
                if tag.is_empty() || tag.len() > MAX_UNIQUE_TAG_BYTES {
                    return Err(invalid(
                        "mail-merge recipient unique tag has invalid length",
                    ));
                }
                value_leaf(&mut xml, "uniqueTag", &BASE64.encode(tag));
            }
            xml.push_str("</w:recipientData>");
        }
        xml.push_str("</w:recipients>");
        Ok(xml)
    }
}

pub(crate) fn parse_settings_mail_merge(xml: &[u8]) -> Result<Option<MailMergeSettings>> {
    let root = parse_tree(xml)?;
    require_word_element(&root, "settings")?;
    let mut found = None;
    let mut mail_index = None;
    for (index, child) in root.children.iter().enumerate() {
        if child.local == "mailMerge" {
            if !child.is_word() {
                return Err(invalid("spoofed mailMerge element namespace"));
            }
            if found.is_some() {
                return Err(invalid("duplicate mailMerge setting"));
            }
            found = Some(parse_mail_merge(child)?);
            mail_index = Some(index);
        }
        reject_nested_mail_merge(child, child.local == "mailMerge")?;
    }
    if let Some(index) = mail_index {
        validate_settings_order(&root.children, index)?;
    }
    Ok(found)
}

fn reject_nested_mail_merge(node: &Node, is_direct: bool) -> Result<()> {
    for child in &node.children {
        if child.local == "mailMerge" && child.is_word() && !is_direct {
            return Err(invalid("mailMerge must be a direct settings child"));
        }
        reject_nested_mail_merge(child, false)?;
    }
    Ok(())
}

fn validate_settings_order(children: &[Node], mail_index: usize) -> Result<()> {
    const BEFORE: &[&str] = &[
        "writeProtection",
        "view",
        "zoom",
        "linkStyles",
        "removePersonalInformation",
        "removeDateAndTime",
        "doNotDisplayPageBoundaries",
        "displayBackgroundShape",
        "printPostScriptOverText",
        "printFractionalCharacterWidth",
        "printFormsData",
        "embedTrueTypeFonts",
        "embedSystemFonts",
        "saveSubsetFonts",
        "saveFormsData",
        "mirrorMargins",
        "alignBordersAndEdges",
        "bordersDoNotSurroundHeader",
        "bordersDoNotSurroundFooter",
        "gutterAtTop",
        "hideSpellingErrors",
        "hideGrammaticalErrors",
        "activeWritingStyle",
        "proofState",
        "formsDesign",
        "attachedTemplate",
        "stylePaneFormatFilter",
        "stylePaneSortMethod",
        "documentType",
    ];
    const AFTER: &[&str] = &[
        "revisionView",
        "trackRevisions",
        "doNotTrackMoves",
        "doNotTrackFormatting",
        "documentProtection",
        "autoFormatOverride",
        "styleLockTheme",
        "styleLockQFSet",
        "defaultTabStop",
    ];
    for child in &children[..mail_index] {
        if child.is_word() && AFTER.contains(&child.local.as_str()) {
            return Err(invalid(format!("{} must follow mailMerge", child.local)));
        }
    }
    for child in &children[mail_index + 1..] {
        if child.is_word() && BEFORE.contains(&child.local.as_str()) {
            return Err(invalid(format!("{} must precede mailMerge", child.local)));
        }
    }
    Ok(())
}

fn parse_mail_merge(node: &Node) -> Result<MailMergeSettings> {
    ensure_no_schema_attrs(node)?;
    let names = [
        "mainDocumentType",
        "linkToQuery",
        "dataType",
        "connectString",
        "query",
        "dataSource",
        "headerSource",
        "doNotSuppressBlankLines",
        "destination",
        "addressFieldName",
        "mailSubject",
        "mailAsAttachment",
        "viewMergedData",
        "activeRecord",
        "checkErrors",
        "odso",
    ];
    let mut seen = [false; 16];
    let mut last = 0usize;
    let mut first = true;
    let mut value = MailMergeSettings::default();
    for child in &node.children {
        if !child.is_word() {
            if names.contains(&child.local.as_str()) {
                return Err(invalid(format!(
                    "spoofed {} element namespace",
                    child.local
                )));
            }
            continue;
        }
        let Some(index) = names.iter().position(|name| *name == child.local) else {
            return Err(invalid(format!(
                "unexpected mailMerge child '{}'",
                child.local
            )));
        };
        if seen[index] {
            return Err(invalid(format!(
                "duplicate mailMerge child '{}'",
                child.local
            )));
        }
        if !first && index < last {
            return Err(invalid(format!(
                "mailMerge child '{}' is out of order",
                child.local
            )));
        }
        first = false;
        last = index;
        seen[index] = true;
        match index {
            0 => {
                value.main_document_type = MailMergeMainDocumentType::parse(&required_val(child)?)?
            },
            1 => value.link_to_query = on_off(child)?,
            2 => value.data_type = Some(MailMergeDataType::parse(&required_val(child)?)?),
            3 => {
                value.connect_string = Some(bounded_string(required_val(child)?, "connectString")?)
            },
            4 => value.query = Some(bounded_string(required_val(child)?, "query")?),
            5 => value.data_source_relationship_id = Some(relationship_id(child)?),
            6 => value.header_source_relationship_id = Some(relationship_id(child)?),
            7 => value.do_not_suppress_blank_lines = on_off(child)?,
            8 => value.destination = MailMergeDestination::parse(&required_val(child)?)?,
            9 => {
                value.address_field_name =
                    Some(bounded_string(required_val(child)?, "addressFieldName")?)
            },
            10 => value.mail_subject = Some(bounded_string(required_val(child)?, "mailSubject")?),
            11 => value.mail_as_attachment = on_off(child)?,
            12 => value.view_merged_data = on_off(child)?,
            13 => {
                value.active_record = decimal(child, "activeRecord")?;
                if value.active_record < 1 {
                    return Err(invalid("activeRecord must be at least 1"));
                }
            },
            14 => value.check_errors = decimal(child, "checkErrors")?,
            15 => value.odso = Some(parse_odso(child)?),
            _ => unreachable!(),
        }
    }
    validate_model(&value)?;
    Ok(value)
}

fn parse_odso(node: &Node) -> Result<MailMergeDataSourceObject> {
    ensure_no_schema_attrs(node)?;
    let names = [
        "udl",
        "table",
        "src",
        "colDelim",
        "type",
        "fHdr",
        "fieldMapData",
        "recipientData",
    ];
    let mut seen = [false; 8];
    let mut last = 0usize;
    let mut first = true;
    let mut value = MailMergeDataSourceObject::default();
    for child in &node.children {
        if !child.is_word() {
            if names.contains(&child.local.as_str()) {
                return Err(invalid(format!(
                    "spoofed ODSO {} element namespace",
                    child.local
                )));
            }
            continue;
        }
        let Some(index) = names.iter().position(|name| *name == child.local) else {
            return Err(invalid(format!("unexpected odso child '{}'", child.local)));
        };
        if index != 6 && seen[index] {
            return Err(invalid(format!("duplicate odso child '{}'", child.local)));
        }
        if !first && index < last {
            return Err(invalid(format!(
                "odso child '{}' is out of order",
                child.local
            )));
        }
        first = false;
        last = index;
        seen[index] = true;
        match index {
            0 => value.udl = Some(bounded_string(required_val(child)?, "odso udl")?),
            1 => value.table = Some(bounded_string(required_val(child)?, "odso table")?),
            2 => value.source_relationship_id = Some(relationship_id(child)?),
            3 => {
                let number = decimal(child, "colDelim")?;
                if number < 0 {
                    return Err(invalid("colDelim cannot be negative"));
                }
                value.column_delimiter = Some(number);
            },
            4 => {
                value.source_type = Some(bounded_string(required_val(child)?, "odso source type")?)
            },
            5 => value.first_row_header = on_off(child)?,
            6 => {
                if value.field_maps.len() >= MAX_FIELD_MAPS {
                    return Err(invalid("too many odso field maps"));
                }
                value.field_maps.push(parse_field_map(child)?);
            },
            7 => value.recipient_data_relationship_id = Some(relationship_id(child)?),
            _ => unreachable!(),
        }
    }
    Ok(value)
}

fn parse_field_map(node: &Node) -> Result<MailMergeFieldMap> {
    ensure_no_schema_attrs(node)?;
    let names = [
        "type",
        "name",
        "mappedName",
        "column",
        "lid",
        "dynamicAddress",
    ];
    let mut seen = [false; 6];
    let mut last = 0usize;
    let mut first = true;
    let mut value = MailMergeFieldMap::default();
    for child in &node.children {
        if !child.is_word() {
            if names.contains(&child.local.as_str()) {
                return Err(invalid(format!(
                    "spoofed fieldMapData {} element namespace",
                    child.local
                )));
            }
            continue;
        }
        let Some(index) = names.iter().position(|name| *name == child.local) else {
            return Err(invalid(format!(
                "unexpected fieldMapData child '{}'",
                child.local
            )));
        };
        if seen[index] {
            return Err(invalid(format!(
                "duplicate fieldMapData child '{}'",
                child.local
            )));
        }
        if !first && index < last {
            return Err(invalid(format!(
                "fieldMapData child '{}' is out of order",
                child.local
            )));
        }
        first = false;
        last = index;
        seen[index] = true;
        match index {
            0 => {
                value.mapping_type = Some(MailMergeFieldMappingType::parse(&required_val(child)?)?)
            },
            1 => value.name = Some(bounded_string(required_val(child)?, "field-map name")?),
            2 => {
                value.mapped_name = Some(bounded_string(required_val(child)?, "mapped field name")?)
            },
            3 => {
                let number = decimal(child, "field-map column")?;
                if number < 0 {
                    return Err(invalid("field-map column cannot be negative"));
                }
                value.column = Some(number);
            },
            4 => {
                value.language_id =
                    Some(bounded_string(required_val(child)?, "field-map language")?)
            },
            5 => value.dynamic_address = on_off(child)?,
            _ => unreachable!(),
        }
    }
    Ok(value)
}

fn parse_recipient(node: &Node) -> Result<MailMergeRecipient> {
    ensure_no_schema_attrs(node)?;
    let names = ["active", "column", "uniqueTag"];
    let mut seen = [false; 3];
    let mut last = 0usize;
    let mut first = true;
    let mut recipient = MailMergeRecipient {
        active: true,
        column: None,
        unique_tag: None,
    };
    for child in &node.children {
        if !child.is_word() {
            if names.contains(&child.local.as_str()) {
                return Err(invalid(format!(
                    "spoofed recipient {} element namespace",
                    child.local
                )));
            }
            continue;
        }
        let Some(index) = names.iter().position(|name| *name == child.local) else {
            return Err(invalid(format!(
                "unexpected recipientData child '{}'",
                child.local
            )));
        };
        if seen[index] {
            return Err(invalid(format!(
                "duplicate recipientData child '{}'",
                child.local
            )));
        }
        if !first && index < last {
            return Err(invalid(format!(
                "recipientData child '{}' is out of order",
                child.local
            )));
        }
        first = false;
        last = index;
        seen[index] = true;
        match index {
            0 => recipient.active = on_off(child)?,
            1 => {
                let number = decimal(child, "recipient column")?;
                if number < 0 {
                    return Err(invalid("recipient column cannot be negative"));
                }
                recipient.column = Some(number);
            },
            2 => recipient.unique_tag = Some(strict_base64(&required_val(child)?)?),
            _ => unreachable!(),
        }
    }
    Ok(recipient)
}

fn validate_model(value: &MailMergeSettings) -> Result<()> {
    for (description, string) in [
        ("connectString", value.connect_string.as_deref()),
        ("query", value.query.as_deref()),
        ("addressFieldName", value.address_field_name.as_deref()),
        ("mailSubject", value.mail_subject.as_deref()),
    ] {
        if string.is_some_and(|text| text.len() > MAX_STRING_BYTES) {
            return Err(invalid(format!("{description} is too large")));
        }
    }
    if value.active_record < 1 {
        return Err(invalid("activeRecord must be at least 1"));
    }
    if let Some(odso) = &value.odso {
        if odso.field_maps.len() > MAX_FIELD_MAPS {
            return Err(invalid("too many odso field maps"));
        }
    }
    Ok(())
}

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

fn write_odso(xml: &mut String, odso: &MailMergeDataSourceObject) {
    xml.push_str("<w:odso>");
    optional_string_leaf(xml, "udl", odso.udl.as_deref());
    optional_string_leaf(xml, "table", odso.table.as_deref());
    relationship_leaf(xml, "src", odso.source_relationship_id.as_deref());
    if let Some(value) = odso.column_delimiter {
        value_leaf(xml, "colDelim", &value.to_string());
    }
    optional_string_leaf(xml, "type", odso.source_type.as_deref());
    on_off_leaf(xml, "fHdr", odso.first_row_header);
    for field in &odso.field_maps {
        xml.push_str("<w:fieldMapData>");
        if let Some(value) = field.mapping_type {
            value_leaf(xml, "type", value.as_str());
        }
        optional_string_leaf(xml, "name", field.name.as_deref());
        optional_string_leaf(xml, "mappedName", field.mapped_name.as_deref());
        if let Some(value) = field.column {
            value_leaf(xml, "column", &value.to_string());
        }
        optional_string_leaf(xml, "lid", field.language_id.as_deref());
        on_off_leaf(xml, "dynamicAddress", field.dynamic_address);
        xml.push_str("</w:fieldMapData>");
    }
    relationship_leaf(
        xml,
        "recipientData",
        odso.recipient_data_relationship_id.as_deref(),
    );
    xml.push_str("</w:odso>");
}

fn value_leaf(xml: &mut String, name: &str, value: &str) {
    xml.push_str("<w:");
    xml.push_str(name);
    xml.push_str(" w:val=\"");
    escape(xml, value);
    xml.push_str("\"/>");
}
fn optional_string_leaf(xml: &mut String, name: &str, value: Option<&str>) {
    if let Some(value) = value {
        value_leaf(xml, name, value);
    }
}
fn relationship_leaf(xml: &mut String, name: &str, value: Option<&str>) {
    if let Some(value) = value {
        xml.push_str("<w:");
        xml.push_str(name);
        xml.push_str(" r:id=\"");
        escape(xml, value);
        xml.push_str("\"/>");
    }
}
fn on_off_leaf(xml: &mut String, name: &str, value: bool) {
    if value {
        xml.push_str("<w:");
        xml.push_str(name);
        xml.push_str("/>");
    }
}
fn escape(xml: &mut String, value: &str) {
    for character in value.chars() {
        match character {
            '&' => xml.push_str("&amp;"),
            '<' => xml.push_str("&lt;"),
            '>' => xml.push_str("&gt;"),
            '"' => xml.push_str("&quot;"),
            '\'' => xml.push_str("&apos;"),
            _ => xml.push(character),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum OwnedNamespace {
    Bound(String),
    Unbound,
    Unknown(String),
}

#[derive(Debug, Clone)]
struct Attribute {
    namespace: OwnedNamespace,
    local: String,
    value: String,
}

#[derive(Debug, Clone)]
struct Node {
    namespace: OwnedNamespace,
    local: String,
    attributes: Vec<Attribute>,
    children: Vec<Node>,
    has_text: bool,
}

impl Node {
    fn is_word(&self) -> bool {
        matches!(&self.namespace, OwnedNamespace::Bound(value) if value == W || value == STRICT_W)
    }
}

fn parse_tree(xml: &[u8]) -> Result<Node> {
    let processed = crate::common::mce::process_ooxml(xml)
        .map_err(|error| invalid(format!("mail-merge MCE error: {error}")))?;
    let mut reader = NsReader::from_reader(processed.as_ref());
    reader.config_mut().trim_text(false);
    let mut stack: Vec<Node> = Vec::new();
    let mut root = None;
    let mut nodes = 0usize;
    loop {
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| OoxmlError::Xml(error.to_string()))?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => {
                if stack.len() >= MAX_DEPTH {
                    return Err(invalid("mail-merge XML nesting is too deep"));
                }
                nodes = nodes
                    .checked_add(1)
                    .ok_or_else(|| invalid("mail-merge XML node count overflow"))?;
                if nodes > MAX_NODES {
                    return Err(invalid("mail-merge XML has too many nodes"));
                }
                stack.push(make_node(namespace, &element, decoder, &resolver)?);
            },
            Event::Empty(element) => {
                nodes = nodes
                    .checked_add(1)
                    .ok_or_else(|| invalid("mail-merge XML node count overflow"))?;
                if nodes > MAX_NODES {
                    return Err(invalid("mail-merge XML has too many nodes"));
                }
                append_node(
                    make_node(namespace, &element, decoder, &resolver)?,
                    &mut stack,
                    &mut root,
                )?;
            },
            Event::End(_) => {
                let node = stack
                    .pop()
                    .ok_or_else(|| invalid("unexpected mail-merge XML end element"))?;
                append_node(node, &mut stack, &mut root)?;
            },
            Event::Text(text) => {
                if !text
                    .decode()
                    .map_err(|error| OoxmlError::Xml(error.to_string()))?
                    .trim()
                    .is_empty()
                {
                    if let Some(node) = stack.last_mut() {
                        node.has_text = true;
                    } else {
                        return Err(invalid("text outside mail-merge XML root"));
                    }
                }
            },
            Event::CData(text) => {
                if !text
                    .decode()
                    .map_err(|error| OoxmlError::Xml(error.to_string()))?
                    .trim()
                    .is_empty()
                {
                    if let Some(node) = stack.last_mut() {
                        node.has_text = true;
                    } else {
                        return Err(invalid("CDATA outside mail-merge XML root"));
                    }
                }
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid(
                    "DTD and processing instructions are rejected in mail-merge XML",
                ));
            },
            Event::Eof if !stack.is_empty() => return Err(invalid("unterminated mail-merge XML")),
            Event::Eof => break,
            _ => {},
        }
    }
    root.ok_or_else(|| invalid("mail-merge XML has no root element"))
}

fn make_node(
    namespace: ResolveResult<'_>,
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    resolver: &NamespaceResolver,
) -> Result<Node> {
    let mut attributes = Vec::new();
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| OoxmlError::Xml(error.to_string()))?;
        let raw = attribute.key.as_ref();
        if raw == b"xmlns" || raw.starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, _) = resolver.resolve_attribute(attribute.key);
        attributes.push(Attribute {
            namespace: own_namespace(namespace),
            local: String::from_utf8_lossy(attribute.key.local_name().as_ref()).into_owned(),
            value: attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .map_err(|error| OoxmlError::Xml(error.to_string()))?
                .into_owned(),
        });
    }
    Ok(Node {
        namespace: own_namespace(namespace),
        local: String::from_utf8_lossy(element.local_name().as_ref()).into_owned(),
        attributes,
        children: Vec::new(),
        has_text: false,
    })
}

fn own_namespace(namespace: ResolveResult<'_>) -> OwnedNamespace {
    match namespace {
        ResolveResult::Bound(Namespace(value)) => {
            OwnedNamespace::Bound(String::from_utf8_lossy(value).into_owned())
        },
        ResolveResult::Unbound => OwnedNamespace::Unbound,
        ResolveResult::Unknown(prefix) => {
            OwnedNamespace::Unknown(String::from_utf8_lossy(&prefix).into_owned())
        },
    }
}

fn append_node(node: Node, stack: &mut [Node], root: &mut Option<Node>) -> Result<()> {
    if let Some(parent) = stack.last_mut() {
        parent.children.push(node);
    } else if root.replace(node).is_some() {
        return Err(invalid("mail-merge XML has multiple root elements"));
    }
    Ok(())
}

fn require_word_element(node: &Node, local: &str) -> Result<()> {
    if !node.is_word() || node.local != local {
        return Err(invalid(format!("expected WordprocessingML {local} root")));
    }
    Ok(())
}

fn ensure_leaf(node: &Node) -> Result<()> {
    if !node.children.is_empty() || node.has_text {
        return Err(invalid(format!(
            "mail-merge leaf '{}' cannot contain content",
            node.local
        )));
    }
    Ok(())
}

fn ensure_no_schema_attrs(node: &Node) -> Result<()> {
    for attribute in &node.attributes {
        if matches!(&attribute.namespace, OwnedNamespace::Bound(value) if value == W || value == STRICT_W)
            || matches!(
                attribute.namespace,
                OwnedNamespace::Unbound | OwnedNamespace::Unknown(_)
            )
        {
            return Err(invalid(format!(
                "unexpected attribute '{}' on {}",
                attribute.local, node.local
            )));
        }
    }
    if node.has_text {
        return Err(invalid(format!("{} cannot contain text", node.local)));
    }
    Ok(())
}

fn required_val(node: &Node) -> Result<String> {
    ensure_leaf(node)?;
    schema_attribute(node, "val", false)?
        .ok_or_else(|| invalid(format!("{} requires w:val", node.local)))
}

fn relationship_id(node: &Node) -> Result<String> {
    ensure_leaf(node)?;
    let value = schema_attribute(node, "id", true)?
        .ok_or_else(|| invalid(format!("{} requires r:id", node.local)))?;
    if value.is_empty() || value.len() > MAX_RELATIONSHIP_ID_BYTES {
        return Err(invalid(format!(
            "{} has invalid relationship id length",
            node.local
        )));
    }
    Ok(value)
}

fn schema_attribute(node: &Node, local: &str, relationship: bool) -> Result<Option<String>> {
    let mut result = None;
    for attribute in &node.attributes {
        let expected_namespace = if relationship {
            matches!(&attribute.namespace, OwnedNamespace::Bound(value) if value == R || value == STRICT_R)
        } else {
            matches!(&attribute.namespace, OwnedNamespace::Bound(value) if value == W || value == STRICT_W)
                || matches!(attribute.namespace, OwnedNamespace::Unbound)
        };
        if attribute.local == local {
            if !expected_namespace {
                return Err(invalid(format!(
                    "{} has spoofed {} attribute namespace",
                    node.local, local
                )));
            }
            if result.replace(attribute.value.clone()).is_some() {
                return Err(invalid(format!(
                    "{} has duplicate {} attribute",
                    node.local, local
                )));
            }
        } else if expected_namespace
            || matches!(
                attribute.namespace,
                OwnedNamespace::Unbound | OwnedNamespace::Unknown(_)
            )
        {
            return Err(invalid(format!(
                "unexpected attribute '{}' on {}",
                attribute.local, node.local
            )));
        }
    }
    Ok(result)
}

fn on_off(node: &Node) -> Result<bool> {
    ensure_leaf(node)?;
    match schema_attribute(node, "val", false)?.as_deref() {
        None | Some("true" | "1" | "on") => Ok(true),
        Some("false" | "0" | "off") => Ok(false),
        Some(value) => Err(invalid(format!("invalid on/off value '{value}'"))),
    }
}

fn decimal(node: &Node, description: &str) -> Result<i32> {
    required_val(node)?.parse::<i32>().map_err(|_| {
        invalid(format!(
            "{description} is outside the supported 32-bit bound"
        ))
    })
}

fn bounded_string(value: String, description: &str) -> Result<String> {
    if value.len() > MAX_STRING_BYTES {
        return Err(invalid(format!("{description} is too large")));
    }
    Ok(value)
}

fn strict_base64(value: &str) -> Result<Vec<u8>> {
    let compact: String = value
        .chars()
        .filter(|character| !character.is_ascii_whitespace())
        .collect();
    let decoded = BASE64
        .decode(compact.as_bytes())
        .map_err(|_| invalid("invalid recipient uniqueTag base64"))?;
    if decoded.is_empty()
        || decoded.len() > MAX_UNIQUE_TAG_BYTES
        || BASE64.encode(&decoded) != compact
    {
        return Err(invalid(
            "recipient uniqueTag base64 is empty, non-canonical, or too large",
        ));
    }
    Ok(decoded)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::docx::Package;
    use std::io::Cursor;
    use std::path::Path;

    const SETTINGS: &str = r#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:mailMerge><w:mainDocumentType w:val="email"/><w:linkToQuery/><w:dataType w:val="native"/><w:connectString w:val="Provider=Inert&amp;X=1"/><w:query w:val="SELECT * FROM inert"/><w:dataSource r:id="source"/><w:destination w:val="email"/><w:mailSubject w:val="A &amp; B"/><w:mailAsAttachment/><w:viewMergedData/><w:activeRecord w:val="3"/><w:checkErrors w:val="3"/><w:odso><w:table w:val="Sheet1$"/><w:src r:id="source"/><w:colDelim w:val="9"/><w:type w:val="database"/><w:fHdr/><w:fieldMapData><w:type w:val="dbColumn"/><w:name w:val="Name"/><w:mappedName w:val="Last Name"/><w:column w:val="0"/><w:lid w:val="en-US"/><w:dynamicAddress/></w:fieldMapData><w:recipientData r:id="recipients"/></w:odso></w:mailMerge><w:trackRevisions/></w:settings>"#;

    #[test]
    fn parses_and_deterministically_writes_complete_strict_metadata() {
        let value = parse_settings_mail_merge(SETTINGS.as_bytes())
            .unwrap()
            .unwrap();
        assert_eq!(value.main_document_type(), MailMergeMainDocumentType::Email);
        assert_eq!(value.data_type(), Some(MailMergeDataType::Native));
        assert_eq!(value.destination(), MailMergeDestination::Email);
        assert_eq!(value.query(), Some("SELECT * FROM inert"));
        assert_eq!(
            value.odso().unwrap().field_maps()[0].mapping_type(),
            Some(MailMergeFieldMappingType::DatabaseColumn)
        );
        let fragment = value.to_xml(MailMergeConformance::Strict).unwrap();
        let wrapped = format!(r#"<s:settings xmlns:s="{STRICT_W}">{fragment}</s:settings>"#);
        let reparsed = parse_settings_mail_merge(wrapped.as_bytes())
            .unwrap()
            .unwrap();
        assert_eq!(reparsed, value);
        assert_eq!(
            reparsed.to_xml(MailMergeConformance::Strict).unwrap(),
            fragment
        );
    }

    #[test]
    fn applies_defaults_and_mce_fallback_and_preservation() {
        let xml = format!(
            r#"<w:settings xmlns:w="{W}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:future" mc:Ignorable="x"><w:mailMerge mc:Ignorable="x" mc:PreserveAttributes="x:*" x:future="kept"><mc:AlternateContent><mc:Choice Requires="x"><x:dataType/></mc:Choice><mc:Fallback><w:viewMergedData/></mc:Fallback></mc:AlternateContent></w:mailMerge></w:settings>"#
        );
        let value = parse_settings_mail_merge(xml.as_bytes()).unwrap().unwrap();
        assert_eq!(
            value.main_document_type(),
            MailMergeMainDocumentType::FormLetters
        );
        assert_eq!(value.destination(), MailMergeDestination::NewDocument);
        assert_eq!(value.active_record(), 1);
        assert_eq!(value.check_errors(), 2);
        assert!(value.view_merged_data());
    }

    #[test]
    fn rejects_malformed_scoped_ordered_bounded_metadata() {
        let invalid = [
            format!(
                r#"<w:settings xmlns:w="{W}"><w:mailMerge><w:dataType w:val="native"/><w:mainDocumentType w:val="email"/></w:mailMerge></w:settings>"#
            ),
            format!(
                r#"<w:settings xmlns:w="{W}"><w:mailMerge><w:dataType w:val="bogus"/></w:mailMerge></w:settings>"#
            ),
            format!(r#"<w:settings xmlns:w="{W}"><w:trackRevisions/><w:mailMerge/></w:settings>"#),
            format!(r#"<w:settings xmlns:w="{W}"><w:zoom><w:mailMerge/></w:zoom></w:settings>"#),
            format!(
                r#"<w:settings xmlns:w="{W}"><w:mailMerge><w:activeRecord w:val="0"/></w:mailMerge></w:settings>"#
            ),
            format!(
                r#"<w:settings xmlns:w="{W}"><w:mailMerge><w:checkErrors w:val="2147483648"/></w:mailMerge></w:settings>"#
            ),
            format!(
                r#"<w:settings xmlns:w="{W}" xmlns:x="urn:fake"><w:mailMerge><w:dataSource x:id="rId1"/></w:mailMerge></w:settings>"#
            ),
            format!(
                r#"<w:settings xmlns:w="{W}"><w:mailMerge><w:query/></w:mailMerge></w:settings>"#
            ),
        ];
        for xml in invalid {
            assert!(
                parse_settings_mail_merge(xml.as_bytes()).is_err(),
                "accepted {xml}"
            );
        }
        let recipients = format!(
            r#"<w:recipients xmlns:w="{W}"><w:recipientData><w:uniqueTag w:val="AQ="/></w:recipientData></w:recipients>"#
        );
        assert!(MailMergeRecipients::extract_from_xml(recipients.as_bytes()).is_err());
    }

    #[test]
    fn opens_libreoffice_and_synthetic_packages_without_accessing_sources() {
        let fixture = Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../..//3rdparty/libreoffice-core/sw/qa/extras/ooxmlexport/data/mailmerge.docx");
        let package = Package::open(fixture).unwrap();
        let settings = package.document().unwrap().settings().unwrap().unwrap();
        let merge = settings.mail_merge().unwrap();
        assert_eq!(merge.data_type(), Some(MailMergeDataType::Native));
        assert_eq!(merge.odso().unwrap().field_maps().len(), 30);
        assert_eq!(merge.data_source_relationship_id(), Some("rId1"));

        let package = Package::from_reader(Cursor::new(synthetic_docx())).unwrap();
        let document = package.document().unwrap();
        let settings = document.settings().unwrap().unwrap();
        assert_eq!(settings.mail_merge().unwrap().mail_subject(), Some("A & B"));
        let recipients = document.mail_merge_recipients().unwrap().unwrap();
        assert_eq!(recipients.recipients().len(), 1);
        assert!(!recipients.recipients()[0].active());
        assert_eq!(
            recipients.recipients()[0].unique_tag(),
            Some(&[1, 2, 3][..])
        );
        let strict = recipients.to_xml(MailMergeConformance::Strict).unwrap();
        assert_eq!(
            MailMergeRecipients::extract_from_xml(strict.as_bytes()).unwrap(),
            recipients
        );
    }

    fn synthetic_docx() -> Vec<u8> {
        let content_types = r#"<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/><Override PartName="/word/recipientData.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.mailMergeRecipientData+xml"/></Types>"#;
        let root_rels = r#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#;
        let document = format!(
            r#"<?xml version="1.0"?><w:document xmlns:w="{W}"><w:body><w:p/></w:body></w:document>"#
        );
        let document_rels = r#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="settings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>"#;
        let settings_rels = r#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="source" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/mailMergeSource" Target="https://example.invalid/inert.csv" TargetMode="External"/><Relationship Id="recipients" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/recipientData" Target="recipientData.xml"/></Relationships>"#;
        let recipients = format!(
            r#"<?xml version="1.0"?><w:recipients xmlns:w="{W}"><w:recipientData><w:active w:val="0"/><w:column w:val="7"/><w:uniqueTag w:val="AQID"/></w:recipientData></w:recipients>"#
        );
        stored_zip(&[
            ("[Content_Types].xml", content_types.as_bytes()),
            ("_rels/.rels", root_rels.as_bytes()),
            ("word/document.xml", document.as_bytes()),
            ("word/_rels/document.xml.rels", document_rels.as_bytes()),
            ("word/settings.xml", SETTINGS.as_bytes()),
            ("word/_rels/settings.xml.rels", settings_rels.as_bytes()),
            ("word/recipientData.xml", recipients.as_bytes()),
        ])
    }

    fn stored_zip(entries: &[(&str, &[u8])]) -> Vec<u8> {
        let mut output = Vec::new();
        let mut central = Vec::new();
        for (name, data) in entries {
            let offset = output.len() as u32;
            let crc = crc32(data);
            output.extend_from_slice(&0x04034b50u32.to_le_bytes());
            output.extend_from_slice(&20u16.to_le_bytes());
            output.extend_from_slice(&0u16.to_le_bytes());
            output.extend_from_slice(&0u16.to_le_bytes());
            output.extend_from_slice(&0u16.to_le_bytes());
            output.extend_from_slice(&0u16.to_le_bytes());
            output.extend_from_slice(&crc.to_le_bytes());
            output.extend_from_slice(&(data.len() as u32).to_le_bytes());
            output.extend_from_slice(&(data.len() as u32).to_le_bytes());
            output.extend_from_slice(&(name.len() as u16).to_le_bytes());
            output.extend_from_slice(&0u16.to_le_bytes());
            output.extend_from_slice(name.as_bytes());
            output.extend_from_slice(data);
            central.extend_from_slice(&0x02014b50u32.to_le_bytes());
            central.extend_from_slice(&20u16.to_le_bytes());
            central.extend_from_slice(&20u16.to_le_bytes());
            central.extend_from_slice(&0u16.to_le_bytes());
            central.extend_from_slice(&0u16.to_le_bytes());
            central.extend_from_slice(&0u16.to_le_bytes());
            central.extend_from_slice(&0u16.to_le_bytes());
            central.extend_from_slice(&crc.to_le_bytes());
            central.extend_from_slice(&(data.len() as u32).to_le_bytes());
            central.extend_from_slice(&(data.len() as u32).to_le_bytes());
            central.extend_from_slice(&(name.len() as u16).to_le_bytes());
            central.extend_from_slice(&0u16.to_le_bytes());
            central.extend_from_slice(&0u16.to_le_bytes());
            central.extend_from_slice(&0u16.to_le_bytes());
            central.extend_from_slice(&0u16.to_le_bytes());
            central.extend_from_slice(&0u32.to_le_bytes());
            central.extend_from_slice(&offset.to_le_bytes());
            central.extend_from_slice(name.as_bytes());
        }
        let central_offset = output.len() as u32;
        let central_size = central.len() as u32;
        output.extend_from_slice(&central);
        output.extend_from_slice(&0x06054b50u32.to_le_bytes());
        output.extend_from_slice(&0u16.to_le_bytes());
        output.extend_from_slice(&0u16.to_le_bytes());
        output.extend_from_slice(&(entries.len() as u16).to_le_bytes());
        output.extend_from_slice(&(entries.len() as u16).to_le_bytes());
        output.extend_from_slice(&central_size.to_le_bytes());
        output.extend_from_slice(&central_offset.to_le_bytes());
        output.extend_from_slice(&0u16.to_le_bytes());
        output
    }

    fn crc32(data: &[u8]) -> u32 {
        let mut crc = !0u32;
        for &byte in data {
            crc ^= u32::from(byte);
            for _ in 0..8 {
                crc = (crc >> 1) ^ (0xedb88320u32 & (0u32.wrapping_sub(crc & 1)));
            }
        }
        !crc
    }
}
