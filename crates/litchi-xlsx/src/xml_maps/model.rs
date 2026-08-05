//! Typed semantic models for SpreadsheetML Custom XML Maps.

pub(super) const NS: &[u8] = b"http://schemas.openxmlformats.org/spreadsheetml/2006/main";
pub(super) const STRICT_NS: &[u8] = b"http://purl.oclc.org/ooxml/spreadsheetml/main";
pub(super) const NS_TEXT: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
pub(super) const STRICT_NS_TEXT: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
pub(super) const REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/xmlMaps";
pub(super) const STRICT_REL: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/xmlMaps";
pub(super) const CONTENT_TYPE: &str = "application/xml";
pub(super) const MAX_PART_BYTES: usize = 32 * 1024 * 1024;
pub(super) const MAX_SCHEMAS: usize = 4_096;
pub(super) const MAX_MAPS: usize = 65_536;
pub(super) const MAX_STRING_BYTES: usize = 1024 * 1024;
pub(super) const MAX_OPAQUE_BYTES: usize = 16 * 1024 * 1024;
pub(super) const MAX_DEPTH: usize = 256;
pub(super) const MAX_EVENTS: usize = 1_000_000;

/// Namespace family used for a Custom XML Maps part and its workbook relationship.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub enum XmlMapConformance {
    #[default]
    Transitional,
    Strict,
}

impl XmlMapConformance {
    pub(super) const fn relationship_type(self) -> &'static str {
        match self {
            Self::Transitional => REL,
            Self::Strict => STRICT_REL,
        }
    }

    /// Whether this conformance uses ISO/IEC 29500 Strict namespace URIs.
    pub const fn is_strict(self) -> bool {
        matches!(self, Self::Strict)
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct XmlMapSchema {
    pub id: String,
    pub schema_reference: Option<String>,
    pub namespace: Option<String>,
    /// One schema-valid `xsd:any` element, stored without interpretation or resolution.
    pub payload_xml: Option<Vec<u8>>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct XmlMapDataBinding {
    pub data_binding_name: Option<String>,
    pub file_binding: Option<bool>,
    pub connection_id: Option<u32>,
    pub file_binding_name: Option<String>,
    pub load_mode: u32,
    /// One schema-valid `xsd:any` element, stored without interpretation or execution.
    pub payload_xml: Option<Vec<u8>>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct XmlMap {
    pub id: u32,
    pub name: String,
    pub root_element: String,
    pub schema_id: String,
    pub show_import_export_validation_errors: bool,
    pub auto_fit: bool,
    pub append: bool,
    pub preserve_sort_auto_filter_layout: bool,
    pub preserve_format: bool,
    pub data_binding: Option<XmlMapDataBinding>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct XmlMapInfo {
    pub selection_namespaces: String,
    pub schemas: Vec<XmlMapSchema>,
    pub maps: Vec<XmlMap>,
}
