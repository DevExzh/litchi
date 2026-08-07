//! Typed semantic models for SpreadsheetML Custom XML Maps.

pub const NS: &[u8] = b"http://schemas.openxmlformats.org/spreadsheetml/2006/main";
pub const STRICT_NS: &[u8] = b"http://purl.oclc.org/ooxml/spreadsheetml/main";
pub const NS_TEXT: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
pub const STRICT_NS_TEXT: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
pub const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/xmlMaps";
pub const STRICT_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/xmlMaps";
pub const CONTENT_TYPE: &str = "application/xml";
pub const MAX_PART_BYTES: usize = 32 * 1024 * 1024;
pub const MAX_SCHEMAS: usize = 4_096;
pub const MAX_MAPS: usize = 65_536;
pub const MAX_STRING_BYTES: usize = 1024 * 1024;
pub const MAX_OPAQUE_BYTES: usize = 16 * 1024 * 1024;
pub const MAX_DEPTH: usize = 256;
pub const MAX_EVENTS: usize = 1_000_000;

/// Namespace family used for a Custom XML Maps part and its workbook relationship.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub enum XmlMapConformance {
    #[default]
    Transitional,
    Strict,
}

impl XmlMapConformance {
    pub const fn relationship_type(self) -> &'static str {
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
pub struct XmlSchema {
    pub id: String,
    pub schema_reference: Option<String>,
    pub namespace: Option<String>,
    /// One schema-valid `xsd:any` element, stored without interpretation or resolution.
    pub payload_xml: Option<Vec<u8>>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DataBinding {
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
    pub data_binding: Option<DataBinding>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct XmlMapInfo {
    pub selection_namespaces: String,
    pub schemas: Vec<XmlMapSchema>,
    pub maps: Vec<XmlMap>,
}

/// A parsed MapInfo value together with the namespace family observed at its root.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ParsedXmlMapInfo {
    pub info: XmlMapInfo,
    pub conformance: XmlMapConformance,
}

/// Backward-compatible SpreadsheetML name for [`XmlSchema`].
pub type XmlMapSchema = XmlSchema;
/// Backward-compatible SpreadsheetML name for [`DataBinding`].
pub type XmlMapDataBinding = DataBinding;

/// Fixed resource ceilings enforced by the bounded XML Maps codec.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct XmlMapLimits {
    pub max_part_bytes: usize,
    pub max_schemas: usize,
    pub max_maps: usize,
    pub max_string_bytes: usize,
    pub max_opaque_bytes: usize,
    pub max_depth: usize,
    pub max_events: usize,
}

impl XmlMapLimits {
    /// The resource ceilings used by all XML Maps parsing and serialization.
    pub const DEFAULT: Self = Self {
        max_part_bytes: MAX_PART_BYTES,
        max_schemas: MAX_SCHEMAS,
        max_maps: MAX_MAPS,
        max_string_bytes: MAX_STRING_BYTES,
        max_opaque_bytes: MAX_OPAQUE_BYTES,
        max_depth: MAX_DEPTH,
        max_events: MAX_EVENTS,
    };
}

impl Default for XmlMapLimits {
    fn default() -> Self {
        Self::DEFAULT
    }
}

/// Borrowed schema descriptor that never clones strings or opaque XML.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct XmlSchemaRef<'a> {
    pub id: &'a str,
    pub schema_reference: Option<&'a str>,
    pub namespace: Option<&'a str>,
    pub payload_xml: Option<&'a [u8]>,
}

impl<'a> From<&'a XmlSchema> for XmlSchemaRef<'a> {
    fn from(value: &'a XmlSchema) -> Self {
        Self {
            id: &value.id,
            schema_reference: value.schema_reference.as_deref(),
            namespace: value.namespace.as_deref(),
            payload_xml: value.payload_xml.as_deref(),
        }
    }
}

/// Borrowed data-binding descriptor.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct DataBindingRef<'a> {
    pub data_binding_name: Option<&'a str>,
    pub file_binding: Option<bool>,
    pub connection_id: Option<u32>,
    pub file_binding_name: Option<&'a str>,
    pub load_mode: u32,
    pub payload_xml: Option<&'a [u8]>,
}

impl<'a> From<&'a DataBinding> for DataBindingRef<'a> {
    fn from(value: &'a DataBinding) -> Self {
        Self {
            data_binding_name: value.data_binding_name.as_deref(),
            file_binding: value.file_binding,
            connection_id: value.connection_id,
            file_binding_name: value.file_binding_name.as_deref(),
            load_mode: value.load_mode,
            payload_xml: value.payload_xml.as_deref(),
        }
    }
}

/// Borrowed map descriptor.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct XmlMapRef<'a> {
    pub id: u32,
    pub name: &'a str,
    pub root_element: &'a str,
    pub schema_id: &'a str,
    pub show_import_export_validation_errors: bool,
    pub auto_fit: bool,
    pub append: bool,
    pub preserve_sort_auto_filter_layout: bool,
    pub preserve_format: bool,
    pub data_binding: Option<DataBindingRef<'a>>,
}

impl<'a> From<&'a XmlMap> for XmlMapRef<'a> {
    fn from(value: &'a XmlMap) -> Self {
        Self {
            id: value.id,
            name: &value.name,
            root_element: &value.root_element,
            schema_id: &value.schema_id,
            show_import_export_validation_errors: value.show_import_export_validation_errors,
            auto_fit: value.auto_fit,
            append: value.append,
            preserve_sort_auto_filter_layout: value.preserve_sort_auto_filter_layout,
            preserve_format: value.preserve_format,
            data_binding: value.data_binding.as_ref().map(DataBindingRef::from),
        }
    }
}

/// Borrowed MapInfo projection with small owned descriptor vectors only.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct XmlMapInfoRef<'a> {
    pub selection_namespaces: &'a str,
    pub schemas: Vec<XmlSchemaRef<'a>>,
    pub maps: Vec<XmlMapRef<'a>>,
}

impl<'a> XmlMapInfoRef<'a> {
    /// Project an owned common model using caller-selected descriptor bounds.
    pub fn from_owned_with_limits(
        value: &'a XmlMapInfo,
        limits: &XmlMapLimits,
    ) -> crate::Result<Self> {
        if value.schemas.len() > limits.max_schemas || value.maps.len() > limits.max_maps {
            return Err(crate::Error::SpreadsheetXmlMaps(
                "custom XML maps descriptor count exceeds configured limit".into(),
            ));
        }
        let mut schemas = Vec::new();
        schemas.try_reserve(value.schemas.len()).map_err(|_| {
            crate::Error::SpreadsheetXmlMaps(
                "custom XML maps schema descriptor allocation failed".into(),
            )
        })?;
        schemas.extend(value.schemas.iter().map(XmlSchemaRef::from));
        let mut maps = Vec::new();
        maps.try_reserve(value.maps.len()).map_err(|_| {
            crate::Error::SpreadsheetXmlMaps(
                "custom XML maps map descriptor allocation failed".into(),
            )
        })?;
        maps.extend(value.maps.iter().map(XmlMapRef::from));
        Ok(Self {
            selection_namespaces: &value.selection_namespaces,
            schemas,
            maps,
        })
    }
}

impl<'a> TryFrom<&'a XmlMapInfo> for XmlMapInfoRef<'a> {
    type Error = crate::Error;

    fn try_from(value: &'a XmlMapInfo) -> Result<Self, Self::Error> {
        Self::from_owned_with_limits(value, &XmlMapLimits::DEFAULT)
    }
}
