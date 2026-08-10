//! XLSX-owned compatibility model for `SpreadsheetML` Custom XML Maps.

use litchi_core::sheet::Result;

use super::invalid;

#[cfg(test)]
pub(super) const NS: &[u8] = litchi_ooxml_common::spreadsheet_xml_maps::NS;
#[cfg(test)]
pub(super) const STRICT_NS: &[u8] = litchi_ooxml_common::spreadsheet_xml_maps::STRICT_NS;
pub(super) const REL: &str = litchi_ooxml_common::spreadsheet_xml_maps::REL;
pub(super) const STRICT_REL: &str = litchi_ooxml_common::spreadsheet_xml_maps::STRICT_REL;
pub(super) const CONTENT_TYPE: &str = litchi_ooxml_common::spreadsheet_xml_maps::CONTENT_TYPE;
pub(super) const MAX_PART_BYTES: usize = litchi_ooxml_common::spreadsheet_xml_maps::MAX_PART_BYTES;
#[cfg(test)]
pub(super) const MAX_OPAQUE_BYTES: usize =
    litchi_ooxml_common::spreadsheet_xml_maps::MAX_OPAQUE_BYTES;

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
    #[must_use]
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

/// Short compatibility name for [`XmlMapSchema`].
pub type XmlSchema = XmlMapSchema;
/// Short compatibility name for [`XmlMapDataBinding`].
pub type DataBinding = XmlMapDataBinding;
pub use litchi_ooxml_common::spreadsheet_xml_maps::XmlMapLimits;

impl From<XmlMapConformance> for litchi_ooxml_common::spreadsheet_xml_maps::XmlMapConformance {
    fn from(value: XmlMapConformance) -> Self {
        match value {
            XmlMapConformance::Transitional => Self::Transitional,
            XmlMapConformance::Strict => Self::Strict,
        }
    }
}

impl From<litchi_ooxml_common::spreadsheet_xml_maps::XmlMapConformance> for XmlMapConformance {
    fn from(value: litchi_ooxml_common::spreadsheet_xml_maps::XmlMapConformance) -> Self {
        match value {
            litchi_ooxml_common::spreadsheet_xml_maps::XmlMapConformance::Transitional => {
                Self::Transitional
            },
            litchi_ooxml_common::spreadsheet_xml_maps::XmlMapConformance::Strict => Self::Strict,
        }
    }
}

impl From<XmlMapSchema> for litchi_ooxml_common::spreadsheet_xml_maps::XmlSchema {
    fn from(value: XmlMapSchema) -> Self {
        Self {
            id: value.id,
            schema_reference: value.schema_reference,
            namespace: value.namespace,
            payload_xml: value.payload_xml,
        }
    }
}

impl From<litchi_ooxml_common::spreadsheet_xml_maps::XmlSchema> for XmlMapSchema {
    fn from(value: litchi_ooxml_common::spreadsheet_xml_maps::XmlSchema) -> Self {
        Self {
            id: value.id,
            schema_reference: value.schema_reference,
            namespace: value.namespace,
            payload_xml: value.payload_xml,
        }
    }
}

impl From<XmlMapDataBinding> for litchi_ooxml_common::spreadsheet_xml_maps::DataBinding {
    fn from(value: XmlMapDataBinding) -> Self {
        Self {
            data_binding_name: value.data_binding_name,
            file_binding: value.file_binding,
            connection_id: value.connection_id,
            file_binding_name: value.file_binding_name,
            load_mode: value.load_mode,
            payload_xml: value.payload_xml,
        }
    }
}

impl From<litchi_ooxml_common::spreadsheet_xml_maps::DataBinding> for XmlMapDataBinding {
    fn from(value: litchi_ooxml_common::spreadsheet_xml_maps::DataBinding) -> Self {
        Self {
            data_binding_name: value.data_binding_name,
            file_binding: value.file_binding,
            connection_id: value.connection_id,
            file_binding_name: value.file_binding_name,
            load_mode: value.load_mode,
            payload_xml: value.payload_xml,
        }
    }
}

impl From<XmlMap> for litchi_ooxml_common::spreadsheet_xml_maps::XmlMap {
    fn from(value: XmlMap) -> Self {
        Self {
            id: value.id,
            name: value.name,
            root_element: value.root_element,
            schema_id: value.schema_id,
            show_import_export_validation_errors: value.show_import_export_validation_errors,
            auto_fit: value.auto_fit,
            append: value.append,
            preserve_sort_auto_filter_layout: value.preserve_sort_auto_filter_layout,
            preserve_format: value.preserve_format,
            data_binding: value.data_binding.map(Into::into),
        }
    }
}

impl From<litchi_ooxml_common::spreadsheet_xml_maps::XmlMap> for XmlMap {
    fn from(value: litchi_ooxml_common::spreadsheet_xml_maps::XmlMap) -> Self {
        Self {
            id: value.id,
            name: value.name,
            root_element: value.root_element,
            schema_id: value.schema_id,
            show_import_export_validation_errors: value.show_import_export_validation_errors,
            auto_fit: value.auto_fit,
            append: value.append,
            preserve_sort_auto_filter_layout: value.preserve_sort_auto_filter_layout,
            preserve_format: value.preserve_format,
            data_binding: value.data_binding.map(Into::into),
        }
    }
}

impl From<XmlMapInfo> for litchi_ooxml_common::spreadsheet_xml_maps::XmlMapInfo {
    fn from(value: XmlMapInfo) -> Self {
        Self {
            selection_namespaces: value.selection_namespaces,
            schemas: value.schemas.into_iter().map(Into::into).collect(),
            maps: value.maps.into_iter().map(Into::into).collect(),
        }
    }
}

impl From<litchi_ooxml_common::spreadsheet_xml_maps::XmlMapInfo> for XmlMapInfo {
    fn from(value: litchi_ooxml_common::spreadsheet_xml_maps::XmlMapInfo) -> Self {
        Self {
            selection_namespaces: value.selection_namespaces,
            schemas: value.schemas.into_iter().map(Into::into).collect(),
            maps: value.maps.into_iter().map(Into::into).collect(),
        }
    }
}

impl XmlMapInfo {
    pub(super) fn to_common_ref(
        &self,
    ) -> Result<litchi_ooxml_common::spreadsheet_xml_maps::XmlMapInfoRef<'_>> {
        use litchi_ooxml_common::spreadsheet_xml_maps::{
            DataBindingRef, XmlMapInfoRef, XmlMapLimits, XmlMapRef, XmlSchemaRef,
        };

        let limits = XmlMapLimits::DEFAULT;
        if self.schemas.len() > limits.max_schemas {
            return Err(invalid("custom XML maps schema limit exceeded"));
        }
        if self.maps.len() > limits.max_maps {
            return Err(invalid("custom XML maps map limit exceeded"));
        }
        let mut schemas = Vec::new();
        schemas
            .try_reserve_exact(self.schemas.len())
            .map_err(|_source| invalid("custom XML maps schema descriptor allocation failed"))?;
        for schema in &self.schemas {
            schemas.push(XmlSchemaRef {
                id: &schema.id,
                schema_reference: schema.schema_reference.as_deref(),
                namespace: schema.namespace.as_deref(),
                payload_xml: schema.payload_xml.as_deref(),
            });
        }
        let mut maps = Vec::new();
        maps.try_reserve_exact(self.maps.len())
            .map_err(|_source| invalid("custom XML maps map descriptor allocation failed"))?;
        for map in &self.maps {
            maps.push(XmlMapRef {
                id: map.id,
                name: &map.name,
                root_element: &map.root_element,
                schema_id: &map.schema_id,
                show_import_export_validation_errors: map.show_import_export_validation_errors,
                auto_fit: map.auto_fit,
                append: map.append,
                preserve_sort_auto_filter_layout: map.preserve_sort_auto_filter_layout,
                preserve_format: map.preserve_format,
                data_binding: map.data_binding.as_ref().map(|binding| DataBindingRef {
                    data_binding_name: binding.data_binding_name.as_deref(),
                    file_binding: binding.file_binding,
                    connection_id: binding.connection_id,
                    file_binding_name: binding.file_binding_name.as_deref(),
                    load_mode: binding.load_mode,
                    payload_xml: binding.payload_xml.as_deref(),
                }),
            });
        }
        Ok(XmlMapInfoRef {
            selection_namespaces: &self.selection_namespaces,
            schemas,
            maps,
        })
    }
}
