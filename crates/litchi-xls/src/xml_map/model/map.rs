//! XML-map behavior and binding ownership.

use super::binding::DataBinding;
use super::identity::{MapId, SchemaId, validate_string};
use super::schema::NamespaceDeclaration;
use crate::{Error, Result};

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Map {
    pub(super) id: MapId,
    pub(super) name: String,
    pub(super) root_element: String,
    pub(super) schema_id: SchemaId,
    pub(super) show_import_export_validation_errors: bool,
    pub(super) auto_fit: bool,
    pub(super) append: bool,
    pub(super) preserve_sort_auto_filter_layout: bool,
    pub(super) preserve_format: bool,
    pub(super) data_binding: Option<DataBinding>,
    pub(super) namespaces: Vec<NamespaceDeclaration>,
}

impl Map {
    #[allow(clippy::too_many_arguments)]
    pub fn try_new(
        id: MapId,
        name: impl Into<String>,
        root_element: impl Into<String>,
        schema_id: SchemaId,
        show_import_export_validation_errors: bool,
        auto_fit: bool,
        append: bool,
        preserve_sort_auto_filter_layout: bool,
        preserve_format: bool,
    ) -> Result<Self> {
        Self::from_parts(
            id.get().to_string(),
            name.into(),
            root_element.into(),
            schema_id.as_str().to_string(),
            show_import_export_validation_errors,
            auto_fit,
            append,
            preserve_sort_auto_filter_layout,
            preserve_format,
            None,
            Vec::new(),
        )
    }

    #[allow(clippy::too_many_arguments)]
    pub(crate) fn from_parts(
        id: String,
        name: String,
        root_element: String,
        schema_id: String,
        show_import_export_validation_errors: bool,
        auto_fit: bool,
        append: bool,
        preserve_sort_auto_filter_layout: bool,
        preserve_format: bool,
        data_binding: Option<DataBinding>,
        namespaces: Vec<(String, String)>,
    ) -> Result<Self> {
        let id = MapId::new(
            id.parse::<u32>()
                .map_err(|_| invalid("Map ID is not an unsigned decimal value"))?,
        )?;
        let name = validate_string(name, 256, "Map Name", false)?;
        let root_element = validate_string(root_element, 65_535, "Map RootElement", false)?;
        let schema_id = SchemaId::new(schema_id)?;
        let namespaces = namespaces
            .into_iter()
            .map(|(prefix, uri)| NamespaceDeclaration::try_new(prefix, uri))
            .collect::<Result<Vec<_>>>()?;
        Ok(Self {
            id,
            name,
            root_element,
            schema_id,
            show_import_export_validation_errors,
            auto_fit,
            append,
            preserve_sort_auto_filter_layout,
            preserve_format,
            data_binding,
            namespaces,
        })
    }

    pub const fn id(&self) -> MapId {
        self.id
    }
    pub fn name(&self) -> &str {
        &self.name
    }
    pub fn root_element(&self) -> &str {
        &self.root_element
    }
    pub fn schema_id(&self) -> &SchemaId {
        &self.schema_id
    }
    pub const fn show_import_export_validation_errors(&self) -> bool {
        self.show_import_export_validation_errors
    }
    pub const fn auto_fit(&self) -> bool {
        self.auto_fit
    }
    pub const fn append(&self) -> bool {
        self.append
    }
    pub const fn preserve_sort_auto_filter_layout(&self) -> bool {
        self.preserve_sort_auto_filter_layout
    }
    pub const fn preserve_format(&self) -> bool {
        self.preserve_format
    }
    pub fn data_binding(&self) -> Option<&DataBinding> {
        self.data_binding.as_ref()
    }
    pub fn namespaces(&self) -> &[NamespaceDeclaration] {
        &self.namespaces
    }

    pub fn with_data_binding(mut self, value: DataBinding) -> Self {
        self.data_binding = Some(value);
        self
    }
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidData(format!("XML map: {}", message.into()))
}
