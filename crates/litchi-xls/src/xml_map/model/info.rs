//! XML-map container model.

use super::identity::validate_string;
use super::map::Map;
use super::schema::{NamespaceDeclaration, Schema};
use crate::Result;

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MapInfo {
    pub(super) selection_namespaces: String,
    pub(super) namespaces: Vec<NamespaceDeclaration>,
    pub(super) schemas: Vec<Schema>,
    pub(super) maps: Vec<Map>,
}

impl MapInfo {
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn try_new(
        selection_namespaces: impl Into<String>,
        schemas: Vec<Schema>,
        maps: Vec<Map>,
    ) -> Result<Self> {
        let value = Self::from_parts(selection_namespaces.into(), Vec::new(), schemas, maps)?;
        super::super::validation::validate(&value)?;
        Ok(value)
    }

    pub(crate) fn from_parts(
        selection_namespaces: String,
        namespaces: Vec<(String, String)>,
        schemas: Vec<Schema>,
        maps: Vec<Map>,
    ) -> Result<Self> {
        let selection_namespaces =
            validate_string(selection_namespaces, 65_535, "SelectionNamespaces", true)?;
        let namespaces = namespaces
            .into_iter()
            .map(|(prefix, uri)| NamespaceDeclaration::try_new(prefix, uri))
            .collect::<Result<Vec<_>>>()?;
        Ok(Self {
            selection_namespaces,
            namespaces,
            schemas,
            maps,
        })
    }

    #[must_use]
    pub fn selection_namespaces(&self) -> &str {
        &self.selection_namespaces
    }
    #[must_use]
    pub fn namespaces(&self) -> &[NamespaceDeclaration] {
        &self.namespaces
    }
    #[must_use]
    pub fn schemas(&self) -> &[Schema] {
        &self.schemas
    }
    #[must_use]
    pub fn maps(&self) -> &[Map] {
        &self.maps
    }
    #[must_use]
    pub fn schema(&self, id: &super::identity::SchemaId) -> Option<&Schema> {
        self.schemas.iter().find(|schema| schema.id() == id)
    }
    #[must_use]
    pub fn map(&self, id: super::identity::MapId) -> Option<&Map> {
        self.maps.iter().find(|map| map.id() == id)
    }
}
