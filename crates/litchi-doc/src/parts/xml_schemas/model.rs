//! Typed Word 2003 XML schema reference metadata.

/// A single XML schema definition reference (`XSDR`, MS-DOC 2.9.352).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Reference {
    /// `wzURI`: the URI of the schema definition.
    pub uri: String,
    /// `wzManifestLocation`: the URI of the expansion-pack manifest the
    /// schema was loaded through, or empty when none was used.
    pub manifest_location: String,
    /// `sttbElements`: the element names of the schema, in table order.
    pub elements: Vec<String>,
    /// `sttbAttributes`: the attribute names of the schema, in table order.
    pub attributes: Vec<String>,
}

/// The XML schema definition references of a document (`Hplxsdr`, MS-DOC
/// 2.9.117), in `rgxsdr` order.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Collection {
    pub(super) schemas: Vec<Reference>,
}

impl Collection {
    pub(super) fn from_schemas(schemas: Vec<Reference>) -> Self {
        Self { schemas }
    }

    /// All schema definition references in `rgxsdr` order.
    pub fn schemas(&self) -> &[Reference] {
        &self.schemas
    }

    /// Resolve a `TIQ` name reference against the element string table of
    /// the addressed schema, or `None` when either index is out of range.
    ///
    /// Per MS-DOC 2.9.325 step 4, the `TIQ` of an `FSDAP` (a structured tag
    /// attribute) names a string in `sttbElements`.
    pub fn element_name(&self, schema_index: u32, name_index: u32) -> Option<&str> {
        self.schemas
            .get(usize::try_from(schema_index).ok()?)?
            .elements
            .get(usize::try_from(name_index).ok()?)
            .map(String::as_str)
    }

    /// Resolve a `TIQ` name reference against the attribute string table of
    /// the addressed schema, or `None` when either index is out of range.
    ///
    /// Per MS-DOC 2.9.325 step 4, the `TIQ` of an `SDTI` (a structured tag
    /// node) names a string in `sttbAttributes`.
    pub fn attribute_name(&self, schema_index: u32, name_index: u32) -> Option<&str> {
        self.schemas
            .get(usize::try_from(schema_index).ok()?)?
            .attributes
            .get(usize::try_from(name_index).ok()?)
            .map(String::as_str)
    }
}
