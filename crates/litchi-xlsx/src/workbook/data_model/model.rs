//! Semantic MS-XLDM descriptor and inert payload model.

use std::sync::Arc;

/// Retained `x15:extLst` descriptor markup that is not interpreted.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OpaqueXml {
    /// A self-contained `x15:extLst` subtree. Readback retains source markup,
    /// closes inherited namespace bindings, and materializes inherited
    /// `xml:base`, `xml:lang`, and `xml:space` context at its root when needed.
    /// Relative root bases are resolved against the source owner; nested base
    /// overrides remain lexical. Prefixes, comments, mixed content, and opaque
    /// attribute values are otherwise not rewritten.
    pub xml: Vec<u8>,
}

/// One workbook Data Model table descriptor.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Table {
    pub id: String,
    pub name: String,
    pub connection: String,
}

/// One Data Model relationship between table columns.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Relationship {
    pub from_table: String,
    pub from_column: String,
    pub to_table: String,
    pub to_column: String,
}

/// Typed inline `x15:dataModel` workbook descriptor.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Definition {
    /// Minimum application version. The MS-XLSX default and floor are `5`.
    pub min_version_load: u8,
    pub tables: Vec<Table>,
    pub relationships: Vec<Relationship>,
    pub extension_list: Option<OpaqueXml>,
}

impl Default for Definition {
    fn default() -> Self {
        Self {
            min_version_load: 5,
            tables: Vec::new(),
            relationships: Vec::new(),
            extension_list: None,
        }
    }
}

/// Inert MS-XLDM storage payload attached to the workbook Data Model.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Payload {
    /// Absolute OPC part name; currently `/xl/model/item.data`.
    pub part_name: String,
    /// Opaque MS-XLDM bytes. No inner-file or credential processing occurs.
    pub data: Vec<u8>,
}

/// Complete typed descriptor plus inert MS-XLDM payload.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Model {
    pub definition: Definition,
    pub payload: Payload,
}

/// An immutable Data Model view sharing the package's opaque payload bytes.
///
/// Cloning a view shares both its descriptor and payload. Use [`Self::to_owned`]
/// when an independently owned authoring value is explicitly needed.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ModelView {
    pub(crate) definition: Arc<Definition>,
    pub(crate) data: Arc<Vec<u8>>,
}

impl ModelView {
    /// The typed workbook descriptor; no model expressions are evaluated.
    #[must_use]
    pub fn definition(&self) -> &Definition {
        &self.definition
    }

    /// Borrow the exact MS-XLDM payload without allocating or copying.
    #[must_use]
    pub fn payload(&self) -> &[u8] {
        self.data.as_slice()
    }

    /// The fixed package path required by the XLSX Data Model specification.
    #[must_use]
    pub const fn part_name(&self) -> &'static str {
        super::DATA_MODEL_PART_NAME
    }

    /// Copy this view into the independently owned low-level authoring model.
    /// This explicitly copies the complete opaque payload.
    #[must_use]
    pub fn to_owned(&self) -> Model {
        Model {
            definition: self.definition.as_ref().clone(),
            payload: Payload {
                part_name: self.part_name().into(),
                data: self.data.as_ref().clone(),
            },
        }
    }
}
