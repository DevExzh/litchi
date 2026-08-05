//! Semantic MS-XLDM descriptor and inert payload model.

/// Retained `x15:extLst` descriptor markup that is not interpreted.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OpaqueXml {
    /// A self-contained `x15:extLst` subtree.
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
