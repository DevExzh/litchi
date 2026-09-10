//! Semantic MS-XLDM descriptor and inert payload model.

use std::sync::Arc;

use crate::error::Result;

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

/// The content type of a calculated column in a data-model time grouping.
///
/// The [`Other`](Self::Other) variant retains a value introduced by a newer
/// producer without assigning it semantics this crate does not know. Such a
/// value is readable and preserved, but cannot be emitted by the typed writer.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ModelTimeGroupingContentType {
    Years,
    Quarters,
    MonthsIndex,
    Months,
    DaysIndex,
    Days,
    Hours,
    Minutes,
    Seconds,
    Other(String),
}

/// One calculated column belonging to a data-model time grouping.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CalculatedTimeColumn {
    pub column_name: String,
    pub column_id: String,
    pub content_type: ModelTimeGroupingContentType,
    pub is_selected: bool,
}

/// One data-model time grouping for a table and source date column.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ModelTimeGrouping {
    pub table_name: String,
    pub column_name: String,
    pub column_id: String,
    pub calculated_time_columns: Vec<CalculatedTimeColumn>,
}

/// The typed `modelTimeGroupings` extension owned by the Data Model
/// descriptor. Unknown sibling extensions remain in [`OpaqueXml`] verbatim.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ModelTimeGroupings {
    pub groupings: Vec<ModelTimeGrouping>,
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

impl Definition {
    /// Read the typed MS-XLSX `modelTimeGroupings` extension, if present.
    ///
    /// The surrounding `x15:extLst` and all unrecognized extensions remain
    /// source bytes; this method only projects the recognized child.
    pub fn model_time_groupings(&self) -> Result<Option<ModelTimeGroupings>> {
        self.extension_list
            .as_ref()
            .map(|extension| super::codec::parse_model_time_groupings_extension(&extension.xml))
            .transpose()
            .map(|value| value.flatten())
    }

    /// Add, replace, or remove the typed `modelTimeGroupings` child while
    /// retaining unrelated extension markup byte-for-byte. This detached
    /// descriptor helper does not prove identity against the XLDM payload;
    /// source-bound package transactions perform that proof before staging a
    /// changed write.
    pub fn set_model_time_groupings(&mut self, value: Option<ModelTimeGroupings>) -> Result<()> {
        if self.model_time_groupings()? == value {
            return Ok(());
        }
        let updated = super::codec::rewrite_model_time_groupings_extension(
            self.extension_list
                .as_ref()
                .map(|extension| extension.xml.as_slice()),
            value.as_ref(),
        )?;
        self.extension_list = updated.map(|xml| OpaqueXml { xml });
        Ok(())
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
