//! Typed workbook-side Data Model records and an inert model-part handle.

use std::sync::Arc;

/// One table declared by a workbook Data Model record family.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Table {
    /// Stable Data Model table identifier.
    pub id: String,
    /// User-visible table name used by relationships.
    pub name: String,
    /// Name of the workbook external connection owning the table.
    pub connection: String,
}

/// One relationship between two Data Model table columns.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Relationship {
    /// Foreign-key table name.
    pub from_table: String,
    /// Foreign-key column name.
    pub from_column: String,
    /// Primary-key table name.
    pub to_table: String,
    /// Primary-key column name.
    pub to_column: String,
}

/// The nine time-grouping granularities defined by MS-XLSB.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum TimeGroupingContentType {
    /// Calendar years.
    Years = 0,
    /// Calendar quarters.
    Quarters = 1,
    /// Month ordinal values.
    MonthsIndex = 2,
    /// Calendar month names.
    Months = 3,
    /// Day ordinal values.
    DaysIndex = 4,
    /// Calendar day names.
    Days = 5,
    /// Hours.
    Hours = 6,
    /// Minutes.
    Minutes = 7,
    /// Seconds.
    Seconds = 8,
}

impl TimeGroupingContentType {
    pub(crate) const fn from_wire(value: u8) -> Option<Self> {
        match value {
            0 => Some(Self::Years),
            1 => Some(Self::Quarters),
            2 => Some(Self::MonthsIndex),
            3 => Some(Self::Months),
            4 => Some(Self::DaysIndex),
            5 => Some(Self::Days),
            6 => Some(Self::Hours),
            7 => Some(Self::Minutes),
            8 => Some(Self::Seconds),
            _ => None,
        }
    }

    pub(crate) const fn wire(self) -> u8 {
        self as u8
    }
}

/// One calculated column emitted below a Data Model time grouping.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct TimeGroupingColumn {
    /// Whether this granularity was selected by the last time-grouping action.
    pub is_selected: bool,
    /// Calendar granularity represented by this generated column.
    pub content_type: TimeGroupingContentType,
    /// Generated column name.
    pub column_name: String,
    /// Immutable generated column identifier.
    pub column_id: String,
}

/// One Data Model time grouping and its generated calculated columns.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct TimeGrouping {
    /// Source Data Model table name.
    pub table_name: String,
    /// Source Data Model column name.
    pub column_name: String,
    /// Source Data Model column immutable identifier.
    pub column_id: String,
    /// Calculated columns generated for this grouping.
    pub columns: Vec<TimeGroupingColumn>,
}

/// Typed metadata carried by the XLSB workbook stream.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Definition {
    /// Minimum application version required to load the model.
    pub min_version_load: u8,
    /// Data Model table declarations.
    pub tables: Vec<Table>,
    /// Data Model relationships.
    pub relationships: Vec<Relationship>,
    /// Calendar time groupings.
    pub time_groupings: Vec<TimeGrouping>,
}

impl Default for Definition {
    fn default() -> Self {
        Self {
            min_version_load: 5,
            tables: Vec::new(),
            relationships: Vec::new(),
            time_groupings: Vec::new(),
        }
    }
}

/// Opaque XLSB model-part bytes with validated OPC ownership metadata.
///
/// The payload is deliberately not interpreted as DAX, SQL, XML, or an
/// executable artifact. Cloning this value shares its allocation; bytes are
/// only copied when [`Self::to_vec`] is requested.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ModelPart {
    pub(crate) part_name: String,
    pub(crate) content_type: String,
    pub(crate) bytes: Arc<Vec<u8>>,
}

impl ModelPart {
    /// Build a model part using the fixed XLSB model-part envelope.
    pub fn from_bytes(data: Vec<u8>) -> crate::package::error::Result<Self> {
        let part = Self {
            part_name: super::package::DATA_MODEL_PART_NAME.to_string(),
            content_type: super::package::DATA_MODEL_CONTENT_TYPE.to_string(),
            bytes: Arc::new(data),
        };
        super::package::validate_payload(&part)?;
        Ok(part)
    }

    /// Absolute OPC part name. The XLSB contract currently requires
    /// `/xl/model/item.data`.
    #[must_use]
    pub fn part_name(&self) -> &str {
        &self.part_name
    }

    /// OPC content type advertised for this part.
    #[must_use]
    pub fn content_type(&self) -> &str {
        &self.content_type
    }

    /// Borrow the inert MS-XLDM payload without parsing it.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        self.bytes.as_slice()
    }

    /// Return the payload length without materializing a second allocation.
    #[must_use]
    pub fn len(&self) -> usize {
        self.bytes.len()
    }

    /// Whether the payload is empty. Valid model parts are never empty.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.bytes.is_empty()
    }

    /// Copy the inert payload for an explicit replacement or export.
    #[must_use]
    pub fn to_vec(&self) -> Vec<u8> {
        self.bytes.as_ref().clone()
    }
}

/// A complete model view combining workbook records with the model part.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Model {
    /// Typed metadata from the workbook binary stream.
    pub definition: Definition,
    /// Relationship-free opaque MS-XLDM payload part.
    pub part: ModelPart,
}

impl Model {
    /// Construct a model from typed records and an inert payload.
    #[must_use]
    pub const fn new(definition: Definition, part: ModelPart) -> Self {
        Self { definition, part }
    }

    /// Construct a model from typed workbook metadata and opaque payload bytes.
    pub fn from_bytes(
        definition: Definition,
        data: Vec<u8>,
    ) -> crate::package::error::Result<Self> {
        Ok(Self::new(definition, ModelPart::from_bytes(data)?))
    }
}
