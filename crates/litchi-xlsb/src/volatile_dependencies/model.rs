//! Typed, inert models for the XLSB Volatile Dependencies part.
//!
//! The model mirrors the hierarchy described by [MS-XLSB] section 2.2.13.
//! It contains cached values and cell references only; it never contacts an
//! RTD server or OLAP connection and never evaluates a formula.

use std::num::FpCategory;

use crate::package::error::{Error, Result};

/// The largest individual Workbook or Volatile Dependencies stream accepted
/// by this owner.
pub const HARD_MAX_PART_BYTES: usize = 64 * 1024 * 1024;
/// The largest record count accepted in one individual Workbook or Volatile
/// Dependencies stream by this owner.
pub const HARD_MAX_RECORDS: usize = 1_000_000;
/// The largest number of mains, topics, subtopics, or references accepted by
/// one caller operation.
pub const HARD_MAX_ITEMS: usize = 1_000_000;
/// The largest UTF-16 string accepted by one record.
pub const HARD_MAX_STRING_UNITS: usize = 1_048_576;
/// The largest aggregate UTF-16 string budget accepted by one operation.
pub const HARD_MAX_TOTAL_STRING_UNITS: usize = 64 * 1024 * 1024;
/// The largest relationship count captured in one owner source closure.
pub const HARD_MAX_RELATIONSHIPS: usize = 1_000_000;

/// Caller-controlled finite limits for Volatile Dependencies inspection and
/// authoring. `max_part_bytes` and `max_records` are per-stream bounds: the
/// the same values apply to the Workbook stream used for sheet-ordinal
/// binding and to the Volatile Dependencies owner stream. The byte bound also
/// caps the relationship and content-types source tokens captured for the
/// ownership check. Values are checked against hard safety ceilings before
/// parsing or reserving semantic collections.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct ReadLimits {
    /// Maximum encoded bytes in each inspected Workbook or owner stream and
    /// each captured relationship/content-types source token.
    pub max_part_bytes: usize,
    /// Maximum BIFF12 records in each inspected Workbook or owner part.
    pub max_records: usize,
    /// Maximum type collections. The format has at most RTD and cube types.
    pub max_types: usize,
    /// Maximum `BrtBeginVolMain` collections.
    pub max_mains: usize,
    /// Maximum `BrtBeginVolTopic` collections.
    pub max_topics: usize,
    /// Maximum `BrtVolSubtopic` records.
    pub max_subtopics: usize,
    /// Maximum `BrtVolRef` records.
    pub max_references: usize,
    /// Maximum UTF-16 units in one string record.
    pub max_string_units: usize,
    /// Maximum UTF-16 units across all string records.
    pub max_total_string_units: usize,
    /// Maximum root plus Workbook relationships captured for source checks.
    pub max_relationships: usize,
}

impl ReadLimits {
    /// Conservative finite defaults for ordinary workbook inspection.
    pub const DEFAULT: Self = Self {
        max_part_bytes: 16 * 1024 * 1024,
        max_records: 262_144,
        max_types: 2,
        max_mains: 65_535,
        max_topics: 262_144,
        max_subtopics: 1_000_000,
        max_references: 1_000_000,
        max_string_units: 32_767,
        max_total_string_units: 8 * 1024 * 1024,
        max_relationships: 65_536,
    };

    /// Validate caller limits against physical and semantic hard ceilings.
    pub const fn validate(self) -> Result<Self> {
        if self.max_part_bytes > HARD_MAX_PART_BYTES {
            return Err(limit(
                "Volatile Dependencies part bytes",
                self.max_part_bytes,
                HARD_MAX_PART_BYTES,
            ));
        }
        if self.max_records > HARD_MAX_RECORDS {
            return Err(limit(
                "Volatile Dependencies records",
                self.max_records,
                HARD_MAX_RECORDS,
            ));
        }
        if self.max_types > 2 {
            return Err(limit(
                "Volatile Dependencies type collections",
                self.max_types,
                2,
            ));
        }
        if self.max_mains > HARD_MAX_ITEMS {
            return Err(limit(
                "Volatile main collections",
                self.max_mains,
                HARD_MAX_ITEMS,
            ));
        }
        if self.max_topics > HARD_MAX_ITEMS {
            return Err(limit(
                "Volatile topic collections",
                self.max_topics,
                HARD_MAX_ITEMS,
            ));
        }
        if self.max_subtopics > HARD_MAX_ITEMS {
            return Err(limit(
                "Volatile subtopic records",
                self.max_subtopics,
                HARD_MAX_ITEMS,
            ));
        }
        if self.max_references > HARD_MAX_ITEMS {
            return Err(limit(
                "Volatile cell references",
                self.max_references,
                HARD_MAX_ITEMS,
            ));
        }
        if self.max_string_units > HARD_MAX_STRING_UNITS {
            return Err(limit(
                "Volatile UTF-16 string",
                self.max_string_units,
                HARD_MAX_STRING_UNITS,
            ));
        }
        if self.max_total_string_units > HARD_MAX_TOTAL_STRING_UNITS {
            return Err(limit(
                "Volatile UTF-16 strings",
                self.max_total_string_units,
                HARD_MAX_TOTAL_STRING_UNITS,
            ));
        }
        if self.max_relationships > HARD_MAX_RELATIONSHIPS {
            return Err(limit(
                "Volatile Dependencies relationships",
                self.max_relationships,
                HARD_MAX_RELATIONSHIPS,
            ));
        }
        Ok(self)
    }
}

impl Default for ReadLimits {
    fn default() -> Self {
        Self::DEFAULT
    }
}

/// Volatile dependency type from `BrtBeginVolType.type`.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum DependencyKind {
    /// Real-time data server dependency.
    Rtd = 0,
    /// Cube-function dependency.
    Cube = 1,
}

impl DependencyKind {
    pub(crate) fn from_wire(value: u32) -> Result<Self> {
        match value {
            0 => Ok(Self::Rtd),
            1 => Ok(Self::Cube),
            value => Err(invalid(format!("unknown volatile dependency type {value}"))),
        }
    }

    pub(crate) const fn wire(self) -> u32 {
        self as u32
    }
}

/// A bounded cell coordinate from `BrtVolRef`.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub struct CellReference {
    row: u32,
    column: u32,
    sheet_index: u32,
}

impl CellReference {
    /// Construct a checked zero-based cell reference.
    pub fn new(row: u32, column: u32, sheet_index: u32) -> Result<Self> {
        if row >= 1_048_576 {
            return Err(invalid(format!("volatile row {row} is outside 0..1048575")));
        }
        if column >= 16_384 {
            return Err(invalid(format!(
                "volatile column {column} is outside 0..16383"
            )));
        }
        Ok(Self {
            row,
            column,
            sheet_index,
        })
    }

    /// Zero-based row index.
    #[must_use]
    pub const fn row(self) -> u32 {
        self.row
    }

    /// Zero-based column index.
    #[must_use]
    pub const fn column(self) -> u32 {
        self.column
    }

    /// Zero-based index into the Workbook `BrtBundleSh` collection.
    #[must_use]
    pub const fn sheet_index(self) -> u32 {
        self.sheet_index
    }
}

/// Closed `[MS-XLSB]` error value carried by `BrtVolErr`.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum ErrorCode {
    /// `#NULL!`.
    Null = 0x00,
    /// `#DIV/0!`.
    Div0 = 0x07,
    /// `#VALUE!`.
    Value = 0x0F,
    /// `#REF!`.
    Ref = 0x17,
    /// `#NAME?`.
    Name = 0x1D,
    /// `#NUM!`.
    Num = 0x24,
    /// `#N/A`.
    Na = 0x2A,
    /// `#GETTING_DATA`.
    GettingData = 0x2B,
}

impl ErrorCode {
    pub(crate) fn from_wire(value: u8) -> Result<Self> {
        match value {
            0x00 => Ok(Self::Null),
            0x07 => Ok(Self::Div0),
            0x0F => Ok(Self::Value),
            0x17 => Ok(Self::Ref),
            0x1D => Ok(Self::Name),
            0x24 => Ok(Self::Num),
            0x2A => Ok(Self::Na),
            0x2B => Ok(Self::GettingData),
            value => Err(invalid(format!(
                "invalid volatile BErr value 0x{value:02X}"
            ))),
        }
    }

    pub(crate) const fn wire(self) -> u8 {
        self as u8
    }
}

/// Cached scalar returned by one volatile dependency topic.
#[derive(Clone, Debug, PartialEq)]
pub enum CachedValue {
    /// Cached `Xnum` value.
    Number(f64),
    /// Cached `BErr` value.
    Error(ErrorCode),
    /// Cached `XLWideString` value.
    String(String),
    /// Cached one-byte Boolean value.
    Bool(bool),
}

/// One topic group sharing one complete set of function parameters.
#[derive(Clone, Debug, PartialEq)]
pub struct Topic {
    /// Additional parameter values in source order.
    pub subtopics: Vec<String>,
    /// Cached return value, exactly one per topic.
    pub value: CachedValue,
    /// Cells that depend on this topic.
    pub references: Vec<CellReference>,
}

/// One main group sharing its RTD ProgID or cube connection name.
#[derive(Clone, Debug, PartialEq)]
pub struct MainTopic {
    /// RTD ProgID or OLAP connection name.
    pub first: String,
    /// Topic groups in source order.
    pub topics: Vec<Topic>,
}

/// One RTD or cube dependency type collection.
#[derive(Clone, Debug, PartialEq)]
pub struct VolatileType {
    /// Whether this collection describes RTD or cube dependencies.
    pub kind: DependencyKind,
    /// Main groups in source order.
    pub mains: Vec<MainTopic>,
}

/// Complete typed Volatile Dependencies stream.
#[derive(Clone, Debug, PartialEq)]
pub struct Dependencies {
    /// RTD and cube collections in source order.
    pub types: Vec<VolatileType>,
    /// Whether a forward or unsupported record was present in the source.
    /// Such a stream can be inspected and round-tripped unchanged, but typed
    /// replacement is refused because its complete ordering is not proven.
    pub(crate) has_unsupported_records: bool,
}

impl Dependencies {
    /// Construct a detached typed dependency collection.
    #[must_use]
    pub const fn new(types: Vec<VolatileType>) -> Self {
        Self {
            types,
            has_unsupported_records: false,
        }
    }

    /// Validate model cardinalities, scalar domains, and caller quotas.
    pub fn validate(&self, limits: ReadLimits) -> Result<()> {
        let limits = limits.validate()?;
        if self.types.len() > limits.max_types {
            return Err(limit(
                "volatile type collections",
                self.types.len(),
                limits.max_types,
            ));
        }
        let mut kinds = 0u8;
        let mut mains = 0usize;
        let mut topics = 0usize;
        let mut subtopics = 0usize;
        let mut references = 0usize;
        let mut strings = 0usize;
        for volatile_type in &self.types {
            let bit = 1u8 << volatile_type.kind.wire();
            if kinds & bit != 0 {
                return Err(invalid("duplicate volatile dependency type"));
            }
            kinds |= bit;
            if volatile_type.mains.len() > limits.max_mains.saturating_sub(mains) {
                return Err(limit(
                    "volatile main collections",
                    mains + volatile_type.mains.len(),
                    limits.max_mains,
                ));
            }
            mains = checked_add(
                mains,
                volatile_type.mains.len(),
                "volatile main collections",
            )?;
            for main in &volatile_type.mains {
                add_string(&mut strings, &main.first, limits, "volatile first string")?;
                if main.topics.len() > limits.max_topics.saturating_sub(topics) {
                    return Err(limit(
                        "volatile topic collections",
                        topics + main.topics.len(),
                        limits.max_topics,
                    ));
                }
                topics = checked_add(topics, main.topics.len(), "volatile topic collections")?;
                for topic in &main.topics {
                    if topic.subtopics.len() > limits.max_subtopics.saturating_sub(subtopics) {
                        return Err(limit(
                            "volatile subtopic records",
                            subtopics + topic.subtopics.len(),
                            limits.max_subtopics,
                        ));
                    }
                    subtopics = checked_add(
                        subtopics,
                        topic.subtopics.len(),
                        "volatile subtopic records",
                    )?;
                    for subtopic in &topic.subtopics {
                        add_string(&mut strings, subtopic, limits, "volatile subtopic")?;
                    }
                    validate_cached_value(&topic.value, &mut strings, limits)?;
                    if topic.references.len() > limits.max_references.saturating_sub(references) {
                        return Err(limit(
                            "volatile cell references",
                            references + topic.references.len(),
                            limits.max_references,
                        ));
                    }
                    references = checked_add(
                        references,
                        topic.references.len(),
                        "volatile cell references",
                    )?;
                    for reference in &topic.references {
                        let _ = CellReference::new(
                            reference.row,
                            reference.column,
                            reference.sheet_index,
                        )?;
                    }
                }
            }
        }
        Ok(())
    }

    /// Validate every `BrtVolRef.ish` against the Workbook `BrtBundleSh`
    /// collection. The wire field is an ordinal into that collection, rather
    /// than a relationship identifier or an `iTabID` value.
    pub fn validate_sheet_references(&self, sheet_count: usize) -> Result<()> {
        for volatile_type in &self.types {
            for main in &volatile_type.mains {
                for topic in &main.topics {
                    for reference in &topic.references {
                        let sheet_index =
                            usize::try_from(reference.sheet_index).map_err(|_error| {
                                Error::CapacityOverflow {
                                    resource: "volatile sheet ordinal",
                                }
                            })?;
                        if sheet_index >= sheet_count {
                            return Err(invalid(format!(
                                "volatile sheet ordinal {} exceeds {} Workbook BrtBundleSh records",
                                reference.sheet_index, sheet_count
                            )));
                        }
                    }
                }
            }
        }
        Ok(())
    }

    /// Whether the source contained records this typed writer cannot safely
    /// place after a semantic edit.
    #[must_use]
    pub const fn has_unsupported_records(&self) -> bool {
        self.has_unsupported_records
    }
}

fn validate_cached_value(
    value: &CachedValue,
    strings: &mut usize,
    limits: ReadLimits,
) -> Result<()> {
    match value {
        CachedValue::Number(value) => validate_xnum(*value),
        CachedValue::Error(_) | CachedValue::Bool(_) => Ok(()),
        CachedValue::String(value) => add_string(strings, value, limits, "volatile cached string"),
    }
}

pub(crate) fn validate_xnum(value: f64) -> Result<()> {
    if matches!(
        value.classify(),
        FpCategory::Nan | FpCategory::Infinite | FpCategory::Subnormal
    ) || (value == 0.0 && value.is_sign_negative())
    {
        return Err(invalid(format!(
            "invalid volatile Xnum bit pattern 0x{:016X}",
            value.to_bits()
        )));
    }
    Ok(())
}

fn add_string(
    value: &mut usize,
    string: &str,
    limits: ReadLimits,
    what: &'static str,
) -> Result<()> {
    let units = string.encode_utf16().count();
    if units > limits.max_string_units {
        return Err(limit(what, units, limits.max_string_units));
    }
    *value = checked_add(*value, units, "volatile UTF-16 strings")?;
    if *value > limits.max_total_string_units {
        return Err(limit(
            "volatile UTF-16 strings",
            *value,
            limits.max_total_string_units,
        ));
    }
    Ok(())
}

fn checked_add(left: usize, right: usize, resource: &'static str) -> Result<usize> {
    left.checked_add(right)
        .ok_or(Error::CapacityOverflow { resource })
}

const fn limit(resource: &'static str, actual: usize, maximum: usize) -> Error {
    Error::LimitExceeded {
        resource,
        actual,
        maximum,
    }
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}
