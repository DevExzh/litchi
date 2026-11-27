//! Typed validation errors for DOC TAP authoring.
//!
//! The semantic model and binary codec share this error vocabulary so callers
//! can distinguish invalid table state from a missing builder row without
//! inspecting encoded bytes.

use crate::parts::tap::TableWidth;

/// Error returned when table row properties cannot be represented in DOC TAP.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum TapBuildError {
    /// Requested row index is not present in the builder.
    RowOutOfBounds(usize),
    /// DOC table rows can contain at most 63 cells.
    InvalidCellCount(usize),
    /// Cumulative cell boundaries exceed the DOC XAS coordinate range.
    CellWidthsOverflow,
    /// DOC row heights use the YAS range of -31680 through 31680 twips.
    InvalidRowHeight(i16),
    /// A merge continuation cannot occur in the first cell.
    MergeWithoutPrecedingCell,
    /// Brc80 spacing is a five-bit value.
    InvalidBorderSpacing(u8),
    /// DOC cell padding cannot exceed 22 inches.
    InvalidCellPadding(u16),
    /// DOC uniform cell spacing cannot exceed 11 inches.
    InvalidCellSpacing(u16),
    /// `PGPInfo.ipgpSelf` identifiers are nonzero.
    InvalidParagraphGroupId,
    /// `PropRMark` stores its revision-author index as a signed 16-bit value.
    InvalidRevisionAuthorIndex(u16),
    /// `PropRMark` contains an invalid packed DTTM.
    InvalidRevisionTimestamp(u32),
    /// Table-style band sizes are limited to one through three cells.
    InvalidStyleBandSize(&'static str, u8),
    /// Style border defaults are only legal inside a `TCnf` property list.
    StyleBorderOutsideConditional,
    /// A conditional nested grpprl is malformed.
    InvalidConditionalProperties(String),
    /// `CNFOperand` uses a one-byte total operand length.
    ConditionalPropertiesTooLong(usize),
    /// A `TCellBrcType` prefix requires four explicit types for every included cell.
    IncompleteCellBorderTypes(usize),
    /// A preferred-width property uses unsupported units or a value outside its context's range.
    InvalidPreferredWidth(&'static str, TableWidth),
    /// TLP contains bits outside the eleven-bit Fatl field.
    InvalidTableLookFlags(u16),
    /// A physical table offset cannot be represented by the plus-one operand.
    InvalidTablePosition(&'static str, i16),
    /// A wrapping distance exceeds the `XAS/YAS_nonNeg` range.
    InvalidWrapDistance(&'static str, u16),
    /// A preserved row state cannot itself contain a `sprmTWall` boundary.
    NestedPreservedState,
}

impl std::fmt::Display for TapBuildError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::RowOutOfBounds(index) => write!(f, "table row {index} does not exist"),
            Self::InvalidCellCount(count) => {
                write!(
                    f,
                    "DOC table rows must contain between 1 and 63 cells, got {count}"
                )
            },
            Self::CellWidthsOverflow => {
                write!(
                    f,
                    "DOC cell widths exceed the 31680-twip XAS coordinate range"
                )
            },
            Self::InvalidRowHeight(height) => {
                write!(f, "DOC row height {height} is outside the YAS range")
            },
            Self::MergeWithoutPrecedingCell => {
                write!(f, "the first DOC table cell cannot be a merge continuation")
            },
            Self::InvalidBorderSpacing(spacing) => {
                write!(f, "DOC Brc80 spacing {spacing} exceeds 31 points")
            },
            Self::InvalidCellPadding(padding) => {
                write!(f, "DOC cell padding {padding} exceeds 31680 twips")
            },
            Self::InvalidCellSpacing(spacing) => {
                write!(f, "DOC cell spacing {spacing} exceeds 15840 twips")
            },
            Self::InvalidParagraphGroupId => {
                write!(f, "DOC paragraph-group identifier cannot be zero")
            },
            Self::InvalidRevisionAuthorIndex(index) => {
                write!(f, "DOC table revision author index {index} exceeds 32767")
            },
            Self::InvalidRevisionTimestamp(timestamp) => {
                write!(f, "DOC table revision DTTM {timestamp:#010x} is invalid")
            },
            Self::InvalidStyleBandSize(axis, size) => {
                write!(
                    f,
                    "DOC table-style {axis} band size {size} is outside 1..=3"
                )
            },
            Self::StyleBorderOutsideConditional => {
                write!(f, "DOC table-style borders must be placed inside sprmTCnf")
            },
            Self::InvalidConditionalProperties(error) => {
                write!(
                    f,
                    "DOC conditional table-style properties are invalid: {error}"
                )
            },
            Self::ConditionalPropertiesTooLong(size) => {
                write!(
                    f,
                    "DOC conditional table-style grpprl is {size} bytes; maximum is 253"
                )
            },
            Self::IncompleteCellBorderTypes(index) => {
                write!(f, "DOC cell {index} has an incomplete border-type override")
            },
            Self::InvalidPreferredWidth(property, width) => {
                write!(f, "DOC {property} has an invalid preferred width {width:?}")
            },
            Self::InvalidTableLookFlags(flags) => {
                write!(f, "DOC table look contains reserved flags {flags:#06x}")
            },
            Self::InvalidTablePosition(axis, value) => {
                write!(f, "DOC {axis} table position {value} cannot be encoded")
            },
            Self::InvalidWrapDistance(side, value) => {
                write!(
                    f,
                    "DOC {side} wrapping distance {value} exceeds 31680 twips"
                )
            },
            Self::NestedPreservedState => {
                write!(
                    f,
                    "DOC table revisions cannot contain nested preserved states"
                )
            },
        }
    }
}

impl std::error::Error for TapBuildError {}
