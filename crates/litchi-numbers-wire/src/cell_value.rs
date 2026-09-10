//! Allocation-free semantic projection of one Numbers cell payload.
//!
//! The Numbers package adapters own sidecar lookup and formula rendering. This
//! module only interprets the source-backed BNC or pre-BNC value envelope, so
//! both readers share the same version dispatch, formula precedence, scalar
//! fallback, and comment-reference semantics.

use std::fmt;

use litchi_iwa_common::formula::FiniteF64;

use crate::pre_bnc::PreBncCellView;
use crate::{BncCellView, CachedScalar, Error as WireError, StoredValue};

/// A source-backed value classification with no sidecar resolution.
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum ValueSource {
    /// The cell has no value-bearing representation.
    Empty,
    /// A finite native Numbers scalar.
    Number(FiniteF64),
    /// A finite Apple-epoch date scalar.
    Date(FiniteF64),
    /// A native Boolean scalar.
    Boolean(bool),
    /// A finite duration scalar in seconds.
    Duration(FiniteF64),
    /// An interned string-table identifier.
    Text(u32),
    /// An interned rich-text identifier.
    RichText(u32),
    /// A formula-table identifier.
    Formula(u32),
    /// A formula error with its optional error-table identifier.
    Error(Option<u32>),
}

/// The value and comment reference projected from one cell payload.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct CellValueSource {
    /// Source value classification. Sidecar identifiers are unresolved.
    pub value: ValueSource,
    /// Optional comment-table identifier, including identifier zero.
    pub comment_identifier: Option<u32>,
}

/// A typed failure while projecting a Numbers cell payload.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum DecodeError {
    /// The caller supplied no cell payload.
    Empty,
    /// The first byte is not a supported pre-BNC or BNC version.
    UnsupportedVersion(u8),
    /// The v5 BNC envelope is malformed.
    Bnc(WireError),
    /// A version 0–4 pre-BNC envelope is malformed.
    PreBnc(WireError),
    /// A numeric BNC cell carried a Boolean, date, or duration scalar cache.
    MismatchedNumericScalar,
    /// The cell type is not represented by the shared source vocabulary.
    UnsupportedCellType(u8),
}

impl fmt::Display for DecodeError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Empty => formatter.write_str("Numbers cell storage is empty"),
            Self::UnsupportedVersion(version) => {
                write!(
                    formatter,
                    "unsupported Numbers cell storage version {version}"
                )
            },
            Self::Bnc(error) => write!(formatter, "invalid Numbers BNC cell: {error}"),
            Self::PreBnc(error) => write!(formatter, "invalid Numbers pre-BNC cell: {error}"),
            Self::MismatchedNumericScalar => {
                formatter.write_str("Numbers numeric BNC cell has a mismatched scalar encoding")
            },
            Self::UnsupportedCellType(cell_type) => {
                write!(formatter, "unsupported Numbers cell type {cell_type}")
            },
        }
    }
}

impl std::error::Error for DecodeError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Bnc(error) | Self::PreBnc(error) => Some(error),
            Self::Empty
            | Self::UnsupportedVersion(_)
            | Self::MismatchedNumericScalar
            | Self::UnsupportedCellType(_) => None,
        }
    }
}

/// Decode one Numbers cell into a source-backed value classification.
///
/// Valid input does not allocate and does not copy the source bytes. Formula
/// and comment identifiers are intentionally left unresolved for the concrete
/// package adapter. Formula identifiers take precedence over cached scalar
/// values in both storage generations.
pub fn decode_cell_value(source: &[u8]) -> Result<CellValueSource, DecodeError> {
    match source.first().copied() {
        None => Err(DecodeError::Empty),
        Some(0..=4) => decode_pre_bnc(source),
        Some(5) => decode_bnc(source),
        Some(version) => Err(DecodeError::UnsupportedVersion(version)),
    }
}

fn decode_bnc(source: &[u8]) -> Result<CellValueSource, DecodeError> {
    let cell = BncCellView::parse(source).map_err(DecodeError::Bnc)?;
    let comment_identifier = cell.comment_identifier();
    let stored_value = cell.stored_value();
    if let StoredValue::Formula(identifier) = stored_value {
        return Ok(CellValueSource {
            value: ValueSource::Formula(identifier),
            comment_identifier,
        });
    }

    let value = match stored_value {
        StoredValue::Empty => ValueSource::Empty,
        StoredValue::Number => decode_bnc_number(cell.cached_scalar())?,
        StoredValue::Text(identifier) => ValueSource::Text(identifier),
        StoredValue::RichText(identifier) => ValueSource::RichText(identifier),
        StoredValue::Date => ValueSource::Date(decode_bnc_date(cell.cached_scalar())?),
        StoredValue::Boolean => ValueSource::Boolean(decode_bnc_boolean(cell.cached_scalar())),
        StoredValue::Duration => ValueSource::Duration(decode_bnc_duration(cell.cached_scalar())?),
        StoredValue::Error => ValueSource::Error(cell.formula_error_identifier()),
        StoredValue::Formula(identifier) => ValueSource::Formula(identifier),
        StoredValue::Unsupported(cell_type) => {
            return Err(DecodeError::UnsupportedCellType(cell_type));
        },
    };
    Ok(CellValueSource {
        value,
        comment_identifier,
    })
}

fn decode_bnc_number(cached_scalar: Option<CachedScalar>) -> Result<ValueSource, DecodeError> {
    match cached_scalar {
        Some(CachedScalar::Number(value)) => Ok(ValueSource::Number(value)),
        Some(CachedScalar::Boolean(_) | CachedScalar::Date(_) | CachedScalar::Duration(_)) => {
            Err(DecodeError::MismatchedNumericScalar)
        },
        Some(CachedScalar::Unsupported(_)) | None => Ok(ValueSource::Number(zero_scalar()?)),
    }
}

fn decode_bnc_date(cached_scalar: Option<CachedScalar>) -> Result<FiniteF64, DecodeError> {
    match cached_scalar {
        Some(CachedScalar::Date(value)) => Ok(value),
        Some(
            CachedScalar::Number(_)
            | CachedScalar::Boolean(_)
            | CachedScalar::Duration(_)
            | CachedScalar::Unsupported(_),
        )
        | None => zero_scalar(),
    }
}

fn decode_bnc_boolean(cached_scalar: Option<CachedScalar>) -> bool {
    matches!(cached_scalar, Some(CachedScalar::Boolean(true)))
}

fn decode_bnc_duration(cached_scalar: Option<CachedScalar>) -> Result<FiniteF64, DecodeError> {
    match cached_scalar {
        Some(CachedScalar::Duration(value)) => Ok(value),
        Some(
            CachedScalar::Number(_)
            | CachedScalar::Boolean(_)
            | CachedScalar::Date(_)
            | CachedScalar::Unsupported(_),
        )
        | None => zero_scalar(),
    }
}

fn decode_pre_bnc(source: &[u8]) -> Result<CellValueSource, DecodeError> {
    let cell = PreBncCellView::parse(source).map_err(DecodeError::PreBnc)?;
    let comment_identifier = cell.comment_identifier();
    if let Some(identifier) = cell.formula_identifier() {
        return Ok(CellValueSource {
            value: ValueSource::Formula(identifier),
            comment_identifier,
        });
    }

    let value = match cell.cell_type() {
        0 => ValueSource::Empty,
        2 => ValueSource::Number(cell.number().unwrap_or(zero_scalar()?)),
        3 => cell
            .string_identifier()
            .map_or(ValueSource::Empty, ValueSource::Text),
        5 => ValueSource::Date(cell.date().unwrap_or(zero_scalar()?)),
        6 => ValueSource::Boolean(cell.number().is_some_and(|value| value.get() != 0.0)),
        7 => ValueSource::Duration(cell.number().unwrap_or(zero_scalar()?)),
        8 => ValueSource::Error(cell.formula_error_identifier()),
        9 => cell
            .rich_text_identifier()
            .map_or(ValueSource::Empty, ValueSource::RichText),
        cell_type => return Err(DecodeError::UnsupportedCellType(cell_type)),
    };
    Ok(CellValueSource {
        value,
        comment_identifier,
    })
}

fn zero_scalar() -> Result<FiniteF64, DecodeError> {
    FiniteF64::new(0.0).map_err(|_| DecodeError::MismatchedNumericScalar)
}
