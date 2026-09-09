//! Borrowed view over legacy Numbers cell storage.
//!
//! Versions zero through four use a fixed header followed by a flag-ordered
//! sequence of four- and eight-byte fields.  This module keeps that legacy
//! representation source-backed while the package adapters migrate away from
//! their local parsers.  Unknown flags and bytes after the recognized fields
//! remain opaque; the source slice is always the preservation authority.

use std::fmt;

use litchi_iwa_common::formula::FiniteF64;

use crate::{Error, Result};

const MAX_SUPPORTED_VERSION: u8 = 4;
const SHORT_HEADER_LEN: usize = 8;
const LONG_HEADER_LEN: usize = 12;

const FLAG_UNKNOWN_0002: u32 = 0x0000_0002;
const FLAG_UNKNOWN_0080: u32 = 0x0000_0080;
const FLAG_UNKNOWN_0400: u32 = 0x0000_0400;
const FLAG_UNKNOWN_0800: u32 = 0x0000_0800;
const FLAG_UNKNOWN_0004: u32 = 0x0000_0004;
const FLAG_FORMULA: u32 = 0x0000_0008;
const FLAG_FORMULA_ERROR: u32 = 0x0000_0100;
const FLAG_RICH_TEXT: u32 = 0x0000_0200;
const FLAG_COMMENT: u32 = 0x0000_1000;
const FLAG_UNKNOWN_2000: u32 = 0x0000_2000;
const FLAG_STRING: u32 = 0x0000_0010;
const FLAG_NUMBER: u32 = 0x0000_0020;
const FLAG_DATE: u32 = 0x0000_0040;
const FLAG_UNKNOWN_010000: u32 = 0x0001_0000;
const FLAG_UNKNOWN_080000: u32 = 0x0008_0000;
const FLAG_UNKNOWN_020000: u32 = 0x0002_0000;
const FLAG_UNKNOWN_040000: u32 = 0x0004_0000;
const FLAG_UNKNOWN_100000: u32 = 0x0010_0000;
const FLAG_UNKNOWN_200000: u32 = 0x0020_0000;
const FLAG_UNKNOWN_400000: u32 = 0x0040_0000;
const FLAG_UNKNOWN_800000: u32 = 0x0080_0000;

/// The exact fixed-width field order used by legacy Numbers cell storage.
///
/// The names intentionally describe only the fields projected by this view;
/// all other recognized slots are validated for width and retained through
/// [`PreBncCellView::source`]. Bytes after this layout are available through
/// [`PreBncCellView::opaque_tail`].
const FIELD_LAYOUT: &[(u32, usize)] = &[
    (FLAG_UNKNOWN_0002, 4),
    (FLAG_UNKNOWN_0080, 4),
    (FLAG_UNKNOWN_0400, 4),
    (FLAG_UNKNOWN_0800, 4),
    (FLAG_UNKNOWN_0004, 4),
    (FLAG_FORMULA, 4),
    (FLAG_FORMULA_ERROR, 4),
    (FLAG_RICH_TEXT, 4),
    (FLAG_COMMENT, 4),
    (FLAG_UNKNOWN_2000, 4),
    (FLAG_STRING, 4),
    (FLAG_NUMBER, 8),
    (FLAG_DATE, 8),
    (FLAG_UNKNOWN_010000, 4),
    (FLAG_UNKNOWN_080000, 4),
    (FLAG_UNKNOWN_020000, 4),
    (FLAG_UNKNOWN_040000, 4),
    (FLAG_UNKNOWN_100000, 4),
    (FLAG_UNKNOWN_200000, 4),
    (FLAG_UNKNOWN_400000, 4),
    (FLAG_UNKNOWN_800000, 4),
];

/// A borrowed, allocation-free view over one legacy Numbers cell payload.
#[derive(Clone, Copy, PartialEq)]
pub struct PreBncCellView<'source> {
    source: &'source [u8],
    version: u8,
    cell_type: u8,
    flags: u32,
    number: Option<FiniteF64>,
    date: Option<FiniteF64>,
    string_identifier: Option<u32>,
    rich_text_identifier: Option<u32>,
    formula_identifier: Option<u32>,
    formula_error_identifier: Option<u32>,
    comment_identifier: Option<u32>,
    opaque_tail: &'source [u8],
}

impl fmt::Debug for PreBncCellView<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("PreBncCellView")
            .field("version", &self.version)
            .field("cell_type", &self.cell_type)
            .field("flags", &format_args!("0x{:08x}", self.flags))
            .field("has_number", &self.number.is_some())
            .field("has_date", &self.date.is_some())
            .field("has_string_identifier", &self.string_identifier.is_some())
            .field(
                "has_rich_text_identifier",
                &self.rich_text_identifier.is_some(),
            )
            .field("has_formula_identifier", &self.formula_identifier.is_some())
            .field(
                "has_formula_error_identifier",
                &self.formula_error_identifier.is_some(),
            )
            .field("has_comment_identifier", &self.comment_identifier.is_some())
            .field("source_bytes", &self.source.len())
            .field("opaque_tail_bytes", &self.opaque_tail.len())
            .finish()
    }
}

impl<'source> PreBncCellView<'source> {
    /// Parse one legacy cell without allocating a field map or copying bytes.
    ///
    /// Versions zero and one use an eight-byte header; versions two through
    /// four use a twelve-byte header.  A version greater than four is
    /// rejected so callers cannot accidentally route BNC v5 through the
    /// legacy layout.  Unknown flag bits and trailing bytes are preserved as
    /// opaque source data.
    ///
    /// # Errors
    ///
    /// Returns [`Error::ParseError`] for an empty/truncated payload,
    /// unsupported version, malformed fixed-width field, or non-finite number
    /// or date field.
    pub fn parse(source: &'source [u8]) -> Result<Self> {
        let version = source
            .first()
            .copied()
            .ok_or_else(|| Error::ParseError("Empty Numbers pre-BNC cell payload".to_owned()))?;
        if version > MAX_SUPPORTED_VERSION {
            return Err(Error::ParseError(format!(
                "Numbers pre-BNC cell view does not support storage version {version}"
            )));
        }

        let header_length = if version <= 1 {
            SHORT_HEADER_LEN
        } else {
            LONG_HEADER_LEN
        };
        if source.len() < header_length {
            return Err(Error::ParseError(
                "Truncated Numbers pre-BNC cell header".to_owned(),
            ));
        }

        let cell_type = source[if version == 4 { 1 } else { 2 }];
        let flags = if version <= 1 {
            u32::from(u16::from_le_bytes([source[4], source[5]]))
        } else {
            u32::from_le_bytes(source[4..8].try_into().map_err(|_error| {
                Error::ParseError("Truncated Numbers pre-BNC flags".to_owned())
            })?)
        };

        let mut cursor = header_length;
        let mut number = None;
        let mut date = None;
        let mut string_identifier = None;
        let mut rich_text_identifier = None;
        let mut formula_identifier = None;
        let mut formula_error_identifier = None;
        let mut comment_identifier = None;

        for &(flag, size) in FIELD_LAYOUT {
            if flags & flag == 0 {
                continue;
            }
            let field = take_field(source, &mut cursor, size, flag)?;
            match flag {
                FLAG_FORMULA => formula_identifier = Some(read_u32(field, flag)?),
                FLAG_FORMULA_ERROR => formula_error_identifier = Some(read_u32(field, flag)?),
                FLAG_COMMENT => comment_identifier = Some(read_u32(field, flag)?),
                FLAG_RICH_TEXT => rich_text_identifier = Some(read_u32(field, flag)?),
                FLAG_STRING => string_identifier = Some(read_u32(field, flag)?),
                FLAG_NUMBER => number = Some(read_finite_f64(field, flag)?),
                FLAG_DATE => date = Some(read_finite_f64(field, flag)?),
                _ => {},
            }
        }

        Ok(Self {
            source,
            version,
            cell_type,
            flags,
            number,
            date,
            string_identifier,
            rich_text_identifier,
            formula_identifier,
            formula_error_identifier,
            comment_identifier,
            opaque_tail: &source[cursor..],
        })
    }

    /// Return the legacy storage version.
    #[must_use]
    pub const fn version(self) -> u8 {
        self.version
    }

    /// Return the native legacy cell type byte.
    #[must_use]
    pub const fn cell_type(self) -> u8 {
        self.cell_type
    }

    /// Return the decoded legacy field flags.
    #[must_use]
    pub const fn flags(self) -> u32 {
        self.flags
    }

    /// Return the finite number field, when present.
    #[must_use]
    pub const fn number(self) -> Option<FiniteF64> {
        self.number
    }

    /// Return the finite date field, when present.
    #[must_use]
    pub const fn date(self) -> Option<FiniteF64> {
        self.date
    }

    /// Return the optional interned string identifier, including zero.
    #[must_use]
    pub const fn string_identifier(self) -> Option<u32> {
        self.string_identifier
    }

    /// Return the optional rich-text identifier, including zero.
    #[must_use]
    pub const fn rich_text_identifier(self) -> Option<u32> {
        self.rich_text_identifier
    }

    /// Return the optional formula identifier, including zero.
    #[must_use]
    pub const fn formula_identifier(self) -> Option<u32> {
        self.formula_identifier
    }

    /// Return the optional formula-error identifier, including zero.
    #[must_use]
    pub const fn formula_error_identifier(self) -> Option<u32> {
        self.formula_error_identifier
    }

    /// Return the optional comment identifier, including zero.
    #[must_use]
    pub const fn comment_identifier(self) -> Option<u32> {
        self.comment_identifier
    }

    /// Return the exact caller-owned source payload.
    #[must_use]
    pub const fn source(self) -> &'source [u8] {
        self.source
    }

    /// Return bytes after the recognized fixed-layout fields.
    ///
    /// Unknown flags are deliberately not rejected.  Their bytes, together
    /// with any producer-specific trailing bytes, remain in this opaque
    /// source-backed suffix.
    #[must_use]
    pub const fn opaque_tail(self) -> &'source [u8] {
        self.opaque_tail
    }
}

fn take_field<'source>(
    source: &'source [u8],
    cursor: &mut usize,
    length: usize,
    flag: u32,
) -> Result<&'source [u8]> {
    let end = cursor.checked_add(length).ok_or_else(|| {
        Error::ParseError(format!(
            "Numbers pre-BNC field offset overflow for flag 0x{flag:08x}"
        ))
    })?;
    let field = source.get(*cursor..end).ok_or_else(|| {
        Error::ParseError(format!("Truncated Numbers pre-BNC field 0x{flag:08x}"))
    })?;
    *cursor = end;
    Ok(field)
}

fn read_u32(field: &[u8], flag: u32) -> Result<u32> {
    let bytes: [u8; 4] = field.try_into().map_err(|_error| {
        Error::ParseError(format!(
            "Numbers pre-BNC field 0x{flag:08x} is not four bytes"
        ))
    })?;
    Ok(u32::from_le_bytes(bytes))
}

fn read_finite_f64(field: &[u8], flag: u32) -> Result<FiniteF64> {
    let bytes: [u8; 8] = field.try_into().map_err(|_error| {
        Error::ParseError(format!(
            "Numbers pre-BNC field 0x{flag:08x} is not eight bytes"
        ))
    })?;
    FiniteF64::new(f64::from_le_bytes(bytes)).map_err(|_error| {
        Error::ParseError(format!(
            "Numbers pre-BNC field 0x{flag:08x} must contain a finite scalar"
        ))
    })
}
