//! Exact binary codec for MS-PPT `OfficeArtClientAnchor` records.

use std::io::{self, Write};

use crate::package::{Error, Result};

use super::model::{Anchor, Data, Limits, RECT_LEN, Rect, SMALL_RECT_LEN, SmallRect};
use super::validation;

/// MS-ODRAW record type assigned to the host-specific client anchor.
pub const RECORD_TYPE: u16 = 0xF010;
pub(super) const HEADER_LEN: usize = 8;

impl Data {
    /// Parse an exact anchor payload using the default bound.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn parse(bytes: impl AsRef<[u8]>) -> Result<Self> {
        Self::parse_with_limits(bytes, Limits::default())
    }

    /// Parse an exact anchor payload with an explicit resource bound.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn parse_with_limits(bytes: impl AsRef<[u8]>, limits: Limits) -> Result<Self> {
        let payload = bytes.as_ref();
        validation::payload_len(payload.len(), limits)?;
        match payload.len() {
            SMALL_RECT_LEN => Ok(Self::Small(SmallRect::new(
                i16_at(payload, 2),
                i16_at(payload, 0),
                i16_at(payload, 4),
                i16_at(payload, 6),
            )?)),
            RECT_LEN => Ok(Self::Full(Rect::new(
                i32_at(payload, 4),
                i32_at(payload, 0),
                i32_at(payload, 8),
                i32_at(payload, 12),
            )?)),
            _ => unreachable!("payload length was validated"),
        }
    }

    /// Serialize only the host-defined payload.
    ///
    /// # Panics
    ///
    /// Panics only if writing to the in-memory buffer fails, which `Write for
    /// Vec<u8>` documents as impossible.
    #[must_use]
    #[allow(
        clippy::expect_used,
        reason = "`Write for Vec<u8>` is documented as infallible, so the result is always `Ok`"
    )]
    pub fn to_bytes(self) -> Vec<u8> {
        let mut bytes = Vec::with_capacity(self.encoded_len());
        self.write_to(&mut bytes)
            .expect("writing to a byte vector cannot fail");
        bytes
    }

    /// Write only the host-defined payload.
    ///
    /// # Errors
    ///
    /// Returns an error if the underlying writer reports an error.
    pub fn write_to<W: Write>(self, writer: &mut W) -> io::Result<()> {
        match self {
            Self::Small(value) => {
                writer.write_all(&value.top.to_le_bytes())?;
                writer.write_all(&value.left.to_le_bytes())?;
                writer.write_all(&value.right.to_le_bytes())?;
                writer.write_all(&value.bottom.to_le_bytes())?;
            },
            Self::Full(value) => {
                writer.write_all(&value.top.to_le_bytes())?;
                writer.write_all(&value.left.to_le_bytes())?;
                writer.write_all(&value.right.to_le_bytes())?;
                writer.write_all(&value.bottom.to_le_bytes())?;
            },
        }
        Ok(())
    }
}

impl Anchor {
    /// Parse one exact complete record using the default bound.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn parse(bytes: impl AsRef<[u8]>) -> Result<Self> {
        Self::parse_with_limits(bytes, Limits::default())
    }

    /// Parse one exact complete record with an explicit resource bound.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn parse_with_limits(bytes: impl AsRef<[u8]>, limits: Limits) -> Result<Self> {
        let record = bytes.as_ref();
        if record.len() < HEADER_LEN {
            return Err(Error::Corrupted(
                "OfficeArtClientAnchor header is truncated".into(),
            ));
        }
        let version_instance = u16::from_le_bytes([record[0], record[1]]);
        validation::header(
            version_instance & 0x000F,
            version_instance >> 4,
            u16::from_le_bytes([record[2], record[3]]),
        )?;
        let payload_len = u32::from_le_bytes([record[4], record[5], record[6], record[7]]) as usize;
        validation::payload_len(payload_len, limits)?;
        let expected_len = HEADER_LEN
            .checked_add(payload_len)
            .ok_or_else(|| Error::Corrupted("OfficeArtClientAnchor length overflows".into()))?;
        if record.len() != expected_len {
            return Err(Error::Corrupted(format!(
                "OfficeArtClientAnchor record length is {}, expected {expected_len}",
                record.len()
            )));
        }
        Ok(Self::new(Data::parse_with_limits(
            &record[HEADER_LEN..],
            limits,
        )?))
    }

    /// Serialize the complete record.
    ///
    /// # Panics
    ///
    /// Panics only if writing to the in-memory buffer fails, which `Write for
    /// Vec<u8>` documents as impossible.
    #[must_use]
    #[allow(
        clippy::expect_used,
        reason = "`Write for Vec<u8>` is documented as infallible, so the result is always `Ok`"
    )]
    pub fn to_bytes(self) -> Vec<u8> {
        let mut bytes = Vec::with_capacity(self.encoded_len());
        self.write_to(&mut bytes)
            .expect("writing to a byte vector cannot fail");
        bytes
    }

    /// Write the complete record without changing its compact/full representation.
    ///
    /// # Errors
    ///
    /// Returns an error if serialization fails or the underlying writer reports an error.
    pub fn write_to<W: Write>(self, writer: &mut W) -> io::Result<()> {
        writer.write_all(&0u16.to_le_bytes())?;
        writer.write_all(&RECORD_TYPE.to_le_bytes())?;
        #[allow(
            clippy::cast_possible_truncation,
            reason = "the anchor payload is exactly `SMALL_RECT_LEN` or `RECT_LEN` bytes, so its length always fits in `u32`"
        )]
        let payload_len = self.data().encoded_len() as u32;
        writer.write_all(&payload_len.to_le_bytes())?;
        self.data().write_to(writer)
    }
}

fn i16_at(bytes: &[u8], offset: usize) -> i16 {
    i16::from_le_bytes([bytes[offset], bytes[offset + 1]])
}

fn i32_at(bytes: &[u8], offset: usize) -> i32 {
    i32::from_le_bytes([
        bytes[offset],
        bytes[offset + 1],
        bytes[offset + 2],
        bytes[offset + 3],
    ])
}
