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
    pub fn parse(bytes: impl AsRef<[u8]>) -> Result<Self> {
        Self::parse_with_limits(bytes, Limits::default())
    }

    /// Parse an exact anchor payload with an explicit resource bound.
    pub fn parse_with_limits(bytes: impl AsRef<[u8]>, limits: Limits) -> Result<Self> {
        let bytes = bytes.as_ref();
        validation::payload_len(bytes.len(), limits)?;
        match bytes.len() {
            SMALL_RECT_LEN => Ok(Self::Small(SmallRect::new(
                i16_at(bytes, 2),
                i16_at(bytes, 0),
                i16_at(bytes, 4),
                i16_at(bytes, 6),
            )?)),
            RECT_LEN => Ok(Self::Full(Rect::new(
                i32_at(bytes, 4),
                i32_at(bytes, 0),
                i32_at(bytes, 8),
                i32_at(bytes, 12),
            )?)),
            _ => unreachable!("payload length was validated"),
        }
    }

    /// Serialize only the host-defined payload.
    pub fn to_bytes(self) -> Vec<u8> {
        let mut bytes = Vec::with_capacity(self.encoded_len());
        self.write_to(&mut bytes)
            .expect("writing to a byte vector cannot fail");
        bytes
    }

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
    pub fn parse(bytes: impl AsRef<[u8]>) -> Result<Self> {
        Self::parse_with_limits(bytes, Limits::default())
    }

    /// Parse one exact complete record with an explicit resource bound.
    pub fn parse_with_limits(bytes: impl AsRef<[u8]>, limits: Limits) -> Result<Self> {
        let bytes = bytes.as_ref();
        if bytes.len() < HEADER_LEN {
            return Err(Error::Corrupted(
                "OfficeArtClientAnchor header is truncated".into(),
            ));
        }
        let version_instance = u16::from_le_bytes([bytes[0], bytes[1]]);
        validation::header(
            version_instance & 0x000F,
            version_instance >> 4,
            u16::from_le_bytes([bytes[2], bytes[3]]),
        )?;
        let payload_len = u32::from_le_bytes([bytes[4], bytes[5], bytes[6], bytes[7]]) as usize;
        validation::payload_len(payload_len, limits)?;
        let expected_len = HEADER_LEN
            .checked_add(payload_len)
            .ok_or_else(|| Error::Corrupted("OfficeArtClientAnchor length overflows".into()))?;
        if bytes.len() != expected_len {
            return Err(Error::Corrupted(format!(
                "OfficeArtClientAnchor record length is {}, expected {expected_len}",
                bytes.len()
            )));
        }
        Ok(Self::new(Data::parse_with_limits(
            &bytes[HEADER_LEN..],
            limits,
        )?))
    }

    /// Serialize the complete record.
    pub fn to_bytes(self) -> Vec<u8> {
        let mut bytes = Vec::with_capacity(self.encoded_len());
        self.write_to(&mut bytes)
            .expect("writing to a byte vector cannot fail");
        bytes
    }

    /// Write the complete record without changing its compact/full representation.
    pub fn write_to<W: Write>(self, writer: &mut W) -> io::Result<()> {
        writer.write_all(&0u16.to_le_bytes())?;
        writer.write_all(&RECORD_TYPE.to_le_bytes())?;
        writer.write_all(&(self.data().encoded_len() as u32).to_le_bytes())?;
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
