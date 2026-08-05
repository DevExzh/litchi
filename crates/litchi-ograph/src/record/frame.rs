//! Chart-element frame (`[MS-OGRAPH]` 2.4.53).

use crate::{Error, Result, record};
use litchi_biff::{Encoder, Kind as RecordKind, RecordRef};

const LEN: usize = 4;
const AUTO_SIZE: u16 = 0x0001;
const AUTO_POSITION: u16 = 0x0002;

/// The frame drawing style.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u16)]
pub enum Kind {
    /// A frame surrounding the chart element.
    Surrounding = 0x0000,
    /// A shadowed frame surrounding the chart element.
    Shadowed = 0x0004,
}

impl Kind {
    fn parse(value: u16) -> Result<Self> {
        match value {
            0x0000 => Ok(Self::Surrounding),
            0x0004 => Ok(Self::Shadowed),
            value => Err(Error::InvalidRecordValue {
                kind: Frame::KIND.get(),
                field: "frt",
                value: u64::from(value),
            }),
        }
    }
}

/// Type, automatic sizing, and automatic positioning of a frame.
///
/// Reserved flag bits are retained when parsing and replayed byte-for-byte.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Frame {
    kind: Kind,
    flags: u16,
}

impl Frame {
    /// BIFF record identifier.
    pub const KIND: RecordKind = RecordKind::from_wire(0x1032);

    /// Creates a frame with explicit size and position.
    pub const fn new(kind: Kind) -> Self {
        Self { kind, flags: 0 }
    }

    /// Decodes a framed record.
    pub fn parse(input: RecordRef<'_>) -> Result<Self> {
        Self::from_payload(record::payload(input, Self::KIND, LEN)?)
    }

    /// Decodes a payload supplied by an embedding host.
    pub fn from_payload(payload: &[u8]) -> Result<Self> {
        let payload = record::payload_bytes(Self::KIND, payload, LEN)?;
        let frame_kind = record::u16_at(payload, 0).ok_or(Error::InvalidRecordLength {
            kind: Self::KIND.get(),
            expected: LEN,
            actual: payload.len(),
        })?;
        let flags = record::u16_at(payload, 2).ok_or(Error::InvalidRecordLength {
            kind: Self::KIND.get(),
            expected: LEN,
            actual: payload.len(),
        })?;
        Ok(Self {
            kind: Kind::parse(frame_kind)?,
            flags,
        })
    }

    /// Selects automatic sizing in builder style.
    pub const fn auto_size(mut self, enabled: bool) -> Self {
        self.flags = set(self.flags, AUTO_SIZE, enabled);
        self
    }

    /// Selects automatic positioning in builder style.
    pub const fn auto_position(mut self, enabled: bool) -> Self {
        self.flags = set(self.flags, AUTO_POSITION, enabled);
        self
    }

    /// Frame drawing style.
    pub const fn kind(self) -> Kind {
        self.kind
    }

    /// Whether the size is calculated automatically.
    pub const fn is_auto_size(self) -> bool {
        self.flags & AUTO_SIZE != 0
    }

    /// Whether the position is calculated automatically.
    pub const fn is_auto_position(self) -> bool {
        self.flags & AUTO_POSITION != 0
    }

    /// Raw flags, including ignored reserved bits retained from input.
    pub const fn raw_flags(self) -> u16 {
        self.flags
    }

    /// Encodes the fixed-size payload without allocating.
    pub fn payload(self) -> [u8; LEN] {
        let mut bytes = [0; LEN];
        bytes[..2].copy_from_slice(&(self.kind as u16).to_le_bytes());
        bytes[2..].copy_from_slice(&self.flags.to_le_bytes());
        bytes
    }

    /// Appends the complete record to a bounded encoder.
    pub fn write(self, out: &mut Encoder) -> Result<()> {
        out.push(Self::KIND, &self.payload())?;
        Ok(())
    }
}

const fn set(flags: u16, mask: u16, enabled: bool) -> u16 {
    if enabled { flags | mask } else { flags & !mask }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn round_trips_reserved_flags_and_builds_concisely() {
        let parsed = Frame::from_payload(&[0, 0, 0xFC, 0xFF]).expect("frame");
        assert_eq!(parsed.raw_flags(), 0xFFFC);
        assert_eq!(parsed.payload(), [0, 0, 0xFC, 0xFF]);

        let frame = Frame::new(Kind::Shadowed)
            .auto_size(true)
            .auto_position(true);
        assert!(frame.is_auto_size());
        assert!(frame.is_auto_position());
        assert_eq!(frame.payload(), [4, 0, 3, 0]);
    }

    #[test]
    fn rejects_bad_length_and_unknown_kind() {
        assert!(Frame::from_payload(&[0, 0, 0]).is_err());
        assert!(matches!(
            Frame::from_payload(&[1, 0, 0, 0]),
            Err(Error::InvalidRecordValue { field: "frt", .. })
        ));
    }
}
