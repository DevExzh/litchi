//! Chart-group lines and their opaque link (`[MS-OGRAPH]` 2.4.32 and
//! 2.4.33).

use crate::{Error, Result, record};
use litchi_biff::{Encoder, Kind as RecordKind, RecordRef};

/// Type of chart-group line.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
#[repr(u16)]
pub enum Kind {
    /// Drop lines below data points.
    Drop = 0,
    /// High-low lines around data points.
    HighLow = 1,
    /// Lines connecting series.
    Series = 2,
    /// Leader lines connecting labels to data points.
    Leader = 3,
}

impl Kind {
    fn parse(value: u16) -> Result<Self> {
        match value {
            0 => Ok(Self::Drop),
            1 => Ok(Self::HighLow),
            2 => Ok(Self::Series),
            3 => Ok(Self::Leader),
            value => Err(Error::InvalidRecordValue {
                kind: Line::KIND.get(),
                field: "id",
                value: u64::from(value),
            }),
        }
    }
}

/// Presence of one line type on a chart group.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Line {
    kind: Kind,
}

impl Line {
    /// BIFF record identifier.
    pub const KIND: RecordKind = RecordKind::from_wire(0x101C);

    /// Creates a line marker.
    pub const fn new(kind: Kind) -> Self {
        Self { kind }
    }

    /// Decodes a framed record.
    pub fn parse(input: RecordRef<'_>) -> Result<Self> {
        Self::from_payload(record::payload(input, Self::KIND, 2)?)
    }

    /// Decodes a payload supplied by an embedding host.
    pub fn from_payload(payload: &[u8]) -> Result<Self> {
        let payload = record::payload_bytes(Self::KIND, payload, 2)?;
        let value = record::u16_at(payload, 0).ok_or(Error::InvalidRecordLength {
            kind: Self::KIND.get(),
            expected: 2,
            actual: payload.len(),
        })?;
        Ok(Self::new(Kind::parse(value)?))
    }

    /// Line type.
    pub const fn kind(self) -> Kind {
        self.kind
    }

    /// Encodes the fixed-size payload without allocating.
    pub fn payload(self) -> [u8; 2] {
        (self.kind as u16).to_le_bytes()
    }

    /// Appends the complete record to a bounded encoder.
    pub fn write(self, out: &mut Encoder) -> Result<()> {
        out.push(Self::KIND, &self.payload())?;
        Ok(())
    }
}

/// Written-but-unused chart link bytes, retained exactly.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Link {
    bytes: [u8; 10],
}

impl Link {
    /// BIFF record identifier.
    pub const KIND: RecordKind = RecordKind::from_wire(0x1022);

    /// Creates a link from the ten opaque bytes that will be preserved.
    pub const fn new(bytes: [u8; 10]) -> Self {
        Self { bytes }
    }

    /// Decodes a framed record.
    pub fn parse(input: RecordRef<'_>) -> Result<Self> {
        Self::from_payload(record::payload(input, Self::KIND, 10)?)
    }

    /// Decodes a payload supplied by an embedding host.
    pub fn from_payload(payload: &[u8]) -> Result<Self> {
        let payload = record::payload_bytes(Self::KIND, payload, 10)?;
        let bytes = <[u8; 10]>::try_from(payload).map_err(|_| Error::InvalidRecordLength {
            kind: Self::KIND.get(),
            expected: 10,
            actual: payload.len(),
        })?;
        Ok(Self::new(bytes))
    }

    /// Opaque bytes retained from input.
    pub const fn bytes(self) -> [u8; 10] {
        self.bytes
    }

    /// Encodes the fixed-size payload without allocating.
    pub const fn payload(self) -> [u8; 10] {
        self.bytes
    }

    /// Appends the complete record to a bounded encoder.
    pub fn write(self, out: &mut Encoder) -> Result<()> {
        out.push(Self::KIND, &self.bytes)?;
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn line_kinds_round_trip() {
        for kind in [Kind::Drop, Kind::HighLow, Kind::Series, Kind::Leader] {
            let line = Line::new(kind);
            assert_eq!(
                Line::from_payload(&line.payload()).expect("round trip"),
                line
            );
        }
        assert!(Line::from_payload(&4_u16.to_le_bytes()).is_err());
    }

    #[test]
    fn link_retains_all_opaque_bytes() {
        let bytes = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10];
        assert_eq!(Link::from_payload(&bytes).expect("link").payload(), bytes);
        assert!(Link::from_payload(&bytes[..9]).is_err());
    }
}
