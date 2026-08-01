//! Pie-series formatting (`[MS-OGRAPH]` 2.4.77).

use litchi_core::BoundedU32;

use crate::raw::{Encoder, Kind, RecordRef};
use crate::{Error, Result, record};

const LEN: usize = 2;

/// A non-negative pie explosion percentage representable by `pcExplode`.
pub type Explode = BoundedU32<0, 0x7FFF>;

/// Distance of a pie data point or series from its center.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Format {
    explode: Explode,
}

impl Format {
    /// BIFF record identifier.
    pub const KIND: Kind = Kind::new(0x100B);

    /// Creates a format from an already range-proven value.
    pub const fn new(explode: Explode) -> Self {
        Self { explode }
    }

    /// Convenient checked construction from a primitive percentage.
    pub fn try_new(explode: u16) -> Result<Self> {
        let value = Explode::new(u32::from(explode)).ok_or(Error::InvalidRecordValue {
            kind: Self::KIND.get(),
            field: "pcExplode",
            value: u64::from(explode),
        })?;
        Ok(Self::new(value))
    }

    /// Decodes a framed record.
    pub fn parse(input: RecordRef<'_>) -> Result<Self> {
        Self::from_payload(record::payload(input, Self::KIND, LEN)?)
    }

    /// Decodes a payload supplied by an embedding host.
    pub fn from_payload(payload: &[u8]) -> Result<Self> {
        let payload = record::payload_bytes(Self::KIND, payload, LEN)?;
        let raw = record::u16_at(payload, 0).ok_or(Error::InvalidRecordLength {
            kind: Self::KIND.get(),
            expected: LEN,
            actual: payload.len(),
        })?;
        Self::try_new(raw)
    }

    /// Range-proven explosion percentage.
    pub const fn explode(self) -> Explode {
        self.explode
    }

    /// Encodes the fixed-size payload without allocating.
    pub fn payload(self) -> [u8; LEN] {
        (self.explode.get() as u16).to_le_bytes()
    }

    /// Appends the complete record to a bounded encoder.
    pub fn write(self, out: &mut Encoder) -> Result<()> {
        out.push(Self::KIND, &self.payload())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn accepts_spec_examples_and_rejects_negative_encoding() {
        for value in [0, 100, 400, 0x7FFF] {
            let format = Format::try_new(value).expect("valid");
            assert_eq!(
                Format::from_payload(&format.payload()).expect("round trip"),
                format
            );
        }
        assert!(Format::try_new(0x8000).is_err());
        assert!(Format::from_payload(&(-1_i16).to_le_bytes()).is_err());
    }
}
