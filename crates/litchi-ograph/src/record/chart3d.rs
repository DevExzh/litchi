//! Three-dimensional chart primitives.

use crate::raw::{Encoder, Kind, RecordRef};
use crate::{Error, Result, record};

/// Base shape of a 3-D bar or column data point.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u8)]
pub enum Riser {
    /// Rectangular base.
    Rectangle = 0,
    /// Elliptical base.
    Ellipse = 1,
}

impl Riser {
    fn parse(value: u8) -> Result<Self> {
        match value {
            0 => Ok(Self::Rectangle),
            1 => Ok(Self::Ellipse),
            value => Err(Error::InvalidRecordValue {
                kind: BarShape::KIND.get(),
                field: "riser",
                value: u64::from(value),
            }),
        }
    }
}

/// Taper from the base to the tip of a 3-D data point.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u8)]
pub enum Taper {
    /// No taper.
    None = 0,
    /// Tapers to a point at the maximum value.
    Point = 1,
    /// Tapers toward the projected maximum and clips at the point value.
    Clipped = 2,
}

impl Taper {
    fn parse(value: u8) -> Result<Self> {
        match value {
            0 => Ok(Self::None),
            1 => Ok(Self::Point),
            2 => Ok(Self::Clipped),
            value => Err(Error::InvalidRecordValue {
                kind: BarShape::KIND.get(),
                field: "taper",
                value: u64::from(value),
            }),
        }
    }
}

/// Shape of data points in a 3-D bar or column chart group.
///
/// See `[MS-OGRAPH]` 2.4.23.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct BarShape {
    riser: Riser,
    taper: Taper,
}

impl BarShape {
    /// BIFF record identifier.
    pub const KIND: Kind = Kind::new(0x105F);

    /// Creates a 3-D bar shape.
    pub const fn new(riser: Riser, taper: Taper) -> Self {
        Self { riser, taper }
    }

    /// Decodes a framed record.
    pub fn parse(input: RecordRef<'_>) -> Result<Self> {
        Self::from_payload(record::payload(input, Self::KIND, 2)?)
    }

    /// Decodes a payload supplied by an embedding host.
    pub fn from_payload(payload: &[u8]) -> Result<Self> {
        let payload = record::payload_bytes(Self::KIND, payload, 2)?;
        let riser = payload.first().copied().ok_or(Error::InvalidRecordLength {
            kind: Self::KIND.get(),
            expected: 2,
            actual: payload.len(),
        })?;
        let taper = payload.get(1).copied().ok_or(Error::InvalidRecordLength {
            kind: Self::KIND.get(),
            expected: 2,
            actual: payload.len(),
        })?;
        Ok(Self::new(Riser::parse(riser)?, Taper::parse(taper)?))
    }

    /// Base shape.
    pub const fn riser(self) -> Riser {
        self.riser
    }

    /// Taper behavior.
    pub const fn taper(self) -> Taper {
        self.taper
    }

    /// Encodes the fixed-size payload without allocating.
    pub const fn payload(self) -> [u8; 2] {
        [self.riser as u8, self.taper as u8]
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
    fn all_spec_combinations_round_trip() {
        for riser in [Riser::Rectangle, Riser::Ellipse] {
            for taper in [Taper::None, Taper::Point, Taper::Clipped] {
                let shape = BarShape::new(riser, taper);
                assert_eq!(
                    BarShape::from_payload(&shape.payload()).expect("round trip"),
                    shape
                );
            }
        }
        assert!(BarShape::from_payload(&[2, 0]).is_err());
        assert!(BarShape::from_payload(&[0, 3]).is_err());
    }
}
