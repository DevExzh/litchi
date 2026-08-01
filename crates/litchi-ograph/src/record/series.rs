//! Series formatting and parent references (`[MS-OGRAPH]` 2.4.88 and
//! 2.4.92).

use litchi_core::BoundedU32;

use crate::raw::{Encoder, Kind, RecordRef};
use crate::{Error, Result, record};

const FORMAT_LEN: usize = 2;
const SMOOTH: u16 = 0x0001;
const BUBBLES_3D: u16 = 0x0002;
const SHADOW: u16 = 0x0004;

/// Properties of associated data points, markers, or lines.
///
/// Reserved bits are retained when parsing and replayed byte-for-byte.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct Format {
    flags: u16,
}

impl Format {
    /// BIFF record identifier.
    pub const KIND: Kind = Kind::new(0x105D);

    /// Creates formatting with all effects disabled.
    pub const fn new() -> Self {
        Self { flags: 0 }
    }

    /// Decodes a framed record.
    pub fn parse(input: RecordRef<'_>) -> Result<Self> {
        Self::from_payload(record::payload(input, Self::KIND, FORMAT_LEN)?)
    }

    /// Decodes a payload supplied by an embedding host.
    pub fn from_payload(payload: &[u8]) -> Result<Self> {
        let payload = record::payload_bytes(Self::KIND, payload, FORMAT_LEN)?;
        let flags = record::u16_at(payload, 0).ok_or(Error::InvalidRecordLength {
            kind: Self::KIND.get(),
            expected: FORMAT_LEN,
            actual: payload.len(),
        })?;
        Ok(Self { flags })
    }

    /// Enables or disables smooth lines in builder style.
    pub const fn smooth(mut self, enabled: bool) -> Self {
        self.flags = set(self.flags, SMOOTH, enabled);
        self
    }

    /// Enables or disables the 3-D bubble effect in builder style.
    pub const fn bubbles_3d(mut self, enabled: bool) -> Self {
        self.flags = set(self.flags, BUBBLES_3D, enabled);
        self
    }

    /// Enables or disables marker shadows in builder style.
    pub const fn shadow(mut self, enabled: bool) -> Self {
        self.flags = set(self.flags, SHADOW, enabled);
        self
    }

    /// Whether smooth lines are enabled.
    pub const fn is_smooth(self) -> bool {
        self.flags & SMOOTH != 0
    }

    /// Whether the 3-D bubble effect is enabled.
    pub const fn has_3d_bubbles(self) -> bool {
        self.flags & BUBBLES_3D != 0
    }

    /// Whether marker shadows are enabled.
    pub const fn has_shadow(self) -> bool {
        self.flags & SHADOW != 0
    }

    /// Raw flags, including ignored reserved bits retained from input.
    pub const fn raw_flags(self) -> u16 {
        self.flags
    }

    /// Encodes the fixed-size payload without allocating.
    pub fn payload(self) -> [u8; FORMAT_LEN] {
        self.flags.to_le_bytes()
    }

    /// Appends the complete record to a bounded encoder.
    pub fn write(self, out: &mut Encoder) -> Result<()> {
        out.push(Self::KIND, &self.payload())
    }
}

/// A one-based index into the current chart's `Series` records.
pub type Index = BoundedU32<1, 0x00FE>;

/// Series associated with a trendline or error bar.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Parent {
    series: Index,
}

impl Parent {
    /// BIFF record identifier.
    pub const KIND: Kind = Kind::new(0x104A);

    /// Creates a parent from an already range-proven series index.
    pub const fn new(series: Index) -> Self {
        Self { series }
    }

    /// Convenient checked construction from a primitive index.
    pub fn try_new(series: u16) -> Result<Self> {
        let index = Index::new(u32::from(series)).ok_or(Error::InvalidRecordValue {
            kind: Self::KIND.get(),
            field: "series",
            value: u64::from(series),
        })?;
        Ok(Self::new(index))
    }

    /// Decodes a framed record.
    pub fn parse(input: RecordRef<'_>) -> Result<Self> {
        Self::from_payload(record::payload(input, Self::KIND, 2)?)
    }

    /// Decodes a payload supplied by an embedding host.
    pub fn from_payload(payload: &[u8]) -> Result<Self> {
        let payload = record::payload_bytes(Self::KIND, payload, 2)?;
        let raw = record::u16_at(payload, 0).ok_or(Error::InvalidRecordLength {
            kind: Self::KIND.get(),
            expected: 2,
            actual: payload.len(),
        })?;
        Self::try_new(raw)
    }

    /// Range-proven one-based series index.
    pub const fn series(self) -> Index {
        self.series
    }

    /// Encodes the fixed-size payload without allocating.
    pub fn payload(self) -> [u8; 2] {
        (self.series.get() as u16).to_le_bytes()
    }

    /// Appends the complete record to a bounded encoder.
    pub fn write(self, out: &mut Encoder) -> Result<()> {
        out.push(Self::KIND, &self.payload())
    }
}

const fn set(flags: u16, mask: u16, enabled: bool) -> u16 {
    if enabled { flags | mask } else { flags & !mask }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn format_builders_preserve_unrecognized_bits() {
        let parsed = Format::from_payload(&0xFFF8_u16.to_le_bytes()).expect("format");
        let edited = parsed.smooth(true).shadow(true);
        assert_eq!(edited.raw_flags(), 0xFFFD);
        assert!(edited.is_smooth());
        assert!(edited.has_shadow());
        assert!(!edited.has_3d_bubbles());
    }

    #[test]
    fn parent_index_is_type_proven() {
        for value in [1, 2, 0x00FE] {
            let parent = Parent::try_new(value).expect("valid");
            assert_eq!(
                Parent::from_payload(&parent.payload()).expect("round trip"),
                parent
            );
        }
        assert!(Parent::try_new(0).is_err());
        assert!(Parent::try_new(0x00FF).is_err());
    }
}
