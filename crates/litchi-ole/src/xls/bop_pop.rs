//! BIFF8 `BopPop` record (0x1061, MS-XLS 2.4.25) of the Chart Sheet
//! substream (MS-XLS 2.1): the attributes of a bar of pie or pie of pie
//! chart group.
//!
//! Everything in this module is INERT: values are stored verbatim and no
//! chart group is split or rendered. Fields ignored for the active `split`
//! mode or when `fAutoSplit` is set are preserved verbatim (MS-XLS 2.4.25).
//!
//! # References
//!
//! - MS-XLS 2.4.25 (BopPop), 2.5.14 (Boolean), 2.5.342 (Xnum)

use super::{XlsError, XlsResult};

/// Record type of the `BopPop` record (MS-XLS 2.4.25).
pub(crate) const BOP_POP_RECORD_TYPE: u16 = 0x1061;

/// Byte length of a `BopPop` record payload.
const PAYLOAD_LEN: usize = 22;
/// Flags bit: `fHasShadow`.
const FLAG_HAS_SHADOW: u16 = 0x0001;
/// Maximum `iSplitPos` value (MS-XLS 2.4.25).
const MAX_SPLIT_POSITION: i16 = 32000;
/// Minimum `pcPie2Size` value (MS-XLS 2.4.25).
const MIN_PIE2_SIZE: i16 = 5;
/// Maximum `pcPie2Size` value (MS-XLS 2.4.25).
const MAX_PIE2_SIZE: i16 = 200;
/// Maximum `pcGap` value (MS-XLS 2.4.25).
const MAX_GAP: i16 = 500;

fn invalid(message: impl Into<String>) -> XlsError {
    XlsError::InvalidRecord {
        record_type: BOP_POP_RECORD_TYPE,
        message: message.into(),
    }
}

/// The `pst` chart group subtype (MS-XLS 2.4.25).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u8)]
pub enum XlsBopPopSubtype {
    /// 0x01: pie of pie chart group.
    PieOfPie = 0x01,
    /// 0x02: bar of pie chart group.
    BarOfPie = 0x02,
}

impl XlsBopPopSubtype {
    fn parse(value: u8) -> XlsResult<Self> {
        match value {
            0x01 => Ok(Self::PieOfPie),
            0x02 => Ok(Self::BarOfPie),
            other => Err(invalid(format!(
                "BopPop pst {other:#04X} is not a defined chart group subtype"
            ))),
        }
    }
}

/// The `split` mode selecting what determines the split between the primary
/// pie and the secondary bar/pie (MS-XLS 2.4.25).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u16)]
pub enum XlsBopPopSplit {
    /// 0x0000: split by data point position (`iSplitPos`).
    Position = 0x0000,
    /// 0x0001: split by threshold value (`numSplitValue`).
    Value = 0x0001,
    /// 0x0002: split by percentage threshold (`pcSplitPercent`).
    Percent = 0x0002,
    /// 0x0003: custom split, arranged by the user in a following
    /// `BopPopCustom` record.
    Custom = 0x0003,
}

impl XlsBopPopSplit {
    fn parse(value: u16) -> XlsResult<Self> {
        match value {
            0x0000 => Ok(Self::Position),
            0x0001 => Ok(Self::Value),
            0x0002 => Ok(Self::Percent),
            0x0003 => Ok(Self::Custom),
            other => Err(invalid(format!(
                "BopPop split {other:#06X} is not a defined split mode"
            ))),
        }
    }
}

/// Typed `BopPop` record content (MS-XLS 2.4.25): the attributes of a bar of
/// pie or pie of pie chart group.
///
/// The 15 `reserved` flags bits (MUST be ignored) are preserved verbatim so
/// the record round-trips unchanged.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct XlsBopPop {
    /// The chart group subtype (`pst`).
    subtype: XlsBopPopSubtype,
    /// Whether the split point is determined automatically (`fAutoSplit`).
    auto_split: bool,
    /// What determines the split (`split`).
    split: XlsBopPopSplit,
    /// Number of trailing data points in the secondary bar/pie (`iSplitPos`),
    /// in 0..=32000.
    split_position: i16,
    /// Percentage below which data points move to the secondary bar/pie
    /// (`pcSplitPercent`).
    split_percent: i16,
    /// Size of the secondary bar/pie as a percentage of the primary pie
    /// (`pcPie2Size`), in 5..=200.
    pie2_size_percent: i16,
    /// Distance between the primary pie and the secondary bar/pie, as a
    /// percentage of their average width (`pcGap`), in 0..=500.
    gap_percent: i16,
    /// Threshold value for a value split (`numSplitValue`).
    split_value: f64,
    /// Raw flags word: `fHasShadow` and the 15 reserved bits.
    flags: u16,
}

impl XlsBopPop {
    /// Parse a `BopPop` record payload.
    pub fn parse(data: &[u8]) -> XlsResult<Self> {
        if data.len() != PAYLOAD_LEN {
            return Err(XlsError::InvalidLength {
                expected: PAYLOAD_LEN,
                found: data.len(),
            });
        }
        // Boolean (MS-XLS 2.5.14): only 0x00 and 0x01 are legal.
        let auto_split = match data[1] {
            0x00 => false,
            0x01 => true,
            other => {
                return Err(invalid(format!(
                    "BopPop fAutoSplit {other:#04X} is not a Boolean"
                )));
            },
        };
        let split_position = i16::from_le_bytes([data[4], data[5]]);
        if !(0..=MAX_SPLIT_POSITION).contains(&split_position) {
            return Err(invalid(format!(
                "BopPop iSplitPos {split_position} is outside 0..={MAX_SPLIT_POSITION}"
            )));
        }
        let pie2_size_percent = i16::from_le_bytes([data[8], data[9]]);
        if !(MIN_PIE2_SIZE..=MAX_PIE2_SIZE).contains(&pie2_size_percent) {
            return Err(invalid(format!(
                "BopPop pcPie2Size {pie2_size_percent} is outside {MIN_PIE2_SIZE}..={MAX_PIE2_SIZE}"
            )));
        }
        let gap_percent = i16::from_le_bytes([data[10], data[11]]);
        if !(0..=MAX_GAP).contains(&gap_percent) {
            return Err(invalid(format!(
                "BopPop pcGap {gap_percent} is outside 0..={MAX_GAP}"
            )));
        }
        Ok(Self {
            subtype: XlsBopPopSubtype::parse(data[0])?,
            auto_split,
            split: XlsBopPopSplit::parse(u16::from_le_bytes([data[2], data[3]]))?,
            split_position,
            split_percent: i16::from_le_bytes([data[6], data[7]]),
            pie2_size_percent,
            gap_percent,
            split_value: f64::from_le_bytes(data[12..20].try_into().expect("length checked")),
            flags: u16::from_le_bytes([data[20], data[21]]),
        })
    }

    /// Serialize back to a complete `BopPop` record payload.
    pub fn to_payload(&self) -> Vec<u8> {
        let mut payload = Vec::with_capacity(PAYLOAD_LEN);
        payload.push(self.subtype as u8);
        payload.push(u8::from(self.auto_split));
        payload.extend_from_slice(&(self.split as u16).to_le_bytes());
        payload.extend_from_slice(&self.split_position.to_le_bytes());
        payload.extend_from_slice(&self.split_percent.to_le_bytes());
        payload.extend_from_slice(&self.pie2_size_percent.to_le_bytes());
        payload.extend_from_slice(&self.gap_percent.to_le_bytes());
        payload.extend_from_slice(&self.split_value.to_le_bytes());
        payload.extend_from_slice(&self.flags.to_le_bytes());
        payload
    }

    /// The chart group subtype (`pst`).
    pub fn subtype(&self) -> XlsBopPopSubtype {
        self.subtype
    }

    /// Whether the split point is determined automatically (`fAutoSplit`).
    pub fn auto_split(&self) -> bool {
        self.auto_split
    }

    /// What determines the split (`split`).
    pub fn split(&self) -> XlsBopPopSplit {
        self.split
    }

    /// Number of trailing data points in the secondary bar/pie (`iSplitPos`).
    pub fn split_position(&self) -> i16 {
        self.split_position
    }

    /// Percentage below which data points move to the secondary bar/pie
    /// (`pcSplitPercent`).
    pub fn split_percent(&self) -> i16 {
        self.split_percent
    }

    /// Size of the secondary bar/pie as a percentage of the primary pie
    /// (`pcPie2Size`).
    pub fn pie2_size_percent(&self) -> i16 {
        self.pie2_size_percent
    }

    /// Distance between the primary pie and the secondary bar/pie (`pcGap`).
    pub fn gap_percent(&self) -> i16 {
        self.gap_percent
    }

    /// Threshold value for a value split (`numSplitValue`).
    pub fn split_value(&self) -> f64 {
        self.split_value
    }

    /// Whether one or more data points have shadows (`fHasShadow`).
    pub fn has_shadow(&self) -> bool {
        self.flags & FLAG_HAS_SHADOW != 0
    }

    /// Raw flags word, including the 15 reserved bits.
    pub fn flags(&self) -> u16 {
        self.flags
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// Build a BopPop payload. `fields` are `iSplitPos`, `pcSplitPercent`,
    /// `pcPie2Size`, and `pcGap`, in record order.
    fn record(
        subtype: u8,
        auto_split: u8,
        split: u16,
        fields: [i16; 4],
        value: f64,
        flags: u16,
    ) -> Vec<u8> {
        let mut data = vec![subtype, auto_split];
        data.extend_from_slice(&split.to_le_bytes());
        for field in fields {
            data.extend_from_slice(&field.to_le_bytes());
        }
        data.extend_from_slice(&value.to_le_bytes());
        data.extend_from_slice(&flags.to_le_bytes());
        data
    }

    #[test]
    fn round_trip_all_subtypes_and_splits() {
        for (subtype, expected_subtype) in [
            (0x01, XlsBopPopSubtype::PieOfPie),
            (0x02, XlsBopPopSubtype::BarOfPie),
        ] {
            for (split, expected_split) in [
                (0x0000, XlsBopPopSplit::Position),
                (0x0001, XlsBopPopSplit::Value),
                (0x0002, XlsBopPopSplit::Percent),
                (0x0003, XlsBopPopSplit::Custom),
            ] {
                let bytes = record(subtype, 0x01, split, [3, 40, 75, 100], 2.5, 0x0001);
                let parsed = XlsBopPop::parse(&bytes).unwrap();
                assert_eq!(parsed.subtype(), expected_subtype);
                assert!(parsed.auto_split());
                assert_eq!(parsed.split(), expected_split);
                assert_eq!(parsed.split_position(), 3);
                assert_eq!(parsed.split_percent(), 40);
                assert_eq!(parsed.pie2_size_percent(), 75);
                assert_eq!(parsed.gap_percent(), 100);
                assert_eq!(parsed.split_value(), 2.5);
                assert!(parsed.has_shadow());
                assert_eq!(parsed.to_payload(), bytes);
            }
        }
    }

    #[test]
    fn accepts_field_bounds_and_preserves_reserved_bits() {
        for (position, pie2, gap) in [(0, 5, 0), (32000, 200, 500)] {
            assert!(
                XlsBopPop::parse(&record(0x01, 0, 0, [position, 0, pie2, gap], 0.0, 0)).is_ok()
            );
        }
        // The 15 reserved flags bits MUST be ignored but round-trip verbatim.
        let bytes = record(0x02, 0, 0x0003, [0, 0, 5, 0], 0.0, 0xFFFE);
        let parsed = XlsBopPop::parse(&bytes).unwrap();
        assert_eq!(parsed.flags(), 0xFFFE);
        assert!(!parsed.has_shadow());
        assert_eq!(parsed.to_payload(), bytes);
    }

    #[test]
    fn rejects_malformed_records() {
        let bytes = record(0x01, 0x00, 0x0000, [0, 0, 5, 0], 0.0, 0);
        // Truncated and overlong payloads.
        assert!(XlsBopPop::parse(&bytes[..21]).is_err());
        assert!(XlsBopPop::parse(&[bytes.as_slice(), &[0]].concat()).is_err());
        // Undefined pst / split values.
        assert!(XlsBopPop::parse(&record(0x00, 0, 0, [0, 0, 5, 0], 0.0, 0)).is_err());
        assert!(XlsBopPop::parse(&record(0x03, 0, 0, [0, 0, 5, 0], 0.0, 0)).is_err());
        assert!(XlsBopPop::parse(&record(0x01, 0, 0x0004, [0, 0, 5, 0], 0.0, 0)).is_err());
        // Non-Boolean fAutoSplit.
        assert!(XlsBopPop::parse(&record(0x01, 0x02, 0, [0, 0, 5, 0], 0.0, 0)).is_err());
        // Out-of-range iSplitPos / pcPie2Size / pcGap.
        assert!(XlsBopPop::parse(&record(0x01, 0, 0, [-1, 0, 5, 0], 0.0, 0)).is_err());
        assert!(XlsBopPop::parse(&record(0x01, 0, 0, [32001, 0, 5, 0], 0.0, 0)).is_err());
        assert!(XlsBopPop::parse(&record(0x01, 0, 0, [0, 0, 4, 0], 0.0, 0)).is_err());
        assert!(XlsBopPop::parse(&record(0x01, 0, 0, [0, 0, 201, 0], 0.0, 0)).is_err());
        assert!(XlsBopPop::parse(&record(0x01, 0, 0, [0, 0, 5, -1], 0.0, 0)).is_err());
        assert!(XlsBopPop::parse(&record(0x01, 0, 0, [0, 0, 5, 501], 0.0, 0)).is_err());
    }
}
