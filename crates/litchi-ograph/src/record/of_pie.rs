//! Pie-of-pie and bar-of-pie settings (`BopPop`, MS-XLS 2.4.25 and
//! MS-OGRAPH 2.4.15). Ignored fields and reserved bits remain replayable.

use crate::{Error, Result, record};
use litchi_biff::{Encoder, Kind as RecordKind, RecordRef};

const LEN: usize = 22;

/// Shape of the secondary plot.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Kind {
    Pie,
    Bar,
}

/// Effective split rule. Inactive wire fields are deliberately not interpreted.
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum Split {
    Automatic,
    /// Number of trailing points assigned to the secondary plot (`0..=32000`).
    Position(u16),
    /// Values below this finite threshold belong to the secondary plot.
    Value(f64),
    /// Percentage threshold, encoded as a signed 16-bit integer.
    Percent(i16),
    /// Point selection is carried by a separate `BopPopCustom` record.
    Custom,
}

/// Validated Excel-compatible settings with lossless ignored-field retention.
/// Standalone Graph has narrower position/gap domains; use `validate_graph`
/// when interpreting this record in a Graph rather than Excel substream.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Options {
    bytes: [u8; LEN],
}

impl Options {
    pub const KIND: RecordKind = RecordKind::from_wire(0x1061);

    /// Creates canonical settings. Secondary size is `5..=200`, gap is `0..=500`.
    pub fn new(
        kind: Kind,
        split: Split,
        secondary_size: u16,
        gap: u16,
        shadow: bool,
    ) -> Result<Self> {
        let mut bytes = [0; LEN];
        bytes[0] = match kind {
            Kind::Pie => 1,
            Kind::Bar => 2,
        };
        match split {
            Split::Automatic => bytes[1] = 1,
            Split::Position(value) => bytes[4..6].copy_from_slice(&value.to_le_bytes()),
            Split::Value(value) => {
                bytes[2] = 1;
                bytes[12..20].copy_from_slice(&value.to_le_bytes());
            },
            Split::Percent(value) => {
                bytes[2] = 2;
                bytes[6..8].copy_from_slice(&value.to_le_bytes());
            },
            Split::Custom => bytes[2] = 3,
        }
        bytes[8..10].copy_from_slice(&secondary_size.to_le_bytes());
        bytes[10..12].copy_from_slice(&gap.to_le_bytes());
        bytes[20] = u8::from(shadow);
        Self::from_payload(&bytes)
    }

    /// Decodes a framed `BopPop` record.
    pub fn parse(input: RecordRef<'_>) -> Result<Self> {
        Self::from_payload(record::payload(input, Self::KIND, LEN)?)
    }

    /// Decodes the fixed payload, checking only active split fields.
    pub fn from_payload(payload: &[u8]) -> Result<Self> {
        let payload = record::payload_bytes(Self::KIND, payload, LEN)?;
        let mut bytes = [0; LEN];
        bytes.copy_from_slice(payload);
        let result = Self { bytes };
        if !matches!(bytes[0], 1 | 2) {
            return Err(invalid("pst", u64::from(bytes[0])));
        }
        if bytes[1] > 1 {
            return Err(invalid("fAutoSplit", u64::from(bytes[1])));
        }
        if !(5..=200).contains(&result.word(8)) {
            return Err(invalid("pcPie2Size", u64::from(result.word(8))));
        }
        if result.word(10) > 500 {
            return Err(invalid("pcGap", u64::from(result.word(10))));
        }
        if bytes[1] == 0 {
            match result.word(2) {
                0 if result.word(4) > 32000 => {
                    return Err(invalid("iSplitPos", u64::from(result.word(4))));
                },
                1 if !result.threshold().is_finite()
                    || result.threshold().is_subnormal()
                    || result.threshold().to_bits() == (-0.0_f64).to_bits() =>
                {
                    return Err(invalid("numSplitValue", result.threshold().to_bits()));
                },
                0..=3 => {},
                value => return Err(invalid("split", u64::from(value))),
            }
        }
        Ok(result)
    }

    /// Checks the narrower MS-OGRAPH 2.4.15 domains after generic decoding.
    pub fn validate_graph(self) -> Result<()> {
        if self.split() == Split::Position(0) {
            return Err(invalid("iSplitPos", 0));
        }
        if self.gap() > 200 {
            return Err(invalid("pcGap", u64::from(self.gap())));
        }
        Ok(())
    }

    #[must_use]
    pub const fn kind(self) -> Kind {
        if self.bytes[0] == 1 {
            Kind::Pie
        } else {
            Kind::Bar
        }
    }

    #[must_use]
    pub fn split(self) -> Split {
        if self.bytes[1] == 1 {
            return Split::Automatic;
        }
        match self.word(2) {
            0 => Split::Position(self.word(4)),
            1 => Split::Value(self.threshold()),
            2 => Split::Percent(i16::from_le_bytes([self.bytes[6], self.bytes[7]])),
            _ => Split::Custom, // from_payload proves the remaining active value is three.
        }
    }

    #[must_use]
    pub fn secondary_size(self) -> u16 {
        self.word(8)
    }

    #[must_use]
    pub fn gap(self) -> u16 {
        self.word(10)
    }

    #[must_use]
    pub const fn has_shadow(self) -> bool {
        self.bytes[20] & 1 != 0
    }

    /// Original payload, including ignored fields. Newly constructed values
    /// have zeroed inactive fields and reserved bits.
    #[must_use]
    pub const fn payload(self) -> [u8; LEN] {
        self.bytes
    }

    pub fn write(self, out: &mut Encoder) -> Result<()> {
        out.push(Self::KIND, &self.bytes)?;
        Ok(())
    }

    fn word(self, offset: usize) -> u16 {
        u16::from_le_bytes([self.bytes[offset], self.bytes[offset + 1]])
    }

    fn threshold(self) -> f64 {
        f64::from_le_bytes([
            self.bytes[12],
            self.bytes[13],
            self.bytes[14],
            self.bytes[15],
            self.bytes[16],
            self.bytes[17],
            self.bytes[18],
            self.bytes[19],
        ])
    }
}

const fn invalid(field: &'static str, value: u64) -> Error {
    Error::InvalidRecordValue {
        kind: Options::KIND.get(),
        field,
        value,
    }
}

#[cfg(test)]
mod tests {
    #![allow(clippy::unwrap_used, reason = "test assertions panic by design")]
    use super::*;
    use litchi_biff::{Limits, Records};

    #[test]
    fn split_rules_and_both_subtypes_round_trip() {
        for kind in [Kind::Pie, Kind::Bar] {
            for split in [
                Split::Automatic,
                Split::Position(0),
                Split::Position(32000),
                Split::Value(-1.25),
                Split::Percent(25),
                Split::Percent(-1),
                Split::Custom,
            ] {
                let value = Options::new(kind, split, 75, 150, true).unwrap();
                assert_eq!(value.kind(), kind);
                assert_eq!(value.split(), split);
                assert_eq!(value.secondary_size(), 75);
                assert_eq!(value.gap(), 150);
                assert!(value.has_shadow());
                let mut encoder = Encoder::new(Limits::default()).unwrap();
                value.write(&mut encoder).unwrap();
                let bytes = encoder.finish();
                assert_eq!(bytes.len(), LEN + 4);
                let record = Records::new(&bytes, Limits::default())
                    .unwrap()
                    .next()
                    .unwrap()
                    .unwrap();
                assert_eq!(Options::parse(record).unwrap(), value);
            }
        }
    }

    #[test]
    fn automatic_and_inactive_split_fields_are_ignored_and_preserved() {
        for split in [
            Split::Automatic,
            Split::Position(2),
            Split::Percent(20),
            Split::Custom,
        ] {
            let mut bytes = Options::new(Kind::Pie, split, 75, 100, false)
                .unwrap()
                .payload();
            bytes[12..20].copy_from_slice(&f64::NAN.to_le_bytes());
            bytes[20..22].copy_from_slice(&0xfffe_u16.to_le_bytes());
            if split == Split::Automatic {
                bytes[2..8].fill(0xff);
            }
            let parsed = Options::from_payload(&bytes).unwrap();
            assert_eq!(parsed.split(), split);
            assert!(!parsed.has_shadow());
            assert_eq!(parsed.payload(), bytes);
        }
    }

    #[test]
    fn malformed_active_fields_and_lengths_are_rejected() {
        let base = Options::new(Kind::Pie, Split::Position(2), 75, 100, false)
            .unwrap()
            .payload();
        for len in [0, 2, 20, 21, 23] {
            assert!(matches!(
                Options::from_payload(&vec![0; len]),
                Err(Error::InvalidRecordLength { kind: 0x1061, .. })
            ));
        }
        for (offset, bytes, field) in [
            (0, [0, 0], "pst"),
            (0, [3, 0], "pst"),
            (1, [2, 0], "fAutoSplit"),
            (2, [4, 0], "split"),
            (4, (-1_i16).to_le_bytes(), "iSplitPos"),
            (4, 32001_u16.to_le_bytes(), "iSplitPos"),
            (8, 4_u16.to_le_bytes(), "pcPie2Size"),
            (8, 201_u16.to_le_bytes(), "pcPie2Size"),
            (10, (-1_i16).to_le_bytes(), "pcGap"),
            (10, 501_u16.to_le_bytes(), "pcGap"),
        ] {
            let mut wire = base;
            wire[offset..offset + 2].copy_from_slice(&bytes);
            assert!(
                matches!(Options::from_payload(&wire), Err(Error::InvalidRecordValue { field: actual, .. }) if actual == field)
            );
        }
        for value in [
            f64::NAN,
            f64::INFINITY,
            f64::NEG_INFINITY,
            -0.0,
            f64::from_bits(1),
        ] {
            assert!(matches!(
                Options::new(Kind::Pie, Split::Value(value), 75, 100, false),
                Err(Error::InvalidRecordValue {
                    field: "numSplitValue",
                    ..
                })
            ));
        }
    }

    #[test]
    fn graph_and_excel_position_and_gap_domains_stay_distinct() {
        for (split, gap, graph_valid) in [
            (Split::Automatic, 200, true),
            (Split::Position(1), 200, true),
            (Split::Position(0), 200, false),
            (Split::Automatic, 201, false),
            (Split::Position(32000), 500, false),
        ] {
            let value = Options::new(Kind::Pie, split, 75, gap, false).unwrap();
            assert_eq!(value.validate_graph().is_ok(), graph_valid);
        }
    }
}
