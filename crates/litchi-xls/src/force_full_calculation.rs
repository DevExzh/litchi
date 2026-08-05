//! BIFF8 `ForceFullCalculation` record (0x08A3, MS-XLS 2.4.125) of the
//! Globals substream (MS-XLS 2.1): the forced calculation mode of the
//! workbook.
//!
//! Everything in this module is INERT: the flag is stored verbatim and no
//! calculation is triggered.
//!
//! # References
//!
//! - MS-XLS 2.4.125 (ForceFullCalculation), 2.5.14 (Boolean), 2.5.134
//!   (FrtFlags), 2.5.135 (FrtHeader)

use super::{Error, Result};

/// Record type of the `ForceFullCalculation` record (MS-XLS 2.4.125); also
/// the required `frtHeader.rt` value.
pub(crate) const FORCE_FULL_CALCULATION_RECORD_TYPE: u16 = 0x08A3;

/// Size in bytes of an `FrtHeader` (MS-XLS 2.5.135).
const FRT_HEADER_LEN: usize = 12;
/// `FrtFlags` bits that MUST be zero in an `FrtHeader` (MS-XLS 2.5.135):
/// `fFrtRef` and `fFrtAlert`.
const FRT_FLAGS_FORBIDDEN: u16 = 0x0003;
/// Byte length of a `ForceFullCalculation` record payload: `FrtHeader` (12) +
/// `fNoDeps` (4).
const PAYLOAD_LEN: usize = 16;

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidRecord {
        record_type: FORCE_FULL_CALCULATION_RECORD_TYPE,
        message: message.into(),
    }
}

/// Typed `ForceFullCalculation` record content (MS-XLS 2.4.125): whether all
/// cell formulas fully calculate on every calculation.
///
/// The `frtHeader` reserved bytes (MUST be ignored) are preserved verbatim so
/// the record round-trips unchanged.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ForceFullCalculation {
    /// Raw `frtHeader.grbitFrt` bitfield (`fFrtRef`/`fFrtAlert` are zero).
    frt_flags: u16,
    /// `frtHeader.reserved` bytes, preserved verbatim.
    frt_reserved: [u8; 8],
    /// Whether dependencies are ignored and all formulas fully calculate
    /// (`fNoDeps`).
    force_full: bool,
}

impl ForceFullCalculation {
    /// Parse a `ForceFullCalculation` record payload.
    pub fn parse(data: &[u8]) -> Result<Self> {
        if data.len() != PAYLOAD_LEN {
            return Err(Error::InvalidLength {
                expected: PAYLOAD_LEN,
                found: data.len(),
            });
        }
        if u16::from_le_bytes([data[0], data[1]]) != FORCE_FULL_CALCULATION_RECORD_TYPE {
            return Err(invalid("ForceFullCalculation FrtHeader.rt mismatch"));
        }
        let frt_flags = u16::from_le_bytes([data[2], data[3]]);
        if frt_flags & FRT_FLAGS_FORBIDDEN != 0 {
            return Err(invalid(format!(
                "ForceFullCalculation FrtHeader.grbitFrt {frt_flags:#06X} sets fFrtRef or fFrtAlert"
            )));
        }
        // Boolean (MS-XLS 2.5.14): only 0 and 1 are legal.
        let force_full = match u32::from_le_bytes(data[12..16].try_into().expect("length checked"))
        {
            0 => false,
            1 => true,
            other => {
                return Err(invalid(format!(
                    "ForceFullCalculation fNoDeps {other:#X} is not a Boolean"
                )));
            },
        };
        Ok(Self {
            frt_flags,
            frt_reserved: data[4..FRT_HEADER_LEN].try_into().expect("length checked"),
            force_full,
        })
    }

    /// Serialize back to a complete `ForceFullCalculation` record payload.
    pub fn to_payload(&self) -> Vec<u8> {
        let mut payload = Vec::with_capacity(PAYLOAD_LEN);
        payload.extend_from_slice(&FORCE_FULL_CALCULATION_RECORD_TYPE.to_le_bytes());
        payload.extend_from_slice(&self.frt_flags.to_le_bytes());
        payload.extend_from_slice(&self.frt_reserved);
        payload.extend_from_slice(&u32::from(self.force_full).to_le_bytes());
        payload
    }

    /// Whether dependencies are ignored and all cell formulas fully calculate
    /// every time a calculation is triggered (`fNoDeps`).
    pub fn force_full(&self) -> bool {
        self.force_full
    }

    /// Raw `frtHeader.grbitFrt` bitfield (`fFrtRef`/`fFrtAlert` are zero).
    pub fn frt_flags(&self) -> u16 {
        self.frt_flags
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn record(frt_flags: u16, value: u32) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&FORCE_FULL_CALCULATION_RECORD_TYPE.to_le_bytes());
        data.extend_from_slice(&frt_flags.to_le_bytes());
        data.extend_from_slice(&[0; 8]);
        data.extend_from_slice(&value.to_le_bytes());
        data
    }

    #[test]
    fn round_trip_both_modes() {
        for (value, expected) in [(0, false), (1, true)] {
            let bytes = record(0, value);
            let parsed = ForceFullCalculation::parse(&bytes).unwrap();
            assert_eq!(parsed.force_full(), expected);
            assert_eq!(parsed.frt_flags(), 0);
            assert_eq!(parsed.to_payload(), bytes);
        }
    }

    #[test]
    fn preserves_reserved_header_bytes() {
        // The 8 reserved FrtHeader bytes and the 14 reserved grbitFrt bits
        // MUST be ignored but round-trip verbatim.
        let mut bytes = record(0xFFFC, 1);
        bytes[4..FRT_HEADER_LEN].copy_from_slice(&[1, 2, 3, 4, 5, 6, 7, 8]);
        let parsed = ForceFullCalculation::parse(&bytes).unwrap();
        assert_eq!(parsed.frt_flags(), 0xFFFC);
        assert_eq!(parsed.to_payload(), bytes);
    }

    #[test]
    fn rejects_malformed_records() {
        let bytes = record(0, 1);
        // Truncated and overlong payloads.
        assert!(ForceFullCalculation::parse(&bytes[..15]).is_err());
        assert!(ForceFullCalculation::parse(&[bytes.as_slice(), &[0]].concat()).is_err());
        // Wrong FrtHeader.rt.
        let mut wrong_rt = bytes.clone();
        wrong_rt[0..2].copy_from_slice(&0x08A4u16.to_le_bytes());
        assert!(ForceFullCalculation::parse(&wrong_rt).is_err());
        // fFrtRef / fFrtAlert set.
        assert!(ForceFullCalculation::parse(&record(0x0001, 1)).is_err());
        assert!(ForceFullCalculation::parse(&record(0x0002, 1)).is_err());
        // Non-Boolean fNoDeps.
        assert!(ForceFullCalculation::parse(&record(0, 2)).is_err());
        assert!(ForceFullCalculation::parse(&record(0, 0xFFFF_FFFF)).is_err());
    }
}
