//! BIFF8 `BopPopCustom` record (0x1067, MS-XLS 2.4.26) of the Chart Sheet
//! substream (MS-XLS 2.1): the custom split of data points between the
//! primary pie and the secondary bar/pie of a bar of pie or pie of pie
//! chart group.
//!
//! Everything in this module is INERT: the split bits are stored verbatim
//! and no chart group is rendered. The record MUST follow a `BopPop` record
//! with `split` Custom (0x0003) (MS-XLS 2.4.26); that cross-record
//! constraint is documented here, not enforced by the record reader.
//!
//! # References
//!
//! - MS-XLS 2.4.26 (BopPopCustom)

use super::{XlsError, XlsResult};

/// Record type of the `BopPopCustom` record (MS-XLS 2.4.26).
pub(crate) const BOP_POP_CUSTOM_RECORD_TYPE: u16 = 0x1067;

/// Maximum `cxi` value (exclusive): MUST be less than 32000 (MS-XLS 2.4.26).
const MAX_CXI_EXCLUSIVE: u16 = 32000;

fn invalid(message: impl Into<String>) -> XlsError {
    XlsError::InvalidRecord {
        record_type: BOP_POP_CUSTOM_RECORD_TYPE,
        message: message.into(),
    }
}

/// Checked reader for the fixed-width fields in a `BopPopCustom` payload.
///
/// The public parser currently rejects any payload whose length is not exact,
/// but keeping offset arithmetic and field extraction checked here prevents a
/// future layout change from reintroducing panic paths.
struct BopPopCustomReader<'a> {
    data: &'a [u8],
    offset: usize,
}

impl<'a> BopPopCustomReader<'a> {
    fn new(data: &'a [u8]) -> Self {
        Self { data, offset: 0 }
    }

    fn read_bytes<const N: usize>(&mut self) -> XlsResult<[u8; N]> {
        let end = self
            .offset
            .checked_add(N)
            .ok_or_else(|| invalid("BopPopCustom field offset overflows usize"))?;
        let bytes = self
            .data
            .get(self.offset..end)
            .ok_or(XlsError::InvalidLength {
                expected: end,
                found: self.data.len(),
            })?;
        let value: [u8; N] = bytes.try_into().map_err(|_| XlsError::InvalidLength {
            expected: N,
            found: bytes.len(),
        })?;
        self.offset = end;
        Ok(value)
    }

    fn read_u16(&mut self) -> XlsResult<u16> {
        Ok(u16::from_le_bytes(self.read_bytes()?))
    }

    fn read_remaining(&mut self) -> XlsResult<&'a [u8]> {
        let bytes = self
            .data
            .get(self.offset..)
            .ok_or_else(|| invalid("BopPopCustom field offset is outside payload"))?;
        self.offset = self.data.len();
        Ok(bytes)
    }
}

/// Typed `BopPopCustom` record content (MS-XLS 2.4.26): which data points of
/// the series are contained in the secondary bar/pie instead of the primary
/// pie.
///
/// The `cxi` field specifies the data point count plus one, and the
/// `rggrbit` bit sequence holds one bit per data point (padding in the most
/// significant bits of the first byte) plus a final least significant bit
/// marking that the secondary bar/pie is empty. The raw bytes are preserved
/// verbatim; use [`Self::is_secondary`] and [`Self::secondary_is_empty`] to
/// decode them.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsBopPopCustom {
    /// Data point count plus one (`cxi`).
    data_point_count_plus_one: u16,
    /// The raw `rggrbit` bytes, preserved verbatim.
    bits: Vec<u8>,
}

impl XlsBopPopCustom {
    /// Parse a `BopPopCustom` record payload.
    pub fn parse(data: &[u8]) -> XlsResult<Self> {
        let mut reader = BopPopCustomReader::new(data);
        let cxi = reader.read_u16()?;
        if cxi >= MAX_CXI_EXCLUSIVE {
            return Err(invalid(format!(
                "BopPopCustom cxi {cxi} is not less than {MAX_CXI_EXCLUSIVE}"
            )));
        }
        // MS-XLS 2.4.26: size of rggrbit in bytes = 1 + floor(cxi / 8).
        let expected = 2 + 1 + usize::from(cxi) / 8;
        if data.len() != expected {
            return Err(XlsError::InvalidLength {
                expected,
                found: data.len(),
            });
        }
        let bits = reader.read_remaining()?.to_vec();
        // MS-XLS 2.4.26: when the final least significant bit is 1 (the
        // secondary bar/pie contains no data points), every other bit MUST
        // be 0.
        let parsed = Self {
            data_point_count_plus_one: cxi,
            bits,
        };
        // MS-XLS 2.4.26: when the final least significant bit is 1 (the
        // secondary bar/pie contains no data points), every other bit MUST
        // be 0.
        if parsed.secondary_is_empty() {
            let mut bytes = parsed.bits.iter().copied();
            let Some(last) = bytes.next_back() else {
                return Err(invalid(
                    "BopPopCustom rggrbit is missing the empty-secondary bit",
                ));
            };
            if bytes.any(|byte| byte != 0) || last & !0x01 != 0 {
                return Err(invalid(
                    "BopPopCustom rggrbit sets data point bits while the empty-secondary bit is 1",
                ));
            }
        }
        Ok(parsed)
    }

    /// Serialize back to a complete `BopPopCustom` record payload.
    pub fn to_payload(&self) -> Vec<u8> {
        let mut payload = Vec::with_capacity(2 + self.bits.len());
        payload.extend_from_slice(&self.data_point_count_plus_one.to_le_bytes());
        payload.extend_from_slice(&self.bits);
        payload
    }

    /// Data point count plus one (`cxi`).
    pub fn data_point_count_plus_one(&self) -> u16 {
        self.data_point_count_plus_one
    }

    /// The raw `rggrbit` bytes, preserved verbatim.
    pub fn bits(&self) -> &[u8] {
        &self.bits
    }

    /// Whether the data point at the zero-based `index` is contained in the
    /// secondary bar/pie instead of the primary pie.
    ///
    /// Returns `None` when `index` is out of range (the series has
    /// `cxi - 1` data points).
    pub fn is_secondary(&self, index: u16) -> Option<bool> {
        if index >= self.data_point_count_plus_one.saturating_sub(1) {
            return None;
        }
        // MS-XLS 2.4.26: padding occupies the most significant bits of the
        // first byte; bit positions count from the MSB of byte 0.
        let total_bits = self.bits.len() * 8;
        let padding = total_bits - usize::from(self.data_point_count_plus_one);
        let position = padding + usize::from(index);
        self.bits
            .get(position / 8)
            .map(|byte| *byte & (0x80 >> (position % 8)) != 0)
    }

    /// Whether the final least significant bit is set, marking that the
    /// secondary bar/pie does not contain data points.
    pub fn secondary_is_empty(&self) -> bool {
        self.bits.last().is_some_and(|last| last & 0x01 != 0)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn record(cxi: u16, bits: &[u8]) -> Vec<u8> {
        let mut data = cxi.to_le_bytes().to_vec();
        data.extend_from_slice(bits);
        data
    }

    #[test]
    fn round_trip_and_bit_positions() {
        // cxi = 5 (4 data points): rggrbit is 1 byte; the 4 data point bits
        // follow the 3 padding bits, and the LSB is the empty marker.
        // Bits: data point 0 secondary (bit 4), data point 3 secondary (bit 1).
        let bytes = record(5, &[0b0001_0010]);
        let parsed = XlsBopPopCustom::parse(&bytes).unwrap();
        assert_eq!(parsed.data_point_count_plus_one(), 5);
        assert_eq!(parsed.is_secondary(0), Some(true));
        assert_eq!(parsed.is_secondary(1), Some(false));
        assert_eq!(parsed.is_secondary(2), Some(false));
        assert_eq!(parsed.is_secondary(3), Some(true));
        assert_eq!(parsed.is_secondary(4), None);
        assert!(!parsed.secondary_is_empty());
        assert_eq!(parsed.to_payload(), bytes);
    }

    #[test]
    fn multi_byte_round_trip() {
        // cxi = 12 (11 data points): rggrbit is 2 bytes, 4 padding bits.
        let bytes = record(12, &[0b0010_1010, 0b1010_1100]);
        let parsed = XlsBopPopCustom::parse(&bytes).unwrap();
        assert_eq!(parsed.is_secondary(0), Some(true));
        assert_eq!(parsed.is_secondary(1), Some(false));
        assert_eq!(parsed.is_secondary(2), Some(true));
        assert_eq!(parsed.is_secondary(5), Some(false));
        assert_eq!(parsed.is_secondary(8), Some(true));
        assert_eq!(parsed.is_secondary(10), Some(false));
        assert_eq!(parsed.is_secondary(11), None);
        assert!(!parsed.secondary_is_empty());
        assert_eq!(parsed.to_payload(), bytes);
    }

    #[test]
    fn empty_secondary_marker() {
        // The empty-secondary bit set with all data bits clear is legal.
        let bytes = record(5, &[0b0000_0001]);
        let parsed = XlsBopPopCustom::parse(&bytes).unwrap();
        assert!(parsed.secondary_is_empty());
        assert_eq!(parsed.is_secondary(0), Some(false));
        assert_eq!(parsed.to_payload(), bytes);
        // The empty-secondary bit with any data bit set is invalid.
        assert!(XlsBopPopCustom::parse(&record(5, &[0b0001_0001])).is_err());
    }

    #[test]
    fn rejects_malformed_records() {
        // Truncated cxi.
        assert!(XlsBopPopCustom::parse(&[0x05]).is_err());
        // cxi at or above 32000.
        let mut too_big = 32000u16.to_le_bytes().to_vec();
        too_big.extend_from_slice(&[0; 4001]);
        assert!(XlsBopPopCustom::parse(&too_big).is_err());
        // rggrbit size mismatch.
        assert!(XlsBopPopCustom::parse(&record(5, &[0, 0])).is_err());
        assert!(XlsBopPopCustom::parse(&record(9, &[0])).is_err());
    }

    #[test]
    fn rejects_every_truncated_payload() {
        let bytes = record(5, &[0b0001_0010]);
        for length in 0..bytes.len() {
            assert!(
                XlsBopPopCustom::parse(&bytes[..length]).is_err(),
                "length {length}"
            );
        }
    }

    #[test]
    fn reader_rejects_offset_overflow() {
        let mut reader = BopPopCustomReader {
            data: &[],
            offset: usize::MAX,
        };

        assert!(reader.read_bytes::<1>().is_err());
    }
}
