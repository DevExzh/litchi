//! BIFF8 `Pls` record (0x004D, MS-XLS 2.4.199) of the worksheet substream
//! (MS-XLS 2.1): printer settings and printer driver information as a
//! `DEVMODE` structure that may span `Continue` records (0x003C).
//!
//! Everything in this module is INERT: the `DEVMODE` bytes are stored
//! verbatim and never parsed or sent to a printer driver. The validation
//! and chunking rules mirror the crate's BIFF8 writer
//! (`writer::biff::worksheet::write_page_settings`), so writer output always
//! parses.
//!
//! # References
//!
//! - MS-XLS 2.4.199 (Pls), 2.4.63 (Continue)

use super::{XlsError, XlsResult};

/// Byte length of the `Pls` `reserved` field.
const RESERVED_LEN: usize = 2;
/// Largest BIFF8 record payload.
const MAX_RECORD_PAYLOAD: usize = 8_224;

/// Typed `Pls` record content (MS-XLS 2.4.199): printer settings and printer
/// driver information.
///
/// The `reserved` field (MUST be zero, and MUST be ignored) is preserved
/// verbatim so the record round-trips unchanged.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsPrinterDriverData {
    /// Raw `reserved` field, preserved verbatim.
    reserved: u16,
    /// Opaque `DEVMODE` bytes, reassembled across `Continue` records.
    devmode: Vec<u8>,
}

impl XlsPrinterDriverData {
    /// Parse a `Pls` record payload plus the payloads of the `Continue`
    /// records that follow it.
    ///
    /// `Continue` payloads carry raw `DEVMODE` sections without any header,
    /// so they are appended verbatim; their record-type identity is
    /// established by the caller's record iteration.
    pub fn parse(data: &[u8], continues: &[Vec<u8>]) -> XlsResult<Self> {
        if data.len() < RESERVED_LEN {
            return Err(XlsError::InvalidLength {
                expected: RESERVED_LEN,
                found: data.len(),
            });
        }
        let mut devmode = data[RESERVED_LEN..].to_vec();
        for continuation in continues {
            devmode.extend_from_slice(continuation);
        }
        Ok(Self {
            reserved: u16::from_le_bytes([data[0], data[1]]),
            devmode,
        })
    }

    /// Serialize as a sequence of complete record payloads: the `Pls` record
    /// followed by `Continue` records when the `DEVMODE` bytes exceed one
    /// record.
    pub fn to_record_payloads(&self) -> Vec<Vec<u8>> {
        let first_chunk = MAX_RECORD_PAYLOAD - RESERVED_LEN;
        let mut chunks = self.devmode.chunks(first_chunk);
        let mut first = Vec::with_capacity(RESERVED_LEN + first_chunk);
        first.extend_from_slice(&self.reserved.to_le_bytes());
        first.extend_from_slice(chunks.next().unwrap_or(&[]));
        let mut records = vec![first];
        records.extend(chunks.map(<[u8]>::to_vec));
        records
    }

    /// Raw `reserved` field value.
    pub fn reserved(&self) -> u16 {
        self.reserved
    }

    /// The opaque `DEVMODE` printer driver data.
    pub fn driver_data(&self) -> &[u8] {
        &self.devmode
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn single_record_round_trip() {
        let mut data = Vec::new();
        data.extend_from_slice(&[0; 2]);
        data.extend_from_slice(b"devmode-bytes");
        let parsed = XlsPrinterDriverData::parse(&data, &[]).unwrap();
        assert_eq!(parsed.reserved(), 0);
        assert_eq!(parsed.driver_data(), b"devmode-bytes");
        assert_eq!(parsed.to_record_payloads(), vec![data]);
    }

    #[test]
    fn reassembles_continuations_and_round_trips() {
        let devmode = vec![0x5Au8; 20_000];
        let mut first = Vec::new();
        first.extend_from_slice(&[0; 2]);
        first.extend_from_slice(&devmode[..6_000]);
        let continues = vec![devmode[6_000..14_000].to_vec(), devmode[14_000..].to_vec()];
        let parsed = XlsPrinterDriverData::parse(&first, &continues).unwrap();
        assert_eq!(parsed.driver_data(), devmode.as_slice());

        let payloads = parsed.to_record_payloads();
        assert!(payloads.len() > 1);
        for payload in &payloads {
            assert!(payload.len() <= MAX_RECORD_PAYLOAD);
        }
        let reparsed = XlsPrinterDriverData::parse(&payloads[0], &payloads[1..]).unwrap();
        assert_eq!(reparsed, parsed);
        assert_eq!(reparsed.driver_data(), devmode.as_slice());
    }

    #[test]
    fn preserves_reserved_and_handles_empty_devmode() {
        // reserved MUST be zero and MUST be ignored; it round-trips verbatim.
        let mut data = Vec::new();
        data.extend_from_slice(&0x7F7Fu16.to_le_bytes());
        let parsed = XlsPrinterDriverData::parse(&data, &[]).unwrap();
        assert_eq!(parsed.reserved(), 0x7F7F);
        assert!(parsed.driver_data().is_empty());
        assert_eq!(parsed.to_record_payloads(), vec![data]);
        // Truncated reserved field.
        assert!(XlsPrinterDriverData::parse(&[0x00], &[]).is_err());
    }

    #[test]
    fn empty_devmode_emits_single_record() {
        let parsed = XlsPrinterDriverData::parse(&[0; 2], &[]).unwrap();
        let payloads = parsed.to_record_payloads();
        assert_eq!(payloads.len(), 1);
        assert_eq!(payloads[0], vec![0; 2]);
    }
}
