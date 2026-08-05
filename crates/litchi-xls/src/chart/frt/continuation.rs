//! BIFF8 chart `CrtMlFrt` record (MS-XLS 2.4.70): additional chart-element
//! properties stored as an `XmlTkChain` structure chain that may span
//! `CrtMlFrtContinue` records (MS-XLS 2.4.71).
//!
//! Everything in this module is INERT: the `XmlTkChain` bytes are stored
//! verbatim and never tokenized, interpreted, or applied to a chart. An
//! application can ignore this record without loss of functionality (MS-XLS
//! 2.4.70).
//!
//! # References
//!
//! - MS-XLS 2.4.70 (CrtMlFrt), 2.4.71 (CrtMlFrtContinue), 2.5.134 (FrtFlags),
//!   2.5.135 (FrtHeader)

use crate::{XlsError, XlsResult};

/// Record type of the `CrtMlFrt` record (MS-XLS 2.4.70); also the required
/// `frtHeader.rt` value. (The neighboring 0x089D is `CrtLayout12`,
/// MS-XLS 2.4.66.)
pub(crate) const CRT_ML_FRT_RECORD_TYPE: u16 = 0x089E;
/// Record type of the `CrtMlFrtContinue` record (MS-XLS 2.4.71).
pub(crate) const CRT_ML_FRT_CONTINUE_RECORD_TYPE: u16 = 0x089F;

/// Size in bytes of an `FrtHeader` (MS-XLS 2.5.135).
const FRT_HEADER_LEN: usize = 12;
/// `FrtFlags` bits that MUST be zero in a `CrtMlFrt`/`CrtMlFrtContinue`
/// header: `fFrtRef` and `fFrtAlert` (MS-XLS 2.5.135).
const FRT_FLAGS_FORBIDDEN: u16 = 0x0003;
/// Largest BIFF8 record payload.
const MAX_RECORD_PAYLOAD: usize = 8_224;
/// Byte length of the `cb` field and of the trailing `unused` field.
const FIELD_LEN: usize = 4;
/// Maximum `cb` value: the largest legal `XmlTkChain` size (MS-XLS 2.4.70).
const MAX_CHAIN_LEN: u64 = 0x7FFF_FFEB;

fn invalid(message: impl Into<String>) -> XlsError {
    XlsError::InvalidRecord {
        record_type: CRT_ML_FRT_RECORD_TYPE,
        message: message.into(),
    }
}

/// Checked cursor for the fixed-width fields and opaque payload sections in
/// `CrtMlFrt` records.
struct CrtMlFrtReader<'a> {
    data: &'a [u8],
    offset: usize,
}

impl<'a> CrtMlFrtReader<'a> {
    fn new(data: &'a [u8]) -> Self {
        Self { data, offset: 0 }
    }

    fn read_bytes<const N: usize>(&mut self) -> XlsResult<[u8; N]> {
        let bytes = self.read_slice(N)?;
        bytes.try_into().map_err(|_| XlsError::InvalidLength {
            expected: N,
            found: bytes.len(),
        })
    }

    fn read_u16(&mut self) -> XlsResult<u16> {
        Ok(u16::from_le_bytes(self.read_bytes()?))
    }

    fn read_u32(&mut self) -> XlsResult<u32> {
        Ok(u32::from_le_bytes(self.read_bytes()?))
    }

    fn read_slice(&mut self, len: usize) -> XlsResult<&'a [u8]> {
        let end = self
            .offset
            .checked_add(len)
            .ok_or_else(|| invalid("CrtMlFrt field offset overflows usize"))?;
        let bytes = self
            .data
            .get(self.offset..end)
            .ok_or(XlsError::InvalidLength {
                expected: end,
                found: self.data.len(),
            })?;
        self.offset = end;
        Ok(bytes)
    }

    fn read_remaining(&mut self) -> XlsResult<&'a [u8]> {
        self.read_slice(self.remaining())
    }

    fn remaining(&self) -> usize {
        self.data.len().saturating_sub(self.offset)
    }
}

/// Validate an `FrtHeader` (MS-XLS 2.5.135): the `rt` field and the
/// `fFrtRef`/`fFrtAlert` bits that MUST be zero, returning the raw flags and
/// reserved bytes for lossless round-tripping.
fn validate_frt_header(
    reader: &mut CrtMlFrtReader<'_>,
    expected_rt: u16,
    context: &str,
) -> XlsResult<(u16, [u8; 8])> {
    if reader.read_u16()? != expected_rt {
        return Err(invalid(format!("{context} FrtHeader.rt mismatch")));
    }
    let flags = reader.read_u16()?;
    if flags & FRT_FLAGS_FORBIDDEN != 0 {
        return Err(invalid(format!(
            "{context} FrtHeader.grbitFrt {flags:#06X} sets fFrtRef or fFrtAlert"
        )));
    }
    Ok((flags, reader.read_bytes()?))
}

fn validate_record_payload_len(data: &[u8], minimum: usize, context: &str) -> XlsResult<()> {
    if data.len() < minimum {
        return Err(XlsError::InvalidLength {
            expected: minimum,
            found: data.len(),
        });
    }
    if data.len() > MAX_RECORD_PAYLOAD {
        return Err(invalid(format!(
            "{context} payload length {} exceeds BIFF8 maximum {MAX_RECORD_PAYLOAD}",
            data.len()
        )));
    }
    Ok(())
}

fn append_chain_bytes(chain: &mut Vec<u8>, bytes: &[u8]) -> XlsResult<()> {
    chain
        .try_reserve(bytes.len())
        .map_err(|_| XlsError::Allocation("reassembling CrtMlFrt XmlTkChain"))?;
    chain.extend_from_slice(bytes);
    Ok(())
}

/// Typed `CrtMlFrt` record content (MS-XLS 2.4.70): additional chart-element
/// properties as an opaque `XmlTkChain` byte chain.
///
/// The header bitfield, the 8 reserved header bytes, and the 4 trailing
/// `unused` bytes are preserved verbatim so the record round-trips unchanged.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CrtMlFrt {
    /// Raw `frtHeader.grbitFrt` bitfield. `fFrtRef` and `fFrtAlert` are
    /// guaranteed zero; the undefined reserved bits are preserved.
    flags: u16,
    /// `frtHeader.reserved` bytes, preserved verbatim (MUST be ignored).
    reserved: [u8; 8],
    /// Opaque `XmlTkChain` bytes, reassembled across `CrtMlFrtContinue`
    /// records.
    chain: Vec<u8>,
    /// Trailing `unused` field, preserved verbatim (MUST be ignored).
    unused: [u8; 4],
}

impl CrtMlFrt {
    /// Parse a `CrtMlFrt` record payload plus the payloads of the
    /// `CrtMlFrtContinue` records that follow it.
    pub fn parse(data: &[u8], continues: &[Vec<u8>]) -> XlsResult<Self> {
        const MIN_LEN: usize = FRT_HEADER_LEN + FIELD_LEN + FIELD_LEN;
        validate_record_payload_len(data, MIN_LEN, "CrtMlFrt")?;
        let mut reader = CrtMlFrtReader::new(data);
        let (flags, reserved) =
            validate_frt_header(&mut reader, CRT_ML_FRT_RECORD_TYPE, "CrtMlFrt")?;
        let declared = u64::from(reader.read_u32()?);
        if declared > MAX_CHAIN_LEN {
            return Err(invalid(format!(
                "CrtMlFrt cb {declared:#X} exceeds {MAX_CHAIN_LEN:#X}"
            )));
        }

        let first_chain_len =
            reader
                .remaining()
                .checked_sub(FIELD_LEN)
                .ok_or(XlsError::InvalidLength {
                    expected: MIN_LEN,
                    found: data.len(),
                })?;
        let first_chain = reader.read_slice(first_chain_len)?;
        let unused = reader.read_bytes::<FIELD_LEN>()?;

        let declared = usize::try_from(declared)
            .map_err(|_| XlsError::Allocation("sizing CrtMlFrt XmlTkChain"))?;
        let mut chain_len = first_chain.len();
        for continuation in continues {
            validate_record_payload_len(continuation, FRT_HEADER_LEN, "CrtMlFrtContinue")?;
            let mut reader = CrtMlFrtReader::new(continuation);
            validate_frt_header(
                &mut reader,
                CRT_ML_FRT_CONTINUE_RECORD_TYPE,
                "CrtMlFrtContinue",
            )?;
            chain_len = chain_len
                .checked_add(reader.remaining())
                .ok_or_else(|| invalid("CrtMlFrt XmlTkChain length overflows usize"))?;
        }
        // MS-XLS 2.4.70: cb specifies the exact size of the XmlTkChain,
        // including the continuation record data.
        if chain_len != declared {
            return Err(invalid(format!(
                "CrtMlFrt cb {declared} does not match its XmlTkChain size {chain_len}"
            )));
        }

        let mut chain = Vec::new();
        chain
            .try_reserve_exact(chain_len)
            .map_err(|_| XlsError::Allocation("reassembling CrtMlFrt XmlTkChain"))?;
        append_chain_bytes(&mut chain, first_chain)?;
        for continuation in continues {
            let mut reader = CrtMlFrtReader::new(continuation);
            validate_frt_header(
                &mut reader,
                CRT_ML_FRT_CONTINUE_RECORD_TYPE,
                "CrtMlFrtContinue",
            )?;
            append_chain_bytes(&mut chain, reader.read_remaining()?)?;
        }
        Ok(Self {
            flags,
            reserved,
            chain,
            unused,
        })
    }

    /// Serialize as a sequence of complete record payloads: the `CrtMlFrt`
    /// record followed by `CrtMlFrtContinue` records when the chain exceeds
    /// one record.
    pub fn to_record_payloads(&self) -> Vec<Vec<u8>> {
        let mut first = Vec::new();
        first.extend_from_slice(&CRT_ML_FRT_RECORD_TYPE.to_le_bytes());
        first.extend_from_slice(&self.flags.to_le_bytes());
        first.extend_from_slice(&self.reserved);
        first.extend_from_slice(&(self.chain.len() as u32).to_le_bytes());
        let first_chunk = MAX_RECORD_PAYLOAD - (FRT_HEADER_LEN + FIELD_LEN + FIELD_LEN);
        let mut chunks = self.chain.chunks(first_chunk);
        first.extend_from_slice(chunks.next().unwrap_or(&[]));
        first.extend_from_slice(&self.unused);
        let mut records = vec![first];
        for chunk in chunks {
            let mut continuation = Vec::with_capacity(FRT_HEADER_LEN + chunk.len());
            continuation.extend_from_slice(&CRT_ML_FRT_CONTINUE_RECORD_TYPE.to_le_bytes());
            continuation.extend_from_slice(&self.flags.to_le_bytes());
            continuation.extend_from_slice(&self.reserved);
            continuation.extend_from_slice(chunk);
            records.push(continuation);
        }
        records
    }

    /// Raw `frtHeader.grbitFrt` bitfield (`fFrtRef`/`fFrtAlert` are zero).
    pub const fn flags(&self) -> u16 {
        self.flags
    }

    /// The eight `FrtHeader.reserved` bytes retained from the source record.
    pub const fn reserved(&self) -> [u8; 8] {
        self.reserved
    }

    /// The opaque `XmlTkChain` bytes.
    pub fn chain(&self) -> &[u8] {
        &self.chain
    }

    /// The preserved trailing `unused` field.
    pub const fn unused(&self) -> [u8; 4] {
        self.unused
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn crt_ml_frt_record(chain: &[u8], unused: [u8; 4]) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&CRT_ML_FRT_RECORD_TYPE.to_le_bytes());
        data.extend_from_slice(&[0; FRT_HEADER_LEN - 2]);
        data.extend_from_slice(&(chain.len() as u32).to_le_bytes());
        data.extend_from_slice(chain);
        data.extend_from_slice(&unused);
        data
    }

    fn continuation(chain: &[u8]) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&CRT_ML_FRT_CONTINUE_RECORD_TYPE.to_le_bytes());
        data.extend_from_slice(&[0; FRT_HEADER_LEN - 2]);
        data.extend_from_slice(chain);
        data
    }

    #[test]
    fn parses_single_record_and_round_trips_exactly() {
        let record = crt_ml_frt_record(b"chain-bytes", [0xDE, 0xAD, 0xBE, 0xEF]);
        let parsed = CrtMlFrt::parse(&record, &[]).unwrap();
        assert_eq!(parsed.chain(), b"chain-bytes");
        assert_eq!(parsed.flags(), 0);
        assert_eq!(parsed.unused(), [0xDE, 0xAD, 0xBE, 0xEF]);
        let payloads = parsed.to_record_payloads();
        assert_eq!(payloads.len(), 1);
        assert_eq!(payloads[0], record);
    }

    #[test]
    fn reassembles_continuations_and_round_trips() {
        let chain = vec![0x5Au8; 30_000];
        let mut first = crt_ml_frt_record(&[], [0; 4]);
        // Splice the declared length and the first chain chunk into the record.
        let first_chunk = MAX_RECORD_PAYLOAD - (FRT_HEADER_LEN + FIELD_LEN + FIELD_LEN);
        first.truncate(FRT_HEADER_LEN);
        first.extend_from_slice(&(chain.len() as u32).to_le_bytes());
        first.extend_from_slice(&chain[..first_chunk]);
        first.extend_from_slice(&[0; 4]);
        let continuation_chunk = MAX_RECORD_PAYLOAD - FRT_HEADER_LEN;
        let continues = chain[first_chunk..]
            .chunks(continuation_chunk)
            .map(continuation)
            .collect::<Vec<_>>();
        let parsed = CrtMlFrt::parse(&first, &continues).unwrap();
        assert_eq!(parsed.chain(), chain.as_slice());

        let payloads = parsed.to_record_payloads();
        assert!(payloads.len() > 1);
        for payload in &payloads {
            assert!(payload.len() <= MAX_RECORD_PAYLOAD);
        }
        let reparsed = CrtMlFrt::parse(&payloads[0], &payloads[1..]).unwrap();
        assert_eq!(reparsed, parsed);
        assert_eq!(reparsed.chain(), chain.as_slice());
    }

    #[test]
    fn empty_chain_round_trips() {
        let record = crt_ml_frt_record(&[], [1, 2, 3, 4]);
        let parsed = CrtMlFrt::parse(&record, &[]).unwrap();
        assert!(parsed.chain().is_empty());
        assert_eq!(parsed.to_record_payloads(), vec![record]);
    }

    #[test]
    fn rejects_malformed_records() {
        let record = crt_ml_frt_record(b"chain", [0; 4]);
        // Truncated.
        assert!(CrtMlFrt::parse(&record[..10], &[]).is_err());
        // Wrong FrtHeader.rt (0x089D is the neighboring CrtLayout12).
        let mut wrong_rt = record.clone();
        wrong_rt[0..2].copy_from_slice(&0x089Du16.to_le_bytes());
        assert!(CrtMlFrt::parse(&wrong_rt, &[]).is_err());
        // fFrtRef / fFrtAlert set.
        for flags in [0x0001u16, 0x0002] {
            let mut bad = record.clone();
            bad[2..4].copy_from_slice(&flags.to_le_bytes());
            assert!(CrtMlFrt::parse(&bad, &[]).is_err());
        }
        // cb exceeds the legal maximum.
        let mut huge = record.clone();
        huge[FRT_HEADER_LEN..FRT_HEADER_LEN + 4].copy_from_slice(&0x8000_0000u32.to_le_bytes());
        assert!(CrtMlFrt::parse(&huge, &[]).is_err());
        // cb does not match the reassembled chain size.
        let mut short = record.clone();
        short[FRT_HEADER_LEN..FRT_HEADER_LEN + 4].copy_from_slice(&99u32.to_le_bytes());
        assert!(CrtMlFrt::parse(&short, &[]).is_err());
        assert!(CrtMlFrt::parse(&record, &[continuation(b"extra")]).is_err());
        // Continuation with a wrong FrtHeader.rt.
        let mut bad = continuation(b"x");
        bad[0..2].copy_from_slice(&0x003Cu16.to_le_bytes());
        let mut first = crt_ml_frt_record(b"ab", [0; 4]);
        first[FRT_HEADER_LEN..FRT_HEADER_LEN + 4].copy_from_slice(&3u32.to_le_bytes());
        assert!(CrtMlFrt::parse(&first, &[bad]).is_err());
    }

    #[test]
    fn rejects_truncated_and_oversized_payloads() {
        let record = crt_ml_frt_record(b"chain", [0; 4]);
        for length in 0..record.len() {
            assert!(
                CrtMlFrt::parse(&record[..length], &[]).is_err(),
                "expected truncated CrtMlFrt payload of {length} bytes to fail"
            );
        }

        let continuation_record = continuation(b"continuation");
        for length in 0..FRT_HEADER_LEN {
            assert!(
                CrtMlFrt::parse(&record, &[continuation_record[..length].to_vec()]).is_err(),
                "expected truncated CrtMlFrtContinue payload of {length} bytes to fail"
            );
        }

        let mut oversized = record.clone();
        oversized.resize(MAX_RECORD_PAYLOAD + 1, 0);
        assert!(CrtMlFrt::parse(&oversized, &[]).is_err());

        let oversized_continuation =
            continuation(&vec![0; MAX_RECORD_PAYLOAD - FRT_HEADER_LEN + 1]);
        assert!(CrtMlFrt::parse(&record, &[oversized_continuation]).is_err());
    }

    #[test]
    fn accepts_empty_continuation_payload() {
        let record = crt_ml_frt_record(b"", [0; 4]);
        let parsed = CrtMlFrt::parse(&record, &[continuation(&[])]).unwrap();
        assert!(parsed.chain().is_empty());
        assert_eq!(parsed.to_record_payloads(), vec![record]);
    }

    #[test]
    fn preserves_reserved_header_bytes() {
        let mut record = crt_ml_frt_record(b"x", [0; 4]);
        // The 8 reserved FrtHeader bytes MUST be ignored but round-trip.
        record[4..FRT_HEADER_LEN].copy_from_slice(&[1, 2, 3, 4, 5, 6, 7, 8]);
        // The 14 reserved grbitFrt bits are also preserved (fFrtRef/fFrtAlert stay 0).
        record[2..4].copy_from_slice(&0xFFFCu16.to_le_bytes());
        let parsed = CrtMlFrt::parse(&record, &[]).unwrap();
        assert_eq!(parsed.flags(), 0xFFFC);
        assert_eq!(parsed.reserved(), [1, 2, 3, 4, 5, 6, 7, 8]);
        assert_eq!(parsed.to_record_payloads(), vec![record]);
    }
}
