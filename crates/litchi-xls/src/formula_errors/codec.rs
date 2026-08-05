//! Bounded BIFF8 codecs and worksheet-scoped collection for formula errors.

use super::model::{Checks, Feature, Header, Range};
use crate::error::{Error, Result};

pub(super) const FEAT_HDR_RECORD_TYPE: u16 = 0x0867;
pub(super) const FEAT_RECORD_TYPE: u16 = 0x0868;
const ISF_FEC2: u16 = 0x0003;
const HEADER_PAYLOAD_LEN: usize = 19;
const FEATURE_FIXED_LEN: usize = 27;
const FEATURE_DATA_LEN: usize = 4;
const MAX_RECORD_PAYLOAD_LEN: usize = 8_224;
pub(super) const MAX_RANGES: usize = 1_024;
const MAX_FEATURES: usize = u16::MAX as usize;

fn invalid(record_type: u16, message: impl Into<String>) -> Error {
    Error::InvalidRecord {
        record_type,
        message: message.into(),
    }
}

fn read_u16(data: &[u8], offset: usize, record_type: u16, field: &str) -> Result<u16> {
    data.get(offset..offset + 2)
        .map(|bytes| u16::from_le_bytes([bytes[0], bytes[1]]))
        .ok_or_else(|| invalid(record_type, format!("truncated {field}")))
}

fn read_u32(data: &[u8], offset: usize, record_type: u16, field: &str) -> Result<u32> {
    data.get(offset..offset + 4)
        .map(|bytes| u32::from_le_bytes(bytes.try_into().unwrap()))
        .ok_or_else(|| invalid(record_type, format!("truncated {field}")))
}

fn validate_frt_header(data: &[u8], record_type: u16) -> Result<()> {
    if read_u16(data, 0, record_type, "frtHeader.rt")? != record_type {
        return Err(invalid(
            record_type,
            "future-record type does not match containing record",
        ));
    }
    if read_u16(data, 2, record_type, "frtHeader.grbitFrt")? != 0 {
        return Err(invalid(record_type, "future-record flags must be zero"));
    }
    let reserved = data
        .get(4..12)
        .ok_or_else(|| invalid(record_type, "truncated future-record reserved bytes"))?;
    if reserved.iter().any(|&byte| byte != 0) {
        return Err(invalid(
            record_type,
            "future-record reserved bytes must be zero",
        ));
    }
    Ok(())
}

fn append_frt_header(data: &mut Vec<u8>, record_type: u16) {
    data.extend_from_slice(&record_type.to_le_bytes());
    data.extend_from_slice(&0u16.to_le_bytes());
    data.extend_from_slice(&[0u8; 8]);
}

fn with_record_header(record_type: u16, payload: Vec<u8>) -> Result<Vec<u8>> {
    let length = u16::try_from(payload.len())
        .map_err(|_| invalid(record_type, "payload length exceeds BIFF u16"))?;
    let mut record = Vec::with_capacity(4 + payload.len());
    record.extend_from_slice(&record_type.to_le_bytes());
    record.extend_from_slice(&length.to_le_bytes());
    record.extend_from_slice(&payload);
    Ok(record)
}

impl Header {
    pub fn parse_payload(data: &[u8]) -> Result<Self> {
        if data.len() != HEADER_PAYLOAD_LEN {
            return Err(invalid(
                FEAT_HDR_RECORD_TYPE,
                format!(
                    "FeatHdr ISFFEC2 payload must be exactly 19 bytes, got {}",
                    data.len()
                ),
            ));
        }
        validate_frt_header(data, FEAT_HDR_RECORD_TYPE)?;
        if read_u16(data, 12, FEAT_HDR_RECORD_TYPE, "FeatHdr.isf")? != ISF_FEC2 {
            return Err(invalid(
                FEAT_HDR_RECORD_TYPE,
                "FeatHdr shared-feature type must be ISFFEC2",
            ));
        }
        if data[14] != 1 {
            return Err(invalid(
                FEAT_HDR_RECORD_TYPE,
                "FeatHdr reserved byte must be one",
            ));
        }
        if read_u32(data, 15, FEAT_HDR_RECORD_TYPE, "FeatHdr.cbHdrData")? != 0 {
            return Err(invalid(
                FEAT_HDR_RECORD_TYPE,
                "ISFFEC2 FeatHdr must not contain header data",
            ));
        }
        Ok(Self)
    }

    /// Serialize the complete `FeatHdr` payload deterministically.
    pub fn to_payload(self) -> Vec<u8> {
        let mut data = Vec::with_capacity(HEADER_PAYLOAD_LEN);
        append_frt_header(&mut data, FEAT_HDR_RECORD_TYPE);
        data.extend_from_slice(&ISF_FEC2.to_le_bytes());
        data.push(1);
        data.extend_from_slice(&0u32.to_le_bytes());
        data
    }

    /// Serialize the complete BIFF record including its four-byte record header.
    pub fn to_record_bytes(self) -> Result<Vec<u8>> {
        with_record_header(FEAT_HDR_RECORD_TYPE, self.to_payload())
    }
}

impl Range {
    pub fn try_new(
        first_row: u16,
        last_row: u16,
        first_column: u16,
        last_column: u16,
    ) -> Result<Self> {
        if first_row > last_row {
            return Err(invalid(
                FEAT_RECORD_TYPE,
                "range first row exceeds last row",
            ));
        }
        if first_column > last_column || last_column > 255 {
            return Err(invalid(
                FEAT_RECORD_TYPE,
                "range columns are reversed or exceed 255",
            ));
        }
        Ok(Self {
            first_row,
            last_row,
            first_column: first_column as u8,
            last_column: last_column as u8,
        })
    }
}

impl Feature {
    pub fn try_new(ranges: Vec<Range>, checks: Checks) -> Result<Self> {
        if ranges.len() > MAX_RANGES {
            return Err(invalid(
                FEAT_RECORD_TYPE,
                format!("Feat ISFFEC2 range count exceeds {MAX_RANGES}"),
            ));
        }
        Ok(Self { ranges, checks })
    }

    pub fn parse_payload(data: &[u8]) -> Result<Self> {
        let minimum = FEATURE_FIXED_LEN + FEATURE_DATA_LEN;
        if !(minimum..=MAX_RECORD_PAYLOAD_LEN).contains(&data.len()) {
            return Err(invalid(
                FEAT_RECORD_TYPE,
                format!(
                    "Feat ISFFEC2 payload must be {minimum}..={MAX_RECORD_PAYLOAD_LEN} bytes, got {}",
                    data.len()
                ),
            ));
        }
        validate_frt_header(data, FEAT_RECORD_TYPE)?;
        if read_u16(data, 12, FEAT_RECORD_TYPE, "Feat.isf")? != ISF_FEC2 {
            return Err(invalid(
                FEAT_RECORD_TYPE,
                "Feat shared-feature type must be ISFFEC2",
            ));
        }
        if data[14] != 0 || read_u32(data, 15, FEAT_RECORD_TYPE, "Feat.reserved2")? != 0 {
            return Err(invalid(
                FEAT_RECORD_TYPE,
                "Feat reserved1 and reserved2 must be zero",
            ));
        }
        let range_count = usize::from(read_u16(data, 19, FEAT_RECORD_TYPE, "Feat.cref")?);
        if range_count > MAX_RANGES {
            return Err(invalid(
                FEAT_RECORD_TYPE,
                format!("Feat range count exceeds {MAX_RANGES}"),
            ));
        }
        if read_u32(data, 21, FEAT_RECORD_TYPE, "Feat.cbFeatData")? != FEATURE_DATA_LEN as u32 {
            return Err(invalid(FEAT_RECORD_TYPE, "ISFFEC2 cbFeatData must be four"));
        }
        if read_u16(data, 25, FEAT_RECORD_TYPE, "Feat.reserved3")? != 0 {
            return Err(invalid(FEAT_RECORD_TYPE, "Feat reserved3 must be zero"));
        }
        let expected = range_count
            .checked_mul(8)
            .and_then(|bytes| bytes.checked_add(minimum))
            .ok_or_else(|| invalid(FEAT_RECORD_TYPE, "Feat payload size overflows"))?;
        if data.len() != expected {
            return Err(invalid(
                FEAT_RECORD_TYPE,
                format!("Feat cref requires {expected} bytes, got {}", data.len()),
            ));
        }

        let mut ranges = Vec::with_capacity(range_count);
        let mut offset = FEATURE_FIXED_LEN;
        for _ in 0..range_count {
            let first_row = read_u16(data, offset, FEAT_RECORD_TYPE, "Ref8U.rwFirst")?;
            let last_row = read_u16(data, offset + 2, FEAT_RECORD_TYPE, "Ref8U.rwLast")?;
            let first_column = read_u16(data, offset + 4, FEAT_RECORD_TYPE, "Ref8U.colFirst")?;
            let last_column = read_u16(data, offset + 6, FEAT_RECORD_TYPE, "Ref8U.colLast")?;
            ranges.push(Range::try_new(
                first_row,
                last_row,
                first_column,
                last_column,
            )?);
            offset += 8;
        }
        let raw_checks = read_u32(
            data,
            offset,
            FEAT_RECORD_TYPE,
            "FeatFormulaErr2.grffecIgnore",
        )?;
        if raw_checks & !0xFF != 0 {
            return Err(invalid(
                FEAT_RECORD_TYPE,
                "FFErrorCheck reserved bits must be zero",
            ));
        }
        Self::try_new(ranges, Checks::from_bits(raw_checks as u8))
    }

    /// Serialize the complete `Feat` payload deterministically.
    pub fn to_payload(&self) -> Result<Vec<u8>> {
        let size = self
            .ranges
            .len()
            .checked_mul(8)
            .and_then(|bytes| bytes.checked_add(FEATURE_FIXED_LEN + FEATURE_DATA_LEN))
            .ok_or_else(|| invalid(FEAT_RECORD_TYPE, "Feat serialized size overflows"))?;
        if self.ranges.len() > MAX_RANGES || size > MAX_RECORD_PAYLOAD_LEN {
            return Err(invalid(FEAT_RECORD_TYPE, "Feat exceeds BIFF payload cap"));
        }
        let mut data = Vec::with_capacity(size);
        append_frt_header(&mut data, FEAT_RECORD_TYPE);
        data.extend_from_slice(&ISF_FEC2.to_le_bytes());
        data.push(0);
        data.extend_from_slice(&0u32.to_le_bytes());
        data.extend_from_slice(&(self.ranges.len() as u16).to_le_bytes());
        data.extend_from_slice(&(FEATURE_DATA_LEN as u32).to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes());
        for range in &self.ranges {
            data.extend_from_slice(&range.first_row.to_le_bytes());
            data.extend_from_slice(&range.last_row.to_le_bytes());
            data.extend_from_slice(&u16::from(range.first_column).to_le_bytes());
            data.extend_from_slice(&u16::from(range.last_column).to_le_bytes());
        }
        data.extend_from_slice(&u32::from(self.checks.bits).to_le_bytes());
        Ok(data)
    }

    /// Serialize the complete BIFF record including its four-byte record header.
    pub fn to_record_bytes(&self) -> Result<Vec<u8>> {
        with_record_header(FEAT_RECORD_TYPE, self.to_payload()?)
    }
}

pub(crate) struct FormulaErrorCollector {
    header_seen: bool,
    features: Vec<Feature>,
}

impl FormulaErrorCollector {
    pub(crate) fn new() -> Self {
        Self {
            header_seen: false,
            features: Vec::new(),
        }
    }

    pub(crate) fn feed_record(&mut self, record_type: u16, data: &[u8]) -> Result<()> {
        if !matches!(record_type, FEAT_HDR_RECORD_TYPE | FEAT_RECORD_TYPE) {
            return Ok(());
        }
        if data.len() < 14 {
            return Err(invalid(
                record_type,
                "shared-feature record is truncated before isf",
            ));
        }
        if read_u16(data, 12, record_type, "shared-feature isf")? != ISF_FEC2 {
            return Ok(());
        }
        match record_type {
            FEAT_HDR_RECORD_TYPE => {
                if self.header_seen {
                    return Err(invalid(record_type, "duplicate ISFFEC2 FeatHdr"));
                }
                Header::parse_payload(data)?;
                self.header_seen = true;
            },
            FEAT_RECORD_TYPE => {
                if !self.header_seen {
                    return Err(invalid(record_type, "ISFFEC2 Feat appears before FeatHdr"));
                }
                if self.features.len() >= MAX_FEATURES {
                    return Err(invalid(
                        record_type,
                        "too many ISFFEC2 features in worksheet",
                    ));
                }
                self.features.push(Feature::parse_payload(data)?);
            },
            _ => unreachable!(),
        }
        Ok(())
    }

    pub(crate) fn finish(self) -> Result<Vec<Feature>> {
        if self.header_seen && self.features.is_empty() {
            return Err(invalid(
                FEAT_HDR_RECORD_TYPE,
                "ISFFEC2 FeatHdr has no following Feat",
            ));
        }
        Ok(self.features)
    }
}
