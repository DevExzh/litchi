//! BIFF8 formula error-checking shared features (`FeatHdr` + `Feat`).

use super::{XlsError, XlsResult};

pub(crate) const FEAT_HDR_RECORD_TYPE: u16 = 0x0867;
pub(crate) const FEAT_RECORD_TYPE: u16 = 0x0868;
const ISF_FEC2: u16 = 0x0003;
const HEADER_PAYLOAD_LEN: usize = 19;
const FEATURE_FIXED_LEN: usize = 27;
const FEATURE_DATA_LEN: usize = 4;
const MAX_RECORD_PAYLOAD_LEN: usize = 8_224;
const MAX_RANGES: usize = 1_024;
const MAX_FEATURES: usize = u16::MAX as usize;

fn invalid(record_type: u16, message: impl Into<String>) -> XlsError {
    XlsError::InvalidRecord {
        record_type,
        message: message.into(),
    }
}

fn read_u16(data: &[u8], offset: usize, record_type: u16, field: &str) -> XlsResult<u16> {
    data.get(offset..offset + 2)
        .map(|bytes| u16::from_le_bytes([bytes[0], bytes[1]]))
        .ok_or_else(|| invalid(record_type, format!("truncated {field}")))
}

fn read_u32(data: &[u8], offset: usize, record_type: u16, field: &str) -> XlsResult<u32> {
    data.get(offset..offset + 4)
        .map(|bytes| u32::from_le_bytes(bytes.try_into().unwrap()))
        .ok_or_else(|| invalid(record_type, format!("truncated {field}")))
}

fn validate_frt_header(data: &[u8], record_type: u16) -> XlsResult<()> {
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

fn with_record_header(record_type: u16, payload: Vec<u8>) -> XlsResult<Vec<u8>> {
    let length = u16::try_from(payload.len())
        .map_err(|_| invalid(record_type, "payload length exceeds BIFF u16"))?;
    let mut record = Vec::with_capacity(4 + payload.len());
    record.extend_from_slice(&record_type.to_le_bytes());
    record.extend_from_slice(&length.to_le_bytes());
    record.extend_from_slice(&payload);
    Ok(record)
}

/// Header starting a worksheet formula error-checking feature collection.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct XlsFormulaErrorHeader;

impl XlsFormulaErrorHeader {
    pub fn new() -> Self {
        Self
    }

    pub fn parse_payload(data: &[u8]) -> XlsResult<Self> {
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
    pub fn to_record_bytes(self) -> XlsResult<Vec<u8>> {
        with_record_header(FEAT_HDR_RECORD_TYPE, self.to_payload())
    }
}

/// Inclusive BIFF8 cell range targeted by a formula error feature.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsFormulaErrorRange {
    first_row: u16,
    last_row: u16,
    first_column: u8,
    last_column: u8,
}

impl XlsFormulaErrorRange {
    pub fn try_new(
        first_row: u16,
        last_row: u16,
        first_column: u16,
        last_column: u16,
    ) -> XlsResult<Self> {
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

    pub fn first_row(self) -> u16 {
        self.first_row
    }
    pub fn last_row(self) -> u16 {
        self.last_row
    }
    pub fn first_column(self) -> u8 {
        self.first_column
    }
    pub fn last_column(self) -> u8 {
        self.last_column
    }
}

/// Formula conditions selected by an `FFErrorCheck` bit field.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct XlsFormulaErrorChecks {
    bits: u8,
}

impl XlsFormulaErrorChecks {
    pub const fn from_bits(bits: u8) -> Self {
        Self { bits }
    }
    pub const fn bits(self) -> u8 {
        self.bits
    }
    pub const fn calculation_errors(self) -> bool {
        self.bits & 0x01 != 0
    }
    pub const fn empty_cell_references(self) -> bool {
        self.bits & 0x02 != 0
    }
    pub const fn numbers_stored_as_text(self) -> bool {
        self.bits & 0x04 != 0
    }
    pub const fn inconsistent_ranges(self) -> bool {
        self.bits & 0x08 != 0
    }
    pub const fn inconsistent_formulas(self) -> bool {
        self.bits & 0x10 != 0
    }
    pub const fn insufficient_date_time_formats(self) -> bool {
        self.bits & 0x20 != 0
    }
    pub const fn unprotected_formulas(self) -> bool {
        self.bits & 0x40 != 0
    }
    pub const fn data_validation(self) -> bool {
        self.bits & 0x80 != 0
    }
}

/// One worksheet formula error-checking shared feature.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsFormulaErrorFeature {
    ranges: Vec<XlsFormulaErrorRange>,
    checks: XlsFormulaErrorChecks,
}

impl XlsFormulaErrorFeature {
    pub fn try_new(
        ranges: Vec<XlsFormulaErrorRange>,
        checks: XlsFormulaErrorChecks,
    ) -> XlsResult<Self> {
        if ranges.len() > MAX_RANGES {
            return Err(invalid(
                FEAT_RECORD_TYPE,
                format!("Feat ISFFEC2 range count exceeds {MAX_RANGES}"),
            ));
        }
        Ok(Self { ranges, checks })
    }

    pub fn parse_payload(data: &[u8]) -> XlsResult<Self> {
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
            ranges.push(XlsFormulaErrorRange::try_new(
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
        Self::try_new(ranges, XlsFormulaErrorChecks::from_bits(raw_checks as u8))
    }

    pub fn ranges(&self) -> &[XlsFormulaErrorRange] {
        &self.ranges
    }
    pub fn checks(&self) -> XlsFormulaErrorChecks {
        self.checks
    }

    /// Serialize the complete `Feat` payload deterministically.
    pub fn to_payload(&self) -> XlsResult<Vec<u8>> {
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
    pub fn to_record_bytes(&self) -> XlsResult<Vec<u8>> {
        with_record_header(FEAT_RECORD_TYPE, self.to_payload()?)
    }
}

pub(crate) struct FormulaErrorCollector {
    header_seen: bool,
    features: Vec<XlsFormulaErrorFeature>,
}

impl FormulaErrorCollector {
    pub(crate) fn new() -> Self {
        Self {
            header_seen: false,
            features: Vec::new(),
        }
    }

    pub(crate) fn feed_record(&mut self, record_type: u16, data: &[u8]) -> XlsResult<()> {
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
                XlsFormulaErrorHeader::parse_payload(data)?;
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
                self.features
                    .push(XlsFormulaErrorFeature::parse_payload(data)?);
            },
            _ => unreachable!(),
        }
        Ok(())
    }

    pub(crate) fn finish(self) -> XlsResult<Vec<XlsFormulaErrorFeature>> {
        if self.header_seen && self.features.is_empty() {
            return Err(invalid(
                FEAT_HDR_RECORD_TYPE,
                "ISFFEC2 FeatHdr has no following Feat",
            ));
        }
        Ok(self.features)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    const POI_HEADER: [u8; 19] = [
        0x67, 0x08, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 3, 0, 1, 0, 0, 0, 0,
    ];
    const POI_FEATURE: [u8; 39] = [
        0x68, 0x08, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 3, 0, 0, 0, 0, 0, 0, 1, 0, 4, 0, 0, 0, 0, 0, 0,
        0, 0, 0, 0, 0, 0, 0, 4, 0, 0, 0,
    ];

    #[test]
    fn parses_and_round_trips_poi_reference_records() {
        let header = XlsFormulaErrorHeader::parse_payload(&POI_HEADER).unwrap();
        assert_eq!(header.to_payload(), POI_HEADER);
        assert_eq!(
            &header.to_record_bytes().unwrap()[0..4],
            &[0x67, 0x08, 19, 0]
        );
        let feature = XlsFormulaErrorFeature::parse_payload(&POI_FEATURE).unwrap();
        assert_eq!(feature.ranges().len(), 1);
        assert_eq!(
            feature.ranges()[0],
            XlsFormulaErrorRange::try_new(0, 0, 0, 0).unwrap()
        );
        assert!(feature.checks().numbers_stored_as_text());
        assert!(!feature.checks().calculation_errors());
        assert_eq!(feature.to_payload().unwrap(), POI_FEATURE);
        assert_eq!(
            &feature.to_record_bytes().unwrap()[0..4],
            &[0x68, 0x08, 39, 0]
        );
    }

    #[test]
    fn constructs_multiple_ranges_and_all_flags() {
        let ranges = vec![
            XlsFormulaErrorRange::try_new(1, 7, 3, 3).unwrap(),
            XlsFormulaErrorRange::try_new(12, 18, 3, 3).unwrap(),
            XlsFormulaErrorRange::try_new(22, 28, 3, 3).unwrap(),
        ];
        let feature =
            XlsFormulaErrorFeature::try_new(ranges.clone(), XlsFormulaErrorChecks::from_bits(0xFF))
                .unwrap();
        let reparsed =
            XlsFormulaErrorFeature::parse_payload(&feature.to_payload().unwrap()).unwrap();
        assert_eq!(reparsed.ranges(), ranges);
        let checks = reparsed.checks();
        assert!(checks.calculation_errors());
        assert!(checks.empty_cell_references());
        assert!(checks.numbers_stored_as_text());
        assert!(checks.inconsistent_ranges());
        assert!(checks.inconsistent_formulas());
        assert!(checks.insufficient_date_time_formats());
        assert!(checks.unprotected_formulas());
        assert!(checks.data_validation());
    }

    #[test]
    fn rejects_malformed_headers_lengths_ranges_flags_and_ordering() {
        let mut bad = POI_HEADER.to_vec();
        bad[0] = 0x68;
        assert!(XlsFormulaErrorHeader::parse_payload(&bad).is_err());
        let mut bad = POI_HEADER.to_vec();
        bad[2] = 1;
        assert!(XlsFormulaErrorHeader::parse_payload(&bad).is_err());
        let mut bad = POI_HEADER.to_vec();
        bad[4] = 1;
        assert!(XlsFormulaErrorHeader::parse_payload(&bad).is_err());
        let mut bad = POI_HEADER.to_vec();
        bad[14] = 0;
        assert!(XlsFormulaErrorHeader::parse_payload(&bad).is_err());
        let mut bad = POI_HEADER.to_vec();
        bad[15] = 1;
        assert!(XlsFormulaErrorHeader::parse_payload(&bad).is_err());
        assert!(XlsFormulaErrorHeader::parse_payload(&POI_HEADER[..18]).is_err());

        let mut bad = POI_FEATURE.to_vec();
        bad[14] = 1;
        assert!(XlsFormulaErrorFeature::parse_payload(&bad).is_err());
        let mut bad = POI_FEATURE.to_vec();
        bad[19] = 2;
        assert!(XlsFormulaErrorFeature::parse_payload(&bad).is_err());
        let mut bad = POI_FEATURE.to_vec();
        bad[21] = 5;
        assert!(XlsFormulaErrorFeature::parse_payload(&bad).is_err());
        let mut bad = POI_FEATURE.to_vec();
        bad[27] = 1;
        assert!(XlsFormulaErrorFeature::parse_payload(&bad).is_err());
        let mut bad = POI_FEATURE.to_vec();
        bad[31] = 0;
        bad[32] = 1;
        assert!(XlsFormulaErrorFeature::parse_payload(&bad).is_err());
        let mut bad = POI_FEATURE.to_vec();
        bad[38] = 1;
        assert!(XlsFormulaErrorFeature::parse_payload(&bad).is_err());
        assert!(XlsFormulaErrorFeature::parse_payload(&POI_FEATURE[..38]).is_err());

        let range = XlsFormulaErrorRange::try_new(0, 0, 0, 0).unwrap();
        assert!(
            XlsFormulaErrorFeature::try_new(
                vec![range; MAX_RANGES + 1],
                XlsFormulaErrorChecks::default(),
            )
            .is_err()
        );

        let mut collector = FormulaErrorCollector::new();
        assert!(
            collector
                .feed_record(FEAT_RECORD_TYPE, &POI_FEATURE)
                .is_err()
        );
        let mut collector = FormulaErrorCollector::new();
        collector
            .feed_record(FEAT_HDR_RECORD_TYPE, &POI_HEADER)
            .unwrap();
        assert!(
            collector
                .feed_record(FEAT_HDR_RECORD_TYPE, &POI_HEADER)
                .is_err()
        );
        let mut collector = FormulaErrorCollector::new();
        collector
            .feed_record(FEAT_HDR_RECORD_TYPE, &POI_HEADER)
            .unwrap();
        assert!(collector.finish().is_err());
    }
}
