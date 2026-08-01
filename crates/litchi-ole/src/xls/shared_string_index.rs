//! BIFF8 extended shared-string table index (`ExtSST`).

use super::{XlsError, XlsResult};

pub(crate) const SST_RECORD_TYPE: u16 = 0x00FC;
pub(crate) const EXT_SST_RECORD_TYPE: u16 = 0x00FF;
const FIXED_PAYLOAD_LEN: usize = 2;
const BUCKET_LEN: usize = 8;
const MAX_BUCKETS: usize = 128;
const MAX_PAYLOAD_LEN: usize = FIXED_PAYLOAD_LEN + MAX_BUCKETS * BUCKET_LEN;

fn invalid(message: impl Into<String>) -> XlsError {
    XlsError::InvalidRecord {
        record_type: EXT_SST_RECORD_TYPE,
        message: message.into(),
    }
}

fn strings_per_bucket(unique_string_count: u32) -> XlsResult<u16> {
    let value = (unique_string_count / 128 + 1).max(8);
    u16::try_from(value)
        .map_err(|_| invalid("SST unique-string count is too large for the ExtSST dsst field"))
}

fn required_bucket_count(unique_string_count: u32, bucket_size: u16) -> XlsResult<usize> {
    if bucket_size == 0 {
        return Err(invalid("ExtSST dsst must not be zero"));
    }
    let count = if unique_string_count == 0 {
        0
    } else {
        (unique_string_count - 1) / u32::from(bucket_size) + 1
    };
    let count =
        usize::try_from(count).map_err(|_| invalid("ExtSST bucket count does not fit usize"))?;
    if count > MAX_BUCKETS {
        return Err(invalid("ExtSST requires more than 128 buckets"));
    }
    Ok(count)
}

/// Pointer to the first shared string in one `ExtSST` bucket.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsSharedStringBucket {
    stream_position: u32,
    record_offset: u16,
}

impl XlsSharedStringBucket {
    pub fn try_new(stream_position: u32, record_offset: u16) -> XlsResult<Self> {
        if u32::from(record_offset) >= stream_position {
            return Err(invalid("ISSTInf cbOffset must be less than ib"));
        }
        Ok(Self {
            stream_position,
            record_offset,
        })
    }

    /// Absolute zero-based position of the bucket's first string.
    pub fn stream_position(self) -> u32 {
        self.stream_position
    }

    /// Offset of the first string within its containing SST or Continue record.
    pub fn record_offset(self) -> u16 {
        self.record_offset
    }

    /// Absolute position of the containing SST or Continue record header.
    pub fn record_position(self) -> u32 {
        self.stream_position - u32::from(self.record_offset)
    }
}

/// Quick-lookup bucket index for a BIFF8 shared string table.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsSharedStringIndex {
    unique_string_count: u32,
    strings_per_bucket: u16,
    buckets: Vec<XlsSharedStringBucket>,
}

impl XlsSharedStringIndex {
    pub fn try_new(
        unique_string_count: u32,
        buckets: Vec<XlsSharedStringBucket>,
    ) -> XlsResult<Self> {
        let strings_per_bucket = strings_per_bucket(unique_string_count)?;
        let expected = required_bucket_count(unique_string_count, strings_per_bucket)?;
        if buckets.len() != expected {
            return Err(invalid(format!(
                "ExtSST requires {expected} buckets for {unique_string_count} unique strings, got {}",
                buckets.len()
            )));
        }
        Ok(Self {
            unique_string_count,
            strings_per_bucket,
            buckets,
        })
    }

    pub fn parse_payload(unique_string_count: u32, data: &[u8]) -> XlsResult<Self> {
        if !(FIXED_PAYLOAD_LEN..=MAX_PAYLOAD_LEN).contains(&data.len()) {
            return Err(invalid(format!(
                "ExtSST payload must be 2..={MAX_PAYLOAD_LEN} bytes, got {}",
                data.len()
            )));
        }
        if !(data.len() - FIXED_PAYLOAD_LEN).is_multiple_of(BUCKET_LEN) {
            return Err(invalid("ExtSST payload has a partial ISSTInf entry"));
        }
        let actual_bucket_size = u16::from_le_bytes([data[0], data[1]]);
        let expected_bucket_size = strings_per_bucket(unique_string_count)?;
        if actual_bucket_size != expected_bucket_size {
            return Err(invalid(format!(
                "ExtSST dsst must be {expected_bucket_size} for {unique_string_count} unique strings, got {actual_bucket_size}"
            )));
        }
        let expected_count = required_bucket_count(unique_string_count, actual_bucket_size)?;
        let actual_count = (data.len() - FIXED_PAYLOAD_LEN) / BUCKET_LEN;
        if actual_count != expected_count {
            return Err(invalid(format!(
                "ExtSST requires {expected_count} buckets, got {actual_count}"
            )));
        }

        let mut buckets = Vec::with_capacity(actual_count);
        for entry in data[FIXED_PAYLOAD_LEN..].chunks_exact(BUCKET_LEN) {
            let stream_position = u32::from_le_bytes(entry[0..4].try_into().unwrap());
            let record_offset = u16::from_le_bytes([entry[4], entry[5]]);
            let reserved = u16::from_le_bytes([entry[6], entry[7]]);
            if reserved != 0 {
                return Err(invalid("ISSTInf reserved field must be zero"));
            }
            buckets.push(XlsSharedStringBucket::try_new(
                stream_position,
                record_offset,
            )?);
        }
        Self::try_new(unique_string_count, buckets)
    }

    pub fn unique_string_count(&self) -> u32 {
        self.unique_string_count
    }
    pub fn strings_per_bucket(&self) -> u16 {
        self.strings_per_bucket
    }
    pub fn buckets(&self) -> &[XlsSharedStringBucket] {
        &self.buckets
    }

    /// Bucket containing the given zero-based shared-string index.
    pub fn bucket_for_string(&self, string_index: u32) -> Option<&XlsSharedStringBucket> {
        if string_index >= self.unique_string_count {
            return None;
        }
        let bucket = string_index / u32::from(self.strings_per_bucket);
        self.buckets.get(bucket as usize)
    }

    /// First shared-string index represented by a bucket.
    pub fn first_string_index(&self, bucket_index: usize) -> Option<u32> {
        if bucket_index >= self.buckets.len() {
            return None;
        }
        u32::try_from(bucket_index)
            .ok()?
            .checked_mul(u32::from(self.strings_per_bucket))
    }

    /// Serialize the complete `ExtSST` payload deterministically.
    pub fn to_payload(&self) -> Vec<u8> {
        let mut data = Vec::with_capacity(FIXED_PAYLOAD_LEN + self.buckets.len() * BUCKET_LEN);
        data.extend_from_slice(&self.strings_per_bucket.to_le_bytes());
        for bucket in &self.buckets {
            data.extend_from_slice(&bucket.stream_position.to_le_bytes());
            data.extend_from_slice(&bucket.record_offset.to_le_bytes());
            data.extend_from_slice(&0u16.to_le_bytes());
        }
        data
    }

    /// Serialize the complete BIFF record including its four-byte record header.
    pub fn to_record_bytes(&self) -> XlsResult<Vec<u8>> {
        let payload = self.to_payload();
        let length = u16::try_from(payload.len())
            .map_err(|_| invalid("ExtSST payload length exceeds BIFF u16"))?;
        let mut data = Vec::with_capacity(4 + payload.len());
        data.extend_from_slice(&EXT_SST_RECORD_TYPE.to_le_bytes());
        data.extend_from_slice(&length.to_le_bytes());
        data.extend_from_slice(&payload);
        Ok(data)
    }
}

pub(crate) struct SharedStringIndexCollector {
    unique_string_count: Option<u32>,
    value: XlsResult<Option<XlsSharedStringIndex>>,
    ext_sst_seen: bool,
}

impl SharedStringIndexCollector {
    pub(crate) fn new() -> Self {
        Self {
            unique_string_count: None,
            value: Ok(None),
            ext_sst_seen: false,
        }
    }

    /// Collect the optional accelerator without making workbook content depend on it.
    pub(crate) fn feed_record(&mut self, record_type: u16, data: &[u8]) {
        match record_type {
            SST_RECORD_TYPE if data.len() >= 8 => {
                self.unique_string_count = Some(u32::from_le_bytes(data[4..8].try_into().unwrap()));
            },
            EXT_SST_RECORD_TYPE => {
                if self.ext_sst_seen {
                    self.value = Err(invalid("duplicate ExtSST record"));
                    return;
                }
                self.ext_sst_seen = true;
                let Some(unique_string_count) = self.unique_string_count else {
                    self.value = Err(invalid("ExtSST appears before SST"));
                    return;
                };
                self.value =
                    XlsSharedStringIndex::parse_payload(unique_string_count, data).map(Some);
            },
            _ => {},
        }
    }

    pub(crate) fn finish(self) -> XlsResult<Option<XlsSharedStringIndex>> {
        self.value
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    const SIMPLE_REFERENCE: [u8; 10] = [0x08, 0x00, 0x8C, 0x05, 0x00, 0x00, 0x0C, 0x00, 0x00, 0x00];

    #[test]
    fn parses_and_round_trips_simple_poi_reference() {
        let index = XlsSharedStringIndex::parse_payload(1, &SIMPLE_REFERENCE).unwrap();
        assert_eq!(index.unique_string_count(), 1);
        assert_eq!(index.strings_per_bucket(), 8);
        assert_eq!(index.buckets().len(), 1);
        let bucket = index.buckets()[0];
        assert_eq!(bucket.stream_position(), 1_420);
        assert_eq!(bucket.record_offset(), 12);
        assert_eq!(bucket.record_position(), 1_408);
        assert_eq!(index.bucket_for_string(0), Some(&bucket));
        assert!(index.bucket_for_string(1).is_none());
        assert_eq!(index.first_string_index(0), Some(0));
        assert_eq!(index.to_payload(), SIMPLE_REFERENCE);
        assert_eq!(
            &index.to_record_bytes().unwrap()[0..4],
            &[0xFF, 0x00, 10, 0],
        );
    }

    #[test]
    fn constructs_spec_derived_multi_bucket_index() {
        let buckets = (0..128)
            .map(|index| XlsSharedStringBucket::try_new(100 + index * 16, 12).unwrap())
            .collect::<Vec<_>>();
        let index = XlsSharedStringIndex::try_new(1_144, buckets).unwrap();
        assert_eq!(index.strings_per_bucket(), 9);
        assert_eq!(index.buckets().len(), 128);
        assert_eq!(index.first_string_index(127), Some(1_143));
        assert!(index.bucket_for_string(1_143).is_some());
        assert!(index.bucket_for_string(1_144).is_none());
        assert_eq!(index.to_payload().len(), MAX_PAYLOAD_LEN);
    }

    #[test]
    fn rejects_bad_formula_lengths_pointers_reserved_and_duplicates() {
        assert!(XlsSharedStringIndex::parse_payload(1, &[]).is_err());
        assert!(XlsSharedStringIndex::parse_payload(1, &[8, 0, 0]).is_err());
        let mut bad = SIMPLE_REFERENCE;
        bad[0] = 9;
        assert!(XlsSharedStringIndex::parse_payload(1, &bad).is_err());
        assert!(XlsSharedStringIndex::parse_payload(9, &SIMPLE_REFERENCE).is_err());
        let mut bad = SIMPLE_REFERENCE;
        bad[8] = 1;
        assert!(XlsSharedStringIndex::parse_payload(1, &bad).is_err());
        let mut bad = SIMPLE_REFERENCE;
        bad[2..6].copy_from_slice(&12u32.to_le_bytes());
        assert!(XlsSharedStringIndex::parse_payload(1, &bad).is_err());
        assert!(XlsSharedStringBucket::try_new(10, 10).is_err());
        assert!(XlsSharedStringIndex::try_new(1, Vec::new()).is_err());
        assert!(XlsSharedStringIndex::try_new(u32::MAX, Vec::new()).is_err());

        let mut collector = SharedStringIndexCollector::new();
        collector.feed_record(EXT_SST_RECORD_TYPE, &SIMPLE_REFERENCE);
        assert!(collector.finish().is_err());
        let mut collector = SharedStringIndexCollector::new();
        collector.feed_record(SST_RECORD_TYPE, &[1, 0, 0, 0, 1, 0, 0, 0]);
        collector.feed_record(EXT_SST_RECORD_TYPE, &SIMPLE_REFERENCE);
        collector.feed_record(EXT_SST_RECORD_TYPE, &SIMPLE_REFERENCE);
        assert!(collector.finish().is_err());
    }
}
