//! Logical RealTimeData record assembly and prefix reconstruction.

use super::codec::{
    CONTINUE_FRT_RECORD_TYPE, FRT_HEADER_LEN, MAX_LOGICAL_PAYLOAD_BYTES, MAX_STRING_CHARACTERS,
    MAX_TOPIC_SEGMENTS, MIN_PREFIXED_TOPIC_SEGMENTS, MIN_TOPIC_SEGMENTS, Payload,
    REAL_TIME_DATA_RECORD_TYPE, RTD_E_ITEM_LEN, RTD_OPER_BOOLEAN, RTD_OPER_ERROR, RTD_OPER_INTEGER,
    RTD_OPER_LONG_TEXT, RTD_OPER_NUMBER, RTD_OPER_SHORT_TEXT, biff_char_count, invalid,
    join_segments, parse_rtd_oper, parse_segmented_topic, read_u16, read_u32, write_chars,
    write_segmented_topic,
};
use super::model::{Cell, Record, UnknownRecord, Value};
use super::validation::{RecordSpan, frame_logical_payload, replace_range, scan_records};
use crate::error::{Error, Result};

#[derive(Debug, Clone, Copy)]
struct RealTimeDataSpan {
    first_record: usize,
    end_record: usize,
}

/// The workbook-global BIFF record owner for real-time data.
///
/// The owner keeps the complete record order and only interprets a
/// `RealTimeData` record together with its immediately following `ContinueFrt`
/// records.  Every other record remains an opaque source span, which lets the
/// detached editor preserve producer-specific records without normalizing the
/// surrounding workbook globals stream.
#[derive(Debug, Clone)]
pub(crate) struct Package {
    records: Vec<RecordSpan>,
    real_time_data: Vec<Record>,
    real_time_data_spans: Vec<RealTimeDataSpan>,
    unknown_records: Vec<usize>,
}

impl Package {
    pub(crate) fn parse(bytes: &[u8]) -> Result<Self> {
        let records = scan_records(bytes)?;
        let mut real_time_data = Vec::new();
        let mut real_time_data_spans = Vec::new();
        let mut unknown_records = Vec::new();
        let mut previous_topic = None;
        let mut record_index = 0usize;

        while record_index < records.len() {
            let record = records[record_index];
            if record.record_type != REAL_TIME_DATA_RECORD_TYPE {
                unknown_records.push(record_index);
                record_index += 1;
                continue;
            }

            let first_record = record_index;
            let mut end_record = record_index
                .checked_add(1)
                .ok_or_else(|| invalid("RealTimeData record span overflows"))?;
            let logical_len = records[first_record]
                .payload_end
                .checked_sub(records[first_record].payload_start)
                .ok_or_else(|| invalid("RealTimeData payload span is inverted"))?;
            let mut logical_payload = Vec::new();
            logical_payload
                .try_reserve(logical_len)
                .map_err(|_| Error::Allocation("assembling RealTimeData records"))?;
            logical_payload.extend_from_slice(record.payload(bytes));
            while let Some(continuation) = records.get(end_record) {
                if continuation.record_type != CONTINUE_FRT_RECORD_TYPE {
                    break;
                }
                let continuation_len = continuation
                    .payload_end
                    .checked_sub(continuation.payload_start)
                    .ok_or_else(|| invalid("ContinueFrt payload span is inverted"))?;
                logical_payload
                    .try_reserve(continuation_len)
                    .map_err(|_| Error::Allocation("assembling RealTimeData continuations"))?;
                logical_payload.extend_from_slice(continuation.payload(bytes));
                end_record = end_record
                    .checked_add(1)
                    .ok_or_else(|| invalid("RealTimeData continuation span overflows"))?;
            }

            let parsed = Record::parse(&logical_payload, previous_topic.as_deref())?;
            previous_topic = Some(parsed.topic.clone());
            real_time_data.push(parsed);
            real_time_data_spans.push(RealTimeDataSpan {
                first_record,
                end_record,
            });
            record_index = end_record;
        }

        Ok(Self {
            records,
            real_time_data,
            real_time_data_spans,
            unknown_records,
        })
    }

    pub(crate) fn real_time_data(&self) -> &[Record] {
        &self.real_time_data
    }

    pub(crate) fn unknown_record_count(&self) -> usize {
        self.unknown_records.len()
    }

    pub(crate) fn unknown_records<'a>(
        &'a self,
        bytes: &'a [u8],
    ) -> impl Iterator<Item = UnknownRecord<'a>> + 'a {
        self.unknown_records.iter().map(move |&record_index| {
            let record = self.records[record_index];
            UnknownRecord::new(record.record_type, record.payload(bytes))
        })
    }

    pub(crate) fn replace_real_time_data(
        &self,
        bytes: &[u8],
        index: usize,
        value: &Record,
    ) -> Result<Vec<u8>> {
        let span = self.real_time_data_spans.get(index).ok_or_else(|| {
            Error::UnsafeEdit(format!(
                "RealTimeData index {index} is outside the source collection"
            ))
        })?;
        let payload = value.to_payload()?;
        let replacement = frame_logical_payload(
            &payload,
            REAL_TIME_DATA_RECORD_TYPE,
            CONTINUE_FRT_RECORD_TYPE,
        )?;
        let start = self.records[span.first_record].record_start;
        let end = self.records[span.end_record - 1].payload_end;
        replace_range(bytes, start, end, &replacement)
    }

    pub(crate) fn insert_real_time_data(
        &self,
        bytes: &[u8],
        index: usize,
        value: &Record,
    ) -> Result<Vec<u8>> {
        if index > self.real_time_data.len() {
            return Err(Error::UnsafeEdit(format!(
                "RealTimeData insertion index {index} is outside the source collection"
            )));
        }
        let payload = value.to_payload()?;
        let replacement = frame_logical_payload(
            &payload,
            REAL_TIME_DATA_RECORD_TYPE,
            CONTINUE_FRT_RECORD_TYPE,
        )?;
        let start = if let Some(span) = self.real_time_data_spans.get(index) {
            self.records[span.first_record].record_start
        } else {
            bytes.len()
        };
        replace_range(bytes, start, start, &replacement)
    }

    pub(crate) fn remove_real_time_data(&self, bytes: &[u8], index: usize) -> Result<Vec<u8>> {
        let span = self.real_time_data_spans.get(index).ok_or_else(|| {
            Error::UnsafeEdit(format!(
                "RealTimeData index {index} is outside the source collection"
            ))
        })?;
        let start = self.records[span.first_record].record_start;
        let end = self.records[span.end_record - 1].payload_end;
        replace_range(bytes, start, end, &[])
    }
}

impl Record {
    /// Parse one logical `RealTimeData` payload: the record body with any
    /// `ContinueFrt` bodies already appended.
    ///
    /// `previous_topic` is the reconstructed [`Record::topic`] of
    /// the preceding `RealTimeData` record in the globals substream, needed
    /// to re-apply prefix compression; pass `None` for the first record.
    pub fn parse(data: &[u8], previous_topic: Option<&str>) -> Result<Self> {
        if data.len() > MAX_LOGICAL_PAYLOAD_BYTES {
            return Err(invalid(format!(
                "RealTimeData payload exceeds {MAX_LOGICAL_PAYLOAD_BYTES} bytes"
            )));
        }
        if data.len() < FRT_HEADER_LEN + 4 {
            return Err(Error::InvalidLength {
                expected: FRT_HEADER_LEN + 4,
                found: data.len(),
            });
        }
        if read_u16(data, 0)? != REAL_TIME_DATA_RECORD_TYPE {
            return Err(invalid("RealTimeData FrtHeader.rt mismatch"));
        }

        let common_prefix_len = read_u32(data, FRT_HEADER_LEN)?;
        let mut offset = FRT_HEADER_LEN + 4;

        // stTopic: XLUnicodeStringSegmentedRTD (MS-XLS 2.5.298).
        let (topic_segments, used) = parse_segmented_topic(
            data.get(offset..).ok_or(Error::InvalidLength {
                expected: offset,
                found: data.len(),
            })?,
            common_prefix_len != 0,
        )?;
        offset = offset
            .checked_add(used)
            .ok_or_else(|| invalid("RealTimeData topic offset overflows usize"))?;

        // rtdOper: RTDOper (MS-XLS 2.5.224).
        let (value, used) = parse_rtd_oper(data.get(offset..).ok_or(Error::InvalidLength {
            expected: offset,
            found: data.len(),
        })?)?;
        offset = offset
            .checked_add(used)
            .ok_or_else(|| invalid("RealTimeData value offset overflows usize"))?;

        // rgRTDE: the rest of the payload in 6-byte RTDEItem entries.
        let remaining = data.get(offset..).ok_or(Error::InvalidLength {
            expected: offset,
            found: data.len(),
        })?;
        if !remaining.len().is_multiple_of(RTD_E_ITEM_LEN) {
            return Err(invalid("RealTimeData rgRTDE size is not a multiple of 6"));
        }
        let mut cells = Vec::new();
        cells
            .try_reserve_exact(remaining.len() / RTD_E_ITEM_LEN)
            .map_err(|_| Error::Allocation("retaining RTD subscriber cells"))?;
        for chunk in remaining.chunks_exact(RTD_E_ITEM_LEN) {
            let column = u8::try_from(read_u16(chunk, 2)?)
                .map_err(|_| invalid("RTD subscriber column exceeds the BIFF8 grid"))?;
            cells.push(Cell {
                row: read_u16(chunk, 0)?,
                column,
                sheet_index: read_u16(chunk, 4)?,
            });
        }

        // Re-apply prefix compression against the previous topic.
        let stored = join_segments(&topic_segments)?;
        let topic = if common_prefix_len == 0 {
            stored
        } else {
            let previous = previous_topic
                .ok_or_else(|| invalid("first RealTimeData record declares a shared prefix"))?;
            let prefix_len = usize::try_from(common_prefix_len)
                .map_err(|_| invalid("RealTimeData ichSamePrefix overflows"))?;
            if prefix_len > previous.chars().count() {
                return Err(invalid(
                    "RealTimeData ichSamePrefix exceeds the previous topic",
                ));
            }
            let capacity = prefix_len
                .checked_add(stored.len())
                .ok_or_else(|| invalid("reconstructed RealTimeData topic length overflows"))?;
            let mut topic = String::new();
            topic
                .try_reserve(capacity)
                .map_err(|_| Error::Allocation("reconstructing RealTimeData topic"))?;
            topic.extend(previous.chars().take(prefix_len));
            topic.push_str(&stored);
            topic
        };

        Ok(Record {
            common_prefix_len,
            topic_segments,
            topic,
            value,
            cells,
        })
    }

    /// Serialize back to a complete logical `RealTimeData` payload (the
    /// record body; the writer chunks it into `ContinueFrt` records when it
    /// exceeds the maximum record size).
    ///
    /// The stored topic sub-strings are written as-is, so a value parsed
    /// from a workbook round-trips exactly; `topic` is re-derived from
    /// `common_prefix_len` and the previous record on the next parse.
    pub(crate) fn to_payload(&self) -> Result<Vec<u8>> {
        let common_prefix_len = usize::try_from(self.common_prefix_len)
            .map_err(|_| invalid("RTD common prefix does not fit in usize"))?;
        if common_prefix_len > MAX_STRING_CHARACTERS {
            return Err(invalid(
                "RTD common prefix exceeds the string resource limit",
            ));
        }
        let minimum_segments = if common_prefix_len == 0 {
            MIN_TOPIC_SEGMENTS
        } else {
            MIN_PREFIXED_TOPIC_SEGMENTS
        };
        if !(minimum_segments..=MAX_TOPIC_SEGMENTS).contains(&self.topic_segments.len()) {
            return Err(invalid(format!(
                "RTD topic must have {minimum_segments}..={MAX_TOPIC_SEGMENTS} segments"
            )));
        }
        if self.cells.len() > MAX_LOGICAL_PAYLOAD_BYTES / RTD_E_ITEM_LEN {
            return Err(invalid(
                "RTD subscriber cell count exceeds the resource limit",
            ));
        }
        let mut payload = Payload::new();
        payload.extend_from_slice(&REAL_TIME_DATA_RECORD_TYPE.to_le_bytes())?;
        payload.extend_from_slice(&[0u8; FRT_HEADER_LEN - 2])?; // grbitFrt + reserved
        payload.extend_from_slice(&self.common_prefix_len.to_le_bytes())?;
        write_segmented_topic(&mut payload, &self.topic_segments, minimum_segments)?;
        match &self.value {
            Value::Number(value) => {
                payload.extend_from_slice(&RTD_OPER_NUMBER.to_le_bytes())?;
                payload.extend_from_slice(&value.to_le_bytes())?;
            },
            Value::Text(text) => {
                let char_count = biff_char_count(text);
                if char_count > MAX_STRING_CHARACTERS {
                    return Err(invalid("RTD text exceeds the string resource limit"));
                }
                let char_count = u32::try_from(char_count)
                    .map_err(|_| invalid("RTD text character count overflows u32"))?;
                let kind = if char_count < 256 {
                    RTD_OPER_SHORT_TEXT
                } else {
                    RTD_OPER_LONG_TEXT
                };
                payload.extend_from_slice(&kind.to_le_bytes())?;
                payload.extend_from_slice(&char_count.to_le_bytes())?;
                write_chars(&mut payload, text)?;
            },
            Value::Boolean(value) => {
                payload.extend_from_slice(&RTD_OPER_BOOLEAN.to_le_bytes())?;
                payload.extend_from_slice(&u32::from(*value).to_le_bytes())?;
            },
            Value::Error(value) => {
                payload.extend_from_slice(&RTD_OPER_ERROR.to_le_bytes())?;
                payload.extend_from_slice(&value.to_le_bytes())?;
            },
            Value::Integer(value) => {
                payload.extend_from_slice(&RTD_OPER_INTEGER.to_le_bytes())?;
                payload.extend_from_slice(&value.to_le_bytes())?;
            },
        }
        for cell in &self.cells {
            payload.extend_from_slice(&cell.row.to_le_bytes())?;
            payload.extend_from_slice(&u16::from(cell.column).to_le_bytes())?;
            payload.extend_from_slice(&cell.sheet_index.to_le_bytes())?;
        }
        Ok(payload.into_vec())
    }
}
