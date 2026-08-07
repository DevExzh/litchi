use litchi_core::error::Result;

use super::types::{
    EMFPLUS_RECORD_HEADER_SIZE, ObjectId, ObjectRecordFlags, ObjectType, ParserLimits, RecordFlags,
    RecordHeader, RecordType, parse_error,
};

const RECORD_HEADER_SIZE_U32: u32 = 12;

/// A zero-copy view of one structurally valid EMF+ record.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct EmfPlusRecord<'a> {
    /// Offset from the beginning of the comment payload passed to the iterator.
    pub offset: usize,
    pub header: RecordHeader,
    /// Record-specific bytes, excluding the invariant 12-byte header.
    pub data: &'a [u8],
    bytes: &'a [u8],
}

impl<'a> EmfPlusRecord<'a> {
    /// The complete serialized record, including its header.
    #[must_use]
    pub const fn bytes(self) -> &'a [u8] {
        self.bytes
    }

    /// Decode an `EmfPlusObject` record and isolate its object-data fragment.
    pub fn object_fragment(self, limits: ParserLimits) -> Result<ObjectFragment<'a>> {
        if self.header.record_type != RecordType::Object {
            return Err(parse_error("record is not an EmfPlusObject record"));
        }
        let flags = ObjectRecordFlags::parse(self.header.flags, limits)?;
        let (total_object_size, data) = if flags.continued {
            if self.data.len() < 4 {
                return Err(parse_error("continued EmfPlusObject lacks TotalObjectSize"));
            }
            let total = read_u32(self.data, 0)?;
            (Some(total), &self.data[4..])
        } else {
            (None, self.data)
        };
        Ok(ObjectFragment {
            flags,
            total_object_size,
            data,
        })
    }
}

/// Object data carried by one `EmfPlusObject` record.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct ObjectFragment<'a> {
    pub flags: ObjectRecordFlags,
    /// Present when the continuation bit is set.
    pub total_object_size: Option<u32>,
    /// Fragment bytes with `TotalObjectSize`, when present, removed.
    pub data: &'a [u8],
}

/// Strict iterator over the EMF+ records stored in one comment payload.
///
/// After the first error the iterator is fused. Empty input is accepted here;
/// use [`EmfPlusStreamValidator`] when Header/EndOfFile semantics are required.
#[derive(Debug)]
pub struct EmfPlusRecordIter<'a> {
    data: &'a [u8],
    limits: ParserLimits,
    offset: usize,
    records_seen: usize,
    failed: bool,
}

impl<'a> EmfPlusRecordIter<'a> {
    pub fn new(data: &'a [u8], limits: ParserLimits) -> Result<Self> {
        limits.validate()?;
        if data.len() > limits.max_bytes {
            return Err(parse_error(format!(
                "EMF+ payload has {} bytes, exceeding limit {}",
                data.len(),
                limits.max_bytes
            )));
        }
        Ok(Self {
            data,
            limits,
            offset: 0,
            records_seen: 0,
            failed: false,
        })
    }

    #[must_use]
    pub const fn offset(&self) -> usize {
        self.offset
    }

    fn next_record(&mut self) -> Result<EmfPlusRecord<'a>> {
        if self.records_seen >= self.limits.max_records {
            return Err(parse_error(format!(
                "EMF+ record count exceeds limit {}",
                self.limits.max_records
            )));
        }

        let remaining = self
            .data
            .get(self.offset..)
            .ok_or_else(|| parse_error("EMF+ record offset is outside payload"))?;
        if remaining.len() < EMFPLUS_RECORD_HEADER_SIZE {
            return Err(parse_error(format!(
                "truncated EMF+ record header at offset {}: {} bytes remain",
                self.offset,
                remaining.len()
            )));
        }

        let raw_type = read_u16(remaining, 0)?;
        let record_type = RecordType::try_from(raw_type)?;
        let flags = RecordFlags::new(read_u16(remaining, 2)?);
        let size = read_u32(remaining, 4)?;
        let data_size = read_u32(remaining, 8)?;

        if size < RECORD_HEADER_SIZE_U32 {
            return Err(parse_error(format!(
                "EMF+ record at offset {} has undersized Size {size}",
                self.offset
            )));
        }
        if size % 4 != 0 || data_size % 4 != 0 {
            return Err(parse_error(format!(
                "EMF+ record at offset {} is not 32-bit aligned",
                self.offset
            )));
        }
        let expected_size = data_size
            .checked_add(RECORD_HEADER_SIZE_U32)
            .ok_or_else(|| parse_error("EMF+ DataSize plus header overflows u32"))?;
        if size != expected_size {
            return Err(parse_error(format!(
                "EMF+ record at offset {} has Size {size} but DataSize {data_size}",
                self.offset
            )));
        }

        let size_usize = usize::try_from(size)
            .map_err(|_| parse_error("EMF+ record Size does not fit usize"))?;
        let end = self
            .offset
            .checked_add(size_usize)
            .ok_or_else(|| parse_error("EMF+ record end offset overflow"))?;
        if end > self.data.len() {
            return Err(parse_error(format!(
                "EMF+ record at offset {} ends at {end}, beyond payload length {}",
                self.offset,
                self.data.len()
            )));
        }
        if size_usize > self.limits.max_bytes {
            return Err(parse_error(format!(
                "EMF+ record Size {size_usize} exceeds byte limit {}",
                self.limits.max_bytes
            )));
        }

        let data_start = self
            .offset
            .checked_add(EMFPLUS_RECORD_HEADER_SIZE)
            .ok_or_else(|| parse_error("EMF+ data offset overflow"))?;
        let bytes = self
            .data
            .get(self.offset..end)
            .ok_or_else(|| parse_error("invalid EMF+ record byte range"))?;
        let record_data = self
            .data
            .get(data_start..end)
            .ok_or_else(|| parse_error("invalid EMF+ record data range"))?;
        let record = EmfPlusRecord {
            offset: self.offset,
            header: RecordHeader {
                record_type,
                flags,
                size,
                data_size,
            },
            data: record_data,
            bytes,
        };

        self.offset = end;
        self.records_seen = self
            .records_seen
            .checked_add(1)
            .ok_or_else(|| parse_error("EMF+ record count overflow"))?;
        Ok(record)
    }
}

impl<'a> Iterator for EmfPlusRecordIter<'a> {
    type Item = Result<EmfPlusRecord<'a>>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.failed || self.offset == self.data.len() {
            return None;
        }
        let result = self.next_record();
        if result.is_err() {
            self.failed = true;
        }
        Some(result)
    }
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum StreamState {
    AwaitingHeader,
    Active,
    Ended,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
struct PendingObject {
    object_id: ObjectId,
    object_type: ObjectType,
    total_size: usize,
    received: usize,
}

/// Validates a logical EMF+ stream, including streams split over many EMF
/// comment records.
#[derive(Debug)]
pub struct EmfPlusStreamValidator {
    limits: ParserLimits,
    state: StreamState,
    pending_object: Option<PendingObject>,
    bytes_seen: usize,
    records_seen: usize,
}

impl EmfPlusStreamValidator {
    pub fn new(limits: ParserLimits) -> Result<Self> {
        Ok(Self {
            limits: limits.validate()?,
            state: StreamState::AwaitingHeader,
            pending_object: None,
            bytes_seen: 0,
            records_seen: 0,
        })
    }

    /// Frame and validate all records in one `EMR_COMMENT_EMFPLUS` payload.
    pub fn push_payload(&mut self, payload: &[u8]) -> Result<usize> {
        let mut count = 0usize;
        for record in EmfPlusRecordIter::new(payload, self.limits)? {
            self.push(record?)?;
            count = count
                .checked_add(1)
                .ok_or_else(|| parse_error("payload record count overflow"))?;
        }
        Ok(count)
    }

    /// Validate one already-framed record.
    pub fn push(&mut self, record: EmfPlusRecord<'_>) -> Result<()> {
        if self.state == StreamState::Ended {
            return Err(parse_error("EMF+ record appears after EndOfFile"));
        }

        let record_size = usize::try_from(record.header.size)
            .map_err(|_| parse_error("EMF+ record Size does not fit usize"))?;
        let next_bytes = self
            .bytes_seen
            .checked_add(record_size)
            .ok_or_else(|| parse_error("logical EMF+ stream byte count overflow"))?;
        if next_bytes > self.limits.max_bytes {
            return Err(parse_error(format!(
                "logical EMF+ stream exceeds byte limit {}",
                self.limits.max_bytes
            )));
        }
        let next_records = self
            .records_seen
            .checked_add(1)
            .ok_or_else(|| parse_error("logical EMF+ record count overflow"))?;
        if next_records > self.limits.max_records {
            return Err(parse_error(format!(
                "logical EMF+ stream exceeds record limit {}",
                self.limits.max_records
            )));
        }

        match self.state {
            StreamState::AwaitingHeader => self.validate_initial_header(record)?,
            StreamState::Active => self.validate_active_record(record)?,
            StreamState::Ended => {
                return Err(parse_error("EMF+ record appears after EndOfFile"));
            },
        }
        self.bytes_seen = next_bytes;
        self.records_seen = next_records;
        Ok(())
    }

    /// Finish a logical stream. An `EndOfFile` record and a completed object are
    /// required; a comment boundary alone is not end-of-stream.
    pub fn finish(&self) -> Result<()> {
        if self.pending_object.is_some() {
            return Err(parse_error("unterminated continued EmfPlusObject"));
        }
        match self.state {
            StreamState::Ended => Ok(()),
            StreamState::AwaitingHeader => Err(parse_error("empty EMF+ stream")),
            StreamState::Active => Err(parse_error("EMF+ stream lacks EndOfFile record")),
        }
    }

    #[must_use]
    pub const fn bytes_seen(&self) -> usize {
        self.bytes_seen
    }

    #[must_use]
    pub const fn records_seen(&self) -> usize {
        self.records_seen
    }

    fn validate_initial_header(&mut self, record: EmfPlusRecord<'_>) -> Result<()> {
        if record.header.record_type != RecordType::Header {
            return Err(parse_error("first EMF+ record is not Header"));
        }
        if record.header.data_size != 16 {
            return Err(parse_error("EmfPlusHeader DataSize must be 16"));
        }
        self.state = StreamState::Active;
        Ok(())
    }

    fn validate_active_record(&mut self, record: EmfPlusRecord<'_>) -> Result<()> {
        if record.header.record_type == RecordType::Header {
            return Err(parse_error("duplicate EmfPlusHeader record"));
        }
        if record.header.record_type.is_reserved() {
            return Err(parse_error(format!(
                "reserved EMF+ record type 0x{:04X} is forbidden",
                record.header.record_type.raw()
            )));
        }
        if self.pending_object.is_some() && record.header.record_type != RecordType::Object {
            return Err(parse_error(
                "continued EmfPlusObject was interrupted by another record type",
            ));
        }

        if record.header.record_type == RecordType::Object {
            return self.validate_object(record);
        }
        if record.header.record_type == RecordType::EndOfFile {
            return self.validate_end_of_file(record);
        }
        Ok(())
    }

    fn validate_end_of_file(&mut self, record: EmfPlusRecord<'_>) -> Result<()> {
        if record.header.data_size != 0 || record.header.flags.raw() != 0 {
            return Err(parse_error(
                "EmfPlusEndOfFile must have zero Flags and DataSize",
            ));
        }
        self.state = StreamState::Ended;
        Ok(())
    }

    fn validate_object(&mut self, record: EmfPlusRecord<'_>) -> Result<()> {
        let fragment = record.object_fragment(self.limits)?;
        match self.pending_object {
            None => self.start_or_complete_object(fragment),
            Some(pending) => self.continue_object(fragment, pending),
        }
    }

    fn start_or_complete_object(&mut self, fragment: ObjectFragment<'_>) -> Result<()> {
        if !fragment.flags.continued {
            return Ok(());
        }
        let total_u32 = fragment
            .total_object_size
            .ok_or_else(|| parse_error("continued object has no TotalObjectSize"))?;
        let total_size = usize::try_from(total_u32)
            .map_err(|_| parse_error("TotalObjectSize does not fit usize"))?;
        if total_size > self.limits.max_bytes {
            return Err(parse_error(format!(
                "TotalObjectSize {total_size} exceeds byte limit {}",
                self.limits.max_bytes
            )));
        }
        let received = fragment.data.len();
        if received > total_size {
            return Err(parse_error("continued object exceeds TotalObjectSize"));
        }
        if received == total_size {
            return Ok(());
        }
        self.pending_object = Some(PendingObject {
            object_id: fragment.flags.object_id,
            object_type: fragment.flags.object_type,
            total_size,
            received,
        });
        Ok(())
    }

    fn continue_object(
        &mut self,
        fragment: ObjectFragment<'_>,
        pending: PendingObject,
    ) -> Result<()> {
        if fragment.flags.object_id != pending.object_id
            || fragment.flags.object_type != pending.object_type
        {
            return Err(parse_error(
                "continued object changed its ObjectID or ObjectType",
            ));
        }
        if let Some(total_u32) = fragment.total_object_size {
            let total_size = usize::try_from(total_u32)
                .map_err(|_| parse_error("TotalObjectSize does not fit usize"))?;
            if total_size != pending.total_size {
                return Err(parse_error("continued object changed its TotalObjectSize"));
            }
        }
        let received = pending
            .received
            .checked_add(fragment.data.len())
            .ok_or_else(|| parse_error("continued object byte count overflow"))?;
        if received > pending.total_size {
            return Err(parse_error("continued object exceeds TotalObjectSize"));
        }

        if received == pending.total_size {
            self.pending_object = None;
        } else if fragment.flags.continued {
            self.pending_object = Some(PendingObject {
                received,
                ..pending
            });
        } else {
            return Err(parse_error(format!(
                "continued object ended at {received} bytes, expected {}",
                pending.total_size
            )));
        }
        Ok(())
    }
}

/// Validate a complete, contiguous logical EMF+ stream.
pub fn validate_complete_stream(data: &[u8], limits: ParserLimits) -> Result<()> {
    let mut validator = EmfPlusStreamValidator::new(limits)?;
    validator.push_payload(data)?;
    validator.finish()
}

fn read_u16(data: &[u8], offset: usize) -> Result<u16> {
    let end = offset
        .checked_add(2)
        .ok_or_else(|| parse_error("16-bit field offset overflow"))?;
    let bytes: [u8; 2] = data
        .get(offset..end)
        .ok_or_else(|| parse_error("truncated 16-bit field"))?
        .try_into()
        .map_err(|_| parse_error("invalid 16-bit field"))?;
    Ok(u16::from_le_bytes(bytes))
}

fn read_u32(data: &[u8], offset: usize) -> Result<u32> {
    let end = offset
        .checked_add(4)
        .ok_or_else(|| parse_error("32-bit field offset overflow"))?;
    let bytes: [u8; 4] = data
        .get(offset..end)
        .ok_or_else(|| parse_error("truncated 32-bit field"))?
        .try_into()
        .map_err(|_| parse_error("invalid 32-bit field"))?;
    Ok(u32::from_le_bytes(bytes))
}
