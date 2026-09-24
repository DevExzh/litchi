//! Bounded BIFF12 codec for the XLSB Volatile Dependencies part.

use std::char;

use crate::package::error::{Error, Result};
use crate::raw::{Cursor, Records, Writer, kind};

use super::model::{
    CachedValue, CellReference, Dependencies, DependencyKind, ErrorCode, MainTopic, ReadLimits,
    Topic, VolatileType,
};

/// Result of parsing one Volatile Dependencies stream.
#[derive(Debug, Clone, PartialEq)]
pub(crate) struct Parsed {
    /// Typed known records.
    pub dependencies: Dependencies,
    /// Whether any record was not understood by this version of the codec.
    pub rewrite_safe: bool,
}

/// Parse one complete Volatile Dependencies stream and bind references to the
/// Workbook sheet catalog cardinality.
pub(crate) fn read(data: &[u8], limits: ReadLimits, sheet_count: usize) -> Result<Parsed> {
    let limits = limits.validate()?;
    if data.is_empty() {
        return Err(invalid("Volatile Dependencies stream is empty"));
    }
    if data.len() > limits.max_part_bytes {
        return Err(limit(
            "Volatile Dependencies part bytes",
            data.len(),
            limits.max_part_bytes,
        ));
    }
    let raw_limits = crate::raw::Limits::new(limits.max_part_bytes, limits.max_string_units);
    let mut iterator = Records::try_with_limits(data, raw_limits)?;
    let mut parser = Parser::new(limits);
    while let Some(record) = iterator.next() {
        let record = record?;
        parser.records = parser
            .records
            .checked_add(1)
            .ok_or(Error::CapacityOverflow {
                resource: "Volatile Dependencies records",
            })?;
        if parser.records > limits.max_records {
            return Err(limit(
                "Volatile Dependencies records",
                parser.records,
                limits.max_records,
            ));
        }
        parser.accept(record.kind(), record.payload())?;
    }
    parser.finish(sheet_count)
}

struct Parser {
    limits: ReadLimits,
    records: usize,
    strings: usize,
    mains: usize,
    topics: usize,
    subtopics: usize,
    references: usize,
    dependencies: Vec<VolatileType>,
    current_type: Option<VolatileType>,
    current_main: Option<MainTopic>,
    current_topic: Option<Topic>,
    value_seen: bool,
    began: bool,
    ended: bool,
    rewrite_safe: bool,
}

impl Parser {
    fn new(limits: ReadLimits) -> Self {
        Self {
            limits,
            records: 0,
            strings: 0,
            mains: 0,
            topics: 0,
            subtopics: 0,
            references: 0,
            dependencies: Vec::new(),
            current_type: None,
            current_main: None,
            current_topic: None,
            value_seen: false,
            began: false,
            ended: false,
            rewrite_safe: true,
        }
    }

    fn accept(&mut self, record_kind: crate::raw::Kind, payload: &[u8]) -> Result<()> {
        if self.ended {
            if is_known(record_kind) {
                return Err(invalid("record follows BrtEndVolDeps"));
            }
            self.rewrite_safe = false;
            return Ok(());
        }
        match record_kind {
            kind if kind == kind::BEGIN_VOL_DEPS => self.begin_deps(payload),
            kind if kind == kind::END_VOL_DEPS => self.end_deps(payload),
            kind if kind == kind::BEGIN_VOL_TYPE => self.begin_type(payload),
            kind if kind == kind::END_VOL_TYPE => self.end_type(payload),
            kind if kind == kind::BEGIN_VOL_MAIN => self.begin_main(payload),
            kind if kind == kind::END_VOL_MAIN => self.end_main(payload),
            kind if kind == kind::BEGIN_VOL_TOPIC => self.begin_topic(payload),
            kind if kind == kind::END_VOL_TOPIC => self.end_topic(payload),
            kind if kind == kind::VOL_SUBTOPIC => self.subtopic(payload),
            kind if kind == kind::VOL_REF => self.reference(payload),
            kind if kind == kind::VOL_NUM
                || kind == kind::VOL_ERR
                || kind == kind::VOL_STR
                || kind == kind::VOL_BOOL =>
            {
                self.value(record_kind, payload)
            },
            _ => {
                self.rewrite_safe = false;
                Ok(())
            },
        }
    }

    fn begin_deps(&mut self, payload: &[u8]) -> Result<()> {
        empty(payload, "BrtBeginVolDeps")?;
        if self.began {
            return Err(invalid("duplicate BrtBeginVolDeps"));
        }
        if !self.dependencies.is_empty()
            || self.current_type.is_some()
            || self.current_main.is_some()
            || self.current_topic.is_some()
        {
            return Err(invalid("BrtBeginVolDeps is not at stream start"));
        }
        self.began = true;
        self.dependencies
            .try_reserve_exact(self.limits.max_types.min(2))
            .map_err(|source| Error::Allocation {
                resource: "Volatile Dependencies type collections",
                source,
            })?;
        Ok(())
    }

    fn end_deps(&mut self, payload: &[u8]) -> Result<()> {
        empty(payload, "BrtEndVolDeps")?;
        if !self.began
            || self.ended
            || self.current_type.is_some()
            || self.current_main.is_some()
            || self.current_topic.is_some()
        {
            return Err(invalid("unbalanced BrtEndVolDeps"));
        }
        self.ended = true;
        Ok(())
    }

    fn begin_type(&mut self, payload: &[u8]) -> Result<()> {
        if !self.began
            || self.ended
            || self.current_type.is_some()
            || self.current_main.is_some()
            || self.current_topic.is_some()
        {
            return Err(invalid("unbalanced BrtBeginVolType"));
        }
        if self.dependencies.len() >= self.limits.max_types {
            return Err(limit(
                "Volatile Dependencies type collections",
                self.dependencies.len() + 1,
                self.limits.max_types,
            ));
        }
        let mut cursor = Cursor::new(payload, "BrtBeginVolType");
        let flags = cursor.read_u32()?;
        // The low bit selects the dependency kind.  The remaining bits are
        // reserved by the record grammar and must not be interpreted.  Keep
        // the source opaque when they are present so an exact no-op/removal
        // remains possible while a typed rewrite is refused.
        let has_reserved_bits = flags & !1 != 0;
        let kind = DependencyKind::from_wire(flags & 1)?;
        if self.dependencies.iter().any(|value| value.kind == kind) {
            return Err(invalid("duplicate volatile dependency type"));
        }
        cursor.finish()?;
        if has_reserved_bits {
            self.rewrite_safe = false;
        }
        self.current_type = Some(VolatileType {
            kind,
            mains: Vec::new(),
        });
        Ok(())
    }

    fn end_type(&mut self, payload: &[u8]) -> Result<()> {
        empty(payload, "BrtEndVolType")?;
        if self.current_main.is_some() || self.current_topic.is_some() {
            return Err(invalid("BrtEndVolType closes an open child"));
        }
        let value = self
            .current_type
            .take()
            .ok_or_else(|| invalid("BrtEndVolType has no matching begin"))?;
        if self.dependencies.len() >= self.limits.max_types {
            return Err(limit(
                "Volatile Dependencies type collections",
                self.dependencies.len() + 1,
                self.limits.max_types,
            ));
        }
        self.dependencies.push(value);
        Ok(())
    }

    fn begin_main(&mut self, payload: &[u8]) -> Result<()> {
        if self.current_type.is_none()
            || self.current_main.is_some()
            || self.current_topic.is_some()
        {
            return Err(invalid("unbalanced BrtBeginVolMain"));
        }
        self.mains = self.mains.checked_add(1).ok_or(Error::CapacityOverflow {
            resource: "Volatile main collections",
        })?;
        if self.mains > self.limits.max_mains {
            return Err(limit(
                "Volatile main collections",
                self.mains,
                self.limits.max_mains,
            ));
        }
        self.current_type
            .as_mut()
            .ok_or_else(|| invalid("BrtBeginVolMain has no enclosing type"))?
            .mains
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: "Volatile main collections",
                source,
            })?;
        let first = self.read_string(payload, "BrtBeginVolMain.first")?;
        self.current_main = Some(MainTopic {
            first,
            topics: Vec::new(),
        });
        Ok(())
    }

    fn end_main(&mut self, payload: &[u8]) -> Result<()> {
        empty(payload, "BrtEndVolMain")?;
        if self.current_topic.is_some() {
            return Err(invalid("BrtEndVolMain closes an open topic"));
        }
        let main = self
            .current_main
            .take()
            .ok_or_else(|| invalid("BrtEndVolMain has no matching begin"))?;
        let current_type = self
            .current_type
            .as_mut()
            .ok_or_else(|| invalid("BrtEndVolMain has no enclosing type"))?;
        if current_type.mains.len() >= self.limits.max_mains {
            return Err(limit(
                "Volatile main collections",
                current_type.mains.len() + 1,
                self.limits.max_mains,
            ));
        }
        current_type.mains.push(main);
        Ok(())
    }

    fn begin_topic(&mut self, payload: &[u8]) -> Result<()> {
        empty(payload, "BrtBeginVolTopic")?;
        if self.current_main.is_none() || self.current_topic.is_some() {
            return Err(invalid("unbalanced BrtBeginVolTopic"));
        }
        self.topics = self.topics.checked_add(1).ok_or(Error::CapacityOverflow {
            resource: "Volatile topic collections",
        })?;
        if self.topics > self.limits.max_topics {
            return Err(limit(
                "Volatile topic collections",
                self.topics,
                self.limits.max_topics,
            ));
        }
        self.current_main
            .as_mut()
            .ok_or_else(|| invalid("BrtBeginVolTopic has no enclosing main"))?
            .topics
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: "Volatile topic collections",
                source,
            })?;
        self.current_topic = Some(Topic {
            subtopics: Vec::new(),
            value: CachedValue::Bool(false),
            references: Vec::new(),
        });
        self.value_seen = false;
        Ok(())
    }

    fn end_topic(&mut self, payload: &[u8]) -> Result<()> {
        empty(payload, "BrtEndVolTopic")?;
        if !self.value_seen {
            return Err(invalid(
                "BrtEndVolTopic must contain exactly one cached value",
            ));
        }
        let topic = self
            .current_topic
            .take()
            .ok_or_else(|| invalid("BrtEndVolTopic has no matching begin"))?;
        let main = self
            .current_main
            .as_mut()
            .ok_or_else(|| invalid("BrtEndVolTopic has no enclosing main"))?;
        if main.topics.len() >= self.limits.max_topics {
            return Err(limit(
                "Volatile topic collections",
                main.topics.len() + 1,
                self.limits.max_topics,
            ));
        }
        main.topics.push(topic);
        self.value_seen = false;
        Ok(())
    }

    fn subtopic(&mut self, payload: &[u8]) -> Result<()> {
        if self.current_topic.is_none() || self.value_seen {
            return Err(invalid(
                "BrtVolSubtopic must occur before the cached topic value",
            ));
        }
        self.subtopics = self
            .subtopics
            .checked_add(1)
            .ok_or(Error::CapacityOverflow {
                resource: "Volatile subtopic records",
            })?;
        if self.subtopics > self.limits.max_subtopics {
            return Err(limit(
                "Volatile subtopic records",
                self.subtopics,
                self.limits.max_subtopics,
            ));
        }
        self.current_topic
            .as_mut()
            .ok_or_else(|| invalid("BrtVolSubtopic has no enclosing topic"))?
            .subtopics
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: "Volatile subtopic records",
                source,
            })?;
        let value = self.read_string(payload, "BrtVolSubtopic")?;
        let topic = self
            .current_topic
            .as_mut()
            .ok_or_else(|| invalid("BrtVolSubtopic has no enclosing topic"))?;
        if topic.subtopics.len() >= self.limits.max_subtopics {
            return Err(limit(
                "Volatile subtopic records",
                topic.subtopics.len() + 1,
                self.limits.max_subtopics,
            ));
        }
        topic.subtopics.push(value);
        Ok(())
    }

    fn value(&mut self, record_kind: crate::raw::Kind, payload: &[u8]) -> Result<()> {
        if self.current_topic.is_none() || self.value_seen {
            return Err(invalid(
                "volatile topic must contain exactly one cached value",
            ));
        }
        let value = match record_kind {
            value_kind if value_kind == kind::VOL_NUM => {
                let mut cursor = Cursor::new(payload, "BrtVolNum");
                let value = cursor.read_f64()?;
                super::model::validate_xnum(value)?;
                cursor.finish()?;
                CachedValue::Number(value)
            },
            value_kind if value_kind == kind::VOL_ERR => {
                let mut cursor = Cursor::new(payload, "BrtVolErr");
                let value = ErrorCode::from_wire(cursor.read_u8()?)?;
                cursor.finish()?;
                CachedValue::Error(value)
            },
            value_kind if value_kind == kind::VOL_STR => {
                CachedValue::String(self.read_string(payload, "BrtVolStr")?)
            },
            value_kind if value_kind == kind::VOL_BOOL => {
                let mut cursor = Cursor::new(payload, "BrtVolBool");
                let value = cursor.read_bool8()?;
                cursor.finish()?;
                CachedValue::Bool(value)
            },
            _ => return Err(invalid("unknown cached volatile value record")),
        };
        let topic = self
            .current_topic
            .as_mut()
            .ok_or_else(|| invalid("cached value has no enclosing topic"))?;
        topic.value = value;
        self.value_seen = true;
        Ok(())
    }

    fn reference(&mut self, payload: &[u8]) -> Result<()> {
        if self.current_topic.is_none() || !self.value_seen {
            return Err(invalid("BrtVolRef must follow the topic cached value"));
        }
        self.references = self
            .references
            .checked_add(1)
            .ok_or(Error::CapacityOverflow {
                resource: "Volatile cell references",
            })?;
        if self.references > self.limits.max_references {
            return Err(limit(
                "Volatile cell references",
                self.references,
                self.limits.max_references,
            ));
        }
        self.current_topic
            .as_mut()
            .ok_or_else(|| invalid("BrtVolRef has no enclosing topic"))?
            .references
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: "Volatile cell references",
                source,
            })?;
        let mut cursor = Cursor::new(payload, "BrtVolRef");
        let row = cursor.read_i32()?;
        let column = cursor.read_i32()?;
        let sheet_index = cursor.read_u32()?;
        cursor.finish()?;
        if row < 0 || column < 0 {
            return Err(invalid("BrtVolRef row and column must be non-negative"));
        }
        let row = u32::try_from(row).map_err(|_error| invalid("BrtVolRef row is negative"))?;
        let column =
            u32::try_from(column).map_err(|_error| invalid("BrtVolRef column is negative"))?;
        let reference = CellReference::new(row, column, sheet_index)?;
        let topic = self
            .current_topic
            .as_mut()
            .ok_or_else(|| invalid("BrtVolRef has no enclosing topic"))?;
        if topic.references.len() >= self.limits.max_references {
            return Err(limit(
                "Volatile cell references",
                topic.references.len() + 1,
                self.limits.max_references,
            ));
        }
        topic.references.push(reference);
        Ok(())
    }

    fn read_string(&mut self, payload: &[u8], context: &'static str) -> Result<String> {
        let mut cursor = Cursor::new(payload, context);
        let units =
            usize::try_from(cursor.read_u32()?).map_err(|_error| Error::CapacityOverflow {
                resource: "Volatile Dependencies UTF-16 string",
            })?;
        if units > self.limits.max_string_units {
            return Err(limit(
                "Volatile UTF-16 string",
                units,
                self.limits.max_string_units,
            ));
        }
        self.strings = self
            .strings
            .checked_add(units)
            .ok_or(Error::CapacityOverflow {
                resource: "Volatile Dependencies UTF-16 strings",
            })?;
        if self.strings > self.limits.max_total_string_units {
            return Err(limit(
                "Volatile UTF-16 strings",
                self.strings,
                self.limits.max_total_string_units,
            ));
        }
        let byte_len = units.checked_mul(2).ok_or(Error::CapacityOverflow {
            resource: "Volatile Dependencies UTF-16 string",
        })?;
        let bytes = cursor.read_bytes(byte_len)?;
        let mut value = String::new();
        let utf8_capacity = units.checked_mul(3).ok_or(Error::CapacityOverflow {
            resource: "Volatile Dependencies UTF-8 string",
        })?;
        value
            .try_reserve_exact(utf8_capacity)
            .map_err(|source| Error::Allocation {
                resource: "Volatile Dependencies UTF-16 string",
                source,
            })?;
        for decoded in char::decode_utf16(
            bytes
                .chunks_exact(2)
                .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]])),
        ) {
            value.push(
                decoded.map_err(|_error| {
                    Error::InvalidFormat(format!("invalid UTF-16 in {context}"))
                })?,
            );
        }
        cursor.finish()?;
        Ok(value)
    }

    fn finish(self, sheet_count: usize) -> Result<Parsed> {
        if !self.began
            || !self.ended
            || self.current_type.is_some()
            || self.current_main.is_some()
            || self.current_topic.is_some()
        {
            return Err(Error::UnexpectedEndOfStream(
                "Volatile Dependencies record hierarchy".to_string(),
            ));
        }
        let dependencies = Dependencies {
            types: self.dependencies,
            has_unsupported_records: !self.rewrite_safe,
        };
        dependencies.validate(self.limits)?;
        dependencies.validate_sheet_references(sheet_count)?;
        Ok(Parsed {
            dependencies,
            rewrite_safe: self.rewrite_safe,
        })
    }
}

/// Serialize a complete typed stream. Unsupported source records are refused
/// because their original positions cannot be proven from the typed model.
pub(crate) fn write(
    value: &Dependencies,
    limits: ReadLimits,
    sheet_count: usize,
) -> Result<Vec<u8>> {
    let limits = limits.validate()?;
    if value.has_unsupported_records {
        return Err(Error::UnsupportedFeature(
            "typed Volatile Dependencies edits cannot rewrite a stream containing unsupported records".to_string(),
        ));
    }
    value.validate(limits)?;
    value.validate_sheet_references(sheet_count)?;
    let record_count = encoded_record_count(value)?;
    if record_count > limits.max_records {
        return Err(limit(
            "Volatile Dependencies records",
            record_count,
            limits.max_records,
        ));
    }
    let capacity = encoded_size(value)?;
    if capacity > limits.max_part_bytes {
        return Err(limit(
            "Volatile Dependencies output bytes",
            capacity,
            limits.max_part_bytes,
        ));
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(capacity)
        .map_err(|source| Error::Allocation {
            resource: "Volatile Dependencies output",
            source,
        })?;
    let mut writer = Writer::new(&mut output);
    writer.write_record(kind::BEGIN_VOL_DEPS, &[])?;
    for volatile_type in &value.types {
        writer.write_record(
            kind::BEGIN_VOL_TYPE,
            &volatile_type.kind.wire().to_le_bytes(),
        )?;
        for main in &volatile_type.mains {
            writer.write_record(
                kind::BEGIN_VOL_MAIN,
                &wide_string_payload(&main.first, limits)?,
            )?;
            for topic in &main.topics {
                writer.write_record(kind::BEGIN_VOL_TOPIC, &[])?;
                for subtopic in &topic.subtopics {
                    writer.write_record(
                        kind::VOL_SUBTOPIC,
                        &wide_string_payload(subtopic, limits)?,
                    )?;
                }
                match &topic.value {
                    CachedValue::Number(value) => {
                        writer.write_record(kind::VOL_NUM, &value.to_le_bytes())?
                    },
                    CachedValue::Error(value) => {
                        writer.write_record(kind::VOL_ERR, &[value.wire()])?
                    },
                    CachedValue::String(value) => {
                        writer.write_record(kind::VOL_STR, &wide_string_payload(value, limits)?)?
                    },
                    CachedValue::Bool(value) => {
                        writer.write_record(kind::VOL_BOOL, &[u8::from(*value)])?
                    },
                }
                for reference in &topic.references {
                    let mut payload = [0u8; 12];
                    let row = i32::try_from(reference.row()).map_err(|_error| {
                        Error::CapacityOverflow {
                            resource: "Volatile Dependencies row",
                        }
                    })?;
                    let column = i32::try_from(reference.column()).map_err(|_error| {
                        Error::CapacityOverflow {
                            resource: "Volatile Dependencies column",
                        }
                    })?;
                    payload[..4].copy_from_slice(&row.to_le_bytes());
                    payload[4..8].copy_from_slice(&column.to_le_bytes());
                    payload[8..].copy_from_slice(&reference.sheet_index().to_le_bytes());
                    writer.write_record(kind::VOL_REF, &payload)?;
                }
                writer.write_record(kind::END_VOL_TOPIC, &[])?;
            }
            writer.write_record(kind::END_VOL_MAIN, &[])?;
        }
        writer.write_record(kind::END_VOL_TYPE, &[])?;
    }
    writer.write_record(kind::END_VOL_DEPS, &[])?;
    if output.len() != capacity {
        return Err(Error::InvalidFormat(
            "Volatile Dependencies writer size preflight disagreed with output".to_string(),
        ));
    }
    Ok(output)
}

fn wide_string_payload(value: &str, limits: ReadLimits) -> Result<Vec<u8>> {
    let units = value.encode_utf16().count();
    if units > limits.max_string_units {
        return Err(limit(
            "Volatile UTF-16 string",
            units,
            limits.max_string_units,
        ));
    }
    let byte_len = units
        .checked_mul(2)
        .and_then(|value| value.checked_add(4))
        .ok_or(Error::CapacityOverflow {
            resource: "Volatile Dependencies UTF-16 string",
        })?;
    let mut payload = Vec::new();
    payload
        .try_reserve_exact(byte_len)
        .map_err(|source| Error::Allocation {
            resource: "Volatile Dependencies UTF-16 string",
            source,
        })?;
    payload.extend_from_slice(
        &(u32::try_from(units).map_err(|_error| Error::CapacityOverflow {
            resource: "Volatile Dependencies UTF-16 string",
        })?)
        .to_le_bytes(),
    );
    for unit in value.encode_utf16() {
        payload.extend_from_slice(&unit.to_le_bytes());
    }
    Ok(payload)
}

fn encoded_size(value: &Dependencies) -> Result<usize> {
    let mut size = record_size(kind::BEGIN_VOL_DEPS, 0)?;
    for volatile_type in &value.types {
        size = add_size(size, record_size(kind::BEGIN_VOL_TYPE, 4)?)?;
        for main in &volatile_type.mains {
            size = add_size(
                size,
                record_size(kind::BEGIN_VOL_MAIN, string_payload_len(&main.first)?)?,
            )?;
            for topic in &main.topics {
                size = add_size(size, record_size(kind::BEGIN_VOL_TOPIC, 0)?)?;
                for subtopic in &topic.subtopics {
                    size = add_size(
                        size,
                        record_size(kind::VOL_SUBTOPIC, string_payload_len(subtopic)?)?,
                    )?;
                }
                let value_size = match &topic.value {
                    CachedValue::Number(_) => 8,
                    CachedValue::Error(_) | CachedValue::Bool(_) => 1,
                    CachedValue::String(value) => string_payload_len(value)?,
                };
                let value_kind = match &topic.value {
                    CachedValue::Number(_) => kind::VOL_NUM,
                    CachedValue::Error(_) => kind::VOL_ERR,
                    CachedValue::String(_) => kind::VOL_STR,
                    CachedValue::Bool(_) => kind::VOL_BOOL,
                };
                size = add_size(size, record_size(value_kind, value_size)?)?;
                for _ in &topic.references {
                    size = add_size(size, record_size(kind::VOL_REF, 12)?)?;
                }
                size = add_size(size, record_size(kind::END_VOL_TOPIC, 0)?)?;
            }
            size = add_size(size, record_size(kind::END_VOL_MAIN, 0)?)?;
        }
        size = add_size(size, record_size(kind::END_VOL_TYPE, 0)?)?;
    }
    add_size(size, record_size(kind::END_VOL_DEPS, 0)?)
}

fn encoded_record_count(value: &Dependencies) -> Result<usize> {
    // The enclosing stream always has its begin/end pair.  Count every
    // record before allocating an output buffer so a caller record quota is
    // enforced even for an otherwise empty stream.
    let mut count = 2usize;
    for volatile_type in &value.types {
        count = add_record_count(count, 2)?; // begin/end type
        for main in &volatile_type.mains {
            count = add_record_count(count, 2)?; // begin/end main
            for topic in &main.topics {
                count = add_record_count(count, 3)?; // begin, value, end
                count = add_record_count(count, topic.subtopics.len())?;
                count = add_record_count(count, topic.references.len())?;
            }
        }
    }
    Ok(count)
}

fn string_payload_len(value: &str) -> Result<usize> {
    value
        .encode_utf16()
        .count()
        .checked_mul(2)
        .and_then(|value| value.checked_add(4))
        .ok_or(Error::CapacityOverflow {
            resource: "Volatile Dependencies output",
        })
}

fn record_size(kind_value: crate::raw::Kind, payload: usize) -> Result<usize> {
    let kind_bytes: usize = if kind_value.get() < 0x80 { 1 } else { 2 };
    let length_bytes: usize = if payload < 0x80 {
        1
    } else if payload < 0x4000 {
        2
    } else if payload < 0x20_0000 {
        3
    } else if payload < 0x1000_0000 {
        4
    } else {
        5
    };
    kind_bytes
        .checked_add(length_bytes)
        .and_then(|value| value.checked_add(payload))
        .ok_or(Error::CapacityOverflow {
            resource: "Volatile Dependencies output",
        })
}

fn add_size(left: usize, right: usize) -> Result<usize> {
    left.checked_add(right).ok_or(Error::CapacityOverflow {
        resource: "Volatile Dependencies output",
    })
}

fn add_record_count(left: usize, right: usize) -> Result<usize> {
    left.checked_add(right).ok_or(Error::CapacityOverflow {
        resource: "Volatile Dependencies records",
    })
}

fn empty(payload: &[u8], record: &'static str) -> Result<()> {
    if payload.is_empty() {
        Ok(())
    } else {
        Err(invalid(format!(
            "{record} has {} unexpected payload bytes",
            payload.len()
        )))
    }
}

fn is_known(record_kind: crate::raw::Kind) -> bool {
    record_kind == kind::BEGIN_VOL_DEPS
        || record_kind == kind::END_VOL_DEPS
        || record_kind == kind::BEGIN_VOL_TYPE
        || record_kind == kind::END_VOL_TYPE
        || record_kind == kind::BEGIN_VOL_MAIN
        || record_kind == kind::END_VOL_MAIN
        || record_kind == kind::BEGIN_VOL_TOPIC
        || record_kind == kind::END_VOL_TOPIC
        || record_kind == kind::VOL_SUBTOPIC
        || record_kind == kind::VOL_REF
        || record_kind == kind::VOL_NUM
        || record_kind == kind::VOL_ERR
        || record_kind == kind::VOL_STR
        || record_kind == kind::VOL_BOOL
}

fn limit(resource: &'static str, actual: usize, maximum: usize) -> Error {
    Error::LimitExceeded {
        resource,
        actual,
        maximum,
    }
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}
