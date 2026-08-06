//! Binary reader for `pivotCacheRecords*.bin` (MS-XLSB 2.1.7.39).
//!
//! `BrtPCRRecord` stores direct values in its payload, while
//! `BrtPCRRecordDt` stores one value per child `BrtPCDI*` record. The field
//! metadata from the definition is required to distinguish shared-item
//! indexes from direct numbers, dates, and strings.

use crate::package::error::{Error, Result};
use crate::package::pivot::model::*;
use crate::package::walker::{RecordWalker, malformed};
use crate::raw::{Cursor, kind as rt};

/// Parse the records part without applying definition-wide semantic checks.
pub(super) fn parse_pivot_cache_records_binary(
    data: &[u8],
    definition: &PivotCacheDefinition,
) -> Result<PivotCacheRecords> {
    let mut walker = RecordWalker::new(data);
    let begin =
        walker.required_begin(rt::BEGIN_PIVOT_CACHE_RECORDS, "BrtBeginPivotCacheRecords")?;
    let declared = read_count(begin.payload(), "BrtBeginPivotCacheRecords")?;
    let mut records = PivotCacheRecords {
        record_count: declared,
        records: Vec::new(),
    };

    while let Some(record) = walker.next()? {
        match record.kind() {
            rt::END_PIVOT_CACHE_RECORDS => {
                validate_record_count(declared, records.records.len())?;
                return Ok(records);
            },
            rt::PCR_RECORD => {
                records
                    .records
                    .push(parse_direct_record(record.payload(), definition)?);
            },
            rt::PCR_RECORD_DT => {
                records
                    .records
                    .push(parse_record_dt(&mut walker, definition)?);
            },
            other => walker.skip_unhandled(other, "PivotCache records stream")?,
        }
    }

    Err(Error::UnexpectedEndOfStream(
        "BrtEndPivotCacheRecords".to_string(),
    ))
}

fn read_count(data: &[u8], context: &'static str) -> Result<u32> {
    let mut cursor = Cursor::new(data, context);
    let count = cursor.read_u32()?;
    cursor.finish()?;
    Ok(count)
}

fn validate_record_count(declared: u32, actual: usize) -> Result<()> {
    if u64::from(declared) != actual as u64 {
        return Err(malformed(
            "BrtBeginPivotCacheRecords",
            format!("declared {declared} records, found {actual}"),
        ));
    }
    Ok(())
}

fn parse_direct_record(data: &[u8], definition: &PivotCacheDefinition) -> Result<PivotCacheRecord> {
    let mut cursor = Cursor::new(data, "BrtPCRRecord");
    let mut values = Vec::with_capacity(source_field_count(definition));
    for field in source_fields(definition) {
        values.push(read_direct_value(&mut cursor, field)?);
    }
    cursor.finish()?;
    Ok(PivotCacheRecord { values })
}

fn parse_record_dt(
    walker: &mut RecordWalker<'_>,
    definition: &PivotCacheDefinition,
) -> Result<PivotCacheRecord> {
    let mut values = Vec::with_capacity(source_field_count(definition));
    for field in source_fields(definition) {
        let record = walker.required("BrtPCRRecordDt value")?;
        values.push(read_record_dt_value(
            record.kind(),
            record.payload(),
            field,
        )?);
    }
    Ok(PivotCacheRecord { values })
}

fn source_fields<'a>(
    definition: &'a PivotCacheDefinition,
) -> impl Iterator<Item = &'a PivotCacheField> {
    definition.fields.iter().filter(|field| field.source_field)
}

fn source_field_count(definition: &PivotCacheDefinition) -> usize {
    definition
        .fields
        .iter()
        .filter(|field| field.source_field)
        .count()
}

fn read_direct_value(
    cursor: &mut Cursor<'_>,
    field: &PivotCacheField,
) -> Result<PivotCacheItemValue> {
    if !field.shared_items.items.is_empty() {
        return Ok(PivotCacheItemValue::Index(cursor.read_u32()?));
    }

    let stats = field.shared_items.stats.as_ref().ok_or_else(|| {
        malformed(
            "BrtPCRRecord",
            format!("field {:?} has no type metadata", field.name),
        )
    })?;
    if stats.numeric_field {
        Ok(PivotCacheItemValue::Number(cursor.read_f64()?))
    } else if stats.date_in_field && !stats.has_text_item {
        Ok(PivotCacheItemValue::DateTime(super::parse::read_date_time(
            cursor,
        )?))
    } else {
        Ok(PivotCacheItemValue::String(
            cursor.read_wide_string()?.into_boxed_str(),
        ))
    }
}

fn read_record_dt_value(
    record_type: crate::raw::Kind,
    data: &[u8],
    field: &PivotCacheField,
) -> Result<PivotCacheItemValue> {
    let mut cursor = Cursor::new(data, "BrtPCDI cache record value");
    let value = if !field.shared_items.items.is_empty() {
        if record_type != rt::PCDI_INDEX {
            return Err(Error::UnexpectedRecord {
                expected: rt::PCDI_INDEX.get(),
                found: record_type.get(),
            });
        }
        PivotCacheItemValue::Index(cursor.read_u32()?)
    } else {
        let stats = field.shared_items.stats.as_ref().ok_or_else(|| {
            malformed(
                "BrtPCRRecordDt",
                format!("field {:?} has no type metadata", field.name),
            )
        })?;
        let expected = if stats.numeric_field {
            rt::PCDI_NUMBER
        } else if stats.date_in_field && !stats.has_text_item {
            rt::PCDI_DATETIME
        } else {
            rt::PCDI_STRING
        };
        if record_type != expected {
            return Err(Error::UnexpectedRecord {
                expected: expected.get(),
                found: record_type.get(),
            });
        }
        match record_type {
            rt::PCDI_NUMBER => PivotCacheItemValue::Number(cursor.read_f64()?),
            rt::PCDI_DATETIME => {
                PivotCacheItemValue::DateTime(super::parse::read_date_time(&mut cursor)?)
            },
            rt::PCDI_STRING => {
                PivotCacheItemValue::String(cursor.read_wide_string()?.into_boxed_str())
            },
            _ => unreachable!("record type was checked above"),
        }
    };
    cursor.finish()?;
    Ok(value)
}
