#![allow(
    clippy::map_err_ignore,
    reason = "legacy module confines normalization into the module's stable typed public error to this codec boundary"
)]

//! BIFF12 stream authoring for one XLSB external-link part.
//!
//! OPC relationships and part placement remain host concerns. This owner
//! emits only the inert `BrtSupBook` stream; source-bound transactions may
//! additionally supply opaque records retained by the codec layer.

use super::validation;
use super::{
    CachedValue, DATA_ITEM_REQUIRED_TRAILING_FLAG, DATA_ITEM_WANT_ADVISE, DATA_ITEM_WANT_PICTURE,
    DDE_ITEM_SUPPORTS_OLE, DdeItem, DefinedName, EXTERNAL_NAME_BUILT_IN, EXTERNAL_REFERENCE_DDE,
    EXTERNAL_REFERENCE_OLE, EXTERNAL_REFERENCE_WORKBOOK, Error, Kind, Link,
    OLE_ITEM_DISPLAY_AS_ICON, OLE_ITEM_REQUIRED_CLASS_FLAG, OleItem, Result, UnknownRecord,
    ValueMatrix,
};
use crate::raw::{Writer, kind};

const NULL_WIDE_STRING_LENGTH: u32 = u32::MAX;

/// Write one canonical external-link stream without opaque records.
pub fn write_external_link_stream(link: &Link, relationship_id: Option<&str>) -> Result<Vec<u8>> {
    write_external_link_stream_with_unknown(link, relationship_id, &[])
}

/// Write one canonical stream while retaining source-owned opaque records.
pub(crate) fn write_external_link_stream_with_unknown(
    link: &Link,
    relationship_id: Option<&str>,
    unknown_records: &[UnknownRecord],
) -> Result<Vec<u8>> {
    validation::validate_link(link)?;
    validation::validate_relationship(link, relationship_id)?;
    validation::validate_unknown_records(unknown_records)?;

    let records = encode_records(link, relationship_id)?;
    let known_bytes = records.iter().try_fold(0usize, |total, record| {
        total
            .checked_add(record.len())
            .ok_or_else(|| Error::InvalidFormula("external-link output size overflow".to_string()))
    })?;
    let unknown_bytes = unknown_records.iter().try_fold(0usize, |total, record| {
        total
            .checked_add(record.bytes().len())
            .ok_or_else(|| Error::InvalidFormula("external-link output size overflow".to_string()))
    })?;
    let output_bytes = known_bytes
        .checked_add(unknown_bytes)
        .ok_or_else(|| Error::InvalidFormula("external-link output size overflow".to_string()))?;
    if output_bytes > super::MAX_LINK_PART_BYTES {
        return Err(Error::InvalidLength {
            expected: super::MAX_LINK_PART_BYTES,
            found: output_bytes,
        });
    }
    let mut bytes = Vec::new();
    bytes
        .try_reserve_exact(output_bytes)
        .map_err(|source| Error::Allocation {
            resource: "external-link output",
            source,
        })?;
    let end_index = records.len().saturating_sub(1);
    let mut unknown_index = 0usize;
    for known_index in 0..records.len() {
        while unknown_index < unknown_records.len()
            && unknown_records[unknown_index].after_known().min(end_index) == known_index
        {
            bytes.extend_from_slice(unknown_records[unknown_index].bytes());
            unknown_index += 1;
        }
        bytes.extend_from_slice(&records[known_index]);
    }
    while unknown_index < unknown_records.len() {
        bytes.extend_from_slice(unknown_records[unknown_index].bytes());
        unknown_index += 1;
    }
    if bytes.len() > super::MAX_LINK_PART_BYTES {
        return Err(Error::InvalidLength {
            expected: super::MAX_LINK_PART_BYTES,
            found: bytes.len(),
        });
    }
    Ok(bytes)
}

fn encode_records(link: &Link, relationship_id: Option<&str>) -> Result<Vec<Vec<u8>>> {
    let mut records = Vec::new();
    let minimum_records = 2usize
        .checked_add(link.entry_count().checked_mul(4).ok_or_else(|| {
            Error::InvalidFormula("external-link record count overflow".to_string())
        })?)
        .ok_or_else(|| Error::InvalidFormula("external-link record count overflow".to_string()))?;
    records
        .try_reserve(minimum_records)
        .map_err(|source| Error::Allocation {
            resource: "external-link records",
            source,
        })?;
    let begin_capacity = begin_capacity(link, relationship_id)?;
    let mut begin = try_buffer(begin_capacity, "external-link header")?;
    let mut payload = Writer::new(&mut begin);
    match link.kind() {
        Kind::Workbook => {
            payload.write_u16(EXTERNAL_REFERENCE_WORKBOOK)?;
            payload.write_wide_string(required_relationship_id(relationship_id, "workbook")?)?;
            payload.write_u32(NULL_WIDE_STRING_LENGTH)?;
        },
        Kind::Dde => {
            if relationship_id.is_some() {
                return Err(Error::InvalidFormula(
                    "DDE external link cannot have a relationship".to_string(),
                ));
            }
            payload.write_u16(EXTERNAL_REFERENCE_DDE)?;
            payload.write_wide_string(link.source())?;
            payload.write_wide_string(
                link.dde_topic()
                    .ok_or_else(|| Error::InvalidFormula("DDE link has no topic".to_string()))?,
            )?;
        },
        Kind::Ole => {
            payload.write_u16(EXTERNAL_REFERENCE_OLE)?;
            payload.write_wide_string(required_relationship_id(relationship_id, "OLE")?)?;
            payload.write_wide_string(link.ole_program_id().ok_or_else(|| {
                Error::InvalidFormula("OLE link has no program ID".to_string())
            })?)?;
        },
    }
    push_record(&mut records, kind::BEGIN_SUP_BOOK, &begin)?;

    if link.is_workbook() {
        let tabs_capacity = link.sheet_names().iter().try_fold(4usize, |total, name| {
            wide_string_wire_size(name)
                .and_then(|size| total.checked_add(size))
                .ok_or_else(|| {
                    Error::InvalidFormula("external sheet-name size overflow".to_string())
                })
        })?;
        let mut tabs = try_buffer(tabs_capacity, "external sheet names")?;
        let mut payload = Writer::new(&mut tabs);
        payload.write_u32(u32::try_from(link.sheet_names().len()).map_err(|_| {
            Error::InvalidFormula("external sheet-name count overflow".to_string())
        })?)?;
        for name in link.sheet_names() {
            payload.write_wide_string(name)?;
        }
        push_record(&mut records, kind::SUP_TABS, &tabs)?;
    }

    for entry in link.defined_names() {
        write_defined_name(&mut records, entry)?;
    }
    for item in link.dde_items() {
        write_dde_item(&mut records, item)?;
    }
    for item in link.ole_items() {
        write_ole_item(&mut records, item)?;
    }
    push_record(&mut records, kind::END_SUP_BOOK, &[])?;
    Ok(records)
}

fn encode_record(record_kind: crate::raw::Kind, payload: &[u8]) -> Result<Vec<u8>> {
    let mut bytes = Vec::new();
    bytes
        .try_reserve(payload.len().checked_add(8).ok_or_else(|| {
            Error::InvalidFormula("external-link record size overflow".to_string())
        })?)
        .map_err(|source| Error::Allocation {
            resource: "external-link record",
            source,
        })?;
    Writer::new(&mut bytes).write_record(record_kind, payload)?;
    Ok(bytes)
}

fn write_entry_start(records: &mut Vec<Vec<u8>>, name: &str) -> Result<()> {
    let mut payload = Vec::new();
    payload
        .try_reserve(
            wide_string_wire_size(name)
                .ok_or_else(|| Error::InvalidFormula("external name size overflow".to_string()))?,
        )
        .map_err(|source| Error::Allocation {
            resource: "external name",
            source,
        })?;
    Writer::new(&mut payload).write_wide_string(name)?;
    push_record(records, kind::SUP_NAME_START, &payload)
}

fn write_defined_name(records: &mut Vec<Vec<u8>>, entry: &DefinedName) -> Result<()> {
    write_entry_start(records, entry.name())?;
    let formula = entry.formula().map_or(&[][..], |formula| formula.tokens());
    let mut formula_payload = Vec::new();
    formula_payload
        .try_reserve(
            formula.len().checked_add(4).ok_or_else(|| {
                Error::InvalidFormula("external formula size overflow".to_string())
            })?,
        )
        .map_err(|source| Error::Allocation {
            resource: "external formula",
            source,
        })?;
    formula_payload.extend_from_slice(
        &u32::try_from(formula.len())
            .map_err(|_| Error::InvalidFormula("external formula size overflow".to_string()))?
            .to_le_bytes(),
    );
    formula_payload.extend_from_slice(formula);
    push_record(records, kind::SUP_NAME_FORMULA, &formula_payload)?;

    let mut bits = entry.wire_bits().unwrap_or([0u8; 7]);
    let scope = entry
        .scope_sheet_index()
        .map_or(0u32, |index| u32::from(index) + 1);
    bits[2..6].copy_from_slice(&scope.to_le_bytes());
    bits[0] = (entry.wire_bits().unwrap_or([0u8; 7])[0] & !EXTERNAL_NAME_BUILT_IN)
        | (u8::from(entry.is_built_in()) * EXTERNAL_NAME_BUILT_IN);
    push_record(records, kind::SUP_NAME_BITS, &bits)?;
    push_record(records, kind::SUP_NAME_END, &[])
}

fn write_dde_item(records: &mut Vec<Vec<u8>>, item: &DdeItem) -> Result<()> {
    write_entry_start(records, item.name())?;
    let mut bits = item.wire_bits().unwrap_or([0u8; 7]);
    let modeled_bits = DATA_ITEM_WANT_ADVISE | DATA_ITEM_WANT_PICTURE | DDE_ITEM_SUPPORTS_OLE;
    bits[0] = (bits[0] & !modeled_bits)
        | (u8::from(item.wants_advise()) * DATA_ITEM_WANT_ADVISE)
        | (u8::from(item.wants_picture()) * DATA_ITEM_WANT_PICTURE)
        | (u8::from(item.supports_ole()) * DDE_ITEM_SUPPORTS_OLE);
    bits[6] |= DATA_ITEM_REQUIRED_TRAILING_FLAG;
    push_record(records, kind::SUP_NAME_BITS, &bits)?;
    if let Some(values) = item.cached_values() {
        write_cached_values(records, values)?;
    }
    push_record(records, kind::SUP_NAME_END, &[])
}

fn write_ole_item(records: &mut Vec<Vec<u8>>, item: &OleItem) -> Result<()> {
    write_entry_start(records, item.name())?;
    let mut bits = item.wire_bits().unwrap_or([0u8; 7]);
    let modeled_bits = DATA_ITEM_WANT_ADVISE
        | DATA_ITEM_WANT_PICTURE
        | OLE_ITEM_REQUIRED_CLASS_FLAG
        | OLE_ITEM_DISPLAY_AS_ICON;
    bits[0] = (bits[0] & !modeled_bits)
        | (u8::from(item.wants_advise()) * DATA_ITEM_WANT_ADVISE)
        | (u8::from(item.wants_picture()) * DATA_ITEM_WANT_PICTURE)
        | OLE_ITEM_REQUIRED_CLASS_FLAG
        | (u8::from(item.displays_as_icon()) * OLE_ITEM_DISPLAY_AS_ICON);
    bits[6] |= DATA_ITEM_REQUIRED_TRAILING_FLAG;
    push_record(records, kind::SUP_NAME_BITS, &bits)?;
    if let Some(values) = item.cached_values() {
        write_cached_values(records, values)?;
    }
    push_record(records, kind::SUP_NAME_END, &[])
}

fn write_cached_values(records: &mut Vec<Vec<u8>>, values: &ValueMatrix) -> Result<()> {
    let mut dimensions = try_buffer(8, "external cache dimensions")?;
    dimensions.extend_from_slice(&values.rows().to_le_bytes());
    dimensions.extend_from_slice(&values.columns().to_le_bytes());
    push_record(records, kind::SUP_NAME_VALUE_START, &dimensions)?;
    for value in values.values() {
        match value {
            CachedValue::Empty => push_record(records, kind::SUP_NAME_NIL, &[])?,
            CachedValue::Number(number) => {
                push_record(records, kind::SUP_NAME_NUM, &number.to_le_bytes())?;
            },
            CachedValue::Boolean(value) => {
                push_record(records, kind::SUP_NAME_BOOL, &[u8::from(*value)])?;
            },
            CachedValue::Error(error) => {
                push_record(records, kind::SUP_NAME_ERROR, &[error.code()])?;
            },
            CachedValue::String(value) => {
                let mut payload = try_buffer(
                    wide_string_wire_size(value).ok_or_else(|| {
                        Error::InvalidFormula("external cached string size overflow".to_string())
                    })?,
                    "external cached string",
                )?;
                Writer::new(&mut payload).write_wide_string(value)?;
                push_record(records, kind::SUP_NAME_STRING, &payload)?;
            },
        }
    }
    push_record(records, kind::SUP_NAME_VALUE_END, &[])
}

fn push_record(
    records: &mut Vec<Vec<u8>>,
    record_kind: crate::raw::Kind,
    payload: &[u8],
) -> Result<()> {
    records.try_reserve(1).map_err(|source| Error::Allocation {
        resource: "external-link records",
        source,
    })?;
    let record = encode_record(record_kind, payload)?;
    records.push(record);
    Ok(())
}

fn try_buffer(capacity: usize, resource: &'static str) -> Result<Vec<u8>> {
    let mut buffer = Vec::new();
    buffer
        .try_reserve(capacity)
        .map_err(|source| Error::Allocation { resource, source })?;
    Ok(buffer)
}

fn wide_string_wire_size(value: &str) -> Option<usize> {
    value
        .encode_utf16()
        .count()
        .checked_mul(2)
        .and_then(|size| size.checked_add(4))
}

fn begin_capacity(link: &Link, relationship_id: Option<&str>) -> Result<usize> {
    let fixed = 2usize;
    match link.kind() {
        Kind::Workbook => {
            let relationship_id = required_relationship_id(relationship_id, "workbook")?;
            fixed
                .checked_add(wide_string_wire_size(relationship_id).ok_or_else(|| {
                    Error::InvalidFormula("external-link relationship size overflow".to_string())
                })?)
                .and_then(|size| size.checked_add(4))
                .ok_or_else(|| {
                    Error::InvalidFormula("external-link header size overflow".to_string())
                })
        },
        Kind::Dde => {
            let topic = link
                .dde_topic()
                .ok_or_else(|| Error::InvalidFormula("DDE link has no topic".to_string()))?;
            fixed
                .checked_add(
                    wide_string_wire_size(link.source()).ok_or_else(|| {
                        Error::InvalidFormula("DDE source size overflow".to_string())
                    })?,
                )
                .and_then(|size| {
                    wide_string_wire_size(topic).and_then(|topic_size| size.checked_add(topic_size))
                })
                .ok_or_else(|| Error::InvalidFormula("DDE header size overflow".to_string()))
        },
        Kind::Ole => {
            let relationship_id = required_relationship_id(relationship_id, "OLE")?;
            let program_id = link
                .ole_program_id()
                .ok_or_else(|| Error::InvalidFormula("OLE link has no program ID".to_string()))?;
            fixed
                .checked_add(wide_string_wire_size(relationship_id).ok_or_else(|| {
                    Error::InvalidFormula("OLE relationship size overflow".to_string())
                })?)
                .and_then(|size| {
                    wide_string_wire_size(program_id)
                        .and_then(|program_size| size.checked_add(program_size))
                })
                .ok_or_else(|| Error::InvalidFormula("OLE header size overflow".to_string()))
        },
    }
}

fn required_relationship_id<'a>(
    relationship_id: Option<&'a str>,
    context: &str,
) -> Result<&'a str> {
    relationship_id
        .filter(|value| !value.is_empty())
        .ok_or_else(|| {
            Error::InvalidFormula(format!("{context} external link has no relationship"))
        })
}
