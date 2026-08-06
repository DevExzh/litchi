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
    let mut bytes = Vec::with_capacity(
        records.iter().map(Vec::len).sum::<usize>()
            + unknown_records
                .iter()
                .map(|record| record.bytes().len())
                .sum::<usize>(),
    );
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
    let mut records = Vec::with_capacity(2 + link.entry_count() * 4);
    let mut begin = Vec::with_capacity(64);
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
    records.push(encode_record(kind::BEGIN_SUP_BOOK, &begin)?);

    if link.is_workbook() {
        let mut tabs = Vec::with_capacity(4 + link.sheet_names().len() * 16);
        let mut payload = Writer::new(&mut tabs);
        payload.write_u32(u32::try_from(link.sheet_names().len()).map_err(|_| {
            Error::InvalidFormula("external sheet-name count overflow".to_string())
        })?)?;
        for name in link.sheet_names() {
            payload.write_wide_string(name)?;
        }
        records.push(encode_record(kind::SUP_TABS, &tabs)?);
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
    records.push(encode_record(kind::END_SUP_BOOK, &[])?);
    Ok(records)
}

fn encode_record(record_kind: crate::raw::Kind, payload: &[u8]) -> Result<Vec<u8>> {
    let mut bytes = Vec::with_capacity(payload.len() + 4);
    Writer::new(&mut bytes).write_record(record_kind, payload)?;
    Ok(bytes)
}

fn write_entry_start(records: &mut Vec<Vec<u8>>, name: &str) -> Result<()> {
    let mut payload = Vec::with_capacity(4 + name.len() * 2);
    Writer::new(&mut payload).write_wide_string(name)?;
    records.push(encode_record(kind::SUP_NAME_START, &payload)?);
    Ok(())
}

fn write_defined_name(records: &mut Vec<Vec<u8>>, entry: &DefinedName) -> Result<()> {
    write_entry_start(records, entry.name())?;
    let formula = entry.formula().map_or(&[][..], |formula| formula.tokens());
    let mut formula_payload = Vec::with_capacity(4 + formula.len());
    formula_payload.extend_from_slice(
        &u32::try_from(formula.len())
            .map_err(|_| Error::InvalidFormula("external formula size overflow".to_string()))?
            .to_le_bytes(),
    );
    formula_payload.extend_from_slice(formula);
    records.push(encode_record(kind::SUP_NAME_FORMULA, &formula_payload)?);

    let mut bits = [0u8; 7];
    bits[0] = u8::from(entry.is_built_in()) * EXTERNAL_NAME_BUILT_IN;
    let scope = entry
        .scope_sheet_index()
        .map_or(0u32, |index| u32::from(index) + 1);
    bits[2..6].copy_from_slice(&scope.to_le_bytes());
    records.push(encode_record(kind::SUP_NAME_BITS, &bits)?);
    records.push(encode_record(kind::SUP_NAME_END, &[])?);
    Ok(())
}

fn write_dde_item(records: &mut Vec<Vec<u8>>, item: &DdeItem) -> Result<()> {
    write_entry_start(records, item.name())?;
    let mut bits = [0u8; 7];
    bits[0] = (u8::from(item.wants_advise()) * DATA_ITEM_WANT_ADVISE)
        | (u8::from(item.wants_picture()) * DATA_ITEM_WANT_PICTURE)
        | (u8::from(item.supports_ole()) * DDE_ITEM_SUPPORTS_OLE);
    bits[6] = DATA_ITEM_REQUIRED_TRAILING_FLAG;
    records.push(encode_record(kind::SUP_NAME_BITS, &bits)?);
    if let Some(values) = item.cached_values() {
        write_cached_values(records, values)?;
    }
    records.push(encode_record(kind::SUP_NAME_END, &[])?);
    Ok(())
}

fn write_ole_item(records: &mut Vec<Vec<u8>>, item: &OleItem) -> Result<()> {
    write_entry_start(records, item.name())?;
    let mut bits = [0u8; 7];
    bits[0] = (u8::from(item.wants_advise()) * DATA_ITEM_WANT_ADVISE)
        | (u8::from(item.wants_picture()) * DATA_ITEM_WANT_PICTURE)
        | OLE_ITEM_REQUIRED_CLASS_FLAG
        | (u8::from(item.displays_as_icon()) * OLE_ITEM_DISPLAY_AS_ICON);
    bits[6] = DATA_ITEM_REQUIRED_TRAILING_FLAG;
    records.push(encode_record(kind::SUP_NAME_BITS, &bits)?);
    if let Some(values) = item.cached_values() {
        write_cached_values(records, values)?;
    }
    records.push(encode_record(kind::SUP_NAME_END, &[])?);
    Ok(())
}

fn write_cached_values(records: &mut Vec<Vec<u8>>, values: &ValueMatrix) -> Result<()> {
    let mut dimensions = Vec::with_capacity(8);
    dimensions.extend_from_slice(&values.rows().to_le_bytes());
    dimensions.extend_from_slice(&values.columns().to_le_bytes());
    records.push(encode_record(kind::SUP_NAME_VALUE_START, &dimensions)?);
    for value in values.values() {
        match value {
            CachedValue::Empty => records.push(encode_record(kind::SUP_NAME_NIL, &[])?),
            CachedValue::Number(number) => {
                records.push(encode_record(kind::SUP_NAME_NUM, &number.to_le_bytes())?);
            },
            CachedValue::Boolean(value) => {
                records.push(encode_record(kind::SUP_NAME_BOOL, &[u8::from(*value)])?);
            },
            CachedValue::Error(error) => {
                records.push(encode_record(kind::SUP_NAME_ERROR, &[error.code()])?);
            },
            CachedValue::String(value) => {
                let mut payload = Vec::with_capacity(4 + value.len() * 2);
                Writer::new(&mut payload).write_wide_string(value)?;
                records.push(encode_record(kind::SUP_NAME_STRING, &payload)?);
            },
        }
    }
    records.push(encode_record(kind::SUP_NAME_VALUE_END, &[])?);
    Ok(())
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
