//! BIFF12 stream authoring for one XLSB external-link part.
//!
//! OPC relationships and part placement remain host concerns. This owner
//! emits only the inert `BrtSupBook` stream.

use super::{
    CachedValue, DATA_ITEM_REQUIRED_TRAILING_FLAG, DATA_ITEM_WANT_ADVISE, DATA_ITEM_WANT_PICTURE,
    DDE_ITEM_SUPPORTS_OLE, DdeItem, DefinedName, EXTERNAL_NAME_BUILT_IN, EXTERNAL_REFERENCE_DDE,
    EXTERNAL_REFERENCE_OLE, EXTERNAL_REFERENCE_WORKBOOK, Error, Kind, Link,
    OLE_ITEM_DISPLAY_AS_ICON, OLE_ITEM_REQUIRED_CLASS_FLAG, OleItem, Result, ValueMatrix,
};
use crate::raw::{Writer, kind};

const NULL_WIDE_STRING_LENGTH: u32 = u32::MAX;

pub fn write_external_link_stream(link: &Link, relationship_id: Option<&str>) -> Result<Vec<u8>> {
    link.validate()?;
    let mut bytes = Vec::with_capacity(256);
    let mut writer = Writer::new(&mut bytes);
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
    writer.write_record(kind::BEGIN_SUP_BOOK, &begin)?;

    if link.is_workbook() {
        let mut tabs = Vec::with_capacity(4 + link.sheet_names().len() * 16);
        let mut payload = Writer::new(&mut tabs);
        payload.write_u32(u32::try_from(link.sheet_names().len()).map_err(|_| {
            Error::InvalidFormula("external sheet-name count overflow".to_string())
        })?)?;
        for name in link.sheet_names() {
            payload.write_wide_string(name)?;
        }
        writer.write_record(kind::SUP_TABS, &tabs)?;
    }

    for entry in link.defined_names() {
        write_defined_name(&mut writer, entry)?;
    }
    for item in link.dde_items() {
        write_dde_item(&mut writer, item)?;
    }
    for item in link.ole_items() {
        write_ole_item(&mut writer, item)?;
    }
    writer.write_record(kind::END_SUP_BOOK, &[])?;
    Ok(bytes)
}

fn write_entry_start<W: std::io::Write>(writer: &mut Writer<W>, name: &str) -> Result<()> {
    let mut name_payload = Vec::with_capacity(4 + name.len() * 2);
    Writer::new(&mut name_payload).write_wide_string(name)?;
    writer.write_record(kind::SUP_NAME_START, &name_payload)?;
    Ok(())
}

fn write_defined_name<W: std::io::Write>(
    writer: &mut Writer<W>,
    entry: &DefinedName,
) -> Result<()> {
    write_entry_start(writer, entry.name())?;
    let formula = entry.formula().map_or(&[][..], |formula| formula.tokens());
    let mut formula_payload = Vec::with_capacity(4 + formula.len());
    formula_payload.extend_from_slice(
        &u32::try_from(formula.len())
            .map_err(|_| Error::InvalidFormula("external formula size overflow".to_string()))?
            .to_le_bytes(),
    );
    formula_payload.extend_from_slice(formula);
    writer.write_record(kind::SUP_NAME_FORMULA, &formula_payload)?;

    let mut bits = [0u8; 7];
    bits[0] = u8::from(entry.is_built_in()) * EXTERNAL_NAME_BUILT_IN;
    let scope = entry
        .scope_sheet_index()
        .map_or(0u32, |index| u32::from(index) + 1);
    bits[2..6].copy_from_slice(&scope.to_le_bytes());
    writer.write_record(kind::SUP_NAME_BITS, &bits)?;
    writer.write_record(kind::SUP_NAME_END, &[])?;
    Ok(())
}

fn write_dde_item<W: std::io::Write>(writer: &mut Writer<W>, item: &DdeItem) -> Result<()> {
    write_entry_start(writer, item.name())?;
    let mut bits = [0u8; 7];
    bits[0] = (u8::from(item.wants_advise()) * DATA_ITEM_WANT_ADVISE)
        | (u8::from(item.wants_picture()) * DATA_ITEM_WANT_PICTURE)
        | (u8::from(item.supports_ole()) * DDE_ITEM_SUPPORTS_OLE);
    bits[6] = DATA_ITEM_REQUIRED_TRAILING_FLAG;
    writer.write_record(kind::SUP_NAME_BITS, &bits)?;
    if let Some(values) = item.cached_values() {
        write_cached_values(writer, values)?;
    }
    writer.write_record(kind::SUP_NAME_END, &[])?;
    Ok(())
}

fn write_ole_item<W: std::io::Write>(writer: &mut Writer<W>, item: &OleItem) -> Result<()> {
    write_entry_start(writer, item.name())?;
    let mut bits = [0u8; 7];
    bits[0] = (u8::from(item.wants_advise()) * DATA_ITEM_WANT_ADVISE)
        | (u8::from(item.wants_picture()) * DATA_ITEM_WANT_PICTURE)
        | OLE_ITEM_REQUIRED_CLASS_FLAG
        | (u8::from(item.displays_as_icon()) * OLE_ITEM_DISPLAY_AS_ICON);
    bits[6] = DATA_ITEM_REQUIRED_TRAILING_FLAG;
    writer.write_record(kind::SUP_NAME_BITS, &bits)?;
    if let Some(values) = item.cached_values() {
        write_cached_values(writer, values)?;
    }
    writer.write_record(kind::SUP_NAME_END, &[])?;
    Ok(())
}

fn write_cached_values<W: std::io::Write>(
    writer: &mut Writer<W>,
    values: &ValueMatrix,
) -> Result<()> {
    let mut dimensions = Vec::with_capacity(8);
    dimensions.extend_from_slice(&values.rows().to_le_bytes());
    dimensions.extend_from_slice(&values.columns().to_le_bytes());
    writer.write_record(kind::SUP_NAME_VALUE_START, &dimensions)?;
    for value in values.values() {
        match value {
            CachedValue::Empty => {
                writer.write_record(kind::SUP_NAME_NIL, &[])?;
            },
            CachedValue::Number(number) => {
                writer.write_record(kind::SUP_NAME_NUM, &number.to_le_bytes())?;
            },
            CachedValue::Boolean(value) => {
                writer.write_record(kind::SUP_NAME_BOOL, &[u8::from(*value)])?;
            },
            CachedValue::Error(error) => {
                writer.write_record(kind::SUP_NAME_ERROR, &[error.code()])?;
            },
            CachedValue::String(value) => {
                let mut payload = Vec::with_capacity(4 + value.len() * 2);
                Writer::new(&mut payload).write_wide_string(value)?;
                writer.write_record(kind::SUP_NAME_STRING, &payload)?;
            },
        }
    }
    writer.write_record(kind::SUP_NAME_VALUE_END, &[])?;
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
