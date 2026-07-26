//! Typed, inert XLSB External Link part authoring (MS-XLSB 2.1.7.25).
//!
//! Workbook and OLE targets are stored as external OPC relationships. DDE
//! metadata has no relationship. No target is opened, contacted, refreshed,
//! instantiated, evaluated, or executed.

use crate::xlsb::error::{XlsbError, XlsbResult};
use crate::xlsb::external_link::{
    DATA_ITEM_REQUIRED_TRAILING_FLAG, DATA_ITEM_WANT_ADVISE, DATA_ITEM_WANT_PICTURE,
    DDE_ITEM_SUPPORTS_OLE, EXTERNAL_NAME_BUILT_IN, EXTERNAL_REFERENCE_DDE, EXTERNAL_REFERENCE_OLE,
    EXTERNAL_REFERENCE_WORKBOOK, OLE_ITEM_DISPLAY_AS_ICON, OLE_ITEM_REQUIRED_CLASS_FLAG,
    XlsbDdeItem, XlsbExternalCachedValue, XlsbExternalDefinedName, XlsbExternalLink,
    XlsbExternalLinkKind, XlsbExternalValueMatrix, XlsbOleItem,
};
use crate::xlsb::records::record_types;
use crate::xlsb::writer::RecordWriter;
use litchi_opc::PackURI;
use litchi_opc::constants::relationship_type;
use litchi_opc::part::{BlobPart, Part};

const EXTERNAL_LINK_CONTENT_TYPE: &str = "application/vnd.ms-excel.externalLink";
const MAX_EXTERNAL_LINK_PART_BYTES: usize = 32 * 1024 * 1024;
const NULL_WIDE_STRING_LENGTH: u32 = u32::MAX;

pub(crate) fn author_external_link_part(
    link: &XlsbExternalLink,
    one_based_index: usize,
) -> XlsbResult<BlobPart> {
    link.validate()?;
    if one_based_index == 0 {
        return Err(XlsbError::InvalidFormula(
            "external-link part index must be one-based".to_string(),
        ));
    }
    let mut part = BlobPart::new(
        PackURI::new(format!(
            "/xl/externalLinks/externalLink{one_based_index}.bin"
        ))?,
        EXTERNAL_LINK_CONTENT_TYPE.to_string(),
        Vec::new(),
    );
    let relationship_id = match link.kind() {
        XlsbExternalLinkKind::Workbook => {
            Some(part.relate_to_ext(link.source(), relationship_type::EXTERNAL_LINK_PATH))
        },
        XlsbExternalLinkKind::Dde => None,
        XlsbExternalLinkKind::Ole => {
            Some(part.relate_to_ext(link.source(), relationship_type::OLE_OBJECT))
        },
    };
    let bytes = write_external_link_stream(link, relationship_id.as_deref())?;
    if bytes.len() > MAX_EXTERNAL_LINK_PART_BYTES {
        return Err(XlsbError::InvalidLength {
            expected: MAX_EXTERNAL_LINK_PART_BYTES,
            found: bytes.len(),
        });
    }
    part.set_blob(bytes);
    Ok(part)
}

fn write_external_link_stream(
    link: &XlsbExternalLink,
    relationship_id: Option<&str>,
) -> XlsbResult<Vec<u8>> {
    let mut bytes = Vec::with_capacity(256);
    let mut writer = RecordWriter::new(&mut bytes);
    let mut begin = Vec::with_capacity(64);
    let mut payload = RecordWriter::new(&mut begin);
    match link.kind() {
        XlsbExternalLinkKind::Workbook => {
            payload.write_u16(EXTERNAL_REFERENCE_WORKBOOK)?;
            payload.write_wide_string(required_relationship_id(relationship_id, "workbook")?)?;
            payload.write_u32(NULL_WIDE_STRING_LENGTH)?;
        },
        XlsbExternalLinkKind::Dde => {
            if relationship_id.is_some() {
                return Err(XlsbError::InvalidFormula(
                    "DDE external link cannot have a relationship".to_string(),
                ));
            }
            payload.write_u16(EXTERNAL_REFERENCE_DDE)?;
            payload.write_wide_string(link.source())?;
            payload.write_wide_string(
                link.dde_topic().ok_or_else(|| {
                    XlsbError::InvalidFormula("DDE link has no topic".to_string())
                })?,
            )?;
        },
        XlsbExternalLinkKind::Ole => {
            payload.write_u16(EXTERNAL_REFERENCE_OLE)?;
            payload.write_wide_string(required_relationship_id(relationship_id, "OLE")?)?;
            payload.write_wide_string(link.ole_program_id().ok_or_else(|| {
                XlsbError::InvalidFormula("OLE link has no program ID".to_string())
            })?)?;
        },
    }
    writer.write_record(record_types::BEGIN_SUP_BOOK, &begin)?;

    if link.is_workbook() {
        let mut tabs = Vec::with_capacity(4 + link.sheet_names().len() * 16);
        let mut payload = RecordWriter::new(&mut tabs);
        payload.write_u32(u32::try_from(link.sheet_names().len()).map_err(|_| {
            XlsbError::InvalidFormula("external sheet-name count overflow".to_string())
        })?)?;
        for name in link.sheet_names() {
            payload.write_wide_string(name)?;
        }
        writer.write_record(record_types::SUP_TABS, &tabs)?;
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
    writer.write_record(record_types::END_SUP_BOOK, &[])?;
    Ok(bytes)
}

fn write_entry_start<W: std::io::Write>(
    writer: &mut RecordWriter<W>,
    name: &str,
) -> XlsbResult<()> {
    let mut name_payload = Vec::with_capacity(4 + name.len() * 2);
    RecordWriter::new(&mut name_payload).write_wide_string(name)?;
    writer.write_record(record_types::SUP_NAME_START, &name_payload)?;
    Ok(())
}

fn write_defined_name<W: std::io::Write>(
    writer: &mut RecordWriter<W>,
    entry: &XlsbExternalDefinedName,
) -> XlsbResult<()> {
    write_entry_start(writer, entry.name())?;
    let formula = entry.formula().map_or(&[][..], |formula| formula.tokens());
    let mut formula_payload = Vec::with_capacity(4 + formula.len());
    formula_payload.extend_from_slice(
        &u32::try_from(formula.len())
            .map_err(|_| XlsbError::InvalidFormula("external formula size overflow".to_string()))?
            .to_le_bytes(),
    );
    formula_payload.extend_from_slice(formula);
    writer.write_record(record_types::SUP_NAME_FORMULA, &formula_payload)?;

    let mut bits = [0u8; 7];
    bits[0] = u8::from(entry.is_built_in()) * EXTERNAL_NAME_BUILT_IN;
    let scope = entry
        .scope_sheet_index()
        .map_or(0u32, |index| u32::from(index) + 1);
    bits[2..6].copy_from_slice(&scope.to_le_bytes());
    writer.write_record(record_types::SUP_NAME_BITS, &bits)?;
    writer.write_record(record_types::SUP_NAME_END, &[])?;
    Ok(())
}

fn write_dde_item<W: std::io::Write>(
    writer: &mut RecordWriter<W>,
    item: &XlsbDdeItem,
) -> XlsbResult<()> {
    write_entry_start(writer, item.name())?;
    let mut bits = [0u8; 7];
    bits[0] = (u8::from(item.wants_advise()) * DATA_ITEM_WANT_ADVISE)
        | (u8::from(item.wants_picture()) * DATA_ITEM_WANT_PICTURE)
        | (u8::from(item.supports_ole()) * DDE_ITEM_SUPPORTS_OLE);
    bits[6] = DATA_ITEM_REQUIRED_TRAILING_FLAG;
    writer.write_record(record_types::SUP_NAME_BITS, &bits)?;
    if let Some(values) = item.cached_values() {
        write_cached_values(writer, values)?;
    }
    writer.write_record(record_types::SUP_NAME_END, &[])?;
    Ok(())
}

fn write_ole_item<W: std::io::Write>(
    writer: &mut RecordWriter<W>,
    item: &XlsbOleItem,
) -> XlsbResult<()> {
    write_entry_start(writer, item.name())?;
    let mut bits = [0u8; 7];
    bits[0] = (u8::from(item.wants_advise()) * DATA_ITEM_WANT_ADVISE)
        | (u8::from(item.wants_picture()) * DATA_ITEM_WANT_PICTURE)
        | OLE_ITEM_REQUIRED_CLASS_FLAG
        | (u8::from(item.displays_as_icon()) * OLE_ITEM_DISPLAY_AS_ICON);
    bits[6] = DATA_ITEM_REQUIRED_TRAILING_FLAG;
    writer.write_record(record_types::SUP_NAME_BITS, &bits)?;
    if let Some(values) = item.cached_values() {
        write_cached_values(writer, values)?;
    }
    writer.write_record(record_types::SUP_NAME_END, &[])?;
    Ok(())
}

fn write_cached_values<W: std::io::Write>(
    writer: &mut RecordWriter<W>,
    values: &XlsbExternalValueMatrix,
) -> XlsbResult<()> {
    let mut dimensions = Vec::with_capacity(8);
    dimensions.extend_from_slice(&values.rows().to_le_bytes());
    dimensions.extend_from_slice(&values.columns().to_le_bytes());
    writer.write_record(record_types::SUP_NAME_VALUE_START, &dimensions)?;
    for value in values.values() {
        match value {
            XlsbExternalCachedValue::Empty => {
                writer.write_record(record_types::SUP_NAME_NIL, &[])?;
            },
            XlsbExternalCachedValue::Number(number) => {
                writer.write_record(record_types::SUP_NAME_NUM, &number.to_le_bytes())?;
            },
            XlsbExternalCachedValue::Boolean(value) => {
                writer.write_record(record_types::SUP_NAME_BOOL, &[u8::from(*value)])?;
            },
            XlsbExternalCachedValue::Error(error) => {
                writer.write_record(record_types::SUP_NAME_ERROR, &[error.code()])?;
            },
            XlsbExternalCachedValue::String(value) => {
                let mut payload = Vec::with_capacity(4 + value.len() * 2);
                RecordWriter::new(&mut payload).write_wide_string(value)?;
                writer.write_record(record_types::SUP_NAME_STRING, &payload)?;
            },
        }
    }
    writer.write_record(record_types::SUP_NAME_VALUE_END, &[])?;
    Ok(())
}

fn required_relationship_id<'a>(
    relationship_id: Option<&'a str>,
    context: &str,
) -> XlsbResult<&'a str> {
    relationship_id
        .filter(|value| !value.is_empty())
        .ok_or_else(|| {
            XlsbError::InvalidFormula(format!("{context} external link has no relationship"))
        })
}
