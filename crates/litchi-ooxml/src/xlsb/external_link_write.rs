//! Typed, inert XLSB External Link part authoring (MS-XLSB 2.1.7.25).
//!
//! Workbook and OLE targets are stored as external OPC relationships. DDE
//! metadata has no relationship. No target is opened, contacted, refreshed,
//! instantiated, evaluated, or executed.

use crate::xlsb::error::{XlsbError, XlsbResult};
use crate::xlsb::formula::{XlsbExternalLink, XlsbExternalLinkKind};
use crate::xlsb::records::record_types;
use crate::xlsb::writer::RecordWriter;
use litchi_opc::PackURI;
use litchi_opc::constants::relationship_type;
use litchi_opc::part::{BlobPart, Part};

const EXTERNAL_LINK_CONTENT_TYPE: &str = "application/vnd.ms-excel.externalLink";
const MAX_EXTERNAL_LINK_PART_BYTES: usize = 32 * 1024 * 1024;
const EXTERNAL_REFERENCE_WORKBOOK: u16 = 0;
const EXTERNAL_REFERENCE_DDE: u16 = 1;
const EXTERNAL_REFERENCE_OLE: u16 = 2;
const NULL_WIDE_STRING_LENGTH: u32 = u32::MAX;
const EMPTY_EXTERNAL_NAME_BITS: [u8; 7] = [0; 7];

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

    for name in link.declared_names() {
        let mut name_payload = Vec::with_capacity(4 + name.len() * 2);
        RecordWriter::new(&mut name_payload).write_wide_string(name)?;
        writer.write_record(record_types::SUP_NAME_START, &name_payload)?;
        if link.is_workbook() {
            // cb=0 means the external name is undefined. We do not invent a
            // formula absent one in the public typed model.
            writer.write_record(record_types::SUP_NAME_FORMULA, &0u32.to_le_bytes())?;
        }
        writer.write_record(record_types::SUP_NAME_BITS, &EMPTY_EXTERNAL_NAME_BITS)?;
        writer.write_record(record_types::SUP_NAME_END, &[])?;
    }
    writer.write_record(record_types::END_SUP_BOOK, &[])?;
    Ok(bytes)
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
