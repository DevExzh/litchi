#![allow(
    clippy::map_err_ignore,
    reason = "legacy module confines normalization into the module's stable typed public error to this codec boundary"
)]

//! Lossless workbook resource interning for structural cell transfer.

#![deny(
    clippy::cast_lossless,
    clippy::cast_possible_truncation,
    clippy::cast_possible_wrap,
    clippy::cast_precision_loss,
    clippy::cast_sign_loss,
    clippy::checked_conversions,
    clippy::expect_used,
    clippy::float_cmp,
    clippy::let_underscore_must_use,
    clippy::unnecessary_unwrap,
    reason = "new resource transfer code uses checked BIFF12 indexes and lengths"
)]

use super::StyleIndex;
use crate::package::error::{Error, Result};
use crate::package::shared_strings::SharedString;
use crate::raw::{Header, Kind, Limits as RawLimits, Records, Writer, kind};
use litchi_core::binary;
use litchi_opc::{BlobPart, OpcPackage, PackURI};

const SST_URI: &str = "/xl/sharedStrings.bin";
const STYLES_URI: &str = "/xl/styles.bin";
const WORKBOOK_URI: &str = "/xl/workbook.bin";
const SST_CONTENT_TYPE: &str = "application/vnd.ms-excel.sharedStrings";
const MAX_SST_COUNT: u32 = 0x7fff_ffff;

#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) struct StylePlan {
    pub(super) font: Vec<u8>,
    pub(super) fill: Vec<u8>,
    pub(super) border: Vec<u8>,
    pub(super) number_format: Option<String>,
    pub(super) xf_tail: [u8; 6],
}

pub(super) fn plan_style(style: &super::AuthoredStyle) -> Result<StylePlan> {
    let alignment = style.alignment.as_ref();
    if alignment
        .is_some_and(|value| value.rotation > 180 || value.indent > 250 || value.text_direction > 2)
    {
        return Err(Error::InvalidFormat(
            "authored style alignment is outside BIFF12 bounds".to_string(),
        ));
    }
    let mut xf_tail = [0_u8; 6];
    if let Some(value) = alignment {
        xf_tail[0] = value.rotation;
        xf_tail[1] = value.indent;
        xf_tail[2] = horizontal_bits(value.horizontal)
            | (vertical_bits(value.vertical) << 3)
            | (u8::from(value.wrap_text) << 6);
        xf_tail[3] = u8::from(value.shrink_to_fit) | (value.text_direction << 2);
    }
    if let Some(code) = &style.number_format {
        let units = code.encode_utf16().count();
        if !(1..=255).contains(&units) || code.contains('\0') {
            return Err(Error::InvalidFormat(
                "authored number format must contain 1..=255 UTF-16 units and no NUL".to_string(),
            ));
        }
    }
    Ok(StylePlan {
        font: crate::writer::StylesWriter::encode_font_payload(&style.font)?,
        fill: crate::writer::StylesWriter::encode_fill_payload(&style.fill)?,
        border: encode_authored_border(&style.border)?,
        number_format: style.number_format.clone(),
        xf_tail,
    })
}

fn encode_authored_border(border: &crate::styles::Border) -> Result<Vec<u8>> {
    if border.vertical.is_some() || border.horizontal.is_some() {
        return Err(Error::UnsupportedFeature(
            "vertical and horizontal table borders are not BrtBorder fields".to_string(),
        ));
    }
    if (border.diagonal_down || border.diagonal_up) && border.diagonal.is_none() {
        return Err(Error::InvalidFormat(
            "authored diagonal direction requires a diagonal border".to_string(),
        ));
    }
    let mut data = Vec::new();
    let mut writer = Writer::new(&mut data);
    writer.write_u8(u8::from(border.diagonal_down) | (u8::from(border.diagonal_up) << 1))?;
    write_authored_border_side(&mut writer, border.top.as_ref())?;
    write_authored_border_side(&mut writer, border.bottom.as_ref())?;
    write_authored_border_side(&mut writer, border.left.as_ref())?;
    write_authored_border_side(&mut writer, border.right.as_ref())?;
    write_authored_border_side(&mut writer, border.diagonal.as_ref())?;
    Ok(data)
}

fn write_authored_border_side<W: std::io::Write>(
    writer: &mut Writer<W>,
    side: Option<&crate::styles::BorderSide>,
) -> Result<()> {
    if let Some(side) = side {
        writer.write_u8(border_style_bits(side.style))?;
        writer.write_u8(0)?;
        write_authored_color(writer, side.color)?;
    } else {
        writer.write_u16(0)?;
        write_authored_color(writer, None)?;
    }
    Ok(())
}

const fn border_style_bits(style: crate::styles::BorderStyle) -> u8 {
    match style {
        crate::styles::BorderStyle::None => 0,
        crate::styles::BorderStyle::Thin => 1,
        crate::styles::BorderStyle::Medium => 2,
        crate::styles::BorderStyle::Dashed => 3,
        crate::styles::BorderStyle::Dotted => 4,
        crate::styles::BorderStyle::Thick => 5,
        crate::styles::BorderStyle::Double => 6,
        crate::styles::BorderStyle::Hair => 7,
        crate::styles::BorderStyle::MediumDashed => 8,
        crate::styles::BorderStyle::DashDot => 9,
        crate::styles::BorderStyle::MediumDashDot => 10,
        crate::styles::BorderStyle::DashDotDot => 11,
        crate::styles::BorderStyle::MediumDashDotDot => 12,
        crate::styles::BorderStyle::SlantDashDot => 13,
    }
}

fn write_authored_color<W: std::io::Write>(
    writer: &mut Writer<W>,
    color: Option<u32>,
) -> Result<()> {
    if let Some(argb) = color {
        writer.write_u8(5)?;
        writer.write_u8(0)?;
        writer.write_u16(0)?;
        writer.write_u8(argb.to_be_bytes()[1])?;
        writer.write_u8(argb.to_be_bytes()[2])?;
        writer.write_u8(argb.to_be_bytes()[3])?;
        writer.write_u8(argb.to_be_bytes()[0])?;
    } else {
        writer.write_u32(0)?;
        writer.write_u32(0)?;
    }
    Ok(())
}

const fn horizontal_bits(value: crate::styles::HorizontalAlignment) -> u8 {
    match value {
        crate::styles::HorizontalAlignment::General => 0,
        crate::styles::HorizontalAlignment::Left => 1,
        crate::styles::HorizontalAlignment::Center => 2,
        crate::styles::HorizontalAlignment::Right => 3,
        crate::styles::HorizontalAlignment::Fill => 4,
        crate::styles::HorizontalAlignment::Justify => 5,
        crate::styles::HorizontalAlignment::CenterContinuous => 6,
        crate::styles::HorizontalAlignment::Distributed => 7,
    }
}

const fn vertical_bits(value: crate::styles::VerticalAlignment) -> u8 {
    match value {
        crate::styles::VerticalAlignment::Top => 0,
        crate::styles::VerticalAlignment::Center => 1,
        crate::styles::VerticalAlignment::Bottom => 2,
        crate::styles::VerticalAlignment::Justify => 3,
        crate::styles::VerticalAlignment::Distributed => 4,
    }
}

pub(super) fn intern_style_plan(package: &mut OpcPackage, plan: &StylePlan) -> Result<StyleIndex> {
    ensure_styles_part(package)?;
    let uri = PackURI::new(STYLES_URI)?;
    let mut styles = package.get_part(&uri)?.blob().to_vec();
    let font = intern_payload(
        &mut styles,
        kind::BEGIN_FONTS,
        kind::END_FONTS,
        kind::FONT,
        &plan.font,
    )?;
    let fill = intern_payload(
        &mut styles,
        kind::BEGIN_FILLS,
        kind::END_FILLS,
        kind::FILL,
        &plan.fill,
    )?;
    let border = intern_payload(
        &mut styles,
        kind::BEGIN_BORDERS,
        kind::END_BORDERS,
        kind::BORDER,
        &plan.border,
    )?;
    let number_format = match &plan.number_format {
        Some(code) => intern_number_format(&mut styles, code)?,
        None => 0,
    };
    let mut xf = vec![0_u8; 16];
    xf[..2].copy_from_slice(&0_u16.to_le_bytes());
    xf[2..4].copy_from_slice(&number_format.to_le_bytes());
    xf[4..6].copy_from_slice(&font.to_le_bytes());
    xf[6..8].copy_from_slice(&fill.to_le_bytes());
    xf[8..10].copy_from_slice(&border.to_le_bytes());
    xf[10..16].copy_from_slice(&plan.xf_tail);
    let index = intern_payload(
        &mut styles,
        kind::BEGIN_CELL_XFS,
        kind::END_CELL_XFS,
        kind::XF,
        &xf,
    )?;
    package.get_part_mut(&uri)?.set_blob(styles);
    StyleIndex::new(u32::from(index))
}

fn ensure_styles_part(package: &mut OpcPackage) -> Result<()> {
    let uri = PackURI::new(STYLES_URI)?;
    if package.get_part(&uri).is_ok() {
        return Ok(());
    }
    let mut bytes = Vec::new();
    crate::writer::StylesWriter::new().write(&mut Writer::new(&mut bytes))?;
    package.try_add_part(Box::new(BlobPart::new(
        uri,
        "application/vnd.ms-excel.styles".to_string(),
        bytes,
    )))?;
    let workbook = package.get_part_mut(&PackURI::new(WORKBOOK_URI)?)?;
    let strict = workbook.rels().iter().any(|relationship| {
        relationship
            .reltype()
            .starts_with("http://purl.oclc.org/ooxml/")
    });
    workbook.rels_mut().get_or_add(
        if strict {
            litchi_opc::constants::relationship_type::STRICT_STYLES
        } else {
            litchi_opc::constants::relationship_type::STYLES
        },
        "styles.bin",
    );
    Ok(())
}

fn intern_payload(
    styles: &mut Vec<u8>,
    begin: Kind,
    end: Kind,
    item: Kind,
    payload: &[u8],
) -> Result<u16> {
    let values = collection(styles, begin, end, item)?;
    let index = if let Some(index) = values.iter().position(|value| value == payload) {
        index
    } else {
        *styles = append_collection_item(styles, begin, end, item, payload)?;
        values.len()
    };
    u16::try_from(index)
        .map_err(|_| Error::UnsupportedFeature("authored style resource index exceeds u16".into()))
}

fn intern_number_format(styles: &mut Vec<u8>, code: &str) -> Result<u16> {
    let formats = collection(styles, kind::BEGIN_FMTS, kind::END_FMTS, kind::FMT)?;
    for payload in &formats {
        let (id, existing) = crate::styles::parse_num_fmt(payload).map_err(map_style_error)?;
        if existing == code {
            return u16::try_from(id).map_err(|_| {
                Error::UnsupportedFeature("number-format ID exceeds u16".to_string())
            });
        }
    }
    let used = formats
        .iter()
        .filter_map(|payload| payload.get(..2))
        .map(|bytes| u16::from_le_bytes([bytes[0], bytes[1]]))
        .collect::<std::collections::BTreeSet<_>>();
    let id = (164_u16..=382)
        .find(|id| !used.contains(id))
        .ok_or_else(|| Error::UnsupportedFeature("no custom number-format ID remains".into()))?;
    let mut payload = id.to_le_bytes().to_vec();
    Writer::new(&mut payload).write_wide_string(code)?;
    *styles = append_collection_item(
        styles,
        kind::BEGIN_FMTS,
        kind::END_FMTS,
        kind::FMT,
        &payload,
    )?;
    Ok(id)
}

pub(super) fn intern_shared_string_for_new_cell(
    package: &mut OpcPackage,
    value: &SharedString,
) -> Result<u32> {
    let encoded = value.encode()?;
    let uri = PackURI::new(SST_URI)?;
    let existing = package.get_part(&uri).ok().map(|part| part.blob().to_vec());
    let (bytes, index) = match existing {
        Some(source) => append_or_reuse_sst(&source, value, &encoded)?,
        None => (new_sst(&encoded)?, 0),
    };
    if package.get_part(&uri).is_ok() {
        package.get_part_mut(&uri)?.set_blob(bytes);
    } else {
        let relationship_type = if package
            .get_part(&PackURI::new(WORKBOOK_URI)?)?
            .rels()
            .iter()
            .any(|relationship| {
                relationship.reltype() == litchi_opc::constants::relationship_type::STRICT_STYLES
                    || relationship
                        .reltype()
                        .starts_with("http://purl.oclc.org/ooxml/")
            }) {
            litchi_opc::constants::relationship_type::STRICT_SHARED_STRINGS
        } else {
            litchi_opc::constants::relationship_type::SHARED_STRINGS
        };
        package.try_add_part(Box::new(BlobPart::new(
            uri,
            SST_CONTENT_TYPE.to_string(),
            bytes,
        )))?;
        package
            .get_part_mut(&PackURI::new(WORKBOOK_URI)?)?
            .rels_mut()
            .get_or_add(relationship_type, "sharedStrings.bin");
    }
    Ok(index)
}

pub(super) fn transfer_style(
    source: &OpcPackage,
    target: &mut OpcPackage,
    source_index: StyleIndex,
) -> Result<StyleIndex> {
    let source_uri = PackURI::new(STYLES_URI)?;
    let Some(source_blob) = source
        .get_part(&source_uri)
        .ok()
        .map(|part| part.blob().to_vec())
    else {
        return if source_index.get() == 0 {
            StyleIndex::new(0)
        } else {
            Err(Error::UnsupportedFeature(
                "source workbook omits styles.bin for a nonzero style".to_string(),
            ))
        };
    };
    ensure_styles_part(target)?;
    let mut target_blob = target.get_part(&source_uri)?.blob().to_vec();
    let source_cell_xfs = collection(
        &source_blob,
        kind::BEGIN_CELL_XFS,
        kind::END_CELL_XFS,
        kind::XF,
    )?;
    let source_slot = usize::try_from(source_index.get()).map_err(|_| {
        Error::UnsupportedFeature("source style index does not fit this platform".to_string())
    })?;
    let source_xf = source_cell_xfs.get(source_slot).ok_or_else(|| {
        Error::UnsupportedFeature(format!(
            "source style index {} is absent",
            source_index.get()
        ))
    })?;
    let rewritten = transfer_xf_dependencies(&source_blob, &mut target_blob, source_xf, false)?;
    let target_cell_xfs = collection(
        &target_blob,
        kind::BEGIN_CELL_XFS,
        kind::END_CELL_XFS,
        kind::XF,
    )?;
    let index = if let Some(index) = target_cell_xfs
        .iter()
        .position(|payload| payload == &rewritten)
    {
        index
    } else {
        target_blob = append_collection_item(
            &target_blob,
            kind::BEGIN_CELL_XFS,
            kind::END_CELL_XFS,
            kind::XF,
            &rewritten,
        )?;
        target_cell_xfs.len()
    };
    let wire_index = u32::try_from(index).map_err(|_| {
        Error::UnsupportedFeature("transferred style index exceeds u32".to_string())
    })?;
    let style = StyleIndex::new(wire_index)?;
    target.get_part_mut(&source_uri)?.set_blob(target_blob);
    Ok(style)
}

pub(super) fn transfer_string_fonts(
    source: &OpcPackage,
    target: &mut OpcPackage,
    value: &SharedString,
) -> Result<SharedString> {
    if value.runs.is_empty() && value.phonetic.is_none() {
        return Ok(value.clone());
    }
    let styles_uri = PackURI::new(STYLES_URI)?;
    let source_blob = source.get_part(&styles_uri).map_err(|_| {
        Error::UnsupportedFeature(
            "source rich string references fonts but omits styles.bin".to_string(),
        )
    })?;
    let mut target_blob = target
        .get_part(&styles_uri)
        .map_err(|_| {
            Error::UnsupportedFeature(
                "target rich-string transfer requires an existing styles.bin".to_string(),
            )
        })?
        .blob()
        .to_vec();
    let mut transferred = value.clone();
    let mut mapped = std::collections::BTreeMap::new();
    for run in &mut transferred.runs {
        run.font_id = transfer_font_index(
            source_blob.blob(),
            &mut target_blob,
            run.font_id,
            &mut mapped,
        )?;
    }
    if let Some(phonetic) = &mut transferred.phonetic {
        phonetic.font_id = transfer_font_index(
            source_blob.blob(),
            &mut target_blob,
            phonetic.font_id,
            &mut mapped,
        )?;
    }
    target.get_part_mut(&styles_uri)?.set_blob(target_blob);
    Ok(transferred)
}

fn transfer_font_index(
    source: &[u8],
    target: &mut Vec<u8>,
    source_index: u16,
    mapped: &mut std::collections::BTreeMap<u16, u16>,
) -> Result<u16> {
    if let Some(index) = mapped.get(&source_index) {
        return Ok(*index);
    }
    let source_fonts = collection(source, kind::BEGIN_FONTS, kind::END_FONTS, kind::FONT)?;
    let payload = source_fonts.get(usize::from(source_index)).ok_or_else(|| {
        Error::UnsupportedFeature(format!("rich-string font index {source_index} is absent"))
    })?;
    let target_fonts = collection(target, kind::BEGIN_FONTS, kind::END_FONTS, kind::FONT)?;
    let target_index = if let Some(index) = target_fonts.iter().position(|font| font == payload) {
        index
    } else {
        *target = append_collection_item(
            target,
            kind::BEGIN_FONTS,
            kind::END_FONTS,
            kind::FONT,
            payload,
        )?;
        target_fonts.len()
    };
    let target_index = u16::try_from(target_index)
        .map_err(|_| Error::UnsupportedFeature("font index exceeds u16".to_string()))?;
    mapped.insert(source_index, target_index);
    Ok(target_index)
}

fn append_or_reuse_sst(
    source: &[u8],
    value: &SharedString,
    encoded: &[u8],
) -> Result<(Vec<u8>, u32)> {
    let mut total = None;
    let mut unique = None;
    let mut items = Vec::new();
    for item in Records::new(source) {
        let record = item?;
        match record.kind() {
            kind::BEGIN_SST => {
                if record.payload().len() != 8 {
                    return Err(Error::InvalidLength {
                        expected: 8,
                        found: record.payload().len(),
                    });
                }
                total = Some(binary::read_u32_le_at(record.payload(), 0)?);
                unique = Some(binary::read_u32_le_at(record.payload(), 4)?);
            },
            kind::SST_ITEM => items.push(SharedString::parse(record.payload())?),
            _ => {},
        }
    }
    let total = total.ok_or_else(|| Error::InvalidFormat("SST has no BrtBeginSst".to_string()))?;
    let unique =
        unique.ok_or_else(|| Error::InvalidFormat("SST has no unique count".to_string()))?;
    let new_total = total
        .checked_add(1)
        .filter(|count| *count <= MAX_SST_COUNT)
        .ok_or_else(|| {
            Error::UnsupportedFeature("shared-string occurrence count overflow".to_string())
        })?;
    let existing = items.iter().position(|item| item == value);
    let new_unique = if existing.is_some() {
        unique
    } else {
        unique
            .checked_add(1)
            .filter(|count| *count <= new_total)
            .ok_or_else(|| {
                Error::UnsupportedFeature("shared-string unique count overflow".to_string())
            })?
    };
    let index = existing.unwrap_or(items.len());
    let bytes = rewrite_sst(
        source,
        new_total,
        new_unique,
        existing.is_none().then_some(encoded),
    )?;
    let index = u32::try_from(index)
        .map_err(|_| Error::UnsupportedFeature("shared-string index exceeds u32".to_string()))?;
    Ok((bytes, index))
}

fn new_sst(encoded: &[u8]) -> Result<Vec<u8>> {
    let mut bytes = Vec::new();
    let mut header = Vec::new();
    Writer::new(&mut header).write_u32(1)?;
    Writer::new(&mut header).write_u32(1)?;
    let mut writer = Writer::new(&mut bytes);
    writer.write_record(kind::BEGIN_SST, &header)?;
    writer.write_record(kind::SST_ITEM, encoded)?;
    writer.write_record(kind::END_SST, &[])?;
    Ok(bytes)
}

fn rewrite_sst(source: &[u8], total: u32, unique: u32, appended: Option<&[u8]>) -> Result<Vec<u8>> {
    let mut output = Vec::new();
    for item in Records::new(source) {
        let record = item?;
        match record.kind() {
            kind::BEGIN_SST => {
                let mut payload = record.payload().to_vec();
                payload[..4].copy_from_slice(&total.to_le_bytes());
                payload[4..8].copy_from_slice(&unique.to_le_bytes());
                Writer::new(&mut output).write_record(kind::BEGIN_SST, &payload)?;
            },
            kind::END_SST => {
                if let Some(payload) = appended {
                    Writer::new(&mut output).write_record(kind::SST_ITEM, payload)?;
                }
                copy_record(source, &record, &mut output)?;
            },
            _ => copy_record(source, &record, &mut output)?,
        }
    }
    Ok(output)
}

fn transfer_xf_dependencies(
    source: &[u8],
    target: &mut Vec<u8>,
    source_xf: &[u8],
    style_xf: bool,
) -> Result<Vec<u8>> {
    if source_xf.len() < 16 {
        return Err(Error::InvalidLength {
            expected: 16,
            found: source_xf.len(),
        });
    }
    let mut rewritten = source_xf.to_vec();
    let font = transfer_indexed_resource(
        source,
        target,
        source_xf,
        kind::BEGIN_FONTS,
        kind::END_FONTS,
        kind::FONT,
        4,
    )?;
    let fill = transfer_indexed_resource(
        source,
        target,
        source_xf,
        kind::BEGIN_FILLS,
        kind::END_FILLS,
        kind::FILL,
        6,
    )?;
    let border = transfer_indexed_resource(
        source,
        target,
        source_xf,
        kind::BEGIN_BORDERS,
        kind::END_BORDERS,
        kind::BORDER,
        8,
    )?;
    rewritten[4..6].copy_from_slice(&font.to_le_bytes());
    rewritten[6..8].copy_from_slice(&fill.to_le_bytes());
    rewritten[8..10].copy_from_slice(&border.to_le_bytes());
    transfer_number_format(source, target, &mut rewritten)?;

    let parent = u16::from_le_bytes([source_xf[0], source_xf[1]]);
    if !style_xf && parent != u16::MAX {
        let source_parents = collection(
            source,
            kind::BEGIN_CELL_STYLE_XFS,
            kind::END_CELL_STYLE_XFS,
            kind::XF,
        )?;
        let parent_payload = source_parents.get(usize::from(parent)).ok_or_else(|| {
            Error::UnsupportedFeature(format!("source cell-style XF {parent} is absent"))
        })?;
        let transferred = transfer_xf_dependencies(source, target, parent_payload, true)?;
        let target_parents = collection(
            target,
            kind::BEGIN_CELL_STYLE_XFS,
            kind::END_CELL_STYLE_XFS,
            kind::XF,
        )?;
        let target_parent = if let Some(index) = target_parents
            .iter()
            .position(|value| value == &transferred)
        {
            index
        } else {
            *target = append_collection_item(
                target,
                kind::BEGIN_CELL_STYLE_XFS,
                kind::END_CELL_STYLE_XFS,
                kind::XF,
                &transferred,
            )?;
            target_parents.len()
        };
        let target_parent = u16::try_from(target_parent).map_err(|_| {
            Error::UnsupportedFeature("cell-style XF index exceeds u16".to_string())
        })?;
        rewritten[..2].copy_from_slice(&target_parent.to_le_bytes());
    }
    Ok(rewritten)
}

fn transfer_indexed_resource(
    source: &[u8],
    target: &mut Vec<u8>,
    source_xf: &[u8],
    begin: Kind,
    end: Kind,
    item: Kind,
    xf_offset: usize,
) -> Result<u16> {
    let source_items = collection(source, begin, end, item)?;
    let source_index = u16::from_le_bytes([
        source_items_index_byte(source_xf, xf_offset)?,
        source_items_index_byte(source_xf, xf_offset + 1)?,
    ]);
    let payload = source_items.get(usize::from(source_index)).ok_or_else(|| {
        Error::UnsupportedFeature(format!("style resource index {source_index} is absent"))
    })?;
    let target_items = collection(target, begin, end, item)?;
    let index = if let Some(index) = target_items.iter().position(|value| value == payload) {
        index
    } else {
        *target = append_collection_item(target, begin, end, item, payload)?;
        target_items.len()
    };
    u16::try_from(index)
        .map_err(|_| Error::UnsupportedFeature("style resource index exceeds u16".to_string()))
}

fn source_items_index_byte(source_xf: &[u8], offset: usize) -> Result<u8> {
    source_xf.get(offset).copied().ok_or(Error::InvalidLength {
        expected: offset.saturating_add(1),
        found: source_xf.len(),
    })
}

fn transfer_number_format(source: &[u8], target: &mut Vec<u8>, xf: &mut [u8]) -> Result<()> {
    let source_id = u16::from_le_bytes([xf[2], xf[3]]);
    let source_formats = collection(source, kind::BEGIN_FMTS, kind::END_FMTS, kind::FMT)?;
    let Some(source_payload) = source_formats
        .iter()
        .find(|payload| payload.get(..2) == Some(source_id.to_le_bytes().as_slice()))
    else {
        return Ok(());
    };
    let (_, source_code) = crate::styles::parse_num_fmt(source_payload).map_err(map_style_error)?;
    let target_formats = collection(target, kind::BEGIN_FMTS, kind::END_FMTS, kind::FMT)?;
    for payload in &target_formats {
        let (id, code) = crate::styles::parse_num_fmt(payload).map_err(map_style_error)?;
        if code == source_code {
            let id = u16::try_from(id).map_err(|_| {
                Error::UnsupportedFeature("number-format ID exceeds u16".to_string())
            })?;
            xf[2..4].copy_from_slice(&id.to_le_bytes());
            return Ok(());
        }
    }
    let used = target_formats
        .iter()
        .filter_map(|payload| payload.get(..2))
        .map(|bytes| u16::from_le_bytes([bytes[0], bytes[1]]))
        .collect::<std::collections::BTreeSet<_>>();
    let new_id = (164_u16..=382)
        .find(|id| !used.contains(id))
        .ok_or_else(|| {
            Error::UnsupportedFeature("no custom number-format ID remains".to_string())
        })?;
    let mut payload = source_payload.clone();
    payload[..2].copy_from_slice(&new_id.to_le_bytes());
    *target = append_collection_item(
        target,
        kind::BEGIN_FMTS,
        kind::END_FMTS,
        kind::FMT,
        &payload,
    )?;
    xf[2..4].copy_from_slice(&new_id.to_le_bytes());
    Ok(())
}

fn collection(source: &[u8], begin: Kind, end: Kind, item: Kind) -> Result<Vec<Vec<u8>>> {
    let mut in_collection = false;
    let mut values = Vec::new();
    for record in Records::new(source) {
        let record = record?;
        match record.kind() {
            found if found == begin => in_collection = true,
            found if found == end => in_collection = false,
            found if found == item && in_collection => values.push(record.payload().to_vec()),
            _ => {},
        }
    }
    Ok(values)
}

fn append_collection_item(
    source: &[u8],
    begin: Kind,
    end: Kind,
    item: Kind,
    payload: &[u8],
) -> Result<Vec<u8>> {
    let mut output = Vec::new();
    let mut in_collection = false;
    let mut inserted = false;
    for record in Records::new(source) {
        let record = record?;
        if record.kind() == begin {
            if record.payload().len() < 4 {
                return Err(Error::InvalidLength {
                    expected: 4,
                    found: record.payload().len(),
                });
            }
            let mut count = record.payload().to_vec();
            let next = binary::read_u32_le_at(&count, 0)?
                .checked_add(1)
                .ok_or_else(|| {
                    Error::UnsupportedFeature("style collection count overflow".to_string())
                })?;
            count[..4].copy_from_slice(&next.to_le_bytes());
            Writer::new(&mut output).write_record(begin, &count)?;
            in_collection = true;
        } else if record.kind() == end && in_collection {
            Writer::new(&mut output).write_record(item, payload)?;
            copy_record(source, &record, &mut output)?;
            in_collection = false;
            inserted = true;
        } else {
            copy_record(source, &record, &mut output)?;
        }
    }
    if !inserted {
        return Err(Error::InvalidFormat(format!(
            "style collection {begin}..{end} is absent"
        )));
    }
    Ok(output)
}

fn copy_record(source: &[u8], record: &crate::raw::Record<'_>, output: &mut Vec<u8>) -> Result<()> {
    let record_source = source.get(record.offset()..).ok_or_else(|| {
        Error::InvalidFormat("record offset is outside its source part".to_string())
    })?;
    let (_, header_len) = Header::parse(record_source, RawLimits::DEFAULT)?;
    let end = record
        .offset()
        .checked_add(header_len)
        .and_then(|offset| offset.checked_add(record.len()))
        .ok_or(Error::CapacityOverflow {
            resource: "resource record range",
        })?;
    output.extend_from_slice(source.get(record.offset()..end).ok_or_else(|| {
        Error::InvalidFormat("record range is outside its source part".to_string())
    })?);
    Ok(())
}

fn map_style_error(error: crate::styles::Error) -> Error {
    Error::InvalidFormat(format!("style resource transfer: {error}"))
}
