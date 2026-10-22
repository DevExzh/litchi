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
use crate::raw::{Encoder, Header, Kind, Limits as RawLimits, Records, kind};
use litchi_core::binary;
use litchi_opc::{BlobPart, OpcError, OpcPackage, PackURI};

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

pub(super) fn plan_style(
    style: &super::structure::AuthoredStyle,
    limits: RawLimits,
) -> Result<StylePlan> {
    limits.validate()?;
    let alignment = style.alignment.as_ref();
    if alignment
        .is_some_and(|value| value.rotation > 180 || value.indent > 250 || value.text_direction > 2)
    {
        return Err(Error::InvalidFormat(
            "authored style alignment is outside BIFF12 bounds",
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
        if !(1..=255).contains(&units) || units > limits.string_units() || code.contains('\0') {
            return Err(Error::InvalidFormat(
                "authored number format must contain 1..=255 UTF-16 units and no NUL",
            ));
        }
    }
    if style.font.name.encode_utf16().count() > limits.string_units() {
        return Err(crate::raw::Error::StringLimit {
            units: style.font.name.encode_utf16().count(),
            limit: limits.string_units(),
            offset: 0,
        }
        .into());
    }
    let font = crate::writer::StylesWriter::encode_font_payload(&style.font)?;
    let fill = crate::writer::StylesWriter::encode_fill_payload(&style.fill)?;
    let border = encode_authored_border(&style.border, limits)?;
    for payload in [&font, &fill, &border] {
        if payload.len() > limits.payload() {
            return Err(crate::raw::Error::PayloadLimit {
                length: payload.len(),
                limit: limits.payload(),
                offset: 0,
            }
            .into());
        }
    }
    Ok(StylePlan {
        font,
        fill,
        border,
        number_format: style.number_format.clone(),
        xf_tail,
    })
}

fn encode_authored_border(border: &crate::styles::Border, limits: RawLimits) -> Result<Vec<u8>> {
    if border.vertical.is_some() || border.horizontal.is_some() {
        return Err(Error::refused(
            litchi_core::edit::RefusalKind::UnsupportedContent,
            "edit",
            "vertical and horizontal table borders are not BrtBorder fields",
        ));
    }
    if (border.diagonal_down || border.diagonal_up) && border.diagonal.is_none() {
        return Err(Error::InvalidFormat(
            "authored diagonal direction requires a diagonal border",
        ));
    }
    let mut data = Vec::new();
    let mut writer = Encoder::new(&mut data, limits);
    writer.write_u8(u8::from(border.diagonal_down) | (u8::from(border.diagonal_up) << 1))?;
    write_authored_border_side(&mut writer, border.top.as_ref())?;
    write_authored_border_side(&mut writer, border.bottom.as_ref())?;
    write_authored_border_side(&mut writer, border.left.as_ref())?;
    write_authored_border_side(&mut writer, border.right.as_ref())?;
    write_authored_border_side(&mut writer, border.diagonal.as_ref())?;
    Ok(data)
}

fn write_authored_border_side<W: std::io::Write>(
    writer: &mut Encoder<W>,
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

pub(super) const fn border_style_bits(style: crate::styles::BorderStyle) -> u8 {
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
    writer: &mut Encoder<W>,
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

pub(super) const fn horizontal_bits(value: crate::styles::HorizontalAlignment) -> u8 {
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

pub(super) const fn vertical_bits(value: crate::styles::VerticalAlignment) -> u8 {
    match value {
        crate::styles::VerticalAlignment::Top => 0,
        crate::styles::VerticalAlignment::Center => 1,
        crate::styles::VerticalAlignment::Bottom => 2,
        crate::styles::VerticalAlignment::Justify => 3,
        crate::styles::VerticalAlignment::Distributed => 4,
    }
}

pub(super) fn intern_style_plan(
    package: &mut OpcPackage,
    plan: &StylePlan,
    limits: RawLimits,
) -> Result<StyleIndex> {
    ensure_styles_part(package, limits)?;
    let uri = PackURI::new(STYLES_URI)?;
    let mut styles = package.get_part(&uri)?.blob().to_vec();
    let font = intern_payload(
        &mut styles,
        kind::BEGIN_FONTS,
        kind::END_FONTS,
        kind::FONT,
        &plan.font,
        limits,
    )?;
    let fill = intern_payload(
        &mut styles,
        kind::BEGIN_FILLS,
        kind::END_FILLS,
        kind::FILL,
        &plan.fill,
        limits,
    )?;
    let border = intern_payload(
        &mut styles,
        kind::BEGIN_BORDERS,
        kind::END_BORDERS,
        kind::BORDER,
        &plan.border,
        limits,
    )?;
    let number_format = match &plan.number_format {
        Some(code) => intern_number_format(&mut styles, code, limits)?,
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
        limits,
    )?;
    package.get_part_mut(&uri)?.set_blob(styles);
    StyleIndex::new(u32::from(index))
}

fn ensure_styles_part(package: &mut OpcPackage, limits: RawLimits) -> Result<()> {
    let uri = PackURI::new(STYLES_URI)?;
    let workbook_uri = PackURI::new(WORKBOOK_URI)?;
    let relationship_type = workbook_resource_relationship_type(package, false)?;
    let relationship_present =
        has_workbook_resource_relationship(package, relationship_type, "styles.bin")?;
    let missing = match package.get_part(&uri) {
        Ok(_) => false,
        Err(OpcError::PartNotFound(_)) => true,
        Err(error) => return Err(error.into()),
    };
    if missing {
        let mut bytes = Vec::new();
        crate::writer::StylesWriter::new().write(&mut Encoder::new(&mut bytes, limits))?;
        package.try_add_part(Box::new(BlobPart::new(
            uri,
            "application/vnd.ms-excel.styles".to_string(),
            bytes,
        )))?;
    }
    if !relationship_present {
        package
            .get_part_mut(&workbook_uri)?
            .rels_mut()
            .get_or_add(relationship_type, "styles.bin");
    }
    Ok(())
}

fn workbook_resource_relationship_type(
    package: &OpcPackage,
    shared_strings: bool,
) -> Result<&'static str> {
    let workbook_uri = PackURI::new(WORKBOOK_URI)?;
    let strict = package
        .get_part(&workbook_uri)?
        .rels()
        .iter()
        .any(|relationship| {
            relationship.reltype() == litchi_opc::constants::relationship_type::STRICT_STYLES
                || relationship
                    .reltype()
                    .starts_with("http://purl.oclc.org/ooxml/")
        });
    Ok(if shared_strings {
        if strict {
            litchi_opc::constants::relationship_type::STRICT_SHARED_STRINGS
        } else {
            litchi_opc::constants::relationship_type::SHARED_STRINGS
        }
    } else if strict {
        litchi_opc::constants::relationship_type::STRICT_STYLES
    } else {
        litchi_opc::constants::relationship_type::STYLES
    })
}

fn has_workbook_resource_relationship(
    package: &OpcPackage,
    relationship_type: &str,
    target: &str,
) -> Result<bool> {
    let workbook_uri = PackURI::new(WORKBOOK_URI)?;
    Ok(package
        .get_part(&workbook_uri)?
        .rels()
        .iter()
        .any(|relationship| {
            !relationship.is_external()
                && relationship.reltype() == relationship_type
                && relationship.target_ref() == target
        }))
}

pub(super) fn intern_payload(
    styles: &mut Vec<u8>,
    begin: Kind,
    end: Kind,
    item: Kind,
    payload: &[u8],
    limits: RawLimits,
) -> Result<u16> {
    let values = collection(styles, begin, end, item, limits)?;
    let index = if let Some(index) = values.iter().position(|value| value == payload) {
        index
    } else {
        *styles = append_collection_item(styles, begin, end, item, payload, limits)?;
        values.len()
    };
    u16::try_from(index).map_err(|_| {
        Error::refused(
            litchi_core::edit::RefusalKind::UnsupportedContent,
            "edit",
            "authored style resource index exceeds u16",
        )
    })
}

pub(super) fn intern_number_format(
    styles: &mut Vec<u8>,
    code: &str,
    limits: RawLimits,
) -> Result<u16> {
    let formats = collection(styles, kind::BEGIN_FMTS, kind::END_FMTS, kind::FMT, limits)?;
    for payload in &formats {
        let (id, existing) =
            crate::styles::parse_num_fmt(payload, limits).map_err(map_style_error)?;
        if existing == code {
            return u16::try_from(id).map_err(|_| {
                Error::refused(
                    litchi_core::edit::RefusalKind::UnsupportedContent,
                    "edit",
                    "number-format ID exceeds u16",
                )
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
        .ok_or_else(|| {
            Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                "edit",
                "no custom number-format ID remains",
            )
        })?;
    let mut payload = id.to_le_bytes().to_vec();
    Encoder::new(&mut payload, limits).write_wide_string(code)?;
    *styles = append_collection_item(
        styles,
        kind::BEGIN_FMTS,
        kind::END_FMTS,
        kind::FMT,
        &payload,
        limits,
    )?;
    Ok(id)
}

pub(super) fn intern_shared_string_for_new_cell(
    package: &mut OpcPackage,
    value: &SharedString,
    raw_limits: RawLimits,
) -> Result<u32> {
    let encoded = value.encode(raw_limits)?;
    let uri = PackURI::new(SST_URI)?;
    let existing = match package.get_part(&uri) {
        Ok(part) => Some(part.blob().to_vec()),
        Err(OpcError::PartNotFound(_)) => None,
        Err(error) => return Err(error.into()),
    };
    let workbook_uri = PackURI::new(WORKBOOK_URI)?;
    let relationship_type = workbook_resource_relationship_type(package, true)?;
    let relationship_present =
        has_workbook_resource_relationship(package, relationship_type, "sharedStrings.bin")?;
    let has_part = existing.is_some();
    let (bytes, index) = match existing {
        Some(source) => append_or_reuse_sst(&source, value, &encoded, raw_limits)?,
        None => (new_sst(&encoded, raw_limits)?, 0),
    };
    if has_part {
        package.get_part_mut(&uri)?.set_blob(bytes);
    } else {
        package.try_add_part(Box::new(BlobPart::new(
            uri,
            SST_CONTENT_TYPE.to_string(),
            bytes,
        )))?;
    }
    if !relationship_present {
        package
            .get_part_mut(&workbook_uri)?
            .rels_mut()
            .get_or_add(relationship_type, "sharedStrings.bin");
    }
    Ok(index)
}

pub(super) fn transfer_style(
    source: &OpcPackage,
    target: &mut OpcPackage,
    source_index: StyleIndex,
    source_limits: RawLimits,
    target_limits: RawLimits,
) -> Result<StyleIndex> {
    let source_uri = PackURI::new(STYLES_URI)?;
    let source_blob = match source.get_part(&source_uri) {
        Ok(part) => part.blob().to_vec(),
        Err(OpcError::PartNotFound(_)) => {
            return if source_index.get() == 0 {
                StyleIndex::new(0)
            } else {
                Err(Error::refused(
                    litchi_core::edit::RefusalKind::UnsupportedContent,
                    "edit",
                    "source workbook omits styles.bin for a nonzero style",
                ))
            };
        },
        Err(error) => return Err(error.into()),
    };
    ensure_styles_part(target, target_limits)?;
    let mut target_blob = target.get_part(&source_uri)?.blob().to_vec();
    let source_cell_xfs = collection(
        &source_blob,
        kind::BEGIN_CELL_XFS,
        kind::END_CELL_XFS,
        kind::XF,
        source_limits,
    )?;
    let source_slot = usize::try_from(source_index.get()).map_err(|_| {
        Error::refused(
            litchi_core::edit::RefusalKind::UnsupportedContent,
            "edit",
            "source style index does not fit this platform",
        )
    })?;
    let source_xf = source_cell_xfs.get(source_slot).ok_or_else(|| {
        Error::refused(
            litchi_core::edit::RefusalKind::UnsupportedContent,
            "edit",
            "source style index is absent",
        )
    })?;
    let rewritten = transfer_xf_dependencies(
        &source_blob,
        &mut target_blob,
        source_xf,
        false,
        source_limits,
        target_limits,
    )?;
    let target_cell_xfs = collection(
        &target_blob,
        kind::BEGIN_CELL_XFS,
        kind::END_CELL_XFS,
        kind::XF,
        target_limits,
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
            target_limits,
        )?;
        target_cell_xfs.len()
    };
    let wire_index = u32::try_from(index).map_err(|_| {
        Error::refused(
            litchi_core::edit::RefusalKind::UnsupportedContent,
            "edit",
            "transferred style index exceeds u32",
        )
    })?;
    let style = StyleIndex::new(wire_index)?;
    target.get_part_mut(&source_uri)?.set_blob(target_blob);
    Ok(style)
}

pub(super) fn transfer_string_fonts(
    source: &OpcPackage,
    target: &mut OpcPackage,
    value: &SharedString,
    source_limits: RawLimits,
    target_limits: RawLimits,
) -> Result<SharedString> {
    if value.runs.is_empty() && value.phonetic.is_none() {
        return Ok(value.clone());
    }
    let styles_uri = PackURI::new(STYLES_URI)?;
    let source_blob = source.get_part(&styles_uri).map_err(|_| {
        Error::refused(
            litchi_core::edit::RefusalKind::UnsupportedContent,
            "edit",
            "source rich string references fonts but omits styles.bin",
        )
    })?;
    let mut target_blob = target
        .get_part(&styles_uri)
        .map_err(|_| {
            Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                "edit",
                "target rich-string transfer requires an existing styles.bin",
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
            source_limits,
            target_limits,
        )?;
    }
    if let Some(phonetic) = &mut transferred.phonetic {
        phonetic.font_id = transfer_font_index(
            source_blob.blob(),
            &mut target_blob,
            phonetic.font_id,
            &mut mapped,
            source_limits,
            target_limits,
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
    source_limits: RawLimits,
    target_limits: RawLimits,
) -> Result<u16> {
    if let Some(index) = mapped.get(&source_index) {
        return Ok(*index);
    }
    let source_fonts = collection(
        source,
        kind::BEGIN_FONTS,
        kind::END_FONTS,
        kind::FONT,
        source_limits,
    )?;
    let payload = source_fonts.get(usize::from(source_index)).ok_or_else(|| {
        Error::refused(
            litchi_core::edit::RefusalKind::UnsupportedContent,
            "edit",
            "rich-string font index is absent",
        )
    })?;
    let target_fonts = collection(
        target,
        kind::BEGIN_FONTS,
        kind::END_FONTS,
        kind::FONT,
        target_limits,
    )?;
    let target_index = if let Some(index) = target_fonts.iter().position(|font| font == payload) {
        index
    } else {
        *target = append_collection_item(
            target,
            kind::BEGIN_FONTS,
            kind::END_FONTS,
            kind::FONT,
            payload,
            target_limits,
        )?;
        target_fonts.len()
    };
    let target_index = u16::try_from(target_index).map_err(|_| {
        Error::refused(
            litchi_core::edit::RefusalKind::UnsupportedContent,
            "edit",
            "font index exceeds u16",
        )
    })?;
    mapped.insert(source_index, target_index);
    Ok(target_index)
}

fn append_or_reuse_sst(
    source: &[u8],
    value: &SharedString,
    encoded: &[u8],
    raw_limits: RawLimits,
) -> Result<(Vec<u8>, u32)> {
    let mut total = None;
    let mut unique = None;
    let mut items = Vec::new();
    for item in Records::new(source, raw_limits) {
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
            kind::SST_ITEM => items.push(SharedString::parse(record.payload(), raw_limits)?),
            _ => {},
        }
    }
    let total = total.ok_or(Error::InvalidFormat("SST has no BrtBeginSst"))?;
    let unique = unique.ok_or(Error::InvalidFormat("SST has no unique count"))?;
    let new_total = total
        .checked_add(1)
        .filter(|count| *count <= MAX_SST_COUNT)
        .ok_or_else(|| {
            Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                "edit",
                "shared-string occurrence count overflow",
            )
        })?;
    let existing = items.iter().position(|item| item == value);
    let new_unique = if existing.is_some() {
        unique
    } else {
        unique
            .checked_add(1)
            .filter(|count| *count <= new_total)
            .ok_or_else(|| {
                Error::refused(
                    litchi_core::edit::RefusalKind::UnsupportedContent,
                    "edit",
                    "shared-string unique count overflow",
                )
            })?
    };
    let index = existing.unwrap_or(items.len());
    let bytes = rewrite_sst(
        source,
        new_total,
        new_unique,
        existing.is_none().then_some(encoded),
        raw_limits,
    )?;
    let index = u32::try_from(index).map_err(|_| {
        Error::refused(
            litchi_core::edit::RefusalKind::UnsupportedContent,
            "edit",
            "shared-string index exceeds u32",
        )
    })?;
    Ok((bytes, index))
}

fn new_sst(encoded: &[u8], raw_limits: RawLimits) -> Result<Vec<u8>> {
    let mut bytes = Vec::new();
    let mut header = Vec::new();
    Encoder::new(&mut header, raw_limits).write_u32(1)?;
    Encoder::new(&mut header, raw_limits).write_u32(1)?;
    let mut writer = Encoder::new(&mut bytes, raw_limits);
    writer.write_record(kind::BEGIN_SST, &header)?;
    writer.write_record(kind::SST_ITEM, encoded)?;
    writer.write_record(kind::END_SST, &[])?;
    Ok(bytes)
}

fn rewrite_sst(
    source: &[u8],
    total: u32,
    unique: u32,
    appended: Option<&[u8]>,
    raw_limits: RawLimits,
) -> Result<Vec<u8>> {
    let mut output = Vec::new();
    for item in Records::new(source, raw_limits) {
        let record = item?;
        match record.kind() {
            kind::BEGIN_SST => {
                let mut payload = record.payload().to_vec();
                payload[..4].copy_from_slice(&total.to_le_bytes());
                payload[4..8].copy_from_slice(&unique.to_le_bytes());
                Encoder::new(&mut output, raw_limits).write_record(kind::BEGIN_SST, &payload)?;
            },
            kind::END_SST => {
                if let Some(payload) = appended {
                    Encoder::new(&mut output, raw_limits).write_record(kind::SST_ITEM, payload)?;
                }
                copy_record(source, &record, raw_limits, &mut output)?;
            },
            _ => copy_record(source, &record, raw_limits, &mut output)?,
        }
    }
    Ok(output)
}

fn transfer_xf_dependencies(
    source: &[u8],
    target: &mut Vec<u8>,
    source_xf: &[u8],
    style_xf: bool,
    source_limits: RawLimits,
    target_limits: RawLimits,
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
        source_limits,
        target_limits,
    )?;
    let fill = transfer_indexed_resource(
        source,
        target,
        source_xf,
        kind::BEGIN_FILLS,
        kind::END_FILLS,
        kind::FILL,
        6,
        source_limits,
        target_limits,
    )?;
    let border = transfer_indexed_resource(
        source,
        target,
        source_xf,
        kind::BEGIN_BORDERS,
        kind::END_BORDERS,
        kind::BORDER,
        8,
        source_limits,
        target_limits,
    )?;
    rewritten[4..6].copy_from_slice(&font.to_le_bytes());
    rewritten[6..8].copy_from_slice(&fill.to_le_bytes());
    rewritten[8..10].copy_from_slice(&border.to_le_bytes());
    transfer_number_format(source, target, &mut rewritten, source_limits, target_limits)?;

    let parent = u16::from_le_bytes([source_xf[0], source_xf[1]]);
    if !style_xf && parent != u16::MAX {
        let source_parents = collection(
            source,
            kind::BEGIN_CELL_STYLE_XFS,
            kind::END_CELL_STYLE_XFS,
            kind::XF,
            source_limits,
        )?;
        let parent_payload = source_parents.get(usize::from(parent)).ok_or_else(|| {
            Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                "edit",
                "source cell-style XF is absent",
            )
        })?;
        let transferred = transfer_xf_dependencies(
            source,
            target,
            parent_payload,
            true,
            source_limits,
            target_limits,
        )?;
        let target_parents = collection(
            target,
            kind::BEGIN_CELL_STYLE_XFS,
            kind::END_CELL_STYLE_XFS,
            kind::XF,
            target_limits,
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
                target_limits,
            )?;
            target_parents.len()
        };
        let target_parent = u16::try_from(target_parent).map_err(|_| {
            Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                "edit",
                "cell-style XF index exceeds u16",
            )
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
    source_limits: RawLimits,
    target_limits: RawLimits,
) -> Result<u16> {
    let source_items = collection(source, begin, end, item, source_limits)?;
    let source_index = u16::from_le_bytes([
        source_items_index_byte(source_xf, xf_offset)?,
        source_items_index_byte(source_xf, xf_offset + 1)?,
    ]);
    let payload = source_items.get(usize::from(source_index)).ok_or_else(|| {
        Error::refused(
            litchi_core::edit::RefusalKind::UnsupportedContent,
            "edit",
            "style resource index is absent",
        )
    })?;
    let target_items = collection(target, begin, end, item, target_limits)?;
    let index = if let Some(index) = target_items.iter().position(|value| value == payload) {
        index
    } else {
        *target = append_collection_item(target, begin, end, item, payload, target_limits)?;
        target_items.len()
    };
    u16::try_from(index).map_err(|_| {
        Error::refused(
            litchi_core::edit::RefusalKind::UnsupportedContent,
            "edit",
            "style resource index exceeds u16",
        )
    })
}

fn source_items_index_byte(source_xf: &[u8], offset: usize) -> Result<u8> {
    source_xf.get(offset).copied().ok_or(Error::InvalidLength {
        expected: offset.saturating_add(1),
        found: source_xf.len(),
    })
}

fn transfer_number_format(
    source: &[u8],
    target: &mut Vec<u8>,
    xf: &mut [u8],
    source_limits: RawLimits,
    target_limits: RawLimits,
) -> Result<()> {
    let source_id = u16::from_le_bytes([xf[2], xf[3]]);
    let source_formats = collection(
        source,
        kind::BEGIN_FMTS,
        kind::END_FMTS,
        kind::FMT,
        source_limits,
    )?;
    let Some(source_payload) = source_formats
        .iter()
        .find(|payload| payload.get(..2) == Some(source_id.to_le_bytes().as_slice()))
    else {
        return Ok(());
    };
    let (_, source_code) =
        crate::styles::parse_num_fmt(source_payload, source_limits).map_err(map_style_error)?;
    let target_formats = collection(
        target,
        kind::BEGIN_FMTS,
        kind::END_FMTS,
        kind::FMT,
        target_limits,
    )?;
    for payload in &target_formats {
        let (id, code) =
            crate::styles::parse_num_fmt(payload, target_limits).map_err(map_style_error)?;
        if code == source_code {
            let id = u16::try_from(id).map_err(|_| {
                Error::refused(
                    litchi_core::edit::RefusalKind::UnsupportedContent,
                    "edit",
                    "number-format ID exceeds u16",
                )
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
            Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                "edit",
                "no custom number-format ID remains",
            )
        })?;
    let mut payload = source_payload.clone();
    payload[..2].copy_from_slice(&new_id.to_le_bytes());
    *target = append_collection_item(
        target,
        kind::BEGIN_FMTS,
        kind::END_FMTS,
        kind::FMT,
        &payload,
        target_limits,
    )?;
    xf[2..4].copy_from_slice(&new_id.to_le_bytes());
    Ok(())
}

pub(crate) fn collection(
    source: &[u8],
    begin: Kind,
    end: Kind,
    item: Kind,
    limits: RawLimits,
) -> Result<Vec<Vec<u8>>> {
    let mut in_collection = false;
    let mut values = Vec::new();
    for record in Records::new(source, limits) {
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

pub(crate) fn append_collection_item(
    source: &[u8],
    begin: Kind,
    end: Kind,
    item: Kind,
    payload: &[u8],
    limits: RawLimits,
) -> Result<Vec<u8>> {
    let mut output = Vec::new();
    let mut in_collection = false;
    let mut inserted = false;
    for record in Records::new(source, limits) {
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
                    Error::refused(
                        litchi_core::edit::RefusalKind::UnsupportedContent,
                        "edit",
                        "style collection count overflow",
                    )
                })?;
            count[..4].copy_from_slice(&next.to_le_bytes());
            Encoder::new(&mut output, limits).write_record(begin, &count)?;
            in_collection = true;
        } else if record.kind() == end && in_collection {
            Encoder::new(&mut output, limits).write_record(item, payload)?;
            copy_record(source, &record, limits, &mut output)?;
            in_collection = false;
            inserted = true;
        } else {
            copy_record(source, &record, limits, &mut output)?;
        }
    }
    if !inserted {
        return Err(Error::InvalidFormat("style collection.. is absent"));
    }
    Ok(output)
}

fn copy_record(
    source: &[u8],
    record: &crate::raw::Record<'_>,
    limits: RawLimits,
    output: &mut Vec<u8>,
) -> Result<()> {
    let record_source = source.get(record.offset()..).ok_or(Error::InvalidFormat(
        "record offset is outside its source part",
    ))?;
    let (_, header_len) = Header::parse(record_source, limits)?;
    let end = record
        .offset()
        .checked_add(header_len)
        .and_then(|offset| offset.checked_add(record.len()))
        .ok_or(Error::CapacityOverflow {
            resource: "resource record range",
        })?;
    output.extend_from_slice(
        source
            .get(record.offset()..end)
            .ok_or(Error::InvalidFormat(
                "record range is outside its source part",
            ))?,
    );
    Ok(())
}

fn map_style_error(error: crate::styles::Error) -> Error {
    match error {
        crate::styles::Error::Wire(error) => Error::Wire(error),
        crate::styles::Error::InvalidLength { expected, found } => {
            Error::InvalidLength { expected, found }
        },
        crate::styles::Error::Unrecognized { typ, val } => Error::Unrecognized { typ, val },
        crate::styles::Error::Allocation { resource, source } => {
            Error::Allocation { resource, source }
        },
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn package_without_resource_relationship(
        package: &mut OpcPackage,
        relationship_type: &str,
        target: &str,
    ) {
        let workbook_uri = PackURI::new(WORKBOOK_URI).unwrap();
        let ids = package
            .get_part(&workbook_uri)
            .unwrap()
            .rels()
            .iter()
            .filter(|relationship| {
                !relationship.is_external()
                    && relationship.reltype() == relationship_type
                    && relationship.target_ref() == target
            })
            .map(|relationship| relationship.r_id().to_string())
            .collect::<Vec<_>>();
        let workbook = package.get_part_mut(&workbook_uri).unwrap();
        for id in ids {
            workbook.rels_mut().remove(&id);
        }
    }

    #[test]
    fn existing_styles_part_without_relationship_is_relinked() {
        let mut package = crate::Package::new().unwrap().into_opc();
        let relationship_type = litchi_opc::constants::relationship_type::STYLES;
        package_without_resource_relationship(&mut package, relationship_type, "styles.bin");

        assert!(
            !has_workbook_resource_relationship(&package, relationship_type, "styles.bin").unwrap()
        );
        assert!(package.get_part(&PackURI::new(STYLES_URI).unwrap()).is_ok());

        ensure_styles_part(&mut package, RawLimits::DEFAULT).unwrap();

        assert!(
            has_workbook_resource_relationship(&package, relationship_type, "styles.bin").unwrap()
        );
    }

    #[test]
    fn style_collection_append_preserves_unknown_records_under_explicit_limits() {
        let limits = RawLimits::new(128, 32);
        let unknown = Kind::new(0x1234).unwrap();
        let mut source = Vec::new();
        let mut writer = Encoder::new(&mut source, limits);
        writer
            .write_record(kind::BEGIN_FMTS, &1_u32.to_le_bytes())
            .unwrap();
        writer.write_record(unknown, &[0xA5, 0x00, 0x7F]).unwrap();
        writer.write_record(kind::END_FMTS, &[]).unwrap();

        let output = append_collection_item(
            &source,
            kind::BEGIN_FMTS,
            kind::END_FMTS,
            kind::FMT,
            &[164, 0, 0, 0, 0, 0],
            limits,
        )
        .unwrap();
        assert!(output.windows(3).any(|window| window == [0xA5, 0x00, 0x7F]));
    }

    #[test]
    fn existing_shared_strings_part_without_relationship_is_relinked() {
        let mut package = crate::Package::new().unwrap().into_opc();
        let value = SharedString {
            text: "existing".to_string(),
            runs: Vec::new(),
            phonetic: None,
        };
        let shared_strings_uri = PackURI::new(SST_URI).unwrap();
        package
            .try_add_part(Box::new(BlobPart::new(
                shared_strings_uri,
                SST_CONTENT_TYPE.to_string(),
                new_sst(
                    &value.encode(RawLimits::DEFAULT).unwrap(),
                    RawLimits::DEFAULT,
                )
                .unwrap(),
            )))
            .unwrap();
        let relationship_type = litchi_opc::constants::relationship_type::SHARED_STRINGS;
        package_without_resource_relationship(&mut package, relationship_type, "sharedStrings.bin");
        assert!(
            !has_workbook_resource_relationship(&package, relationship_type, "sharedStrings.bin")
                .unwrap()
        );

        assert_eq!(
            intern_shared_string_for_new_cell(&mut package, &value, RawLimits::DEFAULT).unwrap(),
            0
        );

        assert!(
            has_workbook_resource_relationship(&package, relationship_type, "sharedStrings.bin")
                .unwrap()
        );
    }
}
