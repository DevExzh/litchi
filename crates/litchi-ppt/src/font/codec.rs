//! PowerPoint font record parsing.

use super::model::{
    EmbeddedPowerPointFont, PowerPointFont, PowerPointFontCollection, PowerPointFontCollections,
    PowerPointFontEmbeddingFlags,
};
use crate::consts::PptRecordType;
use crate::package::{PptError, Result};
use crate::records::PptRecord;

impl PowerPointFontCollections {
    /// Discover and parse both font collections below `root`.
    pub fn parse(root: &PptRecord) -> Result<Self> {
        let mut base_records = Vec::new();
        collect_records(root, PptRecordType::FontCollection, &mut base_records);
        if base_records.len() > 1 {
            return Err(PptError::Corrupted(
                "Record tree contains multiple base font collections".to_string(),
            ));
        }
        let base = base_records
            .first()
            .map(|record| PowerPointFontCollection::parse(record))
            .transpose()?;

        let mut international = None;
        let mut embedding_flags = None;
        for record in root.versioned_binary_tag_records(10)? {
            match record.record_type {
                PptRecordType::FontCollection10 if international.is_some() => {
                    return Err(PptError::Corrupted(
                        "Record tree contains multiple international font collections".to_string(),
                    ));
                },
                PptRecordType::FontCollection10 => {
                    international = Some(PowerPointFontCollection::parse(&record)?);
                },
                PptRecordType::FontEmbedFlags10Atom if embedding_flags.is_some() => {
                    return Err(PptError::Corrupted(
                        "Record tree contains multiple PowerPoint 10 font embedding flags"
                            .to_string(),
                    ));
                },
                PptRecordType::FontEmbedFlags10Atom => {
                    embedding_flags = Some(PowerPointFontEmbeddingFlags::parse(&record)?);
                },
                _ => {},
            }
        }
        Ok(Self {
            base,
            international,
            embedding_flags,
        })
    }
}

impl PowerPointFontEmbeddingFlags {
    /// Parse a `FontEmbedFlags10Atom` record.
    pub fn parse(record: &PptRecord) -> Result<Self> {
        if record.record_type != PptRecordType::FontEmbedFlags10Atom
            || record.version != 0
            || record.instance != 0
            || record.data.len() != 4
        {
            return Err(PptError::Corrupted(
                "FontEmbedFlags10Atom has an invalid record header or size".to_string(),
            ));
        }
        let raw = u32::from_le_bytes(record.data[..4].try_into().map_err(|_| {
            PptError::Corrupted("FontEmbedFlags10Atom payload is truncated".to_string())
        })?);
        Ok(Self {
            raw,
            subset: raw & 0x01 != 0,
            subset_option_confirmed: raw & 0x02 != 0,
        })
    }
}

impl PowerPointFontCollection {
    /// Parse a `FontCollectionContainer` or `FontCollection10Container`.
    pub fn parse(record: &PptRecord) -> Result<Self> {
        let international = match record.record_type {
            PptRecordType::FontCollection => false,
            PptRecordType::FontCollection10 => true,
            _ => {
                return Err(PptError::Corrupted(
                    "Record is not a PowerPoint font collection".to_string(),
                ));
            },
        };
        if record.version != 0x0f || record.instance != 0 {
            return Err(PptError::Corrupted(
                "Font collection has an invalid record header".to_string(),
            ));
        }

        let children = PptRecord::parse_sequence_strict(&record.data, "font collection")?;
        let mut fonts = Vec::new();
        let mut current = None;
        for child in children {
            match child.record_type {
                PptRecordType::FontEntityAtom => {
                    if let Some(font) = current.take() {
                        fonts.push(font);
                    }
                    let font = parse_font_entity(&child)?;
                    if fonts
                        .iter()
                        .any(|existing: &PowerPointFont| existing.index == font.index)
                    {
                        return Err(PptError::Corrupted(
                            "Font collection has a duplicate font index".to_string(),
                        ));
                    }
                    current = Some(font);
                },
                PptRecordType::FontEmbeddedData => {
                    let font = current.as_mut().ok_or_else(|| {
                        PptError::Corrupted(
                            "Embedded font data precedes its FontEntityAtom".to_string(),
                        )
                    })?;
                    if child.version != 0 || child.instance > 3 {
                        return Err(PptError::Corrupted(
                            "Embedded font data has an invalid record header".to_string(),
                        ));
                    }
                    let style = child.instance as u8;
                    if font
                        .embedded_fonts
                        .last()
                        .is_some_and(|previous| previous.style >= style)
                    {
                        return Err(PptError::Corrupted(
                            "Embedded font facets are duplicated or out of order".to_string(),
                        ));
                    }
                    font.embedded_fonts.push(EmbeddedPowerPointFont {
                        style,
                        data: child.data,
                    });
                },
                _ => {
                    return Err(PptError::Corrupted(format!(
                        "Unexpected {:?} record in font collection",
                        child.record_type
                    )));
                },
            }
        }
        if let Some(font) = current {
            fonts.push(font);
        }
        Ok(Self {
            international,
            fonts,
        })
    }
}

fn parse_font_entity(record: &PptRecord) -> Result<PowerPointFont> {
    if record.version != 0 || record.instance > 128 || record.data.len() != 68 {
        return Err(PptError::Corrupted(
            "FontEntityAtom has an invalid record header or size".to_string(),
        ));
    }
    let name_units: Vec<u16> = record.data[..64]
        .chunks_exact(2)
        .map(|bytes| u16::from_le_bytes([bytes[0], bytes[1]]))
        .collect();
    let terminator = name_units
        .iter()
        .position(|unit| *unit == 0)
        .ok_or_else(|| {
            PptError::Corrupted("FontEntityAtom typeface is not null-terminated".to_string())
        })?;
    let name = String::from_utf16(&name_units[..terminator]).map_err(|_| {
        PptError::Corrupted("FontEntityAtom typeface is invalid UTF-16".to_string())
    })?;
    let charset = record.data[64];
    let font_flags = record.data[65];
    let font_type_flags = record.data[66];
    if font_type_flags & 0xf0 != 0 {
        return Err(PptError::Corrupted(
            "FontEntityAtom has nonzero reserved font-type bits".to_string(),
        ));
    }
    Ok(PowerPointFont {
        index: record.instance,
        name,
        charset,
        font_flags,
        embedded_subset: font_flags & 0x01 != 0,
        font_type_flags,
        raster: font_type_flags & 0x01 != 0,
        device: font_type_flags & 0x02 != 0,
        truetype: font_type_flags & 0x04 != 0,
        no_substitution: font_type_flags & 0x08 != 0,
        pitch_and_family: record.data[67],
        embedded_fonts: Vec::new(),
    })
}

fn collect_records<'a>(
    record: &'a PptRecord,
    kind: PptRecordType,
    output: &mut Vec<&'a PptRecord>,
) {
    if record.record_type == kind {
        output.push(record);
        return;
    }
    for child in &record.children {
        collect_records(child, kind, output);
    }
}
