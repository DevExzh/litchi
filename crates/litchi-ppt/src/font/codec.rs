//! PowerPoint font record parsing.

use super::model::{EmbeddedFont, Font, FontCollection, FontCollections, FontEmbeddingFlags};
use crate::consts::RecordType;
use crate::package::{Error, Result};
use crate::records::Record;

impl FontCollections {
    /// Discover and parse both font collections below `root`.
    pub fn parse(root: &Record) -> Result<Self> {
        let mut base_records = Vec::new();
        collect_records(root, RecordType::FontCollection, &mut base_records);
        if base_records.len() > 1 {
            return Err(Error::Corrupted(
                "Record tree contains multiple base font collections".to_string(),
            ));
        }
        let base = base_records
            .first()
            .map(|record| FontCollection::parse(record))
            .transpose()?;

        let mut international = None;
        let mut embedding_flags = None;
        for record in root.versioned_binary_tag_records(10)? {
            match record.record_type {
                RecordType::FontCollection10 if international.is_some() => {
                    return Err(Error::Corrupted(
                        "Record tree contains multiple international font collections".to_string(),
                    ));
                },
                RecordType::FontCollection10 => {
                    international = Some(FontCollection::parse(&record)?);
                },
                RecordType::FontEmbedFlags10Atom if embedding_flags.is_some() => {
                    return Err(Error::Corrupted(
                        "Record tree contains multiple PowerPoint 10 font embedding flags"
                            .to_string(),
                    ));
                },
                RecordType::FontEmbedFlags10Atom => {
                    embedding_flags = Some(FontEmbeddingFlags::parse(&record)?);
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

impl FontEmbeddingFlags {
    /// Parse a `FontEmbedFlags10Atom` record.
    pub fn parse(record: &Record) -> Result<Self> {
        if record.record_type != RecordType::FontEmbedFlags10Atom
            || record.version != 0
            || record.instance != 0
            || record.data.len() != 4
        {
            return Err(Error::Corrupted(
                "FontEmbedFlags10Atom has an invalid record header or size".to_string(),
            ));
        }
        let raw = u32::from_le_bytes(record.data[..4].try_into().map_err(|_| {
            Error::Corrupted("FontEmbedFlags10Atom payload is truncated".to_string())
        })?);
        Ok(Self {
            raw,
            subset: raw & 0x01 != 0,
            subset_option_confirmed: raw & 0x02 != 0,
        })
    }
}

impl FontCollection {
    /// Parse a `FontCollectionContainer` or `FontCollection10Container`.
    pub fn parse(record: &Record) -> Result<Self> {
        let international = match record.record_type {
            RecordType::FontCollection => false,
            RecordType::FontCollection10 => true,
            _ => {
                return Err(Error::Corrupted(
                    "Record is not a PowerPoint font collection".to_string(),
                ));
            },
        };
        if record.version != 0x0f || record.instance != 0 {
            return Err(Error::Corrupted(
                "Font collection has an invalid record header".to_string(),
            ));
        }

        let children = Record::parse_sequence_strict(&record.data, "font collection")?;
        let mut fonts = Vec::new();
        let mut current = None;
        for child in children {
            match child.record_type {
                RecordType::FontEntityAtom => {
                    if let Some(font) = current.take() {
                        fonts.push(font);
                    }
                    let font = parse_font_entity(&child)?;
                    if fonts
                        .iter()
                        .any(|existing: &Font| existing.index == font.index)
                    {
                        return Err(Error::Corrupted(
                            "Font collection has a duplicate font index".to_string(),
                        ));
                    }
                    current = Some(font);
                },
                RecordType::FontEmbeddedData => {
                    let font = current.as_mut().ok_or_else(|| {
                        Error::Corrupted(
                            "Embedded font data precedes its FontEntityAtom".to_string(),
                        )
                    })?;
                    if child.version != 0 || child.instance > 3 {
                        return Err(Error::Corrupted(
                            "Embedded font data has an invalid record header".to_string(),
                        ));
                    }
                    let style = child.instance as u8;
                    if font
                        .embedded_fonts
                        .last()
                        .is_some_and(|previous| previous.style >= style)
                    {
                        return Err(Error::Corrupted(
                            "Embedded font facets are duplicated or out of order".to_string(),
                        ));
                    }
                    font.embedded_fonts.push(EmbeddedFont {
                        style,
                        data: child.data,
                    });
                },
                _ => {
                    return Err(Error::Corrupted(format!(
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

fn parse_font_entity(record: &Record) -> Result<Font> {
    if record.version != 0 || record.instance > 128 || record.data.len() != 68 {
        return Err(Error::Corrupted(
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
            Error::Corrupted("FontEntityAtom typeface is not null-terminated".to_string())
        })?;
    let name = String::from_utf16(&name_units[..terminator])
        .map_err(|_| Error::Corrupted("FontEntityAtom typeface is invalid UTF-16".to_string()))?;
    let charset = record.data[64];
    let font_flags = record.data[65];
    let font_type_flags = record.data[66];
    if font_type_flags & 0xf0 != 0 {
        return Err(Error::Corrupted(
            "FontEntityAtom has nonzero reserved font-type bits".to_string(),
        ));
    }
    Ok(Font {
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

fn collect_records<'a>(record: &'a Record, kind: RecordType, output: &mut Vec<&'a Record>) {
    if record.record_type == kind {
        output.push(record);
        return;
    }
    for child in &record.children {
        collect_records(child, kind, output);
    }
}
