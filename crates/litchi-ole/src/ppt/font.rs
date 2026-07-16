//! Legacy PowerPoint font collection parsing.

use super::package::{PptError, Result};
use super::records::PptRecord;
use crate::consts::PptRecordType;

/// One embedded OpenType font facet.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct EmbeddedPowerPointFont {
    /// Facet index: plain, bold, italic, or bold-italic (`0..=3`).
    pub style: u8,
    /// Embedded font bytes in the format specified by MS-PPT.
    pub data: Vec<u8>,
}

/// Font attributes from a `FontEntityAtom`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PowerPointFont {
    /// Zero-based font index from the atom's record instance.
    pub index: u16,
    /// Null-terminated UTF-16 typeface name.
    pub name: String,
    /// Windows character-set identifier.
    pub charset: u8,
    /// Raw font flags byte.
    pub font_flags: u8,
    /// Whether only a subset of the font is embedded.
    pub embedded_subset: bool,
    /// Raw four-bit font type flags.
    pub font_type_flags: u8,
    /// Whether this is a raster font.
    pub raster: bool,
    /// Whether this is a device font.
    pub device: bool,
    /// Whether this is a TrueType font.
    pub truetype: bool,
    /// Whether font substitution is disabled.
    pub no_substitution: bool,
    /// Windows pitch and family byte.
    pub pitch_and_family: u8,
    /// Optional embedded font facets in record order.
    pub embedded_fonts: Vec<EmbeddedPowerPointFont>,
}

/// Parsed base or PowerPoint 10 font collection.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PowerPointFontCollection {
    /// Whether this is the international `FontCollection10Container`.
    pub international: bool,
    /// Fonts in collection order.
    pub fonts: Vec<PowerPointFont>,
}

/// Base and international font collections resolved from a PPT record tree.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct PowerPointFontCollections {
    /// Base font collection from `DocumentTextInfoContainer`.
    pub base: Option<PowerPointFontCollection>,
    /// PowerPoint 10 international font collection from `___PPT10`.
    pub international: Option<PowerPointFontCollection>,
}

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
        for record in root.versioned_binary_tag_records(10)? {
            if record.record_type != PptRecordType::FontCollection10 {
                continue;
            }
            if international
                .replace(PowerPointFontCollection::parse(&record)?)
                .is_some()
            {
                return Err(PptError::Corrupted(
                    "Record tree contains multiple international font collections".to_string(),
                ));
            }
        }
        Ok(Self {
            base,
            international,
        })
    }

    /// Resolve a base-font reference.
    pub fn get_base(&self, index: u16) -> Option<&PowerPointFont> {
        self.base.as_ref()?.get(index)
    }

    /// Resolve a PowerPoint 10 international-font reference.
    pub fn get_international(&self, index: u16) -> Option<&PowerPointFont> {
        self.international.as_ref()?.get(index)
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

    /// Resolve a zero-based font reference.
    pub fn get(&self, index: u16) -> Option<&PowerPointFont> {
        self.fonts.iter().find(|font| font.index == index)
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

#[cfg(test)]
mod tests {
    use super::*;

    fn record_bytes(version: u16, instance: u16, kind: u16, payload: &[u8]) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&((instance << 4) | version).to_le_bytes());
        data.extend_from_slice(&kind.to_le_bytes());
        data.extend_from_slice(&(payload.len() as u32).to_le_bytes());
        data.extend_from_slice(payload);
        data
    }

    fn collection(kind: PptRecordType, payload: Vec<u8>) -> PptRecord {
        PptRecord {
            record_type: kind,
            record_type_raw: kind.as_u16(),
            version: 0x0f,
            instance: 0,
            data_length: payload.len() as u32,
            data: payload,
            children: Vec::new(),
        }
    }

    fn prog_tags_record(version: u8, blob_payload: &[u8]) -> PptRecord {
        let tag_name: Vec<u8> = format!("___PPT{version}")
            .encode_utf16()
            .flat_map(u16::to_le_bytes)
            .collect();
        let name = record_bytes(0, 0, 4026, &tag_name);
        let blob = record_bytes(0, 0, 0x138b, blob_payload);
        let mut tag_payload = name;
        tag_payload.extend_from_slice(&blob);
        let tag = record_bytes(0x0f, 0, 0x138a, &tag_payload);
        PptRecord {
            record_type: PptRecordType::ProgTags,
            record_type_raw: 0x1388,
            version: 0x0f,
            instance: 0,
            data_length: tag.len() as u32,
            data: tag,
            children: Vec::new(),
        }
    }

    #[test]
    fn parses_font_collections_and_embedded_facets() {
        let mut entity = vec![0u8; 68];
        for (index, unit) in "Noto Sans CJK"
            .encode_utf16()
            .chain(std::iter::once(0))
            .enumerate()
        {
            entity[index * 2..index * 2 + 2].copy_from_slice(&unit.to_le_bytes());
        }
        entity[64] = 0x80;
        entity[65] = 0x01;
        entity[66] = 0x0c;
        entity[67] = 0x22;
        let mut payload = record_bytes(0, 7, 4023, &entity);
        payload.extend_from_slice(&record_bytes(0, 0, 4024, b"plain-font"));
        payload.extend_from_slice(&record_bytes(0, 3, 4024, b"bold-italic-font"));

        let fonts =
            PowerPointFontCollection::parse(&collection(PptRecordType::FontCollection10, payload))
                .unwrap();

        assert!(fonts.international);
        let font = fonts.get(7).unwrap();
        assert_eq!(font.name, "Noto Sans CJK");
        assert_eq!(font.charset, 0x80);
        assert!(font.embedded_subset);
        assert!(font.truetype);
        assert!(font.no_substitution);
        assert_eq!(font.pitch_and_family, 0x22);
        assert_eq!(font.embedded_fonts.len(), 2);
        assert_eq!(font.embedded_fonts[1].style, 3);
    }

    #[test]
    fn rejects_malformed_font_collections() {
        let mut unterminated = vec![b'A'; 68];
        unterminated[66] = 0;
        let data = record_bytes(0, 0, 4023, &unterminated);
        assert!(
            PowerPointFontCollection::parse(&collection(PptRecordType::FontCollection, data,))
                .is_err()
        );

        let embedded_first = record_bytes(0, 0, 4024, b"font");
        assert!(
            PowerPointFontCollection::parse(&collection(
                PptRecordType::FontCollection,
                embedded_first,
            ))
            .is_err()
        );
    }

    #[test]
    fn resolves_base_and_international_font_collections() {
        let mut entity = vec![0u8; 68];
        entity[..4].copy_from_slice(&[b'A', 0, 0, 0]);
        entity[66] = 4;
        let base = collection(
            PptRecordType::FontCollection,
            record_bytes(0, 0, 4023, &entity),
        );
        let international_bytes = record_bytes(0, 9, 4023, &entity);
        let international = record_bytes(0x0f, 0, 2006, &international_bytes);
        let root = PptRecord {
            record_type: PptRecordType::Document,
            record_type_raw: 1000,
            version: 0x0f,
            instance: 0,
            data_length: 0,
            data: Vec::new(),
            children: vec![base, prog_tags_record(10, &international)],
        };

        let fonts = PowerPointFontCollections::parse(&root).unwrap();
        assert_eq!(fonts.get_base(0).unwrap().name, "A");
        assert_eq!(fonts.get_international(9).unwrap().name, "A");
    }
}
