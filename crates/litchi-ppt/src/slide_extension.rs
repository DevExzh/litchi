//! PowerPoint 12 slide and master round-trip metadata.

use super::package::{PptError, Result};
use super::records::PptRecord;
use crate::consts::PptRecordType;

/// Default visibility of header and footer placeholders on newly created slides.
///
/// MS-PPT stores these flags on main, title, handout, and notes masters. They are defaults for
/// new slides rather than the resolved visibility of placeholders on an existing slide.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct HeaderFooterDefaults {
    /// Include a date placeholder in the footer of new slides.
    pub include_date: bool,
    /// Include a footer placeholder on new slides.
    pub include_footer: bool,
    /// Include a header placeholder on new slides.
    pub include_header: bool,
    /// Include a slide-number or page-number placeholder in the footer of new slides.
    pub include_slide_number: bool,
}

/// Header or footer placeholder identity retained for PowerPoint 12 round trips.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum HeaderFooterPlaceholder {
    /// Master date placeholder (`PT_MasterDate`).
    Date,
    /// Master slide- or page-number placeholder (`PT_MasterSlideNumber`).
    SlideNumber,
    /// Master footer placeholder (`PT_MasterFooter`).
    Footer,
    /// Master header placeholder (`PT_MasterHeader`).
    Header,
}

/// Placeholder identities added by PowerPoint 12 and retained for round trips.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum NewPlaceholder {
    /// Vertical object placeholder (`PT_VerticalObject`).
    VerticalObject,
    /// Picture placeholder (`PT_Picture`).
    Picture,
}

/// Application-defined checksums retained for PowerPoint 12 custom-layout round trips.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ShapeChecksums {
    /// Checksum used to detect changes to OfficeArt shape properties.
    pub shape: u32,
    /// Checksum used to detect changes to the shape's text body.
    pub text: u32,
}

/// Inert PowerPoint 12 shape metadata stored in OfficeArt client data.
///
/// MS-PPT recommends ignoring these compatibility atoms while preserving them. The legacy
/// `PlaceholderAtom`, when present, remains authoritative for the resolved placeholder type.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct ShapeMetadata {
    /// Header/footer identity from `RoundTripHFPlaceholder12Atom`.
    pub header_footer: Option<HeaderFooterPlaceholder>,
    /// New identity from `RoundTripNewPlaceholderId12Atom`.
    pub new_placeholder: Option<NewPlaceholder>,
    /// Shape identifier from `RoundTripShapeId12Atom`.
    pub shape_id: Option<u32>,
    /// Application-defined shape and text checksums for custom layouts.
    pub custom_layout_checksums: Option<ShapeChecksums>,
}

/// Metadata stored in a `___PPT12` slide programmable-tag extension.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct SlideExtension {
    /// Optional header and footer defaults for a master slide.
    pub header_footer_defaults: Option<HeaderFooterDefaults>,
}

impl SlideExtension {
    /// Discover and parse slide metadata from every `___PPT12` programmable tag below `root`.
    pub fn parse(root: &PptRecord) -> Result<Self> {
        let mut extension = Self::default();
        for record in root.versioned_binary_tag_records(12)? {
            if record.record_type != PptRecordType::RoundTripHeaderFooterDefaults12Atom {
                continue;
            }
            if extension.header_footer_defaults.is_some() {
                return Err(PptError::Corrupted(
                    "PowerPoint 12 slide extension contains duplicate header/footer defaults"
                        .to_string(),
                ));
            }
            if record.version != 0 || record.instance != 0 || record.data.len() != 1 {
                return Err(PptError::Corrupted(
                    "RoundTripHeaderFooterDefaults12Atom has an invalid record header or size"
                        .to_string(),
                ));
            }
            let flags = record.data[0];
            if flags & 0xf0 != 0 {
                return Err(PptError::Corrupted(
                    "RoundTripHeaderFooterDefaults12Atom has nonzero reserved bits".to_string(),
                ));
            }
            extension.header_footer_defaults = Some(HeaderFooterDefaults {
                include_date: flags & 0x01 != 0,
                include_footer: flags & 0x02 != 0,
                include_header: flags & 0x04 != 0,
                include_slide_number: flags & 0x08 != 0,
            });
        }
        Ok(extension)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn record_bytes(version: u16, instance: u16, record_type: u16, data: &[u8]) -> Vec<u8> {
        let mut bytes = Vec::with_capacity(8 + data.len());
        bytes.extend_from_slice(&((instance << 4) | version).to_le_bytes());
        bytes.extend_from_slice(&record_type.to_le_bytes());
        bytes.extend_from_slice(&(data.len() as u32).to_le_bytes());
        bytes.extend_from_slice(data);
        bytes
    }

    fn prog_tags_record(version: u8, records: &[u8]) -> PptRecord {
        let name: Vec<u8> = format!("___PPT{version}")
            .encode_utf16()
            .flat_map(u16::to_le_bytes)
            .collect();
        let name = record_bytes(0, 0, PptRecordType::CString.as_u16(), &name);
        let blob = record_bytes(0, 0, PptRecordType::BinaryTagData.as_u16(), records);
        let tag_data = [name, blob].concat();
        let tag = record_bytes(0xf, 0, PptRecordType::ProgBinaryTag.as_u16(), &tag_data);
        let tags = record_bytes(0xf, 0, PptRecordType::ProgTags.as_u16(), &tag);
        PptRecord::parse(&tags, 0).unwrap().0
    }

    fn root(children: Vec<PptRecord>) -> PptRecord {
        PptRecord {
            version: 0xf,
            instance: 0,
            record_type: PptRecordType::Document,
            record_type_raw: PptRecordType::Document.as_u16(),
            data_length: 0,
            data: Vec::new(),
            children,
        }
    }

    #[test]
    fn parses_powerpoint12_header_footer_defaults() {
        let defaults = record_bytes(0, 0, 0x0424, &[0b1011]);
        let document = root(vec![prog_tags_record(12, &defaults)]);

        let parsed = SlideExtension::parse(&document).unwrap();
        assert_eq!(
            parsed.header_footer_defaults,
            Some(HeaderFooterDefaults {
                include_date: true,
                include_footer: true,
                include_header: false,
                include_slide_number: true,
            })
        );

        let wrong_version = root(vec![prog_tags_record(11, &defaults)]);
        assert_eq!(
            SlideExtension::parse(&wrong_version).unwrap(),
            SlideExtension::default()
        );
    }

    #[test]
    fn rejects_malformed_or_duplicate_header_footer_defaults() {
        for malformed in [
            record_bytes(1, 0, 0x0424, &[0]),
            record_bytes(0, 1, 0x0424, &[0]),
            record_bytes(0, 0, 0x0424, &[]),
            record_bytes(0, 0, 0x0424, &[0x10]),
        ] {
            let document = root(vec![prog_tags_record(12, &malformed)]);
            assert!(SlideExtension::parse(&document).is_err());
        }

        let defaults = record_bytes(0, 0, 0x0424, &[0]);
        let duplicate = [defaults.clone(), defaults].concat();
        let document = root(vec![prog_tags_record(12, &duplicate)]);
        assert!(SlideExtension::parse(&document).is_err());
    }
}
