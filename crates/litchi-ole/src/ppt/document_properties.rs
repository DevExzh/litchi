//! PowerPoint 10 document metadata and privacy settings.

use super::package::{PptError, Result};
use super::records::PptRecord;
use crate::consts::PptRecordType;

/// Metadata and privacy settings stored in the `___PPT10` document extension.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct PowerPoint10DocumentProperties {
    /// Optional copyright notice.
    pub copyright: Option<String>,
    /// Optional document keywords.
    pub keywords: Option<String>,
    /// Whether personally identifiable information is removed when saving.
    ///
    /// `None` means that the optional `FilterPrivacyFlags10Atom` is absent.
    pub remove_personally_identifiable_information: Option<bool>,
}

impl PowerPoint10DocumentProperties {
    /// Discover and parse document properties from every `___PPT10` programmable tag below `root`.
    pub fn parse(root: &PptRecord) -> Result<Self> {
        let mut properties = Self::default();
        for record in root.versioned_binary_tag_records(10)? {
            match (record.record_type, record.instance) {
                (PptRecordType::CString, 1) => {
                    if properties.copyright.is_some() {
                        return Err(PptError::Corrupted(
                            "PowerPoint 10 document extension contains duplicate copyright"
                                .to_string(),
                        ));
                    }
                    properties.copyright = Some(parse_summary_string(&record, "CopyrightAtom")?);
                },
                (PptRecordType::CString, 2) => {
                    if properties.keywords.is_some() {
                        return Err(PptError::Corrupted(
                            "PowerPoint 10 document extension contains duplicate keywords"
                                .to_string(),
                        ));
                    }
                    properties.keywords = Some(parse_summary_string(&record, "KeywordsAtom")?);
                },
                (PptRecordType::FilterPrivacyFlags10Atom, _) => {
                    if properties
                        .remove_personally_identifiable_information
                        .is_some()
                    {
                        return Err(PptError::Corrupted(
                            "PowerPoint 10 document extension contains duplicate privacy flags"
                                .to_string(),
                        ));
                    }
                    properties.remove_personally_identifiable_information =
                        Some(parse_privacy_flags(&record)?);
                },
                _ => {},
            }
        }
        Ok(properties)
    }
}

fn parse_summary_string(record: &PptRecord, name: &str) -> Result<String> {
    if record.record_type != PptRecordType::CString
        || record.version != 0
        || !matches!(record.instance, 1 | 2)
        || record.data.len() > 510
        || record.data.len() & 1 != 0
    {
        return Err(PptError::Corrupted(format!(
            "{name} has an invalid record header or size"
        )));
    }

    let mut units = Vec::with_capacity(record.data.len() / 2);
    for bytes in record.data.chunks_exact(2) {
        let unit = u16::from_le_bytes([bytes[0], bytes[1]]);
        if unit == 0 {
            break;
        }
        if matches!(unit, 0x0001..=0x001f | 0x007f..=0x009f) {
            return Err(PptError::Corrupted(format!(
                "{name} contains a non-printable character"
            )));
        }
        units.push(unit);
    }
    String::from_utf16(&units)
        .map_err(|_| PptError::Corrupted(format!("{name} contains invalid UTF-16")))
}

fn parse_privacy_flags(record: &PptRecord) -> Result<bool> {
    if record.record_type != PptRecordType::FilterPrivacyFlags10Atom
        || record.version != 0
        || record.instance != 0
        || record.data.len() != 4
    {
        return Err(PptError::Corrupted(
            "FilterPrivacyFlags10Atom has an invalid record header or size".to_string(),
        ));
    }
    let flags = u32::from_le_bytes([
        record.data[0],
        record.data[1],
        record.data[2],
        record.data[3],
    ]);
    if flags & !1 != 0 {
        return Err(PptError::Corrupted(
            "FilterPrivacyFlags10Atom has nonzero reserved bits".to_string(),
        ));
    }
    Ok(flags & 1 != 0)
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

    fn utf16(value: &str) -> Vec<u8> {
        value.encode_utf16().flat_map(u16::to_le_bytes).collect()
    }

    fn prog_tags_record(version: u8, blob_payload: &[u8]) -> PptRecord {
        let tag_name = utf16(&format!("___PPT{version}"));
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

    fn root(children: Vec<PptRecord>) -> PptRecord {
        PptRecord {
            record_type: PptRecordType::Document,
            record_type_raw: 1000,
            version: 0x0f,
            instance: 0,
            data_length: 0,
            data: Vec::new(),
            children,
        }
    }

    #[test]
    fn parses_powerpoint10_document_properties() {
        let mut records = record_bytes(0, 1, 4026, &utf16("Copyright © 2026"));
        records.extend_from_slice(&record_bytes(0, 2, 4026, &utf16("rust; ooxml")));
        records.extend_from_slice(&record_bytes(0, 0, 0x36b0, &1u32.to_le_bytes()));

        let properties =
            PowerPoint10DocumentProperties::parse(&root(vec![prog_tags_record(10, &records)]))
                .unwrap();

        assert_eq!(properties.copyright.as_deref(), Some("Copyright © 2026"));
        assert_eq!(properties.keywords.as_deref(), Some("rust; ooxml"));
        assert_eq!(
            properties.remove_personally_identifiable_information,
            Some(true)
        );
    }

    #[test]
    fn ignores_other_versions_and_preserves_absent_privacy_flags() {
        let copyright = record_bytes(0, 1, 4026, &utf16("old"));
        let keywords = record_bytes(0, 2, 4026, &utf16("current"));
        let properties = PowerPoint10DocumentProperties::parse(&root(vec![
            prog_tags_record(9, &copyright),
            prog_tags_record(10, &keywords),
        ]))
        .unwrap();

        assert_eq!(properties.copyright, None);
        assert_eq!(properties.keywords.as_deref(), Some("current"));
        assert_eq!(properties.remove_personally_identifiable_information, None);
    }

    #[test]
    fn rejects_malformed_document_properties() {
        let malformed = [
            record_bytes(1, 1, 4026, &utf16("copyright")),
            record_bytes(0, 1, 4026, b"A"),
            record_bytes(0, 2, 4026, &vec![0; 512]),
            record_bytes(0, 2, 4026, &[0x01, 0x00]),
            record_bytes(0, 2, 4026, &[0x00, 0xd8]),
            record_bytes(0, 1, 0x36b0, &0u32.to_le_bytes()),
            record_bytes(0, 0, 0x36b0, &[0, 0, 0]),
            record_bytes(0, 0, 0x36b0, &2u32.to_le_bytes()),
        ];
        for record in malformed {
            let document = root(vec![prog_tags_record(10, &record)]);
            assert!(PowerPoint10DocumentProperties::parse(&document).is_err());
        }
    }

    #[test]
    fn rejects_duplicate_document_properties() {
        for record in [
            record_bytes(0, 1, 4026, &utf16("copyright")),
            record_bytes(0, 2, 4026, &utf16("keywords")),
            record_bytes(0, 0, 0x36b0, &0u32.to_le_bytes()),
        ] {
            let mut duplicate = record.clone();
            duplicate.extend_from_slice(&record);
            let document = root(vec![prog_tags_record(10, &duplicate)]);
            assert!(PowerPoint10DocumentProperties::parse(&document).is_err());
        }
    }
}
