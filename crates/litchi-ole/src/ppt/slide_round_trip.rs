//! Inert PowerPoint 12 direct slide round-trip metadata.

use super::package::{PptError, Result};
use super::records::PptRecord;
use crate::consts::PptRecordType;

/// Reference from a slide to its PowerPoint 12 slide layout.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PowerPointContentMasterReference {
    /// Record-instance bits retained because MS-PPT does not constrain them for this atom.
    pub record_instance: u16,
    /// Identifier of the owning main master slide.
    pub main_master_id: u32,
    /// Instance identifier of the slide layout.
    pub layout_instance_id: u16,
    /// Undefined payload value retained for lossless inspection.
    pub unused: u16,
}

/// PowerPoint 12 master references stored directly in a slide container.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct PowerPoint12SlideRoundTripMetadata {
    /// Identifier of the main master merged into this slide layout.
    pub composite_master_id: Option<u32>,
    /// Reference from this slide to its main master and slide layout.
    pub content_master: Option<PowerPointContentMasterReference>,
}

impl PowerPoint12SlideRoundTripMetadata {
    /// Parse direct PowerPoint 12 round-trip records below `root`.
    pub fn parse(root: &PptRecord) -> Result<Self> {
        let mut metadata = Self::default();
        for record in &root.children {
            match record.record_type {
                PptRecordType::RoundTripCompositeMasterId12Atom => {
                    if metadata.composite_master_id.is_some() {
                        return Err(PptError::Corrupted(
                            "Slide contains duplicate RoundTripCompositeMasterId12Atom records"
                                .to_string(),
                        ));
                    }
                    validate_atom(record, "RoundTripCompositeMasterId12Atom", 4, Some(0))?;
                    metadata.composite_master_id = Some(u32::from_le_bytes([
                        record.data[0],
                        record.data[1],
                        record.data[2],
                        record.data[3],
                    ]));
                },
                PptRecordType::RoundTripContentMasterId12Atom => {
                    if metadata.content_master.is_some() {
                        return Err(PptError::Corrupted(
                            "Slide contains duplicate RoundTripContentMasterId12Atom records"
                                .to_string(),
                        ));
                    }
                    validate_atom(record, "RoundTripContentMasterId12Atom", 8, None)?;
                    metadata.content_master = Some(PowerPointContentMasterReference {
                        record_instance: record.instance,
                        main_master_id: u32::from_le_bytes([
                            record.data[0],
                            record.data[1],
                            record.data[2],
                            record.data[3],
                        ]),
                        layout_instance_id: u16::from_le_bytes([record.data[4], record.data[5]]),
                        unused: u16::from_le_bytes([record.data[6], record.data[7]]),
                    });
                },
                _ => {},
            }
        }
        Ok(metadata)
    }
}

fn validate_atom(
    record: &PptRecord,
    name: &str,
    expected_length: usize,
    expected_instance: Option<u16>,
) -> Result<()> {
    if record.version != 0
        || expected_instance.is_some_and(|instance| record.instance != instance)
        || record.data_length as usize != expected_length
        || record.data.len() != expected_length
    {
        return Err(PptError::Corrupted(format!(
            "{name} has an invalid record header or size"
        )));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn record(record_type: PptRecordType, version: u16, instance: u16, data: &[u8]) -> PptRecord {
        PptRecord {
            version,
            instance,
            record_type,
            record_type_raw: record_type.as_u16(),
            data_length: data.len() as u32,
            data: data.to_vec(),
            children: Vec::new(),
        }
    }

    fn root(children: Vec<PptRecord>) -> PptRecord {
        let mut root = record(PptRecordType::Slide, 0x0f, 0, &[]);
        root.children = children;
        root
    }

    #[test]
    fn parses_direct_slide_master_references_and_retains_undefined_values() {
        let composite = record(
            PptRecordType::RoundTripCompositeMasterId12Atom,
            0,
            0,
            &u32::MAX.to_le_bytes(),
        );
        let mut content = Vec::new();
        content.extend_from_slice(&0u32.to_le_bytes());
        content.extend_from_slice(&u16::MAX.to_le_bytes());
        content.extend_from_slice(&0xa55au16.to_le_bytes());
        let content = record(
            PptRecordType::RoundTripContentMasterId12Atom,
            0,
            0x0fff,
            &content,
        );

        let parsed =
            PowerPoint12SlideRoundTripMetadata::parse(&root(vec![composite, content])).unwrap();
        assert_eq!(parsed.composite_master_id, Some(u32::MAX));
        assert_eq!(
            parsed.content_master,
            Some(PowerPointContentMasterReference {
                record_instance: 0x0fff,
                main_master_id: 0,
                layout_instance_id: u16::MAX,
                unused: 0xa55a,
            })
        );
        assert_eq!(
            PowerPoint12SlideRoundTripMetadata::parse(&root(Vec::new())).unwrap(),
            PowerPoint12SlideRoundTripMetadata::default()
        );
    }

    #[test]
    fn rejects_malformed_or_duplicate_direct_slide_master_references() {
        let composite = |version, instance, data: &[u8]| {
            record(
                PptRecordType::RoundTripCompositeMasterId12Atom,
                version,
                instance,
                data,
            )
        };
        let content = |version, instance, data: &[u8]| {
            record(
                PptRecordType::RoundTripContentMasterId12Atom,
                version,
                instance,
                data,
            )
        };
        for malformed in [
            composite(1, 0, &[0; 4]),
            composite(0, 1, &[0; 4]),
            composite(0, 0, &[0; 3]),
            composite(0, 0, &[0; 5]),
            content(1, 0, &[0; 8]),
            content(0, 0, &[0; 7]),
            content(0, 0, &[0; 9]),
        ] {
            assert!(PowerPoint12SlideRoundTripMetadata::parse(&root(vec![malformed])).is_err());
        }

        let mut wrong_declared_length = composite(0, 0, &[0; 4]);
        wrong_declared_length.data_length = 5;
        assert!(
            PowerPoint12SlideRoundTripMetadata::parse(&root(vec![wrong_declared_length])).is_err()
        );

        let duplicate_composite = root(vec![composite(0, 0, &[0; 4]), composite(0, 0, &[1; 4])]);
        assert!(PowerPoint12SlideRoundTripMetadata::parse(&duplicate_composite).is_err());
        let duplicate_content = root(vec![content(0, 0, &[0; 8]), content(0, 1, &[1; 8])]);
        assert!(PowerPoint12SlideRoundTripMetadata::parse(&duplicate_content).is_err());
    }
}
