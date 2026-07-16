//! Inert PowerPoint 12 direct slide round-trip metadata.

use super::package::{PptError, Result};
use super::records::PptRecord;
use crate::consts::PptRecordType;
use litchi_opc::OpcPackage;
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

const PRESENTATIONML_NAMESPACE: &[u8] =
    b"http://schemas.openxmlformats.org/presentationml/2006/main";
const TIMING_INFO_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.timingInfo+xml";
const TIMING_INFO_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/timingInfo";

/// Validated embedded ECMA-376 package containing PowerPoint 12 animation timing data.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PowerPointAnimationPackage {
    /// Original package bytes retained without modification for lossless round trips.
    pub data: Vec<u8>,
    /// Number of parts in the embedded OPC package.
    pub part_count: usize,
    /// Package part name of the PresentationML Timing Info part.
    pub timing_part_name: String,
}

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
    /// Validated embedded PowerPoint 12 animation package.
    pub animation_package: Option<PowerPointAnimationPackage>,
    /// Checksum stored for the animation data.
    pub animation_checksum: Option<u32>,
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
                PptRecordType::RoundTripAnimation12Atom => {
                    if metadata.animation_package.is_some() {
                        return Err(PptError::Corrupted(
                            "Slide contains duplicate RoundTripAnimationAtom records".to_string(),
                        ));
                    }
                    validate_variable_atom(record, "RoundTripAnimationAtom")?;
                    metadata.animation_package = Some(parse_animation_package(&record.data)?);
                },
                PptRecordType::RoundTripAnimationHash12Atom => {
                    if metadata.animation_checksum.is_some() {
                        return Err(PptError::Corrupted(
                            "Slide contains duplicate RoundTripAnimationHashAtom records"
                                .to_string(),
                        ));
                    }
                    validate_atom(record, "RoundTripAnimationHashAtom", 4, Some(0))?;
                    metadata.animation_checksum = Some(u32::from_le_bytes([
                        record.data[0],
                        record.data[1],
                        record.data[2],
                        record.data[3],
                    ]));
                },
                _ => {},
            }
        }
        Ok(metadata)
    }
}

fn validate_variable_atom(record: &PptRecord, name: &str) -> Result<()> {
    if record.version != 0
        || record.instance != 0
        || record.data_length as usize != record.data.len()
    {
        return Err(PptError::Corrupted(format!(
            "{name} has an invalid record header or size"
        )));
    }
    Ok(())
}

fn parse_animation_package(data: &[u8]) -> Result<PowerPointAnimationPackage> {
    let package = OpcPackage::from_bytes(data).map_err(|error| {
        PptError::Corrupted(format!(
            "RoundTripAnimationAtom contains an invalid ECMA-376 package: {error}"
        ))
    })?;
    let mut timing_relationships = package
        .rels()
        .iter()
        .filter(|relationship| relationship.reltype() == TIMING_INFO_RELATIONSHIP_TYPE);
    let timing_relationship = timing_relationships.next().ok_or_else(|| {
        PptError::Corrupted(
            "RoundTripAnimationAtom package has no Timing Info relationship".to_string(),
        )
    })?;
    if timing_relationships.next().is_some() || timing_relationship.is_external() {
        return Err(PptError::Corrupted(
            "RoundTripAnimationAtom package has invalid Timing Info relationships".to_string(),
        ));
    }
    let timing_part_name = timing_relationship.target_partname().map_err(|error| {
        PptError::Corrupted(format!(
            "RoundTripAnimationAtom has an invalid Timing Info target: {error}"
        ))
    })?;
    let timing_part = package.get_part(&timing_part_name).map_err(|error| {
        PptError::Corrupted(format!(
            "RoundTripAnimationAtom Timing Info part is invalid: {error}"
        ))
    })?;
    if timing_part.content_type() != TIMING_INFO_CONTENT_TYPE {
        return Err(PptError::Corrupted(
            "RoundTripAnimationAtom Timing Info part has an invalid content type".to_string(),
        ));
    }
    if !xml_contains_presentation_timing(timing_part.blob()).map_err(|error| {
        PptError::Corrupted(format!(
            "RoundTripAnimationAtom Timing Info XML is invalid: {error}"
        ))
    })? {
        return Err(PptError::Corrupted(
            "RoundTripAnimationAtom Timing Info part has no PresentationML timing element"
                .to_string(),
        ));
    }
    for part in package.iter_parts() {
        if part.partname() == &timing_part_name || !is_xml_content_type(part.content_type()) {
            continue;
        }
        xml_contains_presentation_timing(part.blob()).map_err(|error| {
            PptError::Corrupted(format!(
                "RoundTripAnimationAtom XML part {} is invalid: {error}",
                part.partname()
            ))
        })?;
    }
    Ok(PowerPointAnimationPackage {
        data: data.to_vec(),
        part_count: package.part_count(),
        timing_part_name: timing_part_name.to_string(),
    })
}

fn is_xml_content_type(content_type: &str) -> bool {
    content_type == "application/xml"
        || content_type == "text/xml"
        || content_type.ends_with("+xml")
}

fn xml_contains_presentation_timing(data: &[u8]) -> std::result::Result<bool, String> {
    let mut reader = NsReader::from_reader(data);
    let mut contains_timing = false;
    let mut depth = 0usize;
    let mut root_count = 0usize;
    loop {
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| error.to_string())?;
        match event {
            Event::Start(element) => {
                if depth == 0 {
                    root_count += 1;
                    if root_count > 1 {
                        return Err("XML document has multiple root elements".to_string());
                    }
                }
                if element.local_name().as_ref() == b"timing"
                    && matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == PRESENTATIONML_NAMESPACE)
                {
                    contains_timing = true;
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| "XML nesting is too deep".to_string())?;
            },
            Event::Empty(element) => {
                if depth == 0 {
                    root_count += 1;
                    if root_count > 1 {
                        return Err("XML document has multiple root elements".to_string());
                    }
                }
                if element.local_name().as_ref() == b"timing"
                    && matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == PRESENTATIONML_NAMESPACE)
                {
                    contains_timing = true;
                }
            },
            Event::End(_) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| "XML document has an unmatched closing element".to_string())?;
            },
            Event::Text(text)
                if depth == 0
                    && text
                        .as_ref()
                        .iter()
                        .any(|byte| !matches!(byte, b' ' | b'\t' | b'\n' | b'\r')) =>
            {
                return Err("XML document has text outside its root element".to_string());
            },
            Event::CData(_) | Event::GeneralRef(_) if depth == 0 => {
                return Err("XML document has character data outside its root element".to_string());
            },
            Event::Eof if depth != 0 => {
                return Err("XML document has an unclosed element".to_string());
            },
            Event::Eof if root_count != 1 => {
                return Err("XML document does not have exactly one root element".to_string());
            },
            Event::Eof => return Ok(contains_timing),
            _ => {},
        }
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
    use litchi_opc::{PackURI, XmlPart};
    use std::io::Cursor;

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

    fn animation_package(parts: &[(&str, &str, &str, &[u8])]) -> Vec<u8> {
        let mut package = OpcPackage::new();
        for (index, (name, content_type, relationship_type, data)) in parts.iter().enumerate() {
            package.add_part(Box::new(XmlPart::new(
                PackURI::new(*name).unwrap(),
                (*content_type).to_string(),
                data.to_vec(),
            )));
            package.rels_mut().add_relationship(
                (*relationship_type).to_string(),
                (*name).to_string(),
                format!("rId{}", index + 1),
                false,
            );
        }
        let mut output = Cursor::new(Vec::new());
        package.to_stream(&mut output).unwrap();
        output.into_inner()
    }

    fn valid_animation_package() -> Vec<u8> {
        animation_package(&[
            (
                "/drs/timingInfo.xml",
                TIMING_INFO_CONTENT_TYPE,
                TIMING_INFO_RELATIONSHIP_TYPE,
                br#"<p:timing xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>"#,
            ),
            (
                "/drs/metadata.xml",
                "application/xml",
                "urn:litchi:test:metadata",
                br#"<root xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:timing/></root>"#,
            ),
        ])
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

    #[test]
    fn parses_animation_package_and_checksum() {
        assert!(
            xml_contains_presentation_timing(
                br#"<p:timing xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>"#
            )
            .unwrap()
        );
        let package_data = valid_animation_package();
        let package_record = record(PptRecordType::RoundTripAnimation12Atom, 0, 0, &package_data);
        let checksum_record = record(
            PptRecordType::RoundTripAnimationHash12Atom,
            0,
            0,
            &u32::MAX.to_le_bytes(),
        );

        let parsed =
            PowerPoint12SlideRoundTripMetadata::parse(&root(vec![checksum_record, package_record]))
                .unwrap();
        assert_eq!(parsed.animation_checksum, Some(u32::MAX));
        let package = parsed.animation_package.unwrap();
        assert_eq!(package.data, package_data);
        assert_eq!(package.part_count, 2);
        assert_eq!(package.timing_part_name, "/drs/timingInfo.xml");

        let zero_checksum = record(
            PptRecordType::RoundTripAnimationHash12Atom,
            0,
            0,
            &0u32.to_le_bytes(),
        );
        assert_eq!(
            PowerPoint12SlideRoundTripMetadata::parse(&root(vec![zero_checksum]))
                .unwrap()
                .animation_checksum,
            Some(0)
        );
    }

    #[test]
    fn rejects_malformed_animation_records() {
        let package_data = valid_animation_package();
        let animation = |version, instance, data: &[u8]| {
            record(
                PptRecordType::RoundTripAnimation12Atom,
                version,
                instance,
                data,
            )
        };
        let checksum = |version, instance, data: &[u8]| {
            record(
                PptRecordType::RoundTripAnimationHash12Atom,
                version,
                instance,
                data,
            )
        };
        for malformed in [
            animation(1, 0, &package_data),
            animation(0, 1, &package_data),
            animation(0, 0, b"not a package"),
            checksum(1, 0, &[0; 4]),
            checksum(0, 1, &[0; 4]),
            checksum(0, 0, &[0; 3]),
            checksum(0, 0, &[0; 5]),
        ] {
            assert!(PowerPoint12SlideRoundTripMetadata::parse(&root(vec![malformed])).is_err());
        }

        let mut wrong_package_length = animation(0, 0, &package_data);
        wrong_package_length.data_length -= 1;
        assert!(
            PowerPoint12SlideRoundTripMetadata::parse(&root(vec![wrong_package_length])).is_err()
        );
        let mut wrong_checksum_length = checksum(0, 0, &[0; 4]);
        wrong_checksum_length.data_length = 5;
        assert!(
            PowerPoint12SlideRoundTripMetadata::parse(&root(vec![wrong_checksum_length])).is_err()
        );

        let duplicate_animation = root(vec![
            animation(0, 0, &package_data),
            animation(0, 0, &package_data),
        ]);
        assert!(PowerPoint12SlideRoundTripMetadata::parse(&duplicate_animation).is_err());
        let duplicate_checksum = root(vec![checksum(0, 0, &[0; 4]), checksum(0, 0, &[1; 4])]);
        assert!(PowerPoint12SlideRoundTripMetadata::parse(&duplicate_checksum).is_err());
    }

    #[test]
    fn rejects_animation_packages_without_valid_presentation_timing_xml() {
        for package_data in [
            animation_package(&[(
                "/drs/timingInfo.xml",
                TIMING_INFO_CONTENT_TYPE,
                TIMING_INFO_RELATIONSHIP_TYPE,
                br#"<timing xmlns="urn:not-presentationml"/>"#,
            )]),
            animation_package(&[(
                "/drs/timingInfo.xml",
                TIMING_INFO_CONTENT_TYPE,
                TIMING_INFO_RELATIONSHIP_TYPE,
                br#"<p:timing xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">"#,
            )]),
            animation_package(&[(
                "/drs/timingInfo.xml",
                "application/octet-stream",
                TIMING_INFO_RELATIONSHIP_TYPE,
                br#"<p:timing xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>"#,
            )]),
            animation_package(&[(
                "/drs/timingInfo.xml",
                TIMING_INFO_CONTENT_TYPE,
                "urn:not-timing-info",
                br#"<p:timing xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>"#,
            )]),
            animation_package(&[
                (
                    "/drs/timingInfo.xml",
                    TIMING_INFO_CONTENT_TYPE,
                    TIMING_INFO_RELATIONSHIP_TYPE,
                    br#"<p:timing xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>"#,
                ),
                (
                    "/drs/broken.xml",
                    "application/xml",
                    "urn:litchi:test:metadata",
                    b"<broken>",
                ),
            ]),
            animation_package(&[
                (
                    "/drs/timingInfo1.xml",
                    TIMING_INFO_CONTENT_TYPE,
                    TIMING_INFO_RELATIONSHIP_TYPE,
                    br#"<p:timing xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>"#,
                ),
                (
                    "/drs/timingInfo2.xml",
                    TIMING_INFO_CONTENT_TYPE,
                    TIMING_INFO_RELATIONSHIP_TYPE,
                    br#"<p:timing xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>"#,
                ),
            ]),
        ] {
            let animation = record(
                PptRecordType::RoundTripAnimation12Atom,
                0,
                0,
                &package_data,
            );
            assert!(PowerPoint12SlideRoundTripMetadata::parse(&root(vec![animation])).is_err());
        }
    }
}
