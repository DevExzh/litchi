//! PowerPoint 12 round-trip metadata attached to main master slides.

use super::package::{PptError, Result};
use super::records::PptRecord;
use super::slide_round_trip::{
    PowerPointAnimationPackage, PowerPointColorMapping, PowerPointEmbeddedXmlPackage,
    PowerPointThemePackage, parse_animation_package, parse_color_mapping,
    parse_embedded_xml_package, parse_theme_package, validate_variable_atom,
};
use crate::consts::PptRecordType;
use litchi_opc::constants::content_type;

const PRESENTATIONML_NAMESPACE: &[u8] =
    b"http://schemas.openxmlformats.org/presentationml/2006/main";

/// Position of a PowerPoint 12 main-master text-style package in the container schema.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PowerPointMainMasterTextStylesSource {
    /// Optional direct field before the main master's drawing.
    Direct,
    /// Member of the PowerPoint 12 round-trip array after the drawing.
    RoundTripArray,
}

/// Embedded slide layout associated with a PowerPoint 12 main master.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PowerPointContentMasterInfo {
    /// Record-instance bits retained because MS-PPT does not constrain them.
    pub record_instance: u16,
    /// Validated package containing the PresentationML `sldLayout` part.
    pub package: PowerPointEmbeddedXmlPackage,
}

/// Embedded text styles associated with a PowerPoint 12 main master.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PowerPointMainMasterTextStyles {
    /// Schema position from which this package was read.
    pub source: PowerPointMainMasterTextStylesSource,
    /// Validated package containing the PresentationML `txStyles` part.
    pub package: PowerPointEmbeddedXmlPackage,
}

/// PowerPoint 12 round-trip metadata stored on one main master slide.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct PowerPoint12MainMasterMetadata {
    /// Original PresentationML slide-master identifier.
    pub original_main_master_id: Option<u32>,
    /// Validated embedded theme or theme-override package.
    pub theme_package: Option<PowerPointThemePackage>,
    /// Validated color-mapping XML.
    pub color_mapping: Option<PowerPointColorMapping>,
    /// Repeatable embedded slide-layout packages.
    pub content_masters: Vec<PowerPointContentMasterInfo>,
    /// Direct and/or round-trip-array text-style packages.
    pub text_styles: Vec<PowerPointMainMasterTextStyles>,
    /// Validated embedded animation package.
    pub animation_package: Option<PowerPointAnimationPackage>,
    /// Checksum stored for the animation data.
    pub animation_checksum: Option<u32>,
    /// Identifier of the main master merged into a custom layout.
    pub composite_master_id: Option<u32>,
}

impl PowerPoint12MainMasterMetadata {
    /// Parse PowerPoint 12 round-trip records from one `MainMaster` container.
    pub fn parse(root: &PptRecord) -> Result<Self> {
        if root.record_type != PptRecordType::MainMaster {
            return Err(PptError::Corrupted(
                "PowerPoint 12 main-master metadata requires a MainMaster container".to_string(),
            ));
        }
        let mut metadata = Self::default();
        let mut seen_drawing = false;
        let mut direct_text_styles = false;
        let mut array_text_styles = false;
        for record in &root.children {
            if record.record_type == PptRecordType::PPDrawing {
                seen_drawing = true;
                continue;
            }
            match record.record_type {
                PptRecordType::RoundTripOArtTextStyles12Atom => {
                    let (source, seen) = if seen_drawing {
                        (
                            PowerPointMainMasterTextStylesSource::RoundTripArray,
                            &mut array_text_styles,
                        )
                    } else {
                        (
                            PowerPointMainMasterTextStylesSource::Direct,
                            &mut direct_text_styles,
                        )
                    };
                    if *seen {
                        return Err(PptError::Corrupted(format!(
                            "MainMaster contains duplicate {source:?} text-style records"
                        )));
                    }
                    *seen = true;
                    validate_variable_atom(record, "RoundTripOArtTextStyles12Atom")?;
                    metadata.text_styles.push(PowerPointMainMasterTextStyles {
                        source,
                        package: parse_embedded_xml_package(
                            &record.data,
                            "RoundTripOArtTextStyles12Atom",
                            content_type::PML_SLIDE_MASTER,
                            PRESENTATIONML_NAMESPACE,
                            b"txStyles",
                        )?,
                    });
                },
                PptRecordType::RoundTripOriginalMainMasterId12Atom => {
                    require_array_position(seen_drawing, "RoundTripOriginalMainMasterId12Atom")?;
                    if metadata.original_main_master_id.is_some() {
                        return Err(duplicate("RoundTripOriginalMainMasterId12Atom"));
                    }
                    validate_fixed_atom(record, "RoundTripOriginalMainMasterId12Atom", 4, true)?;
                    metadata.original_main_master_id = Some(read_u32(record));
                },
                PptRecordType::RoundTripTheme12Atom => {
                    require_array_position(seen_drawing, "RoundTripThemeAtom")?;
                    if metadata.theme_package.is_some() {
                        return Err(duplicate("RoundTripThemeAtom"));
                    }
                    validate_variable_atom(record, "RoundTripThemeAtom")?;
                    metadata.theme_package = Some(parse_theme_package(&record.data)?);
                },
                PptRecordType::RoundTripColorMapping12Atom => {
                    require_array_position(seen_drawing, "RoundTripColorMappingAtom")?;
                    if metadata.color_mapping.is_some() {
                        return Err(duplicate("RoundTripColorMappingAtom"));
                    }
                    validate_variable_atom(record, "RoundTripColorMappingAtom")?;
                    metadata.color_mapping = Some(parse_color_mapping(&record.data)?);
                },
                PptRecordType::RoundTripContentMasterInfo12Atom => {
                    require_array_position(seen_drawing, "RoundTripContentMasterInfo12Atom")?;
                    validate_variable_atom_allow_instance(
                        record,
                        "RoundTripContentMasterInfo12Atom",
                    )?;
                    metadata.content_masters.push(PowerPointContentMasterInfo {
                        record_instance: record.instance,
                        package: parse_embedded_xml_package(
                            &record.data,
                            "RoundTripContentMasterInfo12Atom",
                            content_type::PML_SLIDE_LAYOUT,
                            PRESENTATIONML_NAMESPACE,
                            b"sldLayout",
                        )?,
                    });
                },
                PptRecordType::RoundTripAnimation12Atom => {
                    require_array_position(seen_drawing, "RoundTripAnimationAtom")?;
                    if metadata.animation_package.is_some() {
                        return Err(duplicate("RoundTripAnimationAtom"));
                    }
                    validate_variable_atom(record, "RoundTripAnimationAtom")?;
                    metadata.animation_package = Some(parse_animation_package(&record.data)?);
                },
                PptRecordType::RoundTripAnimationHash12Atom => {
                    require_array_position(seen_drawing, "RoundTripAnimationHashAtom")?;
                    if metadata.animation_checksum.is_some() {
                        return Err(duplicate("RoundTripAnimationHashAtom"));
                    }
                    validate_fixed_atom(record, "RoundTripAnimationHashAtom", 4, true)?;
                    metadata.animation_checksum = Some(read_u32(record));
                },
                PptRecordType::RoundTripCompositeMasterId12Atom => {
                    require_array_position(seen_drawing, "RoundTripCompositeMasterId12Atom")?;
                    if metadata.composite_master_id.is_some() {
                        return Err(duplicate("RoundTripCompositeMasterId12Atom"));
                    }
                    validate_fixed_atom(record, "RoundTripCompositeMasterId12Atom", 4, true)?;
                    metadata.composite_master_id = Some(read_u32(record));
                },
                _ => {},
            }
        }
        Ok(metadata)
    }
}

fn require_array_position(seen_drawing: bool, name: &str) -> Result<()> {
    if !seen_drawing {
        return Err(PptError::Corrupted(format!(
            "{name} occurs before the MainMaster drawing"
        )));
    }
    Ok(())
}

fn duplicate(name: &str) -> PptError {
    PptError::Corrupted(format!("MainMaster contains duplicate {name} records"))
}

fn validate_variable_atom_allow_instance(record: &PptRecord, name: &str) -> Result<()> {
    if record.version != 0 || record.data_length as usize != record.data.len() {
        return Err(PptError::Corrupted(format!(
            "{name} has an invalid record header or size"
        )));
    }
    Ok(())
}

fn validate_fixed_atom(
    record: &PptRecord,
    name: &str,
    expected_length: usize,
    require_zero_instance: bool,
) -> Result<()> {
    if record.version != 0
        || (require_zero_instance && record.instance != 0)
        || record.data_length as usize != expected_length
        || record.data.len() != expected_length
    {
        return Err(PptError::Corrupted(format!(
            "{name} has an invalid record header or size"
        )));
    }
    Ok(())
}

fn read_u32(record: &PptRecord) -> u32 {
    u32::from_le_bytes([
        record.data[0],
        record.data[1],
        record.data[2],
        record.data[3],
    ])
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ppt::Package;
    use litchi_opc::{OpcPackage, PackURI, XmlPart};
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
        let mut root = record(PptRecordType::MainMaster, 0x0f, 0, &[]);
        root.children = children;
        root
    }

    fn drawing() -> PptRecord {
        record(PptRecordType::PPDrawing, 0x0f, 0, &[])
    }

    fn xml_package(part_name: &str, content_type: &str, xml: &[u8]) -> Vec<u8> {
        let mut package = OpcPackage::new();
        package.add_part(Box::new(XmlPart::new(
            PackURI::new(part_name).unwrap(),
            content_type.to_string(),
            xml.to_vec(),
        )));
        package.rels_mut().add_relationship(
            "urn:litchi:test:part".to_string(),
            part_name.to_string(),
            "rId1".to_string(),
            false,
        );
        let mut output = Cursor::new(Vec::new());
        package.to_stream(&mut output).unwrap();
        output.into_inner()
    }

    fn text_styles_package() -> Vec<u8> {
        xml_package(
            "/drs/slideMasters/slideMaster1.xml",
            content_type::PML_SLIDE_MASTER,
            br#"<p:txStyles xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>"#,
        )
    }

    fn content_master_package() -> Vec<u8> {
        xml_package(
            "/drs/slideLayouts/slideLayout1.xml",
            content_type::PML_SLIDE_LAYOUT,
            br#"<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>"#,
        )
    }

    #[test]
    fn parses_master_only_metadata_with_schema_position_and_boundaries() {
        let text_styles = text_styles_package();
        let content_master = content_master_package();
        let parsed = PowerPoint12MainMasterMetadata::parse(&root(vec![
            record(
                PptRecordType::RoundTripOArtTextStyles12Atom,
                0,
                0,
                &text_styles,
            ),
            drawing(),
            record(
                PptRecordType::RoundTripOriginalMainMasterId12Atom,
                0,
                0,
                &0u32.to_le_bytes(),
            ),
            record(
                PptRecordType::RoundTripContentMasterInfo12Atom,
                0,
                0,
                &content_master,
            ),
            record(
                PptRecordType::RoundTripContentMasterInfo12Atom,
                0,
                0x0fff,
                &content_master,
            ),
            record(
                PptRecordType::RoundTripOArtTextStyles12Atom,
                0,
                0,
                &text_styles,
            ),
            record(
                PptRecordType::RoundTripCompositeMasterId12Atom,
                0,
                0,
                &u32::MAX.to_le_bytes(),
            ),
        ]))
        .unwrap();

        assert_eq!(parsed.original_main_master_id, Some(0));
        assert_eq!(parsed.composite_master_id, Some(u32::MAX));
        assert_eq!(parsed.content_masters.len(), 2);
        assert_eq!(parsed.content_masters[0].record_instance, 0);
        assert_eq!(parsed.content_masters[1].record_instance, 0x0fff);
        assert_eq!(
            parsed.content_masters[0].package.xml_part_name,
            "/drs/slideLayouts/slideLayout1.xml"
        );
        assert_eq!(parsed.text_styles.len(), 2);
        assert_eq!(
            parsed.text_styles[0].source,
            PowerPointMainMasterTextStylesSource::Direct
        );
        assert_eq!(
            parsed.text_styles[1].source,
            PowerPointMainMasterTextStylesSource::RoundTripArray
        );
        assert_eq!(parsed.text_styles[0].package.data, text_styles);

        assert_eq!(
            PowerPoint12MainMasterMetadata::parse(&root(vec![drawing()])).unwrap(),
            PowerPoint12MainMasterMetadata::default()
        );
    }

    #[test]
    fn rejects_wrong_container_positions_duplicates_and_invalid_headers() {
        let text_styles = text_styles_package();
        let content_master = content_master_package();
        let original = |version, instance, data: &[u8]| {
            record(
                PptRecordType::RoundTripOriginalMainMasterId12Atom,
                version,
                instance,
                data,
            )
        };
        assert!(
            PowerPoint12MainMasterMetadata::parse(&record(PptRecordType::Slide, 0x0f, 0, &[]))
                .is_err()
        );
        for malformed in [
            root(vec![original(0, 0, &u32::MAX.to_le_bytes()), drawing()]),
            root(vec![drawing(), original(1, 0, &u32::MAX.to_le_bytes())]),
            root(vec![drawing(), original(0, 1, &u32::MAX.to_le_bytes())]),
            root(vec![drawing(), original(0, 0, &[0; 3])]),
            root(vec![
                drawing(),
                original(0, 0, &u32::MAX.to_le_bytes()),
                original(0, 0, &u32::MAX.to_le_bytes()),
            ]),
            root(vec![
                record(
                    PptRecordType::RoundTripOArtTextStyles12Atom,
                    0,
                    0,
                    &text_styles,
                ),
                record(
                    PptRecordType::RoundTripOArtTextStyles12Atom,
                    0,
                    0,
                    &text_styles,
                ),
                drawing(),
            ]),
            root(vec![
                drawing(),
                record(
                    PptRecordType::RoundTripOArtTextStyles12Atom,
                    0,
                    0,
                    &text_styles,
                ),
                record(
                    PptRecordType::RoundTripOArtTextStyles12Atom,
                    0,
                    0,
                    &text_styles,
                ),
            ]),
            root(vec![
                drawing(),
                record(
                    PptRecordType::RoundTripContentMasterInfo12Atom,
                    1,
                    0,
                    &content_master,
                ),
            ]),
        ] {
            assert!(PowerPoint12MainMasterMetadata::parse(&malformed).is_err());
        }

        let mut wrong_declared = original(0, 0, &u32::MAX.to_le_bytes());
        wrong_declared.data_length = 5;
        assert!(
            PowerPoint12MainMasterMetadata::parse(&root(vec![drawing(), wrong_declared])).is_err()
        );
    }

    #[test]
    fn rejects_invalid_or_ambiguous_master_embedded_packages() {
        let text_styles = text_styles_package();
        let wrong_text_root = xml_package(
            "/drs/slideMasters/slideMaster1.xml",
            content_type::PML_SLIDE_MASTER,
            br#"<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>"#,
        );
        let wrong_layout_type = xml_package(
            "/drs/slideLayouts/slideLayout1.xml",
            "application/xml",
            br#"<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>"#,
        );
        for (record_type, instance, package) in [
            (
                PptRecordType::RoundTripOArtTextStyles12Atom,
                0,
                b"not a package".as_slice(),
            ),
            (
                PptRecordType::RoundTripOArtTextStyles12Atom,
                1,
                text_styles.as_slice(),
            ),
            (
                PptRecordType::RoundTripOArtTextStyles12Atom,
                0,
                wrong_text_root.as_slice(),
            ),
            (
                PptRecordType::RoundTripContentMasterInfo12Atom,
                0,
                wrong_layout_type.as_slice(),
            ),
        ] {
            let candidate = record(record_type, 0, instance, package);
            assert!(
                PowerPoint12MainMasterMetadata::parse(&root(vec![drawing(), candidate])).is_err()
            );
        }

        let mut package = OpcPackage::new();
        for index in 1..=2 {
            let name = format!("/drs/slideLayouts/slideLayout{index}.xml");
            package.add_part(Box::new(XmlPart::new(
                PackURI::new(&name).unwrap(),
                content_type::PML_SLIDE_LAYOUT.to_string(),
                br#"<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>"#
                    .to_vec(),
            )));
            package.rels_mut().add_relationship(
                "urn:litchi:test:layout".to_string(),
                name,
                format!("rId{index}"),
                false,
            );
        }
        let mut output = Cursor::new(Vec::new());
        package.to_stream(&mut output).unwrap();
        let ambiguous = record(
            PptRecordType::RoundTripContentMasterInfo12Atom,
            0,
            0,
            &output.into_inner(),
        );
        assert!(PowerPoint12MainMasterMetadata::parse(&root(vec![drawing(), ambiguous])).is_err());
    }

    #[test]
    fn presentation_exposes_real_main_master_round_trip_metadata() {
        let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/ole/ppt/SampleShow.ppt");
        let mut package = Package::open(path).unwrap();
        let presentation = package.presentation().unwrap();
        let masters = presentation.powerpoint12_main_master_metadata().unwrap();

        assert!(!masters.is_empty());
        assert!(masters.iter().any(|master| master.theme_package.is_some()));
        assert!(masters.iter().any(|master| !master.text_styles.is_empty()));
        assert!(
            masters
                .iter()
                .any(|master| !master.content_masters.is_empty())
        );
    }
}
