//! Inert `PowerPoint` 12 direct slide round-trip metadata.

use super::package::{Error, Result};
use super::records::Record;
use crate::consts::RecordType;
use litchi_opc::OpcPackage;
use litchi_opc::constants::content_type;
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

const PRESENTATIONML_NAMESPACE: &[u8] =
    b"http://schemas.openxmlformats.org/presentationml/2006/main";
const DRAWINGML_NAMESPACE: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/main";
const TIMING_INFO_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.timingInfo+xml";
const TIMING_INFO_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/timingInfo";

/// Validated embedded ECMA-376 package containing `PowerPoint` 12 animation timing data.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AnimationPackage {
    /// Original package bytes retained without modification for lossless round trips.
    pub data: Vec<u8>,
    /// Number of parts in the embedded OPC package.
    pub part_count: usize,
    /// Package part name of the `PresentationML` Timing Info part.
    pub timing_part_name: String,
}

/// Kind of `DrawingML` theme stored in a `PowerPoint` 12 round-trip package.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ThemeKind {
    /// Full `DrawingML` Theme part with a `theme` root element.
    Theme,
    /// `DrawingML` Theme Override part with a `themeOverride` root element.
    ThemeOverride,
}

/// Validated embedded ECMA-376 package containing a theme or theme override.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ThemePackage {
    /// Original package bytes retained without modification for lossless round trips.
    pub data: Vec<u8>,
    /// Number of parts in the embedded OPC package.
    pub part_count: usize,
    /// Package part name of the Theme or Theme Override part.
    pub theme_part_name: String,
    /// Kind of theme part stored in the package.
    pub kind: ThemeKind,
}

/// Validated embedded ECMA-376 package containing one expected XML part.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct EmbeddedXmlPackage {
    /// Original package bytes retained without modification for lossless round trips.
    pub data: Vec<u8>,
    /// Number of parts in the embedded OPC package.
    pub part_count: usize,
    /// Package part name of the format-specific XML part.
    pub xml_part_name: String,
}

/// XML form stored by a `PowerPoint` 12 color-mapping atom.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ColorMappingKind {
    /// A `DrawingML` `clrMap` element containing a complete color mapping.
    Direct,
    /// A `PresentationML` override that selects the mapping inherited from the master.
    MasterOverride,
    /// A `PresentationML` override containing an explicit complete color mapping.
    ExplicitOverride,
}

/// A `DrawingML` color-scheme slot referenced by a color mapping.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ColorSchemeIndex {
    /// Dark color 1.
    Dark1,
    /// Light color 1.
    Light1,
    /// Dark color 2.
    Dark2,
    /// Light color 2.
    Light2,
    /// Accent color 1.
    Accent1,
    /// Accent color 2.
    Accent2,
    /// Accent color 3.
    Accent3,
    /// Accent color 4.
    Accent4,
    /// Accent color 5.
    Accent5,
    /// Accent color 6.
    Accent6,
    /// Hyperlink color.
    Hyperlink,
    /// Followed-hyperlink color.
    FollowedHyperlink,
}

/// Complete `DrawingML` mapping from presentation roles to color-scheme slots.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ColorMappingValues {
    /// First background role (`bg1`).
    pub background1: ColorSchemeIndex,
    /// First text role (`tx1`).
    pub text1: ColorSchemeIndex,
    /// Second background role (`bg2`).
    pub background2: ColorSchemeIndex,
    /// Second text role (`tx2`).
    pub text2: ColorSchemeIndex,
    /// First accent role.
    pub accent1: ColorSchemeIndex,
    /// Second accent role.
    pub accent2: ColorSchemeIndex,
    /// Third accent role.
    pub accent3: ColorSchemeIndex,
    /// Fourth accent role.
    pub accent4: ColorSchemeIndex,
    /// Fifth accent role.
    pub accent5: ColorSchemeIndex,
    /// Sixth accent role.
    pub accent6: ColorSchemeIndex,
    /// Hyperlink role.
    pub hyperlink: ColorSchemeIndex,
    /// Followed-hyperlink role.
    pub followed_hyperlink: ColorSchemeIndex,
}

/// Validated `PowerPoint` 12 color-mapping XML.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ColorMapping {
    /// Original UTF-8 XML retained without modification for lossless round trips.
    pub xml: String,
    /// Top-level color-mapping form.
    pub kind: ColorMappingKind,
    /// Complete mapping values; absent only when the override inherits its master.
    pub values: Option<ColorMappingValues>,
}

/// Reference from a slide to its `PowerPoint` 12 slide layout.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ContentMasterReference {
    /// Record-instance bits retained because MS-PPT does not constrain them for this atom.
    pub record_instance: u16,
    /// Identifier of the owning main master slide.
    pub main_master_id: u32,
    /// Instance identifier of the slide layout.
    pub layout_instance_id: u16,
    /// Undefined payload value retained for lossless inspection.
    pub unused: u16,
}

/// `PowerPoint` 12 round-trip metadata stored directly in a slide container.
#[allow(
    clippy::module_name_repetitions,
    reason = "`SlideRoundTripMetadata12` is the established public API name for slide-level `PowerPoint` 12 round-trip metadata; renaming it would break downstream crates"
)]
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct SlideRoundTripMetadata12 {
    /// Validated embedded `PowerPoint` 12 theme or theme-override package.
    pub theme_package: Option<ThemePackage>,
    /// Validated `PowerPoint` 12 color-mapping XML.
    pub color_mapping: Option<ColorMapping>,
    /// Identifier of the main master merged into this slide layout.
    pub composite_master_id: Option<u32>,
    /// Reference from this slide to its main master and slide layout.
    pub content_master: Option<ContentMasterReference>,
    /// Validated embedded `PowerPoint` 12 animation package.
    pub animation_package: Option<AnimationPackage>,
    /// Checksum stored for the animation data.
    pub animation_checksum: Option<u32>,
}

impl SlideRoundTripMetadata12 {
    /// Parse direct `PowerPoint` 12 round-trip records below `root`.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn parse(root: &Record) -> Result<Self> {
        let mut metadata = Self::default();
        for record in &root.children {
            #[allow(
                clippy::wildcard_enum_match_arm,
                reason = "`RecordType` mirrors the full MS-PPT record-type enumeration; records other than the six round-trip atoms are intentionally ignored"
            )]
            match record.record_type {
                RecordType::RoundTripTheme12Atom => {
                    if metadata.theme_package.is_some() {
                        return Err(Error::Corrupted(
                            "Slide contains duplicate RoundTripThemeAtom records".to_string(),
                        ));
                    }
                    validate_variable_atom(record, "RoundTripThemeAtom")?;
                    metadata.theme_package = Some(parse_theme_package(&record.data)?);
                },
                RecordType::RoundTripColorMapping12Atom => {
                    if metadata.color_mapping.is_some() {
                        return Err(Error::Corrupted(
                            "Slide contains duplicate RoundTripColorMappingAtom records"
                                .to_string(),
                        ));
                    }
                    validate_variable_atom(record, "RoundTripColorMappingAtom")?;
                    metadata.color_mapping = Some(parse_color_mapping(&record.data)?);
                },
                RecordType::RoundTripCompositeMasterId12Atom => {
                    if metadata.composite_master_id.is_some() {
                        return Err(Error::Corrupted(
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
                RecordType::RoundTripContentMasterId12Atom => {
                    if metadata.content_master.is_some() {
                        return Err(Error::Corrupted(
                            "Slide contains duplicate RoundTripContentMasterId12Atom records"
                                .to_string(),
                        ));
                    }
                    validate_atom(record, "RoundTripContentMasterId12Atom", 8, None)?;
                    metadata.content_master = Some(ContentMasterReference {
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
                RecordType::RoundTripAnimation12Atom => {
                    if metadata.animation_package.is_some() {
                        return Err(Error::Corrupted(
                            "Slide contains duplicate RoundTripAnimationAtom records".to_string(),
                        ));
                    }
                    validate_variable_atom(record, "RoundTripAnimationAtom")?;
                    metadata.animation_package = Some(parse_animation_package(&record.data)?);
                },
                RecordType::RoundTripAnimationHash12Atom => {
                    if metadata.animation_checksum.is_some() {
                        return Err(Error::Corrupted(
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

pub(crate) fn validate_variable_atom(record: &Record, name: &str) -> Result<()> {
    if record.version != 0
        || record.instance != 0
        || record.data_length as usize != record.data.len()
    {
        return Err(Error::Corrupted(format!(
            "{name} has an invalid record header or size"
        )));
    }
    Ok(())
}

pub(crate) fn parse_embedded_xml_package(
    data: &[u8],
    record_name: &str,
    expected_content_type: &str,
    expected_namespace: &[u8],
    expected_root: &[u8],
) -> Result<EmbeddedXmlPackage> {
    let package = OpcPackage::from_bytes(data).map_err(|error| {
        Error::Corrupted(format!(
            "{record_name} contains an invalid ECMA-376 package: {error}"
        ))
    })?;
    let mut found_part_name = None;
    for part in package.iter_parts() {
        if part.content_type() == expected_content_type {
            if found_part_name.is_some() {
                return Err(Error::Corrupted(format!(
                    "{record_name} package has multiple expected XML parts"
                )));
            }
            if !xml_has_root(part.blob(), expected_namespace, expected_root).map_err(|error| {
                Error::Corrupted(format!(
                    "{record_name} XML part {} is invalid: {error}",
                    part.partname()
                ))
            })? {
                return Err(Error::Corrupted(format!(
                    "{record_name} part {} has an invalid root element",
                    part.partname()
                )));
            }
            found_part_name = Some(part.partname().to_string());
        } else if is_xml_content_type(part.content_type()) {
            validate_xml_with(part.blob(), |_, _, _, _| Ok(())).map_err(|error| {
                Error::Corrupted(format!(
                    "{record_name} XML part {} is invalid: {error}",
                    part.partname()
                ))
            })?;
        }
    }
    let xml_part_name = found_part_name.ok_or_else(|| {
        Error::Corrupted(format!("{record_name} package has no expected XML part"))
    })?;
    Ok(EmbeddedXmlPackage {
        data: data.to_vec(),
        part_count: package.part_count(),
        xml_part_name,
    })
}

pub(crate) fn parse_theme_package(data: &[u8]) -> Result<ThemePackage> {
    let package = OpcPackage::from_bytes(data).map_err(|error| {
        Error::Corrupted(format!(
            "RoundTripThemeAtom contains an invalid ECMA-376 package: {error}"
        ))
    })?;
    let mut theme_part = None;
    for part in package.iter_parts() {
        let expected = match part.content_type() {
            content_type::OFC_THEME => Some((ThemeKind::Theme, b"theme".as_slice())),
            content_type::OFC_THEME_OVERRIDE => {
                Some((ThemeKind::ThemeOverride, b"themeOverride".as_slice()))
            },
            _ => None,
        };
        if let Some((kind, expected_root)) = expected {
            if theme_part.is_some() {
                return Err(Error::Corrupted(
                    "RoundTripThemeAtom package has multiple Theme parts".to_string(),
                ));
            }
            if !xml_has_root(part.blob(), DRAWINGML_NAMESPACE, expected_root).map_err(|error| {
                Error::Corrupted(format!(
                    "RoundTripThemeAtom Theme part {} is invalid: {error}",
                    part.partname()
                ))
            })? {
                return Err(Error::Corrupted(format!(
                    "RoundTripThemeAtom part {} has an invalid root element",
                    part.partname()
                )));
            }
            theme_part = Some((part.partname().to_string(), kind));
        } else if is_xml_content_type(part.content_type()) {
            validate_xml_with(part.blob(), |_, _, _, _| Ok(())).map_err(|error| {
                Error::Corrupted(format!(
                    "RoundTripThemeAtom XML part {} is invalid: {error}",
                    part.partname()
                ))
            })?;
        }
    }
    let (theme_part_name, kind) = theme_part.ok_or_else(|| {
        Error::Corrupted(
            "RoundTripThemeAtom package has no Theme or Theme Override part".to_string(),
        )
    })?;
    Ok(ThemePackage {
        data: data.to_vec(),
        part_count: package.part_count(),
        theme_part_name,
        kind,
    })
}

pub(crate) fn parse_color_mapping(data: &[u8]) -> Result<ColorMapping> {
    #[derive(Clone, Copy)]
    enum RootKind {
        Direct,
        Override,
    }

    let xml = std::str::from_utf8(data).map_err(|error| {
        Error::Corrupted(format!(
            "RoundTripColorMappingAtom is not valid UTF-8: {error}"
        ))
    })?;
    let mut root_kind = None;
    let mut result = None;
    validate_xml_with(data, |namespace, element, depth, decoder| {
        if depth == 0 {
            if xml_name(namespace, element, DRAWINGML_NAMESPACE, b"clrMap") {
                root_kind = Some(RootKind::Direct);
                result = Some((
                    ColorMappingKind::Direct,
                    Some(parse_color_mapping_values(element, decoder)?),
                ));
            } else if xml_name(namespace, element, PRESENTATIONML_NAMESPACE, b"clrMapOvr") {
                root_kind = Some(RootKind::Override);
            }
        } else if depth == 1 && matches!(root_kind, Some(RootKind::Override)) {
            let candidate =
                if xml_name(namespace, element, DRAWINGML_NAMESPACE, b"masterClrMapping") {
                    Some((ColorMappingKind::MasterOverride, None))
                } else if xml_name(
                    namespace,
                    element,
                    DRAWINGML_NAMESPACE,
                    b"overrideClrMapping",
                ) {
                    Some((
                        ColorMappingKind::ExplicitOverride,
                        Some(parse_color_mapping_values(element, decoder)?),
                    ))
                } else {
                    None
                };
            if let Some(choice) = candidate
                && result.replace(choice).is_some()
            {
                return Err("clrMapOvr contains multiple color-mapping choices".to_string());
            }
        }
        Ok(())
    })
    .map_err(|error| {
        Error::Corrupted(format!(
            "RoundTripColorMappingAtom contains invalid XML: {error}"
        ))
    })?;
    if root_kind.is_none() {
        return Err(Error::Corrupted(
            "RoundTripColorMappingAtom has an invalid color-mapping root element".to_string(),
        ));
    }
    let (kind, values) = result.ok_or_else(|| {
        Error::Corrupted(
            "RoundTripColorMappingAtom clrMapOvr has no color-mapping choice".to_string(),
        )
    })?;
    Ok(ColorMapping {
        xml: xml.to_string(),
        kind,
        values,
    })
}

fn parse_color_mapping_values(
    element: &BytesStart<'_>,
    decoder: Decoder,
) -> std::result::Result<ColorMappingValues, String> {
    let mut values = [None; 12];
    for attribute_result in element.attributes().with_checks(true) {
        let attribute = attribute_result.map_err(|error| error.to_string())?;
        let index = match attribute.key.as_ref() {
            b"bg1" => 0,
            b"tx1" => 1,
            b"bg2" => 2,
            b"tx2" => 3,
            b"accent1" => 4,
            b"accent2" => 5,
            b"accent3" => 6,
            b"accent4" => 7,
            b"accent5" => 8,
            b"accent6" => 9,
            b"hlink" => 10,
            b"folHlink" => 11,
            b"xmlns" => continue,
            other if other.contains(&b':') => continue,
            other => {
                return Err(format!(
                    "color mapping has unexpected attribute {}",
                    String::from_utf8_lossy(other)
                ));
            },
        };
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| error.to_string())?;
        let scheme_index = parse_color_scheme_index(&value)?;
        if values[index].replace(scheme_index).is_some() {
            return Err(format!(
                "color mapping has duplicate attribute {}",
                String::from_utf8_lossy(attribute.key.as_ref())
            ));
        }
    }
    let [
        Some(background1),
        Some(text1),
        Some(background2),
        Some(text2),
        Some(accent1),
        Some(accent2),
        Some(accent3),
        Some(accent4),
        Some(accent5),
        Some(accent6),
        Some(hyperlink),
        Some(followed_hyperlink),
    ] = values
    else {
        return Err("color mapping does not define all 12 required roles".to_string());
    };
    Ok(ColorMappingValues {
        background1,
        text1,
        background2,
        text2,
        accent1,
        accent2,
        accent3,
        accent4,
        accent5,
        accent6,
        hyperlink,
        followed_hyperlink,
    })
}

fn parse_color_scheme_index(value: &str) -> std::result::Result<ColorSchemeIndex, String> {
    match value {
        "dk1" => Ok(ColorSchemeIndex::Dark1),
        "lt1" => Ok(ColorSchemeIndex::Light1),
        "dk2" => Ok(ColorSchemeIndex::Dark2),
        "lt2" => Ok(ColorSchemeIndex::Light2),
        "accent1" => Ok(ColorSchemeIndex::Accent1),
        "accent2" => Ok(ColorSchemeIndex::Accent2),
        "accent3" => Ok(ColorSchemeIndex::Accent3),
        "accent4" => Ok(ColorSchemeIndex::Accent4),
        "accent5" => Ok(ColorSchemeIndex::Accent5),
        "accent6" => Ok(ColorSchemeIndex::Accent6),
        "hlink" => Ok(ColorSchemeIndex::Hyperlink),
        "folHlink" => Ok(ColorSchemeIndex::FollowedHyperlink),
        _ => Err(format!("invalid color-scheme index '{value}'")),
    }
}

pub(crate) fn parse_animation_package(data: &[u8]) -> Result<AnimationPackage> {
    let package = OpcPackage::from_bytes(data).map_err(|error| {
        Error::Corrupted(format!(
            "RoundTripAnimationAtom contains an invalid ECMA-376 package: {error}"
        ))
    })?;
    let mut timing_relationships = package
        .rels()
        .iter()
        .filter(|relationship| relationship.reltype() == TIMING_INFO_RELATIONSHIP_TYPE);
    let timing_relationship = timing_relationships.next().ok_or_else(|| {
        Error::Corrupted(
            "RoundTripAnimationAtom package has no Timing Info relationship".to_string(),
        )
    })?;
    if timing_relationships.next().is_some() || timing_relationship.is_external() {
        return Err(Error::Corrupted(
            "RoundTripAnimationAtom package has invalid Timing Info relationships".to_string(),
        ));
    }
    let timing_part_name = timing_relationship.target_partname().map_err(|error| {
        Error::Corrupted(format!(
            "RoundTripAnimationAtom has an invalid Timing Info target: {error}"
        ))
    })?;
    let timing_part = package.get_part(&timing_part_name).map_err(|error| {
        Error::Corrupted(format!(
            "RoundTripAnimationAtom Timing Info part is invalid: {error}"
        ))
    })?;
    if timing_part.content_type() != TIMING_INFO_CONTENT_TYPE {
        return Err(Error::Corrupted(
            "RoundTripAnimationAtom Timing Info part has an invalid content type".to_string(),
        ));
    }
    if !xml_contains_presentation_timing(timing_part.blob()).map_err(|error| {
        Error::Corrupted(format!(
            "RoundTripAnimationAtom Timing Info XML is invalid: {error}"
        ))
    })? {
        return Err(Error::Corrupted(
            "RoundTripAnimationAtom Timing Info part has no PresentationML timing element"
                .to_string(),
        ));
    }
    for part in package.iter_parts() {
        if part.partname() == &timing_part_name || !is_xml_content_type(part.content_type()) {
            continue;
        }
        xml_contains_presentation_timing(part.blob()).map_err(|error| {
            Error::Corrupted(format!(
                "RoundTripAnimationAtom XML part {} is invalid: {error}",
                part.partname()
            ))
        })?;
    }
    Ok(AnimationPackage {
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
    let mut contains_timing = false;
    validate_xml_with(data, |namespace, element, _, _| {
        if xml_name(namespace, element, PRESENTATIONML_NAMESPACE, b"timing") {
            contains_timing = true;
        }
        Ok(())
    })?;
    Ok(contains_timing)
}

fn xml_has_root(
    data: &[u8],
    expected_namespace: &[u8],
    expected_local_name: &[u8],
) -> std::result::Result<bool, String> {
    let mut matches_root = false;
    validate_xml_with(data, |namespace, element, depth, _| {
        if depth == 0 {
            matches_root = xml_name(namespace, element, expected_namespace, expected_local_name);
        }
        Ok(())
    })?;
    Ok(matches_root)
}

fn xml_name(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    expected_namespace: &[u8],
    expected_local_name: &[u8],
) -> bool {
    element.local_name().as_ref() == expected_local_name
        && matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == expected_namespace)
}

fn validate_xml_with(
    data: &[u8],
    mut inspect: impl FnMut(
        &ResolveResult<'_>,
        &BytesStart<'_>,
        usize,
        Decoder,
    ) -> std::result::Result<(), String>,
) -> std::result::Result<(), String> {
    let mut reader = NsReader::from_reader(data);
    let mut depth = 0usize;
    let mut root_count = 0usize;
    loop {
        let decoder = reader.decoder();
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
                validate_xml_attributes(&element, decoder)?;
                inspect(&namespace, &element, depth, decoder)?;
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
                validate_xml_attributes(&element, decoder)?;
                inspect(&namespace, &element, depth, decoder)?;
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
            Event::Eof => return Ok(()),
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
    }
}

fn validate_xml_attributes(
    element: &BytesStart<'_>,
    decoder: Decoder,
) -> std::result::Result<(), String> {
    for attribute in element.attributes().with_checks(true) {
        attribute
            .map_err(|error| error.to_string())?
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| error.to_string())?;
    }
    Ok(())
}

fn validate_atom(
    record: &Record,
    name: &str,
    expected_length: usize,
    expected_instance: Option<u16>,
) -> Result<()> {
    if record.version != 0
        || expected_instance.is_some_and(|instance| record.instance != instance)
        || record.data_length as usize != expected_length
        || record.data.len() != expected_length
    {
        return Err(Error::Corrupted(format!(
            "{name} has an invalid record header or size"
        )));
    }
    Ok(())
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
mod tests {
    use super::*;
    use litchi_opc::constants::relationship_type;
    use litchi_opc::{PackURI, Part, XmlPart};
    use soapberry_zip::office::{ArchiveReader, StreamingArchiveWriter};
    use std::io::Cursor;

    const COLOR_MAPPING_ATTRIBUTES: &str = r#"bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink""#;

    fn record(record_type: RecordType, version: u16, instance: u16, data: &[u8]) -> Record {
        Record {
            version,
            instance,
            record_type,
            record_type_raw: record_type.as_u16(),
            data_length: u32::try_from(data.len()).unwrap(),
            data: data.to_vec(),
            children: Vec::new(),
        }
    }

    fn root(children: Vec<Record>) -> Record {
        let mut root = record(RecordType::Slide, 0x0f, 0, &[]);
        root.children = children;
        root
    }

    fn replace_zip_member_raw(package: &[u8], path: &str, replacement: &[u8]) -> Vec<u8> {
        let archive = ArchiveReader::new(package).unwrap();
        let mut writer = StreamingArchiveWriter::new();
        let path = path.trim_start_matches('/');
        let mut replaced = false;
        for name in archive.file_names() {
            let data = if name == path {
                replaced = true;
                replacement.to_vec()
            } else {
                archive.read(name).unwrap()
            };
            writer.write_stored(name, &data).unwrap();
        }
        assert!(replaced, "test ZIP member {path} must exist");
        writer.finish_to_bytes().unwrap()
    }

    fn animation_package(parts: &[(&str, &str, &str, &[u8])]) -> Vec<u8> {
        let mut package = OpcPackage::new();
        for (index, (name, content_type, relationship_type, _data)) in parts.iter().enumerate() {
            package.add_part(Box::new(XmlPart::new(
                PackURI::new(*name).unwrap(),
                (*content_type).to_string(),
                b"<root/>".to_vec(),
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
        let mut output = output.into_inner();
        for (name, _content_type, _relationship_type, data) in parts {
            output = replace_zip_member_raw(&output, name, data);
        }
        output
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

    fn theme_package(kind: ThemeKind, theme_xml: &[u8]) -> Vec<u8> {
        let mut package = OpcPackage::new();
        let mut manager = XmlPart::new(
            PackURI::new("/theme/theme/themeManager.xml").unwrap(),
            "application/vnd.openxmlformats-officedocument.themeManager+xml".to_string(),
            br#"<a:themeManager xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>"#
                .to_vec(),
        );
        let (theme_content_type, theme_relationship_type) = match kind {
            ThemeKind::Theme => (content_type::OFC_THEME, relationship_type::THEME),
            ThemeKind::ThemeOverride => (
                content_type::OFC_THEME_OVERRIDE,
                relationship_type::THEME_OVERRIDE,
            ),
        };
        manager.relate_to("theme1.xml", theme_relationship_type);
        package.add_part(Box::new(manager));
        package.add_part(Box::new(XmlPart::new(
            PackURI::new("/theme/theme/theme1.xml").unwrap(),
            theme_content_type.to_string(),
            b"<root/>".to_vec(),
        )));
        package.rels_mut().add_relationship(
            relationship_type::OFFICE_DOCUMENT.to_string(),
            "theme/theme/themeManager.xml".to_string(),
            "rId1".to_string(),
            false,
        );
        let mut output = Cursor::new(Vec::new());
        package.to_stream(&mut output).unwrap();
        replace_zip_member_raw(&output.into_inner(), "/theme/theme/theme1.xml", theme_xml)
    }

    fn direct_color_mapping_xml() -> Vec<u8> {
        format!(
            r#"<a:clrMap xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" {COLOR_MAPPING_ATTRIBUTES}/>"#
        )
        .into_bytes()
    }

    #[test]
    fn parses_direct_slide_master_references_and_retains_undefined_values() {
        let composite = record(
            RecordType::RoundTripCompositeMasterId12Atom,
            0,
            0,
            &u32::MAX.to_le_bytes(),
        );
        let mut content_payload = Vec::new();
        content_payload.extend_from_slice(&0u32.to_le_bytes());
        content_payload.extend_from_slice(&u16::MAX.to_le_bytes());
        content_payload.extend_from_slice(&0xa55au16.to_le_bytes());
        let content = record(
            RecordType::RoundTripContentMasterId12Atom,
            0,
            0x0fff,
            &content_payload,
        );

        let parsed = SlideRoundTripMetadata12::parse(&root(vec![composite, content])).unwrap();
        assert_eq!(parsed.composite_master_id, Some(u32::MAX));
        assert_eq!(
            parsed.content_master,
            Some(ContentMasterReference {
                record_instance: 0x0fff,
                main_master_id: 0,
                layout_instance_id: u16::MAX,
                unused: 0xa55a,
            })
        );
        assert_eq!(
            SlideRoundTripMetadata12::parse(&root(Vec::new())).unwrap(),
            SlideRoundTripMetadata12::default()
        );
    }

    #[test]
    fn rejects_malformed_or_duplicate_direct_slide_master_references() {
        let composite = |version, instance, data: &[u8]| {
            record(
                RecordType::RoundTripCompositeMasterId12Atom,
                version,
                instance,
                data,
            )
        };
        let content = |version, instance, data: &[u8]| {
            record(
                RecordType::RoundTripContentMasterId12Atom,
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
            assert!(SlideRoundTripMetadata12::parse(&root(vec![malformed])).is_err());
        }

        let mut wrong_declared_length = composite(0, 0, &[0; 4]);
        wrong_declared_length.data_length = 5;
        assert!(SlideRoundTripMetadata12::parse(&root(vec![wrong_declared_length])).is_err());

        let duplicate_composite = root(vec![composite(0, 0, &[0; 4]), composite(0, 0, &[1; 4])]);
        assert!(SlideRoundTripMetadata12::parse(&duplicate_composite).is_err());
        let duplicate_content = root(vec![content(0, 0, &[0; 8]), content(0, 1, &[1; 8])]);
        assert!(SlideRoundTripMetadata12::parse(&duplicate_content).is_err());
    }

    #[test]
    fn parses_theme_packages_and_all_color_mapping_forms() {
        let theme_data = theme_package(
            ThemeKind::Theme,
            br#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Test"/>"#,
        );
        let color_xml = direct_color_mapping_xml();
        let parsed = SlideRoundTripMetadata12::parse(&root(vec![
            record(RecordType::RoundTripTheme12Atom, 0, 0, &theme_data),
            record(RecordType::RoundTripColorMapping12Atom, 0, 0, &color_xml),
        ]))
        .unwrap();
        let theme = parsed.theme_package.unwrap();
        assert_eq!(theme.data, theme_data);
        assert_eq!(theme.part_count, 2);
        assert_eq!(theme.theme_part_name, "/theme/theme/theme1.xml");
        assert_eq!(theme.kind, ThemeKind::Theme);
        let mapping = parsed.color_mapping.unwrap();
        assert_eq!(mapping.xml.as_bytes(), color_xml);
        assert_eq!(mapping.kind, ColorMappingKind::Direct);
        assert_eq!(
            mapping.values,
            Some(ColorMappingValues {
                background1: ColorSchemeIndex::Light1,
                text1: ColorSchemeIndex::Dark1,
                background2: ColorSchemeIndex::Light2,
                text2: ColorSchemeIndex::Dark2,
                accent1: ColorSchemeIndex::Accent1,
                accent2: ColorSchemeIndex::Accent2,
                accent3: ColorSchemeIndex::Accent3,
                accent4: ColorSchemeIndex::Accent4,
                accent5: ColorSchemeIndex::Accent5,
                accent6: ColorSchemeIndex::Accent6,
                hyperlink: ColorSchemeIndex::Hyperlink,
                followed_hyperlink: ColorSchemeIndex::FollowedHyperlink,
            })
        );

        let override_data = theme_package(
            ThemeKind::ThemeOverride,
            br#"<a:themeOverride xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>"#,
        );
        let override_parsed = SlideRoundTripMetadata12::parse(&root(vec![record(
            RecordType::RoundTripTheme12Atom,
            0,
            0,
            &override_data,
        )]))
        .unwrap();
        assert_eq!(
            override_parsed.theme_package.unwrap().kind,
            ThemeKind::ThemeOverride
        );

        let master_override = br#"<p:clrMapOvr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:masterClrMapping/></p:clrMapOvr>"#;
        let master_override_mapping = SlideRoundTripMetadata12::parse(&root(vec![record(
            RecordType::RoundTripColorMapping12Atom,
            0,
            0,
            master_override,
        )]))
        .unwrap()
        .color_mapping
        .unwrap();
        assert_eq!(
            master_override_mapping.kind,
            ColorMappingKind::MasterOverride
        );
        assert_eq!(master_override_mapping.values, None);

        let explicit_override = format!(
            r#"<p:clrMapOvr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:overrideClrMapping {COLOR_MAPPING_ATTRIBUTES}/></p:clrMapOvr>"#
        );
        let explicit_override_mapping = SlideRoundTripMetadata12::parse(&root(vec![record(
            RecordType::RoundTripColorMapping12Atom,
            0,
            0,
            explicit_override.as_bytes(),
        )]))
        .unwrap()
        .color_mapping
        .unwrap();
        assert_eq!(
            explicit_override_mapping.kind,
            ColorMappingKind::ExplicitOverride
        );
        assert!(explicit_override_mapping.values.is_some());
    }

    #[test]
    fn rejects_malformed_or_duplicate_theme_and_color_mapping_records() {
        let theme_data = theme_package(
            ThemeKind::Theme,
            br#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>"#,
        );
        let color_xml = direct_color_mapping_xml();
        let theme = |version, instance, data: &[u8]| {
            record(RecordType::RoundTripTheme12Atom, version, instance, data)
        };
        let color = |version, instance, data: &[u8]| {
            record(
                RecordType::RoundTripColorMapping12Atom,
                version,
                instance,
                data,
            )
        };
        for malformed in [
            theme(1, 0, &theme_data),
            theme(0, 1, &theme_data),
            theme(0, 0, b"not a package"),
            color(1, 0, &color_xml),
            color(0, 1, &color_xml),
            color(0, 0, &[0xff]),
        ] {
            assert!(SlideRoundTripMetadata12::parse(&root(vec![malformed])).is_err());
        }

        let mut wrong_theme_length = theme(0, 0, &theme_data);
        wrong_theme_length.data_length -= 1;
        assert!(SlideRoundTripMetadata12::parse(&root(vec![wrong_theme_length])).is_err());
        let mut wrong_color_length = color(0, 0, &color_xml);
        wrong_color_length.data_length += 1;
        assert!(SlideRoundTripMetadata12::parse(&root(vec![wrong_color_length])).is_err());
        assert!(
            SlideRoundTripMetadata12::parse(&root(vec![
                theme(0, 0, &theme_data),
                theme(0, 0, &theme_data),
            ]))
            .is_err()
        );
        assert!(
            SlideRoundTripMetadata12::parse(&root(vec![
                color(0, 0, &color_xml),
                color(0, 0, &color_xml),
            ]))
            .is_err()
        );
    }

    #[test]
    fn rejects_invalid_theme_packages_and_color_mapping_schema_values() {
        let valid_theme_xml =
            br#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>"#;
        for package_data in [
            theme_package(
                ThemeKind::Theme,
                br#"<a:themeOverride xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>"#,
            ),
            theme_package(ThemeKind::Theme, b"<a:theme>"),
            animation_package(&[(
                "/theme/themeManager.xml",
                "application/vnd.openxmlformats-officedocument.themeManager+xml",
                "urn:litchi:test:manager",
                br#"<a:themeManager xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>"#,
            )]),
            animation_package(&[
                (
                    "/theme/theme1.xml",
                    content_type::OFC_THEME,
                    "urn:litchi:test:theme",
                    valid_theme_xml,
                ),
                (
                    "/theme/themeOverride1.xml",
                    content_type::OFC_THEME_OVERRIDE,
                    "urn:litchi:test:theme-override",
                    br#"<a:themeOverride xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>"#,
                ),
            ]),
            animation_package(&[
                (
                    "/theme/theme1.xml",
                    content_type::OFC_THEME,
                    "urn:litchi:test:theme",
                    valid_theme_xml,
                ),
                (
                    "/theme/broken.xml",
                    "application/xml",
                    "urn:litchi:test:metadata",
                    b"<broken>",
                ),
            ]),
        ] {
            let theme = record(RecordType::RoundTripTheme12Atom, 0, 0, &package_data);
            assert!(SlideRoundTripMetadata12::parse(&root(vec![theme])).is_err());
        }

        let missing_attribute = br#"<a:clrMap xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" bg1="lt1"/>"#;
        let invalid_value = String::from_utf8(direct_color_mapping_xml())
            .unwrap()
            .replace(r#"accent6="accent6""#, r#"accent6="invalid""#)
            .into_bytes();
        let unexpected_attribute = format!(
            r#"<a:clrMap xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" extra="1" {COLOR_MAPPING_ATTRIBUTES}/>"#
        );
        let missing_override_choice = br#"<p:clrMapOvr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>"#;
        let duplicate_override_choice = format!(
            r#"<p:clrMapOvr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:masterClrMapping/><a:overrideClrMapping {COLOR_MAPPING_ATTRIBUTES}/></p:clrMapOvr>"#
        );
        for xml in [
            missing_attribute.as_slice(),
            invalid_value.as_slice(),
            unexpected_attribute.as_bytes(),
            missing_override_choice.as_slice(),
            duplicate_override_choice.as_bytes(),
            br#"<a:clrMap xmlns:a="urn:not-drawingml"/>"#,
            b"<broken>",
            br#"<a:clrMap xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" broken=>"#,
        ] {
            let color = record(RecordType::RoundTripColorMapping12Atom, 0, 0, xml);
            assert!(SlideRoundTripMetadata12::parse(&root(vec![color])).is_err());
        }
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
        let package_record = record(RecordType::RoundTripAnimation12Atom, 0, 0, &package_data);
        let checksum_record = record(
            RecordType::RoundTripAnimationHash12Atom,
            0,
            0,
            &u32::MAX.to_le_bytes(),
        );

        let parsed =
            SlideRoundTripMetadata12::parse(&root(vec![checksum_record, package_record])).unwrap();
        assert_eq!(parsed.animation_checksum, Some(u32::MAX));
        let package = parsed.animation_package.unwrap();
        assert_eq!(package.data, package_data);
        assert_eq!(package.part_count, 2);
        assert_eq!(package.timing_part_name, "/drs/timingInfo.xml");

        let zero_checksum = record(
            RecordType::RoundTripAnimationHash12Atom,
            0,
            0,
            &0u32.to_le_bytes(),
        );
        assert_eq!(
            SlideRoundTripMetadata12::parse(&root(vec![zero_checksum]))
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
                RecordType::RoundTripAnimation12Atom,
                version,
                instance,
                data,
            )
        };
        let checksum = |version, instance, data: &[u8]| {
            record(
                RecordType::RoundTripAnimationHash12Atom,
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
            assert!(SlideRoundTripMetadata12::parse(&root(vec![malformed])).is_err());
        }

        let mut wrong_package_length = animation(0, 0, &package_data);
        wrong_package_length.data_length -= 1;
        assert!(SlideRoundTripMetadata12::parse(&root(vec![wrong_package_length])).is_err());
        let mut wrong_checksum_length = checksum(0, 0, &[0; 4]);
        wrong_checksum_length.data_length = 5;
        assert!(SlideRoundTripMetadata12::parse(&root(vec![wrong_checksum_length])).is_err());

        let duplicate_animation = root(vec![
            animation(0, 0, &package_data),
            animation(0, 0, &package_data),
        ]);
        assert!(SlideRoundTripMetadata12::parse(&duplicate_animation).is_err());
        let duplicate_checksum = root(vec![checksum(0, 0, &[0; 4]), checksum(0, 0, &[1; 4])]);
        assert!(SlideRoundTripMetadata12::parse(&duplicate_checksum).is_err());
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
                RecordType::RoundTripAnimation12Atom,
                0,
                0,
                &package_data,
            );
            assert!(SlideRoundTripMetadata12::parse(&root(vec![animation])).is_err());
        }
    }
}
