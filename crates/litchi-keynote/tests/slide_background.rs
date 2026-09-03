//! Integration coverage for selector-first Keynote slide-background edits.
//!
//! The fixtures deliberately use the smallest multi-component package that
//! still has a real slide-style parent chain.  This keeps the tests focused on
//! the public package API while exercising the same physical ownership shape
//! as a native presentation: slides live in their slide components, and style
//! objects live in the document stylesheet component.

use std::io;

use litchi_iwa_archive::{Limits, package::Catalog};
use litchi_iwa_common::{
    color::{RgbColorSpace, Rgba},
    shape::fill::{Angle, Gradient, Kind, Stop},
    wire::{WireView, append_length_delimited_field, append_varint_field},
};
use litchi_iwa_core::{
    Archive, ArchiveObject, FieldInfo, FieldType, MessageInfo, RawMessage, SnappyStream,
};
use litchi_iwa_protos::{kn, tsa, tsd, tsk, tsp, tss};
use litchi_keynote::{Background, Package, Position, SlideSelector};
use prost::Message as _;

const DOCUMENT_MEMBER: &str = "Index/Document.iwa";
const STYLESHEET_MEMBER: &str = "Index/DocumentStylesheet.iwa";
const FIRST_SLIDE_MEMBER: &str = "Index/Slide-4.iwa";
const SECOND_SLIDE_MEMBER: &str = "Index/Slide-10.iwa";
const METADATA_MEMBER: &str = "Index/Metadata.iwa";

const SHOW_ID: u64 = 2;
const FIRST_NODE_ID: u64 = 3;
const SECOND_NODE_ID: u64 = 9;
const FIRST_SLIDE_ID: u64 = 4;
const SECOND_SLIDE_ID: u64 = 10;
const STYLESHEET_ID: u64 = 81;
const BASE_STYLE_ID: u64 = 40;
const VARIATION_STYLE_ID: u64 = 50;
const STYLE_MESSAGE_TYPE: u32 = 9;
const STYLESHEET_MESSAGE_TYPE: u32 = 401;
const SLIDE_MESSAGE_TYPE: u32 = 5;
const PACKAGE_METADATA_MESSAGE_TYPE: u32 = 11_006;

const UNKNOWN_STYLE_FIELD: u32 = 99;
const UNKNOWN_FILL_FIELD: u32 = 100;
const UNKNOWN_CHILD_FIELD: u32 = 101;
const METADATA_OBJECT_ID: u64 = 300;
const METADATA_LAST_OBJECT_ID: u64 = 1_000;

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..tsp::Reference::default()
    }
}

fn object(identifier: u64, type_: u32, data: Vec<u8>) -> TestResult<ArchiveObject> {
    Ok(ArchiveObject::new(
        identifier,
        vec![RawMessage { type_, data }],
    )?)
}

fn component(objects: Vec<ArchiveObject>) -> TestResult<Vec<u8>> {
    Ok(SnappyStream::compress(&Archive { objects }.to_bytes()?)?)
}

fn solid(red: f32, green: f32, blue: f32, alpha: f32, color_space: RgbColorSpace) -> Background {
    Background::Solid(Rgba::new(red, green, blue, alpha, color_space).expect("test color is valid"))
}

fn white_color() -> Rgba {
    Rgba::new(1.0, 1.0, 1.0, 1.0, RgbColorSpace::Srgb).expect("test color is valid")
}

fn white() -> Background {
    Background::Solid(white_color())
}

fn green_color() -> Rgba {
    Rgba::new(0.1, 0.8, 0.3, 1.0, RgbColorSpace::Srgb).expect("test color is valid")
}

fn red() -> Background {
    solid(0.9, 0.2, 0.1, 0.75, RgbColorSpace::DisplayP3)
}

fn green() -> Background {
    Background::Solid(green_color())
}

fn gradient() -> Gradient {
    Gradient::linear(
        Rgba::new(0.0, 1.0, 1.0, 1.0, RgbColorSpace::Srgb).expect("test color is valid"),
        Rgba::new(0.0, 0.2, 1.0, 1.0, RgbColorSpace::Srgb).expect("test color is valid"),
        Angle::from_degrees(90.0).expect("test angle is valid"),
    )
}

fn native_solid_payload(background: &Background) -> Vec<u8> {
    let Background::Solid(color) = background else {
        panic!("expected solid background")
    };
    tsd::FillArchive {
        color: Some(tsp::Color {
            model: tsp::color::ColorModel::Rgb as i32,
            r: Some(color.red()),
            g: Some(color.green()),
            b: Some(color.blue()),
            a: Some(color.alpha()),
            rgbspace: Some(match color.color_space() {
                RgbColorSpace::Srgb => tsp::color::RgbColorSpace::Srgb as i32,
                RgbColorSpace::DisplayP3 => tsp::color::RgbColorSpace::P3 as i32,
            }),
            ..tsp::Color::default()
        }),
        ..tsd::FillArchive::default()
    }
    .encode_to_vec()
}

fn native_gradient_payload(value: &Gradient) -> Vec<u8> {
    tsd::FillArchive {
        gradient: Some(tsd::GradientArchive {
            r#type: Some(match value.kind() {
                Kind::Linear => tsd::gradient_archive::GradientType::Linear as i32,
                Kind::Radial => tsd::gradient_archive::GradientType::Radial as i32,
            }),
            stops: value
                .stops()
                .iter()
                .map(|stop| tsd::gradient_archive::GradientStop {
                    color: Some(native_color(stop)),
                    fraction: Some(stop.position().get()),
                    inflection: Some(stop.midpoint().get()),
                })
                .collect(),
            opacity: Some(value.opacity().get()),
            advanced_gradient: Some(value.is_advanced()),
            anglegradient: Some(tsd::AngleGradientArchive {
                gradientangle: Some(value.angle().radians()),
            }),
            transformgradient: None,
        }),
        ..tsd::FillArchive::default()
    }
    .encode_to_vec()
}

fn native_color(stop: &Stop) -> tsp::Color {
    let color = stop.color();
    tsp::Color {
        model: tsp::color::ColorModel::Rgb as i32,
        r: Some(color.red()),
        g: Some(color.green()),
        b: Some(color.blue()),
        a: Some(color.alpha()),
        rgbspace: Some(match color.color_space() {
            RgbColorSpace::Srgb => tsp::color::RgbColorSpace::Srgb as i32,
            RgbColorSpace::DisplayP3 => tsp::color::RgbColorSpace::P3 as i32,
        }),
        ..tsp::Color::default()
    }
}

fn fill_payload(background: &Background) -> Vec<u8> {
    match background {
        Background::None => tsd::FillArchive::default().encode_to_vec(),
        Background::Solid(_) => native_solid_payload(background),
        Background::Gradient(value) => native_gradient_payload(value),
        _ => panic!("unsupported test background variant"),
    }
}

fn replace_first_length_delimited(data: &[u8], number: u32, payload: &[u8]) -> TestResult<Vec<u8>> {
    let view = WireView::parse(data)?;
    let mut output = Vec::with_capacity(data.len() + payload.len());
    let mut replaced = false;
    for field in view.fields() {
        if field.number() == number && !replaced {
            append_length_delimited_field(&mut output, number, payload)?;
            replaced = true;
        } else {
            output.extend_from_slice(field.raw());
        }
    }
    if !replaced {
        append_length_delimited_field(&mut output, number, payload)?;
    }
    Ok(output)
}

fn append_unknown(data: &mut Vec<u8>, number: u32, value: u64) -> TestResult<()> {
    append_varint_field(data, number, value)?;
    Ok(())
}

fn style_payload(
    parent: Option<u64>,
    variation: bool,
    fill: Option<&[u8]>,
    unknown_style: Option<u64>,
) -> TestResult<Vec<u8>> {
    let style = kn::SlideStyleArchive {
        super_: tss::StyleArchive {
            name: (!variation || unknown_style.is_some()).then(|| {
                if variation {
                    "direct slide variation".to_owned()
                } else {
                    "layout slide style".to_owned()
                }
            }),
            style_identifier: (!variation || unknown_style.is_some()).then(|| {
                if variation {
                    "variation-style".to_owned()
                } else {
                    "base-style".to_owned()
                }
            }),
            parent: parent.map(reference),
            is_variation: Some(variation),
            stylesheet: Some(reference(STYLESHEET_ID)),
        },
        override_count: Some(u32::from(variation)),
        slide_properties: fill.map(|_| kn::SlideStylePropertiesArchive {
            fill: Some(tsd::FillArchive::default()),
            ..kn::SlideStylePropertiesArchive::default()
        }),
    };
    let mut output = style.encode_to_vec();
    if let Some(fill) = fill {
        let properties = style
            .slide_properties
            .as_ref()
            .expect("fill implies slide properties")
            .encode_to_vec();
        let properties = replace_first_length_delimited(&properties, 1, fill)?;
        output = replace_first_length_delimited(&output, 11, &properties)?;
    }
    if let Some(value) = unknown_style {
        append_unknown(&mut output, UNKNOWN_STYLE_FIELD, value)?;
    }
    Ok(output)
}

fn style_object(
    identifier: u64,
    parent: Option<u64>,
    variation: bool,
    fill: Option<&[u8]>,
    unknown_style: Option<u64>,
) -> TestResult<ArchiveObject> {
    let mut object = object(
        identifier,
        STYLE_MESSAGE_TYPE,
        style_payload(parent, variation, fill, unknown_style)?,
    )?;
    object.archive_info.message_infos[0]
        .object_references
        .extend(parent.into_iter().chain([STYLESHEET_ID]));
    Ok(object)
}

fn slide_object(identifier: u64, style: u64, name: &str) -> TestResult<ArchiveObject> {
    let slide = kn::SlideArchive {
        style: reference(style),
        transition: kn::TransitionArchive::default(),
        name: Some(name.to_owned()),
        in_document: true,
        ..kn::SlideArchive::default()
    };
    let mut object = object(identifier, SLIDE_MESSAGE_TYPE, slide.encode_to_vec())?;
    object.archive_info.message_infos[0]
        .object_references
        .push(style);
    Ok(object)
}

fn metadata_uuid(identifier: u64) -> tsp::ObjectUuidMapEntry {
    tsp::ObjectUuidMapEntry {
        identifier,
        uuid: tsp::Uuid {
            lower: identifier.saturating_add(10_000),
            upper: identifier.saturating_add(20_000),
        },
    }
}

fn metadata_external_reference(
    component_identifier: u64,
    object_identifier: u64,
) -> tsp::ComponentExternalReference {
    tsp::ComponentExternalReference {
        component_identifier,
        object_identifier: Some(object_identifier),
        is_weak: None,
    }
}

fn package_metadata_payload(first_style: u64, second_style: u64) -> Vec<u8> {
    tsp::PackageMetadata {
        last_object_identifier: METADATA_LAST_OBJECT_ID,
        components: vec![
            tsp::ComponentInfo {
                identifier: 1,
                preferred_locator: "Document".to_owned(),
                object_uuid_map_entries: vec![
                    metadata_uuid(SHOW_ID),
                    metadata_uuid(FIRST_NODE_ID),
                    metadata_uuid(SECOND_NODE_ID),
                ],
                ..tsp::ComponentInfo::default()
            },
            tsp::ComponentInfo {
                identifier: STYLESHEET_ID,
                preferred_locator: "DocumentStylesheet".to_owned(),
                object_uuid_map_entries: vec![
                    metadata_uuid(STYLESHEET_ID),
                    metadata_uuid(BASE_STYLE_ID),
                    metadata_uuid(VARIATION_STYLE_ID),
                ],
                ..tsp::ComponentInfo::default()
            },
            tsp::ComponentInfo {
                identifier: FIRST_SLIDE_ID,
                preferred_locator: "Slide".to_owned(),
                locator: Some("Slide-4".to_owned()),
                object_uuid_map_entries: vec![metadata_uuid(FIRST_SLIDE_ID)],
                external_references: vec![metadata_external_reference(STYLESHEET_ID, first_style)],
                ..tsp::ComponentInfo::default()
            },
            tsp::ComponentInfo {
                identifier: SECOND_SLIDE_ID,
                preferred_locator: "Slide".to_owned(),
                locator: Some("Slide-10".to_owned()),
                object_uuid_map_entries: vec![metadata_uuid(SECOND_SLIDE_ID)],
                external_references: vec![metadata_external_reference(STYLESHEET_ID, second_style)],
                ..tsp::ComponentInfo::default()
            },
        ],
        ..tsp::PackageMetadata::default()
    }
    .encode_to_vec()
}

fn synthetic_package_with_style_state(
    first_style: u64,
    second_style: u64,
    first_style_payload: Option<Vec<u8>>,
    second_style_payload: Option<Vec<u8>>,
    shared_variation: bool,
) -> TestResult<Vec<u8>> {
    let document = kn::DocumentArchive {
        super_: tsa::DocumentArchive {
            super_: tsk::DocumentArchive::default(),
            ..tsa::DocumentArchive::default()
        },
        show: reference(SHOW_ID),
        ..kn::DocumentArchive::default()
    };
    let show = kn::ShowArchive {
        theme: reference(80),
        slide_tree: kn::SlideTreeArchive {
            slides: vec![reference(FIRST_NODE_ID), reference(SECOND_NODE_ID)],
            ..kn::SlideTreeArchive::default()
        },
        size: tsp::Size {
            width: 1_024.0,
            height: 768.0,
        },
        stylesheet: reference(STYLESHEET_ID),
        ..kn::ShowArchive::default()
    };
    #[allow(deprecated, reason = "native schema retains cache fields")]
    let first_node = kn::SlideNodeArchive {
        slide: Some(reference(FIRST_SLIDE_ID)),
        is_skipped: false,
        has_builds: false,
        has_transition: false,
        ..kn::SlideNodeArchive::default()
    };
    #[allow(deprecated, reason = "native schema retains cache fields")]
    let second_node = kn::SlideNodeArchive {
        slide: Some(reference(SECOND_SLIDE_ID)),
        is_skipped: false,
        has_builds: false,
        has_transition: false,
        ..kn::SlideNodeArchive::default()
    };

    let mut document_objects = vec![
        object(1, 1, document.encode_to_vec())?,
        object(SHOW_ID, 2, show.encode_to_vec())?,
        object(FIRST_NODE_ID, 4, first_node.encode_to_vec())?,
        object(SECOND_NODE_ID, 4, second_node.encode_to_vec())?,
    ];
    document_objects[0].archive_info.message_infos[0]
        .object_references
        .push(SHOW_ID);
    document_objects[1].archive_info.message_infos[0]
        .object_references
        .extend([STYLESHEET_ID, FIRST_NODE_ID, SECOND_NODE_ID]);
    document_objects[2].archive_info.message_infos[0]
        .object_references
        .push(FIRST_SLIDE_ID);
    document_objects[3].archive_info.message_infos[0]
        .object_references
        .push(SECOND_SLIDE_ID);

    let mut style_objects = vec![style_object(
        BASE_STYLE_ID,
        None,
        false,
        first_style_payload.as_deref(),
        None,
    )?];
    if let Some(payload) = second_style_payload.as_deref() {
        style_objects.push(style_object(
            VARIATION_STYLE_ID,
            Some(BASE_STYLE_ID),
            true,
            Some(payload),
            None,
        )?);
    }
    let styles = if shared_variation {
        vec![reference(BASE_STYLE_ID), reference(VARIATION_STYLE_ID)]
    } else {
        style_objects
            .iter()
            .map(|value| reference(value.archive_info.identifier.unwrap_or_default()))
            .collect()
    };
    let style_children = if second_style_payload.is_some() {
        vec![tss::stylesheet_archive::StyleChildrenEntry {
            parent: reference(BASE_STYLE_ID),
            children: vec![reference(VARIATION_STYLE_ID)],
        }]
    } else {
        Vec::new()
    };
    let stylesheet = tss::StylesheetArchive {
        styles,
        parent_to_children_style_map: style_children,
        can_cull_styles: Some(true),
        ..tss::StylesheetArchive::default()
    };
    let mut stylesheet_object = object(
        STYLESHEET_ID,
        STYLESHEET_MESSAGE_TYPE,
        stylesheet.encode_to_vec(),
    )?;
    stylesheet_object.archive_info.message_infos[0]
        .object_references
        .extend(
            style_objects
                .iter()
                .filter_map(|value| value.archive_info.identifier),
        );
    style_objects.push(stylesheet_object);

    let first = slide_object(FIRST_SLIDE_ID, first_style, "Alpha")?;
    let second = slide_object(SECOND_SLIDE_ID, second_style, "Beta")?;

    let document_component = component(document_objects)?;
    let stylesheet_component = component(style_objects)?;
    let first_component = component(vec![first])?;
    let second_component = component(vec![second])?;
    let metadata_component = component(vec![object(
        METADATA_OBJECT_ID,
        PACKAGE_METADATA_MESSAGE_TYPE,
        package_metadata_payload(first_style, second_style),
    )?])?;
    Ok(litchi_iwa_archive::package::to_bytes(
        [
            ("Data/sentinel.bin", b"unrelated ZIP sentinel".as_slice()),
            (DOCUMENT_MEMBER, document_component.as_slice()),
            (STYLESHEET_MEMBER, stylesheet_component.as_slice()),
            (FIRST_SLIDE_MEMBER, first_component.as_slice()),
            (SECOND_SLIDE_MEMBER, second_component.as_slice()),
            (METADATA_MEMBER, metadata_component.as_slice()),
        ],
        Limits::default(),
    )?)
}

fn synthetic_package() -> TestResult<Vec<u8>> {
    let base = white();
    let direct = red();
    synthetic_package_with_style_state(
        BASE_STYLE_ID,
        VARIATION_STYLE_ID,
        Some(fill_payload(&base)),
        Some(fill_payload(&direct)),
        false,
    )
}

fn synthetic_package_with_inherited_none() -> TestResult<Vec<u8>> {
    synthetic_package_with_style_state(BASE_STYLE_ID, BASE_STYLE_ID, None, None, false)
}

fn synthetic_package_with_shared_variation() -> TestResult<Vec<u8>> {
    let base = white();
    let direct = red();
    synthetic_package_with_style_state(
        VARIATION_STYLE_ID,
        VARIATION_STYLE_ID,
        Some(fill_payload(&base)),
        Some(fill_payload(&direct)),
        true,
    )
}

fn exact_bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn component_stream(package: &[u8], member: &str) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(package)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == member)
        .ok_or_else(|| io::Error::other("missing synthetic component"))?;
    Ok(SnappyStream::decompress(entry.data())?.into_bytes())
}

fn message_payload(
    package: &[u8],
    member: &str,
    identifier: u64,
    type_: u32,
) -> TestResult<Vec<u8>> {
    let archive = Archive::parse(&component_stream(package, member)?)?;
    let object = archive
        .object(identifier)
        .ok_or_else(|| io::Error::other("missing synthetic object"))?;
    object
        .messages
        .iter()
        .find(|message| message.type_ == type_)
        .map(|message| message.data.clone())
        .ok_or_else(|| io::Error::other("missing synthetic message").into())
}

fn package_metadata(package: &[u8]) -> TestResult<tsp::PackageMetadata> {
    Ok(tsp::PackageMetadata::decode(
        message_payload(
            package,
            METADATA_MEMBER,
            METADATA_OBJECT_ID,
            PACKAGE_METADATA_MESSAGE_TYPE,
        )?
        .as_slice(),
    )?)
}

#[test]
fn reads_effective_and_direct_backgrounds_and_selects_by_name_or_position() -> TestResult {
    let package = Package::from_bytes(&synthetic_package()?)?;
    assert_eq!(package.slide_background(SlideSelector::index(0))?, white());
    assert_eq!(package.slide_background("Alpha")?, white());
    assert_eq!(
        package.slide_background_override(SlideSelector::index(0))?,
        None
    );
    assert_eq!(package.slide_background(SlideSelector::index(1))?, red());
    assert_eq!(
        package.slide_background_override(SlideSelector::name("Beta"))?,
        Some(red())
    );
    Ok(())
}

#[test]
fn set_solid_gradient_and_explicit_none_round_trip_with_reversible_patch() -> TestResult<()> {
    let package = Package::from_bytes(&synthetic_package()?)?;

    let solid_commit = package
        .edit_slide_background("Alpha")?
        .set_solid(green_color())?
        .commit()?;
    assert_eq!(solid_commit.package().slide_background(0usize)?, green());
    assert_eq!(
        solid_commit.package().slide_background_override(0usize)?,
        Some(green())
    );

    let gradient_value = gradient();
    let gradient_commit = solid_commit
        .package()
        .edit_slide_background(0usize)?
        .set_gradient(gradient_value.clone())?
        .commit()?;
    assert_eq!(
        gradient_commit.package().slide_background(0usize)?,
        Background::Gradient(gradient_value.clone())
    );

    let none_commit = gradient_commit
        .package()
        .edit_slide_background(0usize)?
        .set(Background::None)?
        .commit()?;
    assert_eq!(
        none_commit.package().slide_background(0usize)?,
        Background::None
    );
    assert_eq!(
        none_commit.package().slide_background_override(0usize)?,
        Some(Background::None)
    );
    Ok(())
}

#[test]
fn no_op_preserves_exact_source_and_reset_restores_inheritance() -> TestResult<()> {
    let package = Package::from_bytes(&synthetic_package()?)?;
    let source = exact_bytes(&package)?;
    let source_metadata = package_metadata(&source)?;
    let noop = package
        .edit_slide_background(0usize)?
        .set_solid(white_color())?
        .commit()?;
    assert!(noop.patch().is_noop());
    assert_eq!(exact_bytes(noop.package())?, source);

    let direct_noop = package
        .edit_slide_background("Beta")?
        .set(red())?
        .commit()?;
    assert!(direct_noop.patch().is_noop());
    assert_eq!(exact_bytes(direct_noop.package())?, source);

    let inherited_none_source = synthetic_package_with_inherited_none()?;
    let inherited_none = Package::from_bytes(&inherited_none_source)?;
    let reset_noop = inherited_none
        .edit_slide_background(0usize)?
        .reset()?
        .commit()?;
    assert!(reset_noop.patch().is_noop());
    assert_eq!(exact_bytes(reset_noop.package())?, inherited_none_source);
    assert_eq!(
        reset_noop.package().slide_background_override(0usize)?,
        None
    );

    let cleared = package.edit_slide_background("Beta")?.reset()?.commit()?;
    assert_eq!(cleared.package().slide_background_override("Beta")?, None);
    assert_eq!(cleared.package().slide_background("Beta")?, white());
    let cleared_bytes = exact_bytes(cleared.package())?;
    let cleared_metadata = package_metadata(&cleared_bytes)?;
    assert!(cleared_metadata.last_object_identifier >= source_metadata.last_object_identifier);
    let cleared_stylesheet = cleared_metadata
        .components
        .iter()
        .find(|component| component.identifier == STYLESHEET_ID)
        .ok_or_else(|| io::Error::other("missing stylesheet metadata component"))?;
    assert!(
        !cleared_stylesheet
            .object_uuid_map_entries
            .iter()
            .any(|entry| entry.identifier == VARIATION_STYLE_ID)
    );
    let cleared_slide = cleared_metadata
        .components
        .iter()
        .find(|component| component.identifier == SECOND_SLIDE_ID)
        .ok_or_else(|| io::Error::other("missing second slide metadata component"))?;
    assert!(cleared_slide.external_references.iter().any(|reference| {
        reference.component_identifier == STYLESHEET_ID
            && reference.object_identifier == Some(BASE_STYLE_ID)
            && reference.is_weak.is_none()
    }));
    assert!(!cleared_slide.external_references.iter().any(|reference| {
        reference.component_identifier == STYLESHEET_ID
            && reference.object_identifier == Some(VARIATION_STYLE_ID)
    }));
    let restored = cleared
        .package()
        .apply_slide_background(&cleared.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, source);
    assert_eq!(restored.package().slide_background("Beta")?, red());
    Ok(())
}

#[test]
fn explicit_none_over_inherited_none_is_a_reversible_direct_override() -> TestResult<()> {
    let source = synthetic_package_with_inherited_none()?;
    let package = Package::from_bytes(&source)?;
    assert_eq!(package.slide_background(0usize)?, Background::None);
    assert_eq!(package.slide_background_override(0usize)?, None);

    let commit = package
        .edit_slide_background(0usize)?
        .set(Background::None)?
        .commit()?;
    assert!(!commit.patch().is_noop());
    assert_eq!(commit.package().slide_background(0usize)?, Background::None);
    assert_eq!(
        commit.package().slide_background_override(0usize)?,
        Some(Background::None)
    );

    let restored = commit
        .package()
        .apply_slide_background(&commit.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, source);
    assert_eq!(
        restored.package().slide_background(0usize)?,
        Background::None
    );
    assert_eq!(restored.package().slide_background_override(0usize)?, None);
    Ok(())
}

#[test]
fn selector_missing_and_ambiguous_fail_without_mutating_source() -> TestResult<()> {
    let mut bytes = synthetic_package()?;
    let mut stream = component_stream(&bytes, SECOND_SLIDE_MEMBER)?;
    let mut archive = Archive::parse(&stream)?;
    let second = archive
        .object_mut(SECOND_SLIDE_ID)
        .ok_or_else(|| io::Error::other("missing second slide"))?;
    let mut slide = kn::SlideArchive::decode(second.messages[0].data.as_slice())?;
    slide.name = Some("Alpha".to_owned());
    second.messages[0].data = slide.encode_to_vec();
    stream = archive.to_bytes()?;
    let compressed = SnappyStream::compress(&stream)?;
    bytes = Catalog::from_bytes(&bytes)?.reassemble_to_bytes(
        &[litchi_iwa_archive::package::EntryEdit::new(
            SECOND_SLIDE_MEMBER,
            &compressed,
        )],
        Limits::default(),
    )?;
    let ambiguous = Package::from_bytes(&bytes)?;
    let source = exact_bytes(&ambiguous)?;
    assert!(ambiguous.slide_background("Missing").is_err());
    assert!(ambiguous.edit_slide_background("Missing").is_err());
    assert!(ambiguous.slide_background("Alpha").is_err());
    assert!(ambiguous.edit_slide_background(Position::new(8)).is_err());
    assert_eq!(exact_bytes(&ambiguous)?, source);
    Ok(())
}

#[test]
fn shared_variation_is_copy_on_write_and_unselected_slide_is_unchanged() -> TestResult<()> {
    let package = Package::from_bytes(&synthetic_package_with_shared_variation()?)?;
    assert_eq!(package.slide_background(0usize)?, red());
    assert_eq!(package.slide_background(1usize)?, red());

    let commit = package
        .edit_slide_background(0usize)?
        .set(green())?
        .commit()?;
    assert_eq!(commit.package().slide_background(0usize)?, green());
    assert_eq!(commit.package().slide_background(1usize)?, red());
    assert_eq!(
        commit.package().slide_background_override(1usize)?,
        Some(red())
    );

    let reset = package.edit_slide_background(0usize)?.reset()?.commit()?;
    assert_eq!(reset.package().slide_background_override(0usize)?, None);
    assert_eq!(reset.package().slide_background(0usize)?, white());
    assert_eq!(reset.package().slide_background(1usize)?, red());
    assert_eq!(
        reset.package().slide_background_override(1usize)?,
        Some(red())
    );
    Ok(())
}

#[test]
fn unsupported_backgrounds_are_preserved_and_refuse_changed_edits_atomically() -> TestResult<()> {
    let canonical = native_gradient_payload(&gradient());
    let view = WireView::parse(&canonical)?;
    let mut gradient_payload = view
        .fields()
        .find(|field| field.number() == 2)
        .ok_or_else(|| io::Error::other("gradient payload missing"))?
        .payload()
        .to_vec();
    append_unknown(&mut gradient_payload, 99, 73)?;
    let future = replace_first_length_delimited(&canonical, 2, &gradient_payload)?;

    let bytes = synthetic_package_with_style_state(
        VARIATION_STYLE_ID,
        BASE_STYLE_ID,
        Some(fill_payload(&white())),
        Some(future),
        false,
    )?;
    let package = Package::from_bytes(&bytes)?;
    assert_eq!(package.slide_background(0usize)?, Background::Unsupported);
    assert_eq!(
        package.slide_background_override(0usize)?,
        Some(Background::Unsupported)
    );

    let source = exact_bytes(&package)?;
    let noop = package.edit_slide_background(0usize)?.commit()?;
    assert!(noop.patch().is_noop());
    assert_eq!(noop.patch().before(), &Background::Unsupported);
    assert_eq!(exact_bytes(noop.package())?, source);

    assert!(matches!(
        package
            .edit_slide_background(0usize)?
            .set(Background::Unsupported),
        Err(litchi_keynote::SlideBackgroundError::UnsupportedSource)
    ));
    assert!(matches!(
        package
            .edit_slide_background(0usize)?
            .set_solid(green_color())?
            .commit(),
        Err(litchi_keynote::SlideBackgroundError::UnsupportedSource)
    ));
    assert!(matches!(
        package.edit_slide_background(0usize)?.clear()?.commit(),
        Err(litchi_keynote::SlideBackgroundError::UnsupportedSource)
    ));
    assert_eq!(exact_bytes(&package)?, source);
    Ok(())
}

#[test]
fn unknown_style_and_fill_fields_survive_a_background_rewrite() -> TestResult<()> {
    let base = white();
    let direct = red();
    let mut base_fill = fill_payload(&base);
    append_unknown(&mut base_fill, UNKNOWN_FILL_FIELD, 7_303)?;
    let mut bytes = synthetic_package_with_style_state(
        BASE_STYLE_ID,
        VARIATION_STYLE_ID,
        Some(base_fill),
        Some(fill_payload(&direct)),
        false,
    )?;
    let stream = component_stream(&bytes, STYLESHEET_MEMBER)?;
    let mut archive = Archive::parse(&stream)?;
    let base_style = archive
        .object_mut(BASE_STYLE_ID)
        .ok_or_else(|| io::Error::other("missing base style"))?;
    append_unknown(&mut base_style.messages[0].data, UNKNOWN_STYLE_FIELD, 7_301)?;
    let stream = archive.to_bytes()?;
    let compressed = SnappyStream::compress(&stream)?;
    bytes = Catalog::from_bytes(&bytes)?.reassemble_to_bytes(
        &[litchi_iwa_archive::package::EntryEdit::new(
            STYLESHEET_MEMBER,
            &compressed,
        )],
        Limits::default(),
    )?;
    let package = Package::from_bytes(&bytes)?;
    let commit = package
        .edit_slide_background(0usize)?
        .set(green())?
        .commit()?;
    let rewritten_slide = message_payload(
        &exact_bytes(commit.package())?,
        FIRST_SLIDE_MEMBER,
        FIRST_SLIDE_ID,
        SLIDE_MESSAGE_TYPE,
    )?;
    let rewritten_slide = kn::SlideArchive::decode(rewritten_slide.as_slice())?;
    let style = message_payload(
        &exact_bytes(commit.package())?,
        STYLESHEET_MEMBER,
        rewritten_slide.style.identifier,
        STYLE_MESSAGE_TYPE,
    )?;
    let style_view = WireView::parse(&style)?;
    let parent_style = message_payload(
        &exact_bytes(commit.package())?,
        STYLESHEET_MEMBER,
        BASE_STYLE_ID,
        STYLE_MESSAGE_TYPE,
    )?;
    assert!(
        WireView::parse(&parent_style)?
            .fields()
            .any(|field| field.number() == UNKNOWN_STYLE_FIELD)
    );
    let fill = style_view
        .fields()
        .find(|field| field.number() == 11)
        .and_then(|field| WireView::parse(field.payload()).ok())
        .and_then(|properties| properties.fields().find(|field| field.number() == 1))
        .ok_or_else(|| io::Error::other("rewritten fill missing"))?;
    assert!(
        WireView::parse(fill.payload())?
            .fields()
            .any(|field| field.number() == UNKNOWN_FILL_FIELD)
    );
    assert_eq!(commit.package().slide_background(0usize)?, green());
    Ok(())
}

#[test]
fn package_metadata_registers_new_style_and_inverse_restores_it() -> TestResult<()> {
    let source = synthetic_package()?;
    let package = Package::from_bytes(&source)?;
    let source_metadata = package_metadata(&source)?;
    let commit = package
        .edit_slide_background(0usize)?
        .set_solid(green_color())?
        .commit()?;
    let changed = exact_bytes(commit.package())?;
    let rewritten_slide = kn::SlideArchive::decode(
        message_payload(
            &changed,
            FIRST_SLIDE_MEMBER,
            FIRST_SLIDE_ID,
            SLIDE_MESSAGE_TYPE,
        )?
        .as_slice(),
    )?;
    let new_style_identifier = rewritten_slide.style.identifier;
    assert!(new_style_identifier > METADATA_LAST_OBJECT_ID);

    let metadata = package_metadata(&changed)?;
    assert_eq!(metadata.last_object_identifier, new_style_identifier);
    let stylesheet = metadata
        .components
        .iter()
        .find(|component| component.identifier == STYLESHEET_ID)
        .ok_or_else(|| io::Error::other("missing stylesheet metadata component"))?;
    assert!(stylesheet.object_uuid_map_entries.iter().any(|entry| {
        entry.identifier == new_style_identifier && (entry.uuid.lower != 0 || entry.uuid.upper != 0)
    }));
    let slide = metadata
        .components
        .iter()
        .find(|component| component.identifier == FIRST_SLIDE_ID)
        .ok_or_else(|| io::Error::other("missing slide metadata component"))?;
    assert!(slide.external_references.iter().any(|reference| {
        reference.component_identifier == STYLESHEET_ID
            && reference.object_identifier == Some(new_style_identifier)
            && reference.is_weak.is_none()
    }));
    assert!(!slide.external_references.iter().any(|reference| {
        reference.component_identifier == STYLESHEET_ID
            && reference.object_identifier == Some(BASE_STYLE_ID)
    }));

    let restored = commit
        .package()
        .apply_slide_background(&commit.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, source);
    assert_eq!(
        package_metadata(&exact_bytes(restored.package())?)?,
        source_metadata
    );
    Ok(())
}

#[test]
fn malformed_background_is_rejected_and_edits_are_atomic() -> TestResult<()> {
    let mut bytes = synthetic_package()?;
    let mut style_stream = component_stream(&bytes, STYLESHEET_MEMBER)?;
    let mut archive = Archive::parse(&style_stream)?;
    let style = archive
        .object_mut(BASE_STYLE_ID)
        .ok_or_else(|| io::Error::other("missing base style"))?;
    let malformed = [0x0a, 0x02, 0x08, 0x01];
    style.messages[0].data = style_payload(None, false, Some(&malformed), Some(7_304))?;
    style_stream = archive.to_bytes()?;
    let compressed = SnappyStream::compress(&style_stream)?;
    bytes = Catalog::from_bytes(&bytes)?.reassemble_to_bytes(
        &[litchi_iwa_archive::package::EntryEdit::new(
            STYLESHEET_MEMBER,
            &compressed,
        )],
        Limits::default(),
    )?;
    let package = Package::from_bytes(&bytes)?;
    let source = exact_bytes(&package)?;
    assert!(package.slide_background(0usize).is_err());
    assert!(package.edit_slide_background(0usize).is_err());
    assert_eq!(exact_bytes(&package)?, source);
    Ok(())
}

#[test]
fn companion_style_messages_fail_closed_without_mutation() -> TestResult<()> {
    let mut bytes = synthetic_package()?;
    let stream = component_stream(&bytes, STYLESHEET_MEMBER)?;
    let mut archive = Archive::parse(&stream)?;
    let style = archive
        .object_mut(BASE_STYLE_ID)
        .ok_or_else(|| io::Error::other("missing base style"))?;
    let payload = b"unknown companion style message".to_vec();
    style
        .archive_info
        .message_infos
        .push(MessageInfo::new(991, u32::try_from(payload.len())?));
    style.messages.push(RawMessage {
        type_: 991,
        data: payload,
    });
    let stream = archive.to_bytes()?;
    let compressed = SnappyStream::compress(&stream)?;
    bytes = Catalog::from_bytes(&bytes)?.reassemble_to_bytes(
        &[litchi_iwa_archive::package::EntryEdit::new(
            STYLESHEET_MEMBER,
            &compressed,
        )],
        Limits::default(),
    )?;
    let package = Package::from_bytes(&bytes)?;
    let source = exact_bytes(&package)?;
    assert!(
        package
            .edit_slide_background(0usize)?
            .set_solid(green_color())?
            .commit()
            .is_err()
    );
    assert_eq!(exact_bytes(&package)?, source);
    Ok(())
}

#[test]
fn source_style_merge_and_diff_metadata_fail_closed_atomically() -> TestResult<()> {
    for use_diff_metadata in [false, true] {
        let mut bytes = synthetic_package()?;
        let stream = component_stream(&bytes, STYLESHEET_MEMBER)?;
        let mut archive = Archive::parse(&stream)?;
        let style = archive
            .object_mut(BASE_STYLE_ID)
            .ok_or_else(|| io::Error::other("missing base style"))?;
        if use_diff_metadata {
            style.archive_info.message_infos[0]
                .diff_merge_version
                .push(14);
        } else {
            style.archive_info.should_merge = Some(true);
        }
        let stream = archive.to_bytes()?;
        let compressed = SnappyStream::compress(&stream)?;
        bytes = Catalog::from_bytes(&bytes)?.reassemble_to_bytes(
            &[litchi_iwa_archive::package::EntryEdit::new(
                STYLESHEET_MEMBER,
                &compressed,
            )],
            Limits::default(),
        )?;

        let package = Package::from_bytes(&bytes)?;
        let source = exact_bytes(&package)?;
        assert!(
            package
                .edit_slide_background(0usize)?
                .set(green())
                .and_then(|edit| edit.commit())
                .is_err(),
            "style metadata variant should reject copy-on-write"
        );
        assert_eq!(exact_bytes(&package)?, source);
    }
    Ok(())
}

#[test]
fn reset_rejects_aggregate_only_fill_resource_metadata_atomically() -> TestResult<()> {
    for data_reference in [false, true] {
        let mut bytes = synthetic_package()?;
        let stream = component_stream(&bytes, STYLESHEET_MEMBER)?;
        let mut archive = Archive::parse(&stream)?;
        let style = archive
            .object_mut(VARIATION_STYLE_ID)
            .ok_or_else(|| io::Error::other("missing variation style"))?;
        let info = &mut style.archive_info.message_infos[0];
        if data_reference {
            info.data_references.push(90_002);
        } else {
            info.object_references.push(90_001);
        }
        let stream = archive.to_bytes()?;
        let compressed = SnappyStream::compress(&stream)?;
        bytes = Catalog::from_bytes(&bytes)?.reassemble_to_bytes(
            &[litchi_iwa_archive::package::EntryEdit::new(
                STYLESHEET_MEMBER,
                &compressed,
            )],
            Limits::default(),
        )?;

        let package = Package::from_bytes(&bytes)?;
        let source = exact_bytes(&package)?;
        assert!(
            package
                .edit_slide_background("Beta")?
                .reset()?
                .commit()
                .is_err()
        );
        assert_eq!(exact_bytes(&package)?, source);
    }
    Ok(())
}

#[test]
fn reset_rejects_identity_edges_attributed_to_fill_metadata_atomically() -> TestResult<()> {
    for reference in [BASE_STYLE_ID, STYLESHEET_ID] {
        let mut bytes = synthetic_package()?;
        let stream = component_stream(&bytes, STYLESHEET_MEMBER)?;
        let mut archive = Archive::parse(&stream)?;
        let style = archive
            .object_mut(VARIATION_STYLE_ID)
            .ok_or_else(|| io::Error::other("missing variation style"))?;
        let mut field = FieldInfo::new(vec![11, 1]);
        field.object_references.push(reference);
        style.archive_info.message_infos[0].field_infos.push(field);
        let stream = archive.to_bytes()?;
        let compressed = SnappyStream::compress(&stream)?;
        bytes = Catalog::from_bytes(&bytes)?.reassemble_to_bytes(
            &[litchi_iwa_archive::package::EntryEdit::new(
                STYLESHEET_MEMBER,
                &compressed,
            )],
            Limits::default(),
        )?;

        let package = Package::from_bytes(&bytes)?;
        let source = exact_bytes(&package)?;
        assert!(
            package
                .edit_slide_background("Beta")?
                .reset()?
                .commit()
                .is_err()
        );
        assert_eq!(exact_bytes(&package)?, source);
    }
    Ok(())
}

#[test]
fn non_selected_stylesheet_owner_prevents_style_culling() -> TestResult<()> {
    let mut bytes = synthetic_package()?;
    let stream = component_stream(&bytes, STYLESHEET_MEMBER)?;
    let mut archive = Archive::parse(&stream)?;
    let owner = object(
        82,
        STYLESHEET_MESSAGE_TYPE,
        tss::StylesheetArchive {
            styles: vec![reference(VARIATION_STYLE_ID)],
            can_cull_styles: Some(true),
            ..tss::StylesheetArchive::default()
        }
        .encode_to_vec(),
    )?;
    let mut owner = owner;
    owner.archive_info.message_infos[0]
        .object_references
        .push(VARIATION_STYLE_ID);
    archive.objects.push(owner);
    let stream = archive.to_bytes()?;
    let compressed = SnappyStream::compress(&stream)?;
    bytes = Catalog::from_bytes(&bytes)?.reassemble_to_bytes(
        &[litchi_iwa_archive::package::EntryEdit::new(
            STYLESHEET_MEMBER,
            &compressed,
        )],
        Limits::default(),
    )?;

    let package = Package::from_bytes(&bytes)?;
    let commit = package
        .edit_slide_background("Beta")?
        .set(green())?
        .commit()?;
    let changed = exact_bytes(commit.package())?;
    let rewritten_slide = kn::SlideArchive::decode(
        message_payload(
            &changed,
            SECOND_SLIDE_MEMBER,
            SECOND_SLIDE_ID,
            SLIDE_MESSAGE_TYPE,
        )?
        .as_slice(),
    )?;
    assert_ne!(rewritten_slide.style.identifier, VARIATION_STYLE_ID);
    assert_eq!(commit.package().slide_background("Beta")?, green());
    let changed_stylesheet = component_stream(&changed, STYLESHEET_MEMBER)?;
    assert!(
        Archive::parse(&changed_stylesheet)?
            .object(VARIATION_STYLE_ID)
            .is_some()
    );
    Ok(())
}

#[test]
fn unknown_other_style_field_prevents_selected_style_culling() -> TestResult<()> {
    let mut bytes = synthetic_package()?;
    let stream = component_stream(&bytes, STYLESHEET_MEMBER)?;
    let mut archive = Archive::parse(&stream)?;
    let mut owner_payload = style_payload(None, false, None, None)?;
    append_length_delimited_field(
        &mut owner_payload,
        UNKNOWN_STYLE_FIELD,
        &reference(VARIATION_STYLE_ID).encode_to_vec(),
    )?;
    archive
        .objects
        .push(object(82, STYLE_MESSAGE_TYPE, owner_payload)?);
    let stream = archive.to_bytes()?;
    let compressed = SnappyStream::compress(&stream)?;
    bytes = Catalog::from_bytes(&bytes)?.reassemble_to_bytes(
        &[litchi_iwa_archive::package::EntryEdit::new(
            STYLESHEET_MEMBER,
            &compressed,
        )],
        Limits::default(),
    )?;

    let package = Package::from_bytes(&bytes)?;
    let commit = package.edit_slide_background("Beta")?.reset()?.commit()?;
    assert_eq!(commit.package().slide_background("Beta")?, white());
    let changed = exact_bytes(commit.package())?;
    let stylesheet = component_stream(&changed, STYLESHEET_MEMBER)?;
    let archive = Archive::parse(&stylesheet)?;
    assert!(archive.object(VARIATION_STYLE_ID).is_some());
    let owner = archive
        .object(82)
        .ok_or_else(|| io::Error::other("missing other style owner"))?;
    assert!(
        WireView::parse(&owner.messages[0].data)?
            .fields()
            .any(|field| field.number() == UNKNOWN_STYLE_FIELD)
    );
    Ok(())
}

#[test]
fn other_style_stylesheet_reference_prevents_selected_style_culling() -> TestResult<()> {
    let mut bytes = synthetic_package()?;
    let stream = component_stream(&bytes, STYLESHEET_MEMBER)?;
    let mut archive = Archive::parse(&stream)?;
    let owner_payload = style_payload(None, false, None, None)?;
    let owner_view = WireView::parse(&owner_payload)?;
    let super_payload = owner_view
        .fields()
        .find(|field| field.number() == 1)
        .ok_or_else(|| io::Error::other("missing style super payload"))?
        .payload();
    let super_payload = replace_first_length_delimited(
        super_payload,
        5,
        &reference(VARIATION_STYLE_ID).encode_to_vec(),
    )?;
    let owner_payload = replace_first_length_delimited(&owner_payload, 1, &super_payload)?;
    archive
        .objects
        .push(object(82, STYLE_MESSAGE_TYPE, owner_payload)?);
    let stream = archive.to_bytes()?;
    let compressed = SnappyStream::compress(&stream)?;
    bytes = Catalog::from_bytes(&bytes)?.reassemble_to_bytes(
        &[litchi_iwa_archive::package::EntryEdit::new(
            STYLESHEET_MEMBER,
            &compressed,
        )],
        Limits::default(),
    )?;

    let package = Package::from_bytes(&bytes)?;
    let commit = package.edit_slide_background("Beta")?.reset()?.commit()?;
    assert_eq!(commit.package().slide_background("Beta")?, white());
    let changed = exact_bytes(commit.package())?;
    let stylesheet = component_stream(&changed, STYLESHEET_MEMBER)?;
    let archive = Archive::parse(&stylesheet)?;
    assert!(archive.object(VARIATION_STYLE_ID).is_some());
    assert!(archive.object(82).is_some());
    Ok(())
}

#[test]
fn unsupported_slide_style_metadata_refs_fail_closed_atomically() -> TestResult<()> {
    for variant in 0..4 {
        let mut bytes = synthetic_package()?;
        let stream = component_stream(&bytes, FIRST_SLIDE_MEMBER)?;
        let mut archive = Archive::parse(&stream)?;
        let slide = archive
            .object_mut(FIRST_SLIDE_ID)
            .ok_or_else(|| io::Error::other("missing first slide"))?;
        match variant {
            0 => slide.archive_info.message_infos[0]
                .object_references
                .push(BASE_STYLE_ID),
            1 => {
                let mut field = FieldInfo::new(vec![99]);
                field.object_references.push(BASE_STYLE_ID);
                slide.archive_info.message_infos[0].field_infos.push(field);
            },
            2 => slide.archive_info.message_infos[0]
                .data_references
                .push(BASE_STYLE_ID),
            _ => {
                let mut field = FieldInfo::new(vec![1]);
                field.r#type = Some(FieldType::DataReference);
                field.object_references.push(BASE_STYLE_ID);
                slide.archive_info.message_infos[0].field_infos.push(field);
            },
        }
        let stream = archive.to_bytes()?;
        let compressed = SnappyStream::compress(&stream)?;
        bytes = Catalog::from_bytes(&bytes)?.reassemble_to_bytes(
            &[litchi_iwa_archive::package::EntryEdit::new(
                FIRST_SLIDE_MEMBER,
                &compressed,
            )],
            Limits::default(),
        )?;

        let package = Package::from_bytes(&bytes)?;
        let source = exact_bytes(&package)?;
        assert!(
            package
                .edit_slide_background(0usize)?
                .set(green())?
                .commit()
                .is_err()
        );
        assert_eq!(exact_bytes(&package)?, source);
    }
    Ok(())
}

#[test]
fn unknown_slide_field_prevents_style_retargeting_atomically() -> TestResult<()> {
    let mut bytes = synthetic_package()?;
    let stream = component_stream(&bytes, FIRST_SLIDE_MEMBER)?;
    let mut archive = Archive::parse(&stream)?;
    let slide = archive
        .object_mut(FIRST_SLIDE_ID)
        .ok_or_else(|| io::Error::other("missing first slide"))?;
    append_length_delimited_field(
        &mut slide.messages[0].data,
        UNKNOWN_STYLE_FIELD,
        &reference(BASE_STYLE_ID).encode_to_vec(),
    )?;
    let stream = archive.to_bytes()?;
    let compressed = SnappyStream::compress(&stream)?;
    bytes = Catalog::from_bytes(&bytes)?.reassemble_to_bytes(
        &[litchi_iwa_archive::package::EntryEdit::new(
            FIRST_SLIDE_MEMBER,
            &compressed,
        )],
        Limits::default(),
    )?;

    let package = Package::from_bytes(&bytes)?;
    let source = exact_bytes(&package)?;
    assert!(
        package
            .edit_slide_background(0usize)?
            .set(green())?
            .commit()
            .is_err()
    );
    assert_eq!(exact_bytes(&package)?, source);
    Ok(())
}

#[test]
fn sparse_versioned_stylesheet_metadata_allows_cow_and_stays_unchanged() -> TestResult<()> {
    let mut bytes = synthetic_package()?;
    let stream = component_stream(&bytes, STYLESHEET_MEMBER)?;
    let mut archive = Archive::parse(&stream)?;
    let stylesheet = archive
        .object_mut(STYLESHEET_ID)
        .ok_or_else(|| io::Error::other("missing stylesheet"))?;
    let mut payload = tss::StylesheetArchive::decode(&*stylesheet.messages[0].data)?;
    payload.styles_for_12_1 = Some(tss::stylesheet_archive::VersionedStyles {
        styles: vec![reference(VARIATION_STYLE_ID)],
        ..tss::stylesheet_archive::VersionedStyles::default()
    });
    stylesheet.messages[0].data = payload.encode_to_vec();
    stylesheet.archive_info.message_infos[0].field_infos = vec![FieldInfo::new(vec![14])];
    let stream = archive.to_bytes()?;
    let compressed = SnappyStream::compress(&stream)?;
    bytes = Catalog::from_bytes(&bytes)?.reassemble_to_bytes(
        &[litchi_iwa_archive::package::EntryEdit::new(
            STYLESHEET_MEMBER,
            &compressed,
        )],
        Limits::default(),
    )?;

    let source_stylesheet = component_stream(&bytes, STYLESHEET_MEMBER)?;
    let source_archive = Archive::parse(&source_stylesheet)?;
    let source_object = source_archive
        .object(STYLESHEET_ID)
        .ok_or_else(|| io::Error::other("missing source stylesheet"))?;
    let source_versioned = WireView::parse(&source_object.messages[0].data)?
        .fields()
        .find(|field| field.number() == 14)
        .ok_or_else(|| io::Error::other("missing versioned stylesheet branch"))?
        .raw()
        .to_vec();
    let source_fields = source_object.archive_info.message_infos[0]
        .field_infos
        .clone();

    let package = Package::from_bytes(&bytes)?;
    let commit = package
        .edit_slide_background(0usize)?
        .set(green())?
        .commit()?;
    assert!(!commit.patch().is_noop());
    assert_eq!(commit.package().slide_background(0usize)?, green());

    let changed = exact_bytes(commit.package())?;
    let changed_stylesheet = component_stream(&changed, STYLESHEET_MEMBER)?;
    let changed_archive = Archive::parse(&changed_stylesheet)?;
    let changed_object = changed_archive
        .object(STYLESHEET_ID)
        .ok_or_else(|| io::Error::other("missing changed stylesheet"))?;
    let changed_versioned = WireView::parse(&changed_object.messages[0].data)?
        .fields()
        .find(|field| field.number() == 14)
        .ok_or_else(|| io::Error::other("missing changed versioned branch"))?
        .raw()
        .to_vec();
    assert_eq!(changed_versioned, source_versioned);
    assert_eq!(
        changed_object.archive_info.message_infos[0].field_infos,
        source_fields
    );

    let rewritten_slide = kn::SlideArchive::decode(
        message_payload(
            &changed,
            FIRST_SLIDE_MEMBER,
            FIRST_SLIDE_ID,
            SLIDE_MESSAGE_TYPE,
        )?
        .as_slice(),
    )?;
    assert_ne!(rewritten_slide.style.identifier, BASE_STYLE_ID);
    assert!(
        changed_archive
            .object(rewritten_slide.style.identifier)
            .is_some()
    );
    Ok(())
}

#[test]
fn unknown_style_child_entry_prevents_cull_and_survives_cow_insertion() -> TestResult<()> {
    let mut bytes = synthetic_package()?;
    let stream = component_stream(&bytes, STYLESHEET_MEMBER)?;
    let mut archive = Archive::parse(&stream)?;
    let stylesheet = archive
        .object_mut(STYLESHEET_ID)
        .ok_or_else(|| io::Error::other("missing stylesheet"))?;
    let stylesheet_data = stylesheet.messages[0].data.clone();
    let child = WireView::parse(&stylesheet_data)?
        .fields()
        .find(|field| field.number() == 5)
        .ok_or_else(|| io::Error::other("missing stylesheet child entry"))?
        .payload()
        .to_vec();
    let mut child_with_unknown = child;
    append_unknown(&mut child_with_unknown, UNKNOWN_CHILD_FIELD, 73)?;
    stylesheet.messages[0].data =
        replace_first_length_delimited(&stylesheet_data, 5, &child_with_unknown)?;
    let stream = archive.to_bytes()?;
    let compressed = SnappyStream::compress(&stream)?;
    bytes = Catalog::from_bytes(&bytes)?.reassemble_to_bytes(
        &[litchi_iwa_archive::package::EntryEdit::new(
            STYLESHEET_MEMBER,
            &compressed,
        )],
        Limits::default(),
    )?;

    let package = Package::from_bytes(&bytes)?;
    let commit = package
        .edit_slide_background(1usize)?
        .set(green())?
        .commit()?;
    assert!(!commit.patch().is_noop());
    assert_eq!(commit.package().slide_background(1usize)?, green());
    assert_eq!(commit.package().slide_background(0usize)?, white());

    let changed = exact_bytes(commit.package())?;
    let rewritten_slide = kn::SlideArchive::decode(
        message_payload(
            &changed,
            SECOND_SLIDE_MEMBER,
            SECOND_SLIDE_ID,
            SLIDE_MESSAGE_TYPE,
        )?
        .as_slice(),
    )?;
    assert_ne!(rewritten_slide.style.identifier, VARIATION_STYLE_ID);

    let changed_stylesheet = component_stream(&changed, STYLESHEET_MEMBER)?;
    let changed_archive = Archive::parse(&changed_stylesheet)?;
    assert!(changed_archive.object(VARIATION_STYLE_ID).is_some());
    let stylesheet = changed_archive
        .object(STYLESHEET_ID)
        .ok_or_else(|| io::Error::other("missing changed stylesheet"))?;
    let child_entries = WireView::parse(&stylesheet.messages[0].data)?
        .fields()
        .filter(|field| field.number() == 5)
        .map(|field| field.payload().to_vec())
        .collect::<Vec<_>>();
    assert!(!child_entries.is_empty(), "culled stylesheet child entry");
    let mut unknown_entry_found = false;
    let mut replacement_entry_found = false;
    for child in child_entries {
        let child_view = WireView::parse(&child)?;
        unknown_entry_found |= child_view
            .fields()
            .any(|field| field.number() == UNKNOWN_CHILD_FIELD && field.payload() == [73]);
        let child_ids = child_view
            .fields()
            .filter(|field| field.number() == 2)
            .map(|field| Ok(tsp::Reference::decode(field.payload())?.identifier))
            .collect::<TestResult<Vec<_>>>()?;
        replacement_entry_found |= child_ids.contains(&rewritten_slide.style.identifier);
    }
    assert!(unknown_entry_found);
    assert!(replacement_entry_found);
    Ok(())
}

#[test]
fn unknown_nested_reference_fields_survive_style_retargeting() -> TestResult<()> {
    let mut bytes = synthetic_package()?;
    let stream = component_stream(&bytes, FIRST_SLIDE_MEMBER)?;
    let mut archive = Archive::parse(&stream)?;
    let slide = archive
        .object_mut(FIRST_SLIDE_ID)
        .ok_or_else(|| io::Error::other("missing first slide"))?;
    let view = WireView::parse(&slide.messages[0].data)?;
    let mut reference_payload = view
        .fields()
        .find(|field| field.number() == 1)
        .ok_or_else(|| io::Error::other("slide style reference missing"))?
        .payload()
        .to_vec();
    append_unknown(&mut reference_payload, 99, 73)?;
    slide.messages[0].data =
        replace_first_length_delimited(&slide.messages[0].data, 1, &reference_payload)?;
    let stream = archive.to_bytes()?;
    let compressed = SnappyStream::compress(&stream)?;
    bytes = Catalog::from_bytes(&bytes)?.reassemble_to_bytes(
        &[litchi_iwa_archive::package::EntryEdit::new(
            FIRST_SLIDE_MEMBER,
            &compressed,
        )],
        Limits::default(),
    )?;

    let package = Package::from_bytes(&bytes)?;
    let commit = package
        .edit_slide_background(0usize)?
        .set_solid(green_color())?
        .commit()?;
    let slide = message_payload(
        &exact_bytes(commit.package())?,
        FIRST_SLIDE_MEMBER,
        FIRST_SLIDE_ID,
        SLIDE_MESSAGE_TYPE,
    )?;
    let slide_view = WireView::parse(&slide)?;
    let reference = slide_view
        .fields()
        .find(|field| field.number() == 1)
        .ok_or_else(|| io::Error::other("rewritten style reference missing"))?;
    assert!(
        WireView::parse(reference.payload())?
            .fields()
            .any(|field| field.number() == 99 && field.payload() == [73])
    );
    Ok(())
}

#[test]
fn unknown_stylesheet_reference_fields_prevent_cull_and_survive_reset() -> TestResult<()> {
    let mut bytes = synthetic_package()?;
    let stream = component_stream(&bytes, STYLESHEET_MEMBER)?;
    let mut archive = Archive::parse(&stream)?;
    let stylesheet = archive
        .object_mut(STYLESHEET_ID)
        .ok_or_else(|| io::Error::other("missing stylesheet"))?;
    let view = WireView::parse(&stylesheet.messages[0].data)?;
    let mut rewritten = Vec::new();
    for field in view.fields() {
        if field.number() == 1
            && tsp::Reference::decode(field.payload())?.identifier == VARIATION_STYLE_ID
        {
            let mut reference_payload = field.payload().to_vec();
            append_unknown(&mut reference_payload, UNKNOWN_STYLE_FIELD, 73)?;
            append_length_delimited_field(&mut rewritten, 1, &reference_payload)?;
        } else {
            rewritten.extend_from_slice(field.raw());
        }
    }
    stylesheet.messages[0].data = rewritten;
    let stream = archive.to_bytes()?;
    let compressed = SnappyStream::compress(&stream)?;
    bytes = Catalog::from_bytes(&bytes)?.reassemble_to_bytes(
        &[litchi_iwa_archive::package::EntryEdit::new(
            STYLESHEET_MEMBER,
            &compressed,
        )],
        Limits::default(),
    )?;

    let package = Package::from_bytes(&bytes)?;
    let commit = package.edit_slide_background("Beta")?.reset()?.commit()?;
    assert_eq!(commit.package().slide_background("Beta")?, white());
    let changed = exact_bytes(commit.package())?;
    let stylesheet = component_stream(&changed, STYLESHEET_MEMBER)?;
    let archive = Archive::parse(&stylesheet)?;
    assert!(archive.object(VARIATION_STYLE_ID).is_some());
    let stylesheet = archive
        .object(STYLESHEET_ID)
        .ok_or_else(|| io::Error::other("missing rewritten stylesheet"))?;
    let retained_extension = WireView::parse(&stylesheet.messages[0].data)?
        .fields()
        .filter(|field| field.number() == 1)
        .any(|field| {
            WireView::parse(field.payload()).is_ok_and(|reference| {
                reference.fields().any(|nested| {
                    nested.number() == UNKNOWN_STYLE_FIELD && nested.payload() == [73]
                })
            })
        });
    assert!(retained_extension);
    Ok(())
}

#[test]
fn stylesheet_cull_flag_is_authoritative() -> TestResult<()> {
    let mut bytes = synthetic_package()?;
    let stream = component_stream(&bytes, STYLESHEET_MEMBER)?;
    let mut archive = Archive::parse(&stream)?;
    let base = archive
        .object_mut(BASE_STYLE_ID)
        .ok_or_else(|| io::Error::other("missing base style"))?;
    base.messages[0].data = WireView::parse(&base.messages[0].data)?
        .fields()
        .filter(|field| field.number() != UNKNOWN_STYLE_FIELD)
        .flat_map(|field| field.raw().iter().copied())
        .collect();
    let stylesheet = archive
        .object_mut(STYLESHEET_ID)
        .ok_or_else(|| io::Error::other("missing stylesheet"))?;
    let mut stylesheet_payload = tss::StylesheetArchive::decode(&*stylesheet.messages[0].data)?;
    stylesheet_payload.can_cull_styles = Some(false);
    stylesheet.messages[0].data = stylesheet_payload.encode_to_vec();
    let stream = archive.to_bytes()?;
    let compressed = SnappyStream::compress(&stream)?;
    bytes = Catalog::from_bytes(&bytes)?.reassemble_to_bytes(
        &[litchi_iwa_archive::package::EntryEdit::new(
            STYLESHEET_MEMBER,
            &compressed,
        )],
        Limits::default(),
    )?;

    let package = Package::from_bytes(&bytes)?;
    let first = package
        .edit_slide_background(0usize)?
        .set_solid(green_color())?
        .commit()?;
    let first_slide = kn::SlideArchive::decode(
        message_payload(
            &exact_bytes(first.package())?,
            FIRST_SLIDE_MEMBER,
            FIRST_SLIDE_ID,
            SLIDE_MESSAGE_TYPE,
        )?
        .as_slice(),
    )?;
    let retained_style = first_slide.style.identifier;
    let second = first
        .package()
        .edit_slide_background(0usize)?
        .set_gradient(gradient())?
        .commit()?;
    let stylesheet = component_stream(&exact_bytes(second.package())?, STYLESHEET_MEMBER)?;
    assert!(
        Archive::parse(&stylesheet)?
            .object(retained_style)
            .is_some()
    );
    Ok(())
}

#[test]
fn output_limits_are_atomic() -> TestResult<()> {
    let bytes = synthetic_package()?;
    let source = Package::from_bytes(&bytes)?;
    let target = source
        .edit_slide_background(0usize)?
        .set(green())?
        .commit()?;
    let target_length = exact_bytes(target.package())?.len();
    let limits = Limits::new(
        u64::try_from(target_length - 1)?,
        32,
        1_024 * 1_024,
        1_024 * 1_024,
        1_024 * 1_024,
    )?;
    let limited = Package::from_bytes_with_limits(&bytes, limits)?;
    let source_bytes = exact_bytes(&limited)?;
    let result = limited
        .edit_slide_background(0usize)?
        .set(green())?
        .commit();
    assert!(result.is_err());
    assert_eq!(exact_bytes(&limited)?, source_bytes);
    Ok(())
}
