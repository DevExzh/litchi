//! Black-box coverage for selector-first Keynote slide-transition transactions.
//!
//! The fixture deliberately keeps opaque records, an unrelated IWA component,
//! and the `SlideNodeArchive.has_transition` cache around the edited payload.
//! That makes the assertions below a regression boundary rather than a
//! convenient round-trip through prost.

use std::io;
use std::sync::Arc;

use litchi_iwa_archive::package::Catalog;
use litchi_iwa_common::wire::WireView;
use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
use litchi_iwa_protos::{
    keynote_slide_transition_codec::{
        DecodeOptions as TransitionDecodeOptions, decode_slide_transition,
    },
    kn, tsa, tsd, tsk, tsp,
};
use litchi_keynote::{
    Limits, Package, Position, SlideSelector,
    transition::{
        Acceleration, AnimationParameters, Commit, CustomParameters, Diagnostics, Direction, Edit,
        Effect, Error, LimitKind, MosaicType, Patch, Settings, TextDelivery, TimingCurveSlot,
    },
};
use prost::Message as _;

const DOCUMENT_MEMBER: &str = "Index/Document.iwa";
const FIRST_NODE: u64 = 3;
const SECOND_NODE: u64 = 5;
const FIRST_SLIDE: u64 = 4;
const SECOND_SLIDE: u64 = 6;
const PRIVATE_MARKER: &str = "private-keynote-transition-marker-998244353";
const SLIDE_MESSAGE_TYPE: u32 = 5;

type TestResult<T> = Result<T, Box<dyn std::error::Error>>;

trait ExactPackageBytes {
    fn exact_bytes(&self) -> &'static [u8];
}

impl ExactPackageBytes for Package {
    fn exact_bytes(&self) -> &'static [u8] {
        let mut bytes = Vec::new();
        self.write_to(&mut bytes)
            .expect("an in-memory Vec accepts every package byte");
        Box::leak(bytes.into_boxed_slice())
    }
}

#[derive(Debug, Clone, Copy)]
enum Malformation {
    Canonical,
    None,
    DuplicateEffect,
    WrongEffectWire,
    NonCanonicalDirection,
}

#[derive(Debug, Clone, Copy)]
enum NodeMalformation {
    DuplicateFlag,
    WrongFlagWire,
    NonCanonicalFlag,
}

fn fixture_path() -> std::path::PathBuf {
    std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/keynote/basic.key")
}

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

fn object_with_references(
    identifier: u64,
    type_: u32,
    data: Vec<u8>,
    references: &[u64],
) -> TestResult<ArchiveObject> {
    let mut object = object(identifier, type_, data)?;
    object.archive_info.message_infos[0].object_references = references.to_vec();
    Ok(object)
}

fn component(objects: Vec<ArchiveObject>) -> TestResult<Vec<u8>> {
    Ok(SnappyStream::compress(&Archive { objects }.to_bytes()?)?)
}

#[allow(clippy::cast_possible_truncation, reason = "protobuf varint byte")]
fn push_varint(mut value: u64, output: &mut Vec<u8>) {
    while value >= 0x80 {
        output.push(((value as u8) & 0x7f) | 0x80);
        value >>= 7;
    }
    output.push(value as u8);
}

fn length_delimited_field(number: u32, payload: &[u8]) -> Vec<u8> {
    let mut result = Vec::with_capacity(payload.len().saturating_add(8));
    push_varint((u64::from(number) << 3) | 2, &mut result);
    push_varint(payload.len() as u64, &mut result);
    result.extend_from_slice(payload);
    result
}

fn fixed32_field(number: u32, value: u32) -> Vec<u8> {
    let mut result = Vec::with_capacity(8);
    push_varint((u64::from(number) << 3) | 5, &mut result);
    result.extend_from_slice(&value.to_le_bytes());
    result
}

fn full_settings() -> TestResult<Settings> {
    let color = tsp::Color {
        model: tsp::color::ColorModel::Rgb as i32,
        r: Some(0.25),
        g: Some(0.5),
        b: Some(0.75),
        a: Some(1.0),
        ..tsp::Color::default()
    }
    .encode_to_vec();
    let first_curve = tsd::PathSourceArchive {
        localization_key: Some("first-curve".to_owned()),
        ..tsd::PathSourceArchive::default()
    }
    .encode_to_vec();
    let second_curve = tsd::PathSourceArchive {
        user_defined_name: Some("second-curve".to_owned()),
        ..tsd::PathSourceArchive::default()
    }
    .encode_to_vec();
    let third_curve = tsd::PathSourceArchive {
        horizontal_flip: Some(true),
        ..tsd::PathSourceArchive::default()
    }
    .encode_to_vec();
    let mut animation = AnimationParameters::new();
    animation.set_color_payload(Some(&color))?;
    animation.set_timing_curve_payload(TimingCurveSlot::First, Some(&first_curve))?;
    animation.set_timing_curve_payload(TimingCurveSlot::Second, Some(&second_curve))?;
    animation.set_timing_curve_payload(TimingCurveSlot::Third, Some(&third_curve))?;
    animation.set_random_number_seed(Some(u32::MAX));
    animation.set_detail(Some(2.5))?;
    animation.set_timing_curve_theme_name(TimingCurveSlot::First, Some("alpha"))?;
    animation.set_timing_curve_theme_name(TimingCurveSlot::Second, Some("beta"))?;
    animation.set_timing_curve_theme_name(TimingCurveSlot::Third, Some("gamma"))?;
    animation.set_writing_direction_is_rtl(Some(true));
    let mut custom = CustomParameters::new();
    custom.set_twist(Some(1.25))?;
    custom.set_mosaic_size(Some(u32::MAX));
    custom.set_mosaic_type(Some(MosaicType::from_native(u32::MAX)));
    custom.set_bounce(Some(true));
    custom.set_magic_move_fade_unmatched_objects(Some(false));
    custom.set_acceleration(Some(Acceleration::from_native(-7)));
    custom.set_text_delivery(Some(TextDelivery::from_native(-9)));
    custom.set_motion_blur(Some(true));
    custom.set_travel_distance(Some(3.5))?;
    Ok(Settings::builder()
        .animation_type(Some("future-animation"))?
        .effect(Some(Effect::unknown("future:effect")?))?
        .duration(Some(1.5))?
        .direction(Some(Direction::from_native(u32::MAX)))
        .delay(Some(2.25))?
        .is_automatic(Some(true))
        .animation_parameters(animation)?
        .custom_parameters(custom)?
        .build()?)
}

fn native_none_from(settings: &Settings) -> TestResult<Settings> {
    let source_animation = settings.animation_parameters();
    let mut animation = AnimationParameters::new();
    animation.set_random_number_seed(source_animation.random_number_seed());
    animation.set_writing_direction_is_rtl(source_animation.writing_direction_is_rtl());
    Settings::builder()
        .animation_type(Some("Transition"))?
        .effect(Some(Effect::None))?
        .duration(Some(1.0))?
        .delay(settings.delay())?
        .is_automatic(settings.is_automatic())
        .animation_parameters(animation)?
        .custom_parameters(CustomParameters::new())?
        .build()
        .map_err(Into::into)
}

fn native_transition(
    settings: Option<&Settings>,
    malformation: Malformation,
) -> TestResult<Vec<u8>> {
    let transition_attributes = if let Some(value) = settings {
        let animation = value.animation_parameters();
        let decode_color = || {
            animation
                .color_payload()
                .map(tsp::Color::decode)
                .transpose()
        };
        let decode_curve = |slot| {
            animation
                .timing_curve_payload(slot)
                .map(tsd::PathSourceArchive::decode)
                .transpose()
        };
        kn::TransitionAttributesArchive {
            animation_attributes: Some(kn::AnimationAttributesArchive {
                animation_type: value.animation_type().map(str::to_owned),
                effect: value.effect().map(|effect| effect.identifier().to_owned()),
                duration: value.duration(),
                direction: value.direction().map(Direction::native_value),
                delay: value.delay(),
                is_automatic: value.is_automatic(),
                color: decode_color()?,
                custom_effect_timing_curve_1: decode_curve(TimingCurveSlot::First)?,
                custom_effect_timing_curve_2: decode_curve(TimingCurveSlot::Second)?,
                custom_effect_timing_curve_3: decode_curve(TimingCurveSlot::Third)?,
                random_number_seed: animation.random_number_seed(),
                custom_detail: animation.detail(),
                custom_effect_timing_curve_theme_name_1: animation
                    .timing_curve_theme_name(TimingCurveSlot::First)
                    .map(str::to_owned),
                custom_effect_timing_curve_theme_name_2: animation
                    .timing_curve_theme_name(TimingCurveSlot::Second)
                    .map(str::to_owned),
                custom_effect_timing_curve_theme_name_3: animation
                    .timing_curve_theme_name(TimingCurveSlot::Third)
                    .map(str::to_owned),
                writing_direction_is_rtl: animation.writing_direction_is_rtl(),
            }),
            custom_twist: value.custom_parameters().twist(),
            custom_mosaic_size: value.custom_parameters().mosaic_size(),
            custom_mosaic_type: value
                .custom_parameters()
                .mosaic_type()
                .map(MosaicType::native_value),
            custom_bounce: value.custom_parameters().bounce(),
            custom_magic_move_fade_unmatched_objects: value
                .custom_parameters()
                .magic_move_fade_unmatched_objects(),
            custom_timing_curve: value
                .custom_parameters()
                .acceleration()
                .map(Acceleration::native_value),
            custom_text_delivery_type: value
                .custom_parameters()
                .text_delivery()
                .map(TextDelivery::native_value),
            custom_motion_blur: value.custom_parameters().motion_blur(),
            custom_travel_distance: value.custom_parameters().travel_distance(),
            ..kn::TransitionAttributesArchive::default()
        }
    } else {
        kn::TransitionAttributesArchive::default()
    };
    let canonical = kn::TransitionArchive {
        attributes: transition_attributes,
    }
    .encode_to_vec();
    if matches!(malformation, Malformation::Canonical) {
        return Ok(canonical);
    }
    if matches!(malformation, Malformation::None) {
        let root = WireView::parse(&canonical)?;
        let envelope_attributes = root
            .fields()
            .find(|field| field.number() == 2)
            .ok_or_else(|| io::Error::other("transition has no attributes"))?;
        return transition_with_unknown_records(envelope_attributes.payload());
    }
    let root = WireView::parse(&canonical)?;
    let envelope_attributes = root
        .fields()
        .find(|field| field.number() == 2)
        .ok_or_else(|| io::Error::other("transition has no attributes"))?;
    let view = WireView::parse(envelope_attributes.payload())?;
    let animation = view
        .fields()
        .find(|field| field.number() == 8)
        .ok_or_else(|| io::Error::other("transition has no animation"))?;
    let animation_view = WireView::parse(animation.payload())?;
    let mut rewritten_animation = Vec::new();
    for field in animation_view.fields() {
        if field.number() == 2 {
            match malformation {
                Malformation::DuplicateEffect => {
                    rewritten_animation.extend_from_slice(field.raw());
                    rewritten_animation.extend_from_slice(field.raw());
                },
                Malformation::WrongEffectWire => {
                    rewritten_animation.extend_from_slice(&fixed32_field(2, 7));
                },
                Malformation::Canonical
                | Malformation::None
                | Malformation::NonCanonicalDirection => {
                    rewritten_animation.extend_from_slice(field.raw());
                },
            }
        } else if field.number() == 4 && matches!(malformation, Malformation::NonCanonicalDirection)
        {
            // field 4's value is deliberately encoded in a non-minimal width.
            rewritten_animation.extend_from_slice(&[0x20, 0xff, 0x00]);
        } else {
            rewritten_animation.extend_from_slice(field.raw());
        }
    }
    let mut rewritten_attributes = Vec::new();
    for field in view.fields() {
        if field.number() == 8 {
            rewritten_attributes
                .extend_from_slice(&length_delimited_field(8, &rewritten_animation));
        } else {
            rewritten_attributes.extend_from_slice(field.raw());
        }
    }
    transition_with_unknown_records(&rewritten_attributes)
}

fn transition_with_unknown_records(attributes: &[u8]) -> TestResult<Vec<u8>> {
    let attributes_view = WireView::parse(attributes)?;
    let mut rewritten_attributes = Vec::with_capacity(attributes.len().saturating_add(16));
    for field in attributes_view.fields() {
        if field.number() == 8 {
            let mut animation = field.payload().to_vec();
            animation.extend_from_slice(&fixed32_field(92, 0x1020_3040));
            rewritten_attributes.extend_from_slice(&length_delimited_field(8, &animation));
        } else {
            rewritten_attributes.extend_from_slice(field.raw());
        }
    }
    rewritten_attributes.extend_from_slice(&length_delimited_field(91, b"unknown-attributes"));
    let mut transition = length_delimited_field(2, &rewritten_attributes);
    transition.extend_from_slice(&fixed32_field(90, 0xaabb_ccdd));
    Ok(transition)
}

fn slide(
    name: &str,
    settings: Option<&Settings>,
    stale_has_transition: bool,
    malformation: Malformation,
) -> TestResult<(Vec<u8>, Vec<u8>)> {
    let canonical = kn::SlideArchive {
        style: reference(80),
        transition: kn::TransitionArchive::default(),
        name: Some(name.to_owned()),
        in_document: true,
        ..kn::SlideArchive::default()
    }
    .encode_to_vec();
    let transition = native_transition(settings, malformation)?;
    let view = WireView::parse(&canonical)?;
    let mut payload = Vec::with_capacity(canonical.len().saturating_add(transition.len()));
    for field in view.fields() {
        if field.number() == 4 {
            payload.extend_from_slice(&length_delimited_field(4, &transition));
        } else {
            payload.extend_from_slice(field.raw());
        }
    }
    payload.extend_from_slice(&length_delimited_field(99, b"unknown-slide"));
    #[allow(
        deprecated,
        reason = "native schema retains the required cached build state"
    )]
    let node = kn::SlideNodeArchive {
        slide: Some(reference(if name == "Alpha" {
            FIRST_SLIDE
        } else {
            SECOND_SLIDE
        })),
        is_skipped: false,
        has_builds: false,
        has_transition: stale_has_transition,
        ..kn::SlideNodeArchive::default()
    }
    .encode_to_vec();
    Ok((payload, node))
}

fn component_with_unknown_header(
    objects: Vec<ArchiveObject>,
    target_identifier: u64,
) -> TestResult<Vec<u8>> {
    let bytes = Archive { objects }.to_bytes()?;
    let archive = Archive::parse(&bytes)?;
    let object = archive
        .object(target_identifier)
        .ok_or_else(|| io::Error::other("synthetic object is missing"))?;
    let header_offset = usize::try_from(object.header_offset)?;
    let data_offset = usize::try_from(object.data_offset)?;
    let (raw_header_length, prefix_length) = litchi_iwa_common::decode_varint_from_bytes(
        bytes
            .get(header_offset..)
            .ok_or_else(|| io::Error::other("synthetic header offset is invalid"))?,
    )?;
    let header_length = usize::try_from(raw_header_length)?;
    let header_start = header_offset
        .checked_add(prefix_length)
        .ok_or_else(|| io::Error::other("synthetic header start overflows"))?;
    let header_end = header_start
        .checked_add(header_length)
        .ok_or_else(|| io::Error::other("synthetic header end overflows"))?;
    if header_end != data_offset {
        return Err(io::Error::other("synthetic header offsets disagree").into());
    }
    let mut header = bytes[header_start..header_end].to_vec();
    litchi_iwa_common::wire::append_varint_field(&mut header, 99, 9_999)?;
    let mut output = Vec::with_capacity(bytes.len().saturating_add(8));
    output.extend_from_slice(&bytes[..header_offset]);
    litchi_iwa_common::encode_varint_into(&mut output, u64::try_from(header.len())?);
    output.extend_from_slice(&header);
    output.extend_from_slice(&bytes[data_offset..]);
    Archive::parse(&output)?;
    Ok(SnappyStream::compress(&output)?)
}

fn package_bytes(
    first: Option<&Settings>,
    second: Option<&Settings>,
    names: [&str; 2],
    node_has_transition: [bool; 2],
    malformation: Malformation,
) -> TestResult<Vec<u8>> {
    let (first_slide, first_node) = slide(names[0], first, node_has_transition[0], malformation)?;
    let (second_slide, second_node) =
        slide(names[1], second, node_has_transition[1], Malformation::None)?;
    let document = kn::DocumentArchive {
        super_: tsa::DocumentArchive {
            super_: tsk::DocumentArchive::default(),
            ..tsa::DocumentArchive::default()
        },
        show: reference(2),
        ..kn::DocumentArchive::default()
    };
    let show = kn::ShowArchive {
        theme: reference(80),
        slide_tree: kn::SlideTreeArchive {
            slides: vec![reference(FIRST_NODE), reference(SECOND_NODE)],
            ..kn::SlideTreeArchive::default()
        },
        size: tsp::Size {
            width: 1024.0,
            height: 768.0,
        },
        stylesheet: reference(81),
        ..kn::ShowArchive::default()
    };
    let mut document_component = component(vec![
        object_with_references(1, 1, document.encode_to_vec(), &[2])?,
        object_with_references(2, 2, show.encode_to_vec(), &[FIRST_NODE, SECOND_NODE])?,
        object_with_references(FIRST_NODE, 4, first_node, &[FIRST_SLIDE])?,
        object_with_references(SECOND_NODE, 4, second_node, &[SECOND_SLIDE])?,
    ])?;
    // Keep a marker in the document component's zip sibling too; transactions
    // must only touch the selected slide component.
    let first_component = component_with_unknown_header(
        vec![ArchiveObject::new(
            FIRST_SLIDE,
            vec![
                RawMessage {
                    type_: 777,
                    data: b"before-transition".to_vec(),
                },
                RawMessage {
                    type_: SLIDE_MESSAGE_TYPE,
                    data: first_slide,
                },
                RawMessage {
                    type_: 778,
                    data: b"after-transition".to_vec(),
                },
            ],
        )?],
        FIRST_SLIDE,
    )?;
    let second_component = component(vec![object(
        SECOND_SLIDE,
        SLIDE_MESSAGE_TYPE,
        second_slide,
    )?])?;
    document_component.shrink_to_fit();
    Ok(litchi_iwa_archive::package::to_bytes(
        [
            ("Data/sentinel.bin", PRIVATE_MARKER.as_bytes()),
            (DOCUMENT_MEMBER, document_component.as_slice()),
            ("Index/Slide-4.iwa", first_component.as_slice()),
            ("Index/Slide-6.iwa", second_component.as_slice()),
        ],
        Limits::default(),
    )?)
}

fn co_located_package_bytes(
    first: Option<&Settings>,
    second: Option<&Settings>,
) -> TestResult<Vec<u8>> {
    let (first_slide, first_node) = slide("Alpha", first, false, Malformation::None)?;
    let (second_slide, second_node) = slide("Beta", second, false, Malformation::None)?;
    let document = kn::DocumentArchive {
        super_: tsa::DocumentArchive {
            super_: tsk::DocumentArchive::default(),
            ..tsa::DocumentArchive::default()
        },
        show: reference(2),
        ..kn::DocumentArchive::default()
    };
    let show = kn::ShowArchive {
        theme: reference(80),
        slide_tree: kn::SlideTreeArchive {
            slides: vec![reference(FIRST_NODE), reference(SECOND_NODE)],
            ..kn::SlideTreeArchive::default()
        },
        size: tsp::Size {
            width: 1024.0,
            height: 768.0,
        },
        stylesheet: reference(81),
        ..kn::ShowArchive::default()
    };
    let component = component_with_unknown_header(
        vec![
            object_with_references(1, 1, document.encode_to_vec(), &[2])?,
            object_with_references(2, 2, show.encode_to_vec(), &[FIRST_NODE, SECOND_NODE])?,
            object_with_references(FIRST_NODE, 4, first_node, &[FIRST_SLIDE])?,
            object_with_references(SECOND_NODE, 4, second_node, &[SECOND_SLIDE])?,
            object(FIRST_SLIDE, SLIDE_MESSAGE_TYPE, first_slide)?,
            object(SECOND_SLIDE, SLIDE_MESSAGE_TYPE, second_slide)?,
        ],
        FIRST_SLIDE,
    )?;
    Ok(litchi_iwa_archive::package::to_bytes(
        [
            ("Data/sentinel.bin", PRIVATE_MARKER.as_bytes()),
            (DOCUMENT_MEMBER, component.as_slice()),
        ],
        Limits::default(),
    )?)
}

fn legacy_package_bytes(flat: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(flat)?;
    let inner_entries = catalog
        .iter()
        .filter(|entry| {
            std::path::Path::new(entry.name())
                .extension()
                .is_some_and(|extension| extension.eq_ignore_ascii_case("iwa"))
        })
        .map(|entry| (entry.name(), entry.data()))
        .collect::<Vec<_>>();
    let inner = litchi_iwa_archive::package::to_bytes(inner_entries, Limits::default())?;
    Ok(litchi_iwa_archive::package::to_bytes(
        [
            ("legacy.key/Index.zip", inner.as_slice()),
            ("legacy.key/Data/sentinel.bin", b"legacy marker".as_slice()),
        ],
        Limits::default(),
    )?)
}

fn malformed_node_package(source: &[u8], malformation: NodeMalformation) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(source)?;
    let mut entries = Vec::with_capacity(catalog.len());
    for entry in catalog.iter() {
        let replacement = if entry.name() == DOCUMENT_MEMBER {
            let stream = SnappyStream::decompress(entry.data())?;
            let mut archive = Archive::parse(stream.as_bytes())?;
            let object = archive
                .object_mut(FIRST_NODE)
                .ok_or_else(|| io::Error::other("synthetic node is missing"))?;
            let message = object
                .messages
                .iter_mut()
                .find(|message| message.type_ == 4)
                .ok_or_else(|| io::Error::other("synthetic node payload is missing"))?;
            let view = WireView::parse(&message.data)?;
            let mut output = Vec::with_capacity(message.data.len().saturating_add(8));
            for field in view.fields() {
                if field.number() != 7 {
                    output.extend_from_slice(field.raw());
                    continue;
                }
                match malformation {
                    NodeMalformation::DuplicateFlag => {
                        output.extend_from_slice(field.raw());
                        output.extend_from_slice(field.raw());
                    },
                    NodeMalformation::WrongFlagWire => {
                        output.extend_from_slice(&length_delimited_field(7, &[1]));
                    },
                    NodeMalformation::NonCanonicalFlag => {
                        output.extend_from_slice(&[0x38, 0x81, 0x00]);
                    },
                }
            }
            message.data = output;
            SnappyStream::compress(&archive.to_bytes()?)?
        } else {
            entry.data().to_vec()
        };
        entries.push((entry.name().to_owned(), replacement));
    }
    Ok(litchi_iwa_archive::package::to_bytes(
        entries
            .iter()
            .map(|(name, data)| (name.as_str(), data.as_slice())),
        Limits::default(),
    )?)
}

fn mutate_component(
    source: &[u8],
    member: &str,
    mutate: impl FnOnce(&mut Archive) -> TestResult<()>,
) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(source)?;
    let mut mutation = Some(mutate);
    let mut entries = Vec::with_capacity(catalog.len());
    for entry in catalog.iter() {
        let data = if entry.name() == member {
            let stream = SnappyStream::decompress(entry.data())?;
            let mut archive = Archive::parse(stream.as_bytes())?;
            mutation
                .take()
                .ok_or_else(|| io::Error::other("test component was mutated twice"))?(
                &mut archive,
            )?;
            SnappyStream::compress(&archive.to_bytes()?)?
        } else {
            entry.data().to_vec()
        };
        entries.push((entry.name().to_owned(), data));
    }
    if mutation.is_some() {
        return Err(io::Error::other("test component was not found").into());
    }
    Ok(litchi_iwa_archive::package::to_bytes(
        entries
            .iter()
            .map(|(name, data)| (name.as_str(), data.as_slice())),
        Limits::default(),
    )?)
}

fn mixed_modern_legacy_package(source: &[u8]) -> TestResult<Vec<u8>> {
    mutate_component(source, "Index/Slide-4.iwa", |archive| {
        let slide = archive
            .object_mut(FIRST_SLIDE)
            .ok_or_else(|| io::Error::other("synthetic slide is missing"))?;
        let message = slide
            .messages
            .iter_mut()
            .find(|message| message.type_ == SLIDE_MESSAGE_TYPE)
            .ok_or_else(|| io::Error::other("synthetic slide payload is missing"))?;
        let root = WireView::parse(&message.data)?;
        let mut rewritten_slide = Vec::with_capacity(message.data.len().saturating_add(16));
        for field in root.fields() {
            if field.number() != 4 {
                rewritten_slide.extend_from_slice(field.raw());
                continue;
            }
            let transition = WireView::parse(field.payload())?;
            let mut rewritten_transition = Vec::with_capacity(field.payload().len() + 16);
            for nested in transition.fields() {
                if nested.number() == 2 {
                    let mut attributes = nested.payload().to_vec();
                    attributes.extend_from_slice(&length_delimited_field(1, b"legacy"));
                    rewritten_transition.extend_from_slice(&length_delimited_field(2, &attributes));
                } else {
                    rewritten_transition.extend_from_slice(nested.raw());
                }
            }
            rewritten_slide.extend_from_slice(&length_delimited_field(4, &rewritten_transition));
        }
        message.data = rewritten_slide;
        Ok(())
    })
}

fn duplicate_rooted_node_package(source: &[u8]) -> TestResult<Vec<u8>> {
    mutate_component(source, DOCUMENT_MEMBER, |archive| {
        let show = archive
            .object_mut(2)
            .ok_or_else(|| io::Error::other("synthetic show is missing"))?;
        let message = show
            .messages
            .iter_mut()
            .find(|message| message.type_ == 2)
            .ok_or_else(|| io::Error::other("synthetic show payload is missing"))?;
        let mut decoded = kn::ShowArchive::decode(message.data.as_slice())?;
        decoded.slide_tree.slides.push(reference(FIRST_NODE));
        message.data = decoded.encode_to_vec();
        Ok(())
    })
}

fn shared_slide_package(source: &[u8]) -> TestResult<Vec<u8>> {
    mutate_component(source, DOCUMENT_MEMBER, |archive| {
        let node = archive
            .object_mut(SECOND_NODE)
            .ok_or_else(|| io::Error::other("synthetic second node is missing"))?;
        let message = node
            .messages
            .iter_mut()
            .find(|message| message.type_ == 4)
            .ok_or_else(|| io::Error::other("synthetic second node payload is missing"))?;
        #[allow(
            deprecated,
            reason = "the native node reference is deliberately made adversarial"
        )]
        {
            let mut decoded = kn::SlideNodeArchive::decode(message.data.as_slice())?;
            decoded.slide = Some(reference(FIRST_SLIDE));
            message.data = decoded.encode_to_vec();
        }
        node.archive_info.message_infos[0].object_references = vec![FIRST_SLIDE];
        Ok(())
    })
}

fn metadata_or_framing_package(
    source: &[u8],
    mutate: impl FnOnce(&mut ArchiveObject) -> TestResult<()>,
) -> TestResult<Vec<u8>> {
    mutate_component(source, "Index/Slide-4.iwa", |archive| {
        let object = archive
            .object_mut(FIRST_SLIDE)
            .ok_or_else(|| io::Error::other("synthetic slide is missing"))?;
        mutate(object)
    })
}

fn noncanonical_object_length_prefix(source: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(source)?;
    let mut entries = Vec::with_capacity(catalog.len());
    for entry in catalog.iter() {
        let data = if entry.name() == "Index/Slide-4.iwa" {
            let stream = SnappyStream::decompress(entry.data())?;
            let archive = Archive::parse(stream.as_bytes())?;
            let object = archive
                .object(FIRST_SLIDE)
                .ok_or_else(|| io::Error::other("synthetic slide is missing"))?;
            let offset = usize::try_from(object.header_offset)?;
            let (_length, prefix_length) = litchi_iwa_common::decode_varint_from_bytes(
                stream
                    .as_bytes()
                    .get(offset..)
                    .ok_or_else(|| io::Error::other("test object offset is invalid"))?,
            )?;
            let last = offset
                .checked_add(prefix_length)
                .and_then(|end| end.checked_sub(1))
                .ok_or_else(|| io::Error::other("test object prefix overflows"))?;
            let mut raw = stream.as_bytes().to_vec();
            raw[last] |= 0x80;
            raw.insert(last + 1, 0);
            Archive::parse(&raw)?;
            SnappyStream::compress(&raw)?
        } else {
            entry.data().to_vec()
        };
        entries.push((entry.name().to_owned(), data));
    }
    Ok(litchi_iwa_archive::package::to_bytes(
        entries
            .iter()
            .map(|(name, data)| (name.as_str(), data.as_slice())),
        Limits::default(),
    )?)
}

fn wrong_message_info_type_package(source: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(source)?;
    let mut entries = Vec::with_capacity(catalog.len());
    for entry in catalog.iter() {
        let data = if entry.name() == "Index/Slide-4.iwa" {
            let stream = SnappyStream::decompress(entry.data())?;
            let archive = Archive::parse(stream.as_bytes())?;
            let object = archive
                .object(FIRST_SLIDE)
                .ok_or_else(|| io::Error::other("synthetic slide is missing"))?;
            let header_offset = usize::try_from(object.header_offset)?;
            let data_offset = usize::try_from(object.data_offset)?;
            let (header_length, prefix_length) = litchi_iwa_common::decode_varint_from_bytes(
                stream
                    .as_bytes()
                    .get(header_offset..)
                    .ok_or_else(|| io::Error::other("test header offset is invalid"))?,
            )?;
            let header_start = header_offset
                .checked_add(prefix_length)
                .ok_or_else(|| io::Error::other("test header start overflows"))?;
            let header_end = header_start
                .checked_add(usize::try_from(header_length)?)
                .ok_or_else(|| io::Error::other("test header end overflows"))?;
            if header_end != data_offset {
                return Err(io::Error::other("test header offsets disagree").into());
            }
            let archive_info = WireView::parse(&stream.as_bytes()[header_start..header_end])?;
            let mut rewritten_header = Vec::with_capacity(header_end - header_start);
            let mut changed = false;
            for field in archive_info.fields() {
                if field.number() != 2 || changed {
                    rewritten_header.extend_from_slice(field.raw());
                    continue;
                }
                let message_info = WireView::parse(field.payload())?;
                let is_selected = message_info.fields().any(|nested| {
                    nested.number() == 1 && nested.payload() == [SLIDE_MESSAGE_TYPE as u8]
                });
                if !is_selected {
                    rewritten_header.extend_from_slice(field.raw());
                    continue;
                }
                let mut rewritten_info = Vec::with_capacity(field.payload().len());
                for nested in message_info.fields() {
                    if nested.number() == 1 {
                        push_varint(8, &mut rewritten_info);
                        push_varint(u64::from(SLIDE_MESSAGE_TYPE + 1), &mut rewritten_info);
                    } else {
                        rewritten_info.extend_from_slice(nested.raw());
                    }
                }
                rewritten_header.extend_from_slice(&length_delimited_field(2, &rewritten_info));
                changed = true;
            }
            if !changed {
                return Err(io::Error::other("synthetic selected MessageInfo is missing").into());
            }
            let mut raw = Vec::with_capacity(
                stream
                    .as_bytes()
                    .len()
                    .saturating_add(rewritten_header.len())
                    .saturating_sub(header_end - header_start),
            );
            raw.extend_from_slice(&stream.as_bytes()[..header_offset]);
            litchi_iwa_common::encode_varint_into(&mut raw, u64::try_from(rewritten_header.len())?);
            raw.extend_from_slice(&rewritten_header);
            raw.extend_from_slice(&stream.as_bytes()[data_offset..]);
            Archive::parse(&raw)?;
            SnappyStream::compress(&raw)?
        } else {
            entry.data().to_vec()
        };
        entries.push((entry.name().to_owned(), data));
    }
    Ok(litchi_iwa_archive::package::to_bytes(
        entries
            .iter()
            .map(|(name, data)| (name.as_str(), data.as_slice())),
        Limits::default(),
    )?)
}

fn component_payload(package: &[u8], member: &str, identifier: u64) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(package)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == member)
        .ok_or_else(|| io::Error::other("missing test component"))?;
    let stream = SnappyStream::decompress(entry.data())?;
    let archive = Archive::parse(stream.as_bytes())?;
    let object = archive
        .object(identifier)
        .ok_or_else(|| io::Error::other("missing test object"))?;
    Ok(object
        .messages
        .iter()
        .find(|message| message.type_ == SLIDE_MESSAGE_TYPE)
        .ok_or_else(|| io::Error::other("missing slide message"))?
        .data
        .clone())
}

fn object_header(package: &[u8], member: &str, identifier: u64) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(package)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == member)
        .ok_or_else(|| io::Error::other("missing test component"))?;
    let stream = SnappyStream::decompress(entry.data())?;
    let archive = Archive::parse(stream.as_bytes())?;
    let object = archive
        .object(identifier)
        .ok_or_else(|| io::Error::other("missing test object"))?;
    let header_offset = usize::try_from(object.header_offset)?;
    let data_offset = usize::try_from(object.data_offset)?;
    let (raw_header_length, prefix_length) = litchi_iwa_common::decode_varint_from_bytes(
        stream
            .as_bytes()
            .get(header_offset..)
            .ok_or_else(|| io::Error::other("invalid test header offset"))?,
    )?;
    let header_length = usize::try_from(raw_header_length)?;
    let start = header_offset
        .checked_add(prefix_length)
        .ok_or_else(|| io::Error::other("test header start overflows"))?;
    let end = start
        .checked_add(header_length)
        .ok_or_else(|| io::Error::other("test header end overflows"))?;
    if end != data_offset {
        return Err(io::Error::other("test header offsets disagree").into());
    }
    Ok(stream.as_bytes()[start..end].to_vec())
}

fn raw_node_has_transition(package: &[u8], identifier: u64) -> TestResult<bool> {
    let catalog = Catalog::from_bytes(package)?;
    let component = catalog
        .iter()
        .find(|entry| entry.name() == DOCUMENT_MEMBER)
        .ok_or_else(|| io::Error::other("missing document component"))?;
    let stream = SnappyStream::decompress(component.data())?;
    let archive = Archive::parse(stream.as_bytes())?;
    let payload = archive
        .object(identifier)
        .and_then(|object| object.messages.iter().find(|message| message.type_ == 4))
        .ok_or_else(|| io::Error::other("missing slide-node message"))?
        .data
        .as_slice();
    #[allow(
        deprecated,
        reason = "the native node cache is deliberately verified at the raw fixture boundary"
    )]
    Ok(kn::SlideNodeArchive::decode(payload)?.has_transition)
}

fn normalize_message_length(header: &[u8]) -> TestResult<Vec<u8>> {
    let view = WireView::parse(header)?;
    let mut output = Vec::with_capacity(header.len());
    for field in view.fields() {
        if field.number() != 2 || field.wire_type() != 2 {
            output.extend_from_slice(field.raw());
            continue;
        }
        output.extend_from_slice(field.key());
        output.extend_from_slice(b"<message-info>");
        let message = WireView::parse(field.payload())?;
        for nested in message.fields() {
            output.extend_from_slice(nested.key());
            if nested.number() == 3 && nested.wire_type() == 0 {
                output.extend_from_slice(b"<payload-length>");
            } else {
                output.extend_from_slice(nested.payload());
            }
        }
        output.extend_from_slice(b"</message-info>");
    }
    Ok(output)
}

fn assert_untouched_members(source: &[u8], target: &[u8]) -> TestResult<()> {
    let source_catalog = Catalog::from_bytes(source)?;
    let target_catalog = Catalog::from_bytes(target)?;
    assert_eq!(source_catalog.len(), target_catalog.len());
    let mut changed = 0usize;
    for (left, right) in source_catalog.iter().zip(target_catalog.iter()) {
        assert_eq!(left.name(), right.name());
        if left.data() == right.data() {
            assert_eq!(
                left.raw_record().local_record(),
                right.raw_record().local_record()
            );
        } else {
            changed += 1;
            assert!(
                matches!(left.name(), DOCUMENT_MEMBER | "Index/Slide-4.iwa"),
                "unexpected changed member: {}",
                left.name()
            );
        }
    }
    assert_eq!(changed, 2, "slide edit also reconciles the node cache");
    Ok(())
}

fn assert_transition_projection_accepts(package: &[u8]) -> TestResult<()> {
    let payload = component_payload(package, "Index/Slide-4.iwa", FIRST_SLIDE)?;
    decode_slide_transition(&payload, TransitionDecodeOptions::new(payload.len(), 4)).map_err(
        |error| io::Error::other(format!("transition projection rejected source: {error:?}")),
    )?;
    Ok(())
}

#[test]
fn native_fixture_reads_and_selector_edit_preserves_semantics() -> TestResult<()> {
    let package = Package::open(fixture_path())?;
    let source = package.exact_bytes().to_vec();
    let before = package.slide_transition(SlideSelector::index(0))?;
    let slide_count = package.show()?.slides().len();
    let commit = package
        .edit_slide_transition(SlideSelector::index(0))?
        .set(full_settings()?)?
        .commit()?;
    assert_eq!(commit.patch().position(), Position::new(0));
    assert_eq!(commit.package().show()?.slides().len(), slide_count);
    assert_eq!(commit.patch().before(), before.as_ref());
    assert_eq!(
        commit.package().slide_transition(SlideSelector::index(0))?,
        Some(full_settings()?)
    );
    assert!(commit.diagnostics().changed());
    let target = commit.package().exact_bytes().to_vec();
    let source_catalog = Catalog::from_bytes(&source)?;
    let target_catalog = Catalog::from_bytes(&target)?;
    let changed = source_catalog
        .iter()
        .zip(target_catalog.iter())
        .filter(|(before, after)| before.data() != after.data())
        .count();
    assert_eq!(changed, commit.diagnostics().touched_components());
    let reopened = Package::from_bytes(&target)?;
    assert_eq!(
        reopened.slide_transition(SlideSelector::index(0))?,
        Some(full_settings()?)
    );
    assert_eq!(
        reopened
            .apply_slide_transition(&commit.patch().inverse())?
            .package()
            .exact_bytes(),
        source
    );
    Ok(())
}

#[test]
fn presence_future_values_stale_cache_and_noop_are_exact() -> TestResult<()> {
    let settings = full_settings()?;
    let bytes = package_bytes(
        Some(&settings),
        None,
        ["Alpha", "Beta"],
        [true, false],
        Malformation::None,
    )?;
    let package = Package::from_bytes(&bytes)?;
    assert_transition_projection_accepts(&bytes)?;
    assert_eq!(package.slide_transition("Alpha")?, Some(settings.clone()));
    let source_snapshot = package.exact_bytes();
    let commit = package
        .edit_slide_transition("Alpha")?
        .set(settings.clone())?
        .commit()?;
    assert!(commit.patch().is_noop());
    assert_eq!(commit.package().exact_bytes(), source_snapshot);
    assert_eq!(commit.package().exact_bytes(), bytes);
    assert!(raw_node_has_transition(
        commit.package().exact_bytes(),
        FIRST_NODE
    )?);

    let cleared = package.edit_slide_transition("Alpha")?.clear()?.commit()?;
    assert_eq!(
        cleared.package().slide_transition("Alpha")?,
        Some(native_none_from(&settings)?)
    );
    assert!(cleared.diagnostics().changed());
    assert_eq!(cleared.diagnostics().touched_components(), 2);
    assert!(cleared.diagnostics().full_reparse_performed());
    assert!(!raw_node_has_transition(
        cleared.package().exact_bytes(),
        FIRST_NODE
    )?);
    Ok(())
}

#[test]
fn clear_absent_and_native_none_are_exact_noops_and_edit_debug_is_redacted() -> TestResult<()> {
    let absent_bytes = package_bytes(
        None,
        None,
        ["Alpha", "Beta"],
        [false, false],
        Malformation::None,
    )?;
    let absent = Package::from_bytes(&absent_bytes)?;
    let absent_edit = absent.edit_slide_transition("Alpha")?;
    assert_eq!(absent_edit.settings(), None);
    let absent_commit = absent_edit.clear()?.commit()?;
    assert!(absent_commit.patch().is_noop());
    assert!(!absent_commit.diagnostics().changed());
    assert_eq!(absent_commit.package().exact_bytes(), absent_bytes);

    let native_none = native_none_from(&full_settings()?)?;
    let none_bytes = package_bytes(
        Some(&native_none),
        None,
        ["Alpha", "Beta"],
        [false, false],
        Malformation::None,
    )?;
    let none = Package::from_bytes(&none_bytes)?;
    let none_commit = none.edit_slide_transition("Alpha")?.clear()?.commit()?;
    assert!(none_commit.patch().is_noop());
    assert!(!none_commit.diagnostics().changed());
    assert_eq!(none_commit.package().exact_bytes(), none_bytes);

    let edit = none.edit_slide_transition("Alpha")?.set(full_settings()?)?;
    let debug = format!("{edit:?}");
    for secret in ["future-animation", "first-curve", "second-curve", "alpha"] {
        assert!(!debug.contains(secret), "Edit Debug leaked staged content");
    }
    Ok(())
}

#[test]
fn selector_patch_inverse_unknowns_and_cache_reconciliation_are_transactional() -> TestResult<()> {
    let before = Settings::new();
    let target = full_settings()?;
    let unchanged = Settings::builder()
        .effect(Some(Effect::Dissolve))?
        .duration(Some(1.0))?
        .build()?;
    let bytes = package_bytes(
        Some(&before),
        Some(&unchanged),
        ["Alpha", "Beta"],
        [false, true],
        Malformation::None,
    )?;
    let package = Package::from_bytes(&bytes)?;
    assert_transition_projection_accepts(&bytes)?;
    let first_before = component_payload(&bytes, "Index/Slide-4.iwa", FIRST_SLIDE)?;
    let first_header_before = object_header(&bytes, "Index/Slide-4.iwa", FIRST_SLIDE)?;
    let second_before = component_payload(&bytes, "Index/Slide-6.iwa", SECOND_SLIDE)?;
    let commit = package
        .edit_slide_transition(SlideSelector::name("Alpha"))?
        .set(target.clone())?
        .commit()?;
    assert_eq!(commit.patch().before(), Some(&before));
    assert_eq!(commit.patch().after(), Some(&target));
    assert_ne!(
        commit.patch().source_fingerprint(),
        commit.patch().target_fingerprint()
    );
    assert_eq!(
        normalize_message_length(&object_header(
            commit.package().exact_bytes(),
            "Index/Slide-4.iwa",
            FIRST_SLIDE
        )?)?,
        normalize_message_length(&first_header_before)?,
        "unknown ArchiveInfo fields and noncanonical framing survive the length change"
    );
    assert_eq!(
        commit.package().slide_transition(1usize)?,
        Some(unchanged.clone())
    );
    assert!(raw_node_has_transition(
        commit.package().exact_bytes(),
        FIRST_NODE
    )?);
    assert!(raw_node_has_transition(
        commit.package().exact_bytes(),
        SECOND_NODE
    )?);
    assert_ne!(
        component_payload(
            commit.package().exact_bytes(),
            "Index/Slide-4.iwa",
            FIRST_SLIDE
        )?,
        first_before
    );
    let first_after = component_payload(
        commit.package().exact_bytes(),
        "Index/Slide-4.iwa",
        FIRST_SLIDE,
    )?;
    for marker in [
        b"unknown-slide".as_slice(),
        b"unknown-attributes".as_slice(),
    ] {
        assert!(
            first_after
                .windows(marker.len())
                .any(|window| window == marker),
            "nested extension marker was not preserved"
        );
    }
    assert_eq!(
        component_payload(
            commit.package().exact_bytes(),
            "Index/Slide-6.iwa",
            SECOND_SLIDE
        )?,
        second_before
    );
    assert_untouched_members(&bytes, commit.package().exact_bytes())?;
    let applied = package.apply_slide_transition(commit.patch())?;
    assert_eq!(
        applied.package().exact_bytes(),
        commit.package().exact_bytes()
    );
    let inverse = commit.patch().inverse();
    assert_eq!(inverse.inverse(), commit.patch().clone());
    assert_eq!(
        commit
            .package()
            .apply_slide_transition(&inverse)?
            .package()
            .exact_bytes(),
        bytes
    );
    let unrelated = Package::from_bytes(&package_bytes(
        Some(&before),
        Some(&unchanged),
        ["Other", "Beta"],
        [false, true],
        Malformation::None,
    )?)?;
    assert!(matches!(
        unrelated.apply_slide_transition(commit.patch()),
        Err(Error::PatchConflict)
    ));
    let debug = format!("{:?}", commit.patch());
    assert!(!debug.contains(PRIVATE_MARKER));
    assert!(!debug.contains("Index/"));
    assert!(!debug.contains("fingerprint"));
    Ok(())
}

#[test]
fn co_located_slide_and_node_are_rewritten_once() -> TestResult<()> {
    let before = Settings::new();
    let bytes = co_located_package_bytes(Some(&before), None)?;
    let package = Package::from_bytes(&bytes)?;
    let commit = package
        .edit_slide_transition("Alpha")?
        .set(full_settings()?)?
        .commit()?;
    assert_eq!(commit.diagnostics().touched_components(), 1);
    assert!(commit.diagnostics().changed());
    assert_eq!(
        commit.package().slide_transition("Alpha")?,
        Some(full_settings()?)
    );
    assert!(raw_node_has_transition(
        commit.package().exact_bytes(),
        FIRST_NODE
    )?);
    Ok(())
}

#[test]
fn selector_malformed_legacy_and_limits_fail_without_publication() -> TestResult<()> {
    let empty = Settings::new();
    let duplicate_names = Package::from_bytes(&package_bytes(
        Some(&empty),
        None,
        ["Same", "Same"],
        [false, false],
        Malformation::None,
    )?)?;
    assert!(matches!(
        duplicate_names.edit_slide_transition("Missing"),
        Err(Error::SlideNameNotFound)
    ));
    assert!(
        matches!(duplicate_names.edit_slide_transition(99usize), Err(Error::SlidePositionNotFound { position }) if position == Position::new(99))
    );
    assert!(matches!(
        duplicate_names.edit_slide_transition("Same"),
        Err(Error::AmbiguousSelector)
    ));

    let flat = package_bytes(
        Some(&empty),
        None,
        ["Alpha", "Beta"],
        [false, false],
        Malformation::None,
    )?;
    let no_modern = Package::from_bytes(&flat)?;
    let no_modern_edit = no_modern.edit_slide_transition("Beta")?;
    assert_eq!(no_modern_edit.settings(), None);
    assert!(matches!(
        no_modern_edit.set(full_settings()?),
        Err(Error::UnsupportedTransition)
    ));

    for malformed in [
        Malformation::DuplicateEffect,
        Malformation::WrongEffectWire,
        Malformation::NonCanonicalDirection,
    ] {
        let bytes = package_bytes(
            Some(&full_settings()?),
            None,
            ["Alpha", "Beta"],
            [true, false],
            malformed,
        )?;
        let package = Package::from_bytes(&bytes)?;
        assert!(matches!(
            package.slide_transition("Alpha"),
            Err(Error::InvalidSource)
        ));
        assert_eq!(package.exact_bytes(), bytes);
    }

    let stale = Package::from_bytes(&package_bytes(
        Some(&full_settings()?),
        None,
        ["Alpha", "Beta"],
        [false, false],
        Malformation::Canonical,
    )?)?;
    assert!(matches!(
        stale.edit_slide_transition("Alpha"),
        Err(Error::InvalidSource)
    ));

    for malformed in [
        NodeMalformation::DuplicateFlag,
        NodeMalformation::WrongFlagWire,
        NodeMalformation::NonCanonicalFlag,
    ] {
        let node_bytes = malformed_node_package(&flat, malformed)?;
        let node_package = Package::from_bytes(&node_bytes)?;
        assert!(matches!(
            node_package.edit_slide_transition("Alpha"),
            Err(Error::InvalidSource)
        ));
        assert_eq!(node_package.exact_bytes(), node_bytes);
    }

    let legacy = Package::from_bytes(&legacy_package_bytes(&flat)?)?;
    let noop = legacy.edit_slide_transition("Alpha")?.set(empty.clone())?;
    assert!(noop.commit()?.patch().is_noop());
    let changed = legacy
        .edit_slide_transition("Alpha")?
        .set(full_settings()?)?;
    assert!(matches!(changed.commit(), Err(Error::UnsupportedSource)));

    let unrestricted = Package::from_bytes(&flat)?;
    let target_len = unrestricted
        .edit_slide_transition("Alpha")?
        .set(full_settings()?)?
        .commit()?
        .package()
        .exact_bytes()
        .len();
    let limits = Limits::new(
        u64::try_from(target_len - 1)?,
        8,
        1024 * 1024,
        1024 * 1024,
        1024 * 1024,
    )?;
    let limited = Package::from_bytes_with_limits(&flat, limits)?;
    let limited_edit = limited
        .edit_slide_transition("Alpha")?
        .set(full_settings()?)?;
    assert!(matches!(
        limited_edit.commit(),
        Err(Error::LimitExceeded { .. })
    ));
    Ok(())
}

#[test]
fn mixed_legacy_ownership_metadata_and_framing_fail_closed() -> TestResult<()> {
    let source = package_bytes(
        Some(&full_settings()?),
        Some(&Settings::new()),
        ["Alpha", "Beta"],
        [true, false],
        Malformation::None,
    )?;
    let mut changed_settings = full_settings()?;
    changed_settings.set_duration(Some(9.0))?;

    let mixed_bytes = mixed_modern_legacy_package(&source)?;
    let mixed_payload = component_payload(&mixed_bytes, "Index/Slide-4.iwa", FIRST_SLIDE)?;
    assert!(
        decode_slide_transition(
            &mixed_payload,
            TransitionDecodeOptions::new(mixed_payload.len(), 4)
        )?
        .settings
        .has_legacy_database_fields
    );
    let mixed = Package::from_bytes(&mixed_bytes)?;
    let mixed_edit = mixed
        .edit_slide_transition("Alpha")?
        .set(changed_settings.clone())?;
    assert!(matches!(mixed_edit.commit(), Err(Error::InvalidSource)));

    for malformed in [
        duplicate_rooted_node_package(&source)?,
        shared_slide_package(&source)?,
    ] {
        let package = Package::from_bytes(&malformed)?;
        match package.edit_slide_transition(0usize) {
            Ok(edit) => assert!(matches!(
                edit.set(changed_settings.clone())?.commit(),
                Err(Error::InvalidSource)
            )),
            Err(Error::InvalidSource) => {},
            Err(other) => return Err(format!("unexpected ownership error: {other:?}").into()),
        }
    }

    let metadata_variants = [
        (
            "MessageInfo type",
            wrong_message_info_type_package(&source)?,
        ),
        (
            "ArchiveInfo should_merge",
            metadata_or_framing_package(&source, |object| {
                object.archive_info.should_merge = Some(true);
                Ok(())
            })?,
        ),
        (
            "MessageInfo base_message_index",
            metadata_or_framing_package(&source, |object| {
                let selected = object
                    .archive_info
                    .message_infos
                    .iter_mut()
                    .find(|info| info.type_ == SLIDE_MESSAGE_TYPE)
                    .ok_or_else(|| io::Error::other("synthetic slide metadata is missing"))?;
                selected.base_message_index = Some(0);
                Ok(())
            })?,
        ),
        (
            "MessageInfo diff_merge_version",
            metadata_or_framing_package(&source, |object| {
                let selected = object
                    .archive_info
                    .message_infos
                    .iter_mut()
                    .find(|info| info.type_ == SLIDE_MESSAGE_TYPE)
                    .ok_or_else(|| io::Error::other("synthetic slide metadata is missing"))?;
                selected.diff_merge_version.push(1);
                Ok(())
            })?,
        ),
        (
            "noncanonical object length prefix",
            noncanonical_object_length_prefix(&source)?,
        ),
    ];
    for (name, malformed) in metadata_variants {
        let package = Package::from_bytes(&malformed)?;
        match package.edit_slide_transition("Alpha") {
            Ok(edit) => match edit.set(changed_settings.clone())?.commit() {
                Err(Error::InvalidSource) => {},
                Ok(_) => return Err(format!("{name} malformed source was published").into()),
                Err(other) => return Err(format!("unexpected {name} error: {other:?}").into()),
            },
            Err(Error::InvalidSource) => {},
            Err(other) => return Err(format!("unexpected metadata error: {other:?}").into()),
        }
    }
    Ok(())
}

#[test]
fn public_transaction_values_are_send_sync() {
    fn assert_send_sync<T: Send + Sync>() {}
    assert_send_sync::<Edit<'static>>();
    assert_send_sync::<Commit>();
    assert_send_sync::<Patch>();
    assert_send_sync::<Diagnostics>();
    assert_send_sync::<Error>();
    assert_send_sync::<LimitKind>();
    assert_send_sync::<Arc<[u8]>>();
}
