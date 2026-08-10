use super::*;
use crate::archive::{Archive, ArchiveObject, FieldInfo, FieldPath};
use crate::media::MediaAssetId;
use crate::package_metadata::{PACKAGE_METADATA_ENTRY, PACKAGE_METADATA_MESSAGE_TYPE};
use crate::protobuf::tsp::{ComponentInfo, ObjectUuidMapEntry, PackageMetadata, Reference, Uuid};
use crate::protobuf::tswp::StorageArchive;
use crate::shapes::{DrawablePoint, DrawableSize};
use litchi_iwa_common::media::playback::{MediaLoopMode, MediaPlaybackSettings, MediaVolume};
use litchi_iwa_common::shape::fill::{Opacity, StopMidpoint, StopPosition};
use litchi_keynote::slide::media::MovieKind;
use litchi_keynote::soundtrack::{Mode as SoundtrackMode, Settings as SoundtrackSettings};
use litchi_keynote::transition::{
    Acceleration, AnimationParameters, CustomParameters, Direction, Effect, MosaicType, Settings,
    TextDelivery, TimingCurveSlot,
};
use litchi_keynote::{Package as FocusedKeynotePackage, SlideSelector};
use std::time::Duration;

const TEST_SLIDE_MESSAGE_TYPE: u32 = 5;
const TEST_SLIDE_NODE_MESSAGE_TYPE: u32 = 4;
const TEST_THEME_MESSAGE_TYPE: u32 = 10;
const TEST_PLACEHOLDER_MESSAGE_TYPE: u32 = 7;
const TEST_IMAGE_MESSAGE_TYPE: u32 = 3_005;
const TEST_MOVIE_MESSAGE_TYPE: u32 = 3_007;
const TEST_LIVE_VIDEO_INFO_FIELD: u32 = 100;
const TEST_LIVE_VIDEO_INFO_PAYLOAD: &[u8] = &[0x08, 0x01];
const TEST_TITLE_PLACEHOLDER_FIELD: u32 = 5;
const TEST_SLIDE_NUMBER_PLACEHOLDER_FIELD: u32 = 20;
const TEST_SLIDE_OWNED_DRAWABLES_FIELD: u32 = 7;
const TEST_SLIDE_DRAWABLES_Z_ORDER_FIELD: u32 = 42;
const TEST_SLIDE_NUMBER_PLACEHOLDER_ID: u64 = 70;
const TEST_SHOW_MODE_FIELD: u32 = 9;
const TEST_SHOW_SOUNDTRACK_FIELD: u32 = 17;
const TEST_SOUNDTRACK_ID: u64 = 80;
const TEST_SOUNDTRACK_MESSAGE_TYPE: u32 = 21;
const TEST_SOUNDTRACK_MODE_FIELD: u32 = 2;
const TEST_SOUNDTRACK_PLAY_ONCE_MODE: i32 = 0;
const TEST_SOUNDTRACK_LOOP_MODE: i32 = 1;
const TEST_SOUNDTRACK_MEDIA_IDS: [u64; 2] = [91, 92];
const TEST_MOVIE_VIDEO: &[u8] = b"\0\0\0\x18ftypqt  movie-data";
const TEST_MOVIE_VIDEO_REPLACEMENT: &[u8] = b"\0\0\0\x18ftypqt  replacement";
const TEST_MOVIE_POSTER: &[u8] = b"\x89PNG\r\n\x1a\nmovie-poster";
const TEST_MOVIE_POSTER_REPLACEMENT: &[u8] = b"\x89PNG\r\n\x1a\nreplacement";

#[test]
fn slide_background_crud_inherits_and_culls_native_variations() {
    let mut editor = KeynoteEditor::from_package(test_package_with_slide_background()).unwrap();
    let white = Background::Solid(Rgba::new(1.0, 1.0, 1.0, 1.0, RgbColorSpace::Srgb).unwrap());
    assert_eq!(editor.slide_background(0).unwrap(), white);
    assert_eq!(editor.slide_background(1).unwrap(), white);
    assert_eq!(editor.slide_background_override(0).unwrap(), None);

    let red = Background::Solid(Rgba::new(0.9, 0.2, 0.1, 0.75, RgbColorSpace::DisplayP3).unwrap());
    editor.set_slide_background(0, red.clone()).unwrap();
    assert_eq!(editor.slide_background(0).unwrap(), red);
    assert_eq!(
        editor.slide_background_override(0).unwrap(),
        Some(red.clone())
    );
    assert_eq!(editor.slide_background(1).unwrap(), white);
    let graph = ObjectGraph::read(editor.package()).unwrap();
    let slide: kn::SlideArchive = graph.decode_type(4, 5, "KN.SlideArchive").unwrap();
    let first_variation = slide.style.identifier;
    let style: kn::SlideStyleArchive = graph
        .decode_type(first_variation, 9, "KN.SlideStyleArchive")
        .unwrap();
    assert_eq!(style.super_.parent.unwrap().identifier, 40);
    assert_eq!(style.super_.is_variation, Some(true));
    assert_eq!(style.override_count, Some(1));

    let no_op = editor.to_bytes().unwrap();
    editor.set_slide_background(0, red).unwrap();
    assert_eq!(editor.to_bytes().unwrap(), no_op);

    let green = Background::Solid(Rgba::new(0.1, 0.8, 0.3, 1.0, RgbColorSpace::Srgb).unwrap());
    editor.set_slide_background(0, green.clone()).unwrap();
    assert_eq!(editor.slide_background(0).unwrap(), green);
    let graph = ObjectGraph::read(editor.package()).unwrap();
    assert!(!graph.objects.contains_key(&first_variation));
    let slide: kn::SlideArchive = graph.decode_type(4, 5, "KN.SlideArchive").unwrap();
    let second_variation = slide.style.identifier;
    let stylesheet: tss::StylesheetArchive =
        graph.decode_type(41, 401, "TSS.StylesheetArchive").unwrap();
    assert_eq!(
        stylesheet
            .styles
            .iter()
            .map(|reference| reference.identifier)
            .collect::<Vec<_>>(),
        vec![40, second_variation]
    );
    assert_eq!(
        stylesheet.parent_to_children_style_map[0]
            .children
            .iter()
            .map(|reference| reference.identifier)
            .collect::<Vec<_>>(),
        vec![second_variation]
    );

    editor.set_slide_background(0, Background::None).unwrap();
    assert_eq!(editor.slide_background(0).unwrap(), Background::None);
    let graph = ObjectGraph::read(editor.package()).unwrap();
    assert!(!graph.objects.contains_key(&second_variation));
    let slide: kn::SlideArchive = graph.decode_type(4, 5, "KN.SlideArchive").unwrap();
    let no_fill_variation = slide.style.identifier;
    let raw = graph
        .message_data_type(no_fill_variation, 9, "KN.SlideStyleArchive")
        .unwrap();
    let properties = required_length_delimited_payload(raw, 11, "slide style").unwrap();
    assert!(
        required_length_delimited_payload(properties, 1, "slide fill")
            .unwrap()
            .is_empty()
    );

    assert!(editor.reset_slide_background(0).unwrap());
    assert_eq!(editor.slide_background(0).unwrap(), white);
    assert_eq!(editor.slide_background_override(0).unwrap(), None);
    let graph = ObjectGraph::read(editor.package()).unwrap();
    assert!(!graph.objects.contains_key(&no_fill_variation));
    let stylesheet: tss::StylesheetArchive =
        graph.decode_type(41, 401, "TSS.StylesheetArchive").unwrap();
    assert_eq!(stylesheet.styles, vec![reference(40)]);
    assert!(stylesheet.parent_to_children_style_map.is_empty());
    let no_op = editor.to_bytes().unwrap();
    assert!(!editor.reset_slide_background(0).unwrap());
    assert_eq!(editor.to_bytes().unwrap(), no_op);
}

#[test]
fn slide_background_gradient_crud_round_trips_native_semantics() {
    let mut editor = KeynoteEditor::from_package(test_package_with_slide_background()).unwrap();
    let cyan = Rgba::new(
        0.337_592_3,
        0.757_728_4,
        0.999_999_94,
        1.0,
        RgbColorSpace::Srgb,
    )
    .unwrap();
    let blue = Rgba::new(
        3.419_328e-7,
        0.462_459_27,
        0.729_136_77,
        1.0,
        RgbColorSpace::Srgb,
    )
    .unwrap();
    let simple = Gradient::linear(cyan, blue, Angle::from_degrees(270.0).unwrap());
    editor
        .set_slide_background(0, Background::Gradient(simple.clone()))
        .unwrap();
    assert_eq!(
        editor.slide_background(0).unwrap(),
        Background::Gradient(simple.clone())
    );
    assert_eq!(
        editor.slide_background_override(0).unwrap(),
        Some(Background::Gradient(simple))
    );

    let graph = ObjectGraph::read(editor.package()).unwrap();
    let slide: kn::SlideArchive = graph.decode_type(4, 5, "KN.SlideArchive").unwrap();
    let simple_style_id = slide.style.identifier;
    let raw = graph
        .message_data_type(simple_style_id, 9, "KN.SlideStyleArchive")
        .unwrap();
    let properties = required_length_delimited_payload(raw, 11, "slide style").unwrap();
    let fill_payload = required_length_delimited_payload(properties, 1, "slide fill").unwrap();
    let fill = tsd::FillArchive::decode(fill_payload).unwrap();
    let native = fill.gradient.unwrap();
    assert_eq!(
        native.r#type,
        Some(tsd::gradient_archive::GradientType::Linear as i32)
    );
    assert_eq!(native.advanced_gradient, Some(false));
    assert_eq!(native.opacity, Some(1.0));
    assert_eq!(
        native.anglegradient.unwrap().gradientangle,
        Some(3.0 * std::f32::consts::FRAC_PI_2)
    );

    let advanced = Gradient::advanced(
        Kind::Radial,
        vec![
            Stop::new(
                cyan,
                StopPosition::new(0.0).unwrap(),
                StopMidpoint::new(0.35).unwrap(),
            ),
            Stop::new(
                Rgba::new(0.4, 0.2, 0.8, 0.9, RgbColorSpace::DisplayP3).unwrap(),
                StopPosition::new(0.4).unwrap(),
                StopMidpoint::new(0.6).unwrap(),
            ),
            Stop::new(blue, StopPosition::new(1.0).unwrap(), StopMidpoint::CENTER),
        ],
        Opacity::new(0.75).unwrap(),
        Angle::from_degrees(315.0).unwrap(),
    )
    .unwrap();
    editor
        .set_slide_background(0, Background::Gradient(advanced.clone()))
        .unwrap();
    assert_eq!(
        editor.slide_background(0).unwrap(),
        Background::Gradient(advanced.clone())
    );
    let no_op = editor.to_bytes().unwrap();
    editor
        .set_slide_background(0, Background::Gradient(advanced))
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), no_op);
    let graph = ObjectGraph::read(editor.package()).unwrap();
    assert!(!graph.objects.contains_key(&simple_style_id));

    assert!(editor.reset_slide_background(0).unwrap());
    assert_eq!(editor.slide_background_override(0).unwrap(), None);
}

#[test]
fn slide_background_gradient_validation_and_unknown_wire_are_lossless() {
    assert!(Angle::from_degrees(-1.0).is_err());
    assert!(Angle::from_degrees(360.0).is_err());
    assert!(Angle::from_degrees(f32::NAN).is_err());

    let black = Rgba::new(0.0, 0.0, 0.0, 1.0, RgbColorSpace::Srgb).unwrap();
    let white = Rgba::new(1.0, 1.0, 1.0, 1.0, RgbColorSpace::Srgb).unwrap();
    let angle = Angle::from_degrees(90.0).unwrap();
    let gradient = Gradient::linear(black, white, angle);
    assert!(StopPosition::new(f32::NAN).is_err());
    assert!(Opacity::new(1.1).is_err());
    let fill = slide_background_gradient_wire::gradient_to_fill(&gradient);
    assert_eq!(
        slide_background::background_from_fill(&fill).unwrap(),
        Background::Gradient(gradient.clone())
    );
    let future = transform_length_delimited_field(&fill, 2, |payload| {
        let mut payload = payload.to_vec();
        append_unknown_varint(&mut payload, 99, 73);
        Ok(payload)
    })
    .unwrap();
    assert_eq!(
        slide_background::background_from_fill(&future).unwrap(),
        Background::Opaque(Opaque::from_slice(&future).unwrap())
    );

    let mut editor = KeynoteEditor::from_package(test_package_with_slide_background()).unwrap();
    editor
        .set_slide_background(0, Background::Opaque(Opaque::from_slice(&future).unwrap()))
        .unwrap();
    assert_eq!(
        editor.slide_background(0).unwrap(),
        Background::Opaque(Opaque::from_slice(&future).unwrap())
    );
    let before = editor.to_bytes().unwrap();
    assert!(
        Gradient::from_parts(
            Kind::Radial,
            gradient.stops().to_vec(),
            litchi_iwa_common::shape::fill::Opacity::OPAQUE,
            false,
            angle,
        )
        .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before);

    let unordered = Gradient::advanced(
        Kind::Linear,
        vec![
            Stop::new(black, StopPosition::new(0.8).unwrap(), StopMidpoint::CENTER),
            Stop::new(white, StopPosition::new(0.2).unwrap(), StopMidpoint::CENTER),
        ],
        Opacity::OPAQUE,
        angle,
    );
    assert!(unordered.is_err());
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn slide_background_reset_preserves_combined_and_unknown_style_properties() {
    let mut editor = KeynoteEditor::from_package(test_package_with_slide_background()).unwrap();
    let red = Background::Solid(Rgba::new(0.9, 0.2, 0.1, 1.0, RgbColorSpace::Srgb).unwrap());
    editor.set_slide_background(0, red).unwrap();
    let mut package = editor.into_package();
    let graph = ObjectGraph::read(&package).unwrap();
    let slide: kn::SlideArchive = graph.decode_type(4, 5, "KN.SlideArchive").unwrap();
    let style_id = slide.style.identifier;
    package
        .update_archive("Index/DocumentStylesheet.iwa", |archive| {
            let object = archive.object_mut(style_id).unwrap();
            let mut style = kn::SlideStyleArchive::decode(object.messages[0].data.as_slice())?;
            style.override_count = Some(2);
            style.slide_properties.as_mut().unwrap().transition_null = Some(true);
            let data =
                transform_length_delimited_field(&style.encode_to_vec(), 11, |properties| {
                    let mut properties = properties.to_vec();
                    append_unknown_varint(&mut properties, 99, 73);
                    Ok(properties)
                })?;
            object.replace_message(0, RawMessage { type_: 9, data })?;
            let info = &mut object.archive_info.message_infos[0];
            info.object_references.extend([88, 89]);
            info.data_references.extend([98, 99]);
            info.field_infos.extend([
                FieldInfo {
                    path: FieldPath { path: vec![11, 1] },
                    object_references: vec![88, 89],
                    data_references: vec![98, 99],
                    ..Default::default()
                },
                FieldInfo {
                    path: FieldPath { path: vec![11, 3] },
                    object_references: vec![89],
                    data_references: vec![99],
                    ..Default::default()
                },
            ]);
            Ok(())
        })
        .unwrap();
    let mut editor = KeynoteEditor::from_package(package).unwrap();

    assert!(editor.reset_slide_background(0).unwrap());
    assert_eq!(editor.slide_background_override(0).unwrap(), None);
    assert_eq!(
        editor.slide_background(0).unwrap(),
        Background::Solid(Rgba::new(1.0, 1.0, 1.0, 1.0, RgbColorSpace::Srgb).unwrap())
    );
    let graph = ObjectGraph::read(editor.package()).unwrap();
    let slide: kn::SlideArchive = graph.decode_type(4, 5, "KN.SlideArchive").unwrap();
    assert_eq!(slide.style.identifier, style_id);
    let style: kn::SlideStyleArchive = graph
        .decode_type(style_id, 9, "KN.SlideStyleArchive")
        .unwrap();
    assert_eq!(style.override_count, Some(1));
    let properties = style.slide_properties.unwrap();
    assert_eq!(properties.fill, None);
    assert_eq!(properties.transition_null, Some(true));
    let raw = graph
        .message_data_type(style_id, 9, "KN.SlideStyleArchive")
        .unwrap();
    let properties = required_length_delimited_payload(raw, 11, "slide style").unwrap();
    assert!(properties.windows(3).any(|bytes| bytes == [0x98, 0x06, 73]));
    let archive = editor
        .package()
        .archive("Index/DocumentStylesheet.iwa")
        .unwrap();
    let info = &archive.object(style_id).unwrap().archive_info.message_infos[0];
    assert_eq!(info.object_references, [40, 89]);
    assert_eq!(info.data_references, [99]);
    assert_eq!(
        info.field_infos
            .iter()
            .map(|field| field.path.path.as_slice())
            .collect::<Vec<_>>(),
        vec![&[11, 3][..]]
    );
}

#[test]
fn slide_background_reset_preserves_shared_background_variations() {
    let mut editor = KeynoteEditor::from_package(test_package_with_slide_background()).unwrap();
    let red = Background::Solid(Rgba::new(0.9, 0.2, 0.1, 1.0, RgbColorSpace::Srgb).unwrap());
    editor.set_slide_background(0, red.clone()).unwrap();
    let mut package = editor.into_package();
    let graph = ObjectGraph::read(&package).unwrap();
    let first: kn::SlideArchive = graph.decode_type(4, 5, "KN.SlideArchive").unwrap();
    let shared_style_id = first.style.identifier;
    package
        .update_archive("Index/Slide-10.iwa", |archive| {
            let object = archive.object_mut(10).unwrap();
            let data = patch_length_delimited_field(
                &object.messages[0].data,
                1,
                true,
                Some(&reference(shared_style_id).encode_to_vec()),
            )?;
            object.replace_message(0, RawMessage { type_: 5, data })?;
            let references = &mut object.archive_info.message_infos[0].object_references;
            *references
                .iter_mut()
                .find(|reference| **reference == 40)
                .unwrap() = shared_style_id;
            Ok(())
        })
        .unwrap();
    let mut editor = KeynoteEditor::from_package(package).unwrap();

    assert!(editor.reset_slide_background(0).unwrap());
    assert_eq!(editor.slide_background_override(0).unwrap(), None);
    assert_eq!(editor.slide_background(1).unwrap(), red);
    assert_eq!(
        editor.slide_background_override(1).unwrap(),
        Some(red.clone())
    );
    let graph = ObjectGraph::read(editor.package()).unwrap();
    assert!(graph.objects.contains_key(&shared_style_id));
    let first: kn::SlideArchive = graph.decode_type(4, 5, "KN.SlideArchive").unwrap();
    let second: kn::SlideArchive = graph.decode_type(10, 5, "KN.SlideArchive").unwrap();
    assert_eq!(first.style.identifier, 40);
    assert_eq!(second.style.identifier, shared_style_id);
}

#[test]
fn slide_background_reset_copy_on_writes_shared_combined_variations() {
    let mut editor = KeynoteEditor::from_package(test_package_with_slide_background()).unwrap();
    let red = Background::Solid(Rgba::new(0.9, 0.2, 0.1, 1.0, RgbColorSpace::Srgb).unwrap());
    editor.set_slide_background(0, red.clone()).unwrap();
    let mut package = editor.into_package();
    let graph = ObjectGraph::read(&package).unwrap();
    let first: kn::SlideArchive = graph.decode_type(4, 5, "KN.SlideArchive").unwrap();
    let shared_style_id = first.style.identifier;
    package
        .update_archive("Index/DocumentStylesheet.iwa", |archive| {
            let object = archive.object_mut(shared_style_id).unwrap();
            let mut style = kn::SlideStyleArchive::decode(object.messages[0].data.as_slice())?;
            style.override_count = Some(2);
            style.slide_properties.as_mut().unwrap().transition_null = Some(true);
            object.replace_message(
                0,
                RawMessage {
                    type_: 9,
                    data: style.encode_to_vec(),
                },
            )?;
            Ok(())
        })
        .unwrap();
    package
        .update_archive("Index/Slide-10.iwa", |archive| {
            let object = archive.object_mut(10).unwrap();
            let data = patch_length_delimited_field(
                &object.messages[0].data,
                1,
                true,
                Some(&reference(shared_style_id).encode_to_vec()),
            )?;
            object.replace_message(0, RawMessage { type_: 5, data })?;
            let references = &mut object.archive_info.message_infos[0].object_references;
            *references
                .iter_mut()
                .find(|reference| **reference == 40)
                .unwrap() = shared_style_id;
            Ok(())
        })
        .unwrap();
    let mut editor = KeynoteEditor::from_package(package).unwrap();

    assert!(editor.reset_slide_background(0).unwrap());
    assert_eq!(editor.slide_background_override(0).unwrap(), None);
    assert_eq!(editor.slide_background(1).unwrap(), red);
    let graph = ObjectGraph::read(editor.package()).unwrap();
    let first: kn::SlideArchive = graph.decode_type(4, 5, "KN.SlideArchive").unwrap();
    let second: kn::SlideArchive = graph.decode_type(10, 5, "KN.SlideArchive").unwrap();
    assert_ne!(first.style.identifier, shared_style_id);
    assert_eq!(second.style.identifier, shared_style_id);
    let replacement: kn::SlideStyleArchive = graph
        .decode_type(first.style.identifier, 9, "KN.SlideStyleArchive")
        .unwrap();
    assert_eq!(replacement.super_.parent.unwrap().identifier, 40);
    assert_eq!(replacement.override_count, Some(1));
    let properties = replacement.slide_properties.unwrap();
    assert_eq!(properties.fill, None);
    assert_eq!(properties.transition_null, Some(true));
    let shared: kn::SlideStyleArchive = graph
        .decode_type(shared_style_id, 9, "KN.SlideStyleArchive")
        .unwrap();
    assert!(shared.slide_properties.unwrap().fill.is_some());
}

#[test]
fn slide_background_reset_and_update_reject_or_preserve_future_style_wire() {
    let red = Background::Solid(Rgba::new(0.9, 0.2, 0.1, 1.0, RgbColorSpace::Srgb).unwrap());
    let green = Background::Solid(Rgba::new(0.1, 0.8, 0.3, 1.0, RgbColorSpace::Srgb).unwrap());
    let mut editor = KeynoteEditor::from_package(test_package_with_slide_background()).unwrap();
    editor.set_slide_background(0, red.clone()).unwrap();
    let mut package = editor.into_package();
    let graph = ObjectGraph::read(&package).unwrap();
    let slide: kn::SlideArchive = graph.decode_type(4, 5, "KN.SlideArchive").unwrap();
    let future_style_id = slide.style.identifier;
    package
        .update_archive("Index/DocumentStylesheet.iwa", |archive| {
            let object = archive.object_mut(future_style_id).unwrap();
            let mut data = object.messages[0].data.clone();
            append_unknown_varint(&mut data, 99, 73);
            object.replace_message(0, RawMessage { type_: 9, data })?;
            Ok(())
        })
        .unwrap();
    let mut editor = KeynoteEditor::from_package(package).unwrap();
    editor.set_slide_background(0, green.clone()).unwrap();
    let graph = ObjectGraph::read(editor.package()).unwrap();
    assert!(graph.objects.contains_key(&future_style_id));
    let slide: kn::SlideArchive = graph.decode_type(4, 5, "KN.SlideArchive").unwrap();
    let replacement: kn::SlideStyleArchive = graph
        .decode_type(slide.style.identifier, 9, "KN.SlideStyleArchive")
        .unwrap();
    assert_eq!(
        replacement.super_.parent.unwrap().identifier,
        future_style_id
    );
    assert_eq!(editor.slide_background(0).unwrap(), green);

    let mut editor = KeynoteEditor::from_package(test_package_with_slide_background()).unwrap();
    editor.set_slide_background(0, red).unwrap();
    let mut package = editor.into_package();
    let graph = ObjectGraph::read(&package).unwrap();
    let slide: kn::SlideArchive = graph.decode_type(4, 5, "KN.SlideArchive").unwrap();
    package
        .update_archive("Index/DocumentStylesheet.iwa", |archive| {
            let object = archive.object_mut(slide.style.identifier).unwrap();
            let data = patch_varint_field(&object.messages[0].data, 10, true, None)?;
            object.replace_message(0, RawMessage { type_: 9, data })?;
            Ok(())
        })
        .unwrap();
    let mut malformed = KeynoteEditor::from_package(package).unwrap();
    let before = malformed.to_bytes().unwrap();
    assert!(malformed.reset_slide_background(0).is_err());
    assert_eq!(malformed.to_bytes().unwrap(), before);
}

#[test]
fn slide_background_preserves_opaque_fills_and_unknown_solid_fields() {
    let mut future_fill = Vec::new();
    append_unknown_varint(&mut future_fill, 100, 73);
    assert_eq!(
        slide_background::background_from_fill(&future_fill).unwrap(),
        Background::Opaque(Opaque::from_slice(&future_fill).unwrap())
    );

    let mut editor = KeynoteEditor::from_package(test_package_with_slide_background()).unwrap();
    let opaque = tsd::FillArchive {
        gradient: Some(tsd::GradientArchive::default()),
        ..Default::default()
    }
    .encode_to_vec();
    let background = Background::Opaque(Opaque::from_slice(&opaque).unwrap());
    editor.set_slide_background(0, background.clone()).unwrap();
    assert_eq!(editor.slide_background(0).unwrap(), background);
    let graph = ObjectGraph::read(editor.package()).unwrap();
    let slide: kn::SlideArchive = graph.decode_type(4, 5, "KN.SlideArchive").unwrap();
    let raw = graph
        .message_data_type(slide.style.identifier, 9, "KN.SlideStyleArchive")
        .unwrap();
    let properties = required_length_delimited_payload(raw, 11, "slide style").unwrap();
    assert_eq!(
        required_length_delimited_payload(properties, 1, "slide fill").unwrap(),
        opaque
    );

    let mut editor = KeynoteEditor::from_package(test_package_with_slide_background()).unwrap();
    let blue = Background::Solid(Rgba::new(0.2, 0.4, 0.8, 1.0, RgbColorSpace::Srgb).unwrap());
    editor.set_slide_background(0, blue).unwrap();
    let graph = ObjectGraph::read(editor.package()).unwrap();
    let slide: kn::SlideArchive = graph.decode_type(4, 5, "KN.SlideArchive").unwrap();
    let raw = graph
        .message_data_type(slide.style.identifier, 9, "KN.SlideStyleArchive")
        .unwrap();
    let properties = required_length_delimited_payload(raw, 11, "slide style").unwrap();
    let fill = required_length_delimited_payload(properties, 1, "slide fill").unwrap();
    assert!(fill.windows(3).any(|bytes| bytes == [0xa0, 0x06, 73]));
}

#[test]
fn slide_background_rejects_invalid_inputs_transactionally() {
    let mut editor = KeynoteEditor::from_package(test_package_with_slide_background()).unwrap();
    let before = editor.to_bytes().unwrap();
    assert!(Rgba::new(f32::NAN, 0.0, 0.0, 1.0, RgbColorSpace::Srgb).is_err());
    assert_eq!(editor.to_bytes().unwrap(), before);
    assert!(
        editor
            .set_slide_background(
                0,
                Background::Opaque(Opaque::from_slice(&[0x0a, 0xff]).unwrap()),
            )
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before);
    assert!(Opaque::from_slice(&[]).is_err());
    assert_eq!(editor.to_bytes().unwrap(), before);
    assert!(editor.slide_background(2).is_err());
}

#[test]
fn reads_placeholder_text_and_edits_navigator_name_by_slide_index() {
    let mut editor = KeynoteEditor::from_package(test_package()).unwrap();
    let slides = editor.slides().unwrap();
    assert_eq!(slides.len(), 2);
    assert_eq!(slides[0].title.as_deref(), Some("Old title"));
    assert_eq!(slides[0].body.as_deref(), Some("Old body 🚀"));
    assert_eq!(slides[0].notes.as_deref(), Some("Speaker 🚀"));

    editor.set_slide_name(0, Some("Agenda 🚀")).unwrap();
    let slides = editor.slides().unwrap();
    assert_eq!(slides[0].title.as_deref(), Some("Old title"));
    assert_eq!(slides[0].body.as_deref(), Some("Old body 🚀"));
    assert_eq!(slides[0].notes.as_deref(), Some("Speaker 🚀"));
    assert_eq!(slides[0].name.as_deref(), Some("Agenda 🚀"));

    let before = editor.to_bytes().unwrap();
    assert!(editor.set_slide_name(2, Some("missing")).is_err());
    assert_eq!(editor.to_bytes().unwrap(), before);
    assert!(editor.set_slide_name(0, Some("bad\0name")).is_err());
    assert_eq!(editor.to_bytes().unwrap(), before);
    editor.set_slide_name(0, None).unwrap();
    assert_eq!(editor.slides().unwrap()[0].name, None);
}

#[test]
fn build_settings_project_to_semantic_values_and_validate_transactionally() {
    let settings = KeynoteBuildSettings::appear_in();
    let semantic = settings.semantic().unwrap();
    assert_eq!(semantic.effect(), &litchi_keynote::build::Effect::Appear);
    assert_eq!(semantic.duration().as_f64(), 1.0);

    let before = settings.clone();
    let mut invalid = settings;
    assert!(
        invalid
            .set_delay(litchi_keynote::Seconds::new(0.25).unwrap())
            .is_err()
    );
    assert_eq!(invalid, before);

    assert!(invalid.set_duration(litchi_keynote::Seconds::ZERO).is_err());
    assert_eq!(invalid, before);

    invalid.set_start(BuildStart::AfterPrevious).unwrap();
    invalid
        .set_delay(litchi_keynote::Seconds::new(0.25).unwrap())
        .unwrap();
    let semantic = invalid.semantic().unwrap();
    assert_eq!(semantic.start(), BuildStart::AfterPrevious);
    assert_eq!(semantic.delay().as_f64(), 0.25);
}

#[test]
fn typed_build_mutators_preserve_transactional_invariants() {
    let mut rotate = KeynoteBuildSettings::rotate_action(90.0, KeynoteRotationDirection::Clockwise);
    let before_invalid_acceleration = rotate.clone();
    assert!(
        rotate
            .set_action_acceleration(BuildAcceleration::from_native(99))
            .is_err()
    );
    assert_eq!(rotate, before_invalid_acceleration);

    rotate
        .set_action_acceleration(BuildAcceleration::EaseIn)
        .unwrap();
    assert!(matches!(
        rotate.semantic().unwrap().effect(),
        litchi_keynote::build::Effect::Action(litchi_keynote::build::Action::Rotate(action))
            if action.acceleration() == BuildAcceleration::EaseIn
    ));

    let mut move_settings = KeynoteBuildSettings::move_action(4.0, 5.0);
    move_settings.set_move_alignment(true).unwrap();
    let path = KeynoteMotionPath::straight(8.0, 13.0);
    move_settings.set_move_path(path.clone()).unwrap();
    assert_eq!(
        move_settings
            .move_action
            .as_ref()
            .map(|action| (&action.path, action.align_to_path)),
        Some((&path, true))
    );

    let mut simple = KeynoteBuildSettings::appear_in();
    simple
        .set_effect(litchi_keynote::build::Effect::Appear)
        .unwrap();
    assert_eq!(
        simple.semantic().unwrap().effect(),
        &litchi_keynote::build::Effect::Appear
    );
    let before_invalid_path = move_settings.clone();
    let invalid_path = KeynoteMotionPath {
        subpaths: Vec::new(),
        natural_width: 0.0,
        natural_height: 0.0,
        horizontal_flip: false,
        vertical_flip: false,
    };
    assert!(move_settings.set_move_path(invalid_path).is_err());
    assert_eq!(move_settings, before_invalid_path);
}

#[test]
fn slide_build_crud_is_transactional_and_updates_native_caches() {
    let mut editor = KeynoteEditor::from_package(test_package()).unwrap();
    assert!(editor.slide_builds(0).unwrap().is_empty());

    let settings = KeynoteBuildSettings::appear_in();
    let created = editor.add_slide_build(0, 5, settings.clone()).unwrap();
    assert_eq!(created.drawable_object_id, 5);
    assert_eq!(created.settings, settings);
    assert_eq!(created.chunks.len(), 1);
    let graph = ObjectGraph::read(editor.package()).unwrap();
    let node: kn::SlideNodeArchive = graph.decode(3, "KN.SlideNodeArchive").unwrap();
    assert_eq!(node.build_event_count, Some(1));
    assert_eq!(node.build_event_count_cache_version, Some(2));
    assert_eq!(node.has_explicit_builds, Some(true));
    assert_eq!(node.has_explicit_builds_cache_version, Some(2));

    let mut updated = KeynoteBuildSettings::dissolve_in();
    updated.duration = 2.5;
    updated.delay = 0.25;
    updated.start = BuildStart::AfterTransition;
    editor
        .set_slide_build(0, created.object_id, updated.clone())
        .unwrap();
    assert_eq!(editor.slide_builds(0).unwrap()[0].settings, updated);

    let before_invalid = editor.to_bytes().unwrap();
    let mut invalid = updated.clone();
    invalid.duration = f64::NAN;
    assert!(
        editor
            .set_slide_build(0, created.object_id, invalid)
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before_invalid);
    assert!(
        editor
            .set_slide_build(1, created.object_id, updated)
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before_invalid);

    let removed = editor.remove_slide_build(0, created.object_id).unwrap();
    assert_eq!(removed.object_id, created.object_id);
    assert!(editor.slide_builds(0).unwrap().is_empty());
    let graph = ObjectGraph::read(editor.package()).unwrap();
    assert!(!graph.objects.contains_key(&created.object_id));
    assert!(!graph.objects.contains_key(&created.chunks[0].object_id));
    let node: kn::SlideNodeArchive = graph.decode(3, "KN.SlideNodeArchive").unwrap();
    assert_eq!(node.build_event_count, None);
    assert_eq!(node.build_event_count_cache_version, Some(u32::MAX));
    assert_eq!(node.has_explicit_builds, Some(false));
    assert_eq!(node.has_explicit_builds_cache_version, Some(2));

    let build_out = editor
        .add_slide_build(0, 5, KeynoteBuildSettings::appear_out())
        .unwrap();
    assert_eq!(build_out.settings.animation_type, "Out");
    assert_eq!(build_out.settings.effect, "apple:bc-appear");
    editor.remove_slide_build(0, build_out.object_id).unwrap();
}

#[test]
fn slide_rotate_action_crud_maps_native_parameters_and_is_transactional() {
    use kn::build_attributes_archive::{
        BuildAttributesAcceleration, BuildAttributesRotationDirection,
    };

    let mut editor = KeynoteEditor::from_package(test_package()).unwrap();
    let mut settings =
        KeynoteBuildSettings::rotate_action(810.0, KeynoteRotationDirection::Clockwise);
    settings.duration = 2.5;
    settings.rotation.as_mut().unwrap().acceleration = BuildAcceleration::EaseIn;
    let created = editor.add_slide_build(0, 5, settings.clone()).unwrap();
    assert_eq!(created.settings, settings);

    let graph = ObjectGraph::read(editor.package()).unwrap();
    let native: kn::BuildArchive = graph
        .decode_type(created.object_id, BUILD_MESSAGE_TYPE, "KN.BuildArchive")
        .unwrap();
    assert_eq!(native.attributes.action_rotation_angle, Some(810.0));
    assert_eq!(
        native.attributes.action_rotation_direction,
        Some(BuildAttributesRotationDirection::KClockwise as i32)
    );
    assert_eq!(
        native.attributes.action_acceleration,
        Some(BuildAttributesAcceleration::KEaseIn as i32)
    );
    assert_eq!(native.attributes.custom_text_delivery, None);
    assert_eq!(native.attributes.custom_delivery_option, None);

    let before_invalid = editor.to_bytes().unwrap();
    let mut invalid = settings.clone();
    invalid.rotation.as_mut().unwrap().total_degrees = f64::NAN;
    assert!(
        editor
            .set_slide_build(0, created.object_id, invalid)
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before_invalid);
    let mut mismatched = settings.clone();
    mismatched.effect = "apple:action-scale".to_owned();
    assert!(
        editor
            .set_slide_build(0, created.object_id, mismatched)
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before_invalid);
    let mut custom = settings.clone();
    custom.rotation.as_mut().unwrap().acceleration = BuildAcceleration::Custom;
    assert!(editor.add_slide_build(0, 6, custom).is_err());
    assert_eq!(editor.to_bytes().unwrap(), before_invalid);

    let mut updated = settings;
    updated.duration = 1.25;
    updated.rotation = Some(KeynoteRotationAction {
        total_degrees: 270.0,
        direction: KeynoteRotationDirection::Counterclockwise,
        acceleration: BuildAcceleration::EaseOut,
    });
    editor
        .set_slide_build(0, created.object_id, updated.clone())
        .unwrap();
    assert_eq!(editor.slide_builds(0).unwrap()[0].settings, updated);

    editor
        .set_slide_build(0, created.object_id, KeynoteBuildSettings::appear_out())
        .unwrap();
    let graph = ObjectGraph::read(editor.package()).unwrap();
    let native: kn::BuildArchive = graph
        .decode_type(created.object_id, BUILD_MESSAGE_TYPE, "KN.BuildArchive")
        .unwrap();
    assert_eq!(native.attributes.action_rotation_angle, None);
    assert_eq!(native.attributes.action_rotation_direction, None);
    assert_eq!(native.attributes.action_acceleration, None);
    assert_eq!(
        editor.slide_builds(0).unwrap()[0].settings,
        KeynoteBuildSettings::appear_out()
    );
}

#[test]
fn slide_custom_timing_curve_actions_create_update_and_clear_natively() {
    use kn::build_attributes_archive::BuildAttributesAcceleration;

    let curve = KeynoteBuildTimingCurve::cubic(
        KeynoteMotionPathPoint::new(0.18, 0.04),
        KeynoteMotionPathPoint::new(0.82, 0.96),
    );
    let settings = KeynoteBuildSettings::rotate_action(720.0, KeynoteRotationDirection::Clockwise)
        .with_custom_timing_curve(curve.clone())
        .unwrap();

    let mut editor = KeynoteEditor::from_package(test_package()).unwrap();
    let created = editor.add_slide_build(0, 5, settings.clone()).unwrap();
    assert_eq!(created.settings, settings);

    let graph = ObjectGraph::read(editor.package()).unwrap();
    let native: kn::BuildArchive = graph
        .decode_type(created.object_id, BUILD_MESSAGE_TYPE, "KN.BuildArchive")
        .unwrap();
    assert_eq!(
        native.attributes.action_acceleration,
        Some(BuildAttributesAcceleration::KCustom as i32)
    );
    let animation = native.attributes.animation_attributes.as_ref().unwrap();
    let stored =
        timing_curve_from_native(animation.custom_effect_timing_curve_1.as_ref().unwrap()).unwrap();
    assert_eq!(stored, curve);
    assert_eq!(animation.custom_effect_timing_curve_theme_name_1, None);
    drop(graph);

    let updated_curve = KeynoteBuildTimingCurve::cubic(
        KeynoteMotionPathPoint::new(0.11, 0.77),
        KeynoteMotionPathPoint::new(0.91, 0.21),
    );
    let mut updated = settings.clone();
    updated
        .set_custom_timing_curve(updated_curve.clone())
        .unwrap();
    editor
        .set_slide_build(0, created.object_id, updated.clone())
        .unwrap();
    assert_eq!(editor.slide_builds(0).unwrap()[0].settings, updated);

    let mut builtin = updated;
    builtin.rotation.as_mut().unwrap().acceleration = BuildAcceleration::EaseOut;
    builtin.timing_curve = None;
    editor
        .set_slide_build(0, created.object_id, builtin.clone())
        .unwrap();
    let graph = ObjectGraph::read(editor.package()).unwrap();
    let native: kn::BuildArchive = graph
        .decode_type(created.object_id, BUILD_MESSAGE_TYPE, "KN.BuildArchive")
        .unwrap();
    let animation = native.attributes.animation_attributes.as_ref().unwrap();
    assert_eq!(animation.custom_effect_timing_curve_1, None);
    assert_eq!(animation.custom_effect_timing_curve_theme_name_1, None);
    drop(graph);
    assert_eq!(editor.slide_builds(0).unwrap()[0].settings, builtin);

    let restored = builtin.with_custom_timing_curve(updated_curve).unwrap();
    editor
        .set_slide_build(0, created.object_id, restored.clone())
        .unwrap();
    assert_eq!(editor.slide_builds(0).unwrap()[0].settings, restored);
}

#[test]
fn slide_custom_timing_curve_rejects_invalid_pairing_and_shape_transactionally() {
    let mut editor = KeynoteEditor::from_package(test_package()).unwrap();
    let settings =
        KeynoteBuildSettings::rotate_action(90.0, KeynoteRotationDirection::Counterclockwise);
    let created = editor.add_slide_build(0, 5, settings.clone()).unwrap();
    let before = editor.to_bytes().unwrap();

    let mut builtin_curve = settings.clone();
    builtin_curve.timing_curve = Some(KeynoteBuildTimingCurve::linear());
    assert!(
        editor
            .set_slide_build(0, created.object_id, builtin_curve)
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before);

    let mut invalid_curve = KeynoteBuildTimingCurve::linear();
    invalid_curve.path.subpaths[0].nodes[1].point.y = 0.9;
    let mut invalid_setter = settings.clone();
    assert!(
        invalid_setter
            .set_custom_timing_curve(invalid_curve.clone())
            .is_err()
    );
    assert_eq!(invalid_setter, settings);
    assert!(
        settings
            .clone()
            .with_custom_timing_curve(invalid_curve.clone())
            .is_err()
    );
    let mut invalid = settings.clone();
    invalid.rotation.as_mut().unwrap().acceleration = BuildAcceleration::Custom;
    invalid.timing_curve = Some(invalid_curve);
    assert!(
        editor
            .set_slide_build(0, created.object_id, invalid)
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before);

    assert!(
        KeynoteBuildSettings::appear_in()
            .with_custom_timing_curve(KeynoteBuildTimingCurve::linear())
            .is_err()
    );
}

#[test]
fn slide_custom_timing_curve_updates_preserve_deep_unknown_path_wire() {
    let settings = KeynoteBuildSettings::rotate_action(360.0, KeynoteRotationDirection::Clockwise)
        .with_custom_timing_curve(KeynoteBuildTimingCurve::cubic(
            KeynoteMotionPathPoint::new(0.15, 0.1),
            KeynoteMotionPathPoint::new(0.85, 0.9),
        ))
        .unwrap();
    let mut editor = KeynoteEditor::from_package(test_package()).unwrap();
    let created = editor.add_slide_build(0, 5, settings).unwrap();
    let mut package = editor.into_package();
    let suffixes = [
        unknown_varint(99, 990),
        unknown_varint(98, 980),
        unknown_varint(97, 970),
        unknown_varint(96, 960),
        unknown_varint(95, 950),
    ];
    package
        .update_archive("Index/Slide-4.iwa", |archive| {
            let build = archive.object_mut(created.object_id).unwrap();
            let message = build.messages[0].clone();
            let data = transform_length_delimited_field(&message.data, 4, |attributes| {
                transform_length_delimited_field(attributes, 18, |animation| {
                    transform_length_delimited_field(animation, 8, |path_source| {
                        let mut path_source =
                            transform_length_delimited_field(path_source, 8, |editable| {
                                let mut editable = transform_length_delimited_fields_at_path(
                                    editable,
                                    &[1],
                                    |subpath| {
                                        let mut subpath =
                                            transform_length_delimited_fields_at_path(
                                                subpath,
                                                &[1],
                                                |node| {
                                                    let mut node =
                                                        transform_length_delimited_field(
                                                            node,
                                                            2,
                                                            |point| {
                                                                let mut point = point.to_vec();
                                                                point.extend_from_slice(
                                                                    &suffixes[4],
                                                                );
                                                                Ok(point)
                                                            },
                                                        )?;
                                                    node.extend_from_slice(&suffixes[3]);
                                                    Ok(node)
                                                },
                                            )?;
                                        subpath.extend_from_slice(&suffixes[2]);
                                        Ok(subpath)
                                    },
                                )?;
                                editable.extend_from_slice(&suffixes[1]);
                                Ok(editable)
                            })?;
                        path_source.extend_from_slice(&suffixes[0]);
                        Ok(path_source)
                    })
                })
            })?;
            build.replace_message(
                0,
                RawMessage {
                    type_: message.type_,
                    data,
                },
            )?;
            Ok(())
        })
        .unwrap();

    let mut editor = KeynoteEditor::from_package(package).unwrap();
    let mut settings = editor.slide_builds(0).unwrap()[0].settings.clone();
    settings.timing_curve.as_mut().unwrap().path.subpaths[0].nodes[0]
        .out_control_point
        .x = 0.22;
    editor
        .set_slide_build(0, created.object_id, settings)
        .unwrap();

    let graph = ObjectGraph::read(editor.package()).unwrap();
    let build = graph
        .message_data_type(created.object_id, BUILD_MESSAGE_TYPE, "KN.BuildArchive")
        .unwrap();
    let attributes = repeated_length_delimited_payloads(build, 4).unwrap()[0];
    let animation = repeated_length_delimited_payloads(attributes, 18).unwrap()[0];
    let path_source = repeated_length_delimited_payloads(animation, 8).unwrap()[0];
    assert!(path_source.ends_with(&suffixes[0]));
    let editable = repeated_length_delimited_payloads(path_source, 8).unwrap()[0];
    assert!(editable.ends_with(&suffixes[1]));
    let subpath = repeated_length_delimited_payloads(editable, 1).unwrap()[0];
    assert!(subpath.ends_with(&suffixes[2]));
    let node = repeated_length_delimited_payloads(subpath, 1).unwrap()[0];
    assert!(node.ends_with(&suffixes[3]));
    let point = repeated_length_delimited_payloads(node, 2).unwrap()[0];
    assert!(point.ends_with(&suffixes[4]));
}

#[test]
fn slide_rotate_action_rejects_missing_and_duplicate_native_parameters() {
    let mut editor = KeynoteEditor::from_package(test_package()).unwrap();
    let settings =
        KeynoteBuildSettings::rotate_action(90.0, KeynoteRotationDirection::Counterclockwise);
    let created = editor.add_slide_build(0, 5, settings.clone()).unwrap();
    let mut package = editor.into_package();
    package
        .update_archive("Index/Slide-4.iwa", |archive| {
            let build = archive.object_mut(created.object_id).unwrap();
            let message = build.messages[0].clone();
            let data = transform_length_delimited_field(&message.data, 4, |attributes| {
                patch_varint_field(attributes, 13, true, None)
            })?;
            build.replace_message(
                0,
                RawMessage {
                    type_: message.type_,
                    data,
                },
            )?;
            Ok(())
        })
        .unwrap();
    assert!(
        KeynoteEditor::from_package(package)
            .unwrap()
            .slide_builds(0)
            .is_err()
    );

    let mut editor = KeynoteEditor::from_package(test_package()).unwrap();
    let created = editor.add_slide_build(0, 5, settings.clone()).unwrap();
    let mut package = editor.into_package();
    package
        .update_archive("Index/Slide-4.iwa", |archive| {
            let build = archive.object_mut(created.object_id).unwrap();
            let message = build.messages[0].clone();
            let data = transform_length_delimited_field(&message.data, 4, |attributes| {
                let mut attributes = attributes.to_vec();
                append_unknown_varint(
                    &mut attributes,
                    10,
                    native_rotation_direction(KeynoteRotationDirection::Clockwise) as u64,
                );
                Ok(attributes)
            })?;
            build.replace_message(
                0,
                RawMessage {
                    type_: message.type_,
                    data,
                },
            )?;
            Ok(())
        })
        .unwrap();
    let mut malformed = KeynoteEditor::from_package(package).unwrap();
    assert!(malformed.slide_builds(0).is_err());
    let before = malformed.to_bytes().unwrap();
    assert!(
        malformed
            .set_slide_build(0, created.object_id, settings)
            .is_err()
    );
    assert_eq!(malformed.to_bytes().unwrap(), before);
}

#[test]
fn slide_scale_and_opacity_action_crud_maps_native_parameters_and_is_transactional() {
    use kn::build_attributes_archive::BuildAttributesAcceleration;

    let mut editor = KeynoteEditor::from_package(test_package()).unwrap();
    let mut scale = KeynoteBuildSettings::scale_action(0.280_904_227_638_190_7);
    scale.duration = 1.75;
    scale.scale.as_mut().unwrap().acceleration = BuildAcceleration::EaseOut;
    let scale_build = editor.add_slide_build(0, 5, scale.clone()).unwrap();

    let mut opacity = KeynoteBuildSettings::opacity_action(37.0);
    opacity.duration = 2.25;
    opacity.opacity.as_mut().unwrap().acceleration = BuildAcceleration::EaseIn;
    let opacity_build = editor.add_slide_build(0, 6, opacity.clone()).unwrap();
    assert_eq!(editor.slide_builds(0).unwrap()[0].settings, scale);
    assert_eq!(editor.slide_builds(0).unwrap()[1].settings, opacity);

    let graph = ObjectGraph::read(editor.package()).unwrap();
    let native_scale: kn::BuildArchive = graph
        .decode_type(scale_build.object_id, BUILD_MESSAGE_TYPE, "KN.BuildArchive")
        .unwrap();
    assert_eq!(
        native_scale.attributes.action_scale_size,
        Some(0.280_904_227_638_190_7)
    );
    assert_eq!(
        native_scale.attributes.action_acceleration,
        Some(BuildAttributesAcceleration::KEaseOut as i32)
    );
    let native_opacity: kn::BuildArchive = graph
        .decode_type(
            opacity_build.object_id,
            BUILD_MESSAGE_TYPE,
            "KN.BuildArchive",
        )
        .unwrap();
    assert_eq!(native_opacity.attributes.action_color_alpha, Some(37.0));
    assert_eq!(
        native_opacity.attributes.action_acceleration,
        Some(BuildAttributesAcceleration::KEaseIn as i32)
    );

    let before_invalid = editor.to_bytes().unwrap();
    let mut invalid_scale = scale.clone();
    invalid_scale.scale.as_mut().unwrap().scale_factor = 0.0;
    assert!(
        editor
            .set_slide_build(0, scale_build.object_id, invalid_scale)
            .is_err()
    );
    let mut invalid_opacity = opacity.clone();
    invalid_opacity.opacity.as_mut().unwrap().opacity_percent = 100.01;
    assert!(
        editor
            .set_slide_build(0, opacity_build.object_id, invalid_opacity)
            .is_err()
    );
    let mut mismatched = scale.clone();
    mismatched.opacity = Some(KeynoteOpacityAction {
        opacity_percent: 50.0,
        acceleration: BuildAcceleration::None,
    });
    assert!(
        editor
            .set_slide_build(0, scale_build.object_id, mismatched)
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before_invalid);

    editor
        .set_slide_build(0, scale_build.object_id, opacity.clone())
        .unwrap();
    let graph = ObjectGraph::read(editor.package()).unwrap();
    let converted: kn::BuildArchive = graph
        .decode_type(scale_build.object_id, BUILD_MESSAGE_TYPE, "KN.BuildArchive")
        .unwrap();
    assert_eq!(converted.attributes.action_rotation_angle, None);
    assert_eq!(converted.attributes.action_rotation_direction, None);
    assert_eq!(converted.attributes.action_scale_size, None);
    assert_eq!(converted.attributes.action_color_alpha, Some(37.0));
    assert_eq!(editor.slide_builds(0).unwrap()[0].settings, opacity);
}

#[test]
fn slide_scale_and_opacity_reject_malformed_native_parameters() {
    for (settings, scalar_field) in [
        (KeynoteBuildSettings::scale_action(1.5), 11),
        (KeynoteBuildSettings::opacity_action(50.0), 12),
    ] {
        let mut editor = KeynoteEditor::from_package(test_package()).unwrap();
        let created = editor.add_slide_build(0, 5, settings.clone()).unwrap();
        let mut package = editor.into_package();
        package
            .update_archive("Index/Slide-4.iwa", |archive| {
                let build = archive.object_mut(created.object_id).unwrap();
                let message = build.messages[0].clone();
                let data = transform_length_delimited_field(&message.data, 4, |attributes| {
                    patch_fixed64_field(attributes, scalar_field, true, None)
                })?;
                build.replace_message(
                    0,
                    RawMessage {
                        type_: message.type_,
                        data,
                    },
                )?;
                Ok(())
            })
            .unwrap();
        assert!(
            KeynoteEditor::from_package(package)
                .unwrap()
                .slide_builds(0)
                .is_err()
        );

        let mut editor = KeynoteEditor::from_package(test_package()).unwrap();
        let created = editor.add_slide_build(0, 5, settings).unwrap();
        let mut package = editor.into_package();
        package
            .update_archive("Index/Slide-4.iwa", |archive| {
                let build = archive.object_mut(created.object_id).unwrap();
                let message = build.messages[0].clone();
                let decoded = kn::BuildArchive::decode(message.data.as_slice()).unwrap();
                let scalar = if scalar_field == 11 {
                    decoded.attributes.action_scale_size.unwrap()
                } else {
                    decoded.attributes.action_color_alpha.unwrap()
                };
                let data = transform_length_delimited_field(&message.data, 4, |attributes| {
                    let mut attributes = attributes.to_vec();
                    append_unknown_fixed64(&mut attributes, scalar_field, scalar.to_bits());
                    Ok(attributes)
                })?;
                build.replace_message(
                    0,
                    RawMessage {
                        type_: message.type_,
                        data,
                    },
                )?;
                Ok(())
            })
            .unwrap();
        assert!(
            KeynoteEditor::from_package(package)
                .unwrap()
                .slide_builds(0)
                .is_err()
        );
    }
}

#[test]
fn slide_emphasis_action_crud_maps_native_parameters_and_cleans_conversions() {
    use kn::build_attributes_archive::ActionBuildAttributesJiggleIntensity;

    let mut editor = KeynoteEditor::from_package(test_package()).unwrap();
    let settings = [
        KeynoteBuildSettings::blink_action(7, true),
        KeynoteBuildSettings::bounce_action(5, false),
        KeynoteBuildSettings::flip_action(6, KeynoteFlipDirection::RightToLeft),
        KeynoteBuildSettings::jiggle_action(KeynoteJiggleIntensity::Large),
        KeynoteBuildSettings::pop_action(165.0),
        KeynoteBuildSettings::pulse_action(6, 135.0),
    ];
    let mut builds = Vec::new();
    for settings in &settings {
        let created = editor.add_slide_build(0, 5, settings.clone()).unwrap();
        assert_eq!(created.settings, *settings);
        builds.push(created.object_id);
    }
    assert_eq!(
        editor
            .slide_builds(0)
            .unwrap()
            .into_iter()
            .map(|build| build.settings)
            .collect::<Vec<_>>(),
        settings
    );

    let graph = ObjectGraph::read(editor.package()).unwrap();
    let native = builds
        .iter()
        .map(|object_id| {
            graph
                .decode_type::<kn::BuildArchive>(*object_id, BUILD_MESSAGE_TYPE, "KN.BuildArchive")
                .unwrap()
        })
        .collect::<Vec<_>>();
    assert_eq!(native[0].attributes.custom_action_decay, Some(true));
    assert_eq!(native[0].attributes.custom_action_repeat_count, Some(7));
    assert_eq!(native[1].attributes.custom_action_decay, Some(false));
    assert_eq!(native[1].attributes.custom_action_repeat_count, Some(5));
    assert_eq!(native[2].attributes.custom_action_repeat_count, Some(6));
    assert_eq!(
        native[2]
            .attributes
            .animation_attributes
            .as_ref()
            .unwrap()
            .direction,
        Some(12)
    );
    assert_eq!(
        native[3].attributes.custom_action_jiggle_intensity,
        Some(ActionBuildAttributesJiggleIntensity::KJiggleIntensityLarge as i32)
    );
    assert_eq!(native[4].attributes.custom_action_scale, Some(165.0));
    assert_eq!(native[5].attributes.custom_action_repeat_count, Some(6));
    assert_eq!(native[5].attributes.custom_action_scale, Some(135.0));
    for native in &native {
        assert_eq!(native.attributes.action_acceleration, None);
        assert_eq!(native.attributes.custom_text_delivery, None);
        assert_eq!(native.attributes.custom_delivery_option, None);
    }
    drop(graph);

    let before_invalid = editor.to_bytes().unwrap();
    let mut invalid = KeynoteBuildSettings::pulse_action(0, 135.0);
    assert!(
        editor
            .set_slide_build(0, builds[5], invalid.clone())
            .is_err()
    );
    invalid.emphasis = Some(KeynoteEmphasisAction::Pulse {
        repeat_count: 2,
        scale_percent: f64::NAN,
    });
    assert!(editor.set_slide_build(0, builds[5], invalid).is_err());
    let mut mismatched = KeynoteBuildSettings::blink_action(2, true);
    mismatched.effect = BOUNCE_ACTION_EFFECT.to_owned();
    assert!(editor.set_slide_build(0, builds[5], mismatched).is_err());
    assert_eq!(editor.to_bytes().unwrap(), before_invalid);

    let jiggle = KeynoteBuildSettings::jiggle_action(KeynoteJiggleIntensity::Small);
    editor
        .set_slide_build(0, builds[5], jiggle.clone())
        .unwrap();
    let graph = ObjectGraph::read(editor.package()).unwrap();
    let converted: kn::BuildArchive = graph
        .decode_type(builds[5], BUILD_MESSAGE_TYPE, "KN.BuildArchive")
        .unwrap();
    assert_eq!(converted.attributes.custom_action_decay, None);
    assert_eq!(converted.attributes.custom_action_repeat_count, None);
    assert_eq!(converted.attributes.custom_action_scale, None);
    assert_eq!(
        converted.attributes.custom_action_jiggle_intensity,
        Some(ActionBuildAttributesJiggleIntensity::KJiggleIntensitySmall as i32)
    );
    assert_eq!(
        converted
            .attributes
            .animation_attributes
            .as_ref()
            .unwrap()
            .direction,
        None
    );
    drop(graph);
    assert_eq!(editor.slide_builds(0).unwrap()[5].settings, jiggle);

    editor
        .set_slide_build(0, builds[5], KeynoteBuildSettings::appear_out())
        .unwrap();
    let graph = ObjectGraph::read(editor.package()).unwrap();
    let converted: kn::BuildArchive = graph
        .decode_type(builds[5], BUILD_MESSAGE_TYPE, "KN.BuildArchive")
        .unwrap();
    assert_eq!(converted.attributes.custom_action_decay, None);
    assert_eq!(converted.attributes.custom_action_repeat_count, None);
    assert_eq!(converted.attributes.custom_action_scale, None);
    assert_eq!(converted.attributes.custom_action_jiggle_intensity, None);
}

#[test]
fn slide_emphasis_actions_reject_missing_duplicate_and_cross_effect_fields() {
    let cases = [
        KeynoteBuildSettings::blink_action(7, true),
        KeynoteBuildSettings::bounce_action(5, false),
        KeynoteBuildSettings::flip_action(6, KeynoteFlipDirection::LeftToRight),
        KeynoteBuildSettings::jiggle_action(KeynoteJiggleIntensity::Medium),
        KeynoteBuildSettings::pop_action(165.0),
        KeynoteBuildSettings::pulse_action(6, 135.0),
        KeynoteBuildSettings::blink_action(2, false),
        KeynoteBuildSettings::rotate_action(90.0, KeynoteRotationDirection::Clockwise),
    ];
    for (mutation, settings) in cases.into_iter().enumerate() {
        let mut editor = KeynoteEditor::from_package(test_package()).unwrap();
        let created = editor.add_slide_build(0, 5, settings.clone()).unwrap();
        let mut package = editor.into_package();
        package
            .update_archive("Index/Slide-4.iwa", |archive| {
                let build = archive.object_mut(created.object_id).unwrap();
                let message = build.messages[0].clone();
                let data =
                    transform_length_delimited_field(
                        &message.data,
                        4,
                        |attributes| match mutation {
                            0 => patch_varint_field(attributes, 23, true, None),
                            1 => patch_varint_field(attributes, 24, true, None),
                            2 => transform_length_delimited_field(attributes, 18, |animation| {
                                patch_varint_field(animation, 4, true, None)
                            }),
                            3 => patch_varint_field(attributes, 26, true, None),
                            4 => patch_fixed64_field(attributes, 25, true, None),
                            5 => {
                                let mut attributes = attributes.to_vec();
                                append_unknown_varint(&mut attributes, 24, 6);
                                Ok(attributes)
                            },
                            6 => {
                                let mut attributes = attributes.to_vec();
                                append_unknown_fixed64(&mut attributes, 9, 90.0_f64.to_bits());
                                Ok(attributes)
                            },
                            7 => {
                                let mut attributes = attributes.to_vec();
                                append_unknown_varint(&mut attributes, 24, 2);
                                Ok(attributes)
                            },
                            _ => unreachable!(),
                        },
                    )?;
                build.replace_message(
                    0,
                    RawMessage {
                        type_: message.type_,
                        data,
                    },
                )?;
                Ok(())
            })
            .unwrap();
        let mut malformed = KeynoteEditor::from_package(package).unwrap();
        assert!(malformed.slide_builds(0).is_err(), "mutation {mutation}");
        let before = malformed.to_bytes().unwrap();
        assert!(
            malformed
                .set_slide_build(0, created.object_id, settings)
                .is_err(),
            "mutation {mutation}"
        );
        assert_eq!(malformed.to_bytes().unwrap(), before);
    }
}

#[test]
fn slide_keyboard_build_crud_maps_native_parameters_and_cleans_conversions() {
    let mut editor = KeynoteEditor::from_package(test_package()).unwrap();
    let mut build_in = KeynoteBuildSettings::keyboard_in(KeynoteKeyboardDirection::Forward, false);
    build_in.duration = 2.25;
    let created_in = editor.add_slide_build(0, 5, build_in.clone()).unwrap();
    let mut build_out =
        KeynoteBuildSettings::keyboard_out(KeynoteKeyboardDirection::Backward, true);
    build_out.duration = 1.75;
    let created_out = editor.add_slide_build(0, 6, build_out.clone()).unwrap();
    assert_eq!(created_in.settings, build_in);
    assert_eq!(created_out.settings, build_out);

    let graph = ObjectGraph::read(editor.package()).unwrap();
    let native_in: kn::BuildArchive = graph
        .decode_type(created_in.object_id, BUILD_MESSAGE_TYPE, "KN.BuildArchive")
        .unwrap();
    let native_out: kn::BuildArchive = graph
        .decode_type(created_out.object_id, BUILD_MESSAGE_TYPE, "KN.BuildArchive")
        .unwrap();
    assert_eq!(native_in.attributes.custom_cursor, Some(false));
    assert_eq!(native_out.attributes.custom_cursor, Some(true));
    assert_eq!(
        native_in
            .attributes
            .animation_attributes
            .as_ref()
            .unwrap()
            .direction,
        Some(111)
    );
    assert_eq!(
        native_out
            .attributes
            .animation_attributes
            .as_ref()
            .unwrap()
            .direction,
        Some(112)
    );
    drop(graph);

    let before_invalid = editor.to_bytes().unwrap();
    let mut mismatched_direction = build_in.clone();
    mismatched_direction.direction = Some(112);
    assert!(
        editor
            .set_slide_build(0, created_in.object_id, mismatched_direction)
            .is_err()
    );
    let mut invalid_phase = build_in.clone();
    invalid_phase.animation_type = "Action".to_owned();
    assert!(
        editor
            .set_slide_build(0, created_in.object_id, invalid_phase)
            .is_err()
    );
    let mut mixed = build_in.clone();
    mixed.emphasis = Some(KeynoteEmphasisAction::Blink {
        repeat_count: 2,
        fade: false,
    });
    assert!(
        editor
            .set_slide_build(0, created_in.object_id, mixed)
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before_invalid);

    let converted = KeynoteBuildSettings::keyboard_out(KeynoteKeyboardDirection::Backward, true);
    editor
        .set_slide_build(0, created_in.object_id, converted.clone())
        .unwrap();
    assert_eq!(editor.slide_builds(0).unwrap()[0].settings, converted);

    editor
        .set_slide_build(0, created_out.object_id, KeynoteBuildSettings::appear_out())
        .unwrap();
    let graph = ObjectGraph::read(editor.package()).unwrap();
    let cleaned: kn::BuildArchive = graph
        .decode_type(created_out.object_id, BUILD_MESSAGE_TYPE, "KN.BuildArchive")
        .unwrap();
    assert_eq!(cleaned.attributes.custom_cursor, None);
    assert_eq!(
        cleaned
            .attributes
            .animation_attributes
            .as_ref()
            .unwrap()
            .direction,
        None
    );
}

#[test]
fn slide_keyboard_build_rejects_missing_duplicate_and_cross_effect_fields() {
    for mutation in 0..5 {
        let mut editor = KeynoteEditor::from_package(test_package()).unwrap();
        let settings = KeynoteBuildSettings::keyboard_in(KeynoteKeyboardDirection::Forward, true);
        let created = editor.add_slide_build(0, 5, settings.clone()).unwrap();
        let mut package = editor.into_package();
        package
            .update_archive("Index/Slide-4.iwa", |archive| {
                let build = archive.object_mut(created.object_id).unwrap();
                let message = build.messages[0].clone();
                let data =
                    transform_length_delimited_field(
                        &message.data,
                        4,
                        |attributes| match mutation {
                            0 => patch_varint_field(attributes, 36, true, None),
                            1 => {
                                let mut attributes = attributes.to_vec();
                                append_unknown_varint(&mut attributes, 36, 0);
                                Ok(attributes)
                            },
                            2 => transform_length_delimited_field(attributes, 18, |animation| {
                                patch_varint_field(animation, 4, true, None)
                            }),
                            3 => transform_length_delimited_field(attributes, 18, |animation| {
                                let mut animation = animation.to_vec();
                                append_unknown_varint(&mut animation, 4, 112);
                                Ok(animation)
                            }),
                            4 => {
                                let mut attributes = attributes.to_vec();
                                append_unknown_varint(&mut attributes, 24, 2);
                                Ok(attributes)
                            },
                            _ => unreachable!(),
                        },
                    )?;
                build.replace_message(
                    0,
                    RawMessage {
                        type_: message.type_,
                        data,
                    },
                )?;
                Ok(())
            })
            .unwrap();
        let mut malformed = KeynoteEditor::from_package(package).unwrap();
        assert!(malformed.slide_builds(0).is_err(), "mutation {mutation}");
        let before = malformed.to_bytes().unwrap();
        assert!(
            malformed
                .set_slide_build(0, created.object_id, settings)
                .is_err(),
            "mutation {mutation}"
        );
        assert_eq!(malformed.to_bytes().unwrap(), before);
    }
}

#[test]
fn slide_dissolve_build_crud_uses_native_text_delivery_defaults() {
    use kn::build_attributes_archive::{
        BuildAttributesDeliveryOption, BuildAttributesTextDelivery,
    };

    let mut editor = KeynoteEditor::from_package(test_package()).unwrap();
    let build_in = KeynoteBuildSettings::dissolve_in();
    let created = editor.add_slide_build(0, 5, build_in.clone()).unwrap();
    assert_eq!(created.settings, build_in);
    assert_eq!(
        editor.slide_builds(0).unwrap()[0].settings,
        KeynoteBuildSettings::dissolve_in()
    );

    let graph = ObjectGraph::read(editor.package()).unwrap();
    let native: kn::BuildArchive = graph
        .decode_type(created.object_id, BUILD_MESSAGE_TYPE, "KN.BuildArchive")
        .unwrap();
    let animation = native.attributes.animation_attributes.as_ref().unwrap();
    assert_eq!(animation.animation_type.as_deref(), Some("In"));
    assert_eq!(animation.effect.as_deref(), Some(DISSOLVE_BUILD_EFFECT));
    assert_eq!(animation.duration, Some(1.0));
    assert_eq!(animation.direction, None);
    assert_eq!(
        native.attributes.custom_text_delivery,
        Some(BuildAttributesTextDelivery::KTextDeliveryByObject as i32)
    );
    assert_eq!(
        native.attributes.custom_delivery_option,
        Some(BuildAttributesDeliveryOption::KDeliveryOptionForward as i32)
    );
    assert_eq!(native.attributes.custom_bounce, None);
    assert_eq!(native.attributes.custom_motion_blur, None);
    assert_eq!(native.attributes.custom_include_endpoints, None);
    assert_eq!(native.attributes.custom_shine, None);
    assert_eq!(native.attributes.custom_scale_amount, None);
    assert_eq!(native.attributes.custom_travel_distance, None);
    drop(graph);

    let build_out = KeynoteBuildSettings::dissolve_out();
    editor
        .set_slide_build(0, created.object_id, build_out.clone())
        .unwrap();
    assert_eq!(editor.slide_builds(0).unwrap()[0].settings, build_out);

    let before_invalid = editor.to_bytes().unwrap();
    let mut invalid_direction = KeynoteBuildSettings::dissolve_out();
    invalid_direction.direction = Some(11);
    assert!(
        editor
            .set_slide_build(0, created.object_id, invalid_direction)
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before_invalid);

    editor.remove_slide_build(0, created.object_id).unwrap();
    assert!(editor.slide_builds(0).unwrap().is_empty());
}

#[test]
fn slide_object_build_crud_maps_native_effects_directions_and_defaults() {
    let mut editor = KeynoteEditor::from_package(test_package()).unwrap();
    let settings = [
        KeynoteBuildSettings::shimmer_in(),
        KeynoteBuildSettings::shimmer_out(),
        KeynoteBuildSettings::skid_in(KeynoteHorizontalBuildDirection::LeftToRight),
        KeynoteBuildSettings::skid_out(KeynoteHorizontalBuildDirection::RightToLeft),
        KeynoteBuildSettings::swoosh_in(KeynoteSwooshDirection::Center),
        KeynoteBuildSettings::swoosh_out(KeynoteSwooshDirection::FromRight),
        KeynoteBuildSettings::trace_in(KeynoteHorizontalBuildDirection::LeftToRight),
        KeynoteBuildSettings::trace_out(KeynoteHorizontalBuildDirection::RightToLeft),
    ];
    let mut object_ids = Vec::new();
    for settings in &settings {
        let created = editor.add_slide_build(0, 5, settings.clone()).unwrap();
        assert_eq!(created.settings, *settings);
        object_ids.push(created.object_id);
    }
    assert_eq!(
        editor
            .slide_builds(0)
            .unwrap()
            .into_iter()
            .map(|build| build.settings)
            .collect::<Vec<_>>(),
        settings
    );

    let graph = ObjectGraph::read(editor.package()).unwrap();
    let native = object_ids
        .iter()
        .map(|object_id| {
            graph
                .decode_type::<kn::BuildArchive>(*object_id, BUILD_MESSAGE_TYPE, "KN.BuildArchive")
                .unwrap()
        })
        .collect::<Vec<_>>();
    let effects = native
        .iter()
        .map(|build| {
            build
                .attributes
                .animation_attributes
                .as_ref()
                .unwrap()
                .effect
                .as_deref()
                .unwrap()
        })
        .collect::<Vec<_>>();
    assert_eq!(
        effects,
        [
            SHIMMER_BUILD_EFFECT,
            SHIMMER_BUILD_EFFECT,
            SKID_BUILD_EFFECT,
            SKID_BUILD_EFFECT,
            SWOOSH_BUILD_EFFECT,
            SWOOSH_BUILD_EFFECT,
            TRACE_BUILD_EFFECT,
            TRACE_BUILD_EFFECT,
        ]
    );
    let directions = native
        .iter()
        .map(|build| {
            build
                .attributes
                .animation_attributes
                .as_ref()
                .unwrap()
                .direction
        })
        .collect::<Vec<_>>();
    assert_eq!(
        directions,
        [
            None,
            None,
            Some(11),
            Some(12),
            None,
            Some(12),
            Some(11),
            Some(12)
        ]
    );
    let durations = native
        .iter()
        .map(|build| {
            build
                .attributes
                .animation_attributes
                .as_ref()
                .unwrap()
                .duration
                .unwrap()
        })
        .collect::<Vec<_>>();
    assert_eq!(durations, [1.5, 1.5, 1.25, 1.25, 1.0, 1.0, 2.0, 2.0]);
    for build in &native {
        assert_eq!(build.attributes.custom_text_delivery, None);
        assert_eq!(build.attributes.custom_delivery_option, None);
        assert_eq!(build.attributes.custom_bounce, None);
        assert_eq!(build.attributes.custom_cursor, None);
        assert_eq!(build.attributes.custom_align_to_path, None);
    }
    drop(graph);

    let before_noop = editor.to_bytes().unwrap();
    editor
        .set_slide_build(0, object_ids[3], settings[3].clone())
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), before_noop);

    let before_invalid = editor.to_bytes().unwrap();
    let mut mismatched =
        KeynoteBuildSettings::trace_in(KeynoteHorizontalBuildDirection::LeftToRight);
    mismatched.effect = SKID_BUILD_EFFECT.to_owned();
    assert!(
        editor
            .set_slide_build(0, object_ids[6], mismatched)
            .is_err()
    );
    let mut wrong_direction = settings[6].clone();
    wrong_direction.direction = Some(12);
    assert!(
        editor
            .set_slide_build(0, object_ids[6], wrong_direction)
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before_invalid);

    let converted = KeynoteBuildSettings::swoosh_out(KeynoteSwooshDirection::FromLeft);
    editor
        .set_slide_build(0, object_ids[6], converted.clone())
        .unwrap();
    assert_eq!(editor.slide_builds(0).unwrap()[6].settings, converted);

    editor
        .set_slide_build(0, object_ids[7], KeynoteBuildSettings::appear_out())
        .unwrap();
    let graph = ObjectGraph::read(editor.package()).unwrap();
    let cleaned: kn::BuildArchive = graph
        .decode_type(object_ids[7], BUILD_MESSAGE_TYPE, "KN.BuildArchive")
        .unwrap();
    let animation = cleaned.attributes.animation_attributes.as_ref().unwrap();
    assert_eq!(animation.effect.as_deref(), Some("apple:bc-appear"));
    assert_eq!(animation.direction, None);
    drop(graph);
    assert_eq!(
        editor.slide_builds(0).unwrap()[7].settings.object_effect,
        None
    );

    editor.remove_slide_build(0, object_ids[0]).unwrap();
    assert!(
        !editor
            .slide_builds(0)
            .unwrap()
            .iter()
            .any(|build| build.object_id == object_ids[0])
    );
}

#[test]
fn slide_object_build_accepts_omitted_native_left_to_right_default_losslessly() {
    let mut editor = KeynoteEditor::from_package(test_package()).unwrap();
    let settings = KeynoteBuildSettings::trace_in(KeynoteHorizontalBuildDirection::LeftToRight);
    let created = editor.add_slide_build(0, 5, settings).unwrap();
    let mut package = editor.into_package();
    package
        .update_archive("Index/Slide-4.iwa", |archive| {
            let build = archive.object_mut(created.object_id).unwrap();
            let message = build.messages[0].clone();
            let data = transform_length_delimited_field(&message.data, 4, |attributes| {
                transform_length_delimited_field(attributes, 18, |animation| {
                    patch_varint_field(animation, 4, true, None)
                })
            })?;
            build.replace_message(
                0,
                RawMessage {
                    type_: message.type_,
                    data,
                },
            )?;
            Ok(())
        })
        .unwrap();
    let mut editor = KeynoteEditor::from_package(package).unwrap();
    let read = editor.slide_builds(0).unwrap().remove(0).settings;
    assert_eq!(read.direction, None);
    assert_eq!(
        read.object_effect,
        Some(KeynoteObjectBuildEffect::Trace {
            direction: KeynoteHorizontalBuildDirection::LeftToRight,
        })
    );
    let before = editor.to_bytes().unwrap();
    editor.set_slide_build(0, created.object_id, read).unwrap();
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn slide_object_build_rejects_duplicate_wrong_wire_and_cross_effect_fields() {
    for mutation in 0..6 {
        let mut editor = KeynoteEditor::from_package(test_package()).unwrap();
        let settings = KeynoteBuildSettings::trace_in(KeynoteHorizontalBuildDirection::RightToLeft);
        let created = editor.add_slide_build(0, 5, settings.clone()).unwrap();
        let mut package = editor.into_package();
        package
            .update_archive("Index/Slide-4.iwa", |archive| {
                let build = archive.object_mut(created.object_id).unwrap();
                let message = build.messages[0].clone();
                let data =
                    transform_length_delimited_field(
                        &message.data,
                        4,
                        |attributes| match mutation {
                            0 => transform_length_delimited_field(attributes, 18, |animation| {
                                let mut animation = animation.to_vec();
                                append_unknown_varint(&mut animation, 4, 12);
                                Ok(animation)
                            }),
                            1 => transform_length_delimited_field(attributes, 18, |animation| {
                                let mut animation = animation.to_vec();
                                append_unknown_fixed64(&mut animation, 4, 12);
                                Ok(animation)
                            }),
                            2 => {
                                let mut attributes = attributes.to_vec();
                                append_unknown_varint(&mut attributes, 24, 2);
                                Ok(attributes)
                            },
                            3 => {
                                let mut attributes = attributes.to_vec();
                                append_unknown_varint(&mut attributes, 19, 1);
                                Ok(attributes)
                            },
                            4 => transform_length_delimited_field(attributes, 18, |animation| {
                                patch_varint_field(animation, 4, true, Some(99))
                            }),
                            5 => transform_length_delimited_field(attributes, 18, |animation| {
                                patch_length_delimited_field(animation, 1, true, Some(b"Action"))
                            }),
                            _ => unreachable!(),
                        },
                    )?;
                build.replace_message(
                    0,
                    RawMessage {
                        type_: message.type_,
                        data,
                    },
                )?;
                Ok(())
            })
            .unwrap();
        let mut malformed = KeynoteEditor::from_package(package).unwrap();
        assert!(malformed.slide_builds(0).is_err(), "mutation {mutation}");
        let before = malformed.to_bytes().unwrap();
        assert!(
            malformed
                .set_slide_build(0, created.object_id, settings)
                .is_err(),
            "mutation {mutation}"
        );
        assert_eq!(malformed.to_bytes().unwrap(), before);
    }
}

#[test]
fn slide_raw_custom_build_parameter_crud_is_lossless_and_transactional() {
    let mut editor = KeynoteEditor::from_package(test_package()).unwrap();
    let mut settings = KeynoteBuildSettings::appear_in();
    settings.effect = "com.example.future-build".to_owned();
    settings.direction = Some(42);
    settings.custom_parameters = KeynoteBuildCustomParameters {
        bounce: Some(false),
        motion_blur: Some(true),
        include_endpoints: Some(false),
        shine: Some(true),
        scale_amount: Some(1.375),
        travel_distance: Some(275.5),
    };
    let created = editor.add_slide_build(0, 5, settings.clone()).unwrap();
    assert_eq!(created.settings, settings);

    let graph = ObjectGraph::read(editor.package()).unwrap();
    let native: kn::BuildArchive = graph
        .decode_type(created.object_id, BUILD_MESSAGE_TYPE, "KN.BuildArchive")
        .unwrap();
    assert_eq!(native.attributes.custom_bounce, Some(false));
    assert_eq!(native.attributes.custom_motion_blur, Some(true));
    assert_eq!(native.attributes.custom_include_endpoints, Some(false));
    assert_eq!(native.attributes.custom_shine, Some(true));
    assert_eq!(native.attributes.custom_scale_amount, Some(1.375));
    assert_eq!(native.attributes.custom_travel_distance, Some(275.5));
    drop(graph);

    let before_noop = editor.to_bytes().unwrap();
    editor
        .set_slide_build(0, created.object_id, settings.clone())
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), before_noop);

    let before_invalid = editor.to_bytes().unwrap();
    let mut invalid = settings.clone();
    invalid.custom_parameters.travel_distance = Some(f64::NAN);
    assert!(
        editor
            .set_slide_build(0, created.object_id, invalid)
            .is_err()
    );
    let mut mixed = KeynoteBuildSettings::keyboard_in(KeynoteKeyboardDirection::Forward, true);
    mixed.custom_parameters.motion_blur = Some(true);
    assert!(editor.set_slide_build(0, created.object_id, mixed).is_err());
    assert_eq!(editor.to_bytes().unwrap(), before_invalid);

    editor
        .set_slide_build(0, created.object_id, KeynoteBuildSettings::appear_out())
        .unwrap();
    let graph = ObjectGraph::read(editor.package()).unwrap();
    let cleaned: kn::BuildArchive = graph
        .decode_type(created.object_id, BUILD_MESSAGE_TYPE, "KN.BuildArchive")
        .unwrap();
    assert_eq!(cleaned.attributes.custom_bounce, None);
    assert_eq!(cleaned.attributes.custom_motion_blur, None);
    assert_eq!(cleaned.attributes.custom_include_endpoints, None);
    assert_eq!(cleaned.attributes.custom_shine, None);
    assert_eq!(cleaned.attributes.custom_scale_amount, None);
    assert_eq!(cleaned.attributes.custom_travel_distance, None);
}

#[test]
fn slide_raw_custom_build_parameters_reject_duplicate_and_wrong_wire_types() {
    for mutation in 0..7 {
        let mut editor = KeynoteEditor::from_package(test_package()).unwrap();
        let mut settings = KeynoteBuildSettings::appear_in();
        settings.effect = "com.example.future-build".to_owned();
        settings.custom_parameters = KeynoteBuildCustomParameters {
            bounce: Some(true),
            motion_blur: Some(false),
            include_endpoints: Some(true),
            shine: Some(false),
            scale_amount: Some(1.25),
            travel_distance: Some(80.0),
        };
        let created = editor.add_slide_build(0, 5, settings.clone()).unwrap();
        let mut package = editor.into_package();
        package
            .update_archive("Index/Slide-4.iwa", |archive| {
                let build = archive.object_mut(created.object_id).unwrap();
                let message = build.messages[0].clone();
                let data = transform_length_delimited_field(&message.data, 4, |attributes| {
                    let mut attributes = attributes.to_vec();
                    match mutation {
                        0 => append_unknown_varint(&mut attributes, 19, 0),
                        1 => append_unknown_varint(&mut attributes, 29, 1),
                        2 => append_unknown_varint(&mut attributes, 30, 0),
                        3 => append_unknown_varint(&mut attributes, 33, 1),
                        4 => append_unknown_fixed64(&mut attributes, 34, 2.0_f64.to_bits()),
                        5 => append_unknown_fixed64(&mut attributes, 35, 120.0_f64.to_bits()),
                        6 => append_unknown_varint(&mut attributes, 34, 1),
                        _ => unreachable!(),
                    }
                    Ok(attributes)
                })?;
                build.replace_message(
                    0,
                    RawMessage {
                        type_: message.type_,
                        data,
                    },
                )?;
                Ok(())
            })
            .unwrap();
        let mut malformed = KeynoteEditor::from_package(package).unwrap();
        assert!(malformed.slide_builds(0).is_err(), "mutation {mutation}");
        let before = malformed.to_bytes().unwrap();
        assert!(
            malformed
                .set_slide_build(0, created.object_id, settings)
                .is_err(),
            "mutation {mutation}"
        );
        assert_eq!(malformed.to_bytes().unwrap(), before);
    }
}

#[test]
fn slide_move_action_crud_maps_editable_bezier_path_and_is_transactional() {
    use kn::build_attributes_archive::BuildAttributesAcceleration;

    let mut editor = KeynoteEditor::from_package(test_package()).unwrap();
    let mut settings = KeynoteBuildSettings::move_action(488.492, -258.172);
    settings.duration = 2.25;
    let move_action = settings.move_action.as_mut().unwrap();
    move_action.align_to_path = true;
    move_action.acceleration = BuildAcceleration::EaseOut;
    let created = editor.add_slide_build(0, 5, settings.clone()).unwrap();
    assert_eq!(created.settings, settings);

    let graph = ObjectGraph::read(editor.package()).unwrap();
    let native: kn::BuildArchive = graph
        .decode_type(created.object_id, BUILD_MESSAGE_TYPE, "KN.BuildArchive")
        .unwrap();
    assert_eq!(native.attributes.custom_align_to_path, Some(true));
    assert_eq!(
        native.attributes.action_acceleration,
        Some(BuildAttributesAcceleration::KEaseOut as i32)
    );
    let native_path = native.attributes.action_motion_path_source.unwrap();
    assert_eq!(native_path.horizontal_flip, Some(false));
    assert_eq!(native_path.vertical_flip, Some(false));
    let editable = native_path.editable_bezier_path_source.unwrap();
    assert_eq!(editable.subpaths.len(), 1);
    assert_eq!(editable.subpaths[0].nodes.len(), 2);
    assert_eq!(editable.subpaths[0].nodes[1].node_point.x, 488.492);
    assert_eq!(editable.subpaths[0].nodes[1].node_point.y, -258.172);
    assert_eq!(editable.natural_size.unwrap().width, 488.492);

    let before_invalid = editor.to_bytes().unwrap();
    let mut invalid = settings.clone();
    invalid.move_action.as_mut().unwrap().path.subpaths[0].nodes[1]
        .point
        .x = f32::NAN;
    assert!(
        editor
            .set_slide_build(0, created.object_id, invalid)
            .is_err()
    );
    let mut invalid_start = settings.clone();
    invalid_start.move_action.as_mut().unwrap().path.subpaths[0].nodes[0]
        .point
        .x = 1.0;
    assert!(
        editor
            .set_slide_build(0, created.object_id, invalid_start)
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before_invalid);

    let mut curved = settings;
    let path = &mut curved.move_action.as_mut().unwrap().path;
    path.subpaths[0].nodes = vec![
        KeynoteMotionPathNode::sharp(0.0, 0.0),
        KeynoteMotionPathNode {
            in_control_point: KeynoteMotionPathPoint::new(120.0, -20.0),
            point: KeynoteMotionPathPoint::new(240.0, -150.0),
            out_control_point: KeynoteMotionPathPoint::new(360.0, -280.0),
            node_type: KeynoteMotionPathNodeType::Bezier,
        },
        KeynoteMotionPathNode::sharp(488.492, -258.172),
    ];
    editor
        .set_slide_build(0, created.object_id, curved.clone())
        .unwrap();
    assert_eq!(editor.slide_builds(0).unwrap()[0].settings, curved);

    editor
        .set_slide_build(0, created.object_id, KeynoteBuildSettings::appear_out())
        .unwrap();
    let graph = ObjectGraph::read(editor.package()).unwrap();
    let native: kn::BuildArchive = graph
        .decode_type(created.object_id, BUILD_MESSAGE_TYPE, "KN.BuildArchive")
        .unwrap();
    assert!(native.attributes.action_motion_path_source.is_none());
    assert!(native.attributes.custom_align_to_path.is_none());
    assert!(native.attributes.action_acceleration.is_none());
}

#[test]
fn slide_move_updates_preserve_deep_unknown_path_wire() {
    let mut editor = KeynoteEditor::from_package(test_package()).unwrap();
    let created = editor
        .add_slide_build(0, 5, KeynoteBuildSettings::move_action(400.0, -200.0))
        .unwrap();
    let mut package = editor.into_package();
    let suffixes = [
        unknown_varint(99, 990),
        unknown_varint(98, 980),
        unknown_varint(97, 970),
        unknown_varint(96, 960),
        unknown_varint(95, 950),
    ];
    package
        .update_archive("Index/Slide-4.iwa", |archive| {
            let build = archive.object_mut(created.object_id).unwrap();
            let message = build.messages[0].clone();
            let data = transform_length_delimited_field(&message.data, 4, |attributes| {
                transform_length_delimited_field(attributes, 22, |path_source| {
                    let mut path_source =
                        transform_length_delimited_field(path_source, 8, |editable| {
                            let mut editable = transform_length_delimited_fields_at_path(
                                editable,
                                &[1],
                                |subpath| {
                                    let mut subpath = transform_length_delimited_fields_at_path(
                                        subpath,
                                        &[1],
                                        |node| {
                                            let mut node = transform_length_delimited_field(
                                                node,
                                                2,
                                                |point| {
                                                    let mut point = point.to_vec();
                                                    point.extend_from_slice(&suffixes[4]);
                                                    Ok(point)
                                                },
                                            )?;
                                            node.extend_from_slice(&suffixes[3]);
                                            Ok(node)
                                        },
                                    )?;
                                    subpath.extend_from_slice(&suffixes[2]);
                                    Ok(subpath)
                                },
                            )?;
                            editable.extend_from_slice(&suffixes[1]);
                            Ok(editable)
                        })?;
                    path_source.extend_from_slice(&suffixes[0]);
                    Ok(path_source)
                })
            })?;
            build.replace_message(
                0,
                RawMessage {
                    type_: message.type_,
                    data,
                },
            )?;
            Ok(())
        })
        .unwrap();

    let mut editor = KeynoteEditor::from_package(package).unwrap();
    let mut settings = editor.slide_builds(0).unwrap()[0].settings.clone();
    let path = &mut settings.move_action.as_mut().unwrap().path;
    path.subpaths[0].nodes[1].point.x = 420.0;
    path.subpaths[0].nodes[1].in_control_point.x = 420.0;
    path.subpaths[0].nodes[1].out_control_point.x = 420.0;
    path.natural_width = 420.0;
    editor
        .set_slide_build(0, created.object_id, settings)
        .unwrap();

    let graph = ObjectGraph::read(editor.package()).unwrap();
    let build = graph
        .message_data_type(created.object_id, BUILD_MESSAGE_TYPE, "KN.BuildArchive")
        .unwrap();
    let attributes = repeated_length_delimited_payloads(build, 4).unwrap()[0];
    let path_source = repeated_length_delimited_payloads(attributes, 22).unwrap()[0];
    assert!(path_source.ends_with(&suffixes[0]));
    let editable = repeated_length_delimited_payloads(path_source, 8).unwrap()[0];
    assert!(editable.ends_with(&suffixes[1]));
    let subpath = repeated_length_delimited_payloads(editable, 1).unwrap()[0];
    assert!(subpath.ends_with(&suffixes[2]));
    let node = repeated_length_delimited_payloads(subpath, 1).unwrap()[0];
    assert!(node.ends_with(&suffixes[3]));
    let point = repeated_length_delimited_payloads(node, 2).unwrap()[0];
    assert!(point.ends_with(&suffixes[4]));
}

#[test]
fn slide_move_rejects_missing_duplicate_and_invalid_native_path_fields() {
    for mutation in 0..3 {
        let mut editor = KeynoteEditor::from_package(test_package()).unwrap();
        let settings = KeynoteBuildSettings::move_action(300.0, 100.0);
        let created = editor.add_slide_build(0, 5, settings.clone()).unwrap();
        let mut package = editor.into_package();
        package
            .update_archive("Index/Slide-4.iwa", |archive| {
                let build = archive.object_mut(created.object_id).unwrap();
                let message = build.messages[0].clone();
                let data =
                    transform_length_delimited_field(
                        &message.data,
                        4,
                        |attributes| match mutation {
                            0 => patch_length_delimited_field(attributes, 22, true, None),
                            1 => {
                                let path = repeated_length_delimited_payloads(attributes, 22)?[0];
                                append_repeated_length_delimited_field(attributes, 22, path)
                            },
                            2 => transform_length_delimited_fields_at_path(
                                attributes,
                                &[22, 8, 1, 1],
                                |node| patch_varint_field(node, 4, true, Some(99)),
                            ),
                            _ => unreachable!(),
                        },
                    )?;
                build.replace_message(
                    0,
                    RawMessage {
                        type_: message.type_,
                        data,
                    },
                )?;
                Ok(())
            })
            .unwrap();
        let mut malformed = KeynoteEditor::from_package(package).unwrap();
        assert!(malformed.slide_builds(0).is_err());
        let before = malformed.to_bytes().unwrap();
        assert!(
            malformed
                .set_slide_build(0, created.object_id, settings)
                .is_err()
        );
        assert_eq!(malformed.to_bytes().unwrap(), before);
    }
}

#[test]
fn unsupported_action_timing_updates_preserve_opaque_native_parameters() {
    let mut editor = KeynoteEditor::from_package(test_package()).unwrap();
    let created = editor
        .add_slide_build(
            0,
            5,
            KeynoteBuildSettings::rotate_action(450.0, KeynoteRotationDirection::Clockwise),
        )
        .unwrap();
    let mut package = editor.into_package();
    package
        .update_archive("Index/Slide-4.iwa", |archive| {
            let build = archive.object_mut(created.object_id).unwrap();
            let message = build.messages[0].clone();
            let data = transform_length_delimited_field(&message.data, 4, |attributes| {
                transform_length_delimited_field(attributes, 18, |animation| {
                    patch_length_delimited_field(animation, 2, true, Some(b"apple:action-future"))
                })
            })?;
            build.replace_message(
                0,
                RawMessage {
                    type_: message.type_,
                    data,
                },
            )?;
            Ok(())
        })
        .unwrap();
    let mut editor = KeynoteEditor::from_package(package).unwrap();
    let mut settings = editor.slide_builds(0).unwrap()[0].settings.clone();
    assert_eq!(settings.effect, "apple:action-future");
    assert!(settings.rotation.is_none());
    settings.duration = 3.25;
    editor
        .set_slide_build(0, created.object_id, settings.clone())
        .unwrap();
    let graph = ObjectGraph::read(editor.package()).unwrap();
    let native: kn::BuildArchive = graph
        .decode_type(created.object_id, BUILD_MESSAGE_TYPE, "KN.BuildArchive")
        .unwrap();
    assert_eq!(native.attributes.action_rotation_angle, Some(450.0));
    assert!(native.attributes.action_rotation_direction.is_some());
    assert!(native.attributes.action_acceleration.is_some());
    assert_eq!(editor.slide_builds(0).unwrap()[0].settings, settings);

    let before = editor.to_bytes().unwrap();
    assert!(
        editor
            .set_slide_build(0, created.object_id, KeynoteBuildSettings::appear_out())
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn slide_build_start_modes_map_to_native_chunks_and_guard_sequence_edges() {
    let mut editor = KeynoteEditor::from_package(test_package()).unwrap();
    let mut first_settings = KeynoteBuildSettings::appear_in();
    first_settings.start = BuildStart::AfterTransition;
    first_settings.delay = 0.25;
    let first = editor.add_slide_build(0, 5, first_settings).unwrap();
    assert_eq!(
        editor.slide_builds(0).unwrap()[0].settings.start,
        BuildStart::AfterTransition
    );

    let mut second_settings = KeynoteBuildSettings::appear_in();
    second_settings.start = BuildStart::WithPrevious;
    let second = editor.add_slide_build(0, 6, second_settings).unwrap();
    let builds = editor.slide_builds(0).unwrap();
    assert_eq!(builds[1].settings.start, BuildStart::WithPrevious);
    assert_eq!(builds[1].chunks[0].automatic, Some(true));
    assert_eq!(builds[1].chunks[0].referent, Some(false));

    let before_invalid = editor.to_bytes().unwrap();
    assert!(editor.move_slide_build(0, second.object_id, 0).is_err());
    assert_eq!(editor.to_bytes().unwrap(), before_invalid);
    assert!(editor.remove_slide_build(0, first.object_id).is_err());
    assert_eq!(editor.to_bytes().unwrap(), before_invalid);

    let mut after_previous = builds[1].settings.clone();
    after_previous.start = BuildStart::AfterPrevious;
    after_previous.delay = 0.5;
    editor
        .set_slide_build(0, second.object_id, after_previous)
        .unwrap();
    let after_builds = editor.slide_builds(0).unwrap();
    let second_after = &after_builds[1];
    assert_eq!(second_after.settings.start, BuildStart::AfterPrevious);
    assert_eq!(second_after.chunks[0].automatic, Some(true));
    assert_eq!(second_after.chunks[0].referent, Some(true));

    let mut invalid_delay = second_after.settings.clone();
    invalid_delay.start = BuildStart::OnClick;
    assert!(
        editor
            .set_slide_build(0, second.object_id, invalid_delay)
            .is_err()
    );

    let mut on_click = second_after.settings.clone();
    on_click.start = BuildStart::OnClick;
    on_click.delay = 0.0;
    editor
        .set_slide_build(0, second.object_id, on_click)
        .unwrap();
    let click_builds = editor.slide_builds(0).unwrap();
    let second_click = &click_builds[1];
    assert_eq!(second_click.settings.start, BuildStart::OnClick);
    assert_eq!(second_click.chunks[0].automatic, Some(false));
    assert_eq!(second_click.chunks[0].referent, Some(true));

    let mut empty = KeynoteEditor::from_package(test_package()).unwrap();
    let mut invalid_first = KeynoteBuildSettings::appear_in();
    invalid_first.start = BuildStart::AfterPrevious;
    assert!(empty.add_slide_build(0, 5, invalid_first).is_err());
    assert!(empty.slide_builds(0).unwrap().is_empty());

    let mut invalid_second = KeynoteBuildSettings::appear_in();
    invalid_second.start = BuildStart::AfterTransition;
    assert!(editor.add_slide_build(0, 5, invalid_second).is_err());
}

#[test]
fn slide_build_reader_rejects_automatic_first_event_without_transition_semantics() {
    let mut editor = KeynoteEditor::from_package(test_package()).unwrap();
    let build = editor
        .add_slide_build(0, 5, KeynoteBuildSettings::appear_in())
        .unwrap();
    let chunk_id = build.chunks[0].object_id;
    let mut package = editor.into_package();
    package
        .update_archive("Index/Slide-4.iwa", |archive| {
            let chunk = archive.object_mut(chunk_id).unwrap();
            let message = chunk.messages[0].clone();
            let data = patch_varint_field(&message.data, 5, true, Some(1))?;
            let data = patch_varint_field(&data, 6, true, Some(0))?;
            chunk.replace_message(
                0,
                RawMessage {
                    type_: message.type_,
                    data,
                },
            )?;
            Ok(())
        })
        .unwrap();

    let malformed = KeynoteEditor::from_package(package).unwrap();
    assert!(malformed.slide_builds(0).is_err());
}

#[test]
fn slide_build_reorder_is_transactional_and_byte_exact_when_restored() {
    let mut editor = KeynoteEditor::from_package(test_package()).unwrap();
    let first = editor
        .add_slide_build(0, 5, KeynoteBuildSettings::appear_in())
        .unwrap();
    let second = editor
        .add_slide_build(0, 6, KeynoteBuildSettings::appear_in())
        .unwrap();
    let baseline = editor.to_bytes().unwrap();

    editor.move_slide_build(0, first.object_id, 1).unwrap();
    assert_eq!(
        editor
            .slide_builds(0)
            .unwrap()
            .iter()
            .map(|build| build.object_id)
            .collect::<Vec<_>>(),
        [second.object_id, first.object_id]
    );
    let graph = ObjectGraph::read(editor.package()).unwrap();
    let slide: kn::SlideArchive = graph.decode(4, "KN.SlideArchive").unwrap();
    assert_eq!(
        slide
            .build_chunks
            .iter()
            .map(|reference| reference.identifier)
            .collect::<Vec<_>>(),
        [second.chunks[0].object_id, first.chunks[0].object_id]
    );

    let reordered = editor.to_bytes().unwrap();
    assert!(editor.move_slide_build(0, first.object_id, 2).is_err());
    assert_eq!(editor.to_bytes().unwrap(), reordered);
    assert!(
        editor
            .reorder_slide_builds(0, &[second.object_id, second.object_id])
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), reordered);
    assert!(editor.move_slide_build(1, first.object_id, 0).is_err());
    assert_eq!(editor.to_bytes().unwrap(), reordered);

    editor.move_slide_build(0, first.object_id, 0).unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);
}

#[test]
fn slide_build_reorder_preserves_unknown_reference_wire() {
    let mut editor = KeynoteEditor::from_package(test_package()).unwrap();
    let first = editor
        .add_slide_build(0, 5, KeynoteBuildSettings::appear_in())
        .unwrap();
    let second = editor
        .add_slide_build(0, 6, KeynoteBuildSettings::appear_in())
        .unwrap();
    let mut package = editor.into_package();
    package
        .update_archive("Index/Slide-4.iwa", |archive| {
            let slide = archive.object_mut(4).unwrap();
            let original = slide.messages[0].data.as_slice();
            let mut data = original.to_vec();
            for field in [2, 43] {
                let replacements = repeated_length_delimited_payloads(&data, field)?
                    .into_iter()
                    .enumerate()
                    .map(|(index, payload)| {
                        let mut payload = payload.to_vec();
                        append_unknown_varint(&mut payload, 99, 9_900 + index as u64);
                        payload
                    })
                    .collect::<Vec<_>>();
                data = rewrite_repeated_length_delimited_fields(&data, field, &replacements)?;
            }
            slide.replace_message(0, RawMessage { type_: 5, data })?;
            Ok(())
        })
        .unwrap();
    let before = package
        .archive("Index/Slide-4.iwa")
        .unwrap()
        .object(4)
        .unwrap()
        .messages[0]
        .data
        .clone();
    let before_builds = repeated_length_delimited_payloads(&before, 2)
        .unwrap()
        .into_iter()
        .map(Vec::from)
        .collect::<Vec<_>>();
    let before_chunks = repeated_length_delimited_payloads(&before, 43)
        .unwrap()
        .into_iter()
        .map(Vec::from)
        .collect::<Vec<_>>();

    let mut editor = KeynoteEditor::from_package(package).unwrap();
    editor
        .reorder_slide_builds(0, &[second.object_id, first.object_id])
        .unwrap();
    let graph = ObjectGraph::read(editor.package()).unwrap();
    let slide = graph.objects.get(&4).unwrap()[0].data.as_slice();
    assert_eq!(
        repeated_length_delimited_payloads(slide, 2).unwrap(),
        [before_builds[1].as_slice(), before_builds[0].as_slice()]
    );
    assert_eq!(
        repeated_length_delimited_payloads(slide, 43).unwrap(),
        [before_chunks[1].as_slice(), before_chunks[0].as_slice()]
    );
}

#[test]
fn build_updates_preserve_unknown_wire_and_normalize_native_merge_payloads() {
    let mut editor = KeynoteEditor::from_package(test_package()).unwrap();
    let created = editor
        .add_slide_build(0, 5, KeynoteBuildSettings::appear_in())
        .unwrap();
    let mut package = editor.into_package();
    let mut build_suffix = Vec::new();
    append_unknown_varint(&mut build_suffix, 99, 9_900);
    let mut chunk_suffix = Vec::new();
    append_unknown_varint(&mut chunk_suffix, 98, 9_800);
    package
        .update_archive("Index/Slide-4.iwa", |archive| {
            let build = archive.object_mut(created.object_id).unwrap();
            let mut message = build.messages[0].clone();
            message.data.extend_from_slice(&build_suffix);
            build.replace_message(0, message)?;
            build.push_message(RawMessage {
                type_: 0,
                data: kn::BuildAttributesArchive {
                    animation_attributes: Some(kn::AnimationAttributesArchive {
                        effect: Some("apple:appear".to_owned()),
                        ..Default::default()
                    }),
                    ..Default::default()
                }
                .encode_to_vec(),
            })?;
            build.archive_info.should_merge = Some(true);

            let chunk = archive.object_mut(created.chunks[0].object_id).unwrap();
            let mut message = chunk.messages[0].clone();
            message.data.extend_from_slice(&chunk_suffix);
            chunk.replace_message(0, message)?;
            Ok(())
        })
        .unwrap();

    let mut editor = KeynoteEditor::from_package(package).unwrap();
    let mut settings = KeynoteBuildSettings::appear_in();
    settings.effect = "apple:unmapped-test-effect".to_owned();
    settings.duration = 2.5;
    editor
        .set_slide_build(0, created.object_id, settings)
        .unwrap();
    let archive = editor.package().archive("Index/Slide-4.iwa").unwrap();
    let build = archive.object(created.object_id).unwrap();
    assert_eq!(
        build
            .messages
            .iter()
            .map(|message| message.type_)
            .collect::<Vec<_>>(),
        [BUILD_MESSAGE_TYPE]
    );
    assert_eq!(build.archive_info.should_merge, None);
    assert!(build.messages[0].data.ends_with(&build_suffix));
    let chunk = archive.object(created.chunks[0].object_id).unwrap();
    assert!(chunk.messages[0].data.ends_with(&chunk_suffix));
}

#[test]
fn build_crud_tracks_only_native_uuid_objects_and_releases_highwater() {
    let mut package = test_package();
    package
        .replace_archive(
            PACKAGE_METADATA_ENTRY,
            &Archive {
                objects: vec![object(
                    20,
                    PACKAGE_METADATA_MESSAGE_TYPE,
                    PackageMetadata {
                        last_object_identifier: 20,
                        components: vec![ComponentInfo {
                            identifier: 4,
                            preferred_locator: "Slide-4".to_owned(),
                            ..Default::default()
                        }],
                        ..Default::default()
                    },
                )],
            },
        )
        .unwrap();
    let mut editor = KeynoteEditor::from_package(package).unwrap();
    let build = editor
        .add_slide_build(0, 5, KeynoteBuildSettings::appear_in())
        .unwrap();
    assert_eq!(build.object_id, 21);
    assert_eq!(build.chunks[0].object_id, 22);
    let mapped = component_uuid_identifiers(editor.package(), 4)
        .unwrap()
        .unwrap();
    assert_eq!(mapped, HashSet::from([21]));
    let metadata = editor.package().archive(PACKAGE_METADATA_ENTRY).unwrap();
    let metadata =
        PackageMetadata::decode(metadata.object(20).unwrap().messages[0].data.as_slice()).unwrap();
    assert_eq!(metadata.last_object_identifier, 22);

    editor.remove_slide_build(0, 21).unwrap();
    assert!(
        component_uuid_identifiers(editor.package(), 4)
            .unwrap()
            .unwrap()
            .is_empty()
    );
    let metadata = editor.package().archive(PACKAGE_METADATA_ENTRY).unwrap();
    let metadata =
        PackageMetadata::decode(metadata.object(20).unwrap().messages[0].data.as_slice()).unwrap();
    assert_eq!(metadata.last_object_identifier, 20);
}

#[test]
fn slide_owned_text_storage_crud_covers_placeholders_and_text_boxes() {
    let mut editor = KeynoteEditor::from_package(test_package_with_text_box()).unwrap();
    let text = editor.slide_text_storages(0).unwrap();
    assert_eq!(
        text.iter().map(|item| item.role).collect::<Vec<_>>(),
        [
            KeynoteSlideTextRole::Title,
            KeynoteSlideTextRole::Body,
            KeynoteSlideTextRole::TextBox,
        ]
    );
    assert_eq!(text[2].drawable_object_id, 17);
    assert_eq!(text[2].storage.id, TextStorageId::new(18).unwrap());
    assert_eq!(text[2].storage.storage.text(), "Independent text box");

    editor
        .replace_slide_text_storage(0, 17, 12..16, "shape 🚀")
        .unwrap();
    assert_eq!(
        editor.slide_text_storages(0).unwrap()[2]
            .storage
            .storage
            .text(),
        "Independent shape 🚀 box"
    );
    editor.set_slide_text_storage(0, 17, "Replacement").unwrap();
    assert_eq!(
        editor.slide_text_storages(0).unwrap()[2]
            .storage
            .storage
            .text(),
        "Replacement"
    );
    editor.clear_slide_text_storage(0, 17).unwrap();
    assert!(
        editor.slide_text_storages(0).unwrap()[2]
            .storage
            .storage
            .is_empty()
    );

    let before = editor.to_bytes().unwrap();
    assert!(editor.set_slide_text_storage(0, 11, "wrong slide").is_err());
    assert!(editor.set_slide_text_storage(2, 17, "missing").is_err());
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn slide_text_storage_updates_preserve_unknown_wire_and_restore_exactly() {
    let mut package = test_package_with_text_box();
    package
        .update_archive("Index/Slide-4.iwa", |archive| {
            let object = archive.object_mut(18).unwrap();
            let mut message = object.messages[0].clone();
            append_unknown_varint(&mut message.data, 99, 990);
            Ok(object.replace_message(0, message).map(|_| ())?)
        })
        .unwrap();
    let mut editor = KeynoteEditor::from_package(package).unwrap();
    let baseline = editor
        .package()
        .archive("Index/Slide-4.iwa")
        .unwrap()
        .to_bytes()
        .unwrap();
    editor
        .set_slide_text_storage(0, 17, "Temporary 東京")
        .unwrap();
    editor
        .set_slide_text_storage(0, 17, "Independent text box")
        .unwrap();
    assert_eq!(
        editor
            .package()
            .archive("Index/Slide-4.iwa")
            .unwrap()
            .to_bytes()
            .unwrap(),
        baseline
    );
}

#[test]
fn ambiguous_slide_text_ownership_fails_transactionally() {
    let mut package = test_package_with_text_box();
    package
        .update_archive("Index/Slide-4.iwa", |archive| {
            let object = archive.object_mut(4).unwrap();
            let mut slide = kn::SlideArchive::decode(object.messages[0].data.as_slice()).unwrap();
            slide.owned_drawables.push(reference(17));
            object.replace_message(
                0,
                RawMessage {
                    type_: 5,
                    data: slide.encode_to_vec(),
                },
            )?;
            Ok(())
        })
        .unwrap();
    let mut editor = KeynoteEditor::from_package(package).unwrap();
    let before = editor.to_bytes().unwrap();
    assert!(editor.slide_text_storages(0).is_err());
    assert!(editor.set_slide_text_storage(0, 17, "rejected").is_err());
    assert_eq!(editor.to_bytes().unwrap(), before);

    let mut package = test_package_with_text_box();
    package
        .update_archive("Index/Slide-10.iwa", |archive| {
            let slide_object = archive.object_mut(10).unwrap();
            let mut slide =
                kn::SlideArchive::decode(slide_object.messages[0].data.as_slice()).unwrap();
            slide.owned_drawables.push(reference(19));
            slide_object.replace_message(
                0,
                RawMessage {
                    type_: 5,
                    data: slide.encode_to_vec(),
                },
            )?;
            Ok(archive.insert_object(object(
                19,
                2011,
                tswp::ShapeInfoArchive {
                    owned_storage: Some(reference(18)),
                    ..Default::default()
                },
            ))?)
        })
        .unwrap();
    let mut editor = KeynoteEditor::from_package(package).unwrap();
    let before = editor.to_bytes().unwrap();
    assert!(editor.slide_text_storages(0).is_err());
    assert!(editor.set_slide_text_storage(0, 17, "rejected").is_err());
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn ordinary_text_box_duplicate_delete_is_independent_and_exact() {
    let mut editor = KeynoteEditor::from_package(test_package_with_text_box()).unwrap();
    let baseline = editor.to_bytes().unwrap();

    let created = editor.duplicate_slide_text_box(0, 17, "Clone 🚀").unwrap();
    assert_eq!(created.drawable_object_id, 23);
    assert_eq!(created.storage.id, TextStorageId::new(26).unwrap());
    assert_eq!(created.storage.storage.text(), "Clone 🚀");
    assert_eq!(
        editor.text_box_graph(0, 23).unwrap().object_ids,
        [23, 24, 25, 26]
    );

    let archive = editor.package().archive("Index/Slide-4.iwa").unwrap();
    let slide =
        kn::SlideArchive::decode(archive.object(4).unwrap().messages[0].data.as_slice()).unwrap();
    assert_eq!(
        slide
            .owned_drawables
            .iter()
            .map(|reference| reference.identifier)
            .collect::<Vec<_>>(),
        [5, 6, 17, 23]
    );
    assert_eq!(
        slide
            .drawables_z_order
            .iter()
            .map(|reference| reference.identifier)
            .collect::<Vec<_>>(),
        [5, 6, 17, 23]
    );
    let shape =
        tswp::ShapeInfoArchive::decode(archive.object(23).unwrap().messages[0].data.as_slice())
            .unwrap();
    let position = shape.super_.super_.geometry.unwrap().position.unwrap();
    assert_eq!((position.x, position.y), (110.0, 110.0));

    editor.set_slide_text_storage(0, 23, "Changed").unwrap();
    assert_eq!(
        editor
            .slide_text_storages(0)
            .unwrap()
            .into_iter()
            .find(|item| item.drawable_object_id == 17)
            .unwrap()
            .storage
            .storage
            .text(),
        "Independent text box"
    );
    let removed = editor.remove_slide_text_box(0, 23).unwrap();
    assert_eq!(removed.text.storage.storage.text(), "Changed");
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    let before = editor.to_bytes().unwrap();
    assert!(
        editor
            .duplicate_slide_text_box(0, 5, "placeholder")
            .is_err()
    );
    assert!(editor.remove_slide_text_box(0, 5).is_err());
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn ordinary_text_box_geometry_updates_are_guarded_and_byte_exact() {
    let mut editor = KeynoteEditor::from_package(test_package_with_text_box()).unwrap();
    let original = editor.slide_text_box_geometry(0, 17).unwrap();
    assert_eq!(
        original,
        DrawableGeometry {
            position: Some(DrawablePoint { x: 100.0, y: 100.0 }),
            size: Some(DrawableSize {
                width: 200.0,
                height: 60.0,
            }),
            flags: Some(0),
            angle: Some(0.0),
        }
    );
    let baseline = editor.to_bytes().unwrap();
    let changed = DrawableGeometry {
        position: Some(DrawablePoint {
            x: 150.5,
            y: 140.25,
        }),
        size: Some(DrawableSize {
            width: 360.0,
            height: 90.0,
        }),
        flags: Some(3),
        angle: Some(0.5),
    };
    editor.set_slide_text_box_geometry(0, 17, changed).unwrap();
    assert_eq!(editor.slide_text_box_geometry(0, 17).unwrap(), changed);
    assert_eq!(
        editor
            .slide_text_storages(0)
            .unwrap()
            .into_iter()
            .find(|text| text.drawable_object_id == 17)
            .unwrap()
            .storage
            .storage
            .text(),
        "Independent text box"
    );
    editor.set_slide_text_box_geometry(0, 17, original).unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    let before = editor.to_bytes().unwrap();
    let mut invalid = original;
    invalid.angle = Some(f32::INFINITY);
    assert!(editor.set_slide_text_box_geometry(0, 17, invalid).is_err());
    assert!(editor.slide_text_box_geometry(0, 5).is_err());
    assert!(editor.slide_text_box_geometry(1, 17).is_err());
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn ordinary_text_box_properties_updates_are_guarded_and_byte_exact() {
    let mut editor = KeynoteEditor::from_package(test_package_with_text_box()).unwrap();
    let original = editor.slide_text_box_properties(0, 17).unwrap();
    assert_eq!(original, DrawableProperties::default());
    let baseline = editor.to_bytes().unwrap();
    let changed = DrawableProperties {
        hyperlink_url: Some("https://example.test/keynote-text-box".to_owned()),
        locked: Some(true),
        aspect_ratio_locked: Some(true),
        accessibility_description: Some("Accessible Keynote text box ✨".to_owned()),
    };

    editor
        .set_slide_text_box_properties(0, 17, changed.clone())
        .unwrap();
    assert_eq!(editor.slide_text_box_properties(0, 17).unwrap(), changed);
    assert_eq!(
        editor
            .slide_text_storages(0)
            .unwrap()
            .into_iter()
            .find(|text| text.drawable_object_id == 17)
            .unwrap()
            .storage
            .storage
            .text(),
        "Independent text box"
    );
    editor
        .set_slide_text_box_properties(0, 17, original)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    let before = editor.to_bytes().unwrap();
    assert!(
        editor
            .set_slide_text_box_properties(0, 5, DrawableProperties::default())
            .is_err()
    );
    assert!(
        editor
            .set_slide_text_box_properties(1, 17, DrawableProperties::default())
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn text_box_duplicate_delete_tracks_package_highwater_and_slide_uuids() {
    let mut package = test_package_with_text_box();
    let uuid_entry = |identifier| ObjectUuidMapEntry {
        identifier,
        uuid: Uuid {
            lower: identifier,
            upper: identifier + 1_000,
        },
    };
    let metadata = PackageMetadata {
        last_object_identifier: 22,
        components: vec![ComponentInfo {
            identifier: 4,
            preferred_locator: "Slide-4".to_owned(),
            object_uuid_map_entries: [17, 21, 22, 18].into_iter().map(uuid_entry).collect(),
            ..Default::default()
        }],
        ..Default::default()
    };
    package
        .replace_archive(
            PACKAGE_METADATA_ENTRY,
            &Archive {
                objects: vec![object(20, PACKAGE_METADATA_MESSAGE_TYPE, metadata)],
            },
        )
        .unwrap();

    let mut editor = KeynoteEditor::from_package(package).unwrap();
    let baseline = editor.to_bytes().unwrap();
    let created = editor
        .duplicate_slide_text_box(0, 17, "Metadata clone")
        .unwrap();
    assert_eq!(created.drawable_object_id, 23);
    let graph = editor
        .text_box_graph(0, created.drawable_object_id)
        .unwrap();
    assert_eq!(graph.object_ids, [23, 24, 25, 26]);
    assert_eq!(graph.uuid_object_ids, [23, 24, 25, 26]);

    let archive = editor.package().archive(PACKAGE_METADATA_ENTRY).unwrap();
    let metadata =
        PackageMetadata::decode(archive.object(20).unwrap().messages[0].data.as_slice()).unwrap();
    assert_eq!(metadata.last_object_identifier, 26);
    let entries = &metadata.components[0].object_uuid_map_entries;
    assert_eq!(entries.len(), 8);
    assert_eq!(
        entries
            .iter()
            .map(|entry| entry.identifier)
            .collect::<HashSet<_>>(),
        [17, 18, 21, 22, 23, 24, 25, 26].into_iter().collect()
    );
    assert_eq!(
        entries
            .iter()
            .map(|entry| (entry.uuid.lower, entry.uuid.upper))
            .collect::<HashSet<_>>()
            .len(),
        entries.len()
    );

    editor
        .remove_slide_text_box(0, created.drawable_object_id)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);
}

#[test]
fn text_box_graph_crud_preserves_unknowns_and_rejects_external_owners() {
    let mut package = test_package_with_text_box();
    package
        .update_archive("Index/Slide-4.iwa", |archive| {
            for identifier in [17, 21, 22, 18] {
                let object = archive.object_mut(identifier).unwrap();
                let mut message = object.messages[0].clone();
                append_unknown_varint(&mut message.data, 99, identifier);
                object.replace_message(0, message)?;
            }
            Ok(())
        })
        .unwrap();
    let mut editor = KeynoteEditor::from_package(package).unwrap();
    let baseline = editor.to_bytes().unwrap();
    let created = editor
        .duplicate_slide_text_box(0, 17, "Unknown clone")
        .unwrap();
    let archive = editor.package().archive("Index/Slide-4.iwa").unwrap();
    for (source, cloned) in [(17, 23), (21, 24), (22, 25), (18, 26)] {
        let source = &archive.object(source).unwrap().messages[0].data;
        let cloned = &archive.object(cloned).unwrap().messages[0].data;
        let source_field = crate::wire::parse_wire_fields(source)
            .unwrap()
            .into_iter()
            .find(|field| field.number() == 99)
            .unwrap();
        let cloned_field = crate::wire::parse_wire_fields(cloned)
            .unwrap()
            .into_iter()
            .find(|field| field.number() == 99)
            .unwrap();
        assert_eq!(
            &source[source_field.start()..source_field.end()],
            &cloned[cloned_field.start()..cloned_field.end()]
        );
    }
    editor
        .remove_slide_text_box(0, created.drawable_object_id)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    let mut package = test_package_with_text_box();
    package
        .update_archive("Index/Slide-4.iwa", |archive| {
            let mut owner = object(19, 9_999, Vec::new());
            owner.archive_info.message_infos[0]
                .object_references
                .push(17);
            Ok(archive.insert_object(owner)?)
        })
        .unwrap();
    let mut editor = KeynoteEditor::from_package(package).unwrap();
    let before = editor.to_bytes().unwrap();
    assert!(editor.duplicate_slide_text_box(0, 17, "rejected").is_err());
    assert!(editor.remove_slide_text_box(0, 17).is_err());
    assert_eq!(editor.to_bytes().unwrap(), before);

    let mut package = test_package_with_text_box();
    package
        .update_archive("Index/Slide-4.iwa", |archive| {
            let slide_object = archive.object_mut(4).unwrap();
            let mut slide =
                kn::SlideArchive::decode(slide_object.messages[0].data.as_slice()).unwrap();
            slide.drawables_z_order.clear();
            slide_object.replace_message(
                0,
                RawMessage {
                    type_: 5,
                    data: slide.encode_to_vec(),
                },
            )?;
            Ok(())
        })
        .unwrap();
    assert!(KeynoteEditor::from_package(package).is_err());
}

#[test]
fn show_settings_and_transitions_are_transactional() {
    let mut editor = KeynoteEditor::from_package(test_package()).unwrap();
    let mut settings = editor.show_settings().unwrap();
    settings.set_size(Size::new(1_920.0, 1_080.0).unwrap());
    settings.set_slide_numbers_visible(Some(true));
    settings.set_loop_presentation(Some(true));
    settings.set_mode(Some(Mode::SelfPlaying)).unwrap();
    settings.set_autoplay_transition_delay(Some(Seconds::new(3.5).unwrap()));
    settings.set_autoplay_build_delay(Some(Seconds::new(1.25).unwrap()));
    settings.set_idle_timer_active(Some(true));
    settings.set_idle_timer_delay(Some(Seconds::new(60.0).unwrap()));
    settings.set_automatically_plays_upon_open(Some(false));
    let before = editor.to_bytes().unwrap();
    assert!(Size::new(f32::NAN, 1_080.0).is_err());
    assert_eq!(editor.to_bytes().unwrap(), before);
    assert!(settings.set_mode(Some(Mode::Unknown(1))).is_err());
    assert_eq!(editor.to_bytes().unwrap(), before);
    editor.set_show_settings(settings).unwrap();
    assert_eq!(editor.show_settings().unwrap(), settings);
    for mode in [Mode::Normal, Mode::LinksOnly, Mode::Unknown(19)] {
        settings.set_mode(Some(mode)).unwrap();
        editor.set_show_settings(settings).unwrap();
        let reparsed = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(reparsed.show_settings().unwrap(), settings);
    }

    let mut transition = editor.slides().unwrap()[0].transition.clone().unwrap();
    transition.set_duration(Some(2.5)).unwrap();
    transition.set_delay(Some(1.0)).unwrap();
    transition.set_is_automatic(Some(true));
    let before = editor.to_bytes().unwrap();
    let mut invalid_transition = transition.clone();
    assert!(
        invalid_transition
            .set_duration(Some(f64::INFINITY))
            .is_err()
    );
    assert_eq!(invalid_transition, transition);
    assert_eq!(editor.to_bytes().unwrap(), before);
    editor.set_slide_transition(0, transition.clone()).unwrap();
    assert_eq!(
        editor.slides().unwrap()[0].transition.as_ref(),
        Some(&transition)
    );

    let reparsed = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(reparsed.show_settings().unwrap(), settings);
    assert_eq!(
        reparsed.slides().unwrap()[0].transition.as_ref(),
        Some(&transition)
    );
}

#[test]
fn soundtrack_settings_are_typed_transactional_and_wire_exact() {
    let mut editor = KeynoteEditor::from_package(test_package_with_soundtrack()).unwrap();
    let original = SoundtrackSettings::new(Some(1.0), Some(SoundtrackMode::PlayOnce)).unwrap();
    assert_eq!(editor.soundtrack_settings().unwrap(), Some(original));
    let original_bytes = editor.to_bytes().unwrap();
    let original_graph = ObjectGraph::read(editor.package()).unwrap();
    let original_wire = original_graph
        .message_data_type(
            TEST_SOUNDTRACK_ID,
            TEST_SOUNDTRACK_MESSAGE_TYPE,
            "KN.Soundtrack",
        )
        .unwrap();
    let original_unknown = crate::wire::parse_wire_fields(original_wire)
        .unwrap()
        .into_iter()
        .find(|field| field.number() == 99)
        .map(|field| original_wire[field.start()..field.end()].to_vec())
        .unwrap();

    for invalid_update in [
        SoundtrackSettings::new(Some(f64::NAN), original.mode()),
        SoundtrackSettings::new(Some(1.01), original.mode()),
        SoundtrackSettings::new(
            original.volume(),
            Some(SoundtrackMode::Unknown(TEST_SOUNDTRACK_LOOP_MODE)),
        ),
    ] {
        assert!(invalid_update.is_err());
        assert_eq!(editor.to_bytes().unwrap(), original_bytes);
    }

    let changed = SoundtrackSettings::new(Some(0.35), Some(SoundtrackMode::Loop)).unwrap();
    editor.set_soundtrack_settings(changed).unwrap();
    assert_eq!(editor.soundtrack_settings().unwrap(), Some(changed));
    let changed_bytes = editor.to_bytes().unwrap();
    editor.set_soundtrack_settings(changed).unwrap();
    assert_eq!(editor.to_bytes().unwrap(), changed_bytes);
    let graph = ObjectGraph::read(editor.package()).unwrap();
    let native: kn::Soundtrack = graph
        .decode_type(
            TEST_SOUNDTRACK_ID,
            TEST_SOUNDTRACK_MESSAGE_TYPE,
            "KN.Soundtrack",
        )
        .unwrap();
    assert_eq!(
        native
            .movie_media
            .iter()
            .map(|reference| reference.identifier)
            .collect::<Vec<_>>(),
        TEST_SOUNDTRACK_MEDIA_IDS
    );
    let changed_wire = graph
        .message_data_type(
            TEST_SOUNDTRACK_ID,
            TEST_SOUNDTRACK_MESSAGE_TYPE,
            "KN.Soundtrack",
        )
        .unwrap();
    let changed_unknown = crate::wire::parse_wire_fields(changed_wire)
        .unwrap()
        .into_iter()
        .find(|field| field.number() == 99)
        .map(|field| changed_wire[field.start()..field.end()].to_vec())
        .unwrap();
    assert_eq!(changed_unknown, original_unknown);
    let archive = editor.package().archive("Index/Document.iwa").unwrap();
    assert_eq!(
        archive
            .object(TEST_SOUNDTRACK_ID)
            .unwrap()
            .archive_info
            .message_infos[0]
            .data_references,
        TEST_SOUNDTRACK_MEDIA_IDS
    );

    let future =
        SoundtrackSettings::new(changed.volume(), Some(SoundtrackMode::Unknown(19))).unwrap();
    editor.set_soundtrack_settings(future).unwrap();
    assert_eq!(editor.soundtrack_settings().unwrap(), Some(future));
    editor.set_soundtrack_settings(original).unwrap();
    assert_eq!(editor.to_bytes().unwrap(), original_bytes);
}

#[test]
fn soundtrack_settings_handle_absent_and_malformed_objects_transactionally() {
    let mut editor = KeynoteEditor::from_package(test_package()).unwrap();
    assert_eq!(editor.soundtrack_settings().unwrap(), None);
    let before = editor.to_bytes().unwrap();
    assert!(
        editor
            .set_soundtrack_settings(
                SoundtrackSettings::new(Some(1.0), Some(SoundtrackMode::PlayOnce)).unwrap()
            )
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before);

    let mut package = test_package_with_soundtrack();
    package
        .update_archive("Index/Document.iwa", |archive| {
            let object = archive.object_mut(TEST_SOUNDTRACK_ID).unwrap();
            let mut message = object.messages[0].clone();
            append_unknown_varint(
                &mut message.data,
                TEST_SOUNDTRACK_MODE_FIELD,
                TEST_SOUNDTRACK_LOOP_MODE as u64,
            );
            Ok(object.replace_message(0, message).map(|_| ())?)
        })
        .unwrap();
    let mut editor = KeynoteEditor::from_package(package).unwrap();
    let before = editor.to_bytes().unwrap();
    assert!(editor.soundtrack_settings().is_err());
    assert!(
        editor
            .set_soundtrack_settings(
                SoundtrackSettings::new(Some(0.5), Some(SoundtrackMode::Loop)).unwrap()
            )
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before);

    let mut package = test_package_with_soundtrack();
    package
        .update_archive("Index/Document.iwa", |archive| {
            let object = archive.object_mut(2).unwrap();
            let mut message = object.messages[0].clone();
            message.data = append_repeated_length_delimited_field(
                &message.data,
                TEST_SHOW_SOUNDTRACK_FIELD,
                &reference(TEST_SOUNDTRACK_ID).encode_to_vec(),
            )?;
            Ok(object.replace_message(0, message).map(|_| ())?)
        })
        .unwrap();
    let mut editor = KeynoteEditor::from_package(package).unwrap();
    let before = editor.to_bytes().unwrap();
    assert!(editor.soundtrack_settings().is_err());
    assert!(
        editor
            .set_soundtrack_settings(
                SoundtrackSettings::new(Some(0.5), Some(SoundtrackMode::Loop)).unwrap()
            )
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn soundtrack_item_crud_is_ordered_transactional_and_wire_exact() {
    const FIRST: &[u8] = b"FORM\0\0\0\x04AIFFfirst";
    const SECOND: &[u8] = b"FORM\0\0\0\x05AIFFsecond";
    const REPLACEMENT: &[u8] = b"FORM\0\0\0\x06AIFCreplacement";

    let mut editor = KeynoteEditor::from_package(test_package_with_empty_soundtrack()).unwrap();
    let baseline = editor.to_bytes().unwrap();
    assert!(editor.soundtrack_items().unwrap().is_empty());

    let first = editor.add_soundtrack_item("first.aiff", FIRST).unwrap();
    assert_eq!(first.index, 0);
    assert_eq!(
        first.asset.media_type,
        litchi_iwa_common::media::Type::Audio
    );
    let second = editor
        .insert_soundtrack_item(0, "second.aiff", SECOND)
        .unwrap();
    let items = editor.soundtrack_items().unwrap();
    assert_eq!(second.index, 0);
    assert_eq!(items.len(), 2);
    assert_eq!(items[0].asset.preferred_filename, "second.aiff");
    assert_eq!(items[1].asset.preferred_filename, "first.aiff");
    assert_eq!(items[0].asset.component_reference_count, 1);
    assert_eq!(items[0].asset.message_reference_count, 1);

    editor.move_soundtrack_item(0, 1).unwrap();
    let moved = editor.soundtrack_items().unwrap();
    assert_eq!(moved[0].asset.data_identifier, first.asset.data_identifier);
    assert_eq!(moved[1].asset.data_identifier, second.asset.data_identifier);
    editor.move_soundtrack_item(1, 0).unwrap();
    assert_eq!(editor.soundtrack_items().unwrap(), items);

    let replaced = editor
        .replace_soundtrack_item(1, "replacement.aif", REPLACEMENT)
        .unwrap();
    assert_eq!(replaced.index, 1);
    assert_eq!(replaced.asset.preferred_filename, "replacement.aif");
    assert!(
        IWorkMediaEditor::from_package(editor.package().clone())
            .unwrap()
            .asset(first.asset.data_identifier)
            .is_none()
    );

    let before_invalid = editor.to_bytes().unwrap();
    assert!(editor.add_soundtrack_item("not-audio.png", FIRST).is_err());
    assert!(
        editor
            .insert_soundtrack_item(3, "bad.aiff", SECOND)
            .is_err()
    );
    assert!(editor.move_soundtrack_item(0, 2).is_err());
    assert_eq!(editor.to_bytes().unwrap(), before_invalid);

    editor.remove_soundtrack_item(1).unwrap();
    editor.remove_soundtrack_item(0).unwrap();
    assert!(editor.soundtrack_items().unwrap().is_empty());
    assert_eq!(editor.to_bytes().unwrap(), baseline);
}

#[test]
fn soundtrack_item_replacement_isolated_from_duplicate_references() {
    const ORIGINAL: &[u8] = b"FORM\0\0\0\x04AIFForiginal";
    const REPLACEMENT: &[u8] = b"FORM\0\0\0\x04AIFFreplacement";

    let mut editor = KeynoteEditor::from_package(test_package_with_empty_soundtrack()).unwrap();
    let original = editor
        .add_soundtrack_item("original.aiff", ORIGINAL)
        .unwrap();
    let mut package = editor.into_package();
    package
        .update_archive("Index/Document.iwa", |archive| {
            let object = archive.object(TEST_SOUNDTRACK_ID).unwrap();
            let mut payloads =
                soundtrack_wire::soundtrack_media_payloads(object.messages[0].data.as_slice())?;
            payloads.push(payloads[0].clone());
            let data = soundtrack_wire::rewrite_soundtrack_media(
                object.messages[0].data.as_slice(),
                &payloads,
            )?;
            soundtrack_wire::replace_soundtrack_message(archive, TEST_SOUNDTRACK_ID, data)
        })
        .unwrap();
    crate::data_reference_registry::add_component_data_reference(
        &mut package,
        1,
        original.asset.data_identifier.get(),
        TEST_SOUNDTRACK_ID,
    )
    .unwrap();

    let mut editor = KeynoteEditor::from_package(package).unwrap();
    let duplicated = editor.soundtrack_items().unwrap();
    assert_eq!(duplicated.len(), 2);
    assert_eq!(duplicated[0].asset.component_reference_count, 2);
    assert_eq!(duplicated[0].asset.message_reference_count, 2);

    let replacement = editor
        .replace_soundtrack_item(0, "replacement.aiff", REPLACEMENT)
        .unwrap();
    let items = editor.soundtrack_items().unwrap();
    assert_eq!(items[0], replacement);
    assert_eq!(
        items[1].asset.data_identifier,
        original.asset.data_identifier
    );
    assert_eq!(items[1].asset.component_reference_count, 1);
    assert_eq!(items[1].asset.message_reference_count, 1);

    editor.remove_soundtrack_item(1).unwrap();
    assert!(
        IWorkMediaEditor::from_package(editor.package().clone())
            .unwrap()
            .asset(original.asset.data_identifier)
            .is_none()
    );
}

#[test]
fn slide_number_visibility_matches_native_ownership_and_round_trips_exactly() {
    let mut editor = KeynoteEditor::from_package(test_package_with_slide_number()).unwrap();
    assert_eq!(
        editor.slides().unwrap()[0].is_slide_number_visible,
        Some(false)
    );
    let document_before = editor
        .package()
        .archive("Index/Document.iwa")
        .unwrap()
        .to_bytes()
        .unwrap();
    let slide_before = editor
        .package()
        .archive("Index/Slide-4.iwa")
        .unwrap()
        .to_bytes()
        .unwrap();
    let raw_placeholder_before = {
        let archive = editor.package().archive("Index/Slide-4.iwa").unwrap();
        let slide = &archive.object(4).unwrap().messages[0].data;
        repeated_length_delimited_payloads(slide, TEST_SLIDE_NUMBER_PLACEHOLDER_FIELD).unwrap()[0]
            .to_vec()
    };

    editor.set_slide_number_visible(0, true).unwrap();
    assert_eq!(
        editor.slides().unwrap()[0].is_slide_number_visible,
        Some(true)
    );
    let graph = ObjectGraph::read(editor.package()).unwrap();
    let slide: kn::SlideArchive = graph
        .decode_type(4, TEST_SLIDE_MESSAGE_TYPE, "KN.SlideArchive")
        .unwrap();
    assert_eq!(
        slide
            .owned_drawables
            .iter()
            .filter(|reference| reference.identifier == TEST_SLIDE_NUMBER_PLACEHOLDER_ID)
            .count(),
        1
    );
    assert_eq!(
        slide
            .drawables_z_order
            .iter()
            .filter(|reference| reference.identifier == TEST_SLIDE_NUMBER_PLACEHOLDER_ID)
            .count(),
        1
    );
    let archive = editor.package().archive("Index/Slide-4.iwa").unwrap();
    let data = &archive.object(4).unwrap().messages[0].data;
    assert_eq!(
        repeated_length_delimited_payloads(data, TEST_SLIDE_OWNED_DRAWABLES_FIELD)
            .unwrap()
            .last()
            .copied(),
        Some(raw_placeholder_before.as_slice())
    );
    assert_eq!(
        repeated_length_delimited_payloads(data, TEST_SLIDE_DRAWABLES_Z_ORDER_FIELD)
            .unwrap()
            .last()
            .copied(),
        Some(raw_placeholder_before.as_slice())
    );

    let visible = editor.to_bytes().unwrap();
    editor.set_slide_number_visible(0, true).unwrap();
    assert_eq!(editor.to_bytes().unwrap(), visible);
    editor.set_slide_number_visible(0, false).unwrap();
    assert_eq!(
        editor
            .package()
            .archive("Index/Document.iwa")
            .unwrap()
            .to_bytes()
            .unwrap(),
        document_before
    );
    assert_eq!(
        editor
            .package()
            .archive("Index/Slide-4.iwa")
            .unwrap()
            .to_bytes()
            .unwrap(),
        slide_before
    );

    let before_invalid = editor.to_bytes().unwrap();
    assert!(editor.set_slide_number_visible(1, true).is_err());
    assert!(editor.set_slide_number_visible(2, true).is_err());
    assert_eq!(editor.to_bytes().unwrap(), before_invalid);
}

#[test]
fn slide_number_visibility_rejects_inconsistent_native_state_transactionally() {
    let mut package = test_package_with_slide_number();
    package
        .update_archive("Index/Slide-4.iwa", |archive| {
            let slide = archive.object_mut(4).unwrap();
            let message = slide.messages[0].clone();
            let mut data = message.data;
            data = append_repeated_length_delimited_field(
                &data,
                TEST_SLIDE_OWNED_DRAWABLES_FIELD,
                &reference(TEST_SLIDE_NUMBER_PLACEHOLDER_ID).encode_to_vec(),
            )?;
            slide.replace_message(
                0,
                RawMessage {
                    type_: message.type_,
                    data,
                },
            )?;
            Ok(())
        })
        .unwrap();
    let mut editor = KeynoteEditor::from_package(package).unwrap();
    let before = editor.to_bytes().unwrap();
    assert!(editor.set_slide_number_visible(0, true).is_err());
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn slide_text_placeholder_visibility_matches_native_ownership_and_preserves_references() {
    let mut package = test_package();
    package
        .update_archive("Index/Slide-4.iwa", |archive| {
            let slide = archive.object_mut(4).unwrap();
            let message = slide.messages[0].clone();
            let data = transform_length_delimited_fields_at_path(
                &message.data,
                &[TEST_TITLE_PLACEHOLDER_FIELD],
                |reference| {
                    let mut reference = reference.to_vec();
                    append_unknown_varint(&mut reference, 98, 980);
                    Ok(reference)
                },
            )?;
            slide.replace_message(
                0,
                RawMessage {
                    type_: message.type_,
                    data,
                },
            )?;
            Ok(())
        })
        .unwrap();
    let mut editor = KeynoteEditor::from_package(package).unwrap();
    let before = editor.slides().unwrap();
    assert_eq!(before[0].is_title_visible, Some(true));
    assert_eq!(before[0].is_body_visible, Some(true));
    let document_before = editor
        .package()
        .archive("Index/Document.iwa")
        .unwrap()
        .to_bytes()
        .unwrap();
    let raw_title = {
        let archive = editor.package().archive("Index/Slide-4.iwa").unwrap();
        repeated_length_delimited_payloads(
            &archive.object(4).unwrap().messages[0].data,
            TEST_TITLE_PLACEHOLDER_FIELD,
        )
        .unwrap()[0]
            .to_vec()
    };

    editor.set_slide_title_visible(0, false).unwrap();
    let hidden = editor.slides().unwrap();
    assert_eq!(hidden[0].is_title_visible, Some(false));
    assert_eq!(hidden[0].is_body_visible, Some(true));
    assert_eq!(hidden[0].title.as_deref(), Some("Old title"));
    let hidden_bytes = editor.to_bytes().unwrap();
    editor.set_slide_title_visible(0, false).unwrap();
    assert_eq!(editor.to_bytes().unwrap(), hidden_bytes);

    editor.set_slide_title_visible(0, true).unwrap();
    let graph = ObjectGraph::read(editor.package()).unwrap();
    let slide: kn::SlideArchive = graph
        .decode_type(4, TEST_SLIDE_MESSAGE_TYPE, "KN.SlideArchive")
        .unwrap();
    assert_eq!(
        slide
            .owned_drawables
            .iter()
            .map(|reference| reference.identifier)
            .collect::<Vec<_>>(),
        vec![6, 5]
    );
    assert_eq!(
        slide
            .drawables_z_order
            .iter()
            .map(|reference| reference.identifier)
            .collect::<Vec<_>>(),
        vec![6, 5]
    );
    let archive = editor.package().archive("Index/Slide-4.iwa").unwrap();
    let data = &archive.object(4).unwrap().messages[0].data;
    for field in [
        TEST_SLIDE_OWNED_DRAWABLES_FIELD,
        TEST_SLIDE_DRAWABLES_Z_ORDER_FIELD,
    ] {
        assert_eq!(
            repeated_length_delimited_payloads(data, field)
                .unwrap()
                .last()
                .copied(),
            Some(raw_title.as_slice())
        );
    }
    assert_eq!(
        editor
            .package()
            .archive("Index/Document.iwa")
            .unwrap()
            .to_bytes()
            .unwrap(),
        document_before
    );

    editor
        .set_slide_text_placeholder_visible(0, KeynoteSlideTextPlaceholder::Body, false)
        .unwrap();
    let updated = editor.slides().unwrap();
    assert_eq!(updated[0].is_title_visible, Some(true));
    assert_eq!(updated[0].is_body_visible, Some(false));
    assert_eq!(updated[0].body.as_deref(), Some("Old body 🚀"));
    let reparsed = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(reparsed.slides().unwrap(), updated);
}

#[test]
fn slide_text_placeholder_visibility_rejects_missing_or_inconsistent_state() {
    let mut editor = KeynoteEditor::from_package(test_package()).unwrap();
    let before = editor.to_bytes().unwrap();
    assert!(editor.set_slide_title_visible(2, false).is_err());
    assert_eq!(editor.to_bytes().unwrap(), before);

    let mut missing = test_package();
    missing
        .update_archive("Index/Slide-4.iwa", |archive| {
            let slide = archive.object_mut(4).unwrap();
            let message = slide.messages[0].clone();
            let mut decoded = kn::SlideArchive::decode(message.data.as_slice())?;
            decoded.body_placeholder = None;
            decoded
                .owned_drawables
                .retain(|reference| reference.identifier != 6);
            decoded
                .drawables_z_order
                .retain(|reference| reference.identifier != 6);
            slide.replace_message(
                0,
                RawMessage {
                    type_: message.type_,
                    data: decoded.encode_to_vec(),
                },
            )?;
            Ok(())
        })
        .unwrap();
    let mut missing_editor = KeynoteEditor::from_package(missing).unwrap();
    assert_eq!(missing_editor.slides().unwrap()[0].is_body_visible, None);
    let missing_before = missing_editor.to_bytes().unwrap();
    assert!(missing_editor.set_slide_body_visible(0, true).is_err());
    assert_eq!(missing_editor.to_bytes().unwrap(), missing_before);

    let mut inconsistent = test_package();
    inconsistent
        .update_archive("Index/Slide-4.iwa", |archive| {
            let slide = archive.object_mut(4).unwrap();
            let message = slide.messages[0].clone();
            let mut decoded = kn::SlideArchive::decode(message.data.as_slice())?;
            decoded
                .drawables_z_order
                .retain(|reference| reference.identifier != 5);
            slide.replace_message(
                0,
                RawMessage {
                    type_: message.type_,
                    data: decoded.encode_to_vec(),
                },
            )?;
            Ok(())
        })
        .unwrap();
    assert!(KeynoteEditor::from_package(inconsistent).is_err());
}

#[test]
fn transition_lifecycle_is_typed_transactional_and_wire_exact() {
    let mut editor = KeynoteEditor::from_package(test_package()).unwrap();
    let original: Settings = editor.slide_transition(0).unwrap().unwrap();
    assert_eq!(original.effect(), Some(&Effect::None));
    assert!(!original.has_effect());

    let mut dissolve = original.clone();
    dissolve.set_effect(Some(Effect::Dissolve)).unwrap();
    dissolve.set_duration(Some(1.5)).unwrap();
    dissolve.set_direction(Some(Direction::from_native(2)));
    let mut animation_parameters = dissolve.animation_parameters().clone();
    animation_parameters.set_detail(Some(0.75)).unwrap();
    dissolve
        .set_animation_parameters(animation_parameters)
        .unwrap();
    let mut custom_parameters = *dissolve.custom_parameters();
    custom_parameters.set_bounce(Some(true));
    dissolve.set_custom_parameters(custom_parameters).unwrap();
    editor.set_slide_transition(0, dissolve.clone()).unwrap();
    assert_eq!(editor.slide_transition(0).unwrap(), Some(dissolve));
    assert!(editor.slide_transition(0).unwrap().unwrap().has_effect());

    assert!(editor.clear_slide_transition(0).unwrap());
    assert_eq!(editor.slide_transition(0).unwrap(), Some(original));
    let cleared = editor.to_bytes().unwrap();
    assert!(!editor.clear_slide_transition(0).unwrap());
    assert!(editor.slide_transition(99).is_err());
    assert!(editor.clear_slide_transition(99).is_err());
    assert_eq!(editor.to_bytes().unwrap(), cleared);

    assert!(Effect::unknown("none").is_err());
    assert_eq!(editor.to_bytes().unwrap(), cleared);
}

#[test]
fn transition_lifecycle_synchronizes_the_separate_slide_node_component() {
    let mut package = test_package();
    package
        .update_archive("Index/Document.iwa", |archive| {
            let object = archive.object_mut(3).unwrap();
            let message = object.messages[0].clone();
            let mut data = patch_varint_field(&message.data, 7, true, Some(0))?;
            append_unknown_varint(&mut data, 97, 9_700);
            object.replace_message(
                0,
                RawMessage {
                    type_: message.type_,
                    data,
                },
            )?;
            object.archive_info.message_infos[0]
                .field_infos
                .push(FieldInfo {
                    path: FieldPath { path: vec![7] },
                    ..Default::default()
                });
            Ok(())
        })
        .unwrap();

    let mut editor = KeynoteEditor::from_package(package).unwrap();
    let baseline = editor.to_bytes().unwrap();
    let original = editor.slide_transition(0).unwrap().unwrap();
    assert!(!original.has_effect());
    let graph = ObjectGraph::read(editor.package()).unwrap();
    assert_ne!(
        graph.archive_name(3).unwrap(),
        graph.archive_name(4).unwrap()
    );
    let original_node_wire = graph
        .message_data_type(3, TEST_SLIDE_NODE_MESSAGE_TYPE, "KN.SlideNodeArchive")
        .unwrap();
    let original_unknown = crate::wire::parse_wire_fields(original_node_wire)
        .unwrap()
        .into_iter()
        .find(|field| field.number() == 97)
        .map(|field| original_node_wire[field.start()..field.end()].to_vec())
        .unwrap();
    let original_header_metadata = editor
        .package()
        .archive("Index/Document.iwa")
        .unwrap()
        .object(3)
        .unwrap()
        .archive_info
        .message_infos[0]
        .field_infos
        .clone();

    let mut dissolve = original.clone();
    dissolve.set_effect(Some(Effect::Dissolve)).unwrap();
    editor.set_slide_transition(0, dissolve.clone()).unwrap();
    assert_eq!(editor.slide_transition(0).unwrap(), Some(dissolve.clone()));
    let graph = ObjectGraph::read(editor.package()).unwrap();
    let node: kn::SlideNodeArchive = graph
        .decode_type(3, TEST_SLIDE_NODE_MESSAGE_TYPE, "KN.SlideNodeArchive")
        .unwrap();
    assert!(node.has_transition);
    let node_wire = graph
        .message_data_type(3, TEST_SLIDE_NODE_MESSAGE_TYPE, "KN.SlideNodeArchive")
        .unwrap();
    let unknown = crate::wire::parse_wire_fields(node_wire)
        .unwrap()
        .into_iter()
        .find(|field| field.number() == 97)
        .map(|field| node_wire[field.start()..field.end()].to_vec())
        .unwrap();
    assert_eq!(unknown, original_unknown);
    assert_eq!(
        editor
            .package()
            .archive("Index/Document.iwa")
            .unwrap()
            .object(3)
            .unwrap()
            .archive_info
            .message_infos[0]
            .field_infos,
        original_header_metadata
    );

    let changed = editor.to_bytes().unwrap();
    editor.set_slide_transition(0, dissolve).unwrap();
    assert_eq!(editor.to_bytes().unwrap(), changed);

    let reopened = KeynoteEditor::from_bytes(&changed).unwrap();
    assert_eq!(
        reopened.slide_transition(0).unwrap().unwrap().has_effect(),
        true
    );
    let reopened_graph = ObjectGraph::read(reopened.package()).unwrap();
    let reopened_node: kn::SlideNodeArchive = reopened_graph
        .decode_type(3, TEST_SLIDE_NODE_MESSAGE_TYPE, "KN.SlideNodeArchive")
        .unwrap();
    assert!(reopened_node.has_transition);

    assert!(editor.clear_slide_transition(0).unwrap());
    assert!(!editor.slide_transition(0).unwrap().unwrap().has_effect());
    let graph = ObjectGraph::read(editor.package()).unwrap();
    let node: kn::SlideNodeArchive = graph
        .decode_type(3, TEST_SLIDE_NODE_MESSAGE_TYPE, "KN.SlideNodeArchive")
        .unwrap();
    assert!(!node.has_transition);
    assert_eq!(editor.to_bytes().unwrap(), baseline);
    assert!(!editor.clear_slide_transition(0).unwrap());
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    let mut malformed = KeynoteEditor::from_package(test_package()).unwrap();
    let mut invalid = malformed.slide_transition(0).unwrap().unwrap();
    invalid.set_effect(Some(Effect::Dissolve)).unwrap();
    let mut corrupted = malformed.into_package();
    corrupted
        .update_archive("Index/Document.iwa", |archive| {
            let object = archive.object_mut(3).unwrap();
            let message = object.messages[0].clone();
            let mut data = message.data;
            append_unknown_varint(&mut data, 7, 0);
            object.replace_message(
                0,
                RawMessage {
                    type_: TEST_SLIDE_NODE_MESSAGE_TYPE,
                    data,
                },
            )?;
            Ok(())
        })
        .unwrap();
    malformed = KeynoteEditor::from_package(corrupted).unwrap();
    let malformed_before = malformed.to_bytes().unwrap();
    assert!(malformed.set_slide_transition(0, invalid).is_err());
    assert_eq!(malformed.to_bytes().unwrap(), malformed_before);
    assert!(!malformed.slide_transition(0).unwrap().unwrap().has_effect());
}

#[test]
fn transition_custom_parameter_crud_is_lossless_and_transactional() {
    let mut editor = KeynoteEditor::from_package(test_package()).unwrap();
    let mut settings = editor.slides().unwrap()[0].transition.clone().unwrap();
    settings
        .set_effect(Some(
            Effect::unknown("com.example.future-transition").unwrap(),
        ))
        .unwrap();
    settings.set_direction(Some(Direction::from_native(42)));
    let mut custom_parameters = CustomParameters::new();
    custom_parameters.set_twist(Some(-0.375)).unwrap();
    custom_parameters.set_mosaic_size(Some(0));
    custom_parameters.set_mosaic_type(Some(MosaicType::from_native(7)));
    custom_parameters.set_bounce(Some(false));
    custom_parameters.set_magic_move_fade_unmatched_objects(Some(true));
    custom_parameters.set_acceleration(Some(Acceleration::Custom));
    custom_parameters.set_text_delivery(Some(TextDelivery::ByCharacter));
    custom_parameters.set_motion_blur(Some(false));
    custom_parameters.set_travel_distance(Some(275.5)).unwrap();
    settings.set_custom_parameters(custom_parameters).unwrap();
    editor.set_slide_transition(0, settings.clone()).unwrap();
    assert_eq!(
        editor.slides().unwrap()[0].transition.as_ref(),
        Some(&settings)
    );

    let graph = ObjectGraph::read(editor.package()).unwrap();
    let native: kn::SlideArchive = graph.decode_type(4, 5, "KN.SlideArchive").unwrap();
    let native = native.transition.attributes;
    assert_eq!(native.custom_twist, Some(-0.375));
    assert_eq!(native.custom_mosaic_size, Some(0));
    assert_eq!(native.custom_mosaic_type, Some(7));
    assert_eq!(native.custom_bounce, Some(false));
    assert_eq!(native.custom_magic_move_fade_unmatched_objects, Some(true));
    assert_eq!(native.custom_timing_curve, Some(5));
    assert_eq!(native.custom_text_delivery_type, Some(3));
    assert_eq!(native.custom_motion_blur, Some(false));
    assert_eq!(native.custom_travel_distance, Some(275.5));
    drop(graph);

    let before_noop = editor.to_bytes().unwrap();
    editor.set_slide_transition(0, settings.clone()).unwrap();
    assert_eq!(editor.to_bytes().unwrap(), before_noop);

    let mut custom_parameters = *settings.custom_parameters();
    custom_parameters.set_acceleration(Some(Acceleration::from_native(19)));
    custom_parameters.set_text_delivery(Some(TextDelivery::from_native(-1)));
    settings.set_custom_parameters(custom_parameters).unwrap();
    editor.set_slide_transition(0, settings.clone()).unwrap();
    let reparsed = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reparsed.slides().unwrap()[0].transition.as_ref(),
        Some(&settings)
    );

    let before_invalid = editor.to_bytes().unwrap();
    let mut invalid_custom_parameters = *settings.custom_parameters();
    assert!(invalid_custom_parameters.set_twist(Some(f32::NAN)).is_err());
    assert_eq!(invalid_custom_parameters, *settings.custom_parameters());
    assert!(
        invalid_custom_parameters
            .set_travel_distance(Some(f32::INFINITY))
            .is_err()
    );
    assert_eq!(invalid_custom_parameters, *settings.custom_parameters());
    assert!(Effect::unknown("apple:dissolve").is_err());
    let mut invalid = settings.clone();
    assert!(
        invalid
            .set_animation_type(Some("invalid\0transition"))
            .is_err()
    );
    assert_eq!(invalid, settings);
    assert_eq!(editor.to_bytes().unwrap(), before_invalid);

    settings
        .set_custom_parameters(CustomParameters::default())
        .unwrap();
    editor.set_slide_transition(0, settings.clone()).unwrap();
    let graph = ObjectGraph::read(editor.package()).unwrap();
    let native: kn::SlideArchive = graph.decode_type(4, 5, "KN.SlideArchive").unwrap();
    let native = native.transition.attributes;
    assert_eq!(native.custom_twist, None);
    assert_eq!(native.custom_mosaic_size, None);
    assert_eq!(native.custom_mosaic_type, None);
    assert_eq!(native.custom_bounce, None);
    assert_eq!(native.custom_magic_move_fade_unmatched_objects, None);
    assert_eq!(native.custom_timing_curve, None);
    assert_eq!(native.custom_text_delivery_type, None);
    assert_eq!(native.custom_motion_blur, None);
    assert_eq!(native.custom_travel_distance, None);
}

#[test]
fn transition_custom_parameters_reject_malformed_wire_transactionally() {
    for mutation in 0..10 {
        let mut editor = KeynoteEditor::from_package(test_package()).unwrap();
        let mut settings = editor.slides().unwrap()[0].transition.clone().unwrap();
        let mut custom_parameters = CustomParameters::new();
        custom_parameters.set_twist(Some(0.25)).unwrap();
        custom_parameters.set_mosaic_size(Some(16));
        custom_parameters.set_mosaic_type(Some(MosaicType::from_native(2)));
        custom_parameters.set_bounce(Some(true));
        custom_parameters.set_magic_move_fade_unmatched_objects(Some(false));
        custom_parameters.set_acceleration(Some(Acceleration::EaseInOut));
        custom_parameters.set_text_delivery(Some(TextDelivery::ByWord));
        custom_parameters.set_motion_blur(Some(true));
        custom_parameters.set_travel_distance(Some(80.0)).unwrap();
        settings.set_custom_parameters(custom_parameters).unwrap();
        editor.set_slide_transition(0, settings.clone()).unwrap();
        let mut package = editor.into_package();
        package
            .update_archive("Index/Slide-4.iwa", |archive| {
                let slide = archive.object_mut(4).unwrap();
                let message = slide.messages[0].clone();
                let data = transform_length_delimited_field(&message.data, 4, |transition| {
                    transform_length_delimited_field(transition, 2, |attributes| {
                        let mut attributes = attributes.to_vec();
                        match mutation {
                            0 => append_unknown_fixed32(&mut attributes, 9, 0.5_f32.to_bits()),
                            1 => append_unknown_varint(&mut attributes, 10, 24),
                            2 => append_unknown_varint(&mut attributes, 11, 3),
                            3 => append_unknown_varint(&mut attributes, 12, 0),
                            4 => append_unknown_varint(&mut attributes, 13, 1),
                            5 => append_unknown_varint(&mut attributes, 15, 3),
                            6 => append_unknown_varint(&mut attributes, 16, 4),
                            7 => append_unknown_varint(&mut attributes, 17, 0),
                            8 => append_unknown_fixed32(&mut attributes, 18, 120.0_f32.to_bits()),
                            9 => append_unknown_varint(&mut attributes, 9, 1),
                            _ => unreachable!(),
                        }
                        Ok(attributes)
                    })
                })?;
                slide.replace_message(
                    0,
                    RawMessage {
                        type_: message.type_,
                        data,
                    },
                )?;
                Ok(())
            })
            .unwrap();
        if let Ok(mut malformed) = KeynoteEditor::from_package(package) {
            assert!(malformed.slides().is_err(), "mutation {mutation}");
            let before = malformed.to_bytes().unwrap();
            assert!(
                malformed.set_slide_transition(0, settings).is_err(),
                "mutation {mutation}"
            );
            assert_eq!(malformed.to_bytes().unwrap(), before);
        }
    }
}

#[test]
fn transition_animation_parameter_crud_preserves_raw_payloads_transactionally() {
    let mut color_payload = tsp::Color {
        model: tsp::color::ColorModel::Rgb as i32,
        r: Some(0.125),
        g: Some(0.5),
        b: Some(0.875),
        a: Some(0.75),
        rgbspace: Some(tsp::color::RgbColorSpace::P3 as i32),
        ..Default::default()
    }
    .encode_to_vec();
    append_unknown_varint(&mut color_payload, 99, 9_900);
    let mut curve_payload =
        native_motion_path(&KeynoteMotionPath::straight(1.0, 1.0)).encode_to_vec();
    append_unknown_varint(&mut curve_payload, 98, 9_800);

    let mut editor = KeynoteEditor::from_package(test_package()).unwrap();
    let mut settings = editor.slides().unwrap()[0].transition.clone().unwrap();
    let mut animation_parameters = AnimationParameters::new();
    animation_parameters
        .set_color_payload(Some(&color_payload))
        .unwrap();
    for slot in TimingCurveSlot::ALL {
        animation_parameters
            .set_timing_curve_payload(slot, Some(&curve_payload))
            .unwrap();
    }
    animation_parameters.set_random_number_seed(Some(0));
    animation_parameters.set_detail(Some(0.0)).unwrap();
    for (slot, name) in TimingCurveSlot::ALL
        .into_iter()
        .zip(["", "Ease In", "Custom Bézier"])
    {
        animation_parameters
            .set_timing_curve_theme_name(slot, Some(name))
            .unwrap();
    }
    animation_parameters.set_writing_direction_is_rtl(Some(false));
    settings
        .set_animation_parameters(animation_parameters)
        .unwrap();
    editor.set_slide_transition(0, settings.clone()).unwrap();
    assert_eq!(
        editor.slides().unwrap()[0].transition.as_ref(),
        Some(&settings)
    );

    let graph = ObjectGraph::read(editor.package()).unwrap();
    let native: kn::SlideArchive = graph.decode_type(4, 5, "KN.SlideArchive").unwrap();
    let animation = native.transition.attributes.animation_attributes.unwrap();
    assert!(animation.color.is_some());
    assert!(animation.custom_effect_timing_curve_1.is_some());
    assert!(animation.custom_effect_timing_curve_2.is_some());
    assert!(animation.custom_effect_timing_curve_3.is_some());
    assert_eq!(animation.random_number_seed, Some(0));
    assert_eq!(animation.custom_detail, Some(0.0));
    assert_eq!(
        animation.custom_effect_timing_curve_theme_name_1.as_deref(),
        Some("")
    );
    assert_eq!(animation.writing_direction_is_rtl, Some(false));
    drop(graph);

    let before_noop = editor.to_bytes().unwrap();
    editor.set_slide_transition(0, settings.clone()).unwrap();
    assert_eq!(editor.to_bytes().unwrap(), before_noop);

    let before_invalid = editor.to_bytes().unwrap();
    let mut invalid_animation_parameters = settings.animation_parameters().clone();
    assert!(
        invalid_animation_parameters
            .set_detail(Some(f64::NAN))
            .is_err()
    );
    assert_eq!(
        invalid_animation_parameters,
        *settings.animation_parameters()
    );
    let mut invalid = settings.clone();
    let mut invalid_animation_parameters = settings.animation_parameters().clone();
    invalid_animation_parameters
        .set_color_payload(Some(&[0xff]))
        .unwrap();
    invalid
        .set_animation_parameters(invalid_animation_parameters)
        .unwrap();
    assert!(editor.set_slide_transition(0, invalid).is_err());
    let mut invalid = settings.clone();
    let mut invalid_animation_parameters = settings.animation_parameters().clone();
    invalid_animation_parameters
        .set_timing_curve_payload(TimingCurveSlot::Second, Some(&[0xff]))
        .unwrap();
    invalid
        .set_animation_parameters(invalid_animation_parameters)
        .unwrap();
    assert!(editor.set_slide_transition(0, invalid).is_err());
    let mut invalid_animation_parameters = settings.animation_parameters().clone();
    assert!(
        invalid_animation_parameters
            .set_timing_curve_theme_name(TimingCurveSlot::Third, Some("bad\0name"))
            .is_err()
    );
    assert_eq!(
        invalid_animation_parameters,
        *settings.animation_parameters()
    );
    assert_eq!(editor.to_bytes().unwrap(), before_invalid);

    settings
        .set_animation_parameters(AnimationParameters::default())
        .unwrap();
    editor.set_slide_transition(0, settings).unwrap();
    let graph = ObjectGraph::read(editor.package()).unwrap();
    let native: kn::SlideArchive = graph.decode_type(4, 5, "KN.SlideArchive").unwrap();
    let animation = native.transition.attributes.animation_attributes.unwrap();
    assert_eq!(animation.color, None);
    assert_eq!(animation.custom_effect_timing_curve_1, None);
    assert_eq!(animation.custom_effect_timing_curve_2, None);
    assert_eq!(animation.custom_effect_timing_curve_3, None);
    assert_eq!(animation.random_number_seed, None);
    assert_eq!(animation.custom_detail, None);
    assert_eq!(animation.custom_effect_timing_curve_theme_name_1, None);
    assert_eq!(animation.custom_effect_timing_curve_theme_name_2, None);
    assert_eq!(animation.custom_effect_timing_curve_theme_name_3, None);
    assert_eq!(animation.writing_direction_is_rtl, None);
}

#[test]
fn transition_animation_parameters_reject_malformed_wire() {
    for mutation in 0..10 {
        let mut editor = KeynoteEditor::from_package(test_package()).unwrap();
        let mut settings = editor.slides().unwrap()[0].transition.clone().unwrap();
        let color_payload = tsp::Color {
            model: tsp::color::ColorModel::White as i32,
            w: Some(0.5),
            ..Default::default()
        }
        .encode_to_vec();
        let curve_payload =
            native_motion_path(&KeynoteMotionPath::straight(1.0, 1.0)).encode_to_vec();
        let mut animation_parameters = AnimationParameters::new();
        animation_parameters
            .set_color_payload(Some(&color_payload))
            .unwrap();
        animation_parameters
            .set_timing_curve_payload(TimingCurveSlot::First, Some(&curve_payload))
            .unwrap();
        animation_parameters.set_random_number_seed(Some(17));
        animation_parameters.set_detail(Some(0.25)).unwrap();
        animation_parameters
            .set_timing_curve_theme_name(TimingCurveSlot::First, Some("Curve"))
            .unwrap();
        animation_parameters.set_writing_direction_is_rtl(Some(true));
        settings
            .set_animation_parameters(animation_parameters)
            .unwrap();
        editor.set_slide_transition(0, settings.clone()).unwrap();
        let mut package = editor.into_package();
        package
            .update_archive("Index/Slide-4.iwa", |archive| {
                let slide = archive.object_mut(4).unwrap();
                let message = slide.messages[0].clone();
                let data = transform_length_delimited_field(&message.data, 4, |transition| {
                    transform_length_delimited_field(transition, 2, |attributes| {
                        transform_length_delimited_field(attributes, 8, |animation| {
                            let mut animation = animation.to_vec();
                            match mutation {
                                0 => {
                                    animation = append_repeated_length_delimited_field(
                                        &animation,
                                        7,
                                        &color_payload,
                                    )?;
                                },
                                1 => {
                                    animation = append_repeated_length_delimited_field(
                                        &animation,
                                        8,
                                        &curve_payload,
                                    )?;
                                },
                                2 => append_unknown_varint(&mut animation, 11, 18),
                                3 => append_unknown_fixed64(&mut animation, 12, 0.5_f64.to_bits()),
                                4 => {
                                    animation = append_repeated_length_delimited_field(
                                        &animation, 13, b"Other",
                                    )?;
                                },
                                5 => append_unknown_varint(&mut animation, 16, 0),
                                6 => append_unknown_varint(&mut animation, 7, 1),
                                7 => append_unknown_varint(&mut animation, 12, 1),
                                8 => append_unknown_fixed64(&mut animation, 11, 17_f64.to_bits()),
                                9 => append_unknown_varint(&mut animation, 13, 1),
                                _ => unreachable!(),
                            }
                            Ok(animation)
                        })
                    })
                })?;
                slide.replace_message(
                    0,
                    RawMessage {
                        type_: message.type_,
                        data,
                    },
                )?;
                Ok(())
            })
            .unwrap();
        if let Ok(mut malformed) = KeynoteEditor::from_package(package) {
            assert!(malformed.slides().is_err(), "mutation {mutation}");
            let before = malformed.to_bytes().unwrap();
            assert!(
                malformed.set_slide_transition(0, settings).is_err(),
                "mutation {mutation}"
            );
            assert_eq!(malformed.to_bytes().unwrap(), before);
        }
    }
}

#[test]
fn scalar_updates_preserve_unknown_wire_and_restore_exact_components() {
    let mut package = test_package();
    package
        .update_archive("Index/Document.iwa", |archive| {
            for (identifier, field_number) in [(2, 99), (3, 98)] {
                let object = archive.object_mut(identifier).unwrap();
                let mut message = object.messages[0].clone();
                append_unknown_varint(&mut message.data, field_number, 900 + identifier);
                object.replace_message(0, message)?;
            }
            Ok(())
        })
        .unwrap();
    package
        .update_archive("Index/Slide-4.iwa", |archive| {
            let object = archive.object_mut(4).unwrap();
            let message = object.messages[0].clone();
            let mut data =
                crate::wire::transform_length_delimited_field(&message.data, 4, |transition| {
                    let mut transition = crate::wire::transform_length_delimited_field(
                        transition,
                        2,
                        |attributes| {
                            let mut attributes = crate::wire::transform_length_delimited_field(
                                attributes,
                                8,
                                |animation| {
                                    let mut animation = animation.to_vec();
                                    append_unknown_varint(&mut animation, 96, 960);
                                    Ok(animation)
                                },
                            )?;
                            append_unknown_varint(&mut attributes, 97, 970);
                            Ok(attributes)
                        },
                    )?;
                    append_unknown_varint(&mut transition, 98, 980);
                    Ok(transition)
                })?;
            append_unknown_varint(&mut data, 99, 990);
            Ok(object
                .replace_message(
                    0,
                    RawMessage {
                        type_: message.type_,
                        data,
                    },
                )
                .map(|_| ())?)
        })
        .unwrap();
    let mut editor = KeynoteEditor::from_package(package).unwrap();
    let document_before = editor
        .package()
        .archive("Index/Document.iwa")
        .unwrap()
        .to_bytes()
        .unwrap();
    let slide_before = editor
        .package()
        .archive("Index/Slide-4.iwa")
        .unwrap()
        .to_bytes()
        .unwrap();
    let original_show = editor.show_settings().unwrap();
    let original_transition = editor.slides().unwrap()[0].transition.clone().unwrap();

    let mut changed_show = original_show;
    changed_show.set_size(Size::new(1_920.0, 1_080.0).unwrap());
    changed_show.set_slide_numbers_visible(Some(true));
    changed_show.set_loop_presentation(Some(true));
    changed_show.set_mode(Some(Mode::SelfPlaying)).unwrap();
    changed_show.set_autoplay_transition_delay(Some(Seconds::new(3.5).unwrap()));
    changed_show.set_autoplay_build_delay(Some(Seconds::new(1.25).unwrap()));
    changed_show.set_idle_timer_active(Some(true));
    changed_show.set_idle_timer_delay(Some(Seconds::new(60.0).unwrap()));
    changed_show.set_automatically_plays_upon_open(Some(true));
    editor.set_show_settings(changed_show).unwrap();
    editor.set_show_settings(original_show).unwrap();

    let mut changed_transition = original_transition.clone();
    changed_transition
        .set_animation_type(Some("Transition"))
        .unwrap();
    changed_transition
        .set_effect(Some(Effect::unknown("dissolve").unwrap()))
        .unwrap();
    changed_transition.set_duration(Some(2.5)).unwrap();
    changed_transition.set_direction(Some(Direction::from_native(2)));
    changed_transition.set_delay(Some(1.0)).unwrap();
    changed_transition.set_is_automatic(Some(true));
    editor.set_slide_transition(0, changed_transition).unwrap();
    editor.set_slide_transition(0, original_transition).unwrap();
    editor.set_slide_name(0, Some("Temporary")).unwrap();
    editor.set_slide_name(0, None).unwrap();

    assert_eq!(
        editor
            .package()
            .archive("Index/Document.iwa")
            .unwrap()
            .to_bytes()
            .unwrap(),
        document_before
    );
    assert_eq!(
        editor
            .package()
            .archive("Index/Slide-4.iwa")
            .unwrap()
            .to_bytes()
            .unwrap(),
        slide_before
    );
}

#[test]
fn show_update_rejects_duplicate_scalar_fields_transactionally() {
    let mut package = test_package();
    package
        .update_archive("Index/Document.iwa", |archive| {
            let object = archive.object_mut(2).unwrap();
            let mut message = object.messages[0].clone();
            append_unknown_varint(&mut message.data, 8, 0);
            append_unknown_varint(&mut message.data, 8, 1);
            Ok(object.replace_message(0, message).map(|_| ())?)
        })
        .unwrap();
    let mut editor = KeynoteEditor::from_package(package).unwrap();
    let before = editor.to_bytes().unwrap();
    let mut settings = editor.show_settings().unwrap();
    settings.set_loop_presentation(Some(false));
    assert!(editor.set_show_settings(settings).is_err());
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn show_update_rejects_duplicate_mode_fields_transactionally() {
    let mut package = test_package();
    package
        .update_archive("Index/Document.iwa", |archive| {
            let object = archive.object_mut(2).unwrap();
            let mut message = object.messages[0].clone();
            append_unknown_varint(&mut message.data, TEST_SHOW_MODE_FIELD, 0);
            append_unknown_varint(&mut message.data, TEST_SHOW_MODE_FIELD, 1);
            Ok(object.replace_message(0, message).map(|_| ())?)
        })
        .unwrap();
    let mut editor = KeynoteEditor::from_package(package).unwrap();
    let before = editor.to_bytes().unwrap();
    let mut settings = editor.show_settings().unwrap();
    settings.set_mode(Some(Mode::LinksOnly)).unwrap();
    assert!(editor.set_show_settings(settings).is_err());
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn removes_slide_tree_node_and_slide_transactionally() {
    let mut editor = KeynoteEditor::from_package(test_package()).unwrap();
    let removed = editor.remove_slide(0).unwrap();
    assert_eq!(removed.slide_id, 4);
    assert!(!editor.package().contains_entry("Index/Slide-4.iwa"));
    assert!(editor.package().contains_entry("Index/Slide-10.iwa"));
    let slides = editor.slides().unwrap();
    assert_eq!(slides.len(), 1);
    assert_eq!(slides[0].slide_id, 10);
    assert_eq!(slides[0].title.as_deref(), Some("Second title"));
    let before = editor.to_bytes().unwrap();
    assert!(editor.remove_slide(0).is_err());
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn reads_current_slide_layout_from_theme_relationship() {
    let editor = KeynoteEditor::from_package(test_package_with_current_layout()).unwrap();
    let slides = editor.slides().unwrap();
    assert_eq!(
        slides[0].layout,
        Some(KeynoteSlideLayoutInfo {
            id: KeynoteSlideLayoutId::new(30).unwrap(),
            name: "Title & Bullets".to_owned(),
            is_default: true,
        })
    );
    assert_eq!(slides[1].layout, None);
}

#[test]
fn slide_layout_ids_reject_null_and_round_trip_nonzero_values() {
    assert_eq!(KeynoteSlideLayoutId::new(0), None);
    let id = KeynoteSlideLayoutId::new(37).expect("nonzero layout IDs are valid");
    assert_eq!(id.as_u64(), 37);
}

#[test]
fn updates_slide_layout_transactionally_without_replacing_user_content() {
    let mut editor = KeynoteEditor::from_package(test_package_with_two_layouts()).unwrap();
    let layouts = editor.slide_layouts().unwrap();
    let title_only = layouts
        .iter()
        .find(|layout| layout.name == "Title Only")
        .unwrap()
        .id;
    let title_and_bullets = layouts
        .iter()
        .find(|layout| layout.name == "Title & Bullets")
        .unwrap()
        .id;

    editor.set_slide_layout(0, title_only).unwrap();
    let slide = &editor.slides().unwrap()[0];
    assert_eq!(
        slide.layout.as_ref().map(|layout| layout.id),
        Some(title_only)
    );
    assert_eq!(slide.is_title_visible, Some(true));
    assert_eq!(slide.is_body_visible, Some(false));
    assert_eq!(slide.title.as_deref(), Some("Old title"));
    assert_eq!(slide.body.as_deref(), Some("Old body 🚀"));
    assert_eq!(slide.notes.as_deref(), Some("Speaker 🚀"));

    let graph = ObjectGraph::read(editor.package()).unwrap();
    let native_slide: kn::SlideArchive = graph.decode_type(4, 5, "KN.SlideArchive").unwrap();
    assert_eq!(native_slide.style.identifier, 60);
    assert_eq!(native_slide.template_slide.unwrap().identifier, 38);
    assert_eq!(native_slide.owned_drawables, vec![reference(5)]);
    assert_eq!(native_slide.drawables_z_order, vec![reference(5)]);
    let node: kn::SlideNodeArchive = graph.decode_type(3, 4, "KN.SlideNodeArchive").unwrap();
    assert_eq!(node.template_slide_id, Some(test_uuid(3, 4)));
    assert_eq!(node.thumbnails_are_dirty, Some(true));
    assert!(node.thumbnails.is_empty());
    let title: kn::PlaceholderArchive = graph
        .decode_type(5, TEST_PLACEHOLDER_MESSAGE_TYPE, "KN.PlaceholderArchive")
        .unwrap();
    let body: kn::PlaceholderArchive = graph
        .decode_type(6, TEST_PLACEHOLDER_MESSAGE_TYPE, "KN.PlaceholderArchive")
        .unwrap();
    assert_eq!(title.super_.owned_storage, Some(reference(7)));
    assert_eq!(body.super_.owned_storage, Some(reference(8)));
    assert_eq!(placeholder_x(&title), Some(300.0));
    assert_eq!(placeholder_x(&body), Some(400.0));
    assert_eq!(title.super_.super_.style, Some(reference(61)));
    assert_eq!(body.super_.super_.style, Some(reference(62)));
    assert_eq!(
        title
            .super_
            .super_
            .pathsource
            .as_ref()
            .and_then(|path| path.horizontal_flip),
        Some(true)
    );
    let package = editor.package();
    let slide_archive = package.archive("Index/Slide-4.iwa").unwrap();
    let slide_object = slide_archive.object(4).unwrap();
    assert_eq!(
        slide_object.archive_info.message_infos[0].object_references,
        [60, 38]
    );

    let no_op = editor.to_bytes().unwrap();
    editor.set_slide_layout(0, title_only).unwrap();
    assert_eq!(editor.to_bytes().unwrap(), no_op);

    editor.set_slide_layout(0, title_and_bullets).unwrap();
    let slide = &editor.slides().unwrap()[0];
    assert_eq!(slide.is_title_visible, Some(true));
    assert_eq!(slide.is_body_visible, Some(true));
    assert_eq!(slide.title.as_deref(), Some("Old title"));
    assert_eq!(slide.body.as_deref(), Some("Old body 🚀"));
    let graph = ObjectGraph::read(editor.package()).unwrap();
    let native_slide: kn::SlideArchive = graph.decode_type(4, 5, "KN.SlideArchive").unwrap();
    assert_eq!(
        native_slide.owned_drawables,
        vec![reference(5), reference(6)]
    );
    assert_eq!(
        native_slide.drawables_z_order,
        vec![reference(5), reference(6)]
    );
    let title: kn::PlaceholderArchive = graph
        .decode_type(5, TEST_PLACEHOLDER_MESSAGE_TYPE, "KN.PlaceholderArchive")
        .unwrap();
    let body: kn::PlaceholderArchive = graph
        .decode_type(6, TEST_PLACEHOLDER_MESSAGE_TYPE, "KN.PlaceholderArchive")
        .unwrap();
    assert_eq!(placeholder_x(&title), Some(100.0));
    assert_eq!(placeholder_x(&body), Some(200.0));
}

#[test]
fn slide_layout_update_materializes_and_removes_image_graphs() {
    let mut editor = KeynoteEditor::from_package(test_package_with_image_layout()).unwrap();
    let layouts = editor.slide_layouts().unwrap();
    let photo = layouts
        .iter()
        .find(|layout| layout.name == "Title & Photo")
        .unwrap()
        .id;
    let bullets = layouts
        .iter()
        .find(|layout| layout.name == "Title & Bullets")
        .unwrap()
        .id;

    editor.set_slide_layout(0, photo).unwrap();
    let graph = ObjectGraph::read(editor.package()).unwrap();
    let slide: kn::SlideArchive = graph.decode_type(4, 5, "KN.SlideArchive").unwrap();
    assert_eq!(
        slide
            .owned_drawables
            .iter()
            .map(|reference| reference.identifier)
            .collect::<Vec<_>>(),
        [74, 5, 6]
    );
    let image: tsd::ImageArchive = graph
        .decode_type(74, TEST_IMAGE_MESSAGE_TYPE, "TSD.ImageArchive")
        .unwrap();
    assert_eq!(image.super_.parent, Some(reference(4)));
    assert_eq!(image.super_.title, Some(reference(75)));
    assert_eq!(image.super_.caption, Some(reference(76)));
    assert_eq!(image.mask, Some(reference(77)));
    assert_eq!(image.flags, Some(1));
    assert_eq!(image.data, Some(tsp::DataReference { identifier: 1_001 }));
    assert_eq!(
        editor.slides().unwrap()[0].title.as_deref(),
        Some("Old title")
    );
    assert_eq!(
        editor.slides().unwrap()[0].body.as_deref(),
        Some("Old body 🚀")
    );

    editor.set_slide_layout(0, bullets).unwrap();
    let graph = ObjectGraph::read(editor.package()).unwrap();
    let slide: kn::SlideArchive = graph.decode_type(4, 5, "KN.SlideArchive").unwrap();
    assert_eq!(slide.owned_drawables, [reference(5), reference(6)]);
    for identifier in 74..=77 {
        assert!(!graph.objects.contains_key(&identifier));
    }
}

#[test]
fn creates_slide_with_materialized_layout_image_graph() {
    let mut editor = KeynoteEditor::from_package(test_package_with_image_layout()).unwrap();
    let photo = editor
        .slide_layouts()
        .unwrap()
        .into_iter()
        .find(|layout| layout.name == "Title & Photo")
        .unwrap()
        .id;
    let created = editor.insert_slide(1, photo).unwrap();
    let graph = ObjectGraph::read(editor.package()).unwrap();
    let slide: kn::SlideArchive = graph
        .decode_type(created.slide_id, TEST_SLIDE_MESSAGE_TYPE, "KN.SlideArchive")
        .unwrap();
    let image_id = slide.owned_drawables[0].identifier;
    let image: tsd::ImageArchive = graph
        .decode_type(image_id, TEST_IMAGE_MESSAGE_TYPE, "TSD.ImageArchive")
        .unwrap();
    assert_eq!(image.super_.parent, Some(reference(created.slide_id)));
    assert_eq!(image.flags, Some(1));
    assert_eq!(created.title, None);
    assert_eq!(created.body, None);
}

#[test]
fn slide_layout_update_materializes_and_removes_live_video_graphs() {
    let mut editor = KeynoteEditor::from_package(test_package_with_live_video_layout()).unwrap();
    let layouts = editor.slide_layouts().unwrap();
    let live_video = layouts
        .iter()
        .find(|layout| layout.name == "Title & Live Video")
        .unwrap()
        .id;
    let bullets = layouts
        .iter()
        .find(|layout| layout.name == "Title & Bullets")
        .unwrap()
        .id;

    editor.set_slide_layout(0, live_video).unwrap();
    let graph = ObjectGraph::read(editor.package()).unwrap();
    let slide: kn::SlideArchive = graph.decode_type(4, 5, "KN.SlideArchive").unwrap();
    assert_eq!(
        slide
            .owned_drawables
            .iter()
            .map(|reference| reference.identifier)
            .collect::<Vec<_>>(),
        [73, 5, 6]
    );
    let movie: tsd::MovieArchive = graph
        .decode_type(73, TEST_MOVIE_MESSAGE_TYPE, "TSD.MovieArchive")
        .unwrap();
    assert_eq!(movie.super_.parent, Some(reference(4)));
    assert_eq!(movie.super_.title, Some(reference(74)));
    assert_eq!(movie.super_.caption, Some(reference(75)));
    assert_eq!(movie.style, Some(reference(60)));
    assert_eq!(
        movie.poster_image_data,
        Some(tsp::DataReference { identifier: 2_001 })
    );
    assert_eq!(movie.is_live_video, Some(true));
    let movie_data = graph
        .message_data_type(73, TEST_MOVIE_MESSAGE_TYPE, "TSD.MovieArchive")
        .unwrap();
    assert_eq!(
        repeated_length_delimited_payloads(movie_data, TEST_LIVE_VIDEO_INFO_FIELD).unwrap(),
        [TEST_LIVE_VIDEO_INFO_PAYLOAD]
    );

    editor.set_slide_layout(0, bullets).unwrap();
    let graph = ObjectGraph::read(editor.package()).unwrap();
    let slide: kn::SlideArchive = graph.decode_type(4, 5, "KN.SlideArchive").unwrap();
    assert_eq!(slide.owned_drawables, [reference(5), reference(6)]);
    for identifier in 73..=75 {
        assert!(!graph.objects.contains_key(&identifier));
    }
}

#[test]
fn creates_slide_with_materialized_layout_live_video_graph() {
    let mut editor = KeynoteEditor::from_package(test_package_with_live_video_layout()).unwrap();
    let live_video = editor
        .slide_layouts()
        .unwrap()
        .into_iter()
        .find(|layout| layout.name == "Title & Live Video")
        .unwrap()
        .id;
    let created = editor.insert_slide(1, live_video).unwrap();
    let graph = ObjectGraph::read(editor.package()).unwrap();
    let slide: kn::SlideArchive = graph
        .decode_type(created.slide_id, TEST_SLIDE_MESSAGE_TYPE, "KN.SlideArchive")
        .unwrap();
    let movie_id = slide.owned_drawables[0].identifier;
    let movie: tsd::MovieArchive = graph
        .decode_type(movie_id, TEST_MOVIE_MESSAGE_TYPE, "TSD.MovieArchive")
        .unwrap();
    assert_eq!(movie.super_.parent, Some(reference(created.slide_id)));
    assert_ne!(movie.super_.title, Some(reference(71)));
    assert_ne!(movie.super_.caption, Some(reference(72)));
    assert_eq!(movie.is_live_video, Some(true));
    assert_eq!(created.title, None);
    assert_eq!(created.body, None);
}

#[test]
fn slide_layout_update_materializes_and_removes_file_movie_placeholders() {
    let mut editor = KeynoteEditor::from_package(test_package_with_file_movie_layout(1)).unwrap();
    let layouts = editor.slide_layouts().unwrap();
    let movie_layout = layouts
        .iter()
        .find(|layout| layout.name == "Title & Movie Placeholder")
        .unwrap()
        .id;
    let bullets = layouts
        .iter()
        .find(|layout| layout.name == "Title & Bullets")
        .unwrap()
        .id;

    editor.set_slide_layout(0, movie_layout).unwrap();
    let graph = ObjectGraph::read(editor.package()).unwrap();
    let slide: kn::SlideArchive = graph.decode_type(4, 5, "KN.SlideArchive").unwrap();
    assert_eq!(
        slide
            .owned_drawables
            .iter()
            .map(|reference| reference.identifier)
            .collect::<Vec<_>>(),
        [73]
    );
    let movie: tsd::MovieArchive = graph
        .decode_type(73, TEST_MOVIE_MESSAGE_TYPE, "TSD.MovieArchive")
        .unwrap();
    assert_eq!(movie.super_.parent, Some(reference(4)));
    assert_eq!(movie.super_.title, Some(reference(74)));
    assert_eq!(movie.super_.caption, Some(reference(75)));
    assert_eq!(
        movie.movie_data,
        Some(tsp::DataReference { identifier: 2_002 })
    );
    assert_eq!(
        movie.poster_image_data,
        Some(tsp::DataReference { identifier: 2_001 })
    );
    assert_eq!(movie.flags, Some(1));
    assert_ne!(movie.is_live_video, Some(true));

    editor.set_slide_layout(0, bullets).unwrap();
    let graph = ObjectGraph::read(editor.package()).unwrap();
    let slide: kn::SlideArchive = graph.decode_type(4, 5, "KN.SlideArchive").unwrap();
    assert_eq!(slide.owned_drawables, [reference(5), reference(6)]);
    for identifier in 73..=75 {
        assert!(!graph.objects.contains_key(&identifier));
    }
}

#[test]
fn creates_slide_with_materialized_file_movie_placeholder_graph() {
    let mut editor = KeynoteEditor::from_package(test_package_with_file_movie_layout(1)).unwrap();
    let movie_layout = editor
        .slide_layouts()
        .unwrap()
        .into_iter()
        .find(|layout| layout.name == "Title & Movie Placeholder")
        .unwrap()
        .id;
    let created = editor.insert_slide(1, movie_layout).unwrap();
    let graph = ObjectGraph::read(editor.package()).unwrap();
    let slide: kn::SlideArchive = graph
        .decode_type(created.slide_id, TEST_SLIDE_MESSAGE_TYPE, "KN.SlideArchive")
        .unwrap();
    let movie_id = slide.owned_drawables[0].identifier;
    let movie: tsd::MovieArchive = graph
        .decode_type(movie_id, TEST_MOVIE_MESSAGE_TYPE, "TSD.MovieArchive")
        .unwrap();
    assert_eq!(movie.super_.parent, Some(reference(created.slide_id)));
    assert_ne!(movie.super_.title, Some(reference(71)));
    assert_ne!(movie.super_.caption, Some(reference(72)));
    assert_eq!(movie.flags, Some(1));
    assert_eq!(created.title, None);
    assert_eq!(created.body, None);
}

#[test]
fn static_layout_movies_remain_template_only_for_update_and_creation() {
    let mut editor = KeynoteEditor::from_package(test_package_with_file_movie_layout(0)).unwrap();
    let movie_layout = editor
        .slide_layouts()
        .unwrap()
        .into_iter()
        .find(|layout| layout.name == "Title & Static Movie")
        .unwrap()
        .id;

    editor.set_slide_layout(0, movie_layout).unwrap();
    let created = editor.insert_slide(1, movie_layout).unwrap();
    let graph = ObjectGraph::read(editor.package()).unwrap();
    for slide_id in [4, created.slide_id] {
        let slide: kn::SlideArchive = graph
            .decode_type(slide_id, TEST_SLIDE_MESSAGE_TYPE, "KN.SlideArchive")
            .unwrap();
        assert_eq!(
            slide
                .owned_drawables
                .iter()
                .filter(|reference| graph
                    .message_data_type(
                        reference.identifier,
                        TEST_MOVIE_MESSAGE_TYPE,
                        "TSD.MovieArchive"
                    )
                    .is_ok())
                .count(),
            0
        );
        assert_eq!(slide.template_slide, Some(reference(38)));
    }
    let template_movie: tsd::MovieArchive = graph
        .decode_type(70, TEST_MOVIE_MESSAGE_TYPE, "TSD.MovieArchive")
        .unwrap();
    assert_eq!(template_movie.flags, Some(0));
}

#[test]
fn slide_movie_crud_preserves_shared_assets_and_culls_final_references() {
    let mut editor = KeynoteEditor::from_package(test_package_with_slide_movie()).unwrap();
    let movies = editor.slide_movies(0).unwrap();
    assert_eq!(movies.len(), 1);
    let original = movies[0].clone();
    assert_eq!(original.kind, MovieKind::File);
    assert_eq!(original.drawable_object_id, 70);
    assert_eq!(original.movie_data_identifier, Some(1));
    assert_eq!(original.poster_image_data_identifier, Some(2));
    assert_eq!(
        original.original_size,
        Some(DrawableSize {
            width: 800.0,
            height: 300.0,
        })
    );
    assert_eq!(
        original.natural_size,
        Some(DrawableSize {
            width: 800.0,
            height: 300.0,
        })
    );
    assert_eq!(editor.slide_builds(0).unwrap()[0].chunks.len(), 1);

    let baseline = editor.to_bytes().unwrap();
    let moved = DrawableGeometry {
        position: Some(DrawablePoint { x: 120.0, y: 130.0 }),
        ..original.geometry
    };
    editor.set_slide_movie_geometry(0, 70, moved).unwrap();
    assert_eq!(editor.slide_movie_geometry(0, 70).unwrap(), moved);
    editor
        .set_slide_movie_geometry(0, 70, original.geometry)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    let duplicate = editor.duplicate_slide_movie(0, 70).unwrap();
    assert_ne!(duplicate.drawable_object_id, 70);
    assert_eq!(
        duplicate.movie_data_identifier,
        original.movie_data_identifier
    );
    assert_eq!(
        duplicate.poster_image_data_identifier,
        original.poster_image_data_identifier
    );
    assert_eq!(
        duplicate.geometry.position,
        Some(DrawablePoint { x: 110.0, y: 210.0 })
    );
    let duplicate_build = editor
        .slide_builds(0)
        .unwrap()
        .into_iter()
        .find(|build| build.drawable_object_id == duplicate.drawable_object_id)
        .unwrap();
    assert_eq!(duplicate_build.chunks.len(), 1);
    let graph = ObjectGraph::read(editor.package()).unwrap();
    let source_chunk: kn::BuildChunkArchive = graph
        .decode_type(74, BUILD_CHUNK_MESSAGE_TYPE, "KN.BuildChunkArchive")
        .unwrap();
    let cloned_chunk: kn::BuildChunkArchive = graph
        .decode_type(
            duplicate_build.chunks[0].object_id,
            BUILD_CHUNK_MESSAGE_TYPE,
            "KN.BuildChunkArchive",
        )
        .unwrap();
    assert_ne!(source_chunk.build_id, cloned_chunk.build_id);
    assert_eq!(
        cloned_chunk.build_id,
        cloned_chunk
            .build_chunk_identifier
            .as_ref()
            .and_then(|identifier| identifier.build_id)
    );
    let assets = editor.media_assets().unwrap();
    for raw_identifier in [1, 2] {
        let identifier = MediaAssetId::try_from(raw_identifier).expect("valid media ID");
        let asset = assets
            .iter()
            .find(|asset| asset.data_identifier == identifier)
            .unwrap_or_else(|| panic!("missing media {raw_identifier} in {assets:?}"));
        assert_eq!(asset.component_reference_count, 2);
        assert_eq!(asset.message_reference_count, 2);
    }

    assert_eq!(
        editor
            .replace_slide_movie_data(
                0,
                duplicate.drawable_object_id,
                TEST_MOVIE_VIDEO_REPLACEMENT,
            )
            .unwrap(),
        TEST_MOVIE_VIDEO
    );
    assert_eq!(
        editor.extract_media(1).unwrap(),
        TEST_MOVIE_VIDEO_REPLACEMENT
    );
    assert_eq!(
        editor
            .replace_slide_movie_poster(0, 70, TEST_MOVIE_POSTER_REPLACEMENT)
            .unwrap(),
        TEST_MOVIE_POSTER
    );
    assert_eq!(
        editor.extract_media(2).unwrap(),
        TEST_MOVIE_POSTER_REPLACEMENT
    );

    let removed_duplicate = editor
        .remove_slide_movie(0, duplicate.drawable_object_id)
        .unwrap();
    assert!(removed_duplicate.removed_data_identifiers.is_empty());
    let remaining = editor.slide_movies(0).unwrap();
    assert_eq!(remaining.len(), 1);
    assert_eq!(remaining[0], original);
    assert_eq!(editor.media_assets().unwrap().len(), 2);
    assert_eq!(editor.slide_builds(0).unwrap().len(), 1);

    let removed_original = editor.remove_slide_movie(0, 70).unwrap();
    assert_eq!(removed_original.movie, original);
    assert_eq!(removed_original.removed_data_identifiers, [1, 2]);
    assert!(editor.slide_movies(0).unwrap().is_empty());
    assert!(editor.slide_builds(0).unwrap().is_empty());
    assert!(editor.media_assets().unwrap().is_empty());
}

#[test]
fn slide_movie_playback_can_initialize_legacy_file_metadata() {
    let mut editor = KeynoteEditor::from_package(test_package_with_slide_movie()).unwrap();
    let movie = editor.slide_movies(0).unwrap().pop().unwrap();
    assert_eq!(movie.kind, MovieKind::File);
    assert_eq!(movie.playback, None);

    let settings = MediaPlaybackSettings::new(Duration::from_secs(8))
        .with_loop_mode(Some(MediaLoopMode::Repeat))
        .with_volume(Some(MediaVolume::new(0.75).unwrap()));
    editor
        .set_slide_movie_playback_settings(0, movie.drawable_object_id, settings)
        .unwrap();
    assert_eq!(
        editor
            .slide_movie_playback_settings(0, movie.drawable_object_id)
            .unwrap(),
        settings
    );
}

#[test]
fn slide_movie_mutations_reject_wrong_targets_transactionally() {
    let mut editor = KeynoteEditor::from_package(test_package_with_slide_movie()).unwrap();
    let before = editor.to_bytes().unwrap();
    assert!(editor.slide_movies(2).is_err());
    assert!(editor.slide_movie_geometry(0, 5).is_err());
    assert!(editor.duplicate_slide_movie(0, 5).is_err());
    assert!(editor.remove_slide_movie(0, 5).is_err());
    assert!(
        editor
            .replace_slide_movie_data(0, 70, TEST_MOVIE_POSTER)
            .is_err()
    );
    assert!(
        editor
            .replace_slide_movie_poster(0, 70, TEST_MOVIE_VIDEO)
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before);

    let mut placeholder =
        KeynoteEditor::from_package(test_package_with_file_movie_layout(1)).unwrap();
    let layout = placeholder
        .slide_layouts()
        .unwrap()
        .into_iter()
        .find(|layout| layout.name == "Title & Movie Placeholder")
        .unwrap()
        .id;
    placeholder.set_slide_layout(0, layout).unwrap();
    let movie = placeholder.slide_movies(0).unwrap().remove(0);
    assert_eq!(movie.kind, MovieKind::Placeholder);
    let before = placeholder.to_bytes().unwrap();
    assert!(
        placeholder
            .set_slide_movie_geometry(0, movie.drawable_object_id, movie.geometry)
            .is_err()
    );
    assert!(
        placeholder
            .duplicate_slide_movie(0, movie.drawable_object_id)
            .is_err()
    );
    assert!(
        placeholder
            .remove_slide_movie(0, movie.drawable_object_id)
            .is_err()
    );
    assert_eq!(placeholder.to_bytes().unwrap(), before);
}

#[test]
fn slide_layout_update_rejects_ambiguous_live_videos_transactionally() {
    let mut editor = KeynoteEditor::from_package(test_package_with_live_video_layout()).unwrap();
    let layouts = editor.slide_layouts().unwrap();
    let live_video = layouts
        .iter()
        .find(|layout| layout.name == "Title & Live Video")
        .unwrap()
        .id;
    let bullets = layouts
        .iter()
        .find(|layout| layout.name == "Title & Bullets")
        .unwrap()
        .id;
    editor.set_slide_layout(0, live_video).unwrap();
    let mut package = editor.into_package();
    package
        .update_archive("Index/Slide-4.iwa", |archive| {
            let source = archive.object(73).unwrap();
            let duplicate =
                clone_object_metadata(source, 76, source.messages.clone(), &HashMap::new(), false)?;
            archive.insert_object(duplicate)?;
            let slide = archive.object_mut(4).unwrap();
            let mut decoded = kn::SlideArchive::decode(slide.messages[0].data.as_slice())?;
            decoded.owned_drawables.push(reference(76));
            decoded.drawables_z_order.push(reference(76));
            slide.replace_message(
                0,
                RawMessage {
                    type_: TEST_SLIDE_MESSAGE_TYPE,
                    data: decoded.encode_to_vec(),
                },
            )?;
            slide.archive_info.message_infos[0]
                .object_references
                .push(76);
            Ok(())
        })
        .unwrap();
    let mut editor = KeynoteEditor::from_package(package).unwrap();
    let before = editor.to_bytes().unwrap();
    assert!(editor.set_slide_layout(0, bullets).is_err());
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn slide_layout_update_rejects_ambiguous_materialized_images_transactionally() {
    let mut editor = KeynoteEditor::from_package(test_package_with_image_layout()).unwrap();
    let layouts = editor.slide_layouts().unwrap();
    let photo = layouts
        .iter()
        .find(|layout| layout.name == "Title & Photo")
        .unwrap()
        .id;
    let bullets = layouts
        .iter()
        .find(|layout| layout.name == "Title & Bullets")
        .unwrap()
        .id;
    editor.set_slide_layout(0, photo).unwrap();
    let mut package = editor.package().clone();
    package
        .update_archive("Index/Slide-4.iwa", |archive| {
            let source = archive.object(74).unwrap();
            let duplicate =
                clone_object_metadata(source, 78, source.messages.clone(), &HashMap::new(), false)?;
            archive.insert_object(duplicate)?;
            let slide = archive.object_mut(4).unwrap();
            let mut decoded = kn::SlideArchive::decode(slide.messages[0].data.as_slice())?;
            decoded.owned_drawables.push(reference(78));
            decoded.drawables_z_order.push(reference(78));
            slide.replace_message(
                0,
                RawMessage {
                    type_: TEST_SLIDE_MESSAGE_TYPE,
                    data: decoded.encode_to_vec(),
                },
            )?;
            slide.archive_info.message_infos[0]
                .object_references
                .push(78);
            Ok(())
        })
        .unwrap();
    editor = KeynoteEditor::from_package(package).unwrap();
    let before = editor.to_bytes().unwrap();
    assert!(editor.set_slide_layout(0, bullets).is_err());
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn slide_layout_update_rejects_invalid_inputs_without_mutation() {
    let mut editor = KeynoteEditor::from_package(test_package_with_two_layouts()).unwrap();
    let before = editor.to_bytes().unwrap();
    assert!(
        editor
            .set_slide_layout(2, KeynoteSlideLayoutId::new(37).unwrap())
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before);
    assert!(
        editor
            .set_slide_layout(0, KeynoteSlideLayoutId::new(999).unwrap())
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn slide_layout_update_hides_a_retained_placeholder_missing_from_the_layout() {
    let mut package = test_package_with_two_layouts();
    package
        .update_archive("Index/TemplateSlide-38.iwa", |archive| {
            let slide = archive.object_mut(38).unwrap();
            let mut decoded = kn::SlideArchive::decode(slide.messages[0].data.as_slice())?;
            decoded.body_placeholder = None;
            slide.replace_message(
                0,
                RawMessage {
                    type_: TEST_SLIDE_MESSAGE_TYPE,
                    data: decoded.encode_to_vec(),
                },
            )?;
            Ok(())
        })
        .unwrap();
    let mut editor = KeynoteEditor::from_package(package).unwrap();
    editor
        .set_slide_layout(0, KeynoteSlideLayoutId::new(37).unwrap())
        .unwrap();
    let slide = &editor.slides().unwrap()[0];
    assert_eq!(slide.is_body_visible, Some(false));
    assert_eq!(slide.body.as_deref(), Some("Old body 🚀"));
    let graph = ObjectGraph::read(editor.package()).unwrap();
    let body: kn::PlaceholderArchive = graph
        .decode_type(6, TEST_PLACEHOLDER_MESSAGE_TYPE, "KN.PlaceholderArchive")
        .unwrap();
    assert_eq!(placeholder_x(&body), Some(200.0));
}

#[test]
fn slide_layout_update_rejects_ambiguous_wire_fields_transactionally() {
    let mut package = test_package_with_two_layouts();
    package
        .update_archive("Index/Slide-4.iwa", |archive| {
            let slide = archive.object_mut(4).unwrap();
            let message = slide.messages[0].clone();
            let data = append_repeated_length_delimited_field(
                &message.data,
                1,
                &reference(50).encode_to_vec(),
            )?;
            slide.replace_message(
                0,
                RawMessage {
                    type_: message.type_,
                    data,
                },
            )?;
            Ok(())
        })
        .unwrap();
    let mut editor = KeynoteEditor::from_package(package).unwrap();
    let before = editor.to_bytes().unwrap();
    assert!(
        editor
            .set_slide_layout(0, KeynoteSlideLayoutId::new(37).unwrap())
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn current_slide_layout_rejects_missing_theme_mapping() {
    let mut package = test_package_with_current_layout();
    package
        .update_archive("Index/Slide-4.iwa", |archive| {
            let slide = archive.object_mut(4).unwrap();
            let mut decoded = kn::SlideArchive::decode(slide.messages[0].data.as_slice())?;
            decoded.template_slide = Some(reference(999));
            slide.replace_message(
                0,
                RawMessage {
                    type_: TEST_SLIDE_MESSAGE_TYPE,
                    data: decoded.encode_to_vec(),
                },
            )?;
            Ok(())
        })
        .unwrap();
    assert!(KeynoteEditor::from_package(package).is_err());
}

#[test]
fn current_slide_layout_rejects_multiple_theme_mappings() {
    let mut package = test_package_with_current_layout();
    package
        .update_archive("Index/Document.iwa", |archive| {
            let theme = archive.object_mut(29).unwrap();
            let mut decoded = kn::ThemeArchive::decode(theme.messages[0].data.as_slice())?;
            decoded.templates.push(reference(37));
            theme.replace_message(
                0,
                RawMessage {
                    type_: TEST_THEME_MESSAGE_TYPE,
                    data: decoded.encode_to_vec(),
                },
            )?;
            archive.insert_object(object(
                37,
                TEST_SLIDE_NODE_MESSAGE_TYPE,
                kn::SlideNodeArchive {
                    slide: Some(reference(31)),
                    ..Default::default()
                },
            ))?;
            Ok(())
        })
        .unwrap();
    assert!(KeynoteEditor::from_package(package).is_err());
}

#[test]
fn creates_empty_slide_from_typed_theme_layout_transactionally() {
    let mut editor = KeynoteEditor::from_package(test_package_with_theme()).unwrap();
    let layouts = editor.slide_layouts().unwrap();
    assert_eq!(
        layouts,
        [KeynoteSlideLayoutInfo {
            id: KeynoteSlideLayoutId::new(30).unwrap(),
            name: "Title & Bullets".to_owned(),
            is_default: true,
        }]
    );
    assert_eq!(editor.default_slide_layout().unwrap(), layouts[0].id);

    let before = editor.to_bytes().unwrap();
    assert!(editor.insert_slide(3, layouts[0].id).is_err());
    assert!(
        editor
            .insert_slide(1, KeynoteSlideLayoutId::new(999).unwrap())
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before);

    let created = editor.insert_slide(1, layouts[0].id).unwrap();
    assert_eq!(created.index, 1);
    assert_eq!(created.layout, Some(layouts[0].clone()));
    assert_eq!(created.title.as_deref(), Some(""));
    assert_eq!(created.body.as_deref(), Some(""));
    assert_eq!(created.notes.as_deref(), Some(""));
    let slides = editor.slides().unwrap();
    assert_eq!(slides.len(), 3);
    assert_eq!(slides[0].title.as_deref(), Some("Old title"));
    assert_eq!(slides[0].notes.as_deref(), Some("Speaker 🚀"));
    let graph = ObjectGraph::read(editor.package()).unwrap();
    let slide: kn::SlideArchive = graph
        .decode_type(created.slide_id, 5, "KN.SlideArchive")
        .unwrap();
    assert_eq!(slide.template_slide, Some(reference(31)));
    assert_eq!(slide.name, None);
    assert!(slide.in_document);
    assert_eq!(slide.title_placeholder, Some(reference(39)));
    assert_eq!(slide.body_placeholder, Some(reference(40)));
    assert_eq!(graph.drawable_storage(39).unwrap(), Some(41));
    assert_eq!(graph.drawable_storage(40).unwrap(), Some(42));
    assert_eq!(graph.storage_text(34).unwrap(), "Slide Title");
    assert_eq!(graph.storage_text(35).unwrap(), "Slide bullet text");

    assert_eq!(editor.remove_slide(1).unwrap().slide_id, created.slide_id);
    let graph = ObjectGraph::read(editor.package()).unwrap();
    assert_eq!(graph.storage_text(34).unwrap(), "Slide Title");
    assert_eq!(graph.storage_text(35).unwrap(), "Slide bullet text");
}

#[test]
fn slide_tree_clone_delete_preserves_unknown_reference_bytes() {
    let mut editor = KeynoteEditor::from_package(test_package_with_unknown_slide_tree()).unwrap();
    let package_before = editor
        .package()
        .entry_names()
        .map(|name| {
            (
                name.to_owned(),
                editor.package().entry(name).unwrap().to_vec(),
            )
        })
        .collect::<Vec<_>>();
    let created = editor.duplicate_slide(0).unwrap();
    assert_eq!(editor.slides().unwrap()[1].slide_id, created.slide_id);
    assert_eq!(editor.remove_slide(1).unwrap().slide_id, created.slide_id);
    let package_after = editor
        .package()
        .entry_names()
        .map(|name| {
            (
                name.to_owned(),
                editor.package().entry(name).unwrap().to_vec(),
            )
        })
        .collect::<Vec<_>>();
    assert_eq!(package_after, package_before);
}

#[test]
fn duplicate_slide_tree_references_fail_transactionally() {
    let mut package = test_package_with_unknown_slide_tree();
    package
        .update_archive("Index/Document.iwa", |archive| {
            let object = archive.object_mut(2).unwrap();
            let original = object.messages[0].data.as_slice();
            let data = transform_length_delimited_field(original, 3, |slide_tree| {
                let mut references = repeated_length_delimited_payloads(slide_tree, 2)?
                    .into_iter()
                    .map(<[u8]>::to_vec)
                    .collect::<Vec<_>>();
                references.push(references[0].clone());
                rewrite_repeated_length_delimited_fields(slide_tree, 2, &references)
            })?;
            object.replace_message(
                0,
                RawMessage {
                    type_: object.messages[0].type_,
                    data,
                },
            )?;
            Ok(())
        })
        .unwrap();
    let mut editor = KeynoteEditor::from_package(package).unwrap();
    let before = editor.to_bytes().unwrap();
    assert!(editor.duplicate_slide(0).is_err());
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn duplicates_slide_graph_with_independent_objects() {
    let mut package = test_package_with_slide_background();
    package
        .update_archive("Index/Slide-4.iwa", |archive| {
            let slide_object = archive.object_mut(4).unwrap();
            slide_object.archive_info.message_infos[0].object_references = vec![40, 5, 6, 15];
            slide_object.archive_info.message_infos[0]
                .field_infos
                .push(FieldInfo {
                    path: FieldPath { path: vec![5] },
                    object_references: vec![5],
                    ..Default::default()
                });
            archive.object_mut(5).unwrap().archive_info.message_infos[0].object_references =
                vec![7];
            archive.object_mut(15).unwrap().archive_info.message_infos[0].object_references =
                vec![16];
            Ok(())
        })
        .unwrap();
    let mut editor = KeynoteEditor::from_package(package).unwrap();
    let source_component = editor
        .package()
        .entry("Index/Slide-4.iwa")
        .unwrap()
        .to_vec();

    let created = editor.duplicate_slide(0).unwrap();
    assert_eq!(created.index, 1);
    assert_ne!(created.node_id, 3);
    assert_ne!(created.slide_id, 4);
    assert_eq!(created.title.as_deref(), Some("Old title"));
    assert_eq!(created.body.as_deref(), Some("Old body 🚀"));
    assert_eq!(created.notes.as_deref(), Some("Speaker 🚀"));
    let component = editor
        .package()
        .archive(&format!("Index/Slide-{}.iwa", created.slide_id))
        .unwrap();
    let slide_object = component.object(created.slide_id).unwrap();
    let cloned_slide = kn::SlideArchive::decode(slide_object.messages[0].data.as_slice()).unwrap();
    let cloned_title = cloned_slide.title_placeholder.unwrap().identifier;
    assert_eq!(
        slide_object.archive_info.message_infos[0].field_infos[0].object_references,
        [cloned_title]
    );
    assert!(
        editor
            .package()
            .contains_entry(&format!("Index/Slide-{}.iwa", created.slide_id))
    );
    assert_eq!(
        editor.package().entry("Index/Slide-4.iwa").unwrap(),
        source_component
    );

    let focused = FocusedKeynotePackage::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    focused.validate().unwrap();
    let selector = SlideSelector::index(1);
    let mut title = focused.edit_slide_title(selector).unwrap();
    title.set("Independent copy").unwrap();
    let title = title.commit().unwrap();
    assert_eq!(
        title.package().slide_title(selector).unwrap().as_deref(),
        Some("Independent copy")
    );
    assert_eq!(
        title.package().slide_notes(selector).unwrap().as_deref(),
        Some("Speaker 🚀")
    );
    let mut notes = title.package().edit_slide_notes(selector).unwrap();
    notes.set("Independent notes").unwrap();
    let notes = notes.commit().unwrap();
    let mut editor = KeynoteEditor::from_bytes(notes.package().source_bytes()).unwrap();
    let slides = editor.slides().unwrap();
    assert_eq!(slides.len(), 3);
    assert_eq!(slides[0].title.as_deref(), Some("Old title"));
    assert_eq!(slides[0].notes.as_deref(), Some("Speaker 🚀"));
    assert_eq!(slides[1].title.as_deref(), Some("Independent copy"));
    assert_eq!(slides[1].notes.as_deref(), Some("Independent notes"));
    assert_eq!(slides[2].title.as_deref(), Some("Second title"));

    let removed = editor.remove_slide(1).unwrap();
    assert_eq!(removed.slide_id, created.slide_id);
    assert_eq!(editor.slides().unwrap().len(), 2);
    let before = editor.to_bytes().unwrap();
    assert!(editor.duplicate_slide(2).is_err());
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn cloned_slide_payloads_preserve_deep_unknown_fields_exactly() {
    let mut package = test_package();
    package
        .update_archive("Index/Slide-4.iwa", |archive| {
            let placeholder = archive.object(5).unwrap().messages[0].data.as_slice();
            let shape = repeated_length_delimited_payloads(placeholder, 1)?[0].to_vec();
            archive.insert_object(ArchiveObject::new(
                17,
                vec![RawMessage {
                    type_: 2011,
                    data: shape,
                }],
            )?)?;

            let storage = archive.object_mut(7).unwrap();
            let mut value = tswp::StorageArchive::decode(storage.messages[0].data.as_slice())?;
            value.style_sheet = Some(reference(8));
            storage.replace_message(
                0,
                RawMessage {
                    type_: 2001,
                    data: value.encode_to_vec(),
                },
            )?;

            for object in &mut archive.objects {
                let message_type = object.messages[0].type_;
                let paths: &[&[u32]] = match message_type {
                    5 => &[&[5], &[6], &[7], &[27]],
                    7 => &[&[1, 4]],
                    15 => &[&[1]],
                    2001 | 2022 => &[&[2]],
                    2011 => &[&[4]],
                    _ => &[],
                };
                let mut data = object.messages[0].data.clone();
                for path in paths {
                    data = transform_length_delimited_fields_at_path(&data, path, |reference| {
                        let mut reference = reference.to_vec();
                        append_unknown_varint(&mut reference, 98, 980);
                        Ok(reference)
                    })?;
                }
                append_unknown_varint(&mut data, 99, 990);
                object.replace_message(
                    0,
                    RawMessage {
                        type_: message_type,
                        data,
                    },
                )?;
            }
            Ok(())
        })
        .unwrap();
    package
        .update_archive("Index/Document.iwa", |archive| {
            let node = archive.object_mut(3).unwrap();
            let message_type = node.messages[0].type_;
            let mut data = transform_length_delimited_fields_at_path(
                node.messages[0].data.as_slice(),
                &[2],
                |reference| {
                    let mut reference = reference.to_vec();
                    append_unknown_varint(&mut reference, 98, 980);
                    Ok(reference)
                },
            )?;
            append_unknown_varint(&mut data, 99, 990);
            node.replace_message(
                0,
                RawMessage {
                    type_: message_type,
                    data,
                },
            )?;
            Ok(())
        })
        .unwrap();

    let mut editor = KeynoteEditor::from_package(package).unwrap();
    let source = editor.package().archive("Index/Slide-4.iwa").unwrap();
    let created = editor.duplicate_slide(0).unwrap();
    let cloned = editor
        .package()
        .archive(&format!("Index/Slide-{}.iwa", created.slide_id))
        .unwrap();
    assert_eq!(source.objects.len(), cloned.objects.len());
    let reverse = source
        .objects
        .iter()
        .zip(&cloned.objects)
        .map(|(source, cloned)| {
            (
                cloned.archive_info.identifier.unwrap(),
                source.archive_info.identifier.unwrap(),
            )
        })
        .collect::<HashMap<_, _>>();

    for (source, cloned) in source.objects.iter().zip(&cloned.objects) {
        assert_eq!(source.messages.len(), cloned.messages.len());
        for (source, cloned) in source.messages.iter().zip(&cloned.messages) {
            let restored = match cloned.type_ {
                5 => remap_slide_archive_wire(&cloned.data, &reverse).unwrap(),
                7 => remap_placeholder_archive_wire(&cloned.data, &reverse).unwrap(),
                15 => remap_note_archive_wire(&cloned.data, &reverse).unwrap(),
                2001 | 2022 => remap_storage_archive_wire(&cloned.data, &reverse).unwrap(),
                2011 => remap_shape_info_wire(&cloned.data, &reverse).unwrap(),
                _ => cloned.data.clone(),
            };
            assert_eq!(restored, source.data);
        }
    }

    let document = editor.package().archive("Index/Document.iwa").unwrap();
    let source_node = &document.object(3).unwrap().messages[0].data;
    let cloned_node = &document.object(created.node_id).unwrap().messages[0].data;
    let source_unknown = crate::wire::parse_wire_fields(source_node)
        .unwrap()
        .into_iter()
        .find(|field| field.number() == 99)
        .unwrap();
    let cloned_unknown = crate::wire::parse_wire_fields(cloned_node)
        .unwrap()
        .into_iter()
        .find(|field| field.number() == 99)
        .unwrap();
    assert_eq!(
        &source_node[source_unknown.start()..source_unknown.end()],
        &cloned_node[cloned_unknown.start()..cloned_unknown.end()]
    );
    let reversed_node = remap_reference_paths(
        cloned_node,
        &[&[2]],
        &HashMap::from([(created.slide_id, 4)]),
    )
    .unwrap();
    assert_eq!(
        repeated_length_delimited_payloads(source_node, 2).unwrap(),
        repeated_length_delimited_payloads(&reversed_node, 2).unwrap()
    );
}

#[test]
fn cloned_slide_rejects_duplicate_reference_identifiers_transactionally() {
    let mut package = test_package();
    package
        .update_archive("Index/Slide-4.iwa", |archive| {
            let slide = archive.object_mut(4).unwrap();
            let message_type = slide.messages[0].type_;
            let data = transform_length_delimited_fields_at_path(
                slide.messages[0].data.as_slice(),
                &[5],
                |reference| {
                    let mut reference = reference.to_vec();
                    reference.extend(litchi_iwa_common::varint::encode_varint(8));
                    reference.extend(litchi_iwa_common::varint::encode_varint(5));
                    Ok(reference)
                },
            )?;
            slide.replace_message(
                0,
                RawMessage {
                    type_: message_type,
                    data,
                },
            )?;
            Ok(())
        })
        .unwrap();
    let mut editor = KeynoteEditor::from_package(package).unwrap();
    let before = editor.to_bytes().unwrap();
    assert!(editor.duplicate_slide(0).is_err());
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn slide_owned_drawable_comment_crud_is_reachability_guarded() {
    let mut editor = KeynoteEditor::from_package(test_package()).unwrap();
    assert_eq!(
        editor
            .slide_drawables(0)
            .unwrap()
            .into_iter()
            .map(|drawable| drawable.id.get())
            .collect::<Vec<_>>(),
        vec![5, 6]
    );
    assert!(editor.slide_drawable_comment(0, 5).unwrap().is_none());
    assert!(
        editor
            .set_slide_drawable_comment(0, 11, "Wrong slide")
            .is_err()
    );

    editor
        .set_slide_drawable_comment(0, 5, "Title annotation")
        .unwrap();
    let comment = editor.slide_drawable_comment(0, 5).unwrap().unwrap();
    assert_eq!(comment.comment.text, "Title annotation");
    assert_eq!(
        editor.slides().unwrap()[0].title.as_deref(),
        Some("Old title")
    );
    let bytes = editor.to_bytes().unwrap();
    editor
        .set_slide_drawable_comment(0, 5, "Title annotation")
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), bytes);

    let mut reparsed = KeynoteEditor::from_bytes(&bytes).unwrap();
    assert_eq!(
        reparsed
            .slide_drawable_comment(0, 5)
            .unwrap()
            .unwrap()
            .comment
            .text,
        "Title annotation"
    );
    reparsed.clear_slide_drawable_comment(0, 5).unwrap();
    assert!(reparsed.slide_drawable_comment(0, 5).unwrap().is_none());
    assert_eq!(
        reparsed.slides().unwrap()[0].title.as_deref(),
        Some("Old title")
    );
}

fn reference(identifier: u64) -> Reference {
    Reference {
        identifier,
        ..Default::default()
    }
}

fn append_unknown_varint(data: &mut Vec<u8>, field_number: u32, value: u64) {
    data.extend(litchi_iwa_common::varint::encode_varint(
        u64::from(field_number) << 3,
    ));
    data.extend(litchi_iwa_common::varint::encode_varint(value));
}

fn unknown_varint(field_number: u32, value: u64) -> Vec<u8> {
    let mut data = Vec::new();
    append_unknown_varint(&mut data, field_number, value);
    data
}

fn append_unknown_fixed64(data: &mut Vec<u8>, field_number: u32, value: u64) {
    data.extend(litchi_iwa_common::varint::encode_varint(
        (u64::from(field_number) << 3) | 1,
    ));
    data.extend(value.to_le_bytes());
}

fn append_unknown_fixed32(data: &mut Vec<u8>, field_number: u32, value: u32) {
    data.extend(litchi_iwa_common::varint::encode_varint(
        (u64::from(field_number) << 3) | 5,
    ));
    data.extend(value.to_le_bytes());
}

#[allow(deprecated)]
fn test_package_with_text_box() -> IWorkPackage {
    let mut package = test_package();
    package
        .update_archive("Index/Slide-4.iwa", |archive| {
            let slide_object = archive.object_mut(4).unwrap();
            let mut slide =
                kn::SlideArchive::decode(slide_object.messages[0].data.as_slice()).unwrap();
            slide.owned_drawables.push(reference(17));
            slide.drawables_z_order.push(reference(17));
            slide_object.replace_message(
                0,
                RawMessage {
                    type_: 5,
                    data: slide.encode_to_vec(),
                },
            )?;
            slide_object.archive_info.message_infos[0]
                .object_references
                .push(17);
            let mut shape = object(
                17,
                2011,
                tswp::ShapeInfoArchive {
                    super_: crate::protobuf::tsd::ShapeArchive {
                        super_: crate::protobuf::tsd::DrawableArchive {
                            geometry: Some(crate::protobuf::tsd::GeometryArchive {
                                position: Some(crate::protobuf::tsp::Point { x: 100.0, y: 100.0 }),
                                size: Some(crate::protobuf::tsp::Size {
                                    width: 200.0,
                                    height: 60.0,
                                }),
                                flags: Some(0),
                                angle: Some(0.0),
                            }),
                            parent: Some(reference(4)),
                            title: Some(reference(22)),
                            caption: Some(reference(21)),
                            ..Default::default()
                        },
                        ..Default::default()
                    },
                    deprecated_storage: Some(reference(18)),
                    owned_storage: Some(reference(18)),
                    is_text_box: Some(true),
                    ..Default::default()
                },
            );
            shape.archive_info.message_infos[0]
                .object_references
                .extend([4, 21, 22, 18]);
            archive.insert_object(shape)?;
            archive.insert_object(object(21, 3097, Vec::new()))?;
            archive.insert_object(object(22, 3097, Vec::new()))?;
            Ok(archive.insert_object(object(
                18,
                2001,
                StorageArchive {
                    text: vec!["Independent text box".to_owned()],
                    ..Default::default()
                },
            ))?)
        })
        .unwrap();
    package
}

fn test_package_with_unknown_slide_tree() -> IWorkPackage {
    let mut package = test_package();
    package
        .update_archive("Index/Document.iwa", |archive| {
            let object = archive.object_mut(2).unwrap();
            let original = object.messages[0].data.as_slice();
            let mut data = transform_length_delimited_field(original, 3, |slide_tree| {
                let references = repeated_length_delimited_payloads(slide_tree, 2)?
                    .into_iter()
                    .enumerate()
                    .map(|(index, raw)| {
                        let mut raw = raw.to_vec();
                        append_unknown_varint(&mut raw, 97, 970 + index as u64);
                        raw
                    })
                    .collect::<Vec<_>>();
                let mut slide_tree =
                    rewrite_repeated_length_delimited_fields(slide_tree, 2, &references)?;
                append_unknown_varint(&mut slide_tree, 98, 980);
                Ok(slide_tree)
            })?;
            append_unknown_varint(&mut data, 99, 990);
            object.replace_message(
                0,
                RawMessage {
                    type_: object.messages[0].type_,
                    data,
                },
            )?;
            Ok(())
        })
        .unwrap();
    package
}

fn test_transition() -> kn::TransitionArchive {
    kn::TransitionArchive {
        attributes: kn::TransitionAttributesArchive {
            animation_attributes: Some(kn::AnimationAttributesArchive {
                animation_type: Some("Transition".to_owned()),
                effect: Some("none".to_owned()),
                duration: Some(1.0),
                delay: Some(0.5),
                is_automatic: Some(false),
                ..Default::default()
            }),
            ..Default::default()
        },
    }
}

fn object<T: Message>(identifier: u64, type_: u32, value: T) -> ArchiveObject {
    ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_,
            data: value.encode_to_vec(),
        }],
    )
    .unwrap()
}

fn test_package() -> IWorkPackage {
    let document = kn::DocumentArchive {
        show: reference(2),
        ..Default::default()
    };
    let show = kn::ShowArchive {
        slide_tree: kn::SlideTreeArchive {
            slides: vec![reference(3), reference(9)],
            ..Default::default()
        },
        size: crate::protobuf::tsp::Size {
            width: 1_024.0,
            height: 768.0,
        },
        ..Default::default()
    };
    let node = kn::SlideNodeArchive {
        slide: Some(reference(4)),
        ..Default::default()
    };
    let slide = kn::SlideArchive {
        title_placeholder: Some(reference(5)),
        body_placeholder: Some(reference(6)),
        owned_drawables: vec![reference(5), reference(6)],
        drawables_z_order: vec![reference(5), reference(6)],
        note: Some(reference(15)),
        transition: test_transition(),
        ..Default::default()
    };
    let title = kn::PlaceholderArchive {
        super_: tswp::ShapeInfoArchive {
            owned_storage: Some(reference(7)),
            ..Default::default()
        },
        ..Default::default()
    };
    let body = kn::PlaceholderArchive {
        super_: tswp::ShapeInfoArchive {
            owned_storage: Some(reference(8)),
            ..Default::default()
        },
        ..Default::default()
    };
    let second_node = kn::SlideNodeArchive {
        slide: Some(reference(10)),
        ..Default::default()
    };
    let second_slide = kn::SlideArchive {
        title_placeholder: Some(reference(11)),
        body_placeholder: Some(reference(12)),
        owned_drawables: vec![reference(11), reference(12)],
        drawables_z_order: vec![reference(11), reference(12)],
        transition: test_transition(),
        ..Default::default()
    };
    let second_title = kn::PlaceholderArchive {
        super_: tswp::ShapeInfoArchive {
            owned_storage: Some(reference(13)),
            ..Default::default()
        },
        ..Default::default()
    };
    let second_body = kn::PlaceholderArchive {
        super_: tswp::ShapeInfoArchive {
            owned_storage: Some(reference(14)),
            ..Default::default()
        },
        ..Default::default()
    };
    let mut package = IWorkPackage::new();
    package
        .replace_archive(
            "Index/Document.iwa",
            &Archive {
                objects: vec![
                    object(1, 1, document),
                    object(2, 2, show),
                    object(3, 4, node),
                    object(9, 4, second_node),
                ],
            },
        )
        .unwrap();
    package
        .replace_archive(
            "Index/Slide-4.iwa",
            &Archive {
                objects: vec![
                    object(4, 5, slide),
                    object(5, 7, title),
                    object(6, 7, body),
                    object(
                        7,
                        2001,
                        StorageArchive {
                            text: vec!["Old title".to_owned()],
                            ..Default::default()
                        },
                    ),
                    object(
                        8,
                        2001,
                        StorageArchive {
                            text: vec!["Old body 🚀".to_owned()],
                            ..Default::default()
                        },
                    ),
                    object(
                        15,
                        15,
                        kn::NoteArchive {
                            contained_storage: reference(16),
                        },
                    ),
                    object(
                        16,
                        2001,
                        StorageArchive {
                            text: vec!["Speaker 🚀".to_owned()],
                            ..Default::default()
                        },
                    ),
                ],
            },
        )
        .unwrap();
    package
        .replace_archive(
            "Index/Slide-10.iwa",
            &Archive {
                objects: vec![
                    object(10, 5, second_slide),
                    object(11, 7, second_title),
                    object(12, 7, second_body),
                    object(
                        13,
                        2001,
                        StorageArchive {
                            text: vec!["Second title".to_owned()],
                            ..Default::default()
                        },
                    ),
                    object(
                        14,
                        2001,
                        StorageArchive {
                            text: vec!["Second body".to_owned()],
                            ..Default::default()
                        },
                    ),
                ],
            },
        )
        .unwrap();
    package
}

fn test_package_with_soundtrack() -> IWorkPackage {
    let mut package = test_package();
    package
        .update_archive("Index/Document.iwa", |archive| {
            let show = archive.object_mut(2).unwrap();
            let mut decoded = kn::ShowArchive::decode(show.messages[0].data.as_slice())?;
            decoded.soundtrack = Some(reference(TEST_SOUNDTRACK_ID));
            show.replace_message(
                0,
                RawMessage {
                    type_: 2,
                    data: decoded.encode_to_vec(),
                },
            )?;
            show.archive_info.message_infos[0]
                .object_references
                .push(TEST_SOUNDTRACK_ID);
            let mut soundtrack = object(
                TEST_SOUNDTRACK_ID,
                TEST_SOUNDTRACK_MESSAGE_TYPE,
                kn::Soundtrack {
                    volume: Some(1.0),
                    mode: Some(TEST_SOUNDTRACK_PLAY_ONCE_MODE),
                    movie_media: TEST_SOUNDTRACK_MEDIA_IDS
                        .into_iter()
                        .map(|identifier| tsp::DataReference { identifier })
                        .collect(),
                },
            );
            append_unknown_varint(&mut soundtrack.messages[0].data, 99, 990);
            soundtrack.archive_info.message_infos[0]
                .data_references
                .extend(TEST_SOUNDTRACK_MEDIA_IDS);
            Ok(archive.insert_object(soundtrack)?)
        })
        .unwrap();
    package
}

fn test_package_with_empty_soundtrack() -> IWorkPackage {
    let mut package = test_package_with_soundtrack();
    package
        .update_archive("Index/Document.iwa", |archive| {
            let soundtrack = archive.object_mut(TEST_SOUNDTRACK_ID).unwrap();
            soundtrack.archive_info.message_infos[0]
                .data_references
                .clear();
            soundtrack.replace_message(
                0,
                RawMessage {
                    type_: TEST_SOUNDTRACK_MESSAGE_TYPE,
                    data: kn::Soundtrack {
                        volume: Some(1.0),
                        mode: Some(TEST_SOUNDTRACK_PLAY_ONCE_MODE),
                        movie_media: Vec::new(),
                    }
                    .encode_to_vec(),
                },
            )?;
            Ok(())
        })
        .unwrap();
    package
        .replace_archive(
            PACKAGE_METADATA_ENTRY,
            &Archive {
                objects: vec![object(
                    69,
                    PACKAGE_METADATA_MESSAGE_TYPE,
                    PackageMetadata {
                        last_object_identifier: 74,
                        components: vec![ComponentInfo {
                            identifier: 1,
                            preferred_locator: "Document".to_owned(),
                            ..Default::default()
                        }],
                        ..Default::default()
                    },
                )],
            },
        )
        .unwrap();
    package
}

fn test_package_with_slide_background() -> IWorkPackage {
    let mut package = test_package();
    for archive_name in ["Index/Slide-4.iwa", "Index/Slide-10.iwa"] {
        package
            .update_archive(archive_name, |archive| {
                let slide = archive
                    .objects
                    .iter_mut()
                    .find(|object| object.messages[0].type_ == TEST_SLIDE_MESSAGE_TYPE)
                    .unwrap();
                let mut native = kn::SlideArchive::decode(slide.messages[0].data.as_slice())?;
                native.style = reference(40);
                slide.replace_message(
                    0,
                    RawMessage {
                        type_: TEST_SLIDE_MESSAGE_TYPE,
                        data: native.encode_to_vec(),
                    },
                )?;
                slide.archive_info.message_infos[0]
                    .object_references
                    .push(40);
                Ok(())
            })
            .unwrap();
    }
    let mut fill_payload = tsd::FillArchive {
        color: Some(tsp::Color {
            model: tsp::color::ColorModel::Rgb as i32,
            r: Some(1.0),
            g: Some(1.0),
            b: Some(1.0),
            rgbspace: Some(tsp::color::RgbColorSpace::Srgb as i32),
            a: Some(1.0),
            ..Default::default()
        }),
        ..Default::default()
    }
    .encode_to_vec();
    append_unknown_varint(&mut fill_payload, 100, 73);
    let style = kn::SlideStyleArchive {
        super_: tss::StyleArchive {
            name: Some("template slide style".to_owned()),
            style_identifier: Some("slide-0-slidestyle".to_owned()),
            stylesheet: Some(reference(41)),
            ..Default::default()
        },
        override_count: Some(1),
        slide_properties: Some(kn::SlideStylePropertiesArchive {
            fill: Some(tsd::FillArchive::default()),
            ..Default::default()
        }),
    };
    let style_data = patch_nested_length_delimited_field(
        &style.encode_to_vec(),
        &[11, 1],
        true,
        Some(&fill_payload),
    )
    .unwrap();
    let mut style_object = ArchiveObject::new(
        40,
        vec![RawMessage {
            type_: 9,
            data: style_data,
        }],
    )
    .unwrap();
    style_object.archive_info.message_infos[0]
        .object_references
        .push(41);
    let mut stylesheet_object = object(
        41,
        401,
        tss::StylesheetArchive {
            styles: vec![reference(40)],
            can_cull_styles: Some(true),
            ..Default::default()
        },
    );
    stylesheet_object.archive_info.message_infos[0]
        .object_references
        .push(40);
    package
        .replace_archive(
            "Index/DocumentStylesheet.iwa",
            &Archive {
                objects: vec![style_object, stylesheet_object],
            },
        )
        .unwrap();
    package
}

fn test_package_with_theme() -> IWorkPackage {
    let mut package = test_package();
    package
        .update_archive("Index/Document.iwa", |archive| {
            let show = archive.object_mut(2).unwrap();
            let mut decoded = kn::ShowArchive::decode(show.messages[0].data.as_slice())?;
            decoded.theme = reference(29);
            show.replace_message(
                0,
                RawMessage {
                    type_: 2,
                    data: decoded.encode_to_vec(),
                },
            )?;
            archive.insert_object(object(
                29,
                10,
                kn::ThemeArchive {
                    templates: vec![reference(30)],
                    default_template_slide_node: Some(reference(30)),
                    ..Default::default()
                },
            ))?;
            archive.insert_object(object(
                30,
                4,
                kn::SlideNodeArchive {
                    slide: Some(reference(31)),
                    ..Default::default()
                },
            ))?;
            Ok(())
        })
        .unwrap();
    package
        .replace_archive(
            "Index/TemplateSlide-31.iwa",
            &Archive {
                objects: vec![
                    object(
                        31,
                        5,
                        kn::SlideArchive {
                            title_placeholder: Some(reference(32)),
                            body_placeholder: Some(reference(33)),
                            owned_drawables: vec![reference(32), reference(33)],
                            drawables_z_order: vec![reference(32), reference(33)],
                            name: Some("Title & Bullets".to_owned()),
                            user_defined_guide_storage: Some(reference(36)),
                            in_document: true,
                            ..Default::default()
                        },
                    ),
                    object(
                        32,
                        7,
                        kn::PlaceholderArchive {
                            super_: tswp::ShapeInfoArchive {
                                owned_storage: Some(reference(34)),
                                ..Default::default()
                            },
                            ..Default::default()
                        },
                    ),
                    object(
                        33,
                        7,
                        kn::PlaceholderArchive {
                            super_: tswp::ShapeInfoArchive {
                                owned_storage: Some(reference(35)),
                                ..Default::default()
                            },
                            ..Default::default()
                        },
                    ),
                    object(
                        34,
                        2_001,
                        StorageArchive {
                            text: vec!["Slide Title".to_owned()],
                            ..Default::default()
                        },
                    ),
                    object(
                        35,
                        2_001,
                        StorageArchive {
                            text: vec!["Slide bullet text".to_owned()],
                            ..Default::default()
                        },
                    ),
                    object(36, 3_047, tsd::GuideStorageArchive::default()),
                ],
            },
        )
        .unwrap();
    package
}

fn test_package_with_current_layout() -> IWorkPackage {
    let mut package = test_package_with_theme();
    package
        .update_archive("Index/Slide-4.iwa", |archive| {
            let slide = archive.object_mut(4).unwrap();
            let mut decoded = kn::SlideArchive::decode(slide.messages[0].data.as_slice())?;
            decoded.template_slide = Some(reference(31));
            slide.replace_message(
                0,
                RawMessage {
                    type_: TEST_SLIDE_MESSAGE_TYPE,
                    data: decoded.encode_to_vec(),
                },
            )?;
            slide.archive_info.message_infos[0]
                .object_references
                .push(31);
            Ok(())
        })
        .unwrap();
    package
}

fn test_package_with_two_layouts() -> IWorkPackage {
    let mut package = test_package_with_current_layout();
    package
        .update_archive("Index/Document.iwa", |archive| {
            let theme = archive.object_mut(29).unwrap();
            let mut decoded = kn::ThemeArchive::decode(theme.messages[0].data.as_slice())?;
            decoded.templates.push(reference(37));
            theme.replace_message(
                0,
                RawMessage {
                    type_: TEST_THEME_MESSAGE_TYPE,
                    data: decoded.encode_to_vec(),
                },
            )?;

            let source_node = archive.object_mut(30).unwrap();
            let mut decoded =
                kn::SlideNodeArchive::decode(source_node.messages[0].data.as_slice())?;
            decoded.template_slide_id = Some(test_uuid(1, 2));
            source_node.replace_message(
                0,
                RawMessage {
                    type_: TEST_SLIDE_NODE_MESSAGE_TYPE,
                    data: decoded.encode_to_vec(),
                },
            )?;

            let live_node = archive.object_mut(3).unwrap();
            let mut decoded = kn::SlideNodeArchive::decode(live_node.messages[0].data.as_slice())?;
            decoded.template_slide_id = Some(test_uuid(1, 2));
            live_node.replace_message(
                0,
                RawMessage {
                    type_: TEST_SLIDE_NODE_MESSAGE_TYPE,
                    data: decoded.encode_to_vec(),
                },
            )?;
            archive.insert_object(object(
                37,
                TEST_SLIDE_NODE_MESSAGE_TYPE,
                kn::SlideNodeArchive {
                    slide: Some(reference(38)),
                    template_slide_id: Some(test_uuid(3, 4)),
                    ..Default::default()
                },
            ))?;
            Ok(())
        })
        .unwrap();
    package
        .update_archive("Index/TemplateSlide-31.iwa", |archive| {
            let slide = archive.object_mut(31).unwrap();
            let mut decoded = kn::SlideArchive::decode(slide.messages[0].data.as_slice())?;
            decoded.style = reference(50);
            slide.replace_message(
                0,
                RawMessage {
                    type_: TEST_SLIDE_MESSAGE_TYPE,
                    data: decoded.encode_to_vec(),
                },
            )?;
            set_test_placeholder_presentation(archive.object_mut(32).unwrap(), 100.0, 51, 2)?;
            set_test_placeholder_presentation(archive.object_mut(33).unwrap(), 200.0, 52, 3)?;
            Ok(())
        })
        .unwrap();
    package
        .update_archive("Index/Slide-4.iwa", |archive| {
            let slide = archive.object_mut(4).unwrap();
            let mut decoded = kn::SlideArchive::decode(slide.messages[0].data.as_slice())?;
            decoded.style = reference(50);
            slide.replace_message(
                0,
                RawMessage {
                    type_: TEST_SLIDE_MESSAGE_TYPE,
                    data: decoded.encode_to_vec(),
                },
            )?;
            slide.archive_info.message_infos[0].object_references = vec![50, 31];
            set_test_placeholder_presentation(archive.object_mut(5).unwrap(), 100.0, 51, 2)?;
            set_test_placeholder_presentation(archive.object_mut(6).unwrap(), 200.0, 52, 3)?;
            Ok(())
        })
        .unwrap();
    package
        .replace_archive(
            "Index/TemplateSlide-38.iwa",
            &Archive {
                objects: vec![
                    object(
                        38,
                        TEST_SLIDE_MESSAGE_TYPE,
                        kn::SlideArchive {
                            style: reference(60),
                            title_placeholder: Some(reference(39)),
                            body_placeholder: Some(reference(40)),
                            owned_drawables: vec![reference(39)],
                            drawables_z_order: vec![reference(39)],
                            name: Some("Title Only".to_owned()),
                            in_document: true,
                            ..Default::default()
                        },
                    ),
                    object(
                        39,
                        TEST_PLACEHOLDER_MESSAGE_TYPE,
                        test_placeholder(None, 300.0, 61, 2),
                    ),
                    object(
                        40,
                        TEST_PLACEHOLDER_MESSAGE_TYPE,
                        test_placeholder(None, 400.0, 62, 3),
                    ),
                ],
            },
        )
        .unwrap();
    package
}

fn test_package_with_image_layout() -> IWorkPackage {
    let mut package = test_package_with_two_layouts();
    package
        .update_archive("Index/TemplateSlide-38.iwa", |archive| {
            let slide = archive.object_mut(38).unwrap();
            let mut decoded = kn::SlideArchive::decode(slide.messages[0].data.as_slice())?;
            decoded.name = Some("Title & Photo".to_owned());
            decoded.owned_drawables = vec![reference(70), reference(39), reference(40)];
            decoded.drawables_z_order = decoded.owned_drawables.clone();
            slide.replace_message(
                0,
                RawMessage {
                    type_: TEST_SLIDE_MESSAGE_TYPE,
                    data: decoded.encode_to_vec(),
                },
            )?;

            let mut image = object(
                70,
                TEST_IMAGE_MESSAGE_TYPE,
                tsd::ImageArchive {
                    super_: tsd::DrawableArchive {
                        geometry: Some(tsd::GeometryArchive {
                            position: Some(tsp::Point { x: 0.0, y: 0.0 }),
                            size: Some(tsp::Size {
                                width: 1_024.0,
                                height: 768.0,
                            }),
                            ..Default::default()
                        }),
                        parent: Some(reference(38)),
                        title: Some(reference(71)),
                        caption: Some(reference(72)),
                        accessibility_description: Some("Test photo".to_owned()),
                        ..Default::default()
                    },
                    data: Some(tsp::DataReference { identifier: 1_001 }),
                    thumbnail_data: Some(tsp::DataReference { identifier: 1_002 }),
                    style: Some(reference(60)),
                    mask: Some(reference(73)),
                    original_size: Some(tsp::Size {
                        width: 1_024.0,
                        height: 768.0,
                    }),
                    natural_size: Some(tsp::Size {
                        width: 2_048.0,
                        height: 1_536.0,
                    }),
                    flags: Some(3),
                    ..Default::default()
                },
            );
            image.archive_info.message_infos[0].object_references = vec![72, 71, 60, 73];
            image.archive_info.message_infos[0].data_references = vec![1_001, 1_002];
            archive.insert_object(image)?;
            archive.insert_object(object(
                71,
                STANDIN_CAPTION_MESSAGE_TYPE,
                tsd::StandinCaptionArchive::default(),
            ))?;
            archive.insert_object(object(
                72,
                STANDIN_CAPTION_MESSAGE_TYPE,
                tsd::StandinCaptionArchive::default(),
            ))?;
            archive.insert_object(object(
                73,
                3_006,
                tsd::MaskArchive {
                    super_: tsd::DrawableArchive {
                        parent: Some(reference(70)),
                        ..Default::default()
                    },
                    ..Default::default()
                },
            ))?;
            Ok(())
        })
        .unwrap();
    package
}

fn test_package_with_live_video_layout() -> IWorkPackage {
    let mut package = test_package_with_two_layouts();
    package
        .update_archive("Index/TemplateSlide-38.iwa", |archive| {
            let slide = archive.object_mut(38).unwrap();
            let mut decoded = kn::SlideArchive::decode(slide.messages[0].data.as_slice())?;
            decoded.name = Some("Title & Live Video".to_owned());
            decoded.owned_drawables = vec![reference(70), reference(39), reference(40)];
            decoded.drawables_z_order = decoded.owned_drawables.clone();
            slide.replace_message(
                0,
                RawMessage {
                    type_: TEST_SLIDE_MESSAGE_TYPE,
                    data: decoded.encode_to_vec(),
                },
            )?;

            let mut movie = object(
                70,
                TEST_MOVIE_MESSAGE_TYPE,
                tsd::MovieArchive {
                    super_: tsd::DrawableArchive {
                        geometry: Some(tsd::GeometryArchive {
                            position: Some(tsp::Point { x: 500.0, y: 100.0 }),
                            size: Some(tsp::Size {
                                width: 400.0,
                                height: 400.0,
                            }),
                            ..Default::default()
                        }),
                        parent: Some(reference(38)),
                        title: Some(reference(71)),
                        caption: Some(reference(72)),
                        accessibility_description: Some("Default Camera".to_owned()),
                        ..Default::default()
                    },
                    poster_image_data: Some(tsp::DataReference { identifier: 2_001 }),
                    style: Some(reference(60)),
                    is_live_video: Some(true),
                    ..Default::default()
                },
            );
            let data = append_repeated_length_delimited_field(
                &movie.messages[0].data,
                TEST_LIVE_VIDEO_INFO_FIELD,
                TEST_LIVE_VIDEO_INFO_PAYLOAD,
            )?;
            movie.replace_message(
                0,
                RawMessage {
                    type_: TEST_MOVIE_MESSAGE_TYPE,
                    data,
                },
            )?;
            movie.archive_info.message_infos[0].object_references = vec![72, 71, 60];
            movie.archive_info.message_infos[0].data_references = vec![2_001];
            archive.insert_object(movie)?;
            archive.insert_object(object(
                71,
                STANDIN_CAPTION_MESSAGE_TYPE,
                tsd::StandinCaptionArchive::default(),
            ))?;
            archive.insert_object(object(
                72,
                STANDIN_CAPTION_MESSAGE_TYPE,
                tsd::StandinCaptionArchive::default(),
            ))?;
            Ok(())
        })
        .unwrap();
    package
}

fn test_package_with_file_movie_layout(flags: u32) -> IWorkPackage {
    let mut package = test_package_with_two_layouts();
    package
        .update_archive("Index/TemplateSlide-38.iwa", |archive| {
            let slide = archive.object_mut(38).unwrap();
            let mut decoded = kn::SlideArchive::decode(slide.messages[0].data.as_slice())?;
            decoded.name = Some(if flags & 1 == 0 {
                "Title & Static Movie".to_owned()
            } else {
                "Title & Movie Placeholder".to_owned()
            });
            decoded.owned_drawables = vec![reference(70)];
            decoded.drawables_z_order = decoded.owned_drawables.clone();
            slide.replace_message(
                0,
                RawMessage {
                    type_: TEST_SLIDE_MESSAGE_TYPE,
                    data: decoded.encode_to_vec(),
                },
            )?;

            let mut movie = object(
                70,
                TEST_MOVIE_MESSAGE_TYPE,
                tsd::MovieArchive {
                    super_: tsd::DrawableArchive {
                        geometry: Some(tsd::GeometryArchive {
                            position: Some(tsp::Point { x: 500.0, y: 100.0 }),
                            size: Some(tsp::Size {
                                width: 800.0,
                                height: 300.0,
                            }),
                            ..Default::default()
                        }),
                        parent: Some(reference(38)),
                        title: Some(reference(71)),
                        caption: Some(reference(72)),
                        accessibility_description: Some("Test movie".to_owned()),
                        ..Default::default()
                    },
                    movie_data: Some(tsp::DataReference { identifier: 2_002 }),
                    poster_image_data: Some(tsp::DataReference { identifier: 2_001 }),
                    style: Some(reference(60)),
                    original_size: Some(tsp::Size {
                        width: 800.0,
                        height: 300.0,
                    }),
                    natural_size: Some(tsp::Size {
                        width: 800.0,
                        height: 300.0,
                    }),
                    flags: Some(flags),
                    ..Default::default()
                },
            );
            movie.archive_info.message_infos[0].object_references = vec![72, 71, 60];
            movie.archive_info.message_infos[0].data_references = vec![2_001, 2_002];
            archive.insert_object(movie)?;
            archive.insert_object(object(
                71,
                STANDIN_CAPTION_MESSAGE_TYPE,
                tsd::StandinCaptionArchive::default(),
            ))?;
            archive.insert_object(object(
                72,
                STANDIN_CAPTION_MESSAGE_TYPE,
                tsd::StandinCaptionArchive::default(),
            ))?;
            Ok(())
        })
        .unwrap();
    package
}

#[allow(deprecated)]
fn test_package_with_slide_movie() -> IWorkPackage {
    let mut package = test_package_with_two_layouts();
    package
        .replace_archive(
            PACKAGE_METADATA_ENTRY,
            &Archive {
                objects: vec![object(
                    100,
                    PACKAGE_METADATA_MESSAGE_TYPE,
                    PackageMetadata {
                        last_object_identifier: 100,
                        components: vec![ComponentInfo {
                            identifier: 4,
                            preferred_locator: "Slide-4".to_owned(),
                            object_uuid_map_entries: [4, 5, 6, 70, 71, 72, 73]
                                .into_iter()
                                .map(|identifier| ObjectUuidMapEntry {
                                    identifier,
                                    uuid: test_uuid(identifier, identifier + 100),
                                })
                                .collect(),
                            ..Default::default()
                        }],
                        ..Default::default()
                    },
                )],
            },
        )
        .unwrap();
    let mut media = IWorkMediaEditor::from_package(package).unwrap();
    let video = media
        .insert_unreferenced("movie.mov", TEST_MOVIE_VIDEO)
        .unwrap();
    let poster = media
        .insert_unreferenced("poster.png", TEST_MOVIE_POSTER)
        .unwrap();
    let mut package = media.into_package();
    package
        .update_archive("Index/Slide-4.iwa", |archive| {
            let slide = archive.object_mut(4).unwrap();
            let mut decoded = kn::SlideArchive::decode(slide.messages[0].data.as_slice())?;
            decoded.owned_drawables.push(reference(70));
            decoded.drawables_z_order.push(reference(70));
            decoded.builds.push(reference(73));
            decoded.build_chunks.push(reference(74));
            slide.replace_message(
                0,
                RawMessage {
                    type_: TEST_SLIDE_MESSAGE_TYPE,
                    data: decoded.encode_to_vec(),
                },
            )?;
            slide.archive_info.message_infos[0]
                .object_references
                .extend([70, 73, 74]);

            let mut movie = object(
                70,
                TEST_MOVIE_MESSAGE_TYPE,
                tsd::MovieArchive {
                    super_: tsd::DrawableArchive {
                        geometry: Some(tsd::GeometryArchive {
                            position: Some(tsp::Point { x: 100.0, y: 200.0 }),
                            size: Some(tsp::Size {
                                width: 800.0,
                                height: 300.0,
                            }),
                            ..Default::default()
                        }),
                        parent: Some(reference(4)),
                        title: Some(reference(71)),
                        caption: Some(reference(72)),
                        accessibility_description: Some("Test movie".to_owned()),
                        ..Default::default()
                    },
                    movie_data: Some(tsp::DataReference {
                        identifier: video.data_identifier.get(),
                    }),
                    poster_image_data: Some(tsp::DataReference {
                        identifier: poster.data_identifier.get(),
                    }),
                    style: Some(reference(50)),
                    original_size: Some(tsp::Size {
                        width: 800.0,
                        height: 300.0,
                    }),
                    natural_size: Some(tsp::Size {
                        width: 800.0,
                        height: 300.0,
                    }),
                    flags: Some(0),
                    ..Default::default()
                },
            );
            movie.archive_info.message_infos[0].object_references = vec![71, 72, 50];
            movie.archive_info.message_infos[0].data_references =
                vec![video.data_identifier.get(), poster.data_identifier.get()];
            archive.insert_object(movie)?;
            archive.insert_object(object(
                71,
                STANDIN_CAPTION_MESSAGE_TYPE,
                tsd::StandinCaptionArchive::default(),
            ))?;
            archive.insert_object(object(
                72,
                STANDIN_CAPTION_MESSAGE_TYPE,
                tsd::StandinCaptionArchive::default(),
            ))?;
            let mut build = object(
                73,
                BUILD_MESSAGE_TYPE,
                kn::BuildArchive {
                    drawable: Some(reference(70)),
                    delivery: "All at Once".to_owned(),
                    duration: Some(0.0),
                    attributes: kn::BuildAttributesArchive {
                        animation_attributes: Some(kn::AnimationAttributesArchive {
                            animation_type: Some("In".to_owned()),
                            effect: Some("apple:movie-start".to_owned()),
                            duration: Some(0.5),
                            delay: Some(0.0),
                            random_number_seed: Some(160_812_089),
                            ..Default::default()
                        }),
                        event_trigger: Some(1),
                        ..Default::default()
                    },
                    chunk_id_seed: Some(1),
                },
            );
            build.archive_info.message_infos[0]
                .object_references
                .push(70);
            archive.insert_object(build)?;
            let build_uuid = test_uuid(7_001, 7_002);
            let mut chunk = object(
                74,
                BUILD_CHUNK_MESSAGE_TYPE,
                kn::BuildChunkArchive {
                    build: Some(reference(73)),
                    delay: Some(0.0),
                    duration: Some(0.5),
                    automatic: Some(false),
                    referent: Some(true),
                    build_chunk_identifier: Some(kn::BuildChunkIdentifierArchive {
                        build_id: Some(build_uuid),
                        build_chunk_id: Some(1),
                    }),
                    build_id: Some(build_uuid),
                    ..Default::default()
                },
            );
            chunk.archive_info.message_infos[0]
                .object_references
                .push(73);
            Ok(archive.insert_object(chunk)?)
        })
        .unwrap();
    crate::data_reference_registry::add_component_data_reference(
        &mut package,
        4,
        video.data_identifier.get(),
        70,
    )
    .unwrap();
    crate::data_reference_registry::add_component_data_reference(
        &mut package,
        4,
        poster.data_identifier.get(),
        70,
    )
    .unwrap();
    package
        .update_archive("Index/Document.iwa", |archive| {
            archive.object_mut(1).unwrap().archive_info.message_infos[0].object_references =
                vec![2];
            archive.object_mut(2).unwrap().archive_info.message_infos[0].object_references =
                vec![3];
            archive.object_mut(3).unwrap().archive_info.message_infos[0].object_references =
                vec![4];
            Ok(())
        })
        .unwrap();
    package
}

fn set_test_placeholder_presentation(
    object: &mut ArchiveObject,
    x: f32,
    style: u64,
    kind: i32,
) -> Result<()> {
    let storage = kn::PlaceholderArchive::decode(object.messages[0].data.as_slice())?
        .super_
        .owned_storage;
    object.replace_message(
        0,
        RawMessage {
            type_: TEST_PLACEHOLDER_MESSAGE_TYPE,
            data: test_placeholder(storage, x, style, kind).encode_to_vec(),
        },
    )?;
    object.archive_info.message_infos[0].object_references = vec![style];
    Ok(())
}

fn test_placeholder(
    storage: Option<Reference>,
    x: f32,
    style: u64,
    kind: i32,
) -> kn::PlaceholderArchive {
    kn::PlaceholderArchive {
        super_: tswp::ShapeInfoArchive {
            super_: tsd::ShapeArchive {
                super_: tsd::DrawableArchive {
                    geometry: Some(tsd::GeometryArchive {
                        position: Some(crate::protobuf::tsp::Point { x, y: 10.0 }),
                        size: Some(crate::protobuf::tsp::Size {
                            width: 500.0,
                            height: 100.0,
                        }),
                        ..Default::default()
                    }),
                    ..Default::default()
                },
                style: Some(reference(style)),
                pathsource: Some(tsd::PathSourceArchive {
                    horizontal_flip: Some(x >= 300.0),
                    ..Default::default()
                }),
                ..Default::default()
            },
            owned_storage: storage,
            ..Default::default()
        },
        kind: Some(kind),
    }
}

fn placeholder_x(placeholder: &kn::PlaceholderArchive) -> Option<f32> {
    placeholder
        .super_
        .super_
        .super_
        .geometry
        .as_ref()?
        .position
        .map(|point| point.x)
}

fn test_uuid(lower: u64, upper: u64) -> Uuid {
    Uuid { lower, upper }
}

fn test_package_with_slide_number() -> IWorkPackage {
    let mut package = test_package();
    package
        .update_archive("Index/Document.iwa", |archive| {
            let node = archive.object_mut(3).unwrap();
            let message = node.messages[0].clone();
            let mut decoded = kn::SlideNodeArchive::decode(message.data.as_slice())?;
            decoded.is_slide_number_visible = Some(false);
            node.replace_message(
                0,
                RawMessage {
                    type_: message.type_,
                    data: decoded.encode_to_vec(),
                },
            )?;
            Ok(())
        })
        .unwrap();
    package
        .update_archive("Index/Slide-4.iwa", |archive| {
            let slide = archive.object_mut(4).unwrap();
            let message = slide.messages[0].clone();
            let mut decoded = kn::SlideArchive::decode(message.data.as_slice())?;
            decoded.slide_number_placeholder = Some(reference(TEST_SLIDE_NUMBER_PLACEHOLDER_ID));
            let data = transform_length_delimited_fields_at_path(
                &decoded.encode_to_vec(),
                &[TEST_SLIDE_NUMBER_PLACEHOLDER_FIELD],
                |reference| {
                    let mut reference = reference.to_vec();
                    append_unknown_varint(&mut reference, 98, 980);
                    Ok(reference)
                },
            )?;
            slide.replace_message(
                0,
                RawMessage {
                    type_: message.type_,
                    data,
                },
            )?;
            archive.insert_object(object(
                TEST_SLIDE_NUMBER_PLACEHOLDER_ID,
                TEST_PLACEHOLDER_MESSAGE_TYPE,
                kn::PlaceholderArchive::default(),
            ))?;
            Ok(())
        })
        .unwrap();
    package
}
