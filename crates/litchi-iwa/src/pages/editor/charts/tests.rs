//! Source-built Pages chart CRUD regression tests.

use std::fs;
use std::path::PathBuf;

use super::*;
use litchi_iwa_common::chart::axis::style::Visibility as AxisVisibility;
use litchi_iwa_common::chart::error_bar::{
    Direction as ErrorBarDirection, FixedValue as ErrorBarFixedValue,
    Percentage as ErrorBarPercentage, Series,
};
use litchi_iwa_common::chart::gaps::{Percentage, Spacing};

use crate::charts::{
    Axis, Bound, Bounds, ChartCornerRadius, ChartDonutInnerRadius, ChartFont, ChartFontSize,
    ChartLegendFill, ChartLegendFont, ChartLegendFontSize, ChartLegendFrame, ChartLegendRect,
    ChartLegendShadow, ChartLegendStroke, ChartPieLabelDistance, ChartPieStartAngle,
    ChartPieWedgeExplosion, ChartPieWedgeIndex, ChartRoundedCorners, ChartSeriesErrorBarAutoFit,
    ChartSeriesStroke, ChartSeriesStrokePattern, ChartSeriesTrendline,
    ChartSeriesTrendlineMovingAveragePeriod, ChartSeriesTrendlinePolynomialOrder,
    ChartSeriesValueLabelAutoFit, ChartSeriesValueLabelLocation, ChartShadow, DecimalPlaces, Index,
    LabelAffixes, LabelVisibility, LeaderLineVisibility, MajorStepCount, MinorStepCount,
    NegativeStyle, NumberFormat, Scale, Steps, TickMarkLocation, Visibility,
};
use crate::package_metadata::{
    add_component_external_reference, add_component_object_uuids, component_identifier_for_entry,
    remove_component_external_reference, remove_component_object_uuids,
};
use crate::pages::PagesDocumentBuilder;
use crate::shapes::{
    Appearance, BlurRadius, Drop, Offset, Pattern, RgbColorSpace, RgbaColor, ShapeFill,
    ShapeImageFillTechnique, Stroke, Width,
};
use litchi_iwa_common::shape::shadow::{Angle, Opacity};

use crate::archive::RawMessage;
use crate::charts::non_style::{
    GENERATED_CHART_NON_STYLE_EXTENSION_FIELD, generated_chart_non_style_extension,
};
use crate::wire::{append_varint_field, parse_wire_fields, patch_length_delimited_field};

const POSITION: DrawablePoint = DrawablePoint { x: 96.0, y: 144.0 };
const SIZE: DrawableSize = DrawableSize {
    width: 360.0,
    height: 240.0,
};

fn sample_data() -> ChartData {
    ChartData::new(
        vec!["North".to_owned(), "South".to_owned()],
        vec!["Q1".to_owned(), "Q2".to_owned(), "Q3".to_owned()],
        vec![
            vec![Some(12.0), Some(18.0), Some(24.0)],
            vec![Some(9.0), Some(21.0), Some(27.0)],
        ],
    )
    .unwrap()
}

fn pie_data() -> ChartData {
    ChartData::new(
        vec!["North".to_owned(), "South".to_owned(), "West".to_owned()],
        vec!["Revenue".to_owned()],
        vec![vec![Some(12.0)], vec![Some(18.0)], vec![Some(24.0)]],
    )
    .unwrap()
}

fn gap_spacing(between_items: f32, between_sets: f32) -> Spacing {
    Spacing::new(
        Percentage::new(between_items).unwrap(),
        Percentage::new(between_sets).unwrap(),
    )
}

fn chart_stroke(pattern: Pattern, width: f32) -> Stroke {
    Stroke::new(
        RgbaColor::new(0.1, 0.3, 0.8, 1.0, RgbColorSpace::Srgb).unwrap(),
        Width::new(width).unwrap(),
        pattern,
    )
}

fn chart_series_stroke(pattern: ChartSeriesStrokePattern, width: f32) -> ChartSeriesStroke {
    ChartSeriesStroke::new(
        RgbaColor::new(0.1, 0.3, 0.8, 1.0, RgbColorSpace::Srgb).unwrap(),
        Width::new(width).unwrap(),
        pattern,
    )
}

fn chart_background_fill() -> ShapeFill {
    ShapeFill::Solid(RgbaColor::new(0.1, 0.3, 0.8, 1.0, RgbColorSpace::Srgb).unwrap())
}

fn chart_shadow() -> ChartShadow {
    ChartShadow::Grouped(Drop::new(
        Appearance::new(
            RgbaColor::new(0.1, 0.3, 0.8, 1.0, RgbColorSpace::Srgb).unwrap(),
            BlurRadius::from_points(15).unwrap(),
            Offset::from_points(8.0).unwrap(),
            Opacity::new(0.6).unwrap(),
        ),
        Angle::from_degrees(60.0).unwrap(),
    ))
}

fn fixture(relative: &str) -> Vec<u8> {
    let root = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../..");
    fs::read(root.join(relative)).unwrap()
}

/// Apply a deliberately malformed generated title payload while retaining the
/// complete Pages body-chart graph. The title transaction must reject it
/// before publication, after the graph checks have succeeded.
fn mutate_pages_chart_non_style(
    editor: &mut PagesEditor,
    drawable_object_id: u64,
    mutate: impl FnOnce(&mut Vec<u8>),
) {
    let graph = body_chart_graph(editor, drawable_object_id).unwrap();
    let archive_name = graph.archive_name.clone();
    let non_style_id = {
        let archive = editor.package().archive(&archive_name).unwrap();
        let object = archive.object(drawable_object_id).unwrap();
        let message = object
            .messages
            .iter()
            .find(|message| message.type_ == CHART_MESSAGE_TYPE)
            .unwrap();
        IWorkChartArchive::decode(&message.data)
            .unwrap()
            .chart
            .unwrap()
            .chart_non_style
            .unwrap()
            .identifier
    };
    let mut package = editor.package().clone();
    package
        .update_archive(&archive_name, |archive| {
            let object = archive.object_mut(non_style_id).unwrap();
            let message_index = object
                .messages
                .iter()
                .position(|message| message.type_ == CHART_NON_STYLE_MESSAGE_TYPE)
                .unwrap();
            let message = object.messages[message_index].clone();
            let mut data = message.data;
            mutate(&mut data);
            object.replace_message(
                message_index,
                RawMessage {
                    type_: CHART_NON_STYLE_MESSAGE_TYPE,
                    data,
                },
            )?;
            Ok(())
        })
        .unwrap();
    *editor = PagesEditor::from_package(package).unwrap();
}

fn pages_chart_non_style_data(editor: &PagesEditor, drawable_object_id: u64) -> Vec<u8> {
    let graph = body_chart_graph(editor, drawable_object_id).unwrap();
    let archive = editor.package().archive(&graph.archive_name).unwrap();
    let chart_object = archive.object(drawable_object_id).unwrap();
    let chart_message = chart_object
        .messages
        .iter()
        .find(|message| message.type_ == CHART_MESSAGE_TYPE)
        .unwrap();
    let non_style_id = IWorkChartArchive::decode(&chart_message.data)
        .unwrap()
        .chart
        .unwrap()
        .chart_non_style
        .unwrap()
        .identifier;
    archive
        .object(non_style_id)
        .unwrap()
        .messages
        .iter()
        .find(|message| message.type_ == CHART_NON_STYLE_MESSAGE_TYPE)
        .unwrap()
        .data
        .clone()
}

fn pages_chart_caption_data(editor: &PagesEditor, drawable_object_id: u64) -> Vec<u8> {
    let graph = body_chart_graph(editor, drawable_object_id).unwrap();
    let archive = editor.package().archive(&graph.archive_name).unwrap();
    archive
        .object(drawable_object_id)
        .unwrap()
        .messages
        .iter()
        .find(|message| message.type_ == CHART_MESSAGE_TYPE)
        .unwrap()
        .data
        .clone()
}

fn mutate_pages_chart_caption(
    editor: &mut PagesEditor,
    drawable_object_id: u64,
    mutate: impl FnOnce(&mut Vec<u8>),
) {
    let graph = body_chart_graph(editor, drawable_object_id).unwrap();
    let archive_name = graph.archive_name.clone();
    let mut package = editor.package().clone();
    package
        .update_archive(&archive_name, |archive| {
            let object = archive.object_mut(drawable_object_id).unwrap();
            let message_index = object
                .messages
                .iter()
                .position(|message| message.type_ == CHART_MESSAGE_TYPE)
                .unwrap();
            let message = object.messages[message_index].clone();
            let mut data = message.data;
            mutate(&mut data);
            object.replace_message(
                message_index,
                RawMessage {
                    type_: CHART_MESSAGE_TYPE,
                    data,
                },
            )?;
            Ok(())
        })
        .unwrap();
    *editor = PagesEditor::from_package(package).unwrap();
}

fn raw_fields(data: &[u8], number: u32) -> Vec<Vec<u8>> {
    parse_wire_fields(data)
        .unwrap()
        .into_iter()
        .filter(|field| field.number() == number)
        .map(|field| data[field.start()..field.end()].to_vec())
        .collect()
}

const UNKNOWN_ARCHIVE_INFO_FIELD: u32 = 4_095;

fn test_encode_varint(mut value: u64) -> Vec<u8> {
    let mut encoded = Vec::new();
    while value >= 0x80 {
        encoded.push((value as u8) | 0x80);
        value >>= 7;
    }
    encoded.push(value as u8);
    encoded
}

fn test_varint_width(value: u64) -> usize {
    test_encode_varint(value).len()
}

fn pages_chart_caption_reference(editor: &PagesEditor, drawable_object_id: u64) -> u64 {
    let data = pages_chart_caption_data(editor, drawable_object_id);
    pages_chart_caption_reference_from_data(&data)
}

fn pages_chart_caption_reference_from_data(data: &[u8]) -> u64 {
    IWorkChartArchive::decode(&data)
        .unwrap()
        .drawable
        .super_
        .unwrap()
        .caption
        .unwrap()
        .identifier
}

fn pages_chart_caption_reference_in_archive(
    editor: &PagesEditor,
    archive_name: &str,
    drawable_object_id: u64,
) -> u64 {
    let archive = editor.package().archive(archive_name).unwrap();
    let data = archive
        .object(drawable_object_id)
        .unwrap()
        .messages
        .iter()
        .find(|message| message.type_ == CHART_MESSAGE_TYPE)
        .unwrap()
        .data
        .as_slice();
    pages_chart_caption_reference_from_data(data)
}

fn pages_chart_message_info(
    editor: &PagesEditor,
    drawable_object_id: u64,
) -> crate::archive::MessageInfo {
    let graph = body_chart_graph(editor, drawable_object_id).unwrap();
    let archive = editor.package().archive(&graph.archive_name).unwrap();
    let object = archive.object(drawable_object_id).unwrap();
    let message_index = object
        .messages
        .iter()
        .position(|message| message.type_ == CHART_MESSAGE_TYPE)
        .unwrap();
    object.archive_info.message_infos[message_index].clone()
}

fn pages_archive_header(
    editor: &PagesEditor,
    archive_name: &str,
    object_identifier: u64,
) -> Vec<u8> {
    let compressed = editor.package().entry(archive_name).unwrap();
    let source = crate::snappy::SnappyStream::decompress(compressed)
        .unwrap()
        .into_bytes();
    let archive = crate::archive::Archive::parse(&source).unwrap();
    let object = archive.object(object_identifier).unwrap();
    let object_start = usize::try_from(object.header_offset).unwrap();
    let prefix_length = test_decode_varint(&source[object_start..]).1;
    let header_start = object_start + prefix_length;
    let header_end = usize::try_from(object.data_offset).unwrap();
    source[header_start..header_end].to_vec()
}

fn test_decode_varint(source: &[u8]) -> (usize, usize) {
    let mut value = 0usize;
    let mut shift = 0usize;
    for (index, byte) in source.iter().copied().enumerate() {
        value |= usize::from(byte & 0x7f) << shift;
        if byte & 0x80 == 0 {
            return (value, index + 1);
        }
        shift += 7;
    }
    panic!("truncated test varint")
}

fn inject_pages_archive_unknown_header(
    editor: &mut PagesEditor,
    archive_name: &str,
    object_identifier: u64,
) -> Vec<u8> {
    let mut package = editor.package().clone();
    let compressed = package.entry(archive_name).unwrap().to_vec();
    let source = crate::snappy::SnappyStream::decompress(&compressed)
        .unwrap()
        .into_bytes();
    let archive = crate::archive::Archive::parse(&source).unwrap();
    let object = archive.object(object_identifier).unwrap();
    let object_start = usize::try_from(object.header_offset).unwrap();
    let (header_length, prefix_length) = test_decode_varint(&source[object_start..]);
    let header_start = object_start + prefix_length;
    let header_end = usize::try_from(object.data_offset).unwrap();
    assert_eq!(header_end - header_start, header_length);
    let payload_end = header_end + usize::try_from(object.data_length).unwrap();

    let mut unknown = Vec::new();
    append_varint_field(&mut unknown, UNKNOWN_ARCHIVE_INFO_FIELD, 42).unwrap();
    let rewritten_header_length = header_length + unknown.len();
    let mut rewritten = Vec::with_capacity(source.len() + unknown.len());
    rewritten.extend_from_slice(&source[..object_start]);
    rewritten.extend_from_slice(&test_encode_varint(rewritten_header_length as u64));
    rewritten.extend_from_slice(&source[header_start..header_end]);
    rewritten.extend_from_slice(&unknown);
    rewritten.extend_from_slice(&source[header_end..payload_end]);
    rewritten.extend_from_slice(&source[payload_end..]);
    package
        .insert_entry(
            archive_name,
            crate::snappy::SnappyStream::compress(&rewritten).unwrap(),
        )
        .unwrap();
    *editor = PagesEditor::from_package(package).unwrap();
    unknown
}

fn reserve_pages_three_byte_caption_identifier(editor: &mut PagesEditor, archive_name: &str) {
    let next = crate::package_metadata::next_object_identifier(editor.package()).unwrap();
    if next < 16_383 {
        let mut package = editor.package().clone();
        package
            .update_archive(archive_name, |archive| {
                archive.insert_object(crate::archive::ArchiveObject::new(
                    16_383,
                    vec![RawMessage {
                        type_: 65_534,
                        data: vec![0],
                    }],
                )?)?;
                Ok(())
            })
            .unwrap();
        *editor = PagesEditor::from_package(package).unwrap();
    }
}

fn mutate_pages_caption_metadata(
    editor: &mut PagesEditor,
    drawable_object_id: u64,
    old_reference_id: u64,
    replacement_id: u64,
    mode: PagesCaptionMetadataMode,
) {
    let graph = body_chart_graph(editor, drawable_object_id).unwrap();
    let mut package = editor.package().clone();
    package
        .update_archive(&graph.archive_name, |archive| {
            let object = archive.object_mut(drawable_object_id).unwrap();
            let message_index = object
                .messages
                .iter()
                .position(|message| message.type_ == CHART_MESSAGE_TYPE)
                .unwrap();
            let info = &mut object.archive_info.message_infos[message_index];
            match mode {
                PagesCaptionMetadataMode::AggregateDuplicate => {
                    info.object_references.push(old_reference_id);
                },
                PagesCaptionMetadataMode::FieldOnly => {
                    info.object_references.retain(|id| *id != old_reference_id);
                    info.field_infos.push(crate::archive::FieldInfo {
                        path: crate::archive::FieldPath::new(vec![1, 11, 1]),
                        object_references: vec![old_reference_id],
                        ..Default::default()
                    });
                },
                PagesCaptionMetadataMode::WrongFieldPath => {
                    info.field_infos.push(crate::archive::FieldInfo {
                        path: crate::archive::FieldPath::new(vec![9, 9]),
                        object_references: vec![old_reference_id],
                        ..Default::default()
                    });
                },
                PagesCaptionMetadataMode::DuplicateField => {
                    info.field_infos.push(crate::archive::FieldInfo {
                        path: crate::archive::FieldPath::new(vec![1, 11, 1]),
                        object_references: vec![old_reference_id, old_reference_id],
                        ..Default::default()
                    });
                },
                PagesCaptionMetadataMode::StaleField => {
                    info.field_infos.push(crate::archive::FieldInfo {
                        path: crate::archive::FieldPath::new(vec![1, 11, 1]),
                        object_references: vec![old_reference_id, 999_999],
                        ..Default::default()
                    });
                },
                PagesCaptionMetadataMode::NewField => {
                    info.field_infos.push(crate::archive::FieldInfo {
                        path: crate::archive::FieldPath::new(vec![1, 11, 1]),
                        object_references: vec![replacement_id],
                        ..Default::default()
                    });
                },
                PagesCaptionMetadataMode::DataReference => {
                    info.field_infos.push(crate::archive::FieldInfo {
                        path: crate::archive::FieldPath::new(vec![1, 11, 1]),
                        data_references: vec![old_reference_id],
                        ..Default::default()
                    });
                },
                PagesCaptionMetadataMode::AuthorizedField => {
                    info.field_infos.push(crate::archive::FieldInfo {
                        path: crate::archive::FieldPath::new(vec![1, 11, 1]),
                        object_references: vec![old_reference_id],
                        ..Default::default()
                    });
                },
            }
            Ok(())
        })
        .unwrap();
    *editor = PagesEditor::from_package(package).unwrap();
}

#[derive(Debug, Clone, Copy)]
enum PagesCaptionMetadataMode {
    AggregateDuplicate,
    FieldOnly,
    WrongFieldPath,
    DuplicateField,
    StaleField,
    NewField,
    DataReference,
    AuthorizedField,
}

fn pages_caption_storage_id(editor: &PagesEditor, drawable_object_id: u64) -> u64 {
    let reference_id = pages_chart_caption_reference(editor, drawable_object_id);
    let archive_name = find_object_archive(editor.package(), reference_id).unwrap();
    let archive = editor.package().archive(&archive_name).unwrap();
    let object = archive.object(reference_id).unwrap();
    let message = object
        .messages
        .iter()
        .find(|message| message.type_ == crate::image_caption::CAPTION_INFO_MESSAGE_TYPE)
        .unwrap();
    litchi_iwa_protos::pages_movie_caption_codec::decode_caption_info(
        &message.data,
        litchi_iwa_protos::pages_movie_caption_codec::DecodeOptions::for_source(&message.data),
    )
    .unwrap()
    .owned_storage_identifier()
    .unwrap()
}

fn make_pages_caption_storage_shared(
    editor: &mut PagesEditor,
    source_drawable_object_id: u64,
    shared_drawable_object_id: u64,
) {
    let source_storage_id = pages_caption_storage_id(editor, source_drawable_object_id);
    let shared_reference_id = pages_chart_caption_reference(editor, shared_drawable_object_id);
    let old_storage_id = pages_caption_storage_id(editor, shared_drawable_object_id);
    let archive_name = find_object_archive(editor.package(), shared_reference_id).unwrap();
    let mut package = editor.package().clone();
    package
        .update_archive(&archive_name, |archive| {
            let object = archive.object_mut(shared_reference_id).unwrap();
            let message_index = object
                .messages
                .iter()
                .position(|message| {
                    message.type_ == crate::image_caption::CAPTION_INFO_MESSAGE_TYPE
                })
                .unwrap();
            let original = object.messages[message_index].data.clone();
            let data = litchi_iwa_protos::pages_movie_caption_codec::rewrite_caption_info(
                &original,
                litchi_iwa_protos::pages_movie_caption_codec::CaptionInfoWrite::new(&[(
                    old_storage_id,
                    source_storage_id,
                )]),
                litchi_iwa_protos::pages_movie_caption_codec::DecodeOptions::new(
                    original.len(),
                    original.len().saturating_mul(4),
                    original.len().saturating_mul(64),
                    8,
                ),
            )
            .unwrap();
            object.replace_message(
                message_index,
                RawMessage {
                    type_: crate::image_caption::CAPTION_INFO_MESSAGE_TYPE,
                    data,
                },
            )?;
            for identifier in
                &mut object.archive_info.message_infos[message_index].object_references
            {
                // The old storage identifier is the only reference in this
                // graph that is not a style or placement object.
                if *identifier == old_storage_id {
                    *identifier = source_storage_id;
                }
            }
            Ok(())
        })
        .unwrap();
    *editor = PagesEditor::from_package(package).unwrap();
}

fn make_pages_caption_info_shared(
    editor: &mut PagesEditor,
    source_drawable_object_id: u64,
    shared_drawable_object_id: u64,
) {
    let source_caption_id = pages_chart_caption_reference(editor, source_drawable_object_id);
    let shared_caption_id = pages_chart_caption_reference(editor, shared_drawable_object_id);
    let graph = body_chart_graph(editor, shared_drawable_object_id).unwrap();
    let limits = editor.package().limits();
    let mut package = editor.package().clone();
    package
        .update_archive(&graph.archive_name, |archive| {
            let object = archive.object_mut(shared_drawable_object_id).unwrap();
            let message_index = object
                .messages
                .iter()
                .position(|message| message.type_ == CHART_MESSAGE_TYPE)
                .unwrap();
            let original = object.messages[message_index].data.clone();
            let data = crate::charts::caption_edge::rewrite_chart_caption_identifier(
                limits,
                &original,
                source_caption_id,
            )?;
            object.replace_message(
                message_index,
                RawMessage {
                    type_: CHART_MESSAGE_TYPE,
                    data,
                },
            )?;
            let info = &mut object.archive_info.message_infos[message_index];
            for identifier in &mut info.object_references {
                if *identifier == shared_caption_id {
                    *identifier = source_caption_id;
                }
            }
            for field in &mut info.field_infos {
                for identifier in &mut field.object_references {
                    if *identifier == shared_caption_id {
                        *identifier = source_caption_id;
                    }
                }
            }
            Ok(())
        })
        .unwrap();
    *editor = PagesEditor::from_package(package).unwrap();
}

fn move_pages_caption_aggregate_owner_to_unrelated_object(
    editor: &mut PagesEditor,
    drawable_object_id: u64,
    unrelated_object_id: u64,
) {
    let caption_id = pages_chart_caption_reference(editor, drawable_object_id);
    let graph = body_chart_graph(editor, unrelated_object_id).unwrap();
    let mut package = editor.package().clone();
    package
        .update_archive(&graph.archive_name, |archive| {
            let selected = archive.object_mut(drawable_object_id).unwrap();
            let selected_message_index = selected
                .messages
                .iter()
                .position(|message| message.type_ == CHART_MESSAGE_TYPE)
                .unwrap();
            selected.archive_info.message_infos[selected_message_index]
                .object_references
                .retain(|identifier| *identifier != caption_id);

            let unrelated = archive.object_mut(unrelated_object_id).unwrap();
            let unrelated_message_index = unrelated
                .messages
                .iter()
                .position(|message| message.type_ == CHART_MESSAGE_TYPE)
                .unwrap();
            unrelated.archive_info.message_infos[unrelated_message_index]
                .object_references
                .push(caption_id);
            Ok(())
        })
        .unwrap();
    *editor = PagesEditor::from_package(package).unwrap();
}

fn assert_only_title_fields_changed(before: &[u8], after: &[u8]) {
    let before_fields = parse_wire_fields(before).unwrap();
    let after_fields = parse_wire_fields(after).unwrap();
    assert_eq!(before_fields.len(), after_fields.len());
    for (before_field, after_field) in before_fields.iter().zip(&after_fields) {
        assert_eq!(before_field.number(), after_field.number());
        if before_field.number() != GENERATED_CHART_NON_STYLE_EXTENSION_FIELD {
            assert_eq!(
                &before[before_field.start()..before_field.end()],
                &after[after_field.start()..after_field.end()]
            );
            continue;
        }
        let before_extension = &before[before_field.payload_start()..before_field.end()];
        let after_extension = &after[after_field.payload_start()..after_field.end()];
        let before_extension_fields = parse_wire_fields(before_extension).unwrap();
        let after_extension_fields = parse_wire_fields(after_extension).unwrap();
        let before_unselected = before_extension_fields
            .iter()
            .filter(|field| field.number() != 21 && field.number() != 23)
            .map(|field| before_extension[field.start()..field.end()].to_vec())
            .collect::<Vec<_>>();
        let after_unselected = after_extension_fields
            .iter()
            .filter(|field| field.number() != 21 && field.number() != 23)
            .map(|field| after_extension[field.start()..field.end()].to_vec())
            .collect::<Vec<_>>();
        assert_eq!(before_unselected, after_unselected);
    }
}

fn mutate_pages_chart_title_standin(editor: &mut PagesEditor, drawable_object_id: u64) {
    let graph = body_chart_graph(editor, drawable_object_id).unwrap();
    let archive_name = graph.archive_name.clone();
    let title_id = {
        let archive = editor.package().archive(&archive_name).unwrap();
        let object = archive.object(drawable_object_id).unwrap();
        let message = object
            .messages
            .iter()
            .find(|message| message.type_ == CHART_MESSAGE_TYPE)
            .unwrap();
        IWorkChartArchive::decode(&message.data)
            .unwrap()
            .drawable
            .super_
            .unwrap()
            .title
            .unwrap()
            .identifier
    };
    let mut package = editor.package().clone();
    package
        .update_archive(&archive_name, |archive| {
            let object = archive.object_mut(title_id).unwrap();
            let message_index = object
                .messages
                .iter()
                .position(|message| message.type_ == STANDIN_MESSAGE_TYPE)
                .unwrap();
            let message = object.messages[message_index].clone();
            object.replace_message(
                message_index,
                RawMessage {
                    type_: CHART_NON_STYLE_MESSAGE_TYPE,
                    data: message.data,
                },
            )?;
            Ok(())
        })
        .unwrap();
    *editor = PagesEditor::from_package(package).unwrap();
}

fn normalize_private_chart_styles_like_pages(
    editor: &PagesEditor,
    drawable_object_id: u64,
) -> PagesEditor {
    let graph = body_chart_graph(editor, drawable_object_id).unwrap();
    let source_group = graph
        .archive_groups
        .iter()
        .find(|group| !group.style_ids.is_empty())
        .unwrap();
    let style_ids = source_group.style_ids.clone();
    let mut package = editor.package().clone();
    let root = pages_document_root_facts(&package).unwrap();
    let theme_id = root.theme.unwrap();
    let theme = chart_theme_context(&package, theme_id).unwrap();
    let stylesheet_archive_name = find_object_archive(&package, theme.stylesheet_id).unwrap();
    let stylesheet_component_id =
        component_identifier_for_entry(&package, &stylesheet_archive_name)
            .unwrap()
            .unwrap();

    let mut moved = Vec::with_capacity(style_ids.len());
    package
        .update_archive(&source_group.archive_name, |archive| {
            for identifier in &style_ids {
                moved.push(archive.remove_object(*identifier).unwrap());
            }
            Ok(())
        })
        .unwrap();
    package
        .update_archive(&stylesheet_archive_name, |archive| {
            for object in moved {
                archive.insert_object(object)?;
            }
            let stylesheet = archive.object_mut(theme.stylesheet_id).unwrap();
            for info in &mut stylesheet.archive_info.message_infos {
                info.object_references
                    .retain(|identifier| !style_ids.contains(identifier));
                for field in &mut info.field_infos {
                    field
                        .object_references
                        .retain(|identifier| !style_ids.contains(identifier));
                }
            }
            Ok(())
        })
        .unwrap();

    remove_component_object_uuids(&mut package, source_group.component_id, &style_ids).unwrap();
    add_component_object_uuids(&mut package, stylesheet_component_id, &style_ids).unwrap();
    for identifier in style_ids {
        remove_component_external_reference(
            &mut package,
            stylesheet_component_id,
            source_group.component_id,
            identifier,
        )
        .unwrap();
        add_component_external_reference(
            &mut package,
            source_group.component_id,
            stylesheet_component_id,
            identifier,
        )
        .unwrap();
    }
    PagesEditor::from_package(package).unwrap()
}

#[test]
fn scratch_document_supports_body_chart_crud() {
    let mut editor = PagesEditor::create_with_text("Quarterly results").unwrap();
    let anchor = "Quarterly results".encode_utf16().count();
    assert!(editor.body_charts().unwrap().is_empty());

    let created = editor
        .add_body_chart(anchor, Kind::Column2d, sample_data(), POSITION, SIZE)
        .unwrap();
    assert_eq!(created.anchor_character_index, anchor as u32);
    assert_eq!(created.kind, Kind::Column2d);
    assert_eq!(created.direction, Direction::Rows);
    assert_eq!(created.data, sample_data());
    assert_eq!(editor.body_text().unwrap(), "Quarterly results\u{fffc}");

    let replacement = ChartData::new(
        vec!["Revenue".to_owned()],
        vec!["2026".to_owned(), "2027".to_owned(), "2028".to_owned()],
        vec![vec![Some(30.0), Some(45.0), None]],
    )
    .unwrap();
    editor
        .set_body_chart_kind(created.drawable_object_id, Kind::Bar2d)
        .unwrap();
    editor
        .set_body_chart_data(created.drawable_object_id, replacement.clone())
        .unwrap();
    editor
        .set_body_chart_direction(created.drawable_object_id, Direction::Columns)
        .unwrap();
    let changed_geometry = chart_geometry(
        "Pages",
        DrawablePoint { x: 72.0, y: 216.0 },
        DrawableSize {
            width: 420.0,
            height: 260.0,
        },
    )
    .unwrap();
    editor
        .set_body_chart_geometry(created.drawable_object_id, changed_geometry)
        .unwrap();

    let reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    let charts = reopened.body_charts().unwrap();
    assert_eq!(charts.len(), 1);
    assert_eq!(charts[0].kind, Kind::Bar2d);
    assert_eq!(charts[0].direction, Direction::Columns);
    assert_eq!(charts[0].data, replacement);
    assert_eq!(charts[0].geometry, changed_geometry);

    let removed = editor
        .remove_body_chart(created.drawable_object_id)
        .unwrap();
    assert_eq!(removed.chart.drawable_object_id, created.drawable_object_id);
    assert_eq!(editor.body_text().unwrap(), "Quarterly results");
    assert!(editor.body_charts().unwrap().is_empty());
    PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
}

#[test]
fn pages_normalized_chart_styles_support_full_lifecycle_crud() {
    let mut editor = PagesEditor::create_with_text("Chart").unwrap();
    let source = editor
        .add_body_chart(5, Kind::Column2d, sample_data(), POSITION, SIZE)
        .unwrap();
    let frame = ChartLegendFrame::Frame(ChartLegendRect::from_points(43.0, 8.0, 0.0, 0.0).unwrap());
    editor
        .set_body_chart_legend_frame(source.drawable_object_id, frame)
        .unwrap();
    let mut editor = normalize_private_chart_styles_like_pages(&editor, source.drawable_object_id);

    let charts = editor.body_charts().unwrap();
    assert_eq!(charts.len(), 1);
    assert_eq!(
        editor
            .body_chart_legend_frame(source.drawable_object_id)
            .unwrap(),
        frame
    );
    let changed_frame =
        ChartLegendFrame::Frame(ChartLegendRect::from_points(56.0, 10.0, 0.0, 0.0).unwrap());
    editor
        .set_body_chart_legend_frame(source.drawable_object_id, changed_frame)
        .unwrap();
    let duplicate_anchor = editor.body_text().unwrap().encode_utf16().count();
    let duplicate = editor
        .duplicate_body_chart(source.drawable_object_id, duplicate_anchor)
        .unwrap();
    editor.remove_body_chart(source.drawable_object_id).unwrap();

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(reopened.body_charts().unwrap().len(), 1);
    assert_eq!(
        reopened
            .body_chart_legend_frame(duplicate.drawable_object_id)
            .unwrap(),
        changed_frame
    );
    reopened
        .remove_body_chart(duplicate.drawable_object_id)
        .unwrap();
    assert!(reopened.body_charts().unwrap().is_empty());
    assert_eq!(reopened.body_text().unwrap(), "Chart");
}

#[test]
fn chart_creation_rejects_invalid_inputs_transactionally() {
    let mut editor = PagesEditor::create_with_text("Body").unwrap();
    let baseline = editor.to_bytes().unwrap();
    assert!(
        editor
            .add_body_chart(4, Kind::Undefined, sample_data(), POSITION, SIZE)
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), baseline);
    assert!(
        editor
            .add_body_chart(5, Kind::Column2d, sample_data(), POSITION, SIZE,)
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), baseline);
    assert!(
        editor
            .add_body_chart(
                4,
                Kind::Column2d,
                sample_data(),
                POSITION,
                DrawableSize {
                    width: 0.0,
                    height: SIZE.height,
                },
            )
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), baseline);
}

#[test]
fn removing_an_earlier_chart_preserves_later_chart_and_anchor() {
    let mut editor = PagesEditor::create_with_text("Body").unwrap();
    let first = editor
        .add_body_chart(4, Kind::Column2d, sample_data(), POSITION, SIZE)
        .unwrap();
    let second = editor
        .add_body_chart(
            5,
            Kind::Line2d,
            sample_data(),
            DrawablePoint { x: 120.0, y: 420.0 },
            SIZE,
        )
        .unwrap();

    editor.remove_body_chart(first.drawable_object_id).unwrap();
    let remaining = editor.body_charts().unwrap();
    assert_eq!(remaining.len(), 1);
    assert_eq!(remaining[0].drawable_object_id, second.drawable_object_id);
    assert_eq!(remaining[0].anchor_character_index, 4);
    assert_eq!(editor.body_text().unwrap(), "Body\u{fffc}");

    editor.remove_body_chart(second.drawable_object_id).unwrap();
    assert!(editor.body_charts().unwrap().is_empty());
    assert_eq!(editor.body_text().unwrap(), "Body");
}

#[test]
fn duplicate_body_chart_clones_the_private_graph_and_inline_data() {
    let mut editor = PagesEditor::create_with_text("Body").unwrap();
    let source = editor
        .add_body_chart(4, Kind::Column2d, sample_data(), POSITION, SIZE)
        .unwrap();
    let source_graph = body_chart_graph(&editor, source.drawable_object_id).unwrap();
    let baseline = editor.to_bytes().unwrap();
    assert!(editor.duplicate_body_chart(u64::MAX, 5).is_err());
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    let duplicate_anchor = editor.body_text().unwrap().encode_utf16().count();
    let duplicate = editor
        .duplicate_body_chart(source.drawable_object_id, duplicate_anchor)
        .unwrap();
    let duplicate_graph = body_chart_graph(&editor, duplicate.drawable_object_id).unwrap();
    let expected_geometry =
        offset_drawable_geometry(source.geometry, BODY_DRAWABLE_DUPLICATE_OFFSET).unwrap();

    assert_ne!(duplicate.drawable_object_id, source.drawable_object_id);
    assert_eq!(duplicate.anchor_character_index, duplicate_anchor as u32);
    assert_eq!(duplicate.kind, source.kind);
    assert_eq!(duplicate.direction, source.direction);
    assert_eq!(duplicate.data, source.data);
    assert_eq!(duplicate.geometry, expected_geometry);
    assert_eq!(
        duplicate_graph.object_ids.len(),
        source_graph.object_ids.len()
    );
    assert!(
        source_graph
            .object_ids
            .iter()
            .all(|identifier| !duplicate_graph.object_ids.contains(identifier))
    );
    assert_eq!(editor.body_text().unwrap(), "Body\u{fffc}\u{fffc}");

    let replacement = ChartData::new(
        vec!["Revenue".to_owned()],
        vec!["2026".to_owned(), "2027".to_owned()],
        vec![vec![Some(30.0), Some(45.0)]],
    )
    .unwrap();
    editor
        .set_body_chart_data(duplicate.drawable_object_id, replacement.clone())
        .unwrap();
    assert_eq!(
        body_chart_graph(&editor, source.drawable_object_id)
            .unwrap()
            .info
            .data,
        source.data
    );
    assert_eq!(
        body_chart_graph(&editor, duplicate.drawable_object_id)
            .unwrap()
            .info
            .data,
        replacement
    );

    editor.remove_body_chart(source.drawable_object_id).unwrap();
    assert_eq!(
        editor
            .body_charts()
            .unwrap()
            .iter()
            .map(|chart| chart.drawable_object_id)
            .collect::<Vec<_>>(),
        vec![duplicate.drawable_object_id]
    );
    editor
        .remove_body_chart(duplicate.drawable_object_id)
        .unwrap();
    assert!(editor.body_charts().unwrap().is_empty());
    assert_eq!(editor.body_text().unwrap(), "Body");
}

#[test]
fn scratch_document_supports_native_chart_caption_crud() {
    let mut editor = PagesEditor::create_with_text("Chart captions").unwrap();
    let source = editor
        .add_body_chart(
            "Chart captions".encode_utf16().count(),
            Kind::Column2d,
            sample_data(),
            POSITION,
            SIZE,
        )
        .unwrap();

    assert_eq!(
        editor
            .body_chart_caption(source.drawable_object_id)
            .unwrap(),
        None
    );
    editor
        .set_body_chart_caption(source.drawable_object_id, "Revenue by region")
        .unwrap();
    assert_eq!(
        editor
            .body_chart_caption(source.drawable_object_id)
            .unwrap(),
        Some("Revenue by region".to_owned())
    );

    let duplicate = editor
        .duplicate_body_chart(
            source.drawable_object_id,
            editor.body_text().unwrap().encode_utf16().count(),
        )
        .unwrap();
    assert_eq!(
        editor
            .body_chart_caption(duplicate.drawable_object_id)
            .unwrap(),
        Some("Revenue by region".to_owned())
    );

    editor
        .set_body_chart_caption(source.drawable_object_id, "Updated source caption")
        .unwrap();
    assert!(
        editor
            .remove_body_chart_caption(source.drawable_object_id)
            .unwrap()
    );
    assert!(
        !editor
            .remove_body_chart_caption(source.drawable_object_id)
            .unwrap()
    );
    assert_eq!(
        editor
            .body_chart_caption(source.drawable_object_id)
            .unwrap(),
        None
    );

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .body_chart_caption(duplicate.drawable_object_id)
            .unwrap(),
        Some("Revenue by region".to_owned())
    );
    reopened
        .remove_body_chart(duplicate.drawable_object_id)
        .unwrap();
    assert!(
        reopened
            .body_charts()
            .unwrap()
            .iter()
            .all(|chart| chart.drawable_object_id != duplicate.drawable_object_id)
    );
}

#[test]
fn pages_chart_caption_rewrite_preserves_unknown_chart_fields() {
    const UNKNOWN_CHART_FIELD: u32 = 4_096;

    let mut editor = PagesEditor::create_with_text("Chart caption wire").unwrap();
    let chart = editor
        .add_body_chart(
            "Chart caption wire".encode_utf16().count(),
            Kind::Column2d,
            sample_data(),
            POSITION,
            SIZE,
        )
        .unwrap();
    mutate_pages_chart_caption(&mut editor, chart.drawable_object_id, |data| {
        append_varint_field(data, UNKNOWN_CHART_FIELD, 42).unwrap();
    });

    let before = pages_chart_caption_data(&editor, chart.drawable_object_id);
    let before_unknown = raw_fields(&before, UNKNOWN_CHART_FIELD);
    editor
        .set_body_chart_caption(chart.drawable_object_id, "Revenue by region")
        .unwrap();
    let after = pages_chart_caption_data(&editor, chart.drawable_object_id);

    assert_eq!(
        raw_fields(&after, UNKNOWN_CHART_FIELD),
        before_unknown,
        "caption edge rewrite must preserve unrelated TSCH fields"
    );
    assert_eq!(
        editor.body_chart_caption(chart.drawable_object_id).unwrap(),
        Some("Revenue by region".to_owned())
    );
}

#[test]
fn pages_chart_caption_retarget_preserves_unknown_archive_header_and_metadata_on_width_growth() {
    let mut editor = PagesEditor::create_with_text("Chart caption header").unwrap();
    let chart = editor
        .add_body_chart(
            "Chart caption header".encode_utf16().count(),
            Kind::Column2d,
            sample_data(),
            POSITION,
            SIZE,
        )
        .unwrap();
    let graph = body_chart_graph(&editor, chart.drawable_object_id).unwrap();
    reserve_pages_three_byte_caption_identifier(&mut editor, &graph.archive_name);

    let old_reference_id = pages_chart_caption_reference(&editor, chart.drawable_object_id);
    let next_identifier =
        crate::package_metadata::next_object_identifier(editor.package()).unwrap();
    let replacement_id = next_identifier + 1;
    assert_ne!(
        test_varint_width(old_reference_id),
        test_varint_width(replacement_id),
        "fixture must exercise a caption-reference varint-width change (old={old_reference_id}, replacement={replacement_id})"
    );
    let before_info = pages_chart_message_info(&editor, chart.drawable_object_id);
    let unknown = inject_pages_archive_unknown_header(
        &mut editor,
        &graph.archive_name,
        chart.drawable_object_id,
    );
    let before_header =
        pages_archive_header(&editor, &graph.archive_name, chart.drawable_object_id);
    let before_unknown = raw_fields(&before_header, UNKNOWN_ARCHIVE_INFO_FIELD);
    assert_eq!(before_unknown, vec![unknown.clone()]);

    editor
        .set_body_chart_caption(chart.drawable_object_id, "Caption after width growth")
        .unwrap();
    let after_data = pages_chart_caption_data(&editor, chart.drawable_object_id);
    let after_info = pages_chart_message_info(&editor, chart.drawable_object_id);
    let mut expected_info = before_info;
    expected_info.length = u32::try_from(after_data.len()).unwrap();
    expected_info
        .object_references
        .retain(|identifier| *identifier != old_reference_id);
    expected_info.object_references.push(replacement_id);
    for field in &mut expected_info.field_infos {
        let had_old_reference = field
            .object_references
            .iter()
            .any(|identifier| *identifier == old_reference_id);
        field
            .object_references
            .retain(|identifier| *identifier != old_reference_id);
        if had_old_reference {
            field.object_references.push(replacement_id);
        }
    }
    assert_eq!(after_info, expected_info);
    let after_header = pages_archive_header(&editor, &graph.archive_name, chart.drawable_object_id);
    assert_eq!(
        raw_fields(&after_header, UNKNOWN_ARCHIVE_INFO_FIELD),
        before_unknown,
        "retarget must retain the complete unknown ArchiveInfo field"
    );
    assert_eq!(
        pages_chart_caption_reference(&editor, chart.drawable_object_id),
        replacement_id
    );
}

#[test]
fn pages_chart_caption_reference_transition_requires_exact_aggregate_and_field_metadata() {
    let modes = [
        PagesCaptionMetadataMode::AggregateDuplicate,
        PagesCaptionMetadataMode::FieldOnly,
        PagesCaptionMetadataMode::WrongFieldPath,
        PagesCaptionMetadataMode::DuplicateField,
        PagesCaptionMetadataMode::StaleField,
        PagesCaptionMetadataMode::NewField,
        PagesCaptionMetadataMode::DataReference,
    ];
    for mode in modes {
        let mut editor = PagesEditor::create_with_text("Chart caption metadata").unwrap();
        let chart = editor
            .add_body_chart(
                "Chart caption metadata".encode_utf16().count(),
                Kind::Column2d,
                sample_data(),
                POSITION,
                SIZE,
            )
            .unwrap();
        let old_reference_id = pages_chart_caption_reference(&editor, chart.drawable_object_id);
        let replacement_id = crate::package_metadata::next_object_identifier(editor.package())
            .unwrap()
            .saturating_add(1);
        mutate_pages_caption_metadata(
            &mut editor,
            chart.drawable_object_id,
            old_reference_id,
            replacement_id,
            mode,
        );
        let before = editor.to_bytes().unwrap();
        assert!(
            editor
                .set_body_chart_caption(chart.drawable_object_id, "must reject")
                .is_err(),
            "malformed caption metadata mode {mode:?} was accepted"
        );
        assert_eq!(editor.to_bytes().unwrap(), before);
    }

    let mut editor = PagesEditor::create_with_text("Chart caption authorized").unwrap();
    let chart = editor
        .add_body_chart(
            "Chart caption authorized".encode_utf16().count(),
            Kind::Column2d,
            sample_data(),
            POSITION,
            SIZE,
        )
        .unwrap();
    let old_reference_id = pages_chart_caption_reference(&editor, chart.drawable_object_id);
    let replacement_id = crate::package_metadata::next_object_identifier(editor.package())
        .unwrap()
        .saturating_add(1);
    mutate_pages_caption_metadata(
        &mut editor,
        chart.drawable_object_id,
        old_reference_id,
        replacement_id,
        PagesCaptionMetadataMode::AuthorizedField,
    );
    editor
        .set_body_chart_caption(chart.drawable_object_id, "authorized")
        .unwrap();
    let info = pages_chart_message_info(&editor, chart.drawable_object_id);
    assert_eq!(
        info.object_references
            .iter()
            .filter(|id| **id == old_reference_id)
            .count(),
        0
    );
    assert_eq!(
        info.object_references
            .iter()
            .filter(|id| **id == replacement_id)
            .count(),
        1
    );
    assert_eq!(info.field_infos.len(), 1);
    assert_eq!(info.field_infos[0].object_references, [replacement_id]);
}

#[test]
fn pages_chart_caption_rejects_two_caption_infos_sharing_one_storage_atomically() {
    let mut editor = PagesEditor::create_with_text("Shared caption storage").unwrap();
    let source = editor
        .add_body_chart(
            "Shared caption storage".encode_utf16().count(),
            Kind::Column2d,
            sample_data(),
            POSITION,
            SIZE,
        )
        .unwrap();
    editor
        .set_body_chart_caption(source.drawable_object_id, "original")
        .unwrap();
    let duplicate = editor
        .duplicate_body_chart(
            source.drawable_object_id,
            editor.body_text().unwrap().encode_utf16().count(),
        )
        .unwrap();
    make_pages_caption_storage_shared(
        &mut editor,
        source.drawable_object_id,
        duplicate.drawable_object_id,
    );
    let before = editor.to_bytes().unwrap();
    assert!(
        editor
            .set_body_chart_caption(source.drawable_object_id, "must reject shared storage")
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn pages_chart_caption_rejects_two_chart_payloads_sharing_one_caption_info_atomically() {
    let mut editor = PagesEditor::create_with_text("Shared caption info").unwrap();
    let source = editor
        .add_body_chart(
            "Shared caption info".encode_utf16().count(),
            Kind::Column2d,
            sample_data(),
            POSITION,
            SIZE,
        )
        .unwrap();
    editor
        .set_body_chart_caption(source.drawable_object_id, "original")
        .unwrap();
    let graph = body_chart_graph(&editor, source.drawable_object_id).unwrap();
    let duplicate = editor
        .duplicate_body_chart(
            source.drawable_object_id,
            editor.body_text().unwrap().encode_utf16().count(),
        )
        .unwrap();
    make_pages_caption_info_shared(
        &mut editor,
        source.drawable_object_id,
        duplicate.drawable_object_id,
    );
    assert_eq!(
        pages_chart_caption_reference_in_archive(
            &editor,
            &graph.archive_name,
            source.drawable_object_id,
        ),
        pages_chart_caption_reference_in_archive(
            &editor,
            &graph.archive_name,
            duplicate.drawable_object_id,
        )
    );

    let before = editor.to_bytes().unwrap();
    assert!(
        editor
            .set_body_chart_caption(source.drawable_object_id, "must reject shared info")
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before);
    assert!(
        editor
            .set_body_chart_caption(duplicate.drawable_object_id, "must reject shared info")
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn pages_chart_caption_rejects_caption_info_owner_on_unrelated_object_atomically() {
    let mut editor = PagesEditor::create_with_text("Unrelated caption owner").unwrap();
    let source = editor
        .add_body_chart(
            "Unrelated caption owner".encode_utf16().count(),
            Kind::Column2d,
            sample_data(),
            POSITION,
            SIZE,
        )
        .unwrap();
    editor
        .set_body_chart_caption(source.drawable_object_id, "original")
        .unwrap();
    let duplicate = editor
        .duplicate_body_chart(
            source.drawable_object_id,
            editor.body_text().unwrap().encode_utf16().count(),
        )
        .unwrap();
    move_pages_caption_aggregate_owner_to_unrelated_object(
        &mut editor,
        source.drawable_object_id,
        duplicate.drawable_object_id,
    );

    let before = editor.to_bytes().unwrap();
    assert!(
        editor
            .set_body_chart_caption(source.drawable_object_id, "must reject stale owner")
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn scratch_document_supports_native_chart_title_crud() {
    let mut editor = PagesEditor::create_with_text("Chart titles").unwrap();
    let source = editor
        .add_body_chart(
            "Chart titles".encode_utf16().count(),
            Kind::Column2d,
            sample_data(),
            POSITION,
            SIZE,
        )
        .unwrap();

    assert_eq!(
        editor.body_chart_title(source.drawable_object_id).unwrap(),
        None
    );
    editor
        .set_body_chart_title(source.drawable_object_id, "Revenue by region")
        .unwrap();
    assert_eq!(
        editor.body_chart_title(source.drawable_object_id).unwrap(),
        Some("Revenue by region".to_owned())
    );

    let duplicate = editor
        .duplicate_body_chart(
            source.drawable_object_id,
            editor.body_text().unwrap().encode_utf16().count(),
        )
        .unwrap();
    assert_eq!(
        editor
            .body_chart_title(duplicate.drawable_object_id)
            .unwrap(),
        Some("Revenue by region".to_owned())
    );

    editor
        .set_body_chart_title(source.drawable_object_id, "Updated source title")
        .unwrap();
    assert_eq!(
        editor.body_chart_title(source.drawable_object_id).unwrap(),
        Some("Updated source title".to_owned())
    );
    assert_eq!(
        editor
            .body_chart_title(duplicate.drawable_object_id)
            .unwrap(),
        Some("Revenue by region".to_owned())
    );
    assert!(
        editor
            .remove_body_chart_title(source.drawable_object_id)
            .unwrap()
    );
    assert!(
        !editor
            .remove_body_chart_title(source.drawable_object_id)
            .unwrap()
    );
    assert_eq!(
        editor.body_chart_title(source.drawable_object_id).unwrap(),
        None
    );

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .body_chart_title(duplicate.drawable_object_id)
            .unwrap(),
        Some("Revenue by region".to_owned())
    );
    reopened
        .remove_body_chart(duplicate.drawable_object_id)
        .unwrap();
    assert!(
        reopened
            .body_charts()
            .unwrap()
            .iter()
            .all(|chart| chart.drawable_object_id != duplicate.drawable_object_id)
    );
}

#[test]
fn pages_chart_title_rewrite_is_field_local_and_lossless() {
    const UNKNOWN_OUTER_FIELD: u32 = 4_096;
    const UNKNOWN_GENERATED_FIELD: u32 = 4_097;

    let mut editor = PagesEditor::create_with_text("Chart titles").unwrap();
    let chart = editor
        .add_body_chart(
            "Chart titles".encode_utf16().count(),
            Kind::Column2d,
            sample_data(),
            POSITION,
            SIZE,
        )
        .unwrap();
    mutate_pages_chart_non_style(&mut editor, chart.drawable_object_id, |data| {
        let extension = generated_chart_non_style_extension(data).unwrap().unwrap();
        let mut replacement = extension.to_vec();
        append_varint_field(&mut replacement, UNKNOWN_GENERATED_FIELD, 42).unwrap();
        let mut patched = patch_length_delimited_field(
            data,
            GENERATED_CHART_NON_STYLE_EXTENSION_FIELD,
            true,
            Some(&replacement),
        )
        .unwrap();
        append_varint_field(&mut patched, UNKNOWN_OUTER_FIELD, 84).unwrap();
        *data = patched;
    });

    let before = pages_chart_non_style_data(&editor, chart.drawable_object_id);
    let before_extension = generated_chart_non_style_extension(&before)
        .unwrap()
        .unwrap()
        .to_vec();
    let before_outer_unknown = raw_fields(&before, UNKNOWN_OUTER_FIELD);
    let before_generated_unknown = raw_fields(&before_extension, UNKNOWN_GENERATED_FIELD);

    editor
        .set_body_chart_title(chart.drawable_object_id, "Revenue by region")
        .unwrap();
    assert_eq!(
        editor.body_chart_title(chart.drawable_object_id).unwrap(),
        Some("Revenue by region".to_owned())
    );
    let after = pages_chart_non_style_data(&editor, chart.drawable_object_id);
    let after_extension = generated_chart_non_style_extension(&after)
        .unwrap()
        .unwrap()
        .to_vec();
    assert_only_title_fields_changed(&before, &after);
    assert_eq!(
        raw_fields(&after, UNKNOWN_OUTER_FIELD),
        before_outer_unknown
    );
    assert_eq!(
        raw_fields(&after_extension, UNKNOWN_GENERATED_FIELD),
        before_generated_unknown
    );

    let no_op = editor.to_bytes().unwrap();
    editor
        .set_body_chart_title(chart.drawable_object_id, "Revenue by region")
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), no_op);
}

#[test]
fn pages_chart_title_rejects_duplicate_selected_field_without_publication() {
    let mut editor = PagesEditor::create_with_text("Chart titles").unwrap();
    let chart = editor
        .add_body_chart(
            "Chart titles".encode_utf16().count(),
            Kind::Column2d,
            sample_data(),
            POSITION,
            SIZE,
        )
        .unwrap();
    mutate_pages_chart_non_style(&mut editor, chart.drawable_object_id, |data| {
        let extension = generated_chart_non_style_extension(data).unwrap().unwrap();
        let mut replacement = extension.to_vec();
        append_varint_field(&mut replacement, 21, 1).unwrap();
        *data = patch_length_delimited_field(
            data,
            GENERATED_CHART_NON_STYLE_EXTENSION_FIELD,
            true,
            Some(&replacement),
        )
        .unwrap();
    });

    let before = editor.to_bytes().unwrap();
    assert!(editor.body_chart_title(chart.drawable_object_id).is_err());
    assert_eq!(editor.to_bytes().unwrap(), before);
    assert!(
        editor
            .set_body_chart_title(chart.drawable_object_id, "rejected")
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn pages_chart_title_rejects_non_standin_title_graph_without_publication() {
    let mut editor = PagesEditor::create_with_text("Chart titles").unwrap();
    let chart = editor
        .add_body_chart(
            "Chart titles".encode_utf16().count(),
            Kind::Column2d,
            sample_data(),
            POSITION,
            SIZE,
        )
        .unwrap();
    mutate_pages_chart_title_standin(&mut editor, chart.drawable_object_id);

    let before = editor.to_bytes().unwrap();
    assert!(editor.body_chart_title(chart.drawable_object_id).is_err());
    assert_eq!(editor.to_bytes().unwrap(), before);
    assert!(
        editor
            .set_body_chart_title(chart.drawable_object_id, "rejected")
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn scratch_document_supports_native_chart_axis_title_crud() {
    let mut editor = PagesEditor::create_with_text("Chart axis titles").unwrap();
    let source = editor
        .add_body_chart(
            "Chart axis titles".encode_utf16().count(),
            Kind::Column2d,
            sample_data(),
            POSITION,
            SIZE,
        )
        .unwrap();

    for axis in [Axis::Category, Axis::Value] {
        assert_eq!(
            editor
                .body_chart_axis_title(source.drawable_object_id, axis)
                .unwrap(),
            None
        );
    }
    editor
        .set_body_chart_axis_title(source.drawable_object_id, Axis::Category, "Month")
        .unwrap();
    editor
        .set_body_chart_axis_title(source.drawable_object_id, Axis::Value, "Revenue")
        .unwrap();

    let duplicate = editor
        .duplicate_body_chart(
            source.drawable_object_id,
            editor.body_text().unwrap().encode_utf16().count(),
        )
        .unwrap();
    for (axis, title) in [(Axis::Category, "Month"), (Axis::Value, "Revenue")] {
        assert_eq!(
            editor
                .body_chart_axis_title(source.drawable_object_id, axis)
                .unwrap()
                .as_deref(),
            Some(title)
        );
        assert_eq!(
            editor
                .body_chart_axis_title(duplicate.drawable_object_id, axis)
                .unwrap()
                .as_deref(),
            Some(title)
        );
    }

    editor
        .set_body_chart_axis_title(source.drawable_object_id, Axis::Category, "Updated month")
        .unwrap();
    editor
        .set_body_chart_axis_title(source.drawable_object_id, Axis::Value, "Updated revenue")
        .unwrap();
    assert_eq!(
        editor
            .body_chart_axis_title(source.drawable_object_id, Axis::Category)
            .unwrap()
            .as_deref(),
        Some("Updated month")
    );
    assert_eq!(
        editor
            .body_chart_axis_title(source.drawable_object_id, Axis::Value)
            .unwrap()
            .as_deref(),
        Some("Updated revenue")
    );
    assert_eq!(
        editor
            .body_chart_axis_title(duplicate.drawable_object_id, Axis::Category)
            .unwrap()
            .as_deref(),
        Some("Month")
    );
    assert_eq!(
        editor
            .body_chart_axis_title(duplicate.drawable_object_id, Axis::Value)
            .unwrap()
            .as_deref(),
        Some("Revenue")
    );

    for axis in [Axis::Category, Axis::Value] {
        assert!(
            editor
                .remove_body_chart_axis_title(source.drawable_object_id, axis)
                .unwrap()
        );
        assert!(
            !editor
                .remove_body_chart_axis_title(source.drawable_object_id, axis)
                .unwrap()
        );
    }

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .body_chart_axis_title(duplicate.drawable_object_id, Axis::Category)
            .unwrap()
            .as_deref(),
        Some("Month")
    );
    assert_eq!(
        reopened
            .body_chart_axis_title(duplicate.drawable_object_id, Axis::Value)
            .unwrap()
            .as_deref(),
        Some("Revenue")
    );
    reopened
        .remove_body_chart(duplicate.drawable_object_id)
        .unwrap();
    assert!(
        reopened
            .body_charts()
            .unwrap()
            .iter()
            .all(|chart| chart.drawable_object_id != duplicate.drawable_object_id)
    );
}

#[test]
fn scratch_document_supports_native_chart_value_axis_bounds_crud() {
    let mut editor = PagesEditor::create_with_text("Chart value-axis bounds").unwrap();
    let source = editor
        .add_body_chart(
            "Chart value-axis bounds".encode_utf16().count(),
            Kind::Column2d,
            sample_data(),
            POSITION,
            SIZE,
        )
        .unwrap();
    let automatic = Bounds::automatic();
    let fixed = Bounds::fixed(Bound::new(-10.0).unwrap(), Bound::new(40.0).unwrap()).unwrap();
    let minimum_only = Bounds::new(Some(Bound::new(-5.0).unwrap()), None).unwrap();

    assert_eq!(
        editor
            .body_chart_value_axis_bounds(source.drawable_object_id)
            .unwrap(),
        automatic
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_body_chart_value_axis_bounds(source.drawable_object_id, automatic)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    editor
        .set_body_chart_value_axis_bounds(source.drawable_object_id, fixed)
        .unwrap();
    let duplicate = editor
        .duplicate_body_chart(
            source.drawable_object_id,
            editor.body_text().unwrap().encode_utf16().count(),
        )
        .unwrap();
    assert_eq!(
        editor
            .body_chart_value_axis_bounds(duplicate.drawable_object_id)
            .unwrap(),
        fixed
    );

    editor
        .set_body_chart_value_axis_bounds(source.drawable_object_id, minimum_only)
        .unwrap();
    assert_eq!(
        editor
            .body_chart_value_axis_bounds(source.drawable_object_id)
            .unwrap(),
        minimum_only
    );

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .body_chart_value_axis_bounds(source.drawable_object_id)
            .unwrap(),
        minimum_only
    );
    assert_eq!(
        reopened
            .body_chart_value_axis_bounds(duplicate.drawable_object_id)
            .unwrap(),
        fixed
    );
    reopened
        .set_body_chart_value_axis_bounds(source.drawable_object_id, automatic)
        .unwrap();
    assert_eq!(
        reopened
            .body_chart_value_axis_bounds(source.drawable_object_id)
            .unwrap(),
        automatic
    );
    reopened
        .remove_body_chart(duplicate.drawable_object_id)
        .unwrap();
}

#[test]
fn scratch_document_supports_native_chart_border_crud() {
    let mut editor = PagesEditor::create_with_text("Chart borders").unwrap();
    let source = editor
        .add_body_chart(
            "Chart borders".encode_utf16().count(),
            Kind::Column2d,
            sample_data(),
            POSITION,
            SIZE,
        )
        .unwrap();

    assert!(
        !editor
            .body_chart_border_visible(source.drawable_object_id)
            .unwrap()
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_body_chart_border_visible(source.drawable_object_id, false)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    editor
        .set_body_chart_border_visible(source.drawable_object_id, true)
        .unwrap();
    assert!(
        editor
            .body_chart_border_visible(source.drawable_object_id)
            .unwrap()
    );

    let duplicate = editor
        .duplicate_body_chart(
            source.drawable_object_id,
            editor.body_text().unwrap().encode_utf16().count(),
        )
        .unwrap();
    assert!(
        editor
            .body_chart_border_visible(duplicate.drawable_object_id)
            .unwrap()
    );
    editor
        .set_body_chart_border_visible(source.drawable_object_id, false)
        .unwrap();

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert!(
        !reopened
            .body_chart_border_visible(source.drawable_object_id)
            .unwrap()
    );
    assert!(
        reopened
            .body_chart_border_visible(duplicate.drawable_object_id)
            .unwrap()
    );
    reopened
        .remove_body_chart(source.drawable_object_id)
        .unwrap();
    reopened
        .remove_body_chart(duplicate.drawable_object_id)
        .unwrap();
    assert!(reopened.body_charts().unwrap().is_empty());
}

#[test]
fn scratch_document_supports_native_chart_rounded_corner_crud() {
    let mut editor = PagesEditor::create_with_text("Rounded chart corners").unwrap();
    let source = editor
        .add_body_chart(
            "Rounded chart corners".encode_utf16().count(),
            Kind::Column2d,
            sample_data(),
            POSITION,
            SIZE,
        )
        .unwrap();
    let rounded = ChartRoundedCorners::new(ChartCornerRadius::new(20.0).unwrap(), true);
    let changed = ChartRoundedCorners::new(ChartCornerRadius::new(35.0).unwrap(), false);

    assert_eq!(
        editor
            .body_chart_rounded_corners(source.drawable_object_id)
            .unwrap(),
        ChartRoundedCorners::NONE
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_body_chart_rounded_corners(source.drawable_object_id, ChartRoundedCorners::NONE)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    editor
        .set_body_chart_rounded_corners(source.drawable_object_id, rounded)
        .unwrap();
    assert_eq!(
        editor
            .body_chart_rounded_corners(source.drawable_object_id)
            .unwrap(),
        rounded
    );

    let duplicate = editor
        .duplicate_body_chart(
            source.drawable_object_id,
            editor.body_text().unwrap().encode_utf16().count(),
        )
        .unwrap();
    assert_eq!(
        editor
            .body_chart_rounded_corners(duplicate.drawable_object_id)
            .unwrap(),
        rounded
    );
    editor
        .set_body_chart_rounded_corners(source.drawable_object_id, changed)
        .unwrap();

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .body_chart_rounded_corners(source.drawable_object_id)
            .unwrap(),
        changed
    );
    assert_eq!(
        reopened
            .body_chart_rounded_corners(duplicate.drawable_object_id)
            .unwrap(),
        rounded
    );
    reopened
        .remove_body_chart(source.drawable_object_id)
        .unwrap();
    reopened
        .remove_body_chart(duplicate.drawable_object_id)
        .unwrap();
    assert!(reopened.body_charts().unwrap().is_empty());
}

#[test]
fn scratch_document_supports_native_chart_gap_crud() {
    let mut editor = PagesEditor::create_with_text("Chart gaps").unwrap();
    let source = editor
        .add_body_chart(
            "Chart gaps".encode_utf16().count(),
            Kind::Column2d,
            sample_data(),
            POSITION,
            SIZE,
        )
        .unwrap();
    let customized = gap_spacing(25.0, 70.0);
    let changed = gap_spacing(30.0, 60.0);

    assert_eq!(
        editor
            .body_chart_gap_spacing(source.drawable_object_id)
            .unwrap(),
        Spacing::DEFAULT
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_body_chart_gap_spacing(source.drawable_object_id, Spacing::DEFAULT)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    editor
        .set_body_chart_gap_spacing(source.drawable_object_id, customized)
        .unwrap();
    let duplicate = editor
        .duplicate_body_chart(
            source.drawable_object_id,
            editor.body_text().unwrap().encode_utf16().count(),
        )
        .unwrap();
    assert_eq!(
        editor
            .body_chart_gap_spacing(duplicate.drawable_object_id)
            .unwrap(),
        customized
    );
    editor
        .set_body_chart_gap_spacing(source.drawable_object_id, changed)
        .unwrap();

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .body_chart_gap_spacing(source.drawable_object_id)
            .unwrap(),
        changed
    );
    assert_eq!(
        reopened
            .body_chart_gap_spacing(duplicate.drawable_object_id)
            .unwrap(),
        customized
    );
    reopened
        .remove_body_chart(source.drawable_object_id)
        .unwrap();
    reopened
        .remove_body_chart(duplicate.drawable_object_id)
        .unwrap();
    assert!(reopened.body_charts().unwrap().is_empty());
}

#[test]
fn scratch_document_supports_native_chart_value_axis_steps_crud() {
    let mut editor = PagesEditor::create_with_text("Chart value-axis steps").unwrap();
    let source = editor
        .add_body_chart(
            "Chart value-axis steps".encode_utf16().count(),
            Kind::Column2d,
            sample_data(),
            POSITION,
            SIZE,
        )
        .unwrap();
    let defaults = Steps::fixed(
        MajorStepCount::new(5).unwrap(),
        MinorStepCount::new(1).unwrap(),
    );
    let fixed = Steps::fixed(
        MajorStepCount::new(6).unwrap(),
        MinorStepCount::new(2).unwrap(),
    );
    let major_only = Steps::new(Some(MajorStepCount::new(4).unwrap()), None);

    assert_eq!(
        editor
            .body_chart_value_axis_steps(source.drawable_object_id)
            .unwrap(),
        defaults
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_body_chart_value_axis_steps(source.drawable_object_id, defaults)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    editor
        .set_body_chart_value_axis_steps(source.drawable_object_id, fixed)
        .unwrap();
    let duplicate = editor
        .duplicate_body_chart(
            source.drawable_object_id,
            editor.body_text().unwrap().encode_utf16().count(),
        )
        .unwrap();
    assert_eq!(
        editor
            .body_chart_value_axis_steps(duplicate.drawable_object_id)
            .unwrap(),
        fixed
    );

    editor
        .set_body_chart_value_axis_steps(source.drawable_object_id, major_only)
        .unwrap();
    assert_eq!(
        editor
            .body_chart_value_axis_steps(source.drawable_object_id)
            .unwrap(),
        major_only
    );

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .body_chart_value_axis_steps(source.drawable_object_id)
            .unwrap(),
        major_only
    );
    assert_eq!(
        reopened
            .body_chart_value_axis_steps(duplicate.drawable_object_id)
            .unwrap(),
        fixed
    );
    reopened
        .set_body_chart_value_axis_steps(source.drawable_object_id, Steps::automatic())
        .unwrap();
    assert_eq!(
        reopened
            .body_chart_value_axis_steps(source.drawable_object_id)
            .unwrap(),
        Steps::automatic()
    );
    reopened
        .remove_body_chart(duplicate.drawable_object_id)
        .unwrap();
}

#[test]
fn scratch_document_supports_native_chart_value_axis_minimum_label_visibility_crud() {
    let mut editor = PagesEditor::create_with_text("Chart minimum label").unwrap();
    let source = editor
        .add_body_chart(
            "Chart minimum label".encode_utf16().count(),
            Kind::Column2d,
            sample_data(),
            POSITION,
            SIZE,
        )
        .unwrap();

    assert!(
        editor
            .body_chart_value_axis_minimum_label_visible(source.drawable_object_id)
            .unwrap()
            .is_visible()
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_body_chart_value_axis_minimum_label_visible(
            source.drawable_object_id,
            AxisVisibility::Visible,
        )
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    editor
        .set_body_chart_value_axis_minimum_label_visible(
            source.drawable_object_id,
            AxisVisibility::Hidden,
        )
        .unwrap();
    assert!(
        !editor
            .body_chart_value_axis_minimum_label_visible(source.drawable_object_id)
            .unwrap()
            .is_visible()
    );
    let duplicate = editor
        .duplicate_body_chart(
            source.drawable_object_id,
            editor.body_text().unwrap().encode_utf16().count(),
        )
        .unwrap();
    assert!(
        !editor
            .body_chart_value_axis_minimum_label_visible(duplicate.drawable_object_id)
            .unwrap()
            .is_visible()
    );

    editor
        .set_body_chart_value_axis_minimum_label_visible(
            source.drawable_object_id,
            AxisVisibility::Visible,
        )
        .unwrap();
    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert!(
        reopened
            .body_chart_value_axis_minimum_label_visible(source.drawable_object_id)
            .unwrap()
            .is_visible()
    );
    assert!(
        !reopened
            .body_chart_value_axis_minimum_label_visible(duplicate.drawable_object_id)
            .unwrap()
            .is_visible()
    );
    reopened
        .set_body_chart_value_axis_minimum_label_visible(
            source.drawable_object_id,
            AxisVisibility::Hidden,
        )
        .unwrap();
    assert!(
        !reopened
            .body_chart_value_axis_minimum_label_visible(source.drawable_object_id)
            .unwrap()
            .is_visible()
    );
    reopened
        .remove_body_chart(duplicate.drawable_object_id)
        .unwrap();
}

#[test]
fn scratch_document_supports_native_chart_category_axis_series_names_visibility_crud() {
    let mut editor = PagesEditor::create_with_text("Chart series names").unwrap();
    let source = editor
        .add_body_chart(
            "Chart series names".encode_utf16().count(),
            Kind::Column2d,
            sample_data(),
            POSITION,
            SIZE,
        )
        .unwrap();

    assert!(
        !editor
            .body_chart_category_axis_series_names_visible(source.drawable_object_id)
            .unwrap()
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_body_chart_category_axis_series_names_visible(source.drawable_object_id, false)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    editor
        .set_body_chart_category_axis_series_names_visible(source.drawable_object_id, true)
        .unwrap();
    assert!(
        editor
            .body_chart_category_axis_series_names_visible(source.drawable_object_id)
            .unwrap()
    );
    let duplicate = editor
        .duplicate_body_chart(
            source.drawable_object_id,
            editor.body_text().unwrap().encode_utf16().count(),
        )
        .unwrap();
    assert!(
        editor
            .body_chart_category_axis_series_names_visible(duplicate.drawable_object_id)
            .unwrap()
    );

    editor
        .set_body_chart_category_axis_series_names_visible(source.drawable_object_id, false)
        .unwrap();
    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert!(
        !reopened
            .body_chart_category_axis_series_names_visible(source.drawable_object_id)
            .unwrap()
    );
    assert!(
        reopened
            .body_chart_category_axis_series_names_visible(duplicate.drawable_object_id)
            .unwrap()
    );
    reopened
        .set_body_chart_category_axis_series_names_visible(source.drawable_object_id, true)
        .unwrap();
    assert!(
        reopened
            .body_chart_category_axis_series_names_visible(source.drawable_object_id)
            .unwrap()
    );
    reopened
        .remove_body_chart(duplicate.drawable_object_id)
        .unwrap();
}

#[test]
fn scratch_document_supports_native_chart_axis_label_visibility_crud() {
    let mut editor = PagesEditor::create_with_text("Chart axis labels").unwrap();
    let source = editor
        .add_body_chart(
            "Chart axis labels".encode_utf16().count(),
            Kind::Column2d,
            sample_data(),
            POSITION,
            SIZE,
        )
        .unwrap();

    for axis in [Axis::Category, Axis::Value] {
        assert!(
            editor
                .body_chart_axis_labels_visible(source.drawable_object_id, axis)
                .unwrap()
        );
    }
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_body_chart_axis_labels_visible(source.drawable_object_id, Axis::Category, true)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    for axis in [Axis::Category, Axis::Value] {
        editor
            .set_body_chart_axis_labels_visible(source.drawable_object_id, axis, false)
            .unwrap();
        assert!(
            !editor
                .body_chart_axis_labels_visible(source.drawable_object_id, axis)
                .unwrap()
        );
    }

    let duplicate = editor
        .duplicate_body_chart(
            source.drawable_object_id,
            editor.body_text().unwrap().encode_utf16().count(),
        )
        .unwrap();
    for axis in [Axis::Category, Axis::Value] {
        assert!(
            !editor
                .body_chart_axis_labels_visible(duplicate.drawable_object_id, axis)
                .unwrap()
        );
        editor
            .set_body_chart_axis_labels_visible(source.drawable_object_id, axis, true)
            .unwrap();
    }

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    for axis in [Axis::Category, Axis::Value] {
        assert!(
            reopened
                .body_chart_axis_labels_visible(source.drawable_object_id, axis)
                .unwrap()
        );
        assert!(
            !reopened
                .body_chart_axis_labels_visible(duplicate.drawable_object_id, axis)
                .unwrap()
        );
        reopened
            .set_body_chart_axis_labels_visible(source.drawable_object_id, axis, false)
            .unwrap();
        assert!(
            !reopened
                .body_chart_axis_labels_visible(source.drawable_object_id, axis)
                .unwrap()
        );
    }
    reopened
        .remove_body_chart(duplicate.drawable_object_id)
        .unwrap();
}

#[test]
fn scratch_document_supports_native_chart_axis_line_visibility_crud() {
    let mut editor = PagesEditor::create_with_text("Chart axis lines").unwrap();
    let source = editor
        .add_body_chart(
            "Chart axis lines".encode_utf16().count(),
            Kind::Column2d,
            sample_data(),
            POSITION,
            SIZE,
        )
        .unwrap();

    for axis in [Axis::Category, Axis::Value] {
        assert!(
            editor
                .body_chart_axis_line_visible(source.drawable_object_id, axis)
                .unwrap()
                .is_visible()
        );
        editor
            .set_body_chart_axis_line_visible(
                source.drawable_object_id,
                axis,
                AxisVisibility::Hidden,
            )
            .unwrap();
        assert!(
            !editor
                .body_chart_axis_line_visible(source.drawable_object_id, axis)
                .unwrap()
                .is_visible()
        );
    }

    let duplicate = editor
        .duplicate_body_chart(
            source.drawable_object_id,
            editor.body_text().unwrap().encode_utf16().count(),
        )
        .unwrap();
    for axis in [Axis::Category, Axis::Value] {
        assert!(
            !editor
                .body_chart_axis_line_visible(duplicate.drawable_object_id, axis)
                .unwrap()
                .is_visible()
        );
        editor
            .set_body_chart_axis_line_visible(
                source.drawable_object_id,
                axis,
                AxisVisibility::Visible,
            )
            .unwrap();
    }

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    for axis in [Axis::Category, Axis::Value] {
        assert!(
            reopened
                .body_chart_axis_line_visible(source.drawable_object_id, axis)
                .unwrap()
                .is_visible()
        );
        assert!(
            !reopened
                .body_chart_axis_line_visible(duplicate.drawable_object_id, axis)
                .unwrap()
                .is_visible()
        );
    }
    reopened
        .remove_body_chart(duplicate.drawable_object_id)
        .unwrap();
}

#[test]
fn scratch_document_supports_native_chart_axis_major_gridline_visibility_crud() {
    let mut editor = PagesEditor::create_with_text("Chart major gridlines").unwrap();
    let source = editor
        .add_body_chart(
            "Chart major gridlines".encode_utf16().count(),
            Kind::Column2d,
            sample_data(),
            POSITION,
            SIZE,
        )
        .unwrap();

    assert!(
        !editor
            .body_chart_axis_major_gridlines_visible(source.drawable_object_id, Axis::Category)
            .unwrap()
            .is_visible()
    );
    assert!(
        editor
            .body_chart_axis_major_gridlines_visible(source.drawable_object_id, Axis::Value)
            .unwrap()
            .is_visible()
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_body_chart_axis_major_gridlines_visible(
            source.drawable_object_id,
            Axis::Category,
            AxisVisibility::Hidden,
        )
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    editor
        .set_body_chart_axis_major_gridlines_visible(
            source.drawable_object_id,
            Axis::Category,
            AxisVisibility::Visible,
        )
        .unwrap();
    editor
        .set_body_chart_axis_major_gridlines_visible(
            source.drawable_object_id,
            Axis::Value,
            AxisVisibility::Hidden,
        )
        .unwrap();

    let duplicate = editor
        .duplicate_body_chart(
            source.drawable_object_id,
            editor.body_text().unwrap().encode_utf16().count(),
        )
        .unwrap();
    for axis in [Axis::Category, Axis::Value] {
        assert_eq!(
            editor
                .body_chart_axis_major_gridlines_visible(duplicate.drawable_object_id, axis)
                .unwrap()
                .is_visible(),
            axis == Axis::Category
        );
    }

    editor
        .set_body_chart_axis_major_gridlines_visible(
            source.drawable_object_id,
            Axis::Category,
            AxisVisibility::Hidden,
        )
        .unwrap();
    editor
        .set_body_chart_axis_major_gridlines_visible(
            source.drawable_object_id,
            Axis::Value,
            AxisVisibility::Visible,
        )
        .unwrap();

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    for axis in [Axis::Category, Axis::Value] {
        assert_eq!(
            reopened
                .body_chart_axis_major_gridlines_visible(source.drawable_object_id, axis)
                .unwrap()
                .is_visible(),
            axis == Axis::Value
        );
        assert_eq!(
            reopened
                .body_chart_axis_major_gridlines_visible(duplicate.drawable_object_id, axis)
                .unwrap()
                .is_visible(),
            axis == Axis::Category
        );
    }
    reopened
        .remove_body_chart(duplicate.drawable_object_id)
        .unwrap();
}

#[test]
fn scratch_document_supports_native_chart_axis_minor_gridline_visibility_crud() {
    let mut editor = PagesEditor::create_with_text("Chart minor gridlines").unwrap();
    let source = editor
        .add_body_chart(
            "Chart minor gridlines".encode_utf16().count(),
            Kind::Column2d,
            sample_data(),
            POSITION,
            SIZE,
        )
        .unwrap();

    for axis in [Axis::Category, Axis::Value] {
        assert!(
            !editor
                .body_chart_axis_minor_gridlines_visible(source.drawable_object_id, axis)
                .unwrap()
                .is_visible()
        );
    }
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_body_chart_axis_minor_gridlines_visible(
            source.drawable_object_id,
            Axis::Category,
            AxisVisibility::Hidden,
        )
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    for axis in [Axis::Category, Axis::Value] {
        editor
            .set_body_chart_axis_minor_gridlines_visible(
                source.drawable_object_id,
                axis,
                AxisVisibility::Visible,
            )
            .unwrap();
    }
    let duplicate = editor
        .duplicate_body_chart(
            source.drawable_object_id,
            editor.body_text().unwrap().encode_utf16().count(),
        )
        .unwrap();
    for axis in [Axis::Category, Axis::Value] {
        assert!(
            editor
                .body_chart_axis_minor_gridlines_visible(duplicate.drawable_object_id, axis)
                .unwrap()
                .is_visible()
        );
        editor
            .set_body_chart_axis_minor_gridlines_visible(
                source.drawable_object_id,
                axis,
                AxisVisibility::Hidden,
            )
            .unwrap();
    }

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    for axis in [Axis::Category, Axis::Value] {
        assert!(
            !reopened
                .body_chart_axis_minor_gridlines_visible(source.drawable_object_id, axis)
                .unwrap()
                .is_visible()
        );
        assert!(
            reopened
                .body_chart_axis_minor_gridlines_visible(duplicate.drawable_object_id, axis)
                .unwrap()
                .is_visible()
        );
    }
    reopened
        .remove_body_chart(duplicate.drawable_object_id)
        .unwrap();
}

#[test]
fn scratch_document_supports_native_chart_axis_minor_tick_mark_visibility_crud() {
    let mut editor = PagesEditor::create_with_text("Chart minor tick marks").unwrap();
    let source = editor
        .add_body_chart(
            "Chart minor tick marks".encode_utf16().count(),
            Kind::Column2d,
            sample_data(),
            POSITION,
            SIZE,
        )
        .unwrap();

    for axis in [Axis::Category, Axis::Value] {
        assert!(
            editor
                .body_chart_axis_minor_tick_marks_visible(source.drawable_object_id, axis)
                .unwrap()
                .is_visible()
        );
    }
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_body_chart_axis_minor_tick_marks_visible(
            source.drawable_object_id,
            Axis::Category,
            AxisVisibility::Visible,
        )
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    for axis in [Axis::Category, Axis::Value] {
        editor
            .set_body_chart_axis_minor_tick_marks_visible(
                source.drawable_object_id,
                axis,
                AxisVisibility::Hidden,
            )
            .unwrap();
        assert!(
            !editor
                .body_chart_axis_minor_tick_marks_visible(source.drawable_object_id, axis)
                .unwrap()
                .is_visible()
        );
    }

    let duplicate = editor
        .duplicate_body_chart(
            source.drawable_object_id,
            editor.body_text().unwrap().encode_utf16().count(),
        )
        .unwrap();
    for axis in [Axis::Category, Axis::Value] {
        assert!(
            !editor
                .body_chart_axis_minor_tick_marks_visible(duplicate.drawable_object_id, axis)
                .unwrap()
                .is_visible()
        );
        editor
            .set_body_chart_axis_minor_tick_marks_visible(
                source.drawable_object_id,
                axis,
                AxisVisibility::Visible,
            )
            .unwrap();
    }

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    for axis in [Axis::Category, Axis::Value] {
        assert!(
            reopened
                .body_chart_axis_minor_tick_marks_visible(source.drawable_object_id, axis)
                .unwrap()
                .is_visible()
        );
        assert!(
            !reopened
                .body_chart_axis_minor_tick_marks_visible(duplicate.drawable_object_id, axis)
                .unwrap()
                .is_visible()
        );
        reopened
            .set_body_chart_axis_minor_tick_marks_visible(
                source.drawable_object_id,
                axis,
                AxisVisibility::Hidden,
            )
            .unwrap();
        assert!(
            !reopened
                .body_chart_axis_minor_tick_marks_visible(source.drawable_object_id, axis)
                .unwrap()
                .is_visible()
        );
    }
    reopened
        .remove_body_chart(duplicate.drawable_object_id)
        .unwrap();
}

#[test]
fn scratch_document_supports_native_chart_axis_tick_mark_location_crud() {
    let mut editor = PagesEditor::create_with_text("Chart tick-mark locations").unwrap();
    let source = editor
        .add_body_chart(
            "Chart tick-mark locations".encode_utf16().count(),
            Kind::Column2d,
            sample_data(),
            POSITION,
            SIZE,
        )
        .unwrap();

    for axis in [Axis::Category, Axis::Value] {
        assert_eq!(
            editor
                .body_chart_axis_tick_mark_location(source.drawable_object_id, axis)
                .unwrap(),
            TickMarkLocation::Centered
        );
    }
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_body_chart_axis_tick_mark_location(
            source.drawable_object_id,
            Axis::Category,
            TickMarkLocation::Centered,
        )
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    editor
        .set_body_chart_axis_tick_mark_location(
            source.drawable_object_id,
            Axis::Category,
            TickMarkLocation::None,
        )
        .unwrap();
    editor
        .set_body_chart_axis_tick_mark_location(
            source.drawable_object_id,
            Axis::Value,
            TickMarkLocation::Outside,
        )
        .unwrap();
    assert_eq!(
        editor
            .body_chart_axis_tick_mark_location(source.drawable_object_id, Axis::Category)
            .unwrap(),
        TickMarkLocation::None
    );
    assert_eq!(
        editor
            .body_chart_axis_tick_mark_location(source.drawable_object_id, Axis::Value)
            .unwrap(),
        TickMarkLocation::Outside
    );

    let duplicate = editor
        .duplicate_body_chart(
            source.drawable_object_id,
            editor.body_text().unwrap().encode_utf16().count(),
        )
        .unwrap();
    assert_eq!(
        editor
            .body_chart_axis_tick_mark_location(duplicate.drawable_object_id, Axis::Category)
            .unwrap(),
        TickMarkLocation::None
    );
    assert_eq!(
        editor
            .body_chart_axis_tick_mark_location(duplicate.drawable_object_id, Axis::Value)
            .unwrap(),
        TickMarkLocation::Outside
    );

    editor
        .set_body_chart_axis_tick_mark_location(
            source.drawable_object_id,
            Axis::Category,
            TickMarkLocation::Inside,
        )
        .unwrap();
    editor
        .set_body_chart_axis_tick_mark_location(
            source.drawable_object_id,
            Axis::Value,
            TickMarkLocation::Centered,
        )
        .unwrap();

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .body_chart_axis_tick_mark_location(source.drawable_object_id, Axis::Category)
            .unwrap(),
        TickMarkLocation::Inside
    );
    assert_eq!(
        reopened
            .body_chart_axis_tick_mark_location(source.drawable_object_id, Axis::Value)
            .unwrap(),
        TickMarkLocation::Centered
    );
    assert_eq!(
        reopened
            .body_chart_axis_tick_mark_location(duplicate.drawable_object_id, Axis::Category)
            .unwrap(),
        TickMarkLocation::None
    );
    assert_eq!(
        reopened
            .body_chart_axis_tick_mark_location(duplicate.drawable_object_id, Axis::Value)
            .unwrap(),
        TickMarkLocation::Outside
    );
    reopened
        .remove_body_chart(duplicate.drawable_object_id)
        .unwrap();
}

#[test]
fn scratch_document_supports_native_chart_legend_visibility_crud() {
    let mut editor = PagesEditor::create_with_text("Chart legends").unwrap();
    let source = editor
        .add_body_chart(
            "Chart legends".encode_utf16().count(),
            Kind::Column2d,
            sample_data(),
            POSITION,
            SIZE,
        )
        .unwrap();

    assert!(
        editor
            .body_chart_legend_visible(source.drawable_object_id)
            .unwrap()
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_body_chart_legend_visible(source.drawable_object_id, true)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);
    editor
        .set_body_chart_legend_visible(source.drawable_object_id, false)
        .unwrap();
    assert!(
        !editor
            .body_chart_legend_visible(source.drawable_object_id)
            .unwrap()
    );

    let duplicate = editor
        .duplicate_body_chart(
            source.drawable_object_id,
            editor.body_text().unwrap().encode_utf16().count(),
        )
        .unwrap();
    assert!(
        !editor
            .body_chart_legend_visible(duplicate.drawable_object_id)
            .unwrap()
    );

    editor
        .set_body_chart_legend_visible(source.drawable_object_id, true)
        .unwrap();
    assert!(
        editor
            .body_chart_legend_visible(source.drawable_object_id)
            .unwrap()
    );
    assert!(
        !editor
            .body_chart_legend_visible(duplicate.drawable_object_id)
            .unwrap()
    );

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert!(
        reopened
            .body_chart_legend_visible(source.drawable_object_id)
            .unwrap()
    );
    assert!(
        !reopened
            .body_chart_legend_visible(duplicate.drawable_object_id)
            .unwrap()
    );
    reopened
        .remove_body_chart(source.drawable_object_id)
        .unwrap();
    reopened
        .remove_body_chart(duplicate.drawable_object_id)
        .unwrap();
    assert!(reopened.body_charts().unwrap().is_empty());
}

#[test]
fn scratch_document_supports_exact_chart_legend_fill_crud() {
    let mut editor = PagesEditor::create_with_text("Chart legend fill").unwrap();
    let chart = editor
        .add_body_chart(
            "Chart legend fill".encode_utf16().count(),
            Kind::Column2d,
            sample_data(),
            POSITION,
            SIZE,
        )
        .unwrap();
    let object_id = chart.drawable_object_id;
    let baseline = editor.to_bytes().unwrap();

    assert_eq!(
        editor.body_chart_legend_fill(object_id).unwrap(),
        ChartLegendFill::Inherited
    );
    let solid = ChartLegendFill::Fill(ShapeFill::Solid(
        RgbaColor::new(0.85, 0.25, 0.2, 1.0, RgbColorSpace::Srgb).unwrap(),
    ));
    editor
        .set_body_chart_legend_fill(object_id, &solid)
        .unwrap();
    assert_eq!(editor.body_chart_legend_fill(object_id).unwrap(), solid);
    assert!(editor.body_chart_legend_visible(object_id).unwrap());

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(reopened.body_chart_legend_fill(object_id).unwrap(), solid);
    reopened
        .set_body_chart_legend_fill(object_id, &ChartLegendFill::Inherited)
        .unwrap();
    assert_eq!(reopened.to_bytes().unwrap(), baseline);
}

#[test]
fn scratch_document_supports_exact_chart_legend_frame_crud() {
    let mut editor = PagesDocumentBuilder::new()
        .body_text("Legend")
        .build()
        .unwrap();
    let chart = editor
        .add_body_chart(
            "Legend".encode_utf16().count(),
            Kind::Column2d,
            sample_data(),
            POSITION,
            SIZE,
        )
        .unwrap();
    let object_id = chart.drawable_object_id;
    let baseline = editor.to_bytes().unwrap();

    assert_eq!(
        editor.body_chart_legend_frame(object_id).unwrap(),
        ChartLegendFrame::Automatic
    );
    let frame =
        ChartLegendFrame::Frame(ChartLegendRect::from_points(36.0, 18.0, 0.0, 0.0).unwrap());
    editor
        .set_body_chart_legend_frame(object_id, frame)
        .unwrap();
    assert_eq!(editor.body_chart_legend_frame(object_id).unwrap(), frame);

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(reopened.body_chart_legend_frame(object_id).unwrap(), frame);
    reopened
        .set_body_chart_legend_frame(object_id, ChartLegendFrame::Automatic)
        .unwrap();
    assert_eq!(reopened.to_bytes().unwrap(), baseline);
}

#[test]
fn scratch_document_supports_exact_chart_legend_typography_crud() {
    let mut editor = PagesDocumentBuilder::new().build().unwrap();
    let chart = editor
        .add_body_chart(0, Kind::Column2d, sample_data(), POSITION, SIZE)
        .unwrap();
    let object_id = chart.drawable_object_id;
    let baseline = editor.to_bytes().unwrap();

    assert_eq!(
        editor.body_chart_legend_font(object_id).unwrap(),
        ChartLegendFont::Inherited
    );
    let bold = ChartLegendFont::Font(ChartFont::named("AvenirNext-Bold").unwrap().with_bold(true));
    editor.set_body_chart_legend_font(object_id, &bold).unwrap();
    assert_eq!(editor.body_chart_legend_font(object_id).unwrap(), bold);

    assert_eq!(
        editor.body_chart_legend_font_size(object_id).unwrap(),
        ChartLegendFontSize::Inherited
    );
    let eighteen = ChartLegendFontSize::Size(ChartFontSize::from_points(18.0).unwrap());
    editor
        .set_body_chart_legend_font_size(object_id, eighteen)
        .unwrap();
    assert_eq!(
        editor.body_chart_legend_font_size(object_id).unwrap(),
        eighteen
    );

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(reopened.body_chart_legend_font(object_id).unwrap(), bold);
    let italic = ChartLegendFont::Font(
        ChartFont::named("AvenirNext-Italic")
            .unwrap()
            .with_italic(true),
    );
    reopened
        .set_body_chart_legend_font(object_id, &italic)
        .unwrap();
    let fifteen = ChartLegendFontSize::Size(ChartFontSize::from_points(15.0).unwrap());
    reopened
        .set_body_chart_legend_font_size(object_id, fifteen)
        .unwrap();
    assert_eq!(
        reopened.body_chart_legend_font_size(object_id).unwrap(),
        fifteen
    );
    reopened
        .set_body_chart_legend_font(object_id, &ChartLegendFont::Inherited)
        .unwrap();
    assert_eq!(
        reopened.body_chart_legend_font_size(object_id).unwrap(),
        fifteen
    );
    reopened
        .set_body_chart_legend_font_size(object_id, ChartLegendFontSize::Inherited)
        .unwrap();
    assert_eq!(reopened.to_bytes().unwrap(), baseline);
}

#[test]
fn scratch_document_supports_exact_chart_legend_stroke_crud() {
    let mut editor = PagesEditor::create_with_text("Chart legend stroke").unwrap();
    let chart = editor
        .add_body_chart(
            "Chart legend stroke".encode_utf16().count(),
            Kind::Column2d,
            sample_data(),
            POSITION,
            SIZE,
        )
        .unwrap();
    let object_id = chart.drawable_object_id;
    let baseline = editor.to_bytes().unwrap();

    assert_eq!(
        editor.body_chart_legend_stroke(object_id).unwrap(),
        ChartLegendStroke::Inherited
    );
    let stroke = ChartLegendStroke::Stroke(Stroke::new(
        RgbaColor::new(0.8, 0.2, 0.15, 1.0, RgbColorSpace::Srgb).unwrap(),
        Width::new(1.5).unwrap(),
        Pattern::RoundedDash,
    ));
    editor
        .set_body_chart_legend_stroke(object_id, stroke)
        .unwrap();
    assert_eq!(editor.body_chart_legend_stroke(object_id).unwrap(), stroke);
    assert!(editor.body_chart_legend_visible(object_id).unwrap());

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened.body_chart_legend_stroke(object_id).unwrap(),
        stroke
    );
    reopened
        .set_body_chart_legend_stroke(object_id, ChartLegendStroke::Inherited)
        .unwrap();
    assert_eq!(reopened.to_bytes().unwrap(), baseline);
}

#[test]
fn scratch_document_supports_exact_chart_legend_shadow_crud() {
    let mut editor = PagesEditor::create_with_text("Chart legend shadow").unwrap();
    let chart = editor
        .add_body_chart(
            "Chart legend shadow".encode_utf16().count(),
            Kind::Column2d,
            sample_data(),
            POSITION,
            SIZE,
        )
        .unwrap();
    let object_id = chart.drawable_object_id;
    let baseline = editor.to_bytes().unwrap();

    assert_eq!(
        editor.body_chart_legend_shadow(object_id).unwrap(),
        ChartLegendShadow::Inherited
    );
    let shadow = ChartLegendShadow::Shadow(Drop::new(
        Appearance::new(
            RgbaColor::black(),
            BlurRadius::from_points(11).unwrap(),
            Offset::from_points(7.0).unwrap(),
            Opacity::new(0.55).unwrap(),
        ),
        Angle::from_degrees(25.0).unwrap(),
    ));
    editor
        .set_body_chart_legend_shadow(object_id, shadow)
        .unwrap();
    assert_eq!(editor.body_chart_legend_shadow(object_id).unwrap(), shadow);
    assert!(editor.body_chart_legend_visible(object_id).unwrap());

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened.body_chart_legend_shadow(object_id).unwrap(),
        shadow
    );
    reopened
        .set_body_chart_legend_shadow(object_id, ChartLegendShadow::Inherited)
        .unwrap();
    assert_eq!(reopened.to_bytes().unwrap(), baseline);
}

#[test]
fn scratch_document_supports_native_chart_value_axis_scale_crud() {
    let mut editor = PagesEditor::create_with_text("Chart value-axis scale").unwrap();
    let source = editor
        .add_body_chart(
            "Chart value-axis scale".encode_utf16().count(),
            Kind::Column2d,
            sample_data(),
            POSITION,
            SIZE,
        )
        .unwrap();

    assert_eq!(
        editor
            .body_chart_value_axis_scale(source.drawable_object_id)
            .unwrap(),
        Scale::Linear
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_body_chart_value_axis_scale(source.drawable_object_id, Scale::Linear)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    editor
        .set_body_chart_value_axis_scale(source.drawable_object_id, Scale::Logarithmic)
        .unwrap();
    assert_eq!(
        editor
            .body_chart_value_axis_scale(source.drawable_object_id)
            .unwrap(),
        Scale::Logarithmic
    );

    let duplicate = editor
        .duplicate_body_chart(
            source.drawable_object_id,
            editor.body_text().unwrap().encode_utf16().count(),
        )
        .unwrap();
    assert_eq!(
        editor
            .body_chart_value_axis_scale(duplicate.drawable_object_id)
            .unwrap(),
        Scale::Logarithmic
    );
    editor
        .set_body_chart_value_axis_scale(source.drawable_object_id, Scale::Linear)
        .unwrap();

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .body_chart_value_axis_scale(source.drawable_object_id)
            .unwrap(),
        Scale::Linear
    );
    assert_eq!(
        reopened
            .body_chart_value_axis_scale(duplicate.drawable_object_id)
            .unwrap(),
        Scale::Logarithmic
    );
    reopened
        .remove_body_chart(source.drawable_object_id)
        .unwrap();
    reopened
        .remove_body_chart(duplicate.drawable_object_id)
        .unwrap();
}

#[test]
fn scratch_document_supports_native_chart_border_stroke_crud() {
    let mut editor = PagesEditor::create_with_text("Chart border stroke").unwrap();
    let source = editor
        .add_body_chart(
            "Chart border stroke".encode_utf16().count(),
            Kind::Column2d,
            sample_data(),
            POSITION,
            SIZE,
        )
        .unwrap();
    let default = Stroke::new(RgbaColor::black(), Width::ONE, Pattern::Solid);
    let customized = chart_stroke(Pattern::MediumDash, 3.0);
    let changed = chart_stroke(Pattern::RoundedDash, 2.0);

    assert_eq!(
        editor
            .body_chart_border_stroke(source.drawable_object_id)
            .unwrap(),
        Some(default)
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_body_chart_border_stroke(source.drawable_object_id, Some(default))
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    editor
        .set_body_chart_border_stroke(source.drawable_object_id, Some(customized))
        .unwrap();
    let duplicate = editor
        .duplicate_body_chart(
            source.drawable_object_id,
            editor.body_text().unwrap().encode_utf16().count(),
        )
        .unwrap();
    assert_eq!(
        editor
            .body_chart_border_stroke(duplicate.drawable_object_id)
            .unwrap(),
        Some(customized)
    );
    editor
        .set_body_chart_border_stroke(source.drawable_object_id, Some(changed))
        .unwrap();
    editor
        .set_body_chart_border_stroke(duplicate.drawable_object_id, None)
        .unwrap();

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .body_chart_border_stroke(source.drawable_object_id)
            .unwrap(),
        Some(changed)
    );
    assert_eq!(
        reopened
            .body_chart_border_stroke(duplicate.drawable_object_id)
            .unwrap(),
        None
    );
    reopened
        .remove_body_chart(source.drawable_object_id)
        .unwrap();
    reopened
        .remove_body_chart(duplicate.drawable_object_id)
        .unwrap();
    assert!(reopened.body_charts().unwrap().is_empty());
}

#[test]
fn scratch_document_supports_native_chart_background_fill_crud() {
    let image_bytes = fixture("test-data/images/png/lena.png");
    let mut editor = PagesEditor::create_with_text("Chart background fill").unwrap();
    let source = editor
        .add_body_chart(
            "Chart background fill".encode_utf16().count(),
            Kind::Column2d,
            sample_data(),
            POSITION,
            SIZE,
        )
        .unwrap();
    let native_default = editor
        .body_chart_background_fill(source.drawable_object_id)
        .unwrap();
    assert!(matches!(native_default, ShapeFill::Gradient(_)));
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_body_chart_background_fill(source.drawable_object_id, &native_default)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    let customized = chart_background_fill();
    let image = editor
        .set_body_chart_background_image_fill(
            source.drawable_object_id,
            "lena.png",
            &image_bytes,
            ShapeImageFillTechnique::ScaleToFill,
            None,
        )
        .unwrap();
    let duplicate = editor
        .duplicate_body_chart(
            source.drawable_object_id,
            editor.body_text().unwrap().encode_utf16().count(),
        )
        .unwrap();
    assert_eq!(
        editor
            .body_chart_background_fill(duplicate.drawable_object_id)
            .unwrap(),
        ShapeFill::Image(image.clone())
    );
    editor
        .set_body_chart_background_fill(source.drawable_object_id, &customized)
        .unwrap();

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .body_chart_background_fill(source.drawable_object_id)
            .unwrap(),
        customized
    );
    assert_eq!(
        reopened
            .body_chart_background_fill(duplicate.drawable_object_id)
            .unwrap(),
        ShapeFill::Image(image.clone())
    );
    assert_eq!(
        reopened
            .extract_media(
                crate::MediaAssetId::try_from(image.data_identifier().unwrap().get()).unwrap(),
            )
            .unwrap(),
        image_bytes
    );
    reopened
        .remove_body_chart(duplicate.drawable_object_id)
        .unwrap();
    assert!(reopened.media_assets().unwrap().is_empty());
    reopened
        .remove_body_chart(source.drawable_object_id)
        .unwrap();
    assert!(reopened.body_charts().unwrap().is_empty());
}

#[test]
fn scratch_document_supports_inherited_series_fill_crud() {
    let image_bytes = fixture("test-data/images/png/lena.png");
    let mut editor = PagesEditor::create_with_text("Chart").unwrap();
    let source = editor
        .add_body_chart(5, Kind::Column2d, sample_data(), POSITION, SIZE)
        .unwrap();
    let defaults = editor
        .body_chart_series_fills(source.drawable_object_id)
        .unwrap();
    assert_eq!(defaults.len(), 2);
    assert!(
        defaults
            .iter()
            .all(|fill| matches!(fill, ShapeFill::Solid(_)))
    );
    assert_ne!(defaults[0], defaults[1]);
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_body_chart_series_fills(source.drawable_object_id, &defaults)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    let first = Index::from_zero_based(0);
    let second = Index::from_zero_based(1);
    editor
        .set_body_chart_series_fill(source.drawable_object_id, first, &ShapeFill::None)
        .unwrap();
    let image = editor
        .set_body_chart_series_image_fill(
            source.drawable_object_id,
            second,
            "lena.png",
            &image_bytes,
            ShapeImageFillTechnique::ScaleToFit,
            None,
        )
        .unwrap();
    let duplicate = editor
        .duplicate_body_chart(source.drawable_object_id, 6)
        .unwrap();
    assert_eq!(
        editor
            .body_chart_series_fills(duplicate.drawable_object_id)
            .unwrap(),
        vec![ShapeFill::None, ShapeFill::Image(image.clone())]
    );
    assert_eq!(
        editor
            .reset_body_chart_series_fill(source.drawable_object_id, first)
            .unwrap(),
        defaults[0]
    );

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .body_chart_series_fill(source.drawable_object_id, first)
            .unwrap(),
        defaults[0]
    );
    assert_eq!(
        reopened
            .body_chart_series_fill(source.drawable_object_id, second)
            .unwrap(),
        ShapeFill::Image(image.clone())
    );
    assert_eq!(
        reopened
            .extract_media(
                crate::MediaAssetId::try_from(image.data_identifier().unwrap().get()).unwrap(),
            )
            .unwrap(),
        image_bytes
    );
    reopened
        .remove_body_chart(source.drawable_object_id)
        .unwrap();
    assert_eq!(reopened.media_assets().unwrap().len(), 1);
    reopened
        .remove_body_chart(duplicate.drawable_object_id)
        .unwrap();
    assert!(reopened.media_assets().unwrap().is_empty());
}

#[test]
fn scratch_document_supports_inherited_series_stroke_crud() {
    let mut editor = PagesEditor::create_with_text("Chart").unwrap();
    let source = editor
        .add_body_chart(5, Kind::Column2d, sample_data(), POSITION, SIZE)
        .unwrap();
    let defaults = vec![None, None];
    assert_eq!(
        editor
            .body_chart_series_strokes(source.drawable_object_id)
            .unwrap(),
        defaults
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_body_chart_series_strokes(source.drawable_object_id, &defaults)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    let first = Index::from_zero_based(0);
    let second = Index::from_zero_based(1);
    let rounded = chart_series_stroke(ChartSeriesStrokePattern::RoundedDash, 3.5);
    let medium = chart_series_stroke(ChartSeriesStrokePattern::MediumDash, 2.0);
    editor
        .set_body_chart_series_strokes(source.drawable_object_id, &[Some(rounded), Some(medium)])
        .unwrap();
    let duplicate = editor
        .duplicate_body_chart(source.drawable_object_id, 6)
        .unwrap();
    assert_eq!(
        editor
            .body_chart_series_strokes(duplicate.drawable_object_id)
            .unwrap(),
        vec![Some(rounded), Some(medium)]
    );
    editor
        .set_body_chart_series_stroke(source.drawable_object_id, first, None)
        .unwrap();
    assert_eq!(
        editor
            .reset_body_chart_series_stroke(source.drawable_object_id, first)
            .unwrap(),
        None
    );

    let reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .body_chart_series_stroke(source.drawable_object_id, first)
            .unwrap(),
        None
    );
    assert_eq!(
        reopened
            .body_chart_series_stroke(source.drawable_object_id, second)
            .unwrap(),
        Some(medium)
    );
    assert_eq!(
        reopened
            .body_chart_series_strokes(duplicate.drawable_object_id)
            .unwrap(),
        vec![Some(rounded), Some(medium)]
    );
}

#[test]
fn scratch_document_supports_native_chart_shadow_crud() {
    let mut editor = PagesEditor::create_with_text("Chart shadow").unwrap();
    let source = editor
        .add_body_chart(
            "Chart shadow".encode_utf16().count(),
            Kind::Column2d,
            sample_data(),
            POSITION,
            SIZE,
        )
        .unwrap();
    let native_default = ChartShadow::native_default();
    assert_eq!(
        editor.body_chart_shadow(source.drawable_object_id).unwrap(),
        native_default
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_body_chart_shadow(source.drawable_object_id, native_default)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    let customized = chart_shadow();
    editor
        .set_body_chart_shadow(source.drawable_object_id, customized)
        .unwrap();
    let duplicate = editor
        .duplicate_body_chart(
            source.drawable_object_id,
            editor.body_text().unwrap().encode_utf16().count(),
        )
        .unwrap();
    assert_eq!(
        editor
            .body_chart_shadow(duplicate.drawable_object_id)
            .unwrap(),
        customized
    );
    editor
        .set_body_chart_shadow(source.drawable_object_id, ChartShadow::None)
        .unwrap();

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .body_chart_shadow(source.drawable_object_id)
            .unwrap(),
        ChartShadow::None
    );
    assert_eq!(
        reopened
            .body_chart_shadow(duplicate.drawable_object_id)
            .unwrap(),
        customized
    );
    reopened
        .set_body_chart_shadow(duplicate.drawable_object_id, native_default)
        .unwrap();
    reopened
        .remove_body_chart(source.drawable_object_id)
        .unwrap();
    reopened
        .remove_body_chart(duplicate.drawable_object_id)
        .unwrap();
    assert!(reopened.body_charts().unwrap().is_empty());
}

#[test]
fn scratch_document_supports_native_pie_start_angle_crud() {
    let mut editor = PagesEditor::create_with_text("Pie rotation").unwrap();
    let source = editor
        .add_body_chart(
            "Pie rotation".encode_utf16().count(),
            Kind::Pie2d,
            pie_data(),
            POSITION,
            SIZE,
        )
        .unwrap();
    assert_eq!(
        editor
            .body_chart_pie_start_angle(source.drawable_object_id)
            .unwrap(),
        ChartPieStartAngle::ZERO
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_body_chart_pie_start_angle(source.drawable_object_id, ChartPieStartAngle::ZERO)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    let customized = ChartPieStartAngle::from_degrees(123.0).unwrap();
    editor
        .set_body_chart_pie_start_angle(source.drawable_object_id, customized)
        .unwrap();
    let duplicate = editor
        .duplicate_body_chart(
            source.drawable_object_id,
            editor.body_text().unwrap().encode_utf16().count(),
        )
        .unwrap();
    editor
        .set_body_chart_kind(duplicate.drawable_object_id, Kind::Donut2d)
        .unwrap();
    assert_eq!(
        editor
            .body_chart_pie_start_angle(duplicate.drawable_object_id)
            .unwrap(),
        customized
    );
    editor
        .set_body_chart_pie_start_angle(source.drawable_object_id, ChartPieStartAngle::HALF_TURN)
        .unwrap();

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .body_chart_pie_start_angle(source.drawable_object_id)
            .unwrap(),
        ChartPieStartAngle::HALF_TURN
    );
    assert_eq!(
        reopened
            .body_chart_pie_start_angle(duplicate.drawable_object_id)
            .unwrap(),
        customized
    );
    reopened
        .set_body_chart_pie_start_angle(duplicate.drawable_object_id, ChartPieStartAngle::ZERO)
        .unwrap();

    let column = reopened
        .add_body_chart(
            reopened.body_text().unwrap().encode_utf16().count(),
            Kind::Column2d,
            sample_data(),
            POSITION,
            SIZE,
        )
        .unwrap();
    let before_rejected_update = reopened.to_bytes().unwrap();
    assert!(
        reopened
            .body_chart_pie_start_angle(column.drawable_object_id)
            .is_err()
    );
    assert!(
        reopened
            .set_body_chart_pie_start_angle(
                column.drawable_object_id,
                ChartPieStartAngle::QUARTER_TURN,
            )
            .is_err()
    );
    assert_eq!(reopened.to_bytes().unwrap(), before_rejected_update);

    reopened
        .remove_body_chart(source.drawable_object_id)
        .unwrap();
    reopened
        .remove_body_chart(duplicate.drawable_object_id)
        .unwrap();
    reopened
        .remove_body_chart(column.drawable_object_id)
        .unwrap();
    assert!(reopened.body_charts().unwrap().is_empty());
}

#[test]
fn scratch_document_supports_native_donut_inner_radius_crud() {
    let mut editor = PagesEditor::create_with_text("Donut radius").unwrap();
    let source = editor
        .add_body_chart(
            "Donut radius".encode_utf16().count(),
            Kind::Donut2d,
            pie_data(),
            POSITION,
            SIZE,
        )
        .unwrap();
    assert_eq!(
        editor
            .body_chart_donut_inner_radius(source.drawable_object_id)
            .unwrap(),
        ChartDonutInnerRadius::DEFAULT
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_body_chart_donut_inner_radius(
            source.drawable_object_id,
            ChartDonutInnerRadius::DEFAULT,
        )
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    let customized = ChartDonutInnerRadius::from_percent(42.0).unwrap();
    editor
        .set_body_chart_donut_inner_radius(source.drawable_object_id, customized)
        .unwrap();
    let duplicate = editor
        .duplicate_body_chart(
            source.drawable_object_id,
            editor.body_text().unwrap().encode_utf16().count(),
        )
        .unwrap();
    assert_eq!(
        editor
            .body_chart_donut_inner_radius(duplicate.drawable_object_id)
            .unwrap(),
        customized
    );
    editor
        .set_body_chart_kind(duplicate.drawable_object_id, Kind::Pie2d)
        .unwrap();
    let before_rejected_update = editor.to_bytes().unwrap();
    assert!(
        editor
            .body_chart_donut_inner_radius(duplicate.drawable_object_id)
            .is_err()
    );
    assert!(
        editor
            .set_body_chart_donut_inner_radius(
                duplicate.drawable_object_id,
                ChartDonutInnerRadius::MAXIMUM,
            )
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before_rejected_update);
    editor
        .set_body_chart_kind(duplicate.drawable_object_id, Kind::Donut3d)
        .unwrap();

    editor
        .set_body_chart_donut_inner_radius(
            source.drawable_object_id,
            ChartDonutInnerRadius::MINIMUM,
        )
        .unwrap();
    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .body_chart_donut_inner_radius(source.drawable_object_id)
            .unwrap(),
        ChartDonutInnerRadius::MINIMUM
    );
    assert_eq!(
        reopened
            .body_chart_donut_inner_radius(duplicate.drawable_object_id)
            .unwrap(),
        customized
    );
    reopened
        .set_body_chart_donut_inner_radius(
            duplicate.drawable_object_id,
            ChartDonutInnerRadius::DEFAULT,
        )
        .unwrap();
    reopened
        .remove_body_chart(source.drawable_object_id)
        .unwrap();
    reopened
        .remove_body_chart(duplicate.drawable_object_id)
        .unwrap();
    assert!(reopened.body_charts().unwrap().is_empty());
}

#[test]
fn scratch_document_supports_native_pie_wedge_explosion_crud() {
    let mut editor = PagesEditor::create_with_text("Revenue").unwrap();
    let source = editor
        .add_body_chart(7, Kind::Pie2d, pie_data(), POSITION, SIZE)
        .unwrap();
    let zeros = vec![ChartPieWedgeExplosion::ZERO; 3];
    assert_eq!(
        editor
            .body_chart_pie_wedge_explosions(source.drawable_object_id)
            .unwrap(),
        zeros
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_body_chart_pie_wedge_explosions(source.drawable_object_id, &zeros)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    let customized = [
        ChartPieWedgeExplosion::from_percent(10.0).unwrap(),
        ChartPieWedgeExplosion::from_percent(25.0).unwrap(),
        ChartPieWedgeExplosion::from_percent(40.0).unwrap(),
    ];
    editor
        .set_body_chart_pie_wedge_explosions(source.drawable_object_id, &customized)
        .unwrap();
    assert_eq!(
        editor
            .body_chart_pie_wedge_explosion(
                source.drawable_object_id,
                ChartPieWedgeIndex::from_zero_based(1),
            )
            .unwrap(),
        customized[1]
    );
    editor
        .set_body_chart_pie_wedge_explosions(source.drawable_object_id, &zeros)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    editor
        .set_body_chart_pie_wedge_explosions(source.drawable_object_id, &customized)
        .unwrap();
    let duplicate = editor
        .duplicate_body_chart(source.drawable_object_id, 7)
        .unwrap();
    editor
        .set_body_chart_kind(duplicate.drawable_object_id, Kind::Donut2d)
        .unwrap();
    assert_eq!(
        editor
            .body_chart_pie_wedge_explosions(duplicate.drawable_object_id)
            .unwrap(),
        customized
    );
    let isolated = ChartPieWedgeExplosion::from_percent(55.0).unwrap();
    editor
        .set_body_chart_pie_wedge_explosion(
            source.drawable_object_id,
            ChartPieWedgeIndex::from_zero_based(0),
            isolated,
        )
        .unwrap();

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .body_chart_pie_wedge_explosion(
                source.drawable_object_id,
                ChartPieWedgeIndex::from_zero_based(0),
            )
            .unwrap(),
        isolated
    );
    assert_eq!(
        reopened
            .body_chart_pie_wedge_explosions(duplicate.drawable_object_id)
            .unwrap(),
        customized
    );

    let before_rejected_updates = reopened.to_bytes().unwrap();
    assert!(
        reopened
            .set_body_chart_pie_wedge_explosions(source.drawable_object_id, &customized[..2],)
            .is_err()
    );
    assert!(
        reopened
            .set_body_chart_pie_wedge_explosion(
                source.drawable_object_id,
                ChartPieWedgeIndex::from_zero_based(3),
                isolated,
            )
            .is_err()
    );
    assert_eq!(reopened.to_bytes().unwrap(), before_rejected_updates);

    let column = reopened
        .add_body_chart(7, Kind::Column2d, sample_data(), POSITION, SIZE)
        .unwrap();
    let before_wrong_kind = reopened.to_bytes().unwrap();
    assert!(
        reopened
            .body_chart_pie_wedge_explosions(column.drawable_object_id)
            .is_err()
    );
    assert!(
        reopened
            .set_body_chart_pie_wedge_explosions(column.drawable_object_id, &customized,)
            .is_err()
    );
    assert_eq!(reopened.to_bytes().unwrap(), before_wrong_kind);

    reopened
        .remove_body_chart(source.drawable_object_id)
        .unwrap();
    reopened
        .remove_body_chart(duplicate.drawable_object_id)
        .unwrap();
    reopened
        .remove_body_chart(column.drawable_object_id)
        .unwrap();
    assert!(reopened.body_charts().unwrap().is_empty());
}

#[test]
fn scratch_document_supports_native_pie_label_visibility_crud() {
    let mut editor = PagesEditor::create_with_text("Revenue").unwrap();
    let source = editor
        .add_body_chart(7, Kind::Pie2d, pie_data(), POSITION, SIZE)
        .unwrap();
    let defaults = vec![LabelVisibility::DEFAULT; 3];
    let customized = [
        LabelVisibility::DATA_POINT_NAMES_ONLY,
        LabelVisibility::ALL,
        LabelVisibility::HIDDEN,
    ];
    assert_eq!(
        editor
            .body_chart_pie_label_visibilities(source.drawable_object_id)
            .unwrap(),
        defaults
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_body_chart_pie_label_visibilities(source.drawable_object_id, &defaults)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    editor
        .set_body_chart_pie_label_visibilities(source.drawable_object_id, &customized)
        .unwrap();
    assert_eq!(
        editor
            .body_chart_pie_label_visibility(
                source.drawable_object_id,
                ChartPieWedgeIndex::from_zero_based(1),
            )
            .unwrap(),
        LabelVisibility::ALL
    );
    let explosions = [
        ChartPieWedgeExplosion::from_percent(10.0).unwrap(),
        ChartPieWedgeExplosion::from_percent(25.0).unwrap(),
        ChartPieWedgeExplosion::from_percent(40.0).unwrap(),
    ];
    editor
        .set_body_chart_pie_wedge_explosions(source.drawable_object_id, &explosions)
        .unwrap();
    editor
        .set_body_chart_pie_label_visibilities(source.drawable_object_id, &defaults)
        .unwrap();
    assert_eq!(
        editor
            .body_chart_pie_wedge_explosions(source.drawable_object_id)
            .unwrap(),
        explosions
    );
    editor
        .set_body_chart_pie_wedge_explosions(
            source.drawable_object_id,
            &[ChartPieWedgeExplosion::ZERO; 3],
        )
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    editor
        .set_body_chart_pie_label_visibilities(source.drawable_object_id, &customized)
        .unwrap();
    let duplicate = editor
        .duplicate_body_chart(source.drawable_object_id, 7)
        .unwrap();
    editor
        .set_body_chart_kind(duplicate.drawable_object_id, Kind::Donut2d)
        .unwrap();
    editor
        .set_body_chart_pie_label_visibility(
            source.drawable_object_id,
            ChartPieWedgeIndex::from_zero_based(0),
            LabelVisibility::VALUES_ONLY,
        )
        .unwrap();
    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .body_chart_pie_label_visibilities(duplicate.drawable_object_id)
            .unwrap(),
        customized
    );
    let before_rejected = reopened.to_bytes().unwrap();
    assert!(
        reopened
            .set_body_chart_pie_label_visibilities(source.drawable_object_id, &customized[..2],)
            .is_err()
    );
    assert_eq!(reopened.to_bytes().unwrap(), before_rejected);
    reopened
        .remove_body_chart(source.drawable_object_id)
        .unwrap();
    reopened
        .remove_body_chart(duplicate.drawable_object_id)
        .unwrap();
    assert!(reopened.body_charts().unwrap().is_empty());
}

#[test]
fn scratch_document_supports_native_pie_label_distance_crud() {
    let mut editor = PagesEditor::create_with_text("Revenue").unwrap();
    let source = editor
        .add_body_chart(7, Kind::Pie2d, pie_data(), POSITION, SIZE)
        .unwrap();
    let defaults = vec![ChartPieLabelDistance::DEFAULT; 3];
    let customized = [
        ChartPieLabelDistance::MINIMUM,
        ChartPieLabelDistance::from_percent(100.0).unwrap(),
        ChartPieLabelDistance::MAXIMUM,
    ];
    assert_eq!(
        editor
            .body_chart_pie_label_distances(source.drawable_object_id)
            .unwrap(),
        defaults
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_body_chart_pie_label_distances(source.drawable_object_id, &defaults)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);
    let leader_line_defaults = [LeaderLineVisibility::Visible; 3];
    let leader_line_customized = [
        LeaderLineVisibility::Hidden,
        LeaderLineVisibility::Visible,
        LeaderLineVisibility::Hidden,
    ];
    assert_eq!(
        editor
            .body_chart_pie_leader_line_visibilities(source.drawable_object_id)
            .unwrap(),
        leader_line_defaults
    );
    editor
        .set_body_chart_pie_leader_line_visibilities(
            source.drawable_object_id,
            &leader_line_defaults,
        )
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);
    editor
        .set_body_chart_pie_leader_line_visibilities(
            source.drawable_object_id,
            &leader_line_customized,
        )
        .unwrap();
    assert_eq!(
        editor
            .body_chart_pie_leader_line_visibility(
                source.drawable_object_id,
                ChartPieWedgeIndex::from_zero_based(0),
            )
            .unwrap(),
        LeaderLineVisibility::Hidden
    );
    editor
        .set_body_chart_pie_leader_line_visibilities(
            source.drawable_object_id,
            &leader_line_defaults,
        )
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    editor
        .set_body_chart_pie_label_distances(source.drawable_object_id, &customized)
        .unwrap();
    assert_eq!(
        editor
            .body_chart_pie_label_distance(
                source.drawable_object_id,
                ChartPieWedgeIndex::from_zero_based(1),
            )
            .unwrap(),
        customized[1]
    );
    let visibilities = [
        LabelVisibility::DATA_POINT_NAMES_ONLY,
        LabelVisibility::ALL,
        LabelVisibility::VALUES_ONLY,
    ];
    editor
        .set_body_chart_pie_label_visibilities(source.drawable_object_id, &visibilities)
        .unwrap();
    editor
        .set_body_chart_pie_label_distances(source.drawable_object_id, &defaults)
        .unwrap();
    assert_eq!(
        editor
            .body_chart_pie_label_visibilities(source.drawable_object_id)
            .unwrap(),
        visibilities
    );
    editor
        .set_body_chart_pie_label_visibilities(
            source.drawable_object_id,
            &[LabelVisibility::DEFAULT; 3],
        )
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    editor
        .set_body_chart_pie_label_distances(source.drawable_object_id, &customized)
        .unwrap();
    editor
        .set_body_chart_pie_leader_line_visibilities(
            source.drawable_object_id,
            &leader_line_customized,
        )
        .unwrap();
    let duplicate = editor
        .duplicate_body_chart(source.drawable_object_id, 7)
        .unwrap();
    editor
        .set_body_chart_kind(duplicate.drawable_object_id, Kind::Donut2d)
        .unwrap();
    editor
        .set_body_chart_pie_label_distance(
            source.drawable_object_id,
            ChartPieWedgeIndex::from_zero_based(0),
            ChartPieLabelDistance::DEFAULT,
        )
        .unwrap();
    editor
        .set_body_chart_pie_leader_line_visibility(
            source.drawable_object_id,
            ChartPieWedgeIndex::from_zero_based(0),
            LeaderLineVisibility::Visible,
        )
        .unwrap();
    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .body_chart_pie_label_distances(duplicate.drawable_object_id)
            .unwrap(),
        customized
    );
    assert_eq!(
        reopened
            .body_chart_pie_leader_line_visibilities(duplicate.drawable_object_id)
            .unwrap(),
        leader_line_customized
    );
    assert_eq!(
        reopened
            .body_chart_pie_leader_line_visibilities(source.drawable_object_id)
            .unwrap(),
        [
            LeaderLineVisibility::Visible,
            LeaderLineVisibility::Visible,
            LeaderLineVisibility::Hidden,
        ]
    );
    let before_rejected = reopened.to_bytes().unwrap();
    assert!(
        reopened
            .set_body_chart_pie_label_distances(source.drawable_object_id, &customized[..2],)
            .is_err()
    );
    assert!(
        reopened
            .set_body_chart_pie_leader_line_visibilities(
                source.drawable_object_id,
                &leader_line_customized[..2],
            )
            .is_err()
    );
    assert!(
        reopened
            .body_chart_pie_leader_line_visibility(
                source.drawable_object_id,
                ChartPieWedgeIndex::from_zero_based(3),
            )
            .is_err()
    );
    assert!(
        reopened
            .body_chart_pie_label_distance(
                source.drawable_object_id,
                ChartPieWedgeIndex::from_zero_based(3),
            )
            .is_err()
    );
    assert_eq!(reopened.to_bytes().unwrap(), before_rejected);
    reopened
        .remove_body_chart(source.drawable_object_id)
        .unwrap();
    reopened
        .remove_body_chart(duplicate.drawable_object_id)
        .unwrap();
    assert!(reopened.body_charts().unwrap().is_empty());
}

#[test]
fn scratch_document_supports_native_series_value_label_crud() {
    const BODY_TEXT: &str = "Chart value labels";

    let mut editor = PagesEditor::create_with_text(BODY_TEXT).unwrap();
    let source = editor
        .add_body_chart(
            BODY_TEXT.encode_utf16().count(),
            Kind::Column2d,
            sample_data(),
            POSITION,
            SIZE,
        )
        .unwrap();
    let defaults = [Visibility::Hidden; 2];
    let customized = [Visibility::Visible, Visibility::Hidden];

    assert_eq!(
        editor
            .body_chart_series_value_label_visibilities(source.drawable_object_id)
            .unwrap(),
        defaults
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_body_chart_series_value_label_visibilities(source.drawable_object_id, &defaults)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    editor
        .set_body_chart_series_value_label_visibilities(source.drawable_object_id, &customized)
        .unwrap();
    assert_eq!(
        editor
            .body_chart_series_value_label_visibility(
                source.drawable_object_id,
                Index::from_zero_based(0),
            )
            .unwrap(),
        Visibility::Visible
    );
    editor
        .set_body_chart_series_value_label_visibilities(source.drawable_object_id, &defaults)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    editor
        .set_body_chart_series_value_label_visibilities(source.drawable_object_id, &customized)
        .unwrap();
    let duplicate = editor
        .duplicate_body_chart(
            source.drawable_object_id,
            editor.body_text().unwrap().encode_utf16().count(),
        )
        .unwrap();
    editor
        .set_body_chart_series_value_label_visibility(
            source.drawable_object_id,
            Index::from_zero_based(0),
            Visibility::Hidden,
        )
        .unwrap();
    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .body_chart_series_value_label_visibilities(source.drawable_object_id)
            .unwrap(),
        defaults
    );
    assert_eq!(
        reopened
            .body_chart_series_value_label_visibilities(duplicate.drawable_object_id)
            .unwrap(),
        customized
    );

    let before_rejected = reopened.to_bytes().unwrap();
    assert!(
        reopened
            .set_body_chart_series_value_label_visibilities(
                source.drawable_object_id,
                &customized[..1],
            )
            .is_err()
    );
    assert!(
        reopened
            .body_chart_series_value_label_visibility(
                source.drawable_object_id,
                Index::from_zero_based(2),
            )
            .is_err()
    );
    assert_eq!(reopened.to_bytes().unwrap(), before_rejected);
    reopened
        .remove_body_chart(source.drawable_object_id)
        .unwrap();
    reopened
        .remove_body_chart(duplicate.drawable_object_id)
        .unwrap();
    assert!(reopened.body_charts().unwrap().is_empty());
}

#[test]
fn scratch_document_supports_native_series_value_label_location_crud() {
    const BODY_TEXT: &str = "Chart value-label Location";

    let mut editor = PagesEditor::create_with_text(BODY_TEXT).unwrap();
    let source = editor
        .add_body_chart(
            BODY_TEXT.encode_utf16().count(),
            Kind::Column2d,
            sample_data(),
            POSITION,
            SIZE,
        )
        .unwrap();
    let defaults = [ChartSeriesValueLabelLocation::Top; 2];
    let customized = [
        ChartSeriesValueLabelLocation::Outside,
        ChartSeriesValueLabelLocation::Top,
    ];

    assert_eq!(
        editor
            .body_chart_series_value_label_locations(source.drawable_object_id)
            .unwrap(),
        defaults
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_body_chart_series_value_label_locations(source.drawable_object_id, &defaults)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    editor
        .set_body_chart_series_value_label_locations(source.drawable_object_id, &customized)
        .unwrap();
    assert_eq!(
        editor
            .body_chart_series_value_label_location(
                source.drawable_object_id,
                Index::from_zero_based(0),
            )
            .unwrap(),
        ChartSeriesValueLabelLocation::Outside
    );
    editor
        .set_body_chart_series_value_label_locations(source.drawable_object_id, &defaults)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    editor
        .set_body_chart_series_value_label_locations(source.drawable_object_id, &customized)
        .unwrap();
    let duplicate = editor
        .duplicate_body_chart(
            source.drawable_object_id,
            editor.body_text().unwrap().encode_utf16().count(),
        )
        .unwrap();
    editor
        .set_body_chart_series_value_label_location(
            source.drawable_object_id,
            Index::from_zero_based(0),
            ChartSeriesValueLabelLocation::Top,
        )
        .unwrap();
    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .body_chart_series_value_label_locations(source.drawable_object_id)
            .unwrap(),
        defaults
    );
    assert_eq!(
        reopened
            .body_chart_series_value_label_locations(duplicate.drawable_object_id)
            .unwrap(),
        customized
    );

    let before_rejected = reopened.to_bytes().unwrap();
    assert!(
        reopened
            .set_body_chart_series_value_label_locations(
                source.drawable_object_id,
                &customized[..1],
            )
            .is_err()
    );
    assert!(
        reopened
            .body_chart_series_value_label_location(
                source.drawable_object_id,
                Index::from_zero_based(2),
            )
            .is_err()
    );
    assert_eq!(reopened.to_bytes().unwrap(), before_rejected);
    reopened
        .remove_body_chart(source.drawable_object_id)
        .unwrap();
    reopened
        .remove_body_chart(duplicate.drawable_object_id)
        .unwrap();
    assert!(reopened.body_charts().unwrap().is_empty());
}

#[test]
fn scratch_document_supports_native_series_value_label_affix_crud() {
    const BODY_TEXT: &str = "Chart value-label affixes";

    let mut editor = PagesEditor::create_with_text(BODY_TEXT).unwrap();
    let source = editor
        .add_body_chart(
            BODY_TEXT.encode_utf16().count(),
            Kind::Column2d,
            sample_data(),
            POSITION,
            SIZE,
        )
        .unwrap();
    let defaults = vec![LabelAffixes::default(); 2];
    let customized = vec![
        LabelAffixes::new("$", " USD").unwrap(),
        LabelAffixes::new("€", " net").unwrap(),
    ];

    assert_eq!(
        editor
            .body_chart_series_value_label_affixes(source.drawable_object_id)
            .unwrap(),
        defaults
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_body_chart_series_value_label_affixes(source.drawable_object_id, &defaults)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    editor
        .set_body_chart_series_value_label_affixes(source.drawable_object_id, &customized)
        .unwrap();
    assert_eq!(
        editor
            .body_chart_series_value_label_affix(
                source.drawable_object_id,
                Index::from_zero_based(0),
            )
            .unwrap()
            .suffix(),
        " USD"
    );
    let duplicate = editor
        .duplicate_body_chart(
            source.drawable_object_id,
            editor.body_text().unwrap().encode_utf16().count(),
        )
        .unwrap();
    for series in 0..2 {
        editor
            .set_body_chart_series_value_label_affix(
                source.drawable_object_id,
                Index::from_zero_based(series),
                LabelAffixes::default(),
            )
            .unwrap();
    }

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .body_chart_series_value_label_affixes(source.drawable_object_id)
            .unwrap(),
        defaults
    );
    assert_eq!(
        reopened
            .body_chart_series_value_label_affixes(duplicate.drawable_object_id)
            .unwrap(),
        customized
    );

    let before_rejected = reopened.to_bytes().unwrap();
    assert!(
        reopened
            .set_body_chart_series_value_label_affixes(source.drawable_object_id, &customized[..1],)
            .is_err()
    );
    assert!(
        reopened
            .body_chart_series_value_label_affix(
                source.drawable_object_id,
                Index::from_zero_based(2),
            )
            .is_err()
    );
    assert_eq!(reopened.to_bytes().unwrap(), before_rejected);
    reopened
        .remove_body_chart(source.drawable_object_id)
        .unwrap();
    reopened
        .remove_body_chart(duplicate.drawable_object_id)
        .unwrap();
    assert!(reopened.body_charts().unwrap().is_empty());
}

#[test]
fn scratch_document_supports_native_series_value_label_number_format_crud() {
    const BODY_TEXT: &str = "Chart value-label number formats";

    let mut editor = PagesEditor::create_with_text(BODY_TEXT).unwrap();
    let source = editor
        .add_body_chart(
            BODY_TEXT.encode_utf16().count(),
            Kind::Column2d,
            sample_data(),
            POSITION,
            SIZE,
        )
        .unwrap();
    let defaults = vec![NumberFormat::SERIES_VALUE_LABEL_NATIVE_DEFAULT; 2];
    let fixed_two = NumberFormat::new(
        DecimalPlaces::fixed(2).unwrap(),
        NegativeStyle::Parentheses,
        false,
    );
    let customized = vec![fixed_two, NumberFormat::SERIES_VALUE_LABEL_NATIVE_DEFAULT];

    assert_eq!(
        editor
            .body_chart_series_value_label_number_formats(source.drawable_object_id)
            .unwrap(),
        defaults
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_body_chart_series_value_label_number_formats(source.drawable_object_id, &defaults)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);
    editor
        .set_body_chart_series_value_label_number_formats(source.drawable_object_id, &customized)
        .unwrap();
    assert_eq!(
        editor
            .body_chart_series_value_label_number_format(
                source.drawable_object_id,
                Index::from_zero_based(0),
            )
            .unwrap(),
        fixed_two
    );

    let duplicate = editor
        .duplicate_body_chart(
            source.drawable_object_id,
            editor.body_text().unwrap().encode_utf16().count(),
        )
        .unwrap();
    editor
        .set_body_chart_series_value_label_number_format(
            source.drawable_object_id,
            Index::from_zero_based(0),
            NumberFormat::SERIES_VALUE_LABEL_NATIVE_DEFAULT,
        )
        .unwrap();
    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .body_chart_series_value_label_number_formats(source.drawable_object_id)
            .unwrap(),
        defaults
    );
    assert_eq!(
        reopened
            .body_chart_series_value_label_number_formats(duplicate.drawable_object_id)
            .unwrap(),
        customized
    );

    let before_rejected = reopened.to_bytes().unwrap();
    assert!(
        reopened
            .set_body_chart_series_value_label_number_formats(
                source.drawable_object_id,
                &customized[..1],
            )
            .is_err()
    );
    assert!(
        reopened
            .body_chart_series_value_label_number_format(
                source.drawable_object_id,
                Index::from_zero_based(2),
            )
            .is_err()
    );
    assert_eq!(reopened.to_bytes().unwrap(), before_rejected);
    reopened
        .remove_body_chart(source.drawable_object_id)
        .unwrap();
    reopened
        .remove_body_chart(duplicate.drawable_object_id)
        .unwrap();
    assert!(reopened.body_charts().unwrap().is_empty());
}

#[test]
fn scratch_document_supports_native_series_value_label_auto_fit_crud() {
    const BODY_TEXT: &str = "Chart value-label Auto-Fit";

    let mut editor = PagesEditor::create_with_text(BODY_TEXT).unwrap();
    let source = editor
        .add_body_chart(
            BODY_TEXT.encode_utf16().count(),
            Kind::Column2d,
            sample_data(),
            POSITION,
            SIZE,
        )
        .unwrap();
    let defaults = vec![ChartSeriesValueLabelAutoFit::Enabled; 2];
    let customized = vec![
        ChartSeriesValueLabelAutoFit::Disabled,
        ChartSeriesValueLabelAutoFit::Enabled,
    ];

    assert_eq!(
        editor
            .body_chart_series_value_label_auto_fits(source.drawable_object_id)
            .unwrap(),
        defaults
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_body_chart_series_value_label_auto_fits(source.drawable_object_id, &defaults)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);
    editor
        .set_body_chart_series_value_label_auto_fits(source.drawable_object_id, &customized)
        .unwrap();
    assert_eq!(
        editor
            .body_chart_series_value_label_auto_fit(
                source.drawable_object_id,
                Index::from_zero_based(0),
            )
            .unwrap(),
        ChartSeriesValueLabelAutoFit::Disabled
    );

    let duplicate = editor
        .duplicate_body_chart(
            source.drawable_object_id,
            editor.body_text().unwrap().encode_utf16().count(),
        )
        .unwrap();
    editor
        .set_body_chart_series_value_label_auto_fit(
            source.drawable_object_id,
            Index::from_zero_based(0),
            ChartSeriesValueLabelAutoFit::Enabled,
        )
        .unwrap();
    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .body_chart_series_value_label_auto_fits(source.drawable_object_id)
            .unwrap(),
        defaults
    );
    assert_eq!(
        reopened
            .body_chart_series_value_label_auto_fits(duplicate.drawable_object_id)
            .unwrap(),
        customized
    );

    let before_rejected = reopened.to_bytes().unwrap();
    assert!(
        reopened
            .set_body_chart_series_value_label_auto_fits(
                source.drawable_object_id,
                &customized[..1],
            )
            .is_err()
    );
    assert!(
        reopened
            .body_chart_series_value_label_auto_fit(
                source.drawable_object_id,
                Index::from_zero_based(2),
            )
            .is_err()
    );
    assert_eq!(reopened.to_bytes().unwrap(), before_rejected);
    reopened
        .remove_body_chart(source.drawable_object_id)
        .unwrap();
    reopened
        .remove_body_chart(duplicate.drawable_object_id)
        .unwrap();
    assert!(reopened.body_charts().unwrap().is_empty());
}

#[test]
fn scratch_document_supports_native_series_trendline_crud() {
    const BODY_TEXT: &str = "Chart series trendlines";

    let mut editor = PagesEditor::create_with_text(BODY_TEXT).unwrap();
    let source = editor
        .add_body_chart(
            BODY_TEXT.encode_utf16().count(),
            Kind::Column2d,
            sample_data(),
            POSITION,
            SIZE,
        )
        .unwrap();
    let defaults = vec![ChartSeriesTrendline::none(); 2];
    let customized = vec![
        ChartSeriesTrendline::linear()
            .with_legend_name("Revenue fit")
            .unwrap()
            .with_equation_visibility(true)
            .unwrap()
            .with_r_squared_visibility(true)
            .unwrap(),
        ChartSeriesTrendline::moving_average(
            ChartSeriesTrendlineMovingAveragePeriod::new(3).unwrap(),
        )
        .with_legend_visibility(true)
        .unwrap(),
    ];

    assert_eq!(
        editor
            .body_chart_series_trendlines(source.drawable_object_id)
            .unwrap(),
        defaults
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_body_chart_series_trendlines(source.drawable_object_id, &defaults)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);
    editor
        .set_body_chart_series_trendlines(source.drawable_object_id, &customized)
        .unwrap();
    assert_eq!(
        editor
            .body_chart_series_trendline(source.drawable_object_id, Index::from_zero_based(1),)
            .unwrap(),
        customized[1]
    );

    let duplicate = editor
        .duplicate_body_chart(
            source.drawable_object_id,
            editor.body_text().unwrap().encode_utf16().count(),
        )
        .unwrap();
    for series in 0..2 {
        editor
            .set_body_chart_series_trendline(
                source.drawable_object_id,
                Index::from_zero_based(series),
                ChartSeriesTrendline::none(),
            )
            .unwrap();
    }
    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .body_chart_series_trendlines(source.drawable_object_id)
            .unwrap(),
        defaults
    );
    assert_eq!(
        reopened
            .body_chart_series_trendlines(duplicate.drawable_object_id)
            .unwrap(),
        customized
    );

    let before_rejected = reopened.to_bytes().unwrap();
    assert!(
        reopened
            .set_body_chart_series_trendlines(source.drawable_object_id, &customized[..1])
            .is_err()
    );
    assert!(
        reopened
            .body_chart_series_trendline(source.drawable_object_id, Index::from_zero_based(2),)
            .is_err()
    );
    assert!(ChartSeriesTrendline::unsupported(1).is_err());
    assert!(ChartSeriesTrendlinePolynomialOrder::new(7).is_err());
    assert_eq!(reopened.to_bytes().unwrap(), before_rejected);
    reopened
        .remove_body_chart(source.drawable_object_id)
        .unwrap();
    reopened
        .remove_body_chart(duplicate.drawable_object_id)
        .unwrap();
    assert!(reopened.body_charts().unwrap().is_empty());
}

#[test]
fn scratch_document_supports_native_series_error_bar_crud() {
    const BODY_TEXT: &str = "Chart series error bars";

    let mut editor = PagesEditor::create_with_text(BODY_TEXT).unwrap();
    let source = editor
        .add_body_chart(
            BODY_TEXT.encode_utf16().count(),
            Kind::Column2d,
            sample_data(),
            POSITION,
            SIZE,
        )
        .unwrap();
    let defaults = vec![Series::None; 2];
    let customized = vec![
        Series::FixedValue {
            direction: ErrorBarDirection::PositiveAndNegative,
            value: ErrorBarFixedValue::new(12.5).unwrap(),
        },
        Series::Percentage {
            direction: ErrorBarDirection::PositiveOnly,
            percentage: ErrorBarPercentage::new(17).unwrap(),
        },
    ];
    let default_auto_fits = vec![ChartSeriesErrorBarAutoFit::Enabled; 2];
    let customized_auto_fits = vec![
        ChartSeriesErrorBarAutoFit::Disabled,
        ChartSeriesErrorBarAutoFit::Enabled,
    ];

    assert_eq!(
        editor
            .body_chart_series_error_bars(source.drawable_object_id)
            .unwrap(),
        defaults
    );
    assert_eq!(
        editor
            .body_chart_series_error_bar_auto_fits(source.drawable_object_id)
            .unwrap(),
        default_auto_fits
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_body_chart_series_error_bars(source.drawable_object_id, &defaults)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);
    editor
        .set_body_chart_series_error_bars(source.drawable_object_id, &customized)
        .unwrap();
    editor
        .set_body_chart_series_error_bar_auto_fits(source.drawable_object_id, &customized_auto_fits)
        .unwrap();
    assert_eq!(
        editor
            .body_chart_series_error_bar(source.drawable_object_id, Index::from_zero_based(1),)
            .unwrap(),
        customized[1]
    );
    assert_eq!(
        editor
            .body_chart_series_error_bar_auto_fit(
                source.drawable_object_id,
                Index::from_zero_based(0),
            )
            .unwrap(),
        ChartSeriesErrorBarAutoFit::Disabled
    );

    let duplicate = editor
        .duplicate_body_chart(
            source.drawable_object_id,
            editor.body_text().unwrap().encode_utf16().count(),
        )
        .unwrap();
    for series in 0..2 {
        editor
            .set_body_chart_series_error_bar(
                source.drawable_object_id,
                Index::from_zero_based(series),
                Series::None,
            )
            .unwrap();
    }
    editor
        .set_body_chart_series_error_bar_auto_fits(source.drawable_object_id, &default_auto_fits)
        .unwrap();
    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .body_chart_series_error_bars(source.drawable_object_id)
            .unwrap(),
        defaults
    );
    assert_eq!(
        reopened
            .body_chart_series_error_bars(duplicate.drawable_object_id)
            .unwrap(),
        customized
    );
    assert_eq!(
        reopened
            .body_chart_series_error_bar_auto_fits(source.drawable_object_id)
            .unwrap(),
        default_auto_fits
    );
    assert_eq!(
        reopened
            .body_chart_series_error_bar_auto_fits(duplicate.drawable_object_id)
            .unwrap(),
        customized_auto_fits
    );

    let before_rejected = reopened.to_bytes().unwrap();
    assert!(
        reopened
            .set_body_chart_series_error_bars(source.drawable_object_id, &customized[..1])
            .is_err()
    );
    assert!(
        reopened
            .set_body_chart_series_error_bar_auto_fits(
                source.drawable_object_id,
                &customized_auto_fits[..1],
            )
            .is_err()
    );
    assert!(
        reopened
            .body_chart_series_error_bar(source.drawable_object_id, Index::from_zero_based(2),)
            .is_err()
    );
    assert_eq!(reopened.to_bytes().unwrap(), before_rejected);
    reopened
        .remove_body_chart(source.drawable_object_id)
        .unwrap();
    reopened
        .remove_body_chart(duplicate.drawable_object_id)
        .unwrap();
    assert!(reopened.body_charts().unwrap().is_empty());
}
