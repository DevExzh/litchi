#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "integration assertions"
)]

//! Public regression coverage for chunked shape edits.

use litchi_core::{OwnedSource, Position};
use litchi_pptx::{
    Error, MAX_SHAPE_PROPERTIES_REPLACEMENTS, MAX_SHAPE_TEXT_REPLACEMENTS, Package,
    ShapePropertiesReplacement, ShapeTextReplacement, SourceBackedPresentation,
    builder::Builder,
    opened::{Limits, Patch as OpenedPatch},
    source::{Package as SourcePackage, Patch as SourcePatch},
};
use soapberry_zip::office::{ArchiveReader, StreamingArchiveWriter};
use std::sync::Arc;

const CHUNK_SIZE: usize = 256;

fn bytes(shape_count: usize) -> Vec<u8> {
    let mut builder = Builder::new().unwrap();
    let slide = builder.presentation().unwrap().add_slide().unwrap();
    for index in 0..shape_count {
        slide.add_text_box(
            &format!("original-{index}"),
            (index as i64) * 10,
            (index as i64) * 20,
            914_400,
            457_200,
        );
    }
    builder.build().unwrap().to_bytes().unwrap()
}

fn open_source(bytes: &[u8]) -> SourcePackage {
    SourcePackage::from_read_at(Arc::new(OwnedSource::new(bytes.to_vec()))).unwrap()
}

fn text_values(package: &Package) -> Vec<String> {
    package
        .presentation()
        .unwrap()
        .slide(0)
        .unwrap()
        .unwrap()
        .shapes()
        .unwrap()
        .iter()
        .map(|shape| shape.text().unwrap_or_default().to_owned())
        .collect()
}

fn source_values(package: &SourcePackage) -> Vec<String> {
    package
        .presentation()
        .unwrap()
        .slide(Position::new(0))
        .unwrap()
        .scene()
        .unwrap()
        .iter()
        .map(|shape| shape.view().text().unwrap_or_default().to_owned())
        .collect()
}

fn eager_properties(package: &Package) -> Vec<(String, (i64, i64, i64, i64))> {
    package
        .presentation()
        .unwrap()
        .slide(0)
        .unwrap()
        .unwrap()
        .shapes()
        .unwrap()
        .iter()
        .map(|shape| {
            let bounds = shape.bounds().unwrap();
            (
                shape.text().unwrap_or_default().to_owned(),
                (bounds.x(), bounds.y(), bounds.width(), bounds.height()),
            )
        })
        .collect()
}

fn source_properties(package: &SourcePackage) -> Vec<(String, (i64, i64, i64, i64))> {
    package
        .presentation()
        .unwrap()
        .slide(Position::new(0))
        .unwrap()
        .scene()
        .unwrap()
        .iter()
        .map(|shape| {
            let bounds = shape.view().bounds().unwrap();
            (
                shape.view().text().unwrap_or_default().to_owned(),
                (bounds.x(), bounds.y(), bounds.width(), bounds.height()),
            )
        })
        .collect()
}

fn text_replacements<'a>(texts: &'a [String]) -> Vec<ShapeTextReplacement<'a>> {
    texts
        .iter()
        .enumerate()
        .map(|(index, text)| ShapeTextReplacement::at(index, text))
        .collect()
}

fn property_replacements<'a>(texts: &'a [String]) -> Vec<ShapePropertiesReplacement<'a>> {
    texts
        .iter()
        .enumerate()
        .map(|(index, text)| {
            ShapePropertiesReplacement::new(Position::new(index))
                .with_text(text)
                .with_bounds((
                    (index as i64) * 3 + 1,
                    (index as i64) * 5 + 2,
                    700_000 + index as i64,
                    300_000 + index as i64,
                ))
        })
        .collect()
}

fn transported_opened(patch: &OpenedPatch) -> OpenedPatch {
    OpenedPatch::from_bytes(&patch.to_bytes().unwrap(), Limits::default()).unwrap()
}

fn transported_source(patch: &SourcePatch) -> SourcePatch {
    SourcePatch::from_bytes(&patch.to_bytes().unwrap()).unwrap()
}

fn malformed_second_rotation(source: &[u8]) -> Vec<u8> {
    let archive =
        ArchiveReader::new(source, soapberry_zip::office::ArchiveLimits::default()).unwrap();
    let mut writer =
        StreamingArchiveWriter::new(soapberry_zip::office::StreamingArchiveLimits::default());
    for name in archive.file_names() {
        let payload = archive.read(name).unwrap();
        if name == "ppt/slides/slide1.xml" {
            let xml = String::from_utf8(payload).unwrap();
            let needle = "<a:xfrm>";
            let first = xml.find(needle).unwrap();
            let second_start =
                first + needle.len() + xml[first + needle.len()..].find(needle).unwrap();
            let xml = format!(
                "{}<a:xfrm rot=\"malformed\">{}",
                &xml[..second_start],
                &xml[second_start + needle.len()..]
            );
            writer.write_stored(name, xml.as_bytes()).unwrap();
        } else {
            writer.write_stored(name, &payload).unwrap();
        }
    }
    writer.finish().unwrap().into_inner()
}

#[test]
fn eager_text_chunks_match_bounded_sequential_oracle_and_replay() {
    let before = bytes(1_000);
    let texts: Vec<_> = (0..1_000).map(|index| format!("chunked-{index}")).collect();
    let replacements = text_replacements(&texts);
    let chunks: Vec<_> = replacements.chunks(CHUNK_SIZE).collect();

    let source = Package::from_bytes(&before).unwrap();
    let mut edit = source.edit().unwrap();
    assert_eq!(edit.set_shape_text_chunks(0_usize, &chunks).unwrap(), 1_000);
    let commit = edit.commit().unwrap();
    let changed = commit.clone().into_package();
    let changed_bytes = changed.to_bytes().unwrap();

    let mut oracle = Package::from_bytes(&before).unwrap();
    let mut sequential = oracle.edit().unwrap();
    for chunk in replacements.chunks(CHUNK_SIZE) {
        assert_eq!(
            sequential.set_shape_texts(0_usize, chunk).unwrap(),
            chunk.len()
        );
    }
    oracle = sequential.commit().unwrap().into_package();
    assert_eq!(changed_bytes, oracle.to_bytes().unwrap());
    assert_eq!(source.to_bytes().unwrap(), before);
    assert_eq!(text_values(&changed)[0], "chunked-0");
    assert_eq!(text_values(&changed)[999], "chunked-999");

    let patch = transported_opened(commit.patch());
    let replayed = Package::from_bytes(&before)
        .unwrap()
        .apply(&patch)
        .unwrap()
        .into_package();
    assert_eq!(replayed.to_bytes().unwrap(), changed_bytes);
    let restored = Package::from_bytes(&changed_bytes)
        .unwrap()
        .apply(&patch.inverse())
        .unwrap()
        .into_package();
    assert_eq!(restored.to_bytes().unwrap(), before);
}

#[test]
fn source_primary_text_chunks_match_bounded_sequential_oracle_and_replay() {
    let before = bytes(1_000);
    let texts: Vec<_> = (0..1_000)
        .map(|index| format!("source-chunked-{index}"))
        .collect();
    let replacements = text_replacements(&texts);
    let chunks: Vec<_> = replacements.chunks(CHUNK_SIZE).collect();

    let source = open_source(&before);
    let mut edit = source.edit().unwrap();
    assert_eq!(
        edit.set_shape_text_chunks(Position::new(0), &chunks)
            .unwrap(),
        1_000
    );
    let (changed, patch) = edit.commit().unwrap().into_parts();
    let changed_bytes = changed.to_bytes().unwrap();

    let mut oracle = open_source(&before);
    let mut sequential = oracle.edit().unwrap();
    for chunk in replacements.chunks(CHUNK_SIZE) {
        assert_eq!(
            sequential.set_shape_texts(Position::new(0), chunk).unwrap(),
            chunk.len()
        );
    }
    oracle = sequential.commit().unwrap().into_package();
    assert_eq!(changed_bytes, oracle.to_bytes().unwrap());
    assert_eq!(source.to_bytes().unwrap(), before);
    assert_eq!(source_values(&changed)[0], "source-chunked-0");
    assert_eq!(source_values(&changed)[999], "source-chunked-999");

    let patch = transported_source(&patch);
    let replayed = open_source(&before).apply(&patch).unwrap().into_package();
    assert_eq!(replayed.to_bytes().unwrap(), changed_bytes);
    let restored = open_source(&changed_bytes)
        .apply(&patch.inverse())
        .unwrap()
        .into_package();
    assert_eq!(restored.to_bytes().unwrap(), before);
}

#[test]
fn focused_source_chunks_publish_and_keep_the_patch_reversible() {
    let before = bytes(3);
    let texts: Vec<_> = (0..3).map(|index| format!("focused-{index}")).collect();
    let replacements = text_replacements(&texts);
    let chunks: Vec<_> = replacements.chunks(CHUNK_SIZE).collect();
    let editor = SourcePackage::from_read_at(Arc::new(OwnedSource::new(before.clone()))).unwrap();
    let mut edit = editor.edit().unwrap();
    assert_eq!(
        edit.set_shape_text_chunks(Position::new(0), &chunks)
            .unwrap(),
        3
    );
    let commit = edit.commit().unwrap();
    assert!(commit.diagnostics().changed());
    assert!(commit.package().apply(&commit.patch().inverse()).is_ok());
    let mut output = Vec::new();
    commit.package().write_to(&mut output).unwrap();
    assert_ne!(output, before);
    let reopened =
        SourceBackedPresentation::from_read_at(Arc::new(OwnedSource::new(output))).unwrap();
    let text = reopened.slide(0).unwrap().text().unwrap();
    assert!(text.contains("focused-0"));
    assert!(text.contains("focused-2"));

    let editor = SourcePackage::from_read_at(Arc::new(OwnedSource::new(before))).unwrap();
    let mut edit = editor.edit().unwrap();
    let properties = [
        ShapePropertiesReplacement::new(Position::new(0))
            .with_text("focused-property-0")
            .with_bounds((11, 22, 333, 444)),
        ShapePropertiesReplacement::new(Position::new(1)).with_bounds((55, 66, 777, 888)),
    ];
    assert_eq!(
        edit.set_shape_property_chunks(Position::new(0), &[&properties])
            .unwrap(),
        2
    );
    let commit = edit.commit().unwrap();
    assert!(commit.diagnostics().changed());
    let mut output = Vec::new();
    commit.package().write_to(&mut output).unwrap();
    let reopened =
        SourceBackedPresentation::from_read_at(Arc::new(OwnedSource::new(output))).unwrap();
    assert!(
        reopened
            .slide(0)
            .unwrap()
            .text()
            .unwrap()
            .contains("focused-property-0")
    );
}

#[test]
fn eager_and_source_property_chunks_mix_text_and_bounds_like_sequential_batches() {
    let before = bytes(1_000);
    let texts: Vec<_> = (0..1_000).map(|index| format!("mixed-{index}")).collect();
    let replacements = property_replacements(&texts);
    let chunks: Vec<_> = replacements.chunks(CHUNK_SIZE).collect();

    let source = Package::from_bytes(&before).unwrap();
    let mut edit = source.edit().unwrap();
    assert_eq!(
        edit.set_shape_property_chunks(0_usize, &chunks).unwrap(),
        1_000
    );
    let commit = edit.commit().unwrap();
    let changed = commit.clone().into_package();
    let changed_bytes = changed.to_bytes().unwrap();
    let mut oracle = Package::from_bytes(&before).unwrap();
    let mut sequential = oracle.edit().unwrap();
    for chunk in replacements.chunks(CHUNK_SIZE) {
        assert_eq!(
            sequential.set_shape_properties(0_usize, chunk).unwrap(),
            chunk.len()
        );
    }
    oracle = sequential.commit().unwrap().into_package();
    assert_eq!(changed_bytes, oracle.to_bytes().unwrap());
    assert_eq!(source.to_bytes().unwrap(), before);
    assert_eq!(eager_properties(&changed)[0].0, "mixed-0");
    assert_eq!(eager_properties(&changed)[999].0, "mixed-999");
    assert_eq!(eager_properties(&changed)[0].1, (1, 2, 700_000, 300_000));

    let source = open_source(&before);
    let mut edit = source.edit().unwrap();
    assert_eq!(
        edit.set_shape_property_chunks(Position::new(0), &chunks)
            .unwrap(),
        1_000
    );
    let (changed, patch) = edit.commit().unwrap().into_parts();
    let mut source_oracle = open_source(&before);
    let mut source_sequential = source_oracle.edit().unwrap();
    for chunk in replacements.chunks(CHUNK_SIZE) {
        assert_eq!(
            source_sequential
                .set_shape_properties(Position::new(0), chunk)
                .unwrap(),
            chunk.len()
        );
    }
    source_oracle = source_sequential.commit().unwrap().into_package();
    let source_changed_bytes = changed.to_bytes().unwrap();
    assert_eq!(source_changed_bytes, source_oracle.to_bytes().unwrap());
    assert_eq!(source.to_bytes().unwrap(), before);
    assert_eq!(source_properties(&changed)[999].0, "mixed-999");
    assert_eq!(
        source_properties(&changed)[999].1,
        (2_998, 4_997, 700_999, 300_999)
    );
    assert_eq!(
        open_source(&before)
            .apply(&transported_source(&patch))
            .unwrap()
            .into_package()
            .to_bytes()
            .unwrap(),
        source_changed_bytes
    );
}

#[test]
fn property_chunks_count_only_items_that_change() {
    let before = bytes(3);
    let replacements = [
        ShapePropertiesReplacement::new(Position::new(0))
            .with_text("original-0")
            .with_bounds((0, 0, 914_400, 457_200)),
        ShapePropertiesReplacement::new(Position::new(1)).with_text("changed-1"),
        ShapePropertiesReplacement::new(Position::new(2)).with_bounds((33, 44, 555, 666)),
    ];
    let chunks: &[&[ShapePropertiesReplacement<'_>]] = &[&replacements[..1], &replacements[1..]];

    let mut eager = Package::from_bytes(&before).unwrap().edit().unwrap();
    assert_eq!(eager.set_shape_property_chunks(0_usize, chunks).unwrap(), 2);
    let eager = eager.commit().unwrap().into_package();
    assert_eq!(eager_properties(&eager)[0].0, "original-0");
    assert_eq!(eager_properties(&eager)[1].0, "changed-1");

    let source = open_source(&before);
    let mut edit = source.edit().unwrap();
    assert_eq!(
        edit.set_shape_property_chunks(Position::new(0), chunks)
            .unwrap(),
        2
    );
    let source = edit.commit().unwrap().into_package();
    assert_eq!(source_properties(&source)[0].0, "original-0");
    assert_eq!(source_properties(&source)[1].0, "changed-1");
}

#[test]
fn ten_thousand_text_chunks_change_every_shape_once_without_changing_input() {
    let before = bytes(10_000);
    let texts: Vec<_> = (0..10_000).map(|index| format!("ten-k-{index}")).collect();
    let replacements = text_replacements(&texts);
    let chunks: Vec<_> = replacements.chunks(CHUNK_SIZE).collect();

    let source = Package::from_bytes(&before).unwrap();
    let mut edit = source.edit().unwrap();
    assert_eq!(
        edit.set_shape_text_chunks(0_usize, &chunks).unwrap(),
        10_000
    );
    let changed = edit.commit().unwrap().into_package();
    assert_eq!(source.to_bytes().unwrap(), before);
    let values = text_values(&changed);
    assert_eq!(values.len(), 10_000);
    assert_eq!(values[0], "ten-k-0");
    assert_eq!(values[9_999], "ten-k-9999");

    let source = open_source(&before);
    let mut edit = source.edit().unwrap();
    assert_eq!(
        edit.set_shape_text_chunks(Position::new(0), &chunks)
            .unwrap(),
        10_000
    );
    let changed = edit.commit().unwrap().into_package();
    assert_eq!(source.to_bytes().unwrap(), before);
    let values = source_values(&changed);
    assert_eq!(values.len(), 10_000);
    assert_eq!(values[0], "ten-k-0");
    assert_eq!(values[9_999], "ten-k-9999");
}

#[test]
fn all_equal_chunks_are_exact_noops_for_eager_source_and_focused_edits() {
    let before = bytes(2);
    let equal_text = [
        ShapeTextReplacement::at(0, "original-0"),
        ShapeTextReplacement::at(1, "original-1"),
    ];
    let text_chunks: &[&[ShapeTextReplacement<'_>]] = &[&equal_text[..1], &equal_text[1..]];

    let mut eager = Package::from_bytes(&before).unwrap().edit().unwrap();
    assert_eq!(
        eager.set_shape_text_chunks(0_usize, text_chunks).unwrap(),
        0
    );
    let commit = eager.commit().unwrap();
    assert!(commit.patch().is_noop());
    assert_eq!(commit.into_package().to_bytes().unwrap(), before);

    let source = open_source(&before);
    let mut edit = source.edit().unwrap();
    assert_eq!(
        edit.set_shape_text_chunks(Position::new(0), text_chunks)
            .unwrap(),
        0
    );
    let (changed, patch) = edit.commit().unwrap().into_parts();
    assert!(patch.is_noop());
    assert_eq!(changed.to_bytes().unwrap(), before);
    assert_eq!(source.to_bytes().unwrap(), before);

    let equal_properties = [
        ShapePropertiesReplacement::new(Position::new(0))
            .with_text("original-0")
            .with_bounds((0, 0, 914_400, 457_200)),
        ShapePropertiesReplacement::new(Position::new(1))
            .with_text("original-1")
            .with_bounds((10, 20, 914_400, 457_200)),
    ];
    let property_chunks: &[&[ShapePropertiesReplacement<'_>]] =
        &[&equal_properties[..1], &equal_properties[1..]];
    let source = open_source(&before);
    let mut edit = source.edit().unwrap();
    assert_eq!(
        edit.set_shape_property_chunks(Position::new(0), property_chunks)
            .unwrap(),
        0
    );
    let (changed, patch) = edit.commit().unwrap().into_parts();
    assert!(patch.is_noop());
    assert_eq!(changed.to_bytes().unwrap(), before);

    let editor = SourcePackage::from_read_at(Arc::new(OwnedSource::new(before.clone()))).unwrap();
    let mut edit = editor.edit().unwrap();
    assert_eq!(
        edit.set_shape_text_chunks(Position::new(0), text_chunks)
            .unwrap(),
        0
    );
    let commit = edit.commit().unwrap();
    assert!(!commit.diagnostics().changed());
    let mut output = Vec::new();
    commit.package().write_to(&mut output).unwrap();
    assert_eq!(output, before);
}

#[test]
fn empty_chunks_do_not_consume_eager_or_primary_source_edits() {
    let before = bytes(2);
    let empty_text: &[ShapeTextReplacement<'_>] = &[];
    let empty_properties: &[ShapePropertiesReplacement<'_>] = &[];

    let mut eager = Package::from_bytes(&before).unwrap().edit().unwrap();
    assert_eq!(
        eager.set_shape_text_chunks(0_usize, &[empty_text]).unwrap(),
        0
    );
    let text = [ShapeTextReplacement::at(0, "after-empty")];
    assert_eq!(eager.set_shape_text_chunks(0_usize, &[&text]).unwrap(), 1);

    let mut eager = Package::from_bytes(&before).unwrap().edit().unwrap();
    assert_eq!(
        eager
            .set_shape_property_chunks(0_usize, &[empty_properties])
            .unwrap(),
        0
    );
    let properties = [ShapePropertiesReplacement::new(Position::new(0)).with_text("after-empty")];
    assert_eq!(
        eager
            .set_shape_property_chunks(0_usize, &[&properties])
            .unwrap(),
        1
    );

    let mut source = open_source(&before).edit().unwrap();
    assert_eq!(
        source
            .set_shape_text_chunks(Position::new(0), &[empty_text])
            .unwrap(),
        0
    );
    assert_eq!(
        source
            .set_shape_text_chunks(Position::new(0), &[&text])
            .unwrap(),
        1
    );

    let mut source = open_source(&before).edit().unwrap();
    assert_eq!(
        source
            .set_shape_property_chunks(Position::new(0), &[empty_properties])
            .unwrap(),
        0
    );
    assert_eq!(
        source
            .set_shape_property_chunks(Position::new(0), &[&properties])
            .unwrap(),
        1
    );
}

#[test]
fn late_invalid_chunks_and_cross_chunk_duplicates_are_atomic() {
    let before = bytes(3);

    let valid = [ShapeTextReplacement::at(0, "valid-first")];
    let invalid = [ShapeTextReplacement::at(1, "bad\0text")];
    let mut edit = Package::from_bytes(&before).unwrap().edit().unwrap();
    assert!(
        edit.set_shape_text_chunks(0_usize, &[&valid, &invalid])
            .is_err()
    );
    assert_eq!(
        edit.commit().unwrap().into_package().to_bytes().unwrap(),
        before
    );

    let duplicate_left = [ShapeTextReplacement::at(0, "left")];
    let duplicate_right = [ShapeTextReplacement::at(0, "right")];
    let mut edit = Package::from_bytes(&before).unwrap().edit().unwrap();
    assert!(matches!(
        edit.set_shape_text_chunks(0_usize, &[&duplicate_left, &duplicate_right]),
        Err(Error::Refused(refusal))
            if refusal.kind() == litchi_core::edit::RefusalKind::InvalidValue
                && refusal.operation() == "set_shape_text"
    ));
    assert_eq!(
        edit.commit().unwrap().into_package().to_bytes().unwrap(),
        before
    );

    let duplicate_left = [ShapePropertiesReplacement::new(Position::new(0)).with_text("left")];
    let duplicate_right = [ShapePropertiesReplacement::new(Position::new(0)).with_text("right")];
    let mut edit = Package::from_bytes(&before).unwrap().edit().unwrap();
    assert!(matches!(
        edit.set_shape_property_chunks(0_usize, &[&duplicate_left, &duplicate_right]),
        Err(Error::Refused(refusal))
            if refusal.kind() == litchi_core::edit::RefusalKind::InvalidValue
                && refusal.operation() == "set_shape_properties"
    ));
    assert_eq!(
        edit.commit().unwrap().into_package().to_bytes().unwrap(),
        before
    );

    let valid = [ShapePropertiesReplacement::new(Position::new(0))
        .with_text("valid-property")
        .with_bounds((1, 2, 3, 4))];
    let invalid = [ShapePropertiesReplacement::new(Position::new(1)).with_bounds((5, 6, 0, 8))];
    let mut edit = Package::from_bytes(&before).unwrap().edit().unwrap();
    assert!(
        edit.set_shape_property_chunks(0_usize, &[&valid, &invalid])
            .is_err()
    );
    assert_eq!(
        edit.commit().unwrap().into_package().to_bytes().unwrap(),
        before
    );

    let source = open_source(&before);
    let mut edit = source.edit().unwrap();
    let source_valid = valid_text_for_source();
    let source_invalid = invalid_text_for_source();
    assert!(
        edit.set_shape_text_chunks(Position::new(0), &[&source_valid, &source_invalid])
            .is_err()
    );
    assert_eq!(source.to_bytes().unwrap(), before);
}

fn valid_text_for_source() -> [ShapeTextReplacement<'static>; 1] {
    [ShapeTextReplacement::at(0, "source-valid")]
}

fn invalid_text_for_source() -> [ShapeTextReplacement<'static>; 1] {
    [ShapeTextReplacement::at(1, "source\0invalid")]
}

#[test]
fn malformed_unselected_rotation_is_preserved_by_a_late_property_chunk() {
    let before = malformed_second_rotation(&bytes(2));
    let mut edit = Package::from_bytes(&before).unwrap().edit().unwrap();
    let selected = [ShapePropertiesReplacement::new(Position::new(0)).with_text("selected")];
    assert_eq!(
        edit.set_shape_property_chunks(0_usize, &[&selected])
            .unwrap(),
        1
    );
    let after = edit.commit().unwrap().into_package().to_bytes().unwrap();
    let archive =
        ArchiveReader::new(&after, soapberry_zip::office::ArchiveLimits::default()).unwrap();
    let slide = String::from_utf8(archive.read("ppt/slides/slide1.xml").unwrap()).unwrap();
    assert!(slide.contains("rot=\"malformed\""));
    assert!(slide.contains("selected"));
}

#[test]
fn one_oversized_chunk_keeps_the_historical_256_cap() {
    let before = bytes(MAX_SHAPE_TEXT_REPLACEMENTS + 1);
    let texts: Vec<_> = (0..MAX_SHAPE_TEXT_REPLACEMENTS + 1)
        .map(|index| format!("over-limit-{index}"))
        .collect();
    let replacements = text_replacements(&texts);
    let mut edit = Package::from_bytes(&before).unwrap().edit().unwrap();
    assert!(matches!(
        edit.set_shape_text_chunks(0_usize, &[&replacements]),
        Err(Error::Limit {
            kind: litchi_pptx::opened::LimitKind::Codec(
                litchi_pptx::opened::CodecLimitKind::TextReplacements,
            ),
            observed,
            maximum,
            ..
        }) if observed == MAX_SHAPE_TEXT_REPLACEMENTS + 1
            && maximum == MAX_SHAPE_TEXT_REPLACEMENTS
    ));
    assert_eq!(
        edit.commit().unwrap().into_package().to_bytes().unwrap(),
        before
    );

    let properties = property_replacements(&texts);
    let mut edit = Package::from_bytes(&before).unwrap().edit().unwrap();
    assert!(matches!(
        edit.set_shape_property_chunks(0_usize, &[&properties]),
        Err(Error::Limit {
            kind: litchi_pptx::opened::LimitKind::Codec(
                litchi_pptx::opened::CodecLimitKind::ShapeProperties,
            ),
            observed,
            maximum,
            ..
        }) if observed == MAX_SHAPE_PROPERTIES_REPLACEMENTS + 1
            && maximum == MAX_SHAPE_PROPERTIES_REPLACEMENTS
    ));
    assert_eq!(
        edit.commit().unwrap().into_package().to_bytes().unwrap(),
        before
    );
}

#[test]
fn chunk_local_text_limits_allow_a_global_budget_below_scene_output_limit() {
    let before = bytes(2);
    let piece = Limits::default().max_text_bytes() / 2 + 1;
    let first_text = "a".repeat(piece);
    let second_text = "b".repeat(piece);
    let replacements = [
        ShapeTextReplacement::at(0, &first_text),
        ShapeTextReplacement::at(1, &second_text),
    ];

    let mut native = Package::from_bytes(&before).unwrap().edit().unwrap();
    assert!(matches!(
        native.set_shape_texts(0_usize, &replacements),
        Err(Error::Limit { .. })
    ));
    assert_eq!(
        native.commit().unwrap().into_package().to_bytes().unwrap(),
        before
    );

    let source = Package::from_bytes(&before).unwrap();
    let mut edit = source.edit().unwrap();
    assert_eq!(
        edit.set_shape_text_chunks(0_usize, &[&replacements[..1], &replacements[1..]])
            .unwrap(),
        2
    );
    let changed = edit.commit().unwrap().into_package();
    let values = text_values(&changed);
    assert_eq!(values[0].len(), piece);
    assert_eq!(values[1].len(), piece);
    assert_eq!(source.to_bytes().unwrap(), before);
}

#[test]
fn an_individual_chunk_over_the_text_cap_is_refused_without_mutation() {
    let before = bytes(1);
    let oversized = "x".repeat(Limits::default().max_text_bytes() + 1);
    let replacement = [ShapeTextReplacement::at(0, &oversized)];
    let mut edit = Package::from_bytes(&before).unwrap().edit().unwrap();
    assert!(matches!(
        edit.set_shape_text_chunks(0_usize, &[&replacement]),
        Err(Error::Limit { .. })
    ));
    assert_eq!(
        edit.commit().unwrap().into_package().to_bytes().unwrap(),
        before
    );
}

#[test]
fn property_chunks_split_the_legacy_text_budget_for_eager_and_source_primary() {
    let before = bytes(2);
    let piece = Limits::default().max_text_bytes() / 2 + 1;
    let first_text = "property-a".repeat(piece / 10 + 1);
    let second_text = "property-b".repeat(piece / 10 + 1);
    let first_text = &first_text[..piece];
    let second_text = &second_text[..piece];
    let replacements = [
        ShapePropertiesReplacement::new(Position::new(0)).with_text(first_text),
        ShapePropertiesReplacement::new(Position::new(1)).with_text(second_text),
    ];

    let mut native = Package::from_bytes(&before).unwrap().edit().unwrap();
    assert!(matches!(
        native.set_shape_properties(0_usize, &replacements),
        Err(Error::Limit { .. })
    ));
    assert_eq!(
        native.commit().unwrap().into_package().to_bytes().unwrap(),
        before
    );

    let source = Package::from_bytes(&before).unwrap();
    let mut edit = source.edit().unwrap();
    assert_eq!(
        edit.set_shape_property_chunks(0_usize, &[&replacements[..1], &replacements[1..]])
            .unwrap(),
        2
    );
    let changed = edit.commit().unwrap().into_package();
    let values = text_values(&changed);
    assert_eq!(values[0].len(), piece);
    assert_eq!(values[1].len(), piece);
    assert_eq!(source.to_bytes().unwrap(), before);

    let mut native = open_source(&before).edit().unwrap();
    assert!(matches!(
        native.set_shape_properties(Position::new(0), &replacements),
        Err(Error::Limit { .. })
    ));
    let source = open_source(&before);
    let mut edit = source.edit().unwrap();
    assert_eq!(
        edit
            .set_shape_property_chunks(
                Position::new(0),
                &[&replacements[..1], &replacements[1..]],
            )
            .unwrap(),
        2
    );
    let changed = edit.commit().unwrap().into_package();
    let values = source_values(&changed);
    assert_eq!(values[0].len(), piece);
    assert_eq!(values[1].len(), piece);
    assert_eq!(source.to_bytes().unwrap(), before);
}

#[test]
fn focused_source_rejects_too_many_empty_chunks_before_the_noop_shortcut() {
    let before = bytes(1);
    let editor = SourcePackage::from_read_at(Arc::new(OwnedSource::new(before))).unwrap();
    let mut edit = editor.edit().unwrap();
    let empty: &[ShapeTextReplacement<'_>] = &[];
    let chunks = vec![empty; litchi_pptx::shape::Limits::DEFAULT.shapes() + 1];
    assert!(matches!(
        edit.set_shape_text_chunks(Position::new(0), &chunks),
        Err(Error::Limit { .. })
    ));
    assert_eq!(
        edit.set_shape_text_chunks(Position::new(0), &[empty])
            .unwrap(),
        0
    );
}

#[test]
fn source_chunk_admission_refuses_before_flatten_and_retains_the_commit_reservation() {
    use litchi_core::{
        Budget, CancellationSource, ExecutionContext, ExecutionError, ExecutionLimits,
        Limits as CoreLimits, Resource,
    };
    use litchi_opc::{OpcError, SourceLimits};
    use std::num::{NonZeroU64, NonZeroUsize};

    let before = bytes(2);
    let budget = Budget::root(
        "pptx-shape-chunk-managed-test",
        CoreLimits::new(
            32 * 1024 * 1024,
            u64::MAX,
            u64::MAX,
            u64::MAX,
            u64::MAX,
            u64::MAX,
        ),
    );
    let (cancel, token) = CancellationSource::pair();
    let context = ExecutionContext::new(
        budget.clone(),
        token,
        ExecutionLimits::new(NonZeroUsize::MIN, NonZeroUsize::MIN, NonZeroU64::MAX, 0).unwrap(),
    );
    let source = SourcePackage::from_read_at_with_options(
        Arc::new(OwnedSource::new(before.clone())),
        litchi_pptx::OpenOptions::default()
            .source_limits(SourceLimits::default())
            .execution_context(context),
    )
    .unwrap();
    let mut edit = source.edit().unwrap();
    let replacements = [
        ShapeTextReplacement::at(0, "managed-first"),
        ShapeTextReplacement::at(1, "managed-second"),
    ];
    let pressure = budget
        .reserve(
            Resource::Memory,
            32 * 1024 * 1024 - budget.used(Resource::Memory),
        )
        .unwrap();
    assert!(matches!(
        edit.set_shape_text_chunks(Position::new(0), &[&replacements]),
        Err(Error::Opc(OpcError::Execution(
            ExecutionError::ResourceLimit(limit)
        ))) if limit.resource == Resource::Memory
    ));
    drop(pressure);

    assert_eq!(
        edit.set_shape_text_chunks(Position::new(0), &[&replacements])
            .unwrap(),
        2
    );
    let commit = edit.commit().unwrap();
    let patch = commit.patch().clone();
    cancel.cancel();
    drop(commit);
    drop(source);
    assert!(budget.used(Resource::Memory) > 0);
    drop(patch);
    assert_eq!(budget.used(Resource::Memory), 0);
}
