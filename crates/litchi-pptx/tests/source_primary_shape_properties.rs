#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "integration assertions"
)]
//! Source-native shape-property batches preserve exact package revisions.

use litchi_core::{OwnedSource, Position};
use litchi_pptx::{
    ShapePropertiesReplacement,
    builder::Builder,
    source::{Package, Patch},
};
use soapberry_zip::office::{ArchiveReader, StreamingArchiveWriter};
use std::collections::BTreeMap;
use std::sync::Arc;

const BOUNDS: (i64, i64, i64, i64) = (10, 20, 914_400, 457_200);

fn bytes(shape_counts: &[usize], strict: bool) -> Vec<u8> {
    let mut builder = Builder::new().unwrap();
    for (slide_index, shape_count) in shape_counts.iter().copied().enumerate() {
        let slide = builder.presentation().unwrap().add_slide().unwrap();
        for shape_index in 0..shape_count {
            slide.add_text_box(
                &format!("original {slide_index}-{shape_index}"),
                (shape_index as i64) * 100,
                (shape_index as i64) * 200,
                914_400,
                914_400,
            );
        }
    }
    let source = builder.build().unwrap().to_bytes().unwrap();
    let archive =
        ArchiveReader::new(&source, soapberry_zip::office::ArchiveLimits::default()).unwrap();
    let mut writer =
        StreamingArchiveWriter::new(soapberry_zip::office::StreamingArchiveLimits::default());
    for name in archive.file_names() {
        let mut xml = String::from_utf8(archive.read(name).unwrap()).unwrap();
        if strict {
            for (old, new) in [
                (
                    "http://schemas.openxmlformats.org/presentationml/2006/main",
                    "http://purl.oclc.org/ooxml/presentationml/main",
                ),
                (
                    "http://schemas.openxmlformats.org/drawingml/2006/main",
                    "http://purl.oclc.org/ooxml/drawingml/main",
                ),
                (
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                    "http://purl.oclc.org/ooxml/officeDocument/relationships",
                ),
            ] {
                xml = xml.replace(old, new);
            }
        }
        if [
            "[Content_Types].xml",
            "ppt/presentation.xml",
            "ppt/_rels/presentation.xml.rels",
        ]
        .contains(&name)
        {
            let at = xml.rfind("</").unwrap();
            xml.insert_str(at, "<!--source insertion sentinel-->");
        }
        writer.write_stored(name, xml.as_bytes()).unwrap();
    }
    writer.finish().unwrap().into_inner()
}

fn bytes_with_unknown_attribute(strict: bool) -> Vec<u8> {
    let source = bytes(&[2, 1], strict);
    let archive =
        ArchiveReader::new(&source, soapberry_zip::office::ArchiveLimits::default()).unwrap();
    let mut writer =
        StreamingArchiveWriter::new(soapberry_zip::office::StreamingArchiveLimits::default());
    for name in archive.file_names() {
        let mut xml = String::from_utf8(archive.read(name).unwrap()).unwrap();
        if name == "ppt/slides/slide1.xml" {
            assert!(xml.contains("<p:cNvPr id="));
            xml = xml.replacen("<p:cNvPr id=", "<p:cNvPr future=\"keep-me\" id=", 1);
            xml = xml.replacen("</p:sld>", "<!--slide opaque sentinel--></p:sld>", 1);
        }
        writer.write_stored(name, xml.as_bytes()).unwrap();
    }
    writer.finish().unwrap().into_inner()
}

fn open(bytes: &[u8]) -> Package {
    Package::from_read_at(Arc::new(OwnedSource::new(bytes.to_vec()))).unwrap()
}

fn members(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
    let archive =
        ArchiveReader::new(bytes, soapberry_zip::office::ArchiveLimits::default()).unwrap();
    archive
        .file_names()
        .map(|name| (name.to_owned(), archive.read(name).unwrap()))
        .collect()
}

fn values(package: &Package, slide: usize) -> Vec<(Option<String>, Option<(i64, i64, i64, i64)>)> {
    let scene = package
        .presentation()
        .unwrap()
        .slide(Position::new(slide))
        .unwrap()
        .scene()
        .unwrap();
    scene
        .iter()
        .map(|handle| {
            let text = handle.view().text().map(ToOwned::to_owned);
            let bounds = handle
                .view()
                .bounds()
                .map(|value| (value.x(), value.y(), value.width(), value.height()));
            (text, bounds)
        })
        .collect()
}

fn transported(patch: &Patch) -> Patch {
    Patch::from_bytes(&patch.to_bytes().unwrap()).unwrap()
}

#[test]
fn malformed_unselected_transform_stays_opaque_during_geometry_and_addition() {
    let input = bytes(&[2], false);
    let archive =
        ArchiveReader::new(&input, soapberry_zip::office::ArchiveLimits::default()).unwrap();
    let mut writer =
        StreamingArchiveWriter::new(soapberry_zip::office::StreamingArchiveLimits::default());
    for name in archive.file_names() {
        let payload = archive.read(name).unwrap();
        if name == "ppt/slides/slide1.xml" {
            let xml = String::from_utf8(payload).unwrap();
            assert!(xml.contains("<a:xfrm>"));
            let xml = xml.replacen("<a:xfrm>", "<a:xfrm rot=\"private-invalid\">", 1);
            writer.write_stored(name, xml.as_bytes()).unwrap();
        } else {
            writer.write_stored(name, &payload).unwrap();
        }
    }
    let before = writer.finish().unwrap().into_inner();
    let source = open(&before);
    let mut edit = source.edit().unwrap();
    let bad = [ShapePropertiesReplacement::new(Position::new(0))
        .with_text("must not publish")
        .with_bounds((1, 2, 3, 4))];
    assert!(edit.set_shape_properties(Position::new(0), &bad).is_err());
    edit.set_shape_text(Position::new(0), 0usize, "allowed text")
        .unwrap();
    let valid = [ShapePropertiesReplacement::new(Position::new(1)).with_bounds((50, 60, 70, 80))];
    edit.set_shape_properties(Position::new(0), &valid).unwrap();
    edit.add_text_box(Position::new(0), "appended", (10, 20, 30, 40))
        .unwrap();
    let (after, patch) = edit.commit().unwrap().into_parts();
    assert_eq!(source.to_bytes().unwrap(), before);
    let after_bytes = after.to_bytes().unwrap();
    assert!(
        String::from_utf8(members(&after_bytes)["ppt/slides/slide1.xml"].clone())
            .unwrap()
            .contains("rot=\"private-invalid\"")
    );
    assert_eq!(
        source
            .apply(&transported(&patch))
            .unwrap()
            .into_package()
            .to_bytes()
            .unwrap(),
        after_bytes
    );
    let restored = open(&after_bytes)
        .apply(&transported(&patch.inverse()))
        .unwrap()
        .into_package();
    assert_eq!(members(&restored.to_bytes().unwrap()), members(&before));
}

#[test]
fn source_shape_properties_compose_text_and_bounds_for_both_conformances() {
    for strict in [false, true] {
        let before = bytes(&[2, 1], strict);
        let source = open(&before);
        let mut edit = source.edit().unwrap();
        let replacements = [
            ShapePropertiesReplacement::new(Position::new(0))
                .with_text("first changed<&\r\n")
                .with_bounds((100, 200, 700, 800)),
            ShapePropertiesReplacement::new(Position::new(1)).with_bounds((300, 400, 900, 1_000)),
        ];
        assert_eq!(
            edit.set_shape_properties(Position::new(0), &replacements)
                .unwrap(),
            2
        );
        let (after, patch) = edit.commit().unwrap().into_parts();
        assert_eq!(
            values(&after, 0)[0].0.as_deref(),
            Some("first changed<&\r\n")
        );
        assert_eq!(values(&after, 0)[0].1, Some((100, 200, 700, 800)));
        assert_eq!(values(&after, 0)[1].1, Some((300, 400, 900, 1_000)));
        assert_eq!(source.to_bytes().unwrap(), before);

        let replayed = source.apply(&transported(&patch)).unwrap().into_package();
        assert_eq!(replayed.to_bytes().unwrap(), after.to_bytes().unwrap());
        assert!(patch.to_bytes().unwrap().starts_with(b"LPSX0002"));

        let reopened = open(&after.to_bytes().unwrap());
        let restored = reopened
            .apply(&transported(&patch.inverse()))
            .unwrap()
            .into_package();
        assert_eq!(members(&restored.to_bytes().unwrap()), members(&before));
        for (name, payload) in members(&before) {
            if name != "ppt/slides/slide1.xml" {
                assert_eq!(
                    members(&after.to_bytes().unwrap()).get(&name),
                    Some(&payload),
                    "{name}"
                );
            }
        }
    }
}

#[test]
fn source_shape_property_noop_is_byte_exact() {
    let before = bytes(&[2], false);
    let source = open(&before);
    let current = values(&source, 0);
    let replacements = [
        ShapePropertiesReplacement::new(Position::new(0))
            .with_text(current[0].0.as_deref().unwrap())
            .with_bounds(current[0].1.unwrap()),
        ShapePropertiesReplacement::new(Position::new(1)).with_bounds(current[1].1.unwrap()),
    ];
    let mut edit = source.edit().unwrap();
    assert_eq!(
        edit.set_shape_properties(Position::new(0), &replacements)
            .unwrap(),
        0
    );
    let commit = edit.commit().unwrap();
    assert!(commit.patch().is_noop());
    assert_eq!(commit.package().to_bytes().unwrap(), before);
}

#[test]
fn appended_text_box_then_geometry_stays_in_the_v2_edit_surface() {
    let before = bytes(&[1], false);
    let source = open(&before);
    let mut edit = source.edit().unwrap();
    edit.add_text_box(Position::new(0), "appended", BOUNDS)
        .unwrap();
    let replacement = [ShapePropertiesReplacement::new(Position::new(1))
        .with_text("appended changed")
        .with_bounds((100, 200, 700, 800))];
    assert_eq!(
        edit.set_shape_properties(Position::new(0), &replacement)
            .unwrap(),
        1
    );
    let (after, patch) = edit.commit().unwrap().into_parts();
    assert!(patch.to_bytes().unwrap().starts_with(b"LPSX0002"));
    assert_eq!(
        values(&after, 0),
        vec![
            (
                Some("original 0-0".to_owned()),
                Some((0, 0, 914_400, 914_400))
            ),
            (
                Some("appended changed".to_owned()),
                Some((100, 200, 700, 800))
            ),
        ]
    );
    assert_eq!(
        source
            .apply(&transported(&patch))
            .unwrap()
            .into_package()
            .to_bytes()
            .unwrap(),
        after.to_bytes().unwrap()
    );
    let restored = open(&after.to_bytes().unwrap())
        .apply(&transported(&patch.inverse()))
        .unwrap()
        .into_package();
    assert_eq!(members(&restored.to_bytes().unwrap()), members(&before));
}

#[test]
fn invalid_shape_property_batches_leave_prior_staging_untouched() {
    let before = bytes(&[2], false);
    let source = open(&before);
    let original = values(&source, 0);
    let mut edit = source.edit().unwrap();
    assert!(
        edit.set_shape_text(Position::new(0), 0usize, "prior staged text")
            .unwrap()
    );

    let duplicate = [
        ShapePropertiesReplacement::new(Position::new(0)).with_text("discarded"),
        ShapePropertiesReplacement::new(Position::new(0)).with_bounds((1, 2, 3, 4)),
    ];
    assert!(
        edit.set_shape_properties(Position::new(0), &duplicate)
            .is_err()
    );

    let invalid_width = [ShapePropertiesReplacement::new(Position::new(1))
        .with_text("discarded")
        .with_bounds((1, 2, 0, 4))];
    assert!(
        edit.set_shape_properties(Position::new(0), &invalid_width)
            .is_err()
    );
    let changed = edit.commit().unwrap().into_package();
    assert_eq!(
        values(&changed, 0)[0].0.as_deref(),
        Some("prior staged text")
    );
    assert_eq!(values(&changed, 0)[0].1, original[0].1);
    assert_eq!(values(&changed, 0)[1], original[1]);
}

#[test]
fn unknown_slide_bytes_stay_exact_and_forged_outside_geometry_is_rejected() {
    let before = bytes_with_unknown_attribute(false);
    let source = open(&before);
    let mut edit = source.edit().unwrap();
    let replacement =
        [ShapePropertiesReplacement::new(Position::new(0)).with_bounds((100, 200, 700, 800))];
    let commit = {
        assert_eq!(
            edit.set_shape_properties(Position::new(0), &replacement)
                .unwrap(),
            1
        );
        edit.commit().unwrap()
    };
    let after = commit.package().to_bytes().unwrap();
    let mut after_members = members(&after);
    let after_slide = after_members.remove("ppt/slides/slide1.xml").unwrap();
    assert!(
        after_slide
            .windows(b"future=\"keep-me\"".len())
            .any(|part| { part == b"future=\"keep-me\"" })
    );
    assert!(
        after_slide
            .windows(b"slide opaque sentinel".len())
            .any(|part| part == b"slide opaque sentinel")
    );
    for (name, payload) in members(&before) {
        if name != "ppt/slides/slide1.xml" {
            assert_eq!(after_members.get(&name), Some(&payload), "{name}");
        }
    }

    let encoded = commit.patch().to_bytes().unwrap();
    let old = b"future=\"keep-me\"";
    let new = b"future=\"forged!\"";
    assert_eq!(old.len(), new.len());
    let start = encoded
        .windows(old.len())
        .rposition(|part| part == old)
        .unwrap();
    let mut forged = encoded;
    forged[start..start + old.len()].copy_from_slice(new);
    assert!(Patch::from_bytes(&forged).is_err());
}

#[test]
fn inserted_slide_text_box_and_properties_replay_as_one_v3_revision() {
    let before = bytes(&[1, 1], false);
    let source = open(&before);
    let mut edit = source.edit().unwrap();
    edit.insert_slide(Position::new(1), "Blank").unwrap();
    edit.add_text_box(Position::new(1), "inserted", BOUNDS)
        .unwrap();
    let replacement = [ShapePropertiesReplacement::new(Position::new(0))
        .with_text("inserted changed")
        .with_bounds((100, 200, 800, 900))];
    assert_eq!(
        edit.set_shape_properties(Position::new(1), &replacement)
            .unwrap(),
        1
    );
    let (after, patch) = edit.commit().unwrap().into_parts();
    let wire = patch.to_bytes().unwrap();
    assert!(wire.starts_with(b"LPSX0003"));
    assert_eq!(
        source
            .apply(&transported(&patch))
            .unwrap()
            .into_package()
            .to_bytes()
            .unwrap(),
        after.to_bytes().unwrap()
    );
    assert_eq!(
        values(&after, 1),
        vec![(
            Some("inserted changed".to_owned()),
            Some((100, 200, 800, 900))
        )]
    );
    let restored = open(&after.to_bytes().unwrap())
        .apply(&transported(&patch.inverse()))
        .unwrap()
        .into_package();
    assert_eq!(members(&restored.to_bytes().unwrap()), members(&before));
}

#[test]
fn managed_shape_property_revision_retains_resources_and_checks_cancellation() {
    use litchi_core::{
        Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits, Resource,
    };
    use litchi_opc::SourceLimits;
    use std::num::{NonZeroU64, NonZeroUsize};

    let budget = Budget::root(
        "primary-shape-properties",
        Limits::new(
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
    let source = Package::from_read_at_with_options(
        Arc::new(OwnedSource::new(bytes(&[2], false))),
        litchi_pptx::OpenOptions::default()
            .source_limits(SourceLimits::default())
            .execution_context(context),
    )
    .unwrap();
    let mut edit = source.edit().unwrap();
    let replacement =
        [ShapePropertiesReplacement::new(Position::new(0)).with_bounds((100, 200, 700, 800))];
    edit.set_shape_properties(Position::new(0), &replacement)
        .unwrap();
    let commit = edit.commit().unwrap();
    let patch = commit.patch().clone();
    let inverse = patch.inverse();
    let package = commit.into_package();

    let mut canceled = package.edit().unwrap();
    let second = [ShapePropertiesReplacement::new(Position::new(1)).with_text("canceled")];
    canceled
        .set_shape_properties(Position::new(0), &second)
        .unwrap();
    cancel.cancel();
    assert!(canceled.commit().is_err());

    drop(package);
    drop(source);
    drop(patch);
    assert!(budget.used(Resource::Memory) > 0);
    drop(inverse);
    assert_eq!(budget.used(Resource::Memory), 0);
}
