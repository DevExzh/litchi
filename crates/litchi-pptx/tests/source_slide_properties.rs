#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "integration fixtures intentionally panic on malformed setup"
)]

//! Native slide-root name and visibility edits preserve source package bytes.

use std::collections::BTreeMap;
use std::num::{NonZeroU64, NonZeroUsize};
use std::sync::Arc;

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits, OwnedSource, Position,
    Resource,
};
use litchi_opc::constants::relationship_type;
use litchi_opc::{BlobPart, OpcPackage, PackURI, PackageWriter, SourceLimits, TargetMode};
use litchi_pptx::builder::Builder;
use litchi_pptx::source::{Package as SourcePackage, Patch};
use litchi_pptx::{Package, ShapePropertiesReplacement};
use soapberry_zip::office::ArchiveReader;

const SLIDE: &str = "ppt/slides/slide1.xml";
const SLIDE_RELS: &str = "ppt/slides/_rels/slide1.xml.rels";
const UNUSED_PART: &str = "ppt/media/unused-slide-properties.bin";
const BOUNDS: (i64, i64, i64, i64) = (10, 20, 914_400, 457_200);

fn fixture(slides: usize, shapes: usize) -> Vec<u8> {
    let mut builder = Builder::new().unwrap();
    for slide_index in 0..slides {
        let slide = builder.presentation().unwrap().add_slide().unwrap();
        for shape_index in 0..shapes {
            slide.add_text_box(
                &format!("slide {slide_index} shape {shape_index}"),
                (shape_index as i64) * 100,
                (shape_index as i64) * 200,
                BOUNDS.2,
                BOUNDS.3,
            );
        }
    }
    builder.build().unwrap().to_bytes().unwrap()
}

fn open_source(bytes: &[u8]) -> SourcePackage {
    SourcePackage::from_read_at(Arc::new(OwnedSource::new(bytes.to_vec()))).unwrap()
}

fn members(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
    let archive =
        ArchiveReader::new(bytes, soapberry_zip::office::ArchiveLimits::default()).unwrap();
    archive
        .file_names()
        .map(|name| (name.to_owned(), archive.read(name).unwrap()))
        .collect()
}

fn member(bytes: &[u8], name: &str) -> Vec<u8> {
    ArchiveReader::new(bytes, soapberry_zip::office::ArchiveLimits::default())
        .unwrap()
        .read(name)
        .unwrap()
}

fn slide_xml(bytes: &[u8]) -> String {
    String::from_utf8(member(bytes, SLIDE)).unwrap()
}

fn transported(patch: &Patch) -> Patch {
    Patch::from_bytes(&patch.to_bytes().unwrap()).unwrap()
}

fn with_unused_relationship(bytes: &[u8]) -> (Vec<u8>, Vec<u8>) {
    let payload = b"unused slide-properties resource".to_vec();
    let mut package = OpcPackage::from_bytes(bytes, SourceLimits::default()).unwrap();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new(format!("/{UNUSED_PART}")).unwrap(),
            "application/octet-stream".to_owned(),
            payload.clone(),
        )))
        .unwrap();
    package
        .get_part_mut(&PackURI::new(format!("/{SLIDE}")).unwrap())
        .unwrap()
        .rels_mut()
        .try_add_relationship(
            relationship_type::IMAGE.to_owned(),
            "../media/unused-slide-properties.bin".to_owned(),
            "rIdUnusedSlideProperties".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
    (PackageWriter::to_bytes(&package).unwrap(), payload)
}

fn managed_context(memory: u64) -> (Budget, CancellationSource, ExecutionContext) {
    let budget = Budget::root(
        "pptx-source-slide-properties-test",
        Limits::new(memory, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
    );
    let (cancellation_source, cancellation) = CancellationSource::pair();
    let limits = ExecutionLimits::new(
        NonZeroUsize::new(1).unwrap(),
        NonZeroUsize::new(1).unwrap(),
        NonZeroU64::new(memory.max(1)).unwrap(),
        0,
    )
    .unwrap();
    (
        budget.clone(),
        cancellation_source,
        ExecutionContext::new(budget, cancellation, limits),
    )
}

#[test]
fn eager_and_source_slide_properties_escape_names_and_use_producer_bool_spellings() {
    let before = fixture(1, 1);
    let name = "A<&>\"'\nB\tC\rD";
    let expected_name = "A&lt;&amp;&gt;&quot;&apos;&#10;B&#9;C&#13;D";

    let mut eager = Package::from_reader(std::io::Cursor::new(before.clone())).unwrap();
    let mut eager_edit = eager.edit().unwrap();
    assert!(eager_edit.set_slide_name(0_usize, name).unwrap());
    assert!(eager_edit.set_slide_hidden(0_usize, true).unwrap());
    eager = eager_edit.commit().unwrap().into_package();
    let eager_xml = slide_xml(&eager.to_bytes().unwrap());
    assert!(
        eager_xml.contains(&format!(r#"name="{expected_name}""#)),
        "{eager_xml}"
    );
    assert!(eager_xml.contains(r#"show="0""#));
    let mut eager_visible_edit = eager.edit().unwrap();
    assert!(eager_visible_edit.set_slide_hidden(0_usize, false).unwrap());
    let eager_visible = eager_visible_edit.commit().unwrap().into_package();
    assert!(slide_xml(&eager_visible.to_bytes().unwrap()).contains(r#"show="1""#));

    let source = open_source(&before);
    let mut source_edit = source.edit().unwrap();
    assert!(source_edit.set_slide_name(0_usize, name).unwrap());
    assert!(source_edit.set_slide_hidden(0_usize, true).unwrap());
    let (changed, patch) = source_edit.commit().unwrap().into_parts();
    let changed_bytes = changed.to_bytes().unwrap();
    assert_eq!(slide_xml(&changed_bytes), eager_xml);
    let slide = changed
        .presentation()
        .unwrap()
        .slide(Position::new(0))
        .unwrap();
    assert_eq!(slide.authored_name().unwrap().as_deref(), Some(name));
    assert!(slide.is_hidden().unwrap());

    let durable = transported(&patch);
    assert_eq!(
        source
            .apply(&durable)
            .unwrap()
            .into_package()
            .to_bytes()
            .unwrap(),
        changed_bytes
    );

    let mut visible_edit = changed.edit().unwrap();
    assert!(visible_edit.set_slide_hidden(0_usize, false).unwrap());
    let visible = visible_edit.commit().unwrap().into_package();
    let visible_xml = slide_xml(&visible.to_bytes().unwrap());
    assert!(visible_xml.contains(r#"show="1""#));
    assert!(
        !visible
            .presentation()
            .unwrap()
            .slide(Position::new(0))
            .unwrap()
            .is_hidden()
            .unwrap()
    );
}

#[test]
fn source_and_eager_slide_property_noops_preserve_exact_bytes() {
    let before = fixture(1, 1);

    let mut eager = Package::from_reader(std::io::Cursor::new(before.clone())).unwrap();
    let mut eager_edit = eager.edit().unwrap();
    assert!(!eager_edit.set_slide_hidden(0_usize, false).unwrap());
    eager = eager_edit.commit().unwrap().into_package();
    assert_eq!(eager.to_bytes().unwrap(), before);

    let source = open_source(&before);
    let mut edit = source.edit().unwrap();
    assert!(!edit.set_slide_hidden(0_usize, false).unwrap());
    let commit = edit.commit().unwrap();
    assert!(commit.patch().is_noop());
    assert_eq!(commit.package().to_bytes().unwrap(), before);

    let mut named_edit = source.edit().unwrap();
    named_edit.set_slide_name(0_usize, "Stable").unwrap();
    named_edit.set_slide_hidden(0_usize, true).unwrap();
    let named = named_edit.commit().unwrap().into_package();
    let named_bytes = named.to_bytes().unwrap();
    let mut repeat = named.edit().unwrap();
    assert!(!repeat.set_slide_name(0_usize, "Stable").unwrap());
    assert!(!repeat.set_slide_hidden(0_usize, true).unwrap());
    let repeated = repeat.commit().unwrap();
    assert!(repeated.patch().is_noop());
    assert_eq!(repeated.package().to_bytes().unwrap(), named_bytes);
}

#[test]
fn source_slide_properties_compose_with_shape_geometry_removal_and_reorder() {
    let before = fixture(3, 2);
    let source = open_source(&before);
    let mut edit = source.edit().unwrap();
    assert!(edit.set_slide_name(0_usize, "Edited first").unwrap());
    assert!(edit.set_slide_hidden(0_usize, true).unwrap());
    assert_eq!(
        edit.set_shape_properties(
            0_usize,
            &[ShapePropertiesReplacement::new(Position::new(0))
                .with_text("edited text")
                .with_bounds((100, 200, 700_000, 800_000))],
        )
        .unwrap(),
        1
    );
    assert!(edit.remove_shape(0_usize, 1_usize).unwrap());
    assert!(edit.move_slide(Position::new(0), Position::new(2)).unwrap());
    let (changed, patch) = edit.commit().unwrap().into_parts();
    let changed_bytes = changed.to_bytes().unwrap();

    let presentation = changed.presentation().unwrap();
    assert_eq!(presentation.slide_count().unwrap(), 3);
    assert_eq!(
        presentation
            .slide(Position::new(2))
            .unwrap()
            .authored_name()
            .unwrap()
            .as_deref(),
        Some("Edited first")
    );
    assert!(
        presentation
            .slide(Position::new(2))
            .unwrap()
            .is_hidden()
            .unwrap()
    );
    assert_eq!(
        presentation
            .slide(Position::new(2))
            .unwrap()
            .text()
            .unwrap(),
        "edited text"
    );
    assert_eq!(
        presentation
            .slide(Position::new(2))
            .unwrap()
            .scene()
            .unwrap()
            .roots()
            .count(),
        1
    );

    let durable = transported(&patch);
    let replayed = source.apply(&durable).unwrap().into_package();
    assert_eq!(replayed.to_bytes().unwrap(), changed_bytes);
    let restored = open_source(&changed_bytes)
        .apply(&transported(&durable.inverse()))
        .unwrap()
        .into_package();
    assert_eq!(members(&restored.to_bytes().unwrap()), members(&before));
}

#[test]
fn malformed_slide_name_is_atomic_for_eager_and_source_edits() {
    let before = fixture(1, 1);
    let invalid = "accepted first\u{0}rejected";

    let mut eager = Package::from_reader(std::io::Cursor::new(before.clone())).unwrap();
    let mut eager_edit = eager.edit().unwrap();
    assert!(eager_edit.set_slide_name(0_usize, invalid).is_err());
    eager = eager_edit.commit().unwrap().into_package();
    assert_eq!(eager.to_bytes().unwrap(), before);

    let source = open_source(&before);
    let mut edit = source.edit().unwrap();
    assert!(edit.set_slide_name(0_usize, "kept").unwrap());
    assert!(edit.set_slide_name(0_usize, invalid).is_err());
    let changed = edit.commit().unwrap().into_package();
    assert_eq!(
        changed
            .presentation()
            .unwrap()
            .slide(Position::new(0))
            .unwrap()
            .authored_name()
            .unwrap()
            .as_deref(),
        Some("kept")
    );
    assert_eq!(source.to_bytes().unwrap(), before);
}

#[test]
fn source_slide_properties_preserve_unused_relationship_parts() {
    let (before, payload) = with_unused_relationship(&fixture(1, 1));
    let source = open_source(&before);
    let mut edit = source.edit().unwrap();
    edit.set_slide_name(0_usize, "named").unwrap();
    edit.set_slide_hidden(0_usize, true).unwrap();
    let changed = edit.commit().unwrap().into_package();
    let after = changed.to_bytes().unwrap();

    assert_eq!(member(&after, UNUSED_PART), payload);
    assert_eq!(member(&after, SLIDE_RELS), member(&before, SLIDE_RELS));
    let old = members(&before);
    let new = members(&after);
    for (name, bytes) in old {
        if name != SLIDE {
            assert_eq!(new.get(&name), Some(&bytes), "{name}");
        }
    }
}

#[test]
fn managed_slide_property_admission_and_cancellation_are_failure_atomic() {
    const MEMORY: u64 = 16 * 1024 * 1024;
    let before = fixture(1, 1);
    let (budget, cancellation_source, context) = managed_context(MEMORY);
    let source = SourcePackage::from_read_at_with_options(
        Arc::new(OwnedSource::new(before.clone())),
        litchi_pptx::OpenOptions::default()
            .source_limits(SourceLimits::default())
            .execution_context(context),
    )
    .unwrap();
    let mut edit = source.edit().unwrap();
    let used = budget.used(Resource::Memory);
    let pressure = budget.reserve(Resource::Memory, MEMORY - used).unwrap();
    assert!(
        edit.set_slide_name(0_usize, "admitted after retry")
            .is_err()
    );
    drop(pressure);
    assert!(
        edit.set_slide_name(0_usize, "admitted after retry")
            .unwrap()
    );
    let commit = edit.commit().unwrap();
    assert_eq!(
        commit
            .package()
            .presentation()
            .unwrap()
            .slide(Position::new(0))
            .unwrap()
            .authored_name()
            .unwrap()
            .as_deref(),
        Some("admitted after retry")
    );
    assert_eq!(source.to_bytes().unwrap(), before);

    let mut canceled = commit.package().edit().unwrap();
    canceled.set_slide_hidden(0_usize, true).unwrap();
    cancellation_source.cancel();
    assert!(canceled.commit().is_err());
}

#[test]
fn eager_rename_updates_staged_name_selection_and_ambiguity() {
    let before = fixture(2, 1);
    let source = Package::from_bytes(&before).unwrap();
    let mut edit = source.edit().unwrap();
    edit.set_slide_name(Position::new(0), "new name").unwrap();
    assert!(edit.set_slide_hidden("new name", true).unwrap());
    edit.set_slide_name(Position::new(1), "new name").unwrap();
    assert!(edit.set_slide_hidden("new name", false).is_err());
    let changed = edit.commit().unwrap().into_package();
    let reopened = open_source(&changed.to_bytes().unwrap());
    assert!(
        reopened
            .presentation()
            .unwrap()
            .slide(Position::new(0))
            .unwrap()
            .is_hidden()
            .unwrap()
    );
    assert!(
        !reopened
            .presentation()
            .unwrap()
            .slide(Position::new(1))
            .unwrap()
            .is_hidden()
            .unwrap()
    );
}

#[test]
fn slide_properties_require_expanded_wire_and_reject_unmanaged_attribute_changes() {
    let before = fixture(1, 1);
    let source = open_source(&before);
    let mut edit = source.edit().unwrap();
    edit.set_slide_name(0_usize, "Renamed").unwrap();
    edit.set_slide_hidden(0_usize, true).unwrap();
    let (_, patch) = edit.commit().unwrap().into_parts();
    let wire = patch.to_bytes().unwrap();
    assert!(wire.starts_with(b"LPSX0004LPSX0002"));
    assert!(Patch::from_bytes(&wire[8..]).is_err());
    assert!(
        transported(&patch)
            .inverse()
            .to_bytes()
            .unwrap()
            .starts_with(b"LPSX0004")
    );

    // Change a retained shape name, outside the managed slide attributes.
    let needle = b"name=\"TextBox\"";
    let offset = wire
        .windows(needle.len())
        .rposition(|bytes| bytes == needle)
        .unwrap();
    let mut forged = wire.clone();
    forged[offset + 6] = b'Z';
    assert!(Patch::from_bytes(&forged).is_err());

    // A malformed managed boolean must not disappear during normalization.
    let needle = b"show=\"0\"";
    let offset = wire
        .windows(needle.len())
        .rposition(|bytes| bytes == needle)
        .unwrap();
    let mut forged = wire;
    forged[offset + 6] = b'X';
    assert!(Patch::from_bytes(&forged).is_err());
}

#[test]
fn combined_primary_properties_match_sequential_edits_and_preserve_prior_staging() {
    let before = fixture(1, 2);
    let source = open_source(&before);
    let mut sequential = source.edit().unwrap();
    sequential.set_slide_name(0, "Roadmap\tQ4").unwrap();
    sequential.set_slide_hidden(0, true).unwrap();
    let sequential = sequential.commit().unwrap().into_package();

    let mut combined = source.edit().unwrap();
    assert!(!combined.set_slide_properties(0, None, None).unwrap());
    assert!(
        combined
            .set_slide_properties(0, Some("Roadmap\tQ4"), Some(true))
            .unwrap()
    );
    assert!(
        combined
            .set_slide_properties(0, Some("bad\u{1}"), Some(false))
            .is_err()
    );
    assert!(
        !combined
            .set_slide_properties(0, Some("Roadmap\tQ4"), Some(true))
            .unwrap()
    );
    let (changed, patch) = combined.commit().unwrap().into_parts();
    assert_eq!(changed.to_bytes().unwrap(), sequential.to_bytes().unwrap());
    assert_eq!(
        source
            .apply(&transported(&patch))
            .unwrap()
            .into_package()
            .to_bytes()
            .unwrap(),
        changed.to_bytes().unwrap()
    );
    assert_eq!(
        changed
            .apply(&transported(&patch).inverse())
            .unwrap()
            .into_package()
            .to_bytes()
            .unwrap(),
        before
    );

    let eager = Package::from_reader(std::io::Cursor::new(before)).unwrap();
    let mut edit = eager.edit().unwrap();
    edit.set_slide_properties(0, Some("Roadmap\tQ4"), Some(true))
        .unwrap();
    let eager = edit.commit().unwrap().into_package();
    assert_eq!(eager.to_bytes().unwrap(), changed.to_bytes().unwrap());
}

#[test]
fn inserted_slide_properties_survive_durable_topology_transport() {
    let before = fixture(1, 1);
    let source = open_source(&before);
    let mut edit = source.edit().unwrap();
    edit.insert_slide(Position::new(0), "Blank").unwrap();
    edit.set_slide_properties(0, Some("Inserted\nDraft"), Some(true))
        .unwrap();
    let (changed, patch) = edit.commit().unwrap().into_parts();
    let durable = transported(&patch);
    let bytes = changed.to_bytes().unwrap();
    assert_eq!(
        source
            .apply(&durable)
            .unwrap()
            .into_package()
            .to_bytes()
            .unwrap(),
        bytes
    );
    let reopened = open_source(&bytes);
    assert_eq!(
        reopened
            .presentation()
            .unwrap()
            .slide(Position::new(0))
            .unwrap()
            .authored_name()
            .unwrap()
            .as_deref(),
        Some("Inserted\nDraft")
    );
    assert_eq!(
        reopened
            .apply(&durable.inverse())
            .unwrap()
            .into_package()
            .to_bytes()
            .unwrap(),
        before
    );
}
