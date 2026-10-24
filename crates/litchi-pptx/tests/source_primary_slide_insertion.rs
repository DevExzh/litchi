#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "integration assertions"
)]
//! Source-native insertion composes with immutable reads and durable topology patches.
use litchi_core::{OwnedSource, Position};
use litchi_pptx::{
    Error,
    builder::Builder,
    source::{Package, Patch},
};
use soapberry_zip::office::{ArchiveReader, StreamingArchiveWriter};
use std::collections::BTreeMap;
use std::sync::Arc;

fn bytes(count: usize, strict: bool) -> Vec<u8> {
    let mut builder = Builder::new().unwrap();
    for index in 0..count {
        builder
            .presentation()
            .unwrap()
            .add_slide()
            .unwrap()
            .add_text_box(&format!("original {index}"), 0, 0, 914_400, 914_400);
    }
    let bytes = builder.build().unwrap().to_bytes().unwrap();
    let archive =
        ArchiveReader::new(&bytes, soapberry_zip::office::ArchiveLimits::default()).unwrap();
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
fn texts(package: &Package) -> Vec<String> {
    let p = package.presentation().unwrap();
    (0..p.slide_count().unwrap())
        .map(|i| p.slide(i).unwrap().text().unwrap())
        .collect()
}
fn transported(patch: &Patch) -> Patch {
    Patch::from_bytes(&patch.to_bytes().unwrap()).unwrap()
}

#[test]
fn insertion_positions_readback_and_source_immutability() {
    let before = bytes(2, false);
    let source = open(&before);
    for position in 0..=2 {
        let mut edit = source.edit().unwrap();
        let inserted = edit.insert_slide(Position::new(position), "Blank").unwrap();
        assert_eq!(inserted.text().unwrap(), "");
        assert!(inserted.layout().unwrap().is_some());
        let (after, patch) = edit.commit().unwrap().into_parts();
        let mut expected = vec!["original 0".to_owned(), "original 1".to_owned()];
        expected.insert(position, String::new());
        assert_eq!(texts(&after), expected);
        assert_eq!(texts(&open(&after.to_bytes().unwrap())), expected);
        assert!(after.source_position(Position::new(position)).is_err());
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
            after
                .apply(&transported(&patch.inverse()))
                .unwrap()
                .into_package()
                .to_bytes()
                .unwrap(),
            before
        );
    }
    assert_eq!(source.to_bytes().unwrap(), before);
}
#[test]
fn empty_source_insertion_and_reopened_inverse_restore_exact_members() {
    for strict in [false, true] {
        let before = bytes(0, strict);
        let source = open(&before);
        assert_eq!(source.presentation().unwrap().slide_count().unwrap(), 0);
        let mut edit = source.edit().unwrap();
        edit.insert_slide(Position::new(0), "Blank").unwrap();
        let (after, patch) = edit.commit().unwrap().into_parts();
        let encoded = after.to_bytes().unwrap();
        let reopened = open(&encoded);
        assert_eq!(texts(&reopened), [""]);
        let restored = reopened
            .apply(&transported(&patch.inverse()))
            .unwrap()
            .into_package();
        assert_eq!(restored.presentation().unwrap().slide_count().unwrap(), 0);
        assert_eq!(members(&restored.to_bytes().unwrap()), members(&before));
        assert_eq!(
            source
                .apply(&transported(&patch))
                .unwrap()
                .into_package()
                .to_bytes()
                .unwrap(),
            encoded
        );
        let mut again = restored.edit().unwrap();
        again.insert_slide(Position::new(0), "Blank").unwrap();
        assert_eq!(texts(&again.commit().unwrap().into_package()), [""]);
        assert!(
            after
                .edit()
                .unwrap()
                .remove_slide(Position::new(0))
                .is_err(),
            "direct final-slide policy stays explicit"
        );
        let archive =
            ArchiveReader::new(&encoded, soapberry_zip::office::ArchiveLimits::default()).unwrap();
        let slide = archive
            .file_names()
            .find(|n| n.starts_with("ppt/slides/slide") && n.ends_with(".xml"))
            .unwrap();
        let xml = String::from_utf8(archive.read(slide).unwrap()).unwrap();
        assert!(xml.contains(if strict {
            "http://purl.oclc.org/ooxml/presentationml/main"
        } else {
            "http://schemas.openxmlformats.org/presentationml/2006/main"
        }));
    }
}
#[test]
fn insertion_composes_with_text_order_and_removal() {
    let before = bytes(2, false);
    let source = open(&before);
    let mut edit = source.edit().unwrap();
    edit.insert_slide(Position::new(1), "Blank").unwrap();
    edit.set_shape_text(0, 0usize, "edited original").unwrap();
    edit.move_slide(Position::new(1), Position::new(2)).unwrap();
    edit.remove_slide(Position::new(1)).unwrap();
    edit.insert_slide(Position::new(0), "Blank").unwrap();
    let (after, patch) = edit.commit().unwrap().into_parts();
    assert_eq!(texts(&after), ["", "edited original", ""]);
    let reopened = open(&after.to_bytes().unwrap());
    let restored = reopened
        .apply(&transported(&patch.inverse()))
        .unwrap()
        .into_package();
    assert_eq!(members(&restored.to_bytes().unwrap()), members(&before));
    assert_eq!(texts(&restored), ["original 0", "original 1"]);
}
#[test]
fn failed_insertion_keeps_prior_staged_changes_and_drop_rolls_back() {
    let before = bytes(1, false);
    let source = open(&before);
    let mut edit = source.edit().unwrap();
    edit.insert_slide(Position::new(0), "Blank").unwrap();
    assert!(matches!(
        edit.insert_slide(Position::new(3), "Blank"),
        Err(Error::SlideIndexOutOfBounds { .. })
    ));
    assert!(matches!(
        edit.insert_slide(Position::new(0), "missing layout"),
        Err(Error::SlideLayoutNameNotFound(_))
    ));
    let after = edit.commit().unwrap().into_package();
    assert_eq!(texts(&after), ["", "original 0"]);
    let mut dropped = after.edit().unwrap();
    dropped.insert_slide(Position::new(0), "Blank").unwrap();
    drop(dropped);
    assert_eq!(texts(&after), ["", "original 0"]);
    assert_eq!(source.to_bytes().unwrap(), before);
}
#[test]
fn insert_then_remove_normalizes_to_original_physical_revision() {
    let before = bytes(1, false);
    let source = open(&before);
    let mut edit = source.edit().unwrap();
    edit.insert_slide(Position::new(0), "Blank").unwrap();
    edit.remove_slide(Position::new(0)).unwrap();
    let (after, patch) = edit.commit().unwrap().into_parts();
    assert!(patch.is_noop());
    assert!(after.same_content(&source).unwrap());
    assert_eq!(after.to_bytes().unwrap(), before);
}

#[test]
fn replacing_every_original_identity_has_complete_durable_resource_proofs() {
    let before = bytes(1, false);
    let source = open(&before);
    let mut edit = source.edit().unwrap();
    edit.insert_slide(Position::new(1), "Blank").unwrap();
    edit.remove_slide(Position::new(0)).unwrap();
    let (after, patch) = edit.commit().unwrap().into_parts();
    assert_eq!(texts(&after), [""]);
    let wire = transported(&patch);
    assert_eq!(texts(&source.apply(&wire).unwrap().into_package()), [""]);
    let reopened = open(&after.to_bytes().unwrap());
    let restored = reopened.apply(&wire.inverse()).unwrap().into_package();
    assert_eq!(members(&restored.to_bytes().unwrap()), members(&before));
}

#[test]
fn managed_insertion_retains_patch_resources_and_cancellation_fences_commit() {
    use litchi_core::{
        Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits, Resource,
    };
    use std::num::{NonZeroU64, NonZeroUsize};
    let budget = Budget::root(
        "primary-insertion",
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
        Arc::new(OwnedSource::new(bytes(1, false))),
        litchi_pptx::OpenOptions::default()
            .source_limits(litchi_opc::SourceLimits::default())
            .execution_context(context),
    )
    .unwrap();
    let mut edit = source.edit().unwrap();
    let pressure = budget
        .reserve(
            Resource::Memory,
            32 * 1024 * 1024 - budget.used(Resource::Memory),
        )
        .unwrap();
    assert!(edit.insert_slide(Position::new(0), "Blank").is_err());
    drop(pressure);
    // A refused admission leaves the edit usable and releases its temporary
    // resources. The retry must still create exactly one slide.
    edit.insert_slide(Position::new(0), "Blank").unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(
        commit
            .package()
            .presentation()
            .unwrap()
            .slide_count()
            .unwrap(),
        2
    );
    let patch = commit.patch().clone();
    let inverse = patch.inverse();
    let mut canceled = commit.package().edit().unwrap();
    canceled.insert_slide(Position::new(0), "Blank").unwrap();
    cancel.cancel();
    assert!(canceled.commit().is_err());
    drop(commit);
    drop(source);
    drop(patch);
    assert!(budget.used(Resource::Memory) > 0);
    drop(inverse);
    assert_eq!(budget.used(Resource::Memory), 0);
}

fn rewrite_member(bytes: &[u8], selected: &str, transform: impl Fn(&str) -> String) -> Vec<u8> {
    let archive =
        ArchiveReader::new(bytes, soapberry_zip::office::ArchiveLimits::default()).unwrap();
    let mut writer =
        StreamingArchiveWriter::new(soapberry_zip::office::StreamingArchiveLimits::default());
    for name in archive.file_names() {
        let payload = archive.read(name).unwrap();
        if name == selected {
            let xml = String::from_utf8(payload).unwrap();
            writer
                .write_stored(name, transform(&xml).as_bytes())
                .unwrap();
        } else {
            writer.write_stored(name, &payload).unwrap();
        }
    }
    writer.finish().unwrap().into_inner()
}
#[test]
fn dangling_case_equivalent_targets_reserve_future_slide_names() {
    let before = rewrite_member(&bytes(0, false), "_rels/.rels", |xml| {
        xml.replace("</Relationships>","<Relationship Id=\"reserved-slide\" Type=\"urn:test\" Target=\"ppt/slides/SLIDE1.XML#bookmark\"/></Relationships>")
    });
    let source = open(&before);
    let mut edit = source.edit().unwrap();
    edit.insert_slide(Position::new(0), "Blank").unwrap();
    let (after, patch) = edit.commit().unwrap().into_parts();
    let output = after.to_bytes().unwrap();
    let archive =
        ArchiveReader::new(&output, soapberry_zip::office::ArchiveLimits::default()).unwrap();
    assert!(archive.read("ppt/slides/slide1.xml").is_err());
    assert!(archive.read("ppt/slides/slide2.xml").is_ok());
    let restored = open(&output)
        .apply(&transported(&patch.inverse()))
        .unwrap()
        .into_package();
    assert_eq!(members(&restored.to_bytes().unwrap()), members(&before));
    assert_eq!(
        after
            .apply(&patch.inverse())
            .unwrap()
            .into_package()
            .to_bytes()
            .unwrap(),
        before
    );
}
#[test]
fn duplicate_layout_names_are_an_atomic_typed_refusal() {
    let before = bytes(0, false);
    let package = litchi_pptx::Package::from_reader(std::io::Cursor::new(before.clone())).unwrap();
    let presentation = package.presentation().unwrap();
    let layout = presentation
        .slide_layouts()
        .unwrap()
        .into_iter()
        .find(|l| l.name().unwrap() != "Blank")
        .unwrap();
    let member = layout.part().part().partname().membername().to_owned();
    let name = layout.name().unwrap();
    let before = rewrite_member(&before, &member, |xml| {
        xml.replace(&format!("name=\"{name}\""), "name=\"Blank\"")
    });
    let source = open(&before);
    let mut edit = source.edit().unwrap();
    assert!(matches!(
        edit.insert_slide(Position::new(0), "Blank"),
        Err(Error::AmbiguousSlideLayoutName { matches: 2, .. })
    ));
    assert!(edit.commit().unwrap().patch().is_noop());
    assert_eq!(source.to_bytes().unwrap(), before);
}

#[test]
fn unsupported_root_metadata_is_refused_without_losing_source() {
    for extra in [
        "<p:custShowLst xmlns:rel=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><p:custShow name=\"dangling\" id=\"1\"><p:sldLst><p:sld rel:id=\"rId99\"/></p:sldLst></p:custShow></p:custShowLst>",
        "<p:showPr><p:sldRg st=\"1\" end=\"2\"/></p:showPr>",
        "<p:extLst><p:ext uri=\"unknown\"/></p:extLst>",
    ] {
        let before = rewrite_member(&bytes(0, false), "ppt/presentation.xml", |xml| {
            xml.replace("</p:presentation>", &format!("{extra}</p:presentation>"))
        });
        let source = open(&before);
        let mut edit = source.edit().unwrap();
        assert!(edit.insert_slide(Position::new(0), "Blank").is_err());
        assert!(edit.commit().unwrap().patch().is_noop());
        assert_eq!(source.to_bytes().unwrap(), before);
    }
}

#[test]
fn insertion_inverse_preserves_unrelated_fragments_with_nonempty_target() {
    let before = rewrite_member(&bytes(1, false), "_rels/.rels", |xml| {
        xml.replace("</Relationships>", "<Relationship Id=\"opaque\" Type=\"urn:test\" Target=\"custom/missing.xml#bookmark\"/></Relationships>")
    });
    let source = open(&before);
    let mut edit = source.edit().unwrap();
    edit.insert_slide(Position::new(1), "Blank").unwrap();
    let (after, patch) = edit.commit().unwrap().into_parts();
    let inverse = transported(&patch.inverse());
    assert_eq!(
        after
            .apply(&inverse)
            .unwrap()
            .into_package()
            .to_bytes()
            .unwrap(),
        before
    );
    let restored = open(&after.to_bytes().unwrap())
        .apply(&inverse)
        .unwrap()
        .into_package();
    assert_eq!(members(&restored.to_bytes().unwrap()), members(&before));
}
