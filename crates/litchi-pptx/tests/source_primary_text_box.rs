#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "integration assertions"
)]
//! Source-native text-box creation preserves immutable revisions and durable proofs.
use litchi_core::{OwnedSource, Position};
use litchi_pptx::{
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

const BOUNDS: (i64, i64, i64, i64) = (10, 20, 914_400, 457_200);

#[test]
fn caller_output_limits_precede_managed_memory_admission() {
    use litchi_core::{
        Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits, Resource,
    };
    use litchi_opc::SourceLimits;
    use litchi_pptx::Error;
    use std::num::{NonZeroU64, NonZeroUsize};

    let input = bytes(1, false);
    let ceiling = members(&input).values().map(Vec::len).sum::<usize>() + 8192;
    let oversized = "&".repeat(ceiling);
    for dimension in 0..4 {
        let limits = SourceLimits::builder();
        let limits = match dimension {
            0 => limits.max_part_bytes(ceiling as u64),
            1 => limits.max_entry_bytes(ceiling as u64),
            2 => limits.max_total_part_bytes(ceiling as u64),
            _ => limits.max_aggregate_bytes(ceiling as u64),
        }
        .unwrap()
        .build()
        .unwrap();
        for managed in [false, true] {
            const MEMORY: u64 = 32 * 1024 * 1024;
            let budget = Budget::root(
                "text-box-output-policy",
                Limits::new(MEMORY, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
            );
            let (_cancel, token) = CancellationSource::pair();
            let context = ExecutionContext::new(
                budget.clone(),
                token,
                ExecutionLimits::new(NonZeroUsize::MIN, NonZeroUsize::MIN, NonZeroU64::MAX, 0)
                    .unwrap(),
            );
            let source = Arc::new(OwnedSource::new(input.clone()));
            let owner = if managed {
                Package::from_read_at_with_options(
                    source,
                    litchi_pptx::OpenOptions::default()
                        .source_limits(limits)
                        .execution_context(context),
                )
            } else {
                Package::from_read_at_with_options(
                    source,
                    litchi_pptx::OpenOptions::default().source_limits(limits),
                )
            }
            .unwrap();
            let expected_id = owner
                .edit()
                .unwrap()
                .add_text_box(0, "retry", BOUNDS)
                .unwrap();
            // Capture first so this assertion isolates authoring admission
            // from the selected-slide read's independent memory admission.
            let mut edit = owner.edit().unwrap();
            let used = budget.used(Resource::Memory);
            let pressure = budget.reserve(Resource::Memory, MEMORY - used).unwrap();
            assert!(matches!(
                edit.add_text_box(0, &oversized, BOUNDS),
                Err(Error::Limit {
                    kind: litchi_pptx::opened::LimitKind::Codec(
                        litchi_pptx::opened::CodecLimitKind::OutputBytes
                    ),
                    maximum,
                    ..
                }) if maximum == ceiling
            ));
            assert_eq!(budget.used(Resource::Memory), MEMORY);
            drop(pressure);
            assert_eq!(budget.used(Resource::Memory), used);
            // Failed policy admission consumes neither the one-operation
            // allowance nor a shape identity. A smaller retry remains valid.
            assert_eq!(edit.add_text_box(0, "retry", BOUNDS).unwrap(), expected_id);
            drop(edit.commit().unwrap());
            drop(owner);
            assert_eq!(budget.used(Resource::Memory), 0);
        }
    }
}

#[test]
fn text_box_addition_reads_replays_and_inverts_after_reopening() {
    for strict in [false, true] {
        let before = bytes(2, strict);
        let source = open(&before);
        let mut edit = source.edit().unwrap();
        let first = edit
            .add_text_box(Position::new(0), "A<&\r\nB\tC", BOUNDS)
            .unwrap();
        let second = edit
            .add_text_box(Position::new(0), "second", BOUNDS)
            .unwrap();
        assert_ne!(first, second);
        let (after, patch) = edit.commit().unwrap().into_parts();
        let scene = after
            .presentation()
            .unwrap()
            .slide(0)
            .unwrap()
            .scene()
            .unwrap();
        assert_eq!(scene.roots().count(), 3);
        assert!(texts(&after)[0].contains("A<&\r\nB\tC"));
        assert!(texts(&after)[0].contains("second"));
        assert_eq!(source.to_bytes().unwrap(), before);
        let replayed = source.apply(&transported(&patch)).unwrap().into_package();
        assert_eq!(replayed.to_bytes().unwrap(), after.to_bytes().unwrap());
        let restored = open(&after.to_bytes().unwrap())
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
        let old = members(&before);
        let new = members(&after.to_bytes().unwrap());
        for (name, payload) in old {
            if name != "ppt/slides/slide1.xml" {
                assert_eq!(new.get(&name), Some(&payload), "{name}");
            }
        }
    }
}

#[test]
fn addition_composes_with_text_order_and_slide_topology() {
    let before = bytes(2, false);
    let source = open(&before);
    let mut edit = source.edit().unwrap();
    edit.insert_slide(Position::new(1), "Blank").unwrap();
    let id = edit.add_text_box(Position::new(1), "new", BOUNDS).unwrap();
    edit.set_shape_text(1, format!("Text Box {id}").as_str(), "changed")
        .unwrap();
    edit.add_text_box(Position::new(0), "on physical", BOUNDS)
        .unwrap();
    edit.set_shape_text(0, 0usize, "original changed").unwrap();
    edit.move_slide(Position::new(1), Position::new(2)).unwrap();
    edit.remove_slide(Position::new(1)).unwrap();
    let (after, patch) = edit.commit().unwrap().into_parts();
    assert_eq!(texts(&after).len(), 2);
    assert!(texts(&after)[0].contains("on physical"));
    assert_eq!(texts(&after)[1], "changed");
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
fn failed_addition_keeps_prior_staging_and_drop_is_atomic() {
    let before = bytes(1, false);
    let source = open(&before);
    let mut edit = source.edit().unwrap();
    edit.add_text_box(Position::new(0), "kept", BOUNDS).unwrap();
    for bounds in [(0, 0, 0, 1), (0, 0, 1, -1)] {
        assert!(edit.add_text_box(Position::new(0), "bad", bounds).is_err());
    }
    assert!(edit.add_text_box(Position::new(9), "bad", BOUNDS).is_err());
    assert!(edit.add_text_box(Position::new(0), "\0", BOUNDS).is_err());
    let changed = edit.commit().unwrap().into_package();
    assert_eq!(
        changed
            .presentation()
            .unwrap()
            .slide(0)
            .unwrap()
            .scene()
            .unwrap()
            .roots()
            .count(),
        2
    );
    let mut abandoned = source.edit().unwrap();
    abandoned
        .add_text_box(Position::new(0), "discard", BOUNDS)
        .unwrap();
    drop(abandoned);
    assert_eq!(source.to_bytes().unwrap(), before);
}

#[test]
fn addition_on_a_reopened_revision_has_an_independent_inverse() {
    let source = open(&bytes(1, false));
    let mut edit = source.edit().unwrap();
    edit.add_text_box(Position::new(0), "first", BOUNDS)
        .unwrap();
    let first = edit.commit().unwrap().into_package();
    let before = first.to_bytes().unwrap();
    let reopened = open(&before);
    let mut edit = reopened.edit().unwrap();
    edit.set_shape_text(0, 1usize, "first changed").unwrap();
    edit.add_text_box(Position::new(0), "second", BOUNDS)
        .unwrap();
    let (after, patch) = edit.commit().unwrap().into_parts();
    let restored = open(&after.to_bytes().unwrap())
        .apply(&transported(&patch.inverse()))
        .unwrap()
        .into_package();
    assert_eq!(members(&restored.to_bytes().unwrap()), members(&before));
}

#[test]
fn managed_addition_retains_patch_resources_and_cancellation_fences_commit() {
    use litchi_core::{
        Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits, Resource,
    };
    use std::num::{NonZeroU64, NonZeroUsize};
    let budget = Budget::root(
        "primary-textbox",
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
    assert!(
        edit.add_text_box(Position::new(0), "managed", BOUNDS)
            .is_err()
    );
    drop(pressure);
    // A refused admission leaves the edit usable and releases its temporary
    // resources. The retry must still create exactly one slide.
    edit.add_text_box(Position::new(0), "managed", BOUNDS)
        .unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(
        commit
            .package()
            .presentation()
            .unwrap()
            .slide_count()
            .unwrap(),
        1
    );
    let patch = commit.patch().clone();
    let inverse = patch.inverse();
    let mut canceled = commit.package().edit().unwrap();
    canceled
        .add_text_box(Position::new(0), "canceled", BOUNDS)
        .unwrap();
    cancel.cancel();
    assert!(canceled.commit().is_err());
    drop(commit);
    drop(source);
    drop(patch);
    assert!(budget.used(Resource::Memory) > 0);
    drop(inverse);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn blank_slide_addition_keeps_existing_related_parts() {
    let mut builder = Builder::new().unwrap();
    builder
        .presentation()
        .unwrap()
        .add_slide()
        .unwrap()
        .set_notes("retained notes");
    let before = builder.build().unwrap().to_bytes().unwrap();
    let source = open(&before);
    let mut edit = source.edit().unwrap();
    edit.add_text_box(Position::new(0), "first shape", BOUNDS)
        .unwrap();
    let (after, patch) = edit.commit().unwrap().into_parts();
    assert_eq!(texts(&after), ["first shape"]);
    assert!(
        after
            .presentation()
            .unwrap()
            .slide(0)
            .unwrap()
            .notes()
            .unwrap()
            .is_some()
    );
    let old = members(&before);
    let new = members(&after.to_bytes().unwrap());
    for (name, payload) in old {
        if name != "ppt/slides/slide1.xml" {
            assert_eq!(new.get(&name), Some(&payload), "{name}");
        }
    }
    let restored = open(&after.to_bytes().unwrap())
        .apply(&transported(&patch.inverse()))
        .unwrap()
        .into_package();
    assert_eq!(members(&restored.to_bytes().unwrap()), members(&before));
}

#[test]
fn durable_addition_refuses_forged_shape_surface_changes() {
    let source = open(&bytes(1, false));
    let mut edit = source.edit().unwrap();
    edit.add_text_box(Position::new(0), "new", BOUNDS).unwrap();
    let commit = edit.commit().unwrap();
    let encoded = commit.patch().to_bytes().unwrap();
    assert!(encoded.starts_with(b"LPSX0002"));
    for (old, new) in [
        (b"txBox=\"1\"".as_slice(), b"txBox=\"0\"".as_slice()),
        (b"Text Box ".as_slice(), b"Evil Box ".as_slice()),
        (b"prst=\"rect\"".as_slice(), b"prst=\"line\"".as_slice()),
    ] {
        let mut forged = encoded.clone();
        let start = forged
            .windows(old.len())
            .rposition(|part| part == old)
            .unwrap();
        assert_eq!(old.len(), new.len());
        forged[start..start + old.len()].copy_from_slice(new);
        assert!(
            Patch::from_bytes(&forged).is_err(),
            "forged {}",
            String::from_utf8_lossy(old)
        );
    }
}

#[test]
fn final_shape_identity_is_usable_and_exhaustion_is_atomic() {
    let original = bytes(1, false);
    let archive =
        ArchiveReader::new(&original, soapberry_zip::office::ArchiveLimits::default()).unwrap();
    let mut writer =
        StreamingArchiveWriter::new(soapberry_zip::office::StreamingArchiveLimits::default());
    for name in archive.file_names() {
        let payload = archive.read(name).unwrap();
        if name == "ppt/slides/slide1.xml" {
            let id = litchi_pptx::shape::Scene::read(&payload, litchi_pptx::shape::Limits::DEFAULT)
                .unwrap()
                .roots()
                .next()
                .unwrap()
                .id()
                .unwrap();
            let mut xml = String::from_utf8(payload).unwrap();
            let marker = format!("id=\"{id}\"");
            let start = xml.rfind(&marker).unwrap();
            xml.replace_range(start..start + marker.len(), "id=\"4294967294\"");
            writer.write_stored(name, xml.as_bytes()).unwrap();
        } else {
            writer.write_stored(name, &payload).unwrap();
        }
    }
    let before = writer.finish().unwrap().into_inner();
    let source = open(&before);
    let mut edit = source.edit().unwrap();
    assert_eq!(
        edit.add_text_box(Position::new(0), "last", BOUNDS).unwrap(),
        u32::MAX
    );
    assert!(
        edit.add_text_box(Position::new(0), "exhausted", BOUNDS)
            .is_err()
    );
    let (after, patch) = edit.commit().unwrap().into_parts();
    let inverse = transported(&patch.inverse());
    let restored = open(&after.to_bytes().unwrap())
        .apply(&inverse)
        .unwrap()
        .into_package();
    assert_eq!(members(&restored.to_bytes().unwrap()), members(&before));
}

fn slide_extra(before: &[u8], marker: &str, extra: &str) -> Vec<u8> {
    let archive =
        ArchiveReader::new(before, soapberry_zip::office::ArchiveLimits::default()).unwrap();
    let mut writer =
        StreamingArchiveWriter::new(soapberry_zip::office::StreamingArchiveLimits::default());
    for name in archive.file_names() {
        let payload = archive.read(name).unwrap();
        if name == "ppt/slides/slide1.xml" {
            let xml = String::from_utf8(payload).unwrap();
            writer
                .write_stored(
                    name,
                    xml.replace(marker, &format!("{extra}{marker}")).as_bytes(),
                )
                .unwrap();
        } else {
            writer.write_stored(name, &payload).unwrap();
        }
    }
    writer.finish().unwrap().into_inner()
}

#[test]
fn additions_do_not_activate_dangling_shape_references() {
    for (marker, extra, expected_id) in [
        (
            "</p:sld>",
            r#"<p:timing><p:tnLst><p:par><p:cTn id="1"><p:stCondLst><p:cond evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="3"/></p:tgtEl></p:cond></p:stCondLst></p:cTn></p:par></p:tnLst></p:timing>"#,
            4,
        ),
        (
            "</p:spTree>",
            r#"<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="3" name="Connector"/><p:cNvCxnSpPr><a:stCxn id="4" idx="0"/><a:endCxn id="5" idx="0"/></p:cNvCxnSpPr><p:nvPr/></p:nvCxnSpPr><p:spPr/></p:cxnSp>"#,
            6,
        ),
        (
            "</p:cSld>",
            r#"<p:controls><p:control spid="9"/></p:controls>"#,
            10,
        ),
        (
            "</p:sld>",
            r#"<p:extLst><p:ext uri="event"><p14:playEvt xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" objId="9"/></p:ext></p:extLst>"#,
            10,
        ),
    ] {
        let before = slide_extra(&bytes(1, false), marker, extra);
        let source = open(&before);
        let mut edit = source.edit().unwrap();
        assert_eq!(
            edit.add_text_box(Position::new(0), "safe identity", BOUNDS)
                .unwrap(),
            expected_id
        );
        let (after, patch) = edit.commit().unwrap().into_parts();
        let restored = open(&after.to_bytes().unwrap())
            .apply(&transported(&patch.inverse()))
            .unwrap()
            .into_package();
        assert_eq!(members(&restored.to_bytes().unwrap()), members(&before));
    }
}

#[test]
fn addition_refuses_mce_branch_selection_before_staging() {
    let extra = r#"<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"><mc:Choice Requires="p14"><p:sp><p:nvSpPr><p:cNvPr id="100" name="Choice"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/></p:sp></mc:Choice><mc:Fallback><p:sp><p:nvSpPr><p:cNvPr id="101" name="Fallback"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/></p:sp></mc:Fallback></mc:AlternateContent>"#;
    let before = slide_extra(&bytes(1, false), "</p:spTree>", extra);
    let source = open(&before);
    let mut edit = source.edit().unwrap();
    assert!(
        edit.add_text_box(Position::new(0), "must refuse", BOUNDS)
            .is_err()
    );
    assert!(edit.commit().unwrap().patch().is_noop());
    assert_eq!(source.to_bytes().unwrap(), before);
}
