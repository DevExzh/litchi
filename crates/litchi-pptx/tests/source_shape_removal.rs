#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "the in-memory fixtures intentionally panic on malformed test setup"
)]

//! Shape removal across the eager and source-backed primary APIs.

use std::collections::BTreeMap;
use std::num::{NonZeroU64, NonZeroUsize};
use std::sync::Arc;

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits as CoreLimits,
    OwnedSource, Position, Resource,
};
use litchi_opc::constants::relationship_type as rt;
use litchi_opc::{BlobPart, OpcPackage, PackURI, PackageWriter, SourceLimits, TargetMode};
use litchi_pptx::Package;
use litchi_pptx::builder::Builder;
use litchi_pptx::shape::Scene;
use litchi_pptx::source::{Package as SourcePackage, Patch as SourcePatch};
use soapberry_zip::office::{ArchiveReader, StreamingArchiveWriter};

const SLIDE: &str = "ppt/slides/slide1.xml";
const SLIDE_RELS: &str = "ppt/slides/_rels/slide1.xml.rels";
const UNUSED_MEDIA: &str = "ppt/media/unused-shape-removal.bin";
const P14: &str = "http://schemas.microsoft.com/office/powerpoint/2010/main";
const SHOW_EVENTS_URI: &str = "{E180D4A7-C9FB-4DFB-919C-405C955672EB}";

fn fixture(slide_count: usize, shape_count: usize) -> Vec<u8> {
    let mut builder = Builder::new().unwrap();
    for slide_index in 0..slide_count {
        let slide = builder.presentation().unwrap().add_slide().unwrap();
        for shape_index in 0..shape_count {
            slide.add_text_box(
                &format!("slide {slide_index} shape {shape_index}"),
                (shape_index as i64) * 10,
                (shape_index as i64) * 20,
                914_400,
                457_200,
            );
        }
    }
    builder.build().unwrap().to_bytes().unwrap()
}

fn open_source(bytes: &[u8]) -> SourcePackage {
    SourcePackage::from_read_at(Arc::new(OwnedSource::new(bytes.to_vec()))).unwrap()
}

fn managed_context(memory: u64) -> (Budget, CancellationSource, ExecutionContext) {
    let budget = Budget::root(
        "pptx-source-shape-removal-test",
        CoreLimits::new(memory, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
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

fn rewrite_slide(bytes: &[u8], transform: impl FnOnce(Vec<u8>) -> Vec<u8>) -> Vec<u8> {
    let archive =
        ArchiveReader::new(bytes, soapberry_zip::office::ArchiveLimits::default()).unwrap();
    let names = archive.file_names().map(str::to_owned).collect::<Vec<_>>();
    let mut transform = Some(transform);
    let mut writer =
        StreamingArchiveWriter::new(soapberry_zip::office::StreamingArchiveLimits::default());
    for name in names {
        let payload = archive.read(&name).unwrap();
        if name == SLIDE {
            writer
                .write_stored(&name, &transform.take().unwrap()(payload))
                .unwrap();
        } else {
            writer.write_stored(&name, &payload).unwrap();
        }
    }
    writer.finish().unwrap().into_inner()
}

fn inject(bytes: &[u8], marker: &str, extra: &str) -> Vec<u8> {
    rewrite_slide(bytes, |payload| {
        let xml = String::from_utf8(payload).unwrap();
        assert_eq!(
            xml.matches(marker).count(),
            1,
            "fixture marker must be unique"
        );
        xml.replacen(marker, &format!("{extra}{marker}"), 1)
            .into_bytes()
    })
}

fn shape_ids(bytes: &[u8]) -> Vec<u32> {
    let xml = member(bytes, SLIDE);
    Scene::read(&xml, litchi_pptx::shape::Limits::DEFAULT)
        .unwrap()
        .iter()
        .map(|shape| shape.id().unwrap())
        .collect()
}

fn eager_texts(package: &Package) -> Vec<String> {
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

fn source_texts(package: &SourcePackage) -> Vec<String> {
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

fn transported(patch: &SourcePatch) -> SourcePatch {
    SourcePatch::from_bytes(&patch.to_bytes().unwrap()).unwrap()
}

fn replace_same_width(mut bytes: Vec<u8>, from: &[u8], to: &[u8]) -> Vec<u8> {
    assert_eq!(from.len(), to.len());
    let mut cursor = 0;
    let mut replaced = false;
    while let Some(offset) = bytes[cursor..]
        .windows(from.len())
        .position(|window| window == from)
    {
        let start = cursor + offset;
        bytes[start..start + from.len()].copy_from_slice(to);
        cursor = start + to.len();
        replaced = true;
    }
    assert!(replaced, "test payload did not contain the expected bytes");
    bytes
}

fn with_unused_relationship(bytes: &[u8]) -> (Vec<u8>, Vec<u8>) {
    let payload = b"unused relationship resource".to_vec();
    let mut package = OpcPackage::from_bytes(bytes, SourceLimits::default()).unwrap();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new(format!("/{UNUSED_MEDIA}")).unwrap(),
            "application/octet-stream".to_owned(),
            payload.clone(),
        )))
        .unwrap();
    package
        .get_part_mut(&PackURI::new(format!("/{SLIDE}")).unwrap())
        .unwrap()
        .rels_mut()
        .try_add_relationship(
            rt::IMAGE.to_owned(),
            "../media/unused-shape-removal.bin".to_owned(),
            "rIdUnusedShapeRemoval".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
    (PackageWriter::to_bytes(&package).unwrap(), payload)
}

fn with_grouped_first_shape(bytes: &[u8]) -> Vec<u8> {
    rewrite_slide(bytes, |payload| {
        let scene = Scene::read(&payload, litchi_pptx::shape::Limits::DEFAULT).unwrap();
        let shape = scene.roots().next().unwrap();
        let id = shape.id().unwrap();
        let span = shape.common().span().unwrap();
        let start = usize::try_from(span.start()).unwrap();
        let end = usize::try_from(span.end().unwrap()).unwrap();
        let original = std::str::from_utf8(&payload[start..end]).unwrap();
        let group = format!(
            r#"<p:grpSp><p:nvGrpSpPr><p:cNvPr id="{}" name="Removal group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/><a:chOff x="0" y="0"/><a:chExt cx="914400" cy="914400"/></a:xfrm></p:grpSpPr>{original}</p:grpSp>"#,
            id.saturating_add(100)
        );
        let mut replacement = payload;
        replacement.splice(start..end, group.as_bytes().iter().copied());
        replacement
    })
}

fn with_grouped_first_shape_and_internal_connector(bytes: &[u8]) -> Vec<u8> {
    rewrite_slide(bytes, |payload| {
        let scene = Scene::read(&payload, litchi_pptx::shape::Limits::DEFAULT).unwrap();
        let shape = scene.roots().next().unwrap();
        let id = shape.id().unwrap();
        let span = shape.common().span().unwrap();
        let start = usize::try_from(span.start()).unwrap();
        let end = usize::try_from(span.end().unwrap()).unwrap();
        let original = std::str::from_utf8(&payload[start..end]).unwrap();
        let group_id = id.saturating_add(100);
        let connector_id = id.saturating_add(101);
        let group = format!(
            r#"<p:grpSp><p:nvGrpSpPr><p:cNvPr id="{group_id}" name="Removal group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/><a:chOff x="0" y="0"/><a:chExt cx="914400" cy="914400"/></a:xfrm></p:grpSpPr>{original}<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="{connector_id}" name="Internal connector"/><p:cNvCxnSpPr><a:stCxn id="{id}" idx="0"/><a:endCxn id="{id}" idx="0"/></p:cNvCxnSpPr><p:nvPr/></p:nvCxnSpPr><p:spPr><a:prstGeom prst="line"><a:avLst/></a:prstGeom></p:spPr></p:cxnSp></p:grpSp>"#
        );
        let mut replacement = payload;
        replacement.splice(start..end, group.as_bytes().iter().copied());
        replacement
    })
}

fn assert_source_remove_refused(bytes: &[u8]) {
    let source = open_source(bytes);
    let original = source.to_bytes().unwrap();
    let result = source.edit().and_then(|mut edit| {
        edit.remove_shape(0, 0_usize).map(|changed| {
            assert!(changed);
        })
    });
    assert!(result.is_err());
    assert_eq!(source.to_bytes().unwrap(), original);
}

#[test]
fn eager_and_source_remove_one_plain_top_level_shape() {
    let before = fixture(1, 2);

    let mut eager = Package::from_reader(std::io::Cursor::new(before.clone())).unwrap();
    let mut eager_edit = eager.edit().unwrap();
    assert!(eager_edit.remove_shape(0, 0_usize).unwrap());
    eager = eager_edit.commit().unwrap().into_package();
    assert_eq!(eager_texts(&eager).len(), 1);
    assert!(eager_texts(&eager)[0].contains("shape 1"));

    let source = open_source(&before);
    let mut source_edit = source.edit().unwrap();
    assert!(source_edit.remove_shape(0, 0_usize).unwrap());
    let source_after = source_edit.commit().unwrap().into_package();
    assert_eq!(source_texts(&source_after), ["slide 0 shape 1"]);
}

#[test]
fn source_shape_removal_patch_transports_inverts_and_reopens() {
    let before = fixture(1, 2);
    let source = open_source(&before);
    let mut edit = source.edit().unwrap();
    assert!(edit.remove_shape(0, 0_usize).unwrap());
    let (after, patch) = edit.commit().unwrap().into_parts();
    let patch = transported(&patch);
    let after_bytes = after.to_bytes().unwrap();

    assert_eq!(
        source
            .apply(&patch)
            .unwrap()
            .into_package()
            .to_bytes()
            .unwrap(),
        after_bytes
    );
    let inverse = transported(&patch.inverse());
    let restored = open_source(&after_bytes)
        .apply(&inverse)
        .unwrap()
        .into_package();
    assert_eq!(members(&restored.to_bytes().unwrap()), members(&before));
}

#[test]
fn source_shape_removal_composes_set_add_and_remove() {
    let before = fixture(1, 3);
    let source = open_source(&before);
    let mut edit = source.edit().unwrap();
    assert!(edit.set_shape_text(0, 0_usize, "updated").unwrap());
    edit.add_text_box(0, "added", (1, 2, 300_000, 200_000))
        .unwrap();
    assert!(edit.remove_shape(0, 1_usize).unwrap());
    let (after, patch) = edit.commit().unwrap().into_parts();
    let patch = transported(&patch);
    assert_eq!(
        source
            .apply(&patch)
            .unwrap()
            .into_package()
            .to_bytes()
            .unwrap(),
        after.to_bytes().unwrap()
    );
    let inverse = transported(&patch.inverse());
    assert_eq!(
        members(
            &open_source(&after.to_bytes().unwrap())
                .apply(&inverse)
                .unwrap()
                .into_package()
                .to_bytes()
                .unwrap()
        ),
        members(&before)
    );

    assert_eq!(
        source_texts(&after),
        ["updated", "slide 0 shape 2", "added"]
    );
}

#[test]
fn source_shape_removal_then_add_replaces_removed_highest_id() {
    let before = fixture(1, 2);
    let ids = shape_ids(&before);
    let source = open_source(&before);
    let mut edit = source.edit().unwrap();
    assert!(edit.remove_shape(0, 1_usize).unwrap());
    let added_id = edit
        .add_text_box(0, "replacement", (1, 2, 300_000, 200_000))
        .unwrap();
    assert_eq!(added_id, ids[1]);

    let (after, patch) = edit.commit().unwrap().into_parts();
    let patch = transported(&patch);
    let after_bytes = after.to_bytes().unwrap();
    assert_eq!(
        source
            .apply(&patch)
            .unwrap()
            .into_package()
            .to_bytes()
            .unwrap(),
        after_bytes
    );
    assert_eq!(source_texts(&after), ["slide 0 shape 0", "replacement"]);

    let inverse = transported(&patch.inverse());
    let restored = open_source(&after_bytes)
        .apply(&inverse)
        .unwrap()
        .into_package();
    assert_eq!(members(&restored.to_bytes().unwrap()), members(&before));
}

#[test]
fn source_v4_rejects_an_added_id_colliding_with_a_dangling_timing_reference() {
    let base = fixture(1, 2);
    let ids = shape_ids(&base);
    let dangling_id = ids[1] + 1;
    let before = inject(
        &base,
        "</p:sld>",
        &format!(
            r#"<p:timing><p:tnLst><p:par><p:cTn id="1"><p:stCondLst><p:cond evt="onClick"><p:tgtEl><p:spTgt spid="{dangling_id}"/></p:tgtEl></p:cond></p:stCondLst></p:cTn></p:par></p:tnLst></p:timing>"#
        ),
    );
    let source = open_source(&before);
    let mut edit = source.edit().unwrap();
    assert!(edit.remove_shape(0, 0_usize).unwrap());
    let added_id = edit
        .add_text_box(0, "safe", (1, 2, 300_000, 200_000))
        .unwrap();
    assert_eq!(added_id, dangling_id + 1);
    let (_, patch) = edit.commit().unwrap().into_parts();

    let forged = replace_same_width(
        replace_same_width(
            patch.to_bytes().unwrap(),
            format!(r#"id="{added_id}""#).as_bytes(),
            format!(r#"id="{dangling_id}""#).as_bytes(),
        ),
        format!("Text Box {added_id}").as_bytes(),
        format!("Text Box {dangling_id}").as_bytes(),
    );
    assert!(SourcePatch::from_bytes(&forged).is_err());
    assert_eq!(source.to_bytes().unwrap(), before);
}

#[test]
fn managed_shape_removal_admission_keeps_prior_stage_and_releases_on_drop() {
    const MEMORY: u64 = 16 * 1024 * 1024;
    let before = fixture(1, 3);
    let (budget, _cancellation, context) = managed_context(MEMORY);
    let source = SourcePackage::from_read_at_with_options(
        Arc::new(OwnedSource::new(before.clone())),
        litchi_pptx::OpenOptions::default()
            .source_limits(SourceLimits::default())
            .execution_context(context),
    )
    .unwrap();
    let mut edit = source.edit().unwrap();
    assert!(edit.set_shape_text(0, 0_usize, "kept").unwrap());

    let used = budget.used(Resource::Memory);
    let pressure = budget.reserve(Resource::Memory, MEMORY - used).unwrap();
    assert!(edit.remove_shape(0, 1_usize).is_err());
    drop(pressure);

    let changed = edit.commit().unwrap().into_package();
    assert_eq!(
        source_texts(&changed),
        ["kept", "slide 0 shape 1", "slide 0 shape 2"]
    );
    assert_eq!(source.to_bytes().unwrap(), before);
    drop(changed);
    drop(source);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn managed_shape_removal_cancellation_leaves_edit_unpublished_and_releases_on_drop() {
    const MEMORY: u64 = 16 * 1024 * 1024;
    let before = fixture(1, 2);
    let (budget, cancellation, context) = managed_context(MEMORY);
    let source = SourcePackage::from_read_at_with_options(
        Arc::new(OwnedSource::new(before)),
        litchi_pptx::OpenOptions::default()
            .source_limits(SourceLimits::default())
            .execution_context(context),
    )
    .unwrap();
    let mut edit = source.edit().unwrap();
    assert!(edit.set_shape_text(0, 0_usize, "kept").unwrap());
    cancellation.cancel();
    assert!(edit.remove_shape(0, 1_usize).is_err());
    assert!(edit.commit().is_err());
    drop(source);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn shape_removal_accepts_an_xml_declaration() {
    let before = inject(
        &fixture(1, 2),
        "<p:sld",
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>",
    );

    let mut eager = Package::from_reader(std::io::Cursor::new(before.clone())).unwrap();
    let mut eager_edit = eager.edit().unwrap();
    assert!(eager_edit.remove_shape(0, 0_usize).unwrap());
    eager = eager_edit.commit().unwrap().into_package();
    assert_eq!(eager_texts(&eager).len(), 1);

    let source = open_source(&before);
    let mut source_edit = source.edit().unwrap();
    assert!(source_edit.remove_shape(0, 0_usize).unwrap());
    assert_eq!(
        source_texts(&source_edit.commit().unwrap().into_package()).len(),
        1
    );
}

#[test]
fn shape_removal_refuses_a_grouped_child() {
    let before = with_grouped_first_shape(&fixture(1, 2));
    let source = open_source(&before);
    let mut edit = source.edit().unwrap();
    assert!(edit.remove_shape(0, 1_usize).is_err());
    assert_eq!(source.to_bytes().unwrap(), before);
}

#[test]
fn removing_a_group_removes_its_internal_connector_with_it() {
    let before = with_grouped_first_shape_and_internal_connector(&fixture(1, 2));
    let source = open_source(&before);
    let mut edit = source.edit().unwrap();
    assert!(edit.remove_shape(0, 0_usize).unwrap());
    let after = edit.commit().unwrap().into_package();
    assert_eq!(source_texts(&after), ["slide 0 shape 1"]);
}

#[test]
fn removing_a_group_refuses_a_retained_connector_to_its_child() {
    let base = fixture(1, 2);
    let child_id = shape_ids(&base)[0];
    let grouped = with_grouped_first_shape(&base);
    let connector = format!(
        r#"<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="{}" name="Retained connector"/><p:cNvCxnSpPr><a:stCxn id="{child_id}" idx="0"/><a:endCxn id="{child_id}" idx="0"/></p:cNvCxnSpPr><p:nvPr/></p:nvCxnSpPr><p:spPr><a:prstGeom prst="line"><a:avLst/></a:prstGeom></p:spPr></p:cxnSp>"#,
        child_id.saturating_add(200)
    );
    let before = inject(&grouped, "</p:spTree>", &connector);
    assert_source_remove_refused(&before);
}

#[test]
fn shape_removal_refuses_all_powerpoint_show_event_targets() {
    let shape_id = shape_ids(&fixture(1, 2))[0];
    let events = [
        format!(r#"<p14:triggerEvt type="onClick" time="1" objId="{shape_id}"/>"#),
        format!(r#"<p14:playEvt time="1" objId="{shape_id}"/>"#),
        format!(r#"<p14:stopEvt time="1" objId="{shape_id}"/>"#),
        format!(r#"<p14:pauseEvt time="1" objId="{shape_id}"/>"#),
        format!(r#"<p14:resumeEvt time="1" objId="{shape_id}"/>"#),
        format!(r#"<p14:seekEvt time="1" objId="{shape_id}" seek="1"/>"#),
        format!(r#"<p14:nullEvt time="1" objId="{shape_id}"/>"#),
    ];
    for event in events {
        let extension = format!(
            r#"<p:extLst xmlns:p14="{P14}"><p:ext uri="{SHOW_EVENTS_URI}"><p14:showEvtLst>{event}</p14:showEvtLst></p:ext></p:extLst>"#
        );
        let before = inject(&fixture(1, 2), "</p:sld>", &extension);
        let source = open_source(&before);
        let mut edit = source.edit().unwrap();
        assert!(edit.remove_shape(0, 0_usize).is_err(), "event: {event}");
        assert_eq!(source.to_bytes().unwrap(), before);
    }
}

#[test]
fn shape_removal_refuses_a_control_target() {
    let shape_id = shape_ids(&fixture(1, 2))[0];
    let extra = format!(
        r#"<p:controls><p:control spid="{shape_id}" name="Target control" showAsIcon="0" imgW="10" imgH="20"/></p:controls>"#
    );
    let before = inject(&fixture(1, 2), "</p:cSld>", &extra);
    let source = open_source(&before);
    let mut edit = source.edit().unwrap();
    assert!(edit.remove_shape(0, 0_usize).is_err());
    assert_eq!(source.to_bytes().unwrap(), before);
}

#[test]
fn shape_removal_refuses_an_unknown_extension_target() {
    let shape_id = shape_ids(&fixture(1, 2))[0];
    let extra =
        format!(r#"<x:opaque xmlns:x="urn:litchi:future-shape-ref" shapeId="{shape_id}"/>"#);
    let before = inject(&fixture(1, 2), "</p:sld>", &extra);
    let source = open_source(&before);
    let mut edit = source.edit().unwrap();
    assert!(edit.remove_shape(0, 0_usize).is_err());
    assert_eq!(source.to_bytes().unwrap(), before);
}

#[test]
fn shape_removal_refuses_missing_or_invalid_known_reference_identities() {
    let cases = [
        r#"<p:spTgt/>"#,
        r#"<p:spTgt spid="not-a-number"/>"#,
        r#"<a:stCxn/>"#,
        r#"<a:endCxn id="not-a-number"/>"#,
    ];
    for reference in cases {
        let before = inject(&fixture(1, 2), "</p:sld>", reference);
        assert_source_remove_refused(&before);
    }
}

#[test]
fn source_shape_removal_keeps_unused_relationship_resources() {
    let (before, payload) = with_unused_relationship(&fixture(1, 2));
    let source = open_source(&before);
    let mut edit = source.edit().unwrap();
    assert!(edit.remove_shape(0, 0_usize).unwrap());
    let after = edit.commit().unwrap().into_package().to_bytes().unwrap();

    assert_eq!(member(&after, UNUSED_MEDIA), payload);
    assert_eq!(member(&after, SLIDE_RELS), member(&before, SLIDE_RELS));
    let before_members = members(&before);
    let after_members = members(&after);
    for (name, bytes) in before_members {
        if name != SLIDE {
            assert_eq!(after_members.get(&name), Some(&bytes), "{name}");
        }
    }
}

#[test]
fn source_focused_slide_edit_removal_publishes() {
    let before = fixture(1, 2);
    let editor =
        litchi_pptx::source::Package::from_read_at(Arc::new(OwnedSource::new(before.clone())))
            .unwrap();
    let mut edit = editor.edit().unwrap();
    assert!(edit.remove_shape(0, 0_usize).unwrap());
    let commit = edit.commit().unwrap();
    let mut output = Vec::new();
    commit.package().write_to(&mut output).unwrap();
    assert_eq!(source_texts(&open_source(&output)), ["slide 0 shape 1"]);
}

#[test]
fn source_multiple_removals_transport_but_retained_name_tampering_does_not() {
    let before = fixture(1, 4);
    let source = open_source(&before);
    let mut edit = source.edit().unwrap();
    edit.remove_shape(0, 3_usize).unwrap();
    edit.remove_shape(0, 1_usize).unwrap();
    let (after, patch) = edit.commit().unwrap().into_parts();
    let patch = transported(&patch);
    assert_eq!(source_texts(&after), ["slide 0 shape 0", "slide 0 shape 2"]);
    let inverse = transported(&patch.inverse());
    assert_eq!(
        members(
            &open_source(&after.to_bytes().unwrap())
                .apply(&inverse)
                .unwrap()
                .into_package()
                .to_bytes()
                .unwrap()
        ),
        members(&before)
    );
    let xml = member(&after.to_bytes().unwrap(), SLIDE);
    let scene = Scene::read(&xml, litchi_pptx::shape::Limits::DEFAULT).unwrap();
    let name = scene.roots().next().unwrap().name().unwrap();
    let needle = format!("name=\"{name}\"");
    let mut wire = patch.to_bytes().unwrap();
    let offset = wire
        .windows(needle.len())
        .rposition(|window| window == needle.as_bytes())
        .unwrap();
    wire[offset + 6] = if wire[offset + 6] == b'Z' { b'Y' } else { b'Z' };
    assert!(SourcePatch::from_bytes(&wire).is_err());
}

#[test]
fn removal_wire_is_explicit_and_cannot_nest_envelopes() {
    let before = fixture(1, 2);
    let source = open_source(&before);
    let mut edit = source.edit().unwrap();
    edit.remove_shape(0, 0_usize).unwrap();
    let (_, patch) = edit.commit().unwrap().into_parts();
    let wire = patch.to_bytes().unwrap();
    assert!(wire.starts_with(b"LPSX0004LPSX0002"));
    assert_eq!(patch.retained_bytes(), wire.len());
    assert!(SourcePatch::from_bytes(&wire[8..]).is_err());
    let mut nested = b"LPSX0004".to_vec();
    nested.extend_from_slice(&wire);
    assert!(SourcePatch::from_bytes(&nested).is_err());
    assert!(
        transported(&patch)
            .inverse()
            .to_bytes()
            .unwrap()
            .starts_with(b"LPSX0004")
    );
}

#[test]
fn source_removal_envelope_carries_slide_topology_and_shape_changes() {
    let before = fixture(2, 2);
    let source = open_source(&before);
    let mut edit = source.edit().unwrap();
    edit.remove_shape(0, 0_usize).unwrap();
    edit.remove_slide(Position::new(1)).unwrap();
    let (after, patch) = edit.commit().unwrap().into_parts();
    let wire = patch.to_bytes().unwrap();
    assert!(wire.starts_with(b"LPSX0004LPSX0003"));
    assert_eq!(patch.retained_bytes(), wire.len());
    let patch = SourcePatch::from_bytes(&wire).unwrap();
    let after_bytes = after.to_bytes().unwrap();
    assert_eq!(
        source
            .apply(&patch)
            .unwrap()
            .into_package()
            .to_bytes()
            .unwrap(),
        after_bytes
    );
    let restored = open_source(&after_bytes)
        .apply(&transported(&patch.inverse()))
        .unwrap()
        .into_package();
    assert_eq!(members(&restored.to_bytes().unwrap()), members(&before));
}
