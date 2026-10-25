mod durable;

use super::super::MAX_SOURCE_BACKED_SLIDE_BATCH;
use super::*;
use crate::source::Package;
use litchi_core::Position;

fn package(bytes: Vec<u8>) -> Package {
    Package::from_read_at(Arc::new(CountingSource::new(bytes))).unwrap()
}
fn text(package: &Package, slide: usize) -> String {
    let bytes = package.to_bytes().unwrap();
    crate::Package::from_bytes(&bytes)
        .unwrap()
        .presentation()
        .unwrap()
        .slides()
        .unwrap()[slide]
        .text()
        .unwrap()
}

#[test]
fn source_owner_retains_custom_shape_profile() {
    let defaults = crate::shape::Limits::DEFAULT;
    let shape_limits = crate::shape::Limits::new(
        defaults.input_bytes(),
        defaults.output_bytes(),
        defaults.depth(),
        defaults.shapes(),
        1,
        defaults.retained_text_bytes(),
    )
    .unwrap();
    let package = Package::from_read_at_with_options(
        Arc::new(CountingSource::new(fixture())),
        crate::OpenOptions::default().shape_limits(shape_limits),
    )
    .unwrap();
    let mut edit = package.edit().unwrap();
    assert!(matches!(
        edit.add_text_box(0, "bounded", (0, 0, 914_400, 914_400)),
        Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::Shapes),
            observed: 2,
            maximum: 1,
        })
    ));
}

#[test]
fn source_owner_retains_custom_transition_profile() {
    let output_limit = crate::transition::Limits::DEFAULT
        .with_output_bytes(1)
        .unwrap();
    let package = Package::from_read_at_with_options(
        Arc::new(CountingSource::new(fixture())),
        crate::OpenOptions::default().transition_limits(output_limit),
    )
    .unwrap();
    let mut edit = package.edit().unwrap();
    let transition =
        crate::transition::Transition::new(crate::transition::Kind::Fade { black: None });
    assert!(matches!(
        edit.set_slide_transition(0, Some(&transition)),
        Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::OutputBytes),
            observed,
                maximum: 1,
        }) if observed > 1
    ));
}

#[test]
fn immutable_commits_compose_replay_inverse_and_keep_original_exact() {
    let bytes = fixture();
    let source = package(bytes.clone());
    let noop = source.edit().unwrap().commit().unwrap();
    assert!(!noop.diagnostics().changed());
    assert!(noop.patch().is_noop());
    assert_eq!(noop.package().to_bytes().unwrap(), bytes);
    let mut edit = source.edit().unwrap();
    assert!(edit.set_shape_text(0, 0_usize, "private-new-text").unwrap());
    assert!(edit.set_shape_text(0, 0_usize, "final text").unwrap());
    let commit = edit.commit().unwrap();
    assert!(commit.diagnostics().changed());
    let (changed, patch) = commit.clone().into_parts();
    assert_eq!(text(&changed, 0), "final text");
    assert_eq!(source.to_bytes().unwrap(), bytes);
    assert_eq!(
        source
            .apply(&patch)
            .unwrap()
            .into_package()
            .to_bytes()
            .unwrap(),
        changed.to_bytes().unwrap()
    );
    assert!(changed.apply(&patch).is_err());
    assert_eq!(
        changed
            .apply(&patch.inverse())
            .unwrap()
            .into_package()
            .to_bytes()
            .unwrap(),
        bytes
    );
    assert!(package(bytes.clone()).apply(&patch).is_err());
    let mut second = changed.edit().unwrap();
    second
        .set_shape_text(1, 0_usize, "second revision")
        .unwrap();
    let (both, second_patch) = second.commit().unwrap().into_parts();
    assert_eq!(text(&both, 0), "final text");
    assert_eq!(text(&both, 1), "second revision");
    assert_eq!(
        both.apply(&second_patch.inverse())
            .unwrap()
            .into_package()
            .to_bytes()
            .unwrap(),
        changed.to_bytes().unwrap()
    );
    for debug in [
        format!("{source:?}"),
        format!("{commit:?}"),
        format!("{patch:?}"),
    ] {
        assert!(!debug.contains("final text"));
        assert!(!debug.contains("ppt/slides"));
    }
}

#[test]
fn shared_projection_can_enter_primary_owner_and_untouched_slide_stays_deferred() {
    let source = Arc::new(CountingSource::new(fixture()));
    let projection = SourceBackedPresentation::from_read_at(source.clone()).unwrap();
    let original_slide = projection.slide(0).unwrap();
    let clone = projection.clone();
    let package = Package::from_presentation(&projection).unwrap();
    let mut edit = package.edit().unwrap();
    edit.set_shape_text(0, 0_usize, "changed").unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(source.second_payload_reads.load(Ordering::SeqCst), 0);
    assert_eq!(original_slide.text().unwrap(), "First slide");
    assert_eq!(clone.slide(0).unwrap().text().unwrap(), "First slide");
    assert_eq!(
        commit
            .package()
            .presentation()
            .unwrap()
            .slide_count()
            .unwrap(),
        2
    );
    let scene = commit
        .package()
        .presentation()
        .unwrap()
        .slide(0)
        .unwrap()
        .scene()
        .unwrap();
    assert_eq!(scene.len(), 1);
    assert_eq!(text(commit.package(), 0), "changed");
}

#[test]
fn order_staging_keeps_slide_payloads_deferred_until_a_logical_read() {
    let source = Arc::new(CountingSource::new(fixture()));
    let package = Package::from_read_at(source.clone()).unwrap();
    let before = source.second_payload_reads.load(Ordering::SeqCst);
    let mut edit = package.edit().unwrap();
    assert!(edit.move_slide(Position::new(0), Position::new(1)).unwrap());
    assert_eq!(source.second_payload_reads.load(Ordering::SeqCst), before);

    let commit = edit.commit().unwrap();
    assert_eq!(source.second_payload_reads.load(Ordering::SeqCst), before);
    assert_eq!(text(commit.package(), 0), "Second slide");
    assert!(source.second_payload_reads.load(Ordering::SeqCst) > before);
}

#[test]
fn removal_dependent_then_insertion_respects_exact_selected_part_limit() {
    let mut builder = crate::builder::Builder::new().unwrap();
    builder
        .presentation()
        .unwrap()
        .add_slide()
        .unwrap()
        .set_notes("owned by the removed slide");
    builder.presentation().unwrap().add_slide().unwrap();
    let bytes = builder.build().unwrap().to_bytes().unwrap();
    let source_catalog =
        litchi_opc::OpcPackage::from_bytes(&bytes, SourceLimits::default()).unwrap();
    let source_parts = source_catalog.iter_parts().count();
    assert!(source_parts > 1);
    let ingress_limits = SourceLimits::builder()
        .max_parts(source_parts)
        .unwrap()
        .build()
        .unwrap();
    let source = Package::from_read_at_with_options(
        Arc::new(litchi_core::OwnedSource::new(bytes)),
        crate::OpenOptions::default().source_limits(ingress_limits),
    )
    .unwrap();
    let layout_name = source
        .presentation()
        .unwrap()
        .slide(1)
        .unwrap()
        .layout()
        .unwrap()
        .unwrap()
        .name()
        .unwrap()
        .to_owned();

    // The removed slide contributes one exclusive notes Part.  Two inserted
    // dependency-free slides exactly replace the removed slide and dependent,
    // so the final catalog remains at the selected original source_parts.
    let mut edit = source.edit().unwrap();
    assert!(edit.remove_slide(Position::new(0)).unwrap());
    edit.insert_slide(Position::new(0), &layout_name).unwrap();
    edit.insert_slide(Position::new(1), &layout_name).unwrap();
    let candidate = edit.commit().unwrap().into_package();
    let output = candidate.to_bytes().unwrap();
    let output_catalog =
        litchi_opc::OpcPackage::from_bytes(&output, SourceLimits::default()).unwrap();
    assert_eq!(output_catalog.iter_parts().count(), source_parts);

    let final_limits = SourceLimits::builder()
        .max_parts(source_parts)
        .unwrap()
        .build()
        .unwrap();
    let reopened = Package::from_read_at_with_options(
        Arc::new(litchi_core::OwnedSource::new(output.clone())),
        crate::OpenOptions::default().source_limits(final_limits),
    )
    .unwrap();
    assert_eq!(reopened.presentation().unwrap().slide_count().unwrap(), 3);
    let layout_name = reopened
        .presentation()
        .unwrap()
        .slide(0)
        .unwrap()
        .layout()
        .unwrap()
        .unwrap()
        .name()
        .unwrap()
        .to_owned();
    let mut rejected = reopened.edit().unwrap();
    assert!(matches!(
        rejected.insert_slide(Position::new(2), &layout_name),
        Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::Parts),
            observed,
            maximum,
        }) if observed == source_parts + 1 && maximum == source_parts
    ));
    assert!(rejected.commit().unwrap().patch().is_noop());
    assert_eq!(reopened.to_bytes().unwrap(), output);
}

#[test]
fn rejected_stage_preserves_prior_changes_and_drop_preserves_source() {
    let bytes = fixture();
    let source = package(bytes.clone());
    let mut edit = source.edit().unwrap();
    edit.set_shape_text(0, 0_usize, "kept").unwrap();
    assert!(edit.set_shape_text(1, 999_usize, "rejected").is_err());
    let changed = edit.commit().unwrap().into_package();
    assert_eq!(text(&changed, 0), "kept");
    let mut dropped = source.edit().unwrap();
    dropped.set_shape_text(0, 0_usize, "dropped").unwrap();
    drop(dropped);
    assert_eq!(source.to_bytes().unwrap(), bytes);
}

#[test]
fn retained_handles_recheck_freshness_and_atomic_save_keeps_destination() {
    let source = Arc::new(CountingSource::new(fixture()));
    let package = Package::from_read_at(source.clone()).unwrap();
    let slide = package.presentation().unwrap().slide(0).unwrap();
    let mut edit = package.edit().unwrap();
    edit.set_shape_text(0, 0_usize, "candidate").unwrap();
    let commit = edit.commit().unwrap();
    source.changed();
    assert!(slide.scene().is_err());
    assert!(slide.text().is_err());
    assert!(slide.text_and_name().is_err());
    let mut report = crate::ToleranceReport::default();
    assert!(slide.text_with_tolerance(&mut report).is_err());
    assert!(report.is_clean());
    assert!(package.edit().is_err());
    assert!(package.apply(commit.patch()).is_err());
    assert!(commit.package().to_bytes().is_err());
    let directory = tempfile::tempdir().unwrap();
    let destination = directory.path().join("output.pptx");
    fs::write(&destination, b"unchanged destination").unwrap();
    assert!(commit.package().save(&destination).is_err());
    assert_eq!(fs::read(destination).unwrap(), b"unchanged destination");
}

#[test]
fn managed_cancellation_applies_to_commits_projections_replay_and_save() {
    let (_budget, cancellation, context) = source_backed_context(16 * 1024 * 1024);
    let package = Package::from_read_at_with_options(
        Arc::new(CountingSource::new(fixture())),
        crate::OpenOptions::default()
            .source_limits(SourceLimits::default())
            .execution_context(context),
    )
    .unwrap();
    let mut edit = package.edit().unwrap();
    edit.set_shape_text(0, 0_usize, "candidate").unwrap();
    let commit = edit.commit().unwrap();
    let slide = commit.package().presentation().unwrap().slide(0).unwrap();
    let pending = package.edit().unwrap();
    cancellation.cancel();
    assert!(pending.commit().is_err());
    assert!(slide.scene().is_err());
    assert!(slide.text().is_err());
    assert!(slide.text_and_name().is_err());
    let mut report = crate::ToleranceReport::default();
    assert!(slide.text_with_tolerance(&mut report).is_err());
    assert!(report.is_clean());
    assert!(package.apply(commit.patch()).is_err());
    assert!(commit.package().to_bytes().is_err());
}

fn fixture() -> Vec<u8> {
    let slide_list = r#"<p:sldId id="256" r:id="rId1"/><p:sldId id="257" r:id="rId2"/>"#;
    let presentation = format!(
        r#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst>{slide_list}</p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/></p:presentation>"#
    );
    let mut writer =
        StreamingArchiveWriter::new(soapberry_zip::office::StreamingArchiveLimits::default());
    writer
            .write_stored(
                "[Content_Types].xml",
                br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>"#,
            )
            .unwrap();
    writer
            .write_stored(
                "_rels/.rels",
                br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>"#,
            )
            .unwrap();
    writer
        .write_stored("ppt/presentation.xml", presentation.as_bytes())
        .unwrap();
    writer
            .write_stored(
                "ppt/_rels/presentation.xml.rels",
                br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/></Relationships>"#,
            )
            .unwrap();
    writer
            .write_stored(
                "ppt/slides/slide1.xml",
                br#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="2" name="private shape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>First slide</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>"#,
            )
            .unwrap();
    let padding = "x".repeat(128 * 1024);
    let second = format!(
        r#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><!--{marker}{padding}--><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="2" name="private shape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Second slide</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>"#,
        marker = std::str::from_utf8(SECOND_MARKER).unwrap(),
    );
    writer
        .write_stored("ppt/slides/slide2.xml", second.as_bytes())
        .unwrap();
    writer.finish_to_bytes().unwrap()
}

#[test]
fn disjoint_replay_preserves_destination_changes_and_atomic_save_reopens() {
    let source = package(fixture());
    let mut left = source.edit().unwrap();
    left.set_shape_text(0, 0_usize, "left").unwrap();
    let left = left.commit().unwrap();
    let mut right = source.edit().unwrap();
    right.set_shape_text(1, 0_usize, "right").unwrap();
    let right = right.commit().unwrap().into_package();
    let combined = right.apply(left.patch()).unwrap().into_package();
    assert_eq!(text(&combined, 0), "left");
    assert_eq!(text(&combined, 1), "right");
    assert_eq!(text(&right, 0), "First slide");
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("out.pptx");
    fs::write(&path, b"old").unwrap();
    combined.save(&path).unwrap();
    let reopened = Package::open(&path).unwrap();
    assert_eq!(text(&reopened, 1), "right");
}

#[test]
fn managed_candidate_reservation_follows_patch_and_detached_shape_handles() {
    let (budget, _cancel, context) = source_backed_context(16 * 1024 * 1024);
    let source = Package::from_read_at_with_options(
        Arc::new(CountingSource::new(fixture())),
        crate::OpenOptions::default()
            .source_limits(SourceLimits::default())
            .execution_context(context),
    )
    .unwrap();
    let mut edit = source.edit().unwrap();
    edit.set_shape_text(0, 0_usize, "reserved text").unwrap();
    let (candidate, patch) = edit.commit().unwrap().into_parts();
    let scene = candidate
        .presentation()
        .unwrap()
        .slide(0)
        .unwrap()
        .scene()
        .unwrap();
    let handle = scene.shape(0_usize).unwrap().unwrap();
    drop((source, candidate, scene));
    assert!(budget.used(Resource::Memory) > 0);
    drop(patch);
    assert!(budget.used(Resource::Memory) > 0);
    assert_eq!(handle.view().name(), Some("private shape"));
    drop(handle);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn rejected_managed_reservation_is_atomic_and_budget_recovers() {
    const MEMORY: u64 = 16 * 1024 * 1024;
    let (budget, _cancel, context) = source_backed_context(MEMORY);
    let source = Package::from_read_at_with_options(
        Arc::new(CountingSource::new(fixture())),
        crate::OpenOptions::default()
            .source_limits(SourceLimits::default())
            .execution_context(context),
    )
    .unwrap();
    let mut edit = source.edit().unwrap();
    edit.set_shape_text(0, 0_usize, "admitted").unwrap();
    let used = budget.used(Resource::Memory);
    let pressure = budget.reserve(Resource::Memory, MEMORY - used).unwrap();
    assert!(matches!(
        edit.set_shape_text(0, 0_usize, "rejected"),
        Err(Error::Opc(litchi_opc::OpcError::Execution(
            ExecutionError::ResourceLimit(_)
        )))
    ));
    drop(pressure);
    assert_eq!(budget.used(Resource::Memory), used);
    let admitted = edit.commit().unwrap().into_package();
    assert_eq!(text(&admitted, 0), "admitted");
    drop((source, admitted));
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn signed_noop_is_exact_and_changed_stage_refuses() {
    let bytes = fs::read("../../test-data/poi/test-data/xmldsign/hello-world-signed.pptx").unwrap();
    let source =
        Package::from_read_at(Arc::new(litchi_core::OwnedSource::new(bytes.clone()))).unwrap();
    let commit = source.edit().unwrap().commit().unwrap();
    assert_eq!(commit.package().to_bytes().unwrap(), bytes);
    assert_eq!(
        source
            .apply(commit.patch())
            .unwrap()
            .into_package()
            .to_bytes()
            .unwrap(),
        bytes
    );
    let mut edit = source.edit().unwrap();
    assert!(matches!(
        edit.set_shape_text(0, 0_usize, "changed"),
        Err(Error::Opc(
            litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy
        ))
    ));
    assert!(edit.commit().unwrap().patch().is_noop());
}

#[test]
fn changed_slide_limit_refuses_atomically_across_successive_commits() {
    let mut builder = crate::builder::Builder::new().unwrap();
    for _ in 0..=MAX_SOURCE_BACKED_SLIDE_BATCH {
        builder
            .presentation()
            .unwrap()
            .add_slide()
            .unwrap()
            .add_text_box("original", 0, 0, 914_400, 914_400);
    }
    let authored = builder.build().unwrap();
    let source = Package::from_read_at(Arc::new(litchi_core::OwnedSource::new(
        authored.to_bytes().unwrap(),
    )))
    .unwrap();
    let mut candidate = source.clone();
    for slide in 0..MAX_SOURCE_BACKED_SLIDE_BATCH {
        let mut edit = candidate.edit().unwrap();
        edit.set_shape_text(slide, 0_usize, "changed").unwrap();
        candidate = edit.commit().unwrap().into_package();
    }
    let mut edit = candidate.edit().unwrap();
    assert!(matches!(
        edit.set_shape_text(MAX_SOURCE_BACKED_SLIDE_BATCH, 0_usize, "refused"),
        Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::Slides),
            ..
        })
    ));
    assert!(edit.commit().unwrap().patch().is_noop());
    assert_eq!(text(&candidate, MAX_SOURCE_BACKED_SLIDE_BATCH), "original");
    assert_eq!(text(&source, 0), "original");
}

#[test]
fn focused_edit_uses_selected_patch_bound_before_candidate_publication() {
    let source = package(fixture());
    let mut edit = source
        .edit_with_options(super::super::primary::EditOptions::default().max_patch_bytes(1))
        .unwrap();
    assert!(matches!(
        edit.set_shape_text(0, 0_usize, "bounded"),
        Err(Error::Limit { .. })
    ));
    assert!(edit.commit().unwrap().patch().is_noop());
    assert_eq!(text(&source, 0), "First slide");
}

fn with_first_slide(xml: &[u8]) -> Vec<u8> {
    let mut package =
        litchi_opc::OpcPackage::from_bytes(&fixture(), SourceLimits::default()).unwrap();
    package
        .get_part_mut(&PackURI::new("/ppt/slides/slide1.xml").unwrap())
        .unwrap()
        .set_blob(xml.to_vec());
    litchi_opc::PackageWriter::to_bytes(&package).unwrap()
}
fn mce_slide(ignored: &str) -> String {
    format!(
        r#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:unknown"><p:cSld><p:spTree><mc:AlternateContent><mc:Choice Requires="x">{ignored}</mc:Choice><mc:Fallback><p:sp><p:nvSpPr><p:cNvPr id="7" name="private MCE shape"/></p:nvSpPr></p:sp></mc:Fallback></mc:AlternateContent></p:spTree></p:cSld></p:sld>"#
    )
}

#[test]
fn valid_mce_reads_remain_available_and_processed_scenes_keep_managed_charge() {
    let (budget, _cancel, context) = source_backed_context(16 * 1024 * 1024);
    let source = Package::from_read_at_with_options(
        Arc::new(litchi_core::OwnedSource::new(with_first_slide(
            mce_slide("").as_bytes(),
        ))),
        crate::OpenOptions::default()
            .source_limits(SourceLimits::default())
            .execution_context(context),
    )
    .unwrap();
    let scene = source
        .presentation()
        .unwrap()
        .slide(0)
        .unwrap()
        .scene()
        .unwrap();
    assert!(scene.is_rewritten());
    let handle = scene.shape("private MCE shape").unwrap().unwrap();
    let mut edit = source.edit().unwrap();
    assert!(matches!(
        edit.set_shape_text(0, 0_usize, "unsupported"),
        Err(Error::Refused(_))
    ));
    assert!(edit.commit().unwrap().patch().is_noop());
    drop((scene, source));
    assert!(budget.used(Resource::Memory) > 0);
    assert_eq!(handle.view().name(), Some("private MCE shape"));
    drop(handle);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn caller_physical_xml_limits_apply_to_reads_and_edits_before_mce() {
    let depth = "<x:n>".repeat(20) + &"</x:n>".repeat(20);
    let events = "<x:n/>".repeat(150);
    let attributes = format!("<x:n value=\"{}\"/>", "x".repeat(200));
    let cases = [
        (
            depth,
            SourceLimits::builder()
                .max_depth(16)
                .unwrap()
                .build()
                .unwrap(),
            SourceLimitKind::Depth,
        ),
        (
            events,
            SourceLimits::builder()
                .max_events(100)
                .unwrap()
                .build()
                .unwrap(),
            SourceLimitKind::Events,
        ),
        (
            attributes,
            SourceLimits::builder()
                .max_xml_attribute_bytes(128)
                .unwrap()
                .max_relationship_target_bytes(128)
                .unwrap()
                .build()
                .unwrap(),
            SourceLimitKind::XmlAttributeBytes,
        ),
    ];
    for (ignored, limits, expected) in cases {
        let projection = SourceBackedPresentation::from_read_at_with_options(
            Arc::new(litchi_core::OwnedSource::new(with_first_slide(
                mce_slide(&ignored).as_bytes(),
            ))),
            crate::OpenOptions::default().source_limits(limits),
        )
        .unwrap();
        let source = Package::from_presentation(&projection).unwrap();
        assert!(matches!(source.presentation().unwrap().slide(0),
            Err(Error::Opc(litchi_opc::OpcError::Limit { kind, .. })) if kind == expected));
        let mut edit = source.edit().unwrap();
        assert!(matches!(edit.set_shape_text(0, 0_usize, "rejected"),
            Err(Error::Opc(litchi_opc::OpcError::Limit { kind, .. })) if kind == expected));
        assert!(edit.commit().unwrap().patch().is_noop());
    }
}

#[test]
fn revision_text_reads_match_reopened_native_parts_and_keep_other_slides_cold() {
    for replacement in ["", "edited & escaped <text>", "line one\nline two"] {
        let counted = Arc::new(CountingSource::new(fixture()));
        let source = Package::from_read_at(counted.clone()).unwrap();
        let original = source.presentation().unwrap().slide(0).unwrap();
        let mut edit = source.edit().unwrap();
        edit.set_shape_text(0, 0_usize, replacement).unwrap();
        let committed = edit.commit().unwrap().into_package();
        let slide = committed.presentation().unwrap().slide(0).unwrap();
        let combined = slide.text_and_name().unwrap();
        assert_eq!(slide.text().unwrap(), combined.0);
        assert_eq!(counted.second_payload_reads.load(Ordering::SeqCst), 0);
        let eager = crate::Package::from_bytes(&committed.to_bytes().unwrap()).unwrap();
        let projection = eager.presentation().unwrap();
        let slides = projection.slides().unwrap();
        assert_eq!(combined, slides[0].part().text_and_name().unwrap());
        assert_eq!(original.text().unwrap(), "First slide");
    }
}

#[test]
fn revision_text_matches_mce_projection_without_rewriting_source() {
    let bytes = with_first_slide(mce_slide("").as_bytes());
    let source =
        Package::from_read_at(Arc::new(litchi_core::OwnedSource::new(bytes.clone()))).unwrap();
    let slide = source.presentation().unwrap().slide(0).unwrap();
    let eager = crate::Package::from_bytes(&bytes).unwrap();
    assert_eq!(
        slide.text().unwrap(),
        eager.presentation().unwrap().slides().unwrap()[0]
            .text()
            .unwrap()
    );
    assert_eq!(source.to_bytes().unwrap(), bytes);
}
