#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use litchi_opc::PackURI;
use std::path::PathBuf;

use super::{Id, Snapshot};
use crate::Package;

const PML: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const STRICT_PML: &str = "http://purl.oclc.org/ooxml/presentationml/main";

fn slide_xml() -> Vec<u8> {
    format!(
        "<p:sld xmlns:p=\"{PML}\"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id=\"2\" name=\"Title\"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/></p:sp><p:pic><p:nvPicPr><p:cNvPr id=\"3\" name=\"Photo\"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill/><p:spPr/></p:pic></p:spTree></p:cSld><p:clrMapOvr/></p:sld>"
    )
    .into_bytes()
}

#[test]
fn detached_edit_round_trips_compact_identifiers_and_inverse() {
    let owner = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    let source = Snapshot::from_source(owner, slide_xml()).unwrap();
    assert_eq!(source.state().creation_id(), None);
    assert_eq!(source.state().shapes().len(), 2);

    let mut edit = source.edit();
    edit.set_creation_id(4_000_000_001_u32);
    edit.set_shape_modification_id("Title", 17_u32).unwrap();
    edit.set_shape_modification_id(1_usize, u32::MAX).unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(commit.diagnostics().changed_identifiers(), 3);
    assert_eq!(
        commit.snapshot().state().creation_id(),
        Some(Id::new(4_000_000_001))
    );
    assert_eq!(
        commit.snapshot().state().shapes()[1].modification_id(),
        Some(Id::new(u32::MAX))
    );
    assert!(!commit.snapshot().source.contains(&b'\n'));
    assert!(!commit.snapshot().source.contains(&b'\r'));
    assert!(
        commit
            .snapshot()
            .source
            .windows(b"<p14:creationId".len())
            .any(|window| window == b"<p14:creationId")
    );
    assert!(
        commit
            .snapshot()
            .source
            .windows(b"<p14:modId".len())
            .any(|window| window == b"<p14:modId")
    );

    let inverse = commit.patch().inverse();
    assert_eq!(inverse.after(), source.state());
    assert_eq!(inverse.target, source.source);
}

#[test]
fn duplicate_shape_modification_ids_are_refused_before_commit() {
    let owner = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    let source = Snapshot::from_source(owner, slide_xml()).unwrap();
    let mut edit = source.edit();
    edit.set_shape_modification_id("Title", 42_u32).unwrap();
    let error = edit.set_shape_modification_id("Photo", 42_u32).unwrap_err();
    assert!(error.to_string().contains("must be unique"));
}

#[test]
fn malformed_producer_duplicates_remain_readable_and_repairable_by_position() {
    let extension = format!(
        "<p:nvPr><p:extLst><p:ext uri=\"{}\"><p14:modId xmlns:p14=\"{}\" val=\"42\"/></p:ext></p:extLst></p:nvPr>",
        super::MODIFICATION_EXTENSION_URI,
        super::NAMESPACE
    );
    let xml = String::from_utf8(slide_xml())
        .unwrap()
        .replace("<p:nvPr/>", &extension)
        .into_bytes();
    let owner = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    let source = Snapshot::from_source(owner, xml).unwrap();
    assert_eq!(
        source.state().shapes()[0].modification_id(),
        source.state().shapes()[1].modification_id()
    );
    assert!(source.edit().commit().is_err());

    let mut repair = source.edit();
    repair.clear_shape_modification_id(1_usize).unwrap();
    let repaired = repair.commit().unwrap();
    assert_eq!(
        repaired.snapshot().state().shapes()[0].modification_id(),
        Some(Id::new(42))
    );
    assert_eq!(
        repaired.snapshot().state().shapes()[1].modification_id(),
        None
    );
}

#[test]
fn strict_aliases_round_trip_without_transitional_namespace_leakage() {
    let xml = String::from_utf8(slide_xml())
        .unwrap()
        .replace(PML, STRICT_PML)
        .replace("xmlns:p=", "xmlns:q=")
        .replace("<p:", "<q:")
        .replace("</p:", "</q:")
        .into_bytes();
    let owner = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    let source = Snapshot::from_source(owner, xml).unwrap();
    let mut edit = source.edit();
    edit.set_creation_id(1_u32);
    edit.set_shape_modification_id("Title", 2_u32).unwrap();
    let commit = edit.commit().unwrap();
    let output = String::from_utf8(commit.snapshot().source.clone()).unwrap();
    assert!(output.contains(STRICT_PML));
    assert!(!output.contains(PML));
    assert_eq!(commit.snapshot().state().creation_id(), Some(Id::new(1)));
}

#[test]
fn facade_publishes_and_reverses_exact_source_checked_commit() {
    let mut authored = Package::new().unwrap();
    let slide = authored.presentation_mut().unwrap().add_slide().unwrap();
    slide.add_text_box("Title", 0, 0, 100, 100);
    slide.add_ellipse(0, 0, 100, 100, None);
    let bytes = authored.to_bytes().unwrap();
    let mut package = Package::from_bytes(&bytes).unwrap();

    let source = package.change_tracking(0_usize).unwrap();
    let mut edit = source.edit();
    edit.set_creation_id(7_u32);
    edit.set_shape_modification_id(0_usize, 11_u32).unwrap();
    let commit = edit.commit().unwrap();
    let inverse = commit.patch().inverse();
    package.apply_change_tracking_commit(commit).unwrap();

    let stored = package.change_tracking(0_usize).unwrap();
    assert_eq!(stored.state().creation_id(), Some(Id::new(7)));
    assert_eq!(
        stored.state().shapes()[0].modification_id(),
        Some(Id::new(11))
    );
    package.apply_change_tracking_patch(&inverse).unwrap();
    assert_eq!(
        package.change_tracking(0_usize).unwrap().state(),
        source.state()
    );

    package.opc.relate_to(
        "_xmlsignatures/origin.sigs",
        litchi_opc::constants::relationship_type::DIGITAL_SIGNATURE_ORIGIN,
    );
    let no_op = package
        .change_tracking(0_usize)
        .unwrap()
        .edit()
        .commit()
        .unwrap();
    assert!(!no_op.is_changed());
    package.apply_change_tracking_commit(no_op).unwrap();
    assert!(package.opc.is_signed());

    let mut stale_edit = source.edit();
    stale_edit.set_creation_id(99_u32);
    let stale_commit = stale_edit.commit().unwrap();
    let mut current = package.change_tracking(0_usize).unwrap().edit();
    current.set_creation_id(100_u32);
    package
        .apply_change_tracking_commit(current.commit().unwrap())
        .unwrap();
    assert!(package.apply_change_tracking_commit(stale_commit).is_err());
}

#[test]
fn unknown_extensions_survive_identifier_add_and_remove_byte_exact() {
    let future = "<p:ext uri=\"urn:future\"><future:data xmlns:future=\"urn:future\" answer=\"42\"/></p:ext>";
    let xml = format!(
        "<p:sld xmlns:p=\"{PML}\"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/></p:spTree><p:extLst>{future}</p:extLst></p:cSld></p:sld>"
    )
    .into_bytes();
    let owner = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    let source = Snapshot::from_source(owner, xml).unwrap();
    let mut edit = source.edit();
    edit.set_creation_id(9_u32);
    let added = edit.commit().unwrap().snapshot().clone();
    assert!(
        added
            .source
            .windows(future.len())
            .any(|window| window == future.as_bytes())
    );
    let mut clear = added.edit();
    clear.clear_creation_id();
    let cleared = clear.commit().unwrap();
    assert_eq!(cleared.snapshot().source, source.source);
}

#[test]
fn empty_element_expansion_does_not_emit_indentation_or_extra_spaces() {
    let xml = String::from_utf8(slide_xml())
        .unwrap()
        .replacen("<p:nvPr/>", "<p:nvPr />", 1)
        .into_bytes();
    let owner = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    let source = Snapshot::from_source(owner, xml).unwrap();
    let mut edit = source.edit();
    edit.set_shape_modification_id(0_usize, 5_u32).unwrap();
    let commit = edit.commit().unwrap();
    assert!(!commit.snapshot().source.contains(&b'\n'));
    assert!(!commit.snapshot().source.contains(&b'\r'));
    assert!(
        !commit
            .snapshot()
            .source
            .windows(2)
            .any(|window| window == b" >")
    );
    assert!(
        !commit
            .snapshot()
            .source
            .windows(3)
            .any(|window| window == b" />")
    );
}

#[test]
fn real_powerpoint_fixture_exposes_bounded_change_tracking_state() {
    let path =
        PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../test-data/ooxml/pptx/shapes.pptx");
    let package = Package::open(path).unwrap();
    let slides = package.presentation().unwrap().slide_count().unwrap();
    assert!(slides > 0);
    let mut creation_ids = 0usize;
    let mut shapes = 0usize;
    for position in 0..slides {
        let snapshot = package.change_tracking(position).unwrap();
        creation_ids += usize::from(snapshot.state().creation_id().is_some());
        shapes += snapshot.state().shapes().len();
    }
    assert!(creation_ids > 0);
    assert!(shapes > 0);
}
