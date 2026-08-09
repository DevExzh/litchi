#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use super::package::is_notes_slide_rel;
use super::*;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::part::{BlobPart, Part};
use litchi_opc::{OpcPackage, PackURI};
const POI: &[u8] = include_bytes!("../../../../test-data/poi/test-data/slideshow/prProps.pptx");
const LO: &[u8] =
    include_bytes!("../../../../test-data/libreoffice-core/oox/qa/unit/data/tdf131082.pptx");
fn presentation() -> PackURI {
    PackURI::new("/ppt/presentation.xml").unwrap()
}

#[test]
fn plain_text_writer_escapes_and_rejects_invalid_xml() {
    let xml = write_text("A < B & C").unwrap();
    assert_eq!(
        xml,
        write_text_with(Conformance::Transitional, "A < B & C").unwrap()
    );
    let xml = std::str::from_utf8(&xml).unwrap();
    assert!(xml.starts_with("<?xml version="));
    assert!(xml.contains("<a:t>A &lt; B &amp; C</a:t>"));
    assert!(xml.ends_with("</p:notes>"));
    assert!(write_text("bad\u{0}text").is_err());
}

#[test]
fn notes_master_template_is_canonical_and_deterministic() {
    assert!(master_xml().contains("<p:notesMaster"));
    assert_eq!(master_xml().as_ptr(), master_xml().as_ptr());
}

#[test]
fn consuming_put_moves_changed_xml_and_preserves_signed_no_ops() {
    let (mut package, name) = synthetic(Conformance::Transitional);
    let graph = load(&package, &name).unwrap().unwrap();
    package.relate_to(
        "_xmlsignatures/origin.sigs",
        litchi_opc::constants::relationship_type::DIGITAL_SIGNATURE_ORIGIN,
    );
    assert!(package.is_signed());

    put(&mut package, &name, graph).unwrap();
    assert!(package.is_signed());

    let mut graph = load(&package, &name).unwrap().unwrap();
    graph.slides_mut()[0].replace_xml(write_text("Updated note").unwrap());
    put(&mut package, &name, graph).unwrap();
    assert!(!package.is_signed());
    assert_eq!(
        load(&package, &name).unwrap().unwrap().slides()[0]
            .text()
            .unwrap()
            .as_deref(),
        Some("Updated note")
    );
}
#[test]
fn poi_and_libreoffice_notes_graphs_load_and_store_deterministically() {
    for bytes in [POI, LO] {
        let mut package = OpcPackage::from_bytes(bytes).unwrap();
        let name = presentation();
        let graph = load(&package, &name).unwrap().unwrap();
        assert_eq!(graph.slides.len(), 1);
        assert_eq!(graph.master.content_type, ct::PML_NOTES_MASTER);
        put(&mut package, &name, graph).unwrap();
        let graph = load(&package, &name).unwrap().unwrap();
        assert_eq!(graph.slides.len(), 1);
        put(&mut package, &name, graph).unwrap();
        assert_eq!(load(&package, &name).unwrap().unwrap().slides.len(), 1);
    }
}
fn synthetic(conformance: Conformance) -> (OpcPackage, PackURI) {
    let p = conformance.p();
    let a = conformance.a();
    let r = conformance.r();
    let mut package = OpcPackage::new();
    let presentation = presentation();
    let mut pres=BlobPart::new(presentation.clone(),ct::PML_PRESENTATION_MAIN.into(),format!("<p:presentation xmlns:p=\"{p}\" xmlns:r=\"{r}\"><p:notesMasterIdLst><p:notesMasterId r:id=\"rIdMaster\"/></p:notesMasterIdLst><p:sldIdLst><p:sldId id=\"256\" r:id=\"rIdSlide\"/></p:sldIdLst><p:notesSz cx=\"1\" cy=\"1\"/></p:presentation>").into_bytes());
    pres.rels_mut().add_relationship(
        conformance.notes_master_rel().into(),
        "notesMasters/notesMaster1.xml".into(),
        "rIdMaster".into(),
        false,
    );
    pres.rels_mut().add_relationship(
        conformance.slide_rel().into(),
        "slides/slide1.xml".into(),
        "rIdSlide".into(),
        false,
    );
    package.add_part(Box::new(pres));
    let slide_uri = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    let mut slide = BlobPart::new(
        slide_uri,
        SLIDE_CT.into(),
        format!("<p:sld xmlns:p=\"{p}\"><p:cSld/></p:sld>").into_bytes(),
    );
    slide.rels_mut().add_relationship(
        conformance.notes_slide_rel().into(),
        "../notesSlides/notesSlide1.xml".into(),
        "rIdNotes".into(),
        false,
    );
    package.add_part(Box::new(slide));
    let master_uri = PackURI::new("/ppt/notesMasters/notesMaster1.xml").unwrap();
    let mut master = BlobPart::new(
        master_uri,
        ct::PML_NOTES_MASTER.into(),
        format!(
            "<p:notesMaster xmlns:p=\"{p}\" xmlns:a=\"{a}\"><p:cSld/><p:clrMap/></p:notesMaster>"
        )
        .into_bytes(),
    );
    master.rels_mut().add_relationship(
        conformance.theme_rel().into(),
        "../theme/theme2.xml".into(),
        "rIdTheme".into(),
        false,
    );
    package.add_part(Box::new(master));
    let notes_uri = PackURI::new("/ppt/notesSlides/notesSlide1.xml").unwrap();
    let mut notes=BlobPart::new(notes_uri,ct::PML_NOTES_SLIDE.into(),format!("<p:notes xmlns:p=\"{p}\" xmlns:a=\"{a}\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:u=\"urn:unsupported\"><mc:AlternateContent><mc:Choice Requires=\"u\"><u:active/></mc:Choice><mc:Fallback><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>Strict note</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></mc:Fallback></mc:AlternateContent><p:clrMapOvr/></p:notes>").into_bytes());
    notes.rels_mut().add_relationship(
        conformance.slide_rel().into(),
        "../slides/slide1.xml".into(),
        "rIdBack".into(),
        false,
    );
    notes.rels_mut().add_relationship(
        conformance.notes_master_rel().into(),
        "../notesMasters/notesMaster1.xml".into(),
        "rIdMaster".into(),
        false,
    );
    package.add_part(Box::new(notes));
    package.add_part(Box::new(BlobPart::new(
        PackURI::new("/ppt/theme/theme2.xml").unwrap(),
        THEME_CT.into(),
        format!("<a:theme xmlns:a=\"{a}\" name=\"Notes\"/>").into_bytes(),
    )));
    (package, presentation)
}
#[test]
fn strict_mce_graph_round_trips_and_projects_text() {
    let (mut package, name) = synthetic(Conformance::Strict);
    let graph = load(&package, &name).unwrap().unwrap();
    assert_eq!(
        graph.slides[0].text().unwrap().as_deref(),
        Some("Strict note")
    );
    put(&mut package, &name, graph).unwrap();
    assert_eq!(
        load(&package, &name).unwrap().unwrap().slides[0]
            .text()
            .unwrap()
            .as_deref(),
        Some("Strict note")
    );
}

#[test]
fn every_presentation_main_profile_accepts_the_same_notes_graph() {
    for content_type in [
        ct::PML_PRESENTATION_MAIN,
        ct::PML_SLIDESHOW_MAIN,
        ct::PML_TEMPLATE_MAIN,
        ct::PML_PRES_MACRO_MAIN,
        ct::PML_SLIDESHOW_MACRO_MAIN,
        ct::PML_TEMPLATE_MACRO_MAIN,
    ] {
        let (mut package, name) = synthetic(Conformance::Transitional);
        package
            .get_part_mut(&name)
            .and_then(|part| part.set_content_type(content_type.to_owned()))
            .unwrap();
        assert!(load(&package, &name).unwrap().is_some());
    }
}

#[test]
fn strict_text_writer_validates_and_round_trips_replacement() {
    let (mut package, name) = synthetic(Conformance::Strict);
    let mut graph = load(&package, &name).unwrap().unwrap();
    let xml = write_text_with(graph.conformance(), "Updated strict note").unwrap();
    let encoded = std::str::from_utf8(&xml).unwrap();
    assert!(encoded.contains(PS));
    assert!(encoded.contains(AS));
    assert!(encoded.contains(RS));

    graph.slides_mut()[0].replace_xml(xml);
    put(&mut package, &name, graph).unwrap();

    let graph = load(&package, &name).unwrap().unwrap();
    assert_eq!(graph.conformance(), Conformance::Strict);
    assert_eq!(
        graph.slides()[0].text().unwrap().as_deref(),
        Some("Updated strict note")
    );
}

#[test]
fn strict_and_transitional_notes_removal_is_idempotent() {
    for conformance in [Conformance::Transitional, Conformance::Strict] {
        let (mut package, name) = synthetic(conformance);
        let slide = PackURI::new("/ppt/slides/slide1.xml").unwrap();
        let notes = PackURI::new("/ppt/notesSlides/notesSlide1.xml").unwrap();
        let master = PackURI::new("/ppt/notesMasters/notesMaster1.xml").unwrap();
        let theme = PackURI::new("/ppt/theme/theme2.xml").unwrap();

        assert!(remove(&mut package, &name, &slide).unwrap());
        assert!(!package.contains_part(&notes));
        assert!(package.contains_part(&master));
        assert!(package.contains_part(&theme));
        assert!(
            package
                .get_part(&slide)
                .unwrap()
                .rels()
                .iter()
                .all(|relationship| !is_notes_slide_rel(relationship.reltype()))
        );
        assert!(load(&package, &name).unwrap().unwrap().slides.is_empty());
        assert!(!remove(&mut package, &name, &slide).unwrap());
        assert_eq!(clear(&mut package, &name).unwrap(), 0);
    }
}

#[test]
fn removal_uses_the_actual_stored_part_name_after_case_folded_lookup() {
    let (mut package, name) = synthetic(Conformance::Transitional);
    let canonical = PackURI::new("/ppt/notesSlides/notesSlide1.xml").unwrap();
    let mixed_case = PackURI::new("/PPT/NOTESSLIDES/NOTESSLIDE1.XML").unwrap();
    let data = package.get_part(&canonical).unwrap().blob().to_vec();
    assert!(package.remove_part(&canonical));
    let mut notes = BlobPart::new(mixed_case.clone(), ct::PML_NOTES_SLIDE.into(), data);
    notes.rels_mut().add_relationship(
        rt::SLIDE.into(),
        "../slides/slide1.xml".into(),
        "rIdBack".into(),
        false,
    );
    notes.rels_mut().add_relationship(
        rt::NOTES_MASTER.into(),
        "../notesMasters/notesMaster1.xml".into(),
        "rIdMaster".into(),
        false,
    );
    package.add_part(Box::new(notes));

    let slide = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    assert!(remove(&mut package, &name, &slide).unwrap());
    assert!(!package.contains_part(&mixed_case));
    assert!(load(&package, &name).unwrap().unwrap().slides.is_empty());
}

#[test]
fn unexpected_inbound_edge_rejects_removal_before_mutation() {
    let (mut package, name) = synthetic(Conformance::Transitional);
    let slide = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    let notes = PackURI::new("/ppt/notesSlides/notesSlide1.xml").unwrap();
    let observer_name = PackURI::new("/ppt/custom/observer.xml").unwrap();
    let mut observer = BlobPart::new(
        observer_name,
        "application/xml".into(),
        b"<observer/>".to_vec(),
    );
    observer.rels_mut().add_relationship(
        "urn:test:observes-notes".into(),
        "../notesSlides/notesSlide1.xml".into(),
        "rIdObserver".into(),
        false,
    );
    package.add_part(Box::new(observer));

    let before_parts = package.part_count();
    let before_relationships = package.get_part(&slide).unwrap().rels().len();
    let error = remove(&mut package, &name, &slide).unwrap_err();
    assert!(
        error
            .to_string()
            .contains("unexpected inbound relationship")
    );
    assert_eq!(package.part_count(), before_parts);
    assert_eq!(
        package.get_part(&slide).unwrap().rels().len(),
        before_relationships
    );
    assert!(package.contains_part(&notes));
}

#[test]
fn malformed_graph_rejects_clear_before_mutation() {
    let (mut package, name) = synthetic(Conformance::Transitional);
    let slide = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    let notes = PackURI::new("/ppt/notesSlides/notesSlide1.xml").unwrap();
    package
        .get_part_mut(&notes)
        .unwrap()
        .set_blob(format!("<p:wrong xmlns:p=\"{P}\"/>").into_bytes());

    let before_parts = package.part_count();
    let before_relationships = package.get_part(&slide).unwrap().rels().len();
    assert!(clear(&mut package, &name).is_err());
    assert_eq!(package.part_count(), before_parts);
    assert_eq!(
        package.get_part(&slide).unwrap().rels().len(),
        before_relationships
    );
    assert!(package.contains_part(&notes));
}

#[test]
fn rejects_external_wrong_root_outbound_orphan_and_caps_before_mutation() {
    let (mut package, name) = synthetic(Conformance::Transitional);
    let notes = PackURI::new("/ppt/notesSlides/notesSlide1.xml").unwrap();
    {
        let part = package.get_part_mut(&notes).unwrap();
        part.rels_mut().remove("rIdBack");
        part.rels_mut().add_relationship(
            rt::SLIDE.into(),
            "https://example.invalid/slide".into(),
            "rIdBack".into(),
            true,
        );
    }
    assert!(load(&package, &name).is_err());
    let (mut package, name) = synthetic(Conformance::Transitional);
    package
        .get_part_mut(&PackURI::new("/ppt/notesMasters/notesMaster1.xml").unwrap())
        .unwrap()
        .set_blob(format!("<p:wrong xmlns:p=\"{P}\"/>").into_bytes());
    assert!(load(&package, &name).is_err());
    let (mut package, name) = synthetic(Conformance::Transitional);
    package
        .get_part_mut(&notes)
        .unwrap()
        .rels_mut()
        .add_relationship(
            rt::IMAGE.into(),
            "../media/image1.png".into(),
            "rIdImage".into(),
            false,
        );
    assert!(load(&package, &name).is_err());
    let (mut package, name) = synthetic(Conformance::Transitional);
    package.add_part(Box::new(BlobPart::new(
        PackURI::new("/ppt/notesSlides/orphan.xml").unwrap(),
        ct::PML_NOTES_SLIDE.into(),
        format!("<p:notes xmlns:p=\"{P}\"><p:cSld/></p:notes>").into_bytes(),
    )));
    assert!(load(&package, &name).is_err());
    let (mut package, name) = synthetic(Conformance::Transitional);
    let mut graph = load(&package, &name).unwrap().unwrap();
    graph.slides[0].data = vec![b' '; MAX_NOTES_XML + 1];
    let before = package.get_part(&name).unwrap().blob().to_vec();
    assert!(put(&mut package, &name, graph).is_err());
    assert_eq!(package.get_part(&name).unwrap().blob(), before);
}

#[test]
fn source_checked_transaction_no_op_is_byte_stable() {
    let (mut package, name) = synthetic(Conformance::Transitional);
    let before = package.get_part(&name).unwrap().blob().to_vec();
    let snapshot = Snapshot::load(&package, &name).unwrap().unwrap();
    let commit = snapshot.edit().commit().unwrap();
    assert!(!commit.is_changed());
    commit.patch().apply(&mut package).unwrap();
    assert_eq!(package.get_part(&name).unwrap().blob(), before);
}

#[test]
fn source_checked_text_edit_preserves_opaque_xml_and_relationships() {
    let (mut package, name) = synthetic(Conformance::Transitional);
    let opaque_name = PackURI::new("/ppt/custom/notes.bin").unwrap();
    package.add_part(Box::new(BlobPart::new(
        opaque_name.clone(),
        "application/octet-stream".into(),
        b"opaque".to_vec(),
    )));
    let notes_name = PackURI::new("/ppt/notesSlides/notesSlide1.xml").unwrap();
    package
        .get_part_mut(&notes_name)
        .unwrap()
        .rels_mut()
        .add_relationship(
            "urn:vendor:notes-extension".into(),
            "../custom/notes.bin".into(),
            "rIdOpaque".into(),
            false,
        );

    let snapshot = Snapshot::load(&package, &name).unwrap().unwrap();
    assert!(
        snapshot.slides()[0]
            .xml()
            .windows(b"AlternateContent".len())
            .any(|w| w == b"AlternateContent")
    );
    let mut edit = snapshot.edit();
    assert!(edit.set_text(0, "Updated <speaker> & note").unwrap());
    let commit = edit.commit().unwrap();
    commit.patch().apply(&mut package).unwrap();

    let notes = package.get_part(&notes_name).unwrap();
    assert!(
        notes
            .blob()
            .windows(b"AlternateContent".len())
            .any(|w| w == b"AlternateContent")
    );
    assert!(
        std::str::from_utf8(notes.blob())
            .unwrap()
            .contains("Updated &lt;speaker&gt; &amp; note")
    );
    assert_eq!(
        notes.rels().get("rIdOpaque").unwrap().target_ref(),
        "../custom/notes.bin"
    );
}

#[test]
fn source_checked_transaction_rejects_stale_source_atomically() {
    let (mut package, name) = synthetic(Conformance::Transitional);
    let snapshot = Snapshot::load(&package, &name).unwrap().unwrap();
    let mut edit = snapshot.edit();
    edit.set_text(0, "Changed").unwrap();
    let commit = edit.commit().unwrap();
    let notes_name = PackURI::new("/ppt/notesSlides/notesSlide1.xml").unwrap();
    package
        .get_part_mut(&notes_name)
        .unwrap()
        .set_blob(b"<stale/>".to_vec());
    let before = package.get_part(&notes_name).unwrap().blob().to_vec();
    assert!(commit.patch().apply(&mut package).is_err());
    assert_eq!(package.get_part(&notes_name).unwrap().blob(), before);
}

#[test]
fn source_checked_transaction_inverse_restores_master_and_theme() {
    let (mut package, name) = synthetic(Conformance::Transitional);
    let snapshot = Snapshot::load(&package, &name).unwrap().unwrap();
    let mut edit = snapshot.edit();
    edit.replace_master_xml(
        format!("<p:notesMaster xmlns:p=\"{P}\"><p:cSld/><p:extLst/></p:notesMaster>").into_bytes(),
    )
    .unwrap();
    edit.replace_theme_xml(format!("<a:theme xmlns:a=\"{A}\"/>").into_bytes())
        .unwrap();
    let commit = edit.commit().unwrap();
    commit.patch().apply(&mut package).unwrap();
    assert_eq!(
        Snapshot::load(&package, &name)
            .unwrap()
            .unwrap()
            .master()
            .xml(),
        commit.snapshot().master().xml()
    );
    commit.patch().undo(&mut package).unwrap();
    assert!(
        Snapshot::load(&package, &name)
            .unwrap()
            .unwrap()
            .same_source(&snapshot)
    );
}

#[test]
fn source_bound_removal_rejects_a_stale_bypass_atomically() {
    let (mut package, presentation) = synthetic(Conformance::Transitional);
    let source = Snapshot::load(&package, &presentation).unwrap().unwrap();
    let notes_name = PackURI::new("/ppt/notesSlides/notesSlide1.xml").unwrap();
    let slide_name = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    package
        .get_part_mut(&notes_name)
        .unwrap()
        .set_blob(write_text("stale source").unwrap());

    let error = remove_checked(&mut package, &source, &slide_name).unwrap_err();
    assert!(matches!(error, crate::Error::Invalid(message) if message.contains("stale")));
    assert!(package.contains_part(&notes_name));
}
