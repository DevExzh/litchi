//! Eager per-slide notes snapshots and source parity.

#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "fixture construction and assertions intentionally panic"
)]

use std::io::Cursor;

use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::part::BlobPart;
use litchi_opc::{OpcPackage, PackURI, PackageWriter};
use litchi_pptx::presentation::related::raw;
use litchi_pptx::{Package, SourceBackedPresentation};

fn deck() -> Vec<u8> {
    let mut builder = litchi_pptx::builder::Builder::new().expect("package");
    let first = builder
        .presentation()
        .expect("presentation")
        .add_slide()
        .expect("slide");
    first.set_title("Overview");
    first.add_text_box("Body", 10, 20, 30, 40);
    first.set_notes("First note");
    builder
        .presentation()
        .expect("presentation")
        .add_slide()
        .expect("slide")
        .set_title("Details");
    let package = builder.build().expect("package");
    package.to_bytes().expect("bytes")
}

fn rewrite(bytes: &[u8], edit: impl FnOnce(&mut OpcPackage)) -> Vec<u8> {
    let mut opc =
        OpcPackage::from_bytes(bytes, litchi_opc::SourceLimits::default()).expect("OPC package");
    edit(&mut opc);
    PackageWriter::to_bytes(&opc).expect("rewritten package bytes")
}

fn remove_relationship(opc: &mut OpcPackage, part_uri: &str, relationship_type: &str) {
    let uri = PackURI::new(part_uri).expect("part URI");
    let relationship_id = opc
        .get_part(&uri)
        .expect("part")
        .rels()
        .iter()
        .find(|relationship| relationship.reltype() == relationship_type)
        .expect("relationship")
        .r_id()
        .to_owned();
    opc.get_part_mut(&uri)
        .expect("mutable part")
        .rels_mut()
        .remove(&relationship_id);
}

#[test]
fn eager_notes_match_source_and_survive_owner_drop() {
    let bytes = deck();
    let eager = {
        let package = Package::from_bytes(&bytes).expect("eager package");
        let presentation = package.presentation().expect("presentation");
        let slide = presentation.slide(0).expect("slide lookup").expect("slide");
        let notes = slide.notes().expect("notes read").expect("notes");
        assert_eq!(notes.text(), Some("First note"));
        assert!(slide.notes().expect("second notes read").is_some());
        assert!(
            presentation
                .slide(1)
                .expect("second slide lookup")
                .expect("second slide")
                .notes()
                .expect("absent notes read")
                .is_none()
        );
        notes
    };

    let source = SourceBackedPresentation::from_reader(Cursor::new(&bytes)).expect("source");
    let source_notes = source
        .slide(0)
        .expect("source slide")
        .notes()
        .expect("source notes read")
        .expect("source notes");
    assert_eq!(eager.text(), source_notes.text());
    assert_eq!(eager.scene().len(), source_notes.scene().len());
    assert_eq!(raw::part_uri(&eager), raw::part_uri(&source_notes));
    assert_eq!(
        raw::relationship_id(&eager),
        raw::relationship_id(&source_notes)
    );
    assert_eq!(raw::data(&eager), raw::data(&source_notes));
}

#[test]
fn eager_notes_do_not_require_the_presentation_notes_graph() {
    let bytes = deck();
    let mut opc =
        OpcPackage::from_bytes(&bytes, litchi_opc::SourceLimits::default()).expect("OPC package");
    let presentation_uri = PackURI::new("/ppt/presentation.xml").expect("presentation URI");
    let master_uri = opc
        .get_part(&presentation_uri)
        .expect("presentation part")
        .rels()
        .iter()
        .find(|relationship| relationship.reltype() == rt::NOTES_MASTER)
        .expect("notes master relationship")
        .target_partname()
        .expect("notes master target");
    let unrelated_uri = PackURI::new("/ppt/notesMasters/unrelated.xml").expect("unrelated URI");
    opc.add_part(Box::new(BlobPart::new(
        unrelated_uri,
        ct::PML_NOTES_MASTER.to_owned(),
        b"<broken-notes-master/>".to_vec(),
    )));
    let unrelated_broken = PackageWriter::to_bytes(&opc).expect("unrelated package bytes");

    let package = Package::from_bytes(&unrelated_broken).expect("eager package");
    let presentation = package.presentation().expect("presentation");
    let slide = presentation.slide(0).expect("slide lookup").expect("slide");
    assert!(slide.notes().expect("leaf notes read").is_some());

    let mut selected_broken = OpcPackage::from_bytes(&bytes, litchi_opc::SourceLimits::default())
        .expect("selected package");
    selected_broken
        .get_part_mut(&master_uri)
        .expect("notes master part")
        .set_blob(b"<broken-notes-master/>".to_vec());
    let selected_broken = PackageWriter::to_bytes(&selected_broken).expect("selected bytes");
    let package = Package::from_bytes(&selected_broken).expect("selected eager package");
    let presentation = package.presentation().expect("presentation");
    let slide = presentation.slide(0).expect("slide lookup").expect("slide");
    assert!(slide.notes().is_err());
    assert!(presentation.notes().is_err());
}

#[test]
fn eager_notes_match_source_closure_refusals() {
    let bytes = deck();
    let missing_backlink = rewrite(&bytes, |opc| {
        remove_relationship(opc, "/ppt/notesSlides/notesSlide1.xml", rt::SLIDE);
    });
    let missing_theme = rewrite(&bytes, |opc| {
        remove_relationship(opc, "/ppt/notesMasters/notesMaster1.xml", rt::THEME);
    });
    let catalog_mismatch = rewrite(&bytes, |opc| {
        let presentation = PackURI::new("/ppt/presentation.xml").expect("presentation URI");
        let relationship_id = opc
            .get_part(&presentation)
            .expect("presentation")
            .rels()
            .iter()
            .find(|relationship| relationship.reltype() == rt::NOTES_MASTER)
            .expect("notes-master relationship")
            .r_id()
            .to_owned();
        opc.get_part_mut(&presentation)
            .expect("mutable presentation")
            .rels_mut()
            .retarget(&relationship_id, "notesMasters/notesMaster2.xml".to_owned())
            .expect("retarget notes master");
    });

    for broken in [missing_backlink, missing_theme, catalog_mismatch] {
        let package = Package::from_bytes(&broken).expect("eager package");
        let presentation = package.presentation().expect("presentation");
        let slide = presentation.slide(0).expect("slide lookup").expect("slide");
        assert!(slide.notes().is_err());
    }
}
