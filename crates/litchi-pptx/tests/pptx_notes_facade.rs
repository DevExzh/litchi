#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, PackURI, Part, TargetMode};
use litchi_pptx::{Error, Package, notes};
use tempfile::NamedTempFile;

const SLIDE: &str = "/ppt/slides/slide1.xml";
const NOTES: &str = "/ppt/notesSlides/notesSlide1.xml";
const CLASSIC_COMMENTS: &str = "/ppt/comments/comment1.xml";
const MODERN_COMMENTS: &str = "/ppt/comments/modernComment1.xml";
const CHILD: &str = "/ppt/media/notes-child.bin";
const GRANDCHILD: &str = "/ppt/embeddings/notes-child.bin";
const MAX_OWNED_PARTS: usize = 4_096;
const IMAGE_REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const PACKAGE_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package";

#[test]
fn package_notes_crud_uses_semantic_slide_selectors() {
    let mut package = opened_deck(&[("Overview", "First note"), ("Appendix", "Second note")]);
    let mut graph = package.notes().unwrap().unwrap();
    graph.slides_mut()[0].replace_xml(notes::write_text("Revised note").unwrap());
    package.put_notes(graph).unwrap();
    assert_eq!(
        package.notes().unwrap().unwrap().slides()[0]
            .text()
            .unwrap(),
        Some("Revised note".into())
    );

    assert!(package.remove_notes("Overview").unwrap());
    assert!(!package.remove_notes("Overview").unwrap());
    assert_eq!(package.notes().unwrap().unwrap().slides().len(), 1);
    assert!(package.remove_notes(1usize).unwrap());
    assert_eq!(package.clear_notes().unwrap(), 0);
    assert!(package.notes().unwrap().unwrap().slides().is_empty());
}

#[test]
fn facade_publishes_source_checked_lossless_edits_and_rejects_stale_patches() {
    let mut package = opened_deck(&[("Overview", "Source note")]);
    package
        .edit_opc(|opc| {
            let notes_name = PackURI::new(NOTES).unwrap();
            let child_name = PackURI::new(CHILD).unwrap();
            opc.try_add_part(Box::new(BlobPart::new(
                child_name.clone(),
                "application/octet-stream".into(),
                b"opaque-child".to_vec(),
            )))?;
            opc.get_part_mut(&notes_name)?
                .rels_mut()
                .try_add_relationship(
                    IMAGE_REL.into(),
                    child_name.relative_ref(notes_name.base_uri()),
                    "rIdOpaque".into(),
                    TargetMode::Internal,
                )?;
            Ok(())
        })
        .unwrap();

    let source = package.notes_snapshot().unwrap().unwrap();
    let mut edit = source.edit();
    assert!(edit.set_text(0, "Committed note").unwrap());
    let commit = edit.commit().unwrap();
    let patch = commit.patch().clone();
    package.apply_notes_commit(commit).unwrap();
    assert_eq!(
        package.notes().unwrap().unwrap().slides()[0]
            .text()
            .unwrap(),
        Some("Committed note".into())
    );
    assert_eq!(
        package
            .opc()
            .unwrap()
            .get_part(&PackURI::new(CHILD).unwrap())
            .unwrap()
            .blob(),
        b"opaque-child"
    );

    let error = package.apply_notes_patch(&patch).unwrap_err();
    assert!(matches!(error, Error::Invalid(message) if message.contains("stale")));
    assert_eq!(
        package.notes().unwrap().unwrap().slides()[0]
            .text()
            .unwrap(),
        Some("Committed note".into())
    );
}

#[test]
fn notes_removal_collects_only_exclusive_descendants_and_preserves_slide_comments() {
    let mut package = opened_deck(&[("Overview", "Source note")]);
    add_comment_and_notes_children(&mut package, false);

    assert!(package.remove_notes("Overview").unwrap());
    let opc = package.opc().unwrap();
    assert!(!opc.contains_part(&PackURI::new(NOTES).unwrap()));
    assert!(!opc.contains_part(&PackURI::new(CHILD).unwrap()));
    assert!(!opc.contains_part(&PackURI::new(GRANDCHILD).unwrap()));
    assert!(opc.contains_part(&PackURI::new(CLASSIC_COMMENTS).unwrap()));
    assert!(opc.contains_part(&PackURI::new(MODERN_COMMENTS).unwrap()));
    let slide = opc.get_part(&PackURI::new(SLIDE).unwrap()).unwrap();
    assert!(
        slide
            .rels()
            .iter()
            .any(|relationship| relationship.reltype() == rt::COMMENTS)
    );
    assert!(
        slide
            .rels()
            .iter()
            .any(|relationship| relationship.reltype()
                == litchi_pptx::modern_comments::MODERN_COMMENT_RELATIONSHIP_TYPE)
    );
}

#[test]
fn shared_notes_descendant_refuses_deletion_without_partial_mutation() {
    let mut package = opened_deck(&[("Overview", "Source note")]);
    add_comment_and_notes_children(&mut package, true);
    let before = package.notes().unwrap().unwrap();

    let error = package.remove_notes("Overview").unwrap_err();
    assert!(matches!(error, Error::Invalid(message) if message.contains("shared inbound")));
    assert_eq!(package.notes().unwrap().unwrap(), before);
    assert!(
        package
            .opc()
            .unwrap()
            .contains_part(&PackURI::new(CHILD).unwrap())
    );
}

#[test]
fn unknown_notes_relationship_refuses_deletion_atomically() {
    let mut package = opened_deck(&[("Overview", "Source note")]);
    package
        .edit_opc(|opc| {
            let notes_name = PackURI::new(NOTES).unwrap();
            let child_name = PackURI::new("/ppt/media/unknown.bin").unwrap();
            opc.try_add_part(Box::new(BlobPart::new(
                child_name.clone(),
                "application/octet-stream".into(),
                Vec::new(),
            )))?;
            opc.get_part_mut(&notes_name)?
                .rels_mut()
                .try_add_relationship(
                    "urn:vendor:unknown-owned".into(),
                    child_name.relative_ref(notes_name.base_uri()),
                    "rIdUnknown".into(),
                    TargetMode::Internal,
                )?;
            Ok(())
        })
        .unwrap();
    let before = package.notes().unwrap().unwrap();

    let error = package.remove_notes("Overview").unwrap_err();
    assert!(
        matches!(error, Error::Invalid(message) if message.contains("unknown relationship type"))
    );
    assert_eq!(package.notes().unwrap().unwrap(), before);
}

#[test]
fn notes_descendant_limit_rejects_n_plus_one_before_mutation() {
    let mut package = opened_deck(&[("Overview", "Source note"), ("Appendix", "Other note")]);
    package
        .edit_opc(|opc| {
            for index in 0..=MAX_OWNED_PARTS {
                let notes_name =
                    PackURI::new(format!("/ppt/notesSlides/notesSlide{}.xml", index % 2 + 1))
                        .unwrap();
                let child_name = PackURI::new(format!("/ppt/media/limit-{index}.bin")).unwrap();
                opc.try_add_part(Box::new(BlobPart::new(
                    child_name.clone(),
                    "application/octet-stream".into(),
                    Vec::new(),
                )))?;
                opc.get_part_mut(&notes_name)?
                    .rels_mut()
                    .try_add_relationship(
                        IMAGE_REL.into(),
                        child_name.relative_ref(notes_name.base_uri()),
                        format!("rIdLimit{index}"),
                        TargetMode::Internal,
                    )?;
            }
            Ok(())
        })
        .unwrap();

    let error = package.clear_notes().unwrap_err();
    assert!(matches!(
        error,
        Error::Limit {
            resource: "notes-owned related parts",
            limit: MAX_OWNED_PARTS
        }
    ));
    assert!(
        package
            .opc()
            .unwrap()
            .contains_part(&PackURI::new(NOTES).unwrap())
    );
    assert!(
        package
            .opc()
            .unwrap()
            .contains_part(&PackURI::new("/ppt/media/limit-4096.bin").unwrap())
    );
}

#[test]
fn authored_notes_xml_and_referenced_master_are_compact() {
    for xml in [
        notes::write_text("Compact").unwrap(),
        notes::master_xml().as_bytes().to_vec(),
    ] {
        assert!(!xml.contains(&b'\n'));
        assert!(!xml.contains(&b'\r'));
        assert!(!xml.windows(3).any(|window| window == b"> <"));
    }
}

fn opened_deck(slides: &[(&str, &str)]) -> Package {
    let output = NamedTempFile::with_suffix(".pptx").unwrap();
    let mut package = Package::new().unwrap();
    for (name, note) in slides {
        let slide = package.presentation_mut().unwrap().add_slide().unwrap();
        slide.set_title(name);
        slide.set_notes(note);
    }
    package.save(output.path()).unwrap();
    let mut package = Package::open(output.path()).unwrap();
    for (index, (name, _)) in slides.iter().enumerate() {
        let slide_name = PackURI::new(format!("/ppt/slides/slide{}.xml", index + 1)).unwrap();
        package
            .edit_opc(|opc| {
                let slide = opc.get_part_mut(&slide_name)?;
                let xml = std::str::from_utf8(slide.blob()).unwrap();
                let named = if xml.contains("<p:cSld>") {
                    xml.replacen("<p:cSld>", &format!(r#"<p:cSld name="{name}">"#), 1)
                } else {
                    let marker = " name=\"";
                    let root = xml.find("<p:cSld ").unwrap();
                    let end = root + xml[root..].find('>').unwrap();
                    let start = root + xml[root..end].find(marker).unwrap() + marker.len();
                    let finish = start + xml[start..end].find('"').unwrap();
                    let mut named = xml.to_owned();
                    named.replace_range(start..finish, name);
                    named
                };
                slide.set_blob(named.into_bytes());
                Ok(())
            })
            .unwrap();
    }
    package
}

fn add_comment_and_notes_children(package: &mut Package, shared: bool) {
    package
        .edit_opc(|opc| {
            let slide_name = PackURI::new(SLIDE).unwrap();
            let notes_name = PackURI::new(NOTES).unwrap();
            let classic_name = PackURI::new(CLASSIC_COMMENTS).unwrap();
            let modern_name = PackURI::new(MODERN_COMMENTS).unwrap();
            let child_name = PackURI::new(CHILD).unwrap();
            let grandchild_name = PackURI::new(GRANDCHILD).unwrap();

            opc.try_add_part(Box::new(BlobPart::new(
                classic_name.clone(),
                ct::PML_COMMENTS.into(),
                b"<p:cmLst xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"/>".to_vec(),
            )))?;
            opc.try_add_part(Box::new(BlobPart::new(
                modern_name.clone(),
                litchi_pptx::modern_comments::MODERN_COMMENT_CONTENT_TYPE.into(),
                b"<p188:cmLst xmlns:p188=\"http://schemas.microsoft.com/office/powerpoint/2018/8/main\"/>".to_vec(),
            )))?;
            opc.try_add_part(Box::new(BlobPart::new(
                child_name.clone(),
                "application/octet-stream".into(),
                b"child".to_vec(),
            )))?;
            let mut grandchild = BlobPart::new(
                grandchild_name.clone(),
                "application/octet-stream".into(),
                b"grandchild".to_vec(),
            );
            grandchild.rels_mut().try_add_relationship(
                PACKAGE_REL.into(),
                child_name.relative_ref(grandchild_name.base_uri()),
                "rIdCycle".into(),
                TargetMode::Internal,
            )?;
            opc.try_add_part(Box::new(grandchild))?;

            let slide = opc.get_part_mut(&slide_name)?;
            slide.rels_mut().try_add_relationship(
                rt::COMMENTS.into(),
                classic_name.relative_ref(slide_name.base_uri()),
                "rIdClassicComment".into(),
                TargetMode::Internal,
            )?;
            slide.rels_mut().try_add_relationship(
                litchi_pptx::modern_comments::MODERN_COMMENT_RELATIONSHIP_TYPE.into(),
                modern_name.relative_ref(slide_name.base_uri()),
                "rIdModernComment".into(),
                TargetMode::Internal,
            )?;
            if shared {
                slide.rels_mut().try_add_relationship(
                    IMAGE_REL.into(),
                    child_name.relative_ref(slide_name.base_uri()),
                    "rIdSharedChild".into(),
                    TargetMode::Internal,
                )?;
            }

            let notes = opc.get_part_mut(&notes_name)?;
            notes.rels_mut().try_add_relationship(
                IMAGE_REL.into(),
                child_name.relative_ref(notes_name.base_uri()),
                "rIdChild".into(),
                TargetMode::Internal,
            )?;
            notes.rels_mut().try_add_relationship(
                PACKAGE_REL.into(),
                grandchild_name.relative_ref(notes_name.base_uri()),
                "rIdGrandchild".into(),
                TargetMode::Internal,
            )?;
            Ok(())
        })
        .unwrap();
}
