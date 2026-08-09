#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use litchi_opc::{BlobPart, OpcPackage, PackURI};

use super::*;

const MAIN_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml";
const PC: &str = "http://schemas.microsoft.com/office/powerpoint/2013/main/command";
const P: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

fn package() -> OpcPackage {
    let mut package = OpcPackage::new();
    package.rels_mut().add_relationship(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument".into(),
        "ppt/presentation.xml".into(),
        "rId1".into(),
        false,
    );
    package.add_part(Box::new(BlobPart::new(
        PackURI::new("/ppt/presentation.xml").unwrap(),
        MAIN_CONTENT_TYPE.into(),
        b"<p:presentation/>".to_vec(),
    )));
    package
}

fn value() -> Part {
    let xml = format!(
        r#"<pc:chgInfo xmlns:pc="{PC}" xmlns:p="{P}" xmlns:v="urn:vendor" xmlns:r="{R}">
  <pc:docChgLst>
    <pc:chgData name="Ada" userId="ada@example.test" email="ada@example.test"><a:extLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:ext uri="urn:review"><v:opaque r:id="rIdNeverFetched" href="https://example.invalid/not-opened"/></a:ext></a:extLst></pc:chgData>
    <pc:docChg chg="addSld"><pc:docMkLst><pc:docMk/></pc:docMkLst><p:extLst><p:ext uri="urn:future"><v:payload r:id="rIdAlsoNeverFetched"/></p:ext></p:extLst></pc:docChg>
    <p:extLst><p:ext uri="urn:list"><v:list><v:unknown>keep</v:unknown></v:list></p:ext></p:extLst>
  </pc:docChgLst>
</pc:chgInfo>"#
    );
    Part {
        relationship_id: "rIdChanges".into(),
        part_name: "/ppt/changesInfo.xml".into(),
        changes_information: Info::parse(xml.as_bytes()).unwrap(),
    }
}

fn stored_package() -> OpcPackage {
    let mut package = package();
    store(&mut package, &value()).unwrap();
    package
}

#[test]
fn no_op_commit_and_patch_preserve_exact_source() {
    let mut package = stored_package();
    let before = load_snapshot(&package).unwrap().unwrap();
    let source = before.source_xml().to_vec();
    let commit = before.edit().commit().unwrap();
    assert!(!commit.changed());
    assert!(commit.patch().is_empty());
    assert_eq!(commit.snapshot().source_xml(), source.as_slice());

    let result = apply_commit(&mut package, commit).unwrap();
    assert_eq!(result.source_xml(), source.as_slice());
    assert_eq!(
        package
            .get_part(&PackURI::new("/ppt/changesInfo.xml").unwrap())
            .unwrap()
            .blob(),
        source.as_slice()
    );
}

#[test]
fn typed_edits_preserve_opaque_commands_and_relationship_looking_xml() {
    let mut package = stored_package();
    let snapshot = load_snapshot(&package).unwrap().unwrap();
    let mut edit = snapshot.edit();
    edit.edit_author(0, |author| {
        author.name = Some("Grace".into());
        Ok(())
    })
    .unwrap();
    assert!(
        edit.set_change_kinds(0, 0, vec![Kind::DeleteSlide])
            .unwrap()
    );
    let commit = edit.commit().unwrap();
    let output = String::from_utf8_lossy(commit.snapshot().source_xml());
    assert!(output.contains(r#"name="Grace""#));
    assert!(output.contains(r#"chg="delSld""#));
    assert!(output.contains(r#"r:id="rIdNeverFetched""#));
    assert!(output.contains(r#"r:id="rIdAlsoNeverFetched""#));
    assert!(output.contains("https://example.invalid/not-opened"));
    assert!(output.contains("<v:unknown>keep</v:unknown>"));

    let relationship_before = {
        let presentation = package.main_document_part().unwrap();
        let relationship = presentation.rels().get("rIdChanges").unwrap();
        (
            relationship.reltype().to_owned(),
            relationship.target_ref().to_owned(),
            relationship.is_external(),
        )
    };
    let after = apply_commit(&mut package, commit).unwrap();
    assert_eq!(
        after.info().change_lists[0].author.as_ref().unwrap().name,
        Some("Grace".into())
    );
    let presentation = package.main_document_part().unwrap();
    let relationship = presentation.rels().get("rIdChanges").unwrap();
    assert_eq!(relationship.reltype(), relationship_before.0.as_str());
    assert_eq!(relationship.target_ref(), relationship_before.1.as_str());
    assert_eq!(relationship.is_external(), relationship_before.2);
}

#[test]
fn stale_patch_rejection_is_atomic_and_inverse_restores_original_bytes() {
    let mut package = stored_package();
    let original = load_snapshot(&package).unwrap().unwrap();
    let mut edit = original.edit();
    edit.edit_author(0, |author| {
        author.name = Some("Grace".into());
        Ok(())
    })
    .unwrap();
    let commit = edit.commit().unwrap();
    let patch = commit.patch().clone();

    let part = package
        .get_part_mut(&PackURI::new("/ppt/changesInfo.xml").unwrap())
        .unwrap();
    let mut stale = part.blob().to_vec();
    stale.extend_from_slice(b" ");
    part.set_blob(stale);
    let stale_source = package
        .get_part(&PackURI::new("/ppt/changesInfo.xml").unwrap())
        .unwrap()
        .blob()
        .to_vec();
    assert!(patch.apply(&mut package).is_err());
    assert_eq!(
        package
            .get_part(&PackURI::new("/ppt/changesInfo.xml").unwrap())
            .unwrap()
            .blob(),
        stale_source.as_slice()
    );

    let mut package = stored_package();
    let committed = patch.apply(&mut package).unwrap();
    assert_eq!(
        committed.info().change_lists[0]
            .author
            .as_ref()
            .unwrap()
            .name,
        Some("Grace".into())
    );
    let inverse = patch.inverse();
    let restored = inverse.apply(&mut package).unwrap();
    assert_eq!(restored.source_xml(), original.source_xml());
}

#[test]
fn failed_typed_edits_do_not_mutate_the_staged_snapshot() {
    let snapshot = load_snapshot(&stored_package()).unwrap().unwrap();
    let mut edit = snapshot.edit();
    let before = edit.snapshot().clone();
    assert!(
        edit.set_change_kinds(0, 0, vec![Kind::AddSlide, Kind::AddSlide])
            .is_err()
    );
    assert_eq!(edit.snapshot(), &before);
    assert!(
        edit.edit_author(0, |author| {
            author.change_id = Some("not-a-guid".into());
            Ok(())
        })
        .is_err()
    );
    assert_eq!(edit.snapshot(), &before);
}

#[test]
fn create_read_update_delete_is_atomic_and_rejects_unexpected_inbound_edges() {
    let mut package = package();
    assert!(load_snapshot(&package).unwrap().is_none());
    store(&mut package, &value()).unwrap();
    assert!(load_snapshot(&package).unwrap().is_some());
    let removed = remove(&mut package).unwrap().unwrap();
    assert_eq!(removed.part_name, "/ppt/changesInfo.xml");
    assert!(load_snapshot(&package).unwrap().is_none());

    let mut package = stored_package();
    let target = PackURI::new("/ppt/changesInfo.xml").unwrap();
    package.add_part(Box::new(BlobPart::new(
        PackURI::new("/ppt/other.xml").unwrap(),
        "application/xml".into(),
        Vec::new(),
    )));
    package
        .get_part_mut(&PackURI::new("/ppt/other.xml").unwrap())
        .unwrap()
        .rels_mut()
        .add_relationship(
            RELATIONSHIP_TYPE.into(),
            target.relative_ref("/ppt/"),
            "rIdOther".into(),
            false,
        );
    assert!(load_snapshot(&package).is_err());
    let part_count = package.part_count();
    assert!(remove(&mut package).is_err());
    assert_eq!(package.part_count(), part_count);
}
