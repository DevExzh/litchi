//! Raw manifest edits must retain their exact publication bytes through replay.

use std::sync::Arc;

use litchi_opc::{OpcPackage, PackURI};
use litchi_pptx::Package;

fn annotate(xml: &[u8]) -> Arc<Vec<u8>> {
    let mut bytes = xml.to_vec();
    bytes.extend_from_slice(b"<!-- retained producer markup -->");
    Arc::new(bytes)
}

fn content_types(opc: &mut OpcPackage) -> litchi_pptx::Result<()> {
    let limits = opc.read_limits();
    let before = opc.source_content_types(limits)?;
    opc.try_replace_content_types_bytes(before.bytes(), annotate(before.bytes()), limits)?;
    Ok(())
}

fn relationships(opc: &mut OpcPackage, owner: &str) -> litchi_pptx::Result<()> {
    let limits = opc.read_limits();
    let owner = PackURI::new(owner).map_err(litchi_pptx::Error::Uri)?;
    let before = opc.source_relationships(&owner, limits)?;
    opc.try_replace_relationships_bytes(&before, annotate(before.bytes()), true, limits)?;
    Ok(())
}

fn assert_exact_replay(operation: impl Fn(&mut OpcPackage) -> litchi_pptx::Result<()>) {
    let original = Package::new().unwrap().to_bytes().unwrap();
    let package = Package::from_bytes(&original).unwrap();
    let mut edit = package.edit_opc().unwrap();
    edit.update(operation).unwrap();
    let expected = litchi_opc::PackageWriter::to_bytes(edit.opc()).unwrap();
    assert_ne!(expected, original);
    let commit = edit.commit().unwrap();
    assert!(commit.diagnostics().changed());
    assert!(!commit.patch().is_noop());
    assert_eq!(commit.package().to_bytes().unwrap(), expected);
    let equal_copy = Package::from_bytes(&original).unwrap();
    assert_eq!(
        equal_copy
            .apply_opc(commit.patch())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
        expected,
    );
    let reopened = Package::from_bytes(&expected).unwrap();
    assert_eq!(
        reopened
            .apply_opc(&commit.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
        original,
    );
    assert_eq!(package.to_bytes().unwrap(), original);
}

#[test]
fn lexical_content_types_survive_commit_replay_and_inverse() {
    assert_exact_replay(content_types);
}

#[test]
fn lexical_root_relationships_survive_commit_replay_and_inverse() {
    assert_exact_replay(|opc| relationships(opc, "/"));
}

#[test]
fn lexical_part_relationships_survive_commit_replay_and_inverse() {
    assert_exact_replay(|opc| relationships(opc, "/ppt/presentation.xml"));
}

fn add_part(opc: &mut OpcPackage, name: &str) -> litchi_pptx::Result<()> {
    let name = PackURI::new(name).map_err(litchi_pptx::Error::Uri)?;
    opc.try_add_part(Box::new(litchi_opc::BlobPart::new(
        name,
        "application/octet-stream".to_owned(),
        vec![1, 2, 3],
    )))?;
    Ok(())
}

#[test]
fn structural_inverse_restores_producer_manifest_bytes() {
    let package = Package::new().unwrap();
    let mut edit = package.edit_opc().unwrap();
    edit.update(content_types).unwrap();
    let original = edit.commit().unwrap().package().to_bytes().unwrap();
    let package = Package::from_bytes(&original).unwrap();
    let mut edit = package.edit_opc().unwrap();
    edit.update(|opc| add_part(opc, "/custom/a.bin")).unwrap();
    let commit = edit.commit().unwrap();
    let reopened = Package::from_bytes(&commit.package().to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .apply_opc(&commit.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
        original,
    );
}

#[test]
fn structural_patch_composes_with_unrelated_part_addition() {
    let package = Package::new().unwrap();
    let mut first = package.edit_opc().unwrap();
    first.update(|opc| add_part(opc, "/custom/a.bin")).unwrap();
    let first = first.commit().unwrap();
    let mut second = package.edit_opc().unwrap();
    second.update(|opc| add_part(opc, "/custom/b.bin")).unwrap();
    let second = second.commit().unwrap();
    let combined = second.package().apply_opc(first.patch()).unwrap();
    let reopened = Package::from_bytes(&combined.package().to_bytes().unwrap()).unwrap();
    for name in ["/custom/a.bin", "/custom/b.bin"] {
        assert_eq!(
            reopened
                .opc()
                .unwrap()
                .get_part(&PackURI::new(name).unwrap())
                .unwrap()
                .blob(),
            &[1, 2, 3]
        );
    }
    let restored = reopened.apply_opc(&first.patch().inverse()).unwrap();
    assert_eq!(
        restored.package().to_bytes().unwrap(),
        second.package().to_bytes().unwrap()
    );
}

#[test]
fn lexical_patch_refuses_a_different_manifest() {
    let package = Package::new().unwrap();
    let mut edit = package.edit_opc().unwrap();
    edit.update(content_types).unwrap();
    let commit = edit.commit().unwrap();
    let error = commit.package().apply_opc(commit.patch()).unwrap_err();
    assert!(matches!(error, litchi_pptx::Error::PatchConflict));
}

#[test]
fn mixed_lexical_and_structural_patch_guards_both_directions() {
    let package = Package::new().unwrap();
    let mut mixed = package.edit_opc().unwrap();
    mixed
        .update(|opc| {
            add_part(opc, "/custom/a.bin")?;
            content_types(opc)
        })
        .unwrap();
    let mixed = mixed.commit().unwrap();
    let original = package.to_bytes().unwrap();
    let expected = mixed.package().to_bytes().unwrap();
    let copy = Package::from_bytes(&original).unwrap();
    assert_eq!(
        copy.apply_opc(mixed.patch())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
        expected,
    );
    let reopened = Package::from_bytes(&expected).unwrap();
    assert_eq!(
        reopened
            .apply_opc(&mixed.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
        original,
    );
    let mut unrelated = package.edit_opc().unwrap();
    unrelated
        .update(|opc| add_part(opc, "/custom/b.bin"))
        .unwrap();
    let unrelated = unrelated.commit().unwrap();
    assert!(matches!(
        unrelated.package().apply_opc(mixed.patch()),
        Err(litchi_pptx::Error::PatchConflict),
    ));
    let mut changed = mixed.package().edit_opc().unwrap();
    changed
        .update(|opc| add_part(opc, "/custom/b.bin"))
        .unwrap();
    let changed = changed.commit().unwrap();
    let original = changed.package().to_bytes().unwrap();
    assert!(matches!(
        changed.package().apply_opc(&mixed.patch().inverse()),
        Err(litchi_pptx::Error::PatchConflict),
    ));
    assert_eq!(changed.package().to_bytes().unwrap(), original);
}

#[test]
fn explicit_patch_matches_reopened_content_after_source_coverage_changes() {
    let package = Package::new().unwrap();
    let mut producer = package.edit_opc().unwrap();
    producer.update(content_types).unwrap();
    let producer = producer.commit().unwrap();
    let mut structural = producer.package().edit_opc().unwrap();
    structural
        .update(|opc| add_part(opc, "/custom/a.bin"))
        .unwrap();
    let structural = structural.commit().unwrap();
    let package = structural.package();
    // The source manifest no longer covers the typed graph; publication
    // uses canonical bytes, which are the recorded content on replay.
    assert!(
        package
            .opc()
            .unwrap()
            .retained_content_types_xml()
            .unwrap()
            .is_none()
    );
    let original = package.to_bytes().unwrap();
    let mut lexical = package.edit_opc().unwrap();
    lexical.update(content_types).unwrap();
    let lexical = lexical.commit().unwrap();
    let expected = lexical.package().to_bytes().unwrap();
    let copy = Package::from_bytes(&original).unwrap();
    assert_eq!(
        copy.apply_opc(lexical.patch())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
        expected,
    );
    let reopened = Package::from_bytes(&expected).unwrap();
    assert_eq!(
        reopened
            .apply_opc(&lexical.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
        original,
    );
}

#[test]
fn structural_inverse_with_producer_markup_composes_with_unrelated_parts() {
    let package = Package::new().unwrap();
    let mut producer = package.edit_opc().unwrap();
    producer.update(content_types).unwrap();
    let package =
        Package::from_bytes(&producer.commit().unwrap().package().to_bytes().unwrap()).unwrap();
    let mut first = package.edit_opc().unwrap();
    first.update(|opc| add_part(opc, "/custom/a.bin")).unwrap();
    let first = first.commit().unwrap();
    let mut second = package.edit_opc().unwrap();
    second.update(|opc| add_part(opc, "/custom/b.bin")).unwrap();
    let second = second.commit().unwrap();
    let combined = second.package().apply_opc(first.patch()).unwrap();
    let reopened = Package::from_bytes(&combined.package().to_bytes().unwrap()).unwrap();
    let restored = reopened.apply_opc(&first.patch().inverse()).unwrap();
    assert_eq!(
        restored.package().to_bytes().unwrap(),
        second.package().to_bytes().unwrap()
    );
}

#[test]
fn lexical_change_then_restore_is_an_exact_noop() {
    let package = Package::new().unwrap();
    let bytes = package.to_bytes().unwrap();
    let mut edit = package.edit_opc().unwrap();
    edit.update(|opc| {
        let limits = opc.read_limits();
        let original = opc.source_content_types(limits)?;
        content_types(opc)?;
        let changed = opc.source_content_types(limits)?;
        opc.try_replace_content_types(changed.bytes(), &original, limits)?;
        Ok(())
    })
    .unwrap();
    let commit = edit.commit().unwrap();
    assert!(!commit.diagnostics().changed());
    assert!(commit.patch().is_noop());
    assert_eq!(commit.package().to_bytes().unwrap(), bytes);
}

#[test]
fn payload_replay_keeps_relationship_spelling() {
    let package = Package::new().unwrap();
    let mut prepare = package.edit_opc().unwrap();
    prepare
        .update(|opc| {
            add_part(opc, "/custom/a.bin")?;
            relationships(opc, "/custom/a.bin")
        })
        .unwrap();
    let original = prepare.commit().unwrap().package().to_bytes().unwrap();
    let package = Package::from_bytes(&original).unwrap();
    let mut edit = package.edit_opc().unwrap();
    edit.update(|opc| {
        let name = PackURI::new("/custom/a.bin").map_err(litchi_pptx::Error::Uri)?;
        opc.get_part_mut(&name)?.set_blob(vec![4, 5, 6]);
        Ok(())
    })
    .unwrap();
    let commit = edit.commit().unwrap();
    let output = commit.package().to_bytes().unwrap();
    assert_eq!(
        package
            .apply_opc(commit.patch())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
        output
    );
    let reopened = Package::from_bytes(&output).unwrap();
    assert_eq!(
        reopened
            .apply_opc(&commit.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
        original
    );
}
