use litchi_core::Metadata;
use litchi_docx::Package;
use litchi_docx::metadata::Error;
use tempfile::NamedTempFile;

#[test]
fn package_owns_core_property_crud_through_the_metadata_family() {
    let package = Package::new().expect("fresh DOCX");
    assert_eq!(package.metadata().cloned(), Some(Metadata::default()));

    let mut edit = package.edit_metadata();
    edit.set(Metadata {
        title: Some("Moved".to_owned()),
        ..Metadata::default()
    });
    let commit = edit.commit().expect("commit");
    assert!(commit.diagnostics().changed());
    let moved = commit.into_package();
    assert_eq!(
        moved.metadata().and_then(|props| props.title.as_deref()),
        Some("Moved")
    );
    // The source snapshot keeps its value.
    assert_eq!(package.metadata().cloned(), Some(Metadata::default()));

    let path = NamedTempFile::with_suffix(".docx").unwrap();
    moved.save(path.path()).unwrap();
    let reopened = Package::open(path.path()).unwrap();
    assert_eq!(
        reopened.metadata().and_then(|props| props.title.as_deref()),
        Some("Moved")
    );
    let mut edit = reopened.edit_metadata();
    edit.clear();
    let cleared = edit.commit().expect("clear").into_package();

    let output = NamedTempFile::with_suffix(".docx").unwrap();
    cleared.save(output.path()).unwrap();
    let reopened = Package::open(output.path()).unwrap();
    assert!(reopened.metadata().is_none());
}

#[test]
fn metadata_patch_replays_and_inverts_on_reopened_bytes() {
    let package = Package::new().expect("fresh DOCX");
    let mut edit = package.edit_metadata();
    edit.update(|props| {
        props.subject = Some("Subject".to_owned());
        props.category = Some("Category".to_owned());
    });
    let (changed, patch) = edit.commit().expect("commit").into_parts();
    let reopened = Package::from_bytes(&changed.to_bytes().unwrap()).unwrap();
    assert_eq!(reopened.metadata(), patch.target());

    let replay = package
        .apply_metadata(&patch)
        .expect("replay")
        .into_package();
    assert_eq!(replay.metadata(), patch.target());
    let inverse = reopened
        .apply_metadata(&patch.inverse())
        .expect("inverse")
        .into_package();
    assert_eq!(inverse.metadata(), package.metadata());
    assert!(matches!(
        package.apply_metadata(&patch.inverse()),
        Err(Error::PatchConflict)
    ));
}
