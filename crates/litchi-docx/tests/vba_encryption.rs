#![cfg(all(feature = "encryption", feature = "vba-inspection"))]

use std::sync::Arc;

use litchi_docx::encryption::Mode;
use litchi_docx::vba_project::VbaSupplementalData;
use litchi_docx::{Error, Package};
use litchi_vba::{Limits, Payload, build};

const PASSWORD: &str = "Litchi VBA boundary 42!";

fn project(name: &str, message: &str) -> build::Project {
    build::Project::new(name).module(build::Module::standard(
        "Module1",
        format!("Public Sub Hello()\r\n  Debug.Print \"{message}\"\r\nEnd Sub\r\n"),
    ))
}

fn payload(name: &str, message: &str) -> Payload {
    project(name, message).finish(&Limits::default()).unwrap()
}

#[test]
fn encrypted_vba_noops_preserve_graph_and_provenance() {
    let directory = tempfile::tempdir().unwrap();
    let macro_path = directory.path().join("macro.docm");
    let plain_path = directory.path().join("plain.docx");

    let mut package = Package::new().unwrap();
    package.set_vba(project("Stable", "same")).unwrap();
    package
        .save_encrypted(&macro_path, PASSWORD, Mode::Agile)
        .unwrap();
    let source = std::fs::read(&macro_path).unwrap();
    let mut opened = Package::open_with_password(&macro_path, PASSWORD).unwrap();
    let before = opened.vba().unwrap().unwrap();
    let project_before = opened
        .opc_package()
        .get_part(before.project_part_name())
        .unwrap()
        .blob_arc();
    let supplemental_before = opened
        .opc_package()
        .get_part(before.supplemental_data_part_name())
        .unwrap()
        .blob_arc();

    opened.set_vba(project("Stable", "same")).unwrap();
    opened
        .put_vba(payload("Stable", "same"), &VbaSupplementalData::new())
        .unwrap();

    let after = opened.vba().unwrap().unwrap();
    assert_eq!(before, after);
    assert!(Arc::ptr_eq(
        &project_before,
        &opened
            .opc_package()
            .get_part(after.project_part_name())
            .unwrap()
            .blob_arc()
    ));
    assert!(Arc::ptr_eq(
        &supplemental_before,
        &opened
            .opc_package()
            .get_part(after.supplemental_data_part_name())
            .unwrap()
            .blob_arc()
    ));
    assert_eq!(opened.encryption(), Some(Mode::Agile));
    assert_eq!(std::fs::read(&macro_path).unwrap(), source);

    let mut plain = Package::new().unwrap();
    plain
        .save_encrypted(&plain_path, PASSWORD, Mode::Standard)
        .unwrap();
    let source = std::fs::read(&plain_path).unwrap();
    let mut opened = Package::open_with_password(&plain_path, PASSWORD).unwrap();
    assert!(!opened.clear_vba().unwrap());
    assert_eq!(opened.encryption(), Some(Mode::Standard));
    assert_eq!(std::fs::read(&plain_path).unwrap(), source);
}

#[test]
fn encrypted_vba_changes_are_refused_atomically() {
    let directory = tempfile::tempdir().unwrap();
    let plain_path = directory.path().join("plain.docx");
    let macro_path = directory.path().join("macro.docm");

    let mut plain = Package::new().unwrap();
    plain
        .save_encrypted(&plain_path, PASSWORD, Mode::Agile)
        .unwrap();
    let source = std::fs::read(&plain_path).unwrap();
    let mut opened = Package::open_with_password(&plain_path, PASSWORD).unwrap();
    assert!(matches!(
        opened.set_vba(project("Added", "set")),
        Err(Error::UnsafeEdit {
            operation: "set_vba",
            ..
        })
    ));
    assert!(matches!(
        opened.put_vba(payload("Added", "put"), &VbaSupplementalData::new()),
        Err(Error::UnsafeEdit {
            operation: "put_vba",
            ..
        })
    ));
    assert!(opened.vba().unwrap().is_none());
    assert_eq!(std::fs::read(&plain_path).unwrap(), source);

    let mut package = Package::new().unwrap();
    package.set_vba(project("Existing", "keep")).unwrap();
    package
        .save_encrypted(&macro_path, PASSWORD, Mode::Standard)
        .unwrap();
    let source = std::fs::read(&macro_path).unwrap();
    let mut opened = Package::open_with_password(&macro_path, PASSWORD).unwrap();
    let before = opened.vba().unwrap().unwrap();
    let project_before = opened
        .opc_package()
        .get_part(before.project_part_name())
        .unwrap()
        .blob_arc();
    assert!(matches!(
        opened.clear_vba(),
        Err(Error::UnsafeEdit {
            operation: "clear_vba",
            ..
        })
    ));
    let after = opened.vba().unwrap().unwrap();
    assert_eq!(before, after);
    assert!(Arc::ptr_eq(
        &project_before,
        &opened
            .opc_package()
            .get_part(after.project_part_name())
            .unwrap()
            .blob_arc()
    ));
    assert_eq!(std::fs::read(&macro_path).unwrap(), source);
}

#[test]
fn explicit_plaintext_copy_allows_vba_changes() {
    let directory = tempfile::tempdir().unwrap();
    let encrypted_path = directory.path().join("source.docx");
    let plaintext_path = directory.path().join("editable.docx");

    let mut package = Package::new().unwrap();
    package
        .save_encrypted(&encrypted_path, PASSWORD, Mode::Agile)
        .unwrap();
    let mut encrypted = Package::open_with_password(&encrypted_path, PASSWORD).unwrap();
    assert!(matches!(
        encrypted.set_vba(project("First", "refused")),
        Err(Error::UnsafeEdit { .. })
    ));
    encrypted.save_plain(&plaintext_path).unwrap();

    let mut plaintext = Package::open(&plaintext_path).unwrap();
    plaintext.set_vba(project("First", "accepted")).unwrap();
    plaintext
        .put_vba(
            payload("Replacement", "accepted"),
            &VbaSupplementalData::new(),
        )
        .unwrap();
    assert_eq!(
        plaintext
            .vba()
            .unwrap()
            .unwrap()
            .project(plaintext.opc_package())
            .unwrap()
            .name(),
        "Replacement"
    );
    assert!(plaintext.clear_vba().unwrap());
    assert!(!plaintext.clear_vba().unwrap());
}
