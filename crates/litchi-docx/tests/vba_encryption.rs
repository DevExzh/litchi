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
    package = {
        let mut edit = package.edit_vba();
        edit.set(project("Stable", "same")).unwrap();
        edit.commit().unwrap().into_package()
    };
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

    opened = {
        let mut edit = opened.edit_vba();
        edit.set(project("Stable", "same")).unwrap();
        edit.commit().unwrap().into_package()
    };
    opened = {
        let mut edit = opened.edit_vba();
        edit.put(payload("Stable", "same"), &VbaSupplementalData::new())
            .unwrap();
        edit.commit().unwrap().into_package()
    };

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

    let plain = Package::new().unwrap();
    plain
        .save_encrypted(&plain_path, PASSWORD, Mode::Standard)
        .unwrap();
    let source = std::fs::read(&plain_path).unwrap();
    let mut opened = Package::open_with_password(&plain_path, PASSWORD).unwrap();
    assert!(!{
        let mut edit = opened.edit_vba();
        let value = edit.clear().unwrap();
        opened = edit.commit().unwrap().into_package();
        value
    });
    assert_eq!(opened.encryption(), Some(Mode::Standard));
    assert_eq!(std::fs::read(&plain_path).unwrap(), source);
}

#[test]
fn encrypted_vba_changes_are_refused_atomically() {
    let directory = tempfile::tempdir().unwrap();
    let plain_path = directory.path().join("plain.docx");
    let macro_path = directory.path().join("macro.docm");

    let plain = Package::new().unwrap();
    plain
        .save_encrypted(&plain_path, PASSWORD, Mode::Agile)
        .unwrap();
    let source = std::fs::read(&plain_path).unwrap();
    let opened = Package::open_with_password(&plain_path, PASSWORD).unwrap();
    assert!(matches!(
                {
    let mut edit = opened.edit_vba();
    edit.set(project("Added", "set")).map(|_| ()).and_then(|()| edit.commit().map(|_| ()))
                },
                Err(Error::Refused(refusal)) if refusal.operation() == "set_vba"
        ));
    assert!(matches!(
            {
                let mut edit = opened.edit_vba();
                edit.put(payload("Added", "put"), &VbaSupplementalData::new()).map(|_| ()).and_then(|()| edit.commit().map(|_| ()))
            },
            Err(Error::Refused(refusal)) if refusal.operation() == "put_vba"
    ));
    assert!(opened.vba().unwrap().is_none());
    assert_eq!(std::fs::read(&plain_path).unwrap(), source);

    let mut package = Package::new().unwrap();
    package = {
        let mut edit = package.edit_vba();
        edit.set(project("Existing", "keep")).unwrap();
        edit.commit().unwrap().into_package()
    };
    package
        .save_encrypted(&macro_path, PASSWORD, Mode::Standard)
        .unwrap();
    let source = std::fs::read(&macro_path).unwrap();
    let opened = Package::open_with_password(&macro_path, PASSWORD).unwrap();
    let before = opened.vba().unwrap().unwrap();
    let project_before = opened
        .opc_package()
        .get_part(before.project_part_name())
        .unwrap()
        .blob_arc();
    assert!(matches!(
            {
                let mut edit = opened.edit_vba();
                edit.clear().map(|_| ()).and_then(|()| edit.commit().map(|_| ()))
            },
            Err(Error::Refused(refusal)) if refusal.operation() == "clear_vba"
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

    let package = Package::new().unwrap();
    package
        .save_encrypted(&encrypted_path, PASSWORD, Mode::Agile)
        .unwrap();
    let encrypted = Package::open_with_password(&encrypted_path, PASSWORD).unwrap();
    assert!(matches!(
        {
            let mut edit = encrypted.edit_vba();
            edit.set(project("First", "refused"))
                .map(|_| ())
                .and_then(|()| edit.commit().map(|_| ()))
        },
        Err(Error::Refused(_))
    ));
    encrypted.save_plain(&plaintext_path).unwrap();

    let mut plaintext = Package::open(&plaintext_path).unwrap();
    plaintext = {
        let mut edit = plaintext.edit_vba();
        edit.set(project("First", "accepted")).unwrap();
        edit.commit().unwrap().into_package()
    };
    plaintext = {
        let mut edit = plaintext.edit_vba();
        edit.put(
            payload("Replacement", "accepted"),
            &VbaSupplementalData::new(),
        )
        .unwrap();
        edit.commit().unwrap().into_package()
    };
    assert_eq!(
        plaintext
            .vba()
            .unwrap()
            .unwrap()
            .project(plaintext.opc_package(), &Limits::default())
            .unwrap()
            .name(),
        "Replacement"
    );
    assert!({
        let mut edit = plaintext.edit_vba();
        let value = edit.clear().unwrap();
        plaintext = edit.commit().unwrap().into_package();
        value
    });
    assert!(!{
        let mut edit = plaintext.edit_vba();
        let value = edit.clear().unwrap();
        plaintext = edit.commit().unwrap().into_package();
        value
    });
    assert!(plaintext.vba().unwrap().is_none());
}
