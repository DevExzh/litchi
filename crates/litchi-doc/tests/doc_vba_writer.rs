use litchi_doc::{DocEncryptionProfile, DocWriter, OpenOptions, Package};
use litchi_vba::{
    Limits,
    build::{Module, Project},
};
use std::io::Cursor;

fn write(writer: &mut DocWriter) -> Vec<u8> {
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

#[test]
fn complete_project_round_trips_from_the_macros_storage() {
    let mut writer = DocWriter::new();
    writer.add_paragraph("Body").unwrap();
    let project = Project::new("DocumentTools")
        .module(Module::standard(
            "Module1",
            "Public Sub RefreshFields()\r\nEnd Sub\r\n",
        ))
        .module(Module::document(
            "ThisDocument",
            0,
            "Private Sub Document_Open()\r\nEnd Sub\r\n",
        ));
    writer.set_vba(project).unwrap();
    assert!(writer.has_vba());

    let mut package = Package::from_reader(Cursor::new(write(&mut writer))).unwrap();
    let storage = package.vba_project_storage().unwrap();
    assert_eq!(storage.project_root_path(), ["Macros"]);
    assert_eq!(storage.vba_storage_path(), ["Macros", "VBA"]);
    assert!(storage.is_structurally_complete());
    assert_eq!(
        storage.candidate_module_stream_names(),
        ["Module1", "ThisDocument"]
    );

    let project = package.vba().unwrap().unwrap();
    assert_eq!(project.name(), "DocumentTools");
    assert_eq!(project.modules().len(), 2);
    assert!(
        project.modules()[0]
            .source()
            .text()
            .contains("Public Sub RefreshFields()")
    );
    assert!(
        project.modules()[1]
            .source()
            .text()
            .contains("Private Sub Document_Open()")
    );
}

#[test]
fn project_storage_remains_clear_when_document_streams_are_encrypted() {
    let mut writer = DocWriter::new();
    writer.add_paragraph("Encrypted body").unwrap();
    let project = Project::new("EncryptedDocument").module(Module::standard(
        "Module1",
        "Sub StoredOnly()\r\nEnd Sub\r\n",
    ));
    writer.set_vba(project).unwrap();
    writer
        .set_password("secret", DocEncryptionProfile::OfficeBinaryRc4)
        .unwrap();

    let mut package = Package::from_reader(Cursor::new(write(&mut writer))).unwrap();
    let project = package.vba().unwrap().unwrap();
    assert_eq!(project.name(), "EncryptedDocument");
    assert_eq!(project.modules()[0].name(), "Module1");

    let document = package
        .document_with_options(OpenOptions {
            password: Some("secret"),
            ..Default::default()
        })
        .unwrap();
    assert!(document.text().unwrap().contains("Encrypted body"));
}

#[test]
fn failed_build_is_atomic_and_project_can_be_cleared() {
    let mut writer = DocWriter::new();
    writer.add_paragraph("Body").unwrap();
    let existing = Project::new("ExistingProject")
        .finish(&Limits::default())
        .unwrap();
    writer.put_vba(existing);
    let replacement =
        Project::new("Replacement").module(Module::standard("Module1", "Sub A()\r\nEnd Sub\r\n"));
    let limits = Limits {
        max_modules: 0,
        ..Limits::default()
    };
    assert!(writer.set_vba_with(replacement, &limits).is_err());

    let mut package = Package::from_reader(Cursor::new(write(&mut writer))).unwrap();
    let project = package.vba().unwrap().unwrap();
    assert_eq!(project.name(), "ExistingProject");
    assert!(project.modules().is_empty());

    writer.clear_vba();
    assert!(!writer.has_vba());
    let package = Package::from_reader(Cursor::new(write(&mut writer))).unwrap();
    assert!(package.vba_project_storage().is_none());
}
