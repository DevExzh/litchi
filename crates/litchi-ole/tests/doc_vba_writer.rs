use litchi_ole::doc::{DocEncryptionProfile, DocOpenOptions, DocWriter, Package};
use litchi_ole::ovba::{VbaLimits, VbaModuleBuilder, VbaProjectBuilder};
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
    let project = VbaProjectBuilder::new("DocumentTools")
        .with_module(VbaModuleBuilder::standard(
            "Module1",
            "Public Sub RefreshFields()\r\nEnd Sub\r\n",
        ))
        .with_module(VbaModuleBuilder::document(
            "ThisDocument",
            0,
            "Private Sub Document_Open()\r\nEnd Sub\r\n",
        ));
    writer
        .set_vba_project(&project, &VbaLimits::default())
        .unwrap();
    assert!(writer.has_vba_project());

    let mut package = Package::from_reader(Cursor::new(write(&mut writer))).unwrap();
    let storage = package.vba_project_storage().unwrap();
    assert_eq!(storage.project_root_path(), ["Macros"]);
    assert_eq!(storage.vba_storage_path(), ["Macros", "VBA"]);
    assert!(storage.is_structurally_complete());
    assert_eq!(
        storage.candidate_module_stream_names(),
        ["Module1", "ThisDocument"]
    );

    let project = package.vba_project(&VbaLimits::default()).unwrap().unwrap();
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
    let project = VbaProjectBuilder::new("EncryptedDocument").with_module(
        VbaModuleBuilder::standard("Module1", "Sub StoredOnly()\r\nEnd Sub\r\n"),
    );
    writer
        .set_vba_project(&project, &VbaLimits::default())
        .unwrap();
    writer
        .set_password("secret", DocEncryptionProfile::OfficeBinaryRc4)
        .unwrap();

    let mut package = Package::from_reader(Cursor::new(write(&mut writer))).unwrap();
    let project = package.vba_project(&VbaLimits::default()).unwrap().unwrap();
    assert_eq!(project.name(), "EncryptedDocument");
    assert_eq!(project.modules()[0].name(), "Module1");

    let document = package
        .document_with_options(DocOpenOptions {
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
    writer.enable_empty_vba_project("ExistingProject").unwrap();
    let replacement = VbaProjectBuilder::new("Replacement").with_module(
        VbaModuleBuilder::standard("Module1", "Sub A()\r\nEnd Sub\r\n"),
    );
    let limits = VbaLimits {
        max_modules: 0,
        ..VbaLimits::default()
    };
    assert!(writer.set_vba_project(&replacement, &limits).is_err());

    let mut package = Package::from_reader(Cursor::new(write(&mut writer))).unwrap();
    let project = package.vba_project(&VbaLimits::default()).unwrap().unwrap();
    assert_eq!(project.name(), "ExistingProject");
    assert!(project.modules().is_empty());

    writer.clear_vba_project();
    assert!(!writer.has_vba_project());
    let package = Package::from_reader(Cursor::new(write(&mut writer))).unwrap();
    assert!(package.vba_project_storage().is_none());
}
