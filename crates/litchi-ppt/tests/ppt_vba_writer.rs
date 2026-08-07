#![cfg(feature = "vba-inspection")]

#[cfg(feature = "encryption")]
use litchi_ppt::OpenOptions;
use litchi_ppt::embedded::storage::{Compression, Kind};
#[cfg(feature = "encryption")]
use litchi_ppt::writer::EncryptionProfile;
use litchi_ppt::writer::Writer;
use litchi_ppt::{Package, VbaProjectCompression, VbaProjectError, VbaProjectLimits};
use litchi_vba::{
    Limits,
    build::{Module, Project},
};
use std::io::Cursor;

fn write(writer: &mut Writer) -> Vec<u8> {
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

#[test]
fn compressed_complete_project_round_trips_as_inert_source() {
    let mut writer = Writer::new();
    let slide = writer.add_slide().unwrap();
    writer
        .add_textbox(slide, 10, 20, 300, 40, "Macro-enabled presentation")
        .unwrap();
    let project = Project::new("PresentationTools")
        .module(Module::standard(
            "Module1",
            "Public Sub RefreshSlides()\r\nEnd Sub\r\n",
        ))
        .module(Module::document(
            "ThisPresentation",
            0,
            "Private Sub Presentation_Open()\r\nEnd Sub\r\n",
        ));
    writer.set_vba(project).unwrap();
    assert!(writer.has_vba());

    let mut package = Package::from_reader(Cursor::new(write(&mut writer))).unwrap();
    let presentation = package.presentation().unwrap();
    let storage = presentation.vba_project_storage().unwrap().unwrap();
    assert!(storage.persist_id_ref() > 0);
    assert!(storage.has_macros());
    assert!(storage.has_persisted_storage());
    assert!(storage.is_compressed());
    assert!(storage.declared_uncompressed_len().unwrap() > 0);

    let outer = presentation
        .ole_storage_as(storage.persist_id_ref(), Kind::VbaProject)
        .unwrap()
        .unwrap();
    assert_eq!(outer.kind(), Kind::VbaProject);
    assert!(matches!(outer.compression(), Compression::Zlib));

    let project = presentation.vba().unwrap().unwrap();
    assert_eq!(project.name(), "PresentationTools");
    assert_eq!(project.modules().len(), 2);
    assert_eq!(project.modules()[0].name(), "Module1");
    assert!(
        project.modules()[0]
            .source()
            .text()
            .contains("Public Sub RefreshSlides()")
    );
    assert_eq!(project.modules()[1].name(), "ThisPresentation");
    assert!(
        project.modules()[1]
            .source()
            .text()
            .contains("Private Sub Presentation_Open()")
    );
}

#[test]
fn uncompressed_empty_project_round_trips_and_can_be_cleared() {
    let mut writer = Writer::new();
    writer.add_slide().unwrap();
    let project = Project::new("EmptyPresentation")
        .finish(&Limits::default())
        .unwrap();
    writer
        .put_vba_with(project, VbaProjectCompression::Uncompressed)
        .unwrap();

    let mut package = Package::from_reader(Cursor::new(write(&mut writer))).unwrap();
    let presentation = package.presentation().unwrap();
    let storage = presentation.vba_project_storage().unwrap().unwrap();
    assert!(storage.has_macros());
    assert!(!storage.is_compressed());
    assert_eq!(storage.declared_uncompressed_len(), None);
    let project = presentation.vba().unwrap().unwrap();
    assert_eq!(project.name(), "EmptyPresentation");
    assert!(project.modules().is_empty());

    writer.clear_vba();
    assert!(!writer.has_vba());
    let mut package = Package::from_reader(Cursor::new(write(&mut writer))).unwrap();
    let presentation = package.presentation().unwrap();
    let storage = presentation.vba_project_storage().unwrap().unwrap();
    assert!(!storage.has_macros());
    assert!(!storage.has_persisted_storage());
    assert!(presentation.vba().unwrap().is_none());
}

#[test]
fn failed_replacement_is_atomic_and_outer_limits_are_enforced() {
    let mut writer = Writer::new();
    writer.add_slide().unwrap();
    writer.set_vba(Project::new("ExistingProject")).unwrap();
    let replacement =
        Project::new("Replacement").module(Module::standard("Module1", "Sub A()\r\nEnd Sub\r\n"));
    let build_limits = Limits {
        max_modules: 0,
        ..Limits::default()
    };
    assert!(
        writer
            .set_vba_with(replacement, &build_limits, VbaProjectCompression::Zlib,)
            .is_err()
    );

    let mut package = Package::from_reader(Cursor::new(write(&mut writer))).unwrap();
    let presentation = package.presentation().unwrap();
    let project = presentation.vba().unwrap().unwrap();
    assert_eq!(project.name(), "ExistingProject");

    let limits = VbaProjectLimits {
        max_cfb_bytes: 1,
        ..VbaProjectLimits::default()
    };
    assert!(matches!(
        presentation.vba_with(&limits),
        Err(VbaProjectError::PowerPoint(_))
    ));

    let limits = VbaProjectLimits {
        max_stored_bytes: 0,
        ..VbaProjectLimits::default()
    };
    assert!(matches!(
        presentation.vba_with(&limits),
        Err(VbaProjectError::PowerPoint(_))
    ));
}

#[test]
#[cfg(feature = "encryption")]
fn project_remains_available_after_presentation_decryption() {
    let mut writer = Writer::new();
    writer.add_slide().unwrap();
    let project = Project::new("EncryptedPresentation").module(Module::standard(
        "Module1",
        "Sub StoredOnly()\r\nEnd Sub\r\n",
    ));
    writer.set_vba(project).unwrap();
    writer
        .set_password("secret", EncryptionProfile::CryptoApiRc4 { key_bits: 128 })
        .unwrap();

    let mut package = Package::from_reader(Cursor::new(write(&mut writer))).unwrap();
    let presentation = package
        .presentation_with_options(OpenOptions {
            password: Some("secret"),
            ..OpenOptions::default()
        })
        .unwrap();
    let project = presentation.vba().unwrap().unwrap();
    assert_eq!(project.name(), "EncryptedPresentation");
    assert_eq!(project.modules()[0].name(), "Module1");
}
