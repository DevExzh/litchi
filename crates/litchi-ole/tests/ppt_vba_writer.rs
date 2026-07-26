use litchi_ole::ovba::{VbaLimits, VbaModuleBuilder, VbaProjectBuilder};
use litchi_ole::ppt::writer::{PptEncryptionProfile, PptWriter};
use litchi_ole::ppt::{
    Package, PowerPointOleStorageCompression, PowerPointOleStorageKind,
    PowerPointVbaProjectCompression, PowerPointVbaProjectError, PowerPointVbaProjectLimits,
    PptOpenOptions,
};
use std::io::Cursor;

fn write(writer: &mut PptWriter) -> Vec<u8> {
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

#[test]
fn compressed_complete_project_round_trips_as_inert_source() {
    let mut writer = PptWriter::new();
    let slide = writer.add_slide().unwrap();
    writer
        .add_textbox(slide, 10, 20, 300, 40, "Macro-enabled presentation")
        .unwrap();
    let project = VbaProjectBuilder::new("PresentationTools")
        .with_module(VbaModuleBuilder::standard(
            "Module1",
            "Public Sub RefreshSlides()\r\nEnd Sub\r\n",
        ))
        .with_module(VbaModuleBuilder::document(
            "ThisPresentation",
            0,
            "Private Sub Presentation_Open()\r\nEnd Sub\r\n",
        ));
    writer
        .set_vba_project(
            &project,
            &VbaLimits::default(),
            PowerPointVbaProjectCompression::Zlib,
        )
        .unwrap();
    assert!(writer.has_vba_project());

    let mut package = Package::from_reader(Cursor::new(write(&mut writer))).unwrap();
    let presentation = package.presentation().unwrap();
    let storage = presentation.vba_project_storage().unwrap().unwrap();
    assert!(storage.persist_id_ref() > 0);
    assert!(storage.has_macros());
    assert!(storage.has_persisted_storage());
    assert!(storage.is_compressed());
    assert!(storage.declared_uncompressed_len().unwrap() > 0);

    let outer = presentation
        .ole_storage_as(
            storage.persist_id_ref(),
            PowerPointOleStorageKind::VbaProject,
        )
        .unwrap()
        .unwrap();
    assert_eq!(outer.kind, PowerPointOleStorageKind::VbaProject);
    assert!(matches!(
        outer.compression,
        PowerPointOleStorageCompression::Zlib { .. }
    ));

    let project = presentation
        .vba_project(&PowerPointVbaProjectLimits::default())
        .unwrap()
        .unwrap();
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
    let mut writer = PptWriter::new();
    writer.add_slide().unwrap();
    writer
        .enable_empty_vba_project(
            "EmptyPresentation",
            PowerPointVbaProjectCompression::Uncompressed,
        )
        .unwrap();

    let mut package = Package::from_reader(Cursor::new(write(&mut writer))).unwrap();
    let presentation = package.presentation().unwrap();
    let storage = presentation.vba_project_storage().unwrap().unwrap();
    assert!(storage.has_macros());
    assert!(!storage.is_compressed());
    assert_eq!(storage.declared_uncompressed_len(), None);
    let project = presentation
        .vba_project(&PowerPointVbaProjectLimits::default())
        .unwrap()
        .unwrap();
    assert_eq!(project.name(), "EmptyPresentation");
    assert!(project.modules().is_empty());

    writer.clear_vba_project();
    assert!(!writer.has_vba_project());
    let mut package = Package::from_reader(Cursor::new(write(&mut writer))).unwrap();
    let presentation = package.presentation().unwrap();
    let storage = presentation.vba_project_storage().unwrap().unwrap();
    assert!(!storage.has_macros());
    assert!(!storage.has_persisted_storage());
    assert!(
        presentation
            .vba_project(&PowerPointVbaProjectLimits::default())
            .unwrap()
            .is_none()
    );
}

#[test]
fn failed_replacement_is_atomic_and_outer_limits_are_enforced() {
    let mut writer = PptWriter::new();
    writer.add_slide().unwrap();
    writer
        .enable_empty_vba_project("ExistingProject", PowerPointVbaProjectCompression::Zlib)
        .unwrap();
    let replacement = VbaProjectBuilder::new("Replacement").with_module(
        VbaModuleBuilder::standard("Module1", "Sub A()\r\nEnd Sub\r\n"),
    );
    let build_limits = VbaLimits {
        max_modules: 0,
        ..VbaLimits::default()
    };
    assert!(
        writer
            .set_vba_project(
                &replacement,
                &build_limits,
                PowerPointVbaProjectCompression::Zlib,
            )
            .is_err()
    );

    let mut package = Package::from_reader(Cursor::new(write(&mut writer))).unwrap();
    let presentation = package.presentation().unwrap();
    let project = presentation
        .vba_project(&PowerPointVbaProjectLimits::default())
        .unwrap()
        .unwrap();
    assert_eq!(project.name(), "ExistingProject");

    let limits = PowerPointVbaProjectLimits {
        max_cfb_bytes: 1,
        ..PowerPointVbaProjectLimits::default()
    };
    assert!(matches!(
        presentation.vba_project(&limits),
        Err(PowerPointVbaProjectError::PowerPoint(_))
    ));
}

#[test]
fn project_remains_available_after_presentation_decryption() {
    let mut writer = PptWriter::new();
    writer.add_slide().unwrap();
    let project = VbaProjectBuilder::new("EncryptedPresentation").with_module(
        VbaModuleBuilder::standard("Module1", "Sub StoredOnly()\r\nEnd Sub\r\n"),
    );
    writer
        .set_vba_project(
            &project,
            &VbaLimits::default(),
            PowerPointVbaProjectCompression::Zlib,
        )
        .unwrap();
    writer
        .set_password(
            "secret",
            PptEncryptionProfile::CryptoApiRc4 { key_bits: 128 },
        )
        .unwrap();

    let mut package = Package::from_reader(Cursor::new(write(&mut writer))).unwrap();
    let presentation = package
        .presentation_with_options(PptOpenOptions {
            password: Some("secret"),
        })
        .unwrap();
    let project = presentation
        .vba_project(&PowerPointVbaProjectLimits::default())
        .unwrap()
        .unwrap();
    assert_eq!(project.name(), "EncryptedPresentation");
    assert_eq!(project.modules()[0].name(), "Module1");
}
