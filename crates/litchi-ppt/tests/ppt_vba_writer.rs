#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
#![cfg(feature = "vba-inspection")]

use litchi_core::OwnedSource;
use litchi_ppt::OpenOptions;
use litchi_ppt::embedded::storage::{Compression, Kind};
use litchi_ppt::writer::Builder;
#[cfg(feature = "encryption")]
use litchi_ppt::writer::EncryptionProfile;
use litchi_ppt::{VbaProjectCompression, VbaProjectError, VbaProjectLimits, raw::Reader};
use litchi_vba::{
    Limits,
    build::{Module, Project},
};
use std::io::Cursor;
use std::sync::Arc;

fn write(writer: &mut Builder) -> Vec<u8> {
    let mut output = Cursor::new(Vec::new());
    litchi_ppt::raw::encode(writer, &mut output).unwrap();
    output.into_inner()
}

#[test]
fn compressed_complete_project_round_trips_as_inert_source() {
    let mut writer = Builder::new();
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

    let mut package = Reader::from_reader(Cursor::new(write(&mut writer))).unwrap();
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

    let read_project = presentation.vba().unwrap().unwrap();
    assert_eq!(read_project.name(), "PresentationTools");
    assert_eq!(read_project.modules().len(), 2);
    assert_eq!(read_project.modules()[0].name(), "Module1");
    assert!(
        read_project.modules()[0]
            .source()
            .text()
            .contains("Public Sub RefreshSlides()")
    );
    assert_eq!(read_project.modules()[1].name(), "ThisPresentation");
    assert!(
        read_project.modules()[1]
            .source()
            .text()
            .contains("Private Sub Presentation_Open()")
    );
}

#[test]
fn uncompressed_empty_project_round_trips_and_can_be_cleared() {
    let mut writer = Builder::new();
    writer.add_slide().unwrap();
    let project = Project::new("EmptyPresentation")
        .finish(&Limits::default())
        .unwrap();
    writer
        .put_vba_with(project, VbaProjectCompression::Uncompressed)
        .unwrap();

    let mut package = Reader::from_reader(Cursor::new(write(&mut writer))).unwrap();
    let presentation = package.presentation().unwrap();
    let storage = presentation.vba_project_storage().unwrap().unwrap();
    assert!(storage.has_macros());
    assert!(!storage.is_compressed());
    assert_eq!(storage.declared_uncompressed_len(), None);
    let read_project = presentation.vba().unwrap().unwrap();
    assert_eq!(read_project.name(), "EmptyPresentation");
    assert!(read_project.modules().is_empty());

    writer.clear_vba();
    assert!(!writer.has_vba());
    let mut cleared_package = Reader::from_reader(Cursor::new(write(&mut writer))).unwrap();
    let cleared_presentation = cleared_package.presentation().unwrap();
    let cleared_storage = cleared_presentation.vba_project_storage().unwrap().unwrap();
    assert!(!cleared_storage.has_macros());
    assert!(!cleared_storage.has_persisted_storage());
    assert!(cleared_presentation.vba().unwrap().is_none());
}

#[test]
fn failed_replacement_is_atomic_and_outer_limits_are_enforced() {
    let mut writer = Builder::new();
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

    let bytes = write(&mut writer);
    let mut package = Reader::from_reader(Cursor::new(bytes.clone())).unwrap();
    let presentation = package.presentation().unwrap();
    let project = presentation.vba().unwrap().unwrap();
    assert_eq!(project.name(), "ExistingProject");

    let limits = VbaProjectLimits {
        max_cfb_bytes: 1,
        ..VbaProjectLimits::default()
    };
    let limited_cfb = package
        .presentation_with_options(OpenOptions::new().vba_project_limits(limits))
        .unwrap();
    assert_eq!(limited_cfb.vba_project_limits(), limits);
    let owned = litchi_ppt::Package::from_bytes_with_options(
        &bytes,
        OpenOptions::new().vba_project_limits(limits),
    )
    .unwrap();
    let owned_presentation = owned.presentation().unwrap();
    assert_eq!(owned_presentation.vba_project_limits(), limits);
    assert!(matches!(
        owned_presentation.vba(),
        Err(VbaProjectError::PowerPoint(_))
    ));
    assert!(matches!(
        limited_cfb.vba(),
        Err(VbaProjectError::PowerPoint(_))
    ));

    let inverse_limits = VbaProjectLimits {
        project: Limits {
            max_cfb_bytes: 1,
            ..Limits::default()
        },
        ..VbaProjectLimits::default()
    };
    let inverse = package
        .presentation_with_options(OpenOptions::new().vba_project_limits(inverse_limits))
        .unwrap();
    assert!(matches!(inverse.vba(), Err(VbaProjectError::PowerPoint(_))));

    let stored_limits = VbaProjectLimits {
        max_stored_bytes: 0,
        ..VbaProjectLimits::default()
    };
    let limited_stored = package
        .presentation_with_options(OpenOptions::new().vba_project_limits(stored_limits))
        .unwrap();
    assert!(matches!(
        limited_stored.vba(),
        Err(VbaProjectError::PowerPoint(_))
    ));
}

#[test]
fn vba_policy_survives_projection_adoption_and_republication() {
    let mut writer = Builder::new();
    writer.add_slide().unwrap();
    writer.set_vba(Project::new("RetainedPolicy")).unwrap();
    let bytes = write(&mut writer);
    let selected = VbaProjectLimits {
        project: Limits {
            max_cfb_bytes: 1,
            ..Limits::default()
        },
        ..VbaProjectLimits::default()
    };

    let mut eager = Reader::from_reader_with_options(
        Cursor::new(bytes.clone()),
        OpenOptions::new().vba_project_limits(selected),
    )
    .unwrap();
    let inherited = eager.presentation_with_options(OpenOptions::new()).unwrap();
    assert_eq!(inherited.vba_project_limits(), selected);
    assert!(matches!(
        inherited.vba(),
        Err(VbaProjectError::PowerPoint(_))
    ));
    let explicitly_replaced = eager
        .presentation_with_options(
            OpenOptions::new().vba_project_limits(VbaProjectLimits::default()),
        )
        .unwrap();
    assert_eq!(
        explicitly_replaced.vba_project_limits(),
        VbaProjectLimits::default()
    );
    assert!(explicitly_replaced.vba().unwrap().is_some());

    let source = litchi_ppt::SourceBackedPackage::from_read_at_with_options(
        Arc::new(OwnedSource::new(bytes)),
        OpenOptions::new().vba_project_limits(selected),
    )
    .unwrap();
    assert_eq!(
        source
            .presentation_with_options(OpenOptions::new())
            .unwrap()
            .vba_project_limits(),
        selected
    );
    let adopted = litchi_ppt::Package::from_source(source, OpenOptions::new()).unwrap();
    assert_eq!(
        adopted.presentation().unwrap().vba_project_limits(),
        selected
    );

    let mut metadata = adopted.edit_metadata().unwrap();
    metadata
        .merge_metadata(&litchi_core::Metadata {
            title: Some("Retained policy".to_owned()),
            ..litchi_core::Metadata::default()
        })
        .unwrap();
    let republished = metadata.commit().unwrap().into_package();
    let republished_presentation = republished.presentation().unwrap();
    assert_eq!(republished_presentation.vba_project_limits(), selected);
    assert!(matches!(
        republished_presentation.vba(),
        Err(VbaProjectError::PowerPoint(_))
    ));
}

#[test]
#[cfg(feature = "encryption")]
fn project_remains_available_after_presentation_decryption() {
    let mut writer = Builder::new();
    writer.add_slide().unwrap();
    let project = Project::new("EncryptedPresentation").module(Module::standard(
        "Module1",
        "Sub StoredOnly()\r\nEnd Sub\r\n",
    ));
    writer.set_vba(project).unwrap();
    writer
        .set_password("secret", EncryptionProfile::CryptoApiRc4 { key_bits: 128 })
        .unwrap();
    writer.set_encryption_entropy(litchi_ppt::EncryptionEntropy::generate().unwrap());

    let mut package = Reader::from_reader(Cursor::new(write(&mut writer))).unwrap();
    let presentation = package
        .presentation_with_options(OpenOptions::new().password("secret"))
        .unwrap();
    let read_project = presentation.vba().unwrap().unwrap();
    assert_eq!(read_project.name(), "EncryptedPresentation");
    assert_eq!(read_project.modules()[0].name(), "Module1");
}
