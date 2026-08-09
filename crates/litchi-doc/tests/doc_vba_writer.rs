#![allow(
    clippy::arbitrary_source_item_ordering,
    clippy::cast_possible_wrap,
    clippy::let_underscore_must_use,
    clippy::manual_midpoint,
    clippy::map_unwrap_or,
    clippy::needless_pass_by_value,
    clippy::shadow_reuse,
    clippy::wildcard_enum_match_arm,
    clippy::bool_assert_comparison,
    clippy::cast_possible_truncation,
    clippy::cast_sign_loss,
    clippy::decimal_bitwise_operands,
    clippy::default_trait_access,
    clippy::doc_markdown,
    clippy::expect_used,
    clippy::field_reassign_with_default,
    clippy::float_cmp,
    clippy::implicit_clone,
    clippy::items_after_statements,
    clippy::manual_let_else,
    clippy::manual_repeat_n,
    clippy::manual_string_new,
    clippy::match_wildcard_for_single_variants,
    clippy::needless_raw_string_hashes,
    clippy::redundant_closure_for_method_calls,
    clippy::shadow_unrelated,
    clippy::similar_names,
    clippy::uninlined_format_args,
    clippy::unreadable_literal,
    clippy::unwrap_used,
    reason = "integration-test fixtures favor explicit wire values and concise panic-driven assertions over production-style ergonomics"
)]

use litchi_doc::{EncryptionProfile, OpenOptions, Package, Writer};
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
fn complete_project_round_trips_from_the_macros_storage() {
    let mut writer = Writer::new();
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
    let mut writer = Writer::new();
    writer.add_paragraph("Encrypted body").unwrap();
    let project = Project::new("EncryptedDocument").module(Module::standard(
        "Module1",
        "Sub StoredOnly()\r\nEnd Sub\r\n",
    ));
    writer.set_vba(project).unwrap();
    writer
        .set_password("secret", EncryptionProfile::OfficeBinaryRc4)
        .unwrap();

    let mut package = Package::from_reader(Cursor::new(write(&mut writer))).unwrap();
    let project = package.vba().unwrap().unwrap();
    assert_eq!(project.name(), "EncryptedDocument");
    assert_eq!(project.modules()[0].name(), "Module1");

    let document = package
        .document_with_options(OpenOptions::default().with_password("secret".to_owned().into()))
        .unwrap();
    assert!(document.text().unwrap().contains("Encrypted body"));
}

#[test]
fn failed_build_is_atomic_and_project_can_be_cleared() {
    let mut writer = Writer::new();
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
