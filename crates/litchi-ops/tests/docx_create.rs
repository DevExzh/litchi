#![cfg(feature = "docx")]

use litchi_core::{FileFormat, Metadata};
use litchi_ops::{
    Code, Conformance, CreateOptions, File, Flavor, NewNode, Placement, create, create_with, props,
};

#[test]
fn create_adopts_the_typed_eager_docx_owner() {
    let file = create(FileFormat::Docx).expect("create DOCX");
    let File::Docx(docx) = &file else {
        panic!("creator returned another format");
    };
    assert!(!docx.is_source_backed());
    assert_eq!(file.text().expect("read created document"), "");
    let metadata = file
        .metadata()
        .expect("read default metadata")
        .expect("DOCX metadata");
    assert!(metadata.created.is_none());
    assert!(metadata.modified.is_none());
}

#[test]
fn create_template_and_metadata_use_leaf_native_state() {
    let metadata = Metadata {
        title: Some("Created title".to_owned()),
        author: Some("Created author".to_owned()),
        ..Metadata::default()
    };
    let options = CreateOptions::new()
        .flavor(Flavor::Template)
        .metadata(metadata);
    let file = create_with(FileFormat::Docx, &options).expect("create template");
    assert_eq!(file.format(), Some(FileFormat::Docx));
    let observed = file
        .metadata()
        .expect("read metadata")
        .expect("DOCX metadata");
    assert_eq!(observed.title.as_deref(), Some("Created title"));
    assert_eq!(observed.author.as_deref(), Some("Created author"));

    let bytes = file.to_bytes().expect("publish template");
    let package = litchi_opc::OpcPackage::from_bytes(&bytes, litchi_opc::SourceLimits::default())
        .expect("inspect OPC package");
    assert_eq!(
        package
            .main_document_part()
            .expect("main document")
            .content_type(),
        "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml"
    );
}

#[test]
fn create_refuses_strict_and_unrepresentable_metadata() {
    let strict = CreateOptions::new().conformance(Conformance::Strict);
    assert_eq!(
        create_with(FileFormat::Docx, &strict)
            .expect_err("Strict is not silently downgraded")
            .code(),
        Code::UnsupportedCapability
    );

    let unsupported = CreateOptions::new().metadata(Metadata {
        page_count: Some(1),
        ..Metadata::default()
    });
    assert_eq!(
        create_with(FileFormat::Docx, &unsupported)
            .expect_err("unsupported metadata is not silently dropped")
            .code(),
        Code::UnsupportedCapability
    );

    let unsupported = CreateOptions::new().metadata(Metadata {
        hyperlink_base: Some("https://example.invalid/".to_owned()),
        ..Metadata::default()
    });
    assert_eq!(
        create_with(FileFormat::Docx, &unsupported)
            .expect_err("unsupported metadata is not silently dropped")
            .code(),
        Code::UnsupportedCapability
    );
}

#[cfg(not(feature = "xlsx"))]
#[test]
fn create_reports_a_disabled_leaf_as_an_unsupported_format() {
    assert_eq!(
        create(FileFormat::Xlsx)
            .expect_err("XLSX creator is disabled in this build")
            .code(),
        Code::UnsupportedFormat
    );
}

#[cfg(feature = "xlsb")]
#[test]
fn create_reports_an_integrated_xlsb_leaf_as_supported() {
    let file = create(FileFormat::Xlsb).expect("XLSB creator is integrated");
    assert_eq!(file.format(), Some(FileFormat::Xlsb));
}

#[cfg(not(feature = "xlsb"))]
#[test]
fn create_reports_a_disabled_xlsb_leaf_as_an_unsupported_format() {
    assert_eq!(
        create(FileFormat::Xlsb)
            .expect_err("XLSB creator is disabled in this build")
            .code(),
        Code::UnsupportedFormat
    );
}

#[test]
fn create_is_deterministic_and_debug_redacts_metadata() {
    let options = CreateOptions::new().metadata(Metadata {
        title: Some("private authored title".to_owned()),
        ..Metadata::default()
    });
    let first = create_with(FileFormat::Docx, &options)
        .expect("first create")
        .to_bytes()
        .expect("first bytes");
    let second = create_with(FileFormat::Docx, &options)
        .expect("second create")
        .to_bytes()
        .expect("second bytes");
    assert_eq!(first, second);
    let debug = format!("{options:?}");
    assert!(!debug.contains("private authored title"));
}

#[test]
fn created_docx_remains_a_normal_editable_owner() {
    let original = create(FileFormat::Docx).expect("create DOCX");
    let first = original
        .add("/body", NewNode::new("paragraph").prop("text", "First"))
        .expect("add first paragraph")
        .into_file();
    let second = first
        .add("/body", NewNode::new("paragraph").prop("text", "Second"))
        .expect("add second paragraph")
        .into_file();
    let with_run = second
        .add(
            "/body/p[1]",
            NewNode::new("run")
                .prop("text", " suffix")
                .at(Placement::End),
        )
        .expect("add run")
        .into_file();
    let set = with_run
        .set("/body/p[1]", props! { "text" => "Revised" })
        .expect("set paragraph text");
    let moved = set
        .move_to("/body/p[1]", Placement::End)
        .expect("move paragraph");
    let swapped = moved
        .swap("/body/p[1]", "/body/p[2]")
        .expect("swap paragraphs");
    let removed = swapped.remove("/body/p[2]").expect("remove paragraph");
    assert_eq!(removed.text().expect("read edited text"), "Revised");

    let bytes = removed.to_bytes().expect("publish edited package");
    let reopened = litchi_ops::open_bytes(bytes.clone()).expect("reopen edited package");
    assert_eq!(reopened.text().expect("read reopened text"), "Revised");

    let directory = tempfile::tempdir().expect("scratch directory");
    let destination = directory.path().join("created.docx");
    let saved = removed.save_as(&destination).expect("save created package");
    assert_eq!(
        saved.bytes,
        std::fs::metadata(&destination).expect("saved file").len()
    );
    let from_path = litchi_ops::open(&destination).expect("open saved package");
    from_path.save().expect("save path-backed package");
    assert_eq!(
        original
            .save()
            .expect_err("created owner has no origin")
            .code(),
        Code::FileRequired
    );
}
