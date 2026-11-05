//! XLSB creation through the unified [`litchi_ops::File`] facade.
#![cfg(feature = "xlsb")]

use litchi_core::FileFormat;
use litchi_ops::capabilities::{Capabilities, Verb};
use litchi_ops::{Code, Conformance, CreateOptions, File, Flavor, create, create_with, props};

#[test]
fn create_adopts_a_native_minimal_workbook_and_reopens_after_edit() {
    let file = create(FileFormat::Xlsb).expect("create XLSB");
    let File::Xlsb(_) = &file else {
        panic!("creator returned another format");
    };

    assert_eq!(file.format(), Some(FileFormat::Xlsb));
    assert_eq!(
        file.workbook()
            .expect("workbook view")
            .worksheet_names()
            .expect("worksheet names"),
        vec!["Sheet1".to_owned()]
    );
    assert!(file.text().expect("read empty workbook").trim().is_empty());
    let metadata = file
        .metadata()
        .expect("read default metadata")
        .expect("XLSB metadata");
    assert!(metadata.author.is_none());
    assert!(metadata.created.is_none());
    assert!(metadata.modified.is_none());

    let original_bytes = file.to_bytes().expect("original bytes");
    let edited = file
        .set("/sheet[1]/A1", props! { "value" => "created" })
        .expect("set a cell in the created workbook");
    assert_eq!(
        file.to_bytes().expect("unchanged original bytes"),
        original_bytes
    );
    assert_eq!(
        edited
            .get("/sheet[1]/A1")
            .expect("read edited cell")
            .view()
            .text()
            .expect("read edited cell text")
            .as_deref(),
        Some("created")
    );

    let bytes = edited.to_bytes().expect("serialize edited XLSB");
    assert!(!bytes.is_empty());
    let reopened = litchi_ops::open_bytes(bytes).expect("reopen created XLSB");
    assert_eq!(
        reopened
            .get("/sheet[1]/A1")
            .expect("read reopened cell")
            .view()
            .text()
            .expect("read reopened cell text")
            .as_deref(),
        Some("created")
    );
}

#[test]
fn xlsb_creation_options_are_explicit_typed_refusals() {
    let cases = [
        CreateOptions::new().flavor(Flavor::Template),
        CreateOptions::new().conformance(Conformance::Strict),
    ];
    for options in cases {
        let error = create_with(FileFormat::Xlsb, &options).expect_err("unsupported option");
        assert_eq!(error.code(), Code::UnsupportedCapability);
        assert_eq!(error.format(), Some(FileFormat::Xlsb));
    }
}

#[test]
fn xlsb_creation_capability_is_advertised() {
    assert!(
        Capabilities::of(FileFormat::Xlsb)
            .verb(Verb::Create)
            .is_supported()
    );
}
