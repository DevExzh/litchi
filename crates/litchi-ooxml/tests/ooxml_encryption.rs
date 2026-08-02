#![cfg(feature = "encryption")]

use litchi_core::sheet::WorkbookTrait;
use litchi_ooxml::encryption::{Error as CryptoError, Limits, Mode};
use litchi_ooxml::{OoxmlError, docx, pptx, xlsx};

const PASSWORD: &str = "Litchi test password 42!";
const NEW_PASSWORD: &str = "Litchi changed password 7!";

#[test]
fn every_ooxml_host_opens_the_canonical_encrypted_container() {
    let directory = tempfile::tempdir().unwrap();

    let docx_encrypted = directory.path().join("encrypted.docx");
    let mut document = docx::Package::new().unwrap();
    document
        .save_encrypted(&docx_encrypted, PASSWORD, Mode::Standard)
        .unwrap();
    let document = docx::Package::open_with_password(&docx_encrypted, PASSWORD).unwrap();
    assert_eq!(document.encryption(), Some(Mode::Standard));
    assert_eq!(document.document().unwrap().paragraph_count().unwrap(), 0);

    let pptx_encrypted = directory.path().join("encrypted.pptx");
    let mut presentation = pptx::Package::new().unwrap();
    presentation
        .presentation_mut()
        .unwrap()
        .add_slide()
        .unwrap();
    presentation
        .save_encrypted(&pptx_encrypted, PASSWORD, Mode::Agile)
        .unwrap();
    let presentation = pptx::Package::open_with_password(&pptx_encrypted, PASSWORD).unwrap();
    assert_eq!(presentation.encryption(), Some(Mode::Agile));
    assert_eq!(
        presentation.presentation().unwrap().slide_count().unwrap(),
        1
    );

    let xlsx_encrypted = directory.path().join("encrypted.xlsx");
    let mut workbook = xlsx::Workbook::create().unwrap();
    workbook
        .save_encrypted(&xlsx_encrypted, PASSWORD, Mode::Agile)
        .unwrap();
    let workbook = xlsx::Workbook::open_with_password(&xlsx_encrypted, PASSWORD).unwrap();
    assert_eq!(workbook.encryption(), Some(Mode::Agile));
    assert_eq!(workbook.worksheet_count(), 1);
}

#[test]
fn a_wrong_password_is_rejected_without_mutating_the_source() {
    let directory = tempfile::tempdir().unwrap();
    let encrypted = directory.path().join("encrypted.docx");
    let mut document = docx::Package::new().unwrap();
    document
        .save_encrypted(&encrypted, PASSWORD, Mode::Agile)
        .unwrap();
    let before = std::fs::read(&encrypted).unwrap();

    assert!(matches!(
        docx::Package::open_with_password(&encrypted, "wrong password"),
        Err(OoxmlError::Crypto(CryptoError::Password))
    ));
    assert_eq!(std::fs::read(encrypted).unwrap(), before);
}

#[test]
fn encrypted_sources_refuse_plain_output_before_touching_the_target() {
    let directory = tempfile::tempdir().unwrap();
    let sentinel = b"existing output must survive";

    let encrypted = directory.path().join("source.docx");
    let target = directory.path().join("target.docx");
    let mut document = docx::Package::new().unwrap();
    document
        .save_encrypted(&encrypted, PASSWORD, Mode::Agile)
        .unwrap();
    let mut document = docx::Package::open_with_password(&encrypted, PASSWORD).unwrap();
    std::fs::write(&target, sentinel).unwrap();
    assert!(document.save(&target).is_err());
    assert_eq!(std::fs::read(&target).unwrap(), sentinel);
    let mut stream = std::io::Cursor::new(sentinel.to_vec());
    assert!(document.to_stream(&mut stream).is_err());
    assert_eq!(stream.get_ref(), sentinel);
    stream.set_position(0);
    document.to_plain_stream(&mut stream).unwrap();
    assert_eq!(&stream.get_ref()[..2], b"PK");
    document.save_plain(&target).unwrap();
    assert_eq!(&std::fs::read(&target).unwrap()[..2], b"PK");

    let encrypted = directory.path().join("source.pptx");
    let target = directory.path().join("target.pptx");
    let mut presentation = pptx::Package::new().unwrap();
    presentation
        .save_encrypted(&encrypted, PASSWORD, Mode::Agile)
        .unwrap();
    let mut presentation = pptx::Package::open_with_password(&encrypted, PASSWORD).unwrap();
    std::fs::write(&target, sentinel).unwrap();
    assert!(presentation.save(&target).is_err());
    assert_eq!(std::fs::read(&target).unwrap(), sentinel);
    presentation.save_plain(&target).unwrap();
    assert_eq!(&std::fs::read(&target).unwrap()[..2], b"PK");

    let encrypted = directory.path().join("source.xlsx");
    let target = directory.path().join("target.xlsx");
    let mut workbook = xlsx::Workbook::create().unwrap();
    workbook
        .save_encrypted(&encrypted, PASSWORD, Mode::Agile)
        .unwrap();
    let mut workbook = xlsx::Workbook::open_with_password(&encrypted, PASSWORD).unwrap();
    std::fs::write(&target, sentinel).unwrap();
    assert!(workbook.save(&target).is_err());
    assert_eq!(std::fs::read(&target).unwrap(), sentinel);
    workbook.save_plain(&target).unwrap();
    assert_eq!(&std::fs::read(&target).unwrap()[..2], b"PK");
}

#[test]
fn preserved_mode_reencrypts_without_retaining_a_password() {
    let directory = tempfile::tempdir().unwrap();
    let source = directory.path().join("source.pptx");
    let changed = directory.path().join("changed.pptx");
    let mut presentation = pptx::Package::new().unwrap();
    presentation
        .presentation_mut()
        .unwrap()
        .add_slide()
        .unwrap();
    presentation
        .save_encrypted(&source, PASSWORD, Mode::Standard)
        .unwrap();

    let mut presentation = pptx::Package::open_with_password(&source, PASSWORD).unwrap();
    presentation
        .save_reencrypted(&changed, NEW_PASSWORD)
        .unwrap();
    assert!(pptx::Package::open_with_password(&changed, PASSWORD).is_err());
    let reopened = pptx::Package::open_with_password(&changed, NEW_PASSWORD).unwrap();
    assert_eq!(reopened.encryption(), Some(Mode::Standard));
    assert_eq!(reopened.presentation().unwrap().slide_count().unwrap(), 1);
}

#[test]
fn bounded_reader_rejects_input_before_package_parsing() {
    let limits = Limits {
        max_input_bytes: 4,
        ..Limits::default()
    };
    let error =
        docx::Package::from_reader_with(std::io::Cursor::new(vec![0u8; 5]), PASSWORD, &limits)
            .err()
            .expect("bounded reader must fail");
    assert!(matches!(
        error,
        OoxmlError::Crypto(CryptoError::Limit {
            resource: "input",
            actual: 5,
            maximum: 4,
        })
    ));
}
