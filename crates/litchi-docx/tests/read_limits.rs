use std::io::Cursor;

use litchi_docx::{Error, Package, ReadLimits};
use litchi_opc::{OpcError, ReadResource};
use tempfile::NamedTempFile;

fn input_limit(bytes: usize) -> ReadLimits {
    ReadLimits::builder()
        .max_input_bytes(u64::try_from(bytes).expect("package size fits in u64"))
        .expect("valid input limit")
        .build()
        .expect("valid resource policy")
}

fn document_bytes() -> Vec<u8> {
    let mut package = Package::new().expect("new DOCX package");
    let mut output = Cursor::new(Vec::new());
    package
        .to_stream(&mut output)
        .expect("serialize DOCX package");
    output.into_inner()
}

fn assert_input_limit(error: Error, actual: usize, maximum: usize) {
    assert!(matches!(
        error,
        Error::Opc(OpcError::ReadLimit {
            resource: ReadResource::InputBytes,
            actual: observed,
            maximum: bound,
        }) if observed == actual as u64 && bound == maximum as u64
    ));
}

#[test]
fn direct_docx_constructors_apply_exact_and_over_input_limits() {
    let bytes = document_bytes();
    let exact = input_limit(bytes.len());
    let over = input_limit(bytes.len() - 1);

    Package::from_reader(Cursor::new(bytes.clone())).expect("default reader accepts valid DOCX");
    Package::from_reader_with_limits(Cursor::new(bytes.clone()), exact)
        .expect("exact reader limit accepts DOCX");
    assert_input_limit(
        Package::from_reader_with_limits(Cursor::new(bytes.clone()), over)
            .err()
            .expect("reader must propagate OPC input limit"),
        bytes.len(),
        bytes.len() - 1,
    );

    let file = NamedTempFile::new().expect("temporary DOCX path");
    std::fs::write(file.path(), &bytes).expect("write DOCX fixture");
    Package::open(file.path()).expect("default path accepts valid DOCX");
    Package::open_with_limits(file.path(), exact).expect("exact path limit accepts DOCX");
    assert_input_limit(
        Package::open_with_limits(file.path(), over)
            .err()
            .expect("path must propagate OPC input limit"),
        bytes.len(),
        bytes.len() - 1,
    );
}

#[cfg(feature = "encryption")]
#[test]
fn decrypted_docx_uses_the_supplied_opc_limits() {
    use litchi_docx::encryption::{Limits as EncryptionLimits, Mode};

    let bytes = document_bytes();
    let mut package = Package::from_reader(Cursor::new(bytes.clone())).expect("valid DOCX");
    let encrypted = package
        .to_encrypted("read-limits", Mode::Standard)
        .expect("encrypt DOCX fixture");
    let encryption_limits = EncryptionLimits {
        max_input_bytes: encrypted.len(),
        ..EncryptionLimits::default()
    };
    let exact = input_limit(bytes.len());
    let over = input_limit(bytes.len() - 1);

    Package::from_reader_with_password_and_limits(
        Cursor::new(encrypted.clone()),
        "read-limits",
        &encryption_limits,
        exact,
    )
    .expect("exact post-decryption OPC limit accepts DOCX");
    assert_input_limit(
        Package::from_reader_with_password_and_limits(
            Cursor::new(encrypted),
            "read-limits",
            &encryption_limits,
            over,
        )
        .err()
        .expect("decrypted OPC parsing must propagate caller limit"),
        bytes.len(),
        bytes.len() - 1,
    );
}
