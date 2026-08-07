use std::io::Cursor;

use litchi_opc::{OpcError, ReadResource};
use litchi_pptx::{Error, Package, ReadLimits};
use tempfile::NamedTempFile;

fn input_limit(bytes: usize) -> ReadLimits {
    ReadLimits::builder()
        .max_input_bytes(u64::try_from(bytes).expect("package size fits in u64"))
        .expect("valid input limit")
        .build()
        .expect("valid resource policy")
}

fn presentation_bytes() -> Vec<u8> {
    Package::new()
        .expect("new PPTX package")
        .to_bytes()
        .expect("serialize PPTX package")
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
fn direct_pptx_constructors_apply_exact_and_over_input_limits() {
    let bytes = presentation_bytes();
    let exact = input_limit(bytes.len());
    let over = input_limit(bytes.len() - 1);

    Package::from_bytes(&bytes).expect("default slice accepts valid PPTX");
    Package::from_bytes_with_limits(&bytes, exact).expect("exact slice limit accepts PPTX");
    assert_input_limit(
        Package::from_bytes_with_limits(&bytes, over)
            .err()
            .expect("slice must propagate OPC input limit"),
        bytes.len(),
        bytes.len() - 1,
    );

    Package::from_vec_with_limits(bytes.clone(), exact).expect("exact owned limit accepts PPTX");
    assert_input_limit(
        Package::from_vec_with_limits(bytes.clone(), over)
            .err()
            .expect("owned bytes must propagate OPC input limit"),
        bytes.len(),
        bytes.len() - 1,
    );
    Package::from_reader_with_limits(Cursor::new(bytes.clone()), exact)
        .expect("exact reader limit accepts PPTX");
    assert_input_limit(
        Package::from_reader_with_limits(Cursor::new(bytes.clone()), over)
            .err()
            .expect("reader must propagate OPC input limit"),
        bytes.len(),
        bytes.len() - 1,
    );

    let file = NamedTempFile::new().expect("temporary PPTX path");
    std::fs::write(file.path(), &bytes).expect("write PPTX fixture");
    Package::open_with_limits(file.path(), exact).expect("exact path limit accepts PPTX");
    assert_input_limit(
        Package::open_with_limits(file.path(), over)
            .err()
            .expect("path must propagate OPC input limit"),
        bytes.len(),
        bytes.len() - 1,
    );
}
