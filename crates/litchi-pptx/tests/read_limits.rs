#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use std::io::Cursor;

use litchi_opc::{OpcError, SourceLimitKind};
use litchi_pptx::{Error, Package, SourceLimits};
use tempfile::NamedTempFile;

fn input_limit(bytes: usize) -> SourceLimits {
    SourceLimits::builder()
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
        Error::Opc(OpcError::Limit {
            kind: SourceLimitKind::InputBytes,
            observed,
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
    Package::from_bytes_with_options(
        &bytes,
        litchi_pptx::OpenOptions::default().source_limits(exact),
    )
    .expect("exact slice limit accepts PPTX");
    assert_input_limit(
        Package::from_bytes_with_options(
            &bytes,
            litchi_pptx::OpenOptions::default().source_limits(over),
        )
        .expect_err("slice must propagate OPC input limit"),
        bytes.len(),
        bytes.len() - 1,
    );

    Package::from_reader_with_options(
        Cursor::new(bytes.clone()),
        litchi_pptx::OpenOptions::default().source_limits(exact),
    )
    .expect("exact owned limit accepts PPTX");
    assert_input_limit(
        Package::from_reader_with_options(
            Cursor::new(bytes.clone()),
            litchi_pptx::OpenOptions::default().source_limits(over),
        )
        .expect_err("owned bytes must propagate OPC input limit"),
        bytes.len(),
        bytes.len() - 1,
    );
    Package::from_reader_with_options(
        Cursor::new(bytes.clone()),
        litchi_pptx::OpenOptions::default().source_limits(exact),
    )
    .expect("exact reader limit accepts PPTX");
    assert_input_limit(
        Package::from_reader_with_options(
            Cursor::new(bytes.clone()),
            litchi_pptx::OpenOptions::default().source_limits(over),
        )
        .expect_err("reader must propagate OPC input limit"),
        bytes.len(),
        bytes.len() - 1,
    );

    let file = NamedTempFile::new().expect("temporary PPTX path");
    std::fs::write(file.path(), &bytes).expect("write PPTX fixture");
    Package::open_with_options(
        file.path(),
        litchi_pptx::OpenOptions::default().source_limits(exact),
    )
    .expect("exact path limit accepts PPTX");
    assert_input_limit(
        Package::open_with_options(
            file.path(),
            litchi_pptx::OpenOptions::default().source_limits(over),
        )
        .expect_err("path must propagate OPC input limit"),
        bytes.len(),
        bytes.len() - 1,
    );
}
