#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "publication tests use direct assertions for short-lived fixtures"
)]

use litchi_doc::{Builder, Error, Package};
use std::io::{self, Write};

struct FailingSink;

impl Write for FailingSink {
    fn write(&mut self, _bytes: &[u8]) -> io::Result<usize> {
        Err(io::Error::new(
            io::ErrorKind::BrokenPipe,
            "publication sink",
        ))
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[test]
fn package_new_and_builder_round_trip_through_exact_publication() {
    let empty = Package::new().expect("minimal package");
    let empty_bytes = empty.to_bytes().expect("minimal package bytes");
    assert!(!empty_bytes.is_empty());

    let mut empty_stream = Vec::new();
    empty
        .write_to(&mut empty_stream)
        .expect("minimal package stream");
    assert_eq!(empty_stream, empty_bytes);

    let mut builder = Builder::new();
    builder
        .add_paragraph("canonical publication")
        .expect("paragraph");
    let package = builder.build().expect("created package");
    let bytes = package.to_bytes().expect("created package bytes");

    let mut streamed = Vec::new();
    package
        .write_to(&mut streamed)
        .expect("created package stream");
    assert_eq!(streamed, bytes);

    let reopened = Package::from_bytes(&bytes).expect("round-trip package");
    assert_eq!(reopened.to_bytes().expect("round-trip bytes"), bytes);
    assert!(
        reopened
            .document()
            .text()
            .expect("round-trip text")
            .contains("canonical publication")
    );
}

#[test]
fn package_publication_keeps_sink_failures_typed() {
    let package = Package::new().expect("minimal package");
    let error = package
        .write_to(FailingSink)
        .expect_err("failing publication sink");
    assert!(matches!(error, Error::Io(_)));
}
