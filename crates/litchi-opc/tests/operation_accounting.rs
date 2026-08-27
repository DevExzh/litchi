#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design"
)]

//! Focused propagation checks for the bounded OPC accounting surface.

use std::io::{self, Write};
use std::sync::Arc;

use litchi_core::OwnedSource;
use litchi_opc::{OpcError, OpcOperationAccounting, PackURI, SourceBackedPackage};
use soapberry_zip::office::StreamingArchiveWriter;

const CONTENT_TYPES: &[u8] = br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>"#;
const ROOT_RELS: &[u8] = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#;
const PART: &[u8] = b"<document><paragraph>accounted</paragraph></document>";

fn source_bytes(deflated: bool) -> Vec<u8> {
    let mut writer = StreamingArchiveWriter::new();
    writer
        .write_stored("[Content_Types].xml", CONTENT_TYPES)
        .unwrap();
    writer.write_stored("_rels/.rels", ROOT_RELS).unwrap();
    if deflated {
        writer.write_deflated("word/document.xml", PART).unwrap();
    } else {
        writer.write_stored("word/document.xml", PART).unwrap();
    }
    writer.finish_to_bytes().unwrap()
}

fn open(bytes: Vec<u8>) -> SourceBackedPackage {
    SourceBackedPackage::from_read_at(Arc::new(OwnedSource::new(bytes))).unwrap()
}

fn document_uri() -> PackURI {
    PackURI::new("/word/document.xml").unwrap()
}

#[test]
fn cold_stored_and_deflated_part_reads_are_accounted_but_cache_hits_are_not() {
    for deflated in [false, true] {
        let package = open(source_bytes(deflated));
        let part = package.part(&document_uri()).unwrap();
        let mut cold = OpcOperationAccounting::default();
        let data = part.data_with_accounting(&mut cold).unwrap();
        assert_eq!(data.as_bytes(), PART);
        if deflated {
            assert!(cold.compressed_deflate_payload_bytes_read() > 0);
            assert_eq!(cold.stored_payload_bytes_read(), 0);
            assert_eq!(cold.stored_payload_bytes_accepted(), 0);
            assert_eq!(cold.deflate_bytes_produced(), PART.len() as u64);
            assert_eq!(cold.deflate_bytes_accepted(), PART.len() as u64);
        } else {
            assert_eq!(cold.compressed_deflate_payload_bytes_read(), 0);
            assert_eq!(cold.stored_payload_bytes_read(), PART.len() as u64);
            assert_eq!(cold.stored_payload_bytes_accepted(), PART.len() as u64);
            assert_eq!(cold.deflate_bytes_produced(), 0);
            assert_eq!(cold.deflate_bytes_accepted(), 0);
        }

        let mut hit = OpcOperationAccounting::default();
        assert_eq!(
            part.data_with_accounting(&mut hit).unwrap().as_bytes(),
            PART
        );
        assert_eq!(hit, OpcOperationAccounting::default());
    }
}

fn corrupt_document_crc(mut source: Vec<u8>) -> Vec<u8> {
    let descriptor_signature = [0x50, 0x4b, 0x07, 0x08];
    if let Some(offset) = source
        .windows(descriptor_signature.len())
        .rposition(|window| window == descriptor_signature)
    {
        source[offset + descriptor_signature.len()] ^= 1;
    } else {
        let crc = soapberry_zip::crc32(PART).to_le_bytes();
        let offset = source
            .windows(crc.len())
            .rposition(|window| window == crc)
            .expect("document CRC in local or central ZIP headers");
        source[offset] ^= 1;
    }
    source
}

#[test]
fn crc_failure_keeps_the_cold_read_counters_and_does_not_cache() {
    let package = open(corrupt_document_crc(source_bytes(false)));
    let part = package.part(&document_uri()).unwrap();
    let mut accounting = OpcOperationAccounting::default();
    let error = part.data_with_accounting(&mut accounting).unwrap_err();
    assert!(
        matches!(error, OpcError::ZipError(_)),
        "unexpected CRC error: {error:?}"
    );
    assert_eq!(accounting.stored_payload_bytes_read(), PART.len() as u64);
    assert_eq!(accounting.stored_payload_bytes_accepted(), 0);

    let mut retry = OpcOperationAccounting::default();
    assert!(part.data_with_accounting(&mut retry).is_err());
    assert_eq!(retry.stored_payload_bytes_read(), PART.len() as u64);
}

#[test]
fn exact_source_publication_reports_actual_raw_and_output_acceptance() {
    let source = source_bytes(false);
    let package = open(source.clone());
    let artifact = package.source_artifact();
    let mut output = Vec::new();
    let mut accounting = OpcOperationAccounting::default();
    artifact
        .write_to_stream_with_accounting(&mut output, &mut accounting)
        .unwrap();
    assert_eq!(output, source);
    assert_eq!(
        accounting.raw_unchanged_source_bytes_accepted(),
        source.len() as u64
    );
    assert_eq!(accounting.output_bytes_accepted(), source.len() as u64);
}

#[test]
fn singular_changed_overlay_distinguishes_raw_and_generated_payloads() {
    for deflated in [false, true] {
        let package = open(source_bytes(deflated));
        let mut output = Vec::new();
        let mut accounting = OpcOperationAccounting::default();
        package
            .write_part_overlay_to_stream_with_accounting(
                &mut output,
                &document_uri(),
                b"<changed/>".to_vec(),
                &mut accounting,
            )
            .unwrap();
        assert!(accounting.raw_unchanged_source_bytes_accepted() > 0);
        assert!(accounting.output_bytes_accepted() > 0);
        if deflated {
            assert!(accounting.generated_deflate_payload_bytes_emitted() > 0);
            assert_eq!(accounting.stored_payload_bytes_emitted(), 0);
        } else {
            assert!(accounting.stored_payload_bytes_emitted() > 0);
            assert_eq!(accounting.generated_deflate_payload_bytes_emitted(), 0);
        }
    }
}

#[derive(Debug)]
struct FailingSink {
    bytes: Vec<u8>,
    remaining: usize,
}

impl Write for FailingSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if self.remaining == 0 {
            return Err(io::Error::new(io::ErrorKind::BrokenPipe, "stop"));
        }
        let accepted = bytes.len().min(self.remaining);
        self.bytes.extend_from_slice(&bytes[..accepted]);
        self.remaining -= accepted;
        Ok(accepted)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[test]
fn exact_source_partial_sink_preserves_checked_progress() {
    let source = source_bytes(false);
    let package = open(source);
    let artifact = package.source_artifact();
    let mut sink = FailingSink {
        bytes: Vec::new(),
        remaining: 7,
    };
    let mut accounting = OpcOperationAccounting::default();
    let error = artifact
        .write_to_stream_with_accounting(&mut sink, &mut accounting)
        .unwrap_err();
    assert!(matches!(
        error,
        OpcError::IncompleteOutput { written: 7, .. }
    ));
    assert_eq!(accounting.output_bytes_accepted(), 7);
    assert_eq!(accounting.raw_unchanged_source_bytes_accepted(), 7);
    assert_eq!(sink.bytes.len(), 7);
}

#[test]
fn changed_overlay_partial_sink_keeps_raw_and_output_progress() {
    let package = open(source_bytes(false));
    let mut sink = FailingSink {
        bytes: Vec::new(),
        remaining: 19,
    };
    let mut accounting = OpcOperationAccounting::default();
    let error = package
        .write_part_overlay_to_stream_with_accounting(
            &mut sink,
            &document_uri(),
            b"<changed/>".to_vec(),
            &mut accounting,
        )
        .unwrap_err();
    assert!(matches!(
        error,
        OpcError::IncompleteOutput { written: 19, .. }
    ));
    assert_eq!(accounting.output_bytes_accepted(), 19);
    assert!(accounting.raw_unchanged_source_bytes_accepted() > 0);
    assert!(accounting.raw_unchanged_source_bytes_accepted() <= 19);
    assert_eq!(sink.bytes.len(), 19);
}

#[test]
fn same_flight_loader_owns_physical_accounting() {
    let package = Arc::new(open(source_bytes(true)));
    let start = Arc::new(std::sync::Barrier::new(3));

    let first_package = Arc::clone(&package);
    let first_start = Arc::clone(&start);
    let first = std::thread::spawn(move || {
        first_start.wait();
        let uri = document_uri();
        let mut accounting = OpcOperationAccounting::default();
        first_package
            .part(&uri)
            .unwrap()
            .data_with_accounting(&mut accounting)
            .unwrap();
        accounting
    });

    let second_package = Arc::clone(&package);
    let second_start = Arc::clone(&start);
    let second = std::thread::spawn(move || {
        second_start.wait();
        let uri = document_uri();
        let mut accounting = OpcOperationAccounting::default();
        second_package
            .part(&uri)
            .unwrap()
            .data_with_accounting(&mut accounting)
            .unwrap();
        accounting
    });

    start.wait();
    let first = first.join().unwrap();
    let second = second.join().unwrap();

    let mut expected = OpcOperationAccounting::default();
    open(source_bytes(true))
        .part(&document_uri())
        .unwrap()
        .data_with_accounting(&mut expected)
        .unwrap();
    let empty = OpcOperationAccounting::default();

    assert!((first == expected && second == empty) || (first == empty && second == expected));
}

#[test]
fn deflate_crc_failure_retains_accounting_and_retries_physical_work() {
    let source = corrupt_document_crc(source_bytes(true));
    let package = open(source);

    let mut first = OpcOperationAccounting::default();
    assert!(
        package
            .part(&document_uri())
            .unwrap()
            .data_with_accounting(&mut first)
            .is_err()
    );
    assert!(first.compressed_deflate_payload_bytes_read() > 0);
    assert!(first.deflate_bytes_produced() > 0);
    assert_eq!(first.deflate_bytes_accepted(), 0);

    let mut retry = OpcOperationAccounting::default();
    assert!(
        package
            .part(&document_uri())
            .unwrap()
            .data_with_accounting(&mut retry)
            .is_err()
    );
    assert!(retry.compressed_deflate_payload_bytes_read() > 0);
    assert!(retry.deflate_bytes_produced() > 0);
    assert_eq!(retry.deflate_bytes_accepted(), 0);
    assert_eq!(
        retry.compressed_deflate_payload_bytes_read(),
        first.compressed_deflate_payload_bytes_read()
    );
    assert_eq!(
        retry.deflate_bytes_produced(),
        first.deflate_bytes_produced()
    );
    assert_eq!(
        retry.deflate_bytes_accepted(),
        first.deflate_bytes_accepted()
    );
}

#[test]
fn singular_exact_noop_overlay_accounts_raw_publication_and_cache_selection() {
    for deflated in [false, true] {
        let source = source_bytes(deflated);
        let uri = document_uri();

        let cold = open(source.clone());
        let mut cold_output = Vec::new();
        let mut cold_accounting = OpcOperationAccounting::default();
        cold.write_part_overlay_to_stream_with_accounting(
            &mut cold_output,
            &uri,
            PART.to_vec(),
            &mut cold_accounting,
        )
        .unwrap();

        assert_eq!(cold_output, source);
        assert_eq!(
            cold_accounting.raw_unchanged_source_bytes_accepted(),
            source.len() as u64
        );
        assert_eq!(cold_accounting.output_bytes_accepted(), source.len() as u64);
        assert_eq!(cold_accounting.generated_deflate_payload_bytes_emitted(), 0);
        assert_eq!(cold_accounting.stored_payload_bytes_emitted(), 0);
        assert_eq!(cold_accounting.precompressed_payload_bytes_emitted(), 0);
        if deflated {
            assert!(cold_accounting.compressed_deflate_payload_bytes_read() > 0);
            assert_eq!(cold_accounting.stored_payload_bytes_read(), 0);
        } else {
            assert!(cold_accounting.stored_payload_bytes_read() > 0);
            assert_eq!(cold_accounting.compressed_deflate_payload_bytes_read(), 0);
        }

        let warm = open(source.clone());
        warm.part(&uri).unwrap().data().unwrap();
        let mut warm_output = Vec::new();
        let mut warm_accounting = OpcOperationAccounting::default();
        warm.write_part_overlay_to_stream_with_accounting(
            &mut warm_output,
            &uri,
            PART.to_vec(),
            &mut warm_accounting,
        )
        .unwrap();

        assert_eq!(warm_output, source);
        assert_eq!(
            warm_accounting.raw_unchanged_source_bytes_accepted(),
            source.len() as u64
        );
        assert_eq!(warm_accounting.output_bytes_accepted(), source.len() as u64);
        assert_eq!(warm_accounting.generated_deflate_payload_bytes_emitted(), 0);
        assert_eq!(warm_accounting.stored_payload_bytes_emitted(), 0);
        assert_eq!(warm_accounting.precompressed_payload_bytes_emitted(), 0);
        assert_eq!(warm_accounting.compressed_deflate_payload_bytes_read(), 0);
        assert_eq!(warm_accounting.stored_payload_bytes_read(), 0);
    }
}

#[test]
fn streamed_stored_and_deflated_parts_are_accounted_without_cache_materialization() {
    for deflated in [false, true] {
        let package = open(source_bytes(deflated));
        let part = package
            .part(&document_uri())
            .expect("document part should be present");
        let mut sink = Vec::new();
        let mut accounting = OpcOperationAccounting::default();

        let written = part
            .stream_to_with_accounting(&mut sink, &mut accounting)
            .expect("selected part should stream");

        assert_eq!(written, PART.len() as u64);
        assert_eq!(sink, PART);
        assert_eq!(accounting.output_bytes_accepted(), PART.len() as u64);
        if deflated {
            assert!(accounting.compressed_deflate_payload_bytes_read() > 0);
            assert!(accounting.deflate_bytes_produced() > 0);
            assert_eq!(accounting.deflate_bytes_accepted(), PART.len() as u64);
            assert_eq!(accounting.stored_payload_bytes_read(), 0);
        } else {
            assert!(accounting.stored_payload_bytes_read() > 0);
            assert_eq!(
                accounting.stored_payload_bytes_accepted(),
                PART.len() as u64
            );
            assert_eq!(accounting.compressed_deflate_payload_bytes_read(), 0);
        }
        assert_eq!(accounting.raw_unchanged_source_bytes_accepted(), 0);
        assert_eq!(accounting.generated_deflate_payload_bytes_emitted(), 0);
        assert_eq!(accounting.stored_payload_bytes_emitted(), 0);
        assert_eq!(accounting.precompressed_payload_bytes_emitted(), 0);

        let materialized = part
            .data()
            .expect("data read should remain available after streaming")
            .into_arc()
            .expect("unmanaged part data should expose its arc");
        assert_eq!(materialized.as_slice(), PART);
    }
}

#[test]
fn streamed_partial_sink_preserves_physical_and_output_accounting() {
    for deflated in [false, true] {
        let package = open(source_bytes(deflated));
        let part = package
            .part(&document_uri())
            .expect("document part should be present");
        let mut sink = FailingSink {
            bytes: Vec::new(),
            remaining: 3,
        };
        let mut accounting = OpcOperationAccounting::default();

        let error = part
            .stream_to_with_accounting(&mut sink, &mut accounting)
            .expect_err("the bounded sink should fail");

        match error {
            OpcError::IncompleteOutput { written, .. } => assert_eq!(written, 3),
            other => panic!("unexpected stream error: {other:?}"),
        }
        assert_eq!(sink.bytes.len(), 3);
        assert_eq!(accounting.output_bytes_accepted(), 3);
        if deflated {
            assert!(accounting.compressed_deflate_payload_bytes_read() >= 3);
            assert!(accounting.deflate_bytes_produced() >= 3);
        } else {
            assert!(accounting.stored_payload_bytes_read() >= 3);
        }
        assert_eq!(accounting.raw_unchanged_source_bytes_accepted(), 0);
        assert_eq!(accounting.generated_deflate_payload_bytes_emitted(), 0);
    }
}

#[test]
fn repeated_streaming_repeats_physical_work_without_cache_hits() {
    let package = open(source_bytes(true));
    let part = package
        .part(&document_uri())
        .expect("document part should be present");
    let mut first_sink = Vec::new();
    let mut first = OpcOperationAccounting::default();
    part.stream_to_with_accounting(&mut first_sink, &mut first)
        .expect("first selected stream should succeed");

    let mut second_sink = Vec::new();
    let mut second = OpcOperationAccounting::default();
    part.stream_to_with_accounting(&mut second_sink, &mut second)
        .expect("second selected stream should succeed");

    assert_eq!(first_sink, PART);
    assert_eq!(second_sink, PART);
    assert!(second.compressed_deflate_payload_bytes_read() > 0);
    assert!(second.deflate_bytes_produced() > 0);
    assert_eq!(second.output_bytes_accepted(), PART.len() as u64);
}

#[test]
fn streaming_bypasses_cache_and_data_remains_cold() {
    let package = open(source_bytes(false));
    let part = package
        .part(&document_uri())
        .expect("document part should be present");
    let before = package.cache_diagnostics();
    let mut sink = Vec::new();
    part.stream_to(&mut sink)
        .expect("selected part should stream");
    let after_stream = package.cache_diagnostics();
    assert_eq!(before, after_stream);

    let data = part
        .data()
        .expect("data read should remain available after streaming")
        .into_arc()
        .expect("unmanaged part data should expose its arc");
    assert_eq!(data.as_slice(), PART);
    let after_data = package.cache_diagnostics();
    assert_ne!(after_stream, after_data);
}
