#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "bounded writer tests use panic-on-fixture-failure assertions"
)]

use std::io::{self, Write};

use sha2::{Digest, Sha256};

use super::BoundedVecWriter;

#[test]
fn bounded_writer_rejects_before_reserving_and_keeps_prefix() {
    let mut writer = BoundedVecWriter::new(3);
    assert_eq!(writer.write(b"abc").expect("prefix write"), 3);
    let prefix = writer.bytes.clone();
    let capacity = writer.bytes.capacity();

    let error = writer
        .write(b"d")
        .expect_err("a write beyond the archive bound must fail");
    assert_eq!(error.kind(), io::ErrorKind::WriteZero);
    assert!(writer.exceeded);
    assert!(writer.allocation_failure.is_none());
    assert_eq!(writer.bytes, prefix);
    assert_eq!(writer.bytes.capacity(), capacity);
}

#[test]
fn bounded_writer_rejects_an_oversized_first_write_without_allocation() {
    let mut writer = BoundedVecWriter::new(2);
    let capacity = writer.bytes.capacity();

    let error = writer
        .write(b"abc")
        .expect_err("the first write must respect the archive bound");
    assert_eq!(error.kind(), io::ErrorKind::WriteZero);
    assert!(writer.exceeded);
    assert!(writer.allocation_failure.is_none());
    assert!(writer.bytes.is_empty());
    assert_eq!(writer.bytes.capacity(), capacity);
}

#[test]
fn bounded_writer_allows_empty_write_at_zero_limit() {
    let mut writer = BoundedVecWriter::new(0);
    let pointer = writer.bytes.as_ptr();
    let capacity = writer.bytes.capacity();

    assert_eq!(writer.write(&[]).expect("an empty write is harmless"), 0);
    assert!(!writer.exceeded);
    assert!(writer.allocation_failure.is_none());
    assert!(writer.bytes.is_empty());
    assert_eq!(writer.bytes.as_ptr(), pointer);
    assert_eq!(writer.bytes.capacity(), capacity);

    let error = writer
        .write(b"x")
        .expect_err("a nonempty write cannot fit a zero limit");
    assert_eq!(error.kind(), io::ErrorKind::WriteZero);
    assert!(writer.exceeded);
    assert!(writer.allocation_failure.is_none());
    assert!(writer.bytes.is_empty());
    assert_eq!(writer.bytes.capacity(), capacity);
}

#[test]
fn bounded_writer_allows_empty_write_after_reaching_the_limit() {
    let mut writer = BoundedVecWriter::new(3);
    assert_eq!(writer.write(b"abc").expect("full write"), 3);
    let pointer = writer.bytes.as_ptr();
    let capacity = writer.bytes.capacity();

    assert_eq!(writer.write(&[]).expect("empty write at the bound"), 0);
    assert!(!writer.exceeded);
    assert!(writer.allocation_failure.is_none());
    assert_eq!(writer.bytes, b"abc");
    assert_eq!(writer.bytes.as_ptr(), pointer);
    assert_eq!(writer.bytes.capacity(), capacity);
}

#[test]
fn bounded_writer_keeps_successful_storage_within_the_limit() {
    let limit = 64;
    let chunks: &[&[u8]] = &[b"a", b"bc", b"defgh", b"ijklmnop", b"qrstuvwxyz"];
    let mut writer = BoundedVecWriter::new(limit);
    let mut expected = Vec::new();
    let mut previous_capacity = writer.bytes.capacity();

    for chunk in chunks {
        assert_eq!(writer.write(chunk).expect("bounded write"), chunk.len());
        expected.extend_from_slice(chunk);
        assert_eq!(writer.bytes, expected);
        assert!(writer.bytes.capacity() >= writer.bytes.len());
        assert!(writer.bytes.capacity() <= limit);
        assert!(writer.bytes.capacity() >= previous_capacity);
        previous_capacity = writer.bytes.capacity();
    }
}

#[test]
fn bounded_writer_into_bytes_reuses_an_exact_buffer() {
    let bytes = b"exact-capacity-buffer".to_vec();
    assert_eq!(bytes.capacity(), bytes.len());
    let pointer = bytes.as_ptr();
    let expected = bytes.clone();
    let length = bytes.len();
    let writer = BoundedVecWriter {
        bytes,
        digest: Sha256::new(),
        limit: length,
        exceeded: false,
        allocation_failure: None,
    };

    let (output, _digest) = writer.into_bytes().expect("exact buffer needs no copy");
    assert_eq!(output.as_ptr(), pointer);
    assert_eq!(output, expected);
}

#[test]
fn bounded_writer_into_bytes_allows_an_empty_archive() {
    let writer = BoundedVecWriter::new(0);
    let (output, digest) = writer.into_bytes().expect("an empty archive needs no copy");
    assert!(output.is_empty());
    let expected: [u8; 32] = Sha256::new().finalize().into();
    assert_eq!(digest, expected);
}

#[test]
fn bounded_writer_into_bytes_copies_when_spare_capacity_exists() {
    let payload = b"spare-capacity-buffer";
    let mut bytes = Vec::with_capacity(payload.len() * 4);
    bytes.extend_from_slice(payload);
    assert!(bytes.capacity() > bytes.len());
    let pointer = bytes.as_ptr();
    let length = bytes.len();
    let writer = BoundedVecWriter {
        bytes,
        digest: Sha256::new(),
        limit: length,
        exceeded: false,
        allocation_failure: None,
    };

    let (output, _digest) = writer
        .into_bytes()
        .expect("spare capacity must be compacted");
    assert_ne!(output.as_ptr(), pointer);
    assert_eq!(output, payload);
}

#[test]
fn bounded_writer_digests_exactly_the_accepted_bytes() {
    let limit = 64;
    let chunks: &[&[u8]] = &[b"a", b"bc", b"defgh", b"ijklmnop", b"qrstuvwxyz"];
    let mut writer = BoundedVecWriter::new(limit);
    let mut expected = Vec::new();
    for chunk in chunks {
        assert_eq!(writer.write(chunk).expect("bounded write"), chunk.len());
        expected.extend_from_slice(chunk);
    }
    // A write that the bound refuses must not reach the digest.
    let refused = vec![b'z'; limit];
    assert!(writer.write(&refused).is_err());

    let (output, digest) = writer.into_bytes().expect("bounded archive");
    assert_eq!(output, expected);
    let direct: [u8; 32] = Sha256::digest(&expected).into();
    assert_eq!(digest, direct);
}
