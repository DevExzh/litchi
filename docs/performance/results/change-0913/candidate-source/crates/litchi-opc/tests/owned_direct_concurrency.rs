#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "focused concurrency assertions intentionally panic on fixture errors"
)]

//! Concurrency contracts for the owned-source deferred OPC payload path.
//!
//! The package and its cloned parts are `Send + Sync`; the first force must
//! therefore have one shared `OnceLock` outcome.  These tests exercise that
//! boundary through the public package API while the ZIP reader uses its
//! specialized owned-source path internally.

use std::sync::{Arc, Barrier};
use std::thread;

use litchi_opc::{OpcError, OpcPackage, PackURI, PackageWriter, ReadLimits, ReadResource};
use soapberry_zip::office::StreamingArchiveWriter;

const MANIFEST: &[u8] = br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="bin" ContentType="application/octet-stream"/></Types>"#;
const ROOT_RELATIONSHIPS: &[u8] =
    br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>"#;
const GOOD_MEMBER: &str = "custom/good.xml";
const BAD_MEMBER: &str = "custom/bad.bin";
const GOOD_URI: &str = "/custom/good.xml";
const BAD_URI: &str = "/custom/bad.bin";
const GOOD_PAYLOAD: &[u8] = b"<good/>";
const BAD_PAYLOAD: &[u8] = b"lazy-corrupt-payload";
const WORKERS: usize = 12;

fn uri(value: &str) -> PackURI {
    PackURI::new(value).expect("valid test part URI")
}

fn archive() -> Vec<u8> {
    let mut writer = StreamingArchiveWriter::new();
    writer
        .write_stored("[Content_Types].xml", MANIFEST)
        .expect("write manifest");
    writer
        .write_stored("_rels/.rels", ROOT_RELATIONSHIPS)
        .expect("write root relationships");
    writer
        .write_deflated_sized(GOOD_MEMBER, GOOD_PAYLOAD)
        .expect("write good member");
    writer
        .write_stored(BAD_MEMBER, BAD_PAYLOAD)
        .expect("write bad member");
    writer.finish_to_bytes().expect("finish test archive")
}

fn corrupt_bad_member(mut bytes: Vec<u8>) -> Vec<u8> {
    let offset = bytes
        .windows(BAD_PAYLOAD.len())
        .position(|window| window == BAD_PAYLOAD)
        .expect("stored test payload must be present in the archive");
    bytes[offset] ^= 1;
    bytes
}

fn owned_package(source: Arc<Vec<u8>>) -> OpcPackage {
    OpcPackage::from_shared_vec_with_limits(source, ReadLimits::default())
        .expect("open shared owned package")
}

#[test]
fn concurrent_first_force_shares_decoded_allocation_across_clones_and_keeps_source_exact() {
    let source = Arc::new(archive());
    let package = owned_package(Arc::clone(&source));
    let cloned = package.clone();
    let part_name = uri(GOOD_URI);
    let barrier = Arc::new(Barrier::new(WORKERS + 1));

    let decoded: Vec<Arc<Vec<u8>>> = thread::scope(|scope| {
        let mut handles = Vec::with_capacity(WORKERS);
        for worker in 0..WORKERS {
            let barrier = Arc::clone(&barrier);
            let target = if worker % 2 == 0 { &package } else { &cloned };
            let part_name = part_name.clone();
            handles.push(scope.spawn(move || {
                barrier.wait();
                let part = target
                    .get_part(&part_name)
                    .expect("good part must decode concurrently");
                let blob = part.blob_arc();
                assert_eq!(blob.as_slice(), GOOD_PAYLOAD);
                blob
            }));
        }
        barrier.wait();
        handles
            .into_iter()
            .map(|handle| handle.join().expect("worker must not panic"))
            .collect()
    });

    assert_eq!(decoded.len(), WORKERS);
    for blob in &decoded {
        assert_eq!(blob.as_slice(), GOOD_PAYLOAD);
        assert!(Arc::ptr_eq(blob, &decoded[0]));
    }
    assert_eq!(
        package.deferred_decode_counters(),
        Some((1, GOOD_PAYLOAD.len() as u64))
    );
    assert_eq!(
        cloned.deferred_decode_counters(),
        Some((1, GOOD_PAYLOAD.len() as u64))
    );

    let package_source = package
        .exact_source_shared()
        .expect("owned package must retain exact source");
    let clone_source = cloned
        .exact_source_shared()
        .expect("cloned package must retain exact source");
    assert!(Arc::ptr_eq(&package_source, &source));
    assert!(Arc::ptr_eq(&clone_source, &source));
    assert_eq!(
        PackageWriter::to_bytes(&package).unwrap(),
        source.as_slice()
    );
    assert_eq!(PackageWriter::to_bytes(&cloned).unwrap(), source.as_slice());
}

#[test]
fn concurrent_corrupt_first_force_records_one_stable_refusal_across_clones() {
    let source = Arc::new(corrupt_bad_member(archive()));
    let package = owned_package(Arc::clone(&source));
    let cloned = package.clone();
    let part_name = uri(BAD_URI);
    let barrier = Arc::new(Barrier::new(WORKERS + 1));

    let refusals: Vec<String> = thread::scope(|scope| {
        let mut handles = Vec::with_capacity(WORKERS);
        for worker in 0..WORKERS {
            let barrier = Arc::clone(&barrier);
            let target = if worker % 2 == 0 { &package } else { &cloned };
            let part_name = part_name.clone();
            handles.push(scope.spawn(move || {
                barrier.wait();
                match target.get_part(&part_name) {
                    Ok(_) => panic!("corrupt part must refuse on first access"),
                    Err(error) => {
                        assert!(matches!(error, OpcError::ZipError(_)));
                        error.to_string()
                    },
                }
            }));
        }
        barrier.wait();
        handles
            .into_iter()
            .map(|handle| handle.join().expect("worker must not panic"))
            .collect()
    });

    assert_eq!(refusals.len(), WORKERS);
    for refusal in &refusals {
        assert_eq!(refusal, &refusals[0]);
    }
    for target in [&package, &cloned] {
        for _ in 0..4 {
            let refusal = match target.get_part(&part_name) {
                Ok(_) => panic!("recorded corrupt refusal must remain stable"),
                Err(error) => {
                    assert!(matches!(error, OpcError::ZipError(_)));
                    error.to_string()
                },
            };
            assert_eq!(refusal, refusals[0]);
        }
    }
    assert_eq!(package.deferred_decode_counters(), Some((0, 0)));
    assert_eq!(cloned.deferred_decode_counters(), Some((0, 0)));
    assert_eq!(
        PackageWriter::to_bytes(&package).unwrap(),
        source.as_slice()
    );
    assert_eq!(PackageWriter::to_bytes(&cloned).unwrap(), source.as_slice());
}

#[test]
fn owned_direct_decode_respects_exact_declared_part_and_total_boundaries() {
    let source = Arc::new(archive());
    let exact_limits = ReadLimits::builder()
        .max_part_bytes(BAD_PAYLOAD.len() as u64)
        .expect("positive part limit")
        .max_total_part_bytes((GOOD_PAYLOAD.len() + BAD_PAYLOAD.len()) as u64)
        .expect("positive total limit")
        .build()
        .expect("consistent exact limits");
    let package = OpcPackage::from_shared_vec_with_limits(Arc::clone(&source), exact_limits)
        .expect("declared exact boundaries admit the archive");
    assert_eq!(
        package.get_part(&uri(GOOD_URI)).unwrap().blob(),
        GOOD_PAYLOAD
    );
    assert_eq!(
        package.deferred_decode_counters(),
        Some((1, GOOD_PAYLOAD.len() as u64))
    );
    assert_eq!(
        PackageWriter::to_bytes(&package).unwrap(),
        source.as_slice()
    );

    let part_under = ReadLimits::builder()
        .max_part_bytes((BAD_PAYLOAD.len() - 1) as u64)
        .expect("positive part limit")
        .max_total_part_bytes((GOOD_PAYLOAD.len() + BAD_PAYLOAD.len()) as u64)
        .expect("positive total limit")
        .build()
        .expect("consistent part-under limit");
    assert!(matches!(
        OpcPackage::from_shared_vec_with_limits(Arc::clone(&source), part_under),
        Err(OpcError::ReadLimit {
            resource: ReadResource::PartBytes,
            actual,
            maximum,
        }) if actual == BAD_PAYLOAD.len() as u64 && maximum == (BAD_PAYLOAD.len() - 1) as u64
    ));

    let total_under = ReadLimits::builder()
        .max_part_bytes(BAD_PAYLOAD.len() as u64)
        .expect("positive part limit")
        .max_total_part_bytes((GOOD_PAYLOAD.len() + BAD_PAYLOAD.len() - 1) as u64)
        .expect("positive total limit")
        .build()
        .expect("consistent total-under limit");
    assert!(matches!(
        OpcPackage::from_shared_vec_with_limits(source, total_under),
        Err(OpcError::ReadLimit {
            resource: ReadResource::TotalPartBytes,
            actual,
            maximum,
        }) if actual == (GOOD_PAYLOAD.len() + BAD_PAYLOAD.len()) as u64
            && maximum == (GOOD_PAYLOAD.len() + BAD_PAYLOAD.len() - 1) as u64
    ));
}
