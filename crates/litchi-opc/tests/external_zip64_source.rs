#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "this ignored interoperability probe uses fixture failures as assertions"
)]

//! Bounded interoperability coverage for an independently produced ZIP64 OPC
//! source.  The test is ignored because the corpus is supplied by the
//! external-producer harness rather than checked into this repository.

use std::collections::HashMap;

use litchi_opc::{PackURI, ReadLimits, SourceBackedPackage};
use soapberry_zip::office::ArchiveLimits;
use soapberry_zip::{PreservationIndex, ZipArchive};

const ZIP64_SOURCE_ENV: &str = "LITCHI_0415_PYTHON_ZIP";
const DOCUMENT_URI: &str = "/document.xml";
const LARGE_URI: &str = "/large.bin";
const LARGE_MEMBER: &[u8] = b"large.bin";
const LARGE_DECLARED_BYTES: u64 = 4_294_967_296;
const MAX_PHYSICAL_SOURCE_BYTES: usize = 64 * 1024 * 1024;
const REPLACEMENT: &[u8] = b"<document>litchi 0415 source edit</document>";

#[derive(Debug, Clone)]
struct RawRecord {
    local: Vec<u8>,
    central: Vec<u8>,
}

fn pack(uri: &str) -> PackURI {
    PackURI::new(uri).unwrap()
}

fn interoperability_limits() -> ReadLimits {
    // The ZIP64 member is admitted by its declared 4 GiB archive/part
    // ceilings, while the physical corpus remains bounded by the ordinary
    // input and compressed-byte policies.  The test never requests the large
    // decoded payload.
    ReadLimits::builder()
        .max_archive_entry_bytes(LARGE_DECLARED_BYTES)
        .unwrap()
        .max_archive_total_bytes(LARGE_DECLARED_BYTES + MAX_PHYSICAL_SOURCE_BYTES as u64)
        .unwrap()
        .max_part_bytes(LARGE_DECLARED_BYTES)
        .unwrap()
        .max_total_part_bytes(LARGE_DECLARED_BYTES + MAX_PHYSICAL_SOURCE_BYTES as u64)
        .unwrap()
        .build()
        .unwrap()
}

fn raw_records(data: &[u8]) -> HashMap<Vec<u8>, RawRecord> {
    let archive = ZipArchive::from_slice(data).unwrap().into_zip_archive();
    let mut scratch = vec![0; soapberry_zip::RECOMMENDED_BUFFER_SIZE];
    let index =
        PreservationIndex::new_with_limits(&archive, &mut scratch, ArchiveLimits::UNBOUNDED)
            .unwrap();

    index
        .entries()
        .iter()
        .map(|entry| {
            let local = entry.local_span();
            let central = entry.central_record();
            (
                entry.raw_name_bytes().to_vec(),
                RawRecord {
                    local: data[local.start as usize..local.end as usize].to_vec(),
                    central: data[central.start as usize..central.end as usize].to_vec(),
                },
            )
        })
        .collect()
}

fn central_without_local_offset(record: &[u8]) -> Vec<u8> {
    let mut record = record.to_vec();
    // The selected XML edit may move the following large member. The central
    // record's local-header offset is therefore the one regenerated field;
    // every other byte, including the ZIP64 extra data, must survive exactly.
    record[42..46].fill(0);
    record
}

#[test]
#[ignore = "requires an independently generated corpus at LITCHI_0415_PYTHON_ZIP"]
fn independent_zip64_source_metadata_and_small_xml_edit_preserve_large_member() {
    let source_path = std::env::var_os(ZIP64_SOURCE_ENV)
        .map(std::path::PathBuf::from)
        .expect("set LITCHI_0415_PYTHON_ZIP to the independent Python ZIP64 corpus");
    let source_length = std::fs::metadata(&source_path)
        .expect("stat independent ZIP64 corpus")
        .len();
    assert!(
        source_length <= MAX_PHYSICAL_SOURCE_BYTES as u64,
        "the probe must remain a bounded physical read; got {source_length} bytes"
    );
    let source_bytes = std::fs::read(&source_path).expect("read independent ZIP64 corpus");
    assert_eq!(source_bytes.len() as u64, source_length);

    let limits = interoperability_limits();
    let package = SourceBackedPackage::from_path_with_limits(&source_path, limits)
        .expect("independent ZIP64 OPC source should open with the declared-size profile");

    package
        .validate_topology_source_boundary()
        .expect("ZIP64 source topology should be preservable without payload reads");
    let physical_names: Vec<_> = package.physical_member_names().collect();
    assert_eq!(
        physical_names,
        vec![
            "[Content_Types].xml",
            "_rels/.rels",
            "document.xml",
            "large.bin"
        ]
    );

    let parts: Vec<_> = package.iter_parts().collect();
    assert_eq!(parts.len(), 2);
    assert_eq!(parts[0].partname().as_str(), DOCUMENT_URI);
    assert_eq!(parts[0].content_type(), "application/xml");
    assert_eq!(parts[0].rels().len(), 0);
    assert_eq!(parts[1].partname().as_str(), LARGE_URI);
    assert_eq!(parts[1].content_type(), "application/octet-stream");
    assert_eq!(
        parts[1].declared_uncompressed_size().unwrap(),
        LARGE_DECLARED_BYTES
    );
    assert_eq!(package.rels().len(), 1);
    let root_relationship = package.rels().get("rId1").unwrap();
    assert_eq!(root_relationship.reltype(), "urn:litchi:zip64-probe");
    assert_eq!(
        root_relationship.target_partname().unwrap().as_str(),
        DOCUMENT_URI
    );
    assert_eq!(package.cache_diagnostics().cold_loads, 0);
    assert_eq!(package.cache_diagnostics().successful_loads, 0);

    // A fresh source-backed package performs the small selected read needed by
    // the overlay. The 4 GiB member remains a raw preservation span throughout.
    let mut output = Vec::new();
    SourceBackedPackage::from_path_with_limits(&source_path, limits)
        .unwrap()
        .write_part_overlay_to_stream(&mut output, &pack(DOCUMENT_URI), REPLACEMENT.to_vec())
        .expect("small XML edit should publish while retaining ZIP64 source members");

    let output_package = SourceBackedPackage::from_vec_with_limits(output.clone(), limits)
        .expect("edited ZIP64 publication should reopen as an OPC package");
    output_package
        .validate_topology_source_boundary()
        .expect("edited ZIP64 publication should retain a valid source topology");
    assert_eq!(
        output_package
            .part(&pack(DOCUMENT_URI))
            .unwrap()
            .data()
            .unwrap()
            .as_bytes(),
        REPLACEMENT
    );
    assert_eq!(
        output_package
            .part(&pack(LARGE_URI))
            .unwrap()
            .declared_uncompressed_size()
            .unwrap(),
        LARGE_DECLARED_BYTES
    );
    assert_eq!(output_package.rels().len(), 1);
    assert_eq!(
        output_package
            .rels()
            .get("rId1")
            .unwrap()
            .target_partname()
            .unwrap()
            .as_str(),
        DOCUMENT_URI
    );

    let source_large = raw_records(&source_bytes)
        .remove(LARGE_MEMBER)
        .expect("source ZIP64 corpus must contain large.bin");
    let output_large = raw_records(&output)
        .remove(LARGE_MEMBER)
        .expect("edited publication must retain large.bin");
    assert_eq!(output_large.local, source_large.local);
    assert_eq!(
        central_without_local_offset(&output_large.central),
        central_without_local_offset(&source_large.central)
    );
}
