#![allow(
    clippy::expect_used,
    clippy::indexing_slicing,
    clippy::unwrap_used,
    reason = "the checked-in interoperability fixture is bounded and deterministic"
)]

//! Source-backed OPC preservation against a Python `force_zip64` corpus.
//!
//! The selected document part is the final physical member.  A targeted edit
//! therefore leaves the preceding package metadata, relationship member, and
//! opaque binary member at their original source offsets, which makes raw
//! local and central record preservation directly observable.

use std::collections::HashMap;

use litchi_opc::{OpcPackage, PackURI, PackageWriter, SourceBackedPackage};
use soapberry_zip::office::ArchiveLimits;
use soapberry_zip::{PreservationIndex, RECOMMENDED_BUFFER_SIZE, ZipArchive};

const SOURCE: &[u8] = include_bytes!(
    "../../../docs/performance/results/change-0416/corpus/opc-local-only-signed.zip"
);
const DOCUMENT_URI: &str = "/word/document.xml";
const DOCUMENT: &[u8] = b"<document>source descriptor fixture</document>";
const REPLACEMENT: &[u8] = b"<document>edited through source overlay</document>";

fn pack(uri: &str) -> PackURI {
    PackURI::new(uri).unwrap()
}

#[derive(Debug, Clone)]
struct RawRecord {
    local: Vec<u8>,
    central: Vec<u8>,
}

fn raw_records(data: &[u8]) -> HashMap<Vec<u8>, RawRecord> {
    let archive = ZipArchive::from_slice(data).unwrap().into_zip_archive();
    let mut scratch = vec![0; RECOMMENDED_BUFFER_SIZE];
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

#[test]
fn python_force_zip64_source_overlay_is_exact_on_noop_and_preserves_untouched_members() {
    let source_raw = raw_records(SOURCE);
    assert_eq!(source_raw.len(), 4);

    // The eager owning package and the source-backed overlay both retain an
    // exact no-op authorization, including ZIP64 local headers and signed
    // descriptors.
    let owning = OpcPackage::from_vec(SOURCE.to_vec()).expect("open Python OPC source");
    assert_eq!(PackageWriter::to_bytes(&owning).unwrap(), SOURCE);

    let mut no_op = Vec::new();
    SourceBackedPackage::from_vec(SOURCE.to_vec())
        .unwrap()
        .write_part_overlay_to_stream(&mut no_op, &pack(DOCUMENT_URI), DOCUMENT.to_vec())
        .expect("source-backed no-op overlay");
    assert_eq!(no_op, SOURCE);

    let package = SourceBackedPackage::from_vec(SOURCE.to_vec()).expect("open source-backed OPC");
    let physical_names: Vec<_> = package.physical_member_names().collect();
    assert_eq!(
        physical_names,
        vec![
            "[Content_Types].xml",
            "_rels/.rels",
            "custom/opaque.bin",
            "word/document.xml",
        ]
    );
    assert_eq!(package.iter_parts().count(), 2);
    assert_eq!(
        package
            .part(&pack(DOCUMENT_URI))
            .unwrap()
            .data()
            .unwrap()
            .as_bytes(),
        DOCUMENT
    );

    let mut output = Vec::new();
    SourceBackedPackage::from_vec(SOURCE.to_vec())
        .unwrap()
        .write_part_overlay_to_stream(&mut output, &pack(DOCUMENT_URI), REPLACEMENT.to_vec())
        .expect("targeted source-backed OPC edit");
    assert_ne!(output, SOURCE);

    let output_raw = raw_records(&output);
    for name in [
        b"[Content_Types].xml".as_slice(),
        b"_rels/.rels".as_slice(),
        b"custom/opaque.bin".as_slice(),
    ] {
        let source_record = source_raw.get(name).expect("source member");
        let output_record = output_raw.get(name).expect("preserved output member");
        assert_eq!(output_record.local, source_record.local);
        assert_eq!(output_record.central, source_record.central);
    }

    let reopened_source = SourceBackedPackage::from_vec(output.clone())
        .expect("reopen targeted source-backed output");
    assert_eq!(
        reopened_source
            .part(&pack(DOCUMENT_URI))
            .unwrap()
            .data()
            .unwrap()
            .as_bytes(),
        REPLACEMENT
    );
    let reopened_owning = OpcPackage::from_vec(output).expect("reopen targeted owning output");
    assert_eq!(
        reopened_owning
            .get_part(&pack(DOCUMENT_URI))
            .unwrap()
            .blob(),
        REPLACEMENT
    );
}
