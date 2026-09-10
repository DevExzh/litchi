#![allow(
    clippy::expect_used,
    clippy::panic,
    clippy::unwrap_used,
    reason = "the bounded deterministic smoke binary fails closed on harness regressions"
)]

use std::fs;
use std::panic::{self, AssertUnwindSafe};
use std::path::{Path, PathBuf};

use soapberry_zip::office::{ArchiveReader, StreamingArchiveWriter};

#[path = "opc_harness.rs"]
mod opc_harness;

const CONTENT_TYPES_NS: &str = "http://schemas.openxmlformats.org/package/2006/content-types";
const RELATIONSHIPS_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const OFFICE_DOCUMENT_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const DOCUMENT_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
const HYPERLINK_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
const EXTERNAL_TARGET: &str = "https://example.invalid/fuzz";
const DATA_REL: &str = "urn:litchi:fuzz:data";
const DATA_TARGET: &str = "../custom/data.bin";
const DOCUMENT_PAYLOAD: &[u8] = br#"<?xml version="1.0"?><document><body>fuzz</body></document>"#;
const DATA_PAYLOAD: &[u8] = &[0xa5; 37];

fn archive_bytes(content_types: &[u8], root_rels: &[u8], document_rels: &[u8]) -> Vec<u8> {
    let mut writer = StreamingArchiveWriter::new();
    writer
        .write_stored("[Content_Types].xml", content_types)
        .expect("deterministic content types member");
    writer
        .write_stored("_rels/.rels", root_rels)
        .expect("deterministic root relationships member");
    writer
        .write_stored("word/_rels/document.xml.rels", document_rels)
        .expect("deterministic document relationships member");
    writer
        .write_deflated_sized("word/document.xml", DOCUMENT_PAYLOAD)
        .expect("deterministic document member");
    writer
        .write_stored("custom/data.bin", DATA_PAYLOAD)
        .expect("deterministic binary member");
    writer.finish_to_bytes().expect("deterministic OPC archive")
}

struct ValidFixture {
    bytes: Vec<u8>,
    boundaries: opc_harness::BoundaryExpectations,
}

fn valid_fixture() -> ValidFixture {
    let content_types = format!(
        r#"<Types xmlns="{CONTENT_TYPES_NS}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="{DOCUMENT_CONTENT_TYPE}"/><Override PartName="/custom/data.bin" ContentType="application/octet-stream"/></Types>"#
    );
    let root_rels = format!(
        r#"<Relationships xmlns="{RELATIONSHIPS_NS}"><Relationship Id="rId1" Type="{OFFICE_DOCUMENT_REL}" Target="word/document.xml"/></Relationships>"#
    );
    let document_rels = format!(
        r#"<Relationships xmlns="{RELATIONSHIPS_NS}"><Relationship Id="rIdExternal" Type="{HYPERLINK_REL}" Target="{EXTERNAL_TARGET}" TargetMode="External"/><Relationship Id="rIdData" Type="{DATA_REL}" Target="{DATA_TARGET}"/></Relationships>"#
    );
    let bytes = archive_bytes(
        content_types.as_bytes(),
        root_rels.as_bytes(),
        document_rels.as_bytes(),
    );
    let input_bytes = bytes.len();
    let member_names = [
        "[Content_Types].xml",
        "_rels/.rels",
        "word/_rels/document.xml.rels",
        "word/document.xml",
        "custom/data.bin",
    ];
    let archive = ArchiveReader::new(&bytes).expect("deterministic archive index");
    let metadata_bytes = member_names
        .iter()
        .map(|name| name.len() + 46)
        .sum::<usize>();
    let metadata = member_names.iter().map(|name| {
        archive
            .metadata(name)
            .expect("deterministic member metadata")
    });
    let metadata = metadata.collect::<Vec<_>>();
    let archive_compressed_bytes = metadata
        .iter()
        .map(|metadata| {
            usize::try_from(metadata.compressed_size()).expect("bounded compressed size")
        })
        .max()
        .expect("deterministic members");
    let archive_entry_bytes = metadata
        .iter()
        .map(|metadata| usize::try_from(metadata.uncompressed_size()).expect("bounded entry size"))
        .max()
        .expect("deterministic members");
    let archive_total_bytes = metadata
        .iter()
        .map(|metadata| usize::try_from(metadata.uncompressed_size()).expect("bounded total size"))
        .sum();
    let archive_member_name_bytes = member_names
        .iter()
        .map(|name| name.len())
        .max()
        .expect("deterministic member names");
    let xml_attribute_bytes = [
        CONTENT_TYPES_NS.len(),
        RELATIONSHIPS_NS.len(),
        OFFICE_DOCUMENT_REL.len(),
        DOCUMENT_CONTENT_TYPE.len(),
        HYPERLINK_REL.len(),
        EXTERNAL_TARGET.len(),
        DATA_REL.len(),
        DATA_TARGET.len(),
        "rIdExternal".len(),
        "rIdData".len(),
    ]
    .into_iter()
    .max()
    .expect("deterministic XML attributes");
    ValidFixture {
        bytes,
        boundaries: opc_harness::BoundaryExpectations {
            input_bytes,
            archive_members: 5,
            parts: 2,
            relationship_parts: 2,
            content_types_bytes: content_types.len(),
            content_type_mappings: 4,
            relationship_xml_bytes: root_rels.len().max(document_rels.len()),
            total_relationship_xml_bytes: root_rels.len() + document_rels.len(),
            relationships_per_part: 2,
            total_relationships: 3,
            archive_member_name_bytes,
            archive_metadata_bytes: metadata_bytes,
            archive_compressed_bytes,
            archive_entry_bytes,
            archive_total_bytes,
            part_bytes: DOCUMENT_PAYLOAD.len().max(DATA_PAYLOAD.len()),
            total_part_bytes: DOCUMENT_PAYLOAD.len() + DATA_PAYLOAD.len(),
            relationship_graph_nodes: 2,
            xml_events: 7,
            total_relationship_xml_events: 9,
            xml_attribute_bytes,
            relationship_target_bytes: EXTERNAL_TARGET.len(),
        },
    }
}

fn malformed_seeds(valid: &[u8]) -> Vec<Vec<u8>> {
    let mut seeds = vec![
        Vec::new(),
        b"not an OPC package".to_vec(),
        b"PK\x03\x04\x14\x00\x00\x00".to_vec(),
        valid[..valid.len() / 2].to_vec(),
        valid[..valid.len() - 1].to_vec(),
    ];

    let root_rels = format!(
        r#"<Relationships xmlns="{RELATIONSHIPS_NS}"><Relationship Id="rId1" Target="word/document.xml"></Relationships>"#
    );
    seeds.push(archive_bytes(
        b"<Types>",
        b"<Relationships>",
        root_rels.as_bytes(),
    ));
    seeds
}

fn seed_files(valid: &[u8], malformed: &[Vec<u8>]) -> Vec<(String, Vec<u8>)> {
    let mut files = vec![("valid-opc.zip".to_owned(), valid.to_vec())];
    files.extend(
        malformed
            .iter()
            .enumerate()
            .map(|(index, bytes)| (format!("malformed-{index:02}.bin"), bytes.clone())),
    );
    files
}

fn write_seeds(path: &Path, valid: &[u8], malformed: &[Vec<u8>]) {
    fs::create_dir_all(path).expect("explicit seed output directory");
    for (name, bytes) in seed_files(valid, malformed) {
        fs::write(path.join(name), bytes).expect("explicit deterministic seed output");
    }
}

fn seed_output_argument() -> Option<PathBuf> {
    std::env::args()
        .skip(1)
        .find_map(|argument| argument.strip_prefix("--write-seeds=").map(PathBuf::from))
}

fn main() {
    let fixture = valid_fixture();
    let valid = &fixture.bytes;
    let malformed = malformed_seeds(valid);
    if let Some(path) = seed_output_argument() {
        write_seeds(&path, valid, &malformed);
    }

    let valid_stats = panic::catch_unwind(AssertUnwindSafe(|| opc_harness::exercise(valid)))
        .expect("valid deterministic seed must not panic");
    assert_eq!(
        valid_stats,
        opc_harness::ExerciseStats {
            source_backed_ok: true,
            source_xml_replacement_ok: true,
            source_xml_addition_ok: true,
            source_topology_write_ok: true,
            source_relationship_resolution_ok: true,
            source_main_document_ok: true,
            source_topology_reopen_ok: true,
            eager_ok: true,
            eager_relationship_resolution_ok: true,
            eager_main_document_ok: true,
            probe_ok: true,
        },
        "valid seed must reach all three public ingress paths"
    );

    for (index, seed) in malformed.iter().enumerate() {
        let stats = panic::catch_unwind(AssertUnwindSafe(|| opc_harness::exercise(seed)))
            .unwrap_or_else(|_| panic!("malformed deterministic seed {index} panicked"));
        assert_eq!(
            stats,
            opc_harness::ExerciseStats::default(),
            "malformed deterministic seed {index} must be rejected by all ingress paths"
        );
    }

    // Exact, one-over, and one-under values for the deterministic fixture's
    // relevant OPC resource ceilings are checked through every ingress path.
    opc_harness::exercise_boundaries(valid, fixture.boundaries);
    println!(
        "bounded OPC smoke passed: valid=1 malformed={} boundaries=22 bytes={}",
        malformed.len(),
        valid.len()
    );
}
