#![allow(
    clippy::expect_used,
    clippy::panic,
    clippy::unwrap_used,
    reason = "the bounded deterministic smoke binary fails closed on harness regressions"
)]

use std::fs;
use std::io::Cursor;
use std::panic::{self, AssertUnwindSafe};
use std::path::{Path, PathBuf};

use litchi_opc::{
    AuthoredXmlFragment, OpcError, OpcPackage, PackURI, ReadLimits, ReadResource,
    SourceBackedPackage, SourceTopologyPlan, probe_package_catalog_from_reader_with_limits,
};
use soapberry_zip::office::{ArchiveReader, StreamingArchiveWriter};
use soapberry_zip::{ZipArchive, ZipArchiveWriter};

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

fn archive_bytes_with_directory(
    content_types: &[u8],
    root_rels: &[u8],
    document_rels: &[u8],
) -> Vec<u8> {
    let mut output = Cursor::new(Vec::new());
    let mut writer = ZipArchiveWriter::new(&mut output);
    writer
        .new_dir("unused/")
        .create()
        .expect("deterministic directory member");
    writer
        .write_stored_file("[Content_Types].xml", content_types)
        .expect("deterministic content types member");
    writer
        .write_stored_file("_rels/.rels", root_rels)
        .expect("deterministic root relationships member");
    writer
        .write_stored_file("word/_rels/document.xml.rels", document_rels)
        .expect("deterministic document relationships member");
    writer
        .write_stored_file("word/document.xml", DOCUMENT_PAYLOAD)
        .expect("deterministic document member");
    writer
        .write_stored_file("custom/data.bin", DATA_PAYLOAD)
        .expect("deterministic binary member");
    writer.finish().expect("deterministic directory archive");
    output.into_inner()
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

fn assert_ingress_admits(data: &[u8], limits: ReadLimits, label: &str) {
    assert!(
        SourceBackedPackage::from_vec_with_limits(data.to_vec(), limits).is_ok(),
        "source-backed ingress must admit {label}"
    );
    assert!(
        OpcPackage::from_bytes_with_limits(data, limits).is_ok(),
        "eager ingress must admit {label}"
    );
    let mut reader = Cursor::new(data);
    assert!(
        probe_package_catalog_from_reader_with_limits(&mut reader, limits).is_ok(),
        "metadata probe ingress must admit {label}"
    );
}

fn assert_ingress_rejects(data: &[u8], limits: ReadLimits, resource: ReadResource, label: &str) {
    let source_error = match SourceBackedPackage::from_vec_with_limits(data.to_vec(), limits) {
        Ok(_) => panic!("source-backed ingress unexpectedly admitted boundary input for {label}"),
        Err(error) => error,
    };
    assert_read_limit(source_error, resource, label, "source-backed");

    let eager_error = match OpcPackage::from_bytes_with_limits(data, limits) {
        Ok(_) => panic!("eager ingress unexpectedly admitted boundary input"),
        Err(error) => error,
    };
    assert_read_limit(eager_error, resource, label, "eager");

    let mut reader = Cursor::new(data);
    let probe_error = match probe_package_catalog_from_reader_with_limits(&mut reader, limits) {
        Ok(_) => panic!("metadata probe unexpectedly admitted boundary input"),
        Err(error) => error,
    };
    assert_read_limit(probe_error, resource, label, "metadata probe");
}

fn assert_read_limit(error: OpcError, resource: ReadResource, label: &str, ingress: &str) {
    assert!(
        matches!(error, OpcError::ReadLimit { resource: actual, .. } if actual == resource),
        "{ingress} ingress returned {error:?}, expected {resource:?} for {label}"
    );
}

fn boundary_manifest() -> (Vec<u8>, Vec<u8>, Vec<u8>) {
    let content_types = format!(
        r#"<Types xmlns="{CONTENT_TYPES_NS}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="{DOCUMENT_CONTENT_TYPE}"/><Override PartName="/custom/data.bin" ContentType="application/octet-stream"/></Types>"#
    );
    let root_rels = format!(
        r#"<Relationships xmlns="{RELATIONSHIPS_NS}"><Relationship Id="rId1" Type="{OFFICE_DOCUMENT_REL}" Target="word/document.xml"/></Relationships>"#
    );
    let document_rels = format!(
        r#"<Relationships xmlns="{RELATIONSHIPS_NS}"><Relationship Id="rIdExternal" Type="{HYPERLINK_REL}" Target="{EXTERNAL_TARGET}" TargetMode="External"/><Relationship Id="rIdData" Type="{DATA_REL}" Target="{DATA_TARGET}"/></Relationships>"#
    );
    (
        content_types.into_bytes(),
        root_rels.into_bytes(),
        document_rels.into_bytes(),
    )
}

fn xml_depth_fixture() -> Vec<u8> {
    let (content_types, _, document_rels) = boundary_manifest();
    let root_rels = format!(
        r#"<Relationships xmlns="{RELATIONSHIPS_NS}"><Relationship Id="rId1" Type="{OFFICE_DOCUMENT_REL}" Target="word/document.xml"></Relationship></Relationships>"#
    );
    archive_bytes(&content_types, root_rels.as_bytes(), &document_rels)
}

fn archive_total_entries_fixture() -> Vec<u8> {
    let (content_types, root_rels, document_rels) = boundary_manifest();
    let bytes = archive_bytes_with_directory(&content_types, &root_rels, &document_rels);
    let archive = ZipArchive::from_slice(&bytes).expect("directory fixture must be a ZIP");
    assert_eq!(
        archive.entries_hint(),
        6,
        "five OPC members plus one directory must be counted physically"
    );
    let reader = ArchiveReader::new(&bytes).expect("directory fixture must be readable");
    assert_eq!(reader.len(), 5, "directory is not an OPC member");
    assert!(
        reader
            .metadata("unused/")
            .expect("directory metadata must be indexed")
            .is_directory(),
        "directory fixture must retain its central-directory classification"
    );
    bytes
}

fn limits_for_xml_depth(maximum: usize) -> ReadLimits {
    ReadLimits::builder()
        .max_xml_depth(maximum)
        .expect("positive XML-depth boundary")
        .build()
        .expect("XML-depth profile must remain consistent")
}

fn limits_for_archive_total_entries(maximum: usize) -> ReadLimits {
    ReadLimits::builder()
        .max_archive_members(5)
        .expect("five non-directory members is a valid ceiling")
        .max_parts(5)
        .expect("five admitted parts is a valid ceiling")
        .max_relationship_parts(5)
        .expect("five relationship parts is a valid ceiling")
        .max_archive_total_entries(maximum)
        .expect("positive total-entry boundary")
        .build()
        .expect("directory-aware archive profile must remain consistent")
}

fn assert_builder_constraints() {
    assert!(matches!(
        ReadLimits::builder().max_xml_depth(0),
        Err(OpcError::InvalidReadLimit {
            resource: ReadResource::XmlDepth,
            value: 0,
        })
    ));
    assert!(matches!(
        ReadLimits::builder().max_archive_total_entries(0),
        Err(OpcError::InvalidReadLimit {
            resource: ReadResource::ArchiveTotalEntries,
            value: 0,
        })
    ));
    assert!(matches!(
        ReadLimits::builder()
            .max_archive_members(6)
            .expect("positive archive-member ceiling")
            .max_parts(6)
            .expect("positive parts ceiling")
            .max_relationship_parts(6)
            .expect("positive relationship-parts ceiling")
            .max_archive_total_entries(5)
            .expect("positive total-entry ceiling")
            .build(),
        Err(OpcError::InvalidReadLimit {
            resource: ReadResource::ArchiveTotalEntries,
            value: 5,
        })
    ));
}

fn exercise_v4_limit_boundaries() {
    let depth = xml_depth_fixture();
    assert_ingress_admits(&depth, limits_for_xml_depth(2), "XML depth exact ceiling");
    assert_ingress_admits(
        &depth,
        limits_for_xml_depth(3),
        "XML depth one-over ceiling",
    );
    assert_ingress_rejects(
        &depth,
        limits_for_xml_depth(1),
        ReadResource::XmlDepth,
        "XML depth one-under ceiling",
    );

    let entries = archive_total_entries_fixture();
    assert_ingress_admits(
        &entries,
        limits_for_archive_total_entries(6),
        "archive total entries exact ceiling",
    );
    assert_ingress_admits(
        &entries,
        limits_for_archive_total_entries(7),
        "archive total entries one-over ceiling",
    );
    assert_ingress_rejects(
        &entries,
        limits_for_archive_total_entries(5),
        ReadResource::ArchiveTotalEntries,
        "archive total entries one-under ceiling",
    );
    assert_builder_constraints();
}

fn assert_reopened_xml_payloads() {
    let (content_types, root_rels, document_rels) = boundary_manifest();
    let input = archive_bytes(&content_types, &root_rels, &document_rels);
    let limits = opc_harness::tight_limits();
    let package = SourceBackedPackage::from_vec_with_limits(input, limits)
        .expect("source payload fixture must open");
    let document_uri = PackURI::new("/word/document.xml").expect("document URI is valid");
    let copy_uri = PackURI::new("/litchi-fuzz-source-copy.xml").expect("copy URI is valid");
    let source = package
        .part(&document_uri)
        .expect("source document must exist")
        .source_xml()
        .expect("source document must be XML-authorized");
    assert_eq!(source.bytes(), DOCUMENT_PAYLOAD);
    let insertion = source
        .bytes()
        .windows(2)
        .rposition(|window| window == b"</")
        .expect("document close tag must provide an insertion point");
    let proof = source
        .checked_range(insertion..insertion, &[])
        .expect("insertion point must be source-authorized");
    let mut expected = source.bytes().to_vec();
    expected.splice(insertion..insertion, b"<litchi-v4/>".iter().copied());
    let mut publication = source
        .into_publication()
        .expect("source document must enter an edit transaction");
    publication
        .replace(
            proof,
            AuthoredXmlFragment::markup(b"<litchi-v4/>".to_vec())
                .expect("authored XML fragment must pass its audit"),
        )
        .expect("source replacement must be accepted");
    let edited = publication
        .finish()
        .expect("source replacement must remain XML");
    assert_eq!(edited.bytes(), expected.as_slice());

    let mut plan = SourceTopologyPlan::new();
    plan.try_replace_source_xml_part(document_uri.clone(), edited.clone())
        .expect("source replacement must enter topology plan");
    plan.try_add_source_xml_part(copy_uri.clone(), edited)
        .expect("source XML copy must enter topology plan");
    let mut output = Vec::new();
    package
        .write_topology_to_stream(&mut output, plan)
        .expect("source replacement and copy must publish");

    let source_reopened = SourceBackedPackage::from_vec_with_limits(output.clone(), limits)
        .expect("source publication must reopen");
    assert_eq!(
        source_reopened
            .part(&document_uri)
            .expect("reopened source document must exist")
            .data()
            .expect("reopened source document must decode")
            .as_bytes(),
        expected.as_slice()
    );
    assert_eq!(
        source_reopened
            .part(&copy_uri)
            .expect("reopened source copy must exist")
            .data()
            .expect("reopened source copy must decode")
            .as_bytes(),
        expected.as_slice()
    );

    let eager_reopened =
        OpcPackage::from_bytes_with_limits(&output, limits).expect("eager publication must reopen");
    assert_eq!(
        eager_reopened
            .get_part(&document_uri)
            .expect("eager document must exist")
            .blob(),
        expected.as_slice()
    );
    assert_eq!(
        eager_reopened
            .get_part(&copy_uri)
            .expect("eager source copy must exist")
            .blob(),
        expected.as_slice()
    );

    let archive = ArchiveReader::new(&output).expect("published ZIP must reopen");
    assert_eq!(
        archive
            .read("word/document.xml")
            .expect("published document member must decode"),
        expected
    );
    assert_eq!(
        archive
            .read("litchi-fuzz-source-copy.xml")
            .expect("published copy member must decode"),
        expected
    );
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
    exercise_v4_limit_boundaries();
    assert_reopened_xml_payloads();
    println!(
        "bounded OPC smoke passed: valid=1 malformed={} boundaries=22 v4_boundaries=6 payload_reopen=2 bytes={}",
        malformed.len(),
        valid.len()
    );
}
