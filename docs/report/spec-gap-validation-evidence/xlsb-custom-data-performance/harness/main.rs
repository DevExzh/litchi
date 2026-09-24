#![allow(
    clippy::expect_used,
    clippy::indexing_slicing,
    clippy::unwrap_used,
    reason = "The bounded authored fixture and its public assertions fail closed."
)]

//! Candidate-only XLSB Custom Data lifecycle profile.
//!
//! Fixture construction is intentionally kept in this standalone binary. It
//! follows the valid source/graph recipe in the committed lifecycle test and
//! does not include the separate user-owned invalid fixture.

mod support;

use std::env;
use std::io::Cursor;
use std::time::Instant;

use litchi_opc::{BlobPart, OpcPackage, PackURI, Part, XmlPart};
use litchi_xlsb::custom_data::{
    Commit, CustomData, Patch, Snapshot, StorageSelector,
};
use litchi_xlsb::raw::{kind as rt, Kind, Writer};
use litchi_xlsb::writer::{MutableWorksheet, WorkbookWriter};
use litchi_xlsb::Package;
use serde_json::{json, Value};
use sha2::{Digest, Sha256};

#[global_allocator]
static ALLOCATOR: support::CountingAllocator = support::CountingAllocator;

const SOURCE_COMMIT: &str = "16102fe751d7c5492042330f1bd1f49c304495f0";
const PROFILE_SCHEMA: &str = "xlsb-custom-data-performance-v1";
const X14: &str = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
const WORKBOOK_URI: &str = "/xl/workbook.bin";
const CONNECTIONS_URI: &str = "/xl/connections.bin";
const CONNECTIONS_CONTENT_TYPE: &str = "application/vnd.ms-excel.connections";
const CONNECTIONS_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections";
const PROPERTIES_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customDataProps";
const DATA_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customData";
const PROPERTIES_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.customDataProperties+xml";
const DATA_CONTENT_TYPE: &str = "application/binary";
const OPAQUE_URI: &str = "/xl/opaque.bin";
const OPAQUE_CONTENT_TYPE: &str = "application/octet-stream";

#[derive(Clone, Copy, Debug)]
struct Level {
    name: &'static str,
    storages: usize,
    references: usize,
    payload_bytes: usize,
}

const LEVELS: [Level; 3] = [
    Level {
        name: "small",
        storages: 2,
        references: 16,
        payload_bytes: 1_024,
    },
    Level {
        name: "medium",
        storages: 4,
        references: 128,
        payload_bytes: 8_192,
    },
    Level {
        name: "large",
        storages: 16,
        references: 512,
        payload_bytes: 32_768,
    },
];

const LANES: [&str; 8] = [
    "read",
    "snapshotclone",
    "noop",
    "editpayload",
    "rename-many-to-one",
    "mixedinsertremove",
    "patchinverse",
    "publicapply",
];

#[derive(Debug)]
struct Fixture {
    package: Package,
    source_bytes: Vec<u8>,
    connections: Vec<u8>,
    opaque: Vec<u8>,
    ids: Vec<String>,
    level: Level,
    source_sha256: String,
    connections_sha256: String,
    opaque_sha256: String,
}

#[derive(Debug)]
enum Prepared {
    None,
    Commit(Commit),
    Patch {
        patch: Patch,
        inverse: Patch,
        forward_package_sha256: String,
    },
}

#[derive(Debug)]
enum Outcome {
    Snapshot(Snapshot),
    Commit(Commit),
    Package(Package, Option<Commit>),
}

#[derive(Debug)]
struct Validation {
    source_exact_noop: bool,
    preserved_unrelated_bytes: bool,
    inverse_correct: bool,
    payloads_exact: bool,
    semantics_correct: bool,
    copies_bytes_observed: u64,
    candidate_package_sha256: String,
}

fn uri(value: &str) -> PackURI {
    PackURI::new(value).expect("authored fixture URI is valid")
}

fn wide(value: &str) -> Vec<u8> {
    let units = value.encode_utf16().collect::<Vec<_>>();
    let mut output = Vec::with_capacity(4 + units.len() * 2);
    output.extend_from_slice(
        &u32::try_from(units.len())
            .expect("authored fixture string length fits u32")
            .to_le_bytes(),
    );
    for unit in units {
        output.extend_from_slice(&unit.to_le_bytes());
    }
    output
}

fn record(kind: Kind, payload: &[u8], output: &mut Vec<u8>) {
    Writer::new(output)
        .write_record(kind, payload)
        .expect("authored BIFF12 record writes");
}

fn ext_conn_payload(connection_id: u32, name: &str) -> Vec<u8> {
    // BrtBeginExtConnection. The source type at offset 10 is DBTOLEDB (5),
    // which is required for an admitted ExtConn14 collection.
    let mut output = vec![7, 5, 2, 0];
    output.extend_from_slice(&30_u16.to_le_bytes());
    output.extend_from_slice(&0_u16.to_le_bytes());
    output.extend_from_slice(&(1_u16 << 3).to_le_bytes());
    output.extend_from_slice(&5_u32.to_le_bytes());
    output.extend_from_slice(&3_u32.to_le_bytes());
    output.extend_from_slice(&connection_id.to_le_bytes());
    output.push(0);
    output.extend_from_slice(&wide(name));
    output
}

fn ext_conn14_payload(culture: &str, uid: &str) -> Vec<u8> {
    let mut output = vec![0; 4];
    output.extend_from_slice(&wide(culture));
    output.extend_from_slice(&wide(uid));
    output
}

fn connections_stream(references: &[String]) -> Vec<u8> {
    let mut output = Vec::new();
    record(rt::BEGIN_EXT_CONNECTIONS, &[], &mut output);
    record(
        rt::BEGIN_EXT_CONNECTION,
        &ext_conn_payload(42, "Synthetic Custom Data"),
        &mut output,
    );
    for (ordinal, uid) in references.iter().enumerate() {
        // FRTBegin, BeginExtConn14, EndExtConn14, FRTEnd is the pinned
        // complete wrapper grammar used by the committed lifecycle test.
        record(rt::FRT_BEGIN, &[1, 0, 1, 0], &mut output);
        record(
            rt::BEGIN_EXT_CONN14,
            &ext_conn14_payload(if ordinal == 0 { "en-US" } else { "fr-FR" }, uid),
            &mut output,
        );
        record(rt::END_EXT_CONN14, &[], &mut output);
        record(rt::FRT_END, &[], &mut output);
    }
    record(rt::END_EXT_CONNECTION, &[], &mut output);
    record(rt::END_EXT_CONNECTIONS, &[], &mut output);
    output
}

fn properties_xml(uid: &str) -> Vec<u8> {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><!-- retained before root --><alias:datastoreItem xmlns:alias="{X14}" xmlns:s="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:v="urn:synthetic-vendor" id="{uid}"><?root-pi?><alias:extLst><?extension-pi?><s:ext uri="urn:synthetic"><v:opaque v:raw="&amp;raw"><v:leaf><![CDATA[<? retained-looking text ]]></v:leaf><!-- opaque comment --></v:opaque></s:ext></alias:extLst><!-- retained after extension --></alias:datastoreItem>"#
    )
    .into_bytes()
}

fn payload_for(level: Level, storage_index: usize) -> Vec<u8> {
    (0..level.payload_bytes)
        .map(|offset| {
            let value = offset
                .wrapping_add(storage_index.wrapping_mul(17))
                .wrapping_add(level.storages);
            value as u8
        })
        .collect()
}

fn add_relationship(
    package: &mut OpcPackage,
    source: &str,
    reltype: &str,
    target: &str,
    id: &str,
) {
    package
        .get_part_mut(&uri(source))
        .expect("authored relationship source exists")
        .rels_mut()
        .add_relationship(
            reltype.to_owned(),
            target.to_owned(),
            id.to_owned(),
            false,
        );
}

fn make_fixture(level: Level) -> Fixture {
    let ids = (0..level.storages)
        .map(|index| format!("uid-{index:04}"))
        .collect::<Vec<_>>();
    // The first UID receives the many-to-one fan-in. Every other UID except
    // the final one receives one edge, leaving the final storage unreferenced
    // for the mixed insert/remove lane.
    let edge_count_for_other_ids = level.storages.saturating_sub(2);
    let first_reference_count = level.references - edge_count_for_other_ids;
    let mut references = vec![ids[0].clone(); first_reference_count];
    references.extend(ids.iter().skip(1).take(edge_count_for_other_ids).cloned());
    assert_eq!(references.len(), level.references);

    let connections = connections_stream(&references);
    let opaque = b"opaque member source bytes retained by the candidate profile".to_vec();

    let mut writer = WorkbookWriter::new();
    writer.add_worksheet(MutableWorksheet::new("Sheet1"));
    let mut output = Cursor::new(Vec::new());
    writer.save(&mut output).expect("base XLSB package writes");
    let base = Package::from_bytes(output.into_inner()).expect("base XLSB package reopens");
    let mut package = base.into_opc();

    for (index, uid) in ids.iter().enumerate() {
        let properties_uri = format!("/xl/customData/props{}.xml", index + 1);
        let data_uri = format!("/xl/customData/data{}.bin", index + 1);
        let mut properties_part = XmlPart::new(
            uri(&properties_uri),
            PROPERTIES_CONTENT_TYPE.to_owned(),
            properties_xml(uid),
        );
        properties_part.rels_mut().add_relationship(
            DATA_RELATIONSHIP_TYPE.to_owned(),
            format!("data{}.bin", index + 1),
            format!("rIdCustomData{}", index + 1),
            false,
        );
        package.add_part(Box::new(properties_part));
        package.add_part(Box::new(BlobPart::new(
            uri(&data_uri),
            DATA_CONTENT_TYPE.to_owned(),
            payload_for(level, index),
        )));
        add_relationship(
            &mut package,
            WORKBOOK_URI,
            PROPERTIES_RELATIONSHIP_TYPE,
            &format!("customData/props{}.xml", index + 1),
            &format!("rIdCustomDataProps{}", index + 1),
        );
    }
    package.add_part(Box::new(BlobPart::new(
        uri(CONNECTIONS_URI),
        CONNECTIONS_CONTENT_TYPE.to_owned(),
        connections.clone(),
    )));
    package.add_part(Box::new(BlobPart::new(
        uri(OPAQUE_URI),
        OPAQUE_CONTENT_TYPE.to_owned(),
        opaque.clone(),
    )));
    add_relationship(
        &mut package,
        WORKBOOK_URI,
        CONNECTIONS_RELATIONSHIP_TYPE,
        "connections.bin",
        "rIdConnections",
    );

    let mut package_bytes = Vec::new();
    package
        .to_stream(&mut package_bytes)
        .expect("authored source package serializes");
    let package = Package::from_bytes(package_bytes)
        .expect("authored XLSB Custom Data graph reopens from physical bytes");
    let source_bytes = package.to_bytes().expect("authored source bytes serialize");
    Fixture {
        package,
        source_sha256: sha256(&source_bytes),
        connections_sha256: sha256(&connections),
        opaque_sha256: sha256(&opaque),
        source_bytes,
        connections,
        opaque,
        ids,
        level,
    }
}

fn sha256(bytes: &[u8]) -> String {
    let mut digest = Sha256::new();
    digest.update(bytes);
    digest
        .finalize()
        .iter()
        .map(|byte| format!("{byte:02x}"))
        .collect()
}

fn fnv1a64(bytes: &[u8]) -> String {
    let mut hash = 0xcbf29ce484222325_u64;
    for byte in bytes {
        hash ^= u64::from(*byte);
        hash = hash.wrapping_mul(0x100000001b3);
    }
    format!("{hash:016x}")
}

fn package_bytes(package: &Package) -> Vec<u8> {
    package.to_bytes().expect("authored package serializes")
}

fn part_bytes(package: &Package, name: &str) -> Vec<u8> {
    package
        .opc_package()
        .get_part(&uri(name))
        .expect("authored fixture member exists")
        .blob()
        .to_vec()
}

fn first_reference_count(level: Level) -> usize {
    level.references - level.storages.saturating_sub(2)
}

fn expected_reference_count(fixture: &Fixture, index: usize) -> usize {
    if index == 0 {
        first_reference_count(fixture.level)
    } else if index + 1 < fixture.ids.len() {
        1
    } else {
        0
    }
}

fn assert_expected_reference_map(
    snapshot: &Snapshot,
    fixture: &Fixture,
    lane: &str,
    expected_first_id: &str,
) {
    let mut expected = Vec::with_capacity(fixture.level.storages);
    for (index, id) in fixture.ids.iter().enumerate() {
        if lane == "mixedinsertremove" && index + 1 == fixture.ids.len() {
            continue;
        }
        let output_id = if index == 0
            && matches!(lane, "rename-many-to-one" | "patchinverse")
        {
            expected_first_id.to_owned()
        } else {
            id.clone()
        };
        expected.push((output_id, expected_reference_count(fixture, index)));
    }
    if lane == "mixedinsertremove" {
        expected.push((
            format!("uid-inserted-{}", fixture.level.name),
            0,
        ));
    }
    assert_eq!(snapshot.len(), expected.len(), "storage count changed");
    for (id, references) in expected {
        assert!(snapshot.find(&id).is_some(), "expected storage is missing: {id}");
        assert_eq!(
            snapshot.connection_references(&id),
            references,
            "reference count changed for {id}"
        );
    }
}

fn assert_snapshot(snapshot: &Snapshot, fixture: &Fixture, expected_first_id: &str) {
    assert_eq!(snapshot.len(), fixture.level.storages);
    for (index, id) in fixture.ids.iter().enumerate() {
        let storage = snapshot.find(id).expect("missing authored storage");
        let expected = payload_for(fixture.level, index);
        assert_eq!(storage.data().len(), expected.len(), "storage payload length changed for {id}");
        assert_eq!(storage.data(), expected.as_slice(), "storage payload changed for {id}");
    }
    assert_expected_reference_map(snapshot, fixture, "unchanged", expected_first_id);
    assert!(!snapshot.has_opaque_reference_candidates());
}

fn assert_expected_payloads(
    snapshot: &Snapshot,
    fixture: &Fixture,
    lane: &str,
    expected_first_id: &str,
) {
    for (index, id) in fixture.ids.iter().enumerate() {
        if lane == "mixedinsertremove" && index + 1 == fixture.ids.len() {
            continue;
        }
        let output_id = if index == 0
            && matches!(lane, "rename-many-to-one" | "patchinverse")
        {
            expected_first_id
        } else {
            id
        };
        let expected = match (lane, index) {
            ("editpayload", 0) => vec![0xA5; fixture.level.payload_bytes / 2 + 1],
            ("publicapply", 0) => vec![0x3C; fixture.level.payload_bytes + 7],
            _ => payload_for(fixture.level, index),
        };
        let storage = snapshot.find(output_id).expect("expected storage is missing");
        assert_eq!(storage.data().len(), expected.len(), "storage payload length changed");
        assert_eq!(storage.data(), expected.as_slice(), "storage payload bytes changed");
    }
    if lane == "mixedinsertremove" {
        let inserted = format!("uid-inserted-{}", fixture.level.name);
        let storage = snapshot.find(&inserted).expect("inserted storage is missing");
        let expected = vec![0x5A; fixture.level.payload_bytes / 3 + 1];
        assert_eq!(storage.data().len(), expected.len(), "inserted payload length changed");
        assert_eq!(storage.data(), expected.as_slice(), "inserted payload bytes changed");
    }
}

fn prepare(fixture: &Fixture, lane: &str) -> Prepared {
    match lane {
        "read" | "snapshotclone" | "noop" => Prepared::None,
        "editpayload" => {
            let mut transaction = fixture
                .package
                .edit_custom_data()
                .expect("payload edit transaction prepares");
            let replacement = vec![0xA5; fixture.level.payload_bytes / 2 + 1];
            transaction
                .set_data(StorageSelector::Id(&fixture.ids[0]), replacement)
                .expect("payload edit stages");
            Prepared::Commit(transaction.commit().expect("payload edit commits"))
        },
        "rename-many-to-one" => {
            let mut transaction = fixture
                .package
                .edit_custom_data()
                .expect("rename transaction prepares");
            transaction
                .rename(
                    StorageSelector::Id(&fixture.ids[0]),
                    format!("renamed-{}", fixture.level.name),
                )
                .expect("rename stages");
            Prepared::Commit(transaction.commit().expect("rename commits"))
        },
        "mixedinsertremove" => {
            let mut transaction = fixture
                .package
                .edit_custom_data()
                .expect("mixed transaction prepares");
            transaction
                .insert(CustomData::new(
                    format!("uid-inserted-{}", fixture.level.name),
                    vec![0x5A; fixture.level.payload_bytes / 3 + 1],
                ))
                .expect("mixed insertion stages");
            transaction
                .remove(StorageSelector::Id(
                    fixture.ids.last().expect("fixture has a removable storage"),
                ))
                .expect("mixed removal stages")
                .expect("unreferenced storage is present");
            Prepared::Commit(transaction.commit().expect("mixed transaction commits"))
        },
        "patchinverse" => {
            let mut transaction = fixture
                .package
                .edit_custom_data()
                .expect("patch transaction prepares");
            transaction
                .rename(
                    StorageSelector::Id(&fixture.ids[0]),
                    format!("inverse-{}", fixture.level.name),
                )
                .expect("patch rename stages");
            let commit = transaction.commit().expect("patch transaction commits");
            let patch = commit.patch().clone();
            let forward = fixture
                .package
                .apply_custom_data_patch(&patch)
                .expect("prepared forward patch applies");
            let expected_first_id = format!("inverse-{}", fixture.level.name);
            let (_, preserved, payloads, _) =
                validate_candidate(&forward, fixture, "rename-many-to-one", &expected_first_id);
            assert!(preserved, "prepared forward patch changed unrelated bytes");
            assert!(payloads, "prepared forward patch changed expected payloads");
            let forward_package_sha256 = sha256(&package_bytes(&forward));
            Prepared::Patch {
                inverse: patch.inverse(),
                patch,
                forward_package_sha256,
            }
        },
        "publicapply" => {
            let mut transaction = fixture
                .package
                .edit_custom_data()
                .expect("public apply transaction prepares");
            transaction
                .set_data(
                    StorageSelector::Id(&fixture.ids[0]),
                    vec![0x3C; fixture.level.payload_bytes + 7],
                )
                .expect("public apply payload stages");
            Prepared::Commit(transaction.commit().expect("public apply commits"))
        },
        other => panic!("unknown lane {other}"),
    }
}

fn cost_scope(lane: &str) -> &'static str {
    match lane {
        "read" => "Package::custom_data only; fixture and source setup are outside timing",
        "snapshotclone" => "Clone of a prepared immutable Snapshot only",
        "noop" => "Package::edit_custom_data, empty commit, and public no-op apply",
        "editpayload" => {
            "Payload staging plus Transaction::commit; caller replacement buffer is built inside the timed call; public apply is validation"
        },
        "rename-many-to-one" => {
            "UID rename, all admitted reference rewrites, and Transaction::commit; caller rename string is built inside the timed call"
        },
        "mixedinsertremove" => {
            "One Custom Data insertion, one unreferenced removal, and Transaction::commit; caller ID and payload are built inside the timed call"
        },
        "patchinverse" => "Prepared public patch apply followed by prepared inverse apply",
        "publicapply" => "Package::apply_custom_data with a prepared public commit",
        _ => "unknown",
    }
}

fn input_preparation(lane: &str) -> &'static str {
    match lane {
        "read" => "prepared fixture package outside timing",
        "snapshotclone" => "prepared immutable snapshot outside timing; clone input is equal for every sample",
        "noop" => "empty transaction is constructed inside the timed call",
        "editpayload" => "replacement payload and transaction are constructed inside the timed call",
        "rename-many-to-one" => "rename string and transaction are constructed inside the timed call",
        "mixedinsertremove" => "inserted ID/payload and transaction are constructed inside the timed call",
        "patchinverse" => "forward patch and inverse are prepared outside timing and reused for every sample",
        "publicapply" => "public commit is prepared outside timing and reused for every sample",
        _ => "unknown",
    }
}

fn validate_candidate(
    candidate: &Package,
    fixture: &Fixture,
    lane: &str,
    expected_first_id: &str,
) -> (bool, bool, bool, u64) {
    let snapshot = candidate.custom_data().expect("candidate reads after publication");
    assert_eq!(snapshot.len(), fixture.level.storages);
    match lane {
        "rename-many-to-one" | "patchinverse" => {
            assert!(snapshot.find(&fixture.ids[0]).is_none());
            for id in fixture.ids.iter().skip(1) {
                assert!(snapshot.find(id).is_some(), "missing unchanged storage {id}");
            }
            assert!(snapshot.find(expected_first_id).is_some());
        },
        "mixedinsertremove" => {
            for id in fixture.ids.iter().take(fixture.ids.len().saturating_sub(1)) {
                assert!(snapshot.find(id).is_some(), "missing unchanged storage {id}");
            }
            assert!(snapshot.find(fixture.ids.last().unwrap()).is_none());
            let inserted = format!("uid-inserted-{}", fixture.level.name);
            assert!(snapshot.find(&inserted).is_some());
        },
        "editpayload" | "publicapply" => {
            for id in &fixture.ids {
                assert!(snapshot.find(id).is_some(), "missing storage {id}");
            }
        },
        _ => assert_snapshot(&snapshot, fixture, expected_first_id),
    }
    assert_expected_reference_map(&snapshot, fixture, lane, expected_first_id);
    assert_expected_payloads(&snapshot, fixture, lane, expected_first_id);
    let mut copies = 0_u64;
    let opaque = part_bytes(candidate, OPAQUE_URI);
    copies = copies.saturating_add(opaque.len() as u64);
    let preserved = opaque == fixture.opaque;
    assert!(preserved, "unrelated opaque member changed");

    match lane {
        "editpayload" | "publicapply" => {
            let data = snapshot
                .find(expected_first_id)
                .expect("changed first storage exists")
                .data();
            assert_ne!(data, fixture.package.custom_data().unwrap().find(&fixture.ids[0]).unwrap().data());
        },
        "rename-many-to-one" | "patchinverse" => {
            assert_eq!(snapshot.connection_references(expected_first_id), first_reference_count(fixture.level));
            assert!(snapshot.find(&fixture.ids[0]).is_none());
        },
        "mixedinsertremove" => {
            let inserted = format!("uid-inserted-{}", fixture.level.name);
            assert!(snapshot.find(&inserted).is_some());
            assert!(snapshot.find(fixture.ids.last().unwrap()).is_none());
        },
        _ => {}
    }
    (true, preserved, true, copies)
}

fn validate_outcome(outcome: &Outcome, fixture: &Fixture, lane: &str) -> Validation {
    let mut source_exact_noop = false;
    let mut inverse_correct = false;
    let payloads_exact;
    let semantics_correct;
    let mut preserved_unrelated_bytes = false;
    let mut copies = 0_u64;
    let candidate_package_sha256;
    match outcome {
        Outcome::Snapshot(snapshot) => {
            assert_snapshot(snapshot, fixture, &fixture.ids[0]);
            payloads_exact = true;
            semantics_correct = true;
            candidate_package_sha256 = fixture.source_sha256.clone();
        },
        Outcome::Commit(commit) => {
            assert!(commit.changed());
            let expected_first_id = match lane {
                "rename-many-to-one" | "patchinverse" => {
                    format!("{}-{}", if lane == "patchinverse" { "inverse" } else { "renamed" }, fixture.level.name)
                },
                _ => fixture.ids[0].clone(),
            };
            let candidate = fixture
                .package
                .apply_custom_data(commit)
                .expect("prepared commit applies for validation");
            let candidate_bytes = package_bytes(&candidate);
            copies = copies.saturating_add(candidate_bytes.len() as u64);
            candidate_package_sha256 = sha256(&candidate_bytes);
            let (semantic, preserved, payloads, observed) =
                validate_candidate(&candidate, fixture, lane, &expected_first_id);
            semantics_correct = semantic;
            preserved_unrelated_bytes = preserved;
            payloads_exact = payloads;
            copies = copies.saturating_add(observed);
            let restored = candidate
                .apply_custom_data_patch(&commit.patch().inverse())
                .expect("public commit inverse applies for validation");
            let restored_bytes = package_bytes(&restored);
            copies = copies.saturating_add(restored_bytes.len() as u64);
            inverse_correct = restored_bytes == fixture.source_bytes;
            assert!(inverse_correct, "commit inverse did not restore source bytes");
        },
        Outcome::Package(package, commit) => {
            let expected_first_id = &fixture.ids[0];
            let candidate_bytes = package_bytes(package);
            copies = copies.saturating_add(candidate_bytes.len() as u64);
            candidate_package_sha256 = sha256(&candidate_bytes);
            let snapshot = package.custom_data().expect("published package reads");
            if lane == "noop" {
                source_exact_noop = candidate_bytes == fixture.source_bytes;
                assert!(source_exact_noop, "public no-op changed source bytes");
                assert_snapshot(&snapshot, fixture, expected_first_id);
                payloads_exact = true;
                semantics_correct = true;
                let opaque = part_bytes(package, OPAQUE_URI);
                copies = copies.saturating_add(opaque.len() as u64);
                preserved_unrelated_bytes = opaque == fixture.opaque;
                assert!(preserved_unrelated_bytes);
            } else if lane == "patchinverse" {
                inverse_correct = candidate_bytes == fixture.source_bytes;
                assert!(inverse_correct, "patch inverse did not restore source bytes");
                assert_snapshot(&snapshot, fixture, &fixture.ids[0]);
                payloads_exact = true;
                semantics_correct = true;
                let opaque = part_bytes(package, OPAQUE_URI);
                copies = copies.saturating_add(opaque.len() as u64);
                preserved_unrelated_bytes = opaque == fixture.opaque;
                assert!(preserved_unrelated_bytes);
            } else if lane == "publicapply" {
                let (semantic, preserved, payloads, observed) =
                    validate_candidate(package, fixture, lane, &fixture.ids[0]);
                semantics_correct = semantic;
                preserved_unrelated_bytes = preserved;
                payloads_exact = payloads;
                copies = copies.saturating_add(observed);
                let commit = commit.as_ref().expect("public apply retains its commit");
                let restored = package
                    .apply_custom_data_patch(&commit.patch().inverse())
                    .expect("public apply inverse applies for validation");
                let restored_bytes = package_bytes(&restored);
                copies = copies.saturating_add(restored_bytes.len() as u64);
                inverse_correct = restored_bytes == fixture.source_bytes;
                assert!(inverse_correct, "public apply inverse did not restore source bytes");
            } else {
                unreachable!("only public package lanes reach this branch");
            }
        },
    }
    Validation {
        source_exact_noop,
        preserved_unrelated_bytes,
        inverse_correct,
        payloads_exact,
        semantics_correct,
        copies_bytes_observed: copies,
        candidate_package_sha256,
    }
}

fn run_operation(
    fixture: &Fixture,
    prepared: &Prepared,
    baseline_snapshot: Option<&Snapshot>,
    lane: &str,
) -> (Outcome, u64, support::Snapshot, support::Snapshot, support::Snapshot, support::Delta) {
    let working = (lane != "snapshotclone").then(|| fixture.package.clone());
    support::reset();
    let before = support::Snapshot::now();
    let started = Instant::now();
    let outcome = match lane {
        "read" => Outcome::Snapshot(
            working
                .as_ref()
                .expect("read working package")
                .custom_data()
                .expect("public read succeeds"),
        ),
        "snapshotclone" => Outcome::Snapshot(
            baseline_snapshot
                .expect("snapshot clone baseline")
                .clone(),
        ),
        "noop" => {
            let transaction = working
                .as_ref()
                .expect("noop working package")
                .edit_custom_data()
                .expect("noop transaction starts");
            let commit = transaction.commit().expect("noop commit succeeds");
            Outcome::Package(
                working
                    .as_ref()
                    .expect("noop working package")
                    .apply_custom_data(&commit)
                    .expect("public noop apply succeeds"),
                None,
            )
        },
        "editpayload" | "rename-many-to-one" | "mixedinsertremove" => {
            let mut transaction = working
                .as_ref()
                .expect("edit working package")
                .edit_custom_data()
                .expect("edit transaction starts");
            match lane {
                "editpayload" => {
                    transaction
                        .set_data(
                            StorageSelector::Id(&fixture.ids[0]),
                            vec![0xA5; fixture.level.payload_bytes / 2 + 1],
                        )
                        .expect("payload edit stages");
                },
                "rename-many-to-one" => {
                    transaction
                        .rename(
                            StorageSelector::Id(&fixture.ids[0]),
                            format!("renamed-{}", fixture.level.name),
                        )
                        .expect("rename stages");
                },
                "mixedinsertremove" => {
                    transaction
                        .insert(CustomData::new(
                            format!("uid-inserted-{}", fixture.level.name),
                            vec![0x5A; fixture.level.payload_bytes / 3 + 1],
                        ))
                        .expect("mixed insertion stages");
                    transaction
                        .remove(StorageSelector::Id(fixture.ids.last().unwrap()))
                        .expect("mixed removal stages")
                        .expect("mixed removable storage exists");
                },
                _ => unreachable!(),
            }
            Outcome::Commit(transaction.commit().expect("edit commit succeeds"))
        },
        "patchinverse" => {
            let Prepared::Patch { patch, inverse, .. } = prepared else {
                panic!("patchinverse requires prepared patches");
            };
            let forward = working
                .as_ref()
                .expect("patch working package")
                .apply_custom_data_patch(patch)
                .expect("forward patch applies");
            Outcome::Package(
                forward
                    .apply_custom_data_patch(inverse)
                    .expect("inverse patch applies"),
                None,
            )
        },
        "publicapply" => {
            let Prepared::Commit(commit) = prepared else {
                panic!("publicapply requires prepared commit");
            };
            Outcome::Package(
                working
                    .as_ref()
                    .expect("public apply working package")
                    .apply_custom_data(commit)
                    .expect("public commit applies"),
                Some(commit.clone()),
            )
        },
        other => panic!("unknown lane {other}"),
    };
    let elapsed = started.elapsed();
    let after_operation = support::Snapshot::now();
    let delta = before.delta(after_operation);
    (outcome, elapsed.as_nanos().try_into().unwrap_or(u64::MAX), before, after_operation, support::Snapshot::now(), delta)
}

fn sample(level: Level, lane: &str, fixture: &Fixture, prepared: &Prepared, baseline_snapshot: Option<&Snapshot>, sample_index: usize) -> Value {
    let (outcome, elapsed_ns, before, after_operation, _, operation_delta) =
        run_operation(fixture, prepared, baseline_snapshot, lane);
    let validation = validate_outcome(&outcome, fixture, lane);
    let prepared_forward_package_sha256 = match prepared {
        Prepared::Patch {
            forward_package_sha256,
            ..
        } => Some(forward_package_sha256.as_str()),
        _ => None,
    };
    drop(outcome);
    let after_drop = support::Snapshot::now();
    let drop_delta = before.delta(after_drop);
    json!({
        "schema": PROFILE_SCHEMA,
        "source_commit": SOURCE_COMMIT,
        "lane": lane,
        "size_class": level.name,
        "sample": sample_index,
        "fixture": {
            "storages": level.storages,
            "references": level.references,
            "payload_bytes_per_storage": level.payload_bytes,
            "source_bytes": fixture.source_bytes.len(),
            "source_sha256": fixture.source_sha256,
            "source_fnv1a64": fnv1a64(&fixture.source_bytes),
            "connections_bytes": fixture.connections.len(),
            "connections_sha256": fixture.connections_sha256,
            "opaque_bytes": fixture.opaque.len(),
            "opaque_sha256": fixture.opaque_sha256
        },
        "cost_scope": cost_scope(lane),
        "input_preparation": input_preparation(lane),
        "timing_ns": elapsed_ns,
        "allocator": {
            "allocation_calls": operation_delta.allocation_calls,
            "reallocation_calls": operation_delta.reallocation_calls,
            "deallocation_calls": operation_delta.deallocation_calls,
            "direct_allocated_bytes": operation_delta.direct_bytes,
            "realloc_old_bytes": operation_delta.realloc_old_bytes,
            "realloc_new_bytes": operation_delta.realloc_new_bytes,
            "requested_alloc_bytes": operation_delta.requested_bytes(),
            "released_alloc_bytes_operation": operation_delta.deallocated_bytes,
            "released_alloc_bytes_through_drop": drop_delta.deallocated_bytes,
            "live_before": before.live_bytes,
            "live_after_operation": after_operation.live_bytes,
            "live_after_drop": after_drop.live_bytes,
            "logical_peak_bytes": operation_delta.peak_live_absolute,
            "peak_live_delta": operation_delta.peak_live_delta,
            "allocation_failed": operation_delta.allocation_failed,
            "alloc_balance_ok": operation_delta.balanced(),
            "alloc_invalid": operation_delta.invalid
        },
        "copies_bytes_observed": validation.copies_bytes_observed,
        "candidate_package_sha256": validation.candidate_package_sha256,
        "prepared_forward_package_sha256": prepared_forward_package_sha256,
        "copy_scope": "explicit harness validation copies after timing; internal production copies are uninstrumented",
        "assertions": {
            "source_exact_noop": validation.source_exact_noop,
            "preserved_unrelated_bytes": validation.preserved_unrelated_bytes,
            "inverse_correct": validation.inverse_correct,
            "payloads_exact": validation.payloads_exact,
            "semantics_correct": validation.semantics_correct
        }
    })
}

fn level_by_name(name: &str) -> Level {
    LEVELS
        .iter()
        .copied()
        .find(|level| level.name == name)
        .unwrap_or_else(|| panic!("unknown size class {name}"))
}

fn run(level: Level, lane: &str, samples: usize, warmups: usize) {
    assert!(LANES.contains(&lane), "unknown lane {lane}");
    let fixture = make_fixture(level);
    let prepared = prepare(&fixture, lane);
    let baseline_snapshot = (lane == "snapshotclone")
        .then(|| fixture.package.custom_data().expect("snapshot clone baseline read"));
    for _ in 0..warmups {
        let _ = sample(level, lane, &fixture, &prepared, baseline_snapshot.as_ref(), usize::MAX);
    }
    for sample_index in 0..samples {
        let row = sample(
            level,
            lane,
            &fixture,
            &prepared,
            baseline_snapshot.as_ref(),
            sample_index,
        );
        println!("{}", serde_json::to_string(&row).expect("receipt serializes"));
    }
}

fn main() {
    let args = env::args().skip(1).collect::<Vec<_>>();
    let mut lane = None;
    let mut level = None;
    let mut samples = 5_usize;
    let mut warmups = 1_usize;
    let mut matrix = false;
    let mut index = 0;
    while index < args.len() {
        match args[index].as_str() {
            "--lane" => {
                index += 1;
                lane = args.get(index).cloned();
            },
            "--level" => {
                index += 1;
                level = args.get(index).cloned();
            },
            "--samples" => {
                index += 1;
                samples = args
                    .get(index)
                    .expect("--samples needs a value")
                    .parse()
                    .expect("--samples is numeric");
            },
            "--warmups" => {
                index += 1;
                warmups = args
                    .get(index)
                    .expect("--warmups needs a value")
                    .parse()
                    .expect("--warmups is numeric");
            },
            "--matrix" => matrix = true,
            "--help" => {
                println!("--lane NAME --level small|medium|large [--samples N] [--warmups N]");
                println!("--matrix runs the complete 24-lane matrix");
                return;
            },
            other => panic!("unknown argument {other}"),
        }
        index += 1;
    }
    if matrix {
        for current_level in LEVELS {
            for current_lane in LANES {
                run(current_level, current_lane, samples, warmups);
            }
        }
        return;
    }
    let lane = lane.expect("--lane is required unless --matrix is used");
    let level = level.expect("--level is required unless --matrix is used");
    run(level_by_name(&level), &lane, samples, warmups);
}
