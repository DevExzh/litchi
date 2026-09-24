#![allow(
    clippy::arbitrary_source_item_ordering,
    clippy::expect_used,
    clippy::indexing_slicing,
    clippy::too_many_lines,
    clippy::unwrap_used,
    reason = "This standalone profile uses deterministic authored fixtures and fail-fast assertions."
)]

//! Matched XLSX Custom Data lifecycle performance profile.
//!
//! The fixture is authored from the bounded graph shape used by the public
//! lifecycle tests. It is not a native Office corpus. Each process runs one
//! lane/size/sample so allocator and RSS observations do not include prior
//! lanes.

mod support;

use std::hint::black_box;
use std::time::Instant;

use litchi_opc::{BlobPart, OpcPackage, PackURI};
use litchi_xlsx::Package;
use litchi_xlsx::custom_data::{RemovalDisposition, Snapshot, Transaction};
use serde_json::json;
use sha2::{Digest, Sha256};
use soapberry_zip::office::StreamingArchiveWriter;

#[global_allocator]
static ALLOCATOR: support::CountingAllocator = support::CountingAllocator;

const PROFILE_SCHEMA: &str = "xlsx-custom-data-performance-v1";
const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const X14: &str = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
const MCE: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const CONNECTIONS_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections";
const PROPERTIES_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customDataProps";
const DATA_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customData";
const CONNECTIONS_CONTENT: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml";
const PROPERTIES_CONTENT: &str =
    "application/vnd.openxmlformats-officedocument.customDataProperties+xml";
const WORKBOOK: &str = "/xl/workbook.xml";
const CONNECTIONS: &str = "/xl/connections.xml";
const EMBEDDED_EXTENSION: &str = "{D79990A0-CA42-45E3-83F4-45C500A0EAA5}";

#[derive(Clone, Copy, Debug)]
struct Level {
    name: &'static str,
    storages: usize,
    references: usize,
    payload_bytes: usize,
}

// The large class is deliberately large enough to expose package-member ×
// storage-count scans while remaining a bounded CI-sized workload.
const LEVELS: [Level; 3] = [
    Level {
        name: "small",
        storages: 2,
        references: 16,
        payload_bytes: 1_024,
    },
    Level {
        name: "medium",
        storages: 16,
        references: 128,
        payload_bytes: 4_096,
    },
    Level {
        name: "large",
        storages: 64,
        references: 512,
        payload_bytes: 16_384,
    },
];

const LANES: [&str; 5] = [
    "read",
    "noop-commit-save",
    "payload-replacement",
    "rename-binding-rewrite",
    "remove-inverse",
];

#[derive(Debug)]
struct Fixture {
    bytes: Vec<u8>,
    source_sha256: [u8; 32],
    level: Level,
    ids: Vec<String>,
    payloads: Vec<Vec<u8>>,
    replacement_payload: Vec<u8>,
    renamed_reference_count: usize,
}

#[derive(Clone, Copy, Debug)]
struct Validation {
    output_sha256: [u8; 32],
    storage_count: usize,
    reference_count: usize,
    source_exact_noop: bool,
    payloads_exact: bool,
    bindings_exact: bool,
    inverse_exact: bool,
}

fn main() {
    let (lane, level_name, sample, warmups) = parse_args();
    let level = LEVELS
        .iter()
        .copied()
        .find(|candidate| candidate.name == level_name)
        .unwrap_or_else(|| panic!("unknown level {level_name:?}"));
    assert!(LANES.contains(&lane.as_str()), "unknown lane {lane}");
    let fixture = make_fixture(level);

    // Warm only the same lane/level path. The warm result is dropped before
    // counters reset, so persistent fixture allocations are the baseline.
    for _ in 0..warmups {
        black_box(run_lane(&lane, &fixture))
            .unwrap_or_else(|error| panic!("warm-up {lane}/{level_name}/{sample} failed: {error}"));
    }
    support::reset();
    let before = support::Snapshot::now();
    let started = Instant::now();
    let validation = run_lane(&lane, &fixture)
        .unwrap_or_else(|error| panic!("{lane}/{level_name}/{sample} failed: {error}"));
    let elapsed_ns = started.elapsed().as_nanos();
    let after = support::Snapshot::now();
    let allocation = before.delta(after);
    assert!(
        allocation.balanced(),
        "allocator balance failed: {allocation:?}"
    );
    assert!(
        !allocation.invalid,
        "allocator live-byte underflow observed"
    );
    assert_eq!(
        allocation.allocation_failed, 0,
        "allocator allocation failed"
    );

    println!(
        "{}",
        json!({
            "schema": PROFILE_SCHEMA,
            "source_commit": std::env::var("LITCHI_PROFILE_COMMIT")
                .unwrap_or_else(|_| "unknown".to_owned()),
            "lane": lane,
            "level": level.name,
            "sample": sample,
            "warmups": warmups,
            "fixture": {
                "storages": level.storages,
                "references": level.references,
                "payload_bytes": level.payload_bytes,
                "source_bytes": fixture.bytes.len(),
                "source_sha256": hex(fixture.source_sha256),
            },
            "elapsed_ns": elapsed_ns,
            "allocator": {
                "allocation_calls": allocation.allocation_calls,
                "reallocation_calls": allocation.reallocation_calls,
                "deallocation_calls": allocation.deallocation_calls,
                "direct_bytes": allocation.direct_bytes,
                "realloc_old_bytes": allocation.realloc_old_bytes,
                "realloc_new_bytes": allocation.realloc_new_bytes,
                "requested_bytes": allocation.requested_bytes(),
                "deallocated_bytes": allocation.deallocated_bytes,
                "live_before": allocation.live_before,
                "live_after": allocation.live_after,
                "peak_live_delta": allocation.peak_live_delta,
                "peak_live_absolute": allocation.peak_live_absolute,
                "balanced": allocation.balanced(),
            },
            "validation": {
                "storage_count": validation.storage_count,
                "reference_count": validation.reference_count,
                "source_exact_noop": validation.source_exact_noop,
                "payloads_exact": validation.payloads_exact,
                "bindings_exact": validation.bindings_exact,
                "inverse_exact": validation.inverse_exact,
                "candidate_sha256": hex(validation.output_sha256),
            },
        })
    );
}

fn parse_args() -> (String, String, u64, u32) {
    let mut lane = None;
    let mut level = None;
    let mut sample = None;
    let mut warmups = 1;
    let mut args = std::env::args().skip(1);
    while let Some(argument) = args.next() {
        match argument.as_str() {
            "--lane" => lane = args.next(),
            "--level" => level = args.next(),
            "--sample" => {
                sample = Some(
                    args.next()
                        .expect("--sample requires a value")
                        .parse::<u64>()
                        .expect("--sample is an integer"),
                )
            },
            "--warmups" => {
                warmups = args
                    .next()
                    .expect("--warmups requires a value")
                    .parse::<u32>()
                    .expect("--warmups is an integer");
            },
            other => panic!("unknown argument {other}"),
        }
    }
    (
        lane.expect("--lane is required"),
        level.expect("--level is required"),
        sample.expect("--sample is required"),
        warmups,
    )
}

fn run_lane(lane: &str, fixture: &Fixture) -> Result<Validation, String> {
    match lane {
        "read" => run_read(fixture),
        "noop-commit-save" => run_noop(fixture),
        "payload-replacement" => run_payload_replacement(fixture),
        "rename-binding-rewrite" => run_rename(fixture),
        "remove-inverse" => run_remove_inverse(fixture),
        _ => Err(format!("unsupported lane {lane}")),
    }
}

fn run_read(fixture: &Fixture) -> Result<Validation, String> {
    let package = open_package(fixture)?;
    let opc = package.into_plain_opc();
    let snapshot = Snapshot::load(&opc).map_err(error_text)?;
    let (storage_count, reference_count) = assert_source_snapshot(&snapshot, fixture)?;
    Ok(Validation {
        output_sha256: fixture.source_sha256,
        storage_count,
        reference_count,
        source_exact_noop: true,
        payloads_exact: true,
        bindings_exact: true,
        inverse_exact: true,
    })
}

fn run_noop(fixture: &Fixture) -> Result<Validation, String> {
    let package = open_package(fixture)?;
    let mut opc = package.into_plain_opc();
    let commit = Transaction::new(&mut opc)
        .map_err(error_text)?
        .commit()
        .map_err(error_text)?;
    if commit.changed() || !commit.patch().is_empty() {
        return Err("empty transaction produced a changed commit".into());
    }
    let output = save_package(opc)?;
    let output_sha256 = sha256(&output);
    let reopened = Package::from_bytes(output)
        .map_err(error_text)?
        .into_plain_opc();
    let snapshot = Snapshot::load(&reopened).map_err(error_text)?;
    let (storage_count, reference_count) = assert_source_snapshot(&snapshot, fixture)?;
    Ok(Validation {
        output_sha256,
        storage_count,
        reference_count,
        source_exact_noop: true,
        payloads_exact: true,
        bindings_exact: true,
        inverse_exact: true,
    })
}

fn run_payload_replacement(fixture: &Fixture) -> Result<Validation, String> {
    let package = open_package(fixture)?;
    let mut opc = package.into_plain_opc();
    let mut transaction = Transaction::new(&mut opc).map_err(error_text)?;
    transaction
        .set_data(0, fixture.replacement_payload.clone())
        .map_err(error_text)?;
    let commit = transaction.commit().map_err(error_text)?;
    if !commit.changed() {
        return Err("payload replacement produced an unchanged commit".into());
    }
    let output = save_package(opc)?;
    let output_sha256 = sha256(&output);
    let reopened = Package::from_bytes(output)
        .map_err(error_text)?
        .into_plain_opc();
    let snapshot = Snapshot::load(&reopened).map_err(error_text)?;
    let (storage_count, reference_count) = assert_payload_replaced_snapshot(&snapshot, fixture)?;
    Ok(Validation {
        output_sha256,
        storage_count,
        reference_count,
        source_exact_noop: false,
        payloads_exact: true,
        bindings_exact: true,
        inverse_exact: false,
    })
}

fn run_rename(fixture: &Fixture) -> Result<Validation, String> {
    let package = open_package(fixture)?;
    let mut opc = package.into_plain_opc();
    let mut transaction = Transaction::new(&mut opc).map_err(error_text)?;
    transaction
        .rename(0, "uid-renamed-0000")
        .map_err(error_text)?;
    let commit = transaction.commit().map_err(error_text)?;
    if !commit.changed() {
        return Err("UID rename produced an unchanged commit".into());
    }
    let output = save_package(opc)?;
    let output_sha256 = sha256(&output);
    let reopened = Package::from_bytes(output)
        .map_err(error_text)?
        .into_plain_opc();
    let snapshot = Snapshot::load(&reopened).map_err(error_text)?;
    let (storage_count, reference_count) = assert_renamed_snapshot(&snapshot, fixture)?;
    let renamed = snapshot
        .find("uid-renamed-0000")
        .ok_or("renamed storage missing")?;
    if snapshot.find("uid-0000").is_some()
        || snapshot.connection_references(renamed.id()) != fixture.renamed_reference_count
    {
        return Err("connection binding rewrite differs".into());
    }
    Ok(Validation {
        output_sha256,
        storage_count,
        reference_count,
        source_exact_noop: false,
        payloads_exact: true,
        bindings_exact: true,
        inverse_exact: false,
    })
}

fn run_remove_inverse(fixture: &Fixture) -> Result<Validation, String> {
    let package = open_package(fixture)?;
    let mut opc = package.into_plain_opc();
    let mut transaction = Transaction::new(&mut opc).map_err(error_text)?;
    transaction
        .remove_with(0, RemovalDisposition::DetachConnections)
        .map_err(error_text)?
        .ok_or("removal did not return a storage")?;
    let commit = transaction.commit().map_err(error_text)?;
    if !commit.changed() {
        return Err("removal produced an unchanged commit".into());
    }
    let inverse = commit.patch().inverse();
    inverse.apply(&mut opc).map_err(error_text)?;
    let output = save_package(opc)?;
    let output_sha256 = sha256(&output);
    let reopened = Package::from_bytes(output)
        .map_err(error_text)?
        .into_plain_opc();
    let snapshot = Snapshot::load(&reopened).map_err(error_text)?;
    let (storage_count, reference_count) = assert_source_snapshot(&snapshot, fixture)?;
    if output_sha256 != fixture.source_sha256 {
        return Err("inverse output does not exactly restore source bytes".into());
    }
    Ok(Validation {
        output_sha256,
        storage_count,
        reference_count,
        source_exact_noop: false,
        payloads_exact: true,
        bindings_exact: true,
        inverse_exact: true,
    })
}

fn open_package(fixture: &Fixture) -> Result<Package, String> {
    Package::from_bytes(fixture.bytes.clone()).map_err(error_text)
}

fn save_package(opc: OpcPackage) -> Result<Vec<u8>, String> {
    Package::from_opc(opc)
        .map_err(error_text)?
        .to_plain_bytes()
        .map_err(error_text)
}

fn assert_source_snapshot(
    snapshot: &Snapshot,
    fixture: &Fixture,
) -> Result<(usize, usize), String> {
    if snapshot.len() != fixture.level.storages {
        return Err(format!(
            "storage count {} != {}",
            snapshot.len(),
            fixture.level.storages
        ));
    }
    for (index, id) in fixture.ids.iter().enumerate() {
        let entry = snapshot.find(id).ok_or_else(|| format!("missing {id}"))?;
        if entry.value().data() != fixture.payloads[index].as_slice() {
            return Err(format!("payload differs for {id}"));
        }
    }
    if snapshot.connection_references("uid-0000") != fixture.renamed_reference_count {
        return Err("source binding count differs".into());
    }
    Ok((snapshot.len(), snapshot.connection_references("uid-0000")))
}

fn assert_renamed_snapshot(
    snapshot: &Snapshot,
    fixture: &Fixture,
) -> Result<(usize, usize), String> {
    if snapshot.len() != fixture.level.storages {
        return Err(format!(
            "storage count {} != {}",
            snapshot.len(),
            fixture.level.storages
        ));
    }
    for (index, id) in fixture.ids.iter().enumerate().skip(1) {
        let entry = snapshot.find(id).ok_or_else(|| format!("missing {id}"))?;
        if entry.value().data() != fixture.payloads[index].as_slice() {
            return Err(format!("payload differs for {id}"));
        }
    }
    let renamed = snapshot
        .find("uid-renamed-0000")
        .ok_or("renamed storage missing")?;
    if renamed.value().data() != fixture.payloads[0].as_slice() {
        return Err("renamed payload differs".into());
    }
    Ok((snapshot.len(), snapshot.connection_references(renamed.id())))
}

fn assert_payload_replaced_snapshot(
    snapshot: &Snapshot,
    fixture: &Fixture,
) -> Result<(usize, usize), String> {
    if snapshot.len() != fixture.level.storages {
        return Err(format!(
            "storage count {} != {}",
            snapshot.len(),
            fixture.level.storages
        ));
    }
    for (index, id) in fixture.ids.iter().enumerate() {
        let entry = snapshot.find(id).ok_or_else(|| format!("missing {id}"))?;
        let expected = if index == 0 {
            fixture.replacement_payload.as_slice()
        } else {
            fixture.payloads[index].as_slice()
        };
        if entry.value().data() != expected {
            return Err(format!("payload differs for {id}"));
        }
    }
    Ok((snapshot.len(), snapshot.connection_references("uid-0000")))
}

fn make_fixture(level: Level) -> Fixture {
    let ids = (0..level.storages)
        .map(|index| format!("uid-{index:04}"))
        .collect::<Vec<_>>();
    let payloads = (0..level.storages)
        .map(|storage| {
            (0..level.payload_bytes)
                .map(|offset| offset.wrapping_add(storage * 17) as u8)
                .collect::<Vec<_>>()
        })
        .collect::<Vec<_>>();
    let secondary = level.storages.saturating_sub(1).min(level.references);
    let references = (0..level.references)
        .map(|slot| {
            if slot < secondary {
                ids[slot + 1].clone()
            } else {
                ids[0].clone()
            }
        })
        .collect::<Vec<_>>();
    let replacement_payload = (0..level.payload_bytes + 257)
        .map(|offset| offset.wrapping_mul(3) as u8)
        .collect::<Vec<_>>();
    let bytes = physical_fixture(&ids, &references, &payloads);
    let source_sha256 = sha256(&bytes);
    Fixture {
        bytes,
        source_sha256,
        level,
        ids,
        payloads,
        replacement_payload,
        renamed_reference_count: level.references.saturating_sub(secondary),
    }
}

fn physical_fixture(ids: &[String], references: &[String], payloads: &[Vec<u8>]) -> Vec<u8> {
    assert_eq!(ids.len(), payloads.len());
    let mut package = Package::create()
        .expect("minimal XLSX package creates")
        .into_plain_opc();
    for (index, (id, payload)) in ids.iter().zip(payloads).enumerate() {
        let ordinal = index + 1;
        let properties_uri = format!("/xl/customData/props{ordinal}.xml");
        let data_uri = format!("/xl/customData/data{ordinal}.bin");
        package.add_part(Box::new(BlobPart::new(
            uri(&properties_uri),
            PROPERTIES_CONTENT.to_owned(),
            properties_xml(id, ordinal),
        )));
        package.add_part(Box::new(BlobPart::new(
            uri(&data_uri),
            "application/binary".to_owned(),
            payload.clone(),
        )));
        add_relationship(
            &mut package,
            &properties_uri,
            DATA_REL,
            &format!("data{ordinal}.bin"),
            &format!("rIdCustomData{ordinal}"),
        );
        add_relationship(
            &mut package,
            WORKBOOK,
            PROPERTIES_REL,
            &format!("customData/props{ordinal}.xml"),
            &format!("rIdCustomDataProps{ordinal}"),
        );
    }
    if !references.is_empty() {
        package.add_part(Box::new(BlobPart::new(
            uri(CONNECTIONS),
            CONNECTIONS_CONTENT.to_owned(),
            connections_xml(references),
        )));
        add_relationship(
            &mut package,
            WORKBOOK,
            CONNECTIONS_REL,
            "connections.xml",
            "rIdConnections",
        );
    }
    physical_source(&package)
}

fn properties_xml(id: &str, ordinal: usize) -> Vec<u8> {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><x14:datastoreItem xmlns:x14="{X14}" xmlns:s="{SML}" xmlns:v="urn:synthetic-vendor" id="{id}"><x14:extLst><s:ext uri="urn:synthetic-{ordinal}"><v:opaque marker="retain-{ordinal}"><v:leaf><![CDATA[opaque-{ordinal}]]></v:leaf></v:opaque></s:ext></x14:extLst></x14:datastoreItem>"#
    )
    .into_bytes()
}

fn connections_xml(references: &[String]) -> Vec<u8> {
    let mut output = format!(
        r#"<s:connections xmlns:s="{SML}" xmlns:x14="{X14}" xmlns:mc="{MCE}" xmlns:q="urn:synthetic-opaque">"#
    );
    for (index, id) in references.iter().enumerate() {
        output.push_str(&format!(
            r#"<s:connection id="{}" type="5" refreshedVersion="3"><s:extLst><s:ext uri="{EMBEDDED_EXTENSION}"><x14:connection embeddedDataId="{id}" culture="en-US"><q:opaque marker="keep-{index}"/></x14:connection></s:ext></s:extLst></s:connection>"#,
            index + 7
        ));
    }
    output.push_str("</s:connections>");
    output.into_bytes()
}

fn add_relationship(package: &mut OpcPackage, source: &str, reltype: &str, target: &str, id: &str) {
    package
        .get_part_mut(&uri(source))
        .expect("fixture relationship source exists")
        .rels_mut()
        .add_relationship(reltype.to_owned(), target.to_owned(), id.to_owned(), false);
}

fn physical_source(package: &OpcPackage) -> Vec<u8> {
    let mut writer = StreamingArchiveWriter::new();
    let mut parts = package.iter_parts().collect::<Vec<_>>();
    parts.sort_by(|left, right| left.partname().as_str().cmp(right.partname().as_str()));
    let mut content_types = String::from(
        r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>"#,
    );
    for part in &parts {
        content_types.push_str(&format!(
            r#"<Override PartName="{}" ContentType="{}"/>"#,
            part.partname(),
            part.content_type()
        ));
        writer
            .write_stored(
                part.partname().as_str().trim_start_matches('/'),
                part.blob(),
            )
            .expect("fixture part writes");
        if !part.rels().is_empty() {
            let rels_uri = part.partname().rels_uri().expect("fixture rels URI");
            writer
                .write_stored(
                    rels_uri.as_str().trim_start_matches('/'),
                    part.rels().to_xml().as_bytes(),
                )
                .expect("fixture relationship writes");
        }
    }
    content_types.push_str("</Types>");
    writer
        .write_stored("[Content_Types].xml", content_types.as_bytes())
        .expect("content types write");
    writer
        .write_stored("_rels/.rels", package.rels().to_xml().as_bytes())
        .expect("package relationships write");
    writer.finish_to_bytes().expect("fixture ZIP closes")
}

fn uri(value: &str) -> PackURI {
    PackURI::new(value).expect("fixture URI is valid")
}

fn sha256(bytes: &[u8]) -> [u8; 32] {
    Sha256::digest(bytes).into()
}

fn hex(value: [u8; 32]) -> String {
    value.iter().map(|byte| format!("{byte:02x}")).collect()
}

fn error_text(error: impl std::fmt::Display) -> String {
    error.to_string()
}
