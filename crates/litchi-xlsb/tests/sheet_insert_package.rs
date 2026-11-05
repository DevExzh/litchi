#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "package fixtures deliberately fail at the assertion that owns the contract"
)]

//! OPC-topology regressions for complete-catalog worksheet insertion.
//!
//! The structural insertion tests cover the workbook record contract.  These
//! fixtures keep the package graph visible as well: the workbook may live
//! outside `/xl`, strict worksheet relationships must retain their flavor,
//! relationship and tab identifiers must remain unique, and optional package
//! properties must stay optional.

use std::collections::BTreeSet;

use litchi_core::Position;
use litchi_opc::constants::relationship_type;
use litchi_opc::{BlobPart, PackURI, Part as _, TargetMode};
use litchi_xlsb::Package;
use litchi_xlsb::builder::chart::{Anchor, Chart};
use litchi_xlsb::builder::{self, Builder};
use litchi_xlsb::raw::{Encoder, Records, kind};
use litchi_xlsb::sheet::State;

const BINARY_INDEX_RELATIONSHIP: &str =
    "http://schemas.microsoft.com/office/2006/relationships/xlBinaryIndex";

#[derive(Debug, Clone)]
struct BundleEntry {
    name: String,
    rel_id: String,
    itab_id: u32,
    part: PackURI,
}

fn worksheets(names: &[&str]) -> Package {
    let mut builder = Builder::new();
    for name in names {
        builder.add_worksheet(builder::Worksheet::new(*name));
    }
    builder.build().expect("worksheet package")
}

fn chart_first_package() -> Package {
    let chart = Chart::bar_chart_with_cache(
        "Sales",
        "Data!$A$1:$A$2",
        &["North", "South"],
        "Data!$B$1:$B$2",
        &[42.0, 55.0],
        Anchor::new(0, 0, 10, 20),
    )
    .expect("chart fixture");
    let mut builder = Builder::new();
    builder
        .add_chart_sheet(builder::ChartSheet::new("Chart", chart))
        .expect("chart sheet fixture");
    builder.add_worksheet(builder::Worksheet::new("Data"));
    builder.build().expect("chart-first package")
}

fn workbook_uri(package: &Package) -> PackURI {
    package
        .opc_package()
        .main_document_part()
        .expect("main workbook relationship")
        .partname()
        .clone()
}

fn workbook_blob(package: &Package) -> Vec<u8> {
    let uri = workbook_uri(package);
    package
        .opc_package()
        .get_part(&uri)
        .expect("workbook part")
        .blob()
        .to_vec()
}

fn bundle_entries(package: &Package) -> Vec<BundleEntry> {
    let workbook_uri = workbook_uri(package);
    let workbook = package
        .opc_package()
        .get_part(&workbook_uri)
        .expect("workbook part");
    Records::new(workbook.blob(), litchi_xlsb::raw::Limits::DEFAULT)
        .map(|item| item.expect("workbook record"))
        .filter(|record| record.kind() == kind::BUNDLE_SH)
        .map(|record| {
            let payload = record.payload();
            let itab_id = u32::from_le_bytes(payload[4..8].try_into().expect("itab id"));
            let strings_offset = if (1..=u32::from(u16::MAX)).contains(&itab_id) {
                8
            } else {
                12
            };
            let rel_units = usize::try_from(u32::from_le_bytes(
                payload[strings_offset..strings_offset + 4]
                    .try_into()
                    .expect("relationship length"),
            ))
            .expect("relationship length fits");
            let rel_start = strings_offset + 4;
            let rel_end = rel_start + rel_units * 2;
            let rel_id = String::from_utf16(
                &payload[rel_start..rel_end]
                    .chunks_exact(2)
                    .map(|unit| u16::from_le_bytes([unit[0], unit[1]]))
                    .collect::<Vec<_>>(),
            )
            .expect("relationship UTF-16");
            let name_units = usize::try_from(u32::from_le_bytes(
                payload[rel_end..rel_end + 4]
                    .try_into()
                    .expect("name length"),
            ))
            .expect("name length fits");
            let name_start = rel_end + 4;
            let name_end = name_start + name_units * 2;
            let name = String::from_utf16(
                &payload[name_start..name_end]
                    .chunks_exact(2)
                    .map(|unit| u16::from_le_bytes([unit[0], unit[1]]))
                    .collect::<Vec<_>>(),
            )
            .expect("sheet name UTF-16");
            let part = workbook
                .rels()
                .get(&rel_id)
                .expect("sheet relationship")
                .target_partname()
                .expect("sheet target");
            BundleEntry {
                name,
                rel_id,
                itab_id,
                part,
            }
        })
        .collect()
}

fn find_entry(package: &Package, name: &str) -> BundleEntry {
    bundle_entries(package)
        .into_iter()
        .find(|entry| entry.name == name)
        .unwrap_or_else(|| panic!("missing sheet {name}"))
}

fn sheet_part(package: &Package, name: &str) -> PackURI {
    find_entry(package, name).part
}

fn sheet_view_flags(package: &Package, name: &str) -> u16 {
    let part = sheet_part(package, name);
    let blob = package
        .opc_package()
        .get_part(&part)
        .expect("worksheet part")
        .blob();
    Records::new(blob, litchi_xlsb::raw::Limits::DEFAULT)
        .map(|item| item.expect("worksheet record"))
        .find(|record| record.kind() == kind::BEGIN_WS_VIEW)
        .map(|record| {
            u16::from_le_bytes(
                record
                    .payload()
                    .get(..2)
                    .expect("worksheet-view flags")
                    .try_into()
                    .expect("worksheet-view flags length"),
            )
        })
        .expect("worksheet view")
}

fn book_view_indices(package: &Package) -> Vec<(u32, u32)> {
    Records::new(&workbook_blob(package), litchi_xlsb::raw::Limits::DEFAULT)
        .filter_map(|item| {
            let record = item.expect("workbook record");
            (record.kind() == kind::BOOK_VIEW).then(|| {
                (
                    u32::from_le_bytes(record.payload()[20..24].try_into().unwrap()),
                    u32::from_le_bytes(record.payload()[24..28].try_into().unwrap()),
                )
            })
        })
        .collect()
}

fn with_book_view_indices(package: &Package, first: u32, active: u32) -> Package {
    let mut opc = package.opc_package().clone();
    let uri = workbook_uri(package);
    let source = opc.get_part(&uri).expect("workbook part").blob().to_vec();
    let mut output = Vec::new();
    let mut writer = Encoder::new(&mut output, litchi_xlsb::raw::Limits::DEFAULT);
    for item in Records::new(&source, litchi_xlsb::raw::Limits::DEFAULT) {
        let record = item.expect("workbook record");
        let mut payload = record.payload().to_vec();
        if record.kind() == kind::BOOK_VIEW {
            assert_eq!(payload.len(), 29, "BrtBookView payload shape");
            payload[20..24].copy_from_slice(&first.to_le_bytes());
            payload[24..28].copy_from_slice(&active.to_le_bytes());
        }
        writer
            .write_record(record.kind(), &payload)
            .expect("rewritten workbook record");
    }
    opc.get_part_mut(&uri)
        .expect("workbook part")
        .set_blob(output);
    Package::from_opc(opc).expect("book-view fixture")
}

fn remove_app(package: &Package) -> Package {
    let mut opc = package.opc_package().clone();
    let app = PackURI::new("/docProps/app.xml").expect("app URI");
    assert!(opc.remove_part(&app), "source fixture should have app.xml");
    Package::from_opc(opc).expect("package without app.xml")
}

/// Move the main workbook while rewriting every internal target relative to
/// the new owner.  This is the package shape that catches code which assumes
/// `/xl/workbook.bin` when resolving a newly authored worksheet relationship.
fn relocate_workbook(package: &Package, destination: &str) -> Package {
    let mut opc = package.opc_package().clone();
    let source = opc
        .main_document_part()
        .expect("source workbook")
        .partname()
        .clone();
    let destination = PackURI::new(destination).expect("destination workbook URI");
    let (content_type, blob, relationships) = {
        let workbook = opc.get_part(&source).expect("source workbook");
        let relationships = workbook
            .rels()
            .iter()
            .map(|relationship| {
                let target = match relationship.target_mode() {
                    TargetMode::Internal => relationship
                        .target_partname()
                        .expect("internal workbook target")
                        .relative_ref(destination.base_uri()),
                    TargetMode::External => relationship.target_ref().to_owned(),
                };
                (
                    relationship.reltype().to_owned(),
                    target,
                    relationship.r_id().to_owned(),
                    relationship.target_mode(),
                )
            })
            .collect::<Vec<_>>();
        (
            workbook.content_type().to_owned(),
            workbook.blob().to_vec(),
            relationships,
        )
    };
    let (root_id, root_type) = opc
        .rels()
        .iter()
        .find(|relationship| {
            matches!(
                relationship.reltype(),
                relationship_type::OFFICE_DOCUMENT | relationship_type::STRICT_OFFICE_DOCUMENT
            )
        })
        .map(|relationship| {
            (
                relationship.r_id().to_owned(),
                relationship.reltype().to_owned(),
            )
        })
        .expect("root workbook relationship");

    let mut replacement = BlobPart::new(destination.clone(), content_type, blob);
    for (reltype, target, id, mode) in relationships {
        replacement
            .rels_mut()
            .try_add_relationship(reltype, target, id, mode)
            .expect("relocated workbook relationship");
    }
    assert!(opc.remove_part(&source), "source workbook removed");
    opc.add_part(Box::new(replacement));
    opc.rels_mut().remove(&root_id);
    opc.rels_mut()
        .try_add_relationship(
            root_type,
            destination.as_str().trim_start_matches('/').to_owned(),
            root_id,
            TargetMode::Internal,
        )
        .expect("relocated root relationship");
    Package::from_opc(opc).expect("relocated workbook package")
}

fn strict_worksheet_relationships(package: &Package) -> Package {
    let mut opc = package.opc_package().clone();
    let workbook_uri = opc
        .main_document_part()
        .expect("workbook part")
        .partname()
        .clone();
    let worksheet_relationships = opc
        .get_part(&workbook_uri)
        .expect("workbook part")
        .rels()
        .iter()
        .filter(|relationship| relationship.reltype() == relationship_type::WORKSHEET)
        .map(|relationship| {
            (
                relationship.r_id().to_owned(),
                relationship.target_ref().to_owned(),
                relationship.target_mode(),
            )
        })
        .collect::<Vec<_>>();
    assert!(
        !worksheet_relationships.is_empty(),
        "strict fixture has worksheets"
    );
    let workbook = opc.get_part_mut(&workbook_uri).expect("workbook part");
    for (id, target, mode) in worksheet_relationships {
        workbook.rels_mut().remove(&id);
        workbook
            .rels_mut()
            .try_add_relationship(
                relationship_type::STRICT_WORKSHEET.to_owned(),
                target,
                id,
                mode,
            )
            .expect("strict worksheet relationship");
    }
    Package::from_opc(opc).expect("strict worksheet package")
}

fn with_relationship_collision(package: &Package) -> (Package, String) {
    let mut opc = package.opc_package().clone();
    let workbook = opc
        .main_document_part()
        .expect("workbook part")
        .partname()
        .clone();
    let candidate = (1_u32..)
        .map(|number| format!("rId{number}"))
        .find(|id| {
            opc.get_part(&workbook)
                .expect("workbook part")
                .rels()
                .get(id)
                .is_none()
        })
        .expect("relationship ID candidate");
    let target = PackURI::new("/xl/insert-collision.bin").expect("collision part URI");
    opc.try_add_part(Box::new(BlobPart::new(
        target.clone(),
        "application/octet-stream".to_owned(),
        b"collision owner".to_vec(),
    )))
    .expect("collision part");
    opc.get_part_mut(&workbook)
        .expect("workbook part")
        .rels_mut()
        .try_add_relationship(
            "urn:litchi:test:occupied-sheet-id".to_owned(),
            target.relative_ref(workbook.base_uri()),
            candidate.clone(),
            TargetMode::Internal,
        )
        .expect("collision relationship");
    (
        Package::from_opc(opc).expect("collision package"),
        candidate,
    )
}

#[test]
fn custom_workbook_uri_uses_relative_targets_and_keeps_absent_app_absent() {
    let source = relocate_workbook(
        &remove_app(&worksheets(&["Active", "Other"])),
        "/custom/workbook.bin",
    );
    assert_eq!(workbook_uri(&source).as_str(), "/custom/workbook.bin");
    assert!(
        source
            .opc_package()
            .get_part(&PackURI::new("/docProps/app.xml").unwrap())
            .is_err()
    );

    let mut edit = source.edit().expect("edit");
    edit.insert_sheet(Position::new(1), "Inserted".to_owned(), State::Visible)
        .expect("insert");
    let target = edit.commit().expect("commit").into_package();
    assert_eq!(workbook_uri(&target).as_str(), "/custom/workbook.bin");
    assert!(
        target
            .opc_package()
            .get_part(&PackURI::new("/docProps/app.xml").unwrap())
            .is_err()
    );

    let workbook_uri = workbook_uri(&target);
    let workbook = target
        .opc_package()
        .get_part(&workbook_uri)
        .expect("custom workbook part");
    for relationship in workbook.rels().iter() {
        if relationship.target_mode() == TargetMode::Internal {
            let part = relationship
                .target_partname()
                .expect("internal relationship resolves");
            assert!(
                target.opc_package().contains_part(&part),
                "relationship {} resolves to missing {}",
                relationship.r_id(),
                part.as_str()
            );
        }
    }
    let inserted = find_entry(&target, "Inserted");
    assert!(inserted.part.as_str().starts_with("/xl/worksheets/"));
    assert!(!inserted.part.as_str().starts_with("/custom/"));
    let sheet = target
        .opc_package()
        .get_part(&inserted.part)
        .expect("inserted worksheet");
    let index_target = sheet
        .rels()
        .iter()
        .find(|relationship| relationship.reltype() == BINARY_INDEX_RELATIONSHIP)
        .expect("inserted binary-index relationship")
        .target_partname()
        .expect("binary-index target");
    assert!(target.opc_package().contains_part(&index_target));
}

#[test]
fn strict_chart_first_insertion_preserves_worksheet_relationship_flavor() {
    let source = strict_worksheet_relationships(&chart_first_package());
    let mut edit = source.edit().expect("edit");
    edit.insert_sheet(Position::new(1), "Inserted".to_owned(), State::Visible)
        .expect("insert");
    let target = edit.commit().expect("commit").into_package();
    let inserted = find_entry(&target, "Inserted");
    let workbook = target
        .opc_package()
        .get_part(&workbook_uri(&target))
        .expect("workbook part");
    let relationship = workbook
        .rels()
        .get(&inserted.rel_id)
        .expect("inserted worksheet relationship");
    assert_eq!(
        relationship.reltype(),
        relationship_type::STRICT_WORKSHEET,
        "chart-first strict source must author strict worksheet relationships"
    );
    assert!(target.opc_package().get_part(&inserted.part).is_ok());
}

#[test]
fn relationship_and_itab_allocators_skip_occupied_candidates() {
    let (source, collision_id) = with_relationship_collision(&worksheets(&["A", "B"]));
    let old_ids = bundle_entries(&source)
        .into_iter()
        .map(|entry| (entry.name, entry.rel_id, entry.itab_id))
        .collect::<Vec<_>>();
    let mut edit = source.edit().expect("edit");
    edit.insert_sheet(Position::new(1), "Inserted".to_owned(), State::Visible)
        .expect("insert");
    let target = edit.commit().expect("commit").into_package();
    let entries = bundle_entries(&target);
    let inserted = entries
        .iter()
        .find(|entry| entry.name == "Inserted")
        .expect("inserted sheet");
    assert_ne!(inserted.rel_id, collision_id);
    assert_eq!(
        target
            .opc_package()
            .get_part(&workbook_uri(&target))
            .expect("workbook")
            .rels()
            .get(&collision_id)
            .expect("collision relationship retained")
            .reltype(),
        "urn:litchi:test:occupied-sheet-id"
    );
    let relation_ids = entries
        .iter()
        .map(|entry| entry.rel_id.as_str())
        .collect::<Vec<_>>();
    assert_eq!(
        relation_ids.len(),
        relation_ids.iter().collect::<BTreeSet<_>>().len()
    );
    let tab_ids = entries
        .iter()
        .map(|entry| entry.itab_id)
        .collect::<Vec<_>>();
    assert_eq!(tab_ids.len(), tab_ids.iter().collect::<BTreeSet<_>>().len());
    for (name, rel_id, itab_id) in old_ids {
        let retained = entries
            .iter()
            .find(|entry| entry.name == name)
            .expect("old sheet retained");
        assert_eq!(retained.rel_id, rel_id);
        assert_eq!(retained.itab_id, itab_id);
    }
}

#[test]
fn inserted_worksheet_is_not_selected_and_existing_active_tab_stays_active() {
    let source = with_book_view_indices(&worksheets(&["Active", "Other"]), 0, 0);
    assert_eq!(book_view_indices(&source), vec![(0, 0)]);
    assert_ne!(sheet_view_flags(&source, "Active") & 0x0040, 0);
    let mut edit = source.edit().expect("edit");
    edit.insert_sheet(Position::new(1), "Inserted".to_owned(), State::Visible)
        .expect("insert");
    let target = edit.commit().expect("commit").into_package();
    assert_eq!(book_view_indices(&target), vec![(0, 0)]);
    assert_eq!(
        target
            .workbook()
            .expect("workbook")
            .active_catalog_position(),
        Some(0)
    );
    assert_ne!(sheet_view_flags(&target, "Active") & 0x0040, 0);
    assert_eq!(sheet_view_flags(&target, "Inserted") & 0x0040, 0);
}
