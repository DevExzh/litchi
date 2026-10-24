#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "structural insertion fixtures deliberately fail at the assertion site"
)]

//! Native complete-catalog worksheet insertion transactions.
//!
//! Insertion is a workbook-structure operation.  These tests keep the source
//! tab identities visible through the operation: existing sheet parts and
//! relationships stay attached to their authored tabs, the new worksheet has
//! a valid empty binary index, and every modeled catalog ordinal either moves
//! with its owner or refuses the transaction before publication.

use std::collections::BTreeSet;

use litchi_core::Position;
use litchi_core::sheet::{Cell as _, Worksheet as _};
use litchi_core::sheet::{CellValue, WorkbookTrait as _};
use litchi_opc::constants::{content_type, relationship_type};
use litchi_opc::{BlobPart, PackURI, TargetMode};
use litchi_xlsb::Package;
use litchi_xlsb::builder::chart::{Anchor, Chart};
use litchi_xlsb::builder::{self, Builder};
use litchi_xlsb::cell_values::TransferLimits;
use litchi_xlsb::cell_values::structure::Patch;
use litchi_xlsb::data_model::{Definition, Model};
use litchi_xlsb::named_ranges::Definition as NamedRange;
use litchi_xlsb::package::PackageError;
use litchi_xlsb::package::connections::{Connection, Connections, SourceType};
use litchi_xlsb::raw::{Encoder, Kind, Records, kind};
use litchi_xlsb::shared_workbook::{
    Catalog, Guid, Header, Info, RawRecord, RevisionHeaders, RevisionLog, ShortDateTime, User,
    UserNames,
};
use litchi_xlsb::sheet::{Kind as SheetKind, State};
use litchi_xlsb::volatile_dependencies::{
    CONTENT_TYPE as VOLATILE_CONTENT_TYPE, CachedValue, CellReference, DEFAULT_PART_NAME,
    Dependencies, DependencyKind, MainTopic, Topic, VolatileType,
};

#[derive(Debug, Clone)]
struct BundleEntry {
    name: String,
    rel_id: String,
    itab_id: u32,
    part: PackURI,
}

fn worksheets(names: &[&str]) -> Package {
    let mut builder = Builder::new();
    for (index, name) in names.iter().enumerate() {
        let mut sheet = builder::Worksheet::new(*name);
        sheet.set_cell(0, 0, format!("{name} identity {index}"));
        builder.add_worksheet(sheet);
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
    let mut data = builder::Worksheet::new("Data");
    data.set_cell(0, 0, "data identity");
    builder.add_worksheet(data);
    let mut tail = builder::Worksheet::new("Tail");
    tail.set_cell(0, 0, "tail identity");
    builder.add_worksheet(tail);
    builder.build().expect("chart-first package")
}

fn catalog(package: &Package) -> Vec<(String, SheetKind, State)> {
    package
        .workbook()
        .expect("workbook projection")
        .sheets()
        .map(|tab| (tab.name().to_owned(), tab.kind(), tab.state()))
        .collect()
}

fn catalog_names(package: &Package) -> Vec<String> {
    catalog(package)
        .into_iter()
        .map(|(name, _, _)| name)
        .collect()
}

fn cell_by_name(package: &Package, name: &str) -> CellValue {
    package
        .workbook()
        .expect("workbook projection")
        .worksheet_by_name(name)
        .expect("worksheet by name")
        .cell_by_coordinate("A1")
        .expect("cell by coordinate")
        .value()
        .clone()
}

fn workbook_blob(package: &Package) -> Vec<u8> {
    let uri = PackURI::new("/xl/workbook.bin").expect("workbook URI");
    package
        .opc_package()
        .get_part(&uri)
        .expect("workbook part")
        .blob()
        .to_vec()
}

fn bundle_entries(package: &Package) -> Vec<BundleEntry> {
    let workbook_uri = PackURI::new("/xl/workbook.bin").expect("workbook URI");
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
            let strings_offset = if (1..=0xffff).contains(&itab_id) {
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

fn sheet_part(package: &Package, name: &str) -> PackURI {
    bundle_entries(package)
        .into_iter()
        .find(|entry| entry.name == name)
        .unwrap_or_else(|| panic!("missing sheet {name}"))
        .part
}

fn app_titles(package: &Package) -> Vec<String> {
    let uri = PackURI::new("/docProps/app.xml").expect("app URI");
    let text = std::str::from_utf8(
        package
            .opc_package()
            .get_part(&uri)
            .expect("app part")
            .blob(),
    )
    .expect("app XML UTF-8");
    text.split_once("<TitlesOfParts>")
        .expect("titles container")
        .1
        .split("<vt:lpstr>")
        .skip(1)
        .filter_map(|value| value.split_once("</vt:lpstr>").map(|(value, _)| value))
        .map(|value| value.to_owned())
        .collect()
}

fn operation_offset(wire: &[u8]) -> usize {
    let before = usize::try_from(u64::from_le_bytes(
        wire[8..16].try_into().expect("before length"),
    ))
    .expect("before length fits usize");
    let after = usize::try_from(u64::from_le_bytes(
        wire[16..24].try_into().expect("after length"),
    ))
    .expect("after length fits usize");
    32 + before + after
}

fn replace_workbook<F>(package: &Package, mut rewrite: F) -> Package
where
    F: FnMut(&[u8], &mut Vec<u8>),
{
    let mut opc = package.opc_package().clone();
    let uri = PackURI::new("/xl/workbook.bin").expect("workbook URI");
    let source = opc.get_part(&uri).expect("workbook part").blob().to_vec();
    let mut output = Vec::new();
    rewrite(&source, &mut output);
    opc.get_part_mut(&uri)
        .expect("workbook part")
        .set_blob(output);
    Package::from_opc(opc).expect("rewritten workbook package")
}

fn rewrite_records<F>(source: &[u8], mut rewrite: F) -> Vec<u8>
where
    F: FnMut(Kind, &[u8], &mut Encoder<&mut Vec<u8>>),
{
    let mut output = Vec::new();
    let mut writer = Encoder::new(&mut output, litchi_xlsb::raw::Limits::DEFAULT);
    for item in Records::new(source, litchi_xlsb::raw::Limits::DEFAULT) {
        let record = item.expect("workbook record");
        rewrite(record.kind(), record.payload(), &mut writer);
    }
    output
}

fn with_book_view_indices(package: &Package, first: u32, active: u32) -> Package {
    replace_workbook(package, |source, output| {
        *output = rewrite_records(source, |record_kind, payload, writer| {
            let mut payload = payload.to_vec();
            if record_kind == kind::BOOK_VIEW {
                assert_eq!(payload.len(), 29, "BrtBookView payload shape");
                payload[20..24].copy_from_slice(&first.to_le_bytes());
                payload[24..28].copy_from_slice(&active.to_le_bytes());
            }
            writer
                .write_record(record_kind, &payload)
                .expect("rewritten workbook record");
        });
    })
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

fn with_opaque_sheet_record(package: &Package, name: &str) -> Package {
    let part = sheet_part(package, name);
    let mut opc = package.opc_package().clone();
    let original = opc.get_part(&part).expect("sheet part").blob().to_vec();
    let opaque_kind = Kind::new(0x0ffe).expect("opaque worksheet record kind");
    let mut opaque = Vec::new();
    Encoder::new(&mut opaque, litchi_xlsb::raw::Limits::DEFAULT)
        .write_record(opaque_kind, b"opaque-insert-preservation")
        .expect("opaque record");
    let end_sheet = Records::new(&original, litchi_xlsb::raw::Limits::DEFAULT)
        .map(|item| item.expect("sheet record"))
        .find(|record| record.kind() == kind::END_SHEET)
        .expect("end sheet record")
        .offset();
    let mut rewritten = Vec::with_capacity(original.len() + opaque.len());
    rewritten.extend_from_slice(&original[..end_sheet]);
    rewritten.extend_from_slice(&opaque);
    rewritten.extend_from_slice(&original[end_sheet..]);
    opc.get_part_mut(&part)
        .expect("sheet part")
        .set_blob(rewritten);
    Package::from_opc(opc).expect("opaque worksheet package")
}

fn named_range_scopes(package: &Package) -> Vec<(String, u32)> {
    Records::new(&workbook_blob(package), litchi_xlsb::raw::Limits::DEFAULT)
        .filter_map(|item| {
            let record = item.expect("workbook record");
            (record.kind() == kind::NAME).then(|| {
                let scope = u32::from_le_bytes(record.payload()[5..9].try_into().unwrap());
                let definition =
                    NamedRange::parse(record.payload(), litchi_xlsb::raw::Limits::DEFAULT)
                        .expect("BrtName");
                (definition.name, scope)
            })
        })
        .collect()
}

fn with_extern_sheet_fixture(first: i32, last: i32) -> Package {
    let mut builder = Builder::new();
    for name in ["A", "B", "C"] {
        builder.add_worksheet(builder::Worksheet::new(name));
    }
    builder
        .add_external_link(
            litchi_xlsb::external_link::Link::workbook(
                "file:///inert/book.xlsx",
                (0..9).map(|i| format!("External{i}")).collect(),
                Vec::new(),
            )
            .expect("external link"),
        )
        .expect("external link");
    let package = builder.build().expect("external fixture");
    replace_workbook(&package, |source, output| {
        *output = rewrite_records(source, |record_kind, payload, writer| {
            let mut payload = payload.to_vec();
            if record_kind == kind::EXTERN_SHEET {
                assert!(payload.len() >= 4 + 3 * 12);
                let offset = 4 + 2 * 12;
                payload[offset..offset + 4].copy_from_slice(&0_u32.to_le_bytes());
                payload[offset + 4..offset + 8].copy_from_slice(&first.to_le_bytes());
                payload[offset + 8..offset + 12].copy_from_slice(&last.to_le_bytes());
                let external_offset = 4 + 3 * 12;
                payload[external_offset..external_offset + 4].copy_from_slice(&1_u32.to_le_bytes());
                payload[external_offset + 4..external_offset + 8]
                    .copy_from_slice(&7_i32.to_le_bytes());
                payload[external_offset + 8..external_offset + 12]
                    .copy_from_slice(&8_i32.to_le_bytes());
            }
            writer
                .write_record(record_kind, &payload)
                .expect("rewritten workbook record");
        });
    })
}

fn extern_sheet_entries(package: &Package) -> Vec<(u32, i32, i32)> {
    Records::new(&workbook_blob(package), litchi_xlsb::raw::Limits::DEFAULT)
        .find_map(|item| {
            let record = item.expect("workbook record");
            (record.kind() == kind::EXTERN_SHEET).then(|| {
                let payload = record.payload();
                let count = usize::try_from(u32::from_le_bytes(payload[..4].try_into().unwrap()))
                    .expect("XTI count");
                (0..count)
                    .map(|index| {
                        let offset = 4 + index * 12;
                        (
                            u32::from_le_bytes(payload[offset..offset + 4].try_into().unwrap()),
                            i32::from_le_bytes(payload[offset + 4..offset + 8].try_into().unwrap()),
                            i32::from_le_bytes(
                                payload[offset + 8..offset + 12].try_into().unwrap(),
                            ),
                        )
                    })
                    .collect()
            })
        })
        .expect("BrtExternSheet")
}

fn extern_sheet_record_count(package: &Package) -> usize {
    Records::new(&workbook_blob(package), litchi_xlsb::raw::Limits::DEFAULT)
        .filter_map(|item| item.ok())
        .filter(|record| record.kind() == kind::EXTERN_SHEET)
        .count()
}

fn with_unknown_workbook_record(package: &Package) -> Package {
    replace_workbook(package, |source, output| {
        let unknown = Kind::new(0x3ffe).expect("unknown workbook record kind");
        let mut writer = Encoder::new(output, litchi_xlsb::raw::Limits::DEFAULT);
        let mut inserted = false;
        for item in Records::new(source, litchi_xlsb::raw::Limits::DEFAULT) {
            let record = item.expect("workbook record");
            if !inserted && record.kind() == kind::END_BOOK {
                writer
                    .write_record(unknown, &[0xde, 0xad, 0xbe, 0xef])
                    .expect("unknown workbook record");
                inserted = true;
            }
            writer
                .write_record(record.kind(), record.payload())
                .expect("workbook record");
        }
        assert!(inserted, "fixture must contain end-book record");
    })
}

fn package_with_data_model() -> Package {
    let mut builder = Builder::new();
    for name in ["A", "B", "C"] {
        builder.add_worksheet(builder::Worksheet::new(name));
    }
    builder
        .set_data_model(
            Model::from_bytes(
                Definition::default(),
                vec![1, 2, 3, 4],
                litchi_xlsb::data_model::SemanticLimits::DEFAULT,
            )
            .unwrap(),
        )
        .unwrap();
    builder.build().unwrap()
}

fn package_with_connections() -> Package {
    let mut builder = Builder::new();
    for name in ["A", "B", "C"] {
        builder.add_worksheet(builder::Worksheet::new(name));
    }
    builder
        .set_connections(Connections {
            connections: vec![Connection {
                connection_id: 1,
                source_type: SourceType::Odbc,
                name: "Connection".to_owned(),
                ..Connection::default()
            }],
        })
        .unwrap();
    builder.build().unwrap()
}

fn package_with_calc_chain() -> Package {
    let mut opc = worksheets(&["A", "B", "C"]).into_opc();
    let workbook = opc.main_document_part().unwrap().partname().clone();
    let target = PackURI::new("/xl/calcChain.bin").unwrap();
    opc.add_part(Box::new(BlobPart::new(
        target.clone(),
        "application/vnd.ms-excel.calcChain".to_owned(),
        Vec::new(),
    )));
    opc.get_part_mut(&workbook)
        .unwrap()
        .rels_mut()
        .try_add_relationship(
            relationship_type::CALC_CHAIN.to_owned(),
            target.relative_ref(workbook.base_uri()),
            "rIdCalcChain".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
    Package::from_opc(opc).unwrap()
}

fn package_with_vba_project() -> Package {
    let mut opc = worksheets(&["A", "B", "C"]).into_opc();
    let workbook = opc.main_document_part().unwrap().partname().clone();
    let target = PackURI::new("/xl/vbaProject.bin").unwrap();
    opc.add_part(Box::new(BlobPart::new(
        target.clone(),
        content_type::OFC_VBA_PROJECT.to_owned(),
        b"opaque-vba-project".to_vec(),
    )));
    opc.get_part_mut(&workbook)
        .unwrap()
        .rels_mut()
        .try_add_relationship(
            relationship_type::VBA_PROJECT.to_owned(),
            target.relative_ref(workbook.base_uri()),
            "rIdVbaProject".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
    Package::from_opc(opc).unwrap()
}

fn package_with_shared_workbook() -> Package {
    let source = worksheets(&["A", "B", "C"]);
    let guid = Guid::from_bytes([7; 16]);
    let date = ShortDateTime {
        year: 2024,
        month: 1,
        day: 1,
        hour: 9,
        minute: 30,
        second: 0,
        weekday: 1,
    };
    let catalog = Catalog {
        users: Some(UserNames::new(vec![User {
            id: 1,
            guid,
            opened_at: date,
            name: "Alice".to_owned(),
        }])),
        headers: Some(RevisionHeaders::new(
            Info {
                guid,
                root_guid: guid,
                revision_id: 0,
                version: 1,
                has_revisions: false,
                no_revision_history: true,
                protected: false,
                revision_history_interval: 0,
            },
            vec![Header {
                guid,
                saved_at: date,
                next_sheet_id: 4,
                revision_min: 0,
                revision_max: 0,
                user_name: "Alice".to_owned(),
                relationship_id: "rIdLog".to_owned(),
                sheet_ids: vec![1, 2, 3],
                reviewed: Vec::new(),
            }],
        )),
        logs: vec![RevisionLog::new(vec![RawRecord::new(
            0x3fff,
            b"opaque-log".to_vec(),
        )])],
    };
    let mut edit = source.edit_shared_workbook().unwrap();
    edit.replace(catalog).unwrap();
    edit.commit().unwrap().into_package()
}

fn package_with_volatile_dependencies() -> Package {
    let package = worksheets(&["A", "B", "C"]);
    let dependencies = Dependencies::new(vec![VolatileType {
        kind: DependencyKind::Rtd,
        mains: vec![MainTopic {
            first: "server".to_owned(),
            topics: vec![Topic {
                subtopics: vec!["topic".to_owned()],
                value: CachedValue::Number(42.5),
                references: vec![
                    CellReference::new(1, 2, 0).unwrap(),
                    CellReference::new(3, 4, 2).unwrap(),
                ],
            }],
        }],
    }]);
    let mut edit = package.edit_volatile_dependencies().unwrap();
    edit.set_dependencies(dependencies).unwrap();
    let commit = edit.commit().unwrap();
    package
        .apply_volatile_dependencies(commit.patch())
        .map(litchi_core::edit::Commit::into_package)
        .unwrap()
}

fn volatile_ref_indices(package: &Package) -> Vec<u32> {
    let uri = PackURI::new(DEFAULT_PART_NAME).unwrap();
    let blob = package
        .opc_package()
        .get_part(&uri)
        .expect("volatile owner")
        .blob();
    Records::new(blob, litchi_xlsb::raw::Limits::DEFAULT)
        .filter_map(|item| {
            let record = item.ok()?;
            (record.kind() == kind::VOL_REF)
                .then(|| u32::from_le_bytes(record.payload()[8..12].try_into().unwrap()))
        })
        .collect()
}

fn add_part_with_content_type(package: Package, path: &str, content: &str) -> Package {
    let mut opc = package.into_opc();
    opc.try_add_part(Box::new(BlobPart::new(
        PackURI::new(path).unwrap(),
        content.to_owned(),
        b"opaque-ordinal-owner".to_vec(),
    )))
    .unwrap();
    Package::from_opc(opc).unwrap()
}

#[test]
fn mixed_chart_catalog_insertion_at_start_middle_and_end_preserves_order_and_titles() {
    for (position, expected) in [
        (0, vec!["Inserted", "Chart", "Data", "Tail"]),
        (2, vec!["Chart", "Data", "Inserted", "Tail"]),
        (3, vec!["Chart", "Data", "Tail", "Inserted"]),
    ] {
        let source = chart_first_package();
        let mut edit = source.edit().expect("edit");
        edit.insert_sheet(
            Position::new(position),
            "Inserted".to_owned(),
            State::Visible,
        )
        .expect("insert sheet");
        let target = edit.commit().expect("insert commit").into_package();
        let expected = expected.into_iter().map(str::to_owned).collect::<Vec<_>>();
        assert_eq!(catalog_names(&target), expected);
        assert_eq!(app_titles(&target), expected);
        assert_eq!(
            cell_by_name(&target, "Data"),
            CellValue::String("data identity".to_owned())
        );
        assert_eq!(
            cell_by_name(&target, "Tail"),
            CellValue::String("tail identity".to_owned())
        );
    }
}

#[test]
fn empty_insert_has_a_valid_binary_index_and_preserves_existing_opaque_members() {
    let source = with_opaque_sheet_record(&worksheets(&["A", "B"]), "B");
    let before = source.to_bytes().expect("source bytes");
    let before_entries = bundle_entries(&source);
    let before_parts = before_entries
        .iter()
        .map(|entry| {
            (
                entry.name.clone(),
                entry.itab_id,
                entry.rel_id.clone(),
                entry.part.clone(),
                source
                    .opc_package()
                    .get_part(&entry.part)
                    .expect("source sheet part")
                    .blob()
                    .to_vec(),
            )
        })
        .collect::<Vec<_>>();

    let mut edit = source.edit().expect("edit");
    edit.insert_sheet(Position::new(1), "Blank".to_owned(), State::Visible)
        .expect("insert blank sheet");
    let target = edit.commit().expect("insert commit").into_package();
    target
        .workbook()
        .expect("workbook")
        .worksheet_binary_index(1)
        .expect("binary index binding")
        .expect("inserted worksheet binary index");
    assert_eq!(cell_by_name(&target, "Blank"), CellValue::Empty);

    let after_entries = bundle_entries(&target);
    let ids = after_entries
        .iter()
        .map(|entry| entry.itab_id)
        .collect::<Vec<_>>();
    let rels = after_entries
        .iter()
        .map(|entry| entry.rel_id.as_str())
        .collect::<Vec<_>>();
    let parts = after_entries
        .iter()
        .map(|entry| entry.part.as_str())
        .collect::<Vec<_>>();
    assert_eq!(ids.len(), ids.iter().collect::<BTreeSet<_>>().len());
    assert_eq!(rels.len(), rels.iter().collect::<BTreeSet<_>>().len());
    assert_eq!(parts.len(), parts.iter().collect::<BTreeSet<_>>().len());
    for (name, itab_id, rel_id, part, blob) in before_parts {
        let entry = after_entries
            .iter()
            .find(|entry| entry.name == name)
            .expect("existing tab retained");
        assert_eq!(entry.itab_id, itab_id);
        assert_eq!(entry.rel_id, rel_id);
        assert_eq!(entry.part, part);
        assert_eq!(
            target.opc_package().get_part(&part).unwrap().blob(),
            blob.as_slice(),
            "existing opaque sheet member changed for {name}"
        );
    }
    assert_eq!(source.to_bytes().expect("source remains"), before);
}

#[test]
fn repeated_inserts_use_current_staged_slots_and_reject_duplicate_names_atomically() {
    let source = worksheets(&["A", "B"]);
    let mut edit = source.edit().expect("edit");
    edit.insert_sheet(Position::new(1), "First".to_owned(), State::Visible)
        .expect("first insert");
    edit.insert_sheet(Position::new(2), "Second".to_owned(), State::Hidden)
        .expect("second insert");
    let target = edit.commit().expect("staged inserts commit").into_package();
    assert_eq!(catalog_names(&target), vec!["A", "First", "Second", "B"]);
    assert_eq!(
        catalog(&target)
            .into_iter()
            .map(|(_, _, state)| state)
            .collect::<Vec<_>>(),
        vec![
            State::Visible,
            State::Visible,
            State::Hidden,
            State::Visible,
        ]
    );

    let source = worksheets(&["A", "B"]);
    let before = source.to_bytes().unwrap();
    let mut duplicate = source.edit().unwrap();
    assert!(
        duplicate
            .insert_sheet(Position::new(1), "Bad\u{1}".to_owned(), State::Visible)
            .is_err()
    );
    assert!(
        duplicate
            .insert_sheet(Position::new(1), "A".to_owned(), State::Visible)
            .is_err()
    );
    assert!(
        duplicate
            .insert_sheet(Position::new(3), "TooFar".to_owned(), State::Visible)
            .is_err()
    );
    let commit = duplicate.commit().unwrap();
    assert!(commit.patch().is_noop());
    assert_eq!(commit.package().to_bytes().unwrap(), before);
}

#[test]
fn hidden_and_very_hidden_insertions_preserve_the_old_active_physical_tab() {
    for state in [State::Hidden, State::VeryHidden] {
        let source = with_book_view_indices(&worksheets(&["A", "B", "C"]), 1, 2);
        let before = source.to_bytes().unwrap();
        let mut edit = source.edit().unwrap();
        edit.insert_sheet(Position::new(0), "Inserted".to_owned(), state)
            .unwrap();
        let target = edit.commit().unwrap().into_package();
        assert_eq!(catalog(&target)[0].2, state);
        assert_eq!(book_view_indices(&target), vec![(2, 3)]);
        assert_eq!(
            target.workbook().unwrap().active_catalog_position(),
            Some(3)
        );
        assert_eq!(source.to_bytes().unwrap(), before);
    }
}

#[test]
fn durable_tag9_replay_inverse_and_malformed_fields_are_rejected() {
    let source = worksheets(&["A", "B"]);
    let before = source.to_bytes().unwrap();
    let mut edit = source.edit().unwrap();
    edit.insert_sheet(Position::new(1), "Inserted".to_owned(), State::Hidden)
        .unwrap();
    let commit = edit.commit().unwrap();
    let wire = commit.patch().to_bytes(TransferLimits::DEFAULT).unwrap();
    let offset = operation_offset(&wire);
    assert_eq!(wire[offset], 9, "insert operation tag");
    let patch = Patch::from_bytes(
        &wire,
        TransferLimits::DEFAULT,
        litchi_xlsb::raw::Limits::DEFAULT,
    )
    .unwrap();
    let replay = source
        .apply(&patch)
        .map(litchi_xlsb::cell_values::structure::Commit::into_package)
        .unwrap();
    assert_eq!(catalog_names(&replay), vec!["A", "Inserted", "B"]);
    assert_eq!(catalog(&replay)[1].2, State::Hidden);
    assert_eq!(
        replay
            .apply(&patch.inverse())
            .map(litchi_xlsb::cell_values::structure::Commit::into_package)
            .unwrap()
            .to_bytes()
            .unwrap(),
        before
    );

    let mut out_of_range_position = wire.clone();
    out_of_range_position[offset + 1..offset + 9].copy_from_slice(&u64::MAX.to_le_bytes());
    assert!(
        Patch::from_bytes(
            &out_of_range_position,
            TransferLimits::DEFAULT,
            litchi_xlsb::raw::Limits::DEFAULT
        )
        .is_err()
    );

    let name_start = wire[offset..]
        .windows("Inserted".len())
        .position(|window| window == b"Inserted")
        .map(|position| offset + position)
        .expect("encoded inserted name");
    let mut malformed_name_count = wire.clone();
    malformed_name_count[name_start - 8..name_start].copy_from_slice(&u64::MAX.to_le_bytes());
    assert!(
        Patch::from_bytes(
            &malformed_name_count,
            TransferLimits::DEFAULT,
            litchi_xlsb::raw::Limits::DEFAULT
        )
        .is_err()
    );

    let mut malformed_name = wire.clone();
    malformed_name[name_start] = 0xff;
    assert!(
        Patch::from_bytes(
            &malformed_name,
            TransferLimits::DEFAULT,
            litchi_xlsb::raw::Limits::DEFAULT
        )
        .is_err()
    );

    assert_eq!(
        name_start,
        offset + 18,
        "tag9 name follows position, state and length"
    );
    let mut malformed_state = wire.clone();
    malformed_state[offset + 9] = 0xff;
    assert!(
        Patch::from_bytes(
            &malformed_state,
            TransferLimits::DEFAULT,
            litchi_xlsb::raw::Limits::DEFAULT
        )
        .is_err()
    );
}

#[test]
fn insertion_merge_conflicts_on_topology_and_preserves_both_sources() {
    let source = worksheets(&["A", "B", "C"]);
    let before = source.to_bytes().unwrap();
    let mut left = source.edit().unwrap();
    left.insert_sheet(Position::new(1), "Left".to_owned(), State::Visible)
        .unwrap();
    let left = left.commit().unwrap();
    let mut right = source.edit().unwrap();
    right
        .insert_sheet(Position::new(1), "Right".to_owned(), State::Visible)
        .unwrap();
    let right = right.commit().unwrap();
    let merge = left.patch().merge_three_way(right.patch()).unwrap();
    assert!(merge.patch().is_none());
    assert!(!merge.conflicts().is_empty());
    assert_eq!(source.to_bytes().unwrap(), before);
}

#[test]
fn named_range_scopes_and_xti_endpoints_shift_with_an_insertion() {
    let mut builder = Builder::new();
    for name in ["A", "B", "C"] {
        builder.add_worksheet(builder::Worksheet::new(name));
    }
    builder.add_named_range(
        NamedRange::new("Scoped".to_owned(), Some(1)).with_formula(vec![
            litchi_xlsb::formula::ptg_types::PTG_INT,
            1,
            0,
        ]),
    );
    builder.add_named_range(
        NamedRange::new("Global".to_owned(), None).with_formula(vec![
            litchi_xlsb::formula::ptg_types::PTG_INT,
            2,
            2,
            0,
        ]),
    );
    let source = builder.build().unwrap();
    let mut edit = source.edit().unwrap();
    edit.insert_sheet(Position::new(0), "Inserted".to_owned(), State::Visible)
        .unwrap();
    let target = edit.commit().unwrap().into_package();
    assert_eq!(
        named_range_scopes(&target),
        vec![("Scoped".to_owned(), 2), ("Global".to_owned(), u32::MAX),]
    );

    let source = with_extern_sheet_fixture(1, 2);
    let before = extern_sheet_entries(&source);
    let mut edit = source.edit().unwrap();
    edit.insert_sheet(Position::new(0), "Inserted".to_owned(), State::Visible)
        .unwrap();
    let target = edit.commit().unwrap().into_package();
    let after = extern_sheet_entries(&target);
    assert_eq!(after[2], (0, 2, 3));
    assert_eq!(after[3], before[3]);
}

#[test]
fn insertion_inside_a_three_d_range_refuses_without_broadening_the_old_set() {
    let source = with_extern_sheet_fixture(0, 1);
    let before = source.to_bytes().unwrap();
    let mut edit = source.edit().unwrap();
    assert!(matches!(
        edit.insert_sheet(Position::new(1), "Inserted".to_owned(), State::Visible),
        Err(PackageError::Refused(_))
    ));
    assert!(edit.commit().unwrap().patch().is_noop());
    assert_eq!(source.to_bytes().unwrap(), before);
}

#[test]
fn insertion_without_external_links_does_not_invent_a_singleton_xti() {
    let source = replace_workbook(&worksheets(&["A", "B"]), |source, output| {
        *output = rewrite_records(source, |record, payload, writer| {
            if !matches!(
                record,
                kind::BEGIN_EXTERNALS | kind::SUP_SELF | kind::EXTERN_SHEET | kind::END_EXTERNALS
            ) {
                writer.write_record(record, payload).unwrap();
            }
        });
    });
    assert_eq!(extern_sheet_record_count(&source), 0);
    let mut edit = source.edit().unwrap();
    edit.insert_sheet(Position::new(1), "Inserted".to_owned(), State::Visible)
        .unwrap();
    let target = edit.commit().unwrap().into_package();
    assert_eq!(extern_sheet_record_count(&target), 0);
}

fn assert_insert_guard(package: Package, label: &str) {
    let before = package.to_bytes().unwrap();
    let mut edit = package.edit().unwrap();
    let error = edit
        .insert_sheet(Position::new(1), "Inserted".to_owned(), State::Visible)
        .expect_err("opaque insertion dependency guard");
    assert!(edit.commit().unwrap().patch().is_noop());
    assert!(
        if label == "orphan volatile owner" {
            matches!(error, PackageError::InvalidFormat(_))
        } else {
            matches!(error, PackageError::Refused(_))
        },
        "{label}: expected typed dependency error, got {error:?}"
    );
    assert_eq!(
        package.to_bytes().unwrap(),
        before,
        "{label}: source changed"
    );
}

#[test]
fn opaque_ordinal_owners_are_typed_insertion_guards() {
    let cases = [
        ("calculation chain", package_with_calc_chain()),
        ("VBA project", package_with_vba_project()),
        ("shared workbook", package_with_shared_workbook()),
        ("connections", package_with_connections()),
        ("data model", package_with_data_model()),
        (
            "orphan volatile owner",
            add_part_with_content_type(
                worksheets(&["A", "B", "C"]),
                "/xl/orphanVolatile.bin",
                VOLATILE_CONTENT_TYPE,
            ),
        ),
    ];
    for (label, package) in cases {
        assert_insert_guard(package, label);
    }
}

#[test]
fn modeled_volatile_owner_shifts_inserted_sheet_references_and_inverse_is_exact() {
    let source = package_with_volatile_dependencies();
    let before = source.to_bytes().unwrap();
    assert_eq!(volatile_ref_indices(&source), vec![0, 2]);
    let mut edit = source.edit().unwrap();
    edit.insert_sheet(Position::new(1), "Inserted".to_owned(), State::Visible)
        .unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(volatile_ref_indices(commit.package()), vec![0, 3]);
    assert_eq!(
        commit
            .package()
            .apply(&commit.patch().inverse())
            .map(litchi_xlsb::cell_values::structure::Commit::into_package)
            .unwrap()
            .to_bytes()
            .unwrap(),
        before
    );
}

#[test]
fn unknown_workbook_records_refuse_insertion_before_publication() {
    assert_insert_guard(
        with_unknown_workbook_record(&worksheets(&["A", "B", "C"])),
        "unknown workbook record",
    );
}

#[test]
fn exact_source_is_untouched_when_insertion_staging_or_commit_fails() {
    let source = worksheets(&["A", "B"]);
    let before = source.to_bytes().unwrap();
    let mut edit = source.edit().unwrap();
    assert!(
        edit.insert_sheet(
            Position::new(usize::MAX),
            "OutOfRange".to_owned(),
            State::Visible
        )
        .is_err()
    );
    assert!(
        edit.insert_sheet(Position::new(1), "A".to_owned(), State::Visible)
            .is_err()
    );
    assert_eq!(source.to_bytes().unwrap(), before);
    let commit = edit.commit().unwrap();
    assert!(commit.patch().is_noop());
    assert_eq!(commit.package().to_bytes().unwrap(), before);
}
