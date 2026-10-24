#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "structural removal fixtures deliberately fail at the assertion site"
)]

//! Native complete-catalog worksheet removal transactions.
//!
//! A removal addresses the immutable source catalog, then publishes one
//! sorted source-position set.  These fixtures keep the physical package
//! graph visible while exercising the durable tag-10 operation and the
//! dependency policies that protect workbook structure.

use std::collections::{BTreeSet, HashMap};

use litchi_core::Position;
use litchi_core::sheet::{Cell as _, Worksheet as _};
use litchi_core::sheet::{CellValue, WorkbookTrait as _};
use litchi_opc::constants::{content_type, relationship_type};
use litchi_opc::{BlobPart, PackURI, Part as _, TargetMode};
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
    for (index, name) in names.iter().enumerate() {
        let mut sheet = builder::Worksheet::new(*name);
        sheet.set_cell(0, 0, format!("{name} identity {index}"));
        builder.add_worksheet(sheet);
    }
    builder.build().expect("worksheet fixture")
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
        .expect("chart fixture");
    let mut data = builder::Worksheet::new("Data");
    data.set_cell(0, 0, "data identity");
    builder.add_worksheet(data);
    let mut tail = builder::Worksheet::new("Tail");
    tail.set_cell(0, 0, "tail identity");
    builder.add_worksheet(tail);
    builder.build().expect("chart-first fixture")
}

/// A chart tab with no outgoing child relationship is still a complete-catalog
/// tab and may be removed.  The chart children from the source fixture become
/// ordinary orphans and are deliberately retained by the package policy.
fn plain_chart_package() -> Package {
    let source = chart_first_package();
    let chart = sheet_part(&source, "Chart");
    let mut opc = source.into_opc();
    let (relationship_ids, original_blob) = {
        let chart_sheet = opc.get_part(&chart).expect("chart sheet part");
        (
            chart_sheet
                .rels()
                .iter()
                .map(|relationship| relationship.r_id().to_owned())
                .collect::<Vec<_>>(),
            chart_sheet.blob().to_vec(),
        )
    };
    let mut stripped_blob = Vec::new();
    let mut writer = Encoder::new(&mut stripped_blob, litchi_xlsb::raw::Limits::DEFAULT);
    for item in Records::new(&original_blob, litchi_xlsb::raw::Limits::DEFAULT) {
        let record = item.expect("chart sheet record");
        if matches!(
            record.kind(),
            kind::DRAWING | kind::LEGACY_DRAWING | kind::LEGACY_DRAWING_HF
        ) {
            continue;
        }
        writer
            .write_record(record.kind(), record.payload())
            .expect("chart sheet record");
    }
    let chart_sheet = opc.get_part_mut(&chart).expect("chart sheet part");
    for relationship_id in relationship_ids {
        chart_sheet.rels_mut().remove(&relationship_id);
    }
    chart_sheet.set_blob(stripped_blob);
    Package::from_opc(opc).expect("dependency-free chart fixture")
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

fn catalog_states(package: &Package) -> Vec<State> {
    catalog(package)
        .into_iter()
        .map(|(_, _, state)| state)
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
        .unwrap_or_else(|| panic!("missing catalog tab {name}"))
}

fn sheet_part(package: &Package, name: &str) -> PackURI {
    find_entry(package, name).part
}

fn app_metadata(package: &Package) -> (usize, usize, Vec<String>) {
    let uri = PackURI::new("/docProps/app.xml").expect("app URI");
    let text = std::str::from_utf8(
        package
            .opc_package()
            .get_part(&uri)
            .expect("app part")
            .blob(),
    )
    .expect("app XML UTF-8");
    let titles_section = text
        .split_once("<TitlesOfParts>")
        .and_then(|(_, value)| value.split_once("</TitlesOfParts>").map(|(value, _)| value))
        .expect("titles container");
    let titles = titles_section
        .split("<vt:lpstr>")
        .skip(1)
        .filter_map(|value| value.split_once("</vt:lpstr>").map(|(value, _)| value))
        .map(str::to_owned)
        .collect::<Vec<_>>();
    let titles_size = text
        .split_once("<TitlesOfParts><vt:vector size=\"")
        .and_then(|(_, value)| value.split_once('"'))
        .and_then(|(value, _)| value.parse().ok())
        .expect("TitlesOfParts vector size");
    let heading_count = text
        .split_once("<HeadingPairs><vt:vector size=\"2\" baseType=\"variant\">")
        .and_then(|(_, value)| value.split_once("<vt:i4>"))
        .and_then(|(_, value)| value.split_once("</vt:i4>"))
        .and_then(|(value, _)| value.parse().ok())
        .expect("HeadingPairs sheet count");
    (titles_size, heading_count, titles)
}

fn operation_offset(wire: &[u8]) -> usize {
    let before = usize::try_from(u64::from_le_bytes(
        wire[8..16].try_into().expect("before length"),
    ))
    .expect("before length fits");
    let after = usize::try_from(u64::from_le_bytes(
        wire[16..24].try_into().expect("after length"),
    ))
    .expect("after length fits");
    32 + before + after
}

fn replace_workbook<F>(package: &Package, mut rewrite: F) -> Package
where
    F: FnMut(&[u8], &mut Vec<u8>),
{
    let mut opc = package.opc_package().clone();
    let uri = workbook_uri(package);
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

fn with_bundle_states(package: &Package, states: &[(usize, u32)]) -> Package {
    let mut position = 0usize;
    replace_workbook(package, |source, output| {
        *output = rewrite_records(source, |record_kind, payload, writer| {
            let mut payload = payload.to_vec();
            if record_kind == kind::BUNDLE_SH {
                if let Some((_, state)) = states.iter().find(|(index, _)| *index == position) {
                    payload[..4].copy_from_slice(&state.to_le_bytes());
                }
                position += 1;
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

fn related_parts(package: &Package, root: &PackURI) -> BTreeSet<String> {
    let mut pending = vec![root.clone()];
    let mut seen = BTreeSet::new();
    while let Some(uri) = pending.pop() {
        if !seen.insert(uri.as_str().to_owned()) {
            continue;
        }
        let Ok(part) = package.opc_package().get_part(&uri) else {
            continue;
        };
        for relationship in part.rels().iter() {
            if relationship.target_mode() == TargetMode::Internal {
                if let Ok(target) = relationship.target_partname() {
                    pending.push(target);
                }
            }
        }
    }
    seen
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

fn package_with_named_ranges() -> Package {
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
    builder.build().expect("named-range fixture")
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
        b"opaque-removal-owner".to_vec(),
    )))
    .unwrap();
    Package::from_opc(opc).unwrap()
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

fn package_with_retained_xml_reference() -> Package {
    let mut opc = worksheets(&["A", "B"]).into_opc();
    opc.try_add_part(Box::new(BlobPart::new(
        PackURI::new("/custom/retained.xml").unwrap(),
        "application/xml".to_owned(),
        br#"<root xmlns:s="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><s:f>A!$A$1</s:f></root>"#.to_vec(),
    )))
    .unwrap();
    Package::from_opc(opc).unwrap()
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

fn assert_removal_guard(package: Package, label: &str) {
    let before = package.to_bytes().expect("source bytes");
    let mut edit = package.edit().expect("edit");
    let error = edit
        .remove_sheet(1usize)
        .expect_err("opaque removal dependency guard");
    assert!(
        matches!(
            error,
            PackageError::Refused(_) | PackageError::InvalidFormat(_)
        ),
        "{label}: unexpected error {error:?}"
    );
    let commit = edit.commit().expect("guarded edit commit");
    assert!(commit.patch().is_noop(), "{label}: guard staged a removal");
    assert_eq!(
        commit.package().to_bytes().unwrap(),
        before,
        "{label}: source changed"
    );
}

fn remove_app(package: &Package) -> Package {
    let mut opc = package.opc_package().clone();
    let app = PackURI::new("/docProps/app.xml").expect("app URI");
    let ids = opc
        .rels()
        .iter()
        .filter(|relationship| {
            relationship.target_mode() == TargetMode::Internal
                && relationship
                    .target_partname()
                    .is_ok_and(|target| target.is_equivalent_to(&app))
        })
        .map(|relationship| relationship.r_id().to_owned())
        .collect::<Vec<_>>();
    for id in ids {
        opc.rels_mut().remove(&id);
    }
    assert!(opc.remove_part(&app), "source fixture should have app.xml");
    Package::from_opc(opc).expect("package without app.xml")
}

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
    for (id, target, mode) in worksheet_relationships {
        let workbook = opc.get_part_mut(&workbook_uri).expect("workbook part");
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

#[test]
fn removes_complete_catalog_tabs_and_updates_app_metadata() {
    {
        let (position, expected) = (2, vec!["Chart", "Data"]);
        let source = chart_first_package();
        let mut edit = source.edit().expect("edit");
        assert_eq!(edit.remove_sheet(position).expect("remove"), Some(true));
        let target = edit.commit().expect("commit").into_package();
        let expected = expected.into_iter().map(str::to_owned).collect::<Vec<_>>();
        assert_eq!(catalog_names(&target), expected);
        assert_eq!(
            app_metadata(&target),
            (expected.len(), expected.len(), expected)
        );
    }

    let source = chart_first_package();
    let before = source.to_bytes().expect("source bytes");
    let mut edit = source.edit().expect("edit");
    assert!(matches!(
        edit.remove_sheet("Data"),
        Err(PackageError::Refused(_))
    ));
    let commit = edit.commit().expect("guarded commit");
    assert!(commit.patch().is_noop());
    assert_eq!(commit.package().to_bytes().unwrap(), before);

    let source = plain_chart_package();
    let mut edit = source.edit().expect("edit");
    assert_eq!(edit.remove_sheet(0usize).expect("remove chart"), Some(true));
    let target = edit.commit().expect("commit").into_package();
    assert_eq!(catalog_names(&target), vec!["Data", "Tail"]);
    assert_eq!(
        app_metadata(&target),
        (2, 2, vec!["Data".to_owned(), "Tail".to_owned()])
    );
}

#[test]
fn selectors_are_immutable_source_positions_and_removals_coalesce_sorted() {
    let source = worksheets(&["A", "B", "C", "D"]);
    let before = source.to_bytes().expect("source bytes");
    let mut edit = source.edit().expect("edit");
    assert_eq!(edit.remove_sheet(3usize).expect("remove D"), Some(true));
    assert_eq!(edit.remove_sheet(1usize).expect("remove B"), Some(true));
    assert_eq!(edit.remove_sheet(3usize).expect("repeat D"), Some(false));
    assert_eq!(edit.remove_sheet("missing").expect("missing"), None);
    let commit = edit.commit().expect("commit");
    assert_eq!(commit.patch().operation_count(), 1);
    assert_eq!(catalog_names(commit.package()), vec!["A", "C"]);
    assert_ne!(commit.package().to_bytes().unwrap(), before);
}

#[test]
fn missing_and_repeated_only_removals_are_exact_noops() {
    let source = worksheets(&["A", "B"]);
    let before = source.to_bytes().expect("source bytes");
    let mut missing = source.edit().expect("edit");
    assert_eq!(missing.remove_sheet("missing").expect("missing"), None);
    let commit = missing.commit().expect("no-op commit");
    assert!(commit.patch().is_noop());
    assert_eq!(commit.package().to_bytes().unwrap(), before);

    let mut repeated = source.edit().expect("edit");
    assert_eq!(repeated.remove_sheet(0usize).expect("remove"), Some(true));
    assert_eq!(repeated.remove_sheet(0usize).expect("repeat"), Some(false));
    let commit = repeated.commit().expect("remove commit");
    assert_eq!(catalog_names(commit.package()), vec!["B"]);
}

#[test]
fn removal_preserves_surviving_parts_and_deletes_removed_sheet_graph() {
    let source = chart_first_package();
    let old_entries = bundle_entries(&source);
    let old_graphs = old_entries
        .iter()
        .map(|entry| (entry.name.clone(), related_parts(&source, &entry.part)))
        .collect::<HashMap<_, _>>();
    let removed = find_entry(&source, "Tail");
    assert!(
        source
            .opc_package()
            .get_part(&removed.part)
            .unwrap()
            .rels()
            .iter()
            .any(|relationship| relationship.reltype() == BINARY_INDEX_RELATIONSHIP),
        "worksheet fixture must expose its owned binary index"
    );
    let before_removed_graph = old_graphs.get("Tail").unwrap().clone();

    let mut edit = source.edit().expect("edit");
    edit.remove_sheet("Tail").expect("remove Tail");
    let target = edit.commit().expect("commit").into_package();
    let new_entries = bundle_entries(&target);
    for old in old_entries.iter().filter(|entry| entry.name != "Tail") {
        let new = new_entries
            .iter()
            .find(|entry| entry.name == old.name)
            .expect("surviving catalog tab");
        assert_eq!(
            new.rel_id, old.rel_id,
            "relationship changed for {}",
            old.name
        );
        assert_eq!(new.itab_id, old.itab_id, "itab changed for {}", old.name);
        assert_eq!(new.part, old.part, "part URI changed for {}", old.name);
        assert_eq!(
            target.opc_package().get_part(&new.part).unwrap().blob(),
            source.opc_package().get_part(&old.part).unwrap().blob(),
            "surviving part changed for {}",
            old.name
        );
    }
    assert!(target.opc_package().get_part(&removed.part).is_err());
    for child in before_removed_graph {
        if child == removed.part.as_str() {
            continue;
        }
        let shared = old_entries
            .iter()
            .filter(|entry| entry.name != "Tail")
            .any(|entry| old_graphs[&entry.name].contains(&child));
        if !shared {
            assert!(
                target
                    .opc_package()
                    .get_part(&PackURI::new(&child).unwrap())
                    .is_err(),
                "removed child survived: {child}"
            );
        }
    }
    target
        .workbook()
        .expect("target workbook")
        .worksheet_binary_index(0)
        .expect("remaining worksheet index binding")
        .expect("remaining worksheet binary index");
    assert_eq!(
        cell_by_name(&target, "Data"),
        CellValue::String("data identity".to_owned())
    );
}

#[test]
fn shared_child_relationships_refuse_removal_atomically() {
    let source = worksheets(&["A", "B"]);
    let a = sheet_part(&source, "A");
    let b = sheet_part(&source, "B");
    let shared = PackURI::new("/xl/shared-removal-child.bin").unwrap();
    let mut opc = source.into_opc();
    opc.try_add_part(Box::new(BlobPart::new(
        shared.clone(),
        "application/octet-stream".to_owned(),
        b"shared child".to_vec(),
    )))
    .unwrap();
    for (sheet, id) in [(a, "rIdSharedA"), (b, "rIdSharedB")] {
        opc.get_part_mut(&sheet)
            .unwrap()
            .rels_mut()
            .try_add_relationship(
                "urn:litchi:test:shared-child".to_owned(),
                shared.relative_ref(sheet.base_uri()),
                id.to_owned(),
                TargetMode::Internal,
            )
            .unwrap();
    }
    let source = Package::from_opc(opc).expect("shared child fixture");
    let before = source.to_bytes().unwrap();
    let mut edit = source.edit().expect("edit");
    assert!(matches!(
        edit.remove_sheet("A"),
        Err(PackageError::Refused(_))
    ));
    let commit = edit.commit().expect("guarded commit");
    assert!(commit.patch().is_noop());
    assert_eq!(commit.package().to_bytes().unwrap(), before);
    assert_eq!(
        source.opc_package().get_part(&shared).unwrap().blob(),
        b"shared child"
    );
    let b_part = sheet_part(&source, "B");
    assert!(
        source
            .opc_package()
            .get_part(&b_part)
            .unwrap()
            .rels()
            .get("rIdSharedB")
            .is_some()
    );
}

#[test]
fn chart_child_relationships_refuse_removal_atomically() {
    let source = chart_first_package();
    let before = source.to_bytes().unwrap();
    let mut edit = source.edit().unwrap();
    assert!(matches!(
        edit.remove_sheet("Chart"),
        Err(PackageError::Refused(_))
    ));
    let commit = edit.commit().unwrap();
    assert!(commit.patch().is_noop());
    assert_eq!(commit.package().to_bytes().unwrap(), before);
}

#[test]
fn dependency_free_chart_removal_retains_orphaned_chart_children() {
    let source_with_children = chart_first_package();
    let chart = find_entry(&source_with_children, "Chart");
    let chart_graph = related_parts(&source_with_children, &chart.part);
    let source = plain_chart_package();
    assert!(
        source
            .opc_package()
            .get_part(&sheet_part(&source, "Chart"))
            .unwrap()
            .rels()
            .iter()
            .next()
            .is_none(),
        "plain chart fixture must have no outgoing child relationships"
    );

    let mut edit = source.edit().unwrap();
    edit.remove_sheet("Chart").unwrap();
    let target = edit.commit().unwrap().into_package();
    assert_eq!(catalog_names(&target), vec!["Data", "Tail"]);
    for child in chart_graph {
        if child != chart.part.as_str() {
            assert!(
                target
                    .opc_package()
                    .get_part(&PackURI::new(&child).unwrap())
                    .is_ok(),
                "orphaned chart child was collected: {child}"
            );
        }
    }
}

#[test]
fn unrelated_orphan_parts_are_retained_after_removal() {
    let orphan = PackURI::new("/xl/orphan-removal-part.bin").unwrap();
    let source = add_part_with_content_type(
        worksheets(&["A", "B"]),
        orphan.as_str(),
        "application/octet-stream",
    );
    let orphan_bytes = source
        .opc_package()
        .get_part(&orphan)
        .unwrap()
        .blob()
        .to_vec();
    let mut edit = source.edit().unwrap();
    edit.remove_sheet("A").unwrap();
    let target = edit.commit().unwrap().into_package();
    assert_eq!(
        target.opc_package().get_part(&orphan).unwrap().blob(),
        orphan_bytes
    );
}

#[test]
fn removing_the_last_tab_or_last_visible_tab_is_atomic() {
    let source = worksheets(&["Only"]);
    let before = source.to_bytes().unwrap();
    let mut edit = source.edit().unwrap();
    assert!(edit.remove_sheet(0usize).is_err());
    let commit = edit.commit().unwrap();
    assert!(commit.patch().is_noop());
    assert_eq!(commit.package().to_bytes().unwrap(), before);

    let source = with_bundle_states(&worksheets(&["Visible", "Hidden"]), &[(1, 1)]);
    let before = source.to_bytes().unwrap();
    let mut edit = source.edit().unwrap();
    assert!(edit.remove_sheet("Visible").is_err());
    let commit = edit.commit().unwrap();
    assert!(commit.patch().is_noop());
    assert_eq!(commit.package().to_bytes().unwrap(), before);
    assert_eq!(
        catalog_states(commit.package()),
        vec![State::Visible, State::Hidden]
    );
}

#[test]
fn active_tab_fallback_chooses_first_visible_after_then_last_before() {
    let source = with_book_view_indices(&worksheets(&["A", "B", "C", "D"]), 0, 1);
    let mut edit = source.edit().unwrap();
    edit.remove_sheet("B").unwrap();
    let target = edit.commit().unwrap().into_package();
    assert_eq!(catalog_names(&target), vec!["A", "C", "D"]);
    assert_eq!(book_view_indices(&target), vec![(0, 1)]);
    assert_eq!(
        target.workbook().unwrap().active_catalog_position(),
        Some(1)
    );

    let source = with_book_view_indices(&worksheets(&["A", "B", "C"]), 0, 2);
    let mut edit = source.edit().unwrap();
    edit.remove_sheet("C").unwrap();
    let target = edit.commit().unwrap().into_package();
    assert_eq!(catalog_names(&target), vec!["A", "B"]);
    assert_eq!(book_view_indices(&target), vec![(0, 1)]);
    assert_eq!(
        target.workbook().unwrap().active_catalog_position(),
        Some(1)
    );

    let source = with_book_view_indices(&worksheets(&["A", "B", "C"]), 1, 1);
    let mut edit = source.edit().unwrap();
    edit.remove_sheet("A").unwrap();
    let target = edit.commit().unwrap().into_package();
    assert_eq!(catalog_names(&target), vec!["B", "C"]);
    assert_eq!(book_view_indices(&target), vec![(0, 0)]);
    assert_eq!(
        target.workbook().unwrap().active_catalog_position(),
        Some(0)
    );

    let source = with_book_view_indices(
        &with_bundle_states(&worksheets(&["A", "B", "Hidden", "D"]), &[(2, 1)]),
        0,
        1,
    );
    let mut edit = source.edit().unwrap();
    edit.remove_sheet("B").unwrap();
    let target = edit.commit().unwrap().into_package();
    assert_eq!(catalog_names(&target), vec!["A", "Hidden", "D"]);
    assert_eq!(
        catalog_states(&target),
        vec![State::Visible, State::Hidden, State::Visible]
    );
    assert_eq!(book_view_indices(&target), vec![(0, 2)]);
    assert_eq!(
        target.workbook().unwrap().active_catalog_position(),
        Some(2)
    );
}

#[test]
fn durable_tag10_is_sorted_and_replays_inverse_exactly() {
    let source = worksheets(&["A", "B", "C", "D"]);
    let before = source.to_bytes().unwrap();
    let mut edit = source.edit().unwrap();
    edit.remove_sheet(3usize).unwrap();
    edit.remove_sheet(1usize).unwrap();
    let commit = edit.commit().unwrap();
    let wire = commit.patch().to_bytes(TransferLimits::DEFAULT).unwrap();
    let offset = operation_offset(&wire);
    assert_eq!(wire[offset], 10, "sheet removal operation tag");
    let count = u64::from_le_bytes(wire[offset + 1..offset + 9].try_into().unwrap());
    assert_eq!(count, 2);
    assert_eq!(
        u64::from_le_bytes(wire[offset + 9..offset + 17].try_into().unwrap()),
        1
    );
    assert_eq!(
        u64::from_le_bytes(wire[offset + 17..offset + 25].try_into().unwrap()),
        3
    );
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
    assert_eq!(catalog_names(&replay), vec!["A", "C"]);
    assert_eq!(
        replay
            .apply(&patch.inverse())
            .map(litchi_xlsb::cell_values::structure::Commit::into_package)
            .unwrap()
            .to_bytes()
            .unwrap(),
        before
    );
}

#[test]
fn durable_tag10_rejects_malformed_count_order_duplicates_and_positions() {
    let source = worksheets(&["A", "B", "C", "D"]);
    let mut edit = source.edit().unwrap();
    edit.remove_sheet(1usize).unwrap();
    edit.remove_sheet(3usize).unwrap();
    let wire = edit
        .commit()
        .unwrap()
        .patch()
        .to_bytes(TransferLimits::DEFAULT)
        .unwrap();
    let offset = operation_offset(&wire);

    let mut huge_count = wire.clone();
    huge_count[offset + 1..offset + 9].copy_from_slice(&u64::MAX.to_le_bytes());
    assert!(
        Patch::from_bytes(
            &huge_count,
            TransferLimits::DEFAULT,
            litchi_xlsb::raw::Limits::DEFAULT
        )
        .is_err()
    );

    let mut truncated = wire.clone();
    truncated.truncate(truncated.len() - 1);
    assert!(
        Patch::from_bytes(
            &truncated,
            TransferLimits::DEFAULT,
            litchi_xlsb::raw::Limits::DEFAULT
        )
        .is_err()
    );

    for positions in [[1_u64, 1], [3, 1], [1, u64::MAX]] {
        let mut malformed = wire.clone();
        malformed[offset + 9..offset + 17].copy_from_slice(&positions[0].to_le_bytes());
        malformed[offset + 17..offset + 25].copy_from_slice(&positions[1].to_le_bytes());
        assert!(
            Patch::from_bytes(
                &malformed,
                TransferLimits::DEFAULT,
                litchi_xlsb::raw::Limits::DEFAULT
            )
            .is_err(),
            "accepted malformed positions {positions:?}"
        );
    }
}

#[test]
fn removal_merge_conflicts_on_divergent_sets_and_coalesces_identical_sets() {
    let source = worksheets(&["A", "B", "C"]);
    let mut left = source.edit().unwrap();
    left.remove_sheet(0usize).unwrap();
    let left = left.commit().unwrap();
    let mut right = source.edit().unwrap();
    right.remove_sheet(1usize).unwrap();
    let right = right.commit().unwrap();
    let merge = left.patch().merge_three_way(right.patch()).unwrap();
    assert!(merge.patch().is_none());
    assert!(!merge.conflicts().is_empty());

    let mut same = source.edit().unwrap();
    same.remove_sheet(0usize).unwrap();
    let same = same.commit().unwrap();
    let merged = left.patch().merge_three_way(same.patch()).unwrap();
    assert!(merged.patch().is_some());
    assert!(merged.conflicts().is_empty());
}

#[test]
fn mixed_catalog_removal_with_other_operation_families_is_rejected() {
    let source = worksheets(&["A", "B", "C"]);
    let mut edit = source.edit().unwrap();
    edit.remove_sheet(1usize).unwrap();
    assert!(edit.set_sheet_name(0usize, "Renamed".to_owned()).is_err());
    assert!(edit.set_sheet_state(0usize, State::Hidden).is_err());
    assert!(
        edit.reorder_sheets(&[0usize.into(), 2usize.into(), 1usize.into()])
            .is_err()
    );
    assert!(
        edit.insert_sheet(Position::new(1), "Inserted".to_owned(), State::Visible)
            .is_err()
    );
}

#[test]
fn named_range_scopes_shift_or_removed_scoped_owners_are_guarded() {
    let source = package_with_named_ranges();
    let before = source.to_bytes().unwrap();
    let mut edit = source.edit().unwrap();
    match edit.remove_sheet(0usize) {
        Ok(Some(true)) => {
            let target = edit.commit().unwrap().into_package();
            assert_eq!(
                named_range_scopes(&target),
                vec![("Scoped".to_owned(), 0), ("Global".to_owned(), u32::MAX)]
            );
        },
        Err(PackageError::Refused(_)) => {
            let commit = edit.commit().unwrap();
            assert!(commit.patch().is_noop());
            assert_eq!(commit.package().to_bytes().unwrap(), before);
        },
        other => panic!("unexpected surviving-scope result: {other:?}"),
    }

    let source = package_with_named_ranges();
    let mut edit = source.edit().unwrap();
    match edit.remove_sheet(1usize) {
        Ok(Some(true)) => {
            let target = edit.commit().unwrap().into_package();
            assert!(
                !named_range_scopes(&target)
                    .iter()
                    .any(|(_, scope)| *scope == 1)
            );
        },
        Err(PackageError::Refused(_)) => {
            assert!(edit.commit().unwrap().patch().is_noop());
        },
        other => panic!("unexpected removed-scope result: {other:?}"),
    }
}

#[test]
fn xti_endpoints_preserve_cardinality_and_shrink_or_ref_the_removed_tab() {
    let source = with_extern_sheet_fixture(1, 2);
    let before = extern_sheet_entries(&source);
    let mut edit = source.edit().unwrap();
    edit.remove_sheet(0usize).unwrap();
    let target = edit.commit().unwrap().into_package();
    let after = extern_sheet_entries(&target);
    assert_eq!(after.len(), before.len());
    assert_eq!(after[2], (0, 0, 1));
    assert_eq!(after[3], before[3]);

    let source = with_extern_sheet_fixture(1, 1);
    let before = extern_sheet_entries(&source);
    let mut edit = source.edit().unwrap();
    edit.remove_sheet(1usize).unwrap();
    let target = edit.commit().unwrap().into_package();
    let after = extern_sheet_entries(&target);
    assert_eq!(after.len(), before.len());
    assert_eq!(after[2], (0, -1, -1));
    assert_eq!(after[3], before[3]);
}

#[test]
fn removal_without_externals_does_not_invent_xti_metadata() {
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
    edit.remove_sheet(0usize).unwrap();
    let target = edit.commit().unwrap().into_package();
    assert_eq!(extern_sheet_record_count(&target), 0);
}

#[test]
fn volatile_dependencies_shift_and_inverse_exactly() {
    let source = package_with_volatile_dependencies();
    let before = source.to_bytes().unwrap();
    assert_eq!(volatile_ref_indices(&source), vec![0, 2]);
    let mut edit = source.edit().unwrap();
    edit.remove_sheet(1usize).unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(volatile_ref_indices(commit.package()), vec![0, 1]);
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
fn surviving_xml_references_to_removed_sheet_refuse_atomically() {
    let source = package_with_retained_xml_reference();
    let before = source.to_bytes().unwrap();
    let mut edit = source.edit().unwrap();
    assert!(matches!(
        edit.remove_sheet("A"),
        Err(PackageError::Refused(_))
    ));
    let commit = edit.commit().unwrap();
    assert!(commit.patch().is_noop());
    assert_eq!(commit.package().to_bytes().unwrap(), before);
}

#[test]
fn opaque_ordinal_owners_are_typed_removal_guards() {
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
        (
            "unknown workbook record",
            with_unknown_workbook_record(&worksheets(&["A", "B", "C"])),
        ),
    ];
    for (label, package) in cases {
        assert_removal_guard(package, label);
    }
}

#[test]
fn custom_workbook_uri_and_absent_app_are_preserved() {
    let source = relocate_workbook(
        &remove_app(&worksheets(&["Active", "Other"])),
        "/custom/workbook.bin",
    );
    let before = source.to_bytes().unwrap();
    let mut edit = source.edit().unwrap();
    edit.remove_sheet("Other").unwrap();
    let target = edit.commit().unwrap().into_package();
    assert_eq!(workbook_uri(&target).as_str(), "/custom/workbook.bin");
    assert!(
        target
            .opc_package()
            .get_part(&PackURI::new("/docProps/app.xml").unwrap())
            .is_err()
    );
    let workbook = target
        .opc_package()
        .get_part(&workbook_uri(&target))
        .unwrap();
    for relationship in workbook.rels().iter() {
        if relationship.target_mode() == TargetMode::Internal {
            let part = relationship.target_partname().unwrap();
            assert!(target.opc_package().contains_part(&part));
        }
    }
    assert_ne!(target.to_bytes().unwrap(), before);
}

#[test]
fn strict_worksheet_relationship_flavor_survives_removal() {
    let source = strict_worksheet_relationships(&chart_first_package());
    let data = find_entry(&source, "Tail");
    let mut edit = source.edit().unwrap();
    edit.remove_sheet("Tail").unwrap();
    let target = edit.commit().unwrap().into_package();
    assert!(target.opc_package().get_part(&data.part).is_err());
    let tail = find_entry(&target, "Data");
    let relationship = target
        .opc_package()
        .get_part(&workbook_uri(&target))
        .unwrap()
        .rels()
        .get(&tail.rel_id)
        .unwrap();
    assert_eq!(relationship.reltype(), relationship_type::STRICT_WORKSHEET);
}

fn selected_flag(package: &Package, name: &str) -> bool {
    let uri = sheet_part(package, name);
    let part = package.opc_package().get_part(&uri).unwrap();
    Records::new(part.blob(), litchi_xlsb::raw::Limits::DEFAULT)
        .find_map(|item| {
            let record = item.unwrap();
            let mask = match record.kind() {
                kind::BEGIN_WS_VIEW => 0x0040,
                kind::BEGIN_CS_VIEW => 0x0001,
                _ => return None,
            };
            Some(u16::from_le_bytes(record.payload()[..2].try_into().unwrap()) & mask != 0)
        })
        .unwrap()
}

#[test]
fn active_removal_selects_successor_and_already_selected_successor_is_valid() {
    for preselected in [false, true] {
        let source = worksheets(&["A", "B", "C"]);
        let uri = sheet_part(&source, "B");
        let mut opc = source.into_opc();
        let blob = opc.get_part(&uri).unwrap().blob().to_vec();
        let blob = rewrite_records(&blob, |record_kind, payload, writer| {
            let mut payload = payload.to_vec();
            if record_kind == kind::BEGIN_WS_VIEW {
                let flags = u16::from_le_bytes(payload[..2].try_into().unwrap());
                let flags = if preselected {
                    flags | 0x40
                } else {
                    flags & !0x40
                };
                payload[..2].copy_from_slice(&flags.to_le_bytes());
            }
            writer.write_record(record_kind, &payload).unwrap();
        });
        opc.replace_part_payloads(vec![(uri, blob)]).unwrap();
        let source = Package::from_opc(opc).unwrap();
        let mut edit = source.edit().unwrap();
        edit.remove_sheet("A").unwrap();
        let target = edit.commit().unwrap().into_package();
        assert!(selected_flag(&target, "B"));
        assert_eq!(
            target.workbook().unwrap().active_catalog_position(),
            Some(0)
        );
    }
}

#[test]
fn absent_workbook_views_preserve_optional_absence_and_refuse_hidden_default() {
    for hidden_first in [false, true] {
        let source = worksheets(&["A", "B", "C"]);
        let source = if hidden_first {
            with_bundle_states(&source, &[(1, 1)])
        } else {
            source
        };
        let source = replace_workbook(&source, |blob, output| {
            *output = rewrite_records(blob, |record_kind, payload, writer| {
                if !matches!(
                    record_kind,
                    kind::BEGIN_BOOK_VIEWS | kind::BOOK_VIEW | kind::END_BOOK_VIEWS
                ) {
                    writer.write_record(record_kind, payload).unwrap();
                }
            });
        });
        let before = source.to_bytes().unwrap();
        let mut edit = source.edit().unwrap();
        if hidden_first {
            assert!(matches!(
                edit.remove_sheet("A"),
                Err(PackageError::Refused(_))
            ));
            assert_eq!(edit.commit().unwrap().package().to_bytes().unwrap(), before);
        } else {
            edit.remove_sheet("A").unwrap();
            let target = edit.commit().unwrap().into_package();
            assert!(book_view_indices(&target).is_empty());
            assert_eq!(catalog_names(&target), ["B", "C"]);
        }
    }
}

#[test]
fn binary_hyperlink_locations_cannot_keep_removed_sheet_names() {
    use litchi_xlsb::hyperlinks::Hyperlink;
    for (name, location) in [
        ("Gone", "Gone!A1"),
        ("A!B", "'A!B'!A1"),
        ("O'Brien", "#'O''Brien'!A1"),
    ] {
        let mut builder = Builder::new();
        builder.add_worksheet(builder::Worksheet::new(name));
        let mut survivor = builder::Worksheet::new("Keep");
        survivor.add_hyperlink(Hyperlink::new_internal(0, 0, 0, 0, location.to_owned()));
        builder.add_worksheet(survivor);
        let source = builder.build().unwrap();
        let before = source.to_bytes().unwrap();
        let mut edit = source.edit().unwrap();
        assert!(matches!(
            edit.remove_sheet(name),
            Err(PackageError::Refused(_))
        ));
        assert_eq!(edit.commit().unwrap().package().to_bytes().unwrap(), before);
    }
}

#[test]
fn opaque_pivot_cache_and_future_worksheet_carriers_are_guarded() {
    let source = add_part_with_content_type(
        worksheets(&["A", "B"]),
        "/xl/pivotCache/pivotCacheDefinition1.bin",
        "application/vnd.ms-excel.pivotCacheDefinition",
    );
    assert_removal_guard(source, "pivot cache");
    for (begin, end) in [
        (kind::FRT_BEGIN, kind::FRT_END),
        (kind::AC_BEGIN, kind::AC_END),
    ] {
        let source = worksheets(&["A", "B"]);
        let uri = sheet_part(&source, "A");
        let mut opc = source.into_opc();
        let blob = opc.get_part(&uri).unwrap().blob().to_vec();
        let blob = rewrite_records(&blob, |record_kind, payload, writer| {
            if record_kind == kind::END_SHEET {
                writer.write_record(begin, &[]).unwrap();
                writer.write_record(end, &[]).unwrap();
            }
            writer.write_record(record_kind, payload).unwrap();
        });
        opc.replace_part_payloads(vec![(uri, blob)]).unwrap();
        assert_removal_guard(Package::from_opc(opc).unwrap(), "future worksheet carrier");
    }
}

#[test]
fn global_defined_name_cannot_lose_its_only_local_target() {
    let mut builder = Builder::new();
    for name in ["A", "B", "C"] {
        builder.add_worksheet(builder::Worksheet::new(name));
    }
    builder.add_named_range(
        NamedRange::new("Target".to_owned(), None)
            .with_formula(litchi_xlsb::named_ranges::area3d_formula(1, 0, 0, 0, 0).unwrap()),
    );
    let source = builder.build().unwrap();
    let before = source.to_bytes().unwrap();
    let mut edit = source.edit().unwrap();
    assert!(matches!(
        edit.remove_sheet("B"),
        Err(PackageError::Refused(_))
    ));
    assert_eq!(edit.commit().unwrap().package().to_bytes().unwrap(), before);
    let mut safe = source.edit().unwrap();
    safe.remove_sheet("C").unwrap();
    assert_eq!(catalog_names(safe.commit().unwrap().package()), ["A", "B"]);
}

#[test]
fn live_formula_references_receive_typed_removal_refusal() {
    use litchi_xlsb::formula::ParsedFormula;
    let mut builder = Builder::new();
    builder.add_worksheet(builder::Worksheet::new("A"));
    let mut keep = builder::Worksheet::new("Keep");
    keep.set_cell_formula_binary(
        0,
        0,
        CellValue::Float(1.0),
        ParsedFormula {
            rgce: litchi_xlsb::named_ranges::area3d_formula(0, 0, 0, 0, 0).unwrap(),
            rgcb: Vec::new(),
        },
        false,
        0,
    );
    builder.add_worksheet(keep);
    builder.add_worksheet(builder::Worksheet::new("C"));
    let source = builder.build().unwrap();
    let before = source.to_bytes().unwrap();
    let mut edit = source.edit().unwrap();
    assert!(matches!(
        edit.remove_sheet("A"),
        Err(PackageError::Refused(_))
    ));
    assert_eq!(edit.commit().unwrap().package().to_bytes().unwrap(), before);
    let mut safe = source.edit().unwrap();
    safe.remove_sheet("C").unwrap();
    assert_eq!(
        catalog_names(safe.commit().unwrap().package()),
        ["A", "Keep"]
    );
}
