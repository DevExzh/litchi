#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "structural-order fixtures deliberately fail at the assertion site"
)]

//! Native complete-catalog sheet-order transactions.
//!
//! These tests keep the workbook's physical relationship identities visible:
//! a reorder changes the `BrtBundleSh` display order, while cells, chart tabs,
//! views, defined-name scopes, and XTI entries continue to point at the same
//! authored tabs.  The raw helpers below create the small record variants
//! needed to exercise dependency boundaries that the writer does not expose
//! as a public authoring surface.

use litchi_core::sheet::{Cell as _, Worksheet as _};
use litchi_core::sheet::{CellValue, WorkbookTrait as _};
use litchi_opc::PackURI;
use litchi_xlsb::Package;
use litchi_xlsb::builder::chart::{Anchor, Chart};
use litchi_xlsb::builder::{self, Builder};
use litchi_xlsb::cell_values::TransferLimits;
use litchi_xlsb::cell_values::structure::Patch;
use litchi_xlsb::named_ranges::Definition as NamedRange;
use litchi_xlsb::package::PackageError;
use litchi_xlsb::raw::{Encoder, Records, kind};
use litchi_xlsb::sheet::{Kind, Selector};

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

fn worksheets(names: &[&str]) -> Package {
    let mut builder = Builder::new();
    for (index, name) in names.iter().enumerate() {
        let mut sheet = builder::Worksheet::new(*name);
        sheet.set_cell(0, 0, format!("{name} identity {index}"));
        builder.add_worksheet(sheet);
    }
    builder.build().expect("worksheet package")
}

fn order_by_names<'a>(names: &[&'a str]) -> Vec<Selector<'a>> {
    names.iter().copied().map(Selector::from).collect()
}

fn catalog(package: &Package) -> Vec<(String, Kind)> {
    package
        .workbook()
        .expect("workbook projection")
        .sheets()
        .map(|tab| (tab.name().to_owned(), tab.kind()))
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

fn workbook_blob(package: &Package) -> Vec<u8> {
    let uri = PackURI::new("/xl/workbook.bin").expect("workbook URI");
    package
        .opc_package()
        .get_part(&uri)
        .expect("workbook part")
        .blob()
        .to_vec()
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
    F: FnMut(litchi_xlsb::raw::Kind, &[u8], usize, &mut Encoder<&mut Vec<u8>>),
{
    let mut output = Vec::new();
    let mut ordinal = 0usize;
    let mut writer = Encoder::new(&mut output, litchi_xlsb::raw::Limits::DEFAULT);
    for item in Records::new(source, litchi_xlsb::raw::Limits::DEFAULT) {
        let record = item.expect("workbook record");
        rewrite(record.kind(), record.payload(), ordinal, &mut writer);
        ordinal = ordinal.saturating_add(1);
    }
    output
}

fn with_book_view_indices(package: &Package, first: u32, active: u32) -> Package {
    replace_workbook(package, |source, output| {
        *output = rewrite_records(source, |record_kind, payload, _ordinal, writer| {
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
    let blob = workbook_blob(package);
    Records::new(&blob, litchi_xlsb::raw::Limits::DEFAULT)
        .filter_map(|item| {
            let record = item.expect("workbook record");
            if record.kind() == kind::BOOK_VIEW {
                assert_eq!(record.payload().len(), 29);
                Some((
                    u32::from_le_bytes(record.payload()[20..24].try_into().unwrap()),
                    u32::from_le_bytes(record.payload()[24..28].try_into().unwrap()),
                ))
            } else {
                None
            }
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
        .expect("add external link");
    let package = builder.build().expect("external fixture");
    replace_workbook(&package, |source, output| {
        *output = rewrite_records(source, |record_kind, payload, _ordinal, writer| {
            let mut payload = payload.to_vec();
            if record_kind == kind::EXTERN_SHEET {
                assert!(payload.len() >= 4 + 3 * 12);
                // XTI 0 and 1 are the workbook and #REF sentinels. The first
                // self-workbook entry is XTI 2, immediately after them.
                let offset = 4 + 2 * 12;
                payload[offset..offset + 4].copy_from_slice(&0_u32.to_le_bytes());
                payload[offset + 4..offset + 8].copy_from_slice(&first.to_le_bytes());
                payload[offset + 8..offset + 12].copy_from_slice(&last.to_le_bytes());
                // Keep one opaque external-workbook XTI between the internal
                // range and the final self entry. Its stable XTI position and
                // endpoints must survive unchanged.
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
    let blob = workbook_blob(package);
    Records::new(&blob, litchi_xlsb::raw::Limits::DEFAULT)
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

fn with_unknown_between_bundle_sheets(package: &Package) -> Package {
    replace_workbook(package, |source, output| {
        let unknown = litchi_xlsb::raw::Kind::new(0x3ffe).expect("unknown record kind");
        let mut inserted = false;
        let mut bundles = 0usize;
        let mut writer = Encoder::new(output, litchi_xlsb::raw::Limits::DEFAULT);
        for item in Records::new(source, litchi_xlsb::raw::Limits::DEFAULT) {
            let record = item.expect("workbook record");
            if record.kind() == kind::BUNDLE_SH {
                if bundles == 1 && !inserted {
                    writer
                        .write_record(unknown, &[0xde, 0xad, 0xbe, 0xef])
                        .expect("unknown record");
                    inserted = true;
                }
                bundles = bundles.saturating_add(1);
            }
            writer
                .write_record(record.kind(), record.payload())
                .expect("workbook record");
        }
        assert!(inserted, "fixture must contain at least two bundle records");
    })
}

fn named_range_scopes(package: &Package) -> Vec<(String, u32)> {
    let blob = workbook_blob(package);
    Records::new(&blob, litchi_xlsb::raw::Limits::DEFAULT)
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

#[test]
fn mixed_chart_first_reorder_preserves_tab_kind_and_cell_identity() {
    let source = chart_first_package();
    let mut edit = source.edit().expect("edit");
    assert!(
        edit.reorder_sheets(&order_by_names(&["Tail", "Chart", "Data"]))
            .expect("reorder")
    );
    let commit = edit.commit().expect("reorder commit");
    let target = commit.package();

    assert_eq!(
        catalog(target),
        vec![
            ("Tail".to_owned(), Kind::Worksheet),
            ("Chart".to_owned(), Kind::Chart),
            ("Data".to_owned(), Kind::Worksheet),
        ]
    );
    assert_eq!(
        cell_by_name(target, "Data"),
        CellValue::String("data identity".to_owned())
    );
    assert_eq!(
        cell_by_name(target, "Tail"),
        CellValue::String("tail identity".to_owned())
    );
}

#[test]
fn repeated_order_staging_replaces_the_previous_permutation_and_can_cancel_exactly() {
    let source = worksheets(&["A", "B", "C"]);
    let before = source.to_bytes().expect("source bytes");
    let mut identity = source.edit().expect("identity edit");
    assert!(
        !identity
            .reorder_sheets(&order_by_names(&["A", "B", "C"]))
            .expect("fresh identity")
    );
    assert!(
        identity
            .commit()
            .expect("identity commit")
            .patch()
            .is_noop()
    );

    let mut edit = source.edit().expect("edit");
    assert!(
        edit.reorder_sheets(&order_by_names(&["C", "A", "B"]))
            .expect("first order")
    );
    assert!(
        edit.reorder_sheets(&order_by_names(&["B", "C", "A"]))
            .expect("second order")
    );
    assert!(
        edit.reorder_sheets(&order_by_names(&["A", "B", "C"]))
            .expect("identity cancellation")
    );
    let commit = edit.commit().expect("cancelled reorder commit");
    assert_eq!(commit.patch().operation_count(), 1);
    assert!(commit.patch().is_noop());
    assert_eq!(
        commit.package().to_bytes().expect("candidate bytes"),
        before
    );
}

#[test]
fn reorder_rejects_incomplete_or_duplicate_source_selectors_without_staging() {
    let source = worksheets(&["A", "B", "C"]);
    let mut edit = source.edit().expect("edit");
    assert!(edit.reorder_sheets(&order_by_names(&["A", "B"])).is_err());
    assert!(
        edit.reorder_sheets(&order_by_names(&["A", "A", "C"]))
            .is_err()
    );
    let commit = edit.commit().expect("empty edit commit");
    assert!(commit.patch().is_noop());
    assert_eq!(
        commit.package().to_bytes().unwrap(),
        source.to_bytes().unwrap()
    );
}

#[test]
fn durable_tag8_replay_and_inverse_preserve_the_complete_catalog() {
    let source = chart_first_package();
    let before = source.to_bytes().expect("source bytes");
    let mut edit = source.edit().expect("edit");
    edit.reorder_sheets(&order_by_names(&["Tail", "Chart", "Data"]))
        .expect("reorder");
    let commit = edit.commit().expect("reorder commit");
    let wire = commit
        .patch()
        .to_bytes(TransferLimits::DEFAULT)
        .expect("durable patch");
    assert_eq!(wire[operation_offset(&wire)], 8, "reorder operation tag");
    let patch = Patch::from_bytes(
        &wire,
        TransferLimits::DEFAULT,
        litchi_xlsb::raw::Limits::DEFAULT,
    )
    .expect("decode tag8");
    let replay = source
        .apply(&patch)
        .map(litchi_xlsb::cell_values::structure::Commit::into_package)
        .expect("replay tag8");
    assert_eq!(
        catalog(&replay),
        vec![
            ("Tail".to_owned(), Kind::Worksheet),
            ("Chart".to_owned(), Kind::Chart),
            ("Data".to_owned(), Kind::Worksheet),
        ]
    );
    assert_eq!(
        replay
            .apply(&patch.inverse())
            .map(litchi_xlsb::cell_values::structure::Commit::into_package)
            .expect("inverse")
            .to_bytes()
            .expect("inverse bytes"),
        before
    );
}

#[test]
fn malformed_tag8_count_is_rejected_before_allocation_or_replay() {
    let source = worksheets(&["A", "B", "C"]);
    let mut edit = source.edit().expect("edit");
    edit.reorder_sheets(&order_by_names(&["C", "A", "B"]))
        .expect("reorder");
    let commit = edit.commit().expect("reorder commit");
    let wire = commit
        .patch()
        .to_bytes(TransferLimits::DEFAULT)
        .expect("durable patch");
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

    let truncated = wire[..wire.len().saturating_sub(1)].to_vec();
    assert!(
        Patch::from_bytes(
            &truncated,
            TransferLimits::DEFAULT,
            litchi_xlsb::raw::Limits::DEFAULT
        )
        .is_err()
    );
}

#[test]
fn duplicate_tag8_operations_are_rejected_as_source_relative_history() {
    let source = worksheets(&["A", "B", "C"]);
    let mut edit = source.edit().expect("edit");
    edit.reorder_sheets(&order_by_names(&["C", "A", "B"]))
        .expect("reorder");
    let commit = edit.commit().expect("reorder commit");
    let one = commit
        .patch()
        .to_bytes(TransferLimits::DEFAULT)
        .expect("durable patch");
    let offset = operation_offset(&one);
    let mut duplicate = one[..offset].to_vec();
    duplicate[24..32].copy_from_slice(&2_u64.to_le_bytes());
    duplicate.extend_from_slice(&one[offset..]);
    duplicate.extend_from_slice(&one[offset..]);
    assert!(
        Patch::from_bytes(
            &duplicate,
            TransferLimits::DEFAULT,
            litchi_xlsb::raw::Limits::DEFAULT
        )
        .is_err()
    );
}

#[test]
fn divergent_reorders_are_a_single_merge_conflict() {
    let source = worksheets(&["A", "B", "C"]);
    let mut left = source.edit().expect("left edit");
    left.reorder_sheets(&order_by_names(&["B", "C", "A"]))
        .expect("left reorder");
    let left = left.commit().expect("left commit");
    let mut right = source.edit().expect("right edit");
    right
        .reorder_sheets(&order_by_names(&["C", "A", "B"]))
        .expect("right reorder");
    let right = right.commit().expect("right commit");
    let merge = left
        .patch()
        .merge_three_way(right.patch())
        .expect("merge result");
    assert!(merge.patch().is_none());
    assert_eq!(merge.conflicts().len(), 1);
    assert_eq!(merge.conflicts()[0].sheet_index(), usize::MAX);
}

#[test]
fn reorder_and_metadata_edits_conflict_at_the_workbook_barrier() {
    let source = worksheets(&["A", "B", "C"]);
    let mut reordered = source.edit().expect("reorder edit");
    reordered
        .reorder_sheets(&order_by_names(&["C", "A", "B"]))
        .expect("reorder");
    let reordered = reordered.commit().expect("reorder commit");

    let mut renamed = source.edit().expect("rename edit");
    renamed
        .set_sheet_name(0, "Renamed".to_owned())
        .expect("rename")
        .expect("resolved sheet");
    let renamed = renamed.commit().expect("rename commit");

    let merge = reordered
        .patch()
        .merge_three_way(renamed.patch())
        .expect("merge result");
    assert!(merge.patch().is_none());
    assert_eq!(merge.conflicts().len(), 1);
}

#[test]
fn book_view_first_and_active_indices_follow_the_physical_tab() {
    let source = with_book_view_indices(&worksheets(&["A", "B", "C"]), 1, 2);
    assert_eq!(book_view_indices(&source), vec![(1, 2)]);
    let mut edit = source.edit().expect("edit");
    edit.reorder_sheets(&order_by_names(&["C", "A", "B"]))
        .expect("reorder");
    let target = edit.commit().expect("reorder commit").into_package();
    // Old positions A=0, B=1, C=2 map to new positions A=1, B=2, C=0.
    assert_eq!(book_view_indices(&target), vec![(2, 0)]);
    assert_eq!(
        target
            .workbook()
            .expect("workbook")
            .active_catalog_position(),
        Some(0)
    );
}

#[test]
fn defined_name_scope_remaps_but_global_scope_stays_global() {
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
            0,
        ]),
    );
    let source = builder.build().expect("named range package");
    assert_eq!(
        named_range_scopes(&source),
        vec![("Scoped".to_owned(), 1), ("Global".to_owned(), u32::MAX)]
    );
    let mut edit = source.edit().expect("edit");
    edit.reorder_sheets(&order_by_names(&["C", "A", "B"]))
        .expect("reorder");
    let target = edit.commit().expect("reorder commit").into_package();
    assert_eq!(
        named_range_scopes(&target),
        vec![("Scoped".to_owned(), 2), ("Global".to_owned(), u32::MAX)]
    );
}

#[test]
fn extern_sheet_xti_order_and_sentinels_are_stable_while_internal_ranges_remap() {
    let source = with_extern_sheet_fixture(0, 1);
    let original = extern_sheet_entries(&source);
    let mut edit = source.edit().expect("edit");
    edit.reorder_sheets(&order_by_names(&["C", "A", "B"]))
        .expect("reorder");
    let target = edit.commit().expect("reorder commit").into_package();
    let entries = extern_sheet_entries(&target);
    assert_eq!(entries.len(), original.len());
    assert_eq!(entries[0], (0, -2, -2));
    assert_eq!(entries[1], (0, -1, -1));
    assert_eq!(entries[2], (0, 1, 2));
    assert_eq!(entries[3], (1, 7, 8));
    assert_eq!(entries[4], (0, 0, 0));
    assert_eq!(original[3], (1, 7, 8));
}

#[test]
fn noncontiguous_internal_three_d_range_refuses_atomically() {
    let source = with_extern_sheet_fixture(0, 1);
    let before = source.to_bytes().expect("source bytes");
    let mut edit = source.edit().expect("edit");
    edit.reorder_sheets(&order_by_names(&["B", "C", "A"]))
        .expect("stage reorder");
    let error = edit.commit().expect_err("noncontiguous range must refuse");
    assert!(matches!(error, PackageError::Refused(_)));
    assert_eq!(source.to_bytes().expect("source remains"), before);
}

#[test]
fn unknown_record_between_bundle_tabs_is_a_safe_reorder_barrier() {
    let source = with_unknown_between_bundle_sheets(&worksheets(&["A", "B", "C"]));
    let before = source.to_bytes().expect("source bytes");
    let mut edit = source.edit().expect("edit");
    edit.reorder_sheets(&order_by_names(&["C", "A", "B"]))
        .expect("stage reorder");
    assert!(edit.commit().is_err());
    assert_eq!(source.to_bytes().expect("source remains"), before);
}

#[test]
fn unknown_workbook_record_outside_catalog_refuses_changed_order_only() {
    let source = replace_workbook(&worksheets(&["A", "B"]), |source, output| {
        let mut writer = Encoder::new(output, litchi_xlsb::raw::Limits::DEFAULT);
        for item in Records::new(source, litchi_xlsb::raw::Limits::DEFAULT) {
            let record = item.expect("record");
            if record.kind() == kind::END_BOOK {
                writer
                    .write_record(litchi_xlsb::raw::Kind::new(0x3ffe).unwrap(), &[1, 2, 3])
                    .unwrap();
            }
            writer
                .write_record(record.kind(), record.payload())
                .unwrap();
        }
    });
    let bytes = source.to_bytes().unwrap();
    let mut edit = source.edit().unwrap();
    edit.reorder_sheets(&order_by_names(&["B", "A"])).unwrap();
    assert!(matches!(edit.commit(), Err(PackageError::Refused(_))));
    let mut edit = source.edit().unwrap();
    assert!(!edit.reorder_sheets(&order_by_names(&["A", "B"])).unwrap());
    assert_eq!(edit.commit().unwrap().package().to_bytes().unwrap(), bytes);
}

#[test]
fn unresolved_xti_supporting_link_refuses_changed_order() {
    let source = replace_workbook(&with_extern_sheet_fixture(0, 1), |source, output| {
        *output = rewrite_records(source, |kind, payload, _, writer| {
            let mut payload = payload.to_vec();
            if kind == kind::EXTERN_SHEET {
                payload[4 + 3 * 12..8 + 3 * 12].copy_from_slice(&99_u32.to_le_bytes());
            }
            writer.write_record(kind, &payload).unwrap();
        });
    });
    let before = source.to_bytes().unwrap();
    let mut edit = source.edit().unwrap();
    edit.reorder_sheets(&order_by_names(&["C", "A", "B"]))
        .unwrap();
    assert!(matches!(
        edit.commit(),
        Err(PackageError::InvalidFormula(_))
    ));
    assert_eq!(source.to_bytes().unwrap(), before);
}
