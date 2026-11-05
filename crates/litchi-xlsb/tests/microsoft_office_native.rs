//! Native Microsoft Excel specimens and independent recorded expectations.

use litchi_core::sheet::{Cell as _, CellValue, WorkbookTrait};
use litchi_xlsb::Workbook;
use std::io::Cursor;

const INTERACTION: &[u8] =
    include_bytes!("../../../test-data/microsoft-office/excel/excel-interaction-print.xlsb");

#[test]
fn native_excel_validation_collection_exposes_all_seven_rules_and_cached_totals() {
    let workbook = Workbook::new(Cursor::new(INTERACTION)).unwrap();
    let sheet = workbook.worksheet(0).unwrap();
    let rules = sheet.data_validations();
    assert_eq!(rules.len(), 7);
    for (range, kind) in [
        ("B2:B5", 1),
        ("C2:C5", 3),
        ("D2:D5", 4),
        ("E2:E5", 2),
        ("G2:G4", 5),
        ("H2:H4", 6),
        ("I2:I4", 7),
    ] {
        let rule = rules.iter().find(|rule| rule.cell_ranges == range).unwrap();
        assert_eq!(rule.validation_type, kind, "{range}");
    }
    let whole = rules
        .iter()
        .find(|rule| rule.cell_ranges == "B2:B5")
        .unwrap();
    assert_eq!(whole.formula1.as_deref(), Some("1"));
    assert_eq!(whole.formula2.as_deref(), Some("10"));
    assert_eq!(whole.input_title.as_deref(), Some("Quantity input"));
    for (column, expected) in [(1, 14.0), (4, 95.0)] {
        let cell = sheet.get_cell(5, column).unwrap();
        let CellValue::Formula {
            cached_value: Some(value),
            ..
        } = cell.value()
        else {
            panic!("expected cached formula")
        };
        assert_eq!(**value, CellValue::Float(expected));
    }
}

fn members(bytes: &[u8]) -> std::collections::BTreeMap<String, Vec<u8>> {
    let reader = litchi_opc::phys_pkg::OwnedPhysPkgReader::from_bytes(
        bytes.to_vec(),
        litchi_opc::SourceLimits::default(),
    )
    .unwrap();
    reader
        .member_names()
        .unwrap()
        .into_iter()
        .map(|name| {
            let bytes = reader.read_member(&name).unwrap();
            (name, bytes)
        })
        .collect()
}

fn alternate_blocks(sheet: &[u8]) -> Vec<Vec<u8>> {
    use litchi_xlsb::raw::{Records, kind};
    let mut records = Records::new(sheet, litchi_xlsb::raw::Limits::DEFAULT);
    let mut blocks = Vec::new();
    let mut start = None;
    while let Some(record) = records.next() {
        let record = record.unwrap();
        if record.kind() == kind::AC_BEGIN {
            assert!(start.replace(record.offset()).is_none());
        } else if record.kind() == kind::AC_END {
            blocks.push(sheet[start.take().unwrap()..records.offset()].to_vec());
        }
    }
    assert!(start.is_none());
    blocks
}

#[test]
fn native_excel_validation_edit_retains_alternate_metadata_and_other_members() {
    let source = litchi_xlsb::Package::from_bytes(INTERACTION).unwrap();
    let before = source.to_bytes().unwrap();
    let mut edit = source.edit_data_validations().unwrap();
    let original = edit.validations(0).unwrap();
    assert_eq!(original.len(), 7);
    assert!(!edit.set(0, 0, original[0].clone()).unwrap());
    let mut changed = original[0].clone();
    changed.error_title = Some("Native validation edit".into());
    assert!(edit.set(0, 0, changed.clone()).unwrap());
    let commit = edit.commit().unwrap();
    let after = commit.package().to_bytes().unwrap();
    let reopened = litchi_xlsb::Package::from_bytes(&after).unwrap();
    let rules = reopened
        .edit_data_validations()
        .unwrap()
        .validations(0)
        .unwrap();
    assert_eq!(rules[0], changed);
    assert_eq!(&rules[1..], &original[1..]);
    let before_parts = members(&before);
    let after_parts = members(&after);
    assert_eq!(
        before_parts.keys().collect::<Vec<_>>(),
        after_parts.keys().collect::<Vec<_>>()
    );
    const SHEET: &str = "xl/worksheets/sheet1.bin";
    for (name, data) in &before_parts {
        if name != SHEET {
            assert_eq!(after_parts.get(name), Some(data), "{name}");
        }
    }
    let blocks = alternate_blocks(&before_parts[SHEET]);
    assert!(blocks.len() >= 7);
    assert_eq!(alternate_blocks(&after_parts[SHEET]), blocks);
    assert_eq!(source.to_bytes().unwrap(), before);
    let undone = commit
        .package()
        .apply_data_validations(&commit.patch().inverse())
        .unwrap()
        .into_package();
    assert_eq!(undone.to_bytes().unwrap(), before);
}

#[test]
fn native_excel_validation_add_and_remove_preserve_existing_rules() {
    let source = litchi_xlsb::Package::from_bytes(INTERACTION).unwrap();
    let before = source.to_bytes().unwrap();
    let mut edit = source.edit_data_validations().unwrap();
    let original = edit.validations(0).unwrap();
    let mut added = original[0].clone();
    added.cell_ranges = "B10".into();
    assert_eq!(edit.add(0, added.clone()).unwrap(), 7);
    let after = edit.commit().unwrap().into_package();
    let mut edit = after.edit_data_validations().unwrap();
    let rules = edit.validations(0).unwrap();
    assert_eq!(&rules[..7], original);
    assert_eq!(rules[7], added);
    edit.remove(0, 7).unwrap();
    let restored = edit.commit().unwrap().into_package();
    assert_eq!(members(&restored.to_bytes().unwrap()), members(&before));
}

fn rewrite_sheet(change: impl FnOnce(&mut Vec<(litchi_xlsb::raw::Kind, Vec<u8>)>)) -> Vec<u8> {
    rewrite_sheet_of(INTERACTION, change)
}

fn rewrite_sheet_of(
    source: &[u8],
    change: impl FnOnce(&mut Vec<(litchi_xlsb::raw::Kind, Vec<u8>)>),
) -> Vec<u8> {
    rewrite_sheet_part(source, "/xl/worksheets/sheet1.bin", change)
}

fn rewrite_sheet_part(
    source: &[u8],
    part: &str,
    change: impl FnOnce(&mut Vec<(litchi_xlsb::raw::Kind, Vec<u8>)>),
) -> Vec<u8> {
    use litchi_xlsb::raw::{Encoder, Limits, Records};
    let mut package =
        litchi_opc::OpcPackage::from_bytes(source, litchi_opc::SourceLimits::default()).unwrap();
    let uri = litchi_opc::PackURI::new(part).unwrap();
    let mut records = Records::new(package.get_part(&uri).unwrap().blob(), Limits::DEFAULT)
        .map(|record| {
            let record = record.unwrap();
            (record.kind(), record.payload().to_vec())
        })
        .collect::<Vec<_>>();
    change(&mut records);
    let mut data = Vec::new();
    let mut encoder = Encoder::new(&mut data, Limits::DEFAULT);
    for (kind, payload) in records {
        encoder.write_record(kind, &payload).unwrap();
    }
    package.get_part_mut(&uri).unwrap().set_blob(data);
    let mut bytes = Vec::new();
    package.write_to(&mut bytes).unwrap();
    bytes
}

#[test]
fn native_validation_stream_refuses_broken_alternate_framing_and_counts() {
    use litchi_xlsb::raw::kind;
    for case in 0..6 {
        let bytes = rewrite_sheet(|records| {
            let begin = records
                .iter()
                .position(|r| r.0 == kind::BEGIN_D_VALS)
                .unwrap();
            let ac = begin + 1;
            assert_eq!(records[ac].0, kind::AC_BEGIN);
            let end = ac
                + records[ac..]
                    .iter()
                    .position(|r| r.0 == kind::AC_END)
                    .unwrap();
            match case {
                0 => {
                    records.remove(end);
                },
                1 => records[end].1.push(0),
                2 => {
                    records[ac].1.pop();
                },
                3 => {
                    let block = records[ac..=end].to_vec();
                    let boundary = records
                        .iter()
                        .position(|r| r.0 == kind::END_D_VALS)
                        .unwrap();
                    records.splice(boundary..boundary, block);
                },
                4 => {
                    let mut blocks = records[ac..=end].to_vec();
                    blocks.extend_from_within(..);
                    records.splice(ac..ac, blocks);
                },
                5 => records[begin].1[14..18].copy_from_slice(&8u32.to_le_bytes()),
                _ => unreachable!(),
            }
        });
        assert!(
            Workbook::new(Cursor::new(&bytes))
                .and_then(|book| book.worksheet(0))
                .is_err(),
            "case {case}"
        );
        assert!(
            litchi_xlsb::Package::from_bytes(&bytes)
                .and_then(|package| package.edit_data_validations())
                .and_then(|edit| edit.validations(0))
                .is_err(),
            "case {case}"
        );
    }
}

#[test]
fn validation_list_alternate_is_retained_or_replaced_without_losing_opaque_prefixes() {
    use litchi_xlsb::raw::kind;
    // Derive a list-prefix variant in memory; the native file stays unchanged.
    let bytes = rewrite_sheet(|records| {
        let mut rules = records
            .iter()
            .enumerate()
            .filter(|(_, r)| r.0 == kind::D_VAL);
        let _ = rules.next().unwrap();
        let index = rules.next().unwrap().0;
        let mut payload = 10u32.to_le_bytes().to_vec();
        for unit in "Alpha,Beta".encode_utf16() {
            payload.extend_from_slice(&unit.to_le_bytes());
        }
        let ac = records
            .iter()
            .find(|r| r.0 == kind::AC_BEGIN)
            .unwrap()
            .1
            .clone();
        records.splice(
            index..index,
            [
                (kind::AC_BEGIN, ac),
                (kind::D_VAL_LIST, payload),
                (kind::AC_END, Vec::new()),
            ],
        );
    });
    let package = litchi_xlsb::Package::from_bytes(&bytes).unwrap();
    let before_parts = members(&package.to_bytes().unwrap());
    const SHEET: &str = "xl/worksheets/sheet1.bin";
    let mut edit = package.edit_data_validations().unwrap();
    let mut rule = edit.validations(0).unwrap()[1].clone();
    assert_eq!(rule.list_formula.as_deref(), Some("Alpha,Beta"));
    rule.error_title = Some("Keep list bytes".into());
    edit.set(0, 1, rule.clone()).unwrap();
    let changed = edit.commit().unwrap().into_package();
    assert_eq!(
        alternate_blocks(&members(&changed.to_bytes().unwrap())[SHEET]),
        alternate_blocks(&before_parts[SHEET])
    );
    rule.list_formula = Some("Gamma,Delta".into());
    rule.formula1 = rule.list_formula.clone();
    let mut edit = changed.edit_data_validations().unwrap();
    edit.set(0, 1, rule).unwrap();
    let changed = edit.commit().unwrap().into_package();
    let read = changed
        .edit_data_validations()
        .unwrap()
        .validations(0)
        .unwrap();
    assert_eq!(read[1].formula1.as_deref(), Some("Gamma,Delta"));
    let original_opaque = alternate_blocks(&members(INTERACTION)[SHEET]);
    let mut list_blocks = 0;
    let opaque = alternate_blocks(&members(&changed.to_bytes().unwrap())[SHEET])
        .into_iter()
        .filter(|block| {
            let list = litchi_xlsb::raw::Records::new(block, litchi_xlsb::raw::Limits::DEFAULT)
                .any(|record| record.unwrap().kind() == kind::D_VAL_LIST);
            list_blocks += usize::from(list);
            !list
        })
        .collect::<Vec<_>>();
    assert_eq!(list_blocks, 1);
    assert_eq!(opaque, original_opaque);
}

#[test]
fn native_excel_conditional_formatting_links_the_wrapped_data_bar_extension() {
    use litchi_xlsb::conditional_formatting::{RecordKind, RuleType};
    let bytes =
        include_bytes!("../../../test-data/microsoft-office/excel/excel-conditional-drawings.xlsb");
    let workbook = Workbook::new(Cursor::new(bytes.as_slice())).unwrap();
    let sheet = workbook.worksheet(0).unwrap();
    let formats = sheet.conditional_formattings();
    assert_eq!(
        sheet
            .tolerance_report()
            .count(litchi_xlsb::LayoutDefect::ConditionalFormulaLength),
        3
    );
    assert_eq!(
        sheet
            .tolerance_report()
            .count(litchi_xlsb::LayoutDefect::ConditionalNullText),
        1
    );
    let rules = formats
        .iter()
        .filter(|format| format.record_kind == RecordKind::Classic)
        .flat_map(|format| &format.rules)
        .collect::<Vec<_>>();
    assert_eq!(rules.len(), 8);
    let mut priorities = rules.iter().map(|rule| rule.priority).collect::<Vec<_>>();
    priorities.sort();
    assert_eq!(priorities, (1..=8).collect::<Vec<_>>());
    let base = rules.iter().find(|rule| rule.priority == 8).unwrap();
    assert_eq!(base.rule_type, RuleType::DataBar);
    assert!(base.data_bar.is_some());
    let extensions = formats
        .iter()
        .filter(|format| format.record_kind == RecordKind::Extension14)
        .flat_map(|format| &format.rules)
        .collect::<Vec<_>>();
    assert_eq!(extensions.len(), 1);
    let metadata = extensions[0].extension14.as_ref().unwrap();
    assert_eq!(metadata.linked_classic_priority, Some(8));
    assert_eq!(base.classic_extension_guid, Some(metadata.guid));
    assert!(extensions[0].data_bar14.is_some());
}

#[test]
fn native_conditional_repairs_do_not_accept_truncated_formulas_or_broken_wrappers() {
    use litchi_xlsb::conditional_formatting::Rule;
    use litchi_xlsb::raw::{Limits, Records, kind};
    let source =
        include_bytes!("../../../test-data/microsoft-office/excel/excel-conditional-drawings.xlsb");
    let sheet = members(source).remove("xl/worksheets/sheet1.bin").unwrap();
    let formula = Records::new(&sheet, Limits::DEFAULT)
        .map(Result::unwrap)
        .find(|record| record.kind() == kind::BEGIN_CF_RULE && record.payload().len() > 46)
        .unwrap();
    assert!(
        Rule::parse(formula.payload(), Limits::DEFAULT).is_err(),
        "single-record codecs cannot silently repair without a report"
    );
    for case in 0..3 {
        let bytes = rewrite_sheet_of(source, |records| {
            if case == 0 {
                let payload = &mut records
                    .iter_mut()
                    .find(|r| r.0 == kind::BEGIN_CF_RULE && r.1.len() > 46)
                    .unwrap()
                    .1;
                let size = u32::from_le_bytes(payload[46..50].try_into().unwrap());
                payload[46..50].copy_from_slice(&(size + 1).to_le_bytes());
            } else if case == 1 {
                let ext = records
                    .iter()
                    .position(|r| r.0 == kind::CF_RULE_EXT)
                    .unwrap();
                assert_eq!(records[ext + 1].0, kind::FRT_END);
                records.remove(ext + 1);
            } else {
                records
                    .iter_mut()
                    .find(|r| r.0 == kind::BEGIN_CF_RULE14)
                    .unwrap()
                    .1
                    .push(0);
            }
        });
        assert!(
            Workbook::new(Cursor::new(bytes))
                .and_then(|book| book.worksheet(0))
                .is_err(),
            "case {case}"
        );
    }
}

#[test]
fn native_excel_frozen_pane_matches_its_paired_xlsx_view() {
    let bytes = include_bytes!(
        "../../../test-data/microsoft-office/excel/excel-values-formulas-chart.xlsb"
    );
    let workbook = Workbook::new(Cursor::new(bytes.as_slice())).unwrap();
    let sheet = workbook.worksheet(0).unwrap();
    let pane = sheet.views()[0].pane.as_ref().unwrap();
    assert_eq!(pane.state, litchi_sheet::view::State::Frozen);
    assert_eq!(pane.position, litchi_sheet::view::Position::BottomRight);
    assert_eq!(pane.horizontal.unwrap().get(), 1.0);
    assert_eq!(pane.vertical.unwrap().get(), 1.0);
    assert_eq!(pane.top_left, litchi_sheet::Cell::from_a1("B2").unwrap());
    assert_eq!(
        sheet
            .tolerance_report()
            .count(litchi_xlsb::LayoutDefect::PaneFreezeFlags),
        1
    );
}

#[test]
fn native_excel_workbooks_materialize_every_declared_worksheet() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/microsoft-office/excel");
    let mut failures = Vec::new();
    for (name, count) in [
        ("excel-values-formulas-chart.xlsb", 2),
        ("excel-interaction-print.xlsb", 1),
        ("excel-recorded-macro.xlsb", 1),
        ("excel-visual-analytics.xlsb", 3),
        ("excel-conditional-drawings.xlsb", 1),
        ("excel-links-objects.xlsb", 1),
        ("excel-pivot-controls.xlsb", 3),
        ("excel-form-controls.xlsb", 1),
        ("excel-formula-owners.xlsb", 2),
        ("excel-connections.xlsb", 2),
        ("excel-date1904-breaks.xlsb", 1),
        ("excel-chart-families.xlsb", 18),
        ("excel-revisions.xlsb", 2),
        ("excel-metadata.xlsb", 5),
        ("excel-legacy-sheets.xlsb", 1),
    ] {
        match Workbook::new(std::fs::File::open(root.join(name)).unwrap()) {
            Ok(workbook) => {
                assert_eq!(workbook.worksheet_count(), count, "{name}");
                for index in 0..count {
                    if let Err(error) = workbook.worksheet(index) {
                        failures.push(format!("{name} worksheet {index}: {error:?}"));
                    }
                }
            },
            Err(error) => failures.push(format!("{name}: {error:?}")),
        }
    }
    assert!(failures.is_empty(), "{}", failures.join("\n"));
}

#[test]
fn native_excel_sparklines_retain_locations_identity_blocks_and_patch_inverse() {
    use litchi_sheet::sparkline::SparklineType;
    use litchi_xlsb::sparkline::{Groups, Limits};
    let bytes =
        include_bytes!("../../../test-data/microsoft-office/excel/excel-visual-analytics.xlsb");
    let package = litchi_xlsb::Package::from_bytes(bytes).unwrap();
    let before = package.to_bytes().unwrap();
    let mut edit = package.edit_sparklines().unwrap();
    let original = edit.groups(2).unwrap().unwrap();
    assert_eq!(original.len(), 2);
    for (group, (kind, column)) in original
        .as_slice()
        .iter()
        .zip([(SparklineType::Column, 6), (SparklineType::Line, 5)])
    {
        assert_eq!(group.kind(), kind);
        assert_eq!(group.sparklines().len(), 2);
        for (index, line) in group.sparklines().iter().enumerate() {
            assert_eq!(
                (line.location().row(), line.location().column()),
                (index as u32 + 1, column)
            );
            assert!(line.formula().is_some());
        }
    }
    let mut changed = original.as_slice().to_vec();
    changed[0] = changed[0].clone().with_line_weight(2.5).unwrap();
    assert!(
        edit.set(2, Groups::new(changed, Limits::DEFAULT).unwrap())
            .unwrap()
    );
    let commit = edit.commit().unwrap();
    let after = commit.package().to_bytes().unwrap();
    let reopened = litchi_xlsb::Package::from_bytes(&after).unwrap();
    let read = reopened
        .edit_sparklines()
        .unwrap()
        .groups(2)
        .unwrap()
        .unwrap();
    assert_eq!(read.as_slice()[0].line_weight(), 2.5);
    assert_eq!(read.as_slice()[1], original.as_slice()[1]);
    let before_parts = members(&before);
    let after_parts = members(&after);
    const SHEET: &str = "xl/worksheets/sheet3.bin";
    assert_eq!(
        alternate_blocks(&before_parts[SHEET]),
        alternate_blocks(&after_parts[SHEET])
    );
    for (name, data) in &before_parts {
        if name != SHEET {
            assert_eq!(after_parts.get(name), Some(data), "{name}");
        }
    }
    assert_eq!(
        commit
            .package()
            .apply_sparklines(&commit.patch().inverse())
            .unwrap()
            .into_package()
            .to_bytes()
            .unwrap(),
        before
    );
    let mut remove = package.edit_sparklines().unwrap();
    assert!(remove.remove(2).unwrap());
    let removed = remove.commit().unwrap().into_package();
    assert!(
        removed
            .edit_sparklines()
            .unwrap()
            .groups(2)
            .unwrap()
            .is_none()
    );
    let raw = &before_parts[SHEET];
    let mut records = litchi_xlsb::raw::Records::new(raw, litchi_xlsb::raw::Limits::DEFAULT);
    let mut start = None;
    let mut end = None;
    while let Some(record) = records.next() {
        let record = record.unwrap();
        if record.kind() == litchi_xlsb::raw::kind::FRT_BEGIN {
            assert!(start.replace(record.offset()).is_none());
        }
        if record.kind() == litchi_xlsb::raw::kind::FRT_END {
            end = Some(records.offset());
        }
    }
    let expected = [&raw[..start.unwrap()], &raw[end.unwrap()..]].concat();
    assert_eq!(members(&removed.to_bytes().unwrap())[SHEET], expected);
}

#[test]
fn native_excel_slicer_views_keep_names_captions_and_alternate_identities() {
    use litchi_xlsb::{
        raw::Limits,
        slicer::{parse_views, write_views},
    };
    let bytes =
        include_bytes!("../../../test-data/microsoft-office/excel/excel-pivot-controls.xlsb");
    let parts = members(bytes);
    let data = &parts["xl/slicers/slicer1.bin"];
    let mut views = parse_views(data, Limits::DEFAULT).unwrap();
    assert_eq!(views.items.len(), 2);
    for (view, name) in views.items.iter().zip(["Region", "Product"]) {
        assert_eq!(view.name, name);
        assert_eq!(view.caption.as_deref(), Some(name));
        assert_eq!(view.cache_name, format!("Slicer_{name}"));
        assert!(!view.alternate_content.is_empty());
    }
    let unchanged = write_views(&views, Limits::DEFAULT).unwrap();
    assert_eq!(alternate_blocks(&unchanged), alternate_blocks(data));
    assert_eq!(parse_views(&unchanged, Limits::DEFAULT).unwrap(), views);
    views.items[0].caption = Some("Regional sales".into());
    views.items.swap(0, 1);
    let changed = write_views(&views, Limits::DEFAULT).unwrap();
    assert_eq!(parse_views(&changed, Limits::DEFAULT).unwrap(), views);
    assert_eq!(
        alternate_blocks(&changed),
        alternate_blocks(data).into_iter().rev().collect::<Vec<_>>()
    );
    views.items[1].alternate_content = views.items[0].alternate_content.clone();
    assert!(write_views(&views, Limits::DEFAULT).is_err());
    let end = litchi_xlsb::raw::Records::new(data, Limits::DEFAULT)
        .last()
        .unwrap()
        .unwrap();
    assert_eq!(end.kind(), litchi_xlsb::raw::kind::END_SLICERS);
    assert!(parse_views(&data[..end.offset()], Limits::DEFAULT).is_err());
}

#[test]
fn native_excel_timeline_resolves_native_relationships_and_retains_february_selection() {
    use litchi_xlsb::timeline;
    let bytes =
        include_bytes!("../../../test-data/microsoft-office/excel/excel-pivot-controls.xlsb");
    let package =
        litchi_opc::OpcPackage::from_bytes(bytes, litchi_opc::SourceLimits::default()).unwrap();
    let caches = timeline::package::load_caches(
        &package,
        &litchi_opc::PackURI::new("/xl/workbook.bin").unwrap(),
        litchi_core::Limits::for_profile(litchi_core::Profile::Server),
    )
    .unwrap();
    assert_eq!(caches.len(), 1);
    let cache = &caches[0].cache;
    assert_eq!(cache.name, "NativeTimeline_Date");
    assert_eq!(cache.source_name, "Date");
    assert_eq!(
        cache.state.selection.as_ref().unwrap().start_date,
        "2026-02-01T00:00:00"
    );
    assert_eq!(
        cache.state.selection.as_ref().unwrap().end_date,
        "2026-02-28T00:00:00"
    );
    let cache_xml = timeline::write_cache(
        cache,
        litchi_core::Limits::for_profile(litchi_core::Profile::Server),
    )
    .unwrap();
    assert_eq!(
        timeline::parse_cache(
            &cache_xml,
            litchi_core::Limits::for_profile(litchi_core::Profile::Server)
        )
        .unwrap(),
        *cache
    );
    let parts = members(bytes);
    let xml = &parts["xl/timelines/timeline1.xml"];
    let views = timeline::parse_views(
        xml,
        litchi_core::Limits::for_profile(litchi_core::Profile::Server),
    )
    .unwrap();
    assert_eq!(views.items.len(), 1);
    let view = &views.items[0];
    assert_eq!(view.name, "Date");
    assert_eq!(view.cache, cache.name);
    assert_eq!(view.caption.as_deref(), Some("Date"));
    assert_eq!(view.level, timeline::Level::Month);
    assert_eq!(view.selection_level, timeline::Level::Month);
    assert_eq!(
        view.uid.as_deref(),
        Some("{A44D620F-849C-8B49-A9E0-1E454E8DE67D}")
    );
    let encoded = timeline::write_views(
        &views,
        litchi_core::Limits::for_profile(litchi_core::Profile::Server),
    )
    .unwrap();
    assert_eq!(
        timeline::parse_views(
            &encoded,
            litchi_core::Limits::for_profile(litchi_core::Profile::Server)
        )
        .unwrap(),
        views
    );
    let spoofed = std::str::from_utf8(xml).unwrap().replace(
        "http://schemas.openxmlformats.org/markup-compatibility/2006",
        "urn:wrong",
    );
    assert!(
        timeline::parse_views(
            spoofed.as_bytes(),
            litchi_core::Limits::for_profile(litchi_core::Profile::Server)
        )
        .is_err()
    );
}

#[test]
fn native_excel_let_parameter_name_and_cached_result_remain_inert() {
    let bytes =
        include_bytes!("../../../test-data/microsoft-office/excel/excel-formula-owners.xlsb");
    let book = Workbook::new(Cursor::new(bytes.as_slice())).unwrap();
    for name in [
        "_xlfn.LET",
        "_xlpm.x",
        "_xlfn.SEQUENCE",
        "_xlfn._xlws.FILTER",
    ] {
        assert!(
            book.defined_names().iter().any(|entry| entry == name),
            "{name}"
        );
    }
    assert_eq!(
        book.tolerance_report()
            .count(litchi_xlsb::LayoutDefect::NameWithoutShortcut),
        1
    );
    let sheet = book.worksheet(1).unwrap();
    let cell = sheet.get_cell(3, 5).unwrap();
    // The renderer deliberately leaves the context-dependent 0x00FF UDF
    // slot opaque. Prove that this remains a formula, with its original
    // name tokens and cache; a scalar-only fallback would lose that data.
    assert!(cell.is_formula());
    assert_eq!(
        cell.formula_resolution_status(),
        Some(litchi_xlsb::FormulaResolutionStatus::Opaque(
            litchi_xlsb::FormulaOpacityReason::Unsupported
        ))
    );
    assert_eq!(cell.cached_value(), Some(&CellValue::Float(50.0)));
    assert_eq!(
        cell.formula_bytes(),
        Some(
            [
                0x23, 2, 0, 0, 0, 0x23, 4, 0, 0, 0, 0x1e, 7, 0, 0x43, 4, 0, 0, 0, 0x43, 4, 0, 0, 0,
                5, 0x1e, 1, 0, 3, 0x42, 4, 0xff, 0
            ]
            .as_slice()
        )
    );
}

#[test]
fn native_excel_external_link_uses_referenced_relative_path_and_cached_values() {
    let bytes =
        include_bytes!("../../../test-data/microsoft-office/excel/excel-links-objects.xlsb");
    let book = Workbook::new(Cursor::new(bytes.as_slice())).unwrap();
    assert_eq!(book.external_link_count(), 1);
    let link = book.external_link(0).unwrap();
    assert_eq!(link.source(), "excel-values-formulas-chart.xlsx");
    assert_eq!(link.sheet_names(), ["Sales chart", "Sheet1", "Sheet2"]);
    let sheet = book.worksheet(0).unwrap();
    for (row, expected) in [(1, 49.75), (2, 14.0), (3, 63.75)] {
        let cell = sheet.get_cell(row, 1).unwrap();
        let CellValue::Formula {
            cached_value: Some(value),
            ..
        } = cell.value()
        else {
            panic!("expected external cached formula")
        };
        assert_eq!(**value, CellValue::Float(expected));
    }
    let package = litchi_xlsb::Package::from_bytes(bytes).unwrap();
    assert_eq!(members(&package.to_bytes().unwrap()), members(bytes));
}

#[test]
fn native_excel_streaming_structured_formula_uses_complete_sheet_catalog_positions() {
    let bytes = include_bytes!(
        "../../../test-data/microsoft-office/excel/excel-values-formulas-chart.xlsb"
    );
    let book =
        litchi_xlsb::SourceBackedWorkbook::from_reader(Cursor::new(bytes.as_slice())).unwrap();
    assert_eq!(
        book.sheet_names().unwrap(),
        ["Sales chart", "Sheet1", "Sheet2"]
    );
    let tables = book.tables().unwrap();
    assert_eq!(tables.len(), 1);
    assert_eq!(tables[0].sheet_index(), 1);
    assert_eq!(tables[0].display_name(), "Table1");
    let sheet = book.worksheet_by_index(1).unwrap().unwrap();
    let mut cursor = sheet.cell_cursor().unwrap();
    let mut found = false;
    while let Some(at) = cursor.next() {
        let at = at.unwrap();
        if at == (3, 1) {
            let cell = cursor.grid().unwrap().cell_at(at).unwrap().unwrap();
            let CellValue::Formula {
                formula,
                cached_value: Some(value),
                ..
            } = cell.value()
            else {
                panic!("expected table total")
            };
            assert_eq!(formula, "SUM(Table1[Quantity])");
            assert_eq!(**value, CellValue::Float(14.0));
            found = true;
        }
    }
    assert!(found);
}

#[test]
fn native_sparkline_alternate_content_accepts_both_positions_but_rejects_broken_envelopes() {
    use litchi_xlsb::raw::kind;
    let bytes =
        include_bytes!("../../../test-data/microsoft-office/excel/excel-visual-analytics.xlsb");
    for case in 0..4 {
        let changed = rewrite_sheet_part(bytes, "/xl/worksheets/sheet3.bin", |records| {
            let group = records
                .iter()
                .position(|r| r.0 == kind::BEGIN_SPARKLINE_GROUPS)
                .unwrap();
            let ac = group + 1;
            assert_eq!(records[ac].0, kind::AC_BEGIN);
            let end = ac
                + records[ac..]
                    .iter()
                    .position(|r| r.0 == kind::AC_END)
                    .unwrap();
            match case {
                0 => {
                    let begin_group = records.remove(end + 1);
                    assert_eq!(begin_group.0, kind::BEGIN_SPARKLINE_GROUP);
                    records.insert(ac, begin_group);
                },
                1 => {
                    records.remove(end);
                },
                2 => {
                    records[end].1.push(0);
                },
                3 => {
                    let end = records.iter().position(|r| r.0 == kind::FRT_END).unwrap();
                    records.remove(end);
                },
                _ => unreachable!(),
            }
        });
        let package = litchi_xlsb::Package::from_bytes(&changed);
        if package.is_err() {
            assert_ne!(case, 0, "valid inside-group prefix: {:?}", package.err());
            continue;
        }
        let package = package.unwrap();
        let result = package.edit_sparklines().and_then(|edit| edit.groups(2));
        if case == 0 {
            assert_eq!(result.unwrap().unwrap().len(), 2);
        } else {
            assert!(result.is_err(), "case {case}");
        }
    }
}
