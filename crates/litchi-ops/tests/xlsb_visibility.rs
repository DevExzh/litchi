//! Unified XLSB worksheet-tab visibility writes through the public facade.
//!
//! The shared workbook vocabulary uses the three Office tab states
//! `visible`, `hidden`, and `veryHidden`.  `hidden` is the convenient boolean
//! projection of the latter two states; a request containing both spellings
//! is deliberately rejected so the result cannot depend on property order.
#![cfg(feature = "xlsb")]

use std::io::Cursor;

use litchi_core::selector::{Path as SelectorPath, Query, Scope};
use litchi_opc::{OpcPackage, PackURI};
use litchi_ops::transaction::SemanticLimits;
use litchi_ops::{
    BatchItem, BatchMode, Code, File, ItemResult, OpenOptions, Props, SourceMode, props,
};
use litchi_xlsb::raw::{Encoder, Records, kind};
use litchi_xlsb::sheet::{Kind, State};

const FIRST: &str = "/sheet[1]";
const SECOND: &str = "/sheet[2]";
const THIRD: &str = "/sheet[3]";

fn fixture() -> Vec<u8> {
    use litchi_xlsb::builder::{self, Builder};

    let mut writer = Builder::new();
    let mut data = builder::Worksheet::new("Data");
    data.set_cell(0, 0, "seed");
    writer.add_worksheet(data);
    let mut archive = builder::Worksheet::new("Archive");
    archive.set_cell(0, 0, "archive");
    writer.add_worksheet(archive);
    let mut report = builder::Worksheet::new("Report");
    report.set_cell(0, 0, "report");
    writer.add_worksheet(report);

    let mut output = Cursor::new(Vec::new());
    writer
        .build()
        .expect("build XLSB visibility fixture")
        .write_to(&mut output)
        .expect("serialize XLSB visibility fixture");
    output.into_inner()
}

/// Rewrite only the workbook tab states while retaining the original package
/// member order and all unrelated workbook records.
fn with_states(bytes: &[u8], states: &[u32]) -> Vec<u8> {
    let mut package = OpcPackage::from_bytes(bytes, litchi_opc::SourceLimits::default())
        .expect("parse XLSB package");
    let workbook_uri = PackURI::new("/xl/workbook.bin").expect("workbook URI");
    let original = package
        .get_part(&workbook_uri)
        .expect("workbook part")
        .blob()
        .to_vec();
    let mut rewritten = Vec::new();
    let mut position = 0;
    {
        let mut writer = Encoder::new(&mut rewritten, litchi_xlsb::raw::Limits::DEFAULT);
        for record in Records::new(&original, litchi_xlsb::raw::Limits::DEFAULT) {
            let record = record.expect("valid workbook record");
            let mut payload = record.payload().to_vec();
            if record.kind() == kind::BUNDLE_SH {
                if let Some(state) = states.get(position) {
                    payload[..4].copy_from_slice(&state.to_le_bytes());
                }
                position += 1;
            }
            writer
                .write_record(record.kind(), &payload)
                .expect("rewrite workbook record");
        }
    }
    package
        .get_part_mut(&workbook_uri)
        .expect("mutable workbook part")
        .set_blob(rewritten);
    let mut output = Vec::new();
    package
        .write_to(&mut output)
        .expect("serialize rewritten workbook");
    output
}

/// Build a complete tab catalog with a chart tab and an unbound, very-hidden
/// tab.  Visibility writes target the worksheet after the chart; the test
/// then verifies that the non-worksheet catalog entries survived unchanged.
fn catalog_fixture() -> Vec<u8> {
    use litchi_xlsb::builder::chart::{Anchor, Chart};
    use litchi_xlsb::builder::{self, Builder};

    let chart = Chart::bar_chart_with_cache(
        "Sales",
        "Data!$A$1:$A$2",
        &["North", "South"],
        "Data!$B$1:$B$2",
        &[42.0, 55.0],
        Anchor::new(0, 0, 10, 20),
    )
    .expect("chart fixture");
    let mut writer = Builder::new();
    writer
        .add_chart_sheet(builder::ChartSheet::new("Sales Chart", chart))
        .expect("add chart tab");
    let mut data = builder::Worksheet::new("Data");
    data.set_cell(0, 0, "data");
    writer.add_worksheet(data);
    writer.add_worksheet(builder::Worksheet::new("Orphan"));
    writer.add_worksheet(builder::Worksheet::new("Tail"));

    let mut output = Cursor::new(Vec::new());
    writer
        .build()
        .expect("build catalog fixture")
        .write_to(&mut output)
        .expect("serialize catalog fixture");
    with_null_relationship(&output.into_inner(), 2)
}

/// Turn one workbook bundle relationship into the validated null
/// relationship form used by the XLSB catalog for an unbound very-hidden tab.
fn with_null_relationship(bytes: &[u8], target_position: usize) -> Vec<u8> {
    let mut package = OpcPackage::from_bytes(bytes, litchi_opc::SourceLimits::default())
        .expect("parse catalog package");
    let workbook_uri = PackURI::new("/xl/workbook.bin").expect("workbook URI");
    let original = package
        .get_part(&workbook_uri)
        .expect("workbook part")
        .blob()
        .to_vec();
    let mut rewritten = Vec::new();
    let mut position = 0;
    {
        let mut writer = Encoder::new(&mut rewritten, litchi_xlsb::raw::Limits::DEFAULT);
        for record in Records::new(&original, litchi_xlsb::raw::Limits::DEFAULT) {
            let record = record.expect("valid workbook record");
            let mut payload = record.payload().to_vec();
            if record.kind() == kind::BUNDLE_SH {
                if position == target_position {
                    payload[..4].copy_from_slice(&2_u32.to_le_bytes());
                    let current_id =
                        u32::from_le_bytes(payload[4..8].try_into().expect("bundle id"));
                    let strings_offset = if (1..=0xFFFF).contains(&current_id) {
                        8
                    } else {
                        12
                    };
                    let rel_units = u32::from_le_bytes(
                        payload[strings_offset..strings_offset + 4]
                            .try_into()
                            .expect("relationship string length"),
                    ) as usize;
                    let rel_end = strings_offset + 4 + rel_units * 2;
                    let mut null_payload = Vec::with_capacity(payload.len() - rel_units * 2);
                    null_payload.extend_from_slice(&payload[..strings_offset]);
                    null_payload.extend_from_slice(&u32::MAX.to_le_bytes());
                    null_payload.extend_from_slice(&payload[rel_end..]);
                    payload = null_payload;
                }
                position += 1;
            }
            writer
                .write_record(record.kind(), &payload)
                .expect("rewrite bundle record");
        }
    }
    package
        .get_part_mut(&workbook_uri)
        .expect("mutable workbook part")
        .set_blob(rewritten);
    let mut output = Vec::new();
    package
        .write_to(&mut output)
        .expect("serialize null relationship package");
    output
}

fn open_owner(bytes: &[u8], eager: bool) -> File {
    litchi_ops::open_bytes_with(
        bytes.to_vec(),
        &OpenOptions::new().source_mode(if eager {
            SourceMode::Eager
        } else {
            SourceMode::Auto
        }),
    )
    .expect("open XLSB owner")
}

fn sheet_property(file: &File, path: &str, key: &str) -> Option<String> {
    file.get(path)
        .unwrap_or_else(|error| panic!("resolve {path}: {error}"))
        .view()
        .format()
        .unwrap_or_else(|error| panic!("read {path} properties: {error}"))
        .find(|property| property.key() == key)
        .map(|property| property.value().to_string())
}

fn scoped_sheet_query(selector: &str) -> Query<'_> {
    Query::parse(selector)
        .expect("parse sheet selector")
        .with_scope(Scope::Path(Box::new(
            SelectorPath::parse("/").expect("root selector scope"),
        )))
}

fn batch_set(path: &'static str, properties: Props<'static>) -> BatchItem<'static> {
    let mut item = BatchItem::new("set");
    item.path = Some(path.into());
    item.props = Some(properties);
    item
}

#[test]
fn worksheet_visibility_round_trips_for_source_and_eager_owners() {
    let states_bytes = with_states(&fixture(), &[0, 1, 2]);
    let bytes = fixture();
    for eager in [false, true] {
        let states = open_owner(&states_bytes, eager);
        assert_eq!(
            sheet_property(&states, FIRST, "visibility").as_deref(),
            None
        );
        assert_eq!(
            sheet_property(&states, SECOND, "visibility").as_deref(),
            Some("hidden")
        );
        assert_eq!(
            sheet_property(&states, THIRD, "visibility").as_deref(),
            Some("veryHidden")
        );
        assert_eq!(sheet_property(&states, FIRST, "hidden"), None);
        assert_eq!(
            sheet_property(&states, SECOND, "hidden").as_deref(),
            Some("true")
        );
        assert_eq!(
            sheet_property(&states, THIRD, "hidden").as_deref(),
            Some("true")
        );

        let source = open_owner(&bytes, eager);
        let before = source.to_bytes().expect("publish source");
        let hidden_second = source
            .set(SECOND, props! { "hidden" => true })
            .expect("hide second worksheet");
        assert_eq!(
            sheet_property(&hidden_second, SECOND, "visibility").as_deref(),
            Some("hidden")
        );
        assert_eq!(
            sheet_property(&hidden_second, SECOND, "hidden").as_deref(),
            Some("true")
        );
        assert_eq!(source.to_bytes().expect("source remains unchanged"), before);

        let very_hidden = hidden_second
            .set(THIRD, props! { "visibility" => "veryHidden" })
            .expect("make third worksheet very hidden");
        assert_eq!(
            sheet_property(&very_hidden, THIRD, "visibility").as_deref(),
            Some("veryHidden")
        );

        let two_visible = very_hidden
            .set(SECOND, props! { "hidden" => false })
            .expect("show second worksheet");
        let hidden_first = two_visible
            .set(FIRST, props! { "hidden" => true })
            .expect("hide first worksheet");
        let very_hidden_first = hidden_first
            .set(FIRST, props! { "visibility" => "veryHidden" })
            .expect("make first worksheet very hidden");
        let visible = very_hidden_first
            .set(FIRST, props! { "hidden" => false })
            .expect("show first worksheet");
        assert_eq!(
            sheet_property(&visible, FIRST, "visibility").as_deref(),
            None
        );
        assert_eq!(sheet_property(&visible, FIRST, "hidden"), None);

        let reopened = open_owner(&visible.to_bytes().expect("publish update"), eager);
        assert_eq!(
            sheet_property(&reopened, FIRST, "visibility").as_deref(),
            None
        );
        assert_eq!(
            sheet_property(&reopened, SECOND, "visibility").as_deref(),
            None
        );
        assert_eq!(
            sheet_property(&reopened, THIRD, "visibility").as_deref(),
            Some("veryHidden")
        );
    }
}

#[test]
fn worksheet_name_and_visibility_are_one_atomic_assignment() {
    let bytes = fixture();
    for eager in [false, true] {
        let source = open_owner(&bytes, eager);
        let before = source.to_bytes().expect("publish source");
        let renamed = source
            .set(
                FIRST,
                props! {
                    "name" => "Input",
                    "visibility" => "hidden",
                },
            )
            .expect("rename and hide worksheet");
        assert_eq!(
            sheet_property(&renamed, FIRST, "name").as_deref(),
            Some("Input")
        );
        assert_eq!(
            sheet_property(&renamed, FIRST, "visibility").as_deref(),
            Some("hidden")
        );
        assert_eq!(source.to_bytes().expect("source remains unchanged"), before);

        let reopened = open_owner(&renamed.to_bytes().expect("publish renamed sheet"), eager);
        assert_eq!(
            sheet_property(&reopened, FIRST, "name").as_deref(),
            Some("Input")
        );
        assert_eq!(
            sheet_property(&reopened, FIRST, "visibility").as_deref(),
            Some("hidden")
        );
    }
}

#[test]
fn worksheet_visibility_queries_update_multiple_tabs_without_touching_the_rest() {
    let bytes = fixture();
    for eager in [false, true] {
        let source = open_owner(&bytes, eager);
        let before = source.to_bytes().expect("publish source");
        let hidden = source
            .set(
                scoped_sheet_query("sheet[name!=Data]"),
                props! { "visibility" => "hidden" },
            )
            .expect("hide both non-Data worksheets");
        assert_eq!(
            sheet_property(&hidden, FIRST, "visibility").as_deref(),
            None
        );
        assert_eq!(
            sheet_property(&hidden, SECOND, "visibility").as_deref(),
            Some("hidden")
        );
        assert_eq!(
            sheet_property(&hidden, THIRD, "visibility").as_deref(),
            Some("hidden")
        );

        let shown = hidden
            .set(
                scoped_sheet_query("sheet[visibility=hidden]"),
                props! { "visibility" => "visible" },
            )
            .expect("show both hidden worksheets");
        for path in [FIRST, SECOND, THIRD] {
            assert_eq!(sheet_property(&shown, path, "visibility").as_deref(), None);
        }
        assert_eq!(source.to_bytes().expect("source remains unchanged"), before);
    }
}

#[test]
fn hiding_the_last_visible_worksheet_is_atomic() {
    let bytes = with_states(&fixture(), &[1, 1, 0]);
    for eager in [false, true] {
        let source = open_owner(&bytes, eager);
        let before = source.to_bytes().expect("publish source");
        for properties in [
            Props::new().with("hidden", true),
            Props::new().with("visibility", "hidden"),
            Props::new()
                .with("name", "Last")
                .with("visibility", "hidden"),
        ] {
            let error = source
                .set(THIRD, properties)
                .expect_err("last visible worksheet must remain visible");
            assert_eq!(error.code(), Code::InvalidArgument);
            assert_eq!(source.to_bytes().expect("source remains unchanged"), before);
            assert_eq!(
                sheet_property(&source, THIRD, "visibility").as_deref(),
                None
            );
        }
    }
}

#[test]
fn malformed_duplicate_and_conflicting_visibility_properties_are_rejected_atomically() {
    let bytes = fixture();
    for eager in [false, true] {
        let source = open_owner(&bytes, eager);
        let before = source.to_bytes().expect("publish source");
        let cases = [
            Props::new().with("visibility", "not-a-state"),
            Props::new().with("hidden", "not-a-bool"),
            Props::new()
                .with("visibility", "hidden")
                .with("visibility", "visible"),
            Props::new()
                .with("hidden", true)
                .with("visibility", "hidden"),
            Props::new().with("unsupported", true),
        ];
        for properties in cases {
            let error = source
                .set(FIRST, properties)
                .expect_err("malformed visibility assignment");
            assert!(matches!(
                error.code(),
                Code::InvalidValue | Code::InvalidArgument | Code::UnsupportedProperty
            ));
            assert_eq!(source.to_bytes().expect("source remains unchanged"), before);
            assert_eq!(
                sheet_property(&source, FIRST, "visibility").as_deref(),
                None
            );
        }
    }
}

#[test]
fn no_op_visibility_bytes_and_inverse_patches_are_exact() {
    let bytes = with_states(&fixture(), &[0, 1, 2]);
    for eager in [false, true] {
        let source = open_owner(&bytes, eager);
        let before = source.to_bytes().expect("publish source");
        for (path, properties) in [
            (FIRST, props! { "visibility" => "visible" }),
            (SECOND, props! { "hidden" => true }),
            (THIRD, props! { "visibility" => "veryHidden" }),
        ] {
            let no_op = source
                .set(path, properties)
                .expect("equal visibility is a no-op");
            assert_eq!(no_op.to_bytes().expect("publish no-op"), before);
        }

        let transition_source = open_owner(&fixture(), eager);
        let transition_before = transition_source
            .to_bytes()
            .expect("publish transition source");
        let commit = transition_source
            .edit(|edit| edit.set(FIRST, props! { "visibility" => "hidden" }))
            .expect("stage visibility transition");
        assert!(commit.diagnostics().changed());
        let replayed = commit
            .patch()
            .apply(&transition_source)
            .expect("replay visibility patch");
        assert_eq!(
            replayed.to_bytes().expect("publish replay"),
            commit.file().to_bytes().expect("publish target")
        );
        let restored = commit
            .patch()
            .inverse()
            .apply(commit.file())
            .expect("apply inverse visibility patch");
        assert_eq!(
            restored.to_bytes().expect("publish inverse"),
            transition_before
        );

        // Two published state changes also move the active tab and clear its
        // selection. Revealing it does not reactivate it; inverse restores all
        // dependencies, while a same-state request remains an exact no-op.
        let round_trip = transition_source
            .edit(|edit| {
                edit.set(FIRST, props! { "visibility" => "hidden" })?;
                edit.set(FIRST, props! { "visibility" => "visible" })
            })
            .expect("stage hide then reveal");
        assert!(round_trip.diagnostics().changed());
        assert_eq!(sheet_property(round_trip.file(), FIRST, "visibility"), None);
        assert_eq!(
            round_trip
                .patch()
                .inverse()
                .apply(round_trip.file())
                .unwrap()
                .to_bytes()
                .unwrap(),
            transition_before
        );
    }
}

#[test]
fn mixed_visibility_cell_batches_keep_family_boundaries_and_atomicity() {
    let bytes = fixture();
    for eager in [false, true] {
        let source = open_owner(&bytes, eager);
        let before = source.to_bytes().expect("publish source");
        let commands = [
            batch_set(FIRST, Props::new().with("visibility", "hidden")),
            batch_set("/sheet[1]/A1", Props::new().with("value", "changed")),
            batch_set(SECOND, Props::new().with("hidden", true)),
        ];
        let commit = source
            .batch(commands, BatchMode::Atomic)
            .expect("mixed visibility and cell batch");
        assert_eq!(commit.value().len(), 3);
        assert!(
            commit.value().iter().all(|result| {
                matches!(result, ItemResult::Applied(applied) if applied.changed)
            })
        );
        assert_eq!(commit.diagnostics().operations(), 3);
        assert!(commit.diagnostics().commits() >= 2);
        assert_eq!(
            sheet_property(commit.file(), FIRST, "visibility").as_deref(),
            Some("hidden")
        );
        assert_eq!(
            sheet_property(commit.file(), SECOND, "visibility").as_deref(),
            Some("hidden")
        );
        assert_eq!(source.to_bytes().expect("source remains unchanged"), before);

        let failed = source
            .batch(
                [
                    batch_set(FIRST, Props::new().with("visibility", "hidden")),
                    batch_set(SECOND, Props::new().with("visibility", "not-a-state")),
                ],
                BatchMode::Atomic,
            )
            .expect_err("atomic visibility batch failure");
        assert_eq!(failed.item(), Some(1));
        assert_eq!(
            source.to_bytes().expect("atomic failure preserves source"),
            before
        );
    }
}

#[test]
fn visibility_queries_and_batches_honor_transaction_quotas() {
    let bytes = fixture();
    for eager in [false, true] {
        let source = open_owner(&bytes, eager);
        let before = source.to_bytes().expect("publish source");

        let one_operation = SemanticLimits::builder()
            .max_operations(1)
            .expect("operation limit")
            .build()
            .expect("valid operation profile");
        let operation_error = source
            .edit_with_options(
                litchi_ops::EditOptions::new().semantic_limits(one_operation),
                |edit| {
                    edit.set(FIRST, props! { "visibility" => "hidden" })?;
                    edit.set(SECOND, props! { "visibility" => "hidden" })
                },
            )
            .expect_err("two visibility operations exceed the bound");
        assert_eq!(operation_error.code(), Code::LimitExceeded);

        let one_target = SemanticLimits::builder()
            .max_expanded_targets(1)
            .expect("expanded target limit")
            .build()
            .expect("valid target profile");
        let target_error = source
            .edit_with_options(
                litchi_ops::EditOptions::new().semantic_limits(one_target),
                |edit| {
                    edit.set(
                        scoped_sheet_query("sheet[name!=Data]"),
                        props! { "visibility" => "hidden" },
                    )
                },
            )
            .expect_err("multi-target visibility query exceeds the bound");
        assert_eq!(target_error.code(), Code::LimitExceeded);

        let tiny_snapshot = SemanticLimits::builder()
            .max_snapshot_bytes(1)
            .expect("snapshot limit")
            .build()
            .expect("valid snapshot profile");
        let snapshot_error = source
            .edit_with_options(
                litchi_ops::EditOptions::new().semantic_limits(tiny_snapshot),
                |edit| edit.set(FIRST, props! { "visibility" => "hidden" }),
            )
            .expect_err("visibility patch exceeds the snapshot bound");
        assert_eq!(snapshot_error.code(), Code::LimitExceeded);
        assert_eq!(
            source
                .to_bytes()
                .expect("all quota failures preserve source"),
            before
        );
    }
}

#[test]
fn complete_catalog_visibility_write_preserves_chart_and_unbound_tabs() {
    let bytes = catalog_fixture();
    let original_tabs = litchi_xlsb::Workbook::new(Cursor::new(bytes.clone()))
        .expect("open complete catalog fixture")
        .sheets()
        .map(|tab| (tab.name().to_owned(), tab.kind(), tab.state()))
        .collect::<Vec<_>>();
    assert_eq!(
        original_tabs,
        vec![
            ("Sales Chart".to_owned(), Kind::Chart, State::Visible),
            ("Data".to_owned(), Kind::Worksheet, State::Visible),
            ("Orphan".to_owned(), Kind::Unbound, State::VeryHidden),
            ("Tail".to_owned(), Kind::Worksheet, State::Visible),
        ]
    );

    for eager in [false, true] {
        let source = open_owner(&bytes, eager);
        let before = source.to_bytes().expect("publish catalog source");
        let unchanged_unbound = source
            .set("/sheet[3]", props! { "visibility" => "veryHidden" })
            .expect("equal very-hidden state is a no-op for an unbound tab");
        assert_eq!(
            unchanged_unbound.to_bytes().expect("publish unbound no-op"),
            before
        );
        let reveal_error = source
            .set("/sheet[3]", props! { "visibility" => "visible" })
            .expect_err("an unbound tab cannot be revealed");
        assert_eq!(reveal_error.code(), Code::UnsupportedCapability);
        assert_eq!(
            source.to_bytes().expect("unbound refusal preserves source"),
            before
        );

        let chart_hidden = source
            .set("/sheet[1]", props! { "visibility" => "hidden" })
            .expect("hide chart tab");
        assert_eq!(
            sheet_property(&chart_hidden, "/sheet[1]", "visibility").as_deref(),
            Some("hidden")
        );
        let updated = chart_hidden
            .set("/sheet[2]", props! { "visibility" => "hidden" })
            .expect("hide worksheet after chart tab");
        assert_eq!(
            sheet_property(&updated, "/sheet[2]", "visibility").as_deref(),
            Some("hidden")
        );
        let output = updated.to_bytes().expect("publish catalog update");
        let tabs = litchi_xlsb::Workbook::new(Cursor::new(output.clone()))
            .expect("reopen catalog update")
            .sheets()
            .map(|tab| (tab.name().to_owned(), tab.kind(), tab.state()))
            .collect::<Vec<_>>();
        assert_eq!(
            tabs[0],
            ("Sales Chart".to_owned(), Kind::Chart, State::Hidden)
        );
        assert_eq!(tabs[1], ("Data".to_owned(), Kind::Worksheet, State::Hidden));
        assert_eq!(
            tabs[2],
            ("Orphan".to_owned(), Kind::Unbound, State::VeryHidden)
        );
        assert_eq!(
            tabs[3],
            ("Tail".to_owned(), Kind::Worksheet, State::Visible)
        );

        let original_chart_parts = chart_parts_except_selection(&bytes);
        let output_chart_parts = chart_parts_except_selection(&output);
        assert_eq!(
            litchi_xlsb::Workbook::new(Cursor::new(output.clone()))
                .unwrap()
                .active_catalog_position(),
            Some(3)
        );
        assert_eq!(output_chart_parts, original_chart_parts);
    }
}

fn chart_parts_except_selection(bytes: &[u8]) -> Vec<Vec<u8>> {
    let package = OpcPackage::from_bytes(bytes, litchi_opc::SourceLimits::default())
        .expect("parse XLSB package");
    package
        .iter_parts()
        .filter(|part| part.content_type() == "application/vnd.ms-excel.chartsheet")
        .map(|part| {
            let part = package.get_part(part.partname()).unwrap();
            let source = part.blob();
            let mut normalized = source.to_vec();
            for record in Records::new(source, litchi_xlsb::raw::Limits::DEFAULT) {
                let record = record.unwrap();
                if record.kind() == kind::BEGIN_CS_VIEW {
                    let (_, header_len) = litchi_xlsb::raw::Header::parse(
                        &source[record.offset()..],
                        Default::default(),
                    )
                    .unwrap();
                    normalized[record.offset() + header_len] &= !1;
                }
            }
            normalized
        })
        .collect()
}

#[test]
fn hiding_active_or_selected_tabs_updates_view_state_and_inverse_restores_it() {
    for eager in [false, true] {
        let source = open_owner(&fixture(), eager);
        let before = source.to_bytes().unwrap();
        let commit = source
            .edit(|edit| edit.set(FIRST, props! { "hidden" => true }))
            .unwrap();
        let after = commit.file().to_bytes().unwrap();
        let workbook = litchi_xlsb::Workbook::new(Cursor::new(after.clone())).unwrap();
        assert_eq!(workbook.active_catalog_position(), Some(1));
        assert_eq!(
            sheet_property(commit.file(), "/", "activeTab").as_deref(),
            Some("1")
        );
        assert!(!worksheet_selected(&after, "/xl/worksheets/sheet1.bin"));
        assert!(worksheet_selected(&after, "/xl/worksheets/sheet2.bin"));
        let restored = commit.patch().inverse().apply(commit.file()).unwrap();
        assert_eq!(restored.to_bytes().unwrap(), before);

        let nonactive = source.set(THIRD, props! { "hidden" => true }).unwrap();
        let bytes = nonactive.to_bytes().unwrap();
        assert!(!worksheet_selected(&bytes, "/xl/worksheets/sheet3.bin"));
        assert_eq!(
            litchi_xlsb::Workbook::new(Cursor::new(bytes))
                .unwrap()
                .active_catalog_position(),
            Some(0)
        );
    }
}

fn worksheet_selected(bytes: &[u8], uri: &str) -> bool {
    let package = OpcPackage::from_bytes(bytes, litchi_opc::SourceLimits::default()).unwrap();
    let part = package.get_part(&PackURI::new(uri).unwrap()).unwrap();
    let record = Records::new(part.blob(), litchi_xlsb::raw::Limits::DEFAULT)
        .map(Result::unwrap)
        .find(|record| record.kind() == kind::BEGIN_WS_VIEW)
        .expect("worksheet view");
    let flags = u16::from_le_bytes(record.payload()[..2].try_into().unwrap());
    flags & 0x0040 != 0
}

#[test]
fn multi_window_active_relocation_is_an_explicit_capability_refusal() {
    let bytes = fixture();
    let mut package = OpcPackage::from_bytes(&bytes, litchi_opc::SourceLimits::default()).unwrap();
    let uri = PackURI::new("/xl/workbook.bin").unwrap();
    let source = package.get_part(&uri).unwrap().blob().to_vec();
    let mut output = Vec::new();
    let mut inserted = false;
    {
        let mut writer = Encoder::new(&mut output, litchi_xlsb::raw::Limits::DEFAULT);
        for record in Records::new(&source, litchi_xlsb::raw::Limits::DEFAULT) {
            let record = record.unwrap();
            writer
                .write_record(record.kind(), record.payload())
                .unwrap();
            if record.kind() == kind::BOOK_VIEW && !inserted {
                writer
                    .write_record(record.kind(), record.payload())
                    .unwrap();
                inserted = true;
            }
        }
    }
    assert!(inserted);
    package.get_part_mut(&uri).unwrap().set_blob(output);
    let mut bytes = Vec::new();
    package.write_to(&mut bytes).unwrap();
    for eager in [false, true] {
        let source = open_owner(&bytes, eager);
        let before = source.to_bytes().unwrap();
        let error = source.set(FIRST, props! { "hidden" => true }).unwrap_err();
        assert_eq!(error.code(), Code::UnsupportedCapability);
        assert_eq!(source.to_bytes().unwrap(), before);
        let equal = source
            .set(FIRST, props! { "visibility" => "visible" })
            .unwrap();
        assert_eq!(equal.to_bytes().unwrap(), before);
        let nonactive = source.set(THIRD, props! { "hidden" => true }).unwrap();
        assert_eq!(
            sheet_property(&nonactive, THIRD, "visibility").as_deref(),
            Some("hidden")
        );
    }
}
