//! Transaction joins, conflicts, and cross-facet edit tests.

use super::super::*;
use super::support::{part_text, styled_workbook};
use crate::cell::{Number, Value};
use crate::{History, HistoryLimits, Margins, PageMargin};
use std::mem::size_of;

#[test]
fn bounded_endpoint_accepts_exact_payload_and_typed_noops() {
    let source = styled_workbook();
    let existing_style = source
        .sheet("Sheet1")
        .expect("sheet lookup")
        .expect("worksheet")
        .style("A1")
        .expect("style lookup")
        .expect("existing style");
    let maximum = source.retained_bytes().saturating_mul(2);
    let mut edit = source
        .edit_with_options(EditOptions::default().with_max_retained_bytes(maximum))
        .expect("endpoint-only budget");
    let mut sheet = edit
        .sheet("Sheet1")
        .expect("sheet lookup")
        .expect("worksheet");
    sheet.set("A1", 1_i32).expect("same scalar payload");
    sheet
        .style("A1", &existing_style)
        .expect("same shared style");
    sheet
        .font(
            "B1",
            &crate::FontChange::new()
                .name(None)
                .expect("empty authored font name"),
        )
        .expect("same default font");
    sheet
        .fill(
            "A1",
            &crate::style::FillChange::new()
                .pattern(Some(crate::style::cell_format::Pattern::Solid)),
        )
        .expect("same source fill pattern");
    let commit = edit.commit().expect("endpoint no-op commit");
    assert!(commit.patch().is_noop());
}

#[test]
fn mixed_join_normalizes_unbounded_typed_noop_before_bounded_admission() {
    let source = styled_workbook();
    let maximum = source.retained_bytes().saturating_mul(2);
    let mut bounded = source
        .edit_with_options(EditOptions::default().with_max_retained_bytes(maximum))
        .expect("bounded edit");
    let mut unbounded = source.edit().expect("unbounded edit");
    unbounded
        .sheet("Sheet1")
        .expect("sheet lookup")
        .expect("worksheet")
        .font(
            "B1",
            &crate::FontChange::new()
                .name(None)
                .expect("empty authored font name"),
        )
        .expect("stage unbounded typed no-op");
    bounded
        .join(unbounded)
        .expect("typed no-op does not consume bounded endpoint budget");
    let commit = bounded.commit().expect("mixed no-op commit");
    assert!(commit.patch().is_noop());
}

#[test]
fn mixed_join_normalizes_typed_noop_against_pending_complete_style() {
    let source = styled_workbook();
    let pending_style = source
        .sheet("Sheet1")
        .expect("sheet lookup")
        .expect("worksheet")
        .style("A1")
        .expect("style lookup")
        .expect("pending style");
    let worksheet_bytes = source
        .inner
        .package
        .get_part(&source.inner.sheets[0].part_uri)
        .expect("worksheet part")
        .blob()
        .len();
    let maximum = source
        .retained_bytes()
        .saturating_mul(2)
        .saturating_add(worksheet_bytes)
        .saturating_add(size_of::<Change>())
        .saturating_add(size_of::<PartChange>())
        .saturating_add(size_of::<StyleEffect>());
    let mut bounded = source
        .edit_with_options(EditOptions::default().with_max_retained_bytes(maximum))
        .expect("bounded edit");
    let mut unbounded = source.edit().expect("unbounded edit");
    unbounded
        .sheet("Sheet1")
        .expect("sheet lookup")
        .expect("worksheet")
        .style("B1", &pending_style)
        .and_then(|sheet| {
            sheet.fill(
                "B1",
                &crate::style::FillChange::new()
                    .pattern(Some(crate::style::cell_format::Pattern::Solid)),
            )
        })
        .expect("stage pending style and matching typed facet");
    bounded
        .join(unbounded)
        .expect("pending-style typed no-op joins within its final capture");
    assert!(
        bounded
            .sheets
            .get(&0)
            .is_none_or(|actions| actions.fills.is_empty())
    );
}

#[test]
fn unbounded_join_normalizes_source_equivalent_effects_before_conflicts() {
    let baseline = crate::Package::new().expect("baseline workbook");
    let mut install = baseline.edit().expect("install edit");
    install
        .sheet("Sheet1")
        .expect("sheet lookup")
        .expect("worksheet")
        .set("A1", "shared value")
        .expect("install value");
    let source = install.commit().expect("install commit").into_package();

    let mut left = source.edit().expect("left edit");
    left.sheet("Sheet1")
        .expect("sheet lookup")
        .expect("worksheet")
        .set("A1", "shared value")
        .expect("left source-equivalent set");
    let mut right = source.edit().expect("right edit");
    right
        .sheet("Sheet1")
        .expect("sheet lookup")
        .expect("worksheet")
        .set("A1", "shared value")
        .and_then(|sheet| sheet.clear("B2"))
        .and_then(|sheet| sheet.remove("C3"))
        .and_then(|sheet| sheet.reset_style("D4"))
        .expect("right source-equivalent effects");

    left.join(right)
        .expect("source-equivalent effects do not overlap");
    let committed = left.commit().expect("source-equivalent join commit");
    assert!(committed.patch().is_noop());
}

#[test]
fn failed_join_preserves_original_self_and_rejected_effects() {
    let source = crate::Package::new().expect("source workbook");
    let mut left = source
        .edit_with_options(EditOptions::default().with_max_retained_bytes(usize::MAX))
        .expect("bounded edit");
    left.sheet("Sheet1")
        .expect("sheet lookup")
        .expect("worksheet")
        .set("A1", "left")
        .expect("left value");
    let left_len = left.len();

    let mut right = source.edit().expect("unbounded edit");
    right
        .sheet("Sheet1")
        .expect("sheet lookup")
        .expect("worksheet")
        .font(
            "B1",
            &crate::FontChange::new()
                .name(None)
                .expect("empty font name"),
        )
        .and_then(|sheet| sheet.set("A1", "right"))
        .and_then(|sheet| sheet.remove("B2"))
        .and_then(|sheet| sheet.clear("C3"))
        .expect("right effects");
    let right_len = right.len();

    let error = left
        .join(right)
        .expect_err("overlapping payloads must refuse");
    assert!(matches!(error.failure(), JoinFailure::Overlap(_)));
    assert_eq!(left.len(), left_len);
    assert_eq!(error.rejected().len(), right_len);
    assert!(matches!(
        left.sheets
            .get(&0)
            .and_then(|actions| {
                actions
                    .cells
                    .get(&Address::from_a1("A1").expect("A1"))
            }),
        Some(Action::Update {
            payload: Some(Payload::Set(content)),
            ..
        }) if matches!(
            content.as_cell(),
            Cell::Value(Value::Text(value)) if value.as_str() == "left"
        )
    ));
    let mut rejected = error.into_rejected();
    assert_eq!(rejected.len(), right_len);
    assert!(rejected.sheets.get(&0).is_some_and(|actions| {
        actions
            .fonts
            .contains_key(&Address::from_a1("B1").expect("B1"))
    }));
    assert!(matches!(
        rejected
            .sheets
            .get(&0)
            .and_then(|actions| {
                actions
                    .cells
                    .get(&Address::from_a1("A1").expect("A1"))
            }),
        Some(Action::Update {
            payload: Some(Payload::Set(content)),
            ..
        }) if matches!(
            content.as_cell(),
            Cell::Value(Value::Text(value)) if value.as_str() == "right"
        )
    ));
    let _ = rejected
        .sheet("Sheet1")
        .expect("sheet lookup")
        .expect("worksheet")
        .state("B2")
        .expect("rejected state remains readable");
}

#[test]
fn primary_edit_capture_budget_rejects_before_cell_staging() {
    let source = crate::Package::new().expect("source workbook");
    // The native ledger admits only the two final endpoints at this
    // boundary. The worksheet before-image and semantic cell delta must be
    // rejected before the action map stages anything.
    let maximum = source.retained_bytes().saturating_mul(2);
    let mut edit = source
        .edit_with_options(EditOptions::default().with_max_retained_bytes(maximum))
        .expect("base capture fits the budget");
    let error = edit
        .sheet("Sheet1")
        .expect("sheet lookup")
        .expect("worksheet")
        .set("A1", "capture")
        .expect_err("worksheet capture should be rejected before staging");
    assert!(matches!(
        error,
        Error::Limit {
            kind: crate::LimitKind::Workbook(LimitKind::SnapshotBytes),
            ..
        }
    ));
}

#[test]
fn primary_edit_capture_budget_replaces_repeated_cell_payload() {
    let source = crate::Package::new().expect("source workbook");
    let sheet = &source.inner.sheets[0];
    let worksheet_bytes = source
        .inner
        .package
        .get_part(&sheet.part_uri)
        .expect("worksheet part")
        .blob()
        .len();
    let final_content = Content::from("final authored value");
    let maximum = source
        .retained_bytes()
        .saturating_mul(2)
        .saturating_add(worksheet_bytes)
        .saturating_add(size_of::<Change>())
        .saturating_add(size_of::<PartChange>())
        .saturating_add(final_content.retained_bytes());
    let mut edit = source
        .edit_with_options(EditOptions::default().with_max_retained_bytes(maximum))
        .expect("base capture fits the budget");
    let mut sheet = edit
        .sheet("Sheet1")
        .expect("sheet lookup")
        .expect("worksheet");
    sheet
        .set("A1", "short")
        .expect("first authored payload fits");
    sheet
        .set("A1", final_content)
        .expect("replacement keeps only final payload charge");
}

#[test]
fn canceled_payload_and_absent_remove_release_capture() {
    let empty = crate::Package::new().expect("source workbook");
    let mut install = empty.edit().expect("install edit");
    install
        .sheet("Sheet1")
        .expect("sheet lookup")
        .expect("worksheet")
        .set("A1", "base")
        .expect("base value");
    let source = install.commit().expect("install commit").into_package();
    let sheet = &source.inner.sheets[0];
    let worksheet_bytes = source
        .inner
        .package
        .get_part(&sheet.part_uri)
        .expect("worksheet part")
        .blob()
        .len();
    let payload = Content::from("after cancellation");
    let maximum = source
        .retained_bytes()
        .saturating_mul(2)
        .saturating_add(worksheet_bytes)
        .saturating_add(size_of::<Change>())
        .saturating_add(size_of::<PartChange>())
        .saturating_add(payload.retained_bytes());
    let mut edit = source
        .edit_with_options(EditOptions::default().with_max_retained_bytes(maximum))
        .expect("base capture fits");
    let mut sheet = edit
        .sheet("Sheet1")
        .expect("sheet lookup")
        .expect("worksheet");
    sheet.set("A1", "temporary").expect("temporary value");
    sheet.set("A1", "base").expect("value back to source");
    sheet.remove("B1").expect("absent remove is a no-op");
    sheet
        .set("C1", payload)
        .expect("canceled entries released their capture");
}

#[test]
fn failed_typed_admission_preserves_a_pending_remove() {
    let source = styled_workbook();
    let sheet = &source.inner.sheets[0];
    let worksheet_bytes = source
        .inner
        .package
        .get_part(&sheet.part_uri)
        .expect("worksheet part")
        .blob()
        .len();
    let maximum = source
        .retained_bytes()
        .saturating_mul(2)
        .saturating_add(worksheet_bytes)
        .saturating_add(size_of::<PartChange>())
        .saturating_add(size_of::<Change>())
        .saturating_add(size_of::<Action>());
    let mut edit = source
        .edit_with_options(EditOptions::default().with_max_retained_bytes(maximum))
        .expect("bounded edit");
    let mut sheet = edit
        .sheet("Sheet1")
        .expect("sheet lookup")
        .expect("worksheet");
    sheet.remove("A1").expect("remove admission");
    let error = sheet
        .font("A1", &crate::FontChange::new().bold(Some(true)))
        .expect_err("typed admission exceeds the remaining capture budget");
    assert!(matches!(error, Error::Limit { .. }));
    assert!(matches!(
        edit.sheets
            .get(&0)
            .and_then(|actions| actions.cells.get(&Address::from_a1("A1").expect("A1"))),
        Some(Action::Remove)
    ));
}

#[test]
fn mixed_bounded_and_unbounded_join_admits_staged_capture() {
    let source = crate::Package::new().expect("source workbook");
    let sheet = &source.inner.sheets[0];
    let worksheet_bytes = source
        .inner
        .package
        .get_part(&sheet.part_uri)
        .expect("worksheet part")
        .blob()
        .len();
    let content = Content::from("mixed join payload");
    let maximum = source
        .retained_bytes()
        .saturating_mul(2)
        .saturating_add(worksheet_bytes)
        .saturating_add(size_of::<Change>())
        .saturating_add(size_of::<Change>())
        .saturating_add(size_of::<PartChange>())
        .saturating_add(content.retained_bytes().saturating_mul(2));
    let mut bounded = source
        .edit_with_options(EditOptions::default().with_max_retained_bytes(maximum))
        .expect("bounded edit");
    bounded
        .sheet("Sheet1")
        .expect("bounded lookup")
        .expect("bounded sheet")
        .set("A1", content.clone())
        .expect("bounded payload");
    let mut unbounded = source.edit().expect("unbounded edit");
    unbounded
        .sheet("Sheet1")
        .expect("unbounded lookup")
        .expect("unbounded sheet")
        .set("C1", content)
        .expect("unbounded payload");
    bounded
        .join(unbounded)
        .expect("mixed join admits both ledgers");
}

#[test]
fn failed_first_cell_capture_rolls_back_worksheet_descriptor() {
    let source = crate::Package::new().expect("source workbook");
    let sheet = &source.inner.sheets[0];
    let worksheet_bytes = source
        .inner
        .package
        .get_part(&sheet.part_uri)
        .expect("worksheet part")
        .blob()
        .len();
    let payload = Content::from("after rollback");
    let maximum = source
        .retained_bytes()
        .saturating_mul(2)
        .saturating_add(worksheet_bytes)
        .saturating_add(size_of::<Change>())
        .saturating_add(size_of::<PartChange>())
        .saturating_add(payload.retained_bytes());
    let mut edit = source
        .edit_with_options(EditOptions::default().with_max_retained_bytes(maximum))
        .expect("base capture fits the budget");
    let mut sheet = edit
        .sheet("Sheet1")
        .expect("sheet lookup")
        .expect("worksheet");
    assert!(
        sheet
            .set_cached("A1", Value::Number(Number::new("7").expect("number")))
            .is_err()
    );
    sheet
        .set("B1", payload)
        .expect("the failed first capture released its worksheet and cell charge");
}

#[test]
fn primary_edit_capture_budget_merges_payload_and_style_facets() {
    let source = crate::Package::new().expect("source workbook");
    let sheet = &source.inner.sheets[0];
    let worksheet_bytes = source
        .inner
        .package
        .get_part(&sheet.part_uri)
        .expect("worksheet part")
        .blob()
        .len();
    let content = Content::from("authored payload");
    let font = crate::style::FontChange::new().bold(Some(true));
    let maximum = source
        .retained_bytes()
        .saturating_mul(2)
        .saturating_add(worksheet_bytes)
        .saturating_add(size_of::<Change>())
        .saturating_add(size_of::<PartChange>())
        .saturating_add(content.retained_bytes())
        .saturating_add(font.retained_bytes());
    let mut edit = source
        .edit_with_options(EditOptions::default().with_max_retained_bytes(maximum))
        .expect("base capture fits the budget");
    let mut sheet = edit
        .sheet("Sheet1")
        .expect("sheet lookup")
        .expect("worksheet");
    sheet.set("A1", content).expect("payload facet fits");
    sheet
        .font("A1", &font)
        .expect("style facet joins the retained payload");
}

#[test]
fn joined_edits_merge_capture_ledgers_and_reject_over_budget_results() {
    let source = crate::Package::new().expect("source workbook");
    let sheet = &source.inner.sheets[0];
    let worksheet_bytes = source
        .inner
        .package
        .get_part(&sheet.part_uri)
        .expect("worksheet part")
        .blob()
        .len();
    let content = Content::from("joined payload");
    let maximum = source
        .retained_bytes()
        .saturating_mul(2)
        .saturating_add(worksheet_bytes)
        .saturating_add(size_of::<Change>())
        .saturating_add(size_of::<PartChange>())
        .saturating_add(content.retained_bytes());
    let mut left = source
        .edit_with_options(EditOptions::default().with_max_retained_bytes(maximum))
        .expect("left base capture fits");
    left.sheet("Sheet1")
        .expect("left lookup")
        .expect("left sheet")
        .set("A1", content.clone())
        .expect("left payload fits");
    let mut right = source
        .edit_with_options(EditOptions::default().with_max_retained_bytes(maximum))
        .expect("right base capture fits");
    right
        .sheet("Sheet1")
        .expect("right lookup")
        .expect("right sheet")
        .set("C3", content)
        .expect("right payload fits");

    let error = left
        .join(right)
        .expect_err("joined payloads exceed the shared cap");
    assert!(matches!(error.failure(), JoinFailure::ResourceLimit(_)));
}

fn strict_unknown_worksheet() -> crate::Package {
    let source = crate::Package::new().expect("source workbook");
    let mut package = source.inner.package.clone();
    package
        .get_part_mut(&source.inner.sheets[0].part_uri)
        .expect("worksheet part")
        .set_blob(
            br#"<x:worksheet xmlns:x="http://purl.oclc.org/ooxml/spreadsheetml/main" xmlns:u="urn:litchi:future"><x:sheetData><x:row r="1"><x:c r="A1"><x:v>7</x:v></x:c></x:row></x:sheetData><u:future marker="keep"/></x:worksheet>"#.to_vec(),
        );
    crate::Package::from_opc(package).expect("strict unknown worksheet")
}

fn test_margins() -> Margins {
    Margins::new(
        PageMargin::from_inches(0.7).expect("left margin"),
        PageMargin::from_inches(0.8).expect("right margin"),
        PageMargin::from_inches(1.0).expect("top margin"),
        PageMargin::from_inches(1.1).expect("bottom margin"),
        PageMargin::from_inches(0.3).expect("header margin"),
        PageMargin::from_inches(0.4).expect("footer margin"),
    )
}

#[test]
fn metadata_only_edit_preserves_strict_namespace_and_unknown_xml() {
    let source = strict_unknown_worksheet();
    let margins = test_margins();
    let mut edit = source.edit().expect("metadata edit");
    edit.put_page_margins("Sheet1", margins)
        .expect("put margins")
        .expect("worksheet selector");
    let committed = edit.commit().expect("metadata commit");
    let xml = part_text(committed.package(), "/xl/worksheets/sheet1.xml");

    assert!(xml.contains("<x:worksheet"));
    assert!(xml.contains("<u:future marker=\"keep\"/>"));
    assert!(xml.contains("<x:pageMargins"));
    assert_eq!(
        committed
            .package()
            .sheet("Sheet1")
            .expect("sheet lookup")
            .expect("worksheet")
            .page_margins()
            .expect("page margins read"),
        Some(margins)
    );
}

#[test]
fn metadata_only_edit_rejects_malformed_worksheet_like_the_source_parser() {
    let source = crate::Package::new().expect("source workbook");
    let mut package = source.inner.package.clone();
    let uri = source.inner.sheets[0].part_uri.clone();
    let malformed = br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData></worksheet>"#;
    package
        .get_part_mut(&uri)
        .expect("worksheet part")
        .set_blob(malformed.to_vec());
    let malformed = crate::Package::from_opc(package).expect("catalog validation");
    let mut edit = malformed.edit().expect("metadata edit");
    edit.put_page_margins("Sheet1", test_margins())
        .expect("stage margins")
        .expect("worksheet selector");
    assert!(edit.commit().is_err());
}

#[test]
fn failed_join_keeps_typed_normalization_source_and_both_edits() {
    let source = crate::Package::new().expect("source workbook");
    let mut package = source.inner.package.clone();
    let uri = source.inner.sheets[0].part_uri.clone();
    package
        .get_part_mut(&uri)
        .expect("worksheet part")
        .set_blob(
            br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData></worksheet>"#.to_vec(),
        );
    let malformed = crate::Package::from_opc(package).expect("catalog validation");
    let mut left = malformed.edit().expect("left edit");
    left.sheet("Sheet1")
        .expect("sheet lookup")
        .expect("worksheet")
        .set("A1", "left")
        .expect("left payload");
    let mut right = malformed.edit().expect("right edit");
    right
        .sheet("Sheet1")
        .expect("sheet lookup")
        .expect("worksheet")
        .set("B1", "right")
        .expect("right payload");

    let error = left
        .join(right)
        .expect_err("malformed source refuses normalization");
    assert!(matches!(
        error.failure(),
        JoinFailure::Normalization(JoinNormalization::Invalid)
    ));
    assert!(matches!(
        error.normalization_source(),
        Some(Error::Invalid(_))
    ));
    assert_eq!(left.len(), 1);
    assert_eq!(error.rejected().len(), 1);
}

#[test]
fn metadata_only_edit_with_the_existing_value_is_a_byte_no_op() {
    let source = crate::Package::new().expect("source workbook");
    let margins = test_margins();
    let mut install = source.edit().expect("install edit");
    install
        .put_page_margins("Sheet1", margins)
        .expect("put margins")
        .expect("worksheet selector");
    let installed = install.commit().expect("install commit");
    let before = installed.package().to_plain_bytes().expect("before bytes");

    let mut no_op = installed.package().edit().expect("no-op edit");
    no_op
        .put_page_margins("Sheet1", margins)
        .expect("put existing margins")
        .expect("worksheet selector");
    let no_op = no_op.commit().expect("no-op commit");
    assert!(no_op.patch().is_noop());
    assert_eq!(
        no_op.package().to_plain_bytes().expect("after bytes"),
        before
    );
}

#[test]
fn page_margins_merge_replay_inverse_and_record_history() {
    let source = crate::Package::new().expect("source workbook");
    let source_bytes = source.to_plain_bytes().expect("source bytes");
    let margins = Margins::new(
        PageMargin::from_inches(0.7).unwrap(),
        PageMargin::from_inches(0.8).unwrap(),
        PageMargin::from_inches(1.0).unwrap(),
        PageMargin::from_inches(1.1).unwrap(),
        PageMargin::from_inches(0.3).unwrap(),
        PageMargin::from_inches(0.4).unwrap(),
    );

    let mut margin_edit = source.edit().expect("margin edit");
    margin_edit
        .put_page_margins("Sheet1", margins)
        .expect("put margins")
        .expect("sheet selector");
    let mut cell_edit = source.edit().expect("cell edit");
    cell_edit
        .sheet("Sheet1")
        .expect("lookup")
        .expect("sheet")
        .set("H8", "disjoint")
        .expect("cell edit");
    let merged = margin_edit
        .plan_three_way(cell_edit, MergeLimits::new(2, 64, 128, 64))
        .expect("plan merge");
    assert!(merged.conflicts().is_empty());
    let commit = merged
        .finish()
        .expect("finish merge")
        .commit()
        .expect("commit margins");
    assert!(commit.patch().changes().iter().any(|change| {
        matches!(change.page_margins(), Some((None, Some(after))) if *after == margins)
    }));
    assert_eq!(
        commit
            .package()
            .sheet("Sheet1")
            .expect("lookup")
            .expect("sheet")
            .page_margins()
            .expect("read margins"),
        Some(margins)
    );

    let durable = commit.patch().durable().expect("durable margins patch");
    let json = durable
        .to_deterministic_json()
        .expect("deterministic margins patch");
    let durable = DurablePatch::from_deterministic_json(&json).expect("parse margins patch");
    let replayed = durable.apply(&source).expect("replay margins patch");
    assert_eq!(
        replayed.to_plain_bytes().expect("replayed bytes"),
        commit.package().to_plain_bytes().expect("committed bytes")
    );
    assert_eq!(
        durable
            .inverse()
            .apply(&replayed)
            .expect("inverse margins patch")
            .to_plain_bytes()
            .expect("restored bytes"),
        source_bytes
    );

    let mut stale_edit = source.edit().expect("stale edit");
    stale_edit
        .sheet("Sheet1")
        .expect("lookup")
        .expect("sheet")
        .set("A1", "stale")
        .expect("stale cell");
    let stale = stale_edit.commit().expect("stale commit").into_package();
    assert!(matches!(
        durable.apply(&stale),
        Err(Error::PatchConflict { .. })
    ));

    let mut history = History::new(source, HistoryLimits::new(2, 256 * 1024 * 1024));
    assert!(
        commit
            .record(&mut history)
            .expect("record margins")
            .is_empty()
    );
    assert!(history.undo());
    assert!(history.redo());
    assert_eq!(
        history
            .current()
            .sheet("Sheet1")
            .expect("lookup")
            .expect("sheet")
            .page_margins()
            .expect("history margins"),
        Some(margins)
    );
}

#[test]
fn independently_prepared_disjoint_edits_join_after_threaded_work() {
    fn assert_send<T: Send>() {}
    assert_send::<Edit>();

    let source = crate::Package::new().expect("source workbook");
    let (mut left, right) = std::thread::scope(|scope| {
        let left_source = source.clone();
        let right_source = source.clone();
        let left = scope.spawn(move || {
            let mut edit = left_source.edit().expect("left edit");
            edit.sheet("Sheet1")
                .expect("lookup")
                .expect("sheet")
                .set("A1", "left")
                .expect("left cell");
            edit
        });
        let right = scope.spawn(move || {
            let mut edit = right_source.edit().expect("right edit");
            edit.sheet(0usize)
                .expect("lookup")
                .expect("sheet")
                .set("C3", 42_i32)
                .expect("right cell");
            edit
        });
        (
            left.join().expect("left worker"),
            right.join().expect("right worker"),
        )
    });

    left.join(right).expect("disjoint join");
    assert_eq!(left.len(), 2);
    let committed = left.commit().expect("joined commit");
    let sheet = committed
        .package()
        .sheet("Sheet1")
        .expect("lookup")
        .expect("sheet");
    assert!(matches!(
        sheet.cell("A1").expect("A1").stored(),
        Some(Cell::Value(Value::Text(text))) if text.as_str() == "left"
    ));
    assert!(matches!(
        sheet.cell("C3").expect("C3").stored(),
        Some(Cell::Value(Value::Number(number))) if number.as_str() == "42"
    ));

    let mut empty = source.edit().expect("empty edit");
    let mut incoming = source.edit().expect("incoming edit");
    incoming
        .sheet("Sheet1")
        .expect("lookup")
        .expect("sheet")
        .set("D4", true)
        .expect("incoming cell");
    empty.join(incoming).expect("adopt incoming sheet map");
    assert_eq!(empty.len(), 1);
}
#[test]
fn join_conflicts_are_structured_and_return_the_rejected_edit() {
    let source = crate::Package::new().expect("source workbook");
    let mut left = source.edit().expect("left edit");
    let mut left_sheet = left.sheet(0usize).expect("lookup").expect("sheet");
    left_sheet
        .set("C3", "left tail")
        .expect("left tail cell")
        .set("A1", "left")
        .expect("left first cell");
    let mut right = source.edit().expect("right edit");
    let mut right_sheet = right.sheet("Sheet1").expect("lookup").expect("sheet");
    right_sheet
        .set("A1", "right")
        .expect("right first cell")
        .set("C3", "right tail")
        .expect("right tail cell");

    let error = match left.join(right) {
        Ok(_) => panic!("overlapping edits must not join"),
        Err(error) => error,
    };
    assert_eq!(left.len(), 2);
    let conflicts = error.conflicts().expect("overlap details");
    assert_eq!(conflicts.len(), 2);
    assert_eq!(conflicts.conflicts().len(), 1);
    assert_eq!(conflicts.conflicts()[0].sheet(), "Sheet1");
    assert_eq!(conflicts.conflicts()[0].position(), 0);
    assert_eq!(
        conflicts.conflicts()[0].cells().expect("cell conflicts"),
        &[
            Address::from_a1("A1").expect("first address"),
            Address::from_a1("C3").expect("tail address"),
        ]
    );
    let rejected = error.into_rejected();
    assert_eq!(rejected.len(), 2);

    let other_source = crate::Package::new().expect("other source");
    let other = other_source.edit().expect("other edit");
    let error = match left.join(other) {
        Ok(_) => panic!("different snapshots must not join"),
        Err(error) => error,
    };
    assert!(matches!(error.failure(), JoinFailure::DifferentSnapshot));
    assert!(error.conflicts().is_none());
    assert_eq!(left.len(), 2);
}

#[test]
fn row_visibility_joins_with_cells_and_conflicts_by_row() {
    let source = crate::Package::new().expect("source workbook");
    let mut cell = source.edit().expect("cell edit");
    cell.sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .set("A2", "same row")
        .expect("cell");
    let mut row = source.edit().expect("row edit");
    row.sheet("Sheet1")
        .expect("lookup")
        .expect("sheet")
        .row(1)
        .expect("row 2")
        .hide();
    cell.join(row).expect("orthogonal row and cell effects");
    let committed = cell.commit().expect("joined commit");
    let sheet = committed
        .package()
        .sheet(0usize)
        .expect("lookup")
        .expect("sheet");
    assert!(sheet.row(1).expect("row 2").hidden());
    assert!(matches!(
        sheet.cell("A2").expect("A2").stored(),
        Some(Cell::Value(Value::Text(text))) if text.as_str() == "same row"
    ));

    let mut left = source.edit().expect("left");
    left.sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .row(4)
        .expect("row 5")
        .hide();
    let mut right = source.edit().expect("right");
    right
        .sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .row(4)
        .expect("row 5")
        .show();
    let error = left.join(right).expect_err("same row must conflict");
    let conflicts = error.conflicts().expect("row conflicts");
    assert_eq!(conflicts.len(), 1);
    assert_eq!(conflicts.conflicts().len(), 1);
    assert_eq!(
        conflicts.conflicts()[0].rows().expect("row conflict"),
        &[RowIndex::new(4).expect("row 5")]
    );
    assert!(conflicts.conflicts()[0].cells().is_none());
}
