use std::sync::Arc;

use litchi_core::OwnedSource;
use litchi_ods::{Builder, SourceBackedSpreadsheet, Spreadsheet, scenario};

const CONTENT: &str = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" office:version="1.4"><office:body><office:spreadsheet><table:table table:name="One"><table:title>Keep title</table:title><table:scenario table:scenario-ranges=".A1:.B2" table:is-active="true" table:comment="Original"/><table:table-row/></table:table><table:table table:name="Two"><table:desc>Keep description</table:desc><table:table-row/></table:table></office:spreadsheet></office:body></office:document-content>"#;

fn range(value: &str) -> scenario::RangeAddress {
    scenario::RangeAddress::new(value).expect("valid range")
}

fn spreadsheet() -> Spreadsheet {
    Spreadsheet::from_bytes(
        Builder::new()
            .content_xml(CONTENT)
            .build()
            .expect("build content"),
    )
    .expect("open content")
}

#[test]
fn scenario_noop_inverse_and_stale_source_are_exact() {
    let before = scenario::Snapshot::parse(CONTENT).expect("parse scenario source");
    let noop = before.edit().commit().expect("no-op commit");
    assert!(!noop.changed());
    assert!(noop.patch().is_empty());
    assert_eq!(noop.snapshot().source_xml(), CONTENT);

    let mut edit = before.edit();
    edit.replace_sheet(
        "One",
        scenario::Scenario::new("One", vec![range(".C1:.D2")], scenario::State::Inactive)
            .expect("scenario")
            .with_comment("Edited"),
    )
    .expect("replace");
    let commit = edit.commit().expect("commit");
    assert!(commit.changed());
    assert!(commit.snapshot().source_xml().contains(".C1:.D2"));
    let restored = commit
        .patch()
        .inverse()
        .apply(commit.snapshot())
        .expect("inverse");
    assert_eq!(restored.snapshot().source_xml(), CONTENT);

    let other =
        scenario::Snapshot::parse(&CONTENT.replace("Keep title", "Other")).expect("other source");
    assert!(commit.patch().apply(&other).is_err());
}

#[test]
fn owned_facade_scenario_edit_is_atomic_and_reopenable() {
    let mut spreadsheet = spreadsheet();
    let original = spreadsheet.content_xml().to_owned();
    spreadsheet
        .edit_scenarios(|edit| {
            edit.replace_sheet(
                "One",
                scenario::Scenario::new("One", vec![range(".C1:.D2")], scenario::State::Inactive)
                    .expect("scenario")
                    .with_display_border(scenario::OptionalSetting::Disabled)
                    .with_comment("Edited & retained"),
            )?;
            edit.add(
                scenario::Scenario::new("Two", vec![range(".E1:.F2")], scenario::State::Active)
                    .expect("scenario"),
            )?;
            Ok(())
        })
        .expect("edit scenarios");

    let target = spreadsheet.content_xml().to_owned();
    assert!(target.contains("Keep title"));
    assert!(target.contains("Keep description"));
    assert!(target.contains("Edited &amp; retained"));
    assert_eq!(
        spreadsheet.scenarios().expect("readback").scenarios().len(),
        2
    );

    let bytes = spreadsheet.into_bytes();
    let reopened = Spreadsheet::from_bytes(bytes.clone()).expect("reopen");
    assert_eq!(
        reopened.scenarios().expect("scenarios").scenarios().len(),
        2
    );

    let mut failed = reopened;
    assert!(
        failed
            .edit_scenarios(|edit| {
                edit.replace_sheet(
                    "One",
                    scenario::Scenario::new("One", vec![range(".G1:.H2")], scenario::State::Active)
                        .expect("scenario"),
                )?;
                Err(litchi_core::Error::InvalidFormat("stop".to_string()))
            })
            .is_err()
    );
    assert_eq!(failed.content_xml(), target);
    assert_ne!(original, target);
}

#[test]
fn source_backed_scenarios_check_source_identity() {
    let bytes = Builder::new()
        .content_xml(CONTENT)
        .build()
        .expect("build content");
    let source = SourceBackedSpreadsheet::from_read_at(Arc::new(OwnedSource::new(bytes)))
        .expect("source open");
    let snapshot = source.scenarios().expect("source scenarios");
    assert_eq!(snapshot.scenarios().len(), 1);
    assert_eq!(snapshot.scenarios()[0].sheet(), "One");
}

#[test]
fn builder_scenario_lifecycle_reopens_typed_metadata() {
    let mut builder = Builder::new().content_xml(CONTENT);
    builder
        .edit_scenarios(|edit| edit.remove_named("One").map(|_| ()))
        .expect("builder edit");
    let bytes = builder.build().expect("build edited package");
    let spreadsheet = Spreadsheet::from_bytes(bytes).expect("reopen edited package");
    assert!(
        spreadsheet
            .scenarios()
            .expect("scenario metadata")
            .scenarios()
            .is_empty()
    );
}

#[test]
fn changed_scenario_refuses_unknown_metadata_loss() {
    let source = CONTENT.replace(
        "table:comment=\"Original\"",
        "table:comment=\"Original\" xmlns:vendor=\"urn:vendor\" vendor:opaque=\"keep\"",
    );
    let snapshot = scenario::Snapshot::parse(&source).expect("parse source");
    let mut edit = snapshot.edit();
    edit.replace_sheet(
        "One",
        scenario::Scenario::new("One", vec![range(".C1:.D2")], scenario::State::Active)
            .expect("scenario"),
    )
    .expect("stage");
    assert!(edit.commit().is_err());
}

#[test]
fn removed_scenario_refuses_unknown_metadata_loss() {
    let source = CONTENT.replace(
        "table:comment=\"Original\"",
        "table:comment=\"Original\" xmlns:vendor=\"urn:vendor\" vendor:opaque=\"keep\"",
    );
    let snapshot = scenario::Snapshot::parse(&source).expect("parse source");
    let mut edit = snapshot.edit();
    edit.remove_named("One").expect("stage removal");
    assert!(edit.commit().is_err());
}

#[test]
fn changed_scenario_preserves_unrelated_document_extensions() {
    let source = CONTENT.replace(
        "</office:document-content>",
        "<vendor:extension xmlns:vendor=\"urn:vendor\">keep</vendor:extension></office:document-content>",
    );
    let snapshot = scenario::Snapshot::parse(&source).expect("parse source");
    let mut edit = snapshot.edit();
    edit.replace_named(
        "One",
        scenario::Scenario::new("One", vec![range(".C1:.D2")], scenario::State::Active)
            .expect("scenario"),
    )
    .expect("stage");
    let commit = edit.commit().expect("commit");
    assert!(commit.snapshot().source_xml().contains("vendor:extension"));
    assert!(
        commit
            .snapshot()
            .source_xml()
            .contains(">keep</vendor:extension>")
    );
}

#[test]
fn changed_scenario_refuses_processing_instruction_markup() {
    let source = CONTENT.replace(
        "table:comment=\"Original\"/>",
        "table:comment=\"Original\"><!-- whitespace is not a scenario value --><?vendor keep?></table:scenario>",
    );
    let snapshot = scenario::Snapshot::parse(&source).expect("parse scenario source");
    let mut edit = snapshot.edit();
    edit.replace_named(
        "One",
        scenario::Scenario::new("One", vec![range(".C1:.D2")], scenario::State::Active)
            .expect("scenario"),
    )
    .expect("stage");
    assert!(edit.commit().is_err());
}

#[test]
fn scenario_edit_uses_the_parent_table_namespace_alias() {
    let mut source = CONTENT.replace(
        "xmlns:table=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\"",
        "xmlns:t=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\"",
    );
    for name in [
        "table:table",
        "table:name",
        "table:title",
        "table:scenario",
        "table:scenario-ranges",
        "table:is-active",
        "table:comment",
        "table:desc",
        "table:table-row",
    ] {
        source = source.replace(name, &name.replacen("table", "t", 1));
    }
    let snapshot = scenario::Snapshot::parse(&source).expect("parse aliased source");
    let mut edit = snapshot.edit();
    edit.replace_named(
        "One",
        scenario::Scenario::new("One", vec![range(".C1:.D2")], scenario::State::Inactive)
            .expect("scenario"),
    )
    .expect("stage");
    let commit = edit.commit().expect("commit aliased source");
    assert!(commit.snapshot().source_xml().contains("<t:scenario"));
    assert_eq!(
        commit.snapshot().scenarios()[0].ranges()[0].as_str(),
        ".C1:.D2"
    );
}

#[test]
fn scenario_commit_retains_the_snapshot_input_budget() {
    let limits = scenario::Limits::default().with_input_bytes(CONTENT.len());
    let before = scenario::Snapshot::parse_with(CONTENT, limits).expect("bounded source");
    let mut edit = before.edit();
    edit.replace_named(
        "One",
        scenario::Scenario::new("One", vec![range(".C1:.D2")], scenario::State::Active)
            .expect("scenario")
            .with_comment("A replacement comment that expands the source"),
    )
    .expect("stage");
    assert!(edit.commit().is_err());
}

#[test]
fn scenario_commit_preflights_exact_and_over_output_budgets() {
    let replacement = || {
        scenario::Scenario::new("One", vec![range(".C1:.D2")], scenario::State::Inactive)
            .expect("scenario")
            .with_comment("Edited & retained")
    };
    let mut unrestricted = scenario::Snapshot::parse(CONTENT)
        .expect("parse source")
        .edit();
    unrestricted
        .replace_named("One", replacement())
        .expect("stage replacement");
    let expected = unrestricted.commit().expect("unrestricted commit");
    let target_len = expected.snapshot().source_xml().len();

    let exact_limits = scenario::Limits::default().with_input_bytes(target_len);
    let mut exact = scenario::Snapshot::parse_with(CONTENT, exact_limits)
        .expect("parse exact-budget source")
        .edit();
    exact
        .replace_named("One", replacement())
        .expect("stage exact replacement");
    let exact_commit = exact.commit().expect("exact output budget should pass");
    assert_eq!(exact_commit.snapshot().source_xml().len(), target_len);

    let over_limits = scenario::Limits::default().with_input_bytes(target_len - 1);
    let before_over = scenario::Snapshot::parse_with(CONTENT, over_limits)
        .expect("parse over-budget source")
        .edit();
    let mut over = before_over;
    over.replace_named("One", replacement())
        .expect("stage over-budget replacement");
    let before_over_xml = over.before().source_xml().to_owned();
    assert!(over.commit().is_err());
    assert_eq!(before_over_xml, CONTENT);
}

#[test]
fn scenario_add_precharges_aggregate_and_count_limits() {
    let aggregate_limits = scenario::Limits::default().with_aggregate_bytes(18);
    let before = scenario::Snapshot::parse_with(CONTENT, aggregate_limits)
        .expect("source aggregate is exactly eighteen bytes");
    let mut aggregate_edit = before.edit();
    assert!(
        aggregate_edit
            .add(
                scenario::Scenario::new("Two", vec![range(".E1:.F2")], scenario::State::Active)
                    .expect("scenario")
            )
            .is_err()
    );
    assert_eq!(aggregate_edit.scenarios().len(), 1);

    let count_limits = scenario::Limits::default().with_scenarios(1);
    let before_count = scenario::Snapshot::parse_with(CONTENT, count_limits)
        .expect("source count is exactly one scenario");
    let mut count_edit = before_count.edit();
    assert!(
        count_edit
            .add(
                scenario::Scenario::new("Two", vec![range(".E1:.F2")], scenario::State::Active)
                    .expect("scenario")
            )
            .is_err()
    );
    assert_eq!(count_edit.scenarios().len(), 1);
}

#[test]
fn scenario_output_preflight_matches_all_typed_attributes() {
    let replacement = scenario::Scenario::new(
        "One",
        vec![range(".C1:.D2"), range(".E3:.F4")],
        scenario::State::Inactive,
    )
    .expect("scenario")
    .with_display_border(scenario::OptionalSetting::Enabled)
    .with_border_color(Some(scenario::RgbColor::new(0x12, 0xAB, 0xF0)))
    .with_copy_back(scenario::OptionalSetting::Disabled)
    .with_copy_styles(scenario::OptionalSetting::Enabled)
    .with_copy_formulas(scenario::OptionalSetting::Disabled)
    .with_comment("<&>\"'")
    .with_protected(scenario::OptionalSetting::Enabled);
    let mut edit = scenario::Snapshot::parse(CONTENT).expect("parse").edit();
    edit.replace_named("One", replacement).expect("stage");
    let commit = edit.commit().expect("commit");
    assert!(commit.snapshot().source_xml().contains("#12ABF0"));
    assert!(
        commit
            .snapshot()
            .source_xml()
            .contains("&lt;&amp;&gt;&quot;&apos;")
    );
    assert_eq!(commit.snapshot().scenarios()[0].ranges().len(), 2);
}

#[test]
fn whitespace_only_range_and_nameless_table_without_scenario_are_rejected_or_accepted_by_grammar() {
    assert!(scenario::RangeAddress::new(" \t\n").is_err());

    let nameless = r#"<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><office:body><office:spreadsheet><table:table><table:table-row/></table:table></office:spreadsheet></office:body></office:document-content>"#;
    let snapshot = scenario::Snapshot::parse(nameless).expect("nameless table has no scenario");
    assert!(snapshot.scenarios().is_empty());

    let nameless_scenario = nameless.replace(
        "<table:table><table:table-row/>",
        "<table:table><table:scenario table:scenario-ranges=\".A1\" table:is-active=\"true\"/>",
    );
    assert!(scenario::Snapshot::parse(&nameless_scenario).is_err());
}

#[test]
fn range_addresses_follow_the_odf_cell_range_address_alternatives() {
    for value in [
        ".A1",
        ".A1:.B2",
        ".$A$1",
        "$Sheet.$A$1",
        "$.A1",
        "$.A1:$.A1",
        "$.1:$.1",
        "$.A:$.A",
        "$.$A$1",
        "'Q1 Sales'.$C$3:.$D$4",
        "Sheet.A1:Other.B2",
        ".1:.3",
        "$Sheet.$1:$Other.$3",
        ".A:.C",
        "'$A''B'.$A:'Q1 Sales'.$C",
        ".A0",
    ] {
        assert!(
            scenario::RangeAddress::new(value).is_ok(),
            "rejected valid address {value}"
        );
    }
    for value in [
        "bogus",
        "123",
        ".a1",
        ".A",
        ".1",
        ".A1:.B",
        ".A1:.1",
        ".A:.1",
        ".1:.A",
        "Sheet A.A1",
        "'' .A1",
        "'Unclosed.A1",
        ".A1:.B2:",
        ".A1 .B2",
        " .A1",
        ".A1 ",
        "\t.A1",
    ] {
        assert!(
            scenario::RangeAddress::new(value).is_err(),
            "accepted invalid address {value}"
        );
    }

    let valid_source = CONTENT.replace(".A1:.B2", "'Q1 Sales'.$C$3:.$D$4");
    let parsed = scenario::Snapshot::parse(&valid_source).expect("valid source address");
    assert_eq!(
        parsed.scenarios()[0].ranges()[0].as_str(),
        "'Q1 Sales'.$C$3:.$D$4"
    );
    let dollar_sheet_source = CONTENT.replace(".A1:.B2", "$.A1:$.A1");
    let dollar_sheet = scenario::Snapshot::parse(&dollar_sheet_source)
        .expect("a dollar-named sheet address is valid source metadata");
    assert_eq!(
        dollar_sheet.scenarios()[0].ranges()[0].as_str(),
        "$.A1:$.A1"
    );
    let invalid_source = CONTENT.replace(".A1:.B2", "bogus");
    assert!(scenario::Snapshot::parse(&invalid_source).is_err());
}

#[test]
fn scenario_count_is_precharged_before_malformed_second_entry_is_parsed() {
    let source = r#"<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><office:body><office:spreadsheet><table:table table:name="One"><table:scenario table:scenario-ranges=".A1" table:is-active="true"/></table:table><table:table table:name="Two"><table:scenario table:is-active="true"/></table:table></office:spreadsheet></office:body></office:document-content>"#;
    let error =
        scenario::Snapshot::parse_with(source, scenario::Limits::default().with_scenarios(1))
            .expect_err("second scenario exceeds the precharged count");
    assert!(matches!(
        error,
        scenario::Error::ResourceLimit {
            resource: "scenarios",
            actual: 2,
            maximum: 1,
        }
    ));
}
