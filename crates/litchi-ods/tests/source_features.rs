use litchi_ods::{
    Builder, Spreadsheet,
    dde::{AutomaticUpdate, ConversionMode},
    scenario::OptionalSetting,
};

const CONTENT: &str = r##"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" office:version="1.4"><office:body><office:spreadsheet><table:table table:name="Base"><office:dde-source office:dde-application="soffice" office:dde-topic="file:///never-contacted.ods" office:dde-item="Sheet1.A1:B2" office:name="Reference" office:conversion-mode="keep-text" office:automatic-update="true"/><table:table-row><table:table-cell office:value-type="string"><text:p>Pre <text:span text:style-name="Bold">styled <text:a xlink:href="https://never-fetched.invalid/" xlink:type="simple">link</text:a></text:span> tail</text:p></table:table-cell></table:table-row></table:table><table:table table:name="Scenario"><table:scenario table:scenario-ranges="$Scenario.$A$1:$B$2 'Q1 Sales'.$C$3:$D$4" table:is-active="true" table:display-border="false" table:border-color="#12AbEF" table:copy-back="true" table:copy-styles="false" table:copy-formulas="true" table:comment="Best &amp; worst" table:protected="false"/><table:table-row/></table:table><table:dde-links><table:dde-link><office:dde-source office:dde-application="calc" office:dde-topic="file:///never-opened.ods" office:dde-item="Prices.A1"/><table:table><table:table-row><table:table-cell office:value-type="float" office:value="42"/></table:table-row></table:table></table:dde-link></table:dde-links></office:spreadsheet></office:body></office:document-content>"##;

fn spreadsheet() -> Spreadsheet {
    Spreadsheet::from_bytes(Builder::new().content_xml(CONTENT).build().unwrap()).unwrap()
}

#[test]
fn inspects_dde_sources_and_cache_without_contacting_them() {
    let spreadsheet = spreadsheet();
    let snapshot = spreadsheet.dde().unwrap();
    assert_eq!(snapshot.source_xml(), CONTENT);
    assert_eq!(snapshot.sheet_sources().len(), 1);
    let source = snapshot.sheet_sources()[0].source();
    assert_eq!(snapshot.sheet_sources()[0].sheet(), "Base");
    assert_eq!(source.application(), "soffice");
    assert_eq!(source.conversion_mode(), ConversionMode::KeepText);
    assert_eq!(source.automatic_update(), AutomaticUpdate::Enabled);
    assert_eq!(snapshot.links().len(), 1);
    assert_eq!(snapshot.links()[0].source().application(), "calc");
    assert_eq!(
        snapshot.links()[0].cached_table_xml(),
        "<table:table><table:table-row><table:table-cell office:value-type=\"float\" office:value=\"42\"/></table:table-row></table:table>"
    );
}

#[test]
fn inspects_scenarios_without_applying_values() {
    let spreadsheet = spreadsheet();
    let snapshot = spreadsheet.scenarios().unwrap();
    assert_eq!(snapshot.source_xml(), CONTENT);
    assert_eq!(snapshot.scenarios().len(), 1);
    let scenario = &snapshot.scenarios()[0];
    assert_eq!(scenario.sheet(), "Scenario");
    assert_eq!(
        scenario
            .ranges()
            .iter()
            .map(litchi_ods::scenario::RangeAddress::as_str)
            .collect::<Vec<_>>(),
        ["$Scenario.$A$1:$B$2", "'Q1 Sales'.$C$3:$D$4"]
    );
    assert!(scenario.is_active());
    assert_eq!(scenario.display_border(), OptionalSetting::Disabled);
    assert_eq!(scenario.comment(), Some("Best & worst"));
}

#[test]
fn detached_metadata_values_reject_invalid_states() {
    assert!(litchi_ods::dde::Source::new("", "topic", "item").is_err());
    let source = litchi_ods::dde::Source::new("calc", "topic", "item")
        .unwrap()
        .with_automatic_update(AutomaticUpdate::Disabled);
    assert_eq!(source.application(), "calc");
    assert_eq!(source.automatic_update(), AutomaticUpdate::Disabled);

    assert!(litchi_ods::scenario::RangeAddress::new(".A1 .B2").is_err());
    assert!(litchi_ods::scenario::RgbColor::from_hex("#12345Z").is_err());
    let range = litchi_ods::scenario::RangeAddress::new(".A1:.B2").unwrap();
    let scenario = litchi_ods::scenario::Scenario::new(
        "Sheet1",
        vec![range],
        litchi_ods::scenario::State::Inactive,
    )
    .unwrap();
    assert!(!scenario.is_active());
    assert_eq!(scenario.display_border(), OptionalSetting::Unspecified);
}

#[test]
fn no_op_package_round_trip_preserves_compact_rich_text_xml_exactly() {
    let spreadsheet = spreadsheet();
    assert_eq!(spreadsheet.content_xml(), CONTENT);
    assert_eq!(
        spreadsheet.sheets()[0].rows[0].cells[0].text,
        "Pre styled link tail"
    );
    let reopened = Spreadsheet::from_bytes(spreadsheet.into_bytes()).unwrap();
    assert_eq!(reopened.content_xml(), CONTENT);
    assert!(!reopened.content_xml().contains("\n  "));
}

#[test]
fn bounded_owners_reject_malformed_or_oversized_metadata() {
    let malformed = CONTENT.replace("office:dde-item=\"Sheet1.A1:B2\" ", "");
    let spreadsheet =
        Spreadsheet::from_bytes(Builder::new().content_xml(malformed).build().unwrap()).unwrap();
    assert!(spreadsheet.dde().is_err());

    let malformed = CONTENT.replace(" table:is-active=\"true\"", "");
    let spreadsheet =
        Spreadsheet::from_bytes(Builder::new().content_xml(malformed).build().unwrap()).unwrap();
    assert!(spreadsheet.scenarios().is_err());

    let limits = litchi_ods::dde::Limits::default().with_input_bytes(CONTENT.len() - 1);
    assert!(litchi_ods::dde::Snapshot::parse_with(CONTENT, limits).is_err());

    let malformed = CONTENT.replace(
        "office:automatic-update=\"true\"/>",
        "office:automatic-update=\"true\"><table:table-row/></office:dde-source>",
    );
    assert!(litchi_ods::dde::Snapshot::parse(&malformed).is_err());

    let malformed = CONTENT.replace(
        " table:protected=\"false\"/>",
        " table:protected=\"false\"><table:table-row/></table:scenario>",
    );
    assert!(litchi_ods::scenario::Snapshot::parse(&malformed).is_err());

    let malformed = CONTENT.replace("</table:dde-link>", "unexpected</table:dde-link>");
    assert!(litchi_ods::dde::Snapshot::parse(&malformed).is_err());
}

#[test]
fn self_closing_dde_cache_obeys_the_exact_byte_quota() {
    const PREFIX: &str = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:tt="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><office:body><office:spreadsheet><table:dde-links><table:dde-link><office:dde-source office:dde-application="app" office:dde-topic="topic" office:dde-item="item"/>"#;
    const SUFFIX: &str = "</table:dde-link></table:dde-links></office:spreadsheet></office:body></office:document-content>";
    const EXACT: &str = "<t:table/>";
    const ONE_OVER: &str = "<tt:table/>";
    assert_eq!(ONE_OVER.len(), EXACT.len() + 1);

    let limits = litchi_ods::dde::Limits::default().with_cached_table_bytes(EXACT.len());
    let exact = format!("{PREFIX}{EXACT}{SUFFIX}");
    assert!(litchi_ods::dde::Snapshot::parse_with(&exact, limits).is_ok());
    let one_over = format!("{PREFIX}{ONE_OVER}{SUFFIX}");
    assert!(litchi_ods::dde::Snapshot::parse_with(&one_over, limits).is_err());
}

#[test]
fn scenario_rejects_entity_content() {
    let scenario = r#"<table:scenario table:scenario-ranges=".A1:.B2" table:is-active="true">&scenario;</table:scenario>"#;
    let malformed = CONTENT.replacen(
        r##"<table:scenario table:scenario-ranges="$Scenario.$A$1:$B$2 'Q1 Sales'.$C$3:$D$4" table:is-active="true" table:display-border="false" table:border-color="#12AbEF" table:copy-back="true" table:copy-styles="false" table:copy-formulas="true" table:comment="Best &amp; worst" table:protected="false"/>"##,
        scenario,
        1,
    );
    assert!(litchi_ods::scenario::Snapshot::parse(&malformed).is_err());
}

#[test]
fn custom_limits_cannot_exceed_hard_ceilings() {
    for limits in [
        litchi_ods::dde::Limits::default().with_input_bytes(usize::MAX),
        litchi_ods::dde::Limits::default().with_links(usize::MAX),
        litchi_ods::dde::Limits::default().with_sheet_sources(usize::MAX),
        litchi_ods::dde::Limits::default().with_text_bytes(usize::MAX),
        litchi_ods::dde::Limits::default().with_cached_table_bytes(usize::MAX),
        litchi_ods::dde::Limits::default().with_depth(usize::MAX),
    ] {
        assert!(litchi_ods::dde::Snapshot::parse_with(CONTENT, limits).is_err());
    }
    for limits in [
        litchi_ods::scenario::Limits::default().with_input_bytes(usize::MAX),
        litchi_ods::scenario::Limits::default().with_scenarios(usize::MAX),
        litchi_ods::scenario::Limits::default().with_ranges(usize::MAX),
        litchi_ods::scenario::Limits::default().with_text_bytes(usize::MAX),
        litchi_ods::scenario::Limits::default().with_range_list_bytes(usize::MAX),
        litchi_ods::scenario::Limits::default().with_depth(usize::MAX),
    ] {
        assert!(litchi_ods::scenario::Snapshot::parse_with(CONTENT, limits).is_err());
    }
}

#[test]
fn scenario_range_preflight_enforces_exact_count_and_byte_limits() {
    const PREFIX: &str = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><office:body><office:spreadsheet><table:table table:name="S"><table:scenario table:scenario-ranges=""#;
    const SUFFIX: &str = r#"" table:is-active="true"/></table:table></office:spreadsheet></office:body></office:document-content>"#;

    let exact_count = format!("{PREFIX}.A1 .B2{SUFFIX}");
    let count_limits = litchi_ods::scenario::Limits::default().with_ranges(2);
    assert!(litchi_ods::scenario::Snapshot::parse_with(&exact_count, count_limits).is_ok());
    let one_over_count = format!("{PREFIX}.A1 .B2 .C3{SUFFIX}");
    assert!(litchi_ods::scenario::Snapshot::parse_with(&one_over_count, count_limits).is_err());

    let exact_bytes = format!("{PREFIX}.A1{SUFFIX}");
    let byte_limits = litchi_ods::scenario::Limits::default().with_range_list_bytes(3);
    assert!(litchi_ods::scenario::Snapshot::parse_with(&exact_bytes, byte_limits).is_ok());
    let one_over_bytes = format!("{PREFIX}.A11{SUFFIX}");
    assert!(litchi_ods::scenario::Snapshot::parse_with(&one_over_bytes, byte_limits).is_err());
}

#[test]
fn paired_dde_cache_obeys_the_exact_byte_quota() {
    const PREFIX: &str = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><office:body><office:spreadsheet><table:dde-links><table:dde-link><office:dde-source office:dde-application="app" office:dde-topic="topic" office:dde-item="item"/>"#;
    const SUFFIX: &str = "</table:dde-link></table:dde-links></office:spreadsheet></office:body></office:document-content>";
    const EXACT: &str = "<table:table><!--x--></table:table>";
    const ONE_OVER: &str = "<table:table><!--xx--></table:table>";
    assert_eq!(ONE_OVER.len(), EXACT.len() + 1);
    let limits = litchi_ods::dde::Limits::default().with_cached_table_bytes(EXACT.len());
    assert!(
        litchi_ods::dde::Snapshot::parse_with(&format!("{PREFIX}{EXACT}{SUFFIX}"), limits).is_ok()
    );
    assert!(
        litchi_ods::dde::Snapshot::parse_with(&format!("{PREFIX}{ONE_OVER}{SUFFIX}"), limits,)
            .is_err()
    );
}
