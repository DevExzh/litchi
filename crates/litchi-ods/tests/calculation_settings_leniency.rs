//! LENIENCY_AUDIT F9.9: an unknown or foreign-namespace child of
//! `table:calculation-settings` is skipped with its subtree, and a
//! `table:null-date` after `table:iteration` is accepted, the way
//! LibreOffice's `ScXMLCalculationSettingsContext` reads them. Every ODS
//! open path records each repair; a conforming file records nothing; a
//! duplicate setting or an undeclared entity stays a hard error.

#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "hand-built fixtures keep tolerance assertions focused"
)]

#[path = "../../litchi-odf-common/tests/support/odf_rezip.rs"]
mod odf_rezip;

use std::sync::Arc;

use litchi_core::OwnedSource;
use litchi_odf_common::{ContentDefect, Defect, ToleranceReport};
use litchi_ods::{SourceBackedSpreadsheet, Spreadsheet};

const SETTINGS: &str = r#"<table:calculation-settings table:automatic-find-labels="false" table:use-regular-expressions="false" table:use-wildcards="true"/>"#;

fn recorded(report: &ToleranceReport) -> Vec<(Defect, u16)> {
    report
        .defects()
        .iter()
        .map(|tolerated| (tolerated.defect, tolerated.context))
        .collect()
}

fn conforming() -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/odf/corpus/calc-two-sheets.ods"),
    )
    .unwrap()
}

/// `calc-two-sheets.ods` with its empty calculation settings replaced by
/// `children` inside the same settings element.
fn with_settings_children(children: &str) -> Vec<u8> {
    odf_rezip::edit_member(&conforming(), "content.xml", |xml| {
        let open = SETTINGS.trim_end_matches("/>");
        xml.replacen(
            SETTINGS,
            &format!("{open}>{children}</table:calculation-settings>"),
            1,
        )
    })
}

fn assert_every_open_path(bytes: &[u8], defects: &[(Defect, u16)], label: &str) {
    let owned = Spreadsheet::from_bytes(bytes.to_vec()).unwrap();
    assert_eq!(
        recorded(owned.tolerance_report()),
        defects,
        "{label}: from_bytes"
    );
    assert_eq!(owned.sheet_names().len(), 2, "{label}");
    assert_eq!(
        owned
            .calculation_settings()
            .and_then(|settings| settings.use_wildcards),
        Some(true),
        "{label}: the remaining settings are kept"
    );

    let prepared = litchi_odf_common::detect::prepared(bytes.to_vec())
        .unwrap_or_else(|| panic!("{label}: package must prepare"));
    let adopted = Spreadsheet::from_prepared_package(prepared).unwrap();
    assert_eq!(
        recorded(adopted.tolerance_report()),
        defects,
        "{label}: prepared"
    );

    let source =
        SourceBackedSpreadsheet::from_read_at(Arc::new(OwnedSource::new(bytes.to_vec()))).unwrap();
    assert_eq!(
        recorded(source.tolerance_report()),
        defects,
        "{label}: source-backed"
    );
}

const UNKNOWN_FOREIGN: (Defect, u16) = (
    Defect::Content(ContentDefect::CalculationSettingsUnknownChild),
    0,
);
const UNKNOWN_TABLE: (Defect, u16) = (
    Defect::Content(ContentDefect::CalculationSettingsUnknownChild),
    1,
);
const REORDERED: (Defect, u16) = (
    Defect::Content(ContentDefect::CalculationSettingsReordered),
    0,
);

#[test]
fn foreign_and_unknown_children_are_skipped_and_recorded() {
    let bytes = with_settings_children(concat!(
        r#"<table:null-date table:date-value="1899-12-30"/>"#,
        r#"<loext:vendor-setting loext:mode="fast"><loext:nested>a &amp; b<loext:deeper/></loext:nested></loext:vendor-setting>"#,
        r#"<table:future-setting/>"#,
        r#"<table:iteration table:status="enable" table:steps="50"/>"#,
    ));
    assert_every_open_path(
        &bytes,
        &[UNKNOWN_FOREIGN, UNKNOWN_TABLE],
        "unknown children",
    );

    let owned = Spreadsheet::from_bytes(bytes).unwrap();
    let settings = owned.calculation_settings().unwrap();
    assert_eq!(
        settings
            .null_date
            .as_ref()
            .and_then(|null_date| null_date.date_value.as_deref()),
        Some("1899-12-30")
    );
    assert_eq!(
        settings
            .iteration
            .as_ref()
            .and_then(|iteration| iteration.steps)
            .map(std::num::NonZeroUsize::get),
        Some(50)
    );
}

#[test]
fn null_date_after_iteration_is_accepted_and_recorded() {
    let bytes = with_settings_children(concat!(
        r#"<table:iteration table:steps="7"/>"#,
        r#"<table:null-date table:date-value="1904-01-01"/>"#,
    ));
    assert_every_open_path(&bytes, &[REORDERED], "reordered");
    let owned = Spreadsheet::from_bytes(bytes).unwrap();
    assert!(owned.calculation_settings().unwrap().null_date.is_some());
}

#[test]
fn the_same_settings_in_schema_order_repair_nothing() {
    assert_every_open_path(&conforming(), &[], "original");
    let bytes = with_settings_children(concat!(
        r#"<table:null-date table:date-value="1899-12-30"/>"#,
        r#"<table:iteration table:status="enable" table:steps="50"/>"#,
    ));
    assert_every_open_path(&bytes, &[], "schema order");
}

#[test]
fn duplicates_and_undeclared_entities_stay_hard_errors() {
    let duplicate = with_settings_children(concat!(
        r#"<table:iteration table:steps="7"/>"#,
        r#"<table:iteration table:steps="8"/>"#,
    ));
    assert!(Spreadsheet::from_bytes(duplicate).is_err());

    let entity =
        with_settings_children(r#"<loext:vendor-setting>&undeclared;</loext:vendor-setting>"#);
    assert!(Spreadsheet::from_bytes(entity.clone()).is_err());
    assert!(SourceBackedSpreadsheet::from_read_at(Arc::new(OwnedSource::new(entity))).is_err());
}

#[test]
fn the_standalone_parser_reports_the_same_repairs() {
    let xml = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:x="urn:example:vendor"><office:body><office:spreadsheet><table:calculation-settings><table:iteration/><x:extra><x:inner/></x:extra><table:null-date/></table:calculation-settings></office:spreadsheet></office:body></office:document-content>"#;
    let (settings, report) = litchi_odf_common::calculation::parse_with_report(xml).unwrap();
    let settings = settings.unwrap();
    assert!(settings.iteration.is_some() && settings.null_date.is_some());
    assert_eq!(recorded(&report), [UNKNOWN_FOREIGN, REORDERED]);
    assert!(litchi_odf_common::calculation::parse(xml).is_ok());
}

#[test]
fn a_skipped_child_is_never_silently_dropped_by_an_edit() {
    let bytes = with_settings_children(r#"<loext:vendor-setting loext:mode="fast"/>"#);
    let owned = Spreadsheet::from_bytes(bytes).unwrap();
    let snapshot = litchi_ods::settings::Snapshot::from_content_xml(owned.content_xml()).unwrap();

    // An unchanged commit borrows the source and keeps the vendor child.
    let unchanged = snapshot.transaction().commit().unwrap();
    assert!(unchanged.content_xml().contains("loext:vendor-setting"));

    // A semantic rewrite refuses rather than dropping the skipped child.
    let mut transaction = snapshot.transaction();
    let mut settings = transaction.calculation().cloned().unwrap();
    settings.case_sensitive = Some(true);
    transaction.replace(settings).unwrap();
    assert!(transaction.commit().is_err());
}
