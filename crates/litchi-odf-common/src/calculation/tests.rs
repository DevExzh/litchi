#![allow(
    clippy::unwrap_used,
    reason = "Test assertions intentionally unwrap known-valid fixture construction failures."
)]

use super::codec::{parse, write};
use super::model::{Iteration, IterationStatus, NullDate, Settings, is_xsd_date, is_xsd_double};
use std::num::NonZeroUsize;

#[test]
fn accepts_lossless_lexical_forms() {
    for value in ["0", ".5", "5.", "-1.25E-3", "INF", "-INF", "NaN"] {
        assert!(is_xsd_double(value), "{value}");
    }
    for value in ["", ".", "inf", "1E", "1 2"] {
        assert!(!is_xsd_double(value), "{value}");
    }
    for value in ["1899-12-30", "2026-07-14Z", "2026-07-14+08:00"] {
        assert!(is_xsd_date(value), "{value}");
    }
    for value in [
        "2026-07-14+14:01",
        "0000-01-01",
        "02026-01-01",
        "2026-02-29",
    ] {
        assert!(!is_xsd_date(value), "{value}");
    }
    assert!(is_xsd_date("2024-02-29"));
}

#[test]
fn writes_nested_settings_in_schema_order() {
    let settings = Settings {
        case_sensitive: Some(true),
        null_year: NonZeroUsize::new(1930),
        null_date: Some(NullDate {
            value_type_date: true,
            date_value: Some("1899-12-30Z".to_string()),
        }),
        iteration: Some(Iteration {
            status: Some(IterationStatus::Enable),
            steps: NonZeroUsize::new(100),
            maximum_difference: Some("1E-6".to_string()),
        }),
        ..Settings::default()
    };
    let mut xml = String::new();
    write(&mut xml, Some(&settings)).unwrap();
    assert!(xml.find("<table:null-date").unwrap() < xml.find("<table:iteration").unwrap());
}

#[test]
fn parses_all_settings_with_namespace_aliases() {
    let xml = r#"<o:document-content
          xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
          xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
          <o:body><o:spreadsheet><t:calculation-settings t:case-sensitive="1"
            t:precision-as-shown="false" t:search-criteria-must-apply-to-whole-cell="true"
            t:automatic-find-labels="false" t:use-regular-expressions="true"
            t:use-wildcards="false" t:null-year="1930">
            <t:null-date t:value-type="date" t:date-value="1899-12-30+08:00"></t:null-date>
            <t:iteration t:status="enable" t:steps="100" t:maximum-difference="NaN"/>
          </t:calculation-settings></o:spreadsheet></o:body>
        </o:document-content>"#;
    let settings = parse(xml).unwrap().unwrap();
    assert_eq!(settings.case_sensitive, Some(true));
    assert_eq!(settings.null_year.unwrap().get(), 1930);
    assert_eq!(
        settings.null_date.unwrap().date_value.as_deref(),
        Some("1899-12-30+08:00")
    );
    let iteration = settings.iteration.unwrap();
    assert_eq!(iteration.status, Some(IterationStatus::Enable));
    assert_eq!(iteration.steps.unwrap().get(), 100);
    assert_eq!(iteration.maximum_difference.as_deref(), Some("NaN"));
}

#[test]
fn rejects_invalid_or_out_of_order_settings() {
    let invalid = r#"<o:document-content
          xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
          xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
          <o:body><o:spreadsheet><t:calculation-settings t:null-year="0">
            <t:iteration t:steps="0"/><t:null-date t:date-value="not-a-date"/>
          </t:calculation-settings></o:spreadsheet></o:body>
        </o:document-content>"#;
    assert!(parse(invalid).is_err());
    let out_of_order = invalid
        .replace("t:null-year=\"0\"", "t:null-year=\"1930\"")
        .replace("t:steps=\"0\"", "t:steps=\"1\"")
        .replace("not-a-date", "1899-12-30");
    assert!(parse(&out_of_order).is_err());
    let duplicate = r#"<o:document-content
          xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
          xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
          <o:body><o:spreadsheet><t:calculation-settings/><t:calculation-settings/>
          </o:spreadsheet></o:body></o:document-content>"#;
    assert!(parse(duplicate).is_err());
    let nested = r#"<o:document-content
          xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
          xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
          <o:body><o:spreadsheet><t:table t:name="Bad"><t:table-row><t:table-cell>
            <t:calculation-settings/></t:table-cell></t:table-row></t:table>
          </o:spreadsheet></o:body></o:document-content>"#;
    assert!(parse(nested).is_err());
}

#[cfg(test)]
mod chart_document_tests {
    use super::*;

    #[test]
    fn parses_calculation_settings_from_chart_document_body() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content
 xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
 xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
 <office:body><office:chart>
  <table:calculation-settings table:case-sensitive="true" table:precision-as-shown="false"/>
 </office:chart></office:body>
</office:document-content>"#;
        let settings = parse(xml).unwrap().unwrap();
        assert_eq!(settings.case_sensitive, Some(true));
        assert_eq!(settings.precision_as_shown, Some(false));
    }

    #[test]
    fn rejects_calculation_settings_below_chart_chart() {
        let xml = r#"<office:document-content
 xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
 xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0"
 xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
 <office:body><office:chart><chart:chart>
  <table:calculation-settings/>
 </chart:chart></office:chart></office:body>
</office:document-content>"#;
        assert!(parse(xml).is_err());
    }
}
