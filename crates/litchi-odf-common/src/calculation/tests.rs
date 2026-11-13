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
    let settings = parse(xml, crate::XmlStreamLimits::default())
        .unwrap()
        .unwrap();
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
fn rejects_invalid_settings_and_records_out_of_order_ones() {
    let invalid = r#"<o:document-content
          xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
          xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
          <o:body><o:spreadsheet><t:calculation-settings t:null-year="0">
            <t:iteration t:steps="0"/><t:null-date t:date-value="not-a-date"/>
          </t:calculation-settings></o:spreadsheet></o:body>
        </o:document-content>"#;
    assert!(parse(invalid, crate::XmlStreamLimits::default()).is_err());
    let out_of_order = invalid
        .replace("t:null-year=\"0\"", "t:null-year=\"1930\"")
        .replace("t:steps=\"0\"", "t:steps=\"1\"")
        .replace("not-a-date", "1899-12-30");
    // LENIENCY_AUDIT F9.9: LibreOffice reads either child order, so the
    // reordered settings are accepted and the repair is recorded.
    let (settings, report) =
        super::codec::parse_with_report(&out_of_order, crate::XmlStreamLimits::default()).unwrap();
    let settings = settings.unwrap();
    assert!(settings.null_date.is_some() && settings.iteration.is_some());
    assert_eq!(report.defects().len(), 1);
    assert_eq!(
        report.defects()[0].defect,
        crate::leniency::Defect::Content(
            crate::leniency::ContentDefect::CalculationSettingsReordered
        )
    );
    let duplicate_child = out_of_order.replace("<t:iteration t:steps=\"1\"/>", "<t:null-date/>");
    assert!(parse(&duplicate_child, crate::XmlStreamLimits::default()).is_err());
    let duplicate = r#"<o:document-content
          xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
          xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
          <o:body><o:spreadsheet><t:calculation-settings/><t:calculation-settings/>
          </o:spreadsheet></o:body></o:document-content>"#;
    assert!(parse(duplicate, crate::XmlStreamLimits::default()).is_err());
    let nested = r#"<o:document-content
          xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
          xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
          <o:body><o:spreadsheet><t:table t:name="Bad"><t:table-row><t:table-cell>
            <t:calculation-settings/></t:table-cell></t:table-row></t:table>
          </o:spreadsheet></o:body></o:document-content>"#;
    assert!(parse(nested, crate::XmlStreamLimits::default()).is_err());
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
        let settings = parse(xml, crate::XmlStreamLimits::default())
            .unwrap()
            .unwrap();
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
        assert!(parse(xml, crate::XmlStreamLimits::default()).is_err());
    }
}

#[cfg(test)]
mod handler_tests {
    use super::super::codec::{CalculationHandler, classify, validate_size};
    use super::parse;
    use litchi_core::{Error, Result};
    use quick_xml::events::Event;
    use quick_xml::reader::NsReader;

    /// Drive the fused-parser handler over one tokenization, mirroring the
    /// family owners' driver loop (classification immediately after the
    /// read, Eof dispatched before the break).
    fn handler_parse(xml: &str) -> Result<Option<super::Settings>> {
        validate_size(xml)?;
        let mut reader = NsReader::from_str(xml);
        let mut buffer = Vec::new();
        let mut handler = CalculationHandler::default();
        loop {
            let (namespace, event) = reader
                .read_resolved_event_into(&mut buffer)
                .map_err(|error| Error::InvalidFormat(format!("XML parsing error: {error}")))?;
            let (is_table, is_office) = classify(&namespace);
            let is_eof = matches!(event, Event::Eof);
            handler.on_event(
                is_table,
                is_office,
                &event,
                reader.resolver(),
                reader.decoder(),
            )?;
            buffer.clear();
            if is_eof {
                break;
            }
        }
        handler.finish()
    }

    #[test]
    fn handler_matches_standalone_parse() {
        const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
        const TABLE: &str = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
        let no_settings = format!(
            r#"<o:document-content xmlns:o="{OFFICE}" xmlns:t="{TABLE}"><o:body><o:spreadsheet/></o:body></o:document-content>"#
        );
        let full_settings = format!(
            r#"<o:document-content xmlns:o="{OFFICE}" xmlns:t="{TABLE}"><o:body><o:spreadsheet><t:calculation-settings t:case-sensitive="1" t:null-year="1930"><t:null-date t:value-type="date" t:date-value="1899-12-30+08:00"/><t:iteration t:status="enable" t:steps="100" t:maximum-difference="NaN"/></t:calculation-settings></o:spreadsheet></o:body></o:document-content>"#
        );
        let duplicate = format!(
            r#"<o:document-content xmlns:o="{OFFICE}" xmlns:t="{TABLE}"><o:body><o:spreadsheet><t:calculation-settings/><t:calculation-settings/></o:spreadsheet></o:body></o:document-content>"#
        );
        let misplaced = format!(
            r#"<o:document-content xmlns:o="{OFFICE}" xmlns:t="{TABLE}"><o:body><t:calculation-settings/></o:body></o:document-content>"#
        );
        let text_inside = format!(
            r#"<o:document-content xmlns:o="{OFFICE}" xmlns:t="{TABLE}"><o:body><o:spreadsheet><t:calculation-settings>text</t:calculation-settings></o:spreadsheet></o:body></o:document-content>"#
        );
        let unterminated = format!(
            r#"<o:document-content xmlns:o="{OFFICE}" xmlns:t="{TABLE}"><o:body><o:spreadsheet><t:calculation-settings>"#
        );
        for xml in [
            no_settings,
            full_settings,
            duplicate,
            misplaced,
            text_inside,
            unterminated,
        ] {
            let expected = parse(&xml, crate::XmlStreamLimits::default())
                .map(|settings| format!("{settings:?}"))
                .map_err(|error| error.to_string());
            let actual = handler_parse(&xml)
                .map(|settings| format!("{settings:?}"))
                .map_err(|error| error.to_string());
            assert_eq!(expected, actual, "handler and standalone disagree");
        }
    }

    /// A one-line summary to pin, and the full value for driver parity.
    fn outcome(
        settings: Option<&super::Settings>,
        tolerance: &crate::leniency::ToleranceReport,
    ) -> (String, String) {
        (
            format!(
                "ok settings={} defects={}",
                settings.is_some(),
                tolerance.total()
            ),
            format!("{settings:?} {tolerance:?}"),
        )
    }

    /// The fused handler's outcome with its repairs, mirroring `handler_parse`.
    fn handler_outcome(xml: &str) -> (String, String) {
        let run = || -> Result<_> {
            validate_size(xml)?;
            let mut reader = NsReader::from_str(xml);
            let mut buffer = Vec::new();
            let mut handler = CalculationHandler::default();
            loop {
                let (namespace, event) = reader
                    .read_resolved_event_into(&mut buffer)
                    .map_err(|error| Error::InvalidFormat(format!("XML parsing error: {error}")))?;
                let (is_table, is_office) = classify(&namespace);
                let is_eof = matches!(event, Event::Eof);
                handler.on_event(
                    is_table,
                    is_office,
                    &event,
                    reader.resolver(),
                    reader.decoder(),
                )?;
                buffer.clear();
                if is_eof {
                    break;
                }
            }
            handler.finish_with_report()
        };
        match run() {
            Ok((settings, tolerance)) => outcome(settings.as_ref(), &tolerance),
            Err(error) => (error.to_string(), String::new()),
        }
    }

    fn standalone_outcome(xml: &str) -> (String, String) {
        match super::super::codec::parse_with_report(xml, crate::XmlStreamLimits::default()) {
            Ok((settings, tolerance)) => outcome(settings.as_ref(), &tolerance),
            Err(error) => (error.to_string(), String::new()),
        }
    }

    /// One case per position the scan can be in when an event arrives: outside
    /// the settings element, directly inside it, inside a known child, inside
    /// a skipped unknown subtree, and after the element has closed.
    #[test]
    fn every_scan_position_reports_the_same_outcome_in_both_drivers() {
        const SETTINGS: &str = "<t:calculation-settings>";
        const END: &str = "</t:calculation-settings>";
        let cases: [(&str, String, &str); 20] = [
            (
                "unknown child subtree is skipped with its content",
                format!(
                    "{SETTINGS}<x:a><x:b>text<![CDATA[c]]>&amp;&#65;<x:c/></x:b><!--c--><?p i?></x:a>{END}"
                ),
                "ok settings=true defects=1",
            ),
            (
                "unknown entity inside a skipped subtree",
                format!("{SETTINGS}<x:a>&unknown;</x:a>{END}"),
                "Invalid format: table:calculation-settings cannot contain entity references",
            ),
            (
                "unknown table child is skipped",
                format!("{SETTINGS}<t:other><t:null-date/></t:other>{END}"),
                "ok settings=true defects=1",
            ),
            (
                "unknown empty child",
                format!("{SETTINGS}<x:a/>{END}"),
                "ok settings=true defects=1",
            ),
            (
                "known child written with an end tag",
                format!("{SETTINGS}<t:null-date></t:null-date><t:iteration> </t:iteration>{END}"),
                "ok settings=true defects=0",
            ),
            (
                "known child with an empty child",
                format!("{SETTINGS}<t:null-date><x:a/></t:null-date>{END}"),
                "Invalid format: calculation setting children must be empty",
            ),
            (
                "known child with an open child",
                format!("{SETTINGS}<t:null-date><x:a></x:a></t:null-date>{END}"),
                "Invalid format: calculation setting children must be empty",
            ),
            (
                "text inside a known child",
                format!("{SETTINGS}<t:null-date>x</t:null-date>{END}"),
                "Invalid format: table:calculation-settings cannot contain text",
            ),
            (
                "character data inside the settings element",
                format!("{SETTINGS}<![CDATA[x]]>{END}"),
                "Invalid format: table:calculation-settings cannot contain CDATA",
            ),
            (
                "blank character data inside the settings element",
                format!("{SETTINGS} <![CDATA[ ]]><!--c--><?p i?>{END}"),
                "ok settings=true defects=0",
            ),
            (
                "entity reference inside the settings element",
                format!("{SETTINGS}&amp;{END}"),
                "Invalid format: table:calculation-settings cannot contain entity references",
            ),
            (
                "entity reference inside a known child",
                format!("{SETTINGS}<t:null-date>&#65;</t:null-date>{END}"),
                "Invalid format: table:calculation-settings cannot contain entity references",
            ),
            (
                "second settings element after a closed one",
                format!("{SETTINGS}{END}{SETTINGS}{END}"),
                "Invalid format: duplicate table:calculation-settings",
            ),
            (
                "empty settings element after a closed one",
                format!("{SETTINGS}{END}<t:calculation-settings/>"),
                "Invalid format: duplicate table:calculation-settings",
            ),
            (
                "settings element nested in the settings element",
                format!("{SETTINGS}{SETTINGS}{END}{END}"),
                "Invalid format: table:calculation-settings must be a direct office document-body child",
            ),
            (
                "settings element nested in a known child",
                format!("{SETTINGS}<t:null-date><t:calculation-settings/></t:null-date>{END}"),
                "Invalid format: table:calculation-settings must be a direct office document-body child",
            ),
            (
                "settings element nested in a skipped subtree",
                format!("{SETTINGS}<x:a><t:calculation-settings/></x:a>{END}"),
                "Invalid format: table:calculation-settings must be a direct office document-body child",
            ),
            (
                "settings element below a table",
                format!("<t:table>{SETTINGS}{END}</t:table>"),
                "Invalid format: table:calculation-settings must be a direct office document-body child",
            ),
            (
                "null-date after iteration is repaired",
                format!("{SETTINGS}<t:iteration/><t:null-date/>{END}"),
                "ok settings=true defects=1",
            ),
            (
                "content after the closed settings element is ignored",
                format!("{SETTINGS}{END}<t:table>text&amp;<t:null-date/></t:table>"),
                "ok settings=true defects=0",
            ),
        ];
        for (label, body, expected) in cases {
            let xml = format!(
                r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:x="urn:example:other"><o:body><o:spreadsheet>{body}</o:spreadsheet></o:body></o:document-content>"#
            );
            let standalone = standalone_outcome(&xml);
            assert_eq!(standalone.0, expected, "standalone: {label}");
            assert_eq!(handler_outcome(&xml), standalone, "handler: {label}");
        }
    }
}
