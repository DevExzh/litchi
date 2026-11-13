use super::codec::MAX_DEPTH;
use super::*;
use crate::error::Result;
use litchi_opc::{OpcPackage, PackURI};
use quick_xml::events::Event;
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;

const NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const MCE: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

fn parse(child: &str) -> Result<Option<IgnoredErrors>> {
    parse_worksheet_ignored_errors(
        format!(r#"<worksheet xmlns="{NS}">{child}</worksheet>"#).as_bytes(),
    )
}

#[test]
fn parses_all_flags_ranges_and_defaults() {
    let value = parse(concat!(
        r#"<ignoredErrors><ignoredError sqref="A1 $B$2:C3" calculatedColumn="1" "#,
        r#"emptyCellReference="true" evalError="1" formula="true" formulaRange="1" "#,
        r#"listDataValidation="true" numberStoredAsText="1" twoDigitTextYear="true" unlockedFormula="1"/>"#,
        r#"<ignoredError sqref="XFD1048576"/></ignoredErrors>"#,
    )).unwrap().unwrap();
    assert_eq!(value.entries().len(), 2);
    assert_eq!(value.entries()[0].ranges()[1].as_str(), "$B$2:C3");
    for kind in [
        IgnoredErrorType::CalculatedColumn,
        IgnoredErrorType::EmptyCellReference,
        IgnoredErrorType::EvaluationError,
        IgnoredErrorType::Formula,
        IgnoredErrorType::FormulaRange,
        IgnoredErrorType::ListDataValidation,
        IgnoredErrorType::NumberStoredAsText,
        IgnoredErrorType::TwoDigitTextYear,
        IgnoredErrorType::UnlockedFormula,
    ] {
        assert!(value.entries()[0].ignores(kind));
        assert!(!value.entries()[1].ignores(kind));
    }
}

#[test]
fn supports_strict_mce_and_extension_retention() {
    let strict = br#"<worksheet xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main"><ignoredErrors><ignoredError sqref="A1" formula="1"/></ignoredErrors></worksheet>"#;
    assert!(
        parse_worksheet_ignored_errors(strict)
            .unwrap()
            .unwrap()
            .entries()[0]
            .ignores(IgnoredErrorType::Formula)
    );
    let xml = format!(
        concat!(
            r#"<worksheet xmlns="{}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" "#,
            r#"xmlns:x="urn:unsupported" mc:Ignorable="x"><ignoredErrors><ignoredError sqref="A1" x:drop="1"/>"#,
            r#"<extLst><ext uri="urn:test"><x:payload value="safe"/></ext></extLst></ignoredErrors></worksheet>"#,
        ),
        NS
    );
    let value = parse_worksheet_ignored_errors(xml.as_bytes())
        .unwrap()
        .unwrap();
    assert_eq!(value.extensions().len(), 1);
    assert_eq!(value.extensions()[0].uri(), "urn:test");
    assert!(
        std::str::from_utf8(value.extensions()[0].markup())
            .unwrap()
            .contains("<ext")
    );
}

#[test]
fn rejects_bad_references_structure_attributes_and_limits() {
    for child in [
        "<ignoredErrors/>",
        "<ignoredErrors><ignoredError/></ignoredErrors>",
        r#"<ignoredErrors><ignoredError sqref=""/></ignoredErrors>"#,
        r#"<ignoredErrors><ignoredError sqref="A0"/></ignoredErrors>"#,
        r#"<ignoredErrors><ignoredError sqref="XFE1"/></ignoredErrors>"#,
        r#"<ignoredErrors><ignoredError sqref="A1048577"/></ignoredErrors>"#,
        r#"<ignoredErrors><ignoredError sqref="A1" formula="yes"/></ignoredErrors>"#,
        r#"<ignoredErrors><ignoredError sqref="A1" mystery="1"/></ignoredErrors>"#,
        r#"<ignoredErrors><ignoredError sqref="A1"><child/></ignoredError></ignoredErrors>"#,
        r#"<ignoredErrors><extLst><ext uri="x"/></extLst><ignoredError sqref="A1"/></ignoredErrors>"#,
        r#"<ignoredErrors><ignoredError sqref="A1"/><extLst/></ignoredErrors>"#,
    ] {
        assert!(parse(child).is_err(), "expected rejection for {child}");
    }
    assert!(parse("<ignoredErrors><ignoredError sqref=\"A1\"/></ignoredErrors><ignoredErrors><ignoredError sqref=\"A1\"/></ignoredErrors>").is_err());
    let entries = (0..10)
        .map(|index| format!(r#"<ignoredError sqref="A{}"/>"#, index + 1))
        .collect::<String>();
    assert!(parse(&format!("<ignoredErrors>{entries}</ignoredErrors>")).is_err());
}

#[test]
fn rejects_multiple_roots_direct_text_and_excessive_depth() {
    for xml in [
        format!(r#"<worksheet xmlns="{NS}"/><worksheet xmlns="{NS}"/>"#),
        format!(r#"text<worksheet xmlns="{NS}"></worksheet>"#),
        format!(r#"<worksheet xmlns="{NS}">text</worksheet>"#),
        format!(r#"<worksheet xmlns="{NS}"></worksheet>tail"#),
    ] {
        assert!(
            parse_worksheet_ignored_errors(xml.as_bytes()).is_err(),
            "expected rejection for {xml}"
        );
    }

    let mut xml = format!(r#"<worksheet xmlns="{NS}">"#);
    for _ in 0..MAX_DEPTH {
        xml.push_str("<extension>");
    }
    for _ in 0..MAX_DEPTH {
        xml.push_str("</extension>");
    }
    xml.push_str("</worksheet>");
    assert!(parse_worksheet_ignored_errors(xml.as_bytes()).is_err());
}

fn fixture(bytes: &[u8]) -> IgnoredErrors {
    let package = OpcPackage::from_bytes(bytes, litchi_opc::SourceLimits::default()).unwrap();
    let part = package
        .get_part(&PackURI::new("/xl/worksheets/sheet1.xml").unwrap())
        .unwrap();
    parse_worksheet_ignored_errors(part.blob())
        .unwrap()
        .unwrap()
}

#[test]
fn reads_poi_ignored_error_fixtures() {
    let format = fixture(include_bytes!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/poi/test-data/spreadsheet/FormatKM.xlsx"
    )));
    assert_eq!(
        format.entries()[0]
            .ranges()
            .iter()
            .map(IgnoredErrorRangeReference::as_str)
            .collect::<Vec<_>>(),
        vec!["C2:C5", "E2:E4", "E5"]
    );
    assert!(format.entries()[0].ignores(IgnoredErrorType::NumberStoredAsText));

    let large = fixture(include_bytes!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/poi/test-data/spreadsheet/no_drawing_patriarch.xlsx"
    )));
    assert_eq!(large.entries()[0].ranges()[0].as_str(), "A1:J7577");
    assert!(large.entries()[0].ignores(IgnoredErrorType::NumberStoredAsText));
}

/// Every name in `markup` that does not resolve the way it resolves in the part
/// the fragment was cut from: an element whose name reaches no namespace at
/// all, or a prefixed name whose prefix has no declaration in the fragment.
///
/// A fragment cut out of a processed part has to carry the declarations it
/// inherits, so this is empty for the markup the accessors publish.
fn unresolved_names(markup: &[u8]) -> Vec<String> {
    let mut reader = NsReader::from_reader(markup);
    let mut unresolved = Vec::new();
    loop {
        let event = reader.read_event().unwrap().into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) | Event::Empty(element) => {
                if !matches!(namespace, ResolveResult::Bound(_)) {
                    unresolved.push(String::from_utf8_lossy(element.name().as_ref()).into_owned());
                }
                for attribute in element.attributes() {
                    let attribute = attribute.unwrap();
                    if matches!(
                        resolver.resolve_attribute(attribute.key).0,
                        ResolveResult::Unknown(_)
                    ) {
                        unresolved
                            .push(String::from_utf8_lossy(attribute.key.as_ref()).into_owned());
                    }
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }
    unresolved
}

#[test]
fn retained_extension_markup_resolves_standalone() {
    let xml = format!(
        concat!(
            r#"<worksheet xmlns="{}" xmlns:mc="{}" xmlns:p="urn:payload"><ignoredErrors>"#,
            r#"<ignoredError sqref="A1" formula="1"/><extLst><ext uri="urn:open">"#,
            r#"<p:payload p:flag="1"/></ext><ext uri="urn:empty"/></extLst>"#,
            r#"</ignoredErrors></worksheet>"#,
        ),
        NS, MCE
    );
    let value = parse_worksheet_ignored_errors(xml.as_bytes())
        .unwrap()
        .unwrap();
    assert_eq!(value.extensions().len(), 2);
    for extension in value.extensions() {
        let markup = extension.markup();
        assert!(
            unresolved_names(markup).is_empty(),
            "{}",
            String::from_utf8_lossy(markup)
        );
    }
    assert!(
        String::from_utf8_lossy(value.extensions()[0].markup())
            .contains(r#"xmlns:p="urn:payload""#)
    );
}

fn collection_outcome(body: &str) -> String {
    match parse(body) {
        Ok(Some(value)) => format!(
            "ok entries={} extensions={}",
            value.entries.len(),
            value.extensions.len()
        ),
        Ok(None) => "none".to_owned(),
        Err(error) => format!("{error:?}"),
    }
}

#[test]
fn collection_child_order_matrix_is_stable() {
    let error = ("error", r#"<ignoredError sqref="A1" evalError="1"/>"#);
    let open_error = (
        "open error",
        r#"<ignoredError sqref="B2" evalError="1"></ignoredError>"#,
    );
    let list = ("extLst", r#"<extLst><ext uri="urn:a"/></extLst>"#);
    let open_list = (
        "open extLst",
        r#"<extLst><ext uri="urn:b"><y:data xmlns:y="urn:y"/></ext></extLst>"#,
    );
    let empty_list = ("empty extLst", "<extLst/>");
    let unknown = ("unknown", "<unknown/>");
    let children = [error, open_error, list, open_list];
    let mut cases: Vec<Vec<(&str, &str)>> = vec![Vec::new()];
    for first in children {
        cases.push(vec![first]);
    }
    for first in children {
        for second in children {
            cases.push(vec![first, second]);
        }
    }
    // Every pair again behind one entry, so that an accepted order is not
    // hidden by the refusal of an extension list without entries.
    for second in children {
        for third in children {
            cases.push(vec![error, second, third]);
        }
    }
    cases.push(vec![empty_list]);
    cases.push(vec![error, empty_list]);
    cases.push(vec![error, list, empty_list]);
    cases.push(vec![unknown]);
    cases.push(vec![error, unknown]);
    cases.push(vec![error, list, unknown]);
    cases.push(vec![(
        "foreign error",
        r#"<q:ignoredError xmlns:q="urn:other" sqref="A1"/>"#,
    )]);
    let mut actual: Vec<(String, String)> = cases
        .iter()
        .map(|case| {
            let label = case.iter().map(|(name, _)| *name).collect::<Vec<_>>();
            let body: String = case.iter().map(|(_, xml)| *xml).collect();
            (
                label.join(","),
                collection_outcome(&format!("<ignoredErrors>{body}</ignoredErrors>")),
            )
        })
        .collect();
    // A second collection is refused whatever the first one reached.
    for (label, body) in [
        ("empty collection", "<ignoredErrors/>".to_owned()),
        (
            "two collections",
            format!(
                "<ignoredErrors>{0}</ignoredErrors><ignoredErrors>{0}</ignoredErrors>",
                error.1
            ),
        ),
        (
            "collection with extLst, then collection",
            format!(
                "<ignoredErrors>{}{}</ignoredErrors><ignoredErrors>{}</ignoredErrors>",
                error.1, list.1, error.1
            ),
        ),
    ] {
        actual.push((label.to_owned(), collection_outcome(&body)));
    }
    let expected: &[(&str, &str)] = &[
        (
            "",
            "Invalid(\"ignoredErrors requires at least one ignoredError\")",
        ),
        ("error", "ok entries=1 extensions=0"),
        ("open error", "ok entries=1 extensions=0"),
        (
            "extLst",
            "Invalid(\"ignoredErrors extLst appears before ignoredError\")",
        ),
        (
            "open extLst",
            "Invalid(\"ignoredErrors extLst appears before ignoredError\")",
        ),
        ("error,error", "ok entries=2 extensions=0"),
        ("error,open error", "ok entries=2 extensions=0"),
        ("error,extLst", "ok entries=1 extensions=1"),
        ("error,open extLst", "ok entries=1 extensions=1"),
        ("open error,error", "ok entries=2 extensions=0"),
        ("open error,open error", "ok entries=2 extensions=0"),
        ("open error,extLst", "ok entries=1 extensions=1"),
        ("open error,open extLst", "ok entries=1 extensions=1"),
        (
            "extLst,error",
            "Invalid(\"ignoredErrors extLst appears before ignoredError\")",
        ),
        (
            "extLst,open error",
            "Invalid(\"ignoredErrors extLst appears before ignoredError\")",
        ),
        (
            "extLst,extLst",
            "Invalid(\"ignoredErrors extLst appears before ignoredError\")",
        ),
        (
            "extLst,open extLst",
            "Invalid(\"ignoredErrors extLst appears before ignoredError\")",
        ),
        (
            "open extLst,error",
            "Invalid(\"ignoredErrors extLst appears before ignoredError\")",
        ),
        (
            "open extLst,open error",
            "Invalid(\"ignoredErrors extLst appears before ignoredError\")",
        ),
        (
            "open extLst,extLst",
            "Invalid(\"ignoredErrors extLst appears before ignoredError\")",
        ),
        (
            "open extLst,open extLst",
            "Invalid(\"ignoredErrors extLst appears before ignoredError\")",
        ),
        ("error,error,error", "ok entries=3 extensions=0"),
        ("error,error,open error", "ok entries=3 extensions=0"),
        ("error,error,extLst", "ok entries=2 extensions=1"),
        ("error,error,open extLst", "ok entries=2 extensions=1"),
        ("error,open error,error", "ok entries=3 extensions=0"),
        ("error,open error,open error", "ok entries=3 extensions=0"),
        ("error,open error,extLst", "ok entries=2 extensions=1"),
        ("error,open error,open extLst", "ok entries=2 extensions=1"),
        (
            "error,extLst,error",
            "Invalid(\"ignoredError appears after ignoredErrors extLst\")",
        ),
        (
            "error,extLst,open error",
            "Invalid(\"ignoredError appears after ignoredErrors extLst\")",
        ),
        (
            "error,extLst,extLst",
            "Invalid(\"duplicate ignoredErrors extLst\")",
        ),
        (
            "error,extLst,open extLst",
            "Invalid(\"duplicate ignoredErrors extLst\")",
        ),
        (
            "error,open extLst,error",
            "Invalid(\"ignoredError appears after ignoredErrors extLst\")",
        ),
        (
            "error,open extLst,open error",
            "Invalid(\"ignoredError appears after ignoredErrors extLst\")",
        ),
        (
            "error,open extLst,extLst",
            "Invalid(\"duplicate ignoredErrors extLst\")",
        ),
        (
            "error,open extLst,open extLst",
            "Invalid(\"duplicate ignoredErrors extLst\")",
        ),
        (
            "empty extLst",
            "Invalid(\"ignoredErrors extLst appears before ignoredError\")",
        ),
        (
            "error,empty extLst",
            "Invalid(\"ignoredErrors extLst requires at least one ext\")",
        ),
        (
            "error,extLst,empty extLst",
            "Invalid(\"duplicate ignoredErrors extLst\")",
        ),
        (
            "unknown",
            "Invalid(\"unexpected ignoredErrors element 'unknown'\")",
        ),
        (
            "error,unknown",
            "Invalid(\"unexpected ignoredErrors element 'unknown'\")",
        ),
        (
            "error,extLst,unknown",
            "Invalid(\"unexpected ignoredErrors element 'unknown'\")",
        ),
        (
            "foreign error",
            "Invalid(\"unexpected ignoredErrors element 'ignoredError'\")",
        ),
        (
            "empty collection",
            "Invalid(\"ignoredErrors requires at least one ignoredError\")",
        ),
        (
            "two collections",
            "Invalid(\"duplicate worksheet ignoredErrors element\")",
        ),
        (
            "collection with extLst, then collection",
            "Invalid(\"duplicate worksheet ignoredErrors element\")",
        ),
    ];
    let same = actual.len() == expected.len()
        && actual.iter().zip(expected).all(
            |((label, outcome), (expected_label, expected_outcome))| {
                label == expected_label && outcome == expected_outcome
            },
        );
    let observed: String = actual
        .iter()
        .map(|(label, outcome)| format!("({label:?}, {outcome:?}),\n"))
        .collect();
    assert!(same, "observed rows:\n{observed}");
}
