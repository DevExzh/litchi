mod support;

use litchi_ods::names::{Definition, Expression, Range, Scope};
use litchi_ods::{Builder, Spreadsheet};

const CONTENT: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content
    xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
    office:version="1.3">
  <office:body><office:spreadsheet><table:table table:name="Sheet1"/></office:spreadsheet></office:body>
</office:document-content>"#;

fn global() -> Scope {
    Scope::Global
}

fn sheet(name: &str) -> Scope {
    Scope::sheet(name)
}

#[test]
fn named_definitions_round_trip_in_order_and_scope() {
    let range = Range::new("Revenue", "$Sheet1.$A$1:.$B$2", global())
        .expect("test fixture or operation should succeed");
    let expression = Expression::new("TaxRate", "of:=0.2", global())
        .expect("test fixture or operation should succeed");
    let local = Range::new("Input", "$Sheet1.$C$1", sheet("Sheet1"))
        .expect("test fixture or operation should succeed");

    let compact_content = support::compact_xml_fixture(CONTENT);
    let bytes = support::raw_package(&[("content.xml", compact_content.as_bytes(), "text/xml")]);
    let mut spreadsheet =
        Spreadsheet::from_bytes(bytes).expect("test fixture or operation should succeed");
    spreadsheet
        .add_range(range)
        .expect("test fixture or operation should succeed");
    spreadsheet
        .add_expression(expression)
        .expect("test fixture or operation should succeed");
    spreadsheet
        .add_range(local)
        .expect("test fixture or operation should succeed");
    assert_eq!(
        spreadsheet
            .definitions()
            .iter()
            .map(Definition::name)
            .collect::<Vec<_>>(),
        ["Revenue", "TaxRate", "Input"]
    );
    assert_eq!(spreadsheet.ranges().count(), 2);
    assert_eq!(spreadsheet.expressions().count(), 1);
    assert_eq!(
        spreadsheet
            .range("Input", &sheet("Sheet1"))
            .map(|range| range.name.as_str()),
        Some("Input")
    );
    assert_eq!(
        spreadsheet
            .expression("TaxRate", &global())
            .map(|expression| expression.expression.as_str()),
        Some("of:=0.2")
    );

    let added = Expression::new("Total", "of:=SUM([.A1:.A2])", global())
        .expect("test fixture or operation should succeed");
    spreadsheet
        .add_expression(added)
        .expect("test fixture or operation should succeed");
    let reopened = Spreadsheet::from_bytes(spreadsheet.into_bytes())
        .expect("test fixture or operation should succeed");
    assert_eq!(
        reopened
            .definitions()
            .iter()
            .map(Definition::name)
            .collect::<Vec<_>>(),
        ["Revenue", "TaxRate", "Total", "Input"]
    );
}

#[test]
fn duplicate_and_invalid_definitions_are_rejected_without_mutation() {
    let mut spreadsheet = Spreadsheet::from_bytes(
        Builder::new()
            .build()
            .expect("test fixture or operation should succeed"),
    )
    .expect("test fixture or operation should succeed");
    let first = Range::new("Total", "$Sheet1.$A$1", global())
        .expect("test fixture or operation should succeed");
    spreadsheet
        .add_range(first)
        .expect("test fixture or operation should succeed");
    let before_xml = spreadsheet.content_xml().to_owned();
    let before_catalog = spreadsheet.definitions().to_vec();

    let duplicate = Range::new("Total", "$Sheet1.$B$1", global())
        .expect("test fixture or operation should succeed");
    assert!(spreadsheet.add_range(duplicate).is_err());
    assert_eq!(spreadsheet.content_xml(), before_xml);
    assert_eq!(spreadsheet.definitions(), before_catalog);

    let mut invalid = Range::new("Broken", "$Sheet1.$C$1", global())
        .expect("test fixture or operation should succeed");
    invalid.cell_range_address = " ".to_owned();
    assert!(spreadsheet.add_range(invalid).is_err());
    assert_eq!(spreadsheet.content_xml(), before_xml);
    assert_eq!(spreadsheet.definitions(), before_catalog);

    let missing_sheet = Range::new("Missing", "$Missing.$A$1", sheet("Missing"))
        .expect("test fixture or operation should succeed");
    assert!(spreadsheet.add_range(missing_sheet).is_err());
    assert_eq!(spreadsheet.content_xml(), before_xml);
    assert_eq!(spreadsheet.definitions(), before_catalog);
}

#[test]
fn builder_rejects_duplicate_definitions_before_building() {
    let mut builder = Builder::new();
    builder
        .add_range(
            Range::new("Total", "$Sheet1.$A$1", global())
                .expect("test fixture or operation should succeed"),
        )
        .expect("test fixture or operation should succeed");
    assert!(
        builder
            .add_range(
                Range::new("Total", "$Sheet1.$B$1", global())
                    .expect("test fixture or operation should succeed")
            )
            .is_err()
    );
    assert_eq!(builder.definitions().len(), 1);
}
