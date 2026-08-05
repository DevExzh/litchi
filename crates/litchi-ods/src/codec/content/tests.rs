//! Regression coverage for the ODS `content.xml` owner.

use super::*;

const TEST_SHEETS_XML: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
    <office:body>
        <office:spreadsheet>
            <table:table table:name="Sheet1">
                <table:table-row>
                    <table:table-cell office:value-type="string">
                        <text:p>Hello</text:p>
                    </table:table-cell>
                    <table:table-cell office:value-type="float" office:value="42">
                        <text:p>42</text:p>
                    </table:table-cell>
                </table:table-row>
            </table:table>
        </office:spreadsheet>
    </office:body>
</office:document-content>"#;

const TEST_MULTIPLE_SHEETS_XML: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
    <office:body>
        <office:spreadsheet>
            <table:table table:name="Sheet1">
                <table:table-row>
                    <table:table-cell office:value-type="string">
                        <text:p>First Sheet</text:p>
                    </table:table-cell>
                </table:table-row>
            </table:table>
            <table:table table:name="Sheet2">
                <table:table-row>
                    <table:table-cell office:value-type="string">
                        <text:p>Second Sheet</text:p>
                    </table:table-cell>
                </table:table-row>
            </table:table>
        </office:spreadsheet>
    </office:body>
</office:document-content>"#;

const TEST_CELL_TYPES_XML: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
    <office:body>
        <office:spreadsheet>
            <table:table table:name="TypesTest">
                <table:table-row>
                    <table:table-cell office:value-type="string"><text:p>Text</text:p></table:table-cell>
                    <table:table-cell office:value-type="float" office:value="3.14"><text:p>3.14</text:p></table:table-cell>
                    <table:table-cell office:value-type="currency" office:value="100" office:currency="EUR"><text:p>€100</text:p></table:table-cell>
                    <table:table-cell office:value-type="percentage" office:value="0.5"><text:p>50%</text:p></table:table-cell>
                    <table:table-cell office:value-type="boolean" office:value="true"><text:p>TRUE</text:p></table:table-cell>
                    <table:table-cell office:value-type="date" office:value="2024-03-15"><text:p>2024-03-15</text:p></table:table-cell>
                    <table:table-cell office:value-type="time" office:value="PT12H30M00S"><text:p>12:30:00</text:p></table:table-cell>
                </table:table-row>
            </table:table>
        </office:spreadsheet>
    </office:body>
</office:document-content>"#;

const TEST_FORMULA_XML: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
    <office:body>
        <office:spreadsheet>
            <table:table table:name="FormulaTest">
                <table:table-row>
                    <table:table-cell office:value-type="float" office:value="10"><text:p>10</text:p></table:table-cell>
                    <table:table-cell office:value-type="float" office:value="20"><text:p>20</text:p></table:table-cell>
                    <table:table-cell table:formula="=SUM([.A1]:[.B1])" office:value-type="float" office:value="30">
                        <text:p>30</text:p>
                    </table:table-cell>
                </table:table-row>
            </table:table>
        </office:spreadsheet>
    </office:body>
</office:document-content>"#;

const TEST_REPEATED_CELLS_XML: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
    <office:body>
        <office:spreadsheet>
            <table:table table:name="RepeatedTest">
                <table:table-row>
                    <table:table-cell table:number-columns-repeated="3" office:value-type="string">
                        <text:p>Repeated</text:p>
                    </table:table-cell>
                    <table:table-cell office:value-type="string">
                        <text:p>Single</text:p>
                    </table:table-cell>
                </table:table-row>
            </table:table>
        </office:spreadsheet>
    </office:body>
</office:document-content>"#;

const TEST_EMPTY_SHEET_XML: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
    <office:body>
        <office:spreadsheet>
            <table:table table:name="EmptySheet">
            </table:table>
        </office:spreadsheet>
    </office:body>
</office:document-content>"#;

const TEST_SPAN_TEXT_XML: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
    <office:body>
        <office:spreadsheet>
            <table:table table:name="SpanTest">
                <table:table-row>
                    <table:table-cell office:value-type="string">
                        <text:p>Normal text <text:span>spanned text</text:span> more text</text:p>
                    </table:table-cell>
                </table:table-row>
            </table:table>
        </office:spreadsheet>
    </office:body>
</office:document-content>"#;

#[test]
fn test_parse_sheets_basic() {
    let sheets = Parser::parse_sheets(TEST_SHEETS_XML).unwrap();
    assert_eq!(sheets.len(), 1);
    assert_eq!(sheets[0].name, "Sheet1");
    assert_eq!(sheets[0].rows.len(), 1);
}

#[test]
fn test_parse_multiple_sheets() {
    let sheets = Parser::parse_sheets(TEST_MULTIPLE_SHEETS_XML).unwrap();
    assert_eq!(sheets.len(), 2);
    assert_eq!(sheets[0].name, "Sheet1");
    assert_eq!(sheets[1].name, "Sheet2");
}

#[test]
fn test_parse_cell_types() {
    let sheets = Parser::parse_sheets(TEST_CELL_TYPES_XML).unwrap();
    assert_eq!(sheets.len(), 1);

    let row = &sheets[0].rows[0];
    assert_eq!(row.cells.len(), 7);

    // Text cell
    match &row.cells[0].value {
        CellValue::Text(t) => assert_eq!(t, "Text"),
        _ => panic!("Expected Text"),
    }

    // Float/Number cell
    match &row.cells[1].value {
        CellValue::Number(n) => {
            let expected = (std::f64::consts::PI * 100.0).trunc() / 100.0;
            assert!((n - expected).abs() < f64::EPSILON);
        },
        _ => panic!("Expected Number"),
    }

    // Currency cell
    match &row.cells[2].value {
        CellValue::Currency(amount, currency) => {
            assert!((amount - 100.0).abs() < f64::EPSILON);
            assert_eq!(currency, "EUR");
        },
        _ => panic!("Expected Currency"),
    }

    // Percentage cell
    match &row.cells[3].value {
        CellValue::Percentage(p) => assert!((p - 0.5).abs() < f64::EPSILON),
        _ => panic!("Expected Percentage"),
    }

    // Boolean cell
    match &row.cells[4].value {
        CellValue::Boolean(b) => assert!(*b),
        _ => panic!("Expected Boolean"),
    }

    // Date cell
    match &row.cells[5].value {
        CellValue::Date(d) => assert_eq!(d, "2024-03-15"),
        _ => panic!("Expected Date"),
    }

    // Time cell
    match &row.cells[6].value {
        CellValue::Time(t) => assert_eq!(t, "PT12H30M00S"),
        _ => panic!("Expected Time"),
    }
}

#[test]
fn test_parse_formula() {
    let sheets = Parser::parse_sheets(TEST_FORMULA_XML).unwrap();
    assert_eq!(sheets.len(), 1);

    let row = &sheets[0].rows[0];
    assert_eq!(row.cells.len(), 3);

    // Cell with formula
    assert_eq!(row.cells[2].formula, Some("=SUM([.A1]:[.B1])".to_string()));
    match &row.cells[2].value {
        CellValue::Number(n) => assert!((n - 30.0).abs() < f64::EPSILON),
        _ => panic!("Expected Number for formula result"),
    }
}

#[test]
fn test_parse_repeated_cells() {
    let sheets = Parser::parse_sheets(TEST_REPEATED_CELLS_XML).unwrap();
    assert_eq!(sheets.len(), 1);

    let row = &sheets[0].rows[0];
    // 3 repeated cells + 1 single = 4 cells
    assert_eq!(row.cells.len(), 4);

    for i in 0..3 {
        match &row.cells[i].value {
            CellValue::Text(t) => assert_eq!(t, "Repeated"),
            _ => panic!("Expected Text for repeated cell {i}"),
        }
    }

    match &row.cells[3].value {
        CellValue::Text(t) => assert_eq!(t, "Single"),
        _ => panic!("Expected Text for single cell"),
    }
}

#[test]
fn parses_cell_range_sources_with_namespace_aliases_and_repetition() {
    let xml = r#"<o:document-content
            xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
            xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
            xmlns:x="http://www.w3.org/1999/xlink">
          <o:body><o:spreadsheet><t:table t:name="Links"><t:table-row>
            <t:table-cell t:number-columns-repeated="2">
              <t:cell-range-source t:name="Named &amp; Range"
                t:last-column-spanned="4" t:last-row-spanned="3"
                t:filter-name="calc8" t:filter-options="A&amp;B"
                t:refresh-delay="PT15M" x:type="simple"
                x:href="../Data&amp;More.ods" x:actuate="onRequest"></t:cell-range-source>
            </t:table-cell>
          </t:table-row></t:table></o:spreadsheet></o:body>
        </o:document-content>"#;

    let sheets = Parser::parse_sheets(xml).unwrap();
    let cells = &sheets[0].rows[0].cells;
    assert_eq!(cells.len(), 2);
    for cell in cells {
        let source = cell.range_source().unwrap();
        assert_eq!(source.name(), "Named & Range");
        assert_eq!(source.href(), "../Data&More.ods");
        assert_eq!((source.rows(), source.columns()), (3, 4));
        assert!(source.actuate_on_request());
        assert_eq!(source.filter_name(), Some("calc8"));
        assert_eq!(source.filter_options(), Some("A&B"));
        assert_eq!(source.refresh_delay(), Some("PT15M"));
    }
}

#[test]
fn rejects_incomplete_or_duplicate_cell_range_sources() {
    let missing_type = r#"<o:document-content
            xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
            xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
            xmlns:x="http://www.w3.org/1999/xlink">
          <o:body><o:spreadsheet><t:table t:name="Links"><t:table-row><t:table-cell>
            <t:cell-range-source t:name="R" t:last-column-spanned="1"
              t:last-row-spanned="1" x:href="source.ods"/>
          </t:table-cell></t:table-row></t:table></o:spreadsheet></o:body>
        </o:document-content>"#;
    assert!(Parser::parse_sheets(missing_type).is_err());

    let duplicate = r#"<o:document-content
            xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
            xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
            xmlns:x="http://www.w3.org/1999/xlink">
          <o:body><o:spreadsheet><t:table t:name="Links"><t:table-row><t:table-cell>
            <t:cell-range-source t:name="R1" t:last-column-spanned="1"
              t:last-row-spanned="1" x:type="simple" x:href="one.ods"/>
            <t:cell-range-source t:name="R2" t:last-column-spanned="1"
              t:last-row-spanned="1" x:type="simple" x:href="two.ods"/>
          </t:table-cell></t:table-row></t:table></o:spreadsheet></o:body>
        </o:document-content>"#;
    assert!(Parser::parse_sheets(duplicate).is_err());
}

#[test]
fn parses_typed_detective_ranges_and_operations() {
    let xml = r#"<o:document-content
            xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
            xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
          <o:body><o:spreadsheet><t:table t:name="Audit"><t:table-row>
            <t:table-cell t:number-columns-repeated="2"><t:detective>
              <t:highlighted-range t:cell-range-address=".A1:.B2"
                t:direction="from-same-table" t:contains-error="true"/>
              <t:highlighted-range t:marked-invalid="false"></t:highlighted-range>
              <t:operation t:name="trace-precedents" t:index="0"/>
              <t:operation t:name="trace-errors" t:index="7"></t:operation>
            </t:detective></t:table-cell>
          </t:table-row></t:table></o:spreadsheet></o:body>
        </o:document-content>"#;

    let sheets = Parser::parse_sheets(xml).unwrap();
    let cells = &sheets[0].rows[0].cells;
    assert_eq!(cells.len(), 2);
    for cell in cells {
        let detective = cell.detective().unwrap();
        assert_eq!(detective.highlighted_ranges().len(), 2);
        assert_eq!(detective.operations().len(), 2);
        let range = &detective.highlighted_ranges()[0];
        assert_eq!(range.cell_range_address(), Some(".A1:.B2"));
        assert_eq!(range.direction(), Some(DetectiveDirection::FromSameTable));
        assert_eq!(range.contains_error(), Some(true));
        assert_eq!(range.marked_invalid(), None);
        assert_eq!(
            detective.highlighted_ranges()[1].marked_invalid(),
            Some(false)
        );
        assert_eq!(
            detective.operations()[1],
            DetectiveOperation::new(DetectiveOperationKind::TraceErrors, 7)
        );
    }
}

#[test]
fn rejects_schema_invalid_detective_metadata() {
    let operation_before_range = r#"<o:document-content
            xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
            xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
          <o:body><o:spreadsheet><t:table t:name="Audit"><t:table-row><t:table-cell>
            <t:detective><t:operation t:name="trace-errors" t:index="0"/>
              <t:highlighted-range t:direction="from-same-table"/></t:detective>
          </t:table-cell></t:table-row></t:table></o:spreadsheet></o:body>
        </o:document-content>"#;
    assert!(Parser::parse_sheets(operation_before_range).is_err());

    let mixed_range = r#"<o:document-content
            xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
            xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
          <o:body><o:spreadsheet><t:table t:name="Audit"><t:table-row><t:table-cell>
            <t:detective><t:highlighted-range t:marked-invalid="true"
              t:direction="from-same-table"/></t:detective>
          </t:table-cell></t:table-row></t:table></o:spreadsheet></o:body>
        </o:document-content>"#;
    assert!(Parser::parse_sheets(mixed_range).is_err());

    let negative_index = operation_before_range
        .replace(
            r#"t:name="trace-errors" t:index="0""#,
            r#"t:name="trace-errors" t:index="-1""#,
        )
        .replace(
            r#"<t:highlighted-range t:direction="from-same-table"/>"#,
            "",
        );
    assert!(Parser::parse_sheets(&negative_index).is_err());

    let nested_child = operation_before_range
            .replace(
                r#"<t:operation t:name="trace-errors" t:index="0"/>"#,
                "",
            )
            .replace(
                r#"<t:highlighted-range t:direction="from-same-table"/>"#,
                r#"<t:highlighted-range t:direction="from-same-table"><t:operation t:name="trace-errors" t:index="1"/></t:highlighted-range>"#,
            );
    assert!(Parser::parse_sheets(&nested_child).is_err());
}

#[test]
fn test_parse_empty_sheet() {
    let sheets = Parser::parse_sheets(TEST_EMPTY_SHEET_XML).unwrap();
    assert_eq!(sheets.len(), 1);
    assert_eq!(sheets[0].name, "EmptySheet");
    assert_eq!(sheets[0].rows.len(), 0);
}

#[test]
fn test_parse_span_text() {
    let sheets = Parser::parse_sheets(TEST_SPAN_TEXT_XML).unwrap();
    assert_eq!(sheets.len(), 1);

    let row = &sheets[0].rows[0];
    assert_eq!(row.cells.len(), 1);

    // Text should include content from both text:p and text:span
    assert!(row.cells[0].text.contains("Normal text"));
    assert!(row.cells[0].text.contains("spanned text"));
}

#[test]
fn parses_rich_annotations_without_mixing_them_into_cell_text() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
    xmlns:dc="http://purl.org/dc/elements/1.1/"
    xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
    xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0">
  <o:body><o:spreadsheet><table:table table:name="Notes"><table:table-row>
    <table:table-cell table:number-columns-repeated="2" o:value-type="string">
      <o:annotation o:display="true" draw:style-name="gr1" svg:width="3.2cm">
        <dc:creator>A &amp; B</dc:creator><dc:date>2026-07-13T12:34:56Z</dc:date>
        <text:p text:style-name="P1"><text:span text:style-name="T1">first</text:span><text:line-break/>second</text:p>
        <text:list><text:list-item><text:p>item</text:p></text:list-item></text:list>
      </o:annotation>
      <text:p>cell <text:span>value</text:span></text:p><text:p>line two</text:p>
    </table:table-cell>
  </table:table-row></table:table></o:spreadsheet></o:body>
</o:document-content>"#;

    let sheets = Parser::parse_sheets(xml).unwrap();
    let cells = &sheets[0].rows[0].cells;
    assert_eq!(cells.len(), 2);
    assert_eq!(cells[0].text, "cell value\nline two");
    assert_eq!(cells[1].text, "cell value\nline two");

    for cell in cells {
        let annotation = cell.annotation().unwrap();
        assert_eq!(annotation.creator().as_deref(), Some("A & B"));
        assert_eq!(annotation.date().as_deref(), Some("2026-07-13T12:34:56Z"));
        assert_eq!(annotation.display(), Some(true));
        assert_eq!(annotation.attribute("draw:style-name"), Some("gr1"));
        assert_eq!(annotation.attribute("svg:width"), Some("3.2cm"));
        assert_eq!(annotation.text(), "first\nsecond\nitem");
        assert_eq!(annotation.children()[2].name(), "text:p");
        assert_eq!(annotation.children()[3].name(), "text:list");
    }
}

#[test]
fn test_extract_table_name_default() {
    // XML without table:name attribute
    let xml = r#"<?xml version="1.0"?>
<office:document-content
    xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
    <office:body><office:spreadsheet><table:table/></office:spreadsheet></office:body>
</office:document-content>"#;

    let sheets = Parser::parse_sheets(xml).unwrap();
    assert_eq!(sheets.len(), 1);
    assert_eq!(sheets[0].name, "Sheet1"); // Default name
}

#[test]
fn parses_repeated_rows_and_merged_cell_coordinates() {
    let xml = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><office:body><office:spreadsheet><table:table table:name="Merged"><table:table-row table:number-rows-repeated="2"><table:table-cell office:value-type="string"><text:p>A</text:p></table:table-cell></table:table-row><table:table-row><table:table-cell table:number-rows-spanned="2" table:number-columns-spanned="2" office:value-type="string"><text:p>anchor</text:p></table:table-cell><table:covered-table-cell/><table:table-cell table:number-matrix-rows-spanned="3" table:number-matrix-columns-spanned="2" office:value-type="string"><text:p>C</text:p></table:table-cell></table:table-row><table:table-row><table:covered-table-cell table:number-columns-repeated="2"/></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>"#;
    let sheets = Parser::parse_sheets(xml).unwrap();
    let rows = &sheets[0].rows;
    assert_eq!(rows.len(), 4);
    assert_eq!(rows[0].cells[0].text, "A");
    assert_eq!(rows[1].cells[0].coordinates(), (1, 0));
    assert_eq!(rows[2].cells[0].span(), Some((2, 2)));
    assert_eq!(rows[2].cells[1].merge(), CellMerge::Covered);
    assert_eq!(rows[2].cells[2].text, "C");
    assert_eq!(
        rows[2].cells[2]
            .matrix_span()
            .map(|span| (span.rows(), span.columns())),
        Some((3, 2))
    );
    assert_eq!(rows[2].cells[2].coordinates(), (2, 2));
    assert_eq!(rows[3].cells.len(), 2);
    assert!(
        rows[3]
            .cells
            .iter()
            .all(|cell| cell.merge() == CellMerge::Covered)
    );
}

#[test]
fn parses_sheet_content_with_arbitrary_namespace_prefixes() {
    let xml = r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:x="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><o:body><o:spreadsheet><t:table t:name="A&amp;B"><t:table-row t:number-rows-repeated="2"><t:table-cell o:value-type="string" t:style-name="Style&amp;One" t:protected="1"><x:p>one<x:s x:c="2"/>two<x:tab/>three<x:line-break/>four</x:p></t:table-cell><t:covered-table-cell/></t:table-row></t:table></o:spreadsheet></o:body></o:document-content>"#;
    let sheets = Parser::parse_sheets(xml).unwrap();
    assert_eq!(sheets[0].name, "A&B");
    assert_eq!(sheets[0].rows.len(), 2);
    for (row_index, row) in sheets[0].rows.iter().enumerate() {
        assert_eq!(row.cells[0].coordinates(), (row_index, 0));
        assert_eq!(row.cells[0].style_name(), Some("Style&One"));
        assert_eq!(row.cells[0].protected(), Some(true));
        assert_eq!(row.cells[0].text, "one  two\tthree\nfour");
        assert_eq!(row.cells[1].merge(), CellMerge::Covered);
    }
}

#[test]
fn parses_repeated_row_and_column_structural_metadata() {
    let xml = r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><o:body><o:spreadsheet><t:table t:name="Structure"><t:table-column t:number-columns-repeated="2" t:style-name="Col&amp;Style" t:default-cell-style-name="CellStyle" t:visibility="collapse"/><t:table-column t:visibility="filter"></t:table-column><t:table-row t:number-rows-repeated="2" t:style-name="RowStyle" t:default-cell-style-name="RowCell" t:visibility="filter"><t:table-cell/></t:table-row></t:table></o:spreadsheet></o:body></o:document-content>"#;
    let sheets = Parser::parse_sheets(xml).unwrap();
    let sheet = &sheets[0];
    assert_eq!(sheet.columns.len(), 3);
    assert_eq!(sheet.columns[0].index, 0);
    assert_eq!(sheet.columns[1].index, 1);
    assert_eq!(sheet.columns[0].style_name.as_deref(), Some("Col&Style"));
    assert_eq!(
        sheet.columns[0].default_cell_style_name.as_deref(),
        Some("CellStyle")
    );
    assert_eq!(sheet.columns[0].visibility, TableVisibility::Collapse);
    assert_eq!(sheet.columns[2].visibility, TableVisibility::Filter);
    assert_eq!(sheet.rows.len(), 2);
    assert!(sheet.rows.iter().all(|row| {
        row.style_name.as_deref() == Some("RowStyle")
            && row.default_cell_style_name.as_deref() == Some("RowCell")
            && row.visibility == TableVisibility::Filter
    }));
}

#[test]
fn parses_nested_groups_and_header_ranges() {
    let xml = r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><o:body><o:spreadsheet><t:table t:name="Outline"><t:table-column-group t:display="false"><t:table-header-columns><t:table-column/></t:table-header-columns><t:table-column-group><t:table-column t:number-columns-repeated="2"/></t:table-column-group></t:table-column-group><t:table-row-group t:display="false"><t:table-header-rows><t:table-row/></t:table-header-rows><t:table-row-group><t:table-row t:number-rows-repeated="2"/></t:table-row-group></t:table-row-group></t:table></o:spreadsheet></o:body></o:document-content>"#;
    let sheets = Parser::parse_sheets(xml).unwrap();
    assert_eq!(
        sheets[0].column_structure,
        vec![TableStructure::Group(TableGroup {
            display: false,
            children: vec![
                TableStructure::Header(TableRange { start: 0, end: 1 }),
                TableStructure::Group(TableGroup {
                    display: true,
                    children: vec![TableStructure::Range(TableRange { start: 1, end: 3 })],
                }),
            ],
        })]
    );
    assert_eq!(
        sheets[0].row_structure,
        vec![TableStructure::Group(TableGroup {
            display: false,
            children: vec![
                TableStructure::Header(TableRange { start: 0, end: 1 }),
                TableStructure::Group(TableGroup {
                    display: true,
                    children: vec![TableStructure::Range(TableRange { start: 1, end: 3 })],
                }),
            ],
        })]
    );
}

#[test]
fn parses_sheet_style_and_print_settings() {
    let xml = r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><o:body><o:spreadsheet><t:table t:name="Print" t:style-name="Sheet&amp;Style" t:template-name="TemplateOne" t:use-first-row-styles="true" t:use-last-row-styles="0" t:use-first-column-styles="1" t:use-last-column-styles="false" t:use-banding-rows-styles="true" t:use-banding-columns-styles="false" t:print="false" t:print-ranges="$Print.$A$1:$B$2 'Q1 Sales'.$C$3:$D$4"></t:table></o:spreadsheet></o:body></o:document-content>"#;
    let sheets = Parser::parse_sheets(xml).unwrap();
    let sheet = &sheets[0];
    assert_eq!(sheet.style.style_name.as_deref(), Some("Sheet&Style"));
    assert_eq!(sheet.style.template_name.as_deref(), Some("TemplateOne"));
    assert_eq!(sheet.style.usage.use_first_row_styles, Some(true));
    assert_eq!(sheet.style.usage.use_last_row_styles, Some(false));
    assert_eq!(sheet.style.usage.use_first_column_styles, Some(true));
    assert_eq!(sheet.style.usage.use_last_column_styles, Some(false));
    assert_eq!(sheet.style.usage.use_banding_row_styles, Some(true));
    assert_eq!(sheet.style.usage.use_banding_column_styles, Some(false));
    assert!(!sheet.print_settings.printable);
    assert_eq!(
        sheet.print_settings.ranges,
        ["$Print.$A$1:$B$2", "'Q1 Sales'.$C$3:$D$4"]
    );
}

#[test]
fn parses_sheet_title_description_and_scenario() {
    let xml = r##"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:l="http://www.w3.org/1999/xlink"><o:body><o:spreadsheet><t:table t:name="Scenario"><t:title>Quarter &amp; Forecast</t:title><t:desc><![CDATA[Best < worst]]></t:desc><t:table-source l:type="simple" l:href="../Q1&amp;Q2.ods" l:actuate="onRequest" t:mode="copy-results-only" t:table-name="Source Sheet" t:filter-name="calc8" t:filter-options="A&amp;B" t:refresh-delay="P1DT2H3.5S"/><t:scenario t:scenario-ranges="$Scenario.$A$1:$B$2 'Q1 Sales'.$C$3:$D$4" t:is-active="true" t:display-border="0" t:border-color="#12AbEF" t:copy-back="1" t:copy-styles="false" t:copy-formulas="true" t:comment="Best &amp; worst" t:protected="false"/></t:table></o:spreadsheet></o:body></o:document-content>"##;
    let sheets = Parser::parse_sheets(xml).unwrap();
    let sheet = &sheets[0];
    assert_eq!(sheet.title.as_deref(), Some("Quarter & Forecast"));
    assert_eq!(sheet.description.as_deref(), Some("Best < worst"));
    let source = sheet.table_source.as_ref().unwrap();
    assert_eq!(source.href, "../Q1&Q2.ods");
    assert_eq!(source.mode, Some(TableSourceMode::CopyResultsOnly));
    assert_eq!(source.table_name.as_deref(), Some("Source Sheet"));
    assert!(source.actuate_on_request);
    assert_eq!(source.filter_name.as_deref(), Some("calc8"));
    assert_eq!(source.filter_options.as_deref(), Some("A&B"));
    assert_eq!(source.refresh_delay.as_deref(), Some("P1DT2H3.5S"));
    let scenario = sheet.scenario.as_ref().unwrap();
    assert_eq!(
        scenario.ranges,
        ["$Scenario.$A$1:$B$2", "'Q1 Sales'.$C$3:$D$4"]
    );
    assert!(scenario.is_active);
    assert_eq!(scenario.display_border, Some(false));
    assert_eq!(scenario.border_color.as_deref(), Some("#12AbEF"));
    assert_eq!(scenario.copy_back, Some(true));
    assert_eq!(scenario.copy_styles, Some(false));
    assert_eq!(scenario.copy_formulas, Some(true));
    assert_eq!(scenario.comment.as_deref(), Some("Best & worst"));
    assert_eq!(scenario.protected, Some(false));
}

#[test]
fn rejects_invalid_or_dangerous_repetition_counts() {
    let zero = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><office:body><office:spreadsheet><table:table><table:table-row table:number-rows-repeated="0"/></table:table></office:spreadsheet></office:body></office:document-content>"#;
    assert!(Parser::parse_sheets(zero).is_err());

    let excessive = format!(
        r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><office:body><office:spreadsheet><table:table><table:table-row table:number-rows-repeated="{}"/></table:table></office:spreadsheet></office:body></office:document-content>"#,
        MAX_EXPANDED_ROWS_PER_SHEET + 1
    );
    assert!(Parser::parse_sheets(&excessive).is_err());

    let excessive_columns = format!(
        r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><office:body><office:spreadsheet><table:table><table:table-column table:number-columns-repeated="{}"/></table:table></office:spreadsheet></office:body></office:document-content>"#,
        MAX_EXPANDED_COLUMNS_PER_SHEET + 1
    );
    assert!(Parser::parse_sheets(&excessive_columns).is_err());

    let invalid_visibility = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><office:body><office:spreadsheet><table:table><table:table-column table:visibility="hidden"/></table:table></office:spreadsheet></office:body></office:document-content>"#;
    assert!(Parser::parse_sheets(invalid_visibility).is_err());

    let invalid_group_display = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><office:body><office:spreadsheet><table:table><table:table-row-group table:display="collapsed"><table:table-row/></table:table-row-group></table:table></office:spreadsheet></office:body></office:document-content>"#;
    assert!(Parser::parse_sheets(invalid_group_display).is_err());

    let empty_group = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><office:body><office:spreadsheet><table:table><table:table-column-group/></table:table></office:spreadsheet></office:body></office:document-content>"#;
    assert!(Parser::parse_sheets(empty_group).is_err());

    let invalid_print = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><office:body><office:spreadsheet><table:table table:print="yes"></table:table></office:spreadsheet></office:body></office:document-content>"#;
    assert!(Parser::parse_sheets(invalid_print).is_err());

    let invalid_print_ranges = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><office:body><office:spreadsheet><table:table table:print-ranges="'Unclosed Sheet.$A$1"></table:table></office:spreadsheet></office:body></office:document-content>"#;
    assert!(Parser::parse_sheets(invalid_print_ranges).is_err());

    let incomplete_scenario = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><office:body><office:spreadsheet><table:table><table:scenario table:scenario-ranges=".A1:.B2"/></table:table></office:spreadsheet></office:body></office:document-content>"#;
    assert!(Parser::parse_sheets(incomplete_scenario).is_err());

    let invalid_scenario_color = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><office:body><office:spreadsheet><table:table><table:scenario table:scenario-ranges=".A1:.B2" table:is-active="false" table:border-color="red"/></table:table></office:spreadsheet></office:body></office:document-content>"#;
    assert!(Parser::parse_sheets(invalid_scenario_color).is_err());

    let duplicate_scenarios = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><office:body><office:spreadsheet><table:table><table:scenario table:scenario-ranges=".A1:.B2" table:is-active="false"/><table:scenario table:scenario-ranges=".C1:.D2" table:is-active="true"/></table:table></office:spreadsheet></office:body></office:document-content>"#;
    assert!(Parser::parse_sheets(duplicate_scenarios).is_err());

    let duplicate_titles = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><office:body><office:spreadsheet><table:table><table:title>First</table:title><table:title>Second</table:title></table:table></office:spreadsheet></office:body></office:document-content>"#;
    assert!(Parser::parse_sheets(duplicate_titles).is_err());

    let duplicate_descriptions = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><office:body><office:spreadsheet><table:table><table:desc>First</table:desc><table:desc>Second</table:desc></table:table></office:spreadsheet></office:body></office:document-content>"#;
    assert!(Parser::parse_sheets(duplicate_descriptions).is_err());

    let invalid_sources = [
        r#"<table:table-source xlink:href="a.ods"/>"#,
        r#"<table:table-source xlink:type="extended" xlink:href="a.ods"/>"#,
        r#"<table:table-source xlink:type="simple"/>"#,
        r#"<table:table-source xlink:type="simple" xlink:href="a.ods" xlink:actuate="onLoad"/>"#,
        r#"<table:table-source xlink:type="simple" xlink:href="a.ods" table:mode="values"/>"#,
        r#"<table:table-source xlink:type="simple" xlink:href="a.ods" table:refresh-delay="15 minutes"/>"#,
    ];
    for source in invalid_sources {
        let xml = format!(
            r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:xlink="http://www.w3.org/1999/xlink"><office:body><office:spreadsheet><table:table>{source}</table:table></office:spreadsheet></office:body></office:document-content>"#
        );
        assert!(Parser::parse_sheets(&xml).is_err(), "{source}");
    }

    let duplicate_sources = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:xlink="http://www.w3.org/1999/xlink"><office:body><office:spreadsheet><table:table><table:table-source xlink:type="simple" xlink:href="a.ods"/><table:table-source xlink:type="simple" xlink:href="b.ods"/></table:table></office:spreadsheet></office:body></office:document-content>"#;
    assert!(Parser::parse_sheets(duplicate_sources).is_err());
}

#[test]
fn test_sheet_builder() {
    let mut builder = SheetBuilder::new("TestSheet".to_string());

    let row1 = Row {
        cells: vec![],
        index: 0,
        style_name: None,
        default_cell_style_name: None,
        visibility: Default::default(),
    };
    builder.add_row(row1);

    let row2 = Row {
        cells: vec![Cell {
            value: CellValue::Text("A1".to_string()),
            text: "A1".to_string(),
            formula: None,
            annotation: None,
            hyperlinks: Vec::new(),
            rich_text: None,
            range_source: None,
            detective: None,
            validation_name: None,
            style_name: None,
            matrix_span: None,
            merge: Default::default(),
            protect: None,
            protected: None,
            row: 0,
            col: 0,
        }],
        index: 0,
        style_name: None,
        default_cell_style_name: None,
        visibility: Default::default(),
    };
    builder.add_row(row2);

    let sheet = builder.build().unwrap();
    assert_eq!(sheet.name, "TestSheet");
    assert_eq!(sheet.rows.len(), 2);
    assert_eq!(sheet.rows[0].index, 0);
    assert_eq!(sheet.rows[1].index, 1);
}

#[test]
fn test_row_builder() {
    let mut builder = RowBuilder::new();

    let cell1 = Cell {
        value: CellValue::Text("A".to_string()),
        text: "A".to_string(),
        formula: None,
        annotation: None,
        hyperlinks: Vec::new(),
        rich_text: None,
        range_source: None,
        detective: None,
        validation_name: None,
        style_name: None,
        matrix_span: None,
        merge: Default::default(),
        protect: None,
        protected: None,
        row: 0,
        col: 0,
    };
    builder.add_cell(cell1);

    let cell2 = Cell {
        value: CellValue::Number(42.0),
        text: "42".to_string(),
        formula: None,
        annotation: None,
        hyperlinks: Vec::new(),
        rich_text: None,
        range_source: None,
        detective: None,
        validation_name: None,
        style_name: None,
        matrix_span: None,
        merge: Default::default(),
        protect: None,
        protected: None,
        row: 0,
        col: 0,
    };
    builder.add_cell(cell2);

    let row = builder.build();
    assert_eq!(row.cells.len(), 2);
    assert_eq!(row.cells[0].col, 0);
    assert_eq!(row.cells[1].col, 1);
}

#[test]
fn test_cell_builder_float_types() {
    // Test "float" value type
    let builder = CellBuilder {
        value_type: Some("float".to_string()),
        value_str: Some("123.45".to_string()),
        currency: None,
        formula: None,
        annotation: None,
        hyperlinks: Vec::new(),
        range_source: None,
        detective: None,
        validation_name: None,
        style_name: None,
        matrix_span: None,
        merge: Default::default(),
        protect: None,
        protected: None,
        repeated: 1,
    };
    let cell = builder.build("123.45", None);
    match cell.value {
        CellValue::Number(n) => assert!((n - 123.45).abs() < f64::EPSILON),
        _ => panic!("Expected Number for float"),
    }

    // Test "double" value type
    let builder = CellBuilder {
        value_type: Some("double".to_string()),
        value_str: Some("99.99".to_string()),
        currency: None,
        formula: None,
        annotation: None,
        hyperlinks: Vec::new(),
        range_source: None,
        detective: None,
        validation_name: None,
        style_name: None,
        matrix_span: None,
        merge: Default::default(),
        protect: None,
        protected: None,
        repeated: 1,
    };
    let cell = builder.build("99.99", None);
    match cell.value {
        CellValue::Number(n) => assert!((n - 99.99).abs() < f64::EPSILON),
        _ => panic!("Expected Number for double"),
    }

    // Test "decimal" value type
    let builder = CellBuilder {
        value_type: Some("decimal".to_string()),
        value_str: Some("0.001".to_string()),
        currency: None,
        formula: None,
        annotation: None,
        hyperlinks: Vec::new(),
        range_source: None,
        detective: None,
        validation_name: None,
        style_name: None,
        matrix_span: None,
        merge: Default::default(),
        protect: None,
        protected: None,
        repeated: 1,
    };
    let cell = builder.build("0.001", None);
    match cell.value {
        CellValue::Number(n) => assert!((n - 0.001).abs() < f64::EPSILON),
        _ => panic!("Expected Number for decimal"),
    }
}

#[test]
fn test_cell_builder_invalid_number_fallback() {
    let builder = CellBuilder {
        value_type: Some("float".to_string()),
        value_str: Some("not-a-number".to_string()),
        currency: None,
        formula: None,
        annotation: None,
        hyperlinks: Vec::new(),
        range_source: None,
        detective: None,
        validation_name: None,
        style_name: None,
        matrix_span: None,
        merge: Default::default(),
        protect: None,
        protected: None,
        repeated: 1,
    };
    let cell = builder.build("some text", None);
    match cell.value {
        CellValue::Text(t) => assert_eq!(t, "some text"),
        _ => panic!("Expected Text fallback for invalid number"),
    }
}

#[test]
fn test_cell_builder_boolean_variations() {
    // Test "false" boolean
    let builder = CellBuilder {
        value_type: Some("boolean".to_string()),
        value_str: Some("false".to_string()),
        currency: None,
        formula: None,
        annotation: None,
        hyperlinks: Vec::new(),
        range_source: None,
        detective: None,
        validation_name: None,
        style_name: None,
        matrix_span: None,
        merge: Default::default(),
        protect: None,
        protected: None,
        repeated: 1,
    };
    let cell = builder.build("FALSE", None);
    match cell.value {
        CellValue::Boolean(b) => assert!(!b),
        _ => panic!("Expected Boolean false"),
    }

    // Test invalid boolean value (falls back to text)
    let builder = CellBuilder {
        value_type: Some("boolean".to_string()),
        value_str: Some("maybe".to_string()),
        currency: None,
        formula: None,
        annotation: None,
        hyperlinks: Vec::new(),
        range_source: None,
        detective: None,
        validation_name: None,
        style_name: None,
        matrix_span: None,
        merge: Default::default(),
        protect: None,
        protected: None,
        repeated: 1,
    };
    let cell = builder.build("maybe", None);
    match cell.value {
        CellValue::Text(t) => assert_eq!(t, "maybe"),
        _ => panic!("Expected Text for invalid boolean"),
    }
}

#[test]
fn test_cell_builder_empty_text() {
    let builder = CellBuilder {
        value_type: None,
        value_str: None,
        currency: None,
        formula: None,
        annotation: None,
        hyperlinks: Vec::new(),
        range_source: None,
        detective: None,
        validation_name: None,
        style_name: None,
        matrix_span: None,
        merge: Default::default(),
        protect: None,
        protected: None,
        repeated: 1,
    };
    let cell = builder.build("   ", None);
    match cell.value {
        CellValue::Empty => {},
        _ => panic!("Expected Empty for whitespace-only text"),
    }
}

#[test]
fn test_cell_builder_currency_default() {
    let builder = CellBuilder {
        value_type: Some("currency".to_string()),
        value_str: Some("50".to_string()),
        currency: None, // No currency specified
        formula: None,
        annotation: None,
        hyperlinks: Vec::new(),
        range_source: None,
        detective: None,
        validation_name: None,
        style_name: None,
        matrix_span: None,
        merge: Default::default(),
        protect: None,
        protected: None,
        repeated: 1,
    };
    let cell = builder.build("$50", None);
    match cell.value {
        CellValue::Currency(amount, currency) => {
            assert!((amount - 50.0).abs() < f64::EPSILON);
            assert_eq!(currency, "USD"); // Default
        },
        _ => panic!("Expected Currency with default USD"),
    }
}

#[test]
fn test_parse_invalid_xml() {
    let invalid_xml = "<invalid>unclosed tag";
    let result = Parser::parse_sheets(invalid_xml);
    // The parser may return Ok with empty sheets or Err depending on implementation
    // Either behavior is acceptable - we just verify it doesn't panic
    match result {
        Ok(sheets) => {
            // If parsing succeeds, we should get 0 sheets
            assert_eq!(sheets.len(), 0);
        },
        Err(_) => {
            // Error is also acceptable
        },
    }
}

#[test]
fn parses_global_and_sheet_local_named_definitions_with_namespace_aliases() {
    let xml = r#"<?xml version="1.0"?>
            <o:document-content
                xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
                xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
                xmlns:f="urn:example:formula">
              <o:body><o:spreadsheet>
                <t:table t:name="Sales &amp; Tax">
                  <t:table-column/><t:table-row><t:table-cell/></t:table-row>
                  <t:named-expressions>
                    <t:named-range t:name="LocalRange"
                      t:cell-range-address="$'Sales &amp; Tax'.$A$1:.$B$2"
                      t:base-cell-address="$'Sales &amp; Tax'.$A$1"
                      t:range-usable-as="print-range filter repeat-row repeat-column"/>
                  </t:named-expressions>
                </t:table>
                <t:named-expressions>
                  <t:named-expression t:name="TaxRate"
                    t:expression="f:=0.2" t:base-cell-address="$'Sales &amp; Tax'.$A$1"/>
                </t:named-expressions>
              </o:spreadsheet></o:body>
            </o:document-content>"#;

    let definitions = Parser::parse_named_definitions(xml).unwrap();
    assert_eq!(definitions.len(), 2);
    let NamedDefinition::Range(range) = &definitions[0] else {
        panic!("expected named range");
    };
    assert_eq!(range.name, "LocalRange");
    assert_eq!(
        range.scope,
        NamedDefinitionScope::Sheet("Sales & Tax".to_string())
    );
    assert_eq!(range.usable_as.len(), 4);
    assert_eq!(
        range.base_cell_address.as_deref(),
        Some("$'Sales & Tax'.$A$1")
    );

    let NamedDefinition::Expression(expression) = &definitions[1] else {
        panic!("expected named expression");
    };
    assert_eq!(expression.name, "TaxRate");
    assert_eq!(expression.expression, "f:=0.2");
    assert_eq!(
        expression.formula_namespace.as_ref().unwrap().uri,
        "urn:example:formula"
    );
    assert_eq!(expression.scope, NamedDefinitionScope::Global);
}

#[test]
fn named_definition_parser_rejects_missing_attributes_and_invalid_usage() {
    let missing_address = r#"<office:spreadsheet
            xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
            xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
            <table:named-expressions><table:named-range table:name="Broken"/>
            </table:named-expressions></office:spreadsheet>"#;
    assert!(Parser::parse_named_definitions(missing_address).is_err());

    let invalid_usage = r#"<office:spreadsheet
            xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
            xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
            <table:named-expressions><table:named-range table:name="Broken"
              table:cell-range-address="$Sheet1.$A$1" table:range-usable-as="chart"/>
            </table:named-expressions></office:spreadsheet>"#;
    assert!(Parser::parse_named_definitions(invalid_usage).is_err());
}

#[test]
fn sheet_parser_ignores_dde_cache_tables() {
    let xml = r#"<o:document-content
            xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
            xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
          <o:body><o:spreadsheet>
            <t:dde-links><t:dde-link>
              <o:dde-source o:dde-application="soffice" o:dde-topic="topic" o:dde-item="item"/>
              <t:table t:name="Cached"><t:table-row><t:table-cell o:value-type="string"><text:p xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">cached</text:p></t:table-cell></t:table-row></t:table>
            </t:dde-link></t:dde-links>
            <t:table t:name="Visible"><t:table-row><t:table-cell o:value-type="string"/></t:table-row></t:table>
            <t:table t:name="Empty"/>
          </o:spreadsheet></o:body>
        </o:document-content>"#;

    let sheets = Parser::parse_sheets(xml).unwrap();
    assert_eq!(sheets.len(), 2);
    assert_eq!(sheets[0].name, "Visible");
    assert_eq!(sheets[1].name, "Empty");
}

const HYPERLINK_DOCUMENT_PREFIX: &str = r#"<office:document-content
        xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
        xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
        xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
        xmlns:xlink="http://www.w3.org/1999/xlink"
        xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0">
      <office:body><office:spreadsheet>
        <table:table table:name="Links"><table:table-row>"#;
const HYPERLINK_DOCUMENT_SUFFIX: &str =
    "</table:table-row></table:table></office:spreadsheet></office:body></office:document-content>";

fn hyperlink_document(cells: &str) -> String {
    format!("{HYPERLINK_DOCUMENT_PREFIX}{cells}{HYPERLINK_DOCUMENT_SUFFIX}")
}

#[test]
fn parses_cell_hyperlinks_with_metadata_and_document_order() {
    let attributes = concat!(
        r#"xlink:href="https://example.com/" xlink:type="simple" "#,
        r#"office:name="Example" office:title="Example site" "#,
        r#"office:target-frame-name="_blank" text:style-name="Internet_20_link" "#,
        r#"xlink:show="new" xlink:actuate="onRequest" "#,
        r#"text:visited-style-name="Visited_20_Internet_20_Link""#,
    );
    let xml = hyperlink_document(&format!(
        concat!(
            r#"<table:table-cell office:value-type="string">"#,
            "<text:p>See <text:a {attributes}>the ",
            "<text:span>example</text:span> site</text:a> and ",
            r##"<text:a xlink:href="#Sheet2.B10">an internal target</text:a>.</text:p>"##,
            "</table:table-cell>",
        ),
        attributes = attributes,
    ));

    let sheets = Parser::parse_sheets(&xml).unwrap();
    let cell = &sheets[0].rows[0].cells[0];
    assert_eq!(cell.text, "See the example site and an internal target.");
    assert_eq!(cell.hyperlinks().len(), 2);

    let first = cell.hyperlink().unwrap();
    assert_eq!(first.href(), "https://example.com/");
    assert_eq!(first.text(), "the example site");
    assert_eq!(first.range(), 4..20);
    assert_eq!(first.name.as_deref(), Some("Example"));
    assert_eq!(first.title.as_deref(), Some("Example site"));
    assert_eq!(first.target_frame_name.as_deref(), Some("_blank"));
    assert_eq!(first.show, Some(TextHyperlinkShow::New));
    assert_eq!(first.actuate, Some(TextHyperlinkActuate::OnRequest));
    assert_eq!(first.style_name.as_deref(), Some("Internet_20_link"));
    assert_eq!(
        first.visited_style_name.as_deref(),
        Some("Visited_20_Internet_20_Link")
    );

    let second = &cell.hyperlinks()[1];
    assert_eq!(second.href, "#Sheet2.B10");
    assert_eq!(second.text, "an internal target");
    assert_eq!(second.range(), 25..43);
    assert!(second.name.is_none());
    assert!(second.target_frame_name.is_none());
    assert!(second.show.is_none());
    assert!(second.actuate.is_none());
}

#[test]
fn parses_hyperlinks_with_namespace_aliases_and_repeated_cells() {
    let xml = r#"<o:document-content
            xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
            xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
            xmlns:tx="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
            xmlns:x="http://www.w3.org/1999/xlink">
          <o:body><o:spreadsheet><t:table t:name="Links"><t:table-row>
            <t:table-cell t:number-columns-repeated="2" o:value-type="string">
              <tx:p><tx:a x:href="mailto:someone@example.com">mail</tx:a></tx:p>
            </t:table-cell>
          </t:table-row></t:table></o:spreadsheet></o:body></o:document-content>"#;

    let sheets = Parser::parse_sheets(xml).unwrap();
    let row = &sheets[0].rows[0];
    assert_eq!(row.cells.len(), 2);
    for cell in &row.cells {
        assert!(cell.has_hyperlinks());
        let link = cell.hyperlink().unwrap();
        assert_eq!(link.href(), "mailto:someone@example.com");
        assert_eq!(link.text(), "mail");
    }
}

#[test]
fn parses_self_closing_hyperlink_with_empty_text() {
    let xml = hyperlink_document(
        r#"<table:table-cell office:value-type="string">
              <text:p>before <text:a xlink:href="https://example.com/"/> after</text:p>
            </table:table-cell>"#,
    );

    let sheets = Parser::parse_sheets(&xml).unwrap();
    let cell = &sheets[0].rows[0].cells[0];
    assert_eq!(cell.text, "before  after");
    assert_eq!(cell.hyperlinks().len(), 1);
    assert_eq!(cell.hyperlinks()[0].href, "https://example.com/");
    assert_eq!(cell.hyperlinks()[0].text, "");
    assert_eq!(cell.hyperlinks()[0].range(), 7..7);
}

#[test]
fn preserves_mixed_text_anchor_range_from_libreoffice_fods() {
    let source = include_str!(
        "../../../../test-data/libreoffice-core/sc/qa/unit/data/functions/text/fods/encodeurl.fods"
    );
    let anchor = r#"<text:a xlink:href="http://www.test/libreOffice" xlink:type="simple">"#;
    let anchor_start = source.find(anchor).unwrap();
    let cell_start = source[..anchor_start].rfind("<table:table-cell").unwrap();
    let cell_end = anchor_start
        + source[anchor_start..].find("</table:table-cell>").unwrap()
        + "</table:table-cell>".len();
    let sheets = Parser::parse_sheets(&hyperlink_document(&source[cell_start..cell_end])).unwrap();
    let cell = sheets
        .iter()
        .flat_map(|sheet| sheet.rows.iter())
        .flat_map(|row| row.cells.iter())
        .find(|cell| {
            cell.hyperlinks()
                .iter()
                .any(|link| link.href() == "http://www.test/libreOffice")
        })
        .unwrap();
    let link = cell
        .hyperlinks()
        .iter()
        .find(|link| link.href() == "http://www.test/libreOffice")
        .unwrap();

    assert_eq!(link.range(), 0..link.text().len());
    assert!(cell.text.starts_with(link.text()));
    assert!(cell.text.ends_with("agJohn01Czech Republic"));
}

#[test]
fn hyperlink_text_includes_whitespace_and_break_elements() {
    let xml = hyperlink_document(
        r#"<table:table-cell office:value-type="string">
              <text:p><text:a xlink:href="https://example.com/">a<text:s text:c="2"/>b<text:line-break/>c</text:a></text:p>
            </table:table-cell>"#,
    );

    let sheets = Parser::parse_sheets(&xml).unwrap();
    let cell = &sheets[0].rows[0].cells[0];
    assert_eq!(cell.hyperlinks()[0].text, "a  b\nc");
}

#[test]
fn rejects_hyperlink_without_href() {
    let xml = hyperlink_document(
        r#"<table:table-cell office:value-type="string">
              <text:p><text:a office:name="broken">no target</text:a></text:p>
            </table:table-cell>"#,
    );

    let error = Parser::parse_sheets(&xml).err().expect("parse must fail");
    assert!(error.to_string().contains("xlink:href"));
}

#[test]
fn rejects_nested_hyperlinks() {
    let xml = hyperlink_document(
        r#"<table:table-cell office:value-type="string">
              <text:p><text:a xlink:href="https://a.example/">outer
                <text:a xlink:href="https://b.example/">inner</text:a></text:a></text:p>
            </table:table-cell>"#,
    );

    let error = Parser::parse_sheets(&xml).err().expect("parse must fail");
    assert!(error.to_string().contains("nested"));
}

#[test]
fn rejects_cell_rich_text_beyond_the_depth_limit() {
    let nested = format!(
        "{}x{}",
        "<text:span>".repeat(128),
        "</text:span>".repeat(128)
    );
    let xml = hyperlink_document(&format!(
        r#"<table:table-cell office:value-type="string"><text:p>{nested}</text:p></table:table-cell>"#
    ));

    let error = Parser::parse_sheets(&xml)
        .err()
        .expect("overly deep rich text must fail");
    assert!(error.to_string().contains("depth limit"));
}

#[test]
fn rejects_hyperlink_with_invalid_xlink_type() {
    let xml = hyperlink_document(
        r#"<table:table-cell office:value-type="string">
              <text:p><text:a xlink:href="https://example.com/" xlink:type="extended">x</text:a></text:p>
            </table:table-cell>"#,
    );

    let error = Parser::parse_sheets(&xml).err().expect("parse must fail");
    assert!(error.to_string().contains("xlink:type"));
}

#[test]
fn rejects_hyperlink_with_invalid_xlink_show_or_actuate() {
    for attributes in [r#"xlink:show="embed""#, r#"xlink:actuate="onLoad""#] {
        let xml = hyperlink_document(&format!(
            r#"<table:table-cell office:value-type="string"><text:p><text:a xlink:href="https://example.com/" {attributes}>x</text:a></text:p></table:table-cell>"#
        ));
        assert!(Parser::parse_sheets(&xml).is_err());
    }
}

#[test]
fn annotation_hyperlinks_are_not_reported_as_cell_hyperlinks() {
    let xml = hyperlink_document(
        r#"<table:table-cell office:value-type="string">
              <office:annotation><text:p><text:a xlink:href="https://note.example/">note link</text:a></text:p></office:annotation>
              <text:p>plain</text:p>
            </table:table-cell>"#,
    );

    let sheets = Parser::parse_sheets(&xml).unwrap();
    let cell = &sheets[0].rows[0].cells[0];
    assert!(cell.annotation().is_some());
    assert!(!cell.has_hyperlinks());
    assert_eq!(cell.text, "plain");
}
