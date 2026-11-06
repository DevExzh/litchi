use super::model::{parse_cell_ref, parse_range};
use super::*;
use crate::conditional_formatting::IconSet;
use crate::sort::{SortBy, SortCondition, SortMethod, SortState};

#[test]
fn test_table_style_info_new() {
    let style = TableStyleInfo::new();
    assert!(style.name.is_none());
    assert!(style.show_first_column.is_none());
    assert!(style.show_last_column.is_none());
    assert!(style.show_row_stripes.is_none());
    assert!(style.show_column_stripes.is_none());
}

#[test]
fn test_table_style_info_default() {
    let style: TableStyleInfo = Default::default();
    assert!(style.name.is_none());
}

#[test]
fn test_table_style_info_parse() {
    let tag = r#"name="TableStyleMedium2" showFirstColumn="1" showLastColumn="0" showRowStripes="1" showColumnStripes="0""#;
    let style = TableStyleInfo::parse(tag).unwrap();
    assert_eq!(style.name, Some("TableStyleMedium2".to_string()));
    assert_eq!(style.show_first_column, Some(true));
    assert_eq!(style.show_last_column, Some(false));
    assert_eq!(style.show_row_stripes, Some(true));
    assert_eq!(style.show_column_stripes, Some(false));
}

#[test]
fn test_table_style_info_parse_partial() {
    let tag = r#"name="TableStyleLight1" showRowStripes="true""#;
    let style = TableStyleInfo::parse(tag).unwrap();
    assert_eq!(style.name, Some("TableStyleLight1".to_string()));
    assert_eq!(style.show_row_stripes, Some(true));
    assert!(style.show_first_column.is_none());
}

#[test]
fn test_totals_row_function_as_str() {
    assert_eq!(TotalsRowFunction::Sum.as_str(), "sum");
    assert_eq!(TotalsRowFunction::Min.as_str(), "min");
    assert_eq!(TotalsRowFunction::Max.as_str(), "max");
    assert_eq!(TotalsRowFunction::Average.as_str(), "average");
    assert_eq!(TotalsRowFunction::Count.as_str(), "count");
    assert_eq!(TotalsRowFunction::CountNums.as_str(), "countNums");
    assert_eq!(TotalsRowFunction::StdDev.as_str(), "stdDev");
    assert_eq!(TotalsRowFunction::Var.as_str(), "var");
    assert_eq!(TotalsRowFunction::Custom.as_str(), "custom");
}

#[test]
fn test_totals_row_function_parse() {
    assert_eq!(
        TotalsRowFunction::parse("sum"),
        Some(TotalsRowFunction::Sum)
    );
    assert_eq!(
        TotalsRowFunction::parse("min"),
        Some(TotalsRowFunction::Min)
    );
    assert_eq!(
        TotalsRowFunction::parse("max"),
        Some(TotalsRowFunction::Max)
    );
    assert_eq!(
        TotalsRowFunction::parse("average"),
        Some(TotalsRowFunction::Average)
    );
    assert_eq!(
        TotalsRowFunction::parse("count"),
        Some(TotalsRowFunction::Count)
    );
    assert_eq!(
        TotalsRowFunction::parse("countNums"),
        Some(TotalsRowFunction::CountNums)
    );
    assert_eq!(
        TotalsRowFunction::parse("stdDev"),
        Some(TotalsRowFunction::StdDev)
    );
    assert_eq!(
        TotalsRowFunction::parse("var"),
        Some(TotalsRowFunction::Var)
    );
    assert_eq!(
        TotalsRowFunction::parse("custom"),
        Some(TotalsRowFunction::Custom)
    );
    assert_eq!(TotalsRowFunction::parse("invalid"), None);
}

#[test]
fn test_table_column_new() {
    let col = TableColumn::new(1u32, "Sales");
    assert_eq!(col.id, 1);
    assert_eq!(col.name, "Sales");
    assert!(col.unique_name.is_none());
    assert!(col.totals_row_function.is_none());
    assert!(col.totals_row_label.is_none());
    assert!(col.calculated_column_formula.is_none());
    assert!(col.totals_row_formula.is_none());
}

#[test]
fn test_table_type_as_str() {
    assert_eq!(TableType::Worksheet.as_str(), "worksheet");
    assert_eq!(TableType::Xml.as_str(), "xml");
    assert_eq!(TableType::QueryTable.as_str(), "queryTable");
}

#[test]
fn test_table_type_parse() {
    assert_eq!(TableType::parse("worksheet"), Some(TableType::Worksheet));
    assert_eq!(TableType::parse("xml"), Some(TableType::Xml));
    assert_eq!(TableType::parse("queryTable"), Some(TableType::QueryTable));
    assert_eq!(TableType::parse("invalid"), None);
}

#[test]
fn test_table_new() {
    let table = Table::new(1u32, "Table1", "A1:D10");
    assert_eq!(table.id, 1);
    assert_eq!(table.name, "Table1");
    assert_eq!(table.display_name, "Table1");
    assert_eq!(table.ref_range, "A1:D10");
    assert_eq!(table.header_row_count, Some(1));
    assert!(table.columns.is_empty());
    assert!(table.comment.is_none());
    assert!(table.table_type.is_none());
}

#[test]
fn test_table_initialize_columns() {
    let mut table = Table::new(1u32, "Table1", "A1:D10");
    table.initialize_columns();
    assert_eq!(table.columns.len(), 4);
    assert_eq!(table.columns[0].name, "Column1");
    assert_eq!(table.columns[3].name, "Column4");
    assert!(table.auto_filter_range.is_some());
}

#[test]
fn test_table_column_names() {
    let mut table = Table::new(1u32, "Table1", "A1:C5");
    table.initialize_columns();
    let names = table.column_names();
    assert_eq!(names, vec!["Column1", "Column2", "Column3"]);
}

#[test]
fn test_parse_cell_ref() {
    assert_eq!(parse_cell_ref("A1"), Some((1, 1)));
    assert_eq!(parse_cell_ref("B2"), Some((2, 2)));
    assert_eq!(parse_cell_ref("Z10"), Some((26, 10)));
    assert_eq!(parse_cell_ref("AA1"), Some((27, 1)));
    assert_eq!(parse_cell_ref("AB100"), Some((28, 100)));
}

#[test]
fn test_parse_range() {
    assert_eq!(parse_range("A1:D10"), Some((1, 1, 4, 10)));
    assert_eq!(parse_range("B2:C5"), Some((2, 2, 3, 5)));
    assert_eq!(parse_range("A1"), None); // Missing colon
    assert_eq!(parse_range(""), None);
}

#[test]
fn parses_prefixed_strict_table_definition() {
    let xml = r#"<s:table xmlns:s="http://purl.oclc.org/ooxml/spreadsheetml/main"
            xmlns:f="urn:foreign" id="3" name="Sales_Internal" displayName="Sales &amp; Margin"
            ref="$A$1:$B$3" tableType="worksheet" headerRowCount="1"
            totalsRowCount="1" totalsRowShown="true" published="0">
            <f:tableColumns count="1"><s:tableColumn id="99" name="Ignored"/></f:tableColumns>
            <s:autoFilter ref="$A$1:$B$3"><s:sortState ref="$A$2:$B$3" caseSensitive="0" sortMethod="pinYin">
                <s:sortCondition ref="$B$2:$B$3" descending="1" sortBy="value" customList="High,Low"/>
            </s:sortState></s:autoFilter>
            <s:tableColumns count="2">
                <s:tableColumn id="1" name="Region &amp; Area" uniqueName="Region" totalsRowLabel="Total"/>
                <s:tableColumn id="2" name="Sales" totalsRowFunction="sum">
                    <s:calculatedColumnFormula array="false">SUM([Sales])&amp;1</s:calculatedColumnFormula>
                    <s:totalsRowFormula>SUBTOTAL(109,[Sales])</s:totalsRowFormula>
                </s:tableColumn>
            </s:tableColumns>
            <s:tableStyleInfo name="TableStyleMedium2" showFirstColumn="0"
                showLastColumn="1" showRowStripes="true" showColumnStripes="false"/>
        </s:table>"#;
    let table = parse_table_xml(xml).unwrap().unwrap();

    assert_eq!(table.id, 3);
    assert_eq!(table.display_name, "Sales & Margin");
    assert_eq!(table.columns.len(), 2);
    assert_eq!(table.columns[0].name, "Region & Area");
    assert_eq!(
        table.columns[1]
            .calculated_column_formula
            .as_ref()
            .unwrap()
            .text,
        "SUM([Sales])&1"
    );
    assert_eq!(table.sort_state.as_ref().unwrap().conditions.len(), 1);
    assert_eq!(
        table.style_info.as_ref().unwrap().show_last_column,
        Some(true)
    );
    assert_eq!(table.published, Some(false));
}

#[test]
fn rejects_malformed_table_definitions() {
    const S: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    let root = |body: &str, attributes: &str| {
        format!(
            r#"<table xmlns="{S}" id="1" displayName="Table1" ref="A1:B3" {attributes}>{body}</table>"#
        )
    };
    for xml in [
        root("", ""),
        root(
            r#"<tableColumns count="2"><tableColumn id="1" name="One"/></tableColumns>"#,
            "",
        ),
        root(
            r#"<tableColumns count="2"><tableColumn id="1" name="One"/><tableColumn id="1" name="Two"/></tableColumns>"#,
            "",
        ),
        root(
            r#"<tableColumns count="2"><tableColumn id="1" name="Same"/><tableColumn id="2" name="same"/></tableColumns>"#,
            "",
        ),
        root(
            r#"<autoFilter ref="A1:C3"/><tableColumns count="2"><tableColumn id="1" name="One"/><tableColumn id="2" name="Two"/></tableColumns>"#,
            "",
        ),
        root(
            r#"<tableColumns count="2"><tableColumn id="1" name="One"/><tableColumn id="2" name="Two" totalsRowFunction="median"/></tableColumns>"#,
            "",
        ),
        root(
            r#"<tableColumns count="2"><tableColumn id="1" name="One"/><tableColumn id="2" name="Two"/></tableColumns><tableStyleInfo showRowStripes="yes"/>"#,
            "",
        ),
        root(
            r#"<tableColumns count="2"><tableColumn id="1" name="One"/><tableColumn id="2" name="Two"/></tableColumns>"#,
            r#"tableType="bad""#,
        ),
    ] {
        assert!(parse_table_xml(&xml).is_err(), "accepted {xml}");
    }
    assert!(
        parse_table_xml(r#"<table xmlns="urn:foreign" id="1" displayName="T" ref="A1"/>"#)
            .unwrap()
            .is_none()
    );
}

fn writable_table() -> Table {
    let mut table = Table::new(1, "TestTable", "A1:B5");
    table.display_name = "Test Table".into();
    table.comment = Some("Test comment".into());
    table.header_row_count = Some(1);
    table.totals_row_count = Some(1);
    table.totals_row_shown = Some(true);
    table.columns = vec![
        TableColumn::new(1, "Column A"),
        TableColumn::new(2, "Column B"),
    ];
    table
}

#[test]
fn writes_table_attributes_columns_and_style() {
    let mut table = writable_table();
    let mut style = TableStyleInfo::new();
    style.name = Some("TableStyleMedium2".into());
    style.show_first_column = Some(true);
    style.show_last_column = Some(false);
    style.show_row_stripes = Some(true);
    style.show_column_stripes = Some(false);
    table.style_info = Some(style);

    let xml = serialize_table(&table).unwrap();
    assert!(xml.contains(r#"id="1""#));
    assert!(xml.contains(r#"name="TestTable""#));
    assert!(xml.contains(r#"displayName="Test Table""#));
    assert!(xml.contains(r#"ref="A1:B5""#));
    assert!(xml.contains(r#"comment="Test comment""#));
    assert!(xml.contains(r#"<tableColumns count="2">"#));
    assert!(xml.contains(r#"<tableColumn id="1" name="Column A"/>"#));
    assert!(xml.contains(r#"<tableColumn id="2" name="Column B"/>"#));
    assert!(xml.contains(r#"<tableStyleInfo name="TableStyleMedium2" showFirstColumn="1" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>"#));
}

#[test]
fn writes_auto_filter_sort_and_formulas_losslessly() {
    let mut table = writable_table();
    table.published = Some(true);
    table.auto_filter_range = Some("A1:B5".into());
    table.sort_state = Some(SortState {
        ref_range: "A2:B5".into(),
        column_sort: Some(false),
        case_sensitive: Some(true),
        sort_method: Some(SortMethod::PinYin),
        conditions: vec![SortCondition {
            ref_range: "B2:B5".into(),
            descending: Some(true),
            sort_by: Some(SortBy::Icon),
            custom_list: Some("High,Low".into()),
            dxf_id: None,
            icon_set: Some(IconSet::ThreeArrows),
            icon_id: Some(2),
        }],
    });
    table.columns[0].calculated_column_formula = Some(TableFormula {
        array: Some(false),
        text: "=[@[Column B]]*2".into(),
    });
    table.columns[1].totals_row_function = Some(TotalsRowFunction::Sum);
    table.columns[1].totals_row_formula = Some(TableFormula {
        array: None,
        text: "SUBTOTAL(109,[Column B])".into(),
    });

    let xml = serialize_table(&table).unwrap();
    let parsed = parse_table_xml(xml.as_bytes()).unwrap().unwrap();
    assert_eq!(parsed.published, Some(true));
    assert_eq!(parsed.auto_filter_range.as_deref(), Some("A1:B5"));
    assert_eq!(parsed.sort_state.as_ref().unwrap().conditions.len(), 1);
    assert_eq!(
        parsed.sort_state.unwrap().sort_method,
        Some(SortMethod::PinYin)
    );
    assert_eq!(
        parsed.columns[0]
            .calculated_column_formula
            .as_ref()
            .unwrap()
            .text,
        "=[@[Column B]]*2"
    );
    assert_eq!(
        parsed.columns[1].totals_row_function,
        Some(TotalsRowFunction::Sum)
    );
    assert_eq!(
        parsed.columns[1].totals_row_formula.as_ref().unwrap().text,
        "SUBTOTAL(109,[Column B])"
    );
}

#[test]
fn escapes_writable_table_values() {
    let mut table = writable_table();
    table.name = "Table<>&\"'".into();
    table.display_name = "Test <Table>".into();
    table.columns[0].name = "A & B".into();
    table.columns[0].calculated_column_formula = Some(TableFormula {
        array: None,
        text: "IF(A1<2,\"yes\",\"no\")".into(),
    });

    let xml = serialize_table(&table).unwrap();
    assert!(xml.contains("Table&lt;&gt;&amp;&quot;&apos;"));
    assert!(xml.contains("A &amp; B"));
    assert!(xml.contains("IF(A1&lt;2,&quot;yes&quot;,&quot;no&quot;)</calculatedColumnFormula>"));
    assert!(parse_table_xml(xml.as_bytes()).is_ok());
}

#[test]
fn rejects_invalid_writable_table_models() {
    let mut table = writable_table();
    table.columns[1].id = 1;
    assert!(serialize_table(&table).is_err());

    let mut table = writable_table();
    table.auto_filter_range = Some("A1:C5".into());
    assert!(serialize_table(&table).is_err());

    let mut table = writable_table();
    table.sort_state = Some(SortState::new("A2:B5"));
    assert!(serialize_table(&table).is_err());

    let mut table = writable_table();
    table.columns.clear();
    assert!(serialize_table(&table).is_err());
}

#[test]
fn native_direct_sort_state_allows_owner_bounded_full_column_keys() {
    let xml = r#"<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="T" displayName="T" ref="A1:B5"><autoFilter ref="A1:B5"/><sortState ref="A2:B4"><sortCondition ref="B1:B5" descending="1"/></sortState><tableColumns count="2"><tableColumn id="1" name="A"/><tableColumn id="2" name="B"/></tableColumns></table>"#;
    let table = parse_table_xml(xml).unwrap().unwrap();
    let sort = table.sort_state.as_ref().unwrap();
    assert_eq!(sort.ref_range, "A2:B4");
    assert_eq!(sort.conditions[0].ref_range, "B1:B5");
    let emitted = serialize_table(&table).unwrap();
    assert_eq!(
        parse_table_xml(&emitted).unwrap().unwrap().sort_state,
        table.sort_state
    );
    for invalid in [
        xml.replace("ref=\"B1:B5\"", "ref=\"B5:B1\""),
        xml.replace("ref=\"B1:B5\"", "ref=\"XFE1:XFE5\""),
        xml.replace(
            "<autoFilter ref=\"A1:B5\"/>",
            "<autoFilter ref=\"A1:B5\"><sortState ref=\"A2:B4\"/></autoFilter>",
        ),
    ] {
        assert!(parse_table_xml(invalid).is_err());
    }
    let outside_owner = xml.replace("ref=\"B1:B5\"", "ref=\"C1:C5\"");
    let retained = parse_table_xml(outside_owner).unwrap().unwrap();
    assert_eq!(
        retained.sort_state.as_ref().unwrap().conditions[0].ref_range,
        "C1:C5"
    );
    assert!(serialize_table(&retained).is_err());
}

fn rect(reference: &str) -> litchi_sheet::Rect {
    litchi_sheet::Rect::from_a1(reference).expect("test range")
}

#[test]
fn table_bands_apply_schema_default_row_counts() {
    let mut table = Table::new(1, "Sales", "B2:D6");
    table.header_row_count = None;
    assert_eq!(table.range().unwrap(), rect("B2:D6"));
    assert_eq!(table.header_range().unwrap(), Some(rect("B2:D2")));
    assert_eq!(table.data_range().unwrap(), Some(rect("B3:D6")));
    assert_eq!(table.totals_range().unwrap(), None);

    table.header_row_count = Some(0);
    table.totals_row_count = Some(2);
    assert_eq!(table.header_range().unwrap(), None);
    assert_eq!(table.data_range().unwrap(), Some(rect("B2:D4")));
    assert_eq!(table.totals_range().unwrap(), Some(rect("B5:D6")));
}

#[test]
fn table_bands_accept_absolute_references_and_empty_data() {
    let mut table = Table::new(1, "Header", "$A$1:$C$1");
    assert_eq!(table.range().unwrap(), rect("A1:C1"));
    assert_eq!(table.header_range().unwrap(), Some(rect("A1:C1")));
    assert_eq!(table.data_range().unwrap(), None);
    assert_eq!(table.totals_range().unwrap(), None);

    table.ref_range = "A1:C2".into();
    table.totals_row_count = Some(1);
    assert_eq!(table.data_range().unwrap(), None);
    assert_eq!(table.totals_range().unwrap(), Some(rect("A2:C2")));
}

#[test]
fn table_bands_refuse_invalid_ranges_and_row_counts() {
    let mut table = Table::new(1, "Broken", "A1:C2");
    table.totals_row_count = Some(2);
    for band in [
        table.header_range(),
        table.data_range(),
        table.totals_range(),
    ] {
        assert!(band.is_err(), "three rows of edges in a two-row range");
    }
    table.totals_row_count = Some(u32::MAX);
    assert!(table.data_range().is_err());

    for reference in ["", "A0:B2", "B2:A1", "A1:B2:C3", "Sheet1!A1:B2"] {
        let table = Table::new(1, "Broken", reference);
        assert!(table.range().is_err(), "accepted {reference:?}");
        assert!(table.header_range().is_err(), "accepted {reference:?}");
    }
}

#[test]
fn table_columns_resolve_without_regard_to_case() {
    let mut table = Table::new(1, "Names", "A1:B3");
    table.columns = vec![TableColumn::new(1, "Straße"), TableColumn::new(2, "Amount")];
    assert_eq!(
        table.column("AMOUNT").unwrap().map(|column| column.id),
        Some(2)
    );
    assert_eq!(
        table.column("STRASSE").unwrap().map(|column| column.id),
        Some(1)
    );
    assert_eq!(
        table.column("strasse").unwrap().map(|column| column.id),
        Some(1)
    );
    assert!(table.column("Amounts").unwrap().is_none());
}

#[test]
fn table_column_lookup_refuses_unicode_ambiguity() {
    let mut table = Table::new(1, "Names", "A1:B3");
    table.columns = vec![
        TableColumn::new(1, "Straße"),
        TableColumn::new(2, "STRASSE"),
    ];
    assert!(matches!(
        table.column("strasse"),
        Err(crate::Error::Ambiguous {
            kind: "table column",
            matches: 2
        })
    ));
}

#[test]
fn unicode_duplicate_column_names_are_rejected_at_admission() {
    let mut table = Table::new(1, "Names", "A1:B3");
    table.columns = vec![
        TableColumn::new(1, "Straße"),
        TableColumn::new(2, "STRASSE"),
    ];
    assert!(serialize_table(&table).is_err());
    let xml = r#"<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Names" displayName="Names" ref="A1:B3"><tableColumns count="2"><tableColumn id="1" name="Straße"/><tableColumn id="2" name="STRASSE"/></tableColumns></table>"#;
    assert!(parse_table_xml(xml).is_err());
}
