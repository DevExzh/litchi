use super::model::MAX_XML_DEPTH;
use super::{parse, parse_defaults, x14ac};
use crate::cell::{Cell, Text, Value};
use crate::column;
use crate::formula::Cache;
use crate::layout;
use crate::row;
use litchi_sheet::{Cell as Address, Column as ColumnIndex, Rect, Row as RowIndex};

const S: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT_S: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";

#[test]
fn plain_worksheets_skip_only_the_unneeded_extension_capture() {
    let plain = format!(
        r#"<worksheet xmlns="{S}"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>"#
    );
    assert!(!x14ac::may_contain_descent(plain.as_bytes()));
    assert!(parse(plain.as_bytes(), || Ok(None)).is_ok());

    let conservative_false_positive =
        format!(r#"<worksheet xmlns="{S}"><!-- dyDescent --><sheetData/></worksheet>"#);
    assert!(x14ac::may_contain_descent(
        conservative_false_positive.as_bytes()
    ));
    assert!(parse(conservative_false_positive.as_bytes(), || Ok(None)).is_ok());
}

#[test]
fn rejected_plain_worksheets_keep_extension_error_precedence() {
    let malformed =
        format!(r#"<worksheet xmlns="{S}"><sheetData><row r="1"></sheetData></worksheet>"#);
    assert!(!x14ac::may_contain_descent(malformed.as_bytes()));
    let error = parse(malformed.as_bytes(), || Ok(None)).expect_err("malformed worksheet");
    assert!(
        error
            .to_string()
            .contains("invalid worksheet extension XML"),
        "unexpected error precedence: {error}"
    );
}

#[test]
fn malformed_alternate_content_keeps_x14ac_error_precedence() {
    let malformed = format!(
        r#"<x:worksheet xmlns:x="{S}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac"><mc:AlternateContent><mc:Choice Requires="future"><x:sheetFormatPr x14ac:dyDescent="0.2"/></mc:AlternateContent></x:worksheet>"#
    );
    let error = parse(malformed.as_bytes(), || Ok(None)).expect_err("malformed AlternateContent");
    assert!(
        error
            .to_string()
            .contains("invalid worksheet extension XML"),
        "unexpected error precedence: {error}"
    );
}

#[test]
fn parses_exact_sparse_values_formulas_and_explicit_empty_cells() {
    let xml = format!(
        r#"<worksheet xmlns="{S}"><sheetData>
                <row r="1"><c r="A1" t="s"><v>0</v></c><c r="C1" s="3"/></row>
                <row r="3"><c r="A3"><v>-0.000</v></c><c r="B3" t="b"><v>1</v></c>
                <c r="C3" t="inlineStr"><is><r><t>Hello </t></r><r><t>world</t></r></is></c></row>
                <row r="4"><c r="A4" t="str"><f>CONCAT(A1,"!")</f><v>cached</v></c></row>
            </sheetData></worksheet>"#
    );
    let strings = [Text::from("shared")];
    let store = parse(xml.as_bytes(), || Ok(Some(&strings))).expect("valid worksheet");
    assert!(matches!(
        store.get(Address::at(0, 0).expect("address")),
        Some(Cell::Value(Value::Text(value))) if value.as_str() == "shared"
    ));
    assert!(store.get(Address::at(0, 1).expect("address")).is_none());
    assert!(matches!(
        store.get(Address::at(0, 2).expect("address")),
        Some(Cell::Empty)
    ));
    assert!(matches!(
        store.get(Address::at(2, 0).expect("address")),
        Some(Cell::Value(Value::Number(number))) if number.as_str() == "-0.000"
    ));
    assert!(matches!(
        store.get(Address::at(2, 2).expect("address")),
        Some(Cell::Value(Value::Text(value))) if value.as_str() == "Hello world"
    ));
    let Some(Cell::Formula(formula)) = store.get(Address::at(3, 0).expect("address")) else {
        panic!("expected formula")
    };
    assert_eq!(formula.text(), "CONCAT(A1,\"!\")");
    assert!(matches!(
        formula.cached().map(Cache::value),
        Some(Value::Text(value)) if value.as_str() == "cached"
    ));
}

#[test]
fn namespace_equivalent_worksheet_events_have_identical_results() {
    let plain = format!(
        r#"<worksheet xmlns="{S}"><sheetData><row r="1"><c r="A1"><v>7</v></c></row></sheetData></worksheet>"#
    );
    let prefixed = format!(
        r#"<x:worksheet xmlns:x="{S}" xmlns:y="{S}" xmlns="urn:not-spreadsheet"><y:sheetData><y:row r="1"><y:c r="A1"><y:v>7</y:v></y:c></y:row></y:sheetData></x:worksheet>"#
    );
    let strict = format!(
        r#"<worksheet xmlns="{STRICT_S}"><sheetData><row r="1"><c r="A1"><v>7</v></c></row></sheetData></worksheet>"#
    );

    let plain = parse(plain.as_bytes(), || Ok(None)).expect("plain worksheet");
    let prefixed = parse(prefixed.as_bytes(), || Ok(None)).expect("prefixed worksheet");
    let strict = parse(strict.as_bytes(), || Ok(None)).expect("strict worksheet");
    let address = Address::from_a1("A1").expect("address");
    assert_eq!(plain.entries().len(), prefixed.entries().len());
    assert_eq!(plain.get(address), prefixed.get(address));
    assert_eq!(plain.get(address), strict.get(address));
}

#[test]
fn namespace_rebinding_to_another_uri_is_not_treated_as_core_content() {
    let xml = format!(
        r#"<x:worksheet xmlns:x="{S}" xmlns:y="urn:not-spreadsheet"><x:sheetData><x:row r="1"><y:c r="A1"><y:v>7</y:v></y:c></x:row></x:sheetData></x:worksheet>"#
    );
    let store = parse(xml.as_bytes(), || Ok(None)).expect("unknown namespace should be ignored");
    assert!(store.entries().is_empty());
}

#[test]
fn parses_sparse_merges_and_rejects_ambiguous_merge_markup() {
    let xml = format!(
        r#"<worksheet xmlns="{S}"><sheetData><row r="1"><c r="A1"><v>1</v></c></row><row r="2"><c r="B2"><v>2</v></c></row></sheetData><mergeCells count="2"><mergeCell ref="A1:C3"/><mergeCell ref="E5:F5"/></mergeCells></worksheet>"#
    );
    let store = parse(xml.as_bytes(), || Ok(None)).expect("valid merged ranges");
    assert_eq!(store.entries().len(), 2, "merges must stay sparse");
    assert_eq!(
        store.merges().map(Rect::a1).collect::<Vec<_>>(),
        ["A1:C3", "E5:F5"]
    );
    assert!(matches!(
        store.view(Address::from_a1("A1").expect("anchor")),
        crate::cell::View::Stored(Cell::Value(_))
    ));
    assert!(matches!(
        store.view(Address::from_a1("B2").expect("covered")),
        crate::cell::View::Covered(range) if range == Rect::from_a1("A1:C3").expect("range")
    ));
    assert!(matches!(
        store.view(Address::from_a1("D4").expect("missing")),
        crate::cell::View::Missing
    ));

    for malformed in [
        format!(
            r#"<worksheet xmlns="{S}"><sheetData/><mergeCells count="2"><mergeCell ref="A1:B2"/></mergeCells></worksheet>"#
        ),
        format!(
            r#"<worksheet xmlns="{S}"><sheetData/><mergeCells><mergeCell ref="A1:C3"/><mergeCell ref="C3:D4"/></mergeCells></worksheet>"#
        ),
        format!(
            r#"<worksheet xmlns="{S}"><sheetData/><mergeCells><mergeCell ref="A1"/></mergeCells></worksheet>"#
        ),
        format!(
            r#"<worksheet xmlns="{S}"><sheetData/><mergeCells><future/></mergeCells></worksheet>"#
        ),
        format!(r#"<worksheet xmlns="{S}"><sheetData/><mergeCell ref="A1:B2"/></worksheet>"#),
        format!(
            r#"<worksheet xmlns="{S}"><sheetData/><hyperlinks/><mergeCells><mergeCell ref="A1:B2"/></mergeCells></worksheet>"#
        ),
    ] {
        assert!(parse(malformed.as_bytes(), || Ok(None)).is_err());
    }
}

#[test]
fn expands_shared_formulas_and_preserves_cached_values() {
    let xml = format!(
        r#"<worksheet xmlns="{S}"><sheetData>
                <row r="1"><c r="A1"><f t="shared" ref="A1:A2" si="7">B1+$C$1</f><v>1</v></c></row>
                <row r="2"><c r="A2"><f t="shared" si="7"/><v>2</v></c></row>
            </sheetData></worksheet>"#
    );
    let store = parse(xml.as_bytes(), || Ok(None)).expect("valid shared formula");
    let Some(Cell::Formula(formula)) = store.get(Address::at(1, 0).expect("address")) else {
        panic!("expected formula")
    };
    assert_eq!(formula.text(), "B2+$C$1");
    assert!(matches!(
        formula.cached().map(Cache::value),
        Some(Value::Number(number)) if number.as_str() == "2"
    ));
}

#[test]
fn keeps_declared_stored_content_and_style_extents_distinct() {
    let xml = format!(
        r#"<worksheet xmlns="{S}"><dimension ref="$B$2:F9"/><sheetData><row r="2"><c r="B2"><v>1</v></c></row><row r="4" ht="30.5" s="1" customFormat="1" customHeight="true" hidden="true" outlineLevel="2" collapsed="1" thickTop="1" thickBot="true" ph="1"><c r="D4"/></row><row r="9" hidden="0"><c r="F9" s="1"/></row></sheetData></worksheet>"#
    );
    let store = parse(xml.as_bytes(), || Ok(None)).expect("valid extents");
    let extents = store.extents();
    assert_eq!(extents.declared().map(Rect::a1).as_deref(), Some("B2:F9"));
    assert_eq!(extents.stored().map(Rect::a1).as_deref(), Some("B2:F9"));
    assert_eq!(extents.content().map(Rect::a1).as_deref(), Some("B2"));
    assert_eq!(extents.styled().map(Rect::a1).as_deref(), Some("F9"));
    assert_eq!(extents.used().map(Rect::a1).as_deref(), Some("B2:F9"));
    let row = store.row(RowIndex::new(3).expect("row 4"));
    assert!(row.hidden());
    assert_eq!(row.height().map(row::Height::get), Some(30.5));
    assert!(row.custom_height());
    assert_eq!(row.outline().get(), 2);
    assert!(row.collapsed());
    assert!(row.thick_top());
    assert!(row.thick_bottom());
    assert!(row.phonetic());
    assert!(row.custom_format());
    assert_eq!(
        store.row_entry(row.index()).unwrap().properties.style,
        Some(1)
    );
    assert!(!store.row(RowIndex::new(8).expect("row 9")).hidden());
    let implicit = store.row(RowIndex::new(5).expect("row 6"));
    assert!(!implicit.stored());
    assert!(!implicit.hidden());
    assert_eq!(store.rows().count(), 3);
}

#[test]
fn parses_checked_grid_defaults_and_effective_x14ac_descent() {
    let xml = format!(
        r#"<x:worksheet xmlns:x="{S}" xmlns:future="urn:future"
                xmlns:compat="http://schemas.openxmlformats.org/markup-compatibility/2006"
                xmlns:ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
                compat:Ignorable="ac future">
                <x:sheetFormatPr baseColWidth="10" defaultColWidth="12.5"
                    defaultRowHeight="16" customHeight="false" zeroHeight="true"
                    thickTop="1" thickBottom="true" outlineLevelRow="3"
                    outlineLevelCol="2" ac:dyDescent="0.2" future:keep="yes"/>
                <x:sheetData><x:row r="2" customHeight="0" ac:dyDescent="0.3"/></x:sheetData>
            </x:worksheet>"#
    );
    let store = parse(xml.as_bytes(), || Ok(None)).expect("valid worksheet defaults");
    let defaults = store.defaults().expect("stored defaults");
    assert_eq!(defaults.base_width(), 10);
    assert_eq!(defaults.stored_base_width(), Some(10));
    assert_eq!(defaults.width().map(layout::Width::get), Some(12.5));
    assert_eq!(defaults.height().get(), 16.0);
    assert!(defaults.custom_height());
    assert!(defaults.hidden());
    assert!(defaults.thick_top());
    assert!(defaults.thick_bottom());
    assert_eq!(defaults.row_outline().get(), 3);
    assert_eq!(defaults.column_outline().get(), 2);
    assert_eq!(defaults.descent().map(layout::Descent::get), Some(0.2));

    let row = store.row(RowIndex::new(1).expect("row 2"));
    assert_eq!(row.descent().map(layout::Descent::get), Some(0.3));
    assert!(row.custom_height());
}

#[test]
fn focused_defaults_path_does_not_materialize_unrelated_cells() {
    let xml = format!(
        r#"<worksheet xmlns="{S}"><sheetFormatPr defaultRowHeight="16"/><sheetData>
                <row r="1"><c r="A1"><v>not-a-number</v></c></row>
            </sheetData></worksheet>"#
    );
    let defaults = parse_defaults(xml.as_bytes())
        .expect("focused defaults parser should ignore cell payloads")
        .expect("sheetFormatPr");
    assert_eq!(defaults.height().get(), 16.0);
}

#[test]
fn focused_defaults_follow_the_active_mce_branch() {
    let xml = format!(
        r#"<worksheet xmlns="{S}"
                xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
                xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
                xmlns:future="urn:future" mc:Ignorable="x14ac future">
                <mc:AlternateContent>
                    <mc:Choice Requires="future">
                        <sheetFormatPr xmlns:x14ac="urn:not-x14ac" defaultRowHeight="20" x14ac:dyDescent="0.5"/>
                    </mc:Choice>
                    <mc:Fallback>
                        <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
                    </mc:Fallback>
                </mc:AlternateContent>
            </worksheet>"#
    );
    let defaults = parse_defaults(xml.as_bytes())
        .expect("active MCE branch should parse")
        .expect("active fallback defaults");
    assert_eq!(defaults.height().get(), 15.0);
    assert_eq!(defaults.descent().map(layout::Descent::get), Some(0.25));
}

#[test]
fn focused_defaults_ignore_malformed_row_extensions() {
    let xml = format!(
        r#"<worksheet xmlns="{S}"
                xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
                <sheetFormatPr defaultRowHeight="16" x14ac:dyDescent="0.2"/>
                <sheetData><row r="1" x14ac:dyDescent="-1"/><row r="1" x14ac:dyDescent="-1"/></sheetData>
            </worksheet>"#
    );
    let defaults = parse_defaults(xml.as_bytes())
        .expect("row extensions are outside the focused view")
        .expect("sheetFormatPr");
    assert_eq!(defaults.descent().map(layout::Descent::get), Some(0.2));
}

#[test]
fn focused_defaults_enforce_root_shape_and_xml_depth() {
    assert!(parse_defaults(format!(r#"<worksheet xmlns="{S}"/>"#).as_bytes()).is_err());
    assert!(
        parse_defaults(
            format!(r#"<worksheet xmlns="{S}"><sheetFormatPr defaultRowHeight="15"></worksheet>"#)
                .as_bytes()
        )
        .is_err()
    );

    let mut xml = format!(r#"<worksheet xmlns="{S}">"#);
    for _ in 0..=MAX_XML_DEPTH {
        xml.push_str("<future>");
    }
    for _ in 0..=MAX_XML_DEPTH {
        xml.push_str("</future>");
    }
    xml.push_str("</worksheet>");
    assert!(parse_defaults(xml.as_bytes()).is_err());
}

#[test]
fn rejects_malformed_grid_defaults_and_descent() {
    for body in [
        r#"<sheetFormatPr/>"#,
        r#"<sheetFormatPr defaultRowHeight="-1"/>"#,
        r#"<sheetFormatPr defaultRowHeight="NaN"/>"#,
        r#"<sheetFormatPr defaultRowHeight="15" baseColWidth="256"/>"#,
        r#"<sheetFormatPr defaultRowHeight="15" defaultColWidth="65536"/>"#,
        r#"<sheetFormatPr defaultRowHeight="15" outlineLevelRow="8"/>"#,
        r#"<sheetFormatPr defaultRowHeight="15" outlineLevelCol="8"/>"#,
        r#"<sheetFormatPr defaultRowHeight="15">text</sheetFormatPr>"#,
        r#"<sheetFormatPr defaultRowHeight="15"><future/></sheetFormatPr>"#,
        r#"<sheetFormatPr defaultRowHeight="15"/><sheetFormatPr defaultRowHeight="16"/>"#,
        r#"<sheetData/><sheetFormatPr defaultRowHeight="15"/>"#,
    ] {
        let xml = format!(r#"<worksheet xmlns="{S}">{body}<sheetData/></worksheet>"#);
        assert!(
            parse(xml.as_bytes(), || Ok(None)).is_err(),
            "accepted {body}"
        );
    }

    for value in ["-0.1", "NaN", "inf"] {
        let xml = format!(
            r#"<worksheet xmlns="{S}"
                    xmlns:a="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
                    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
                    mc:Ignorable="a"><sheetFormatPr defaultRowHeight="15" a:dyDescent="{value}"/><sheetData/></worksheet>"#
        );
        assert!(
            parse(xml.as_bytes(), || Ok(None)).is_err(),
            "accepted dyDescent={value}"
        );
    }
}

#[test]
fn rejects_extension_xml_beyond_the_depth_limit() {
    let mut xml = format!(r#"<worksheet xmlns="{S}">"#);
    for _ in 0..256 {
        xml.push_str("<future>");
    }
    for _ in 0..256 {
        xml.push_str("</future>");
    }
    xml.push_str("</worksheet>");

    assert!(parse(xml.as_bytes(), || Ok(None)).is_err());
}

#[test]
fn resolves_complete_last_matching_column_records() {
    let xml = format!(
        r#"<worksheet xmlns="{S}"><cols><col min="2" max="4" width="20" style="1" hidden="1" bestFit="true" customWidth="1" phonetic="1" outlineLevel="2" collapsed="1"/><col min="3" max="3" width="10"/></cols><sheetData/></worksheet>"#
    );
    let store = parse(xml.as_bytes(), || Ok(None)).expect("valid columns");

    let a = store.column(ColumnIndex::new(0).expect("A"));
    assert!(!a.stored());
    assert!(!a.hidden());
    let b = store.column(ColumnIndex::new(1).expect("B"));
    assert!(b.stored());
    assert!(b.hidden());
    assert_eq!(b.width().map(column::Width::get), Some(20.0));
    assert!(b.best_fit());
    assert!(b.custom_width());
    assert!(b.phonetic());
    assert_eq!(b.outline().get(), 2);
    assert!(b.collapsed());
    assert_eq!(
        store.column_entry(b.index()).unwrap().properties.style,
        Some(1)
    );

    let c = store.column(ColumnIndex::new(2).expect("C"));
    assert!(c.stored());
    assert!(!c.hidden());
    assert_eq!(c.width().map(column::Width::get), Some(10.0));
    assert!(!c.best_fit());
    assert!(!c.custom_width());
    assert!(!c.phonetic());
    assert_eq!(c.outline(), column::Outline::NONE);
    assert!(!c.collapsed());
    assert_eq!(
        store.column_entry(c.index()).unwrap().properties.style,
        None
    );

    let d = store.column(ColumnIndex::new(3).expect("D"));
    assert!(d.hidden());
    assert_eq!(
        store
            .columns()
            .map(|column| column.index())
            .collect::<Vec<_>>(),
        [
            ColumnIndex::new(1).expect("B"),
            ColumnIndex::new(2).expect("C"),
            ColumnIndex::new(3).expect("D"),
        ]
    );
}

#[test]
fn rejects_malformed_column_property_records() {
    for body in [
        "<cols/>",
        "<cols></cols>",
        r#"<cols><col max="1"/></cols>"#,
        r#"<cols><col min="1"/></cols>"#,
        r#"<cols><col min="0" max="1"/></cols>"#,
        r#"<cols><col min="2" max="1"/></cols>"#,
        r#"<cols><col min="1" max="16385"/></cols>"#,
        r#"<cols><col min="1" max="1" width="256"/></cols>"#,
        r#"<cols><col min="1" max="1" width="NaN"/></cols>"#,
        r#"<cols><col min="1" max="1" style="65430"/></cols>"#,
        r#"<cols><col min="1" max="1" outlineLevel="8"/></cols>"#,
        r#"<cols><col min="1" max="1" hidden="yes"/></cols>"#,
        r#"<cols><col min="1" max="1"/></cols><cols><col min="2" max="2"/></cols>"#,
        r#"<sheetData/><cols><col min="1" max="1"/></cols>"#,
    ] {
        let xml = format!(r#"<worksheet xmlns="{S}">{body}<sheetData/></worksheet>"#);
        assert!(
            parse(xml.as_bytes(), || Ok(None)).is_err(),
            "accepted {body}"
        );
    }
}

#[test]
fn rejects_malformed_dimensions_and_row_properties() {
    for body in [
        "<dimension/><sheetData/>",
        r#"<dimension ref="A0"/><sheetData/>"#,
        r#"<dimension ref="A1"/><dimension ref="B2"/><sheetData/>"#,
        r#"<sheetData/><dimension ref="A1"/>"#,
        r#"<sheetData><row r="1" hidden="yes"/></sheetData>"#,
        r#"<sheetData><row r="1" ht="NaN"/></sheetData>"#,
        r#"<sheetData><row r="1" ht="409.1"/></sheetData>"#,
        r#"<sheetData><row r="1" s="65491"/></sheetData>"#,
        r#"<sheetData><row r="1" outlineLevel="8"/></sheetData>"#,
        r#"<sheetData><row r="1" thickTop="yes"/></sheetData>"#,
    ] {
        let xml = format!(r#"<worksheet xmlns="{S}">{body}</worksheet>"#);
        assert!(
            parse(xml.as_bytes(), || Ok(None)).is_err(),
            "accepted {body}"
        );
    }
}

#[test]
fn rejects_grid_escape_duplicates_and_broken_shared_groups() {
    for body in [
        r#"<row r="1048577"/>"#,
        r#"<row r="1"><c r="XFE1"/></row>"#,
        r#"<row r="1"><c r="A1"/><c r="A1"/></row>"#,
        r#"<row r="2"/><row r="1"/>"#,
        r#"<row r="1"><c r="A1" s="65491"/></row>"#,
        r#"<row r="1"><c r="A1" cm="0"/></row>"#,
        r#"<row r="1"><c r="A1" vm="2147483648"/></row>"#,
        r#"<row r="1"><c r="A1"><f bx="1">1</f></c></row>"#,
        r#"<row r="1"><c r="A1"><f t="shared" si="0"/></c></row>"#,
    ] {
        let xml = format!(r#"<worksheet xmlns="{S}"><sheetData>{body}</sheetData></worksheet>"#);
        assert!(
            parse(xml.as_bytes(), || Ok(None)).is_err(),
            "accepted {body}"
        );
    }
}

#[test]
fn rejects_missing_shared_strings_bad_indexes_and_formula_markers() {
    let shared = format!(
        r#"<worksheet xmlns="{S}"><sheetData><row r="1"><c r="A1" t="s"><v>1</v></c></row></sheetData></worksheet>"#
    );
    assert!(parse(shared.as_bytes(), || Ok(None)).is_err());
    let strings = [Text::from("only index zero")];
    assert!(parse(shared.as_bytes(), || Ok(Some(&strings))).is_err());

    let marked = format!(
        r#"<worksheet xmlns="{S}"><sheetData><row r="1"><c r="A1"><f>=1+1</f><v>2</v></c></row></sheetData></worksheet>"#
    );
    assert!(parse(marked.as_bytes(), || Ok(None)).is_err());
}

#[test]
fn validates_typed_date_cells_without_normalizing_the_lexeme() {
    let valid = br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="d"><v>2026-07-31T12:34:56.250-07:00</v></c></row></sheetData></worksheet>"#;
    let store = parse(valid, || Ok(None)).expect("valid date cell");
    assert!(matches!(
        store.get(Address::from_a1("A1").expect("address")),
        Some(Cell::Value(Value::Date(date)))
            if date.as_str() == "2026-07-31T12:34:56.250-07:00"
    ));

    let invalid = br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="d"><v>2026-02-29</v></c></row></sheetData></worksheet>"#;
    assert!(parse(invalid, || Ok(None)).is_err());
}

#[test]
fn numeric_sheets_do_not_load_an_unneeded_shared_string_table() {
    let xml = format!(
        r#"<worksheet xmlns="{S}"><sheetData><row r="1"><c r="A1"><v>7</v></c></row></sheetData></worksheet>"#
    );
    let called = std::cell::Cell::new(false);
    let store = parse(xml.as_bytes(), || {
        called.set(true);
        Ok(None)
    })
    .expect("numeric worksheet");
    assert!(!called.get());
    assert!(matches!(
        store.get(Address::at(0, 0).expect("address")),
        Some(Cell::Value(Value::Number(number))) if number.as_str() == "7"
    ));
}
