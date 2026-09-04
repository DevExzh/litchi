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

mod streaming_0361_tests {
    use std::io::{self, BufRead, Cursor, Read};

    use litchi_ooxml_common::mce::{Capabilities, Error as MceError, StreamError, StreamLimits};
    use litchi_sheet::ROWS;

    use super::super::{x14ac, x14ac::Values};

    const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    const S: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    const X14AC: &str = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac";

    fn worksheet(body: &str) -> String {
        format!(
            r#"<worksheet xmlns="{S}" xmlns:mc="{MC}" xmlns:x14ac="{X14AC}" mc:Ignorable="x14ac">{body}</worksheet>"#
        )
    }

    fn choices_capabilities() -> Capabilities {
        let mut capabilities = Capabilities::default();
        capabilities.understand_namespace(X14AC);
        capabilities
    }

    fn signature(values: &Values) -> (Option<f64>, Vec<(u32, f64)>) {
        (
            values.defaults.map(|value| value.get()),
            values
                .rows
                .iter()
                .map(|(&row, &value)| (row, value.get()))
                .collect(),
        )
    }

    fn stream_reader<R: BufRead>(
        reader: &mut R,
        capabilities: &Capabilities,
        limits: &StreamLimits,
        capture_rows: bool,
    ) -> x14ac::StreamResult<Values> {
        x14ac::capture_stream(reader, capabilities, limits, capture_rows)
    }

    fn stream(
        xml: &[u8],
        capabilities: &Capabilities,
        limits: &StreamLimits,
        capture_rows: bool,
    ) -> x14ac::StreamResult<Values> {
        let mut reader = Cursor::new(xml);
        stream_reader(&mut reader, capabilities, limits, capture_rows)
    }

    fn default_stream(xml: &[u8], capture_rows: bool) -> x14ac::StreamResult<Values> {
        stream(
            xml,
            &Capabilities::default(),
            &StreamLimits::default(),
            capture_rows,
        )
    }

    #[test]
    fn streaming_0361_captures_start_and_empty_structural_events() {
        let xml = worksheet(
            r#"<sheetFormatPr x14ac:dyDescent="0.2"></sheetFormatPr><sheetData><row r="2" x14ac:dyDescent="0.3"></row><row r="3" x14ac:dyDescent="0.4"/></sheetData>"#,
        );
        let values = default_stream(xml.as_bytes(), true).expect("valid extension stream");

        assert_eq!(signature(&values), (Some(0.2), vec![(2, 0.3), (3, 0.4)]));
    }

    #[test]
    fn streaming_0361_focused_defaults_skip_malformed_rows_but_keep_raw_guards() {
        let malformed_rows = worksheet(
            r#"<sheetFormatPr x14ac:dyDescent="0.2"/><sheetData><row r="0" x14ac:dyDescent="-1"/><row r="not-a-row" x14ac:dyDescent="NaN"/></sheetData>"#,
        );
        let values = x14ac::capture_stream_defaults(
            &mut Cursor::new(malformed_rows.as_bytes()),
            &Capabilities::default(),
            &StreamLimits::default(),
        )
        .expect("focused defaults must not parse row values")
        .expect("default descent");
        assert_eq!(values.get(), 0.2);

        let reserved = worksheet(
            r#"<sheetFormatPr x14ac:dyDescent="0.2"/><sheetData><row r="0" litchi_x14ac_dyDescent="0.3"/></sheetData>"#,
        );
        assert!(
            x14ac::capture_stream_defaults(
                &mut Cursor::new(reserved.as_bytes()),
                &Capabilities::default(),
                &StreamLimits::default(),
            )
            .is_err()
        );

        let duplicate = worksheet(&format!(
            r#"<sheetFormatPr x14ac:dyDescent="0.2"/><sheetData><row r="0" x14ac:dyDescent="-1" xmlns:a="{X14AC}" a:dyDescent="NaN"/></sheetData>"#
        ));
        assert!(
            x14ac::capture_stream_defaults(
                &mut Cursor::new(duplicate.as_bytes()),
                &Capabilities::default(),
                &StreamLimits::default(),
            )
            .is_err()
        );
    }

    #[test]
    fn streaming_0361_selects_supported_choice_or_fallback_and_ignores_inactive_bad_values() {
        let xml = format!(
            r#"<worksheet xmlns="{S}" xmlns:mc="{MC}" xmlns:x14ac="{X14AC}" mc:Ignorable="x14ac"><mc:AlternateContent><mc:Choice Requires="x14ac"><sheetFormatPr x14ac:dyDescent="NaN"/><sheetData><row r="0" x14ac:dyDescent="-1"/></sheetData></mc:Choice><mc:Fallback><sheetFormatPr x14ac:dyDescent="0.25"/><sheetData><row r="2" x14ac:dyDescent="0.35"/></sheetData></mc:Fallback></mc:AlternateContent></worksheet>"#
        );

        let fallback = default_stream(xml.as_bytes(), true).expect("fallback branch");
        assert_eq!(signature(&fallback), (Some(0.25), vec![(2, 0.35)]));

        let choice_xml = format!(
            r#"<worksheet xmlns="{S}" xmlns:mc="{MC}" xmlns:x14ac="{X14AC}" mc:Ignorable="x14ac"><mc:AlternateContent><mc:Choice Requires="x14ac"><sheetFormatPr x14ac:dyDescent="0.2"/><sheetData><row r="1" x14ac:dyDescent="0.3"/></sheetData></mc:Choice><mc:Fallback><sheetFormatPr/><sheetData/></mc:Fallback></mc:AlternateContent></worksheet>"#
        );

        let choice = stream(
            choice_xml.as_bytes(),
            &choices_capabilities(),
            &StreamLimits::default(),
            true,
        )
        .expect("supported choice branch");
        assert_eq!(signature(&choice), (Some(0.2), vec![(1, 0.3)]));
    }

    #[test]
    fn streaming_0361_raw_errors_survive_inactive_branches_and_later_malformed_tail() {
        let duplicate = format!(
            r#"<worksheet xmlns="{S}" xmlns:mc="{MC}" xmlns:x14ac="{X14AC}" xmlns:a="{X14AC}" xmlns:future="urn:future" mc:Ignorable="x14ac"><mc:AlternateContent><mc:Choice Requires="future"><sheetFormatPr x14ac:dyDescent="-1" a:dyDescent="NaN"/></mc:Choice><mc:Fallback><sheetFormatPr/></mc:Fallback></mc:AlternateContent></worksheet><tail/>"#
        );
        let duplicate_error = default_stream(duplicate.as_bytes(), true).expect_err("duplicate");
        assert!(matches!(
            duplicate_error,
            StreamError::Mce {
                raw_error: Some(raw),
                ..
            } if raw.to_string().contains("duplicate x14ac:dyDescent")
        ));

        let reserved = format!(
            r#"<worksheet xmlns="{S}" xmlns:mc="{MC}" xmlns:x14ac="{X14AC}" xmlns:future="urn:future" mc:Ignorable="x14ac"><mc:AlternateContent><mc:Choice Requires="future"><sheetFormatPr litchi_x14ac_dyDescent="NaN"/></mc:Choice><mc:Fallback><sheetFormatPr/></mc:Fallback></mc:AlternateContent></worksheet><tail/>"#
        );
        let reserved_error = default_stream(reserved.as_bytes(), true).expect_err("reserved");
        assert!(matches!(
            reserved_error,
            StreamError::Mce {
                raw_error: Some(raw),
                ..
            } if raw.to_string().contains("reserved internal marker")
        ));

        let legacy = x14ac::capture_stream_legacy(
            &mut Cursor::new(reserved.as_bytes()),
            &Capabilities::default(),
            &StreamLimits::default(),
            true,
        )
        .expect_err("legacy adapter must retain raw precedence");
        assert!(legacy.to_string().contains("reserved internal marker"));
    }

    #[test]
    fn streaming_0361_uses_expanded_namespace_after_prefix_rebinding() {
        let xml = format!(
            r#"<worksheet xmlns="{S}" xmlns:mc="{MC}" xmlns:x14ac="{X14AC}" mc:Ignorable="x14ac"><sheetFormatPr xmlns:x14ac="urn:not-x14ac" x14ac:dyDescent="NaN"/><sheetData><row xmlns:x14ac="{X14AC}" x14ac:dyDescent="0.3"/></sheetData></worksheet>"#
        );
        let values = default_stream(xml.as_bytes(), true).expect("rebinding fixture");
        assert_eq!(signature(&values), (None, vec![(1, 0.3)]));
    }

    #[test]
    fn streaming_0361_preserves_legacy_row_inference_and_order_semantics() {
        let continuity = worksheet(
            r#"<sheetData><row r="3" x14ac:dyDescent="0.3"/><row x14ac:dyDescent="0.4"/></sheetData>"#,
        );
        let continuity_stream = default_stream(continuity.as_bytes(), true)
            .expect("explicit and inferred rows")
            .rows;
        let continuity_legacy = x14ac::capture(continuity.as_bytes())
            .expect("legacy explicit and inferred rows")
            .rows;
        assert_eq!(
            continuity_stream
                .iter()
                .map(|(&row, &value)| (row, value.get()))
                .collect::<Vec<_>>(),
            continuity_legacy
                .iter()
                .map(|(&row, &value)| (row, value.get()))
                .collect::<Vec<_>>()
        );

        let decreasing = worksheet(
            r#"<sheetData><row r="3" x14ac:dyDescent="0.3"/><row r="2" x14ac:dyDescent="0.2"/></sheetData>"#,
        );
        assert_eq!(
            signature(&default_stream(decreasing.as_bytes(), true).expect("decreasing rows")),
            signature(&x14ac::capture(decreasing.as_bytes()).expect("legacy decreasing rows")),
        );

        let current = worksheet(
            r#"<sheetData><row r="1"/><row x14ac:dyDescent="0.2"/><row r="4"/><row x14ac:dyDescent="0.4"/></sheetData>"#,
        );
        assert_eq!(
            signature(&default_stream(current.as_bytes(), true).expect("current row state")),
            (None, vec![(2, 0.2), (5, 0.4)]),
        );

        let duplicate = worksheet(
            r#"<sheetData><row r="3" x14ac:dyDescent="0.3"/><row r="3" x14ac:dyDescent="0.4"/></sheetData>"#,
        );
        assert!(default_stream(duplicate.as_bytes(), true).is_err());
        assert!(x14ac::capture(duplicate.as_bytes()).is_err());
    }

    #[test]
    fn streaming_0361_rejects_zero_and_grid_overflow_only_when_rows_are_captured() {
        for row in ["0".to_owned(), (ROWS + 1).to_string()] {
            let xml = worksheet(&format!(
                r#"<sheetData><row r="{row}" x14ac:dyDescent="0.2"/></sheetData>"#
            ));
            assert!(
                default_stream(xml.as_bytes(), true).is_err(),
                "accepted row {row}"
            );
            assert!(
                default_stream(xml.as_bytes(), false).is_ok(),
                "focused defaults parsed row {row}"
            );
        }

        let inferred = worksheet(&format!(
            r#"<sheetData><row r="{ROWS}"/><row x14ac:dyDescent="0.2"/></sheetData>"#
        ));
        assert!(default_stream(inferred.as_bytes(), true).is_err());
    }

    #[test]
    fn streaming_0361_accepts_exact_depth_event_and_input_limits_but_rejects_one_under() {
        let xml = worksheet(r#"<sheetData><row r="1" x14ac:dyDescent="0.2"/></sheetData>"#);
        let exact_events = StreamLimits {
            max_events: 5,
            ..StreamLimits::default()
        };
        default_stream_with_limits(&xml, &exact_events, true).expect("five exact events");
        let under_events = StreamLimits {
            max_events: 4,
            ..exact_events.clone()
        };
        assert!(default_stream_with_limits(&xml, &under_events, true).is_err());

        let exact_depth = StreamLimits {
            processing: litchi_ooxml_common::mce::Limits {
                max_depth: 3,
                ..StreamLimits::default().processing
            },
            ..StreamLimits::default()
        };
        default_stream_with_limits(&xml, &exact_depth, true).expect("three exact levels");
        let under_depth = StreamLimits {
            processing: litchi_ooxml_common::mce::Limits {
                max_depth: 2,
                ..exact_depth.processing.clone()
            },
            ..exact_depth.clone()
        };
        assert!(default_stream_with_limits(&xml, &under_depth, true).is_err());

        let exact_input = StreamLimits {
            processing: litchi_ooxml_common::mce::Limits {
                max_input_bytes: xml.len(),
                ..StreamLimits::default().processing
            },
            ..StreamLimits::default()
        };
        default_stream_with_limits(&xml, &exact_input, true).expect("exact input size");
        let under_input = StreamLimits {
            processing: litchi_ooxml_common::mce::Limits {
                max_input_bytes: xml.len() - 1,
                ..exact_input.processing.clone()
            },
            ..exact_input
        };
        assert!(default_stream_with_limits(&xml, &under_input, true).is_err());

        let event = format!(r#"<worksheet xmlns="{S}"/>"#);
        let exact_event = StreamLimits {
            max_event_bytes: event.len(),
            ..StreamLimits::default()
        };
        default_stream_with_limits(event.as_str(), &exact_event, true).expect("exact event");
        let under_event = StreamLimits {
            max_event_bytes: event.len() - 1,
            ..exact_event
        };
        assert!(default_stream_with_limits(event.as_str(), &under_event, true).is_err());
    }

    fn default_stream_with_limits(
        xml: &str,
        limits: &StreamLimits,
        capture_rows: bool,
    ) -> x14ac::StreamResult<Values> {
        stream(
            xml.as_bytes(),
            &Capabilities::default(),
            limits,
            capture_rows,
        )
    }

    struct ChunkedInput {
        bytes: Vec<u8>,
        position: usize,
        chunk_size: usize,
    }

    impl ChunkedInput {
        fn new(bytes: &[u8], chunk_size: usize) -> Self {
            Self {
                bytes: bytes.to_owned(),
                position: 0,
                chunk_size,
            }
        }
    }

    impl Read for ChunkedInput {
        fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
            let available = self.fill_buf()?;
            let amount = available.len().min(output.len());
            output[..amount].copy_from_slice(&available[..amount]);
            self.consume(amount);
            Ok(amount)
        }
    }

    impl BufRead for ChunkedInput {
        fn fill_buf(&mut self) -> io::Result<&[u8]> {
            let end = self
                .position
                .saturating_add(self.chunk_size)
                .min(self.bytes.len());
            Ok(&self.bytes[self.position..end])
        }

        fn consume(&mut self, amount: usize) {
            self.position = self.position.saturating_add(amount).min(self.bytes.len());
        }
    }

    #[test]
    fn streaming_0361_split_bom_and_chunk_sizes_match_cursor() {
        let body = worksheet(
            r#"<sheetFormatPr x14ac:dyDescent="0.2"/><sheetData><row r="1" x14ac:dyDescent="0.3"/></sheetData>"#,
        );
        let mut xml = vec![0xef, 0xbb, 0xbf];
        xml.extend_from_slice(body.as_bytes());

        let cursor = signature(
            &stream(
                &xml,
                &Capabilities::default(),
                &StreamLimits::default(),
                true,
            )
            .expect("cursor stream"),
        );
        for chunk_size in [1, 2, 7] {
            let mut reader = ChunkedInput::new(&xml, chunk_size);
            let values = stream_reader(
                &mut reader,
                &Capabilities::default(),
                &StreamLimits::default(),
                true,
            )
            .expect("chunked stream");
            assert_eq!(signature(&values), cursor, "chunk size {chunk_size}");
        }
    }

    #[test]
    fn streaming_0361_valid_fixture_matches_legacy_capture_and_defaults() {
        let xml = worksheet(
            r#"<sheetFormatPr x14ac:dyDescent="0.2"/><sheetData><row r="2" x14ac:dyDescent="0.3"/><row r="4" x14ac:dyDescent="0.4"/></sheetData>"#,
        );
        let streamed = default_stream(xml.as_bytes(), true).expect("streamed fixture");
        let legacy = x14ac::capture(xml.as_bytes()).expect("legacy fixture");
        assert_eq!(signature(&streamed), signature(&legacy));

        let streamed_defaults = x14ac::capture_stream_defaults(
            &mut Cursor::new(xml.as_bytes()),
            &Capabilities::default(),
            &StreamLimits::default(),
        )
        .expect("streamed defaults");
        let legacy_defaults = x14ac::capture_defaults(xml.as_bytes()).expect("legacy defaults");
        assert_eq!(
            streamed_defaults.map(|value| value.get()),
            legacy_defaults.map(|value| value.get())
        );
    }

    struct FailAfter {
        bytes: Vec<u8>,
        position: usize,
    }

    impl FailAfter {
        fn new(bytes: &[u8]) -> Self {
            Self {
                bytes: bytes.to_owned(),
                position: 0,
            }
        }
    }

    impl Read for FailAfter {
        fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
            let available = self.fill_buf()?;
            let amount = available.len().min(output.len());
            output[..amount].copy_from_slice(&available[..amount]);
            self.consume(amount);
            Ok(amount)
        }
    }

    impl BufRead for FailAfter {
        fn fill_buf(&mut self) -> io::Result<&[u8]> {
            if self.position == self.bytes.len() {
                return Err(io::Error::other("streaming 0361 input failure"));
            }
            Ok(&self.bytes[self.position..])
        }

        fn consume(&mut self, amount: usize) {
            self.position = self.position.saturating_add(amount).min(self.bytes.len());
        }
    }

    fn dual_callback_failure_xml() -> String {
        worksheet(
            r#"<sheetFormatPr x14ac:dyDescent="not-a-number"/><sheetData><row r="1" litchi_x14ac_dyDescent="0.2"/></sheetData>"#,
        )
    }

    #[test]
    fn streaming_0361_diagnostics_keep_primary_input_or_mce_and_callback_secondaries() {
        let body = dual_callback_failure_xml();
        let mut input = FailAfter::new(body.as_bytes());
        let input_error = stream_reader(
            &mut input,
            &Capabilities::default(),
            &StreamLimits::default(),
            true,
        )
        .expect_err("input failure");
        assert!(matches!(
            input_error,
            StreamError::Input {
                raw_error: Some(raw),
                active_error: Some(active),
                ..
            } if raw.to_string().contains("reserved internal marker")
                && active.to_string().contains("invalid x14ac:dyDescent")
        ));

        let malformed_tail = format!("{body}<tail/>");
        let mce_error =
            default_stream(malformed_tail.as_bytes(), true).expect_err("malformed tail");
        assert!(matches!(
            mce_error,
            StreamError::Mce {
                error: MceError::NonConformant(_) | MceError::Xml(_),
                raw_error: Some(raw),
                active_error: Some(active),
                ..
            } if raw.to_string().contains("reserved internal marker")
                && active.to_string().contains("invalid x14ac:dyDescent")
        ));

        let legacy = x14ac::capture_stream_legacy(
            &mut Cursor::new(malformed_tail.as_bytes()),
            &Capabilities::default(),
            &StreamLimits::default(),
            true,
        )
        .expect_err("legacy raw precedence");
        assert!(legacy.to_string().contains("reserved internal marker"));
    }

    #[test]
    fn streaming_0361_legacy_adapter_maps_clean_eof_raw_then_active_errors() {
        let body = dual_callback_failure_xml();
        let error = default_stream(body.as_bytes(), true).expect_err("two callbacks");
        assert!(matches!(
            error,
            StreamError::Callback {
                raw_error: Some(raw),
                active_error: Some(active),
            } if raw.to_string().contains("reserved internal marker")
                && active.to_string().contains("invalid x14ac:dyDescent")
        ));
        let legacy = x14ac::capture_stream_legacy(
            &mut Cursor::new(body.as_bytes()),
            &Capabilities::default(),
            &StreamLimits::default(),
            true,
        )
        .expect_err("raw callback precedence");
        assert!(legacy.to_string().contains("reserved internal marker"));

        let active_only =
            worksheet(r#"<sheetData><row r="1" x14ac:dyDescent="not-a-number"/></sheetData>"#);
        let active_error = default_stream(active_only.as_bytes(), true).expect_err("active error");
        assert!(matches!(
            active_error,
            StreamError::Callback {
                raw_error: None,
                active_error: Some(active),
            } if active.to_string().contains("invalid x14ac:dyDescent")
        ));
        let legacy_active = x14ac::capture_stream_legacy(
            &mut Cursor::new(active_only.as_bytes()),
            &Capabilities::default(),
            &StreamLimits::default(),
            true,
        )
        .expect_err("active callback mapping");
        assert!(
            legacy_active
                .to_string()
                .contains("invalid x14ac:dyDescent")
        );
    }
}

#[cfg(test)]
mod streaming_0362_selected_tests {
    use std::io::{self, BufRead, Cursor, Read};

    use litchi_ooxml_common::mce::{Capabilities, Limits, StreamLimits};
    use litchi_sheet::Cell as Address;

    use super::super::{
        selected::{NotEligibleReason, ScanOutcome, SelectedCell, scan},
        x14ac,
    };
    use crate::cell::{Cell, Value};

    const SPREADSHEETML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    const X14AC: &str = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac";

    fn address(reference: &str) -> Address {
        Address::from_a1(reference).expect("valid worksheet address")
    }

    fn worksheet(body: &str) -> String {
        format!(r#"<worksheet xmlns="{SPREADSHEETML}">{body}</worksheet>"#)
    }

    fn x14ac_worksheet(body: &str) -> String {
        format!(
            r#"<worksheet xmlns="{SPREADSHEETML}" xmlns:mc="{MC}" xmlns:x14ac="{X14AC}" mc:Ignorable="x14ac">{body}</worksheet>"#
        )
    }

    fn alternate_content(choice: &str, fallback: &str) -> String {
        format!(
            r#"<worksheet xmlns="{SPREADSHEETML}" xmlns:mc="{MC}" xmlns:x14ac="{X14AC}" mc:Ignorable="x14ac"><mc:AlternateContent><mc:Choice Requires="x14ac">{choice}</mc:Choice><mc:Fallback>{fallback}</mc:Fallback></mc:AlternateContent></worksheet>"#
        )
    }

    fn scan_xml(xml: &str, target: &str) -> x14ac::StreamResult<ScanOutcome> {
        let capabilities = Capabilities::default();
        let limits = StreamLimits::default();
        scan_xml_with(xml, target, &capabilities, &limits)
    }

    fn scan_xml_with(
        xml: &str,
        target: &str,
        capabilities: &Capabilities,
        limits: &StreamLimits,
    ) -> x14ac::StreamResult<ScanOutcome> {
        let mut input = Cursor::new(xml.as_bytes());
        scan(&mut input, capabilities, limits, address(target))
    }

    fn eligible(xml: &str, target: &str) -> SelectedCell {
        match scan_xml(xml, target) {
            Ok(ScanOutcome::Eligible(selected)) => selected,
            other => panic!("expected eligible selection, got {other:?}"),
        }
    }

    fn assert_not_eligible(xml: &str, target: &str, expected: NotEligibleReason) {
        match scan_xml(xml, target) {
            Ok(ScanOutcome::NotEligible(reason)) => assert_eq!(reason, expected),
            other => panic!("expected {expected:?}, got {other:?}"),
        }
    }

    fn assert_number(selected: SelectedCell, expected: &str) {
        match selected.cell {
            Some(Cell::Value(Value::Number(value))) => assert_eq!(value.as_str(), expected),
            other => panic!("expected number {expected}, got {other:?}"),
        }
    }

    fn assert_text(selected: SelectedCell, expected: &str) {
        match selected.cell {
            Some(Cell::Value(Value::Text(value))) => assert_eq!(value.as_str(), expected),
            other => panic!("expected text {expected}, got {other:?}"),
        }
    }

    #[test]
    fn streaming_0362_selected_outcomes_cover_first_last_missing_and_empty() {
        let xml = worksheet(
            r#"<sheetData><row r="1"><c r="A1"><v>-0.000</v></c></row><row r="3"><c r="C3"/></row></sheetData>"#,
        );

        let first = eligible(&xml, "A1");
        assert_eq!(first.address, address("A1"));
        assert_number(first, "-0.000");

        let last = eligible(&xml, "C3");
        assert_eq!(last.address, address("C3"));
        assert!(matches!(last.cell, Some(Cell::Empty)));

        let missing = eligible(&xml, "B2");
        assert_eq!(missing.address, address("B2"));
        assert!(missing.cell.is_none());
    }

    #[test]
    fn streaming_0362_selected_payloads_cover_scalar_variants() {
        let boolean = eligible(
            &worksheet(r#"<sheetData><row r="1"><c r="A1" t="b"><v>1</v></c></row></sheetData>"#),
            "A1",
        );
        assert!(matches!(boolean.cell, Some(Cell::Value(Value::Bool(true)))));

        let error = eligible(
            &worksheet(
                r#"<sheetData><row r="1"><c r="A1" t="e"><v>#DIV/0!</v></c></row></sheetData>"#,
            ),
            "A1",
        );
        assert!(matches!(
            error.cell,
            Some(Cell::Value(Value::Error(value))) if value.as_str() == "#DIV/0!"
        ));

        let plain = eligible(
            &worksheet(
                r#"<sheetData><row r="1"><c r="A1" t="str"><v>plain</v></c></row></sheetData>"#,
            ),
            "A1",
        );
        assert_text(plain, "plain");

        let formula_string = eligible(
            &worksheet(
                r#"<sheetData><row r="1"><c r="A1" t="str"><f>CONCAT("a","b")</f><v>ab</v></c></row></sheetData>"#,
            ),
            "A1",
        );
        match formula_string.cell {
            Some(Cell::Formula(formula)) => {
                assert_eq!(formula.text(), "CONCAT(\"a\",\"b\")");
                assert!(formula.cached().is_some());
            },
            other => panic!("expected formula-string cell, got {other:?}"),
        }

        let inline = eligible(
            &worksheet(
                r#"<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>inline text</t></is></c></row></sheetData>"#,
            ),
            "A1",
        );
        assert_text(inline, "inline text");

        let scalar_formula = eligible(
            &worksheet(
                r#"<sheetData><row r="1"><c r="A1"><f>SUM(B1:B2)</f><v>3</v></c></row></sheetData>"#,
            ),
            "A1",
        );
        match scalar_formula.cell {
            Some(Cell::Formula(formula)) => {
                assert_eq!(formula.text(), "SUM(B1:B2)");
                assert!(formula.cached().is_some());
            },
            other => panic!("expected scalar formula cell, got {other:?}"),
        }
    }

    #[test]
    fn streaming_0362_selected_rejects_ordering_and_adjacent_duplicates() {
        let decreasing_rows = worksheet(
            r#"<sheetData><row r="2"><c r="A2"><v>2</v></c></row><row r="1"><c r="A1"><v>1</v></c></row></sheetData>"#,
        );
        assert_not_eligible(&decreasing_rows, "A1", NotEligibleReason::Ordering);

        let decreasing_cells = worksheet(
            r#"<sheetData><row r="1"><c r="B1"><v>2</v></c><c r="A1"><v>1</v></c></row></sheetData>"#,
        );
        assert_not_eligible(&decreasing_cells, "A1", NotEligibleReason::Ordering);

        let duplicate_row = worksheet(r#"<sheetData><row r="1"/><row r="1"/></sheetData>"#);
        assert!(scan_xml(&duplicate_row, "A1").is_err());

        let duplicate_cell =
            worksheet(r#"<sheetData><row r="1"><c r="A1"/><c r="A1"/></row></sheetData>"#);
        assert!(scan_xml(&duplicate_cell, "A1").is_err());
    }

    #[test]
    fn streaming_0362_selected_marks_unsupported_structures_not_eligible() {
        let cases = [
            (
                worksheet(
                    r#"<sheetData/><mergeCells futureAttr="preserve" count="1"><mergeCell ref="A1:B1"/></mergeCells>"#,
                ),
                NotEligibleReason::UnsupportedStructure,
            ),
            (
                worksheet(
                    r#"<cols><col min="1" max="1" width="10" customWidth="1"/></cols><sheetData/>"#,
                ),
                NotEligibleReason::Styles,
            ),
            (
                worksheet(
                    r#"<sheetData><row r="1"><c r="A1"><f t="shared" si="0" ref="A1:A2">A1+1</f><v>2</v></c></row></sheetData>"#,
                ),
                NotEligibleReason::FormulaSemantics,
            ),
            (
                worksheet(
                    r#"<sheetData><row r="1"><c r="A1" t="inlineStr"><is><r><rPr><b/></rPr><t>rich</t></r></is></c></row></sheetData>"#,
                ),
                NotEligibleReason::RichInlineText,
            ),
            (
                worksheet(r#"<sheetPr/><sheetData/>"#),
                NotEligibleReason::UnsupportedStructure,
            ),
        ];

        for (xml, reason) in cases {
            assert_not_eligible(&xml, "A1", reason);
        }

        let shared_string =
            worksheet(r#"<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData>"#);
        let ScanOutcome::Eligible(selected) =
            scan_xml(&shared_string, "A1").expect("valid shared-string cell")
        else {
            panic!("expected eligible deferred shared-string target");
        };
        assert!(selected.cell.is_none());
        assert_eq!(selected.dependencies.target_shared_string_index, Some(0));
        assert_eq!(selected.dependencies.max_shared_string_index, Some(0));
    }

    #[test]
    fn streaming_0362_selected_chooses_supported_choice_or_fallback() {
        let choice_supported = alternate_content(
            r#"<sheetData><row r="1"><c r="A1"><v>7</v></c></row></sheetData>"#,
            r#"<mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>"#,
        );
        let mut capabilities = Capabilities::default();
        capabilities.understand_namespace(X14AC);
        let selected_choice = scan_xml_with(
            &choice_supported,
            "A1",
            &capabilities,
            &StreamLimits::default(),
        )
        .expect("supported choice stream");
        match selected_choice {
            ScanOutcome::Eligible(selected) => assert_number(selected, "7"),
            other => panic!("expected supported choice, got {other:?}"),
        }

        let fallback_selected = alternate_content(
            r#"<mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>"#,
            r#"<sheetData><row r="1"><c r="A1"><v>8</v></c></row></sheetData>"#,
        );
        match scan_xml(&fallback_selected, "A1").expect("fallback stream") {
            ScanOutcome::Eligible(selected) => assert_number(selected, "8"),
            other => panic!("expected fallback branch, got {other:?}"),
        }
    }

    #[test]
    fn streaming_0362_selected_validates_x14ac_without_retaining_rows() {
        let valid = x14ac_worksheet(
            r#"<sheetData><row r="1" x14ac:dyDescent="0.3"><c r="A1"><v>9</v></c></row></sheetData>"#,
        );
        assert_number(eligible(&valid, "A1"), "9");

        let values = x14ac::capture_stream_with_active(
            &mut Cursor::new(valid.as_bytes()),
            &Capabilities::default(),
            &StreamLimits::default(),
            x14ac::RowMode::ValidateOnly,
            |_| Ok(()),
        )
        .expect("valid x14ac descent");
        assert!(values.rows.is_empty());

        let malformed = x14ac_worksheet(
            r#"<sheetData><row r="1" x14ac:dyDescent="not-a-number"><c r="A1"><v>9</v></c></row></sheetData>"#,
        );
        assert!(scan_xml(&malformed, "A1").is_err());
        assert!(
            x14ac::capture_stream_with_active(
                &mut Cursor::new(malformed.as_bytes()),
                &Capabilities::default(),
                &StreamLimits::default(),
                x14ac::RowMode::ValidateOnly,
                |_| Ok(()),
            )
            .is_err()
        );
    }

    #[test]
    fn streaming_0362_selected_does_not_return_early_before_malformed_tail() {
        let xml = format!(
            r#"<worksheet xmlns="{SPREADSHEETML}"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet><tail>"#
        );
        assert!(scan_xml(&xml, "A1").is_err());
    }

    #[test]
    fn streaming_0362_selected_respects_exact_limits_and_split_input() {
        let explicit_empty = worksheet(r#"<sheetData><row r="1"><c r="A1"/></row></sheetData>"#);
        let exact_events = StreamLimits {
            max_events: 7,
            ..StreamLimits::default()
        };
        assert!(
            scan_xml_with(
                &explicit_empty,
                "A1",
                &Capabilities::default(),
                &exact_events,
            )
            .is_ok()
        );
        let under_events = StreamLimits {
            max_events: 6,
            ..exact_events.clone()
        };
        assert!(
            scan_xml_with(
                &explicit_empty,
                "A1",
                &Capabilities::default(),
                &under_events,
            )
            .is_err()
        );

        let exact_input = StreamLimits {
            processing: Limits {
                max_input_bytes: explicit_empty.len(),
                ..StreamLimits::default().processing
            },
            ..StreamLimits::default()
        };
        assert!(
            scan_xml_with(
                &explicit_empty,
                "A1",
                &Capabilities::default(),
                &exact_input,
            )
            .is_ok()
        );
        let under_input = StreamLimits {
            processing: Limits {
                max_input_bytes: explicit_empty.len() - 1,
                ..exact_input.processing.clone()
            },
            ..exact_input.clone()
        };
        assert!(
            scan_xml_with(
                &explicit_empty,
                "A1",
                &Capabilities::default(),
                &under_input,
            )
            .is_err()
        );

        let empty_root = format!(r#"<worksheet xmlns="{SPREADSHEETML}"><sheetData/></worksheet>"#,);
        let root_start = format!(r#"<worksheet xmlns="{SPREADSHEETML}">"#);
        let exact_event = StreamLimits {
            max_event_bytes: root_start.len(),
            ..StreamLimits::default()
        };
        assert!(scan_xml_with(&empty_root, "A1", &Capabilities::default(), &exact_event,).is_ok());
        let under_event = StreamLimits {
            max_event_bytes: root_start.len() - 1,
            ..exact_event.clone()
        };
        assert!(scan_xml_with(&empty_root, "A1", &Capabilities::default(), &under_event,).is_err());

        let split_xml =
            worksheet(r#"<sheetData><row r="1"><c r="A1"><v>11</v></c></row></sheetData>"#);
        for chunk_size in [1, 2, 7] {
            let mut input = ChunkedInput::new(split_xml.as_bytes(), chunk_size);
            match scan(
                &mut input,
                &Capabilities::default(),
                &StreamLimits::default(),
                address("A1"),
            )
            .expect("split selected stream")
            {
                ScanOutcome::Eligible(selected) => assert_number(selected, "11"),
                other => panic!("expected split eligible result, got {other:?}"),
            }
        }
    }

    struct ChunkedInput {
        bytes: Vec<u8>,
        position: usize,
        chunk_size: usize,
    }

    impl ChunkedInput {
        fn new(bytes: &[u8], chunk_size: usize) -> Self {
            Self {
                bytes: bytes.to_vec(),
                position: 0,
                chunk_size,
            }
        }
    }

    impl Read for ChunkedInput {
        fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
            let available = self.fill_buf()?;
            let amount = available.len().min(output.len());
            output[..amount].copy_from_slice(&available[..amount]);
            self.consume(amount);
            Ok(amount)
        }
    }

    impl BufRead for ChunkedInput {
        fn fill_buf(&mut self) -> io::Result<&[u8]> {
            let end = self
                .position
                .saturating_add(self.chunk_size)
                .min(self.bytes.len());
            Ok(&self.bytes[self.position..end])
        }

        fn consume(&mut self, amount: usize) {
            self.position = self.position.saturating_add(amount).min(self.bytes.len());
        }
    }
}

#[cfg(test)]
mod streaming_0364_dependency_metadata_tests {
    use std::io::Cursor;

    use litchi_ooxml_common::mce::{Capabilities, Error as MceError, StreamError, StreamLimits};
    use litchi_sheet::Cell as Address;

    use super::super::selected::{NotEligibleReason, ScanOutcome, scan};
    use crate::cell::{Cell, Value};

    const SPREADSHEETML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    fn address(reference: &str) -> Address {
        Address::from_a1(reference).expect("valid worksheet address")
    }

    fn worksheet(body: &str) -> String {
        format!(r#"<worksheet xmlns="{SPREADSHEETML}">{body}</worksheet>"#)
    }

    fn scan_xml(xml: &str, target: &str) -> super::super::selected::StreamResult<ScanOutcome> {
        let mut input = Cursor::new(xml.as_bytes());
        scan(
            &mut input,
            &Capabilities::default(),
            &StreamLimits::default(),
            address(target),
        )
    }

    #[test]
    fn streaming_0364_dependency_metadata_tracks_global_maxima_and_target_states() {
        let xml = worksheet(
            r#"<sheetData>
                <row r="1">
                    <c r="A1" t="s" s="4"><v>7</v></c>
                    <c r="B1" t="s" s="13"><v>11</v></c>
                    <c r="C1" s="8"><v>1</v></c>
                </row>
                <row r="3"><c r="C3"/></row>
            </sheetData>"#,
        );

        let ScanOutcome::Eligible(deferred) = scan_xml(&xml, "A1").expect("deferred target") else {
            panic!("expected eligible deferred target");
        };
        assert!(deferred.cell.is_none());
        assert_eq!(deferred.dependencies.max_shared_string_index, Some(11));
        assert_eq!(deferred.dependencies.max_direct_style_index, Some(13));
        assert_eq!(deferred.dependencies.target_shared_string_index, Some(7));

        let ScanOutcome::Eligible(missing) = scan_xml(&xml, "B2").expect("missing target") else {
            panic!("expected eligible missing target");
        };
        assert!(missing.cell.is_none());
        assert_eq!(missing.dependencies.max_shared_string_index, Some(11));
        assert_eq!(missing.dependencies.max_direct_style_index, Some(13));
        assert_eq!(missing.dependencies.target_shared_string_index, None);

        let ScanOutcome::Eligible(empty) = scan_xml(&xml, "C3").expect("explicit empty target")
        else {
            panic!("expected eligible explicit empty target");
        };
        assert!(matches!(empty.cell, Some(Cell::Empty)));
        assert_eq!(empty.dependencies.max_shared_string_index, Some(11));
        assert_eq!(empty.dependencies.max_direct_style_index, Some(13));
        assert_eq!(empty.dependencies.target_shared_string_index, None);
    }

    #[test]
    fn streaming_0364_dependency_metadata_keeps_valid_direct_styles_eligible() {
        let xml = worksheet(
            r#"<sheetData><row r="1">
                <c r="A1" s="2"><v>1</v></c>
                <c r="B1" s="17"><v>2</v></c>
                <c r="C1" s="5"><v>3</v></c>
            </row></sheetData>"#,
        );
        let ScanOutcome::Eligible(selected) = scan_xml(&xml, "A1").expect("valid direct styles")
        else {
            panic!("valid direct styles must remain eligible");
        };
        assert!(matches!(
            selected.cell,
            Some(Cell::Value(Value::Number(value))) if value.as_str() == "1"
        ));
        assert_eq!(selected.dependencies.max_shared_string_index, None);
        assert_eq!(selected.dependencies.max_direct_style_index, Some(17));
        assert_eq!(selected.dependencies.target_shared_string_index, None);
    }

    #[test]
    fn streaming_0364_dependency_metadata_marks_invalid_lexicals_not_eligible() {
        let cases = [
            (
                worksheet(
                    r#"<sheetData><row r="1"><c r="A1" s="not-a-style"><v>1</v></c></row></sheetData>"#,
                ),
                NotEligibleReason::Styles,
            ),
            (
                worksheet(
                    r#"<sheetData><row r="1"><c r="A1" t="s"><v>not-an-index</v></c></row></sheetData>"#,
                ),
                NotEligibleReason::SharedStrings,
            ),
            (
                worksheet(
                    r#"<sheetData><row r="1"><c r="A1" t="s"><f>1+1</f><v>0</v></c></row></sheetData>"#,
                ),
                NotEligibleReason::FormulaSemantics,
            ),
        ];

        for (xml, expected) in cases {
            match scan_xml(&xml, "A1").expect("metadata scan") {
                ScanOutcome::NotEligible(actual) => assert_eq!(actual, expected),
                other => panic!("expected {expected:?}, got {other:?}"),
            }
        }
    }

    #[test]
    fn streaming_0364_dependency_metadata_keeps_late_malformed_xml_primary() {
        let xml = format!(
            r#"<worksheet xmlns="{SPREADSHEETML}"><sheetData><row r="1"><c r="A1" s="not-a-style"><v>1</v></c></row></sheetData></worksheet><tail>"#
        );
        let error = scan_xml(&xml, "A1").expect_err("late malformed XML");
        assert!(matches!(
            error,
            StreamError::Mce {
                error: MceError::NonConformant(_) | MceError::Xml(_),
                ..
            }
        ));
    }
}

#[cfg(test)]
mod streaming_0365_range_tests {
    use std::io::Cursor;

    use litchi_ooxml_common::mce::{Capabilities, Error as MceError, StreamError, StreamLimits};
    use litchi_sheet::{Cell as Address, Rect};

    use super::super::model::MAX_CELL_STYLE;
    use super::super::selected::{
        NotEligibleReason, RangeScanOutcome, ScanOutcome, SelectedCells, scan, scan_range,
    };
    use crate::cell::{Cell, Value};

    const SPREADSHEETML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    fn address(reference: &str) -> Address {
        Address::from_a1(reference).expect("valid worksheet address")
    }

    fn worksheet(body: &str) -> String {
        format!(r#"<worksheet xmlns="{SPREADSHEETML}">{body}</worksheet>"#)
    }

    fn scan_range_xml(
        xml: &str,
        reference: &str,
    ) -> super::super::selected::StreamResult<RangeScanOutcome> {
        let mut input = Cursor::new(xml.as_bytes());
        scan_range(
            &mut input,
            &Capabilities::default(),
            &StreamLimits::default(),
            Rect::from_a1(reference).expect("valid worksheet range"),
        )
    }

    fn scan_xml(xml: &str, target: &str) -> super::super::selected::StreamResult<ScanOutcome> {
        let mut input = Cursor::new(xml.as_bytes());
        scan(
            &mut input,
            &Capabilities::default(),
            &StreamLimits::default(),
            address(target),
        )
    }

    fn eligible_range(xml: &str, reference: &str) -> SelectedCells {
        match scan_range_xml(xml, reference) {
            Ok(RangeScanOutcome::Eligible(selected)) => selected,
            other => panic!("expected eligible range, got {other:?}"),
        }
    }

    fn assert_range_not_eligible(xml: &str, expected: NotEligibleReason) {
        match scan_range_xml(xml, "A1") {
            Ok(RangeScanOutcome::NotEligible(actual)) => assert_eq!(actual, expected),
            other => panic!("expected {expected:?}, got {other:?}"),
        }
    }

    fn assert_number(cell: Option<&Cell>, expected: &str) {
        match cell {
            Some(Cell::Value(Value::Number(value))) => assert_eq!(value.as_str(), expected),
            other => panic!("expected number {expected}, got {other:?}"),
        }
    }

    fn assert_same_cell(left: Option<&Cell>, right: Option<&Cell>) {
        match (left, right) {
            (None, None) | (Some(Cell::Empty), Some(Cell::Empty)) => {},
            (Some(Cell::Value(Value::Number(left))), Some(Cell::Value(Value::Number(right)))) => {
                assert_eq!(left.as_str(), right.as_str())
            },
            (Some(Cell::Value(Value::Text(left))), Some(Cell::Value(Value::Text(right)))) => {
                assert_eq!(left.as_str(), right.as_str());
            },
            other => panic!("single-cell and range results differ: {other:?}"),
        }
    }

    #[test]
    fn streaming_0365_range_is_sparse_row_major_with_explicit_empty_records() {
        let xml = worksheet(
            r#"<sheetData>
                <row r="1"><c r="C1"><v>31</v></c><c r="E1"/></row>
                <row r="2"><c r="B2"><v>22</v></c></row>
                <row r="4"><c r="A4" t="str"/></row>
            </sheetData>"#,
        );
        let selected = eligible_range(&xml, "A1:E4");

        assert_eq!(
            selected
                .cells
                .iter()
                .map(|record| record.address)
                .collect::<Vec<_>>(),
            vec![address("C1"), address("E1"), address("B2"), address("A4")]
        );
        assert_number(selected.cells[0].cell.as_ref(), "31");
        assert!(matches!(selected.cells[1].cell.as_ref(), Some(Cell::Empty)));
        assert_number(selected.cells[2].cell.as_ref(), "22");
        assert!(matches!(
            selected.cells[3].cell.as_ref(),
            Some(Cell::Value(Value::Text(value))) if value.as_str().is_empty()
        ));
        assert!(
            selected
                .cells
                .iter()
                .all(|record| record.shared_string_index.is_none())
        );
    }

    #[test]
    fn streaming_0365_range_handles_first_last_grid_edges_and_empty_selection() {
        let xml = worksheet(
            r#"<sheetData>
                <row r="1"><c r="A1"><v>1</v></c></row>
                <row r="1048576"><c r="XFD1048576"><v>9</v></c></row>
            </sheetData>"#,
        );

        let first = eligible_range(&xml, "A1:A1");
        assert_eq!(first.cells.len(), 1);
        assert_eq!(first.cells[0].address, address("A1"));
        assert_number(first.cells[0].cell.as_ref(), "1");

        let last = eligible_range(&xml, "XFD1048576:XFD1048576");
        assert_eq!(last.cells.len(), 1);
        assert_eq!(last.cells[0].address, address("XFD1048576"));
        assert_number(last.cells[0].cell.as_ref(), "9");

        let edge = eligible_range(&xml, "XFC1048575:XFD1048576");
        assert_eq!(edge.cells.len(), 1);
        assert_eq!(edge.cells[0].address, address("XFD1048576"));

        let empty = eligible_range(&xml, "B2:C3");
        assert!(empty.cells.is_empty());
        assert_eq!(empty.dependencies, Default::default());
    }

    #[test]
    fn streaming_0365_range_defers_multiple_shared_strings_and_keeps_global_maxima() {
        let xml = worksheet(
            r#"<sheetData>
                <row r="1">
                    <c r="A1" t="s" s="4"><v>3</v></c>
                    <c r="B1" t="s" s="8"><v>9</v></c>
                </row>
                <row r="2">
                    <c r="A2" t="s" s="12"><v>11</v></c>
                    <c r="C2" s="37"><v>1</v></c>
                </row>
                <row r="3"><c r="D3" t="s"><v>19</v></c></row>
            </sheetData>"#,
        );
        let selected = eligible_range(&xml, "A1:B2");

        assert_eq!(selected.cells.len(), 3);
        assert_eq!(
            selected
                .cells
                .iter()
                .map(|record| record.shared_string_index)
                .collect::<Vec<_>>(),
            vec![Some(3), Some(9), Some(11)]
        );
        assert!(selected.cells.iter().all(|record| record.cell.is_none()));
        assert_eq!(selected.dependencies.max_shared_string_index, Some(19));
        assert_eq!(selected.dependencies.max_direct_style_index, Some(37));
        assert_eq!(selected.dependencies.target_shared_string_index, None);
    }

    #[test]
    fn streaming_0365_single_cell_scan_wrapper_matches_one_cell_ranges() {
        let xml = worksheet(
            r#"<sheetData><row r="1">
                <c r="A1"><v>7</v></c>
                <c r="B1"/>
                <c r="C1" t="str"/>
                <c r="D1" t="s"><v>5</v></c>
            </row></sheetData>"#,
        );

        for target in ["A1", "B1", "C1", "D1", "E1"] {
            let range = match scan_range_xml(&xml, target).expect("one-cell range scan") {
                RangeScanOutcome::Eligible(selected) => selected,
                other => panic!("expected eligible one-cell range, got {other:?}"),
            };
            let single = match scan_xml(&xml, target).expect("single-cell scan") {
                ScanOutcome::Eligible(selected) => selected,
                other => panic!("expected eligible single-cell scan, got {other:?}"),
            };

            assert_eq!(range.dependencies, single.dependencies);
            let range_record = range.cells.first();
            assert_same_cell(
                range_record.and_then(|record| record.cell.as_ref()),
                single.cell.as_ref(),
            );
            if let Some(record) = range_record {
                assert_eq!(record.address, single.address);
                assert_eq!(
                    record.shared_string_index,
                    single.dependencies.target_shared_string_index
                );
            } else {
                assert!(single.cell.is_none());
                assert_eq!(single.dependencies.target_shared_string_index, None);
            }
        }
    }

    #[test]
    fn streaming_0365_range_validates_global_duplicates_order_and_late_xml_errors() {
        let duplicate_row =
            worksheet(r#"<sheetData><row r="1"/><row r="2"/><row r="2"/></sheetData>"#);
        assert!(scan_range_xml(&duplicate_row, "A1").is_err());

        let duplicate_cell = worksheet(
            r#"<sheetData><row r="1"/><row r="2"><c r="B2"/><c r="B2"/></row></sheetData>"#,
        );
        assert!(scan_range_xml(&duplicate_cell, "A1").is_err());

        let decreasing_rows = worksheet(
            r#"<sheetData><row r="2"><c r="A2"><v>2</v></c></row><row r="1"><c r="A1"><v>1</v></c></row></sheetData>"#,
        );
        match scan_range_xml(&decreasing_rows, "A1").expect("ordering stream") {
            RangeScanOutcome::NotEligible(NotEligibleReason::Ordering) => {},
            other => panic!("expected global row-order fallback, got {other:?}"),
        }

        let decreasing_cells = worksheet(
            r#"<sheetData><row r="1"><c r="B1"><v>2</v></c><c r="A1"><v>1</v></c></row></sheetData>"#,
        );
        match scan_range_xml(&decreasing_cells, "A1").expect("cell ordering stream") {
            RangeScanOutcome::NotEligible(NotEligibleReason::Ordering) => {},
            other => panic!("expected global cell-order fallback, got {other:?}"),
        }

        let malformed_tail = format!(
            "{}<tail>",
            worksheet(r#"<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>"#)
        );
        let error = scan_range_xml(&malformed_tail, "A1").expect_err("late malformed XML");
        assert!(matches!(
            error,
            StreamError::Mce {
                error: MceError::NonConformant(_) | MceError::Xml(_),
                ..
            }
        ));
    }

    #[test]
    fn streaming_0365_range_marks_formula_and_inherited_styles_not_eligible() {
        let merged = worksheet(
            r#"<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData><mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>"#,
        );
        let selected = eligible_range(&merged, "A1:B1");
        assert_eq!(selected.cells.len(), 1);
        assert_number(selected.cells[0].cell.as_ref(), "1");

        let cases = [
            (
                worksheet(
                    r#"<sheetData><row r="1"><c r="A1"><f t="shared" si="0" ref="A1:A2">A1+1</f><v>2</v></c></row></sheetData>"#,
                ),
                NotEligibleReason::FormulaSemantics,
            ),
            (
                worksheet(
                    r#"<sheetData><row r="1" s="1"><c r="A1"><v>1</v></c></row></sheetData>"#,
                ),
                NotEligibleReason::Styles,
            ),
            (
                worksheet(r#"<cols><col min="1" max="1" style="1"/></cols><sheetData/>"#),
                NotEligibleReason::Styles,
            ),
        ];

        for (xml, expected) in cases {
            assert_range_not_eligible(&xml, expected);
        }
    }

    #[test]
    fn streaming_0365_range_accepts_max_cell_style_and_falls_back_above_it() {
        let valid = worksheet(&format!(
            r#"<sheetData><row r="1"><c r="A1" s="{MAX_CELL_STYLE}"><v>1</v></c></row></sheetData>"#
        ));
        let selected = eligible_range(&valid, "A1");
        assert_number(selected.cells[0].cell.as_ref(), "1");
        assert_eq!(
            selected.dependencies.max_direct_style_index,
            Some(MAX_CELL_STYLE)
        );

        let too_large = MAX_CELL_STYLE + 1;
        let invalid = worksheet(&format!(
            r#"<sheetData><row r="1"><c r="A1" s="{too_large}"><v>1</v></c></row></sheetData>"#
        ));
        assert_range_not_eligible(&invalid, NotEligibleReason::Styles);
    }

    #[test]
    fn streaming_0365_range_preserves_inline_and_escape_heavy_empty_boundaries() {
        let xml = worksheet(
            r#"<sheetData><row r="1">
                <c r="A1" t="inlineStr"/>
                <c r="B1" t="inlineStr"><is><t/></is></c>
                <c r="C1" t="str"/>
                <c r="D1" t="inlineStr"><is><t>_x0026__xD83D__xDE00__x003C__x0041_</t></is></c>
            </row></sheetData>"#,
        );
        let selected = eligible_range(&xml, "A1:D1");

        assert!(matches!(
            selected.cells[0].cell.as_ref(),
            Some(Cell::Value(Value::Text(value))) if value.as_str().is_empty()
        ));
        assert!(matches!(
            selected.cells[1].cell.as_ref(),
            Some(Cell::Value(Value::Text(value))) if value.as_str().is_empty()
        ));
        assert!(matches!(
            selected.cells[2].cell.as_ref(),
            Some(Cell::Value(Value::Text(value))) if value.as_str().is_empty()
        ));
        assert!(matches!(
            selected.cells[3].cell.as_ref(),
            Some(Cell::Value(Value::Text(value)))
                if value.as_str() == "&😀<A" && value.as_str().chars().count() == 4
        ));

        let formula_without_inline_payload = worksheet(
            r#"<sheetData><row r="1"><c r="A1" t="inlineStr"><f>1+1</f></c></row></sheetData>"#,
        );
        assert!(
            scan_range_xml(&formula_without_inline_payload, "A1").is_err(),
            "formula-bearing inlineStr without is must be rejected"
        );
    }
}

#[cfg(test)]
mod streaming_0367_merge_tests {
    use std::fmt::{Debug, Write as _};
    use std::io::Cursor;

    use litchi_ooxml_common::mce::{Capabilities, Error as MceError, StreamError, StreamLimits};
    use litchi_sheet::{Cell as Address, Rect};

    use super::super::selected::{
        NotEligibleReason, RangeScanOutcome, ScanOutcome, SelectedCell, scan, scan_range,
    };
    use crate::cell::{Cell, Value};

    const SPREADSHEETML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    const STREAMING_MERGE_CAP: usize = 16_384;

    type StreamResult<T> = super::super::selected::StreamResult<T>;

    fn address(reference: &str) -> Address {
        Address::from_a1(reference).expect("valid worksheet address")
    }

    fn range(reference: &str) -> Rect {
        Rect::from_a1(reference).expect("valid worksheet range")
    }

    fn worksheet(body: &str) -> String {
        format!(r#"<worksheet xmlns="{SPREADSHEETML}">{body}</worksheet>"#)
    }

    fn scan_xml(xml: &str, target: &str) -> StreamResult<ScanOutcome> {
        let mut input = Cursor::new(xml.as_bytes());
        scan(
            &mut input,
            &Capabilities::default(),
            &StreamLimits::default(),
            address(target),
        )
    }

    fn scan_range_xml(xml: &str, reference: &str) -> StreamResult<RangeScanOutcome> {
        let mut input = Cursor::new(xml.as_bytes());
        scan_range(
            &mut input,
            &Capabilities::default(),
            &StreamLimits::default(),
            range(reference),
        )
    }

    fn eligible(xml: &str, target: &str) -> SelectedCell {
        match scan_xml(xml, target) {
            Ok(ScanOutcome::Eligible(selected)) => selected,
            other => panic!("expected eligible selection, got {other:?}"),
        }
    }

    fn eligible_range(xml: &str, reference: &str) -> super::super::selected::SelectedCells {
        match scan_range_xml(xml, reference) {
            Ok(RangeScanOutcome::Eligible(selected)) => selected,
            other => panic!("expected eligible range, got {other:?}"),
        }
    }

    fn assert_merge_fallback(xml: &str, target: &str) {
        match scan_xml(xml, target) {
            Ok(ScanOutcome::NotEligible(_)) => {},
            other => panic!("expected merge fallback, got {other:?}"),
        }
    }

    fn assert_number(selected: SelectedCell, expected: &str) {
        match selected.cell {
            Some(Cell::Value(Value::Number(value))) => assert_eq!(value.as_str(), expected),
            other => panic!("expected number {expected}, got {other:?}"),
        }
    }

    fn assert_stream_xml_error<T: Debug>(result: StreamResult<T>) {
        match result {
            Err(StreamError::Mce {
                error: MceError::Xml(_) | MceError::NonConformant(_),
                ..
            }) => {},
            other => panic!("expected typed XML stream error, got {other:?}"),
        }
    }

    fn cap_matrix(count: usize) -> Result<String, std::collections::TryReserveError> {
        let mut xml = String::new();
        xml.try_reserve(count.saturating_mul(40).saturating_add(256))?;
        write!(
            &mut xml,
            r#"<worksheet xmlns="{SPREADSHEETML}"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData><mergeCells count="{count}">"#
        )
        .expect("String formatting cannot fail");
        for index in 0..count {
            let start = index * 2 + 1;
            let end = start + 1;
            write!(&mut xml, r#"<mergeCell ref="A{start}:A{end}"/>"#)
                .expect("String formatting cannot fail");
        }
        xml.push_str("</mergeCells></worksheet>");
        Ok(xml)
    }

    #[test]
    fn streaming_0367_accepts_unsorted_adjacent_merges() {
        let xml = worksheet(
            r#"<sheetData><row r="1"><c r="A1"><v>1</v></c><c r="C1"><v>3</v></c></row></sheetData><mergeCells count="2"><mergeCell ref="C1:D1"/><mergeCell ref="A1:B1"/></mergeCells>"#,
        );

        let left = eligible(&xml, "A1");
        assert_eq!(left.covering_merge, None);
        assert_number(left, "1");

        let right = eligible(&xml, "C1");
        assert_eq!(right.covering_merge, None);
        assert_number(right, "3");
    }

    #[test]
    fn streaming_0367_single_cell_wrapper_classifies_anchor_covered_missing_and_follower() {
        let xml = worksheet(
            r#"<sheetData><row r="1"><c r="A1"><v>7</v></c></row><row r="2"><c r="B2"><v>9</v></c></row></sheetData><mergeCells count="1"><mergeCell ref="A1:B2"/></mergeCells>"#,
        );

        let anchor = eligible(&xml, "A1");
        assert_eq!(anchor.covering_merge, None);
        assert_number(anchor, "7");

        let covered = eligible(&xml, "B1");
        assert_eq!(covered.covering_merge, Some(range("A1:B2")));
        assert!(covered.cell.is_none());

        let missing = eligible(&xml, "C3");
        assert_eq!(missing.covering_merge, None);
        assert!(missing.cell.is_none());

        let follower = eligible(&xml, "B2");
        assert_eq!(follower.covering_merge, Some(range("A1:B2")));
        assert_number(follower, "9");
    }

    #[test]
    fn streaming_0367_range_selection_remains_sparse_and_physical() {
        let xml = worksheet(
            r#"<sheetData><row r="1"><c r="A1"><v>1</v></c><c r="E1"><v>5</v></c></row><row r="2"><c r="C2"><v>3</v></c></row></sheetData><mergeCells count="2"><mergeCell ref="A1:C2"/><mergeCell ref="E1:F1"/></mergeCells>"#,
        );
        let selected = eligible_range(&xml, "A1:F2");

        assert_eq!(
            selected
                .cells
                .iter()
                .map(|record| record.address)
                .collect::<Vec<_>>(),
            vec![address("A1"), address("E1"), address("C2")]
        );
        assert!(matches!(
            selected.cells[0].cell.as_ref(),
            Some(Cell::Value(Value::Number(value))) if value.as_str() == "1"
        ));
        assert!(matches!(
            selected.cells[1].cell.as_ref(),
            Some(Cell::Value(Value::Number(value))) if value.as_str() == "5"
        ));
        assert!(matches!(
            selected.cells[2].cell.as_ref(),
            Some(Cell::Value(Value::Number(value))) if value.as_str() == "3"
        ));
    }

    #[test]
    fn streaming_0367_merge_count_is_optional_exact_and_nonempty_but_not_mismatched() {
        let optional =
            worksheet(r#"<sheetData/><mergeCells><mergeCell ref="A1:B1"/></mergeCells>"#);
        assert!(matches!(
            scan_xml(&optional, "A1"),
            Ok(ScanOutcome::Eligible(_))
        ));

        let exact = worksheet(
            r#"<sheetData/><mergeCells count="2"><mergeCell ref="A1:B1"/><mergeCell ref="C1:D1"/></mergeCells>"#,
        );
        assert!(matches!(
            scan_xml(&exact, "A1"),
            Ok(ScanOutcome::Eligible(_))
        ));

        let empty = worksheet(r#"<sheetData/><mergeCells count="0"/>"#);
        assert!(scan_xml(&empty, "A1").is_err());

        let mismatch =
            worksheet(r#"<sheetData/><mergeCells count="2"><mergeCell ref="A1:B1"/></mergeCells>"#);
        assert!(scan_xml(&mismatch, "A1").is_err());
    }

    #[test]
    fn streaming_0367_rejects_invalid_merge_references_and_placement() {
        let cases = [
            worksheet(
                r#"<sheetData/><mergeCells><mergeCell ref="A1:B1"/></mergeCells><mergeCells><mergeCell ref="C1:D1"/></mergeCells>"#,
            ),
            worksheet(
                r#"<sheetData/><mergeCells><mergeCell ref="A1:B1"/><mergeCell ref="A1:B1"/></mergeCells>"#,
            ),
            worksheet(
                r#"<sheetData/><mergeCells><mergeCell ref="A1:B2"/><mergeCell ref="B2:C3"/></mergeCells>"#,
            ),
            worksheet(r#"<sheetData/><mergeCells><mergeCell ref="A1"/></mergeCells>"#),
            worksheet(r#"<mergeCells><mergeCell ref="A1:B1"/></mergeCells><sheetData/>"#),
            worksheet(
                r#"<sheetData/><hyperlinks/><mergeCells><mergeCell ref="A1:B1"/></mergeCells>"#,
            ),
            worksheet(r#"<sheetData/><mergeCells><mergeCell ref="XFE1:XFF1"/></mergeCells>"#),
            worksheet(r#"<sheetData/><mergeCells><mergeCell/></mergeCells>"#),
        ];

        for (index, xml) in cases.iter().enumerate() {
            assert!(
                scan_xml(xml, "A1").is_err(),
                "accepted invalid merge case {index}"
            );
        }
    }

    #[test]
    fn streaming_0367_merge_nested_and_payload_markup_falls_back() {
        let nested = worksheet(
            r#"<sheetData/><mergeCells><mergeCell ref="A1:B1"><mergeCell ref="C1:D1"/></mergeCell></mergeCells>"#,
        );
        assert_merge_fallback(&nested, "A1");

        let payload = worksheet(
            r#"<sheetData/><mergeCells><mergeCell ref="A1:B1">payload</mergeCell></mergeCells>"#,
        );
        assert_merge_fallback(&payload, "A1");
    }

    #[test]
    fn streaming_0367_unknown_merge_attributes_and_children_fall_back() {
        let attribute = format!(
            r#"<worksheet xmlns="{SPREADSHEETML}" xmlns:z="urn:future"><sheetData/><mergeCells z:opaque="yes"><mergeCell ref="A1:B1"/></mergeCells></worksheet>"#
        );
        assert_merge_fallback(&attribute, "A1");

        let child = format!(
            r#"<worksheet xmlns="{SPREADSHEETML}" xmlns:z="urn:future"><sheetData/><mergeCells><mergeCell ref="A1:B1"/><z:future/></mergeCells></worksheet>"#
        );
        assert_merge_fallback(&child, "A1");
    }

    #[test]
    fn streaming_0367_late_malformed_tail_is_not_published_as_eligible() {
        let xml = format!(
            r#"{}<tail>"#,
            worksheet(
                r#"<sheetData><row r="1"><c r="A1"><v>7</v></c></row></sheetData><mergeCells><mergeCell ref="A1:B1"/></mergeCells>"#,
            )
        );
        assert_stream_xml_error(scan_xml(&xml, "A1"));
    }

    #[test]
    fn streaming_0367_merge_cap_is_bounded_without_area_expansion() {
        for (count, expected_eligible) in [
            (STREAMING_MERGE_CAP, true),
            (STREAMING_MERGE_CAP + 1, false),
        ] {
            let xml = cap_matrix(count).expect("bounded merge matrix");
            match scan_xml(&xml, "A1").expect("merge cap stream") {
                ScanOutcome::Eligible(selected) if expected_eligible => {
                    assert_eq!(selected.covering_merge, None);
                    assert_number(selected, "1");
                },
                ScanOutcome::NotEligible(NotEligibleReason::MergeSemantics)
                    if !expected_eligible => {},
                other => panic!("unexpected result at merge count {count}: {other:?}"),
            }
        }
    }
}

#[cfg(test)]
mod streaming_0366_general_reference_tests {
    use std::fmt::Debug;
    use std::io::Cursor;

    use litchi_ooxml_common::mce::{Capabilities, Error as MceError, StreamError, StreamLimits};
    use litchi_sheet::{Cell as Address, Rect};

    use super::super::selected::{
        NotEligibleReason, RangeScanOutcome, ScanOutcome, SelectedCell, scan, scan_range,
    };
    use crate::cell::{Cell, Value};
    use crate::formula::Cache;

    const SPREADSHEETML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    type StreamResult<T> = super::super::selected::StreamResult<T>;

    fn address(reference: &str) -> Address {
        Address::from_a1(reference).expect("valid worksheet address")
    }

    fn worksheet(body: &str) -> String {
        format!(r#"<worksheet xmlns="{SPREADSHEETML}">{body}</worksheet>"#)
    }

    fn value_worksheet(value: &str) -> String {
        worksheet(&format!(
            r#"<sheetData><row r="1"><c r="A1" t="str"><v>{value}</v></c></row></sheetData>"#
        ))
    }

    fn unselected_value_worksheet(value: &str) -> String {
        worksheet(&format!(
            r#"<sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1" t="str"><v>{value}</v></c></row></sheetData>"#
        ))
    }

    fn scan_xml(xml: &str, target: &str) -> StreamResult<ScanOutcome> {
        let mut input = Cursor::new(xml.as_bytes());
        scan(
            &mut input,
            &Capabilities::default(),
            &StreamLimits::default(),
            address(target),
        )
    }

    fn scan_range_xml(xml: &str, reference: &str) -> StreamResult<RangeScanOutcome> {
        let mut input = Cursor::new(xml.as_bytes());
        scan_range(
            &mut input,
            &Capabilities::default(),
            &StreamLimits::default(),
            Rect::from_a1(reference).expect("valid worksheet range"),
        )
    }

    fn eligible(xml: &str, target: &str) -> SelectedCell {
        match scan_xml(xml, target) {
            Ok(ScanOutcome::Eligible(selected)) => selected,
            other => panic!("expected eligible selection, got {other:?}"),
        }
    }

    fn assert_text(selected: SelectedCell, expected: &str) {
        match selected.cell {
            Some(Cell::Value(Value::Text(value))) => assert_eq!(value.as_str(), expected),
            other => panic!("expected text {expected:?}, got {other:?}"),
        }
    }

    fn assert_formula(selected: SelectedCell, expected_formula: &str, expected_cached: &str) {
        let Some(Cell::Formula(formula)) = selected.cell else {
            panic!("expected formula cell")
        };
        assert_eq!(formula.text(), expected_formula);
        assert!(matches!(
            formula.cached().map(Cache::value),
            Some(Value::Text(value)) if value.as_str() == expected_cached
        ));
    }

    fn assert_not_eligible(xml: &str, target: &str, expected: NotEligibleReason) {
        match scan_xml(xml, target) {
            Ok(ScanOutcome::NotEligible(actual)) => assert_eq!(actual, expected),
            other => panic!("expected {expected:?}, got {other:?}"),
        }
    }

    fn assert_not_eligible_range(xml: &str, reference: &str, expected: NotEligibleReason) {
        match scan_range_xml(xml, reference) {
            Ok(RangeScanOutcome::NotEligible(actual)) => assert_eq!(actual, expected),
            other => panic!("expected {expected:?}, got {other:?}"),
        }
    }

    fn assert_one_cell_parity(xml: &str, target: &str) {
        let range = match scan_range_xml(xml, target).expect("one-cell range scan") {
            RangeScanOutcome::Eligible(selected) => selected,
            other => panic!("expected eligible one-cell range, got {other:?}"),
        };
        let single = match scan_xml(xml, target).expect("single-cell scan") {
            ScanOutcome::Eligible(selected) => selected,
            other => panic!("expected eligible single-cell scan, got {other:?}"),
        };

        assert_eq!(range.dependencies, single.dependencies);
        match range.cells.as_slice() {
            [] => {
                assert!(single.cell.is_none());
                assert_eq!(single.dependencies.target_shared_string_index, None);
            },
            [record] => {
                assert_eq!(record.address, single.address);
                assert_eq!(
                    record.shared_string_index,
                    single.dependencies.target_shared_string_index
                );
                assert_eq!(record.cell.as_ref(), single.cell.as_ref());
            },
            other => panic!("one-cell range retained multiple records: {other:?}"),
        }
    }

    fn assert_mce_xml_error<T: Debug>(result: StreamResult<T>) {
        match result {
            Err(StreamError::Mce {
                error: MceError::Xml(_),
                ..
            }) => {},
            other => panic!("expected typed XML stream error, got {other:?}"),
        }
    }

    fn assert_mce_nonconformant<T: Debug>(result: StreamResult<T>) {
        match result {
            Err(StreamError::Mce {
                error: MceError::NonConformant(_),
                ..
            }) => {},
            other => panic!("expected typed MCE conformance error, got {other:?}"),
        }
    }

    fn assert_mce_error<T: Debug>(result: StreamResult<T>) {
        match result {
            Err(StreamError::Mce {
                error: MceError::Xml(_) | MceError::NonConformant(_),
                ..
            }) => {},
            other => panic!("expected typed MCE stream error, got {other:?}"),
        }
    }

    #[test]
    fn streaming_0366_accepts_predefined_and_numeric_references_in_scalar_payloads() {
        let xml = worksheet(
            r#"<sheetData><row r="1">
                <c r="A1" t="str"><v>value &amp; &lt; &gt; &apos; &quot;</v></c>
                <c r="B1" t="str"><f>CONCAT(&quot;A&quot;,&quot;B&quot;)</f><v>cached &amp; value</v></c>
                <c r="C1" t="inlineStr"><is><t>inline &amp; &lt; &gt; &apos; &quot;</t></is></c>
                <c r="D1" t="str"><v>decimal &#65; hex &#x42;</v></c>
                <c r="E1" t="str"><f>CONCAT(&#34;A&#34;,&#x22;B&#x22;)</f><v>cached &#67; &#x44;</v></c>
            </row></sheetData>"#,
        );

        assert_text(eligible(&xml, "A1"), "value & < > ' \"");
        assert_formula(
            eligible(&xml, "B1"),
            "CONCAT(\"A\",\"B\")",
            "cached & value",
        );
        assert_text(eligible(&xml, "C1"), "inline & < > ' \"");
        assert_text(eligible(&xml, "D1"), "decimal A hex B");
        assert_formula(eligible(&xml, "E1"), "CONCAT(\"A\",\"B\")", "cached C D");

        for target in ["A1", "B1", "C1", "D1", "E1", "F1"] {
            assert_one_cell_parity(&xml, target);
        }
    }

    #[test]
    fn streaming_0366_bounds_valid_general_reference_tokens() {
        let twelve = value_worksheet("prefix&#x00000041;suffix");
        assert_text(eligible(&twelve, "A1"), "prefixAsuffix");

        let thirteen = value_worksheet("prefix&#x000000041;suffix");
        assert_not_eligible(&thirteen, "A1", NotEligibleReason::GeneralReference);
        assert_not_eligible_range(&thirteen, "A1", NotEligibleReason::GeneralReference);

        let xml_illegal = value_worksheet("prefix&#x1;suffix");
        assert_not_eligible(&xml_illegal, "A1", NotEligibleReason::GeneralReference);

        let xml_legal = value_worksheet("prefix&#x20;suffix");
        assert_text(eligible(&xml_legal, "A1"), "prefix suffix");
    }

    #[test]
    fn streaming_0366_rejects_numeric_inline_references_as_not_eligible() {
        let xml = worksheet(
            r#"<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>inline &#65;</t></is></c></row></sheetData>"#,
        );
        assert_not_eligible(&xml, "A1", NotEligibleReason::GeneralReference);
        assert_not_eligible_range(&xml, "A1", NotEligibleReason::GeneralReference);
    }

    #[test]
    fn streaming_0366_keeps_invalid_references_as_typed_errors_without_early_publication() {
        assert_mce_xml_error(scan_xml(&value_worksheet("bad &#xZZ;"), "A1"));
        assert_mce_nonconformant(scan_xml(&value_worksheet("bad &custom;"), "A1"));
        assert_mce_xml_error(scan_xml(&value_worksheet("bad &#x110000;"), "A1"));

        assert_mce_xml_error(scan_xml(&unselected_value_worksheet("bad &#xZZ;"), "A1"));
        assert_mce_nonconformant(scan_xml(&unselected_value_worksheet("bad &custom;"), "A1"));
        assert_mce_xml_error(scan_xml(
            &unselected_value_worksheet("bad &#x110000;"),
            "A1",
        ));

        let late_tail = worksheet(
            r#"<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData><future>&#xZZ;</future>"#,
        );
        assert_mce_error(scan_xml(&late_tail, "A1"));
    }
}

#[cfg(test)]
mod streaming_0400_numeric_scratch_tests {
    use std::fmt::Debug;
    use std::io::Cursor;

    use litchi_ooxml_common::mce::{Capabilities, Error as MceError, StreamError, StreamLimits};
    use litchi_sheet::{Cell as Address, Rect};

    use super::super::model::MAX_CELL_CHARACTERS;
    use super::super::selected::{
        NotEligibleReason, RangeScanOutcome, ScanOutcome, SelectedCell, scan, scan_range,
    };
    use crate::cell::{Cell, Number, Value};
    use crate::formula::Cache;

    const SPREADSHEETML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    const STRICT_SPREADSHEETML: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";

    type StreamResult<T> = super::super::selected::StreamResult<T>;

    fn address(reference: &str) -> Address {
        Address::from_a1(reference).expect("valid worksheet address")
    }

    fn worksheet(body: &str) -> String {
        format!(r#"<worksheet xmlns="{SPREADSHEETML}">{body}</worksheet>"#)
    }

    fn scan_xml(xml: &str, target: &str) -> StreamResult<ScanOutcome> {
        let mut input = Cursor::new(xml.as_bytes());
        scan(
            &mut input,
            &Capabilities::default(),
            &StreamLimits::default(),
            address(target),
        )
    }

    fn scan_range_xml(xml: &str, reference: &str) -> StreamResult<RangeScanOutcome> {
        let mut input = Cursor::new(xml.as_bytes());
        scan_range(
            &mut input,
            &Capabilities::default(),
            &StreamLimits::default(),
            Rect::from_a1(reference).expect("valid worksheet range"),
        )
    }

    fn eligible(xml: &str, target: &str) -> SelectedCell {
        match scan_xml(xml, target) {
            Ok(ScanOutcome::Eligible(selected)) => selected,
            other => panic!("expected eligible selection, got {other:?}"),
        }
    }

    fn assert_number(selected: SelectedCell, expected: &str) {
        match selected.cell {
            Some(Cell::Value(Value::Number(value))) => assert_eq!(value.as_str(), expected),
            other => panic!("expected number {expected}, got {other:?}"),
        }
    }

    fn assert_stream_xml_error<T: Debug>(result: StreamResult<T>) {
        match result {
            Err(StreamError::Mce {
                error: MceError::Xml(_) | MceError::NonConformant(_),
                ..
            }) => {},
            other => panic!("expected typed XML stream error, got {other:?}"),
        }
    }

    fn assert_not_eligible(xml: &str, expected: NotEligibleReason) {
        match scan_xml(xml, "A1") {
            Ok(ScanOutcome::NotEligible(actual)) => assert_eq!(actual, expected),
            other => panic!("expected {expected:?}, got {other:?}"),
        }
    }

    fn assert_number_error(xml: &str, expected: &str) {
        let error = scan_xml(xml, "A1").expect_err("invalid worksheet number");
        assert!(
            error.to_string().contains(expected),
            "expected error containing {expected:?}, got {error}"
        );
    }

    fn assert_unselected_inline_error_matches_selected(inline: &str, expected: Option<&str>) {
        let selected = worksheet(&format!(
            r#"<sheetData><row r="1"><c r="A1">{inline}</c></row></sheetData>"#
        ));
        let unselected = worksheet(&format!(
            r#"<sheetData><row r="1"><c r="A1"><v>7</v></c><c r="B1">{inline}</c></row></sheetData>"#
        ));
        let selected_error = scan_xml(&selected, "A1")
            .expect_err("selected inline payload must fail")
            .to_string();
        let unselected_error = scan_xml(&unselected, "A1")
            .expect_err("unselected inline payload must fail")
            .to_string();
        assert_eq!(unselected_error, selected_error);
        if let Some(expected) = expected {
            assert!(
                unselected_error.contains(expected),
                "expected error containing {expected:?}, got {unselected_error}"
            );
        }
    }

    #[test]
    fn streaming_0400_numeric_scratch_preserves_long_to_short_and_exponent_lexemes() {
        let long = "1234567890123456789012345678901234567890.123456789";
        let xml = worksheet(&format!(
            r#"<sheetData><row r="1"><c r="A1"><v>{long}</v></c><c r="B1"><v>-0.000</v></c><c r="C1" t="n"><v>6.02E+23</v></c></row></sheetData>"#
        ));

        assert_number(eligible(&xml, "B1"), "-0.000");
        assert_number(eligible(&xml, "C1"), "6.02E+23");
    }

    #[test]
    fn streaming_0400_numeric_scratch_handles_scalar_type_transitions() {
        let xml = worksheet(
            r#"<sheetData><row r="1">
                <c r="A1"><v>1234567890123456789012345678901234567890.123456789</v></c>
                <c r="B1" t="str"><v>text</v></c>
                <c r="C1"><f>1+1</f><v>3.00e+2</v></c>
                <c r="D1" t="inlineStr"><is><t>inline</t></is></c>
                <c r="E1" t="e"><v>#N/A</v></c>
                <c r="F1"/>
            </row></sheetData>"#,
        );
        let selected = match scan_range_xml(&xml, "A1:F1").expect("scalar transition stream") {
            RangeScanOutcome::Eligible(selected) => selected,
            other => panic!("expected eligible scalar transition range, got {other:?}"),
        };

        assert_eq!(selected.cells.len(), 6);
        assert!(matches!(
            selected.cells[0].cell.as_ref(),
            Some(Cell::Value(Value::Number(value)))
                if value.as_str() == "1234567890123456789012345678901234567890.123456789"
        ));
        assert!(matches!(
            selected.cells[1].cell.as_ref(),
            Some(Cell::Value(Value::Text(value))) if value.as_str() == "text"
        ));
        match selected.cells[2].cell.as_ref() {
            Some(Cell::Formula(formula)) => {
                assert_eq!(formula.text(), "1+1");
                assert!(matches!(
                    formula.cached().map(Cache::value),
                    Some(Value::Number(value)) if value.as_str() == "3.00e+2"
                ));
            },
            other => panic!("expected cached numeric formula, got {other:?}"),
        }
        assert!(matches!(
            selected.cells[3].cell.as_ref(),
            Some(Cell::Value(Value::Text(value))) if value.as_str() == "inline"
        ));
        assert!(matches!(
            selected.cells[4].cell.as_ref(),
            Some(Cell::Value(Value::Error(value))) if value.as_str() == "#N/A"
        ));
        assert!(matches!(selected.cells[5].cell.as_ref(), Some(Cell::Empty)));
    }

    #[test]
    fn streaming_0400_numeric_scratch_decodes_entity_and_cdata_numeric_fragments() {
        let xml = worksheet(
            r#"<sheetData><row r="1"><c r="A1"><v>&#x31;<![CDATA[.25E+3]]></v></c></row></sheetData>"#,
        );
        assert_number(eligible(&xml, "A1"), "1.25E+3");
    }

    #[test]
    fn streaming_0400_numeric_scratch_rejects_oversized_selected_and_unselected_values() {
        let oversized = "7".repeat(MAX_CELL_CHARACTERS + 1);
        let selected = worksheet(&format!(
            r#"<sheetData><row r="1"><c r="A1"><v>{oversized}</v></c></row></sheetData>"#
        ));
        assert!(scan_xml(&selected, "A1").is_err());

        let unselected = worksheet(&format!(
            r#"<sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>{oversized}</v></c></row></sheetData>"#
        ));
        assert!(scan_xml(&unselected, "A1").is_err());
    }

    #[test]
    fn streaming_0400_numeric_scratch_does_not_publish_before_malformed_tail() {
        let xml = format!(
            r#"{}<tail>"#,
            worksheet(
                r#"<sheetData><row r="1"><c r="A1"><v>1234567890123456789012345678901234567890.123456789</v></c><c r="B1"><v>-0.000</v></c></row></sheetData>"#,
            )
        );
        assert_stream_xml_error(scan_xml(&xml, "B1"));
    }

    #[test]
    fn streaming_0401_numeric_elision_preserves_selected_lexemes_and_unselected_validation() {
        for (type_attribute, value) in [("", "  -0.000  "), (r#" t="n""#, "6.02E+23")] {
            let xml = worksheet(&format!(
                r#"<sheetData><row r="1"><c r="A1"><v>7</v></c><c r="B1"{type_attribute}><v>{value}</v></c></row></sheetData>"#
            ));
            assert_number(eligible(&xml, "A1"), "7");

            let selected = eligible(
                &worksheet(&format!(
                    r#"<sheetData><row r="1"><c r="A1"{type_attribute}><v>{value}</v></c></row></sheetData>"#
                )),
                "A1",
            );
            assert_number(selected, value);
        }

        let whitespace = worksheet(
            r#"<sheetData><row r="1"><c r="A1"><v>7</v></c><c r="B1"><v>   </v></c></row></sheetData>"#,
        );
        assert_number(eligible(&whitespace, "A1"), "7");
        let selected_whitespace = eligible(
            &worksheet(r#"<sheetData><row r="1"><c r="A1"><v>   </v></c></row></sheetData>"#),
            "A1",
        );
        assert!(matches!(selected_whitespace.cell, Some(Cell::Empty)));
    }

    #[test]
    fn streaming_0401_numeric_elision_keeps_selected_and_unselected_errors_exact() {
        for (value, expected) in [
            ("not-a-number", "invalid worksheet number 'not-a-number'"),
            ("NaN", "non-finite worksheet number 'NaN'"),
            ("1e999", "non-finite worksheet number '1e999'"),
        ] {
            let selected = worksheet(&format!(
                r#"<sheetData><row r="1"><c r="A1"><v>{value}</v></c></row></sheetData>"#
            ));
            assert_number_error(&selected, expected);

            let unselected = worksheet(&format!(
                r#"<sheetData><row r="1"><c r="A1"><v>7</v></c><c r="B1"><v>{value}</v></c></row></sheetData>"#
            ));
            assert_number_error(&unselected, expected);

            let error = Number::new(value)
                .expect_err("owned constructor must reject invalid value")
                .to_string();
            assert!(
                error.contains(expected),
                "expected owned error containing {expected:?}, got {error}"
            );
        }
    }

    #[test]
    fn streaming_0401_numeric_elision_keeps_formula_cache_ownership_and_parity() {
        let value = "  6.02E+23  ";
        let xml = worksheet(&format!(
            r#"<sheetData><row r="1"><c r="A1"><v>-0.000</v></c><c r="B1"><f>1+1</f><v>{value}</v></c></row></sheetData>"#
        ));
        assert_number(eligible(&xml, "A1"), "-0.000");

        let selected = eligible(&xml, "B1");
        match selected.cell {
            Some(Cell::Formula(formula)) => {
                assert_eq!(formula.text(), "1+1");
                assert!(matches!(
                    formula.cached().map(Cache::value),
                    Some(Value::Number(number)) if number.as_str() == value
                ));
            },
            other => panic!("expected cached formula, got {other:?}"),
        }
    }

    #[test]
    fn streaming_0401_numeric_elision_does_not_skip_unselected_inline_validation() {
        let valid = worksheet(
            r#"<sheetData><row r="1"><c r="A1"><v>7</v></c><c r="B1"><is><t>unselected inline</t></is></c></row></sheetData>"#,
        );
        assert_number(eligible(&valid, "A1"), "7");

        assert_unselected_inline_error_matches_selected(r#"<is><t>broken</is>"#, None);
        let overlong = "a".repeat(MAX_CELL_CHARACTERS + 1);
        assert_unselected_inline_error_matches_selected(
            &format!(r#"<is><t>{overlong}</t></is>"#),
            Some("inline string exceeds"),
        );
        assert_unselected_inline_error_matches_selected(
            r#"<is><t>_xD83D_</t></is>"#,
            Some("unpaired high surrogate"),
        );
    }

    #[test]
    fn streaming_0400_dimension_keeps_scalar_and_range_queries_eligible() {
        let empty = worksheet(
            r#"<dimension ref="$B$2:F9"/><sheetData><row r="2"><c r="B2"><v>-0.000</v></c><c r="F2"><v>6.02E+23</v></c></row></sheetData>"#,
        );
        assert_number(eligible(&empty, "B2"), "-0.000");
        let selected = match scan_range_xml(&empty, "A1:F9").expect("dimension range scan") {
            RangeScanOutcome::Eligible(selected) => selected,
            other => panic!("expected eligible dimension range, got {other:?}"),
        };
        assert_eq!(selected.cells.len(), 2);
        assert_eq!(selected.cells[0].address.a1(), "B2");
        assert_eq!(selected.cells[1].address.a1(), "F2");

        let explicit = worksheet(
            r#"<dimension ref="A1:C3">
            </dimension><sheetData><row r="1"><c r="A1"><v>7</v></c></row></sheetData>"#,
        );
        assert_number(eligible(&explicit, "A1"), "7");
    }

    #[test]
    fn streaming_0400_dimension_supports_strict_prefixed_and_ignorable_metadata() {
        let strict = format!(
            r#"<s:worksheet xmlns:s="{STRICT_SPREADSHEETML}"><s:dimension ref="A1:C3"/><s:sheetData><s:row r="1"><s:c r="A1"><s:v>7</s:v></s:c></s:row></s:sheetData></s:worksheet>"#,
        );
        assert_number(eligible(&strict, "A1"), "7");

        let ignorable = format!(
            r#"<worksheet xmlns="{SPREADSHEETML}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:future" mc:Ignorable="x"><dimension ref="A1" x:hint="ignored"/><sheetData><row r="1"><c r="A1"><v>8</v></c></row></sheetData></worksheet>"#,
        );
        assert_number(eligible(&ignorable, "A1"), "8");
    }

    #[test]
    fn streaming_0400_dimension_rejects_malformed_duplicate_and_late_metadata() {
        for xml in [
            worksheet(r#"<dimension/><sheetData/>"#),
            worksheet(r#"<dimension ref="A0"/><sheetData/>"#),
            worksheet(r#"<dimension ref="B2:A1"/><sheetData/>"#),
            worksheet(r#"<dimension ref="XFE1"/><sheetData/>"#),
            worksheet(r#"<dimension ref="A1"/><dimension ref="B2"/><sheetData/>"#),
            worksheet(r#"<sheetData/><dimension ref="A1"/>"#),
        ] {
            assert!(
                scan_xml(&xml, "A1").is_err(),
                "accepted invalid dimension: {xml}"
            );
        }
    }

    #[test]
    fn streaming_0400_dimension_content_and_unknown_attributes_fall_back() {
        let attribute = worksheet(r#"<dimension ref="A1" future="keep"/><sheetData/>"#);
        assert_not_eligible(&attribute, NotEligibleReason::UnsupportedStructure);

        let nested = worksheet(
            r#"<dimension ref="A1"><future/></dimension><sheetData><row r="1"><c r="A1"><v>7</v></c></row></sheetData>"#,
        );
        assert_not_eligible(&nested, NotEligibleReason::UnsupportedStructure);

        let attribute_and_nested = worksheet(
            r#"<dimension ref="A1" future="keep"><future>payload</future></dimension><sheetData/>"#,
        );
        assert_not_eligible(
            &attribute_and_nested,
            NotEligibleReason::UnsupportedStructure,
        );

        let text = worksheet(r#"<dimension ref="A1">payload</dimension><sheetData/>"#);
        assert_not_eligible(&text, NotEligibleReason::UnsupportedStructure);
    }
}
