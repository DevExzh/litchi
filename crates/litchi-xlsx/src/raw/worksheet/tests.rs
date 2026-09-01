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
                    r#"<mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells><sheetData/>"#,
                ),
                NotEligibleReason::MergeSemantics,
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
