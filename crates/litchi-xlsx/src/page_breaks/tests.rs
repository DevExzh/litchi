#![allow(
    clippy::unwrap_used,
    reason = "focused unit tests use panic-on-failure assertions"
)]

use super::*;

const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

fn worksheet(body: &str) -> Vec<u8> {
    format!(r#"<worksheet xmlns="{SML}">{body}</worksheet>"#).into_bytes()
}

#[test]
fn parses_defaults_and_manual_counts() {
    let xml = worksheet(concat!(
        r#"<rowBreaks count="2" manualBreakCount="1"><brk/><brk id="8" max="16383" man="1"/></rowBreaks>"#,
        r#"<colBreaks count="1"><brk id="3" max="1048575" pt="true"/></colBreaks>"#,
    ));
    let value = parse(&xml).unwrap();
    let horizontal = value.horizontal().unwrap();
    assert_eq!(horizontal.len(), 2);
    assert_eq!(horizontal.manual_break_count(), 1);
    assert_eq!(horizontal.breaks()[1].id(), 8);
    assert!(horizontal.breaks()[1].is_manual());
    assert!(value.vertical().unwrap().breaks()[0].is_pivot());
}

#[test]
fn writer_is_byte_minimal_and_round_trips() {
    let mut value = PageBreaks::new();
    value
        .set_horizontal(
            Collection::horizontal([
                Break::new(0, 0, 0).unwrap(),
                Break::new(8, 0, 16_383).unwrap().with_manual(true),
            ])
            .unwrap(),
        )
        .unwrap();
    value
        .set_vertical(
            Collection::vertical([Break::new(3, 0, 1_048_575).unwrap().with_pivot(true)]).unwrap(),
        )
        .unwrap();
    let fragment = write(&value).unwrap();
    assert_eq!(
        fragment,
        br#"<rowBreaks count="2" manualBreakCount="1"><brk/><brk id="8" max="16383" man="1"/></rowBreaks><colBreaks count="1"><brk id="3" max="1048575" pt="1"/></colBreaks>"#,
    );
    let parsed = parse(&worksheet(std::str::from_utf8(&fragment).unwrap())).unwrap();
    assert_eq!(parsed, value);
}

#[test]
fn rewrite_preserves_unrelated_bytes_and_schema_order() {
    let source = worksheet("<sheetData/><headerFooter/><customProperties/>");
    let mut value = PageBreaks::new();
    value
        .set_horizontal(
            Collection::horizontal([Break::new(9, 0, 16_383).unwrap().with_manual(true)]).unwrap(),
        )
        .unwrap();
    let output = replace(&source, &value).unwrap();
    assert_eq!(
        output,
        worksheet(concat!(
            "<sheetData/><headerFooter/>",
            r#"<rowBreaks count="1" manualBreakCount="1"><brk id="9" max="16383" man="1"/></rowBreaks>"#,
            "<customProperties/>",
        )),
    );
    assert_eq!(replace(&output, &value).unwrap(), output);
    assert_eq!(replace(&output, &PageBreaks::new()).unwrap(), source);
}

#[test]
fn changed_pretty_xml_is_compact_without_losing_text_whitespace() {
    let source = format!(
        concat!(
            r#"<worksheet xmlns="{SML}">"#,
            "\n  <sheetData><row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is>",
            r#"<t xml:space="preserve"> </t>"#,
            "</is></c></row></sheetData>\n",
            "</worksheet>"
        ),
        SML = SML,
    );
    let mut value = PageBreaks::new();
    value
        .set_horizontal(Collection::horizontal([Break::new(1, 0, 1).unwrap()]).unwrap())
        .unwrap();
    let output = replace(source.as_bytes(), &value).unwrap();
    assert!(!output.windows(2).any(|window| window == b">\n"));
    assert!(
        output
            .windows(b"> </t>".len())
            .any(|window| window == b"> </t>")
    );
}

#[test]
fn rejects_count_bounds_grid_and_unsafe_xml() {
    for body in [
        r#"<rowBreaks count="2"><brk/></rowBreaks>"#,
        r#"<rowBreaks count="1" manualBreakCount="1"><brk/></rowBreaks>"#,
        r#"<rowBreaks count="1"><brk id="1048576"/></rowBreaks>"#,
        r#"<colBreaks count="1"><brk id="16384"/></colBreaks>"#,
        r#"<rowBreaks count="1"><brk min="2" max="1"/></rowBreaks>"#,
        r"<rowBreaks><child/></rowBreaks>",
    ] {
        assert!(parse(&worksheet(body)).is_err(), "accepted {body}");
    }
    assert!(parse(br"<!DOCTYPE x><worksheet/>").is_err());
    assert!(Collection::horizontal((0..=MAX_HORIZONTAL_BREAKS).map(|_| Break::default())).is_err());
}

#[test]
fn supports_strict_and_reads_but_refuses_mce_projected_breaks() {
    let strict = br#"<worksheet xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main"><rowBreaks count="1"><brk id="4"/></rowBreaks></worksheet>"#;
    assert_eq!(
        parse(strict).unwrap().horizontal().unwrap().breaks()[0].id(),
        4
    );

    let mce = format!(
        concat!(
            r#"<worksheet xmlns="{SML}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:future">"#,
            r#"<mc:AlternateContent><mc:Choice Requires="x"><x:future/></mc:Choice><mc:Fallback>"#,
            r#"<rowBreaks count="1"><brk id="6"/></rowBreaks>"#,
            r#"</mc:Fallback></mc:AlternateContent></worksheet>"#,
        ),
        SML = SML,
    );
    let parsed = parse(mce.as_bytes()).unwrap();
    assert_eq!(parsed.horizontal().unwrap().breaks()[0].id(), 6);
    assert_eq!(replace(mce.as_bytes(), &parsed).unwrap(), mce.as_bytes());
    assert!(replace(mce.as_bytes(), &PageBreaks::new()).is_err());
}

#[test]
fn changed_prefixed_worksheet_uses_the_root_namespace_prefix() {
    let source = br#"<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><x:sheetData/><x:rowBreaks count="1"><x:brk id="4"/></x:rowBreaks></x:worksheet>"#;
    let mut value = PageBreaks::new();
    value
        .set_horizontal(Collection::horizontal([Break::new(7, 0, 0).unwrap()]).unwrap())
        .unwrap();

    let output = replace(source, &value).unwrap();
    assert!(
        output
            .windows(b"<x:rowBreaks".len())
            .any(|window| { window == b"<x:rowBreaks" })
    );
    assert!(
        !output
            .windows(b"<rowBreaks".len())
            .any(|window| { window == b"<rowBreaks" })
    );
    assert_eq!(parse(&output).unwrap(), value);

    let insertion_source = br#"<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><x:sheetData/><x:customProperties/></x:worksheet>"#;
    let inserted = replace(insertion_source, &value).unwrap();
    assert!(
        inserted
            .windows(b"<x:rowBreaks".len())
            .any(|window| { window == b"<x:rowBreaks" })
    );
    assert_eq!(parse(&inserted).unwrap(), value);
}

#[test]
fn changed_collections_refuse_lossy_wire_metadata() {
    let sources = [
        worksheet(
            r#"<rowBreaks count="1">
                <!-- retained producer note -->
                <brk id="4"/>
            </rowBreaks>"#,
        ),
        worksheet(r#"<rowBreaks count="1" xmlns:future="urn:future"><brk id="4"/></rowBreaks>"#),
        worksheet(
            r#"<rowBreaks count="1"><brk id="4" future:keep="yes" xmlns:future="urn:future"/></rowBreaks>"#,
        ),
        worksheet(r#"<rowBreaks count="1">&#x20;<brk id="4"/></rowBreaks>"#),
    ];
    let mut value = PageBreaks::new();
    value
        .set_horizontal(Collection::horizontal([Break::new(5, 0, 0).unwrap()]).unwrap())
        .unwrap();

    for source in sources {
        assert!(replace(&source, &value).is_err());
        assert_eq!(
            parse(&source).unwrap().horizontal().unwrap().breaks()[0].id(),
            4
        );
    }
}

#[test]
fn strict_and_transitional_page_break_dialects_must_match_the_root() {
    let mixed_child = br#"<worksheet xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main"><rowBreaks xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1"><brk id="4"/></rowBreaks></worksheet>"#;
    let mixed_break = br#"<worksheet xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main"><rowBreaks count="1"><brk xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="4"/></rowBreaks></worksheet>"#;
    assert!(parse(mixed_child).is_err());
    assert!(parse(mixed_break).is_err());
}

#[test]
fn undeclared_element_and_attribute_prefixes_are_rejected() {
    let undeclared_element = br#"<x:worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><rowBreaks count="0"/></x:worksheet>"#;
    let undeclared_attribute = br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><rowBreaks count="1"><brk id="4" bad:keep="yes"/></rowBreaks></worksheet>"#;
    assert!(parse(undeclared_element).is_err());
    assert!(parse(undeclared_attribute).is_err());
}

#[test]
fn package_transaction_reads_back_and_reverses_page_breaks() {
    let mut package = crate::Package::create().unwrap();
    let mut transaction = package.edit_page_breaks("Sheet1").unwrap();
    let collection = Collection::horizontal([Break::new(7, 0, 0).unwrap()]).unwrap();
    assert!(transaction.set_horizontal(collection.clone()).unwrap());
    let commit = transaction.commit().unwrap();

    assert_eq!(
        commit.snapshot().page_breaks().horizontal(),
        Some(&collection)
    );
    assert_eq!(
        package
            .page_breaks("Sheet1")
            .unwrap()
            .page_breaks()
            .horizontal(),
        Some(&collection)
    );

    package
        .apply_page_breaks_patch(&commit.patch().inverse())
        .unwrap();
    assert_eq!(
        package.page_breaks("Sheet1").unwrap().page_breaks(),
        &PageBreaks::new()
    );
}
