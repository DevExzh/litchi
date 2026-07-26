use litchi_odf::{CellHyperlink, CellValue, FlatSpreadsheet};

const RICH_TEXT_FODS: &str = include_str!("../../../test-data/odf/ods/rich-text-hyperlinks.fods");

fn open_fixture() -> FlatSpreadsheet {
    FlatSpreadsheet::from_bytes(RICH_TEXT_FODS.as_bytes().to_vec()).unwrap()
}

#[test]
fn rich_cell_content_survives_flat_spreadsheet_rewrite() {
    let document = open_fixture();
    let mut mutable = document.into_mutable().unwrap();
    let first = &mutable.spreadsheet().sheets()[0].rows[0].cells[0];
    assert_eq!(first.text, "Pre styled link tail  XQ\nsecond\tline\nend");
    assert_eq!(first.hyperlinks().len(), 1);
    assert_eq!(first.hyperlinks()[0].range(), 11..15);
    assert_eq!(
        first
            .rich_text()
            .expect("mixed content must be retained")
            .paragraphs()
            .len(),
        2
    );
    assert_eq!(
        first.rich_text().unwrap().namespaces().get("ext"),
        Some(&"urn:litchi:test:rich-text".to_string())
    );
    let expected_text = first.text.clone();
    let expected_links = first.hyperlinks().to_vec();

    mutable
        .spreadsheet_mut()
        .set_cell(0, 1, 0, CellValue::Number(42.0))
        .unwrap();
    let bytes = mutable.to_bytes().unwrap();
    let xml = String::from_utf8(bytes.clone()).unwrap();
    assert!(xml.contains(r#"<text:span text:style-name="Bold">"#));
    assert!(xml.contains(r#"<tx:span tx:style-name="Italic">link</tx:span>"#));
    assert!(xml.contains(r#"xmlns:tx="urn:oasis:names:tc:opendocument:xmlns:text:1.0""#));
    assert!(xml.contains(r#"<text:s text:c="2"/>"#));
    assert!(xml.contains("<t2:tab/>"));
    assert!(xml.contains("<t2:line-break/>"));
    assert!(xml.contains(r#"xmlns:t2="urn:oasis:names:tc:opendocument:xmlns:text:1.0""#));
    assert!(xml.contains(r#"xmlns:ext="urn:litchi:test:rich-text""#));
    assert!(xml.contains(r#"<ext:token ext:kind="field">X</ext:token>"#));
    assert!(xml.contains("<ext:s>Q</ext:s>"));

    let mut reopened = FlatSpreadsheet::from_bytes(bytes).unwrap();
    let sheets = reopened.spreadsheet_mut().sheets().unwrap();
    let cell = &sheets[0].rows[0].cells[0];
    assert_eq!(cell.text, expected_text);
    assert_eq!(cell.hyperlinks(), expected_links);
    assert_eq!(cell.rich_text().unwrap().plain_text(), expected_text);
    assert_eq!(cell.rich_text().unwrap().paragraphs().len(), 2);
    assert_eq!(
        cell.rich_text().unwrap().namespaces().get("ext"),
        Some(&"urn:litchi:test:rich-text".to_string())
    );
}

#[test]
fn hyperlink_mutations_preserve_nested_span_styling() {
    let document = open_fixture();
    let mut mutable = document.into_mutable().unwrap();
    let cell = &mut mutable.spreadsheet_mut().sheets_mut()[0].rows[0].cells[1];
    assert_eq!(cell.text, "AboldmixtailZ");
    cell.add_hyperlink(
        3..10,
        CellHyperlink::with_text("https://range.example/", "ldmixta").unwrap(),
    )
    .unwrap();

    let bytes = mutable.to_bytes().unwrap();
    let xml = String::from_utf8(bytes.clone()).unwrap();
    assert!(xml.contains(
        r#"<text:a xlink:href="https://range.example/" xlink:type="simple"><text:span text:style-name="Bold">ld<text:span text:style-name="Italic">mix</text:span>ta</text:span></text:a>"#
    ));

    let reopened = FlatSpreadsheet::from_bytes(bytes).unwrap();
    let mut reopened = reopened.into_mutable().unwrap();
    let cell = &mut reopened.spreadsheet_mut().sheets_mut()[0].rows[0].cells[1];
    assert_eq!(cell.hyperlinks()[0].range(), 3..10);
    assert_eq!(cell.remove_hyperlink(0).unwrap().text(), "ldmixta");
    assert_eq!(cell.text, "AboldmixtailZ");

    let xml = String::from_utf8(reopened.to_bytes().unwrap()).unwrap();
    assert!(!xml.contains("https://range.example/"));
    assert!(xml.matches(r#"text:style-name="Bold""#).count() >= 1);
    assert!(xml.contains(r#"text:style-name="Italic">mix</text:span>"#));
}

#[test]
fn invalid_structural_ranges_fail_without_mutating_the_cell() {
    let document = open_fixture();
    let mut mutable = document.into_mutable().unwrap();
    let cell = &mut mutable.spreadsheet_mut().sheets_mut()[0].rows[0].cells[0];
    let original_links = cell.hyperlinks().to_vec();
    let original_rich_text = cell.rich_text().cloned();

    let paragraph_boundary = cell.text.find('\n').unwrap();
    assert!(
        cell.add_hyperlink(
            paragraph_boundary - 1..paragraph_boundary + 2,
            CellHyperlink::with_text(
                "https://cross.example/",
                &cell.text[paragraph_boundary - 1..paragraph_boundary + 2],
            )
            .unwrap(),
        )
        .is_err()
    );

    let spaces = cell.text.find("  ").unwrap();
    assert!(
        cell.add_hyperlink(
            spaces + 1..spaces + 2,
            CellHyperlink::with_text("https://space.example/", " ").unwrap(),
        )
        .is_err()
    );
    assert_eq!(cell.hyperlinks(), original_links);
    assert_eq!(cell.rich_text(), original_rich_text.as_ref());
}
