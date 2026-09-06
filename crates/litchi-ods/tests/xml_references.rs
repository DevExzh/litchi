#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "These integration fixtures are fixed and assertions are the test oracle."
)]

//! Regression coverage for XML general references in ODS cell text.
//!
//! `quick_xml` exposes predefined and numeric references as separate
//! `Event::GeneralRef` events.  The standalone worksheet parser used by
//! `Builder::sheets` and the fused parser used by `Spreadsheet::from_bytes`
//! must append those events exactly once and reject unresolved or illegal
//! references consistently.

use litchi_odf_common::{constants, core::PackageWriter};
use litchi_ods::{Builder, Spreadsheet};

const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TABLE: &str = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";

// The literal CR is normalized to LF by XML 1.0 text decoding.  The numeric
// &#13; remains CR because it is a separate general-reference event.  The
// nested &amp;lt; case proves that reference decoding is not recursive.
const MIXED_REFERENCES: &str = concat!(
    r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content "#,
    r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
    r#"xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" "#,
    r#"xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" "#,
    r#"xmlns:xlink="http://www.w3.org/1999/xlink" office:version="1.3">"#,
    r#"<office:body><office:spreadsheet><table:table table:name="Refs"><table:table-row>"#,
    r#"<table:table-cell office:value-type="string"><text:p>ordinary"#,
    "\r",
    r#"line&#13;numeric&#10;line&#9;tab&#x1F642;|&amp; &lt; &gt; &quot; &apos;|&amp;lt;|tail</text:p></table:table-cell>"#,
    r#"<table:table-cell office:value-type="string"><text:p>before<text:a "#,
    r#"xlink:type="simple" xlink:href="https://example.test/ref">link&#13;&amp;lt;&#x1F642;</text:a>"#,
    r#"after</text:p></table:table-cell></table:table-row></table:table></office:spreadsheet>"#,
    r#"</office:body></office:document-content>"#,
);

const EXPECTED_CELL_TEXT: &str = "ordinary\nline\rnumeric\nline\ttab🙂|& < > \" '|&lt;|tail";
const EXPECTED_LINK_CELL_TEXT: &str = "beforelink\r&lt;🙂after";
const EXPECTED_LINK_TEXT: &str = "link\r&lt;🙂";

fn invalid_reference_content(reference: &str) -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="{OFFICE}" xmlns:table="{TABLE}" xmlns:text="{TEXT}" office:version="1.3"><office:body><office:spreadsheet><table:table table:name="Invalid"><table:table-row><table:table-cell office:value-type="string"><text:p>bad{reference}</text:p></table:table-cell></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>"#
    )
}

fn raw_package(content_xml: &str) -> litchi_core::Result<Vec<u8>> {
    let mut writer = PackageWriter::new();
    writer.set_mimetype(constants::ODF_SPREADSHEET)?;
    writer.add_file("content.xml", content_xml.as_bytes())?;
    writer.finish_to_bytes()
}

#[test]
fn standalone_and_fused_parsers_decode_mixed_references_once() -> litchi_core::Result<()> {
    let standalone = Builder::new().content_xml(MIXED_REFERENCES).sheets()?;
    assert_eq!(standalone.len(), 1);
    assert_eq!(standalone[0].rows.len(), 1);
    assert_eq!(standalone[0].rows[0].cells.len(), 2);
    assert_eq!(standalone[0].rows[0].cells[0].text, EXPECTED_CELL_TEXT);
    assert_eq!(standalone[0].rows[0].cells[1].text, EXPECTED_LINK_CELL_TEXT);

    let link = standalone[0].rows[0].cells[1]
        .hyperlinks
        .first()
        .expect("the reference-bearing hyperlink should be retained");
    assert_eq!(link.href(), "https://example.test/ref");
    assert_eq!(link.text(), EXPECTED_LINK_TEXT);
    let link_start = "before".len();
    assert_eq!(
        link.range(),
        link_start..link_start + EXPECTED_LINK_TEXT.len()
    );

    let bytes = Builder::new().content_xml(MIXED_REFERENCES).build()?;
    let spreadsheet = Spreadsheet::from_bytes(bytes.clone())?;
    assert_eq!(spreadsheet.sheets(), standalone.as_slice());
    assert_eq!(spreadsheet.into_bytes(), bytes);
    Ok(())
}

#[test]
fn nested_entity_reference_decodes_once_and_keeps_numeric_cr_distinct() -> litchi_core::Result<()> {
    let sheets = Builder::new().content_xml(MIXED_REFERENCES).sheets()?;
    let text = &sheets[0].rows[0].cells[0].text;

    assert_eq!(text, EXPECTED_CELL_TEXT);
    assert_eq!(
        text.as_bytes()
            .iter()
            .filter(|&&byte| byte == b'\n')
            .count(),
        2
    );
    assert_eq!(
        text.as_bytes()
            .iter()
            .filter(|&&byte| byte == b'\r')
            .count(),
        1
    );
    assert!(text.ends_with("|&lt;|tail"));
    assert!(!text.contains("|<|tail"));
    Ok(())
}

#[test]
fn unknown_and_illegal_numeric_references_are_rejected_by_both_open_paths() {
    for reference in [
        "&unknown;",
        "&#0;",
        "&#xD800;",
        "&#x110000;",
        "&#+65;",
        "&#x+41;",
    ] {
        let content = invalid_reference_content(reference);
        assert!(
            Builder::new()
                .content_xml(content.clone())
                .sheets()
                .is_err(),
            "standalone parser accepted {reference}"
        );
        assert!(
            Builder::new().content_xml(content.clone()).build().is_err(),
            "builder accepted {reference}"
        );

        let bytes = raw_package(&content).expect("raw ODS package");
        assert!(
            Spreadsheet::from_bytes(bytes).is_err(),
            "fused parser accepted {reference}"
        );
    }
}
