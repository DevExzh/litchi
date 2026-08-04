use super::*;

use crate::Result;
use litchi_opc::{OpcPackage, PackURI};

const START: &str =
    r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"#;

fn parse(body: &str) -> Result<Option<Settings>> {
    parse_worksheet_header_footer(format!("{START}{body}</worksheet>").as_bytes())
}

fn parse_fixture(path: &str) -> Settings {
    let package = OpcPackage::open(path).unwrap();
    let uri = PackURI::new("/xl/worksheets/sheet1.xml").unwrap();
    parse_worksheet_header_footer(package.get_part(&uri).unwrap().blob())
        .unwrap()
        .unwrap()
}

#[test]
fn parses_all_variants_defaults_and_sections() {
    let settings = parse(r#"<headerFooter differentOddEven="1" differentFirst="true" scaleWithDoc="0" alignWithMargins="false"><oddHeader>&amp;Lleft&amp;Ccenter&amp;Rright</oddHeader><oddFooter>&amp;P</oddFooter><evenHeader>even</evenHeader><evenFooter/><firstHeader>first</firstHeader><firstFooter>last</firstFooter></headerFooter>"#).unwrap().unwrap();
    assert!(settings.different_odd_even());
    assert!(settings.different_first());
    assert!(!settings.scale_with_document());
    assert!(!settings.align_with_margins());
    let header = settings.odd_header().unwrap();
    assert_eq!(header.raw(), "&Lleft&Ccenter&Rright");
    assert_eq!(header.left(), Some("left"));
    assert_eq!(header.center(), Some("center"));
    assert_eq!(header.right(), Some("right"));
    assert_eq!(settings.even_footer().unwrap().center(), Some(""));
}

#[test]
fn preserves_ampersands_and_unrecognized_formatting() {
    let settings = parse(r#"<headerFooter><oddHeader>&amp;Cone &amp;&amp; two &amp;&amp;&amp;&amp;&amp;K01+000</oddHeader></headerFooter>"#).unwrap().unwrap();
    let header = settings.odd_header().unwrap();
    assert_eq!(header.center(), Some("one && two &&&&&K01+000"));
}

#[test]
fn rejects_duplicates_order_and_nested_markup() {
    assert!(parse("<headerFooter><oddFooter/><oddHeader/></headerFooter>").is_err());
    assert!(parse("<headerFooter><oddHeader/><oddHeader/></headerFooter>").is_err());
    assert!(parse("<headerFooter><oddHeader><b/></oddHeader></headerFooter>").is_err());
    assert!(parse("<headerFooter scaleWithDoc=\"yes\"/>").is_err());
}

#[test]
fn loads_poi_ampersand_fixture() {
    let path = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/poi/test-data/spreadsheet/AmpersandHeader.xlsx"
    );
    let settings = parse_fixture(path);
    assert_eq!(
        settings.odd_header().unwrap().center(),
        Some("one && two &&&&")
    );
}

#[test]
fn loads_libreoffice_color_sections_fixture() {
    let path = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/libreoffice-core/sc/qa/unit/data/xlsx/tdf134459_HeaderFooterColor.xlsx"
    );
    let settings = parse_fixture(path);
    let header = settings.odd_header().unwrap();
    assert_eq!(header.left(), Some("&KC06040l"));
    assert_eq!(header.center(), Some("&K4C3789c"));
    assert_eq!(header.right(), Some("r"));
}
