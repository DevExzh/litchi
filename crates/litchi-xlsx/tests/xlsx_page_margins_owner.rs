use litchi_ooxml::xlsx::page_margins::{Margins, PageMargin, parse_page_margins};

#[test]
fn host_reexports_the_canonical_page_margins_owner() {
    let document = concat!(
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"#,
        r#"<pageMargins left="0.75" right="0.8" top="1" bottom="1.1" header="0.5" footer="0.6"/>"#,
        r#"</worksheet>"#,
    );
    let value = parse_page_margins(document.as_bytes()).unwrap().unwrap();

    fn accepts_canonical_owner(_: &litchi_xlsx::page_margins::Margins) {}
    accepts_canonical_owner(&value);

    let _: &Margins = &value;
    let _: PageMargin = value.left();
    assert_eq!(value.left().inches(), 0.75);
    assert!((value.footer().points() - 43.2).abs() < 1e-12);
}
