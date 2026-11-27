use litchi_xlsx::page_setup::{Fit, parse_worksheet_page_setup};

#[test]
fn page_setup_parser_reads_typed_fit_bounds() {
    let xml = br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><pageSetup fitToWidth="1" fitToHeight="0"/></worksheet>"#;
    let setup = parse_worksheet_page_setup(xml).unwrap().unwrap();
    assert_eq!(setup.fit_to_width, Some(Fit::new(1).unwrap()));
    assert_eq!(setup.fit_to_height, Some(Fit::new(0).unwrap()));
}
