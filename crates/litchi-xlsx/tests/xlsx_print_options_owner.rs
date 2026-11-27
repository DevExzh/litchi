use litchi_xlsx::print_options::{PrintOptions, parse_print_options};

#[test]
fn standalone_print_options_parse_through_the_owner() {
    let document = concat!(
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"#,
        r#"<printOptions horizontalCentered="1" gridLines="1" gridLinesSet="true"/>"#,
        r#"</worksheet>"#,
    );
    let value = parse_print_options(document.as_bytes()).unwrap().unwrap();

    fn accepts_canonical_owner(_: &PrintOptions) {}
    accepts_canonical_owner(&value);

    let _: &PrintOptions = &value;
    assert!(value.horizontal_centered());
    assert!(value.prints_grid_lines());
}
