use litchi_ooxml::xlsx::ignored_errors::{
    IgnoredErrorType, IgnoredErrors, parse_worksheet_ignored_errors,
};

#[test]
fn host_reexports_the_canonical_ignored_errors_owner() {
    let document = concat!(
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"#,
        r#"<ignoredErrors><ignoredError sqref="A1" formula="1"/></ignoredErrors></worksheet>"#,
    );
    let value = parse_worksheet_ignored_errors(document.as_bytes())
        .unwrap()
        .unwrap();

    fn accepts_canonical_owner(_: &litchi_xlsx::ignored_errors::IgnoredErrors) {}
    accepts_canonical_owner(&value);

    let _: &IgnoredErrors = &value;
    assert!(value.entries()[0].ignores(IgnoredErrorType::Formula));
}
