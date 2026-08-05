use litchi_docx::Package;
use litchi_ooxml_common::Props;
use tempfile::NamedTempFile;

#[test]
fn package_facade_owns_core_property_crud() {
    let mut package = Package::new().expect("fresh DOCX");
    assert_eq!(package.props().cloned(), Some(Props::new()));

    let previous = package.put_props(Props::new().title("Moved"));
    assert!(previous.is_some());
    assert_eq!(
        package.props().and_then(|props| props.title.as_deref()),
        Some("Moved")
    );

    let path = NamedTempFile::with_suffix(".docx").unwrap();
    package.save(path.path()).unwrap();
    let mut reopened = Package::open(path.path()).unwrap();
    assert_eq!(
        reopened.props().and_then(|props| props.title.as_deref()),
        Some("Moved")
    );
    assert!(reopened.clear_props().is_some());

    let cleared = NamedTempFile::with_suffix(".docx").unwrap();
    reopened.save(cleared.path()).unwrap();
    let reopened = Package::open(cleared.path()).unwrap();
    assert!(reopened.props().is_none());
}
