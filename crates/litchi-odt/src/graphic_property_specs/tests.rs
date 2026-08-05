use super::model::{Kind, Namespace, Value};

#[test]
fn table_contains_all_unique_graphic_property_names() {
    assert_eq!(Kind::ALL.len(), 174);

    let mut names = std::collections::BTreeSet::new();
    for kind in Kind::ALL {
        assert!(names.insert((kind.namespace(), kind.local_name())));
    }
    assert!(
        names
            .iter()
            .any(|(namespace, _)| *namespace == Namespace::Dr3d)
    );
    assert!(
        names
            .iter()
            .any(|(namespace, _)| *namespace == Namespace::Draw)
    );
    assert!(
        names
            .iter()
            .any(|(namespace, _)| *namespace == Namespace::Fo)
    );
    assert!(
        names
            .iter()
            .any(|(namespace, _)| *namespace == Namespace::Style)
    );
    assert!(
        names
            .iter()
            .any(|(namespace, _)| *namespace == Namespace::Svg)
    );
    assert!(
        names
            .iter()
            .any(|(namespace, _)| *namespace == Namespace::Text)
    );
}

#[test]
fn table_uses_typed_lexical_values() {
    assert_eq!(
        Kind::DrawFill.parse_value("solid").unwrap(),
        Value::Keyword("solid".to_owned())
    );
    assert_eq!(
        Kind::DrawOpacity.parse_value("50%").unwrap(),
        Value::Percent("50%".to_owned())
    );
    assert!(Kind::DrawOpacity.parse_value("not-a-percent").is_err());
}
