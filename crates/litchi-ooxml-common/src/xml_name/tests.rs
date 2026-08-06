use super::{NameError, NcName, QualifiedName, is_ncname, is_xml_name, parse, write};

#[test]
fn validates_unicode_ncname_and_name_grammar() {
    for value in ["rId1", "关系一", "éclair.一"] {
        assert!(is_ncname(value), "{value}");
        assert!(is_xml_name(value), "{value}");
    }
    for value in ["", "1relationship", "r:id", "relationship id"] {
        assert!(!is_ncname(value), "{value}");
    }
    assert!(is_xml_name(":leading-colon"));
    assert!(!is_xml_name("1leading"));
}

#[test]
fn parses_and_writes_prefixed_and_unprefixed_qnames_without_duplicate_storage() {
    let prefixed = parse("w:document").expect("prefixed QName");
    assert_eq!(prefixed.as_str(), "w:document");
    assert_eq!(prefixed.prefix(), Some("w"));
    assert_eq!(prefixed.local(), "document");

    let unprefixed = QualifiedName::try_from("document").expect("unprefixed QName");
    assert_eq!(unprefixed.prefix(), None);
    assert_eq!(unprefixed.local(), "document");

    let mut output = String::new();
    write(&prefixed, &mut output).expect("write QName");
    assert_eq!(output, "w:document");
    assert_eq!(NcName::try_from("w").expect("NCName").as_str(), "w");
}

#[test]
fn rejects_empty_parts_multiple_colons_and_invalid_starts() {
    for value in ["", ":local", "prefix:", "a:b:c", "1:local", "prefix:1"] {
        assert!(
            matches!(parse(value), Err(NameError::InvalidQualifiedName(_))),
            "{value}"
        );
    }
    assert!(matches!(
        NcName::try_from("a:b"),
        Err(NameError::InvalidNcName(_))
    ));
}
