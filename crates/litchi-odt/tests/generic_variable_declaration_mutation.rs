use litchi_odt::variable_declaration::{Body, Declaration, Group, Kind, Part, Scope, ValueType};
mod support;

const MIMETYPE: &str = "application/vnd.oasis.opendocument.text";
const CONTENT: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><office:body><office:text/></office:body></office:document-content>"#;
const REFERENCED_CONTENT: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><office:body><office:text><text:variable-decls><text:variable-decl text:name="counter" office:value-type="string"/></text:variable-decls><text:p><text:variable-get text:name="counter">cached</text:variable-get></text:p></office:text></office:body></office:document-content>"#;

fn package(content: &str) -> Vec<u8> {
    support::package(
        MIMETYPE,
        &[
            ("content.xml", content.as_bytes()),
            ("meta.xml", b"<meta-preserved/>"),
            ("settings.xml", b"<settings-preserved/>"),
            ("Objects/asset.bin", b"\x00preserved\xff"),
        ],
    )
}

fn group(name: &str) -> Group {
    Group {
        kind: Kind::Simple,
        part: Part::Content,
        scope: Scope::Body(Body::Text),
        declarations: vec![Declaration::Simple {
            name: name.to_string(),
            value_type: ValueType::String,
        }],
    }
}

#[test]
fn mutates_packaged_declarations_and_preserves_auxiliary_parts() {
    let mut document = litchi_odt::generic::Package::from_bytes(package(CONTENT)).unwrap();
    let first = group("counter");
    assert_eq!(
        document.set_variable_declaration_group(&first).unwrap(),
        None
    );
    assert_eq!(
        document.variable_declarations().unwrap().groups,
        vec![first.clone()]
    );

    let replacement = group("replacement");
    assert_eq!(
        document
            .set_variable_declaration_group(&replacement)
            .unwrap(),
        Some(first)
    );
    assert_eq!(
        document.get_file("Objects/asset.bin").unwrap(),
        b"\x00preserved\xff"
    );
    assert_eq!(document.get_file("meta.xml").unwrap(), b"<meta-preserved/>");
    assert_eq!(
        document.get_file("settings.xml").unwrap(),
        b"<settings-preserved/>"
    );

    assert_eq!(
        document
            .remove_variable_declaration_group(
                Part::Content,
                &Scope::Body(Body::Text),
                Kind::Simple,
            )
            .unwrap(),
        Some(replacement)
    );
    assert!(document.variable_declarations().unwrap().groups.is_empty());
}

#[test]
fn referenced_declaration_removal_is_atomic() {
    let mut document =
        litchi_odt::generic::Package::from_bytes(package(REFERENCED_CONTENT)).unwrap();
    let before = document.as_bytes().to_vec();
    assert!(
        document
            .remove_variable_declaration_group(
                Part::Content,
                &Scope::Body(Body::Text),
                Kind::Simple,
            )
            .is_err()
    );
    assert_eq!(document.as_bytes(), before);
    assert_eq!(document.variable_declarations().unwrap().groups.len(), 1);
}
