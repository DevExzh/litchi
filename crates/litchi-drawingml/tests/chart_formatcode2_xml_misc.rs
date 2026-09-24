use litchi_drawingml::chart::extension::formatcode2::{self, NAMESPACE};

fn with_markup(markup: &[u8]) -> Vec<u8> {
    let mut source = format!("<f:formatcode2 xmlns:f=\"{NAMESPACE}\">General").into_bytes();
    source.extend_from_slice(markup);
    source.extend_from_slice(b"</f:formatcode2>");
    source
}

#[test]
fn opaque_comments_and_pis_still_require_xml_characters_and_utf8() {
    for markup in [
        b"<!--\x01-->".as_slice(),
        b"<!--\xff-->",
        b"<?keep \x01?>",
        b"<?keep \xff?>",
    ] {
        assert!(
            formatcode2::read(&with_markup(markup)).is_err(),
            "{markup:?}"
        );
    }
    let source = with_markup(b"<!-- <? --><?keep legitimate?>");
    let parsed = formatcode2::read(&source).unwrap();
    assert_eq!(formatcode2::write(&parsed).unwrap(), source);
}

#[test]
fn processing_instruction_targets_must_be_valid_and_not_reserved_xml() {
    for markup in [
        b"<?1bad?>".as_slice(),
        b"<?XML reserved?>",
        b"<?xMl reserved?>",
    ] {
        assert!(
            formatcode2::read(&with_markup(markup)).is_err(),
            "{markup:?}"
        );
    }
    let source = with_markup(b"<?xml-stylesheet inert?>");
    assert_eq!(
        formatcode2::write(&formatcode2::read(&source).unwrap()).unwrap(),
        source
    );
}

#[test]
fn inherited_namespace_view_has_an_explicit_combined_byte_limit() {
    let prefix = "<f:owner f:formatcode2='x' padding='";
    let suffix = "'/>";
    let declaration = format!(" xmlns:f=\"{NAMESPACE}\"");
    let padding = formatcode2::MAX_XML_BYTES - prefix.len() - suffix.len() - declaration.len();
    for extra in [0, 1] {
        let xml = format!("{prefix}{}{suffix}", "p".repeat(padding + extra));
        assert!(xml.len() < formatcode2::MAX_XML_BYTES);
        let result = formatcode2::read_attribute_with_bindings(xml.as_bytes(), &[("f", NAMESPACE)]);
        if extra == 0 {
            assert!(result.is_ok(), "{result:?}");
        } else {
            assert!(
                matches!(
                    result,
                    Err(litchi_drawingml::Error::Limit {
                        limit: formatcode2::MAX_XML_BYTES,
                        ..
                    })
                ),
                "{result:?}"
            );
        }
    }
}

#[test]
fn empty_named_namespace_prefix_is_not_a_default_declaration() {
    let invalid =
        format!("<f:formatcode2 xmlns:f=\"{NAMESPACE}\" xmlns:=\"urn:bad\">x</f:formatcode2>");
    assert!(formatcode2::read(invalid.as_bytes()).is_err());
    let attribute =
        format!("<owner xmlns:f=\"{NAMESPACE}\" xmlns:=\"urn:bad\" f:formatcode2='x'/>");
    assert!(formatcode2::read_attribute(attribute.as_bytes()).is_err());
    assert!(
        formatcode2::read_attribute_with_bindings(
            b"<owner xmlns:='urn:bad' f:formatcode2='x'/>",
            &[("f", NAMESPACE)]
        )
        .is_err()
    );
    let valid = format!("<f:formatcode2 xmlns:f=\"{NAMESPACE}\" xmlns=\"\">x</f:formatcode2>");
    assert_eq!(formatcode2::read(valid.as_bytes()).unwrap().value(), "x");
}
