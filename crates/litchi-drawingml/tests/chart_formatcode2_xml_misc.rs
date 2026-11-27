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
            formatcode2::read(&with_markup(markup), &formatcode2::Limits::DEFAULT).is_err(),
            "{markup:?}"
        );
    }
    let source = with_markup(b"<!-- <? --><?keep legitimate?>");
    let parsed = formatcode2::read(&source, &formatcode2::Limits::DEFAULT).unwrap();
    assert_eq!(
        formatcode2::write(&parsed, &formatcode2::Limits::DEFAULT).unwrap(),
        source
    );
}

#[test]
fn processing_instruction_targets_must_be_valid_and_not_reserved_xml() {
    for markup in [
        b"<?1bad?>".as_slice(),
        b"<?XML reserved?>",
        b"<?xMl reserved?>",
    ] {
        assert!(
            formatcode2::read(&with_markup(markup), &formatcode2::Limits::DEFAULT).is_err(),
            "{markup:?}"
        );
    }
    let source = with_markup(b"<?xml-stylesheet inert?>");
    assert_eq!(
        formatcode2::write(
            &formatcode2::read(&source, &formatcode2::Limits::DEFAULT).unwrap(),
            &formatcode2::Limits::DEFAULT,
        )
        .unwrap(),
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
        let result = formatcode2::read_attribute_with_bindings(
            xml.as_bytes(),
            &[("f", NAMESPACE)],
            &formatcode2::Limits::DEFAULT,
        );
        if extra == 0 {
            assert!(result.is_ok(), "{result:?}");
        } else {
            assert!(
                matches!(
                    result,
                    Err(litchi_drawingml::Error::Limit {
                        kind: litchi_drawingml::LimitKind::XmlBytes,
                        maximum: formatcode2::MAX_XML_BYTES,
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
    assert!(formatcode2::read(invalid.as_bytes(), &formatcode2::Limits::DEFAULT).is_err());
    let attribute =
        format!("<owner xmlns:f=\"{NAMESPACE}\" xmlns:=\"urn:bad\" f:formatcode2='x'/>");
    assert!(
        formatcode2::read_attribute(attribute.as_bytes(), &formatcode2::Limits::DEFAULT).is_err()
    );
    assert!(
        formatcode2::read_attribute_with_bindings(
            b"<owner xmlns:='urn:bad' f:formatcode2='x'/>",
            &[("f", NAMESPACE)],
            &formatcode2::Limits::DEFAULT,
        )
        .is_err()
    );
    let valid = format!("<f:formatcode2 xmlns:f=\"{NAMESPACE}\" xmlns=\"\">x</f:formatcode2>");
    assert_eq!(
        formatcode2::read(valid.as_bytes(), &formatcode2::Limits::DEFAULT)
            .unwrap()
            .value(),
        "x"
    );
}

const ORDER: &str = "invalid: formatcode2 XML declaration has invalid attribute order";
const UNKNOWN: &str = "invalid: formatcode2 XML declaration has an unknown attribute";
const NO_VERSION: &str = "invalid: formatcode2 XML declaration must declare version";
const BAD_VERSION: &str = "invalid: formatcode2 XML declaration must use version 1.0";
const BAD_ENCODING: &str = "invalid: formatcode2 XML declaration must use UTF-8 encoding";
const BAD_STANDALONE: &str = "invalid: formatcode2 XML declaration standalone must be yes or no";
const DUPLICATE: &str = "xml: duplicated attribute";

/// Read a document whose XML declaration carries the given pseudo-attributes
/// and name the exact result: `ok`, or the error variant with its message.
///
/// Each letter of `attributes` is one pseudo-attribute: `v` version 1.0,
/// `e` encoding UTF-8, `s` standalone yes, `u` an unknown name, and the
/// rejected values `V` version 1.1, `E` encoding UTF-16, `S` standalone maybe.
/// `l` is the lower-case spelling `utf-8` and `n` is standalone no.
fn declaration_outcome(attributes: &str) -> String {
    let mut source = String::from("<?xml");
    for attribute in attributes.chars() {
        source.push(' ');
        source.push_str(match attribute {
            'v' => "version=\"1.0\"",
            'e' => "encoding=\"UTF-8\"",
            's' => "standalone=\"yes\"",
            'u' => "unknown=\"x\"",
            'V' => "version=\"1.1\"",
            'E' => "encoding=\"UTF-16\"",
            'S' => "standalone=\"maybe\"",
            'l' => "encoding=\"utf-8\"",
            'n' => "standalone=\"no\"",
            other => panic!("unknown declaration attribute key {other:?}"),
        });
    }
    source.push_str("?>");
    let mut document = source.into_bytes();
    document.extend_from_slice(&with_markup(b""));
    match formatcode2::read(&document, &formatcode2::Limits::DEFAULT) {
        Ok(value) => {
            assert_eq!(value.value(), "General");
            "ok".to_string()
        },
        Err(litchi_drawingml::Error::Invalid(message)) => format!("invalid: {message}"),
        // The XML reader reports a repeated name before the declaration
        // check sees it; its wording belongs to the reader, so only the kind
        // of refusal is pinned.
        Err(litchi_drawingml::Error::Xml(message)) if message.contains("duplicated attribute") => {
            DUPLICATE.to_string()
        },
        Err(litchi_drawingml::Error::Xml(message)) => format!("xml: {message}"),
        Err(other) => panic!("unexpected declaration error {other:?}"),
    }
}

fn assert_declaration_outcomes(rows: &[(&str, &str)]) {
    let actual: Vec<(&str, String)> = rows
        .iter()
        .map(|(attributes, _)| (*attributes, declaration_outcome(attributes)))
        .collect();
    let expected: Vec<(&str, String)> = rows
        .iter()
        .map(|(attributes, outcome)| (*attributes, (*outcome).to_string()))
        .collect();
    assert_eq!(actual, expected, "{actual:#?}");
}

#[test]
fn xml_declaration_accepts_only_version_then_encoding_then_standalone() {
    // Every sequence of up to three of version, encoding and standalone.
    assert_declaration_outcomes(&[
        ("", NO_VERSION),
        ("v", "ok"),
        ("vv", DUPLICATE),
        ("vvv", DUPLICATE),
        ("vve", DUPLICATE),
        ("vvs", DUPLICATE),
        ("ve", "ok"),
        ("vev", DUPLICATE),
        ("vee", DUPLICATE),
        ("ves", "ok"),
        ("vs", "ok"),
        ("vsv", DUPLICATE),
        ("vse", ORDER),
        ("vss", DUPLICATE),
        ("e", ORDER),
        ("ev", ORDER),
        ("evv", ORDER),
        ("eve", ORDER),
        ("evs", ORDER),
        ("ee", ORDER),
        ("eev", ORDER),
        ("eee", ORDER),
        ("ees", ORDER),
        ("es", ORDER),
        ("esv", ORDER),
        ("ese", ORDER),
        ("ess", ORDER),
        ("s", ORDER),
        ("sv", ORDER),
        ("svv", ORDER),
        ("sve", ORDER),
        ("svs", ORDER),
        ("se", ORDER),
        ("sev", ORDER),
        ("see", ORDER),
        ("ses", ORDER),
        ("ss", ORDER),
        ("ssv", ORDER),
        ("sse", ORDER),
        ("sss", ORDER),
    ]);
}

#[test]
fn xml_declaration_unknown_attributes_and_rejected_values() {
    assert_declaration_outcomes(&[
        // An unknown name is refused wherever it stands.
        ("u", UNKNOWN),
        ("uv", UNKNOWN),
        ("ue", UNKNOWN),
        ("vu", UNKNOWN),
        ("veu", UNKNOWN),
        ("vus", UNKNOWN),
        ("vsu", UNKNOWN),
        ("vesu", UNKNOWN),
        ("vesv", DUPLICATE),
        // A rejected value in a legal position.
        ("V", BAD_VERSION),
        ("Vu", BAD_VERSION),
        ("Ve", BAD_VERSION),
        ("vE", BAD_ENCODING),
        ("vEs", BAD_ENCODING),
        ("vS", BAD_STANDALONE),
        ("veS", BAD_STANDALONE),
        // The position is checked before the value.
        ("E", ORDER),
        ("S", ORDER),
        ("eV", ORDER),
        ("sV", ORDER),
        ("vsE", ORDER),
        ("vsS", DUPLICATE),
        ("veE", DUPLICATE),
        // Accepted spellings.
        ("vl", "ok"),
        ("vn", "ok"),
        ("vln", "ok"),
        ("vnl", ORDER),
        ("vel", DUPLICATE),
        ("vsn", DUPLICATE),
    ]);
}
