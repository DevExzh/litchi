//! Namespace URI normalization and reserved-binding contract tests.

#![allow(
    clippy::unwrap_used,
    reason = "small fixed namespace values keep these assertions focused"
)]

use litchi_odf_common::namespace::{
    MAX_NAMESPACE_URI_BYTES, XMLNS, XMLNS_DECLARATION_NAMESPACE, namespace_matches,
    resolved_namespace_uri, validate_namespace_binding,
};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::borrow::Cow;

fn decoder() -> quick_xml::Decoder {
    NsReader::from_str("").decoder()
}

#[test]
fn equivalent_namespace_lexical_forms_match_without_losing_raw_bytes() {
    let raw = b"urn:oasis:names:tc:opendocument:xmlns:office&#58;1.0";
    let namespace = ResolveResult::Bound(Namespace(raw));
    let resolved = resolved_namespace_uri(&namespace, decoder(), "office element")
        .unwrap()
        .unwrap();
    assert_eq!(resolved, "urn:oasis:names:tc:opendocument:xmlns:office:1.0");
    assert!(matches!(resolved, Cow::Owned(_)));
    assert_eq!(raw, b"urn:oasis:names:tc:opendocument:xmlns:office&#58;1.0");

    assert!(
        namespace_matches(
            &namespace,
            "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
            decoder(),
            "office element",
        )
        .unwrap()
    );
    assert!(
        namespace_matches(
            &namespace,
            b"urn:oasis:names:tc:opendocument:xmlns:office:1.0",
            decoder(),
            "office element",
        )
        .unwrap()
    );
}

#[test]
fn ordinary_namespace_uri_uses_borrowed_fast_path() {
    let raw = b"urn:oasis:names:tc:opendocument:xmlns:text:1.0";
    let namespace = ResolveResult::Bound(Namespace(raw));
    let resolved = resolved_namespace_uri(&namespace, decoder(), "text element")
        .unwrap()
        .unwrap();
    assert!(matches!(resolved, Cow::Borrowed(_)));
    assert_eq!(
        resolved.as_ref(),
        "urn:oasis:names:tc:opendocument:xmlns:text:1.0"
    );
}

#[test]
fn xml10_attribute_normalization_handles_predefined_numeric_and_eol_values() {
    let raw = b"urn:example:foo&amp;&#x3A;&#58;\tbar\r\nbaz";
    let namespace = ResolveResult::Bound(Namespace(raw));
    let resolved = resolved_namespace_uri(&namespace, decoder(), "extension element")
        .unwrap()
        .unwrap();
    assert_eq!(resolved, "urn:example:foo&:: bar baz");
}

#[test]
fn malformed_entities_and_xml10_controls_fail_closed() {
    for raw in [
        b"urn:example:&unknown;".as_slice(),
        b"urn:example:&unknown".as_slice(),
        b"urn:example:&#x;".as_slice(),
        b"urn:example:&#+65;".as_slice(),
        b"urn:example:&#x+41;".as_slice(),
        b"urn:example:&#x110000;".as_slice(),
        b"urn:example:\x01".as_slice(),
        b"urn:example:&#x1;".as_slice(),
        b"urn:example:<foreign".as_slice(),
    ] {
        let namespace = ResolveResult::Bound(Namespace(raw));
        assert!(resolved_namespace_uri(&namespace, decoder(), "malformed URI").is_err());
    }
}

#[test]
fn reserved_namespace_bindings_require_their_reserved_prefixes() {
    let xml_namespace = ResolveResult::Bound(Namespace(XMLNS.as_bytes()));
    assert!(
        validate_namespace_binding(None, &xml_namespace, decoder(), "default binding").is_err()
    );
    assert!(
        validate_namespace_binding(
            Some(b"other"),
            &xml_namespace,
            decoder(),
            "prefixed binding",
        )
        .is_err()
    );
    assert!(
        validate_namespace_binding(Some(b"xml"), &xml_namespace, decoder(), "xml binding",)
            .unwrap()
            .is_some()
    );

    let xmlns_namespace = ResolveResult::Bound(Namespace(XMLNS_DECLARATION_NAMESPACE.as_bytes()));
    assert!(resolved_namespace_uri(&xmlns_namespace, decoder(), "xmlns binding").is_err());
    assert!(
        validate_namespace_binding(Some(b"xmlns"), &xmlns_namespace, decoder(), "xmlns binding",)
            .is_err()
    );

    let empty = ResolveResult::Bound(Namespace(b""));
    assert!(
        validate_namespace_binding(Some(b""), &empty, decoder(), "named empty binding").is_err()
    );
    assert!(
        validate_namespace_binding(
            Some(b"xmlns"),
            &ResolveResult::Unbound,
            decoder(),
            "unbound xmlns binding",
        )
        .is_err()
    );
    assert!(
        validate_namespace_binding(None, &empty, decoder(), "default empty binding")
            .unwrap()
            .is_some()
    );
    assert!(
        validate_namespace_binding(Some(b"future"), &empty, decoder(), "prefixed empty binding",)
            .is_err()
    );
}

#[test]
fn unbound_and_unknown_prefixes_are_distinguished() {
    let unbound = ResolveResult::Unbound;
    assert!(!namespace_matches(&unbound, b"urn:example", decoder(), "unbound",).unwrap());

    let unknown = ResolveResult::Unknown(b"future".to_vec());
    assert!(resolved_namespace_uri(&unknown, decoder(), "unknown").is_err());
}

#[test]
fn lexical_uri_and_normalized_output_are_bounded_before_destination_allocation() {
    let raw = vec![b'x'; MAX_NAMESPACE_URI_BYTES + 1];
    let namespace = ResolveResult::Bound(Namespace(raw.as_slice()));
    assert!(resolved_namespace_uri(&namespace, decoder(), "oversized URI").is_err());

    let raw = vec![b'x'; MAX_NAMESPACE_URI_BYTES];
    let namespace = ResolveResult::Bound(Namespace(raw.as_slice()));
    assert!(resolved_namespace_uri(&namespace, decoder(), "maximum URI").is_ok());
}
