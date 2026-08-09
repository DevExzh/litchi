#![allow(
    clippy::unwrap_used,
    reason = "focused tests fail at the unexpected audit result"
)]

use xml_minifier::audit::{self, Kind, Limits, Resource, package};

fn limits_for(input: &[u8]) -> Limits {
    Limits::new(input.len(), 16, 128, 32, input.len(), input.len()).unwrap()
}

#[test]
fn accepts_semantically_meaningful_content_without_rewriting() {
    let xml = br#"<?xml version="1.0"?><root a='a"b'>Hello <b>world</b> !<!--c--><?keep x?><![CDATA[  ]]>&amp;&#32;&#x20;</root>"#;
    let report = audit::verify(xml, limits_for(xml)).unwrap();
    assert_eq!(report.max_depth(), 2);
    assert_eq!(report.attributes(), 1);
}

#[test]
fn preserves_plain_space_nodes_and_mixed_boundaries() {
    for xml in [
        b"<p><b>a</b> <i>b</i></p>".as_slice(),
        b"<p>   </p>".as_slice(),
        b"<p>a  b</p>".as_slice(),
    ] {
        let _report = audit::verify(xml, limits_for(xml)).unwrap();
    }
}

#[test]
fn authored_publication_refuses_ambiguous_spaces_but_preserves_explicit_content() {
    for xml in [
        b"<root> <child/></root>".as_slice(),
        b"<p><b>a</b> <i>b</i></p>".as_slice(),
        b"<p>   </p>".as_slice(),
    ] {
        let error = audit::verify_authored(xml, limits_for(xml)).unwrap_err();
        assert!(matches!(
            error,
            audit::Error::NotCompact(violation)
                if violation.kind() == Kind::AmbiguousWhitespace
        ));
    }

    let explicit = b"<p xml:space=\"preserve\"><b>a</b> <i>b</i></p>";
    let _report = audit::verify_authored(explicit, limits_for(explicit)).unwrap();

    let entity_split = b"<p>boxed &lt;text&gt; &amp; more</p>";
    let _report = audit::verify_authored(entity_split, limits_for(entity_split)).unwrap();
}

#[test]
fn package_xml_classification_uses_names_and_media_types() {
    for (name, media_type) in [
        ("[Content_Types].xml", "application/octet-stream"),
        ("_rels/.rels", "application/octet-stream"),
        ("manifest.rdf", "application/octet-stream"),
        ("custom/data", "application/rdf+xml; charset=UTF-8"),
        (
            "signature",
            "application/vnd.oasis.opendocument.digital-signature+xml",
        ),
    ] {
        assert!(package::is_xml_part(name, media_type));
    }
    assert!(!package::is_xml_part("Pictures/image.png", "image/png"));
}

#[test]
fn rejects_all_whitespace_outside_the_document_element() {
    for xml in [
        b" <root/>".as_slice(),
        b"<root/> ".as_slice(),
        b"<?xml version=\"1.0\"?> <root/>".as_slice(),
        b"<root/><!--comment--> ".as_slice(),
    ] {
        let error = audit::verify(xml, limits_for(xml)).unwrap_err();
        assert!(matches!(
            error,
            audit::Error::NotCompact(violation)
                if violation.kind() == Kind::FormattingWhitespace
        ));
    }
}

#[test]
fn rejects_cdata_outside_the_document_element() {
    for xml in [
        b"<![CDATA[ ]]><root/>".as_slice(),
        b"<root/><![CDATA[ ]]>".as_slice(),
    ] {
        assert!(matches!(
            audit::verify(xml, limits_for(xml)),
            Err(audit::Error::Malformed { .. })
        ));
    }
}

#[test]
fn inherited_xml_space_preserves_whitespace_and_default_resets_it() {
    let preserved = b"<a xml:space=\"preserve\">\n<b> </b></a>";
    let _preserved_report = audit::verify(preserved, limits_for(preserved)).unwrap();

    let reset = b"<a xml:space=\"preserve\"><b xml:space=\"default\">\n</b></a>";
    let error = audit::verify(reset, limits_for(reset)).unwrap_err();
    assert!(matches!(
        error,
        audit::Error::NotCompact(violation) if violation.kind() == Kind::FormattingWhitespace
    ));

    let nested = b"<a xml:space=\"preserve\">\n<b><c>\t</c></b></a>";
    let _nested_report = audit::verify(nested, limits_for(nested)).unwrap();

    let normalized = b"<a xml:space=\"pre&#115;erve\">\n</a>";
    let _normalized_report = audit::verify(normalized, limits_for(normalized)).unwrap();

    let invalid = b"<a xml:space=\"keep\">text</a>";
    assert!(matches!(
        audit::verify(invalid, limits_for(invalid)),
        Err(audit::Error::Malformed { .. })
    ));
}

#[test]
fn preserves_pi_comment_cdata_and_entity_events() {
    let xml = b"<a><?keep x\n?><!--note\n--><![CDATA[\n ]]>&amp;&#32;&#x20;</a>";
    let report = audit::verify(xml, limits_for(xml)).unwrap();
    assert!(report.text_bytes() >= 3);
}

#[test]
fn rejects_doctype_without_resolving_entities() {
    let xml = b"<!DOCTYPE a [<!ENTITY x 'value'>]><a>&x;</a>";
    assert!(matches!(
        audit::verify(xml, limits_for(xml)),
        Err(audit::Error::Doctype { offset: 0 })
    ));
}

#[test]
fn rejects_each_lexical_whitespace_defect() {
    for xml in [
        b"<a>\n</a>".as_slice(),
        b"<a  b=\"c\"/>".as_slice(),
        b"<a\tb=\"c\"/>".as_slice(),
        b"<a b = \"c\"/>".as_slice(),
    ] {
        assert!(audit::verify(xml, limits_for(xml)).is_err(), "{xml:?}");
    }

    for xml in [
        b"<a />".as_slice(),
        b"<a b=\"c\" />".as_slice(),
        b"<a></a >".as_slice(),
    ] {
        let error = audit::verify(xml, limits_for(xml)).unwrap_err();
        assert!(matches!(
            error,
            audit::Error::NotCompact(violation) if violation.kind() == Kind::WhitespaceBeforeClose
        ));
    }
}

#[test]
fn limits_are_inclusive_and_fail_with_typed_accounting() {
    let xml = b"<a x=\"1\">t</a>";
    let boundary = Limits::new(xml.len(), 1, 4, 1, 9, 1).unwrap();
    let _boundary_report = audit::verify(xml, boundary).unwrap();

    let below = Limits::new(xml.len() - 1, 1, 4, 1, 9, 1).unwrap();
    let error = audit::verify(xml, below).unwrap_err();
    assert!(matches!(
        error,
        audit::Error::Limit {
            resource: Resource::Bytes,
            limit,
            actual,
            ..
        } if limit == xml.len() - 1 && actual == xml.len()
    ));
}

#[test]
fn configuration_ceiling_is_inclusive_and_ceiling_plus_one_fails() {
    for resource in [
        Resource::Attributes,
        Resource::Bytes,
        Resource::Depth,
        Resource::Events,
        Resource::TextBytes,
        Resource::TokenBytes,
    ] {
        let ceiling = Limits::ceiling(resource);
        let _exact = Limits::builder().limit(resource, ceiling).unwrap().build();
        let error = Limits::builder().limit(resource, ceiling + 1).unwrap_err();
        assert_eq!(error.resource(), resource);
        assert_eq!(error.requested(), ceiling + 1);
        assert_eq!(error.ceiling(), ceiling);
    }

    assert!(Limits::builder().depth(usize::MAX).is_err());
    let defaults = Limits::default();
    assert!(
        Limits::new(
            usize::MAX,
            defaults.max_depth(),
            defaults.max_events(),
            defaults.max_attributes(),
            defaults.max_token_bytes(),
            defaults.max_text_bytes(),
        )
        .is_err()
    );
    let limits = Limits::builder()
        .bytes(Limits::ceiling(Resource::Bytes))
        .unwrap()
        .build();
    let narrowed = limits.narrow(Resource::Bytes, 17);
    assert_eq!(narrowed.max_bytes(), 17);
    assert_eq!(narrowed.narrow(Resource::Bytes, usize::MAX).max_bytes(), 17);
}

#[test]
fn package_helper_borrows_parts_and_enforces_aggregate_limits() {
    let first = b"<a/>";
    let second = b"<b>text</b>";
    let parts = [
        package::Part::new("word/document.xml", first),
        package::Part::new("content.xml", second),
    ];
    let document = Limits::new(second.len(), 2, 8, 2, second.len(), second.len()).unwrap();
    let report = package::verify(
        parts,
        package::Limits::new(document, 2, first.len() + second.len()),
    )
    .unwrap();
    assert_eq!(report.parts(), 2);

    let error = package::verify(parts, package::Limits::new(document, 1, usize::MAX)).unwrap_err();
    assert!(matches!(
        error,
        package::Error::Limit {
            resource: "parts",
            limit: 1,
            actual: 2
        }
    ));
}

#[test]
fn malformed_and_invalid_encoding_are_typed() {
    assert!(matches!(
        audit::verify(b"<a>", Limits::default()),
        Err(audit::Error::Malformed { .. })
    ));
    assert!(matches!(
        audit::verify(b"<a>\xff</a>", Limits::default()),
        Err(audit::Error::Encoding { valid_up_to: 3 })
    ));
    assert!(matches!(
        audit::verify(b"<a/><b/>", Limits::default()),
        Err(audit::Error::Malformed { .. })
    ));
}
