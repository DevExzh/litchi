//! Change 0750: the source policy refuses every well-formedness defect of XML
//! 1.0 (Fifth Edition) and Namespaces in XML 1.0 (Third Edition) that the
//! tokenizer-level checks accepted, each as `Error::Malformed` at its offset,
//! and still accepts every well-formed spelling. The authored and default
//! policies keep their tokenizer-level verdicts, which fragment audits rely on.
#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "each test states one fixed expected verdict"
)]

use xml_minifier::audit::{self, Error, Limits, ReplacementError};

/// Asserts that the source policy refuses `xml` as malformed at `offset`
/// with a diagnostic containing `detail`, and that the pair audit reports the
/// same error for it as a replacement of a valid original.
fn refused(xml: &[u8], offset: usize, detail: &str) {
    let error = audit::verify_source(xml, Limits::default()).expect_err(&format!(
        "{:?} must be refused",
        String::from_utf8_lossy(xml)
    ));
    match &error {
        Error::Malformed {
            offset: actual,
            detail: text,
        } => {
            assert_eq!(
                *actual,
                offset,
                "{:?}: {text}",
                String::from_utf8_lossy(xml)
            );
            assert!(
                text.contains(detail),
                "{:?}: expected {detail:?}, got {text:?}",
                String::from_utf8_lossy(xml)
            );
        },
        other => panic!(
            "{:?} must be malformed, got {other:?}",
            String::from_utf8_lossy(xml)
        ),
    }
    assert_eq!(
        audit::verify_source_replacement(b"<valid/>", xml, Limits::default()),
        Err(ReplacementError::Replacement(error))
    );
}

fn accepted(xml: &[u8]) {
    if let Err(error) = audit::verify_source(xml, Limits::default()) {
        panic!(
            "{:?} is well-formed and must be accepted: {error:?}",
            String::from_utf8_lossy(xml)
        );
    }
}

// ------------------------------------------------ the reviewed gaps, one each

#[test]
fn a_reference_outside_the_document_element_is_refused() {
    // Its text is a space, so it once passed as whitespace.
    refused(b"<r/>& ;", 4, "character data outside the document element");
    refused(
        b"<r/>&#32;",
        4,
        "character data outside the document element",
    );
    refused(
        b"&amp;<r/>",
        0,
        "character data outside the document element",
    );
}

#[test]
fn an_empty_reference_is_refused() {
    refused(b"<r>&;</r>", 3, "malformed reference");
}

#[test]
fn a_reference_that_is_not_a_name_is_refused() {
    refused(b"<r>&a b;</r>", 3, "malformed reference");
}

#[test]
fn a_malformed_or_illegal_character_reference_is_refused() {
    refused(b"<r>&#xZZ;&#0;</r>", 3, "malformed character reference");
    refused(
        b"<r>&#0;</r>",
        3,
        "character reference to a character XML does not allow",
    );
    for reference in [
        "&#x1;",
        "&#xD800;",
        "&#xDFFF;",
        "&#xFFFE;",
        "&#xFFFF;",
        "&#x110000;",
        "&#99999999999999999999;",
    ] {
        refused(
            format!("<r>{reference}</r>").as_bytes(),
            3,
            "character reference to a character XML does not allow",
        );
    }
    for reference in ["&#;", "&#x;", "&#X41;", "&#-1;", "&#1a;"] {
        refused(
            format!("<r>{reference}</r>").as_bytes(),
            3,
            "malformed character reference",
        );
    }
}

#[test]
fn an_xml_declaration_inside_the_document_is_refused() {
    refused(
        b"<r><?xml version=\"1.0\"?></r>",
        3,
        "an XML declaration is allowed only at the start of the document",
    );
}

#[test]
fn a_cdata_section_close_in_character_data_is_refused() {
    refused(b"<r>]]></r>", 3, "']]>' is not allowed in character data");
    refused(
        b"<r>ab]]>c</r>",
        5,
        "']]>' is not allowed in character data",
    );
}

#[test]
fn an_undeclared_namespace_prefix_is_refused() {
    refused(b"<p:r/>", 1, "undeclared namespace prefix");
    refused(b"<r p:a=\"1\"/>", 3, "undeclared namespace prefix");
    // A binding ends with the element that declared it.
    refused(
        b"<r><a xmlns:p=\"u\"/><p:b/></r>",
        20,
        "undeclared namespace prefix",
    );
    refused(
        b"<r><a xmlns:p=\"u\"><p:b/></a><p:c/></r>",
        29,
        "undeclared namespace prefix",
    );
}

// ------------------------------------------------- further gaps of the kind

#[test]
fn a_character_xml_does_not_allow_is_refused_wherever_it_stands() {
    for (xml, offset) in [
        (b"<r>\x01</r>".as_slice(), 3),
        (b"<r>\x00</r>", 3),
        (b"<r>a\x0bb</r>", 4),
        (b"<r>\x0c</r>", 3),
        (b"<r>\x1f</r>", 3),
        (b"<r a=\"\x01\"/>", 6),
        (b"<r><!--\x01--></r>", 7),
        (b"<r><?pi \x01?></r>", 8),
        (b"<r><![CDATA[\x01]]></r>", 12),
        (b"<r/>\x01", 4),
        (b"<r>\xef\xbf\xbe</r>", 3),
        (b"<r>\xef\xbf\xbf</r>", 3),
        (b"<r a=\"\xef\xbf\xbf\"/>", 6),
    ] {
        refused(xml, offset, "is not allowed in XML");
    }
    // A lone surrogate is not UTF-8, so it stays an encoding refusal.
    assert_eq!(
        audit::verify_source(b"<r>\xed\xa0\x80</r>", Limits::default()),
        Err(Error::Encoding { valid_up_to: 3 })
    );
}

#[test]
fn an_attribute_value_holds_no_lt_and_only_well_formed_references() {
    refused(
        b"<r a=\"<\"/>",
        6,
        "'<' is not allowed in an attribute value",
    );
    refused(
        b"<r a='x<y'/>",
        7,
        "'<' is not allowed in an attribute value",
    );
    refused(
        b"<r a=\"&\"/>",
        6,
        "unterminated reference in an attribute value",
    );
    refused(
        b"<r a=\"a&b\"/>",
        7,
        "unterminated reference in an attribute value",
    );
    refused(
        b"<r a=\"&bogus;\"/>",
        6,
        "reference to an undeclared entity",
    );
    refused(
        b"<r a=\"&#0;\"/>",
        6,
        "character reference to a character XML does not allow",
    );
    refused(
        b"<r a=\"&#xD800;\"/>",
        6,
        "character reference to a character XML does not allow",
    );
}

#[test]
fn attributes_are_unique_by_namespace_name_and_local_name() {
    refused(
        b"<r xmlns:a=\"u\" xmlns:b=\"u\" a:x=\"1\" b:x=\"2\"/>",
        35,
        "two attributes have the same namespace name and local name",
    );
    // Among three prefixes, two of which alias one namespace name.
    refused(
        b"<r xmlns:a=\"u\" xmlns:b=\"v\" xmlns:c=\"u\" a:x=\"1\" b:x=\"2\" c:x=\"3\"/>",
        55,
        "two attributes have the same namespace name and local name",
    );
    // The same qualified name twice is refused by the tokenizer, as before.
    assert!(matches!(
        audit::verify_source(b"<r a=\"1\" a=\"2\"/>", Limits::default()),
        Err(Error::Malformed { detail, .. }) if detail.contains("duplicated attribute")
    ));
    // Aliases with distinct local names, and an unprefixed attribute beside a
    // prefixed one with the same local name, are distinct expanded names.
    accepted(b"<r xmlns:a=\"u\" xmlns:b=\"u\" a:x=\"1\" b:y=\"2\"/>");
    accepted(b"<r xmlns:a=\"u\" x=\"1\" a:x=\"2\"/>");
}

#[test]
fn the_reserved_prefixes_and_namespace_names_are_used_only_as_reserved() {
    let xml_name = "http://www.w3.org/XML/1998/namespace";
    let xmlns_name = "http://www.w3.org/2000/xmlns/";
    refused(
        b"<r xmlns:xml=\"urn:other\"/>",
        3,
        "the xml prefix must not be bound to another namespace name",
    );
    refused(
        format!("<r xmlns:xmlns=\"{xmlns_name}\"/>").as_bytes(),
        3,
        "the xmlns prefix must not be declared",
    );
    for name in [xml_name, xmlns_name] {
        refused(
            format!("<r xmlns:p=\"{name}\"/>").as_bytes(),
            3,
            "only a reserved prefix may be bound to a reserved namespace name",
        );
        refused(
            format!("<r xmlns=\"{name}\"/>").as_bytes(),
            3,
            "the default namespace must not be a reserved namespace name",
        );
    }
    // The namespace name is the normalized value, references resolved.
    refused(
        b"<r xmlns:p=\"http://www.w3.org/XML/1998/&#110;amespace\"/>",
        3,
        "only a reserved prefix may be bound to a reserved namespace name",
    );
    refused(
        b"<xmlns:r/>",
        1,
        "an element name must not use the xmlns prefix",
    );
    accepted(format!("<r xmlns:xml=\"{xml_name}\" xml:lang=\"en\"/>").as_bytes());
    accepted(b"<xml:r/>");
}

#[test]
fn a_prefix_must_not_be_undeclared() {
    refused(
        b"<r xmlns:p=\"\"/>",
        3,
        "a namespace prefix must not be undeclared",
    );
    // Undeclaring the default namespace is allowed.
    accepted(b"<r xmlns=\"u\"><a xmlns=\"\"/></r>");
}

#[test]
fn a_mismatched_end_tag_stays_refused() {
    for xml in [b"<r><a></b></r>".as_slice(), b"<r></R>", b"<r><a></></r>"] {
        assert!(
            matches!(
                audit::verify_source(xml, Limits::default()),
                Err(Error::Malformed { detail, .. }) if detail.contains("expected `</")
            ),
            "{xml:?}"
        );
    }
}

#[test]
fn a_comment_holds_no_double_hyphen_and_does_not_end_with_one() {
    refused(
        b"<r><!-- a -- b --></r>",
        10,
        "'--' is not allowed in a comment",
    );
    refused(b"<r><!-- a ---></r>", 10, "a comment must not end with '-'");
    refused(b"<r><!-----></r>", 7, "a comment must not end with '-'");
    accepted(b"<r><!----><!-- a-b - c --></r>");
}

#[test]
fn element_and_attribute_names_are_qualified_names() {
    for (xml, offset) in [
        (b"<1r/>".as_slice(), 1),
        (b"<-r/>", 1),
        (b"<r><a<b/></r>", 4),
        (b"<r><a\"b\"/></r>", 4),
        (b"<r>< a/></r>", 4),
        (b"<r>< a=\"1\"/></r>", 4),
        (b"<r><></></r>", 4),
        (b"<r 1a=\"1\"/>", 3),
        (b"<a/ >", 1),
    ] {
        refused(xml, offset, "invalid XML name");
    }
    for (xml, offset) in [
        (b"<r:/>".as_slice(), 1),
        (b"<:r/>", 1),
        (b"<a:b:c xmlns:a=\"u\"/>", 1),
        (b"<r a:=\"1\"/>", 3),
        (b"<r :a=\"1\"/>", 3),
        (b"<r xmlns:a=\"u\" a:b:c=\"1\"/>", 15),
        (b"<r xmlns:=\"u\"/>", 3),
    ] {
        refused(xml, offset, "not namespace-well-formed");
    }
    // Fifth Edition name characters.
    refused("<\u{d7}/>".as_bytes(), 1, "invalid XML name");
    refused("<\u{300}a/>".as_bytes(), 1, "invalid XML name");
    accepted("<\u{e9}t\u{e9} \u{e9}=\"1\" a\u{b7}-b.c=\"2\"/>".as_bytes());
    accepted("<\u{4e2d}\u{6587}/>".as_bytes());
}

#[test]
fn a_processing_instruction_target_is_a_name_other_than_xml() {
    for xml in [b"<r><?XML x?></r>".as_slice(), b"<r><?xMl?></r>"] {
        refused(xml, 5, "processing-instruction target 'xml' is reserved");
    }
    for (xml, offset) in [
        (b"<r><? x?></r>".as_slice(), 5),
        (b"<r><?1x?></r>", 5),
        (b"<r><?a:b x?></r>", 5),
        (b"<r><?a\"b?></r>", 5),
        (b"<??><r/>", 2),
        (b"<?xmlversion=\"1.0\"?><r/>", 2),
    ] {
        refused(xml, offset, "invalid processing-instruction target");
    }
    accepted(b"<?xml-stylesheet href=\"s\"?><r><?pi?><?pi ?><?xmlfoo x?></r>");
}

#[test]
fn the_xml_declaration_follows_its_grammar() {
    for (xml, offset, detail) in [
        (b"<?xml?><r/>".as_slice(), 2, "must declare a version"),
        (
            b"<?xml encoding=\"UTF-8\"?><r/>",
            6,
            "must declare its version first",
        ),
        (
            b"<?xml standalone=\"yes\" version=\"1.0\"?><r/>",
            6,
            "must declare its version first",
        ),
        (b"<?xml version=\"2.0\"?><r/>", 6, "version must be 1.x"),
        (b"<?xml version=\"1.\"?><r/>", 6, "version must be 1.x"),
        (b"<?xml version=\"1.0 \"?><r/>", 6, "version must be 1.x"),
        (
            b"<?xml version=\"1.0\" standalone=\"maybe\"?><r/>",
            20,
            "standalone must be 'yes' or 'no'",
        ),
        (
            b"<?xml version=\"1.0\" foo=\"bar\"?><r/>",
            20,
            "unexpected attribute",
        ),
        (
            b"<?xml version=\"1.0\" standalone=\"yes\" encoding=\"UTF-8\"?><r/>",
            37,
            "unexpected attribute",
        ),
        (
            b"<?xml version=\"1.0\" encoding=\"\"?><r/>",
            20,
            "invalid encoding name",
        ),
        (
            b"<?xml version=\"1.0\" encoding=\"8bit\"?><r/>",
            20,
            "invalid encoding name",
        ),
    ] {
        refused(xml, offset, detail);
    }
    // The audit reads UTF-8, so a declaration of another encoding is refused:
    // OPC rule M1.17 forbids naming any encoding but UTF-8 or UTF-16, even
    // over bytes that are all ASCII as these are, and UTF-16 over UTF-8 bytes
    // is the mismatch XML 1.0 section 4.3.3 makes fatal.
    for encoding in ["ISO-8859-1", "UTF-16", "windows-1252", "US-ASCII", "UTF8"] {
        refused(
            format!("<?xml version=\"1.0\" encoding=\"{encoding}\"?><r/>").as_bytes(),
            20,
            "names an encoding other than UTF-8",
        );
    }
    for xml in [
        b"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><r/>".as_slice(),
        b"<?xml version='1.0' encoding='utf-8'?><r/>",
        b"<?xml version=\"1.1\"?><r/>",
        b"<?xml\tversion = \"1.0\"\tstandalone='no' ?><r/>",
        b"\xef\xbb\xbf<?xml version=\"1.0\"?><r/>",
    ] {
        accepted(xml);
    }
}

#[test]
fn the_xml_declaration_stands_only_at_the_start_of_the_document() {
    let declaration = "an XML declaration is allowed only at the start of the document";
    refused(b"\n<?xml version=\"1.0\"?><r/>", 1, declaration);
    refused(b" <?xml version=\"1.0\"?><r/>", 1, declaration);
    refused(b"<!--c--><?xml version=\"1.0\"?><r/>", 8, declaration);
    refused(b"<r/><?xml version=\"1.0\"?>", 4, declaration);
    refused(
        b"<?xml version=\"1.0\"?><?xml version=\"1.0\"?><r/>",
        21,
        declaration,
    );
    refused(b"\xef\xbb\xbf <?xml version=\"1.0\"?><r/>", 4, declaration);
}

#[test]
fn only_the_predefined_entities_are_declared() {
    for (xml, offset) in [
        (b"<r>&nbsp;</r>".as_slice(), 3),
        (b"<r>&AMP;</r>", 3),
        (b"<r>a&unknown;</r>", 4),
    ] {
        refused(xml, offset, "reference to an undeclared entity");
    }
    accepted(b"<r a=\"&lt;&gt;&amp;&apos;&quot;&#9;&#x10FFFF;\">&amp;&lt;&gt;&apos;&quot;&#65;&#x41;&#0065;</r>");
}

#[test]
fn well_formed_spellings_stay_accepted() {
    for xml in [
        b"<r>]]&gt;]] ]>x</r>".as_slice(),
        b"<r><![CDATA[]]]]><![CDATA[>]]></r>",
        b"<r a=\"'>\" b='\">' c=\"\t\n\"/>",
        "<r>\u{85}\u{2028}\u{7f}\u{f8ff}\u{fffd}\u{feff}\u{10000}</r>".as_bytes(),
        b"<p:r xmlns:p=\"u\"/>",
        b"<r p:a=\"1\" xmlns:p=\"u\"/>",
        b"<r xmlns:p=\"u\"><p:a xmlns:p=\"v\"><p:b/></p:a><p:c/></r>",
        b"<r xmlns:p=\" \"><p:a/></r>",
        b"<r xmlns:p=\"u\" p:xmlns=\"1\" xml:foo=\"2\"/>",
        b"<r xmlns=\"\"/>",
    ] {
        accepted(xml);
    }
}

// ------------------------------------------ fragment callers keep their audit

/// The authored and default policies audit fragments too, whose namespace
/// declarations belong to an enclosing document, and whole documents wrapped
/// in a synthetic element. They keep their tokenizer-level verdicts.
#[test]
fn the_authored_and_default_policies_keep_their_verdicts_for_fragment_callers() {
    for xml in [
        b"<w:p><w:r><w:t>x</w:t></w:r></w:p>".as_slice(),
        b"<wrapper><?xml version=\"1.0\"?><w:document/></wrapper>",
        b"<r>&unknown;</r>",
    ] {
        let _authored = audit::verify_authored(xml, Limits::default()).unwrap();
        let _default = audit::verify(xml, Limits::default()).unwrap();
        assert!(audit::verify_source(xml, Limits::default()).is_err());
    }
}
