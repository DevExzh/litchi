//! Integration coverage for the shared DrawingML 2012 `themeFamily` owner.
//!
//! The native fragment is extracted from LibreOffice's `tdf131082.pptx`; the
//! provenance and hashes live beside the fixture.  These tests deliberately
//! exercise the fragment owner only.  Package placement remains the
//! responsibility of a host format.

use litchi_drawingml::theme::family::{
    self, DRAWINGML_NAMESPACE, Family, Guid, MAX_ATTRIBUTES, MAX_DEPTH, MAX_NAME_BYTES,
    MAX_NAMESPACE_BYTES, MAX_NAMESPACE_DECLARATIONS, MAX_XML_BYTES, NAMESPACE, Snapshot,
    XML_NAMESPACE,
};

const NATIVE: &[u8] = include_bytes!("fixtures/theme-family-native.xml");
const OFFICE_ID: &str = "{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}";
const OFFICE_VID: &str = "{4A3C46E8-61CC-4603-A589-7422A47A8E4A}";

const FAMILY_WITH_OPAQUE_EXTENSION: &[u8] = br#"<thm15:themeFamily xmlns:thm15="http://schemas.microsoft.com/office/thememl/2012/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:x="urn:example:future" x:opaque="keep" name="Office Theme" id="{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}" vid="{4A3C46E8-61CC-4603-A589-7422A47A8E4A}">
  <thm15:extLst><a:ext uri="urn:example:future"><x:future x:flag="keep">opaque</x:future></a:ext></thm15:extLst>
</thm15:themeFamily>"#;

fn family_value(name: &str) -> Family {
    Family::new(name, OFFICE_ID, OFFICE_VID).expect("valid theme-family value")
}

fn assert_send_sync<T: Send + Sync>() {}

fn root_open(extra_attributes: &str) -> String {
    format!(
        "<thm15:themeFamily xmlns:thm15=\"{NAMESPACE}\" xmlns:x=\"urn:theme-family-test\" name=\"Office\" id=\"{OFFICE_ID}\" vid=\"{OFFICE_VID}\"{extra_attributes}>"
    )
}

fn source_with_body(body: &str) -> Vec<u8> {
    format!("{}{body}</thm15:themeFamily>", root_open("")).into_bytes()
}

fn source_with_prolog(prefix: &str, declaration: &str) -> Vec<u8> {
    let mut source = prefix.as_bytes().to_vec();
    source.extend_from_slice(declaration.as_bytes());
    source.extend(source_with_body(""));
    source
}

fn source_with_raw_name(name: &[u8]) -> Vec<u8> {
    let mut source = format!("<thm15:themeFamily xmlns:thm15=\"{NAMESPACE}\" name=\"").into_bytes();
    source.extend_from_slice(name);
    source.extend_from_slice(format!("\" id=\"{OFFICE_ID}\" vid=\"{OFFICE_VID}\"/>").as_bytes());
    source
}

fn nested_source(depth: usize) -> Vec<u8> {
    let mut source = root_open("");
    for index in 0..depth {
        source.push_str(&format!("<x:n{index}>"));
    }
    source.push_str("<x:leaf/>");
    for index in (0..depth).rev() {
        source.push_str(&format!("</x:n{index}>"));
    }
    source.push_str("</thm15:themeFamily>");
    source.into_bytes()
}

fn source_with_root_attributes(extra_count: usize) -> Vec<u8> {
    let mut source = root_open("");
    source.pop();
    for index in 0..extra_count {
        source.push_str(&format!(" x:a{index}=\"value\""));
    }
    source.push_str("/>");
    source.into_bytes()
}

fn source_with_namespace_declarations(count: usize) -> Vec<u8> {
    let mut source = root_open("");
    source.push_str("<x:child");
    for index in 0..count {
        source.push_str(&format!(" xmlns:p{index}=\"urn:theme-family-{index}\""));
    }
    source.push_str("/></thm15:themeFamily>");
    source.into_bytes()
}

fn source_with_size(size: usize) -> Vec<u8> {
    let open = root_open("");
    let close = "</thm15:themeFamily>";
    assert!(size >= open.len() + close.len());
    let mut source = Vec::with_capacity(size);
    source.extend_from_slice(open.as_bytes());
    source.resize(size - close.len(), b' ');
    source.extend_from_slice(close.as_bytes());
    source
}

#[test]
fn native_fragment_reads_the_schema_fields_and_replays_exactly() {
    let parsed = family::read(NATIVE).expect("native theme-family fragment must parse");

    assert_eq!(parsed.name(), "Office Theme");
    assert_eq!(parsed.id().as_str(), OFFICE_ID);
    assert_eq!(parsed.variant_id().as_str(), OFFICE_VID);
    assert_eq!(parsed.source(), Some(NATIVE));
    assert_eq!(family::write(&parsed).expect("source-backed write"), NATIVE);
}

#[test]
fn detached_authoring_escapes_name_and_round_trips() {
    let authored = family_value("\t  Office & <Theme>\r\n");
    let xml = family::write(&authored).expect("detached value must write");
    let text = std::str::from_utf8(&xml).expect("theme-family XML is UTF-8");

    assert!(text.contains("name=\"&#9;  Office &amp; &lt;Theme&gt;&#13;&#10;\""));
    assert!(text.contains(NAMESPACE));
    assert_eq!(
        family::read(&xml).expect("authored value must parse"),
        authored
    );

    let empty = family_value("");
    let empty_xml = family::write(&empty).expect("empty xsd:string name must write");
    assert!(
        std::str::from_utf8(&empty_xml)
            .unwrap()
            .contains("name=\"\"")
    );
    assert_eq!(family::read(&empty_xml).unwrap(), empty);
}

#[test]
fn source_attribute_whitespace_normalizes_but_character_references_survive() {
    let literal = source_with_raw_name(b"a\tb\r\nc\rd\ne");
    let parsed = family::read(&literal).expect("literal XML attribute whitespace is valid");
    assert_eq!(parsed.name(), "a b c d e");

    let references = source_with_raw_name(br"a&#9;b&#13;c&#10;d");
    let parsed = family::read(&references).expect("XML character references are valid");
    assert_eq!(parsed.name(), "a\tb\rc\nd");
}

#[test]
fn guid_domain_is_braced_hyphenated_uppercase_hex() {
    let guid = Guid::new(OFFICE_ID).expect("valid ST_Guid must be accepted");
    assert_eq!(guid.as_str(), OFFICE_ID);
    assert_eq!(guid.to_string(), OFFICE_ID);

    let padded =
        Guid::new(format!(" \t{OFFICE_ID}\n")).expect("XML token edge whitespace is accepted");
    assert_eq!(padded.as_str(), OFFICE_ID);

    for invalid in [
        "62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F",
        "{62f939b6-93af-4db8-9c6b-d6c7dfdc589f}",
        "{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589}",
        "{62F939B6_93AF_4DB8_9C6B_D6C7DFDC589F}",
        "{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589G}",
        "{62F939B6-93AF-4C6B-D6C7DFDC589F}",
        "{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}\u{00A0}",
        "{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}\u{0B}",
        "{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}\u{2003}",
        "{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F} inner",
    ] {
        assert!(
            Guid::new(invalid).is_err(),
            "invalid GUID accepted: {invalid}"
        );
    }

    for whitespace in ['\u{00A0}', '\u{0B}', '\u{2003}'] {
        let source = format!(
            "<thm15:themeFamily xmlns:thm15=\"{NAMESPACE}\" name=\"Office\" id=\"{OFFICE_ID}{whitespace}\" vid=\"{OFFICE_VID}\"/>"
        );
        assert!(
            family::read(source.as_bytes()).is_err(),
            "accepted non-XML-token GUID whitespace U+{:04X}",
            u32::from(whitespace)
        );
    }
}

#[test]
fn namespace_and_required_attribute_grammar_is_strict() {
    let default_namespace = format!(
        "<themeFamily xmlns=\"{NAMESPACE}\" name=\"Office\" id=\"{OFFICE_ID}\" vid=\"{OFFICE_VID}\"/>"
    );
    assert!(family::read(default_namespace.as_bytes()).is_ok());

    let cases = [
        format!(
            "<themeFamily xmlns=\"{DRAWINGML_NAMESPACE}\" name=\"Office\" id=\"{OFFICE_ID}\" vid=\"{OFFICE_VID}\"/>"
        ),
        format!(
            "<thm15:other xmlns:thm15=\"{NAMESPACE}\" name=\"Office\" id=\"{OFFICE_ID}\" vid=\"{OFFICE_VID}\"/>"
        ),
        format!(
            "<thm15:themeFamily xmlns:thm15=\"{NAMESPACE}\" name=\"Office\" id=\"{OFFICE_ID}\"/>"
        ),
        format!(
            "<thm15:themeFamily xmlns:thm15=\"{NAMESPACE}\" thm15:name=\"confused\" id=\"{OFFICE_ID}\" vid=\"{OFFICE_VID}\"/>"
        ),
        format!(
            "<themeFamily xmlns=\"{XML_NAMESPACE}\" name=\"Office\" id=\"{OFFICE_ID}\" vid=\"{OFFICE_VID}\"/>"
        ),
        format!(
            "<thm15:themeFamily xmlns:thm15=\"{NAMESPACE}\" xmlns:x=\"{XML_NAMESPACE}\" name=\"Office\" id=\"{OFFICE_ID}\" vid=\"{OFFICE_VID}\"><x:child/></thm15:themeFamily>"
        ),
        format!(
            "<thm15:themeFamily xmlns:thm15=\"{NAMESPACE}\" xmlns:x=\"http:&#x2F;&#x2F;www.w3.org&#x2F;XML&#x2F;1998&#x2F;namespace\" name=\"Office\" id=\"{OFFICE_ID}\" vid=\"{OFFICE_VID}\"><x:child/></thm15:themeFamily>"
        ),
    ];
    for source in cases {
        assert!(
            family::read(source.as_bytes()).is_err(),
            "accepted invalid source: {source}"
        );
    }

    let qualified_lookalike = format!(
        "<thm15:themeFamily xmlns:thm15=\"{NAMESPACE}\" name=\"Office\" id=\"{OFFICE_ID}\" vid=\"{OFFICE_VID}\" thm15:name=\"opaque\"/>"
    );
    let qualified = family::read(qualified_lookalike.as_bytes()).unwrap_or_else(|error| {
        panic!(
            "qualified lookalike should remain opaque beside the required unqualified name: {error:?}"
        )
    });
    assert_eq!(qualified.name(), "Office");

    let foreign_ext_list = format!(
        "<thm15:themeFamily xmlns:thm15=\"{NAMESPACE}\" xmlns:a=\"{DRAWINGML_NAMESPACE}\" name=\"Office\" id=\"{OFFICE_ID}\" vid=\"{OFFICE_VID}\"><a:extLst/></thm15:themeFamily>"
    );
    let _ = family::read(foreign_ext_list.as_bytes()).unwrap_or_else(|error| {
        panic!("foreign direct root children remain opaque: {error:?}; source={foreign_ext_list}")
    });

    let duplicate_name = format!(
        "<thm15:themeFamily xmlns:thm15=\"{NAMESPACE}\" name=\"Office\" name=\"Other\" id=\"{OFFICE_ID}\" vid=\"{OFFICE_VID}\"/>"
    );
    assert!(family::read(duplicate_name.as_bytes()).is_err());
}

#[test]
fn malformed_nested_qnames_comments_declarations_and_references_fail_closed() {
    for source in [
        source_with_body("<x:child></x:other>"),
        source_with_body("<q:child/>"),
        source_with_body("<1bad/>"),
        source_with_body("<x:bad:name/>"),
        source_with_body("<x:child q:attr=\"value\"/>"),
        source_with_body("<x:child>&foo;</x:child>"),
        source_with_body("<x:child>&#xZZ;</x:child>"),
        source_with_body("<x:child>bad & text</x:child>"),
        source_with_body("<!-- bad -- comment -->"),
    ] {
        assert!(
            family::read(&source).is_err(),
            "accepted malformed nested source: {}",
            String::from_utf8_lossy(&source)
        );
    }

    let mut declaration = b"<?xml version=\"1.0\"?>".to_vec();
    declaration.extend(source_with_body(""));
    assert!(family::read(&declaration).is_ok());

    let mut late_declaration = source_with_body("");
    late_declaration.extend_from_slice(b"<?xml version=\"1.0\"?>");
    assert!(family::read(&late_declaration).is_err());

    for forbidden in ["<?future?>", "<!DOCTYPE themeFamily>"] {
        let source = source_with_body(forbidden);
        assert!(family::read(&source).is_err(), "accepted {forbidden}");
    }

    let commented = source_with_body("<!-- preserved --><x:child/><!-- tail -->");
    let parsed = family::read(&commented).expect("well-formed comments are opaque markup");
    assert_eq!(family::write(&parsed).unwrap(), commented);
}

#[test]
fn raw_text_delimiter_is_rejected_but_comments_and_attributes_allow_it() {
    let raw_text = source_with_body("<x:child>]]></x:child>");
    assert!(family::read(&raw_text).is_err());

    let comment = source_with_body("<!-- ]]> -->");
    assert!(family::read(&comment).is_ok());

    let raw_attribute = source_with_body(r#"<x:child marker="]]>"></x:child>"#);
    assert!(family::read(&raw_attribute).is_ok());

    let escaped_attribute = source_with_body(r#"<x:child marker="]]&gt;"></x:child>"#);
    assert!(family::read(&escaped_attribute).is_ok());
}

#[test]
fn xml_prolog_is_single_utf8_version_10_and_first() {
    let valid = source_with_prolog("", "<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
    assert!(family::read(&valid).is_ok());
    let mut bom_valid = b"\xEF\xBB\xBF".to_vec();
    bom_valid.extend(valid);
    assert!(family::read(&bom_valid).is_ok());

    let cases = [
        (
            "duplicate declaration",
            source_with_prolog("", "<?xml version=\"1.0\"?><?xml version=\"1.0\"?>"),
        ),
        (
            "declaration after whitespace",
            source_with_prolog(" \n", "<?xml version=\"1.0\"?>"),
        ),
        (
            "declaration after comment",
            source_with_prolog("<!-- prolog comment -->", "<?xml version=\"1.0\"?>"),
        ),
        ("XML 1.1", source_with_prolog("", "<?xml version=\"1.1\"?>")),
        (
            "ISO-8859-1 encoding",
            source_with_prolog("", "<?xml version=\"1.0\" encoding=\"ISO-8859-1\"?>"),
        ),
        (
            "UTF-16 encoding",
            source_with_prolog("", "<?xml version=\"1.0\" encoding=\"UTF-16\"?>"),
        ),
    ];
    let accepted = cases
        .iter()
        .filter_map(|(label, source)| family::read(source).is_ok().then_some(*label))
        .collect::<Vec<_>>();
    assert!(
        accepted.is_empty(),
        "accepted invalid XML prologs: {accepted:?}"
    );
}

#[test]
fn comments_reject_xml_forbidden_controls() {
    let mut accepted = Vec::new();
    for control in ['\0', '\u{1}', '\u{0B}'] {
        let body = format!("<!-- before{control}after -->");
        let source = source_with_body(&body);
        if family::read(&source).is_ok() {
            accepted.push(u32::from(control));
        }
    }
    assert!(
        accepted.is_empty(),
        "accepted forbidden comment controls: {accepted:?}"
    );
}

#[test]
fn extension_list_shape_and_text_are_checked_while_foreign_children_stay_opaque() {
    let valid = format!(
        "{}<thm15:extLst><a:ext uri=\"urn:future\"><x:future x:flag=\"keep\"/></a:ext></thm15:extLst></thm15:themeFamily>",
        root_open(" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"")
    );
    assert!(family::read(valid.as_bytes()).is_ok());

    let standard_entities = source_with_body("<x:child>&amp;&#10;</x:child>");
    assert!(family::read(&standard_entities).is_ok());

    for source in [
        format!(
            "{}<thm15:extLst><a:noChild/></thm15:extLst></thm15:themeFamily>",
            root_open(&format!(" xmlns:a=\"{DRAWINGML_NAMESPACE}\""))
        ),
        format!(
            "{}<thm15:extLst><a:ext><x:future/></a:ext></thm15:extLst></thm15:themeFamily>",
            root_open(&format!(" xmlns:a=\"{DRAWINGML_NAMESPACE}\""))
        ),
        format!(
            "{}<thm15:extLst><a:ext a:uri=\"urn:future\"/></thm15:extLst></thm15:themeFamily>",
            root_open(&format!(" xmlns:a=\"{DRAWINGML_NAMESPACE}\""))
        ),
        format!(
            "{}<thm15:extLst>text</thm15:extLst></thm15:themeFamily>",
            root_open(&format!(" xmlns:a=\"{DRAWINGML_NAMESPACE}\""))
        ),
        format!(
            "{}<thm15:extLst><a:ext uri=\"urn:future\">text</a:ext></thm15:extLst></thm15:themeFamily>",
            root_open(&format!(" xmlns:a=\"{DRAWINGML_NAMESPACE}\""))
        ),
    ] {
        assert!(
            family::read(source.as_bytes()).is_err(),
            "accepted malformed extension list: {source}"
        );
    }

    assert!(family::read(source_with_body("text").as_slice()).is_err());
}

#[test]
fn opaque_extensions_and_unknown_attributes_survive_known_edits() {
    let snapshot = Snapshot::from_xml(FAMILY_WITH_OPAQUE_EXTENSION)
        .expect("theme-family extension fixture must parse");
    assert_eq!(snapshot.value().name(), "Office Theme");
    assert_eq!(snapshot.xml_bytes(), FAMILY_WITH_OPAQUE_EXTENSION);
    assert_eq!(
        family::write(snapshot.value()).unwrap(),
        FAMILY_WITH_OPAQUE_EXTENSION
    );

    let mut edit = snapshot.edit();
    edit.set_name("Edited Theme")
        .expect("name edit must validate");
    let commit = edit.commit().expect("known edit must commit");
    let changed = std::str::from_utf8(commit.snapshot().xml_bytes()).unwrap();
    assert!(changed.contains("name=\"Edited Theme\""));
    assert!(changed.contains("x:opaque=\"keep\""));
    assert!(changed.contains("<x:future x:flag=\"keep\">opaque</x:future>"));
    assert_eq!(commit.snapshot().value().name(), "Edited Theme");
}

#[test]
fn exact_no_op_inverse_and_stale_source_guards_are_enforced() {
    let snapshot = Snapshot::from_xml(FAMILY_WITH_OPAQUE_EXTENSION)
        .expect("theme-family extension fixture must parse");

    let unchanged = snapshot.edit().commit().expect("no-op must commit");
    assert_eq!(
        unchanged.snapshot().xml_bytes(),
        FAMILY_WITH_OPAQUE_EXTENSION
    );
    assert_eq!(unchanged.patch().before_xml(), FAMILY_WITH_OPAQUE_EXTENSION);
    assert_eq!(unchanged.patch().after_xml(), FAMILY_WITH_OPAQUE_EXTENSION);

    let mut edit = snapshot.edit();
    edit.set_name("Changed").unwrap();
    edit.set_id("{00000000-0000-0000-0000-000000000000}")
        .unwrap();
    edit.set_variant_id("{00000000-0000-0000-0000-000000000001}")
        .unwrap();
    let changed = edit.commit().expect("changed edit must commit");
    let reopened = Snapshot::from_xml(snapshot.xml_bytes())
        .expect("equal bytes must reopen as a separate source allocation");
    let reopened_changed = changed
        .patch()
        .apply(&reopened)
        .expect("forward patch must apply to equal reopened bytes");
    assert_eq!(reopened_changed.xml_bytes(), changed.snapshot().xml_bytes());
    let reopened_restored = changed
        .patch()
        .clone()
        .inverse()
        .apply(&reopened_changed)
        .expect("inverse patch must apply to equal reopened changed bytes");
    assert_eq!(reopened_restored.xml_bytes(), FAMILY_WITH_OPAQUE_EXTENSION);

    let restored = changed
        .patch()
        .clone()
        .inverse()
        .apply(changed.snapshot())
        .expect("inverse patch must apply to its commit");
    assert_eq!(restored.xml_bytes(), FAMILY_WITH_OPAQUE_EXTENSION);

    let stale = Snapshot::from_xml(
        std::str::from_utf8(FAMILY_WITH_OPAQUE_EXTENSION)
            .unwrap()
            .replace("Office Theme", "Other Theme")
            .into_bytes(),
    )
    .unwrap();
    assert!(changed.patch().apply(&stale).is_err());
}

#[test]
fn depth_bound_accepts_empty_leaf_at_the_limit_and_rejects_one_more() {
    let exact = nested_source(MAX_DEPTH - 2);
    let parsed = family::read(&exact).expect("depth with an empty leaf at the limit");
    assert_eq!(family::write(&parsed).unwrap(), exact);

    let over = nested_source(MAX_DEPTH - 1);
    assert!(
        family::read(&over).is_err(),
        "accepted one level over depth"
    );
}

#[test]
fn attribute_bound_accepts_exact_count_and_rejects_one_more() {
    const ROOT_ATTRIBUTE_COUNT: usize = 5; // two xmlns declarations + three required attrs
    let exact = source_with_root_attributes(MAX_ATTRIBUTES - ROOT_ATTRIBUTE_COUNT);
    assert!(family::read(&exact).is_ok());

    let over = source_with_root_attributes(MAX_ATTRIBUTES - ROOT_ATTRIBUTE_COUNT + 1);
    assert!(
        family::read(&over).is_err(),
        "accepted one attribute over limit"
    );
}

#[test]
fn namespace_declaration_bound_accepts_exact_count_and_rejects_one_more() {
    let exact = source_with_namespace_declarations(MAX_NAMESPACE_DECLARATIONS);
    assert!(family::read(&exact).is_ok());

    let over = source_with_namespace_declarations(MAX_NAMESPACE_DECLARATIONS + 1);
    assert!(
        family::read(&over).is_err(),
        "accepted one namespace declaration over limit"
    );
}

#[test]
fn namespace_uri_bound_accepts_exact_bytes_and_rejects_one_more() {
    let exact_uri = "u".repeat(MAX_NAMESPACE_BYTES);
    let exact = format!(
        "<thm15:themeFamily xmlns:thm15=\"{NAMESPACE}\" xmlns:x=\"{exact_uri}\" name=\"Office\" id=\"{OFFICE_ID}\" vid=\"{OFFICE_VID}\"><x:child/></thm15:themeFamily>"
    );
    assert!(family::read(exact.as_bytes()).is_ok());

    let over_uri = "u".repeat(MAX_NAMESPACE_BYTES + 1);
    let over = format!(
        "<thm15:themeFamily xmlns:thm15=\"{NAMESPACE}\" xmlns:x=\"{over_uri}\" name=\"Office\" id=\"{OFFICE_ID}\" vid=\"{OFFICE_VID}\"><x:child/></thm15:themeFamily>"
    );
    assert!(
        family::read(over.as_bytes()).is_err(),
        "accepted oversized URI"
    );
}

#[test]
fn namespace_prefix_bound_accepts_exact_bytes_and_rejects_one_more() {
    for length in [MAX_NAMESPACE_BYTES, MAX_NAMESPACE_BYTES + 1] {
        let prefix = "p".repeat(length);
        for body in [
            String::new(),
            format!("<{prefix}:child/>"),
            format!("<x:child {prefix}:attr=\"value\"/>"),
        ] {
            let xml = format!(
                "<thm15:themeFamily xmlns:thm15=\"{NAMESPACE}\" xmlns:x=\"urn:child\" xmlns:{prefix}=\"urn:prefix\" name=\"Office\" id=\"{OFFICE_ID}\" vid=\"{OFFICE_VID}\">{body}</thm15:themeFamily>"
            );
            let result = family::read(xml.as_bytes());
            if length == MAX_NAMESPACE_BYTES {
                assert!(result.is_ok(), "exact prefix boundary: {result:?}");
            } else {
                assert!(
                    matches!(
                        result,
                        Err(litchi_drawingml::Error::Limit {
                            resource: "theme family namespace prefix bytes",
                            limit: MAX_NAMESPACE_BYTES,
                        })
                    ),
                    "oversized declared prefix: {result:?}"
                );
            }
        }
    }
}

#[test]
fn name_and_source_output_bounds_accept_exact_and_reject_over() {
    let exact_name = "N".repeat(MAX_NAME_BYTES);
    let exact_value = Family::new(&exact_name, OFFICE_ID, OFFICE_VID).unwrap();
    let exact_xml = family::write(&exact_value).unwrap();
    assert_eq!(family::read(&exact_xml).unwrap().name(), exact_name);

    let over_name = "N".repeat(MAX_NAME_BYTES + 1);
    assert!(Family::new(&over_name, OFFICE_ID, OFFICE_VID).is_err());

    let exact_source = source_with_size(MAX_XML_BYTES);
    let mut parsed = family::read(&exact_source).expect("exact source bound must parse");
    assert_eq!(family::write(&parsed).unwrap(), exact_source);

    parsed.set_name("Office+").unwrap();
    assert!(
        family::write(&parsed).is_err(),
        "accepted oversized patched output"
    );
    let over_source = source_with_size(MAX_XML_BYTES + 1);
    assert!(
        family::read(&over_source).is_err(),
        "accepted oversized source"
    );
}

#[test]
fn invalid_edits_do_not_leak_into_a_transaction() {
    let snapshot = Snapshot::from_xml(NATIVE).unwrap();
    let mut edit = snapshot.edit();
    assert!(edit.set_id("not-a-guid").is_err());
    assert!(edit.set_variant_id("{bad}").is_err());
    assert!(edit.set_name("bad\u{1}").is_err());
    assert!(!edit.is_changed());
    assert_eq!(edit.value(), snapshot.value());
}

#[test]
fn bounded_malformed_inputs_fail_closed() {
    let oversized = vec![b'x'; MAX_XML_BYTES + 1];
    assert!(family::read(&oversized).is_err());
    assert!(Snapshot::from_xml(oversized).is_err());
    assert!(family::read(br#"<thm15:themeFamily"#).is_err());
    assert!(family::read(br#"<thm15:themeFamily xmlns:thm15="http://schemas.microsoft.com/office/thememl/2012/main" name="Office" id="{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}" vid="{4A3C46E8-61CC-4603-A589-7422A47A8E4A}"><a:extLst/>"#).is_err());

    let mut deep = format!(
        "<thm15:themeFamily xmlns:thm15=\"{NAMESPACE}\" xmlns:a=\"{DRAWINGML_NAMESPACE}\" xmlns:x=\"urn:future\" name=\"Office\" id=\"{OFFICE_ID}\" vid=\"{OFFICE_VID}\"><thm15:extLst><a:ext uri=\"future\">"
    );
    for index in 0..(MAX_DEPTH + 1) {
        deep.push_str(&format!("<x:n{index}>"));
    }
    for index in (0..(MAX_DEPTH + 1)).rev() {
        deep.push_str(&format!("</x:n{index}>"));
    }
    deep.push_str("</a:ext></thm15:extLst></thm15:themeFamily>");
    assert!(family::read(deep.as_bytes()).is_err());
}

#[test]
fn immutable_snapshots_are_send_sync() {
    assert_send_sync::<Snapshot>();
    assert_send_sync::<Family>();
    assert_send_sync::<Guid>();
}
