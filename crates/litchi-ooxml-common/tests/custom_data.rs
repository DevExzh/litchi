#![allow(
    clippy::expect_used,
    clippy::indexing_slicing,
    clippy::unwrap_used,
    reason = "focused XML fixtures make source spans and rejected probes explicit"
)]

//! Public and source-bound coverage for the X14 Custom Data Properties leaf.
//!
//! The fixtures follow [MS-XLSX] §§2.1.2–2.1.3, 2.4.35, and 2.6.66: the
//! properties root is `x14:datastoreItem`, its required value is an
//! `ST_Xstring` `id`, and its only direct child is one `x14:extLst`.

use litchi_ooxml_common::custom_data::codec::{
    Limits, parse_properties_with_limits, rewrite_extension_list,
    rewrite_extension_list_with_limits, rewrite_id, rewrite_id_with_limits,
    validate_source_properties, write_properties_with_limits,
};
use litchi_ooxml_common::custom_data::{
    ExtensionList, Properties, parse_properties, write_properties,
};
use litchi_opc::{OpcPackage, PackURI, PackageWriter, XmlPart};

const X14: &str = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT_SML: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
const PROPERTIES_URI: &str = "/xl/customData/item1.xml";
const PROPERTIES_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.customDataProperties+xml";

fn source_part(xml: &[u8]) -> litchi_opc::OwnedXmlPart {
    let name = PackURI::new(PROPERTIES_URI).expect("properties URI is valid");
    let mut package = OpcPackage::new();
    package.add_part(Box::new(XmlPart::new(
        name.clone(),
        PROPERTIES_CONTENT_TYPE.to_owned(),
        xml.to_vec(),
    )));
    let bytes = PackageWriter::to_bytes(&package).expect("fixture package writes");
    let package = OpcPackage::from_bytes(&bytes).expect("fixture package reopens");
    package
        .source_xml_part(&name)
        .expect("reopened part retains exact XML provenance")
}

fn source_with_extension() -> (Vec<u8>, litchi_opc::OwnedXmlPart) {
    const COMPACT: &[u8] = br#"<d:datastoreItem xmlns:d="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:s="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:v="urn:vendor" id="seed"><d:extLst><s:ext uri="urn:seed"><v:opaque/></s:ext></d:extLst></d:datastoreItem>"#;
    const ENRICHED: &[u8] = br#"<d:datastoreItem xmlns:d="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:s="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:v="urn:vendor" id='old&#45;id'><!-- before-root --><?root-pi?><!-- before-extension --><d:extLst><?extension-pi?><s:ext uri='urn:old'><v:opaque v:raw='&amp;raw'><v:leaf><![CDATA[<? cdata-looking text ]]></v:leaf><!-- opaque-comment --></v:opaque></s:ext></d:extLst><!-- after-extension --></d:datastoreItem>"#;
    let compact = source_part(COMPACT);
    let opening_end = COMPACT
        .iter()
        .position(|byte| *byte == b'>')
        .expect("compact root has an opening tag")
        + 1;
    let enriched = compact
        .replace_element(0..opening_end, ENRICHED)
        .expect("compact fixture can issue an enriched source token");
    (enriched.bytes().to_vec(), enriched)
}

#[test]
fn public_surface_accepts_schema_minimum_and_retains_opaque_extension_bytes() {
    let (source, _) = source_with_extension();
    let properties = parse_properties(&source).expect("valid datastoreItem fixture");
    assert_eq!(properties.id, "old-id");
    let extension = properties
        .extension_list
        .as_ref()
        .expect("direct extLst is projected")
        .xml
        .as_slice();
    assert!(extension.starts_with(b"<d:extLst"));
    assert!(
        extension
            .windows(
                b"xmlns:d=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\"".len()
            )
            .any(|window| {
                window
                    == b"xmlns:d=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\""
            })
    );
    assert!(
        extension
            .windows(b"<?extension-pi?>".len())
            .any(|window| { window == b"<?extension-pi?>" })
    );
    assert!(
        extension
            .windows(b"<![CDATA[<? cdata-looking text ]]>".len())
            .any(|window| { window == b"<![CDATA[<? cdata-looking text ]]>" })
    );
    assert!(
        extension
            .windows(b"v:raw='&amp;raw'".len())
            .any(|window| { window == b"v:raw='&amp;raw'" })
    );

    let writable = Properties {
        id: properties.id.clone(),
        extension_list: Some(ExtensionList {
            xml: format!(
                r#"<x14:extLst xmlns:x14="{X14}" xmlns:s="{SML}"><s:ext uri="urn:writable"><v:opaque xmlns:v="urn:vendor"><![CDATA[<? retained ]]></v:opaque></s:ext></x14:extLst>"#
            )
            .into_bytes(),
        }),
    };
    let serialized = write_properties(&writable).expect("typed properties serialize");
    let reread = parse_properties(&serialized).expect("serialized properties reread");
    assert_eq!(reread, writable);
    assert!(serialized.starts_with(b"<x14:datastoreItem xmlns:x14=\""));
}

#[test]
fn extension_grammar_accepts_transitional_sml_ext_optional_uri_and_one_wildcard() {
    let source = format!(
        r#"<x14:datastoreItem xmlns:x14="{X14}" xmlns:s="{SML}" xmlns:v="urn:vendor" id="core"><x14:extLst><s:ext><v:opaque/></s:ext></x14:extLst></x14:datastoreItem>"#
    );
    let properties = parse_properties(source.as_bytes())
        .expect("CT_ExtensionList admits a core Transitional SpreadsheetML ext");
    let extension = properties
        .extension_list
        .as_ref()
        .expect("direct extension list projects")
        .xml
        .as_slice();
    assert!(extension.windows(b"<s:ext>".len()).any(|w| w == b"<s:ext>"));
    assert!(
        extension
            .windows(b"<v:opaque/>".len())
            .any(|w| w == b"<v:opaque/>")
    );
    let serialized = write_properties(&properties).expect("parsed extension is standalone-ready");
    assert_eq!(parse_properties(&serialized).unwrap(), properties);
}

#[test]
fn extension_grammar_rejects_strict_sml_direct_ext() {
    let source = format!(
        r#"<x14:datastoreItem xmlns:x14="{X14}" xmlns:strict="{STRICT_SML}" xmlns:v="urn:vendor" id="strict"><x14:extLst><strict:ext><v:opaque/></strict:ext></x14:extLst></x14:datastoreItem>"#
    );
    assert!(parse_properties(source.as_bytes()).is_err());
}

#[test]
fn empty_id_and_st_xstring_escape_round_trip_including_supplementary_scalar() {
    let values = [
        "",
        "A & < ' \"",
        "_x0041_",
        "\u{0}\u{1}\u{fffe}\u{ffff}\t\n\r😀",
    ];
    for id in values {
        let properties = Properties {
            id: id.to_owned(),
            extension_list: None,
        };
        let serialized = write_properties(&properties).expect("ST_Xstring value writes");
        let reread = parse_properties(&serialized).expect("ST_Xstring value rereads");
        assert_eq!(reread, properties, "decoded id differs for {id:?}");
    }

    let escaped = write_properties(&Properties {
        id: "_x0041_\u{1}\u{fffe}\u{ffff}😀".to_owned(),
        extension_list: None,
    })
    .expect("escaped value writes");
    assert!(
        escaped
            .windows(b"_x005F_x0041_".len())
            .any(|window| { window == b"_x005F_x0041_" })
    );
    assert!(
        escaped
            .windows(b"_x0001_".len())
            .any(|window| window == b"_x0001_")
    );
    assert!(
        escaped
            .windows(b"_xFFFE_".len())
            .any(|window| window == b"_xFFFE_")
    );
    assert!(
        escaped
            .windows(b"_xFFFF_".len())
            .any(|window| window == b"_xFFFF_")
    );
    assert!(
        escaped
            .windows("😀".len())
            .any(|window| window == "😀".as_bytes())
    );

    let source =
        format!(r#"<x14:datastoreItem xmlns:x14="{X14}" id="A &amp; _x0041_ &#x1F600;"/>"#);
    let parsed = parse_properties(source.as_bytes()).expect("source ST_Xstring parses");
    assert_eq!(parsed.id, "A & A 😀");
}

#[test]
fn raw_less_than_in_attribute_values_is_rejected_but_escaped_less_than_is_valid() {
    let escaped_id = format!(r#"<x14:datastoreItem xmlns:x14="{X14}" id="a &lt; b"/>"#);
    assert_eq!(parse_properties(escaped_id.as_bytes()).unwrap().id, "a < b");

    let raw_id = format!(r#"<x14:datastoreItem xmlns:x14="{X14}" id="a < b"/>"#);
    assert!(parse_properties(raw_id.as_bytes()).is_err());

    let escaped_opaque = format!(
        r#"<x14:datastoreItem xmlns:x14="{X14}" xmlns:s="{SML}" xmlns:v="urn:vendor" id="x"><x14:extLst><s:ext><v:opaque value="a &lt; b"/></s:ext></x14:extLst></x14:datastoreItem>"#
    );
    assert!(parse_properties(escaped_opaque.as_bytes()).is_ok());

    let raw_opaque = format!(
        r#"<x14:datastoreItem xmlns:x14="{X14}" xmlns:s="{SML}" xmlns:v="urn:vendor" id="x"><x14:extLst><s:ext><v:opaque value="a < b"/></s:ext></x14:extLst></x14:datastoreItem>"#
    );
    assert!(parse_properties(raw_opaque.as_bytes()).is_err());
}

#[test]
fn uid_utf16_units_are_exact_for_one_supplementary_scalar() {
    let value = Properties {
        id: "😀".to_owned(),
        extension_list: None,
    };
    let serialized = write_properties(&value).expect("supplementary UID writes");
    let exact = Limits::standard().with_uid_units(2);
    assert!(write_properties_with_limits(&value, &exact).is_ok());
    assert!(parse_properties_with_limits(&serialized, &exact).is_ok());

    let source = source_part(&serialized);
    assert!(rewrite_id_with_limits(&source, "😀", &exact).is_ok());

    let one_under = exact.with_uid_units(1);
    assert!(write_properties_with_limits(&value, &one_under).is_err());
    assert!(parse_properties_with_limits(&serialized, &one_under).is_err());
    assert!(rewrite_id_with_limits(&source, "😀", &one_under).is_err());
}

#[test]
fn parser_event_limit_is_exact_and_one_under_for_minimal_properties() {
    let value = Properties {
        id: "event".to_owned(),
        extension_list: None,
    };
    let serialized = write_properties(&value).expect("baseline properties write");
    let exact = Limits::standard().with_events(2);
    assert!(parse_properties_with_limits(&serialized, &exact).is_ok());
    let one_under = exact.with_events(1);
    assert!(parse_properties_with_limits(&serialized, &one_under).is_err());
}

#[test]
fn lowered_limits_accept_exact_boundaries_and_reject_one_under() {
    let value = Properties {
        id: "abcd".to_owned(),
        extension_list: None,
    };
    let serialized = write_properties(&value).expect("baseline properties write");

    let exact_properties = Limits::standard().with_properties_xml_bytes(serialized.len());
    assert!(write_properties_with_limits(&value, &exact_properties).is_ok());
    assert!(parse_properties_with_limits(&serialized, &exact_properties).is_ok());
    let under_properties = exact_properties.with_properties_xml_bytes(serialized.len() - 1);
    assert!(write_properties_with_limits(&value, &under_properties).is_err());
    assert!(parse_properties_with_limits(&serialized, &under_properties).is_err());

    let exact_string = Limits::standard().with_string_bytes(X14.len());
    assert!(parse_properties_with_limits(&serialized, &exact_string).is_ok());
    let under_string = exact_string.with_string_bytes(X14.len() - 1);
    assert!(parse_properties_with_limits(&serialized, &under_string).is_err());
    let exact_id_write = Limits::standard().with_string_bytes(value.id.len());
    assert!(write_properties_with_limits(&value, &exact_id_write).is_ok());
    let under_id_write = exact_id_write.with_string_bytes(value.id.len() - 1);
    assert!(write_properties_with_limits(&value, &under_id_write).is_err());

    let exact_nodes = Limits::standard().with_nodes(1);
    assert!(parse_properties_with_limits(&serialized, &exact_nodes).is_ok());
    let under_nodes = exact_nodes.with_nodes(0);
    assert!(parse_properties_with_limits(&serialized, &under_nodes).is_err());

    let exact_depth = Limits::standard().with_depth(1);
    assert!(parse_properties_with_limits(&serialized, &exact_depth).is_ok());
    let under_depth = exact_depth.with_depth(0);
    assert!(parse_properties_with_limits(&serialized, &under_depth).is_err());

    let exact_namespace = Limits::standard().with_namespace_bytes(3 + X14.len());
    assert!(parse_properties_with_limits(&serialized, &exact_namespace).is_ok());
    let under_namespace = exact_namespace.with_namespace_bytes(3 + X14.len() - 1);
    assert!(parse_properties_with_limits(&serialized, &under_namespace).is_err());

    let exact_attributes = Limits::standard().with_attributes(2);
    assert!(parse_properties_with_limits(&serialized, &exact_attributes).is_ok());
    let under_attributes = exact_attributes.with_attributes(1);
    assert!(parse_properties_with_limits(&serialized, &under_attributes).is_err());
}

#[test]
fn extension_and_source_rewrite_limits_use_exact_output_boundaries() {
    let extension_xml = format!(
        r#"<x14:extLst xmlns:x14="{X14}" xmlns:s="{SML}"><s:ext uri="urn:cap"><v:opaque xmlns:v="urn:vendor"/></s:ext></x14:extLst>"#
    )
    .into_bytes();
    let value = Properties {
        id: "cap".to_owned(),
        extension_list: Some(ExtensionList {
            xml: extension_xml.clone(),
        }),
    };
    let serialized = write_properties(&value).expect("baseline extension write");
    let exact_extension = Limits::standard().with_extension_xml_bytes(extension_xml.len());
    assert!(write_properties_with_limits(&value, &exact_extension).is_ok());
    assert!(parse_properties_with_limits(&serialized, &exact_extension).is_ok());
    let under_extension = exact_extension.with_extension_xml_bytes(extension_xml.len() - 1);
    assert!(write_properties_with_limits(&value, &under_extension).is_err());
    assert!(parse_properties_with_limits(&serialized, &under_extension).is_err());

    let source_xml = format!(r#"<x14:datastoreItem xmlns:x14="{X14}" id="a"/>"#);
    let source = source_part(source_xml.as_bytes());
    let rewritten_id = rewrite_id(&source, "longer-id").expect("baseline id rewrite");
    let exact_id_output = Limits::standard().with_properties_xml_bytes(rewritten_id.bytes().len());
    assert!(rewrite_id_with_limits(&source, "longer-id", &exact_id_output).is_ok());
    let under_id_output = exact_id_output.with_properties_xml_bytes(rewritten_id.bytes().len() - 1);
    assert!(rewrite_id_with_limits(&source, "longer-id", &under_id_output).is_err());

    let source_xml = format!(
        r#"<x14:datastoreItem xmlns:x14="{X14}" xmlns:s="{SML}" id="a"><x14:extLst><s:ext uri="urn:old"><v:opaque xmlns:v="urn:vendor"/></s:ext></x14:extLst></x14:datastoreItem>"#
    );
    let source = source_part(source_xml.as_bytes());
    let replacement = ExtensionList { xml: extension_xml };
    let rewritten_extension =
        rewrite_extension_list(&source, Some(&replacement)).expect("baseline extension rewrite");
    let exact_output =
        Limits::standard().with_properties_xml_bytes(rewritten_extension.bytes().len());
    assert!(rewrite_extension_list_with_limits(&source, Some(&replacement), &exact_output).is_ok());
    let under_output =
        exact_output.with_properties_xml_bytes(rewritten_extension.bytes().len() - 1);
    assert!(
        rewrite_extension_list_with_limits(&source, Some(&replacement), &under_output).is_err()
    );
}

#[test]
fn namespace_aliases_and_entity_escaped_uri_resolve_by_expanded_name() {
    let escaped_uri = X14.replace('/', "&#x2F;");
    let source = format!(
        r#"<alias:datastoreItem xmlns:alias="{escaped_uri}" xmlns:s="{SML}" id="alias-id"><alias:extLst><s:ext uri="urn:alias"><foreign:opaque xmlns:foreign="urn:foreign"><foreign:leaf/></foreign:opaque></s:ext></alias:extLst></alias:datastoreItem>"#
    );
    let properties = parse_properties(source.as_bytes()).expect("aliased X14 names parse");
    assert_eq!(properties.id, "alias-id");
    assert!(properties.extension_list.is_some());
}

#[test]
fn inherited_bindings_cover_opaque_descendants_and_replacement_fragments() {
    let source = format!(
        r#"<datastoreItem xmlns="{X14}" xmlns:s="{SML}" xmlns:v="urn:vendor" id="old"><extLst><s:ext uri="urn:old"><v:opaque><v:leaf><![CDATA[<? legal ]]></v:leaf></v:opaque></s:ext></extLst></datastoreItem>"#
    );
    let parsed = parse_properties(source.as_bytes()).expect("default X14 binding parses");
    let fragment = ExtensionList {
        xml: br#"<extLst xmlns:s="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><s:ext uri="urn:new"><v:opaque xmlns:v="urn:vendor"><v:leaf/></v:opaque></s:ext></extLst>"#
            .to_vec(),
    };
    let source_part = source_part(source.as_bytes());
    let replaced = rewrite_extension_list(&source_part, Some(&fragment))
        .expect("fragment may use bindings inherited from its insertion root");
    assert!(
        replaced
            .bytes()
            .windows(b"urn:new".len())
            .any(|window| window == b"urn:new")
    );
    assert!(
        !replaced
            .bytes()
            .windows(b"urn:old".len())
            .any(|window| window == b"urn:old")
    );
    assert_eq!(parse_properties(replaced.bytes()).unwrap().id, "old");
    let replaced_properties = parse_properties(replaced.bytes()).unwrap();
    let replaced_extension = replaced_properties
        .extension_list
        .as_ref()
        .expect("replacement extLst remains projected");
    assert!(
        replaced_extension
            .xml
            .windows(b"urn:new".len())
            .any(|window| window == b"urn:new")
    );
    assert!(
        replaced_extension
            .xml
            .windows(b"xmlns:v=\"urn:vendor\"".len())
            .any(|window| window == b"xmlns:v=\"urn:vendor\"")
    );
    assert_eq!(parsed.id, "old");

    let serialized = write_properties(&parsed).expect("parsed properties can be serialized");
    assert_eq!(
        parse_properties(&serialized).expect("serialized parsed properties reread"),
        parsed
    );
}

#[test]
fn identical_raw_extensions_keep_their_own_inherited_namespace_context() {
    let first = format!(
        r#"<x:datastoreItem xmlns:x="{X14}" xmlns:s="{SML}" xmlns:v="urn:first-context" id="first"><x:extLst><s:ext uri="urn:raw"><v:opaque/></s:ext></x:extLst></x:datastoreItem>"#
    );
    let second = format!(
        r#"<x:datastoreItem xmlns:x="{X14}" xmlns:s="{SML}" xmlns:v="urn:second-context" id="second"><x:extLst><s:ext uri="urn:raw"><v:opaque/></s:ext></x:extLst></x:datastoreItem>"#
    );
    let first_properties = parse_properties(first.as_bytes()).expect("first source parses");
    let second_properties = parse_properties(second.as_bytes()).expect("second source parses");
    let first_extension = first_properties
        .extension_list
        .as_ref()
        .expect("first extension projects");
    let second_extension = second_properties
        .extension_list
        .as_ref()
        .expect("second extension projects");
    assert!(
        first_extension
            .xml
            .windows(b"xmlns:v=\"urn:first-context\"".len())
            .any(|window| window == b"xmlns:v=\"urn:first-context\"")
    );
    assert!(
        !first_extension
            .xml
            .windows(b"urn:second-context".len())
            .any(|window| { window == b"urn:second-context" })
    );
    assert!(
        second_extension
            .xml
            .windows(b"xmlns:v=\"urn:second-context\"".len())
            .any(|window| window == b"xmlns:v=\"urn:second-context\"")
    );
    assert!(
        !second_extension
            .xml
            .windows(b"urn:first-context".len())
            .any(|window| { window == b"urn:first-context" })
    );

    let serialized_first =
        write_properties(&first_properties).expect("first parsed fragment is standalone-ready");
    let reread_first =
        parse_properties(&serialized_first).expect("serialized first properties reread");
    assert_eq!(reread_first, first_properties);
    assert!(
        serialized_first
            .windows(b"urn:first-context".len())
            .any(|window| { window == b"urn:first-context" })
    );
    assert!(
        !serialized_first
            .windows(b"urn:second-context".len())
            .any(|window| { window == b"urn:second-context" })
    );
}

#[test]
fn reserved_namespace_bindings_and_empty_prefixed_bindings_are_rejected() {
    let cases = [
        format!(r#"<x14:datastoreItem xmlns:x14="{X14}" xmlns:xmlns="urn:bad" id="x"/>"#),
        format!(
            r#"<x14:datastoreItem xmlns:x14="{X14}" xmlns:p="http://www.w3.org/2000/xmlns/" id="x"/>"#
        ),
        format!(r#"<x14:datastoreItem xmlns:x14="{X14}" xmlns:xml="urn:bad" id="x"/>"#),
        format!(r#"<x14:datastoreItem xmlns:x14="{X14}" xmlns:p="" id="x"/>"#),
    ];
    for source in cases {
        assert!(
            parse_properties(source.as_bytes()).is_err(),
            "accepted {source}"
        );
    }
}

#[test]
fn source_validation_rejects_unbound_prefix_without_real_context() {
    let value = Properties {
        id: "source".to_owned(),
        extension_list: Some(ExtensionList {
            xml: format!(
                r#"<x14:extLst xmlns:x14="{X14}" xmlns:s="{SML}"><s:ext><p:opaque/></s:ext></x14:extLst>"#
            )
            .into_bytes(),
        }),
    };
    assert!(validate_source_properties(&value).is_err());
}

#[test]
fn full_documents_allow_bom_and_declaration_but_embedded_fragments_do_not() {
    let full = format!(
        "\u{feff}<?xml version=\"1.0\" encoding=\"UTF-8\"?><x14:datastoreItem xmlns:x14=\"{X14}\" id=\"bom\"/>"
    );
    assert_eq!(parse_properties(full.as_bytes()).unwrap().id, "bom");

    let declaration = ExtensionList {
        xml: format!(
            "<?xml version=\"1.0\"?><x14:extLst xmlns:x14=\"{X14}\" xmlns:s=\"{SML}\"><s:ext uri=\"urn:x\"><v:opaque xmlns:v=\"urn:vendor\"/></s:ext></x14:extLst>"
        )
        .into_bytes(),
    };
    assert!(
        write_properties(&Properties {
            id: "x".to_owned(),
            extension_list: Some(declaration),
        })
        .is_err()
    );

    let bom = ExtensionList {
        xml: format!(
            "\u{feff}<x14:extLst xmlns:x14=\"{X14}\" xmlns:s=\"{SML}\"><s:ext uri=\"urn:x\"><v:opaque xmlns:v=\"urn:vendor\"/></s:ext></x14:extLst>"
        )
        .into_bytes(),
    };
    assert!(
        write_properties(&Properties {
            id: "x".to_owned(),
            extension_list: Some(bom),
        })
        .is_err()
    );
}

#[test]
fn malformed_whole_documents_and_direct_container_grammar_fail_closed() {
    let valid = format!(r#"<x14:datastoreItem xmlns:x14="{X14}" id="x"/>"#);
    let cases = [
        format!("text{valid}"),
        format!("{valid}tail"),
        format!(
            r#"<x14:datastoreItem xmlns:x14="{X14}" xmlns:s="{SML}" id="x"><x14:extLst><s:ext uri="urn:x"><v:opaque xmlns:v="urn:vendor"/></x14:extLst></x14:datastoreItem>"#
        ),
        format!(r#"<x14:datastoreItem xmlns:x14="{X14}" id="x"><x14:other/></x14:datastoreItem>"#),
        format!(
            r#"<x14:datastoreItem xmlns:x14="{X14}" id="x"><![CDATA[text]]></x14:datastoreItem>"#
        ),
        format!(r#"<x14:datastoreItem xmlns:x14="{X14}" id="x">&amp;</x14:datastoreItem>"#),
        format!(
            r#"<x14:datastoreItem xmlns:x14="{X14}" id="x"><x14:extLst foo="x"/></x14:datastoreItem>"#
        ),
        format!(
            r#"<x14:datastoreItem xmlns:x14="{X14}" id="x"><x14:extLst><x14:other/></x14:extLst></x14:datastoreItem>"#
        ),
        format!(
            r#"<x14:datastoreItem xmlns:x14="{X14}" xmlns:s="{SML}" id="x"><x14:extLst><s:ext/></x14:extLst></x14:datastoreItem>"#
        ),
        format!(
            r#"<x14:datastoreItem xmlns:x14="{X14}" xmlns:s="{SML}" xmlns:v="urn:vendor" id="x"><x14:extLst><s:ext><v:first/><v:second/></s:ext></x14:extLst></x14:datastoreItem>"#
        ),
        format!(
            r#"<x14:datastoreItem xmlns:x14="{X14}" id="x"/><x14:datastoreItem xmlns:x14="{X14}" id="y"/>"#
        ),
        format!(r#"<!DOCTYPE datastoreItem><x14:datastoreItem xmlns:x14="{X14}" id="x"/>"#),
    ];
    for source in cases {
        assert!(
            parse_properties(source.as_bytes()).is_err(),
            "accepted {source}"
        );
    }
}

#[test]
fn non_xml_nbsp_is_not_accepted_as_container_whitespace() {
    let root_text =
        format!("<x14:datastoreItem xmlns:x14=\"{X14}\" id=\"x\">\u{00a0}</x14:datastoreItem>");
    assert!(parse_properties(root_text.as_bytes()).is_err());

    let ext_list_text = format!(
        r#"<x14:extLst xmlns:x14="{X14}" xmlns:s="{SML}" xmlns:v="urn:vendor"> <s:ext><v:opaque/></s:ext></x14:extLst>"#
    );
    assert!(
        write_properties(&Properties {
            id: "x".to_owned(),
            extension_list: Some(ExtensionList {
                xml: ext_list_text.into_bytes(),
            }),
        })
        .is_err()
    );

    let ext_text = format!(
        r#"<x14:extLst xmlns:x14="{X14}" xmlns:s="{SML}" xmlns:v="urn:vendor"><s:ext> <v:opaque/></s:ext></x14:extLst>"#
    );
    assert!(
        write_properties(&Properties {
            id: "x".to_owned(),
            extension_list: Some(ExtensionList {
                xml: ext_text.into_bytes(),
            }),
        })
        .is_err()
    );
}

#[test]
fn duplicate_required_attributes_and_direct_extension_owners_are_rejected() {
    let cases = [
        format!(r#"<x14:datastoreItem xmlns:x14="{X14}" id="a" id="b"/>"#),
        format!(
            r#"<x14:datastoreItem xmlns:x14="{X14}" id="a"><x14:extLst/><x14:extLst/></x14:datastoreItem>"#
        ),
        format!(
            r#"<x14:datastoreItem xmlns:x14="{X14}" xmlns:s="{SML}" id="a"><x14:extLst><s:ext uri="a" uri="b"><v:opaque xmlns:v="urn:vendor"/></s:ext></x14:extLst></x14:datastoreItem>"#
        ),
    ];
    for source in cases {
        assert!(
            parse_properties(source.as_bytes()).is_err(),
            "accepted {source}"
        );
    }
}

#[test]
fn source_rewrites_change_only_the_requested_root_id_or_extension_span() {
    let (source, part) = source_with_extension();
    assert_eq!(parse_properties(&source).unwrap().id, "old-id");

    let renamed = rewrite_id(&part, "new-id").expect("source id rewrites");
    let expected = String::from_utf8_lossy(&source)
        .replace("old&#45;id", "new-id")
        .into_bytes();
    assert_eq!(renamed.bytes(), expected.as_slice());
    assert!(
        renamed
            .bytes()
            .windows(b"<!-- before-root -->".len())
            .any(|window| { window == b"<!-- before-root -->" })
    );
    assert!(
        renamed
            .bytes()
            .windows(b"<?root-pi?>".len())
            .any(|window| window == b"<?root-pi?>")
    );
    assert!(
        renamed
            .bytes()
            .windows(b"<![CDATA[<? cdata-looking text ]]>".len())
            .any(|window| { window == b"<![CDATA[<? cdata-looking text ]]>" })
    );

    let replacement = ExtensionList {
        xml: br#"<d:extLst xmlns:d="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:s="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><s:ext uri="urn:new"><v:opaque xmlns:v="urn:vendor"><![CDATA[<? replacement ]]></v:opaque></s:ext></d:extLst>"#
            .to_vec(),
    };
    let replaced =
        rewrite_extension_list(&part, Some(&replacement)).expect("direct extension span rewrites");
    assert!(
        replaced
            .bytes()
            .windows(b"<!-- before-extension -->".len())
            .any(|window| { window == b"<!-- before-extension -->" })
    );
    assert!(
        replaced
            .bytes()
            .windows(b"<!-- after-extension -->".len())
            .any(|window| { window == b"<!-- after-extension -->" })
    );
    assert!(
        replaced
            .bytes()
            .windows(b"urn:new".len())
            .any(|window| window == b"urn:new")
    );
    assert!(
        !replaced
            .bytes()
            .windows(b"urn:old".len())
            .any(|window| window == b"urn:old")
    );
    assert!(
        replaced
            .bytes()
            .windows(b"<?extension-pi?>".len())
            .all(|window| { window != b"<?extension-pi?>" })
    );
    assert_eq!(parse_properties(replaced.bytes()).unwrap().id, "old-id");
}

#[test]
fn extension_removal_removes_the_whole_opening_range_and_keeps_siblings() {
    let (source, part) = source_with_extension();
    let removed = rewrite_extension_list(&part, None).expect("direct extension removes");
    let source_text = String::from_utf8_lossy(&source);
    let extension = "<d:extLst><?extension-pi?><s:ext uri='urn:old'><v:opaque v:raw='&amp;raw'><v:leaf><![CDATA[<? cdata-looking text ]]></v:leaf><!-- opaque-comment --></v:opaque></s:ext></d:extLst>";
    let expected = source_text.replace(extension, "");
    assert_eq!(removed.bytes(), expected.as_bytes());
    assert!(
        !removed
            .bytes()
            .windows(b"extLst".len())
            .any(|window| window == b"extLst")
    );
    assert!(
        removed
            .bytes()
            .windows(b"<!-- before-extension -->".len())
            .any(|window| { window == b"<!-- before-extension -->" })
    );
    assert!(
        removed
            .bytes()
            .windows(b"<!-- after-extension -->".len())
            .any(|window| { window == b"<!-- after-extension -->" })
    );
    assert_eq!(
        parse_properties(removed.bytes()).unwrap().extension_list,
        None
    );
}

#[test]
fn scalar_semantic_noop_preserves_lexical_source_bytes() {
    let source = format!(
        r#"<x14:datastoreItem xmlns:x14="{X14}" xmlns:s="{SML}" id="same&#45;id"><x14:extLst><s:ext uri="urn:same"><v:opaque xmlns:v="urn:vendor"/></s:ext></x14:extLst></x14:datastoreItem>"#
    );
    let part = source_part(source.as_bytes());
    let properties = parse_properties(source.as_bytes()).expect("source parses");
    let renamed = rewrite_id(&part, &properties.id).expect("semantic same id is accepted");
    assert_eq!(renamed.bytes(), part.bytes());

    let extension = properties
        .extension_list
        .as_ref()
        .expect("extension exists");
    let same_extension = rewrite_extension_list(&part, Some(extension))
        .expect("semantic same extension is accepted");
    assert_eq!(same_extension.bytes(), part.bytes());
}

#[test]
fn malformed_extension_fragments_are_rejected_before_source_splice() {
    let source = format!(r#"<x14:datastoreItem xmlns:x14="{X14}" id="x"/>"#);
    let part = source_part(source.as_bytes());
    let cases = [
        ("declaration", b"<?xml version=\"1.0\"?><x14:extLst xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\"/>".to_vec()),
        ("foreign-attribute", format!(r#"<x14:extLst xmlns:x14="{X14}" bad="x"/>"#).into_bytes()),
        ("foreign-direct-child", format!(r#"<x14:extLst xmlns:x14="{X14}"><x14:other/></x14:extLst>"#).into_bytes()),
        ("direct-text", format!(r#"<x14:extLst xmlns:x14="{X14}" xmlns:s="{SML}"><s:ext uri="x"><v:opaque xmlns:v="urn:vendor"/>text</s:ext></x14:extLst>"#).into_bytes()),
    ];
    for (label, xml) in cases {
        let extension = ExtensionList { xml };
        assert!(
            rewrite_extension_list(&part, Some(&extension)).is_err(),
            "accepted malformed {label} extension"
        );
        assert_eq!(part.bytes(), source.as_bytes());
    }
}
