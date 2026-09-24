#![allow(clippy::expect_used, reason = "regression fixture assertions")]

//! Exercise complete-part grammar independently of the family fragment codec.
use litchi_drawingml::theme::family::{Family, part};

fn theme(body: &str) -> String {
    format!(
        r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">{body}</a:theme>"#
    )
}

fn assert_rejected(source: &[u8]) {
    let family = Family::new(
        "Added",
        "{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}",
        "{4A3C46E8-61CC-4603-A589-7422A47A8E4A}",
    )
    .expect("valid family");
    assert!(part::read(source).is_err(), "read accepted {source:?}");
    assert!(
        part::read_family(source).is_err(),
        "projection accepted {source:?}"
    );
    assert!(
        part::add_family(source, &family).is_err(),
        "add accepted {source:?}"
    );
    assert!(
        part::replace_family(source, &family).is_err(),
        "replace accepted {source:?}"
    );
    assert!(
        part::remove_family(source).is_err(),
        "remove accepted {source:?}"
    );
}

#[test]
fn malformed_xml_is_rejected_even_without_a_family_owner() {
    for body in [
        "<1bad/>",
        "<1bad></1bad>",
        "<a:1bad/>",
        "<a:b:c/>",
        "&foo;",
        "\u{1}",
        "&#x1;",
        "<opaque>&foo;</opaque>",
        "<opaque>&#0;</opaque>",
        "<opaque>&#xD800;</opaque>",
        "<opaque>&#x110000;</opaque>",
        "<opaque>&unterminated</opaque>",
        "<opaque>]]></opaque>",
        "<opaque>\u{1}</opaque>",
        "<opaque><![CDATA[\u{b}]]></opaque>",
        "<!--\u{1}-->",
        "<opaque value=\"\u{1}\"/>",
        "<!-- malformed -- comment -->",
        "<?broken",
    ] {
        assert_rejected(theme(body).as_bytes());
    }
    for prefix in ["outside", "&#32;", "<![CDATA[ ]]>"] {
        assert_rejected(format!("{prefix}{}", theme("")).as_bytes());
        assert_rejected(format!("{}{prefix}", theme("")).as_bytes());
    }
    for body in ["text", "<![CDATA[abc]]>", "&#65;"] {
        assert_rejected(theme(body).as_bytes());
        assert_rejected(theme(&format!("<a:extLst>{body}</a:extLst>")).as_bytes());
        assert_rejected(
            theme(&format!(
                r#"<a:extLst><a:ext uri="{}">{body}</a:ext></a:extLst>"#,
                part::EXTENSION_URI
            ))
            .as_bytes(),
        );
    }
}

#[test]
fn declarations_require_xml_10_utf8_and_correct_placement() {
    for declaration in [
        "<?xml version=\"1.1\"?>",
        "<?xml version=\"2.0\"?>",
        "<?xml version=\"1.0\" encoding=\"UTF-16\"?>",
        "<?xml version=\"1.0\" encoding=\"ISO-8859-1\"?>",
        "<?xml encoding=\"UTF-8\"?>",
        "<?xml version=\"1.0\" encoding=\"wat\"?>",
        "<?xml ?>",
        "<?xml version=\"1.0\"version=\"1.0\"?>",
        "<?xml version=\"1.0\" standalone=\"maybe\"?>",
        " <?xml version=\"1.0\"?>",
        "<!--first--><?xml version=\"1.0\"?>",
        "<?xml version=\"1.0\"?><?xml version=\"1.0\"?>",
    ] {
        assert_rejected(format!("{declaration}{}", theme("")).as_bytes());
    }
    for declaration in [
        "",
        "<?xml version='1.0'?>",
        "<?xml version='1.0' encoding='utf-8' standalone='no'?>",
    ] {
        let source = format!("{declaration}{}", theme(""));
        assert!(part::read(source.as_bytes()).is_ok(), "rejected {source}");
    }
}

#[test]
fn valid_opaque_text_and_xml_whitespace_remain_lexically_intact() {
    for body in [
        "<opaque>&amp;&lt;&gt;&quot;&apos;&#65;&#x1F600;</opaque>",
        "<opaque><![CDATA[&foo; <1bad/>]]></opaque>",
        "<a:extLst> \t\r\n&#32;<![CDATA[\t]]></a:extLst>",
    ] {
        let source = format!("\u{feff}{}\n", theme(body));
        let snapshot = part::read(source.as_bytes()).expect("valid opaque XML");
        assert_eq!(snapshot.xml_bytes(), source.as_bytes());
        assert_eq!(
            snapshot.remove_family().expect("absent no-op"),
            source.as_bytes()
        );
    }
}

#[test]
fn namespace_prefix_byte_limit_covers_declarations_elements_and_attributes() {
    use litchi_drawingml::theme::family::MAX_NAMESPACE_BYTES;

    for length in [MAX_NAMESPACE_BYTES, MAX_NAMESPACE_BYTES + 1] {
        let prefix = "p".repeat(length);
        for body in [
            format!(r#"<opaque xmlns:{prefix}="urn:vendor"/>"#),
            format!(r#"<{prefix}:opaque xmlns:{prefix}="urn:vendor"/>"#),
            format!(r#"<opaque xmlns:{prefix}="urn:vendor" {prefix}:attr="value"/>"#),
            format!(
                r#"<opaque xmlns:{prefix}="urn:vendor"><{prefix}:child {prefix}:attr="value"/></opaque>"#
            ),
        ] {
            let source = theme(&body);
            if length == MAX_NAMESPACE_BYTES {
                assert!(
                    part::read(source.as_bytes()).is_ok(),
                    "exact bound rejected"
                );
            } else {
                assert_rejected(source.as_bytes());
            }
        }
    }
    // Undeclared prefixes must be bounded before lookup/error construction too.
    let prefix = "p".repeat(MAX_NAMESPACE_BYTES + 1);
    for body in [
        format!("<{prefix}:opaque/>"),
        format!("<opaque {prefix}:attr='v'/>"),
    ] {
        let source = theme(&body);
        let error = part::read(source.as_bytes()).expect_err("oversized undeclared prefix");
        assert!(
            error.to_string().contains("namespace prefix bytes"),
            "{error}"
        );
    }
}

#[test]
fn inherited_namespace_decoration_respects_fragment_output_limit() {
    use litchi_drawingml::theme::family::{MAX_XML_BYTES, NAMESPACE};
    let opening = format!(
        r#"<f:themeFamily xmlns:f="{NAMESPACE}" name="A" id="{{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}}" vid="{{4A3C46E8-61CC-4603-A589-7422A47A8E4A}}"><!--"#
    );
    let closing = "--></f:themeFamily>";
    let inherited = r#" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main""#;
    let padding = MAX_XML_BYTES - opening.len() - closing.len() - inherited.len();
    for extra in [0, 1] {
        let fragment = format!("{opening}{}{closing}", "x".repeat(padding + extra));
        assert!(fragment.len() < MAX_XML_BYTES);
        let source = theme(&format!(
            r#"<a:extLst><a:ext uri="{}">{fragment}</a:ext></a:extLst>"#,
            part::EXTENSION_URI
        ));
        if extra == 0 {
            assert!(
                part::read(source.as_bytes()).is_ok(),
                "exact decorated bound"
            );
        } else {
            let error = part::read(source.as_bytes()).expect_err("decoration exceeds bound");
            assert!(
                error
                    .to_string()
                    .contains("decorated Theme Family XML bytes"),
                "{error}"
            );
        }
    }
}

#[test]
fn insertion_wrappers_obey_exact_complete_part_output_limit() {
    let family = Family::new(
        "Added & escaped",
        "{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}",
        "{4A3C46E8-61CC-4603-A589-7422A47A8E4A}",
    )
    .expect("valid family");
    for body in [
        String::new(),
        "<a:extLst/>".to_owned(),
        format!(
            r#"<a:extLst><a:ext uri="{}"/></a:extLst>"#,
            part::EXTENSION_URI
        ),
    ] {
        let source = theme(&body);
        let output = part::add_family(source.as_bytes(), &family).expect("unbounded insertion");
        let exact = part::add_family_with_limit(source.as_bytes(), &family, output.len())
            .expect("exact output limit");
        assert_eq!(output, exact);
        assert!(part::add_family_with_limit(source.as_bytes(), &family, output.len() - 1).is_err());
    }
}

#[test]
fn native_family_name_edit_preserves_exact_future_cdata_payload() {
    let native = include_str!("fixtures/theme-part-native.xml");
    let start = native.find("<thm15:themeFamily").expect("native family");
    let end = start + native[start..].find("/>").expect("self-closing family");
    let payload = r#"<x:future xmlns:x="urn:vendor"><![CDATA[<?x]]></x:future>"#;
    let source = format!(
        "{}>{payload}</thm15:themeFamily>{}",
        &native[..end],
        &native[end + 2..]
    );
    let snapshot = part::read(source.as_bytes()).expect("native family with CDATA");
    let mut replacement = snapshot.family().expect("family projection").clone();
    replacement
        .set_name("Updated CDATA Theme")
        .expect("valid name");
    let changed = snapshot
        .replace_family(&replacement)
        .expect("CDATA is not a PI");
    let expected = format!(
        "{}{}",
        &source[..start],
        source[start..].replacen("name=\"Office Theme\"", "name=\"Updated CDATA Theme\"", 1)
    );
    assert_eq!(changed, expected.as_bytes());
    assert_eq!(
        part::read(&changed).expect("readback").family(),
        Some(&replacement)
    );
}

#[test]
fn caller_limit_preflight_handles_source_fragments_and_escaped_replacement_names() {
    use litchi_drawingml::theme::family;
    let native = include_str!("fixtures/theme-part-native.xml");
    let snapshot = part::read(native.as_bytes()).expect("native family");
    let mut replacement = snapshot.family().expect("family").clone();
    replacement
        .set_name("Expanded & <quoted> \"name\"\r\n\t")
        .expect("escaped name");
    let changed = part::replace_family(native.as_bytes(), &replacement).expect("replace");
    assert!(changed.len() > native.len());
    assert_eq!(
        part::replace_family_with_limit(native.as_bytes(), &replacement, changed.len())
            .expect("exact limit"),
        changed
    );
    assert!(
        part::replace_family_with_limit(native.as_bytes(), &replacement, changed.len() - 1)
            .is_err()
    );

    let fragment = format!(
        r#"<f:themeFamily xmlns:f="{}" name="Large" id="{{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}}" vid="{{4A3C46E8-61CC-4603-A589-7422A47A8E4A}}"><!--{}--></f:themeFamily>"#,
        family::NAMESPACE,
        "x".repeat(64 * 1024)
    );
    let incoming = family::read(fragment.as_bytes()).expect("source-backed incoming family");
    let source = theme("");
    assert!(part::add_family_with_limit(source.as_bytes(), &incoming, 1024).is_err());
    let added = part::add_family(source.as_bytes(), &incoming).expect("large add");
    assert_eq!(
        part::add_family_with_limit(source.as_bytes(), &incoming, added.len())
            .expect("exact add limit"),
        added
    );
}

#[test]
fn empty_prefixed_bindings_are_rejected_before_attribute_projection() {
    let source = r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:x="" x:y="z"/>"#;
    assert_rejected(source.as_bytes());
    for body in [
        r#"<opaque xmlns:x="" x:y="z"/>"#,
        r#"<opaque xmlns:x=""><child x:y="z"/></opaque>"#,
        r#"<opaque xmlns:x="urn:vendor"><child xmlns:x="" x:y="z"/></opaque>"#,
    ] {
        assert_rejected(theme(body).as_bytes());
    }
    let valid =
        theme(r#"<opaque xmlns="urn:vendor"><child xmlns="" y="z" xml:lang="en"/></opaque>"#);
    assert!(
        part::read(valid.as_bytes()).is_ok(),
        "default namespace reset is legal"
    );
}

#[test]
fn standalone_document_markup_cannot_be_embedded_as_a_family_child() {
    use litchi_drawingml::theme::family;
    let value = Family::new(
        "Standalone",
        "{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}",
        "{4A3C46E8-61CC-4603-A589-7422A47A8E4A}",
    )
    .expect("family");
    let fragment = String::from_utf8(family::write(&value).expect("write family")).expect("UTF-8");
    let source = theme("");
    for prefix in [
        "<?xml version='1.0'?>",
        "\u{feff}<?xml version='1.0'?>",
        "<?xml version='1.0'?><!-- prolog -->",
    ] {
        let standalone = format!("{prefix}{fragment}");
        let parsed = family::read(standalone.as_bytes()).expect("valid standalone declaration");
        assert!(part::add_family(source.as_bytes(), &parsed).is_err());
        assert!(part::replace_family(source.as_bytes(), &parsed).is_err());
    }
    for standalone in [
        format!("<?future data?>{fragment}"),
        fragment.replace("/>", "><?future data?></thm15:themeFamily>"),
    ] {
        assert!(
            family::read(standalone.as_bytes()).is_err(),
            "actual PI must be rejected"
        );
    }
    let commented = family::read(format!("<!-- <?x is comment text -->{fragment}").as_bytes())
        .expect("legal comment");
    let added = part::add_family(source.as_bytes(), &commented).expect("comment can be embedded");
    assert!(
        std::str::from_utf8(&added)
            .expect("UTF-8")
            .contains("<!-- <?x is comment text -->")
    );
}

#[test]
fn namespace_heavy_foreign_families_stay_opaque_and_duplicate_owners_fail_early() {
    use litchi_drawingml::theme::family;
    let uri = format!("urn:{}", "x".repeat(family::MAX_NAMESPACE_BYTES - 4));
    let declarations = (0..120)
        .map(|index| format!(" xmlns:p{index}=\"{uri}\""))
        .collect::<String>();
    let root = format!(
        r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:f="{}"{declarations}><a:extLst>"#,
        family::NAMESPACE
    );
    // Each foreign family is small; copying the ~0.5 MiB active scope for each
    // one would amplify a sub-1-MiB document into hundreds of MiB of retention.
    let foreign = r#"<a:ext uri="urn:foreign"><f:themeFamily/></a:ext>"#.repeat(1000);
    let source = format!("{root}{foreign}</a:extLst></a:theme>");
    assert!(source.len() < family::MAX_XML_BYTES);
    assert!(
        part::read_family(source.as_bytes())
            .expect("opaque foreign families")
            .is_none()
    );

    // The second admitted family must be rejected as it is reached, before
    // retaining another scope or continuing to the deliberately malformed tail.
    for second in ["<f:themeFamily/>", "<f:themeFamily></f:themeFamily>"] {
        let source = format!(
            r#"{root}<a:ext uri="{}"><f:themeFamily name="A" id="{{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}}" vid="{{4A3C46E8-61CC-4603-A589-7422A47A8E4A}}"/></a:ext><a:ext uri="{}">{second}</a:ext><broken"#,
            part::EXTENSION_URI,
            part::NATIVE_EXTENSION_URI
        );
        let error = part::read_family(source.as_bytes()).expect_err("duplicate supported owner");
        assert!(
            error
                .to_string()
                .contains("multiple direct Theme Family owners"),
            "{error}"
        );
    }
}

#[test]
fn active_namespace_cap_counts_nested_declarations_and_restores_on_scope_exit() {
    let root_count = part::MAX_ACTIVE_NAMESPACE_DECLARATIONS / 2;
    // The Theme's own a binding is one of the root's active declarations.
    let root_declarations = (1..root_count)
        .map(|index| format!(" xmlns:p{index}=\"urn:root:{index}\""))
        .collect::<String>();
    let root = format!(
        r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"{root_declarations}>"#
    );
    let child_count = part::MAX_ACTIVE_NAMESPACE_DECLARATIONS - root_count;
    for extra in [0, 1] {
        // Rebinding root prefixes still consumes scope entries; those entries
        // must be released for both paired and self-closing siblings.
        let declarations = (0..child_count + extra)
            .map(|index| format!(" xmlns:p{index}=\"urn:child:{index}\""))
            .collect::<String>();
        let source = format!(
            "{root}<opaque{declarations}/><opaque{declarations}><child/></opaque><opaque{declarations}/></a:theme>"
        );
        if extra == 0 {
            assert!(
                part::read(source.as_bytes()).is_ok(),
                "exact active declaration cap"
            );
        } else {
            let error = part::read(source.as_bytes()).expect_err("one over active declaration cap");
            assert!(
                error.to_string().contains("active namespace bindings"),
                "{error}"
            );
        }
    }
}

#[test]
fn direct_extension_list_rejects_foreign_children_but_extensions_keep_opaque_payload() {
    for child in [
        r#"<x:foreign xmlns:x="urn:vendor"/>"#,
        r#"<x:foreign xmlns:x="urn:vendor"><x:child/></x:foreign>"#,
        r#"<x:ext xmlns:x="urn:vendor" uri="urn:vendor"/>"#,
        "<a:themeElements/>",
    ] {
        let source = theme(&format!("<a:extLst>{child}</a:extLst>"));
        assert_rejected(source.as_bytes());
    }
    for uri in [part::NATIVE_EXTENSION_URI, "urn:vendor"] {
        let source = theme(&format!(
            r#"<a:extLst><!--keep--><a:ext uri="{uri}"><x:foreign xmlns:x="urn:vendor"><x:child/></x:foreign></a:ext></a:extLst>"#
        ));
        let snapshot = part::read(source.as_bytes()).expect("foreign payload inside extension");
        assert_eq!(snapshot.xml_bytes(), source.as_bytes());
    }
}
