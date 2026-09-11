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
