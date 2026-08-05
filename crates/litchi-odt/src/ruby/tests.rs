//! Regression tests for the ruby XML codec.

use super::{MAX_DEPTH, parse_rubies};

const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";

#[test]
fn parses_nested_ruby_pairs_with_styles_and_exact_text() {
    let xml = format!(
        r#"<x:p xmlns:x="{TEXT}"><x:ruby x:style-name="Outer"><x:ruby-base>A&amp;<x:s x:c="2"/><x:ruby><x:ruby-base>B</x:ruby-base><x:ruby-text x:style-name="InnerText">b</x:ruby-text></x:ruby>C</x:ruby-base><x:ruby-text x:style-name="Pronunciation"><![CDATA[abc]]></x:ruby-text></x:ruby></x:p>"#
    );
    let rubies = parse_rubies(&xml).unwrap();
    assert_eq!(rubies.len(), 2);
    assert_eq!(rubies[0].style_name(), Some("Outer"));
    assert_eq!(rubies[0].base(), "A&  BC");
    assert_eq!(rubies[0].text(), "abc");
    assert_eq!(rubies[0].text_style_name(), Some("Pronunciation"));
    assert_eq!(rubies[1].style_name(), None);
    assert_eq!(rubies[1].base(), "B");
    assert_eq!(rubies[1].text(), "b");
    assert_eq!(rubies[1].text_style_name(), Some("InnerText"));
}

#[test]
fn rubies_reject_invalid_structure_and_ambiguous_attributes() {
    let missing = format!(r#"<x:ruby xmlns:x="{TEXT}"><x:ruby-base>A</x:ruby-base></x:ruby>"#);
    assert!(parse_rubies(&missing).is_err());
    let wrong_order = format!(
        r#"<x:ruby xmlns:x="{TEXT}"><x:ruby-text>a</x:ruby-text><x:ruby-base>A</x:ruby-base></x:ruby>"#
    );
    assert!(parse_rubies(&wrong_order).is_err());
    let duplicate = format!(
        r#"<x:ruby xmlns:x="{TEXT}"><x:ruby-base>A</x:ruby-base><x:ruby-base>B</x:ruby-base><x:ruby-text>a</x:ruby-text></x:ruby>"#
    );
    assert!(parse_rubies(&duplicate).is_err());
    let text_child = format!(
        r#"<x:ruby xmlns:x="{TEXT}"><x:ruby-base>A</x:ruby-base><x:ruby-text><x:span>a</x:span></x:ruby-text></x:ruby>"#
    );
    assert!(parse_rubies(&text_child).is_err());
    let aliases = format!(
        r#"<x:ruby xmlns:x="{TEXT}" xmlns:y="{TEXT}" x:style-name="A" y:style-name="B"><x:ruby-base>A</x:ruby-base><x:ruby-text>a</x:ruby-text></x:ruby>"#
    );
    assert!(parse_rubies(&aliases).is_err());
    let empty = format!(r#"<x:ruby xmlns:x="{TEXT}"/>"#);
    assert!(parse_rubies(&empty).is_err());
    assert!(parse_rubies("<x:ruby>").is_err());
}

#[test]
fn rubies_enforce_nesting_bound() {
    let mut xml = format!(r#"<x:ruby xmlns:x="{TEXT}"><x:ruby-base>"#);
    for _ in 0..MAX_DEPTH {
        xml.push_str("<x:span>");
    }
    for _ in 0..MAX_DEPTH {
        xml.push_str("</x:span>");
    }
    xml.push_str("</x:ruby-base><x:ruby-text>a</x:ruby-text></x:ruby>");
    assert!(parse_rubies(&xml).is_err());
}
