//! Regression tests for the note configuration model and XML codec.

use super::{
    Class, Configuration, MAX_VALUE_BYTES, NumberingScope, Position, parse, remove_xml, set_xml,
};
use crate::line_numbering::Format;
use crate::{FlatDocument, GenericPackage as Package};
use litchi_odf_common::XmlStreamLimits;

const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const STYLE: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";

fn styles(body: &str) -> String {
    format!(
        r#"<o:document-styles xmlns:o="{OFFICE}" xmlns:t="{TEXT}" xmlns:s="{STYLE}"><o:styles>{body}</o:styles><o:automatic-styles/><o:master-styles/></o:document-styles>"#
    )
}

fn limits() -> XmlStreamLimits {
    XmlStreamLimits::owner_default_value()
}

#[test]
fn parses_and_round_trips_footnote_and_endnote_configurations() {
    // ODF 1.2/1.3 text:notes-configuration grammar; values mirror
    // LibreOffice styles.xml footnote/endnote declarations.
    let xml = styles(
        r#"<t:notes-configuration t:note-class="footnote" t:citation-style-name="Footnote_20_Symbol" t:citation-body-style-name="Footnote_20_anchor" t:default-style-name="Footnote" t:master-page-name="Standard" t:start-value="2" s:num-prefix="[" s:num-suffix="]" s:num-format="a" s:num-letter-sync="true" t:start-numbering-at="chapter" t:footnotes-position="page"><t:note-continuation-notice-forward>Continued &amp; next</t:note-continuation-notice-forward><t:note-continuation-notice-backward><![CDATA[From <previous>]]></t:note-continuation-notice-backward></t:notes-configuration><t:notes-configuration t:note-class="endnote" s:num-format="I" t:start-numbering-at="document"/>"#,
    );
    let configurations = parse(&xml, &limits()).unwrap();
    let footnote = configurations.get(Class::Footnote).unwrap();
    assert_eq!(footnote.start_value, Some(2));
    assert_eq!(footnote.number_format, Some(Format::LowerAlpha));
    assert_eq!(footnote.letter_sync, Some(true));
    assert_eq!(
        footnote.continuation_notice_forward.as_deref(),
        Some("Continued & next")
    );
    assert_eq!(
        footnote.continuation_notice_backward.as_deref(),
        Some("From <previous>")
    );
    assert!(configurations.get(Class::Endnote).is_some());

    let serialized = configurations.to_xml_fragment(&limits()).unwrap();
    let reparsed = parse(&styles(&serialized), &limits()).unwrap();
    assert_eq!(reparsed, configurations);
    let shallow =
        XmlStreamLimits::new(xml.len() as u64 + 1, 3, 100_000, 100_000, 64 * 1024).unwrap();
    assert!(parse(&xml, &shallow).is_err());
}

#[test]
fn rejects_malformed_or_duplicate_note_configuration() {
    for body in [
        r"<t:notes-configuration/>",
        r#"<t:notes-configuration t:note-class="margin"/>"#,
        r#"<t:notes-configuration t:note-class="footnote" s:num-format="1" s:num-letter-sync="true"/>"#,
        r#"<t:notes-configuration t:note-class="footnote" t:start-value="-1"/>"#,
        r#"<t:notes-configuration t:note-class="footnote" t:start-numbering-at="section"/>"#,
        r#"<t:notes-configuration t:note-class="footnote"><t:note-continuation-notice-forward/><t:note-continuation-notice-forward/></t:notes-configuration>"#,
        r#"<t:notes-configuration t:note-class="footnote"><t:note-continuation-notice-forward><t:span/></t:note-continuation-notice-forward></t:notes-configuration>"#,
        r#"<t:notes-configuration t:note-class="footnote"/><t:notes-configuration t:note-class="footnote"/>"#,
    ] {
        assert!(parse(&styles(body), &limits()).is_err(), "{body}");
    }
}

#[test]
fn rejects_misplaced_note_configuration() {
    let xml = format!(
        r#"<o:document-content xmlns:o="{OFFICE}" xmlns:t="{TEXT}"><o:body><t:notes-configuration t:note-class="footnote"/></o:body></o:document-content>"#
    );
    assert!(parse(&xml, &limits()).is_err());
}

#[test]
fn exhaustive_enums_and_number_formats_round_trip() {
    for note_class in Class::ALL {
        for scope in NumberingScope::ALL {
            for position in Position::ALL {
                let mut value = Configuration::new(note_class);
                value.citation_style_name = Some(String::new());
                value.default_style_name = Some("Style_1".to_string());
                value.number_format = Some(Format::LowerAlpha);
                value.letter_sync = Some(false);
                value.start_numbering_at = Some(scope);
                value.footnotes_position = Some(position);
                let parsed = parse(&styles(&value.to_xml(&limits()).unwrap()), &limits()).unwrap();
                assert_eq!(parsed.get(note_class), Some(&value));
            }
        }
    }
    for format in [
        Format::Empty,
        Format::Arabic,
        Format::LowerRoman,
        Format::UpperRoman,
        Format::LowerAlpha,
        Format::UpperAlpha,
        Format::Custom("①".to_string()),
    ] {
        let mut value = Configuration::new(Class::Footnote);
        value.letter_sync =
            matches!(format, Format::LowerAlpha | Format::UpperAlpha).then_some(true);
        value.number_format = Some(format);
        assert_eq!(
            parse(&styles(&value.to_xml(&limits()).unwrap()), &limits())
                .unwrap()
                .footnote,
            Some(value)
        );
    }
}

#[test]
fn accepts_interleaved_notice_order_and_real_libreoffice() {
    let reverse = styles(
        r#"<t:notes-configuration t:note-class="footnote"><t:note-continuation-notice-backward>back</t:note-continuation-notice-backward><t:note-continuation-notice-forward>forward</t:note-continuation-notice-forward></t:notes-configuration>"#,
    );
    let parsed = parse(&reverse, &limits()).unwrap();
    assert_eq!(
        parsed
            .footnote
            .as_ref()
            .unwrap()
            .continuation_notice_forward
            .as_deref(),
        Some("forward")
    );
    assert_eq!(
        parsed
            .footnote
            .as_ref()
            .unwrap()
            .continuation_notice_backward
            .as_deref(),
        Some("back")
    );

    let fixture =
        include_str!("../../../../test-data/libreoffice-core/sw/qa/uitest/data/tdf145178.fodt");
    let real = parse(fixture, &limits()).unwrap();
    assert_eq!(
        real.footnote.as_ref().unwrap().number_format,
        Some(Format::Arabic)
    );
    assert_eq!(
        real.endnote.as_ref().unwrap().number_format,
        Some(Format::LowerRoman)
    );
    let flat = FlatDocument::from_bytes(fixture.as_bytes().to_vec()).unwrap();
    assert_eq!(flat.notes_configurations().unwrap(), real);
}

#[test]
fn rejects_style_refs_namespaces_booleans_and_caps() {
    for body in [
        r#"<t:notes-configuration t:note-class="footnote" t:citation-style-name="bad name"/>"#,
        r#"<t:notes-configuration t:note-class="footnote" t:default-style-name="1bad"/>"#,
        r#"<t:notes-configuration t:note-class="footnote" s:num-format="a" s:num-letter-sync="1"/>"#,
        r#"<t:notes-configuration t:note-class="footnote" s:num-format="custom" s:num-letter-sync="true"/>"#,
        r#"<t:notes-configuration t:note-class="footnote"><t:note-continuation-notice-forward t:bad="1"/></t:notes-configuration>"#,
        r#"<t:notes-configuration xmlns:x="urn:wrong" t:note-class="footnote" x:note-class="endnote"/>"#,
    ] {
        assert!(parse(&styles(body), &limits()).is_err(), "accepted {body}");
    }
    let wrong_namespace = format!(
        r#"<o:document-styles xmlns:o="{OFFICE}" xmlns:t="{TEXT}" xmlns:x="urn:wrong"><o:styles><x:notes-configuration t:note-class="footnote"/></o:styles></o:document-styles>"#
    );
    assert!(parse(&wrong_namespace, &limits()).is_err());
    let mut capped = Configuration::new(Class::Footnote);
    capped.continuation_notice_forward = Some("x".repeat(MAX_VALUE_BYTES + 1));
    assert!(capped.validate().is_err());
}

#[test]
fn lossless_mutation_and_builder_package_access() {
    use crate::Builder;

    let original = styles(
        r#"<!--keep--><style:list-style xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" style:name="L"/>"#,
    );
    let mut value = Configuration::new(Class::Footnote);
    value.start_value = Some(2);
    let inserted = set_xml(&original, &value, &limits()).unwrap();
    assert!(inserted.contains("<!--keep--><style:list-style"));
    value.start_value = Some(3);
    let replaced = set_xml(&inserted, &value, &limits()).unwrap();
    assert!(replaced.contains("text:start-value=\"3\""));
    assert!(!replaced.contains("text:start-value=\"2\""));
    assert_eq!(
        remove_xml(&replaced, Class::Footnote, &limits()).unwrap(),
        original
    );

    let mut builder = Builder::new();
    builder.set_notes_configuration(value.clone()).unwrap();
    let package = Package::from_bytes(&builder.build().unwrap().to_bytes().unwrap()).unwrap();
    assert_eq!(
        package.notes_configurations().unwrap().footnote,
        Some(value)
    );
}

#[test]
fn note_writer_honors_exact_and_tiny_output_budgets() {
    let configuration = Configuration::new(Class::Footnote);
    let expected = configuration
        .to_xml(&limits())
        .expect("default note output");
    let exact =
        XmlStreamLimits::new(expected.len() as u64, 256, 100_000, 100_000, 64 * 1024).unwrap();
    assert_eq!(configuration.to_xml(&exact).unwrap(), expected);
    let tiny = XmlStreamLimits::new(1, 256, 100_000, 100_000, 64 * 1024).unwrap();
    assert!(configuration.to_xml(&tiny).is_err());
}
