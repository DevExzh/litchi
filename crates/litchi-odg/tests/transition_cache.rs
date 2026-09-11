#![allow(clippy::unwrap_used, reason = "test fixtures use unwrap for clarity")]

use litchi_core::Error;
use litchi_odf_common::core::PackageWriter;
use litchi_odg::Drawing;
use std::fmt::Write as _;

const PRESENTATION: &str = "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0";
const SMIL: &str = "urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0";

fn content_with_pages(style_name: &str, page_count: usize, automatic_styles: &str) -> String {
    let mut content = String::from(
        r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:p="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:xml="http://www.w3.org/XML/1998/namespace" office:version="1.4">"#,
    );
    if !automatic_styles.is_empty() {
        content.push_str("<office:automatic-styles>");
        content.push_str(automatic_styles);
        content.push_str("</office:automatic-styles>");
    }
    content.push_str("<office:body><office:drawing>");
    for page in 0..page_count {
        write!(
            &mut content,
            r#"<draw:page draw:name="Page {page}" draw:style-name="{style_name}"/>"#
        )
        .unwrap();
    }
    content.push_str("</office:drawing></office:body></office:document-content>");
    content
}

fn styles_document(definitions: &str) -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:p="{PRESENTATION}" xmlns:s="{SMIL}" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:xml="http://www.w3.org/XML/1998/namespace" office:version="1.4"><office:styles>{definitions}</office:styles></office:document-styles>"#,
        PRESENTATION = PRESENTATION,
        SMIL = SMIL,
    )
}

fn package(content: &str, styles: Option<&str>) -> Vec<u8> {
    let mut writer = PackageWriter::new();
    writer
        .set_mimetype("application/vnd.oasis.opendocument.graphics")
        .unwrap();
    writer.add_file("content.xml", content.as_bytes()).unwrap();
    writer.add_file("media/transition.wav", b"inert").unwrap();
    if let Some(styles) = styles {
        writer.add_file("styles.xml", styles.as_bytes()).unwrap();
    }
    writer.finish_to_bytes().unwrap()
}

fn direct_style(
    name: &str,
    transition_type: &str,
    transition_style: &str,
    sound_id: &str,
) -> String {
    format!(
        r##"<style:style style:name="{name}" style:family="drawing-page"><style:drawing-page-properties p:transition-type="{transition_type}" p:transition-style="{transition_style}" p:transition-speed="fast" s:type="fade" s:subtype="crossfade" s:direction="forward" s:fadeColor="#010203" p:duration="PT2S"><p:sound xlink:type="simple" xlink:href="media/transition.wav" xlink:actuate="onRequest" p:play-full="true" xlink:show="replace" xml:id="{sound_id}"/></style:drawing-page-properties></style:style>"##
    )
}

fn assert_invalid(content: &str, styles: Option<&str>) {
    assert!(matches!(
        Drawing::from_bytes(package(content, styles)),
        Err(Error::InvalidFormat(_))
    ));
}

#[test]
fn shared_direct_transition_reads_and_exact_noop_stay_consistent() {
    let automatic = direct_style("shared", "automatic", "dissolve", "shared-sound");
    let content = content_with_pages("shared", 64, &automatic);
    let source = Drawing::from_bytes(package(&content, None)).unwrap();
    let expected = source.pages()[0].transition().cloned().unwrap();

    assert_eq!(expected.transition_type(), Some("automatic"));
    assert_eq!(expected.style(), Some("dissolve"));
    assert_eq!(expected.speed(), Some("fast"));
    assert_eq!(expected.sound().unwrap().href(), "media/transition.wav");
    assert_eq!(expected.sound().unwrap().xml_id(), Some("shared-sound"));
    assert!(
        source
            .pages()
            .iter()
            .all(|page| page.transition() == Some(&expected))
    );

    let mut edit = source.edit();
    edit.set_page_transition(0, Some(expected.clone())).unwrap();
    let commit = edit.commit().unwrap();
    assert!(!commit.changed());
    assert_eq!(commit.snapshot().as_bytes(), source.as_bytes());
}

#[test]
fn inherited_alias_transition_and_sound_reads_for_many_pages() {
    let definitions = r##"<style:style style:name="base" style:family="drawing-page"><style:drawing-page-properties p:transition-style="fade-from-left" p:transition-speed="slow" s:type="fade" s:subtype="crossfade" s:direction="forward" s:fadeColor="#010203" p:duration="PT2S"><p:sound xlink:type="simple" xlink:href="media/transition.wav" xlink:actuate="onRequest" p:play-full="true" xlink:show="replace" xml:id="base-sound"/></style:drawing-page-properties></style:style><style:style style:name="child" style:family="drawing-page" style:parent-style-name="base"><style:drawing-page-properties p:transition-type="automatic"/></style:style>"##;
    let styles = styles_document(definitions);
    let content = content_with_pages("child", 48, "");
    let source = Drawing::from_bytes(package(&content, Some(&styles))).unwrap();

    for page in source.pages() {
        let transition = page.transition().unwrap();
        assert_eq!(transition.transition_type(), Some("automatic"));
        assert_eq!(transition.style(), Some("fade-from-left"));
        assert_eq!(transition.speed(), Some("slow"));
        assert_eq!(transition.smil_type(), Some("fade"));
        assert_eq!(transition.duration(), Some("PT2S"));
        let sound = transition.sound().unwrap();
        assert_eq!(sound.href(), "media/transition.wav");
        assert_eq!(sound.play_full(), Some(true));
        assert!(sound.actuate_on_request());
        assert_eq!(sound.show(), Some("replace"));
        assert_eq!(sound.xml_id(), Some("base-sound"));
    }
}

#[test]
fn content_style_shadowing_and_independent_sources_keep_values_local() {
    let named = styles_document(&direct_style(
        "shared",
        "manual",
        "fade-from-left",
        "named-sound",
    ));
    let content_style = direct_style("shared", "automatic", "dissolve", "content-sound");
    let content = content_with_pages("shared", 12, &content_style);
    let shadowed = Drawing::from_bytes(package(&content, Some(&named))).unwrap();
    for page in shadowed.pages() {
        let transition = page.transition().unwrap();
        assert_eq!(transition.transition_type(), Some("automatic"));
        assert_eq!(transition.style(), Some("dissolve"));
        assert_eq!(transition.sound().unwrap().xml_id(), Some("content-sound"));
    }

    let first_styles = styles_document(&direct_style(
        "shared",
        "manual",
        "fade-from-left",
        "first-sound",
    ));
    let second_styles = styles_document(&direct_style(
        "shared",
        "automatic",
        "random",
        "second-sound",
    ));
    let first_content = content_with_pages("shared", 4, "");
    let second_content = content_with_pages("shared", 4, "");
    let first = Drawing::from_bytes(package(&first_content, Some(&first_styles))).unwrap();
    let second = Drawing::from_bytes(package(&second_content, Some(&second_styles))).unwrap();
    assert_eq!(
        first.pages()[0].transition().unwrap().style(),
        Some("fade-from-left")
    );
    assert_eq!(
        first.pages()[0]
            .transition()
            .unwrap()
            .sound()
            .unwrap()
            .xml_id(),
        Some("first-sound")
    );
    assert_eq!(
        second.pages()[0].transition().unwrap().style(),
        Some("random")
    );
    assert_eq!(
        second.pages()[0]
            .transition()
            .unwrap()
            .sound()
            .unwrap()
            .xml_id(),
        Some("second-sound")
    );
}

#[test]
fn malformed_referenced_transition_owners_are_not_accepted_from_cache() {
    let valid_named = styles_document(&direct_style(
        "shared",
        "manual",
        "fade-from-left",
        "valid-sound",
    ));
    let malformed_content = direct_style("shared", "automatic", "not-a-transition", "bad-sound");
    let content = content_with_pages("shared", 16, &malformed_content);
    assert_invalid(&content, Some(&valid_named));

    let malformed_named = styles_document(&direct_style(
        "broken",
        "automatic",
        "not-a-transition",
        "bad-sound",
    ));
    let content = content_with_pages("broken", 16, "");
    assert_invalid(&content, Some(&malformed_named));

    let malformed_sound = styles_document(
        r#"<style:style style:name="bad-sound" style:family="drawing-page"><style:drawing-page-properties p:transition-type="automatic"><p:sound xlink:type="extended" xlink:href="media/transition.wav"/></style:drawing-page-properties></style:style>"#,
    );
    let content = content_with_pages("bad-sound", 16, "");
    assert_invalid(&content, Some(&malformed_sound));
}

#[test]
fn cyclic_and_overdeep_transition_inheritance_remain_rejected() {
    let cyclic = r#"<style:style style:name="first" style:family="drawing-page" style:parent-style-name="second"/><style:style style:name="second" style:family="drawing-page" style:parent-style-name="first"/>"#;
    let content = content_with_pages("first", 1, "");
    let styles = styles_document(cyclic);
    assert_invalid(&content, Some(&styles));

    let mut chain = String::new();
    for index in 0..258 {
        let parent = if index + 1 < 258 {
            format!(
                r#" style:parent-style-name="style-{next}""#,
                next = index + 1
            )
        } else {
            String::new()
        };
        write!(
            &mut chain,
            r#"<style:style style:name="style-{index}" style:family="drawing-page"{parent}/>"#,
        )
        .unwrap();
    }
    let styles = styles_document(&chain);
    let shallow = content_with_pages("style-1", 1, "");
    Drawing::from_bytes(package(&shallow, Some(&styles))).unwrap();
    // The first page populates direct values along a valid depth-256 chain.
    // Reaching those same owners one level deeper must still fail on page two.
    let content = content_with_pages("style-1", 2, "").replacen(
        r#"draw:name="Page 1" draw:style-name="style-1""#,
        r#"draw:name="Page 1" draw:style-name="style-0""#,
        1,
    );
    assert_invalid(&content, Some(&styles));
}
