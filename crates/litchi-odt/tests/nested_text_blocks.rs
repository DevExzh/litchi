//! Paragraphs and headings nested inside frames, annotations, and sections.
//!
//! ODF lets a `text:p` carry further paragraphs through `draw:frame` /
//! `draw:text-box`, `draw:custom-shape`, inline `office:annotation`, and the
//! tables and sections those frames contain. Every one of those inner
//! paragraphs is real user content, so the readers surface each as its own
//! block in document order rather than rejecting the document.

use litchi_odt::Document;
use litchi_odt::elements::text::TextElements;
use litchi_odt::generic::Family;

/// Namespace preamble shared by the hand-written content fragments.
const CONTENT_PREFIX: &str = concat!(
    r#"<office:document-content"#,
    r#" xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0""#,
    r#" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0""#,
    r#" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0""#,
    r#" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">"#,
    r#"<office:body><office:text>"#,
);

/// Closing tags matching [`CONTENT_PREFIX`].
const CONTENT_SUFFIX: &str = "</office:text></office:body></office:document-content>";

fn content(body: &str) -> String {
    format!("{CONTENT_PREFIX}{body}{CONTENT_SUFFIX}")
}

#[test]
fn text_box_paragraphs_inside_a_paragraph_become_their_own_blocks() {
    let xml = content(concat!(
        "<text:p>before ",
        "<draw:frame><draw:text-box>",
        "<text:p>framed</text:p>",
        "</draw:text-box></draw:frame>",
        " after</text:p>",
    ));

    // Document order is start order: the enclosing paragraph precedes the frame
    // content it carries, and the outer run keeps the text on both sides.
    assert_eq!(
        TextElements::extract_text(&xml).unwrap(),
        "before  after\nframed"
    );
    assert_eq!(TextElements::parse_paragraphs(&xml).unwrap().len(), 2);
}

#[test]
fn nested_headings_tables_and_sections_are_all_reached() {
    let xml = content(concat!(
        "<text:p>outer",
        "<draw:custom-shape><text:h text:outline-level=\"2\">shape heading</text:h>",
        "<table:table><table:table-row><table:table-cell>",
        "<text:p>celled</text:p>",
        "</table:table-cell></table:table-row></table:table>",
        "<text:section><text:p>sectioned</text:p></text:section>",
        "</draw:custom-shape></text:p>",
    ));

    assert_eq!(
        TextElements::extract_text(&xml).unwrap(),
        "outer\nshape heading\ncelled\nsectioned"
    );
    let headings = TextElements::parse_headings(&xml).unwrap();
    assert_eq!(headings.len(), 1);
    assert_eq!(headings[0].text().unwrap(), "shape heading");
    assert_eq!(headings[0].level(), Some(2));
}

#[test]
fn inline_annotation_paragraphs_are_reached_and_do_not_swallow_the_host() {
    let xml = content(concat!(
        "<text:p>host",
        r#"<office:annotation office:name="a1">"#,
        "<text:p>commented</text:p>",
        "</office:annotation>",
        " tail</text:p>",
    ));

    assert_eq!(
        TextElements::extract_text(&xml).unwrap(),
        "host tail\ncommented"
    );
}

#[test]
fn note_bodies_and_ruby_text_stay_out_of_the_visible_flow() {
    // Note bodies and ruby pronunciation runs have dedicated readers, so they
    // must not leak into the paragraph flow even though they nest paragraphs.
    let xml = content(concat!(
        "<text:p>cited",
        r#"<text:note text:note-class="footnote">"#,
        "<text:note-citation>1</text:note-citation>",
        "<text:note-body><text:p>note text</text:p></text:note-body>",
        "</text:note>",
        "<text:ruby><text:ruby-base>base</text:ruby-base>",
        "<text:ruby-text>ruby</text:ruby-text></text:ruby>",
        "</text:p>",
    ));

    assert_eq!(TextElements::extract_text(&xml).unwrap(), "cited1base");
    assert_eq!(TextElements::parse_paragraphs(&xml).unwrap().len(), 1);
}

#[test]
fn tracked_change_regions_remain_excluded_when_they_nest() {
    let xml = content(concat!(
        "<text:tracked-changes><text:changed-region>",
        "<text:deletion><text:p>deleted</text:p></text:deletion>",
        "</text:changed-region></text:tracked-changes>",
        "<text:p>kept</text:p>",
    ));

    assert_eq!(TextElements::extract_text(&xml).unwrap(), "kept");
}

#[test]
fn deeply_nested_frames_keep_every_level() {
    let xml = content(concat!(
        "<text:p>l0",
        "<draw:frame><draw:text-box><text:p>l1",
        "<draw:frame><draw:text-box><text:p>l2</text:p></draw:text-box></draw:frame>",
        "</text:p></draw:text-box></draw:frame></text:p>",
    ));

    assert_eq!(TextElements::extract_text(&xml).unwrap(), "l0\nl1\nl2");
}

#[test]
fn sibling_paragraphs_in_one_shape_all_close_against_the_right_element() {
    // A shape holding several paragraphs is the shape that exposes depth
    // book-keeping mistakes: each nested block must consume exactly one end tag,
    // so the anchoring paragraph still closes on its own.
    let xml = content(concat!(
        "<text:p>anchor",
        "<draw:custom-shape>",
        "<text:p>one</text:p><text:p>two</text:p><text:p>three</text:p>",
        "<draw:enhanced-geometry/>",
        "</draw:custom-shape>",
        "<text:span>tail</text:span></text:p>",
        "<text:p>next</text:p>",
    ));

    assert_eq!(
        TextElements::extract_text(&xml).unwrap(),
        "anchortail\none\ntwo\nthree\nnext"
    );
    assert_eq!(TextElements::parse_paragraphs(&xml).unwrap().len(), 5);
}

#[test]
fn unbalanced_nested_paragraph_markup_is_a_typed_error() {
    let xml = format!("{CONTENT_PREFIX}<text:p>open<draw:text-box><text:p>inner</draw:text-box>");
    assert!(TextElements::extract_text(&xml).is_err());
}

#[test]
fn producer_written_frame_captions_round_trip_through_the_flat_reader() {
    // LibreOffice writes an anchored caption frame as a `text:p` that contains a
    // `draw:text-box` with its own `text:p`. Before frames were reachable this
    // shape made `text()` fail outright for the whole document.
    let flat = litchi_odt::generic::FlatDocument::open(
        "../../test-data/odf/odt/nested-frame-paragraphs.fodt",
    )
    .unwrap();
    assert_eq!(flat.family(), Family::Text);
    assert!(flat.xml().contains("bar"));
    assert!(flat.xml().contains("Illustration"));
    assert!(flat.xml().contains(": foo"));
}

#[test]
fn packaged_documents_reach_shape_paragraphs_too() {
    // A shape anchored to a paragraph stores its label as a `text:p` inside the
    // `draw:custom-shape`, which sits inside the anchoring `text:p`.
    let document = Document::open("../../test-data/odf/odt/shape-text-in-paragraph.odt").unwrap();

    assert_eq!(document.paragraph_count().unwrap(), 2);
    assert_eq!(document.text().unwrap(), "\nText Inside The Rectangle");
}
