//! Focused regression coverage for ruby range structural selection.

use super::*;
use std::ops::Range;

const TEXT_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";

fn paragraph(body: &str) -> String {
    format!(r#"<text:p xmlns:text="{TEXT_NAMESPACE}">{body}</text:p>"#)
}

#[test]
fn locates_a_balanced_inline_range() {
    let xml = paragraph(r#"A<text:span>漢</text:span><text:span>字</text:span>Z"#);
    let spans = locate_balanced_ruby_ranges(&xml, 0, &(1.."漢字".len() + 1)).unwrap();
    assert_eq!(spans.len(), 1);
    assert_eq!(
        &xml[spans[0].start..spans[0].end],
        "<text:span>漢</text:span><text:span>字</text:span>"
    );
}

#[test]
fn rejects_a_range_crossing_an_existing_ruby() {
    let xml = paragraph(
        r#"A<text:ruby><text:ruby-base>X</text:ruby-base><text:ruby-text>x</text:ruby-text></text:ruby>B"#,
    );
    let range = Range { start: 0, end: 2 };
    assert!(
        locate_balanced_ruby_ranges(&xml, 0, &range)
            .unwrap()
            .is_empty()
    );
}
