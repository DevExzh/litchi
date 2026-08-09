#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use super::Hyperlink;

#[test]
fn constructors_and_codec_keep_target_contexts_distinct() {
    let external = Hyperlink::url_with_tooltip("https://example.com", "Open site");
    assert!(external.is_external());
    assert_eq!(external.tooltip(), Some("Open site"));
    assert_eq!(external.target(), "https://example.com");

    let slide = Hyperlink::from_xml("ppaction://hlinksldjump?sldNum=7", None).unwrap();
    assert_eq!(slide, Hyperlink::slide(7));
    assert!(!slide.is_external());

    let email = Hyperlink::from_xml(
        "mailto:test@example.com?cc=other@example.com&subject=Hello",
        None,
    )
    .unwrap();
    assert_eq!(
        email,
        Hyperlink::Email {
            email: "test@example.com".to_owned(),
            subject: Some("Hello".to_owned()),
            tooltip: None,
        }
    );
}

#[test]
fn parser_rejects_unbounded_target_and_tooltip_inputs() {
    let target = "x".repeat((1 << 20) + 1);
    assert!(Hyperlink::from_xml(&target, None).is_err());
    assert!(Hyperlink::from_xml("https://example.com", Some("x".repeat((1 << 20) + 1))).is_err());
}
