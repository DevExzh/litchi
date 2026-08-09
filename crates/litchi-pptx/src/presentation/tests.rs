#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use super::codec;
use crate::hyperlinks::Hyperlink;

#[test]
fn scans_inline_hyperlinks_with_tooltips() {
    let xml = br#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:hlinkClick action="https://example.com" tooltip="Open site"/></p:sld>"#;

    assert!(matches!(
        codec::parse_inline_hyperlinks(xml),
        Ok(values) if values == vec![Hyperlink::url_with_tooltip("https://example.com", "Open site")]
    ));
}

#[test]
fn rejects_empty_inline_hyperlink_actions() {
    let xml = br#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:hlinkClick action=""/></p:sld>"#;

    assert!(codec::parse_inline_hyperlinks(xml).is_err());
}
