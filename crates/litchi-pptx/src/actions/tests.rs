//! Focused semantic and bounded-scanner coverage for action settings.

use super::codec;
use super::model::{Jump, Kind, Trigger};
use super::package::Limits;

#[test]
fn classifies_reserved_powerpoint_action_values() {
    assert_eq!(
        super::model::classify(Some("ppaction://hlinkshowjump?jump=nextslide"), false),
        Kind::SlideShowJump(Jump::NextSlide)
    );
    assert_eq!(
        super::model::classify(Some("ppaction://customshow?id=42"), false),
        Kind::CustomShow { id: 42 }
    );
    assert_eq!(
        super::model::classify(Some("ppaction://hlinkpres?slideindex=7"), true),
        Kind::Presentation {
            start_slide_index: 7
        }
    );
    assert_eq!(
        super::model::classify(Some("ppaction://macro?name=Module1.Run"), false),
        Kind::Macro
    );
    assert_eq!(
        super::model::classify(Some("urn:vendor:custom"), false),
        Kind::Unknown
    );
    assert_eq!(super::model::classify(None, true), Kind::Hyperlink);
    assert_eq!(super::model::classify(None, false), Kind::None);
}

#[test]
fn scans_strict_click_and_hover_namespaces() {
    let xml = br#"<p:sld xmlns:p="http://purl.oclc.org/ooxml/presentationml/main" xmlns:a="http://purl.oclc.org/ooxml/drawingml/main" xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships"><a:hlinkClick r:id="rId1"/><a:hlinkHover action="ppaction://program"/></p:sld>"#;
    let actions = codec::scan(xml, &mut Limits::default()).expect("strict action scan");
    assert_eq!(actions.len(), 2);
    assert_eq!(actions[0].trigger, Trigger::Click);
    assert_eq!(actions[0].relationship_id.as_deref(), Some("rId1"));
    assert_eq!(actions[1].trigger, Trigger::Hover);
}

#[test]
fn rejects_duplicate_relationship_attributes() {
    let xml = br#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><a:hlinkClick r:id="one" r:id="two"/></p:sld>"#;
    assert!(codec::scan(xml, &mut Limits::default()).is_err());
}

#[test]
fn rejects_non_slide_roots() {
    let xml = br#"<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>"#;
    assert!(codec::scan(xml, &mut Limits::default()).is_err());
}
