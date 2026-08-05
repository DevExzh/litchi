//! Focused tests for the layered web codec facade.

use super::*;
use crate::Result;

#[test]
fn namespace_codec_keeps_extension_fragments_self_contained() -> Result<()> {
    let document = parse_xml(
        br#"<we:webextension xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11"><we:extLst><a:ext xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:vendor="urn:vendor" uri="urn:test"><vendor:item>value</vendor:item></a:ext></we:extLst></we:webextension>"#,
    )?;
    let root = document.root()?;
    let fragment = document.self_contained_fragment(&root.children[0])?;
    assert!(fragment.contains("xmlns:vendor="));
    assert!(fragment.contains("<vendor:item>"));
    Ok(())
}

#[test]
fn relationship_codec_rejects_mixed_conformance_attributes() -> Result<()> {
    let document = parse_xml(
        br#"<we:snapshot xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:s="http://purl.oclc.org/ooxml/officeDocument/relationships" r:embed="rId1" s:embed="sId1"/>"#,
    )?;
    assert!(relationship_attr(document.root()?, "embed").is_err());
    Ok(())
}

#[test]
fn semantic_codec_uses_canonical_floating_point_lexical_form() {
    assert_eq!(format_f64(1.5), "1.5");
    assert_eq!(format_f64(-0.0), "-0.0");
}
