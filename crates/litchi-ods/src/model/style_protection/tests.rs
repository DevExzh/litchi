//! Snapshot and invariant tests for the table-cell protection owner.

use super::model::MAX_CONDITIONAL_ATTRIBUTE_BYTES;
use super::*;
use crate::model::names::formula;
use std::collections::{BTreeMap, HashSet};

#[test]
fn conditional_editor_preserves_unrelated_automatic_styles_with_arbitrary_prefixes() {
    let xml = r#"<o:automatic-styles xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:x="urn:example:formula"><draw:gradient xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" draw:name="keep&amp;exact"/><s:style s:name="old" s:family="table-cell"><s:map s:condition="x:old()" s:apply-style-name="Red"/></s:style><s:style s:name="plain" s:family="table-cell"><s:table-cell-properties/></s:style></o:automatic-styles>"#;
    let fragment = PreservedXmlFragment {
        xml: xml.to_string(),
        namespaces: BTreeMap::new(),
    };
    let style = ConditionalStyle::new(
        "new&style",
        vec![
            Rule::new("x:test()<2", "Red").with_formula_namespace(formula::Namespace {
                prefix: "x".to_string(),
                uri: "urn:example:formula".to_string(),
            }),
        ],
    );
    let common = HashSet::from(["Red".to_string()]);
    validate_conditional_style_collection(std::slice::from_ref(&style), &common)
        .expect("test fixture or operation should succeed");
    let rewritten = rewrite_conditional_styles(Some(&fragment), &[style])
        .expect("test fixture or operation should succeed");
    assert!(rewritten.xml.contains(
            r#"<draw:gradient xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" draw:name="keep&amp;exact"/>"#
        ));
    assert!(rewritten.xml.contains(
        r#"<s:style s:name="plain" s:family="table-cell"><s:table-cell-properties/></s:style>"#
    ));
    assert!(!rewritten.xml.contains("s:name=\"old\""));
    assert!(rewritten.xml.contains("style:name=\"new&amp;style\""));
}

#[test]
fn managed_protection_editor_preserves_unrelated_xml_and_merges_maps() {
    let xml = r#"<o:automatic-styles xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0"><draw:gradient xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" draw:name="keep"/><s:style s:name="combo" s:family="table-cell"><s:table-cell-properties s:cell-protect="protected"/><s:map s:condition="cell-content()>0" s:apply-style-name="Red"/></s:style><s:style s:name="plain" s:family="table-cell"><s:table-cell-properties/></s:style></o:automatic-styles>"#;
    let fragment = PreservedXmlFragment {
        xml: xml.to_string(),
        namespaces: BTreeMap::new(),
    };
    let conditional = ConditionalStyle::new("combo", vec![Rule::new("cell-content()>0", "Red")]);
    let protection = TableStyle::new("combo", Protection::HiddenAndProtected);
    let rewritten = rewrite_managed_cell_styles(Some(&fragment), &[conditional], &[protection])
        .expect("test fixture or operation should succeed");
    assert!(rewritten.xml.contains(
            r#"<draw:gradient xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" draw:name="keep"/>"#
        ));
    assert!(rewritten.xml.contains(
        r#"<s:style s:name="plain" s:family="table-cell"><s:table-cell-properties/></s:style>"#
    ));
    assert_eq!(rewritten.xml.matches("name=\"combo\"").count(), 1);
    assert!(
        rewritten
            .xml
            .contains("cell-protect=\"hidden-and-protected\"")
    );
    assert!(rewritten.xml.contains("condition=\"cell-content()&gt;0\""));
}

#[test]
fn resolves_defaults_named_parents_and_automatic_overrides() {
    let named = r#"<o:styles xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0"><s:default-style s:family="table-cell"><s:table-cell-properties s:cell-protect="none"/></s:default-style><s:style s:name="Locked" s:family="table-cell"><s:table-cell-properties s:cell-protect="protected"/></s:style><s:style s:name="Child" s:family="table-cell" s:parent-style-name="Locked"/></o:styles>"#;
    let content = r#"<o:automatic-styles xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0"><s:style s:name="Auto" s:family="table-cell" s:parent-style-name="Child"><s:table-cell-properties s:cell-protect="formula-hidden protected"/></s:style></o:automatic-styles>"#;
    let registry = CellStyleRegistry::parse(Some(named), content)
        .expect("test fixture or operation should succeed");
    assert_eq!(
        registry
            .resolve(None)
            .expect("test fixture or operation should succeed"),
        Some(Protection::None)
    );
    assert_eq!(
        registry
            .resolve(Some("Child"))
            .expect("test fixture or operation should succeed"),
        Some(Protection::Protected)
    );
    assert_eq!(
        registry
            .resolve(Some("Auto"))
            .expect("test fixture or operation should succeed"),
        Some(Protection::ProtectedFormulaHidden)
    );
}

#[test]
fn rejects_invalid_values_missing_parents_and_cycles() {
    let invalid = r#"<s:style xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0" s:name="Bad" s:family="table-cell"><s:table-cell-properties s:cell-protect="locked"/></s:style>"#;
    assert!(CellStyleRegistry::parse(None, invalid).is_err());

    let styles = r#"<o:styles xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0"><s:style s:name="Missing" s:family="table-cell" s:parent-style-name="Nope"/><s:style s:name="A" s:family="table-cell" s:parent-style-name="B"/><s:style s:name="B" s:family="table-cell" s:parent-style-name="A"/></o:styles>"#;
    let registry =
        CellStyleRegistry::parse(None, styles).expect("test fixture or operation should succeed");
    assert!(registry.resolve(Some("Missing")).is_err());
    assert!(registry.resolve(Some("A")).is_err());
}

#[test]
fn extracts_automatic_styles_and_required_namespace_bindings() {
    let xml = r##"<?xml version="1.0"?><o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:f="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"><o:font-face-decls/><o:automatic-styles><s:style s:name="ce1" s:family="table-cell"><s:table-cell-properties f:background-color="#fff"/></s:style></o:automatic-styles><o:body/></o:document-content>"##;
    let fragment = extract_automatic_styles(xml)
        .expect("test fixture or operation should succeed")
        .expect("test fixture or operation should succeed");
    assert!(fragment.xml.starts_with("<o:automatic-styles>"));
    assert!(fragment.xml.ends_with("</o:automatic-styles>"));
    let mut declarations = String::new();
    fragment.write_missing_namespaces(&mut declarations, ["o"]);
    assert!(!declarations.contains("xmlns:o="));
    assert!(declarations.contains("xmlns:s="));
    assert!(declarations.contains("xmlns:f="));
    let font_faces = extract_font_face_decls(xml)
        .expect("test fixture or operation should succeed")
        .expect("test fixture or operation should succeed");
    assert!(font_faces.xml.starts_with("<o:font-face-decls"));
}

#[test]
fn parses_ordered_inert_conditional_cell_styles_and_overrides() {
    let named = r#"<o:document-styles xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:f="urn:example:formula"><o:styles><s:style s:name="Red" s:family="table-cell"/><s:style s:name="Blue" s:family="table-cell"/></o:styles></o:document-styles>"#;
    let content = r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:f="urn:example:formula"><o:automatic-styles><s:style s:name="ce1" s:family="table-cell" s:parent-style-name="Default"><s:map s:condition="cell-content()&lt;1" s:apply-style-name="Red" s:base-cell-address="Sheet1.A1"/><s:map s:condition="f:is-true-formula([.A1]&gt;0)" s:apply-style-name="Blue"></s:map></s:style></o:automatic-styles></o:document-content>"#;
    let registry = CellStyleRegistry::parse(Some(named), content)
        .expect("test fixture or operation should succeed");
    let style = registry
        .conditional_style("ce1")
        .expect("test fixture or operation should succeed");
    assert_eq!(style.parent_style_name.as_deref(), Some("Default"));
    assert_eq!(style.rules.len(), 2);
    assert_eq!(style.rules[0].condition, "cell-content()<1");
    assert_eq!(style.rules[0].apply_style_name, "Red");
    assert_eq!(
        style.rules[0].base_cell_address.as_deref(),
        Some("Sheet1.A1")
    );
    assert_eq!(
        style.rules[1].formula_namespace,
        Some(formula::Namespace {
            prefix: "f".to_string(),
            uri: "urn:example:formula".to_string(),
        })
    );
    assert_eq!(registry.conditional_styles(), std::slice::from_ref(style));

    let override_content = format!(
        "{content}<o:automatic-styles xmlns:o=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:s=\"urn:oasis:names:tc:opendocument:xmlns:style:1.0\"><s:style s:name=\"ce1\" s:family=\"table-cell\"><s:map s:condition=\"cell-content()=2\" s:apply-style-name=\"Blue\"/></s:style></o:automatic-styles>"
    );
    let overridden = CellStyleRegistry::parse(Some(named), &override_content)
        .expect("test fixture or operation should succeed");
    assert_eq!(
        overridden
            .conditional_style("ce1")
            .expect("test fixture or operation should succeed")
            .rules[0]
            .condition,
        "cell-content()=2"
    );
}

#[test]
fn rejects_malformed_or_active_conditional_style_inputs() {
    let common = r#"<o:styles xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0"><s:style s:name="Target" s:family="table-cell"/></o:styles>"#;
    let missing_target = r#"<s:style xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0" s:name="ce1" s:family="table-cell"><s:map s:condition="cell-content()=1" s:apply-style-name="Missing"/></s:style>"#;
    assert!(CellStyleRegistry::parse(Some(common), missing_target).is_err());

    let non_empty_map = r#"<s:style xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0" s:name="ce1" s:family="table-cell"><s:map s:condition="cell-content()=1" s:apply-style-name="Target">run</s:map></s:style>"#;
    assert!(CellStyleRegistry::parse(Some(common), non_empty_map).is_err());

    let dtd = r#"<!DOCTYPE x [<!ENTITY run SYSTEM "file:///etc/passwd">]><s:style xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0" s:name="ce1" s:family="table-cell"/>"#;
    assert!(CellStyleRegistry::parse(Some(common), dtd).is_err());

    let oversized = "x".repeat(MAX_CONDITIONAL_ATTRIBUTE_BYTES + 1);
    let oversized = format!(
        "<s:style xmlns:s=\"urn:oasis:names:tc:opendocument:xmlns:style:1.0\" s:name=\"ce1\" s:family=\"table-cell\"><s:map s:condition=\"{oversized}\" s:apply-style-name=\"Target\"/></s:style>"
    );
    assert!(CellStyleRegistry::parse(Some(common), &oversized).is_err());

    let extension_only = r#"<s:style xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:c="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0" s:name="ce1" s:family="table-cell"><c:conditional-formats><c:condition c:value="1"/></c:conditional-formats></s:style>"#;
    let registry = CellStyleRegistry::parse(Some(common), extension_only)
        .expect("test fixture or operation should succeed");
    assert!(registry.conditional_styles().is_empty());
}

#[test]
fn parses_libreoffice_flat_conditional_style_fixture_when_available() {
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join(
        "../../test-data/libreoffice-core/sc/qa/unit/data/functions/financial/fods/couppcd.fods",
    );
    if !path.exists() {
        return;
    }
    let xml = std::fs::read_to_string(path).expect("test fixture or operation should succeed");
    let registry =
        CellStyleRegistry::parse(None, &xml).expect("test fixture or operation should succeed");
    let ce6 = registry
        .conditional_style("ce6")
        .expect("test fixture or operation should succeed");
    assert_eq!(ce6.rules.len(), 3);
    assert_eq!(ce6.rules[0].condition, "cell-content()=\"\"");
    assert_eq!(ce6.rules[1].apply_style_name, "Untitled1");
    assert_eq!(ce6.rules[2].base_cell_address.as_deref(), Some("Sheet1.B3"));
}
