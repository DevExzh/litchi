#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use super::{EXTENSION_URI, NAMESPACE, codec, model::Outcome, transaction};

const PML: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";

fn shape_with_classification(outcome: &str) -> Vec<u8> {
    format!(
        r#"<p:sp xmlns:p="{PML}" xmlns:p184="{NAMESPACE}"><p:nvSpPr><p:cNvPr id="1" name="Box"/><p:cNvSpPr/><p:nvPr><p:extLst><p:ext uri="{EXTENSION_URI}"><p184:classification val="{outcome}"/></p:ext></p:extLst></p:nvPr></p:nvSpPr><p:spPr/></p:sp>"#
    )
    .into_bytes()
}

#[test]
fn reads_and_replaces_typed_classification_without_normalizing_the_shape() {
    let original = shape_with_classification("hdr");
    let source = codec::read(&original).unwrap();
    assert_eq!(
        source
            .snapshot
            .as_ref()
            .and_then(super::model::Snapshot::outcome),
        Some(Outcome::Header)
    );

    let updated = transaction::set(
        &original,
        &source,
        Outcome::Watermark,
        crate::shape::Limits::DEFAULT.output_bytes(),
    )
    .unwrap();
    let parsed = codec::read(&updated).unwrap();
    assert_eq!(
        parsed
            .snapshot
            .as_ref()
            .and_then(super::model::Snapshot::outcome),
        Some(Outcome::Watermark)
    );
    assert!(
        updated
            .windows(b"<p:spPr/>".len())
            .any(|window| window == b"<p:spPr/>")
            || updated
                .windows(b"<p:spPr>".len())
                .any(|window| window == b"<p:spPr>")
    );
}

#[test]
fn unknown_extensions_survive_first_class_authoring() {
    let xml = format!(
        r#"<p:sp xmlns:p="{PML}" xmlns:f="urn:future"><p:nvSpPr><p:cNvPr id="1" name="Box"/><p:cNvSpPr/><p:nvPr><p:extLst><p:ext uri="urn:future"><f:payload answer="42"/></p:ext></p:extLst></p:nvPr></p:nvSpPr><p:spPr/></p:sp>"#
    )
    .into_bytes();
    let source = codec::read(&xml).unwrap();
    assert!(source.snapshot.is_none());

    let updated = transaction::set(
        &xml,
        &source,
        Outcome::Footer,
        crate::shape::Limits::DEFAULT.output_bytes(),
    )
    .unwrap();
    assert!(
        updated
            .windows(b"<f:payload answer=\"42\">".len())
            .any(|window| window == b"<f:payload answer=\"42\">")
            || updated
                .windows(b"<f:payload answer=\"42\"/>".len())
                .any(|window| window == b"<f:payload answer=\"42\"/>")
    );
    assert_eq!(
        codec::read(&updated)
            .unwrap()
            .snapshot
            .and_then(|snapshot| snapshot.outcome()),
        Some(Outcome::Footer)
    );
}

#[test]
fn owner_scope_resolves_aliases_defaults_and_empty_sibling_rebindings() {
    for namespace in [PML, "http://purl.oclc.org/ooxml/presentationml/main"] {
        for prefix in ["s:", ""] {
            let declaration = if prefix.is_empty() {
                format!("xmlns=\"{namespace}\"")
            } else {
                format!("xmlns:s=\"{namespace}\"")
            };
            let shadow = if prefix.is_empty() {
                "xmlns=\"urn:foreign\""
            } else {
                "xmlns:s=\"urn:foreign\""
            };
            let shape = format!(
                "<{prefix}sp><{prefix}nvSpPr><{prefix}cNvPr id=\"1\"/><{prefix}cNvSpPr/><{prefix}nvPr><{prefix}extLst><{prefix}ext uri=\"{EXTENSION_URI}\"><k:classification val=\"hdr\"/></{prefix}ext></{prefix}extLst></{prefix}nvPr></{prefix}nvSpPr><{prefix}spPr/></{prefix}sp>"
            );
            let owner = format!(
                "<{prefix}sld {declaration} xmlns:k=\"{NAMESPACE}\"><ignored {shadow}/>{shape}</{prefix}sld>"
            );
            let start = owner.find(&shape).unwrap();
            let span = start..start + shape.len();
            let source = codec::read_in_owner(owner.as_bytes(), span.clone()).unwrap();
            assert_eq!(
                source.snapshot.as_ref().unwrap().outcome(),
                Some(Outcome::Header)
            );
            let changed =
                transaction::set(shape.as_bytes(), &source, Outcome::Watermark, 1 << 20).unwrap();
            let next = codec::replace(owner.as_bytes(), span.clone(), &changed, 1 << 20).unwrap();
            assert_eq!(
                codec::read_in_owner(&next, start..start + changed.len())
                    .unwrap()
                    .snapshot
                    .unwrap()
                    .outcome(),
                Some(Outcome::Watermark)
            );
            let foreign = owner.replace(NAMESPACE, "urn:foreign-classification");
            let start = foreign.find(&shape).unwrap();
            assert!(
                codec::read_in_owner(foreign.as_bytes(), start..start + shape.len())
                    .unwrap()
                    .snapshot
                    .is_none()
            );
        }
    }
}
