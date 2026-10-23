use super::*;

use std::sync::Arc;

use litchi_core::OwnedSource;
use litchi_opc::{OpcPackage, PackURI, PackageWriter, SourceLimitKind, SourceLimits};

const REGISTERED_GRAPH_MEMBERS: &[&str] = &[
    "/ppt/slideLayouts/slideLayout1.xml",
    "/ppt/slideMasters/slideMaster1.xml",
    "/ppt/theme/theme1.xml",
];

#[test]
fn registered_layout_graph_checks_attribute_limits_before_semantic_projection() {
    let original = authored_fixture().unwrap();
    for member in REGISTERED_GRAPH_MEMBERS {
        let bytes = replace_part(&original, member, oversized_attribute);
        let limits = SourceLimits::builder()
            .max_xml_attribute_bytes(8_192)
            .unwrap()
            .build()
            .unwrap();
        let source = SourceBackedPresentation::from_read_at_with_options(
            Arc::new(OwnedSource::new(bytes)),
            crate::OpenOptions::default().source_limits(limits),
        )
        .unwrap();
        let destination = SourceBackedPresentationEditor::from_read_at(Arc::new(OwnedSource::new(
            original.clone(),
        )))
        .unwrap();

        assert!(
            matches!(
                destination.plan_cross_slide_copy(&source, 0, 0, 1),
                Err(Error::Opc(litchi_opc::OpcError::Limit {
                    kind: SourceLimitKind::XmlAttributeBytes,
                    ..
                }))
            ),
            "{member}"
        );
    }
}

#[test]
fn registered_layout_graph_checks_event_limits_before_semantic_projection() {
    let original = authored_fixture().unwrap();
    for member in REGISTERED_GRAPH_MEMBERS {
        let bytes = replace_part(&original, member, oversized_events);
        let limits = SourceLimits::builder()
            .max_events(4_096)
            .unwrap()
            .build()
            .unwrap();
        let source = SourceBackedPresentation::from_read_at_with_options(
            Arc::new(OwnedSource::new(bytes)),
            crate::OpenOptions::default().source_limits(limits),
        )
        .unwrap();
        let destination = SourceBackedPresentationEditor::from_read_at(Arc::new(OwnedSource::new(
            original.clone(),
        )))
        .unwrap();

        assert!(
            matches!(
                destination.plan_cross_slide_copy(&source, 0, 0, 1),
                Err(Error::Opc(litchi_opc::OpcError::Limit {
                    kind: SourceLimitKind::Events,
                    ..
                }))
            ),
            "{member}"
        );
    }
}

fn authored_fixture() -> Result<Vec<u8>> {
    let mut builder = crate::builder::Builder::new()?;
    {
        let presentation = builder.presentation()?;
        presentation
            .add_slide()?
            .add_text_box("source", 0, 0, 914_400, 914_400);
    }
    builder.build()?.to_bytes()
}

fn replace_part(bytes: &[u8], path: &str, rewrite: impl FnOnce(&[u8]) -> Vec<u8>) -> Vec<u8> {
    let mut package = OpcPackage::from_bytes(bytes, SourceLimits::default()).unwrap();
    let part = package.get_part_mut(&PackURI::new(path).unwrap()).unwrap();
    let replacement = rewrite(part.blob());
    part.set_blob(replacement);
    PackageWriter::to_bytes(&package).unwrap()
}

fn oversized_attribute(xml: &[u8]) -> Vec<u8> {
    let root_start = xml
        .windows(2)
        .position(|window| window == b"<p" || window == b"<a")
        .unwrap();
    let root_end = root_start
        + xml[root_start..]
            .iter()
            .position(|byte| *byte == b'>')
            .unwrap();
    let mut replacement = Vec::with_capacity(xml.len() + 16 * 1024);
    replacement.extend_from_slice(&xml[..root_end]);
    replacement.extend_from_slice(br#" data=""#);
    replacement.extend(std::iter::repeat_n(b'x', 16 * 1024));
    replacement.push(b'"');
    replacement.extend_from_slice(&xml[root_end..]);
    replacement
}

fn oversized_events(xml: &[u8]) -> Vec<u8> {
    let closing = xml.windows(2).rposition(|window| window == b"</").unwrap();
    let mut replacement = Vec::with_capacity(xml.len() + 256 * 1024);
    replacement.extend_from_slice(&xml[..closing]);
    for _ in 0..8_192 {
        replacement.extend_from_slice(b"<!-- source-limit-event -->");
    }
    replacement.extend_from_slice(&xml[closing..]);
    replacement
}
