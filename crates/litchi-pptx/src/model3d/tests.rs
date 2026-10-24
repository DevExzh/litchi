#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use super::{codec, model::Data};
use crate::Error;

#[test]
fn codec_rejects_an_incomplete_model_sequence() {
    let xml = br#"<am3d:model3d xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d"><am3d:spPr/><am3d:camera/></am3d:model3d>"#;
    assert!(matches!(
        codec::read(xml, codec::Limits::DEFAULT),
        Err(Error::Drawing(_))
    ));
}

#[test]
fn codec_policy_admits_input_before_namespace_preparation() {
    let xml = br#"<am3d:model3d xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d"><am3d:spPr/></am3d:model3d>"#;
    let limits = codec::Limits::new(8, 64, 64).expect("valid narrow policy");
    assert!(matches!(
        codec::read(xml, limits),
        Err(Error::Drawing(litchi_drawingml::Error::Limit {
            kind: litchi_drawingml::LimitKind::XmlBytes,
            observed,
            maximum: 8,
        })) if observed == xml.len()
    ));
    assert!(codec::Limits::new(0, 1, 1).is_none());
}

#[test]
fn codec_reports_typed_prepared_size_limit_at_the_adapter_boundary() {
    let xml = br#"<am3d:model3d r:embed="rId" ><am3d:spPr/></am3d:model3d>"#;
    let limits = codec::Limits::new(xml.len(), 64, 64).expect("valid narrow policy");
    assert!(matches!(
        codec::read(xml, limits),
        Err(Error::Drawing(litchi_drawingml::Error::Limit {
            kind: litchi_drawingml::LimitKind::XmlBytes,
            observed,
            maximum,
        })) if observed > xml.len() && maximum == xml.len()
    ));
}

#[test]
fn codec_write_refuses_output_before_retaining_an_opaque_copy() {
    let future = "x".repeat(32 * 1024);
    let xml = format!(
        r#"<am3d:model3d xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d" xmlns:f="urn:future"><am3d:spPr/><am3d:camera/><am3d:trans/><am3d:objViewport/><f:future>{future}</f:future></am3d:model3d>"#
    );
    let scene = codec::read(xml.as_bytes(), codec::Limits::DEFAULT).expect("model3d scene");
    let limits = codec::Limits::new(1 << 20, 1 << 20, 64).expect("valid output policy");
    assert!(matches!(
        codec::write(&scene, limits),
        Err(Error::Drawing(litchi_drawingml::Error::Limit {
            kind: litchi_drawingml::LimitKind::OutputBytes,
            maximum: 64,
            observed,
        })) if observed > 64
    ));
}

#[test]
fn data_rejects_an_unbounded_payload() {
    let data = vec![0_u8; super::MAX_MODEL_BYTES.saturating_add(1)];
    assert!(
        matches!(Data::new(data), Err(Error::Limit { resource, .. }) if resource == "model3d payload bytes")
    );
}
