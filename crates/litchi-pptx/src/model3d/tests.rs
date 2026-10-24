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
        Err(Error::Limit { resource, .. }) if resource == "model3d XML bytes"
    ));
    assert!(codec::Limits::new(0, 1, 1).is_none());
}

#[test]
fn data_rejects_an_unbounded_payload() {
    let data = vec![0_u8; super::MAX_MODEL_BYTES.saturating_add(1)];
    assert!(
        matches!(Data::new(data), Err(Error::Limit { resource, .. }) if resource == "model3d payload bytes")
    );
}
