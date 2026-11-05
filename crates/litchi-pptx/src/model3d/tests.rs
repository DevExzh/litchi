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
fn codec_writes_alternate_prefixes_with_existing_bindings_and_roundtrips() {
    let xml = br#"<am3d:model3d xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d" xmlns:rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:f="urn:future" rel:embed="rIdModel"><am3d:spPr/><am3d:camera/><am3d:trans/><am3d:objViewport/><f:future f:value="keep"/></am3d:model3d>"#;
    let scene = codec::read(xml, codec::Limits::DEFAULT).expect("alternate-prefix scene");
    let output = codec::write(&scene, codec::Limits::DEFAULT).expect("scene must serialize");
    assert!(output.starts_with(b"<am3d:model3d"));
    assert!(
        output
            .windows(b"xmlns:am3d=".len())
            .any(|window| window == b"xmlns:am3d=")
    );
    assert!(
        output
            .windows(b"xmlns:rel=".len())
            .any(|window| window == b"xmlns:rel=")
    );
    assert!(
        !output
            .windows(b"xmlns:m3d=".len())
            .any(|window| window == b"xmlns:m3d=")
    );
    assert!(
        !output
            .windows(b"xmlns:r=".len())
            .any(|window| window == b"xmlns:r=")
    );
    assert!(
        output
            .windows(b"f:future".len())
            .any(|window| window == b"f:future")
    );

    let reparsed = codec::read(&output, codec::Limits::DEFAULT).expect("stable output must parse");
    assert_eq!(reparsed.child_count(), scene.child_count());
    assert_eq!(reparsed.has_raster(), scene.has_raster());
}

#[test]
fn codec_preserves_foreign_colliding_prefixes_in_opaque_children() {
    let xml = br#"<am3d:model3d xmlns:m3d="urn:foreign-model" xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d" xmlns:r="urn:foreign-relationship" xmlns:rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships" rel:embed="rIdModel"><am3d:spPr/><am3d:camera/><am3d:trans/><am3d:objViewport/><m3d:foreign/><r:foreign/></am3d:model3d>"#;
    let scene = codec::read(xml, codec::Limits::DEFAULT).expect("colliding-prefix scene");
    let unknown = scene
        .unknown_children()
        .map(|value| (value.local_name(), value.namespace()))
        .collect::<Vec<_>>();
    assert_eq!(
        unknown,
        vec![
            ("foreign", "urn:foreign-model"),
            ("foreign", "urn:foreign-relationship"),
        ]
    );

    let output = codec::write(&scene, codec::Limits::DEFAULT).expect("scene must serialize");
    assert!(output.starts_with(b"<am3d:model3d"));
    assert_eq!(
        output
            .windows(b"xmlns:m3d=".len())
            .filter(|window| *window == b"xmlns:m3d=")
            .count(),
        1
    );
    assert_eq!(
        output
            .windows(b"xmlns:r=".len())
            .filter(|window| *window == b"xmlns:r=")
            .count(),
        1
    );
    assert!(
        output
            .windows(b"<m3d:foreign/>".len())
            .any(|window| window == b"<m3d:foreign/>")
    );
    assert!(
        output
            .windows(b"<r:foreign/>".len())
            .any(|window| window == b"<r:foreign/>")
    );

    let reparsed = codec::read(&output, codec::Limits::DEFAULT).expect("collision-safe output");
    let reparsed_unknown = reparsed
        .unknown_children()
        .map(|value| (value.local_name(), value.namespace()))
        .collect::<Vec<_>>();
    assert_eq!(reparsed_unknown, unknown);
}

#[test]
fn codec_counts_selected_prefix_output_before_alternate_prefix_refusal() {
    let xml = br#"<am3d:model3d xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d" xmlns:rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships" rel:embed="rIdModel"><am3d:spPr/><am3d:camera/><am3d:trans/><am3d:objViewport/></am3d:model3d>"#;
    let scene = codec::read(xml, codec::Limits::DEFAULT).expect("alternate-prefix scene");
    let output = codec::write(&scene, codec::Limits::DEFAULT).expect("scene must serialize");
    let limits = codec::Limits::new(1 << 20, 1 << 20, output.len() - 1)
        .expect("one-byte narrower output policy");
    assert!(matches!(
        codec::write(&scene, limits),
        Err(Error::Drawing(litchi_drawingml::Error::Limit {
            kind: litchi_drawingml::LimitKind::OutputBytes,
            observed,
            maximum,
        })) if observed == output.len() && maximum + 1 == observed
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
            observed,
            maximum: 64,
        })) if observed > 64
    ));
}

#[test]
fn owner_source_limits_tighten_xml_tree_and_payload_budgets() {
    let source = litchi_opc::SourceLimits::builder()
        .max_part_bytes(1024)
        .unwrap()
        .max_total_part_bytes(2048)
        .unwrap()
        .max_events(7)
        .unwrap()
        .max_depth(3)
        .unwrap()
        .max_relationship_target_bytes(5)
        .unwrap()
        .build()
        .unwrap();
    let limits = codec::Limits::from_source_limits(source);
    assert_eq!(limits.max_xml_bytes(), 1024);
    assert_eq!(limits.max_output_bytes(), 1024);
    assert_eq!(limits.max_nodes(), 7);
    assert_eq!(limits.max_depth(), 3);
    assert_eq!(limits.max_model_bytes(), 1024);
    assert_eq!(limits.max_preview_bytes(), 1024);
    assert_eq!(limits.max_total_payload_bytes(), 2048);
    assert_eq!(limits.max_link_bytes(), 5);
}

#[test]
fn replay_validator_rejects_model_payload_above_recipient_profile() {
    let source = litchi_opc::SourceLimits::builder()
        .max_part_bytes(8)
        .unwrap()
        .max_total_part_bytes(16)
        .unwrap()
        .build()
        .unwrap();
    let limits = codec::Limits::from_source_limits(source);
    let name = litchi_opc::PackURI::new("/ppt/media/model3d.glb").unwrap();
    let part = litchi_opc::part::BlobPart::new(
        name.clone(),
        super::MODEL_CONTENT_TYPE.to_owned(),
        vec![0; limits.max_model_bytes() + 1],
    );
    assert!(matches!(
        super::package::validate_replay_part(
            &litchi_opc::OpcPackage::new(),
            &name,
            &part,
            limits,
        ),
        Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::PayloadBytes),
            maximum,
            ..
        }) if maximum == 8
    ));
}

#[test]
fn data_rejects_an_unbounded_payload() {
    let data = vec![0_u8; super::MAX_MODEL_BYTES.saturating_add(1)];
    assert!(matches!(
        Data::new(data),
        Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::PayloadBytes),
            ..
        })
    ));
}
