use litchi_core::Position;
use litchi_docx::document::{Error, LimitKind, raw};

#[test]
fn depth_error_exposes_typed_kind_and_canonical_counts() {
    let xml = format!(
        "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>{}{}</w:body></w:document>",
        "<w:customXml>".repeat(256),
        "</w:customXml>".repeat(256)
    );
    assert!(matches!(
        raw::Tree::from_xml(xml.into_bytes(), raw::Limits::default()),
        Err(Error::Limit {
            kind: LimitKind::XmlDepth,
            observed: 257,
            maximum: 256
        })
    ));
}

#[test]
fn operation_limit_preserves_staging_and_has_content_free_diagnostics() {
    let xml = b"<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body><w:p><w:r><w:t>initial</w:t></w:r></w:p></w:body></w:document>".to_vec();
    let mut edit = raw::Tree::from_xml(xml, raw::Limits::default()).unwrap();
    for i in 0..65_536 {
        edit.replace_paragraph_text(
            Position::new(0),
            if i % 2 == 0 {
                "private-author-a"
            } else {
                "private-author-b"
            },
        )
        .unwrap();
    }
    let before = edit.projected().xml_bytes().to_vec();
    let error = edit
        .replace_paragraph_text(Position::new(0), "private-author-c")
        .unwrap_err();
    assert!(matches!(
        error,
        Error::Limit {
            kind: LimitKind::Operations,
            observed: 65_537,
            maximum: 65_536
        }
    ));
    assert!(!format!("{error} {error:?}").contains("private-author"));
    assert_eq!(
        format!("{error:?}"),
        "Limit { kind: Operations, observed: 65537, maximum: 65536 }"
    );
    assert_eq!(edit.projected().xml_bytes(), before);
    assert_eq!(edit.patch().unwrap().operations().len(), 65_536);
}

fn raw_limits(
    depth: usize,
    nodes: usize,
    attributes: usize,
    attribute_bytes: usize,
) -> raw::Limits {
    raw::Limits::new(
        32 * 1024 * 1024,
        depth,
        nodes,
        attributes,
        attribute_bytes,
        65_536,
        16 * 1024 * 1024,
    )
    .unwrap()
}

#[test]
fn raw_xml_policy_admits_selected_depth_nodes_and_attributes_before_layout() {
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p a="1" b="2"/></w:body></w:document>"#;
    assert!(matches!(
        raw::Tree::from_xml(xml.to_vec(), raw_limits(2, 100, 8, 100)),
        Err(Error::Limit {
            kind: LimitKind::XmlDepth,
            observed: 3,
            maximum: 2
        })
    ));
    assert!(matches!(
        raw::Tree::from_xml(xml.to_vec(), raw_limits(8, 2, 8, 100)),
        Err(Error::Limit {
            kind: LimitKind::XmlElements,
            observed: 3,
            maximum: 2
        })
    ));
    assert!(matches!(
        raw::Tree::from_xml(xml.to_vec(), raw_limits(8, 100, 1, 100)),
        Err(Error::Limit {
            kind: LimitKind::XmlAttributes,
            observed: 2,
            maximum: 1
        })
    ));
    assert!(matches!(
        raw::Tree::from_xml(xml.to_vec(), raw_limits(8, 100, 8, 8)),
        Err(Error::Limit {
            kind: LimitKind::XmlAttributeBytes,
            maximum: 8,
            ..
        })
    ));
}

#[test]
fn raw_output_replay_cannot_widen_limits_and_failed_mutation_is_atomic() {
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>before</w:t></w:r></w:p></w:body></w:document>"#;
    let limits = raw_limits(8, 100, 8, 100);
    let mut tree = raw::Tree::from_xml(xml.to_vec(), limits).unwrap();
    let before = tree.projected().xml_bytes().to_vec();
    assert!(
        tree.replace_paragraph_text(Position::new(99), "after")
            .is_err()
    );
    assert_eq!(tree.projected().xml_bytes(), before.as_slice());
    assert_eq!(tree.to_bytes(limits).unwrap(), before);
    assert_eq!(
        tree.to_bytes(limits).unwrap(),
        tree.to_bytes(limits).unwrap()
    );

    let wider = raw::Limits::default();
    assert!(matches!(
        tree.to_bytes(wider),
        Err(Error::Limit {
            kind: LimitKind::XmlDepth,
            observed: 256,
            maximum: 8
        })
    ));
    let patch = tree.patch().unwrap();
    assert!(matches!(
        tree.apply_patch(&patch, wider),
        Err(Error::Limit {
            kind: LimitKind::XmlDepth,
            observed: 256,
            maximum: 8
        })
    ));
}

#[test]
fn raw_reopen_forward_inverse_and_second_edit_preserve_the_bounded_codec_contract() {
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>before</w:t></w:r></w:p></w:body></w:document>"#;
    let limits = raw_limits(8, 100, 8, 100);
    let mut authored = raw::Tree::from_xml(xml.to_vec(), limits).unwrap();
    authored
        .replace_paragraph_text(Position::new(0), "forward")
        .unwrap();
    let forward = authored.patch().unwrap();
    let inverse = forward.inverse();

    let reopened = raw::Tree::from_xml(xml.to_vec(), limits).unwrap();
    let forwarded = reopened.apply_patch(&forward, limits).unwrap();
    assert_eq!(
        forwarded.to_bytes(limits).unwrap(),
        authored.to_bytes(limits).unwrap()
    );
    let restored = forwarded.apply_patch(&inverse, limits).unwrap();
    assert_eq!(restored.to_bytes(limits).unwrap(), xml);

    let mut second = restored;
    second
        .replace_paragraph_text(Position::new(0), "second")
        .unwrap();
    assert!(
        std::str::from_utf8(&second.to_bytes(limits).unwrap())
            .unwrap()
            .contains("second")
    );
}
