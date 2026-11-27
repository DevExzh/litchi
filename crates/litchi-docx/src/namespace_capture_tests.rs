use super::*;

#[test]
fn unchanged_namespace_scopes_share_storage_and_shadowed_scopes_remain_distinct() {
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="urn:original"><w:body><w:p/><w:p/><w:p xmlns:x="urn:changed"/><w:p/></w:body></w:document>"#;
    let mut scopes = Vec::new();
    scan_word_element_ranges_with_context(xml, &[], &[b"p"], |_, _, _, scope| {
        scopes.push(scope);
        Ok(())
    })
    .unwrap();
    assert_eq!(scopes.len(), 4);
    assert!(Arc::ptr_eq(&scopes[0], &scopes[1]));
    assert_ne!(scopes[1], scopes[2]);
    assert_eq!(scopes[0], scopes[3]);
    assert!(!Arc::ptr_eq(&scopes[2], &scopes[3]));

    let blocks = crate::parts::document_part::body_block_ranges(xml).unwrap();
    assert!(Arc::ptr_eq(&blocks[0].3, &blocks[1].3));
    let snapshot = crate::document::Snapshot::from_xml(xml.to_vec()).unwrap();
    assert_eq!(snapshot.paragraph_count(), 4);
}

#[test]
fn namespace_capture_budget_charges_new_storage_before_capture_and_reuses_at_limit() {
    let mut resolver = NamespaceResolver::default();
    resolver
        .add(PrefixDeclaration::Named(b"x"), Namespace(b"urn:one"))
        .unwrap();
    let (count, bytes) = namespace_requirements(&resolver).unwrap();
    let exact = count * size_of::<(Option<Vec<u8>>, Vec<u8>)>() + bytes;
    let mut cache = NamespaceCapture {
        maximum: exact,
        ..NamespaceCapture::default()
    };
    let first = cache.capture(&resolver).unwrap();
    let second = cache.capture(&resolver).unwrap();
    assert!(Arc::ptr_eq(&first, &second));
    assert_eq!(cache.used, exact);
    let mut short = NamespaceCapture {
        maximum: exact - 1,
        ..NamespaceCapture::default()
    };
    assert!(short.capture(&resolver).is_err());
    assert_eq!(short.used, 0);
    assert!(short.last.is_none());
    resolver
        .add(PrefixDeclaration::Named(b"y"), Namespace(b"urn:two"))
        .unwrap();
    assert!(cache.capture(&resolver).is_err());
    assert_eq!(cache.used, exact);
    assert!(Arc::ptr_eq(cache.last.as_ref().unwrap(), &first));
}

#[test]
fn namespace_free_empty_root_is_selected_without_enabling_late_foreign_fallback() {
    let mut selected = 0;
    scan_word_element_ranges(b"<w:p/>", &[b"p"], |_, _, _| {
        selected += 1;
        Ok(())
    })
    .unwrap();
    assert_eq!(selected, 1);
    let mut selected = 0;
    scan_word_element_ranges(
        b"<x:root xmlns:x=\"urn:foreign\"><w:p/></x:root>",
        &[b"p"],
        |_, _, _| {
            selected += 1;
            Ok(())
        },
    )
    .unwrap();
    assert_eq!(selected, 0);
}
