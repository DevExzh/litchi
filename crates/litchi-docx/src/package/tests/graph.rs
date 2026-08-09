use super::*;

#[test]
fn saves_and_reopens_package() {
    let file = NamedTempFile::with_suffix(".docx").unwrap();
    let mut package = Package::new().unwrap();
    package
        .document_mut()
        .unwrap()
        .add_paragraph_with_text("round-trip text");
    package.save(file.path()).unwrap();

    let mut reopened = Package::open(file.path()).unwrap();
    assert!(
        reopened
            .document()
            .unwrap()
            .text()
            .unwrap()
            .contains("round-trip text")
    );

    reopened
        .document_mut()
        .unwrap()
        .add_paragraph_with_text("appended after reopen");
    reopened.save(file.path()).unwrap();
    let reopened_again = Package::open(file.path()).unwrap();
    let text = reopened_again.document().unwrap().text().unwrap();
    assert!(text.contains("round-trip text"));
    assert!(text.contains("appended after reopen"));
}

#[test]
fn numbering_patch_publishes_through_the_package_graph() {
    let source_xml = br#"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="1"/></w:numbering>"#;
    let numbering_uri = PackURI::new("/word/numbering.xml").unwrap();
    let mut package = Package::new().unwrap();
    package
        .edit_opc(|opc| {
            opc.get_part_mut(&numbering_uri)
                .map_err(Error::from)?
                .set_blob(source_xml.to_vec());
            Ok(())
        })
        .unwrap();

    let source = package.numbering_snapshot().unwrap().unwrap();
    let mut edit = source.edit();
    edit.set_restart_numbering_after_break(1, Some(false))
        .unwrap();
    let commit = edit.commit().unwrap();
    let published = package
        .apply_numbering_patch(&source, commit.patch())
        .unwrap();
    assert_eq!(
        published.restart_numbering_after_break(1).unwrap(),
        Some(false)
    );

    let mut output = Cursor::new(Vec::new());
    package.to_plain_stream(&mut output).unwrap();
    let reopened = Package::from_reader(Cursor::new(output.into_inner())).unwrap();
    assert_eq!(
        reopened
            .numbering_snapshot()
            .unwrap()
            .unwrap()
            .restart_numbering_after_break(1)
            .unwrap(),
        Some(false)
    );
}

#[test]
fn document_patch_publishes_atomically_and_reopens() {
    let source_xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>before</w:t></w:r></w:p><w:sectPr/></w:body></w:document>"#;
    let document_uri = PackURI::new("/word/document.xml").unwrap();
    let mut package = Package::new().unwrap();
    package
        .edit_opc(|opc| {
            opc.get_part_mut(&document_uri)?
                .set_blob(source_xml.to_vec());
            Ok(())
        })
        .unwrap();

    let source = package.document_snapshot().unwrap();
    let mut edit = source.edit();
    edit.replace_paragraph_text(litchi_core::Position::new(0), "after and longer")
        .unwrap();
    let commit = edit.commit().unwrap();
    let published = package.apply_document_patch(commit.patch()).unwrap();
    assert_eq!(
        published
            .paragraph(litchi_core::Position::new(0))
            .unwrap()
            .text()
            .unwrap(),
        "after and longer"
    );

    let before_stale = package.opc.main_document_part().unwrap().blob().to_vec();
    assert!(package.apply_document_patch(commit.patch()).is_err());
    assert_eq!(
        package.opc.main_document_part().unwrap().blob(),
        before_stale
    );

    let mut output = Cursor::new(Vec::new());
    package.to_plain_stream(&mut output).unwrap();
    let reopened = Package::from_reader(Cursor::new(output.into_inner())).unwrap();
    assert_eq!(
        reopened.document().unwrap().text().unwrap(),
        "after and longer"
    );

    package
        .apply_document_patch(&commit.patch().inverse())
        .unwrap();
    assert_eq!(package.opc.main_document_part().unwrap().blob(), source_xml);
}

#[test]
fn failed_stream_keeps_document_and_properties_retryable() {
    let mut package = Package::new().unwrap();
    package
        .document_mut()
        .unwrap()
        .add_paragraph_with_text("retryable document");
    package.put_props(Props::new().title("retryable properties"));
    let document_before = package.opc.main_document_part().unwrap().blob_arc();
    let core_properties_uri = PackURI::new("/docProps/core.xml").unwrap();
    let core_properties_before = package
        .opc
        .get_part(&core_properties_uri)
        .unwrap()
        .blob_arc();

    assert!(package.to_plain_stream(FailingWriter).is_err());
    assert_eq!(
        package.opc.main_document_part().unwrap().blob(),
        document_before.as_slice()
    );
    assert!(std::sync::Arc::ptr_eq(
        &document_before,
        &package.opc.main_document_part().unwrap().blob_arc()
    ));
    assert_eq!(
        package.opc.get_part(&core_properties_uri).unwrap().blob(),
        core_properties_before.as_slice()
    );
    assert!(std::sync::Arc::ptr_eq(
        &core_properties_before,
        &package
            .opc
            .get_part(&core_properties_uri)
            .unwrap()
            .blob_arc()
    ));
    assert!(
        package
            .mutable_doc
            .as_ref()
            .is_some_and(MutableDocument::is_modified)
    );
    assert!(package.properties.is_dirty());

    package
        .document_mut()
        .unwrap()
        .add_paragraph_with_text("second attempt");
    let mut output = Cursor::new(Vec::new());
    package.to_plain_stream(&mut output).unwrap();
    assert!(!output.into_inner().is_empty());
    assert!(!package.properties.is_dirty());
}

#[test]
fn panicking_stream_restores_package_and_retryable_state() {
    let mut package = Package::new().unwrap();
    {
        let document = package.document_mut().unwrap();
        document.add_heading("panic-safe heading", 1).unwrap();
        document.add_toc(crate::TableOfContents::new()).unwrap();
    }
    package.put_props(Props::new().title("panic-safe properties"));
    package
        .custom_props_mut()
        .insert("RetryMarker", "panic-safe custom property")
        .unwrap();

    let package_before = litchi_opc::PackageWriter::to_bytes(&package.opc).unwrap();
    let document_before = package.mutable_doc.as_ref().unwrap().to_xml().unwrap();

    let unwind = catch_unwind(AssertUnwindSafe(|| {
        package.to_plain_stream(PanickingWriter).unwrap();
    }));
    assert!(unwind.is_err());
    assert_eq!(
        litchi_opc::PackageWriter::to_bytes(&package.opc).unwrap(),
        package_before
    );
    let document_after = package.mutable_doc.as_ref().unwrap().to_xml().unwrap();
    // TOC materialization is a one-shot semantic mutation: it consumes the
    // pending configuration and inserts the generated field before the sink
    // runs. The rollback restores that same writer value and its edit
    // intent, so retryability is checked through the retained heading/TOC
    // semantics rather than an impossible byte-identical writer snapshot.
    assert_ne!(document_after, document_before);
    assert!(document_after.contains("panic-safe heading"));
    assert!(document_after.contains("TOC"));
    assert!(
        package
            .mutable_doc
            .as_ref()
            .is_some_and(MutableDocument::is_modified)
    );
    assert!(package.properties.is_dirty());

    let mut output = Cursor::new(Vec::new());
    package.to_plain_stream(&mut output).unwrap();
    let output = output.into_inner();
    assert!(!output.is_empty());
    assert!(!package.properties.is_dirty());
    let reopened = Package::from_opc_package(OpcPackage::from_bytes(&output).unwrap()).unwrap();
    assert!(
        reopened
            .document()
            .unwrap()
            .text()
            .unwrap()
            .contains("panic-safe heading")
    );
    assert_eq!(
        reopened
            .document()
            .unwrap()
            .table_of_contents_count()
            .unwrap(),
        1
    );
    assert!(reopened.custom_props().contains("RetryMarker"));
}

#[test]
fn unchanged_stream_preserves_exact_bytes_and_part_payload_sharing() {
    let file = NamedTempFile::with_suffix(".docx").unwrap();
    let mut source = Package::new().unwrap();
    source
        .document_mut()
        .unwrap()
        .add_paragraph_with_text("unchanged package");
    source.save(file.path()).unwrap();

    let mut package = Package::open(file.path()).unwrap();
    let before = litchi_opc::PackageWriter::to_bytes(package.opc_package()).unwrap();
    let document_uri = PackURI::new("/word/document.xml").unwrap();
    let core_properties_uri = PackURI::new("/docProps/core.xml").unwrap();
    let document_before = package.opc.get_part(&document_uri).unwrap().blob_arc();
    let core_properties_before = package
        .opc
        .get_part(&core_properties_uri)
        .unwrap()
        .blob_arc();

    let mut output = Cursor::new(Vec::new());
    package.to_plain_stream(&mut output).unwrap();
    assert_eq!(output.into_inner(), before);
    assert!(std::sync::Arc::ptr_eq(
        &document_before,
        &package.opc.get_part(&document_uri).unwrap().blob_arc()
    ));
    assert!(std::sync::Arc::ptr_eq(
        &core_properties_before,
        &package
            .opc
            .get_part(&core_properties_uri)
            .unwrap()
            .blob_arc()
    ));
}

#[cfg(feature = "automatic-fonts")]
#[test]
fn raw_opc_rejects_automatic_font_embedding_policy() {
    let file = NamedTempFile::with_suffix(".docx").unwrap();
    let mut source = Package::new().unwrap();
    source
        .document_mut()
        .unwrap()
        .add_paragraph_with_text("preserved text");
    source.save(file.path()).unwrap();

    let mut opened = Package::open(file.path()).unwrap();
    let _ = opened.document_mut().unwrap();
    opened
        .set_font_embedding(litchi_fonts::embedding::Mode::Subset)
        .unwrap();
    assert!(matches!(
        opened.edit_opc(|_| Ok(())),
        Err(Error::UnsafeEdit {
            operation: "edit_opc",
            ..
        })
    ));
}

#[test]
fn raw_opc_transaction_publishes_candidate_and_disables_writer() {
    let mut package = Package::new().unwrap();
    let marker = PackURI::new("/custom/raw-edit-marker.bin").unwrap();

    package
        .edit_opc(|candidate| {
            candidate.try_add_part(Box::new(BlobPart::new(
                marker.clone(),
                "application/octet-stream".to_string(),
                b"raw edit".to_vec(),
            )))?;
            Ok::<_, Error>(())
        })
        .unwrap();

    assert_eq!(
        package.opc_package().get_part(&marker).unwrap().blob(),
        b"raw edit"
    );
    assert!(matches!(
        package.document_mut(),
        Err(Error::UnsafeEdit {
            operation: "document_mut",
            ..
        })
    ));
}

#[test]
fn failed_raw_opc_transaction_preserves_graph_and_writer_state() {
    let mut package = Package::new().unwrap();
    let document_uri = PackURI::new("/word/document.xml").unwrap();
    let original = package
        .opc_package()
        .get_part(&document_uri)
        .unwrap()
        .blob_arc();

    let error = package
        .edit_opc(|candidate| {
            candidate.remove_part(&document_uri);
            Ok::<_, Error>(())
        })
        .unwrap_err();
    assert!(matches!(error, Error::PartNotFound(_)));
    assert!(std::sync::Arc::ptr_eq(
        &original,
        &package
            .opc_package()
            .get_part(&document_uri)
            .unwrap()
            .blob_arc()
    ));
    assert!(package.document_mut().is_ok());
}

#[test]
fn raw_opc_transaction_rejects_pending_managed_state() {
    let mut package = Package::new().unwrap();
    package
        .document_mut()
        .unwrap()
        .add_paragraph_with_text("managed edit");

    assert!(matches!(
        package.edit_opc(|_| Ok::<_, Error>(())),
        Err(Error::UnsafeEdit {
            operation: "edit_opc",
            ..
        })
    ));
}
