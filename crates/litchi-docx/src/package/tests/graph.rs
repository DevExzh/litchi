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

    let mut edit = package.edit_document().unwrap();
    edit.replace_paragraph_text(litchi_core::Position::new(0), "after and longer")
        .unwrap();
    let commit = package.publish_document_edit(edit).unwrap();
    let published = commit.snapshot();
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
fn durable_document_patch_preserves_hyperlink_graph_and_reopens_exactly() {
    let source_xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>multi</w:t></w:r><w:r><w:t> run</w:t></w:r></w:p><w:p><w:hyperlink r:id="rIdHyper" w:tooltip="kept"><w:r><w:rPr><w:u/></w:rPr><w:t>linked</w:t></w:r></w:hyperlink></w:p><w:tbl><w:tr><w:tc><w:tcPr><w:shd w:fill="FFFF00"/></w:tcPr><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>"#;
    let document_uri = PackURI::new("/word/document.xml").unwrap();
    let mut package = Package::new().unwrap();
    package
        .edit_opc(|opc| {
            let main = opc.get_part_mut(&document_uri)?;
            main.set_blob(source_xml.to_vec());
            main.rels_mut().try_add_relationship(
                litchi_opc::constants::relationship_type::HYPERLINK.to_owned(),
                "https://example.invalid/preserved".to_owned(),
                "rIdHyper".to_owned(),
                TargetMode::External,
            )?;
            Ok(())
        })
        .unwrap();
    let relationships_before = document_relationship_inventory(&package);

    let source = package.document_snapshot().unwrap();
    let mut edit = source.edit();
    edit.replace_paragraph_text(litchi_core::Position::new(0), "changed multi run")
        .unwrap()
        .replace_hyperlink_text(
            litchi_core::Position::new(1),
            litchi_core::Position::new(0),
            "changed link",
        )
        .unwrap()
        .replace_table_cell_text(
            litchi_core::Position::new(0),
            litchi_core::Position::new(0),
            litchi_core::Position::new(0),
            "changed cell",
        )
        .unwrap();
    let commit = edit.commit().unwrap();
    let limits = litchi_core::patch::PatchLimits::new(
        litchi_core::patch::BlobLimits::new(1, 1024 * 1024, 1024 * 1024),
        2 * 1024 * 1024,
        16,
        8,
        64 * 1024,
        1024 * 1024,
    );
    let durable = commit.patch().to_durable(limits).unwrap();
    let wire = durable.to_deterministic_json().unwrap();
    let decoded =
        litchi_core::patch::Patch::<litchi_core::patch::Reversible>::from_deterministic_json(
            &wire, limits,
        )
        .unwrap();

    let published = package.apply_durable_document_patch(&decoded).unwrap();
    assert_eq!(published.xml_bytes(), commit.snapshot().xml_bytes());
    assert_eq!(
        document_relationship_inventory(&package),
        relationships_before
    );

    let mut output = Cursor::new(Vec::new());
    package.to_plain_stream(&mut output).unwrap();
    let mut reopened = Package::from_reader(Cursor::new(output.into_inner())).unwrap();
    assert_eq!(
        reopened.document_snapshot().unwrap().xml_bytes(),
        commit.snapshot().xml_bytes()
    );
    assert_eq!(
        document_relationship_inventory(&reopened),
        relationships_before
    );

    reopened
        .apply_durable_document_patch(&decoded.inverse())
        .unwrap();
    assert_eq!(
        reopened.opc.main_document_part().unwrap().blob(),
        source_xml
    );
    assert_eq!(
        document_relationship_inventory(&reopened),
        relationships_before
    );
}

#[test]
fn durable_hyperlink_edit_round_trips_real_open_xml_sdk_fixture() {
    let fixture = include_bytes!(
        "../../../../../3rdparty/Open-XML-SDK/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestFiles/Hyperlink.docx"
    );
    let mut package = Package::from_reader(Cursor::new(fixture.as_slice())).unwrap();
    let relationships_before = document_relationship_inventory(&package);
    let source = package.document_snapshot().unwrap();
    let mut selected = None;
    'paragraphs: for paragraph in 0..source.paragraph_count() {
        for hyperlink in 0..16 {
            let mut edit = source.edit();
            if edit
                .replace_hyperlink_text(
                    litchi_core::Position::new(paragraph),
                    litchi_core::Position::new(hyperlink),
                    "fixture hyperlink edited",
                )
                .is_ok()
            {
                selected = Some(edit.commit().unwrap());
                break 'paragraphs;
            }
        }
    }
    let commit = selected.expect("fixture must contain one directly editable hyperlink");
    let limits = litchi_core::patch::PatchLimits::new(
        litchi_core::patch::BlobLimits::new(1, 4 * 1024 * 1024, 4 * 1024 * 1024),
        8 * 1024 * 1024,
        32,
        8,
        64 * 1024,
        4 * 1024 * 1024,
    );
    let durable = commit.patch().to_durable(limits).unwrap();
    package.apply_durable_document_patch(&durable).unwrap();
    assert_eq!(
        document_relationship_inventory(&package),
        relationships_before
    );

    let mut changed = Cursor::new(Vec::new());
    package.to_plain_stream(&mut changed).unwrap();
    let reopened = Package::from_reader(Cursor::new(changed.into_inner())).unwrap();
    assert_eq!(
        reopened.document_snapshot().unwrap().xml_bytes(),
        commit.snapshot().xml_bytes()
    );
    assert_eq!(
        document_relationship_inventory(&reopened),
        relationships_before
    );

    package
        .apply_durable_document_patch(&durable.inverse())
        .unwrap();
    assert_eq!(
        package.document_snapshot().unwrap().xml_bytes(),
        source.xml_bytes()
    );
    assert_eq!(
        document_relationship_inventory(&package),
        relationships_before
    );
}

#[test]
fn package_root_three_way_transfer_history_and_durable_reopen_are_coupled() {
    let donor_xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><w:body><w:p><w:hyperlink r:id="donorLink"><w:r><w:t>linked transfer</w:t></w:r></w:hyperlink><w:r><w:drawing><a:blip r:embed="donorImage"/></w:drawing><w:t> image</w:t></w:r><w:sdt><w:sdtPr><w:tag w:val="outer-transfer"/></w:sdtPr><w:sdtContent><w:sdt><w:sdtPr><w:tag w:val="inner-transfer"/></w:sdtPr><w:sdtContent><w:hyperlink r:id="donorLink" w:tooltip="transferred nested link"><w:r><w:t>control transfer</w:t></w:r></w:hyperlink></w:sdtContent></w:sdt></w:sdtContent></w:sdt></w:p><w:sectPr/></w:body></w:document>"#;
    let receiver_xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>receiver</w:t></w:r></w:p><w:sectPr/></w:body></w:document>"#;
    let donor = transfer_package(donor_xml, "donorLink", "donorImage", b"same image");
    let mut receiver =
        transfer_package(receiver_xml, "receiverLink", "receiverImage", b"same image");
    let plan = receiver
        .plan_paragraph_transfer_from(&donor, litchi_core::Position::new(0))
        .unwrap();
    assert!(
        std::str::from_utf8(plan.xml_bytes())
            .unwrap()
            .contains("r:id=\"receiverLink\"")
    );
    assert!(
        std::str::from_utf8(plan.xml_bytes())
            .unwrap()
            .contains("r:embed=\"receiverImage\"")
    );
    assert!(
        !std::str::from_utf8(plan.xml_bytes())
            .unwrap()
            .contains("donorLink")
    );
    assert!(
        std::str::from_utf8(plan.xml_bytes())
            .unwrap()
            .contains("xmlns:a=")
    );

    let source = receiver.document_snapshot().unwrap();
    let composition_limits = litchi_core::patch::CompositionLimits::new(8, 8, 32, 8);
    let mut text = source.edit();
    text.replace_run_text(
        litchi_core::Position::new(0),
        litchi_core::Position::new(0),
        "receiver edited",
    )
    .unwrap();
    let mut left = source.compose(composition_limits);
    left.join(text.prepare(composition_limits, "receiver-text").unwrap())
        .unwrap();
    let mut transfer = source.edit();
    transfer
        .insert_paragraph_transfer(litchi_core::Position::new(1), &plan)
        .unwrap()
        .replace_nested_content_control_hyperlink_text(
            litchi_core::Position::new(1),
            &[litchi_core::Position::new(0), litchi_core::Position::new(0)],
            litchi_core::Position::new(0),
            "transferred control edited",
        )
        .unwrap();
    let mut right = source.compose(composition_limits);
    right
        .join(
            transfer
                .prepare(composition_limits, "paragraph-transfer")
                .unwrap(),
        )
        .unwrap();
    let three_way = source.plan_three_way(left, right).unwrap();
    assert!(three_way.is_clean());
    let commit = three_way.finish().unwrap().commit().unwrap();
    assert!(
        std::str::from_utf8(commit.snapshot().xml_bytes())
            .unwrap()
            .contains("transferred control edited")
    );
    assert!(
        std::str::from_utf8(commit.snapshot().xml_bytes())
            .unwrap()
            .contains("w:tooltip=\"transferred nested link\"")
    );

    let patch_limits = litchi_core::patch::PatchLimits::new(
        litchi_core::patch::BlobLimits::new(2, 4 * 1024 * 1024, 8 * 1024 * 1024),
        16 * 1024 * 1024,
        32,
        8,
        4 * 1024 * 1024,
        8 * 1024 * 1024,
    );
    let durable = commit.patch().to_durable(patch_limits).unwrap();
    let wire = durable.to_deterministic_json().unwrap();
    let decoded =
        litchi_core::patch::Patch::<litchi_core::patch::Reversible>::from_deterministic_json(
            &wire,
            patch_limits,
        )
        .unwrap();
    let relationships_before = document_relationship_inventory(&receiver);
    let history_budget = u64::try_from(commit.snapshot().xml_bytes().len()).unwrap() * 2;
    let mut history = source.history(litchi_core::patch::HistoryLimits::new(4, history_budget));
    receiver
        .publish_document_commit_with_history(commit.clone(), &mut history)
        .unwrap();
    assert_eq!(
        document_relationship_inventory(&receiver),
        relationships_before
    );
    assert!(receiver.undo_document(&mut history).unwrap());
    assert_eq!(
        receiver.document_snapshot().unwrap().xml_bytes(),
        source.xml_bytes()
    );
    assert!(receiver.redo_document(&mut history).unwrap());

    let mut output = Cursor::new(Vec::new());
    receiver.to_plain_stream(&mut output).unwrap();
    let reopened = Package::from_reader(Cursor::new(output.into_inner())).unwrap();
    assert_eq!(
        reopened.document_snapshot().unwrap().xml_bytes(),
        commit.snapshot().xml_bytes()
    );
    assert_eq!(
        document_relationship_inventory(&reopened),
        relationships_before
    );

    let mut durable_receiver =
        transfer_package(receiver_xml, "receiverLink", "receiverImage", b"same image");
    durable_receiver
        .apply_durable_document_patch(&decoded)
        .unwrap();
    let mut durable_output = Cursor::new(Vec::new());
    durable_receiver
        .to_plain_stream(&mut durable_output)
        .unwrap();
    let durable_reopened = Package::from_reader(Cursor::new(durable_output.into_inner())).unwrap();
    assert_eq!(
        durable_reopened.document_snapshot().unwrap().xml_bytes(),
        commit.snapshot().xml_bytes()
    );

    let mut stale_receiver =
        transfer_package(receiver_xml, "receiverLink", "receiverImage", b"same image");
    let stale_plan = stale_receiver
        .plan_paragraph_transfer_from(&donor, litchi_core::Position::new(0))
        .unwrap();
    let mut stale_edit = stale_receiver.edit_document().unwrap();
    stale_edit
        .insert_paragraph_transfer(litchi_core::Position::new(1), &stale_plan)
        .unwrap();
    let stale_commit = stale_edit.commit().unwrap();
    let stale_durable = stale_commit.patch().to_durable(patch_limits).unwrap();
    stale_receiver
        .edit_opc(|opc| {
            opc.get_part_mut(&PackURI::new("/word/document.xml").unwrap())?
                .rels_mut()
                .try_add_relationship(
                    litchi_opc::constants::relationship_type::HYPERLINK.to_owned(),
                    "https://example.invalid/unrelated".to_owned(),
                    "lateRelationship".to_owned(),
                    TargetMode::External,
                )?;
            Ok(())
        })
        .unwrap();
    let stale_xml = stale_receiver
        .document_snapshot()
        .unwrap()
        .xml_bytes()
        .to_vec();
    assert!(matches!(
        stale_receiver.publish_document_commit(stale_commit),
        Err(crate::document::TransactionError::StaleSource)
    ));
    assert_eq!(
        stale_receiver.document_snapshot().unwrap().xml_bytes(),
        stale_xml
    );
    assert!(matches!(
        stale_receiver.apply_durable_document_patch(&stale_durable),
        Err(crate::document::TransactionError::StaleSource)
    ));
    assert_eq!(
        stale_receiver.document_snapshot().unwrap().xml_bytes(),
        stale_xml
    );

    let mut mismatched = transfer_package(
        receiver_xml,
        "receiverLink",
        "receiverImage",
        b"different image",
    );
    let mismatched_before = document_relationship_inventory(&mismatched);
    let mismatched_source = mismatched.document_snapshot().unwrap();
    let copied_plan = mismatched
        .plan_paragraph_transfer_from(&donor, litchi_core::Position::new(0))
        .unwrap();
    let mut copied_edit = mismatched_source.edit();
    copied_edit
        .insert_paragraph_transfer(litchi_core::Position::new(1), &copied_plan)
        .unwrap();
    let copied_commit = copied_edit.commit().unwrap();
    let copied_durable = copied_commit.patch().to_durable(patch_limits).unwrap();
    let mut copied_history =
        mismatched_source.history(litchi_core::patch::HistoryLimits::new(4, 4 * 1024 * 1024));
    mismatched
        .publish_document_commit_with_history(copied_commit, &mut copied_history)
        .unwrap();
    let copied_image = PackURI::new("/word/media/shared-transfer1.png").unwrap();
    let copied_metadata = PackURI::new("/word/media/shared-meta-transfer1.xml").unwrap();
    assert_eq!(
        mismatched.opc.get_part(&copied_image).unwrap().blob(),
        b"same image"
    );
    assert_eq!(
        mismatched.opc.get_part(&copied_metadata).unwrap().blob(),
        b"<metadata source=\"nested\"/>"
    );
    assert_eq!(
        mismatched
            .opc
            .get_part(&copied_image)
            .unwrap()
            .rels()
            .get("nestedMetadata")
            .unwrap()
            .target_partname()
            .unwrap(),
        copied_metadata
    );
    assert_ne!(
        document_relationship_inventory(&mismatched),
        mismatched_before
    );
    assert!(mismatched.undo_document(&mut copied_history).unwrap());
    assert!(mismatched.opc.get_part(&copied_image).is_err());
    assert!(mismatched.opc.get_part(&copied_metadata).is_err());
    assert_eq!(
        mismatched.document_snapshot().unwrap().xml_bytes(),
        receiver_xml
    );
    assert_eq!(
        document_relationship_inventory(&mismatched),
        mismatched_before
    );
    assert!(mismatched.redo_document(&mut copied_history).unwrap());
    assert_eq!(
        mismatched.opc.get_part(&copied_image).unwrap().blob(),
        b"same image"
    );
    assert!(mismatched.undo_document(&mut copied_history).unwrap());
    mismatched
        .apply_durable_document_patch(&copied_durable)
        .unwrap();
    assert_eq!(
        mismatched.opc.get_part(&copied_image).unwrap().blob(),
        b"same image"
    );
    assert!(mismatched.opc.get_part(&copied_metadata).is_ok());
    mismatched
        .apply_durable_document_patch(&copied_durable.inverse())
        .unwrap();
    assert!(mismatched.opc.get_part(&copied_image).is_err());
    assert!(mismatched.opc.get_part(&copied_metadata).is_err());
    assert_eq!(
        document_relationship_inventory(&mismatched),
        mismatched_before
    );

    let mut ambiguous =
        transfer_package(receiver_xml, "receiverLink", "receiverImage", b"same image");
    ambiguous
        .edit_opc(|opc| {
            opc.get_part_mut(&PackURI::new("/word/document.xml").unwrap())?
                .rels_mut()
                .try_add_relationship(
                    litchi_opc::constants::relationship_type::HYPERLINK.to_owned(),
                    "https://example.invalid/transfer".to_owned(),
                    "ambiguousLink".to_owned(),
                    TargetMode::External,
                )?;
            Ok(())
        })
        .unwrap();
    assert!(matches!(
        ambiguous.plan_paragraph_transfer_from(&donor, litchi_core::Position::new(0)),
        Err(crate::document::TransactionError::Transfer(
            crate::document::TransferRefusal::AmbiguousEquivalentDependency { .. }
        ))
    ));
}

fn transfer_package(
    document_xml: &[u8],
    hyperlink_id: &str,
    image_id: &str,
    image: &[u8],
) -> Package {
    let document_uri = PackURI::new("/word/document.xml").unwrap();
    let image_uri = PackURI::new("/word/media/shared.png").unwrap();
    let metadata_uri = PackURI::new("/word/media/shared-meta.xml").unwrap();
    let mut package = Package::new().unwrap();
    package
        .edit_opc(|opc| {
            if opc.get_part(&image_uri).is_err() {
                opc.try_add_part(Box::new(BlobPart::new(
                    image_uri.clone(),
                    "image/png".to_owned(),
                    image.to_vec(),
                )))?;
            }
            if opc.get_part(&metadata_uri).is_err() {
                opc.try_add_part(Box::new(BlobPart::new(
                    metadata_uri.clone(),
                    "application/xml".to_owned(),
                    b"<metadata source=\"nested\"/>".to_vec(),
                )))?;
            }
            opc.get_part_mut(&image_uri)?
                .rels_mut()
                .try_add_relationship(
                    "urn:litchi:test:image-metadata".to_owned(),
                    "shared-meta.xml".to_owned(),
                    "nestedMetadata".to_owned(),
                    TargetMode::Internal,
                )?;
            let main = opc.get_part_mut(&document_uri)?;
            main.set_blob(document_xml.to_vec());
            main.rels_mut().try_add_relationship(
                litchi_opc::constants::relationship_type::HYPERLINK.to_owned(),
                "https://example.invalid/transfer".to_owned(),
                hyperlink_id.to_owned(),
                TargetMode::External,
            )?;
            main.rels_mut().try_add_relationship(
                litchi_opc::constants::relationship_type::IMAGE.to_owned(),
                "media/shared.png".to_owned(),
                image_id.to_owned(),
                TargetMode::Internal,
            )?;
            Ok(())
        })
        .unwrap();
    package
}

fn document_relationship_inventory(package: &Package) -> Vec<(String, String, String, bool)> {
    let mut relationships = package
        .opc
        .main_document_part()
        .unwrap()
        .rels()
        .iter()
        .map(|relationship| {
            (
                relationship.r_id().to_owned(),
                relationship.reltype().to_owned(),
                relationship.target_ref().to_owned(),
                relationship.target_mode() == TargetMode::External,
            )
        })
        .collect::<Vec<_>>();
    relationships.sort();
    relationships
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
