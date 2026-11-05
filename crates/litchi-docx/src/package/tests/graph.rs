use super::*;

#[derive(Default)]
struct ForwardOnlySink(Vec<u8>);

impl Write for ForwardOnlySink {
    fn write(&mut self, buffer: &[u8]) -> std::io::Result<usize> {
        self.0.extend_from_slice(buffer);
        Ok(buffer.len())
    }

    fn flush(&mut self) -> std::io::Result<()> {
        Ok(())
    }
}

#[test]
fn plain_stream_accepts_a_forward_only_sink() {
    let mut builder = crate::Builder::new().unwrap();
    builder
        .document()
        .unwrap()
        .add_paragraph_with_text("forward-only DOCX output");

    let mut sink = ForwardOnlySink::default();
    let package = builder.build().unwrap();
    package.to_plain_stream(&mut sink).unwrap();

    let reopened = Package::from_reader(Cursor::new(sink.0)).unwrap();
    assert_eq!(
        reopened.document().unwrap().text().unwrap(),
        "forward-only DOCX output"
    );
}

fn with_eocd_comment(mut archive: Vec<u8>, comment: &[u8]) -> Vec<u8> {
    let comment_len = u16::try_from(comment.len()).expect("ZIP comment fits in EOCD");
    let eocd = archive.len().checked_sub(22).expect("archive has an EOCD");
    assert_eq!(&archive[eocd..eocd + 4], b"PK\x05\x06");
    archive[eocd + 20..eocd + 22].copy_from_slice(&comment_len.to_le_bytes());
    archive.extend_from_slice(comment);
    archive
}

#[test]
fn owned_document_noop_output_preserves_exact_archive() {
    let authored = Package::new().expect("new document");
    let mut initial = Cursor::new(Vec::new());
    authored
        .to_plain_stream(&mut initial)
        .expect("serialize document");
    let source = with_eocd_comment(initial.into_inner(), b"docx exact source");
    let package = Package::from_reader(Cursor::new(source.clone())).expect("open document");
    let mut output = Cursor::new(Vec::new());

    package
        .to_plain_stream(&mut output)
        .expect("serialize no-op document");

    assert_eq!(output.into_inner(), source);
}

#[test]
fn saves_and_reopens_package() {
    let file = NamedTempFile::with_suffix(".docx").unwrap();
    let mut builder = crate::Builder::new().unwrap();
    builder
        .document()
        .unwrap()
        .add_paragraph_with_text("round-trip text");
    let package = builder.build().unwrap();
    package.save(file.path()).unwrap();

    let reopened = Package::open(file.path()).unwrap();
    assert!(
        reopened
            .document()
            .unwrap()
            .text()
            .unwrap()
            .contains("round-trip text")
    );
}

#[test]
fn numbering_patch_publishes_through_the_package_graph() {
    let source_xml = br#"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="1"/></w:numbering>"#;
    let numbering_uri = PackURI::new("/word/numbering.xml").unwrap();
    let mut package = Package::new().unwrap();
    package
        .edit_opc_in_place(|opc| {
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
        .edit_opc_in_place(|opc| {
            opc.get_part_mut(&document_uri)?
                .set_blob(source_xml.to_vec());
            Ok(())
        })
        .unwrap();

    let mut edit = package.edit().unwrap();
    edit.replace_paragraph_text(litchi_core::Position::new(0), "after and longer")
        .unwrap();
    let commit = edit.commit().unwrap();
    package = commit.package().clone();
    let published = commit.patch().target();
    assert_eq!(
        published
            .paragraph(litchi_core::Position::new(0))
            .unwrap()
            .text()
            .unwrap(),
        "after and longer"
    );

    let before_stale = package.opc.main_document_part().unwrap().blob().to_vec();
    assert!(package.apply(commit.patch()).is_err());
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

    package = package
        .apply(&commit.patch().inverse())
        .unwrap()
        .into_package();
    assert_eq!(package.opc.main_document_part().unwrap().blob(), source_xml);
}

#[test]
fn durable_document_patch_preserves_hyperlink_graph_and_reopens_exactly() {
    let source_xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>multi</w:t></w:r><w:r><w:t> run</w:t></w:r></w:p><w:p><w:hyperlink r:id="rIdHyper" w:tooltip="kept"><w:r><w:rPr><w:u/></w:rPr><w:t>linked</w:t></w:r></w:hyperlink></w:p><w:tbl><w:tr><w:tc><w:tcPr><w:shd w:fill="FFFF00"/></w:tcPr><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>"#;
    let document_uri = PackURI::new("/word/document.xml").unwrap();
    let mut package = Package::new().unwrap();
    package
        .edit_opc_in_place(|opc| {
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

    package = package
        .apply_durable_document(&decoded)
        .unwrap()
        .into_package();
    let published = package.document_snapshot().unwrap();
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

    reopened = reopened
        .apply_durable_document(&decoded.inverse())
        .unwrap()
        .into_package();
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
    let fixture = include_bytes!("../../../../../test-data/ooxml/docx/Hyperlink.docx");
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
    package = package
        .apply_durable_document(&durable)
        .unwrap()
        .into_package();
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

    package = package
        .apply_durable_document(&durable.inverse())
        .unwrap()
        .into_package();
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
    let receiver_xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t>receiver</w:t></w:r><w:hyperlink r:id="receiverLink" w:tooltip="receiver link"><w:r><w:t>receiver linked</w:t></w:r></w:hyperlink></w:p><w:sectPr/></w:body></w:document>"#;
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
    let composition_limits = litchi_core::patch::CompositionLimits::new(8, 16, 48, 8);
    let build_composition = || {
        let mut text = source.edit();
        text.replace_run_text(
            litchi_core::Position::new(0),
            litchi_core::Position::new(0),
            "receiver edited",
        )
        .unwrap();
        let mut left = receiver.compose(composition_limits).unwrap();
        left.join(text.prepare(composition_limits, "receiver-text").unwrap())
            .unwrap();
        let mut transfer = source.edit();
        transfer
            .insert_paragraph_transfer(litchi_core::Position::new(1), &plan)
            .unwrap()
            .replace_body_hyperlink_texts(&[
                crate::document::HyperlinkTextReplacement::new(
                    crate::document::ParagraphHyperlinkAddress::new(
                        litchi_core::Position::new(0),
                        litchi_core::Position::new(0),
                    ),
                    "receiver link edited",
                ),
                crate::document::HyperlinkTextReplacement::new(
                    crate::document::ParagraphHyperlinkAddress::new(
                        litchi_core::Position::new(1),
                        litchi_core::Position::new(0),
                    ),
                    "transferred outer link edited",
                ),
            ])
            .unwrap()
            .replace_nested_content_control_hyperlink_text(
                litchi_core::Position::new(1),
                &[litchi_core::Position::new(0), litchi_core::Position::new(0)],
                litchi_core::Position::new(0),
                "transferred control edited",
            )
            .unwrap();
        let mut right = receiver.compose(composition_limits).unwrap();
        right
            .join(
                transfer
                    .prepare(composition_limits, "paragraph-transfer")
                    .unwrap(),
            )
            .unwrap();
        source
            .plan_three_way(left, right)
            .unwrap()
            .finish()
            .unwrap()
    };
    let wrong_artifact = transfer_package(
        receiver_xml,
        "receiverLink",
        "receiverImage",
        b"different image",
    );
    assert_eq!(
        wrong_artifact.document_snapshot().unwrap().xml_bytes(),
        source.xml_bytes()
    );
    assert!(matches!(
        wrong_artifact.apply_composition(build_composition()),
        Err(crate::document::Error::PatchConflict)
    ));
    let commit = receiver.apply_composition(build_composition()).unwrap();
    assert!(
        std::str::from_utf8(commit.patch().target().xml_bytes())
            .unwrap()
            .contains("transferred control edited")
    );
    assert!(
        std::str::from_utf8(commit.patch().target().xml_bytes())
            .unwrap()
            .contains("receiver link edited")
    );
    assert!(
        std::str::from_utf8(commit.patch().target().xml_bytes())
            .unwrap()
            .contains("transferred outer link edited")
    );
    assert!(
        std::str::from_utf8(commit.patch().target().xml_bytes())
            .unwrap()
            .contains("w:tooltip=\"receiver link\"")
    );

    let reopened_receiver =
        Package::from_reader(Cursor::new(receiver.to_bytes().unwrap())).unwrap();
    let reopened_commit = reopened_receiver
        .apply_composition(build_composition())
        .unwrap();
    assert_eq!(
        reopened_commit.patch().target().xml_bytes(),
        commit.patch().target().xml_bytes()
    );
    assert!(
        std::str::from_utf8(commit.patch().target().xml_bytes())
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
    let history_budget = u64::try_from(commit.patch().target().xml_bytes().len()).unwrap() * 2;
    let mut history = source.history(litchi_core::patch::HistoryLimits::new(4, history_budget));
    receiver = receiver
        .publish_document_commit_with_history(commit.clone(), &mut history)
        .unwrap()
        .into_package();
    assert_eq!(
        document_relationship_inventory(&receiver),
        relationships_before
    );
    receiver = {
        let commit = receiver.undo_document(&mut history).unwrap();
        assert!(commit.diagnostics().changed());
        commit.into_package()
    };
    assert_eq!(
        receiver.document_snapshot().unwrap().xml_bytes(),
        source.xml_bytes()
    );
    receiver = {
        let commit = receiver.redo_document(&mut history).unwrap();
        assert!(commit.diagnostics().changed());
        commit.into_package()
    };

    let mut output = Cursor::new(Vec::new());
    receiver.to_plain_stream(&mut output).unwrap();
    let reopened = Package::from_reader(Cursor::new(output.into_inner())).unwrap();
    assert_eq!(
        reopened.document_snapshot().unwrap().xml_bytes(),
        commit.patch().target().xml_bytes()
    );
    assert_eq!(
        document_relationship_inventory(&reopened),
        relationships_before
    );

    let mut durable_receiver =
        transfer_package(receiver_xml, "receiverLink", "receiverImage", b"same image");
    durable_receiver = durable_receiver
        .apply_durable_document(&decoded)
        .unwrap()
        .into_package();
    let mut durable_output = Cursor::new(Vec::new());
    durable_receiver
        .to_plain_stream(&mut durable_output)
        .unwrap();
    let durable_reopened = Package::from_reader(Cursor::new(durable_output.into_inner())).unwrap();
    assert_eq!(
        durable_reopened.document_snapshot().unwrap().xml_bytes(),
        commit.patch().target().xml_bytes()
    );

    let mut stale_receiver =
        transfer_package(receiver_xml, "receiverLink", "receiverImage", b"same image");
    let stale_plan = stale_receiver
        .plan_paragraph_transfer_from(&donor, litchi_core::Position::new(0))
        .unwrap();
    let mut stale_edit = stale_receiver.edit().unwrap();
    stale_edit
        .insert_paragraph_transfer(litchi_core::Position::new(1), &stale_plan)
        .unwrap();
    let stale_commit = stale_edit.commit().unwrap();
    let stale_durable = stale_commit.patch().to_durable(patch_limits).unwrap();
    stale_receiver
        .edit_opc_in_place(|opc| {
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
        stale_receiver.apply(stale_commit.patch()),
        Err(crate::document::Error::PatchConflict)
    ));
    assert_eq!(
        stale_receiver.document_snapshot().unwrap().xml_bytes(),
        stale_xml
    );
    assert!(matches!(
        stale_receiver.apply_durable_document(&stale_durable),
        Err(crate::document::Error::PatchConflict)
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
    let mut copied_edit = mismatched.edit().unwrap();
    copied_edit
        .insert_paragraph_transfer(litchi_core::Position::new(1), &copied_plan)
        .unwrap();
    let copied_commit = copied_edit.commit().unwrap();
    let copied_durable = copied_commit.patch().to_durable(patch_limits).unwrap();
    let mut copied_history =
        mismatched_source.history(litchi_core::patch::HistoryLimits::new(4, 4 * 1024 * 1024));
    mismatched = mismatched
        .publish_document_commit_with_history(copied_commit, &mut copied_history)
        .unwrap()
        .into_package();
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
    mismatched = {
        let commit = mismatched.undo_document(&mut copied_history).unwrap();
        assert!(commit.diagnostics().changed());
        commit.into_package()
    };
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
    mismatched = {
        let commit = mismatched.redo_document(&mut copied_history).unwrap();
        assert!(commit.diagnostics().changed());
        commit.into_package()
    };
    assert_eq!(
        mismatched.opc.get_part(&copied_image).unwrap().blob(),
        b"same image"
    );
    mismatched = {
        let commit = mismatched.undo_document(&mut copied_history).unwrap();
        assert!(commit.diagnostics().changed());
        commit.into_package()
    };
    mismatched = mismatched
        .apply_durable_document(&copied_durable)
        .unwrap()
        .into_package();
    assert_eq!(
        mismatched.opc.get_part(&copied_image).unwrap().blob(),
        b"same image"
    );
    assert!(mismatched.opc.get_part(&copied_metadata).is_ok());
    mismatched = mismatched
        .apply_durable_document(&copied_durable.inverse())
        .unwrap()
        .into_package();
    assert!(mismatched.opc.get_part(&copied_image).is_err());
    assert!(mismatched.opc.get_part(&copied_metadata).is_err());
    assert_eq!(
        document_relationship_inventory(&mismatched),
        mismatched_before
    );

    let mut ambiguous =
        transfer_package(receiver_xml, "receiverLink", "receiverImage", b"same image");
    ambiguous
        .edit_opc_in_place(|opc| {
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
        Err(crate::document::Error::Transfer(
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
        .edit_opc_in_place(|opc| {
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
fn unchanged_stream_preserves_exact_bytes_and_part_payload_sharing() {
    let file = NamedTempFile::with_suffix(".docx").unwrap();
    let mut builder = crate::Builder::new().unwrap();
    builder
        .document()
        .unwrap()
        .add_paragraph_with_text("unchanged package");
    let source = builder.build().unwrap();
    source.save(file.path()).unwrap();

    let package = Package::open(file.path()).unwrap();
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
fn raw_opc_edit_publishes_candidate_and_keeps_the_source() {
    let package = Package::new().unwrap();
    let marker = PackURI::new("/custom/raw-edit-marker.bin").unwrap();

    let mut edit = package.edit_opc();
    edit.package_mut()
        .try_add_part(Box::new(BlobPart::new(
            marker.clone(),
            "application/octet-stream".to_string(),
            b"raw edit".to_vec(),
        )))
        .unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.diagnostics().changed());
    assert_eq!(
        commit
            .package()
            .opc_package()
            .get_part(&marker)
            .unwrap()
            .blob(),
        b"raw edit"
    );
    assert!(!package.opc_package().contains_part(&marker));
    let undone = commit
        .package()
        .apply_opc(&commit.patch().inverse())
        .ok()
        .unwrap();
    assert!(!undone.package().opc_package().contains_part(&marker));
}

#[test]
fn failed_raw_opc_edit_preserves_the_source_graph() {
    let package = Package::new().unwrap();
    let document_uri = PackURI::new("/word/document.xml").unwrap();
    let original = package
        .opc_package()
        .get_part(&document_uri)
        .unwrap()
        .blob_arc();

    let mut edit = package.edit_opc();
    edit.package_mut().remove_part(&document_uri);
    let error = edit.commit().err().unwrap();
    assert!(matches!(error, Error::PartNotFound(_)));
    assert!(std::sync::Arc::ptr_eq(
        &original,
        &package
            .opc_package()
            .get_part(&document_uri)
            .unwrap()
            .blob_arc()
    ));
}

#[test]
fn exact_source_patch_reuse_keeps_noop_arcs_stale_conflicts_and_scan_errors() {
    let source_xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>before</w:t></w:r></w:p><w:sectPr/></w:body></w:document>"#;
    let document_uri = PackURI::new("/word/document.xml").unwrap();
    let mut package = Package::new().unwrap();
    package
        .edit_opc_in_place(|opc| {
            opc.get_part_mut(&document_uri)?
                .set_blob(source_xml.to_vec());
            Ok(())
        })
        .unwrap();

    // An exact no-op publishes the retained source without replacing the
    // main-part payload allocation.
    let main_before = package.opc.main_document_part().unwrap().blob_arc();
    let noop = package.edit().unwrap().commit().unwrap();
    assert!(noop.patch().is_noop());
    let next = package.apply(noop.patch()).unwrap().into_package();
    let published = next.document_snapshot().unwrap();
    let main_after_noop = package.opc.main_document_part().unwrap().blob_arc();
    assert!(std::sync::Arc::ptr_eq(&main_before, &main_after_noop));
    assert_eq!(published.xml_bytes(), source_xml);
    assert_eq!(
        published
            .paragraph(litchi_core::Position::new(0))
            .unwrap()
            .text()
            .unwrap(),
        "before"
    );

    // A patch whose source no longer matches the package still conflicts, and
    // leaves the package untouched.
    let mut edit = package.edit().unwrap();
    edit.replace_paragraph_text(litchi_core::Position::new(0), "after and longer")
        .unwrap();
    let commit = edit.commit().unwrap();
    package = commit.package().clone();
    let republished = package.opc.main_document_part().unwrap().blob().to_vec();
    assert!(matches!(
        package.apply(commit.patch()),
        Err(crate::document::Error::PatchConflict)
    ));
    assert_eq!(
        package.opc.main_document_part().unwrap().blob(),
        republished
    );

    // Reapplying the inverse still reaches the exact original bytes through the
    // reuse proof.
    package = package
        .apply(&commit.patch().inverse())
        .unwrap()
        .into_package();
    assert_eq!(package.opc.main_document_part().unwrap().blob(), source_xml);

    // When the main part no longer parses, the byte proof declines and the full
    // snapshot route decides, so the scan error keeps its identity instead of
    // becoming a stale-source refusal.
    package
        .edit_opc_in_place(|opc| {
            opc.get_part_mut(&document_uri)?
                .set_blob(b"<w:document>".to_vec());
            Ok(())
        })
        .unwrap();
    assert!(matches!(
        package.apply(commit.patch()),
        Err(crate::document::Error::Document(_))
    ));
}

#[test]
fn the_document_snapshot_shares_the_main_part_payload_instead_of_copying_it() {
    let source_xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>before</w:t></w:r></w:p><w:sectPr/></w:body></w:document>"#;
    let document_uri = PackURI::new("/word/document.xml").unwrap();
    let mut package = Package::new().unwrap();
    package
        .edit_opc_in_place(|opc| {
            opc.get_part_mut(&document_uri)?
                .set_blob(source_xml.to_vec());
            Ok(())
        })
        .unwrap();

    let retained = package.opc.main_document_part().unwrap().blob_arc();
    let snapshot = package.document_snapshot().unwrap();
    assert!(
        std::ptr::eq(snapshot.xml_bytes().as_ptr(), retained.as_ptr()),
        "the snapshot must borrow the part's allocation rather than copy it"
    );

    // Publishing a different main document replaces the part's `Arc` instead of
    // mutating through it, so the snapshot taken before still reads its own
    // bytes afterwards.
    let mut edit = package.edit().unwrap();
    edit.replace_paragraph_text(litchi_core::Position::new(0), "after and longer")
        .unwrap();
    package = edit.commit().unwrap().into_package();
    assert_eq!(snapshot.xml_bytes(), source_xml);
    assert!(!std::ptr::eq(
        package.opc.main_document_part().unwrap().blob().as_ptr(),
        retained.as_ptr()
    ));
}

#[test]
fn the_compaction_policy_reaches_publication_through_the_package_route() {
    let source_xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:pStyle w:val="it's"/></w:pPr><w:r><w:t>one</w:t></w:r></w:p><w:p><w:r><w:t>two</w:t></w:r></w:p><w:sectPr/></w:body></w:document>"#;
    let document_uri = PackURI::new("/word/document.xml").unwrap();
    let build = || {
        let mut package = Package::new().unwrap();
        package
            .edit_opc_in_place(|opc| {
                opc.get_part_mut(&document_uri)?
                    .set_blob(source_xml.to_vec());
                Ok(())
            })
            .unwrap();
        package
    };

    let mut preserving = build();
    let mut edit = preserving.edit().unwrap();
    edit.replace_paragraph_text(litchi_core::Position::new(1), "two edited")
        .unwrap();
    preserving = edit.commit().unwrap().into_package();
    let preserved = preserving.opc.main_document_part().unwrap().blob().to_vec();

    let mut compacting = build();
    let mut edit = compacting
        .edit()
        .unwrap()
        .with_compaction_policy(crate::document::CompactionPolicy::WholeDocument);
    edit.replace_paragraph_text(litchi_core::Position::new(1), "two edited")
        .unwrap();
    compacting = edit.commit().unwrap().into_package();
    let compacted = compacting.opc.main_document_part().unwrap().blob().to_vec();

    assert!(
        preserved.windows(6).any(|window| window == b"it's\"/"),
        "the default policy republishes the untouched paragraph's bytes"
    );
    assert!(
        compacted.windows(10).any(|window| window == b"it&apos;s\""),
        "the opt-in policy re-serializes the untouched paragraph"
    );
    assert_ne!(preserved, compacted);

    // Both publish: the preserved bytes satisfy the writer's authored-XML
    // contract, so neither route is refused at save.
    let mut published = Vec::new();
    preserving.write_to(&mut published).unwrap();
    assert!(!published.is_empty());
    let mut published = Vec::new();
    compacting.write_to(&mut published).unwrap();
    assert!(!published.is_empty());
}

#[test]
fn immutable_apply_shares_noops_and_keeps_each_package_revision_readable() {
    fn assert_shared<T: Clone + Send + Sync + 'static>() {}
    assert_shared::<Package>();
    let package = Package::new().unwrap();
    let source_xml = package.document_snapshot().unwrap();
    let noop = package.edit().unwrap().commit().unwrap();
    let unchanged = package.apply(noop.patch()).unwrap().into_package();
    assert!(std::ptr::eq(&*package.opc, &*unchanged.opc));
    assert!(std::ptr::eq(&*package.properties, &*unchanged.properties));
    assert!(std::ptr::eq(
        &*package.custom_props,
        &*unchanged.custom_props
    ));

    unchanged.to_plain_stream(Vec::new()).unwrap();
    assert!(std::ptr::eq(&*package.opc, &*unchanged.opc));

    let mut edit = package.edit().unwrap();
    edit.insert_paragraph(litchi_core::Position::new(0), "new paragraph")
        .unwrap();
    let commit = edit.commit().unwrap();
    let changed = package.apply(commit.patch()).unwrap().into_package();
    assert_eq!(package.document().unwrap().text().unwrap(), "");
    assert_eq!(changed.document().unwrap().text().unwrap(), "new paragraph");
    assert!(!std::ptr::eq(&*package.opc, &*changed.opc));
    changed.document_snapshot().unwrap();
    // The shared admission memo must never confuse a changed allocation
    // with the still-live original package.
    assert_eq!(
        package.document_snapshot().unwrap().xml_bytes(),
        source_xml.xml_bytes()
    );
    assert!(matches!(
        changed.apply(commit.patch()),
        Err(crate::document::Error::PatchConflict)
    ));
    let restored = changed
        .apply(&commit.patch().inverse())
        .unwrap()
        .into_package();
    assert_eq!(
        restored.document_snapshot().unwrap().xml_bytes(),
        source_xml.xml_bytes()
    );
    assert_eq!(changed.document().unwrap().text().unwrap(), "new paragraph");
}
