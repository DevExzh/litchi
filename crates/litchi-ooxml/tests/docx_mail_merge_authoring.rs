use litchi_ooxml::docx::{
    MailMergeConformance, MailMergeDataSourceObject, MailMergeDataType, MailMergeDestination,
    MailMergeFieldMap, MailMergeFieldMappingType, MailMergeRecipient, MailMergeRecipients,
    MailMergeSettings, MailMergeSource, MailMergeTarget, Package,
};
use litchi_opc::constants::content_type as ct;
use litchi_opc::packuri::PackURI;
use litchi_opc::part::{BlobPart, Part};

fn settings_model() -> MailMergeSettings {
    let mut field = MailMergeFieldMap::new();
    field
        .set_mapping_type(Some(MailMergeFieldMappingType::DatabaseColumn))
        .set_name(Some("Name".into()))
        .set_mapped_name(Some("Full Name".into()))
        .set_column(Some(0))
        .set_dynamic_address(true);
    let mut odso = MailMergeDataSourceObject::new();
    odso.set_table(Some("Sheet1$".into()))
        .set_source_type(Some("spreadsheet".into()))
        .set_first_row_header(true)
        .add_field_map(field);
    let mut settings = MailMergeSettings::new();
    settings
        .set_data_type(Some(MailMergeDataType::Spreadsheet))
        .set_connect_string(Some("Provider=Inert;Data Source=never-opened".into()))
        .set_query(Some("SELECT * FROM [Sheet1$]".into()))
        .set_destination(MailMergeDestination::Email)
        .set_mail_subject(Some("Subject & inert".into()))
        .set_view_merged_data(true)
        .set_active_record(2)
        .set_odso(Some(odso));
    settings
}

fn recipients() -> MailMergeRecipients {
    let mut recipients = MailMergeRecipients::new();
    recipients
        .add_recipient(MailMergeRecipient::new(true, Some(0), Some(vec![1, 2, 3])))
        .unwrap();
    recipients
        .add_recipient(MailMergeRecipient::new(false, Some(1), Some(vec![4, 5])))
        .unwrap();
    recipients
}

#[test]
fn generated_graph_round_trips_without_fetching_or_interpreting_sources() {
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("mail-merge.docx");
    let mut package = Package::new().unwrap();
    package
        .set_mail_merge(
            settings_model(),
            Some(MailMergeSource::External(
                "https://example.invalid/never-fetch.csv".into(),
            )),
            Some(MailMergeSource::Internal {
                bytes: b"opaque header bytes".to_vec(),
                content_type: "text/csv".into(),
                extension: "csv".into(),
            }),
            Some(recipients()),
            MailMergeConformance::Transitional,
        )
        .unwrap();
    package.save(&path).unwrap();

    let mut reopened = Package::open(&path).unwrap();
    let merge = reopened.mail_merge_settings().unwrap().unwrap();
    assert_eq!(merge.query(), Some("SELECT * FROM [Sheet1$]"));
    assert_eq!(merge.odso().unwrap().field_maps().len(), 1);
    match reopened
        .mail_merge_target(merge.data_source_relationship_id().unwrap())
        .unwrap()
    {
        MailMergeTarget::External(uri) => {
            assert_eq!(uri, "https://example.invalid/never-fetch.csv")
        },
        _ => panic!("expected inert external URI"),
    }
    let loaded_recipients = reopened
        .document()
        .unwrap()
        .mail_merge_recipients()
        .unwrap()
        .unwrap();
    assert!(loaded_recipients.recipients()[0].active());
    assert!(!loaded_recipients.recipients()[1].active());

    let mut updated = loaded_recipients.clone();
    updated.set_recipient_active(1, true).unwrap();
    reopened
        .update_mail_merge_recipients(updated, MailMergeConformance::Transitional)
        .unwrap();
    assert!(
        reopened
            .document()
            .unwrap()
            .mail_merge_recipients()
            .unwrap()
            .unwrap()
            .recipients()[1]
            .active()
    );
}

#[test]
fn strict_relationships_and_settings_order_are_emitted() {
    let mut package = Package::new().unwrap();
    package
        .set_mail_merge(
            settings_model(),
            Some(MailMergeSource::External("urn:inert:data".into())),
            None,
            Some(recipients()),
            MailMergeConformance::Strict,
        )
        .unwrap();
    let merge = package.mail_merge_settings().unwrap().unwrap();
    let document = package.opc_package().main_document_part().unwrap();
    let settings_rel = document
        .rels()
        .iter()
        .find(|relationship| relationship.reltype().ends_with("/settings"))
        .unwrap();
    let settings_part = package
        .opc_package()
        .get_part(&settings_rel.target_partname().unwrap())
        .unwrap();
    let source_rel = settings_part
        .rels()
        .get(merge.data_source_relationship_id().unwrap())
        .unwrap();
    assert!(
        source_rel
            .reltype()
            .starts_with("http://purl.oclc.org/ooxml/")
    );
    let xml = std::str::from_utf8(settings_part.blob()).unwrap();
    assert!(xml.contains("http://purl.oclc.org/ooxml/wordprocessingml/main"));
    assert!(xml.find("<w:mailMerge").unwrap() < xml.find("<w:defaultTabStop").unwrap());
}

#[test]
fn unrelated_settings_xml_is_preserved_and_invalid_updates_are_atomic() {
    let mut package = Package::new().unwrap();
    let settings_uri = PackURI::new("/word/settings.xml").unwrap();
    package
        .edit_opc(|opc| {
            let settings_part = opc.get_part_mut(&settings_uri)?;
            let original = std::str::from_utf8(settings_part.blob()).map_err(|error| {
                litchi_ooxml::error::OoxmlError::InvalidFormat(error.to_string())
            })?;
            let marked = original.replace(
                "</w:settings>",
                r#"<x:sentinel xmlns:x="urn:test" keep="exact"/></w:settings>"#,
            );
            settings_part.set_blob(marked.into_bytes());
            Ok(())
        })
        .unwrap();
    package
        .set_mail_merge(
            settings_model(),
            None,
            None,
            None,
            MailMergeConformance::Transitional,
        )
        .unwrap();
    let before_count = package.opc_package().part_count();
    let before_xml = package
        .opc_package()
        .get_part(&settings_uri)
        .unwrap()
        .blob()
        .to_vec();
    assert!(
        std::str::from_utf8(&before_xml)
            .unwrap()
            .contains(r#"keep="exact""#)
    );

    let mut invalid = settings_model();
    invalid.set_active_record(0);
    assert!(
        package
            .update_mail_merge(
                invalid,
                Some(MailMergeSource::External("https://example.invalid".into())),
                None,
                None,
                MailMergeConformance::Transitional,
            )
            .is_err()
    );
    assert_eq!(package.opc_package().part_count(), before_count);
    assert_eq!(
        package
            .opc_package()
            .get_part(&settings_uri)
            .unwrap()
            .blob(),
        before_xml
    );
}

#[test]
fn clear_preserves_an_internal_source_still_shared_elsewhere() {
    let mut package = Package::new().unwrap();
    package
        .set_mail_merge(
            settings_model(),
            Some(MailMergeSource::Internal {
                bytes: b"shared opaque source".to_vec(),
                content_type: "application/octet-stream".into(),
                extension: "bin".into(),
            }),
            None,
            None,
            MailMergeConformance::Transitional,
        )
        .unwrap();
    let merge = package.mail_merge_settings().unwrap().unwrap();
    let target = match package
        .mail_merge_target(merge.data_source_relationship_id().unwrap())
        .unwrap()
    {
        MailMergeTarget::Internal { part_name, .. } => part_name,
        _ => panic!("expected internal source"),
    };
    let footer_uri = PackURI::new("/word/footerMailMerge.xml").unwrap();
    let mut footer = BlobPart::new(
        footer_uri.clone(),
        ct::WML_FOOTER.to_string(),
        br#"<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"#
            .to_vec(),
    );
    footer.rels_mut().add_relationship(
        "urn:test:shared".into(),
        target.relative_ref(footer_uri.base_uri()),
        "rIdShared".into(),
        false,
    );
    package
        .edit_opc(|opc| {
            opc.add_part(Box::new(footer));
            Ok(())
        })
        .unwrap();
    assert!(package.clear_mail_merge().unwrap());
    assert!(package.mail_merge_settings().unwrap().is_none());
    assert!(package.opc_package().get_part(&target).is_ok());
    assert!(!package.clear_mail_merge().unwrap());
}

#[test]
fn malformed_uri_and_recipient_limits_fail_before_package_mutation() {
    let mut package = Package::new().unwrap();
    let part_count = package.opc_package().part_count();
    assert!(
        package
            .set_mail_merge(
                settings_model(),
                Some(MailMergeSource::External("bad\nuri".into())),
                None,
                None,
                MailMergeConformance::Transitional,
            )
            .is_err()
    );
    assert_eq!(package.opc_package().part_count(), part_count);
    assert!(package.mail_merge_settings().unwrap().is_none());
}
