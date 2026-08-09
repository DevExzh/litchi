mod support;

use litchi_core::Metadata;
use litchi_odf_common::calculation::{Iteration, IterationStatus, Settings};
use litchi_odf_common::core::OwnedPackage;
use litchi_ods::{Builder, MutableSpreadsheet, Spreadsheet};
use std::num::NonZeroUsize;

const CONTENT: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content
    xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
    xmlns:vendor="urn:example:vendor"
    office:version="1.3">
  <office:body><office:spreadsheet>
    <vendor:extension vendor:flag="keep"><vendor:value>opaque</vendor:value></vendor:extension>
    <table:calculation-settings table:case-sensitive="false"/>
    <table:table table:name="Data"/>
  </office:spreadsheet></office:body>
</office:document-content>"#;

const META: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-meta
    xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:dc="http://purl.org/dc/elements/1.1/"
    xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"
    xmlns:vendor="urn:example:vendor">
  <office:meta>
    <vendor:opaque vendor:flag="keep"><vendor:value>untouched</vendor:value></vendor:opaque>
    <dc:title>Before</dc:title>
    <dc:creator>Author</dc:creator>
  </office:meta>
</office:document-meta>"#;

const SETTINGS: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-settings
    xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0"
    xmlns:vendor="urn:example:vendor"
    office:version="1.3">
  <office:settings><vendor:opaque vendor:flag="keep"/></office:settings>
</office:document-settings>"#;

#[test]
fn facade_round_trips_typed_metadata_and_calculation_settings() {
    let bytes = package(CONTENT, Some(META), Some(SETTINGS));
    let spreadsheet =
        Spreadsheet::from_bytes(bytes.clone()).expect("test fixture or operation should succeed");
    assert_eq!(spreadsheet.metadata().title.as_deref(), Some("Before"));
    assert_eq!(spreadsheet.metadata().author.as_deref(), Some("Author"));
    assert_eq!(spreadsheet.odf_metadata().title.as_deref(), Some("Before"));
    assert_eq!(
        spreadsheet
            .settings()
            .expect("test fixture or operation should succeed")
            .case_sensitive,
        Some(false)
    );
    assert!(
        spreadsheet
            .metadata_snapshot()
            .xml()
            .expect("test fixture or operation should succeed")
            .contains("vendor:opaque")
    );

    let mut mutable =
        MutableSpreadsheet::from_bytes(bytes).expect("test fixture or operation should succeed");
    mutable
        .update_metadata(|metadata| {
            metadata.title = Some("After".to_string());
            metadata.description = Some("Edited safely".to_string());
            Ok(())
        })
        .expect("test fixture or operation should succeed");
    mutable
        .update_settings(|settings| {
            settings.case_sensitive = Some(true);
            settings.iteration = Some(Iteration {
                status: Some(IterationStatus::Enable),
                steps: NonZeroUsize::new(20),
                maximum_difference: Some("1E-6".to_string()),
            });
            Ok(())
        })
        .expect("test fixture or operation should succeed");

    let output = mutable.to_bytes();
    let reopened =
        Spreadsheet::from_bytes(output.clone()).expect("test fixture or operation should succeed");
    assert_eq!(reopened.metadata().title.as_deref(), Some("After"));
    assert_eq!(
        reopened.metadata().description.as_deref(),
        Some("Edited safely")
    );
    assert_eq!(
        reopened
            .settings()
            .expect("test fixture or operation should succeed")
            .case_sensitive,
        Some(true)
    );
    assert_eq!(
        reopened
            .settings()
            .expect("test fixture or operation should succeed")
            .iteration
            .as_ref()
            .expect("test fixture or operation should succeed")
            .steps,
        NonZeroUsize::new(20)
    );
    assert!(reopened.content_xml().contains("vendor:extension"));
    assert!(
        reopened
            .metadata_snapshot()
            .xml()
            .expect("test fixture or operation should succeed")
            .contains("vendor:opaque")
    );

    let archive =
        OwnedPackage::from_bytes(output).expect("test fixture or operation should succeed");
    assert_eq!(
        archive
            .get_file("settings.xml")
            .expect("test fixture or operation should succeed"),
        support::compact_xml_fixture(SETTINGS).as_bytes()
    );

    let mut clear = MutableSpreadsheet::from_bytes(package(CONTENT, Some(META), None))
        .expect("test fixture or operation should succeed");
    clear
        .clear_metadata()
        .expect("test fixture or operation should succeed");
    let cleared = Spreadsheet::from_bytes(clear.to_bytes())
        .expect("test fixture or operation should succeed");
    assert!(!cleared.metadata_snapshot().is_present());
}

#[test]
fn facade_edits_are_atomic_and_builder_writes_new_parts() {
    let mut mutable = MutableSpreadsheet::from_bytes(package(CONTENT, Some(META), None))
        .expect("test fixture or operation should succeed");
    let before = mutable.metadata().title.clone();
    let result = mutable.update_metadata(|metadata| {
        metadata.identifier = Some("unsupported".to_string());
        Ok(())
    });
    assert!(result.is_err());
    assert_eq!(mutable.metadata().title, before);
    assert!(mutable.settings().is_some());

    let mut builder = Builder::new();
    builder
        .set_metadata(Metadata {
            title: Some("Built".to_string()),
            author: Some("Builder".to_string()),
            ..Metadata::default()
        })
        .expect("test fixture or operation should succeed");
    builder
        .set_settings(Some(Settings {
            precision_as_shown: Some(true),
            ..Settings::default()
        }))
        .expect("test fixture or operation should succeed");
    let spreadsheet = Spreadsheet::from_bytes(
        builder
            .build()
            .expect("test fixture or operation should succeed"),
    )
    .expect("test fixture or operation should succeed");
    assert_eq!(spreadsheet.metadata().title.as_deref(), Some("Built"));
    assert_eq!(spreadsheet.metadata().author.as_deref(), Some("Builder"));
    assert_eq!(
        spreadsheet
            .settings()
            .expect("test fixture or operation should succeed")
            .precision_as_shown,
        Some(true)
    );
}

#[test]
fn malformed_metadata_and_settings_are_rejected_at_package_boundary() {
    let malformed_meta = r#"<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"><office:meta><meta:user-defined>missing name</meta:user-defined></office:meta></office:document-meta>"#;
    assert!(Spreadsheet::from_bytes(package(CONTENT, Some(malformed_meta), None)).is_err());

    let malformed_content = CONTENT.replace(
        "<table:calculation-settings table:case-sensitive=\"false\"/>",
        "<table:calculation-settings/><table:calculation-settings/>",
    );
    assert!(Spreadsheet::from_bytes(package(&malformed_content, Some(META), None)).is_err());

    let invalid_iteration = CONTENT.replace(
        "<table:calculation-settings table:case-sensitive=\"false\"/>",
        "<table:calculation-settings><table:iteration table:steps=\"0\"/></table:calculation-settings>",
    );
    assert!(Spreadsheet::from_bytes(package(&invalid_iteration, Some(META), None)).is_err());

    let unknown_owned_attribute = CONTENT.replace(
        "<table:calculation-settings table:case-sensitive=\"false\"/>",
        "<table:calculation-settings table:case-sensitive=\"false\" table:future=\"x\"/>",
    );
    assert!(Spreadsheet::from_bytes(package(&unknown_owned_attribute, Some(META), None)).is_err());
}

fn package(content: &str, metadata: Option<&str>, settings: Option<&str>) -> Vec<u8> {
    let compact_content = support::compact_xml_fixture(content);
    let compact_metadata = metadata.map(support::compact_xml_fixture);
    let compact_settings = settings.map(support::compact_xml_fixture);
    let mut entries = vec![("content.xml", compact_content.as_bytes(), "text/xml")];
    if let Some(metadata_xml) = compact_metadata.as_deref() {
        entries.push(("meta.xml", metadata_xml.as_bytes(), "text/xml"));
    }
    if let Some(settings_xml) = compact_settings.as_deref() {
        entries.push(("settings.xml", settings_xml.as_bytes(), "text/xml"));
    }
    support::raw_package(&entries)
}
