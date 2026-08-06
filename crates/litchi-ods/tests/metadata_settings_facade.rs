use litchi_core::Metadata;
use litchi_odf_common::calculation::{Iteration, IterationStatus, Settings};
use litchi_odf_common::core::{OwnedPackage, PackageWriter};
use litchi_ods::{Builder, MutableSpreadsheet, Spreadsheet};
use std::num::NonZeroUsize;

const MIMETYPE: &str = "application/vnd.oasis.opendocument.spreadsheet";
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
    let spreadsheet = Spreadsheet::from_bytes(bytes.clone()).unwrap();
    assert_eq!(spreadsheet.metadata().title.as_deref(), Some("Before"));
    assert_eq!(spreadsheet.metadata().author.as_deref(), Some("Author"));
    assert_eq!(spreadsheet.odf_metadata().title.as_deref(), Some("Before"));
    assert_eq!(spreadsheet.settings().unwrap().case_sensitive, Some(false));
    assert!(
        spreadsheet
            .metadata_snapshot()
            .xml()
            .unwrap()
            .contains("vendor:opaque")
    );

    let mut mutable = MutableSpreadsheet::from_bytes(bytes).unwrap();
    mutable
        .update_metadata(|metadata| {
            metadata.title = Some("After".to_string());
            metadata.description = Some("Edited safely".to_string());
            Ok(())
        })
        .unwrap();
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
        .unwrap();

    let output = mutable.to_bytes();
    let reopened = Spreadsheet::from_bytes(output.clone()).unwrap();
    assert_eq!(reopened.metadata().title.as_deref(), Some("After"));
    assert_eq!(
        reopened.metadata().description.as_deref(),
        Some("Edited safely")
    );
    assert_eq!(reopened.settings().unwrap().case_sensitive, Some(true));
    assert_eq!(
        reopened
            .settings()
            .unwrap()
            .iteration
            .as_ref()
            .unwrap()
            .steps,
        NonZeroUsize::new(20)
    );
    assert!(reopened.content_xml().contains("vendor:extension"));
    assert!(
        reopened
            .metadata_snapshot()
            .xml()
            .unwrap()
            .contains("vendor:opaque")
    );

    let archive = OwnedPackage::from_bytes(output).unwrap();
    assert_eq!(
        archive.get_file("settings.xml").unwrap(),
        SETTINGS.as_bytes()
    );

    let mut clear = MutableSpreadsheet::from_bytes(package(CONTENT, Some(META), None)).unwrap();
    clear.clear_metadata().unwrap();
    let cleared = Spreadsheet::from_bytes(clear.to_bytes()).unwrap();
    assert!(!cleared.metadata_snapshot().is_present());
}

#[test]
fn facade_edits_are_atomic_and_builder_writes_new_parts() {
    let mut mutable = MutableSpreadsheet::from_bytes(package(CONTENT, Some(META), None)).unwrap();
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
        .unwrap();
    builder
        .set_settings(Some(Settings {
            precision_as_shown: Some(true),
            ..Settings::default()
        }))
        .unwrap();
    let spreadsheet = Spreadsheet::from_bytes(builder.build().unwrap()).unwrap();
    assert_eq!(spreadsheet.metadata().title.as_deref(), Some("Built"));
    assert_eq!(spreadsheet.metadata().author.as_deref(), Some("Builder"));
    assert_eq!(
        spreadsheet.settings().unwrap().precision_as_shown,
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
    let mut writer = PackageWriter::new();
    writer.set_mimetype(MIMETYPE).unwrap();
    writer.add_file("content.xml", content.as_bytes()).unwrap();
    if let Some(metadata) = metadata {
        writer.add_file("meta.xml", metadata.as_bytes()).unwrap();
    }
    if let Some(settings) = settings {
        writer
            .add_file("settings.xml", settings.as_bytes())
            .unwrap();
    }
    writer.finish_to_bytes().unwrap()
}
