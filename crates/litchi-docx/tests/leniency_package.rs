//! The OPC packaging and core/custom-properties repair reports are folded
//! into the WordprocessingML document report on every open path.

use std::io;

use litchi_docx::{Defect, Package, ToleranceReport, source_backed};
use litchi_ooxml_common::leniency::{
    CorePropertiesDefect, CustomPropertiesDefect, Defect as PropertiesDefect,
};
use litchi_opc::leniency::{Defect as PackageDefect, RelationshipDefect};
use soapberry_zip::office::StreamingArchiveWriter;

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const HYPERLINK: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
const FMTID: &str = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}";

/// What each package carries.
#[derive(Default)]
struct Fixture {
    duplicate_relationship_id: bool,
    bad_custom_property: bool,
    untyped_created_date: bool,
    ill_formed_document_rels: bool,
}

fn package_bytes(fixture: &Fixture) -> Vec<u8> {
    let content_types = r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/></Types>"#;
    let package_rels = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/></Relationships>"#;
    let document = format!(
        r#"<?xml version="1.0"?><w:document xmlns:w="{W}"><w:body><w:p><w:r><w:t>body</w:t></w:r></w:p><w:sectPr/></w:body></w:document>"#
    );
    let duplicate = if fixture.duplicate_relationship_id {
        format!(
            r#"<Relationship Id="rIdLink" Type="{HYPERLINK}" Target="https://example.com/second" TargetMode="External"/>"#
        )
    } else {
        String::new()
    };
    let unclosed = if fixture.ill_formed_document_rels {
        "<Relationship"
    } else {
        ""
    };
    let document_rels = format!(
        r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdLink" Type="{HYPERLINK}" Target="https://example.com/first" TargetMode="External"/>{duplicate}{unclosed}</Relationships>"#
    );
    let created_type = if fixture.untyped_created_date {
        ""
    } else {
        r#" xsi:type="dcterms:W3CDTF""#
    };
    let core = format!(
        r#"<?xml version="1.0"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>Title</dc:title><dcterms:created{created_type}>2024-01-02T03:04:05Z</dcterms:created></cp:coreProperties>"#
    );
    let bad = if fixture.bad_custom_property {
        format!(
            r#"<property fmtid="{FMTID}" pid="3" name="Bad"><vt:bool>maybe</vt:bool></property>"#
        )
    } else {
        String::new()
    };
    let custom = format!(
        r#"<?xml version="1.0"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><property fmtid="{FMTID}" pid="2" name="Good"><vt:lpwstr>ok</vt:lpwstr></property>{bad}</Properties>"#
    );

    let mut writer = StreamingArchiveWriter::new();
    for (name, bytes) in [
        ("[Content_Types].xml", content_types.as_bytes()),
        ("_rels/.rels", package_rels.as_bytes()),
        ("word/document.xml", document.as_bytes()),
        ("word/_rels/document.xml.rels", document_rels.as_bytes()),
        ("docProps/core.xml", core.as_bytes()),
        ("docProps/custom.xml", custom.as_bytes()),
    ] {
        writer.write_stored(name, bytes).expect("archive member");
    }
    writer.finish_to_bytes().expect("finish DOCX archive")
}

fn open(bytes: &[u8]) -> Package {
    Package::from_reader(io::Cursor::new(bytes.to_vec())).expect("package opens")
}

fn open_source_backed(bytes: &[u8]) -> source_backed::Package {
    source_backed::Package::from_reader(io::Cursor::new(bytes.to_vec()))
        .expect("source-backed package opens")
}

/// `(defect, part, item, value)` for every located entry.
fn located(report: &ToleranceReport) -> Vec<(Defect, String, Option<String>, Option<String>)> {
    report
        .defects()
        .iter()
        .map(|tolerated| {
            let location = tolerated
                .location
                .as_deref()
                .expect("package and metadata defects are located");
            assert_eq!(tolerated.context, 0);
            (
                tolerated.defect,
                location.part().to_owned(),
                location.item().map(str::to_owned),
                location.value().map(str::to_owned),
            )
        })
        .collect()
}

fn duplicate_id_entry() -> (Defect, String, Option<String>, Option<String>) {
    (
        Defect::Package(PackageDefect::Relationship(RelationshipDefect::DuplicateId)),
        "/word/document.xml".to_owned(),
        Some("rIdLink".to_owned()),
        Some("https://example.com/second".to_owned()),
    )
}

#[test]
fn duplicate_relationship_id_is_listed_in_every_document_report() {
    let bytes = package_bytes(&Fixture {
        duplicate_relationship_id: true,
        ..Fixture::default()
    });
    let expected = vec![duplicate_id_entry()];

    let package = open(&bytes);
    assert_eq!(located(&package.tolerance_report()), expected);
    let document = package.document().expect("document");
    assert_eq!(located(&document.tolerance_report()), expected);
    assert_eq!(document.text().expect("text"), "body");

    let source = open_source_backed(&bytes);
    assert_eq!(located(&source.tolerance_report()), expected);
    let pinned = source.document().expect("pinned document");
    assert_eq!(located(&pinned.tolerance_report()), expected);

    // Materializing the source-backed package hands the report over.
    let owned = source.to_owned_package().expect("owned package");
    assert_eq!(
        located(&owned.document().expect("document").tolerance_report()),
        expected
    );
}

#[test]
fn malformed_custom_property_is_listed_once_metadata_is_read() {
    let bytes = package_bytes(&Fixture {
        bad_custom_property: true,
        ..Fixture::default()
    });

    // The owning package reads its metadata at open.
    let package = open(&bytes);
    assert!(package.custom_props().get("Good").is_some());
    assert!(package.custom_props().get("Bad").is_none());
    let expected = vec![(
        Defect::Properties(PropertiesDefect::CustomProperties(
            CustomPropertiesDefect::InvalidValue,
        )),
        "/docProps/custom.xml".to_owned(),
        Some("Bad".to_owned()),
        Some("maybe".to_owned()),
    )];
    assert_eq!(located(&package.tolerance_report()), expected);
    let document = package.document().expect("document");
    assert_eq!(located(&document.tolerance_report()), expected);
}

#[test]
fn core_property_repairs_join_the_report_when_metadata_is_read() {
    let bytes = package_bytes(&Fixture {
        untyped_created_date: true,
        ..Fixture::default()
    });
    let expected = vec![(
        Defect::Properties(PropertiesDefect::CoreProperties(
            CorePropertiesDefect::MissingDateType,
        )),
        "/docProps/core.xml".to_owned(),
        Some("created".to_owned()),
        None,
    )];

    let package = open(&bytes);
    assert_eq!(located(&package.tolerance_report()), expected);

    // The source-backed package reads metadata on demand.
    let source = open_source_backed(&bytes);
    assert!(source.tolerance_report().is_clean());
    let metadata = source.metadata().expect("metadata");
    assert_eq!(metadata.title.as_deref(), Some("Title"));
    assert!(metadata.created.is_some());
    assert_eq!(located(&source.tolerance_report()), expected);
}

#[test]
fn conformant_package_and_metadata_report_clean() {
    let bytes = package_bytes(&Fixture::default());

    let package = open(&bytes);
    assert!(package.tolerance_report().is_clean());
    assert!(
        package
            .document()
            .expect("document")
            .tolerance_report()
            .is_clean()
    );
    assert!(package.custom_props().get("Good").is_some());

    let source = open_source_backed(&bytes);
    source.metadata().expect("metadata");
    assert!(source.tolerance_report().is_clean());
    assert!(
        source
            .document()
            .expect("pinned")
            .tolerance_report()
            .is_clean()
    );
}

#[test]
fn ill_formed_relationship_manifest_stays_fatal() {
    let bytes = package_bytes(&Fixture {
        duplicate_relationship_id: true,
        ill_formed_document_rels: true,
        ..Fixture::default()
    });
    assert!(Package::from_reader(io::Cursor::new(bytes.clone())).is_err());
    assert!(source_backed::Package::from_reader(io::Cursor::new(bytes)).is_err());
}
