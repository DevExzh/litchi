//! Phase-2 (P1) leniency acceptance tests for WordprocessingML reads.
//!
//! Each finding gets a hand-built minimal package carrying the deviation,
//! asserting that the read succeeds under the default posture with the repair
//! recorded, plus a conformant control that records nothing. Ill-formed XML
//! stays fatal.

use std::io;

use litchi_docx::{Defect, Package, SettingsDefect, ToleratedDefect};
use soapberry_zip::office::StreamingArchiveWriter;

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const STRICT_W: &str = "http://purl.oclc.org/ooxml/wordprocessingml/main";

const SETTINGS_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings";
const SETTINGS_CT: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml";

/// One extra part of the main document: `(zip path, content type,
/// relationship type, XML)`.
struct Extra<'a> {
    path: &'a str,
    content_type: &'a str,
    relationship: &'a str,
    xml: String,
}

fn document_xml(body: &str) -> String {
    format!(
        r#"<?xml version="1.0"?><w:document xmlns:w="{W}"><w:body>{body}<w:sectPr/></w:body></w:document>"#
    )
}

fn package_bytes(document: &str, extras: &[Extra<'_>]) -> Vec<u8> {
    let mut overrides = String::new();
    let mut relationships = String::new();
    for (index, extra) in extras.iter().enumerate() {
        overrides.push_str(&format!(
            r#"<Override PartName="/{}" ContentType="{}"/>"#,
            extra.path, extra.content_type
        ));
        let target = extra.path.strip_prefix("word/").expect("word part");
        relationships.push_str(&format!(
            r#"<Relationship Id="rIdX{index}" Type="{}" Target="{target}"/>"#,
            extra.relationship
        ));
    }
    let content_types = format!(
        r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>{overrides}</Types>"#
    );
    let mut writer = StreamingArchiveWriter::new();
    writer
        .write_stored("[Content_Types].xml", content_types.as_bytes())
        .expect("content types");
    writer
        .write_stored(
            "_rels/.rels",
            br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#,
        )
        .expect("package relationships");
    writer
        .write_stored("word/document.xml", document.as_bytes())
        .expect("main document");
    if !extras.is_empty() {
        let rels = format!(
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{relationships}</Relationships>"#
        );
        writer
            .write_stored("word/_rels/document.xml.rels", rels.as_bytes())
            .expect("document relationships");
    }
    for extra in extras {
        writer
            .write_stored(extra.path, extra.xml.as_bytes())
            .expect("extra part");
    }
    writer.finish_to_bytes().expect("finish DOCX archive")
}

fn open(bytes: &[u8]) -> Package {
    Package::from_reader(io::Cursor::new(bytes.to_vec())).expect("package opens")
}

/// Offset of the first occurrence of `needle` in `haystack`.
fn offset_of(haystack: &str, needle: &str) -> u64 {
    haystack.find(needle).expect("needle in fixture") as u64
}

fn settings_package(settings_xml: &str) -> Vec<u8> {
    package_bytes(
        &document_xml(r"<w:p><w:r><w:t>body</w:t></w:r></w:p>"),
        &[Extra {
            path: "word/settings.xml",
            content_type: SETTINGS_CT,
            relationship: SETTINGS_REL,
            xml: settings_xml.to_owned(),
        }],
    )
}

// ---------------------------------------------------------------------------
// F2.2 — unknown `w:compat` flag name
// ---------------------------------------------------------------------------

#[test]
fn unknown_compat_flag_is_skipped_and_recorded() {
    let settings_xml = format!(
        r#"<w:settings xmlns:w="{W}"><w:trackRevisions/><w:compat><w:useFELayout/><w:someFutureFlag w:val="1"/><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat><w:view w:val="print"/></w:settings>"#
    );
    let expected = offset_of(&settings_xml, "<w:someFutureFlag");
    let package = open(&settings_package(&settings_xml));
    let document = package.document().expect("document");

    let settings = document
        .settings()
        .expect("unknown compat flag no longer fails settings")
        .expect("settings present");
    assert!(settings.track_revisions());
    assert_eq!(settings.compatibility_options().len(), 1);
    assert_eq!(settings.compatibility_mode(), Some(15));
    assert!(settings.view().is_some());
    let recorded = [ToleratedDefect {
        defect: Defect::Settings(SettingsDefect::UnknownCompatFlag),
        context: expected,
    }];
    assert_eq!(settings.tolerance_report().defects(), &recorded);
    assert_eq!(document.tolerance_report().defects(), &recorded);

    // The package-level shortcut that used to fail now reads through.
    assert!(
        package
            .attached_template()
            .expect("settings read")
            .is_none()
    );
}

#[test]
fn transitional_compat_flag_in_strict_settings_is_kept_and_recorded() {
    let settings_xml = format!(
        r#"<w:settings xmlns:w="{STRICT_W}"><w:compat><w:useFELayout/></w:compat></w:settings>"#
    );
    let expected = offset_of(&settings_xml, "<w:useFELayout");
    let package = open(&settings_package(&settings_xml));
    let settings = package
        .document()
        .expect("document")
        .settings()
        .expect("settings read")
        .expect("settings present");
    assert_eq!(settings.compatibility_options().len(), 1);
    assert_eq!(
        settings.tolerance_report().defects(),
        &[ToleratedDefect {
            defect: Defect::Settings(SettingsDefect::TransitionalCompatFlagInStrict),
            context: expected,
        }]
    );
}

#[test]
fn conformant_compat_flags_record_nothing() {
    let settings_xml = format!(
        r#"<w:settings xmlns:w="{W}"><w:compat><w:useFELayout/><w:doNotExpandShiftReturn w:val="off"/></w:compat></w:settings>"#
    );
    let package = open(&settings_package(&settings_xml));
    let document = package.document().expect("document");
    let settings = document.settings().expect("settings").expect("present");
    assert_eq!(settings.compatibility_options().len(), 2);
    assert!(settings.tolerance_report().is_clean());
    assert!(document.tolerance_report().is_clean());
}

#[test]
fn ill_formed_settings_xml_stays_fatal() {
    let settings_xml = format!(
        r#"<w:settings xmlns:w="{W}"><w:compat><w:someFutureFlag></w:compat></w:settings>"#
    );
    let package = open(&settings_package(&settings_xml));
    assert!(package.document().expect("document").settings().is_err());
}
