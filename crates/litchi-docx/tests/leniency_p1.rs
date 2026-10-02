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

// ---------------------------------------------------------------------------
// F2.3 — duplicate singleton settings elements
// ---------------------------------------------------------------------------

#[test]
fn duplicate_singleton_settings_take_the_later_value_and_are_recorded() {
    let settings_xml = format!(
        r#"<w:settings xmlns:w="{W}"><w:view w:val="print"/><w:defaultTabStop w:val="720"/><w:view w:val="web"/><w:compat><w:spaceForUL/></w:compat><w:defaultTabStop w:val="360"/></w:settings>"#
    );
    let second_view = offset_of(&settings_xml, r#"<w:view w:val="web""#);
    let second_tab = offset_of(&settings_xml, r#"<w:defaultTabStop w:val="360""#);
    let package = open(&settings_package(&settings_xml));
    let document = package.document().expect("document");
    let settings = document
        .settings()
        .expect("duplicate settings no longer fail the part")
        .expect("settings present");
    assert_eq!(settings.view(), Some(litchi_docx::settings::View::Web));
    assert_eq!(settings.default_tab_stop_twips(), Some(360));
    assert_eq!(settings.compatibility_options().len(), 1);
    let recorded = [
        ToleratedDefect {
            defect: Defect::Settings(SettingsDefect::DuplicateSetting),
            context: second_view,
        },
        ToleratedDefect {
            defect: Defect::Settings(SettingsDefect::DuplicateSetting),
            context: second_tab,
        },
    ];
    assert_eq!(settings.tolerance_report().defects(), &recorded);
    assert_eq!(document.tolerance_report().defects(), &recorded);
}

#[test]
fn singleton_settings_record_nothing() {
    let settings_xml = format!(
        r#"<w:settings xmlns:w="{W}"><w:view w:val="print"/><w:defaultTabStop w:val="720"/><w:footnotePr><w:numFmt w:val="lowerRoman"/></w:footnotePr></w:settings>"#
    );
    let package = open(&settings_package(&settings_xml));
    let document = package.document().expect("document");
    let settings = document.settings().expect("settings").expect("present");
    assert!(settings.tolerance_report().is_clean());
    assert!(document.tolerance_report().is_clean());
}

// ---------------------------------------------------------------------------
// F2.4 — numbering model degrades per item
// ---------------------------------------------------------------------------

const NUMBERING_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering";
const NUMBERING_CT: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml";

fn numbered_paragraph(num_id: u32, text: &str) -> String {
    format!(
        r#"<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="{num_id}"/></w:numPr></w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>"#
    )
}

const STYLES_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
const STYLES_CT: &str = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";

fn styles_extra(styles: &str) -> Extra<'static> {
    Extra {
        path: "word/styles.xml",
        content_type: STYLES_CT,
        relationship: STYLES_REL,
        xml: format!(r#"<w:styles xmlns:w="{W}">{styles}</w:styles>"#),
    }
}

fn numbering_package(numbering_xml: &str, body: &str) -> Vec<u8> {
    package_bytes(
        &document_xml(body),
        &[
            Extra {
                path: "word/numbering.xml",
                content_type: NUMBERING_CT,
                relationship: NUMBERING_REL,
                xml: numbering_xml.to_owned(),
            },
            styles_extra(""),
        ],
    )
}

#[test]
fn bad_numbering_values_degrade_per_item_and_list_items_still_resolve() {
    use litchi_docx::NumberingDefect;
    let numbering_xml = format!(
        r#"<w:numbering xmlns:w="{W}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="hangulChosung"/><w:lvlText w:val="%1."/><w:suff w:val="colon"/><w:isLgl w:val="yes"/></w:lvl></w:abstractNum><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/></w:lvl></w:abstractNum><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="upperLetter"/><w:lvlText w:val="%1)"/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num><w:num w:numId="3"><w:abstractNumId w:val="9"/></w:num></w:numbering>"#
    );
    let body = [
        numbered_paragraph(1, "first"),
        numbered_paragraph(2, "second"),
        numbered_paragraph(3, "dangling"),
        numbered_paragraph(1, "third"),
    ]
    .concat();
    let package = open(&numbering_package(&numbering_xml, &body));
    let document = package.document().expect("document");

    let items = document
        .list_items()
        .expect("one bad numbering value no longer kills list_items");
    let rendered = items
        .iter()
        .map(|item| (item.paragraph_index, item.text.as_str()))
        .collect::<Vec<_>>();
    assert_eq!(rendered, [(0, "first"), (1, "second"), (3, "third")]);
    assert_eq!(
        items[0].marker,
        litchi_docx::list::ListMarker::Text("1.".to_owned())
    );
    assert_eq!(
        items[1].marker,
        litchi_docx::list::ListMarker::Text("A)".to_owned())
    );

    let at = |needle: &str| offset_of(&numbering_xml, needle);
    let recorded = document
        .tolerance_report()
        .defects()
        .iter()
        .map(|tolerated| (tolerated.defect, tolerated.context))
        .collect::<Vec<_>>();
    assert_eq!(
        recorded,
        [
            (
                Defect::Numbering(NumberingDefect::UnknownNumberFormat),
                at(r#"<w:numFmt w:val="hangulChosung"/>"#)
            ),
            (
                Defect::Numbering(NumberingDefect::UnknownLevelSuffix),
                at(r#"<w:suff w:val="colon"/>"#)
            ),
            (
                Defect::Numbering(NumberingDefect::InvalidOnOff),
                at(r#"<w:isLgl w:val="yes"/>"#)
            ),
            (
                Defect::Numbering(NumberingDefect::DuplicateDefinition),
                at(
                    r#"<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/>"#
                )
            ),
            (
                Defect::Numbering(NumberingDefect::MissingAbstractDefinition),
                at(r#"<w:num w:numId="3">"#)
            ),
            // Context: the paragraph index in the main story.
            (
                Defect::Numbering(NumberingDefect::MissingNumberingInstance),
                2
            ),
        ]
    );

    // The snapshot accessor carries the same part-level repairs.
    let snapshot = document
        .numbering_snapshot()
        .expect("numbering snapshot")
        .expect("numbering present");
    assert_eq!(snapshot.tolerance_report().total(), 5);
}

#[test]
fn conformant_numbering_records_nothing() {
    let numbering_xml = format!(
        r#"<w:numbering xmlns:w="{W}"><w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="singleLevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:suff w:val="tab"/><w:isLgl w:val="0"/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>"#
    );
    let package = open(&numbering_package(
        &numbering_xml,
        &numbered_paragraph(1, "only"),
    ));
    let document = package.document().expect("document");
    assert_eq!(document.list_items().expect("list items").len(), 1);
    assert!(document.tolerance_report().is_clean());
}

#[test]
fn ill_formed_numbering_xml_stays_fatal() {
    let numbering_xml = format!(
        r#"<w:numbering xmlns:w="{W}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"></w:abstractNum></w:numbering>"#
    );
    let package = open(&numbering_package(
        &numbering_xml,
        &numbered_paragraph(1, "only"),
    ));
    assert!(package.document().expect("document").list_items().is_err());
}
