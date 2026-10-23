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
        location: None,
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
            location: None,
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
            location: None,
        },
        ToleratedDefect {
            defect: Defect::Settings(SettingsDefect::DuplicateSetting),
            context: second_tab,
            location: None,
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

// ---------------------------------------------------------------------------
// F2.5 — tracked revision without w:author
// ---------------------------------------------------------------------------

#[test]
fn revision_without_author_is_kept_unattributed_and_recorded() {
    use litchi_docx::RevisionDefect;
    use litchi_docx::source_backed::Package as SourceBackedPackage;
    let document = document_xml(
        r#"<w:p><w:ins w:id="1" w:author="Alice" w:date="2026-01-01T00:00:00Z"><w:r><w:t>kept</w:t></w:r></w:ins><w:ins w:id="2" w:date="2026-01-02T00:00:00Z"><w:r><w:t>anon</w:t></w:r></w:ins><w:del w:id="3" w:author="Bob" w:date="yesterday"><w:r><w:delText>gone</w:delText></w:r></w:del></w:p>"#,
    );
    let anonymous = offset_of(&document, r#"<w:ins w:id="2""#);
    let undated = offset_of(&document, r#"<w:del w:id="3""#);
    let bytes = package_bytes(&document, &[]);

    let package = open(&bytes);
    let opened = package.document().expect("document");
    let revisions = opened
        .revisions()
        .expect("a missing author no longer kills the listing");
    assert_eq!(revisions.len(), 3);
    assert_eq!(revisions[0].author(), Some("Alice"));
    assert_eq!(revisions[1].author(), None);
    assert_eq!(revisions[1].text(), "anon");
    assert_eq!(revisions[2].date(), None);
    assert_eq!(revisions[2].text(), "gone");
    let recorded = [
        ToleratedDefect {
            defect: Defect::Revision(RevisionDefect::MissingAuthor),
            context: anonymous,
            location: None,
        },
        ToleratedDefect {
            defect: Defect::Revision(RevisionDefect::InvalidDate),
            context: undated,
            location: None,
        },
    ];
    assert_eq!(opened.tolerance_report().defects(), &recorded);

    // The pinned source-backed view reports the same repairs.
    let source = SourceBackedPackage::from_read_at(std::sync::Arc::new(
        litchi_core::OwnedSource::new(bytes),
    ))
    .expect("source package");
    let pinned = source.document().expect("pinned document");
    assert_eq!(pinned.revisions().expect("revisions").len(), 3);
    assert_eq!(pinned.tolerance_report().defects(), &recorded);
}

#[test]
fn attributed_revisions_record_nothing() {
    let document = document_xml(
        r#"<w:p><w:ins w:id="1" w:author="Alice" w:date="2026-01-01T00:00:00Z"><w:r><w:t>kept</w:t></w:r></w:ins></w:p>"#,
    );
    let package = open(&package_bytes(&document, &[]));
    let opened = package.document().expect("document");
    assert_eq!(opened.revisions().expect("revisions").len(), 1);
    assert!(opened.tolerance_report().is_clean());
}

#[test]
fn ill_formed_revision_xml_stays_fatal() {
    let document = document_xml(r#"<w:p><w:ins w:id="1"><w:r><w:t>x</w:r></w:ins></w:p>"#);
    let package = open(&package_bytes(&document, &[]));
    assert!(package.document().expect("document").revisions().is_err());
}

// ---------------------------------------------------------------------------
// F2.6 — w:numPr with only w:ilvl
// ---------------------------------------------------------------------------

#[test]
fn ilvl_only_numpr_inherits_style_numbering_and_keeps_style_id() {
    use litchi_docx::NumberingDefect;
    let numbering_xml = format!(
        r#"<w:numbering xmlns:w="{W}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl><w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%2."/></w:lvl><w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="lowerRoman"/><w:lvlText w:val="%3)"/></w:lvl></w:abstractNum><w:num w:numId="4"><w:abstractNumId w:val="0"/></w:num></w:numbering>"#
    );
    let styles = r#"<w:style w:type="paragraph" w:styleId="ListPara"><w:name w:val="List Para"/><w:pPr><w:numPr><w:numId w:val="4"/></w:numPr></w:pPr></w:style>"#;
    let body = r#"<w:p><w:pPr><w:pStyle w:val="ListPara"/><w:numPr><w:ilvl w:val="2"/></w:numPr></w:pPr><w:r><w:t>override</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val="ListPara"/><w:numPr/></w:pPr><w:r><w:t>empty</w:t></w:r></w:p>"#;
    let bytes = package_bytes(
        &document_xml(body),
        &[
            Extra {
                path: "word/numbering.xml",
                content_type: NUMBERING_CT,
                relationship: NUMBERING_REL,
                xml: numbering_xml,
            },
            styles_extra(styles),
        ],
    );
    let package = open(&bytes);
    let document = package.document().expect("document");

    let paragraphs = document.paragraphs().expect("paragraphs");
    assert_eq!(paragraphs[0].numbering().expect("numbering"), None);
    assert_eq!(
        paragraphs[0].style_id().expect("style id is not poisoned"),
        Some("ListPara".to_owned())
    );

    let items = document.list_items().expect("list items");
    assert_eq!(items.len(), 2);
    assert_eq!(items[0].numbering.num_id, 4);
    assert_eq!(items[0].numbering.level, 2);
    assert_eq!(
        items[0].marker,
        litchi_docx::list::ListMarker::Text("i)".to_owned())
    );
    assert_eq!(items[1].numbering.level, 0);
    assert_eq!(
        document
            .tolerance_report()
            .defects()
            .iter()
            .map(|tolerated| (tolerated.defect, tolerated.context))
            .collect::<Vec<_>>(),
        [
            (Defect::Numbering(NumberingDefect::NumPrMissingNumId), 0),
            (Defect::Numbering(NumberingDefect::NumPrMissingNumId), 1),
        ]
    );
}

#[test]
fn malformed_numpr_value_does_not_poison_style_id() {
    let document = document_xml(
        r#"<w:p><w:pPr><w:pStyle w:val="Body"/><w:numPr><w:numId w:val="many"/></w:numPr></w:pPr><w:r><w:t>x</w:t></w:r></w:p>"#,
    );
    let package = open(&package_bytes(&document, &[]));
    let paragraphs = package
        .document()
        .expect("document")
        .paragraphs()
        .expect("paragraphs");
    assert!(paragraphs[0].numbering().is_err());
    assert_eq!(
        paragraphs[0].style_id().expect("style id"),
        Some("Body".to_owned())
    );
}

#[test]
fn complete_numpr_records_nothing() {
    let numbering_xml = format!(
        r#"<w:numbering xmlns:w="{W}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>"#
    );
    let package = open(&numbering_package(
        &numbering_xml,
        &numbered_paragraph(1, "only"),
    ));
    let document = package.document().expect("document");
    let paragraphs = document.paragraphs().expect("paragraphs");
    assert_eq!(
        paragraphs[0]
            .numbering()
            .expect("numbering")
            .map(|n| n.num_id),
        Some(1)
    );
    assert_eq!(document.list_items().expect("list items").len(), 1);
    assert!(document.tolerance_report().is_clean());
}

// ---------------------------------------------------------------------------
// F2.7 — section column without w:w / column-model validation
// ---------------------------------------------------------------------------

#[test]
fn sparse_section_columns_are_repaired_and_recorded_in_the_inventory() {
    use litchi_docx::SectionDefect;
    let document = format!(
        r#"<?xml version="1.0"?><w:document xmlns:w="{W}"><w:body><w:p><w:pPr><w:sectPr><w:cols w:num="2" w:equalWidth="0"><w:col w:space="720"/><w:col w:w="4000"/></w:cols></w:sectPr></w:pPr><w:r><w:t>one</w:t></w:r></w:p><w:p><w:r><w:t>two</w:t></w:r></w:p><w:sectPr><w:cols w:num="99"/></w:sectPr></w:body></w:document>"#
    );
    let missing_width = offset_of(&document, r#"<w:col w:space="720"/>"#);
    let out_of_range = offset_of(&document, r#"<w:cols w:num="99"/>"#);
    let package = open(&package_bytes(&document, &[]));
    let opened = package.document().expect("document");

    let inventory = opened
        .section_inventory()
        .expect("a sparse w:col no longer kills the section inventory");
    assert_eq!(inventory.sections().len(), 2);
    let first = inventory.sections()[0].columns().expect("first columns");
    assert_eq!(first.columns[0].width, litchi_docx::Emu(0));
    assert_eq!(inventory.sections()[1].columns().expect("final").count, 45);
    let recorded = [
        ToleratedDefect {
            defect: Defect::Section(SectionDefect::ColumnWidthMissing),
            context: missing_width,
            location: None,
        },
        ToleratedDefect {
            defect: Defect::Section(SectionDefect::ColumnCountOutOfRange),
            context: out_of_range,
            location: None,
        },
    ];
    assert_eq!(inventory.tolerance_report().defects(), &recorded);
    assert_eq!(opened.tolerance_report().defects(), &recorded);

    // The section model reports the same repairs per section.
    let mut sections = opened.sections().expect("sections");
    let section = sections.get_mut(0).expect("first section");
    assert_eq!(
        section
            .columns()
            .expect("columns")
            .expect("present")
            .columns[0]
            .width,
        litchi_docx::Emu(0)
    );
    assert_eq!(section.tolerance_report().expect("report").total(), 1);
}

#[test]
fn conformant_section_columns_record_nothing() {
    let document = format!(
        r#"<?xml version="1.0"?><w:document xmlns:w="{W}"><w:body><w:p><w:r><w:t>one</w:t></w:r></w:p><w:sectPr><w:cols w:num="2" w:equalWidth="0"><w:col w:w="4000" w:space="720"/><w:col w:w="4000"/></w:cols></w:sectPr></w:body></w:document>"#
    );
    let package = open(&package_bytes(&document, &[]));
    let opened = package.document().expect("document");
    let inventory = opened.section_inventory().expect("inventory");
    assert!(inventory.tolerance_report().is_clean());
    assert!(opened.tolerance_report().is_clean());
}

// ---------------------------------------------------------------------------
// F2.8 — one malformed w:sdtPr degrades only its control
// ---------------------------------------------------------------------------

fn sdt(properties: &str, text: &str) -> String {
    format!(
        r#"<w:sdt><w:sdtPr>{properties}</w:sdtPr><w:sdtContent><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:sdtContent></w:sdt>"#
    )
}

#[test]
fn malformed_sdt_properties_degrade_one_control_and_keep_the_listing() {
    use litchi_docx::ContentControlDefect;
    let body = [
        sdt(r#"<w:id w:val="1"/><w:tag w:val="first"/><w:text/>"#, "one"),
        sdt(
            r#"<w:id w:val="2"/><w:tag w:val="second"/><w:lock w:val="frozen"/><w:dataBinding w:xpath="/root"/>"#,
            "two",
        ),
        sdt(r#"<w:id w:val="3"/><w:tag w:val="third"/>"#, "three"),
    ]
    .concat();
    let document = document_xml(&body);
    let lock = offset_of(&document, r#"<w:lock w:val="frozen"/>"#);
    let binding = offset_of(&document, r#"<w:dataBinding w:xpath="/root"/>"#);
    let package = open(&package_bytes(&document, &[]));
    let opened = package.document().expect("document");

    let controls = opened
        .content_controls()
        .expect("one malformed sdtPr no longer kills the listing");
    assert_eq!(
        controls
            .iter()
            .map(|control| (control.id(), control.tag().map(str::to_owned)))
            .collect::<Vec<_>>(),
        [
            (1, Some("first".to_owned())),
            (2, Some("second".to_owned())),
            (3, Some("third".to_owned())),
        ]
    );
    assert!(!controls[1].is_lock_delete());
    assert!(controls[1].data_bindings().is_empty());
    assert_eq!(
        opened.tolerance_report().defects(),
        &[
            ToleratedDefect {
                defect: Defect::ContentControl(ContentControlDefect::UnknownLockValue),
                context: lock,
                location: None,
            },
            ToleratedDefect {
                defect: Defect::ContentControl(ContentControlDefect::InvalidProperty),
                context: binding,
                location: None,
            },
        ]
    );
}

#[test]
fn conformant_sdt_properties_record_nothing() {
    let body = sdt(
        r#"<w:id w:val="1"/><w:tag w:val="first"/><w:lock w:val="sdtLocked"/><w:text/>"#,
        "one",
    );
    let package = open(&package_bytes(&document_xml(&body), &[]));
    let opened = package.document().expect("document");
    assert_eq!(opened.content_controls().expect("controls").len(), 1);
    assert!(opened.tolerance_report().is_clean());
}

#[test]
fn nested_sdt_properties_stay_fatal() {
    let body = r#"<w:sdt><w:sdtPr><w:id w:val="1"/><w:sdtPr/></w:sdtPr></w:sdt>"#;
    let package = open(&package_bytes(&document_xml(body), &[]));
    assert!(
        package
            .document()
            .expect("document")
            .content_controls()
            .is_err()
    );
}
