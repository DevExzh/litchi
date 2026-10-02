#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

//! Packages carrying cataloged OPC non-conformities (docs/LENIENCY_AUDIT.md
//! section 8) open by default, and every repair is recorded in the package's
//! tolerance report. A conformant control package reports clean through every
//! ingress path.

use litchi_opc::leniency::{ContentTypeDefect, Defect, RelationshipDefect, ToleranceReport};
use litchi_opc::{OpcPackage, PackURI, SourceBackedPackage, TargetMode};
use soapberry_zip::office::StreamingArchiveWriter;

const CONTENT_TYPES: &str = r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#;

const ROOT_RELS: &str = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#;

const DOCUMENT_RELS: &str = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/" TargetMode="External"/></Relationships>"#;

const DOCUMENT: &str = r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body/></w:document>"#;
const STYLES: &str =
    r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"#;

/// Build a minimal package from the given manifests.
fn package(content_types: &str, root_rels: &str, document_rels: &str) -> Vec<u8> {
    let mut writer = StreamingArchiveWriter::new();
    writer
        .write_stored("[Content_Types].xml", content_types.as_bytes())
        .unwrap();
    writer
        .write_stored("_rels/.rels", root_rels.as_bytes())
        .unwrap();
    writer
        .write_stored("word/_rels/document.xml.rels", document_rels.as_bytes())
        .unwrap();
    writer
        .write_stored("word/document.xml", DOCUMENT.as_bytes())
        .unwrap();
    writer
        .write_stored("word/styles.xml", STYLES.as_bytes())
        .unwrap();
    writer.finish_to_bytes().unwrap()
}

/// Open `bytes` through every ingress path and return each path's report,
/// asserting they agree.
fn reports(bytes: &[u8]) -> ToleranceReport {
    let eager = OpcPackage::from_bytes(bytes).expect("eager open");
    let owned = OpcPackage::from_vec(bytes.to_vec()).expect("owned open");
    let source = SourceBackedPackage::from_vec(bytes.to_vec()).expect("source-backed open");
    let materialized = source.to_opc_package().expect("materialize");

    let report = eager.tolerance_report().clone();
    assert_eq!(owned.tolerance_report(), &report, "owned ingress");
    assert_eq!(source.tolerance_report(), &report, "source-backed ingress");
    assert_eq!(
        materialized.tolerance_report(),
        &report,
        "materialized source-backed package"
    );
    report
}

/// Defect classes and parts, in discovery order.
fn classes(report: &ToleranceReport) -> Vec<(Defect, String, Option<String>)> {
    report
        .iter()
        .map(|tolerated| {
            (
                tolerated.defect(),
                tolerated.part().to_owned(),
                tolerated.item().map(str::to_owned),
            )
        })
        .collect()
}

#[test]
fn conformant_package_reports_clean_through_every_ingress() {
    let report = reports(&package(CONTENT_TYPES, ROOT_RELS, DOCUMENT_RELS));
    assert!(report.is_clean(), "{:?}", classes(&report));
    assert_eq!(report.total(), 0);
    assert_eq!(report.unrecorded(), 0);
}

#[test]
fn new_packages_report_clean() {
    assert!(OpcPackage::new().tolerance_report().is_clean());
}

/// A repaired package still saves, and the saved package reopens with the
/// same repairs recorded (an untouched manifest publishes its source bytes).
fn assert_round_trips(bytes: &[u8]) {
    let package = OpcPackage::from_bytes(bytes).expect("open");
    let mut saved = Vec::new();
    package
        .to_stream(std::io::Cursor::new(&mut saved))
        .expect("save");
    let reopened = OpcPackage::from_bytes(&saved).expect("reopen");
    assert_eq!(reopened.tolerance_report(), package.tolerance_report());
}

fn relationship_type(package: &OpcPackage, source: &str, id: &str) -> Option<String> {
    let source = PackURI::new(source).unwrap();
    package
        .get_part(&source)
        .unwrap()
        .rels()
        .get(id)
        .map(|relationship| relationship.reltype().to_owned())
}

const STYLES_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";

/// LENIENCY_AUDIT 8.5: a duplicated relationship `Id` keeps the first
/// relationship (LibreOffice's first match) and records the dropped one.
#[test]
fn duplicate_relationship_id_keeps_the_first_and_is_recorded() {
    let rels = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/" TargetMode="External"/></Relationships>"#;
    let bytes = package(CONTENT_TYPES, ROOT_RELS, rels);
    let report = reports(&bytes);
    assert_eq!(
        classes(&report),
        vec![(
            Defect::Relationship(RelationshipDefect::DuplicateId),
            "/word/document.xml".to_owned(),
            Some("rId1".to_owned()),
        )]
    );
    assert_eq!(report.defects()[0].value(), Some("https://example.com/"));

    let package = OpcPackage::from_bytes(&bytes).unwrap();
    assert_eq!(
        relationship_type(&package, "/word/document.xml", "rId1").as_deref(),
        Some(STYLES_TYPE)
    );
    assert_round_trips(&bytes);
}

/// LENIENCY_AUDIT 8.6: an internal target that cannot be resolved to a part
/// name (literal space, backslash, absolute URI, root escape) stays visible
/// as a dangling relationship, is not followed, and is recorded.
#[test]
fn unresolvable_internal_targets_stay_dangling_and_are_recorded() {
    let rels = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image 1.png"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media\image2.png"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../../../escape.png"/><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="http://example.com/a.png"/></Relationships>"#;
    let bytes = package(CONTENT_TYPES, ROOT_RELS, rels);
    let report = reports(&bytes);
    let unresolvable = Defect::Relationship(RelationshipDefect::UnresolvableTarget);
    assert_eq!(
        report
            .iter()
            .map(|tolerated| (tolerated.defect(), tolerated.item(), tolerated.value()))
            .collect::<Vec<_>>(),
        vec![
            (unresolvable, Some("rId2"), Some("media/image 1.png")),
            (unresolvable, Some("rId3"), Some("media\\image2.png")),
            (unresolvable, Some("rId4"), Some("../../../escape.png")),
            (unresolvable, Some("rId5"), Some("http://example.com/a.png")),
        ]
    );
    assert!(report.iter().all(|t| t.part() == "/word/document.xml"));

    let package = OpcPackage::from_bytes(&bytes).unwrap();
    let document = package
        .get_part(&PackURI::new("/word/document.xml").unwrap())
        .unwrap();
    assert_eq!(
        document.rels().len(),
        5,
        "dangling relationships stay visible"
    );
    assert!(
        document
            .rels()
            .get("rId2")
            .unwrap()
            .target_partname()
            .is_err()
    );
    assert_eq!(
        relationship_type(&package, "/word/document.xml", "rId1").as_deref(),
        Some(STYLES_TYPE)
    );
    assert_round_trips(&bytes);
}

/// LENIENCY_AUDIT 8.7: `TargetMode` is matched case-insensitively (recorded
/// when non-canonical); any other value reads as `External` (Apache POI) so
/// it can never resolve to a part, and is recorded.
#[test]
fn target_mode_is_case_insensitive_and_unknown_modes_read_as_external() {
    let rels = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml" TargetMode="internal"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/" TargetMode="EXTERNAL"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.org/" TargetMode="Remote"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.net/" TargetMode="External"/></Relationships>"#;
    let bytes = package(CONTENT_TYPES, ROOT_RELS, rels);
    let report = reports(&bytes);
    assert_eq!(
        report
            .iter()
            .map(|tolerated| (tolerated.defect(), tolerated.item(), tolerated.value()))
            .collect::<Vec<_>>(),
        vec![
            (
                Defect::Relationship(RelationshipDefect::NonCanonicalTargetMode),
                Some("rId1"),
                Some("internal"),
            ),
            (
                Defect::Relationship(RelationshipDefect::NonCanonicalTargetMode),
                Some("rId2"),
                Some("EXTERNAL"),
            ),
            (
                Defect::Relationship(RelationshipDefect::UnrecognizedTargetMode),
                Some("rId3"),
                Some("Remote"),
            ),
        ]
    );

    let package = OpcPackage::from_bytes(&bytes).unwrap();
    let document = package
        .get_part(&PackURI::new("/word/document.xml").unwrap())
        .unwrap();
    let rels = document.rels();
    assert_eq!(
        rels.get("rId1").unwrap().target_mode(),
        TargetMode::Internal
    );
    assert_eq!(
        rels.get("rId2").unwrap().target_mode(),
        TargetMode::External
    );
    assert_eq!(
        rels.get("rId3").unwrap().target_mode(),
        TargetMode::External
    );
    assert_eq!(
        rels.get("rId4").unwrap().target_mode(),
        TargetMode::External
    );
    assert_round_trips(&bytes);
}

/// LENIENCY_AUDIT 8.8: a relationship with a missing or empty `Target` or
/// `Type` is skipped and recorded; its Id stays free for a later
/// well-formed relationship.
#[test]
fn relationships_without_target_or_type_are_skipped_and_recorded() {
    let rels = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" TargetMode="External"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="" TargetMode="External"/><Relationship Id="rId4" Target="https://example.com/" TargetMode="External"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.org/" TargetMode="External"/></Relationships>"#;
    let bytes = package(CONTENT_TYPES, ROOT_RELS, rels);
    let report = reports(&bytes);
    let hyperlink = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
    assert_eq!(
        report
            .iter()
            .map(|tolerated| (tolerated.defect(), tolerated.item(), tolerated.value()))
            .collect::<Vec<_>>(),
        vec![
            (
                Defect::Relationship(RelationshipDefect::MissingTarget),
                Some("rId2"),
                Some(hyperlink),
            ),
            (
                Defect::Relationship(RelationshipDefect::MissingTarget),
                Some("rId3"),
                Some(hyperlink),
            ),
            (
                Defect::Relationship(RelationshipDefect::MissingType),
                Some("rId4"),
                Some("https://example.com/"),
            ),
        ]
    );

    let package = OpcPackage::from_bytes(&bytes).unwrap();
    let document = package
        .get_part(&PackURI::new("/word/document.xml").unwrap())
        .unwrap();
    assert_eq!(document.rels().len(), 2);
    assert_eq!(
        document.rels().get("rId2").unwrap().target_ref(),
        "https://example.org/"
    );
    assert_round_trips(&bytes);
}

/// LENIENCY_AUDIT 8.9: duplicate `Default`/`Override` mappings keep the first
/// mapping (LibreOffice's first-match lookup) and record the dropped ones.
#[test]
fn duplicate_content_type_mappings_keep_the_first_and_are_recorded() {
    let content_types = r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="XML" ContentType="text/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/document.xml" ContentType="application/xml"/></Types>"#;
    let bytes = package(content_types, ROOT_RELS, DOCUMENT_RELS);
    let report = reports(&bytes);
    assert_eq!(
        report
            .iter()
            .map(|tolerated| (
                tolerated.defect(),
                tolerated.part(),
                tolerated.item(),
                tolerated.value()
            ))
            .collect::<Vec<_>>(),
        vec![
            (
                Defect::ContentType(ContentTypeDefect::DuplicateDefault),
                "/[Content_Types].xml",
                Some("XML"),
                Some("text/xml"),
            ),
            (
                Defect::ContentType(ContentTypeDefect::DuplicateOverride),
                "/[Content_Types].xml",
                Some("/word/document.xml"),
                Some("application/xml"),
            ),
        ]
    );

    let package = OpcPackage::from_bytes(&bytes).unwrap();
    let content_type = |name: &str| {
        package
            .get_part(&PackURI::new(name).unwrap())
            .unwrap()
            .content_type()
            .to_owned()
    };
    assert_eq!(
        content_type("/word/document.xml"),
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
    );
    assert_eq!(content_type("/word/styles.xml"), "application/xml");
    assert_round_trips(&bytes);
}

/// Every recorded defect as `(class, part, item, value)`.
fn entries(report: &ToleranceReport) -> Vec<(Defect, &str, Option<&str>, Option<&str>)> {
    report
        .iter()
        .map(|tolerated| {
            (
                tolerated.defect(),
                tolerated.part(),
                tolerated.item(),
                tolerated.value(),
            )
        })
        .collect()
}

const CONTENT_TYPES_PART: &str = "/[Content_Types].xml";

/// LENIENCY_AUDIT 8.10: relationships parts declared with another content
/// type are still read as relationships (identified by name, as LibreOffice
/// and Apache POI do) and the declaration is recorded per part.
#[test]
fn relationships_parts_declared_with_another_type_are_read_and_recorded() {
    let content_types = CONTENT_TYPES.replace(
        r#"<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>"#,
        r#"<Default Extension="rels" ContentType="application/xml"/>"#,
    );
    let bytes = package(&content_types, ROOT_RELS, DOCUMENT_RELS);
    let report = reports(&bytes);
    let defect = Defect::ContentType(ContentTypeDefect::RelationshipPartType);
    assert_eq!(
        entries(&report),
        vec![
            (
                defect,
                CONTENT_TYPES_PART,
                Some("/_rels/.rels"),
                Some("application/xml")
            ),
            (
                defect,
                CONTENT_TYPES_PART,
                Some("/word/_rels/document.xml.rels"),
                Some("application/xml"),
            ),
        ]
    );

    let package = OpcPackage::from_bytes(&bytes).unwrap();
    assert_eq!(
        relationship_type(&package, "/word/document.xml", "rId1").as_deref(),
        Some(STYLES_TYPE)
    );
    assert_round_trips(&bytes);
}

/// LENIENCY_AUDIT 8.11: unknown elements, unknown attributes, character
/// data, and CDATA in either manifest are ignored and recorded.
#[test]
fn manifest_shape_deviations_are_ignored_and_recorded() {
    let content_types = CONTENT_TYPES
        .replace(
            r#"<Default Extension="xml" ContentType="application/xml"/>"#,
            r#"<Default Extension="xml" ContentType="application/xml" Producer="tool"/><x:Hint xmlns:x="urn:vendor"><Default Extension="bin" ContentType="application/x-hidden"/></x:Hint>"#,
        )
        .replace("</Types>", "stray</Types>");
    let document_rels = DOCUMENT_RELS
        .replace(
            r#"Target="styles.xml"/>"#,
            r#"Target="styles.xml" Extra="1"><![CDATA[note]]></Relationship>"#,
        )
        .replace(
            "</Relationships>",
            r#"<Wrapper><Relationship Id="rId9" Type="urn:hidden" Target="hidden.xml"/></Wrapper></Relationships>"#,
        );
    let bytes = package(&content_types, ROOT_RELS, &document_rels);
    let report = reports(&bytes);
    let content = |defect| Defect::ContentType(defect);
    let relationship = |defect| Defect::Relationship(defect);
    assert_eq!(
        entries(&report),
        vec![
            (
                content(ContentTypeDefect::UnknownAttribute),
                CONTENT_TYPES_PART,
                Some("xml"),
                Some("Producer"),
            ),
            (
                content(ContentTypeDefect::UnknownElement),
                CONTENT_TYPES_PART,
                None,
                Some("x:Hint"),
            ),
            (
                content(ContentTypeDefect::UnexpectedContent),
                CONTENT_TYPES_PART,
                None,
                Some("#text"),
            ),
            (
                relationship(RelationshipDefect::UnknownAttribute),
                "/word/document.xml",
                Some("rId1"),
                Some("Extra"),
            ),
            (
                relationship(RelationshipDefect::UnexpectedContent),
                "/word/document.xml",
                Some("rId1"),
                Some("#cdata-section"),
            ),
            (
                relationship(RelationshipDefect::UnknownElement),
                "/word/document.xml",
                None,
                Some("Wrapper"),
            ),
        ]
    );

    let package = OpcPackage::from_bytes(&bytes).unwrap();
    assert_eq!(
        relationship_type(&package, "/word/document.xml", "rId1").as_deref(),
        Some(STYLES_TYPE)
    );
    assert_eq!(
        relationship_type(&package, "/word/document.xml", "rId9"),
        None
    );
    assert_round_trips(&bytes);
}

/// LENIENCY_AUDIT 8.12: a `ContentType` or `Extension` outside the RFC 2616
/// grammar is kept verbatim (LibreOffice's opaque strings) and recorded.
#[test]
fn content_type_grammar_deviations_are_kept_verbatim_and_recorded() {
    const MAIN: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
    let raw = format!("{MAIN}; charset=utf-8");
    let content_types = CONTENT_TYPES
        .replace(
            &format!(r#"ContentType="{MAIN}""#),
            &format!(r#"ContentType="{raw}""#),
        )
        .replace(
            "</Types>",
            r#"<Default Extension="my ext" ContentType="text/plain"/></Types>"#,
        );
    let bytes = package(&content_types, ROOT_RELS, DOCUMENT_RELS);
    let report = reports(&bytes);
    assert_eq!(
        entries(&report),
        vec![
            (
                Defect::ContentType(ContentTypeDefect::InvalidMediaType),
                CONTENT_TYPES_PART,
                Some("/word/document.xml"),
                Some(raw.as_str()),
            ),
            (
                Defect::ContentType(ContentTypeDefect::InvalidExtension),
                CONTENT_TYPES_PART,
                Some("my ext"),
                Some("text/plain"),
            ),
        ]
    );

    let package = OpcPackage::from_bytes(&bytes).unwrap();
    assert_eq!(
        package
            .get_part(&PackURI::new("/word/document.xml").unwrap())
            .unwrap()
            .content_type(),
        raw
    );
    assert_round_trips(&bytes);

    // Empty content types and control characters stay fatal.
    for fatal in [
        CONTENT_TYPES.replace(&format!(r#"ContentType="{MAIN}""#), r#"ContentType="""#),
        CONTENT_TYPES.replace(
            &format!(r#"ContentType="{MAIN}""#),
            r#"ContentType="a/&#1;b""#,
        ),
    ] {
        assert!(
            OpcPackage::from_bytes(&package_bytes(&fatal)).is_err(),
            "{fatal}"
        );
        assert!(
            SourceBackedPackage::from_vec(package_bytes(&fatal)).is_err(),
            "{fatal}"
        );
    }
}

fn package_bytes(content_types: &str) -> Vec<u8> {
    package(content_types, ROOT_RELS, DOCUMENT_RELS)
}

/// Structural corruption stays fatal: a malformed relationships manifest,
/// a relationship without an Id, and an ill-formed content-types manifest
/// are still rejected.
#[test]
fn structural_manifest_corruption_stays_fatal() {
    let unclosed = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="urn:x" Target="styles.xml">"#;
    assert!(OpcPackage::from_bytes(&package(CONTENT_TYPES, ROOT_RELS, unclosed)).is_err());
    let no_id = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Type="urn:x" Target="styles.xml"/></Relationships>"#;
    assert!(OpcPackage::from_bytes(&package(CONTENT_TYPES, ROOT_RELS, no_id)).is_err());
    let ill_formed = r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"></Types>"#;
    assert!(OpcPackage::from_bytes(&package(ill_formed, ROOT_RELS, DOCUMENT_RELS)).is_err());
    // A DTD, an unbound prefix, and a wrong document element stay fatal in
    // both manifests.
    let doctype_rels = format!("<!DOCTYPE Relationships>{DOCUMENT_RELS}");
    let unbound_rels = DOCUMENT_RELS.replace("</Relationships>", "<x:Ext/></Relationships>");
    for rels in [doctype_rels.as_str(), unbound_rels.as_str()] {
        let bytes = package(CONTENT_TYPES, ROOT_RELS, rels);
        assert!(OpcPackage::from_bytes(&bytes).is_err(), "{rels}");
        assert!(SourceBackedPackage::from_vec(bytes).is_err(), "{rels}");
    }
    let doctype_types = format!("<!DOCTYPE Types>{CONTENT_TYPES}");
    let unbound_types = CONTENT_TYPES.replace("</Types>", "<x:Ext/></Types>");
    let wrong_root = CONTENT_TYPES
        .replace("<Types ", "<Kinds ")
        .replace("</Types>", "</Kinds>");
    for types in [doctype_types, unbound_types, wrong_root] {
        let bytes = package(&types, ROOT_RELS, DOCUMENT_RELS);
        assert!(OpcPackage::from_bytes(&bytes).is_err(), "{types}");
        assert!(SourceBackedPackage::from_vec(bytes).is_err(), "{types}");
    }
}

/// An edited package whose manifest is regenerated carries a source content
/// type admitted verbatim forward unchanged, while a content type the caller
/// authors keeps the MIME grammar check.
#[test]
fn edited_packages_keep_source_content_types_verbatim_but_check_authored_ones() {
    const MAIN: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
    let content_types = CONTENT_TYPES.replace(
        &format!(r#"ContentType="{MAIN}""#),
        &format!(r#"ContentType="{MAIN}; charset=utf-8""#),
    );
    let bytes = package(&content_types, ROOT_RELS, DOCUMENT_RELS);
    let add = |package: &mut OpcPackage, content_type: &str| {
        package
            .try_add_part(Box::new(litchi_opc::BlobPart::new(
                PackURI::new("/custom/item.bin").unwrap(),
                content_type.to_owned(),
                vec![1, 2, 3],
            )))
            .unwrap();
    };

    let mut edited = OpcPackage::from_bytes(&bytes).unwrap();
    add(&mut edited, "application/octet-stream");
    let mut saved = Vec::new();
    edited
        .to_stream(std::io::Cursor::new(&mut saved))
        .expect("a source content type admitted verbatim is carried forward");
    let reopened = OpcPackage::from_bytes(&saved).unwrap();
    assert_eq!(reopened.tolerance_report(), edited.tolerance_report());
    assert_eq!(
        reopened
            .get_part(&PackURI::new("/word/document.xml").unwrap())
            .unwrap()
            .content_type(),
        format!("{MAIN}; charset=utf-8")
    );

    let mut authored = OpcPackage::from_bytes(&bytes).unwrap();
    add(&mut authored, "application/octet-stream; x=1");
    assert!(
        authored
            .to_stream(std::io::Cursor::new(&mut Vec::new()))
            .is_err()
    );
}
