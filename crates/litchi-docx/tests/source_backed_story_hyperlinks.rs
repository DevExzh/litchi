use std::io;
use std::num::{NonZeroU64, NonZeroUsize};
use std::sync::Arc;
use std::sync::atomic::{AtomicU64, Ordering};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits as CoreLimits,
    OwnedSource, ReadAt, Resource,
};
use litchi_docx::package::StoryLimits;
use litchi_docx::source_backed;
use litchi_docx::story_hyperlinks::{Limits as StoryHyperlinkLimits, Mode, UnsupportedClass};
use litchi_docx::{Error, ReadLimits};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcPackage, PackURI, PackageWriter, Part, TargetMode};
use soapberry_zip::office::StreamingArchiveWriter;

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_W: &str = "http://purl.oclc.org/ooxml/wordprocessingml/main";
const STRICT_R: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const TRANSITIONAL_RELATIONSHIPS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
const STRICT_RELATIONSHIPS: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/";
const GLOSSARY_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/glossaryDocument";

fn story_xml(root: &str, label: &str) -> Vec<u8> {
    format!(
        r#"<w:{root} xmlns:w="{W}" xmlns:r="{R}"><w:p><w:hyperlink r:id="rLink"><w:r><w:t>{label}</w:t></w:r></w:hyperlink></w:p></w:{root}>"#
    )
    .into_bytes()
}

fn package_fixture(field: bool) -> OpcPackage {
    let main_xml = if field {
        format!(
            r#"<w:document xmlns:w="{W}" xmlns:r="{R}"><w:body><w:p><w:fldSimple w:instr="HYPERLINK https://field.invalid/"><w:r><w:t>field</w:t></w:r></w:fldSimple><w:hyperlink r:id="rMain"><w:r><w:t>main</w:t></w:r></w:hyperlink></w:p></w:body></w:document>"#
        )
    } else {
        format!(
            r#"<w:document xmlns:w="{W}" xmlns:r="{R}"><w:body><w:p><w:hyperlink r:id="rMain"><w:r><w:t>main</w:t></w:r></w:hyperlink></w:p></w:body></w:document>"#
        )
    };
    let mut package = OpcPackage::new();
    let mut main = BlobPart::new(
        PackURI::new("/word/document.xml").unwrap(),
        ct::WML_DOCUMENT_MAIN.to_owned(),
        main_xml.into_bytes(),
    );
    main.rels_mut()
        .try_add_relationship(
            rt::HYPERLINK.to_owned(),
            "https://shared.invalid/".to_owned(),
            "rMain".to_owned(),
            TargetMode::External,
        )
        .unwrap();
    let stories = [
        ("header1.xml", ct::WML_HEADER, rt::HEADER, "hdr", "header"),
        ("footer1.xml", ct::WML_FOOTER, rt::FOOTER, "ftr", "footer"),
        (
            "footnotes.xml",
            ct::WML_FOOTNOTES,
            rt::FOOTNOTES,
            "footnotes",
            "footnotes",
        ),
        (
            "endnotes.xml",
            ct::WML_ENDNOTES,
            rt::ENDNOTES,
            "endnotes",
            "endnotes",
        ),
        (
            "comments.xml",
            ct::WML_COMMENTS,
            rt::COMMENTS,
            "comments",
            "comments",
        ),
        (
            "glossary.xml",
            ct::WML_DOCUMENT_GLOSSARY,
            GLOSSARY_REL,
            "glossaryDocument",
            "glossary",
        ),
    ];
    for (name, content_type, relationship_type, root, label) in stories {
        main.rels_mut()
            .try_add_relationship(
                relationship_type.to_owned(),
                name.to_owned(),
                format!("r{name}"),
                TargetMode::Internal,
            )
            .unwrap();
        let mut story = BlobPart::new(
            PackURI::new(format!("/word/{name}")).unwrap(),
            content_type.to_owned(),
            story_xml(root, label),
        );
        story
            .rels_mut()
            .try_add_relationship(
                rt::HYPERLINK.to_owned(),
                "https://shared.invalid/".to_owned(),
                "rLink".to_owned(),
                TargetMode::External,
            )
            .unwrap();
        package.try_add_part(Box::new(story)).unwrap();
    }
    package.try_add_part(Box::new(main)).unwrap();
    package.relate_to("word/document.xml", rt::OFFICE_DOCUMENT);
    package
}

fn fixture(field: bool) -> Vec<u8> {
    PackageWriter::to_bytes(&package_fixture(field)).unwrap()
}

fn open(bytes: Vec<u8>) -> source_backed::Package {
    let source: Arc<dyn ReadAt> = Arc::new(OwnedSource::new(bytes));
    source_backed::Package::from_read_at(source).unwrap()
}

fn empty_inventory_fixture() -> Vec<u8> {
    let mut package = OpcPackage::from_bytes(&fixture(false)).unwrap();
    for (part_name, relationship_id) in [
        ("/word/document.xml", "rMain"),
        ("/word/header1.xml", "rLink"),
        ("/word/footer1.xml", "rLink"),
        ("/word/footnotes.xml", "rLink"),
        ("/word/endnotes.xml", "rLink"),
        ("/word/comments.xml", "rLink"),
        ("/word/glossary.xml", "rLink"),
    ] {
        let part = package
            .get_part_mut(&PackURI::new(part_name).unwrap())
            .unwrap();
        let opening = format!(r#"<w:hyperlink r:id="{relationship_id}">"#);
        let blob = std::str::from_utf8(part.blob())
            .unwrap()
            .replace(&opening, "")
            .replace("</w:hyperlink>", "")
            .into_bytes();
        part.set_blob(blob);
        part.rels_mut().remove(relationship_id);
    }
    PackageWriter::to_bytes(&package).unwrap()
}

struct VersionedSource {
    bytes: Vec<u8>,
    revision: AtomicU64,
}

impl ReadAt for VersionedSource {
    fn len(&self) -> io::Result<u64> {
        Ok(self.bytes.len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let offset = usize::try_from(offset)
            .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "offset"))?;
        if offset >= self.bytes.len() {
            return Ok(0);
        }
        let end = offset.saturating_add(output.len()).min(self.bytes.len());
        output[..end - offset].copy_from_slice(&self.bytes[offset..end]);
        Ok(end - offset)
    }

    fn version(&self) -> io::Result<litchi_core::SourceVersion> {
        Ok(litchi_core::SourceVersion::new(
            251,
            self.revision.load(Ordering::SeqCst),
        ))
    }
}

fn nested_owner_fixture() -> Vec<u8> {
    let mut package = OpcPackage::from_bytes(&fixture(false)).unwrap();
    package
        .get_part_mut(&PackURI::new("/word/header1.xml").unwrap())
        .unwrap()
        .rels_mut()
        .try_add_relationship(
            rt::FOOTER.to_owned(),
            "footer1.xml".to_owned(),
            "rNestedFooter".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
    PackageWriter::to_bytes(&package).unwrap()
}

fn mixed_dialect_fixture() -> Vec<u8> {
    let mut package = OpcPackage::from_bytes(&fixture(false)).unwrap();
    let header = package
        .get_part_mut(&PackURI::new("/word/header1.xml").unwrap())
        .unwrap();
    let mixed = std::str::from_utf8(header.blob())
        .unwrap()
        .replace(
            r#"<w:p>"#,
            r#"<w:p xmlns:s="http://purl.oclc.org/ooxml/wordprocessingml/main"><s:r><w:t>mixed</w:t></s:r>"#,
        );
    header.set_blob(mixed.into_bytes());
    PackageWriter::to_bytes(&package).unwrap()
}

fn prolog_fixture() -> Vec<u8> {
    let mut package = OpcPackage::from_bytes(&fixture(false)).unwrap();
    let document = package
        .get_part_mut(&PackURI::new("/word/document.xml").unwrap())
        .unwrap();
    let source = document.blob().to_vec();
    let mut with_prolog = br#"<?xml version="1.0"?><!--bounded prolog-->"#.to_vec();
    with_prolog.extend_from_slice(&source);
    document.set_blob(with_prolog);
    PackageWriter::to_bytes(&package).unwrap()
}

fn package_signature_fixture() -> Vec<u8> {
    let mut package = package_fixture(false);
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new("/_xmlsignatures/origin.sigs").unwrap(),
            ct::OPC_DIGITAL_SIGNATURE_ORIGIN.to_owned(),
            b"<origin/>".to_vec(),
        )))
        .unwrap();
    package
        .rels_mut()
        .try_add_relationship(
            rt::DIGITAL_SIGNATURE_ORIGIN.to_owned(),
            "_xmlsignatures/origin.sigs".to_owned(),
            "rSignature".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
    PackageWriter::to_bytes(&package).unwrap()
}

fn protected_fixture() -> Vec<u8> {
    let mut package = OpcPackage::from_bytes(&fixture(false)).unwrap();
    let main = package
        .get_part_mut(&PackURI::new("/word/document.xml").unwrap())
        .unwrap();
    main.rels_mut()
        .try_add_relationship(
            rt::SETTINGS.to_owned(),
            "settings.xml".to_owned(),
            "rSettings".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new("/word/settings.xml").unwrap(),
            ct::WML_SETTINGS.to_owned(),
            format!(
                r#"<w:settings xmlns:w="{W}"><w:documentProtection w:enforcement="1"/></w:settings>"#
            )
            .into_bytes(),
        )))
        .unwrap();
    PackageWriter::to_bytes(&package).unwrap()
}

fn delete_instruction_fixture() -> Vec<u8> {
    let mut package = OpcPackage::from_bytes(&fixture(false)).unwrap();
    let main = package
        .get_part_mut(&PackURI::new("/word/document.xml").unwrap())
        .unwrap();
    let source = std::str::from_utf8(main.blob()).unwrap().replace(
        r#"<w:hyperlink r:id="rMain">"#,
        r#"<w:delInstrText>HYPERLINK https://deleted.invalid/</w:delInstrText><w:hyperlink r:id="rMain">"#,
    );
    main.set_blob(source.into_bytes());
    PackageWriter::to_bytes(&package).unwrap()
}

fn mixed_transitional_hyperlink_fixture() -> Vec<u8> {
    let mut package = OpcPackage::from_bytes(&fixture(false)).unwrap();
    let header = package
        .get_part_mut(&PackURI::new("/word/header1.xml").unwrap())
        .unwrap();
    header.rels_mut().remove("rLink");
    header
        .rels_mut()
        .try_add_relationship(
            rt::STRICT_HYPERLINK.to_owned(),
            "https://shared.invalid/".to_owned(),
            "rLink".to_owned(),
            TargetMode::External,
        )
        .unwrap();
    PackageWriter::to_bytes(&package).unwrap()
}

fn strict_package_with_transitional_hyperlink_fixture() -> Vec<u8> {
    let mut package = OpcPackage::from_bytes(&fixture(false)).unwrap();
    let part_names: Vec<PackURI> = package
        .iter_parts()
        .map(|part| part.partname().clone())
        .collect();
    for part_name in part_names {
        let (blob, relationships): (Vec<u8>, Vec<(String, String, String, TargetMode)>) = {
            let part = package.get_part(&part_name).unwrap();
            let relationships = part
                .rels()
                .iter()
                .map(|relationship| {
                    (
                        relationship.r_id().to_owned(),
                        relationship.reltype().to_owned(),
                        relationship.target_ref().to_owned(),
                        relationship.target_mode(),
                    )
                })
                .collect();
            (part.blob().to_vec(), relationships)
        };
        let blob = String::from_utf8(blob)
            .unwrap()
            .replace(W, STRICT_W)
            .replace(R, STRICT_R)
            .into_bytes();
        let part = package.get_part_mut(&part_name).unwrap();
        part.set_blob(blob);
        for (id, reltype, target, mode) in relationships {
            part.rels_mut().remove(&id);
            let reltype = if let Some(suffix) = reltype.strip_prefix(TRANSITIONAL_RELATIONSHIPS) {
                format!("{STRICT_RELATIONSHIPS}{suffix}")
            } else {
                reltype
            };
            part.rels_mut()
                .try_add_relationship(reltype, target, id, mode)
                .unwrap();
        }
    }
    let root_relationships: Vec<(String, String, String, TargetMode)> = package
        .rels()
        .iter()
        .map(|relationship| {
            (
                relationship.r_id().to_owned(),
                relationship.reltype().to_owned(),
                relationship.target_ref().to_owned(),
                relationship.target_mode(),
            )
        })
        .collect();
    for (id, reltype, target, mode) in root_relationships {
        package.rels_mut().remove(&id);
        let reltype = if let Some(suffix) = reltype.strip_prefix(TRANSITIONAL_RELATIONSHIPS) {
            format!("{STRICT_RELATIONSHIPS}{suffix}")
        } else {
            reltype
        };
        package
            .rels_mut()
            .try_add_relationship(reltype, target, id, mode)
            .unwrap();
    }
    let header = package
        .get_part_mut(&PackURI::new("/word/header1.xml").unwrap())
        .unwrap();
    header.rels_mut().remove("rLink");
    header
        .rels_mut()
        .try_add_relationship(
            rt::HYPERLINK.to_owned(),
            "https://shared.invalid/".to_owned(),
            "rLink".to_owned(),
            TargetMode::External,
        )
        .unwrap();
    PackageWriter::to_bytes(&package).unwrap()
}

fn custom_inbound_owner_fixture() -> Vec<u8> {
    let mut package = OpcPackage::from_bytes(&fixture(false)).unwrap();
    let mut custom = BlobPart::new(
        PackURI::new("/custom/custom.xml").unwrap(),
        "application/xml".to_owned(),
        b"<custom/>".to_vec(),
    );
    custom
        .rels_mut()
        .try_add_relationship(
            rt::HEADER.to_owned(),
            "../word/header1.xml".to_owned(),
            "rLateHeader".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
    package.try_add_part(Box::new(custom)).unwrap();
    PackageWriter::to_bytes(&package).unwrap()
}

fn suspicious_non_part_fixture() -> Vec<u8> {
    suspicious_non_part_fixture_with_members(&[
        "word/vbaProject.bin",
        "word/embeddings/oleObject1.bin",
        "word/activeX/activeX1.bin",
        "_xmlsignatures/sig1.bin",
    ])
}

fn suspicious_non_part_fixture_with_members(names: &[&str]) -> Vec<u8> {
    let mut writer = StreamingArchiveWriter::new();
    writer
        .write_stored(
            "[Content_Types].xml",
            br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#,
        )
        .unwrap();
    writer
        .write_stored(
            "_rels/.rels",
            br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rDocument" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#,
        )
        .unwrap();
    writer
        .write_stored(
            "word/document.xml",
            format!(
                r#"<w:document xmlns:w="{W}" xmlns:r="{R}"><w:body><w:p><w:hyperlink r:id="rLink"><w:r><w:t>main</w:t></w:r></w:hyperlink></w:p></w:body></w:document>"#
            )
            .as_bytes(),
        )
        .unwrap();
    writer
        .write_stored(
            "word/_rels/document.xml.rels",
            br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://untyped.invalid/" TargetMode="External"/></Relationships>"#,
        )
        .unwrap();
    for name in names {
        writer
            .write_stored(name, b"untyped security surface")
            .unwrap();
    }
    writer.finish_to_bytes().unwrap()
}

fn child_relationship_limit_fixture(extra_relationships: usize) -> Vec<u8> {
    let mut package = OpcPackage::from_bytes(&fixture(false)).unwrap();
    for part_name in [
        "/word/header1.xml",
        "/word/footer1.xml",
        "/word/footnotes.xml",
        "/word/endnotes.xml",
        "/word/comments.xml",
        "/word/glossary.xml",
    ] {
        let story = package
            .get_part_mut(&PackURI::new(part_name).unwrap())
            .unwrap();
        for index in 0..extra_relationships {
            story
                .rels_mut()
                .try_add_relationship(
                    rt::HYPERLINK.to_owned(),
                    format!("https://child-limit-{index}.invalid/"),
                    format!("rChildLimit{index}"),
                    TargetMode::External,
                )
                .unwrap();
        }
    }
    PackageWriter::to_bytes(&package).unwrap()
}

fn with_story_limits(story: StoryLimits) -> StoryHyperlinkLimits {
    StoryHyperlinkLimits::default()
        .with_story_limits(story)
        .unwrap()
}

fn managed_context() -> (Budget, ExecutionContext) {
    let budget = Budget::root(
        "docx-story-hyperlink-managed-test",
        CoreLimits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
    );
    let (_, cancellation) = CancellationSource::pair();
    let execution_limits = ExecutionLimits::new(
        NonZeroUsize::MIN,
        NonZeroUsize::MIN,
        NonZeroU64::new(u64::MAX).unwrap(),
        0,
    )
    .unwrap();
    let context = ExecutionContext::new(budget.clone(), cancellation, execution_limits);
    (budget, context)
}

#[test]
fn inventories_all_relationship_owned_story_kinds_and_redacts_them() {
    let bytes = fixture(false);
    let package = open(bytes);
    let snapshot = package.story_hyperlinks_only_snapshot().unwrap();
    assert_eq!(snapshot.inventory().story_count(), 7);
    assert_eq!(snapshot.inventory().relationship_count(), 7);
    assert!(snapshot.inventory().is_complete());
    assert!(
        snapshot
            .relationships()
            .iter()
            .all(|relationship| relationship.wrapper_count() == 1)
    );

    let commit = package
        .plan_story_hyperlink_redaction(&["https://shared.invalid/"], Mode::Strict)
        .unwrap()
        .apply()
        .unwrap();
    assert_eq!(commit.effect_report().removed_relationships(), 7);
    assert_eq!(commit.effect_report().unwrapped_hyperlinks(), 7);
    let mut output = Vec::new();
    package
        .publish_story_hyperlink_redaction_to_stream(&mut output, &commit)
        .unwrap();
    let reopened = OpcPackage::from_bytes(&output).unwrap();
    for name in [
        "document.xml",
        "header1.xml",
        "footer1.xml",
        "footnotes.xml",
        "endnotes.xml",
        "comments.xml",
        "glossary.xml",
    ] {
        let part = reopened
            .get_part(&PackURI::new(format!("/word/{name}")).unwrap())
            .unwrap();
        assert!(part.rels().get("rLink").is_none(), "{name}");
        assert!(std::str::from_utf8(part.blob()).unwrap().contains("<w:t>"));
    }
}

#[test]
fn duplicate_target_selectors_are_deduplicated_and_missing_targets_are_typed() {
    let package = open(fixture(false));
    let plan = package
        .plan_story_hyperlink_redaction(
            &["https://shared.invalid/", "https://shared.invalid/"],
            Mode::Strict,
        )
        .unwrap();
    assert_eq!(plan.effect_report().selected_targets(), 1);
    assert_eq!(plan.effect_report().removed_relationships(), 7);

    assert!(matches!(
        package.plan_story_hyperlink_redaction(&["https://missing.invalid/"], Mode::Strict),
        Err(Error::InvalidFormat(message)) if message.contains("target is not present")
    ));
}

#[test]
fn empty_relationship_inventory_fails_closed_and_publishes_an_exact_noop() {
    let bytes = empty_inventory_fixture();
    let package = open(bytes.clone());
    let snapshot = package.story_hyperlinks_only_snapshot().unwrap();
    assert_eq!(snapshot.inventory().story_count(), 7);
    assert_eq!(snapshot.inventory().relationship_count(), 0);
    assert!(snapshot.inventory().is_complete());

    assert!(matches!(
        package.plan_story_hyperlink_redaction(&["https://missing.invalid/"], Mode::Strict),
        Err(Error::InvalidFormat(message)) if message.contains("target is not present")
    ));
    let commit = package
        .plan_story_hyperlink_redaction(&[], Mode::Strict)
        .unwrap()
        .apply()
        .unwrap();
    assert!(commit.effect_report().is_noop());
    let mut output = Vec::new();
    package
        .publish_story_hyperlink_redaction_to_stream(&mut output, &commit)
        .unwrap();
    assert_eq!(output, bytes);
}

#[test]
fn stale_and_foreign_commits_fail_before_output() {
    let source = Arc::new(VersionedSource {
        bytes: fixture(false),
        revision: AtomicU64::new(0),
    });
    let package = source_backed::Package::from_read_at(source.clone()).unwrap();
    let commit = package
        .plan_story_hyperlink_redaction(&["https://shared.invalid/"], Mode::Strict)
        .unwrap()
        .apply()
        .unwrap();
    source.revision.fetch_add(1, Ordering::SeqCst);
    let mut output = Vec::new();
    assert!(matches!(
        package.publish_story_hyperlink_redaction_to_stream(&mut output, &commit),
        Err(Error::Opc(litchi_opc::OpcError::SourceChanged { .. }))
            | Err(Error::ExternalHyperlinkRedactionConflict)
    ));
    assert!(output.is_empty());

    let bytes = fixture(false);
    let first = open(bytes.clone());
    let second = open(bytes);
    let commit = first
        .plan_story_hyperlink_redaction(&["https://shared.invalid/"], Mode::Strict)
        .unwrap()
        .apply()
        .unwrap();
    let mut output = Vec::new();
    assert!(matches!(
        second.publish_story_hyperlink_redaction_to_stream(&mut output, &commit),
        Err(Error::ExternalHyperlinkRedactionConflict)
    ));
    assert!(output.is_empty());
}

#[test]
fn redaction_limits_cover_selected_relationships_and_changed_stories() {
    let package = open(fixture(false));
    let selected_limit = StoryHyperlinkLimits::default()
        .with_max_selected_relationships(6)
        .unwrap();
    assert!(matches!(
        package.plan_story_hyperlink_redaction_with_limits(
            &["https://shared.invalid/"],
            Mode::Strict,
            selected_limit,
        ),
        Err(Error::ExternalHyperlinkRedactionLimit {
            resource: "selected story hyperlink relationships",
            maximum: 6,
            actual: 7,
        })
    ));

    let changed_story_limit = StoryHyperlinkLimits::default()
        .with_max_changed_stories(6)
        .unwrap();
    assert!(matches!(
        package.plan_story_hyperlink_redaction_with_limits(
            &["https://shared.invalid/"],
            Mode::Strict,
            changed_story_limit,
        ),
        Err(Error::ExternalHyperlinkRedactionLimit {
            resource: "changed story parts",
            maximum: 6,
            actual: 7,
        })
    ));
}

#[test]
fn exact_noop_copies_source_and_strict_fields_refuse_before_output() {
    let bytes = fixture(false);
    let package = open(bytes.clone());
    let commit = package
        .plan_story_hyperlink_redaction(&[], Mode::Strict)
        .unwrap()
        .apply()
        .unwrap();
    assert!(commit.effect_report().is_noop());
    let mut output = Vec::new();
    package
        .publish_story_hyperlink_redaction_to_stream(&mut output, &commit)
        .unwrap();
    assert_eq!(output, bytes);

    let field = open(fixture(true));
    let snapshot = field.story_hyperlinks_only_snapshot().unwrap();
    assert!(
        snapshot
            .diagnostics()
            .iter()
            .any(|diagnostic| diagnostic.class() == UnsupportedClass::FieldOrDde)
    );
    assert!(
        field
            .plan_story_hyperlink_redaction(&["https://shared.invalid/"], Mode::Strict)
            .is_err()
    );
}

#[test]
fn nested_story_owner_is_inventory_diagnostic_and_strict_refusal() {
    let package = open(nested_owner_fixture());
    let snapshot = package.story_hyperlinks_only_snapshot().unwrap();
    assert!(
        snapshot
            .diagnostics()
            .iter()
            .any(|diagnostic| diagnostic.class() == UnsupportedClass::UnknownOwner)
    );
    assert!(
        package
            .plan_story_hyperlink_redaction(&["https://shared.invalid/"], Mode::Strict)
            .is_err()
    );
}

#[test]
fn closure_capture_enforces_package_part_and_prolog_bounds_at_the_boundary() {
    let package = open(fixture(false));
    let story = StoryLimits {
        max_package_parts: 7,
        ..StoryLimits::default()
    };
    assert!(
        package
            .story_hyperlinks_only_snapshot_with_limits(with_story_limits(story))
            .is_ok()
    );
    let story = StoryLimits {
        max_package_parts: 6,
        ..StoryLimits::default()
    };
    assert!(
        package
            .story_hyperlinks_only_snapshot_with_limits(with_story_limits(story))
            .is_err()
    );

    let package = open(prolog_fixture());
    let story = StoryLimits {
        max_xml_prolog_events: 3,
        ..StoryLimits::default()
    };
    assert!(
        package
            .story_hyperlinks_only_snapshot_with_limits(with_story_limits(story))
            .is_ok()
    );
    let story = StoryLimits {
        max_xml_prolog_events: 2,
        ..StoryLimits::default()
    };
    assert!(
        package
            .story_hyperlinks_only_snapshot_with_limits(with_story_limits(story))
            .is_err()
    );
}

#[test]
fn closure_capture_enforces_exact_and_one_under_topology_limit() {
    let package = open(fixture(false));
    let mut low = 1usize;
    let mut high = StoryLimits::default().max_topology_bytes;
    while low < high {
        let middle = low + (high - low) / 2;
        let story = StoryLimits {
            max_topology_bytes: middle,
            ..StoryLimits::default()
        };
        if package
            .story_hyperlinks_only_snapshot_with_limits(with_story_limits(story))
            .is_ok()
        {
            high = middle;
        } else {
            low = middle + 1;
        }
    }
    let exact = StoryLimits {
        max_topology_bytes: low,
        ..StoryLimits::default()
    };
    assert!(
        package
            .story_hyperlinks_only_snapshot_with_limits(with_story_limits(exact))
            .is_ok()
    );
    let one_under = StoryLimits {
        max_topology_bytes: low - 1,
        ..exact
    };
    assert!(
        package
            .story_hyperlinks_only_snapshot_with_limits(with_story_limits(one_under))
            .is_err()
    );
}

#[test]
fn mixed_strict_child_qname_is_rejected_and_root_signature_is_diagnosed() {
    let mixed = open(mixed_dialect_fixture());
    assert!(mixed.story_hyperlinks_only_snapshot().is_err());

    let signed = open(package_signature_fixture());
    let snapshot = signed.story_hyperlinks_only_snapshot().unwrap();
    assert!(
        snapshot
            .diagnostics()
            .iter()
            .any(|diagnostic| diagnostic.class() == UnsupportedClass::Signature)
    );
    assert!(
        signed
            .plan_story_hyperlink_redaction(&["https://shared.invalid/"], Mode::Strict)
            .is_err()
    );
}

#[test]
fn mismatched_external_hyperlink_dialects_are_rejected_in_both_directions() {
    assert!(
        open(mixed_transitional_hyperlink_fixture())
            .story_hyperlinks_only_snapshot()
            .is_err()
    );
    assert!(
        open(strict_package_with_transitional_hyperlink_fixture())
            .story_hyperlinks_only_snapshot()
            .is_err()
    );
}

#[test]
fn late_custom_inbound_story_owner_is_diagnosed_and_refused() {
    let package = open(custom_inbound_owner_fixture());
    let snapshot = package.story_hyperlinks_only_snapshot().unwrap();
    assert!(
        snapshot
            .diagnostics()
            .iter()
            .any(|diagnostic| diagnostic.class() == UnsupportedClass::UnknownOwner)
    );
    assert!(
        package
            .plan_story_hyperlink_redaction(&["https://shared.invalid/"], Mode::Strict)
            .is_err()
    );
}

#[test]
fn foreign_identical_source_lineage_cannot_publish_a_noop_commit() {
    let bytes = fixture(false);
    let first = open(bytes.clone());
    let second = open(bytes);
    let commit = first
        .plan_story_hyperlink_redaction(&[], Mode::Strict)
        .unwrap()
        .apply()
        .unwrap();
    let mut output = Vec::new();
    assert!(
        second
            .publish_story_hyperlink_redaction_to_stream(&mut output, &commit)
            .is_err()
    );
    assert!(output.is_empty());
}

#[test]
fn empty_selector_copies_signed_source_exactly_despite_security_diagnostics() {
    let bytes = package_signature_fixture();
    let package = open(bytes.clone());
    let commit = package
        .plan_story_hyperlink_redaction(&[], Mode::Strict)
        .unwrap()
        .apply()
        .unwrap();
    assert!(commit.effect_report().is_noop());
    assert!(!commit.report().is_complete());
    let mut output = Vec::new();
    package
        .publish_story_hyperlink_redaction_to_stream(&mut output, &commit)
        .unwrap();
    assert_eq!(output, bytes);

    let bytes = protected_fixture();
    let package = open(bytes.clone());
    let commit = package
        .plan_story_hyperlink_redaction(&[], Mode::Strict)
        .unwrap()
        .apply()
        .unwrap();
    assert!(commit.effect_report().is_noop());
    assert!(!commit.report().is_complete());
    let mut output = Vec::new();
    package
        .publish_story_hyperlink_redaction_to_stream(&mut output, &commit)
        .unwrap();
    assert_eq!(output, bytes);
}

#[test]
fn signed_exact_noop_is_byte_exact_and_changed_plan_has_empty_output() {
    let bytes = package_signature_fixture();
    let package = open(bytes.clone());
    let commit = package
        .plan_story_hyperlink_redaction(&[], Mode::Strict)
        .unwrap()
        .apply()
        .unwrap();
    let mut output = Vec::new();
    package
        .publish_story_hyperlink_redaction_to_stream(&mut output, &commit)
        .unwrap();
    assert_eq!(output, bytes);

    let package = open(package_signature_fixture());
    let output = Vec::<u8>::new();
    assert!(matches!(
        package.plan_story_hyperlink_redaction(&["https://shared.invalid/"], Mode::Strict),
        Err(Error::UnsafeEdit { .. })
    ));
    assert!(output.is_empty());
}

#[test]
fn untyped_non_part_security_surfaces_are_diagnosed_and_refused() {
    let package = open(suspicious_non_part_fixture());
    let snapshot = package.story_hyperlinks_only_snapshot().unwrap();
    assert!(
        snapshot
            .diagnostics()
            .iter()
            .any(|diagnostic| diagnostic.class() == UnsupportedClass::Macro)
    );
    assert!(
        snapshot
            .diagnostics()
            .iter()
            .any(|diagnostic| diagnostic.class() == UnsupportedClass::EmbeddedContent)
    );
    assert!(
        snapshot
            .diagnostics()
            .iter()
            .any(|diagnostic| diagnostic.class() == UnsupportedClass::Signature)
    );
    assert!(
        package
            .plan_story_hyperlink_redaction(&["https://untyped.invalid/"], Mode::Strict)
            .is_err()
    );
}

#[test]
fn root_level_embeddings_non_part_is_diagnosed_and_refused() {
    let package = open(suspicious_non_part_fixture_with_members(&[
        "embeddings/oleObject1.bin",
    ]));
    let snapshot = package.story_hyperlinks_only_snapshot().unwrap();
    assert_eq!(
        snapshot
            .diagnostics()
            .iter()
            .filter(|diagnostic| diagnostic.class() == UnsupportedClass::EmbeddedContent)
            .map(|diagnostic| diagnostic.count())
            .sum::<usize>(),
        1
    );
    assert!(
        package
            .plan_story_hyperlink_redaction(&["https://untyped.invalid/"], Mode::Strict)
            .is_err()
    );
}

#[test]
fn root_level_activex_non_part_is_diagnosed_and_refused() {
    let package = open(suspicious_non_part_fixture_with_members(&[
        "activex/activeX1.bin",
    ]));
    let snapshot = package.story_hyperlinks_only_snapshot().unwrap();
    assert_eq!(
        snapshot
            .diagnostics()
            .iter()
            .filter(|diagnostic| diagnostic.class() == UnsupportedClass::EmbeddedContent)
            .map(|diagnostic| diagnostic.count())
            .sum::<usize>(),
        1
    );
    assert!(
        package
            .plan_story_hyperlink_redaction(&["https://untyped.invalid/"], Mode::Strict)
            .is_err()
    );
}

#[test]
fn delete_instruction_text_is_an_active_field_diagnostic() {
    let package = open(delete_instruction_fixture());
    let snapshot = package.story_hyperlinks_only_snapshot().unwrap();
    assert!(
        snapshot
            .diagnostics()
            .iter()
            .any(|diagnostic| diagnostic.class() == UnsupportedClass::FieldOrDde)
    );
}

#[test]
fn relationship_limits_are_checked_before_main_payload_reads() {
    let package = open(fixture(false));
    let before = package.cache_diagnostics();
    let limits = StoryLimits {
        max_relationships_per_owner: 6,
        ..StoryLimits::default()
    };
    assert!(
        package
            .story_hyperlinks_only_snapshot_with_limits(with_story_limits(limits))
            .is_err()
    );
    let after = package.cache_diagnostics();
    assert_eq!(before.cold_loads, after.cold_loads);
    assert_eq!(before.successful_loads, after.successful_loads);

    let exact = StoryLimits {
        max_relationships_per_owner: 7,
        max_total_relationships: 13,
        ..StoryLimits::default()
    };
    assert!(
        package
            .story_hyperlinks_only_snapshot_with_limits(with_story_limits(exact))
            .is_ok()
    );
}

#[test]
fn child_relationship_owner_limit_is_preflighted_before_payload_reads() {
    let package = open(child_relationship_limit_fixture(7));
    let before = package.cache_diagnostics();
    let one_under = StoryLimits {
        max_relationships_per_owner: 7,
        ..StoryLimits::default()
    };
    assert!(
        package
            .story_hyperlinks_only_snapshot_with_limits(with_story_limits(one_under))
            .is_err()
    );
    let after = package.cache_diagnostics();
    assert_eq!(after.cold_loads, before.cold_loads + 1);
    assert_eq!(after.successful_loads, before.successful_loads + 1);

    let exact = StoryLimits {
        max_relationships_per_owner: 8,
        ..StoryLimits::default()
    };
    assert!(
        package
            .story_hyperlinks_only_snapshot_with_limits(with_story_limits(exact))
            .is_ok()
    );
}

#[test]
fn child_relationship_total_is_preflighted_before_the_offending_payload_read() {
    let package = open(fixture(false));
    let before = package.cache_diagnostics();
    let one_under = StoryLimits {
        max_relationships_per_owner: 7,
        max_total_relationships: 12,
        ..StoryLimits::default()
    };
    assert!(
        package
            .story_hyperlinks_only_snapshot_with_limits(with_story_limits(one_under))
            .is_err()
    );
    let after = package.cache_diagnostics();
    assert_eq!(after.cold_loads, before.cold_loads + 6);
    assert_eq!(after.successful_loads, before.successful_loads + 6);

    let exact = StoryLimits {
        max_relationships_per_owner: 7,
        max_total_relationships: 13,
        ..StoryLimits::default()
    };
    assert!(
        package
            .story_hyperlinks_only_snapshot_with_limits(with_story_limits(exact))
            .is_ok()
    );
}

#[test]
fn managed_story_snapshot_retains_budgeted_part_data() {
    let (budget, context) = managed_context();
    let package = source_backed::Package::from_read_at_with_execution_context(
        Arc::new(OwnedSource::new(fixture(false))),
        ReadLimits::default(),
        context,
    )
    .unwrap();
    let snapshot = package.story_hyperlinks_only_snapshot().unwrap();
    assert_eq!(snapshot.inventory().story_count(), 7);
    assert!(budget.used(Resource::Memory) > 0);
    drop(snapshot);
    drop(package);
    assert_eq!(budget.used(Resource::Memory), 0);
}
