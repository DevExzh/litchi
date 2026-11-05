use std::io::{self, Write};
use std::sync::Arc;

use litchi_core::{OwnedSource, Position, ReadAt};
use litchi_docx::sanitize::Limits;
use litchi_docx::{Error, Package, source_backed};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcError, OpcPackage, PackURI, PackageWriter, Part};

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const MAIN: &str = "/word/document.xml";
const COMMENTS: &str = "/word/comments.xml";
const VBA: &str = "/word/vbaProject.bin";

#[derive(Default)]
struct FixtureOptions {
    relationships: Vec<(&'static str, &'static str, &'static str, bool)>,
    settings: Option<Vec<u8>>,
    settings_relationship: Option<&'static str>,
    comments: Option<Vec<u8>>,
    vba: Option<Vec<u8>>,
    signed: bool,
}

fn fixture(document: Vec<u8>, options: FixtureOptions) -> Vec<u8> {
    let mut package = OpcPackage::new();
    let mut main = BlobPart::new(
        PackURI::new(MAIN).unwrap(),
        ct::WML_DOCUMENT_MAIN.to_owned(),
        document,
    );
    for (id, kind, target, external) in options.relationships {
        main.rels_mut()
            .try_add_relationship(
                kind.to_owned(),
                target.to_owned(),
                id.to_owned(),
                if external {
                    litchi_opc::TargetMode::External
                } else {
                    litchi_opc::TargetMode::Internal
                },
            )
            .unwrap();
    }
    if options.settings.is_some() {
        main.rels_mut()
            .try_add_relationship(
                options
                    .settings_relationship
                    .unwrap_or(rt::SETTINGS)
                    .to_owned(),
                "settings.xml".to_owned(),
                "rSettings".to_owned(),
                litchi_opc::TargetMode::Internal,
            )
            .unwrap();
    }
    if options.comments.is_some() {
        main.rels_mut()
            .try_add_relationship(
                rt::COMMENTS.to_owned(),
                "comments.xml".to_owned(),
                "rComments".to_owned(),
                litchi_opc::TargetMode::Internal,
            )
            .unwrap();
    }
    if options.vba.is_some() {
        main.rels_mut()
            .try_add_relationship(
                rt::VBA_PROJECT.to_owned(),
                "vbaProject.bin".to_owned(),
                "rVba".to_owned(),
                litchi_opc::TargetMode::Internal,
            )
            .unwrap();
    }
    package.try_add_part(Box::new(main)).unwrap();

    if let Some(settings) = options.settings {
        package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new("/word/settings.xml").unwrap(),
                ct::WML_SETTINGS.to_owned(),
                settings,
            )))
            .unwrap();
    }
    if let Some(comments) = options.comments {
        package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new(COMMENTS).unwrap(),
                ct::WML_COMMENTS.to_owned(),
                comments,
            )))
            .unwrap();
    }
    if let Some(vba) = options.vba {
        package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new(VBA).unwrap(),
                ct::OFC_VBA_PROJECT.to_owned(),
                vba,
            )))
            .unwrap();
    }
    package.relate_to("word/document.xml", rt::OFFICE_DOCUMENT);
    if options.signed {
        package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new("/_xmlsignatures/origin.sigs").unwrap(),
                ct::OPC_DIGITAL_SIGNATURE_ORIGIN.to_owned(),
                b"<origin/>".to_vec(),
            )))
            .unwrap();
        package.relate_to("_xmlsignatures/origin.sigs", rt::DIGITAL_SIGNATURE_ORIGIN);
    }
    PackageWriter::to_bytes(&package).unwrap()
}

fn open_source(bytes: Vec<u8>) -> source_backed::Package {
    let source: Arc<dyn ReadAt> = Arc::new(OwnedSource::new(bytes));
    source_backed::Package::from_read_at(source).unwrap()
}

fn external_document(text: &str) -> Vec<u8> {
    format!(
        r#"<w:document xmlns:w="{W}" xmlns:r="{R}"><w:body><w:p><w:hyperlink r:id="rLink"><w:r><w:t>{text}</w:t></w:r></w:hyperlink></w:p></w:body></w:document>"#
    )
    .into_bytes()
}

fn external_options(target: &'static str) -> FixtureOptions {
    FixtureOptions {
        relationships: vec![("rLink", rt::HYPERLINK, target, true)],
        ..FixtureOptions::default()
    }
}

#[test]
fn plan_is_non_mutating_and_detaches_first_middle_last_while_retaining_text() {
    let document = format!(
        r#"<?xml version="1.0"?><w:document xmlns:w="{W}" xmlns:r="{R}" xmlns:x="urn:litchi:opaque"><w:body><w:p><w:hyperlink r:id="rOne"><w:r><w:t>first</w:t></w:r></w:hyperlink><w:r><w:t>|</w:t></w:r><w:hyperlink r:id="rTwo"><w:r><w:t>middle</w:t></w:r></w:hyperlink><w:hyperlink w:anchor="local"><w:r><w:t>|local|</w:t></w:r></w:hyperlink><w:commentRangeStart w:id="8"/><w:hyperlink r:id="rOne"><w:r><w:t>last</w:t></w:r></w:hyperlink><w:commentRangeEnd w:id="8"/><w:r><w:commentReference w:id="8"/></w:r><x:opaque x:value="keep"><![CDATA[a < b]]></x:opaque></w:p></w:body></w:document>"#
    )
    .into_bytes();
    let comments = format!(
        r#"<w:comments xmlns:w="{W}"><w:comment w:id="8" w:author="Author"><w:p><w:r><w:t>comment</w:t></w:r></w:p></w:comment></w:comments>"#
    )
    .into_bytes();
    let vba = b"inert macro bytes\0\xFF".to_vec();
    let options = FixtureOptions {
        relationships: vec![
            ("rOne", rt::HYPERLINK, "https://one.invalid/", true),
            ("rTwo", rt::HYPERLINK, "https://two.invalid/", true),
        ],
        comments: Some(comments.clone()),
        vba: Some(vba.clone()),
        ..FixtureOptions::default()
    };
    let bytes = fixture(document, options);
    let package = open_source(bytes);
    let source = package.external_hyperlink_sanitization_snapshot().unwrap();
    let source_xml = source.xml_bytes().to_vec();
    let plan = source.plan();

    assert_eq!(source.xml_bytes(), source_xml);
    assert_eq!(plan.effect_report().detached_hyperlinks(), 3);
    assert_eq!(plan.effect_report().referenced_external_relationships(), 2);
    assert_eq!(plan.effect_report().retained_relationships(), 2);

    let patch = plan.apply().unwrap();
    let sanitized_snapshot = patch.apply(&source).unwrap();
    assert_eq!(source.xml_bytes(), source_xml);
    assert_eq!(sanitized_snapshot.external_hyperlink_count(), 0);
    let sanitized = std::str::from_utf8(sanitized_snapshot.xml_bytes()).unwrap();
    assert!(!sanitized.contains("r:id=\"rOne\""));
    assert!(!sanitized.contains("r:id=\"rTwo\""));
    assert!(sanitized.contains("<w:hyperlink w:anchor=\"local\">"));
    assert!(sanitized.contains("<w:commentRangeStart w:id=\"8\"/>"));
    assert!(sanitized.contains("<x:opaque x:value=\"keep\"><![CDATA[a < b]]></x:opaque>"));
    let parsed = litchi_docx::document::raw::Tree::from_xml(
        sanitized_snapshot.xml_bytes().to_vec(),
        litchi_docx::document::raw::Limits::default(),
    )
    .unwrap();
    assert_eq!(
        parsed
            .projected()
            .paragraph(Position::new(0))
            .unwrap()
            .text()
            .unwrap(),
        "first|middle|local|last"
    );

    let package = open_source(bytes_for_same_fixture(
        &source_xml,
        comments.clone(),
        vba.clone(),
    ));
    let fresh_commit = package
        .edit_external_hyperlink_detachment()
        .unwrap()
        .commit()
        .unwrap();
    let mut output = Vec::new();
    fresh_commit.package().write_to(&mut output).unwrap();
    let reopened = OpcPackage::from_bytes(&output, litchi_opc::SourceLimits::default()).unwrap();
    let main = reopened.main_document_part().unwrap();
    assert_eq!(
        main.rels().get("rOne").unwrap().target_ref(),
        "https://one.invalid/"
    );
    assert_eq!(
        main.rels().get("rTwo").unwrap().target_ref(),
        "https://two.invalid/"
    );
    assert_eq!(
        reopened
            .get_part(&PackURI::new(COMMENTS).unwrap())
            .unwrap()
            .blob(),
        comments
    );
    assert_eq!(
        reopened
            .get_part(&PackURI::new(VBA).unwrap())
            .unwrap()
            .blob(),
        vba
    );
    let reopened = Package::from_opc_package(reopened).unwrap();
    assert_eq!(
        reopened.document().unwrap().text().unwrap(),
        "first|middle|local|last"
    );
}

fn bytes_for_same_fixture(document: &[u8], comments: Vec<u8>, vba: Vec<u8>) -> Vec<u8> {
    fixture(
        document.to_vec(),
        FixtureOptions {
            relationships: vec![
                ("rOne", rt::HYPERLINK, "https://one.invalid/", true),
                ("rTwo", rt::HYPERLINK, "https://two.invalid/", true),
            ],
            comments: Some(comments),
            vba: Some(vba),
            ..FixtureOptions::default()
        },
    )
}

#[test]
fn exact_noop_shares_snapshot_and_reproduces_signed_source_bytes() {
    let document = format!(
        r#"<w:document xmlns:w="{W}"><w:body><w:p><w:hyperlink w:anchor="local"><w:r><w:t>local</w:t></w:r></w:hyperlink></w:p></w:body></w:document>"#
    )
    .into_bytes();
    let bytes = fixture(
        document,
        FixtureOptions {
            signed: true,
            ..FixtureOptions::default()
        },
    );
    let package = open_source(bytes.clone());
    let source = package.external_hyperlink_sanitization_snapshot().unwrap();
    let patch = source.plan().apply().unwrap();
    let applied = patch.apply(&source).unwrap();
    assert!(patch.effect_report().is_noop());
    assert!(patch.is_noop());
    assert!(source.shares_xml_allocation_with(&applied));

    let mut output = Vec::new();
    package
        .edit_external_hyperlink_detachment()
        .unwrap()
        .commit()
        .unwrap()
        .package()
        .write_to(&mut output)
        .unwrap();
    assert_eq!(output, bytes);
}

#[test]
fn patch_inverse_is_exact_and_foreign_xml_or_relationships_conflict() {
    let first_bytes = fixture(
        external_document("visible"),
        external_options("https://one.invalid/"),
    );
    let first = open_source(first_bytes)
        .external_hyperlink_sanitization_snapshot()
        .unwrap();
    let patch = first.plan().apply().unwrap();
    let applied = patch.apply(&first).unwrap();
    assert_eq!(
        applied.xml_bytes(),
        patch.apply(&first).unwrap().xml_bytes()
    );
    let restored = patch.inverse().apply(&applied).unwrap();
    assert_eq!(restored.xml_bytes(), first.xml_bytes());
    let inverse_report = patch.inverse().effect_report();
    assert_eq!(inverse_report.detached_hyperlinks(), 0);
    assert_eq!(inverse_report.restored_hyperlinks(), 1);
    assert!(!inverse_report.is_noop());

    let foreign_xml = open_source(fixture(
        external_document("different"),
        external_options("https://one.invalid/"),
    ))
    .external_hyperlink_sanitization_snapshot()
    .unwrap();
    assert!(matches!(
        patch.apply(&foreign_xml),
        Err(Error::PatchConflict { .. })
    ));

    let foreign_relationship = open_source(fixture(
        external_document("visible"),
        external_options("https://different.invalid/"),
    ))
    .external_hyperlink_sanitization_snapshot()
    .unwrap();
    assert!(matches!(
        patch.apply(&foreign_relationship),
        Err(Error::PatchConflict { .. })
    ));
}

#[test]
fn canonical_replay_rejects_an_unrelated_part_mutation() {
    let document = external_document("visible");
    let original_comments = format!(
        r#"<w:comments xmlns:w="{W}"><w:comment w:id="1"><w:p><w:r><w:t>original</w:t></w:r></w:p></w:comment></w:comments>"#
    )
    .into_bytes();
    let original = fixture(
        document.clone(),
        FixtureOptions {
            relationships: external_options("https://one.invalid/").relationships,
            comments: Some(original_comments),
            ..FixtureOptions::default()
        },
    );
    let commit = open_source(original)
        .edit_external_hyperlink_detachment()
        .unwrap()
        .commit()
        .unwrap();

    let changed_comments = format!(
        r#"<w:comments xmlns:w="{W}"><w:comment w:id="1"><w:p><w:r><w:t>changed</w:t></w:r></w:p></w:comment></w:comments>"#
    )
    .into_bytes();
    let changed = fixture(
        document,
        FixtureOptions {
            relationships: external_options("https://one.invalid/").relationships,
            comments: Some(changed_comments),
            ..FixtureOptions::default()
        },
    );
    assert!(matches!(
        open_source(changed).apply_external_hyperlink_detachment(commit.patch()),
        Err(Error::PatchConflict { .. })
    ));
}

#[test]
fn protected_malformed_active_and_over_limit_sources_fail_closed() {
    let protected = fixture(
        external_document("protected"),
        FixtureOptions {
            relationships: external_options("https://one.invalid/").relationships,
            settings: Some(
                format!(
                    r#"<w:settings xmlns:w="{W}"><w:documentProtection w:edit="readOnly" w:enforcement="1"/></w:settings>"#
                )
                .into_bytes(),
            ),
            ..FixtureOptions::default()
        },
    );
    assert!(matches!(
        open_source(protected).edit_external_hyperlink_detachment(),
        Err(Error::Refused(_))
    ));

    let write_protected = fixture(
        external_document("write protected"),
        FixtureOptions {
            relationships: external_options("https://one.invalid/").relationships,
            settings: Some(
                format!(r#"<w:settings xmlns:w="{W}"><w:writeProtection/></w:settings>"#)
                    .into_bytes(),
            ),
            ..FixtureOptions::default()
        },
    );
    assert!(matches!(
        open_source(write_protected).edit_external_hyperlink_detachment(),
        Err(Error::Refused(_))
    ));

    const STRICT_W: &str = "http://purl.oclc.org/ooxml/wordprocessingml/main";
    const STRICT_SETTINGS: &str =
        "http://purl.oclc.org/ooxml/officeDocument/relationships/settings";
    let strict_write_protected = fixture(
        external_document("strict write protected"),
        FixtureOptions {
            relationships: external_options("https://one.invalid/").relationships,
            settings: Some(
                format!(r#"<s:settings xmlns:s="{STRICT_W}"><s:writeProtection/></s:settings>"#)
                    .into_bytes(),
            ),
            settings_relationship: Some(STRICT_SETTINGS),
            ..FixtureOptions::default()
        },
    );
    assert!(matches!(
        open_source(strict_write_protected).edit_external_hyperlink_detachment(),
        Err(Error::Refused(_))
    ));

    let mce = format!(
        r#"<w:document xmlns:w="{W}" xmlns:r="{R}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:unsupported"><w:body><mc:AlternateContent><mc:Choice Requires="x"><w:p><w:hyperlink r:id="rLink"><w:r><w:t>choice</w:t></w:r></w:hyperlink></w:p></mc:Choice><mc:Fallback><w:p><w:r><w:t>fallback</w:t></w:r></w:p></mc:Fallback></mc:AlternateContent></w:body></w:document>"#
    )
    .into_bytes();
    assert!(matches!(
        open_source(fixture(mce, external_options("https://one.invalid/"),))
            .edit_external_hyperlink_detachment(),
        Err(Error::Refused(_))
    ));

    let missing = fixture(external_document("missing"), FixtureOptions::default());
    assert!(
        open_source(missing)
            .edit_external_hyperlink_detachment()
            .is_err()
    );

    let nested = format!(
        r#"<w:document xmlns:w="{W}" xmlns:r="{R}"><w:body><w:p><w:hyperlink r:id="rLink"><w:hyperlink r:id="rLink"/></w:hyperlink></w:p></w:body></w:document>"#
    )
    .into_bytes();
    assert!(
        open_source(fixture(nested, external_options("https://one.invalid/"),))
            .edit_external_hyperlink_detachment()
            .is_err()
    );

    let three = format!(
        r#"<w:document xmlns:w="{W}" xmlns:r="{R}"><w:body><w:p><w:hyperlink r:id="rLink"/><w:hyperlink r:id="rLink"/><w:hyperlink r:id="rLink"/></w:p></w:body></w:document>"#
    )
    .into_bytes();
    let package = open_source(fixture(three, external_options("https://one.invalid/")));
    assert_eq!(
        package
            .with_semantic_limits(
                litchi_docx::SemanticLimits::default()
                    .hyperlinks(Limits::default().with_max_external_hyperlinks(3).unwrap())
            )
            .edit_external_hyperlink_detachment()
            .unwrap()
            .effect_report()
            .detached_hyperlinks(),
        3
    );
    assert!(matches!(
        package
            .with_semantic_limits(
                litchi_docx::SemanticLimits::default()
                    .hyperlinks(Limits::default().with_max_external_hyperlinks(2).unwrap())
            )
            .edit_external_hyperlink_detachment(),
        Err(Error::ExternalHyperlinkDetachmentLimit {
            kind: litchi_docx::sanitize::LimitKind::ExternalHyperlinks,
            maximum: 2,
            actual: 3,
        })
    ));
}

#[test]
fn main_declared_size_is_preflighted_before_raw_hyperlink_payload_reads() {
    let package = open_source(fixture(
        external_document("preflight"),
        external_options("https://one.invalid/"),
    ));
    let before = package.cache_diagnostics();
    let limits = Limits::default().with_max_document_bytes(1).unwrap();
    let bounded =
        package.with_semantic_limits(litchi_docx::SemanticLimits::default().hyperlinks(limits));
    assert!(matches!(
        bounded.external_hyperlink_sanitization_snapshot(),
        Err(Error::ExternalHyperlinkDetachmentLimit {
            kind: litchi_docx::sanitize::LimitKind::MainDocumentXmlBytes,
            maximum: 1,
            actual,
        }) if actual > 1
    ));
    assert!(matches!(
        bounded.external_hyperlink_redaction_snapshot(),
        Err(Error::ExternalHyperlinkDetachmentLimit {
            kind: litchi_docx::sanitize::LimitKind::MainDocumentXmlBytes,
            maximum: 1,
            actual,
        }) if actual > 1
    ));
    let after = package.cache_diagnostics();
    assert_eq!(after.cold_loads, before.cold_loads);
    assert_eq!(after.successful_loads, before.successful_loads);
}

#[test]
fn hyperlink_replay_uses_recipient_limits_for_restored_owner() {
    let document = format!(
        r#"<w:document xmlns:w="{W}" xmlns:r="{R}"><w:body><w:p><w:hyperlink r:id="rLink"/><w:hyperlink r:id="rLink"/><w:hyperlink r:id="rLink"/></w:p></w:body></w:document>"#
    )
    .into_bytes();
    let source = fixture(document, external_options("https://one.invalid/"));
    let producer = open_source(source.clone()).with_semantic_limits(
        litchi_docx::SemanticLimits::default()
            .hyperlinks(Limits::default().with_max_external_hyperlinks(3).unwrap()),
    );
    let owner_commit = producer
        .edit_external_hyperlink_detachment()
        .unwrap()
        .commit()
        .unwrap();

    let recipient = open_source(source).with_semantic_limits(
        litchi_docx::SemanticLimits::default()
            .hyperlinks(Limits::default().with_max_external_hyperlinks(2).unwrap()),
    );
    assert!(matches!(
        recipient.apply_external_hyperlink_detachment(owner_commit.patch()),
        Err(Error::ExternalHyperlinkDetachmentLimit {
            kind: litchi_docx::sanitize::LimitKind::ExternalHyperlinks,
            maximum: 2,
            actual: 3,
        })
    ));
}

#[test]
fn wrapper_scoped_namespaces_xml_attributes_and_unknown_attributes_refuse_detachment() {
    let wrapper_scoped_prefix = format!(
        r#"<w:document xmlns:w="{W}" xmlns:r="{R}"><w:body><w:p><w:hyperlink r:id="rLink" xmlns:x="urn:wrapper-only"><w:r><w:t>visible</w:t></w:r><x:opaque/></w:hyperlink></w:p></w:body></w:document>"#
    )
    .into_bytes();
    assert!(matches!(
        open_source(fixture(
            wrapper_scoped_prefix,
            external_options("https://one.invalid/"),
        ))
        .edit_external_hyperlink_detachment(),
        Err(Error::Refused(_))
    ));

    let wrapper_scoped_relationship_prefix = format!(
        r#"<w:document xmlns:w="{W}"><w:body><w:p><w:hyperlink xmlns:r="{R}" r:id="rLink"><w:r><w:t>visible</w:t></w:r></w:hyperlink></w:p></w:body></w:document>"#
    )
    .into_bytes();
    assert!(matches!(
        open_source(fixture(
            wrapper_scoped_relationship_prefix,
            external_options("https://one.invalid/"),
        ))
        .edit_external_hyperlink_detachment(),
        Err(Error::Refused(_))
    ));

    let xml_scope = format!(
        r#"<w:document xmlns:w="{W}" xmlns:r="{R}"><w:body><w:p><w:hyperlink r:id="rLink" xml:space="preserve"><w:r><w:t> visible </w:t></w:r></w:hyperlink></w:p></w:body></w:document>"#
    )
    .into_bytes();
    assert!(matches!(
        open_source(fixture(xml_scope, external_options("https://one.invalid/"),))
            .edit_external_hyperlink_detachment(),
        Err(Error::Refused(_))
    ));

    let unknown = format!(
        r#"<w:document xmlns:w="{W}" xmlns:r="{R}" xmlns:x="urn:unknown"><w:body><w:p><w:hyperlink r:id="rLink" x:policy="keep"><w:r><w:t>visible</w:t></w:r></w:hyperlink></w:p></w:body></w:document>"#
    )
    .into_bytes();
    assert!(matches!(
        open_source(fixture(unknown, external_options("https://one.invalid/"),))
            .edit_external_hyperlink_detachment(),
        Err(Error::Refused(_))
    ));

    let known = format!(
        r#"<w:document xmlns:w="{W}" xmlns:r="{R}"><w:body><w:p><w:hyperlink r:id="rLink" w:tooltip="drop with wrapper" w:history="1"><w:r><w:t>visible</w:t></w:r></w:hyperlink></w:p></w:body></w:document>"#
    )
    .into_bytes();
    let package = open_source(fixture(known, external_options("https://one.invalid/")));
    let edit = package.edit_external_hyperlink_detachment().unwrap();
    assert_eq!(edit.effect_report().detached_hyperlinks(), 1);
    let commit = edit.commit().unwrap();
    let mut output = Vec::new();
    commit.package().write_to(&mut output).unwrap();
    let snapshot = open_source(output)
        .external_hyperlink_sanitization_snapshot()
        .unwrap();
    assert!(
        !std::str::from_utf8(snapshot.xml_bytes())
            .unwrap()
            .contains("tooltip")
    );
}

#[test]
fn changed_signed_publication_and_partial_sink_are_typed_failures() {
    let signed = fixture(
        external_document("signed"),
        FixtureOptions {
            relationships: external_options("https://one.invalid/").relationships,
            signed: true,
            ..FixtureOptions::default()
        },
    );
    let package = open_source(signed);
    let result = package
        .edit_external_hyperlink_detachment()
        .unwrap()
        .commit();
    assert!(matches!(
        result,
        Err(Error::Opc(OpcError::SignedSourceRequiresExplicitPolicy))
    ));

    let bytes = fixture(
        external_document("sink"),
        external_options("https://one.invalid/"),
    );
    let package = open_source(bytes);
    let commit = package
        .edit_external_hyperlink_detachment()
        .unwrap()
        .commit()
        .unwrap();
    let mut sink = FailingSink {
        accepted: 0,
        limit: 128,
    };
    assert!(matches!(
        commit.package().write_to(&mut sink),
        Err(Error::Opc(OpcError::IncompleteOutput { .. }))
    ));
    assert_eq!(sink.accepted, 128);
}

struct FailingSink {
    accepted: usize,
    limit: usize,
}

impl Write for FailingSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if self.accepted >= self.limit {
            return Err(io::Error::other("injected sink failure"));
        }
        let count = bytes.len().min(self.limit - self.accepted);
        self.accepted += count;
        Ok(count)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}
