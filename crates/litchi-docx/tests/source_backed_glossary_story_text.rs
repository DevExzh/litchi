use std::io::{self, Write};
use std::sync::Arc;

use litchi_core::{OwnedSource, Position};
use litchi_docx::glossary::{Id, Name};
use litchi_docx::source_backed::{self, story_text};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcPackage, PackURI, PackageWriter, Part, TargetMode};

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const WS: &str = "http://purl.oclc.org/ooxml/wordprocessingml/main";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const VML: &str = "urn:schemas-microsoft-com:vml";
const STRICT_COMMENTS: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/comments";
const GLOSSARY_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/glossaryDocument";
const STRICT_GLOSSARY_REL: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/glossaryDocument";
const ALPHA_ID: &str = "{12345678-1234-4ABC-8DEF-1234567890AB}";
const BETA_ID: &str = "{12345678-1234-4ABC-8DEF-1234567890AC}";
const GAMMA_ID: &str = "{12345678-1234-4ABC-8DEF-1234567890AD}";
const MAX_GLOSSARY_SELECTOR_NAME_BYTES: usize = 1024 * 1024;

#[derive(Clone, Copy)]
enum Dialect {
    Transitional,
    Strict,
}

impl Dialect {
    const fn word(self) -> &'static str {
        match self {
            Self::Transitional => W,
            Self::Strict => WS,
        }
    }

    const fn office_document(self) -> &'static str {
        match self {
            Self::Transitional => rt::OFFICE_DOCUMENT,
            Self::Strict => rt::STRICT_OFFICE_DOCUMENT,
        }
    }

    const fn glossary(self) -> &'static str {
        match self {
            Self::Transitional => GLOSSARY_REL,
            Self::Strict => STRICT_GLOSSARY_REL,
        }
    }

    const fn opposite(self) -> Self {
        match self {
            Self::Transitional => Self::Strict,
            Self::Strict => Self::Transitional,
        }
    }
}

#[derive(Clone, Copy, Debug)]
enum Mutation {
    Valid,
    Unnamed,
    DuplicateName,
    DuplicateId,
    MissingBody,
    DuplicateBody,
    WrongContentType,
    ExternalRelationship,
    DanglingRelationship,
    WrongOwner,
    Orphan,
    RelationshipDialectMismatch,
    RootDialectMismatch,
    OutboundStoryRelationship,
    InboundRelationship,
    Signed,
    LegalReferences,
    TwoParagraphs,
    BatchThreeEntries,
    UnselectedManyParagraphs,
    UnknownEntity,
    LargeNamespace,
    DocPartPrAfterBody,
    UnexpectedDirectChild,
    StrictTransitionalAuxiliary,
    TransitionalStrictAuxiliary,
    TransitionalRelationshipAttribute,
    TransitionalVml,
    IllegalNameReference,
    QuotedNamespace,
}

fn compact(xml: impl AsRef<[u8]>) -> Vec<u8> {
    xml.as_ref().to_vec()
}

fn fixture(dialect: Dialect, mutation: Mutation) -> Vec<u8> {
    let word = dialect.word();
    let root_word = if matches!(mutation, Mutation::RootDialectMismatch) {
        dialect.opposite().word()
    } else {
        word
    };
    let beta_name = if matches!(mutation, Mutation::DuplicateName) {
        "Alpha"
    } else {
        "Beta"
    };
    let beta_id = if matches!(mutation, Mutation::DuplicateId) {
        ALPHA_ID
    } else {
        BETA_ID
    };
    let alpha_name = if matches!(mutation, Mutation::Unnamed) {
        ""
    } else if matches!(mutation, Mutation::IllegalNameReference) {
        "Alpha&#x1;"
    } else {
        "Alpha"
    };
    let alpha_namespace = if matches!(mutation, Mutation::LargeNamespace) {
        format!(r#" xmlns:u="urn:{}""#, "n".repeat(4096))
    } else if matches!(mutation, Mutation::QuotedNamespace) {
        String::from(r##" xmlns:u="urn:opaque&quot;x""##)
    } else if matches!(mutation, Mutation::TransitionalRelationshipAttribute) {
        format!(r#" xmlns:u="urn:opaque" xmlns:r="{R}""#)
    } else {
        String::from(r#" xmlns:u="urn:opaque""#)
    };
    let visible = if matches!(mutation, Mutation::UnknownEntity) {
        "&unknown;"
    } else if matches!(mutation, Mutation::LegalReferences) {
        "&amp;&#x41;&lt;"
    } else if matches!(mutation, Mutation::LargeNamespace) {
        "namespace"
    } else {
        "target"
    };
    let vml = if matches!(mutation, Mutation::TransitionalVml) {
        format!(r#"<v:shape xmlns:v="{VML}" v:style="x"/>"#)
    } else {
        String::new()
    };
    let paragraph_attribute = if matches!(mutation, Mutation::TransitionalRelationshipAttribute) {
        r#" r:embed="rId1""#
    } else {
        ""
    };
    let paragraphs = if matches!(mutation, Mutation::TwoParagraphs) {
        format!(
            r#"{vml}<w:p{paragraph_attribute}><w:r><w:t>{visible}</w:t></w:r></w:p><w:p><w:r><w:t>second</w:t></w:r></w:p>"#
        )
    } else {
        format!(r#"{vml}<w:p{paragraph_attribute}><w:r><w:t>{visible}</w:t></w:r></w:p>"#)
    };
    let beta_paragraphs = if matches!(mutation, Mutation::UnselectedManyParagraphs) {
        r#"<w:p><w:r><w:t>sibling</w:t></w:r></w:p><w:p><w:r><w:t>unselected-second</w:t></w:r></w:p><w:p><w:r><w:t>unselected-third</w:t></w:r></w:p>"#
    } else {
        r#"<w:p><w:r><w:t>sibling</w:t></w:r></w:p>"#
    };
    let gamma_entry = if matches!(
        mutation,
        Mutation::BatchThreeEntries | Mutation::Signed | Mutation::UnselectedManyParagraphs
    ) {
        format!(
            r#"<w:docPart><w:docPartPr><w:name w:val="Gamma"/><w:guid w:val="{GAMMA_ID}"/></w:docPartPr><w:docPartBody><w:p><w:r><w:t>third</w:t></w:r></w:p></w:docPartBody></w:docPart>"#
        )
    } else {
        String::new()
    };
    let alpha_body = if matches!(mutation, Mutation::MissingBody) {
        String::new()
    } else {
        format!(
            r#"<w:docPartBody xmlns:w="{root_word}"{alpha_namespace}><!--before--><u:opaque u:value="keep"/>{paragraphs}<!--after--></w:docPartBody>"#
        )
    };
    let duplicate_body = if matches!(mutation, Mutation::DuplicateBody) {
        format!(r#"<w:docPartBody xmlns:w="{root_word}"><w:p/></w:docPartBody>"#)
    } else {
        String::new()
    };
    let alpha_pr = format!(
        r#"<w:docPartPr><w:name w:val="{alpha_name}"/><w:guid w:val="{ALPHA_ID}"/></w:docPartPr>"#
    );
    let alpha_extra = if matches!(mutation, Mutation::UnexpectedDirectChild) {
        "<w:unexpected/>"
    } else {
        ""
    };
    let alpha_entry = if matches!(mutation, Mutation::DocPartPrAfterBody) {
        format!(r#"<w:docPart>{alpha_extra}{alpha_body}{duplicate_body}{alpha_pr}</w:docPart>"#)
    } else {
        format!(r#"<w:docPart>{alpha_extra}{alpha_pr}{alpha_body}{duplicate_body}</w:docPart>"#)
    };
    let glossary_xml = format!(
        r#"<w:glossaryDocument xmlns:w="{root_word}" xmlns:x="urn:root-opaque"><!--root-before--><x:rootOpaque x:value="preserve"/><w:docParts>{alpha_entry}<w:docPart><w:docPartPr><w:name w:val="{beta_name}"/><w:guid w:val="{beta_id}"/></w:docPartPr><w:docPartBody>{beta_paragraphs}</w:docPartBody></w:docPart>{gamma_entry}</w:docParts><!--root-after--></w:glossaryDocument>"#
    );
    // The unknown entity and the illegal character reference are not
    // well-formed, so the writer publishes a stand-in without them.
    let published_glossary = glossary_xml
        .replace("&unknown;", "unknown")
        .replace("&#x1;", "");
    let mut package = OpcPackage::new();
    let mut main = BlobPart::new(
        PackURI::new("/word/document.xml").unwrap(),
        ct::WML_DOCUMENT_MAIN.to_owned(),
        compact(format!(
            r#"<w:document xmlns:w="{word}"><w:body><w:p><w:r><w:t>main</w:t></w:r></w:p></w:body></w:document>"#
        )),
    );
    let relationship = if matches!(mutation, Mutation::RelationshipDialectMismatch) {
        dialect.opposite().glossary()
    } else {
        dialect.glossary()
    };
    let target_mode = if matches!(mutation, Mutation::ExternalRelationship) {
        TargetMode::External
    } else {
        TargetMode::Internal
    };
    if !matches!(mutation, Mutation::WrongOwner | Mutation::Orphan) {
        main.rels_mut()
            .try_add_relationship(
                relationship.to_owned(),
                "glossary/document.xml".to_owned(),
                "rGlossary".to_owned(),
                target_mode,
            )
            .unwrap();
    }
    package.try_add_part(Box::new(main)).unwrap();
    if !matches!(mutation, Mutation::WrongOwner | Mutation::Orphan) {
        package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new("/word/glossary/document.xml").unwrap(),
                if matches!(mutation, Mutation::WrongContentType) {
                    ct::WML_DOCUMENT_MAIN.to_owned()
                } else {
                    ct::WML_DOCUMENT_GLOSSARY.to_owned()
                },
                compact(published_glossary.as_bytes()),
            )))
            .unwrap();
    }
    if matches!(
        mutation,
        Mutation::StrictTransitionalAuxiliary | Mutation::TransitionalStrictAuxiliary
    ) {
        let glossary = package
            .get_part_mut(&PackURI::new("/word/glossary/document.xml").unwrap())
            .unwrap();
        let auxiliary_relationship = if matches!(mutation, Mutation::StrictTransitionalAuxiliary) {
            rt::COMMENTS
        } else {
            STRICT_COMMENTS
        };
        glossary
            .rels_mut()
            .try_add_relationship(
                auxiliary_relationship.to_owned(),
                "comments.xml".to_owned(),
                "rComments".to_owned(),
                TargetMode::Internal,
            )
            .unwrap();
        package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new("/word/comments.xml").unwrap(),
                ct::WML_COMMENTS.to_owned(),
                compact(format!(r#"<w:comments xmlns:w="{word}"/>"#)),
            )))
            .unwrap();
    }
    if matches!(mutation, Mutation::DanglingRelationship) {
        let mut owner = BlobPart::new(
            PackURI::new("/word/glossary/document.xml").unwrap(),
            ct::WML_DOCUMENT_GLOSSARY.to_owned(),
            compact(format!(r#"<w:glossaryDocument xmlns:w="{word}"/>"#)),
        );
        owner
            .rels_mut()
            .try_add_relationship(
                rt::COMMENTS.to_owned(),
                "missing-comments.xml".to_owned(),
                "rDangling".to_owned(),
                TargetMode::Internal,
            )
            .unwrap();
        package
            .get_part_mut(&PackURI::new("/word/glossary/document.xml").unwrap())
            .unwrap()
            .rels_mut()
            .try_add_relationship(
                rt::COMMENTS.to_owned(),
                "missing-comments.xml".to_owned(),
                "rDangling".to_owned(),
                TargetMode::Internal,
            )
            .unwrap();
        drop(owner);
    }
    if matches!(mutation, Mutation::WrongOwner) {
        package
            .rels_mut()
            .try_add_relationship(
                relationship.to_owned(),
                "word/glossary/document.xml".to_owned(),
                "rPackageGlossary".to_owned(),
                TargetMode::Internal,
            )
            .unwrap();
    }
    if matches!(mutation, Mutation::Orphan) {
        package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new("/word/glossary/document.xml").unwrap(),
                ct::WML_DOCUMENT_GLOSSARY.to_owned(),
                compact(published_glossary.as_bytes()),
            )))
            .unwrap();
    }
    if matches!(mutation, Mutation::OutboundStoryRelationship) {
        let glossary = package
            .get_part_mut(&PackURI::new("/word/glossary/document.xml").unwrap())
            .unwrap();
        glossary
            .rels_mut()
            .try_add_relationship(
                rt::HEADER.to_owned(),
                "../header1.xml".to_owned(),
                "rHeader".to_owned(),
                TargetMode::Internal,
            )
            .unwrap();
        let mut header = BlobPart::new(
            PackURI::new("/word/header1.xml").unwrap(),
            ct::WML_HEADER.to_owned(),
            compact(format!(r#"<w:hdr xmlns:w="{word}"/>"#)),
        );
        header
            .rels_mut()
            .try_add_relationship(
                rt::FOOTER.to_owned(),
                "footer1.xml".to_owned(),
                "rOutbound".to_owned(),
                TargetMode::Internal,
            )
            .unwrap();
        package.try_add_part(Box::new(header)).unwrap();
        package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new("/word/footer1.xml").unwrap(),
                ct::WML_FOOTER.to_owned(),
                compact(format!(r#"<w:ftr xmlns:w="{word}"/>"#)),
            )))
            .unwrap();
    }
    if matches!(mutation, Mutation::InboundRelationship) {
        package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new("/word/header1.xml").unwrap(),
                ct::WML_HEADER.to_owned(),
                compact(format!(r#"<w:hdr xmlns:w="{word}"/>"#)),
            )))
            .unwrap();
        package
            .get_part_mut(&PackURI::new("/word/header1.xml").unwrap())
            .unwrap()
            .rels_mut()
            .try_add_relationship(
                "urn:litchi:test:inbound".to_owned(),
                "glossary/document.xml".to_owned(),
                "rInbound".to_owned(),
                TargetMode::Internal,
            )
            .unwrap();
    }
    if matches!(mutation, Mutation::Signed) {
        package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new("/_xmlsignatures/origin.sigs").unwrap(),
                ct::OPC_DIGITAL_SIGNATURE_ORIGIN.to_owned(),
                Vec::new(),
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
    }
    package.relate_to("word/document.xml", dialect.office_document());
    let bytes = PackageWriter::to_bytes(&package).unwrap();
    if published_glossary == glossary_xml {
        bytes
    } else {
        with_unaudited_member(
            &bytes,
            "word/glossary/document.xml",
            glossary_xml.as_bytes(),
        )
    }
}

/// Rewrites `archive` with `member` holding `bytes`, without the publication
/// audit. Change 0750: the eager writer refuses to publish XML that is not
/// well-formed, so a fixture that must carry such XML to a reader is written
/// with a well-formed stand-in and its bytes are put back here.
fn with_unaudited_member(archive: &[u8], member: &str, bytes: &[u8]) -> Vec<u8> {
    let reader =
        litchi_opc::phys_pkg::PhysPkgReader::new(archive, litchi_opc::SourceLimits::default())
            .unwrap();
    let mut writer = soapberry_zip::office::StreamingArchiveWriter::new(
        soapberry_zip::office::StreamingArchiveLimits::default(),
    );
    for name in reader.member_names().unwrap() {
        let data = if name == member {
            bytes.to_vec()
        } else {
            reader.read_member(&name).unwrap()
        };
        writer.write_deflated_sized(&name, &data).unwrap();
    }
    writer.finish_to_bytes().unwrap()
}

fn open(bytes: &[u8]) -> source_backed::Package {
    source_backed::Package::from_read_at(Arc::new(OwnedSource::new(bytes.to_vec()))).unwrap()
}

fn selector(selector: story_text::GlossarySelector) -> story_text::Selector {
    story_text::Selector::Glossary(selector)
}

fn alpha_by_name() -> story_text::Selector {
    selector(story_text::GlossarySelector::ByName("Alpha".to_owned()))
}

fn alpha_by_id() -> story_text::Selector {
    selector(story_text::GlossarySelector::ById(
        Id::new(ALPHA_ID).unwrap(),
    ))
}

fn glossary_by_index(index: usize) -> story_text::Selector {
    selector(story_text::GlossarySelector::ByIndex(index))
}

fn part(bytes: &[u8], name: &str) -> Vec<u8> {
    OpcPackage::from_bytes(bytes, litchi_opc::SourceLimits::default())
        .unwrap()
        .get_part(&PackURI::new(name).unwrap())
        .unwrap()
        .blob()
        .to_vec()
}

fn replace_once(bytes: &[u8], from: &str, to: &str) -> Vec<u8> {
    let mut result = bytes.to_vec();
    let index = result
        .windows(from.len())
        .position(|window| window == from.as_bytes())
        .unwrap();
    result.splice(index..index + from.len(), to.as_bytes().iter().copied());
    result
}

fn quoted_wrapper_bytes(source: &[u8], word_namespace: &str) -> usize {
    let raw = part(source, "/word/glossary/document.xml");
    let open = raw
        .windows(b"><!--before-->".len())
        .position(|window| window == b"><!--before-->")
        .unwrap()
        + 1;
    let close = raw
        .windows(b"</w:docPartBody>".len())
        .position(|window| window == b"</w:docPartBody>")
        .unwrap();
    b"<w:document xmlns:u=\"urn:opaque&quot;x\" xmlns:x=\"urn:root-opaque\" xmlns:w=\"".len()
        + word_namespace.len()
        + b"\"><w:body>".len()
        + close
        - open
        + b"</w:body></w:document>".len()
}

#[test]
fn glossary_selectors_read_unique_entries_in_both_dialects() {
    for dialect in [Dialect::Transitional, Dialect::Strict] {
        let source = fixture(dialect, Mutation::Valid);
        let package = open(&source);
        for selected in [
            alpha_by_name(),
            alpha_by_id(),
            selector(story_text::GlossarySelector::ByNameAndId {
                name: "Alpha".to_owned(),
                id: Id::new(ALPHA_ID).unwrap(),
            }),
            selector(story_text::GlossarySelector::ByIndex(0)),
        ] {
            let snapshot = package.story_text_snapshot(selected).unwrap();
            assert_eq!(snapshot.paragraph_count().unwrap(), 1);
            assert_eq!(
                snapshot.paragraph_text(0).unwrap().as_deref(),
                Some("target")
            );
        }
    }
}

#[test]
fn glossary_selection_rejects_missing_ambiguous_and_malformed_identity() {
    assert!(
        open(&fixture(Dialect::Transitional, Mutation::Valid))
            .story_text_snapshot(selector(story_text::GlossarySelector::ByName(
                "Missing".to_owned()
            )))
            .is_err()
    );
    assert!(
        open(&fixture(Dialect::Transitional, Mutation::DuplicateName))
            .story_text_snapshot(alpha_by_name())
            .is_err()
    );
    assert!(
        open(&fixture(Dialect::Transitional, Mutation::DuplicateId))
            .story_text_snapshot(alpha_by_id())
            .is_err()
    );
    assert!(
        open(&fixture(Dialect::Transitional, Mutation::Unnamed))
            .story_text_snapshot(alpha_by_name())
            .is_err()
    );
    assert!(
        open(&fixture(Dialect::Transitional, Mutation::MissingBody))
            .story_text_snapshot(alpha_by_id())
            .is_err()
    );
    assert!(
        open(&fixture(Dialect::Transitional, Mutation::DuplicateBody))
            .story_text_snapshot(alpha_by_id())
            .is_err()
    );
    assert!(
        open(&fixture(
            Dialect::Transitional,
            Mutation::DocPartPrAfterBody
        ))
        .story_text_snapshot(glossary_by_index(0))
        .is_err()
    );
    assert!(
        open(&fixture(
            Dialect::Transitional,
            Mutation::UnexpectedDirectChild
        ))
        .story_text_snapshot(glossary_by_index(0))
        .is_err()
    );
    assert!(
        open(&fixture(
            Dialect::Transitional,
            Mutation::IllegalNameReference
        ))
        .story_text_snapshot(alpha_by_name())
        .is_err()
    );
}

#[test]
fn glossary_selector_name_limits_accept_exact_and_refuse_one_over() {
    let exact = "n".repeat(MAX_GLOSSARY_SELECTOR_NAME_BYTES);
    assert!(matches!(
        open(&fixture(Dialect::Transitional, Mutation::Valid))
            .story_text_snapshot(story_text::Selector::glossary_by_name(exact)),
        Err(story_text::Error::MissingGlossaryEntry { .. })
    ));
    let over = "n".repeat(MAX_GLOSSARY_SELECTOR_NAME_BYTES + 1);
    assert!(matches!(
        open(&fixture(Dialect::Transitional, Mutation::Valid))
            .story_text_snapshot(story_text::Selector::glossary_by_name(over)),
        Err(story_text::Error::Limit {
            kind: story_text::LimitKind::GlossarySelectorNameBytes,
            actual,
            maximum: MAX_GLOSSARY_SELECTOR_NAME_BYTES,
        }) if actual == MAX_GLOSSARY_SELECTOR_NAME_BYTES + 1
    ));

    let exact = "n".repeat(MAX_GLOSSARY_SELECTOR_NAME_BYTES);
    assert!(matches!(
        open(&fixture(Dialect::Transitional, Mutation::Valid)).story_text_snapshot(
            story_text::Selector::glossary_by_name_and_id(exact, Id::new(ALPHA_ID).unwrap())
        ),
        Err(story_text::Error::MissingGlossaryEntry { .. })
    ));
    let over = "n".repeat(MAX_GLOSSARY_SELECTOR_NAME_BYTES + 1);
    assert!(matches!(
        open(&fixture(Dialect::Transitional, Mutation::Valid)).story_text_snapshot(
            story_text::Selector::glossary_by_name_and_id(over, Id::new(ALPHA_ID).unwrap())
        ),
        Err(story_text::Error::Limit {
            kind: story_text::LimitKind::GlossarySelectorNameBytes,
            actual,
            maximum: MAX_GLOSSARY_SELECTOR_NAME_BYTES,
        }) if actual == MAX_GLOSSARY_SELECTOR_NAME_BYTES + 1
    ));
    assert!(
        open(&fixture(Dialect::Transitional, Mutation::Valid))
            .story_text_snapshot(alpha_by_id())
            .is_ok()
    );
}

#[test]
fn glossary_entry_edit_splices_only_selected_body_and_round_trips() {
    for dialect in [Dialect::Transitional, Dialect::Strict] {
        let source = fixture(dialect, Mutation::Valid);
        let package = open(&source);
        let snapshot = package.story_text_snapshot(alpha_by_name()).unwrap();
        let mut edit = snapshot.raw_tree(snapshot.limits_value()).unwrap();
        edit.replace_paragraph_text(Position::new(0), "changed")
            .unwrap();
        let patch = edit.patch().unwrap();
        assert_eq!(
            patch.target().paragraph_text(0).unwrap().as_deref(),
            Some("changed")
        );
        let mut output = Vec::new();
        let publication = package
            .apply_story_text(&patch.clone().into())
            .and_then(|committed| {
                committed.package().write_to(&mut output)?;
                Ok(committed)
            })
            .unwrap();
        let reopened = open(&output);
        assert_eq!(
            reopened
                .story_text_snapshot(alpha_by_name())
                .unwrap()
                .extract_text()
                .unwrap(),
            "changed"
        );
        let raw = String::from_utf8(part(&output, "/word/glossary/document.xml")).unwrap();
        assert!(raw.contains("root-before") && raw.contains("root-after"));
        assert!(raw.contains("rootOpaque") && raw.contains("u:value=\"keep\""));
        assert!(raw.contains("urn:root-opaque") && raw.contains("urn:opaque"));
        assert!(
            raw.contains("sibling")
                && raw.contains("<!--before-->")
                && raw.contains("<!--after-->")
        );
        assert_eq!(
            part(&output, "/word/glossary/document.xml"),
            replace_once(
                &part(&source, "/word/glossary/document.xml"),
                "target",
                "changed"
            )
        );
        assert_eq!(
            part(&source, "/word/document.xml"),
            part(&output, "/word/document.xml")
        );
        assert_eq!(
            publication
                .package()
                .story_text_snapshot(alpha_by_name())
                .unwrap()
                .extract_text()
                .unwrap(),
            "changed"
        );
    }
}

#[test]
fn glossary_noop_and_inverse_preserve_the_complete_artifact() {
    let source = fixture(Dialect::Transitional, Mutation::Valid);
    let package = open(&source);
    let snapshot = package.story_text_snapshot(alpha_by_name()).unwrap();
    let patch = snapshot
        .raw_tree(snapshot.limits_value())
        .unwrap()
        .patch()
        .unwrap();
    assert!(patch.is_noop());
    let mut exact = Vec::new();
    package
        .apply_story_text(&patch.clone().into())
        .and_then(|committed| Ok(committed.package().write_to(&mut exact)?))
        .unwrap();
    assert_eq!(exact, source);

    let source = fixture(Dialect::Transitional, Mutation::Valid);
    let package = open(&source);
    let snapshot = package.story_text_snapshot(alpha_by_name()).unwrap();
    let mut edit = snapshot.raw_tree(snapshot.limits_value()).unwrap();
    edit.replace_paragraph_text(Position::new(0), "changed")
        .unwrap();
    let patch = edit.patch().unwrap();
    let mut output = Vec::new();
    let publication = package
        .apply_story_text(&patch.clone().into())
        .and_then(|committed| {
            committed.package().write_to(&mut output)?;
            Ok(committed)
        })
        .unwrap();
    let mut restored = Vec::new();
    open(&output)
        .apply_story_text(&publication.patch().inverse())
        .and_then(|committed| Ok(committed.package().write_to(&mut restored)?))
        .unwrap();
    assert_eq!(restored, source);
}

#[test]
fn raw_glossary_tree_caches_projection_and_invalidates_after_mutation() {
    let package = open(&fixture(Dialect::Transitional, Mutation::BatchThreeEntries));
    let snapshot = package
        .glossary_text_batch_snapshot(glossary_triplet())
        .unwrap();
    let mut tree = snapshot.raw_tree(snapshot.limits_value()).unwrap();

    let first = tree.projected().unwrap();
    let second = tree.projected().unwrap();
    assert_eq!(first.part_fingerprint(), second.part_fingerprint());
    assert_eq!(
        first.paragraph_text(0, 0).unwrap(),
        second.paragraph_text(0, 0).unwrap()
    );

    tree.replace_paragraph_text(0, Position::new(0), "cached glossary mutation")
        .unwrap();
    let changed = tree.projected().unwrap();
    assert_ne!(first.part_fingerprint(), changed.part_fingerprint());
    assert_eq!(
        changed.paragraph_text(0, 0).unwrap().as_deref(),
        Some("cached glossary mutation")
    );
    let patch = tree.patch().unwrap();
    assert_eq!(
        patch.target().part_fingerprint(),
        changed.part_fingerprint()
    );
}

#[test]
fn raw_glossary_replay_reopens_matches_and_chains_edits() {
    let source = fixture(Dialect::Transitional, Mutation::BatchThreeEntries);
    let author = open(&source);
    let snapshot = author
        .glossary_text_batch_snapshot(glossary_triplet())
        .unwrap();
    let mut tree = snapshot.raw_tree(snapshot.limits_value()).unwrap();
    tree.replace_paragraph_text(0, Position::new(0), "first raw glossary")
        .unwrap();
    let patch = tree.patch().unwrap();

    let recipient = open(&source)
        .glossary_text_batch_snapshot(glossary_triplet())
        .unwrap();
    let first = patch.apply(&recipient).unwrap();
    assert_eq!(
        first.paragraph_text(0, 0).unwrap().as_deref(),
        Some("first raw glossary")
    );

    let restored = patch.inverse().apply(&first).unwrap();
    assert_eq!(restored.part_fingerprint(), recipient.part_fingerprint());
    assert_eq!(
        restored.paragraph_text(0, 0).unwrap().as_deref(),
        Some("target")
    );

    let mut second_tree = first.raw_tree(first.limits_value()).unwrap();
    second_tree
        .replace_paragraph_text(1, Position::new(0), "second raw glossary")
        .unwrap();
    let second_patch = second_tree.patch().unwrap();
    let second = second_patch.apply(&first).unwrap();
    assert_eq!(
        second.paragraph_text(1, 0).unwrap().as_deref(),
        Some("second raw glossary")
    );
    let second_restored = second_patch.inverse().apply(&second).unwrap();
    assert_eq!(
        second_restored.paragraph_text(0, 0).unwrap().as_deref(),
        Some("first raw glossary")
    );
}

#[test]
fn glossary_replay_uses_recipient_limits_for_broad_raw_patches() {
    let source = batch_source();
    let producer_limits =
        story_text::Limits::new(128 * 1024, 8, 10_000, 64, 128 * 1024, 8, 128 * 1024)
            .unwrap()
            .with_secondary_entry_limits(3, 16 * 1024, 256)
            .unwrap();
    let producer = open(&source)
        .with_semantic_limits(litchi_docx::SemanticLimits::default().story_text(producer_limits));
    let snapshot = producer
        .glossary_text_batch_snapshot(glossary_triplet())
        .unwrap();
    let mut tree = snapshot.raw_tree(snapshot.limits_value()).unwrap();
    let glossary_bytes = part(&source, "/word/glossary/document.xml").len();
    tree.replace_paragraph_text(0, Position::new(0), "wide".repeat(glossary_bytes))
        .unwrap();
    let patch = tree.patch().unwrap();

    let accepted_limits =
        story_text::Limits::new(128 * 1024, 8, 10_000, 64, 64 * 1024, 8, 128 * 1024)
            .unwrap()
            .with_secondary_entry_limits(3, 16 * 1024, 256)
            .unwrap();
    let accepted = open(&source)
        .with_semantic_limits(litchi_docx::SemanticLimits::default().story_text(accepted_limits))
        .glossary_text_batch_snapshot(glossary_triplet())
        .unwrap();
    let accepted_target = patch.apply(&accepted).unwrap();
    assert_eq!(accepted_target.limits_value(), accepted_limits);

    let narrow_limits =
        story_text::Limits::new(128 * 1024, 8, 10_000, 64, glossary_bytes, 8, 128 * 1024)
            .unwrap()
            .with_secondary_entry_limits(3, 16 * 1024, 256)
            .unwrap();
    let narrow = open(&source)
        .with_semantic_limits(litchi_docx::SemanticLimits::default().story_text(narrow_limits));
    let narrow_snapshot = narrow
        .glossary_text_batch_snapshot(glossary_triplet())
        .unwrap();
    assert!(matches!(
        patch.apply(&narrow_snapshot),
        Err(story_text::Error::Limit {
            kind: story_text::LimitKind::GlossaryBatchOutputBytes,
            ..
        })
    ));
    let package_patch = patch.clone().into();
    assert!(matches!(
        narrow.apply_glossary_text_batch(&package_patch),
        Err(story_text::Error::Limit {
            kind: story_text::LimitKind::GlossaryBatchOutputBytes,
            ..
        })
    ));
}

#[test]
fn glossary_topology_and_dialect_boundaries_refuse_before_editing() {
    for (mutation, expected) in [
        (Mutation::WrongContentType, "invalid DOCX content type"),
        (
            Mutation::ExternalRelationship,
            "glossary relationship cannot be external",
        ),
        (Mutation::DanglingRelationship, "missing-comments.xml"),
        (
            Mutation::WrongOwner,
            "package root cannot source a glossary-document relationship",
        ),
        (
            Mutation::Orphan,
            "main document glossary relationship is missing",
        ),
        (
            Mutation::RelationshipDialectMismatch,
            "Word story relationship uses a dialect different from the package root",
        ),
        (
            Mutation::RootDialectMismatch,
            "glossary relationship and XML dialects differ",
        ),
        (
            Mutation::OutboundStoryRelationship,
            "selected story part has an outbound Word story relationship",
        ),
        (
            Mutation::InboundRelationship,
            "glossary target has an ambiguous inbound relationship closure",
        ),
    ] {
        let source = fixture(Dialect::Transitional, mutation);
        let rejection =
            match source_backed::Package::from_read_at(Arc::new(OwnedSource::new(source))) {
                Err(error) => error.to_string(),
                Ok(package) => package
                    .story_text_snapshot(alpha_by_name())
                    .err()
                    .unwrap_or_else(|| panic!("mutation should be rejected: {mutation:?}"))
                    .to_string(),
            };
        assert!(
            rejection.contains(expected),
            "unexpected rejection for {mutation:?}: {rejection}"
        );
    }
    for (dialect, mutation, expected) in [
        (
            Dialect::Strict,
            Mutation::StrictTransitionalAuxiliary,
            "Word story relationship uses a dialect different from the package root",
        ),
        (
            Dialect::Transitional,
            Mutation::TransitionalStrictAuxiliary,
            "Word story relationship uses a dialect different from the package root",
        ),
        (
            Dialect::Strict,
            Mutation::TransitionalRelationshipAttribute,
            "mixes Strict and Transitional relationship attributes",
        ),
        (
            Dialect::Strict,
            Mutation::TransitionalVml,
            "Strict glossary content cannot contain VML",
        ),
    ] {
        let source = fixture(dialect, mutation);
        let rejection =
            match source_backed::Package::from_read_at(Arc::new(OwnedSource::new(source))) {
                Err(error) => error.to_string(),
                Ok(package) => package
                    .story_text_snapshot(alpha_by_name())
                    .err()
                    .unwrap_or_else(|| panic!("mutation should be rejected: {mutation:?}"))
                    .to_string(),
            };
        assert!(
            rejection.contains(expected),
            "unexpected rejection for {mutation:?}: {rejection}"
        );
    }
}

#[test]
fn glossary_signed_change_refuses_but_signed_noop_is_exact() {
    let source = fixture(Dialect::Transitional, Mutation::Signed);
    let package = open(&source);
    let snapshot = package.story_text_snapshot(alpha_by_name()).unwrap();
    let patch = snapshot
        .raw_tree(snapshot.limits_value())
        .unwrap()
        .patch()
        .unwrap();
    let mut exact = Vec::new();
    package
        .apply_story_text(&patch.clone().into())
        .and_then(|committed| Ok(committed.package().write_to(&mut exact)?))
        .unwrap();
    assert_eq!(exact, source);

    let source = fixture(Dialect::Transitional, Mutation::Signed);
    let package = open(&source);
    let snapshot = package.story_text_snapshot(alpha_by_name()).unwrap();
    let mut edit = snapshot.raw_tree(snapshot.limits_value()).unwrap();
    edit.replace_paragraph_text(Position::new(0), "changed")
        .unwrap();
    let patch = edit.patch().unwrap();
    let mut refused = Vec::new();
    assert!(
        package
            .apply_story_text(&patch.clone().into())
            .and_then(|committed| {
                committed.package().write_to(&mut refused)?;
                Ok(committed)
            })
            .is_err()
    );
    assert!(refused.is_empty());
}

#[test]
fn glossary_references_and_limits_are_failure_atomic() {
    let legal = open(&fixture(Dialect::Transitional, Mutation::LegalReferences))
        .story_text_snapshot(alpha_by_name())
        .unwrap();
    assert_eq!(legal.extract_text().unwrap(), "&A<");
    assert!(
        open(&fixture(Dialect::Transitional, Mutation::UnknownEntity))
            .story_text_snapshot(alpha_by_name())
            .is_err()
    );
    let source = fixture(Dialect::Transitional, Mutation::LargeNamespace);
    let limits = story_text::Limits::new(16 * 1024, 8, 10_000, 64, 4096, 4096, 4096)
        .unwrap()
        .with_max_namespace_bytes(64)
        .unwrap();
    assert!(
        open(&source)
            .with_semantic_limits(litchi_docx::SemanticLimits::default().story_text(limits))
            .story_text_snapshot(alpha_by_name())
            .is_err()
    );

    let paragraph_limits = story_text::Limits::new(4096, 1, 10_000, 64, 4096, 4096, 4096).unwrap();
    assert!(
        open(&fixture(Dialect::Transitional, Mutation::TwoParagraphs))
            .with_semantic_limits(
                litchi_docx::SemanticLimits::default().story_text(paragraph_limits)
            )
            .story_text_snapshot(alpha_by_name())
            .is_err()
    );
    let output_limits = story_text::Limits::new(4096, 8, 10_000, 64, 1, 4096, 4096).unwrap();
    assert!(
        open(&fixture(Dialect::Transitional, Mutation::Valid))
            .with_semantic_limits(litchi_docx::SemanticLimits::default().story_text(output_limits))
            .story_text_snapshot(alpha_by_name())
            .unwrap()
            .extract_text()
            .is_err()
    );
    let xml_limits = story_text::Limits::new(1, 8, 10_000, 64, 4096, 4096, 4096).unwrap();
    assert!(
        open(&source)
            .with_semantic_limits(litchi_docx::SemanticLimits::default().story_text(xml_limits))
            .story_text_snapshot(alpha_by_name())
            .is_err()
    );
    let entry_limits = story_text::Limits::default()
        .with_secondary_entry_limits(1, 1, 256)
        .unwrap();
    assert!(
        open(&source)
            .with_semantic_limits(litchi_docx::SemanticLimits::default().story_text(entry_limits))
            .story_text_snapshot(alpha_by_name())
            .is_err()
    );

    let source = fixture(Dialect::Transitional, Mutation::Valid);
    let limits = story_text::Limits::new(4096, 2, 10_000, 64, 4096, 1, 4096).unwrap();
    let package = open(&source);
    let snapshot = package
        .with_semantic_limits(litchi_docx::SemanticLimits::default().story_text(limits))
        .story_text_snapshot(alpha_by_name())
        .unwrap();
    let mut edit = snapshot.raw_tree(snapshot.limits_value()).unwrap();
    assert!(edit.replace_paragraph_text(Position::new(0), "one").is_ok());
    assert!(
        edit.replace_paragraph_text(Position::new(0), "two")
            .is_err()
    );
    assert_eq!(
        edit.projected()
            .unwrap()
            .paragraph_text(0)
            .unwrap()
            .as_deref(),
        Some("one")
    );
}

#[test]
fn glossary_quote_namespace_wrapper_has_exact_and_one_under_boundaries() {
    let source = fixture(Dialect::Transitional, Mutation::QuotedNamespace);
    let wrapper_bytes = quoted_wrapper_bytes(&source, Dialect::Transitional.word());
    let exact =
        story_text::Limits::new(16 * 1024, 8, 10_000, 64, wrapper_bytes, 4096, 4096).unwrap();
    let snapshot = open(&source)
        .with_semantic_limits(litchi_docx::SemanticLimits::default().story_text(exact))
        .story_text_snapshot(alpha_by_name())
        .unwrap();
    assert!(snapshot.raw_tree(snapshot.limits_value()).is_ok());

    let one_under =
        story_text::Limits::new(16 * 1024, 8, 10_000, 64, wrapper_bytes - 1, 4096, 4096).unwrap();
    let snapshot = open(&source)
        .with_semantic_limits(litchi_docx::SemanticLimits::default().story_text(one_under))
        .story_text_snapshot(alpha_by_name())
        .unwrap();
    assert!(matches!(
        snapshot.raw_tree(snapshot.limits_value()),
        Err(story_text::Error::Limit {
            kind: story_text::LimitKind::WrappedStoryXmlBytes,
            ..
        })
    ));
    assert_eq!(
        snapshot.paragraph_text(0).unwrap().as_deref(),
        Some("target")
    );
}

struct PartialWriter {
    writes: usize,
}

impl Write for PartialWriter {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if bytes.is_empty() {
            Ok(0)
        } else if self.writes == 0 {
            self.writes = 1;
            Ok(bytes.len().min(1))
        } else {
            Err(io::Error::other("intentional partial sink failure"))
        }
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[test]
fn glossary_publication_sink_failure_does_not_panic_and_complex_edit_is_refused() {
    let source = fixture(Dialect::Transitional, Mutation::Valid);
    let package = open(&source);
    let snapshot = package.story_text_snapshot(alpha_by_name()).unwrap();
    let mut edit = snapshot.raw_tree(snapshot.limits_value()).unwrap();
    assert!(
        edit.replace_paragraph_text(Position::new(0), "a\u{0}b")
            .is_err()
    );
    let mut edit = snapshot.raw_tree(snapshot.limits_value()).unwrap();
    edit.replace_paragraph_text(Position::new(0), "changed")
        .unwrap();
    let patch = edit.patch().unwrap();
    assert!(
        open(&source)
            .apply_story_text(&patch.clone().into())
            .and_then(|committed| {
                committed.package().write_to(PartialWriter { writes: 0 })?;
                Ok(committed)
            })
            .is_err()
    );
}

// This adapter is intentionally small: it proves that a source revision
// change is checked before a glossary-entry overlay is emitted.
struct VersionedSource {
    bytes: Vec<u8>,
    revision: std::sync::atomic::AtomicU64,
}

impl litchi_core::ReadAt for VersionedSource {
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
            378,
            self.revision.load(std::sync::atomic::Ordering::SeqCst),
        ))
    }
}

#[test]
fn glossary_source_mutation_is_rejected_and_matching_copy_is_accepted() {
    let source = fixture(Dialect::Transitional, Mutation::Valid);
    let versioned = Arc::new(VersionedSource {
        bytes: source.clone(),
        revision: std::sync::atomic::AtomicU64::new(0),
    });
    let package = source_backed::Package::from_read_at(versioned.clone()).unwrap();
    let snapshot = package.story_text_snapshot(alpha_by_name()).unwrap();
    let mut edit = snapshot.raw_tree(snapshot.limits_value()).unwrap();
    edit.replace_paragraph_text(Position::new(0), "changed")
        .unwrap();
    let patch = edit.patch().unwrap();
    versioned
        .revision
        .fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let mut refused = Vec::new();
    assert!(
        package
            .apply_story_text(&patch.clone().into())
            .and_then(|committed| {
                committed.package().write_to(&mut refused)?;
                Ok(committed)
            })
            .is_err()
    );
    assert!(refused.is_empty());
    let foreign = open(&source).story_text_snapshot(alpha_by_name()).unwrap();
    let applied = patch.apply(&foreign).unwrap();
    assert_eq!(applied.extract_text().unwrap(), "changed");
}

#[test]
fn glossary_managed_entry_edit_and_cancellation_are_refused() {
    // The source-backed API must retain its established managed boundary:
    // selection may be bounded, but detaching a glossary body for editing is
    // refused. Cancellation is checked on the same reachable operation.
    let source = fixture(Dialect::Transitional, Mutation::Valid);
    let (cancel_source, token) = litchi_core::CancellationSource::pair();
    let budget = litchi_core::Budget::root(
        "source-backed-glossary-test",
        litchi_core::Limits::new(
            2 * source.len() as u64,
            u64::MAX,
            u64::MAX,
            u64::MAX,
            u64::MAX,
            u64::MAX,
        ),
    );
    let context = litchi_core::ExecutionContext::new(
        budget,
        token,
        litchi_core::ExecutionLimits::new(
            std::num::NonZeroUsize::MIN,
            std::num::NonZeroUsize::MIN,
            std::num::NonZeroU64::new(2 * source.len() as u64).unwrap(),
            0,
        )
        .unwrap(),
    );
    let package = source_backed::Package::from_read_at_with_options(
        Arc::new(OwnedSource::new(source)),
        litchi_docx::OpenOptions::default()
            .source_limits(litchi_docx::SourceLimits::default())
            .execution_context(context),
    )
    .unwrap();
    let snapshot = package.story_text_snapshot(alpha_by_name()).unwrap();
    assert!(snapshot.raw_tree(snapshot.limits_value()).is_err());
    cancel_source.cancel();
    assert!(package.story_text_snapshot(alpha_by_name()).is_err());
}

fn glossary_triplet() -> story_text::GlossaryBatchSelector {
    story_text::GlossaryBatchSelector::try_new(vec![
        story_text::GlossarySelector::ById(Id::new(BETA_ID).unwrap()),
        story_text::GlossarySelector::ByName("Gamma".to_owned()),
        story_text::GlossarySelector::ByName("Alpha".to_owned()),
    ])
    .unwrap()
}

fn batch_source() -> Vec<u8> {
    fixture(Dialect::Transitional, Mutation::BatchThreeEntries)
}

#[test]
fn glossary_batch_does_not_charge_unselected_entry_paragraphs() {
    let source = fixture(Dialect::Transitional, Mutation::UnselectedManyParagraphs);
    let limits = story_text::Limits::new(16 * 1024, 2, 10_000, 64, 32 * 1024, 8, 4096)
        .unwrap()
        .with_secondary_entry_limits(3, 16 * 1024, 256)
        .unwrap();
    let selected = story_text::GlossaryBatchSelector::try_new(vec![
        story_text::GlossarySelector::ByName("Gamma".to_owned()),
        story_text::GlossarySelector::ByName("Alpha".to_owned()),
    ])
    .unwrap();
    let snapshot = open(&source)
        .with_semantic_limits(litchi_docx::SemanticLimits::default().story_text(limits))
        .glossary_text_batch_snapshot(selected)
        .unwrap();
    assert_eq!(snapshot.entry_count().unwrap(), 2);
    assert_eq!(snapshot.extract_text(0).unwrap().as_deref(), Some("target"));
    assert_eq!(snapshot.extract_text(1).unwrap().as_deref(), Some("third"));
}

#[test]
fn glossary_batch_is_general_n_canonical_and_preserves_siblings() {
    let source = batch_source();
    let package = open(&source);
    let snapshot = package
        .glossary_text_batch_snapshot(glossary_triplet())
        .unwrap();
    assert_eq!(snapshot.entry_count().unwrap(), 3);
    assert!(matches!(
        snapshot.entry_selector(0),
        Some(story_text::GlossarySelector::ByName(name)) if name == "Alpha"
    ));
    assert!(matches!(
        snapshot.entry_selector(1),
        Some(story_text::GlossarySelector::ById(id)) if *id == Id::new(BETA_ID).unwrap()
    ));
    assert!(matches!(
        snapshot.entry_selector(2),
        Some(story_text::GlossarySelector::ByName(name)) if name == "Gamma"
    ));
    let forward = story_text::GlossaryBatchSelector::try_new(vec![
        story_text::GlossarySelector::ByName("Alpha".to_owned()),
        story_text::GlossarySelector::ById(Id::new(BETA_ID).unwrap()),
        story_text::GlossarySelector::ByName("Gamma".to_owned()),
    ])
    .unwrap();
    let forward_snapshot = package.glossary_text_batch_snapshot(forward).unwrap();
    assert_eq!(
        snapshot.selector().selectors(),
        forward_snapshot.selector().selectors()
    );
    assert_eq!(snapshot.extract_text(0).unwrap().as_deref(), Some("target"));
    assert_eq!(
        snapshot.extract_text(1).unwrap().as_deref(),
        Some("sibling")
    );
    assert_eq!(snapshot.extract_text(2).unwrap().as_deref(), Some("third"));

    let mut edit = snapshot.raw_tree(snapshot.limits_value()).unwrap();
    edit.replace_paragraph_text(0, Position::new(0), "alpha changed")
        .unwrap();
    edit.replace_paragraph_text(2, Position::new(0), "gamma changed")
        .unwrap();
    let patch = edit.patch().unwrap();
    let mut output = Vec::new();
    let publication = package
        .apply_glossary_text_batch(&patch.clone().into())
        .and_then(|committed| {
            committed.package().write_to(&mut output)?;
            Ok(committed)
        })
        .unwrap();
    let raw = String::from_utf8(part(&output, "/word/glossary/document.xml")).unwrap();
    assert!(raw.contains("alpha changed") && raw.contains("gamma changed"));
    assert!(raw.contains("root-before") && raw.contains("root-after"));
    assert!(raw.contains("rootOpaque") && raw.contains("u:value=\"keep\""));
    assert!(
        raw.contains("sibling") && raw.contains("<!--before-->") && raw.contains("<!--after-->")
    );
    assert_eq!(
        part(&source, "/word/document.xml"),
        part(&output, "/word/document.xml")
    );
    assert_eq!(
        open(&output)
            .glossary_text_batch_snapshot(glossary_triplet())
            .unwrap()
            .extract_text(0)
            .unwrap()
            .as_deref(),
        Some("alpha changed")
    );
    assert_eq!(
        open(&output)
            .glossary_text_batch_snapshot(glossary_triplet())
            .unwrap()
            .extract_text(2)
            .unwrap()
            .as_deref(),
        Some("gamma changed")
    );

    let deterministic_package = open(&source);
    let deterministic_snapshot = deterministic_package
        .glossary_text_batch_snapshot(glossary_triplet())
        .unwrap();
    let mut deterministic_edit = deterministic_snapshot
        .raw_tree(snapshot.limits_value())
        .unwrap();
    deterministic_edit
        .replace_paragraph_text(0, Position::new(0), "alpha changed")
        .unwrap();
    deterministic_edit
        .replace_paragraph_text(2, Position::new(0), "gamma changed")
        .unwrap();
    let deterministic_patch = deterministic_edit.patch().unwrap();
    let mut deterministic = Vec::new();
    deterministic_package
        .apply_glossary_text_batch(&deterministic_patch.clone().into())
        .and_then(|committed| Ok(committed.package().write_to(&mut deterministic)?))
        .unwrap();
    assert_eq!(output, deterministic);

    let recaptured = open(&output)
        .glossary_text_batch_snapshot(glossary_triplet())
        .unwrap();
    let committed = publication
        .package()
        .glossary_text_batch_snapshot(glossary_triplet())
        .unwrap();
    assert_eq!(committed.part_fingerprint(), recaptured.part_fingerprint());
    let inverse_target = publication
        .package()
        .apply_glossary_text_batch(&publication.patch().inverse())
        .unwrap()
        .package()
        .glossary_text_batch_snapshot(glossary_triplet())
        .unwrap();
    assert_eq!(
        inverse_target.extract_text(0).unwrap().as_deref(),
        Some("target")
    );
    assert_eq!(
        inverse_target.extract_text(2).unwrap().as_deref(),
        Some("third")
    );
}

#[test]
fn glossary_batch_rejects_empty_alias_missing_and_ambiguous_selectors() {
    assert!(story_text::GlossaryBatchSelector::try_new(Vec::new()).is_err());
    assert!(
        story_text::GlossaryBatchSelector::try_new(vec![story_text::GlossarySelector::ByIndex(0)])
            .is_ok()
    );
    assert!(
        story_text::GlossaryBatchSelector::try_new(vec![
            story_text::GlossarySelector::ByIndex(0),
            story_text::GlossarySelector::ByName("Alpha".to_owned()),
        ])
        .is_ok()
    );
    assert!(
        open(&batch_source())
            .glossary_text_batch_snapshot(
                story_text::GlossaryBatchSelector::try_new(vec![
                    story_text::GlossarySelector::ByIndex(0),
                    story_text::GlossarySelector::ByName("Alpha".to_owned()),
                ])
                .unwrap()
            )
            .is_err()
    );
    assert!(
        open(&batch_source())
            .glossary_text_batch_snapshot(
                story_text::GlossaryBatchSelector::try_new(vec![
                    story_text::GlossarySelector::ByName("Alpha".to_owned()),
                    story_text::GlossarySelector::ByNameAndId {
                        name: "Alpha".to_owned(),
                        id: Id::new(ALPHA_ID).unwrap(),
                    },
                ])
                .unwrap()
            )
            .is_err()
    );
    assert!(
        open(&batch_source())
            .glossary_text_batch_snapshot(
                story_text::GlossaryBatchSelector::try_new(vec![
                    story_text::GlossarySelector::ByName("Missing".to_owned()),
                    story_text::GlossarySelector::ById(Id::new(BETA_ID).unwrap()),
                ])
                .unwrap()
            )
            .is_err()
    );
    assert!(
        open(&fixture(Dialect::Transitional, Mutation::DuplicateName))
            .glossary_text_batch_snapshot(
                story_text::GlossaryBatchSelector::try_new(vec![
                    story_text::GlossarySelector::ByName("Alpha".to_owned()),
                    story_text::GlossarySelector::ById(Id::new(BETA_ID).unwrap()),
                ])
                .unwrap()
            )
            .is_err()
    );
    assert!(
        open(&fixture(Dialect::Transitional, Mutation::DuplicateId))
            .glossary_text_batch_snapshot(
                story_text::GlossaryBatchSelector::try_new(vec![
                    story_text::GlossarySelector::ById(Id::new(ALPHA_ID).unwrap()),
                    story_text::GlossarySelector::ByName("Gamma".to_owned()),
                ])
                .unwrap()
            )
            .is_err()
    );
}

#[test]
fn glossary_batch_noop_inverse_and_limits_are_atomic() {
    let source = batch_source();
    let package = open(&source);
    let snapshot = package
        .glossary_text_batch_snapshot(glossary_triplet())
        .unwrap();
    let patch = snapshot
        .raw_tree(snapshot.limits_value())
        .unwrap()
        .patch()
        .unwrap();
    assert!(patch.is_noop());
    let mut exact = Vec::new();
    package
        .apply_glossary_text_batch(&patch.clone().into())
        .and_then(|committed| Ok(committed.package().write_to(&mut exact)?))
        .unwrap();
    assert_eq!(exact, source);

    let package = open(&source);
    let snapshot = package
        .glossary_text_batch_snapshot(glossary_triplet())
        .unwrap();
    let mut edit = snapshot.raw_tree(snapshot.limits_value()).unwrap();
    edit.replace_paragraph_text(0, Position::new(0), "changed")
        .unwrap();
    let patch = edit.patch().unwrap();
    let mut output = Vec::new();
    let publication = package
        .apply_glossary_text_batch(&patch.clone().into())
        .and_then(|committed| {
            committed.package().write_to(&mut output)?;
            Ok(committed)
        })
        .unwrap();
    let mut restored = Vec::new();
    open(&output)
        .apply_glossary_text_batch(&publication.patch().inverse())
        .and_then(|committed| Ok(committed.package().write_to(&mut restored)?))
        .unwrap();
    assert_eq!(restored, source);

    let count_limited = story_text::Limits::default()
        .with_secondary_entry_limits(1, 16 * 1024, 256)
        .unwrap();
    assert!(
        open(&source)
            .with_semantic_limits(litchi_docx::SemanticLimits::default().story_text(count_limited))
            .glossary_text_batch_snapshot(glossary_triplet())
            .is_err()
    );
    let replacement_limited =
        story_text::Limits::new(16 * 1024, 8, 10_000, 64, 32 * 1024, 1, 4096).unwrap();
    let snapshot = open(&source)
        .with_semantic_limits(
            litchi_docx::SemanticLimits::default().story_text(replacement_limited),
        )
        .glossary_text_batch_snapshot(glossary_triplet())
        .unwrap();
    let mut edit = snapshot.raw_tree(snapshot.limits_value()).unwrap();
    edit.replace_paragraph_text(0, Position::new(0), "one")
        .unwrap();
    assert!(
        edit.replace_paragraph_text(1, Position::new(0), "two")
            .is_err()
    );
    assert!(edit.projected().is_ok());
    let output_limited = story_text::Limits::new(16 * 1024, 8, 10_000, 64, 1, 4096, 4096).unwrap();
    assert!(matches!(
        open(&source)
            .with_semantic_limits(litchi_docx::SemanticLimits::default().story_text(output_limited))
            .glossary_text_batch_snapshot(glossary_triplet()),
        Err(story_text::Error::Limit {
            kind: story_text::LimitKind::GlossaryBatchEntryBytes,
            maximum: 1,
            ..
        })
    ));
}

#[test]
fn glossary_batch_rejects_repeated_targets_and_aggregate_limits() {
    let source = batch_source();
    let package = open(&source);
    let snapshot = package
        .glossary_text_batch_snapshot(glossary_triplet())
        .unwrap();
    let mut edit = snapshot.raw_tree(snapshot.limits_value()).unwrap();
    edit.replace_paragraph_text(0, Position::new(0), "one")
        .unwrap();
    assert!(
        edit.replace_paragraph_text(0, Position::new(0), "again")
            .is_err()
    );
    assert_eq!(
        edit.projected()
            .unwrap()
            .extract_text(0)
            .unwrap()
            .as_deref(),
        Some("one")
    );

    let replacement_limited =
        story_text::Limits::new(16 * 1024, 8, 10_000, 64, 32 * 1024, 8, 5).unwrap();
    let snapshot = open(&source)
        .with_semantic_limits(
            litchi_docx::SemanticLimits::default().story_text(replacement_limited),
        )
        .glossary_text_batch_snapshot(glossary_triplet())
        .unwrap();
    let mut edit = snapshot.raw_tree(snapshot.limits_value()).unwrap();
    edit.replace_paragraph_text(0, Position::new(0), "one")
        .unwrap();
    assert!(
        edit.replace_paragraph_text(1, Position::new(0), "four")
            .is_err()
    );

    let entry_limited = story_text::Limits::default()
        .with_secondary_entry_limits(3, 1, 256)
        .unwrap();
    assert!(matches!(
        open(&source)
            .with_semantic_limits(litchi_docx::SemanticLimits::default().story_text(entry_limited))
            .glossary_text_batch_snapshot(glossary_triplet()),
        Err(story_text::Error::Limit { .. })
    ));

    let glossary_bytes = part(&source, "/word/glossary/document.xml").len();
    let output_limited =
        story_text::Limits::new(16 * 1024, 8, 10_000, 64, glossary_bytes - 1, 8, 4096).unwrap();
    let snapshot = open(&source)
        .with_semantic_limits(litchi_docx::SemanticLimits::default().story_text(output_limited))
        .glossary_text_batch_snapshot(glossary_triplet())
        .unwrap();
    let mut edit = snapshot.raw_tree(snapshot.limits_value()).unwrap();
    edit.replace_paragraph_text(0, Position::new(0), "x".repeat(512))
        .unwrap();
    assert!(edit.patch().is_err());
}

#[test]
fn glossary_batch_rejects_stale_managed_cancelled_signed_and_sink_paths_and_accepts_matching_copy()
{
    let source = batch_source();
    let versioned = Arc::new(VersionedSource {
        bytes: source.clone(),
        revision: std::sync::atomic::AtomicU64::new(0),
    });
    let package = source_backed::Package::from_read_at(versioned.clone()).unwrap();
    let snapshot = package
        .glossary_text_batch_snapshot(glossary_triplet())
        .unwrap();
    let mut edit = snapshot.raw_tree(snapshot.limits_value()).unwrap();
    edit.replace_paragraph_text(0, Position::new(0), "changed")
        .unwrap();
    let patch = edit.patch().unwrap();
    let noop = snapshot
        .raw_tree(snapshot.limits_value())
        .unwrap()
        .patch()
        .unwrap();
    assert!(noop.is_noop());
    versioned
        .revision
        .fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let mut refused = Vec::new();
    assert!(
        package
            .apply_glossary_text_batch(&patch.clone().into())
            .and_then(|committed| {
                committed.package().write_to(&mut refused)?;
                Ok(committed)
            })
            .is_err()
    );
    assert!(refused.is_empty());
    let foreign = open(&source)
        .glossary_text_batch_snapshot(glossary_triplet())
        .unwrap();
    let applied = patch.apply(&foreign).unwrap();
    assert_eq!(applied.extract_text(0).unwrap().as_deref(), Some("changed"));

    let managed = source_backed::Package::from_read_at_with_options(
        Arc::new(OwnedSource::new(source.clone())),
        litchi_docx::OpenOptions::default()
            .source_limits(litchi_docx::SourceLimits::default())
            .execution_context(litchi_core::ExecutionContext::new(
                litchi_core::Budget::root(
                    "source-backed-glossary-batch-test",
                    litchi_core::Limits::new(
                        2 * source.len() as u64,
                        u64::MAX,
                        u64::MAX,
                        u64::MAX,
                        u64::MAX,
                        u64::MAX,
                    ),
                ),
                litchi_core::CancellationSource::pair().1,
                litchi_core::ExecutionLimits::new(
                    std::num::NonZeroUsize::MIN,
                    std::num::NonZeroUsize::MIN,
                    std::num::NonZeroU64::new(2 * source.len() as u64).unwrap(),
                    0,
                )
                .unwrap(),
            )),
    )
    .unwrap();
    let managed_snapshot = managed
        .glossary_text_batch_snapshot(glossary_triplet())
        .unwrap();
    assert!(
        managed_snapshot
            .raw_tree(managed_snapshot.limits_value())
            .is_err()
    );

    let (cancel_source, token) = litchi_core::CancellationSource::pair();
    let cancelled = source_backed::Package::from_read_at_with_options(
        Arc::new(OwnedSource::new(source.clone())),
        litchi_docx::OpenOptions::default()
            .source_limits(litchi_docx::SourceLimits::default())
            .execution_context(litchi_core::ExecutionContext::new(
                litchi_core::Budget::root(
                    "source-backed-glossary-batch-cancel-test",
                    litchi_core::Limits::new(
                        2 * source.len() as u64,
                        u64::MAX,
                        u64::MAX,
                        u64::MAX,
                        u64::MAX,
                        u64::MAX,
                    ),
                ),
                token,
                litchi_core::ExecutionLimits::new(
                    std::num::NonZeroUsize::MIN,
                    std::num::NonZeroUsize::MIN,
                    std::num::NonZeroU64::new(2 * source.len() as u64).unwrap(),
                    0,
                )
                .unwrap(),
            )),
    )
    .unwrap();
    let captured = cancelled
        .glossary_text_batch_snapshot(glossary_triplet())
        .unwrap();
    cancel_source.cancel();
    assert!(captured.entry_count().is_err());
    assert!(matches!(
        patch.apply(&captured),
        Err(story_text::Error::Document(litchi_docx::Error::Opc(
            litchi_opc::OpcError::Cancelled
        )))
    ));
    assert!(matches!(
        noop.apply(&captured),
        Err(story_text::Error::Document(litchi_docx::Error::Opc(
            litchi_opc::OpcError::Cancelled
        )))
    ));
    assert!(
        cancelled
            .glossary_text_batch_snapshot(glossary_triplet())
            .is_err()
    );

    let signed = open(&fixture(Dialect::Transitional, Mutation::Signed));
    let signed_snapshot = signed
        .glossary_text_batch_snapshot(glossary_triplet())
        .unwrap();
    let noop = signed_snapshot
        .raw_tree(signed_snapshot.limits_value())
        .unwrap()
        .patch()
        .unwrap();
    let mut exact = Vec::new();
    signed
        .apply_glossary_text_batch(&noop.clone().into())
        .and_then(|committed| Ok(committed.package().write_to(&mut exact)?))
        .unwrap();
    assert_eq!(exact, fixture(Dialect::Transitional, Mutation::Signed));

    let signed = open(&fixture(Dialect::Transitional, Mutation::Signed));
    let signed_snapshot = signed
        .glossary_text_batch_snapshot(glossary_triplet())
        .unwrap();
    let mut edit = signed_snapshot
        .raw_tree(signed_snapshot.limits_value())
        .unwrap();
    edit.replace_paragraph_text(0, Position::new(0), "changed")
        .unwrap();
    let patch = edit.patch().unwrap();
    let mut refused = Vec::new();
    assert!(
        signed
            .apply_glossary_text_batch(&patch.clone().into())
            .and_then(|committed| {
                committed.package().write_to(&mut refused)?;
                Ok(committed)
            })
            .is_err()
    );
    assert!(refused.is_empty());

    let package = open(&source);
    let snapshot = package
        .glossary_text_batch_snapshot(glossary_triplet())
        .unwrap();
    let mut edit = snapshot.raw_tree(snapshot.limits_value()).unwrap();
    edit.replace_paragraph_text(0, Position::new(0), "changed")
        .unwrap();
    let patch = edit.patch().unwrap();
    assert!(
        package
            .apply_glossary_text_batch(&patch.clone().into())
            .and_then(|committed| {
                committed.package().write_to(PartialWriter { writes: 0 })?;
                Ok(committed)
            })
            .is_err()
    );
}

#[allow(dead_code)]
fn _identity_types_are_constructible() {
    let _ = Name::new("Alpha").unwrap();
    let _ = Id::new(ALPHA_ID).unwrap();
}
