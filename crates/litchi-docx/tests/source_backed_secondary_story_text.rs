use std::io::{self, Write};
use std::num::{NonZeroU64, NonZeroUsize};
use std::sync::Arc;
use std::sync::atomic::{AtomicU64, Ordering};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits as CoreLimits,
    OwnedSource, Position, ReadAt, SourceVersion,
};
use litchi_docx::source_backed::{self, StorySelector};
use litchi_docx::{Error, ReadLimits};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcPackage, PackURI, PackageWriter, Part, TargetMode};

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const STRICT_W: &str = "http://purl.oclc.org/ooxml/wordprocessingml/main";
const STRICT_OFFICE_DOCUMENT: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument";
const STRICT_FOOTNOTES: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/footnotes";
const STRICT_ENDNOTES: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/endnotes";
const STRICT_COMMENTS: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/comments";
const STRICT_HEADER: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/header";
const STRICT_FOOTER: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/footer";

#[derive(Clone, Copy)]
enum Dialect {
    Transitional,
    Strict,
}

impl Dialect {
    const fn word(self) -> &'static str {
        match self {
            Self::Transitional => W,
            Self::Strict => STRICT_W,
        }
    }

    const fn office_document(self) -> &'static str {
        match self {
            Self::Transitional => rt::OFFICE_DOCUMENT,
            Self::Strict => STRICT_OFFICE_DOCUMENT,
        }
    }

    const fn footnotes(self) -> &'static str {
        match self {
            Self::Transitional => rt::FOOTNOTES,
            Self::Strict => STRICT_FOOTNOTES,
        }
    }

    const fn endnotes(self) -> &'static str {
        match self {
            Self::Transitional => rt::ENDNOTES,
            Self::Strict => STRICT_ENDNOTES,
        }
    }

    const fn comments(self) -> &'static str {
        match self {
            Self::Transitional => rt::COMMENTS,
            Self::Strict => STRICT_COMMENTS,
        }
    }

    const fn header(self) -> &'static str {
        match self {
            Self::Transitional => rt::HEADER,
            Self::Strict => STRICT_HEADER,
        }
    }

    const fn footer(self) -> &'static str {
        match self {
            Self::Transitional => rt::FOOTER,
            Self::Strict => STRICT_FOOTER,
        }
    }

    const fn opposite_word(self) -> &'static str {
        match self {
            Self::Transitional => STRICT_W,
            Self::Strict => W,
        }
    }
}

#[derive(Clone, Copy)]
enum Mutation {
    Valid,
    DuplicateFootnote,
    DuplicateComment,
    TwoParagraphFootnote,
    WrongFootnoteContentType,
    ExternalFootnoteRelationship,
    MissingFootnotePart,
    SecondFootnoteRelationship,
    FootnoteRelationshipDialectMismatch,
    SecondaryOwnsStoryRelationship,
    MainStrictStoryRelationship,
    HeaderRootDialectMismatch,
    FooterRootDialectMismatch,
    HeaderOwnsStoryRelationship,
    FooterOwnsStoryRelationship,
    DeepNamespaceFootnote,
    LongNamespaceFootnote,
    LargeTextFootnote,
    EscapedTextFootnote,
    UnknownNamedEntityFootnote,
    InstrTextReferenceFootnote,
    MarkupHeavyFootnote,
    Signed,
}

fn compact_xml(xml: String) -> Vec<u8> {
    let bytes = xml.as_bytes();
    let mut compact = Vec::with_capacity(bytes.len());
    let mut index = 0;
    while index < bytes.len() {
        if bytes[index].is_ascii_whitespace() && compact.last() == Some(&b'>') {
            let start = index;
            while index < bytes.len() && bytes[index].is_ascii_whitespace() {
                index += 1;
            }
            if index < bytes.len() && bytes[index] == b'<' {
                continue;
            }
            compact.extend_from_slice(&bytes[start..index]);
        } else {
            compact.push(bytes[index]);
            index += 1;
        }
    }
    compact
}

fn deep_footnotes_xml(word: &str) -> String {
    let namespace = format!("urn:deep-{}", "q".repeat(16 * 1024));
    let mut xml = format!(r#"<w:footnotes xmlns:w="{word}"><w:footnote w:id="41">"#);
    for index in 0..32 {
        xml.push_str(&format!(
            r#"<n{index}:opaque xmlns:n{index}="{namespace}">"#
        ));
    }
    xml.push_str(r#"<w:p><w:r><w:t>deep</w:t></w:r></w:p>"#);
    for index in (0..32).rev() {
        xml.push_str(&format!(r#"</n{index}:opaque>"#));
    }
    xml.push_str(r#"</w:footnote></w:footnotes>"#);
    xml
}

fn long_namespace_footnotes_xml(word: &str) -> String {
    let prefix = format!("long{}", "p".repeat(1024));
    let value = format!("urn:long-{}", "q".repeat(4096));
    format!(
        r#"<w:footnotes xmlns:w="{word}"><w:footnote w:id="41"><{prefix}:opaque xmlns:{prefix}="{value}"><w:p><w:r><w:t>long</w:t></w:r></w:p></{prefix}:opaque></w:footnote></w:footnotes>"#
    )
}

fn large_text_footnotes_xml(word: &str) -> String {
    let text = "t".repeat(24 * 1024);
    format!(
        r#"<w:footnotes xmlns:w="{word}"><w:footnote w:id="41"><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:footnote></w:footnotes>"#
    )
}

fn escaped_text_footnotes_xml(word: &str) -> String {
    let text = "&#x41;&amp;&lt;&gt;&quot;&apos;".repeat(64);
    format!(
        r#"<w:footnotes xmlns:w="{word}"><w:footnote w:id="41"><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:footnote></w:footnotes>"#
    )
}

fn unknown_named_entity_footnotes_xml(word: &str) -> String {
    format!(
        r#"<w:footnotes xmlns:w="{word}"><w:footnote w:id="41"><w:p><w:r><w:t>&unknown;</w:t></w:r></w:p></w:footnote></w:footnotes>"#
    )
}

fn instr_text_reference_footnotes_xml(word: &str) -> String {
    format!(
        r#"<w:footnotes xmlns:w="{word}"><w:footnote w:id="41"><w:p><w:r><w:instrText>&amp;</w:instrText><w:t>visible</w:t></w:r></w:p></w:footnote></w:footnotes>"#
    )
}

fn markup_heavy_footnotes_xml(word: &str) -> String {
    let opaque = "<x:opaque/>".repeat(64);
    format!(
        r#"<w:footnotes xmlns:w="{word}" xmlns:x="urn:opaque"><w:footnote w:id="41"><w:p>{opaque}<w:r><w:t>fit</w:t></w:r></w:p></w:footnote></w:footnotes>"#
    )
}

fn fixture(dialect: Dialect, mutation: Mutation) -> Vec<u8> {
    let word = dialect.word();
    let main_xml = format!(
        r#"<w:document xmlns:w="{word}" xmlns:x="urn:opaque"><w:body>
<x:opaque/><w:p><w:r><w:t>main</w:t></w:r><w:r><w:footnoteReference w:id="41"/><w:endnoteReference w:id="51"/></w:r></w:p>
<w:p><w:commentRangeStart w:id="61"/><w:r><w:t>anchor</w:t></w:r><w:commentRangeEnd w:id="61"/><w:r><w:commentReference w:id="61"/></w:r></w:p>
</w:body></w:document>"#
    );
    let footnotes_xml = format!(
        r#"<w:footnotes xmlns:w="{word}" xmlns:x="urn:opaque"><x:opaque/>
<w:footnote w:id="-1" w:type="separator"/><w:footnote w:id="0" w:type="continuationSeparator"/>
<w:footnote w:id="41">
<!-- selected-footnote-comment -->
<x:opaque/>
<w:p><w:r><w:t>footnote-target</w:t></w:r></w:p>
<!-- trailing-footnote-comment -->
<x:opaque/>
</w:footnote>
<w:footnote w:id="42"><w:p><w:r><w:t>footnote-sibling</w:t></w:r></w:p></w:footnote>
</w:footnotes>"#
    );
    let footnotes_xml = if matches!(mutation, Mutation::DuplicateFootnote) {
        footnotes_xml.replace(
            "<w:footnote w:id=\"42\">",
            "<w:footnote w:id=\"41\"><w:p><w:r><w:t>duplicate</w:t></w:r></w:p></w:footnote><w:footnote w:id=\"42\">",
        )
    } else {
        footnotes_xml
    };
    let footnotes_xml = if matches!(mutation, Mutation::TwoParagraphFootnote) {
        footnotes_xml.replace(
            "</w:p>\n<!-- trailing-footnote-comment -->",
            "</w:p><w:p><w:r><w:t>footnote-second</w:t></w:r></w:p>\n<!-- trailing-footnote-comment -->",
        )
    } else {
        footnotes_xml
    };
    let footnotes_xml = if matches!(mutation, Mutation::DeepNamespaceFootnote) {
        deep_footnotes_xml(word)
    } else {
        footnotes_xml
    };
    let footnotes_xml = if matches!(mutation, Mutation::LongNamespaceFootnote) {
        long_namespace_footnotes_xml(word)
    } else {
        footnotes_xml
    };
    let footnotes_xml = if matches!(mutation, Mutation::LargeTextFootnote) {
        large_text_footnotes_xml(word)
    } else {
        footnotes_xml
    };
    let footnotes_xml = if matches!(mutation, Mutation::EscapedTextFootnote) {
        escaped_text_footnotes_xml(word)
    } else {
        footnotes_xml
    };
    let footnotes_xml = if matches!(mutation, Mutation::UnknownNamedEntityFootnote) {
        unknown_named_entity_footnotes_xml(word)
    } else {
        footnotes_xml
    };
    let footnotes_xml = if matches!(mutation, Mutation::InstrTextReferenceFootnote) {
        instr_text_reference_footnotes_xml(word)
    } else {
        footnotes_xml
    };
    let footnotes_xml = if matches!(mutation, Mutation::MarkupHeavyFootnote) {
        markup_heavy_footnotes_xml(word)
    } else {
        footnotes_xml
    };
    let endnotes_xml = format!(
        r#"<w:endnotes xmlns:w="{word}" xmlns:x="urn:opaque"><x:opaque/>
<w:endnote w:id="-1" w:type="separator"/><w:endnote w:id="0" w:type="continuationSeparator"/>
<w:endnote w:id="51"><w:p><w:r><w:t>endnote-target</w:t></w:r></w:p></w:endnote>
<w:endnote w:id="52"><w:p><w:r><w:t>endnote-sibling</w:t></w:r></w:p></w:endnote>
</w:endnotes>"#
    );
    let comments_xml = format!(
        r#"<w:comments xmlns:w="{word}" xmlns:x="urn:opaque"><x:opaque/>
<w:comment w:id="61" w:author="Alice"><w:p><w:r><w:t>comment-target</w:t></w:r></w:p><x:opaque/></w:comment>
<w:comment w:id="62" w:author="Bob"><w:p><w:r><w:t>comment-sibling</w:t></w:r></w:p></w:comment>
</w:comments>"#
    );
    let comments_xml = if matches!(mutation, Mutation::DuplicateComment) {
        comments_xml.replace(
            "<w:comment w:id=\"62\"",
            "<w:comment w:id=\"61\"><w:p><w:r><w:t>duplicate</w:t></w:r></w:p></w:comment><w:comment w:id=\"62\"",
        )
    } else {
        comments_xml
    };

    let mut package = OpcPackage::new();
    let mut main = BlobPart::new(
        PackURI::new("/word/document.xml").unwrap(),
        ct::WML_DOCUMENT_MAIN.to_owned(),
        compact_xml(main_xml),
    );
    let mode = if matches!(mutation, Mutation::ExternalFootnoteRelationship) {
        TargetMode::External
    } else {
        TargetMode::Internal
    };
    let footnote_relationship = if matches!(mutation, Mutation::FootnoteRelationshipDialectMismatch)
    {
        match dialect {
            Dialect::Transitional => STRICT_FOOTNOTES,
            Dialect::Strict => rt::FOOTNOTES,
        }
    } else {
        dialect.footnotes()
    };
    main.rels_mut()
        .try_add_relationship(
            footnote_relationship.to_owned(),
            "footnotes.xml".to_owned(),
            "rFootnotes".to_owned(),
            mode,
        )
        .unwrap();
    main.rels_mut()
        .try_add_relationship(
            dialect.endnotes().to_owned(),
            "endnotes.xml".to_owned(),
            "rEndnotes".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
    main.rels_mut()
        .try_add_relationship(
            dialect.comments().to_owned(),
            "comments.xml".to_owned(),
            "rComments".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
    if matches!(mutation, Mutation::SecondFootnoteRelationship) {
        main.rels_mut()
            .try_add_relationship(
                dialect.footnotes().to_owned(),
                "footnotes2.xml".to_owned(),
                "rFootnotes2".to_owned(),
                TargetMode::Internal,
            )
            .unwrap();
    }
    if matches!(mutation, Mutation::MainStrictStoryRelationship) {
        main.rels_mut()
            .try_add_relationship(
                STRICT_HEADER.to_owned(),
                "header1.xml".to_owned(),
                "rStrictHeader".to_owned(),
                TargetMode::Internal,
            )
            .unwrap();
    }
    if matches!(
        mutation,
        Mutation::HeaderRootDialectMismatch | Mutation::HeaderOwnsStoryRelationship
    ) {
        main.rels_mut()
            .try_add_relationship(
                dialect.header().to_owned(),
                "header1.xml".to_owned(),
                "rHeader".to_owned(),
                TargetMode::Internal,
            )
            .unwrap();
    }
    if matches!(
        mutation,
        Mutation::FooterRootDialectMismatch | Mutation::FooterOwnsStoryRelationship
    ) || matches!(mutation, Mutation::MainStrictStoryRelationship)
    {
        main.rels_mut()
            .try_add_relationship(
                dialect.footer().to_owned(),
                "footer1.xml".to_owned(),
                "rFooter".to_owned(),
                TargetMode::Internal,
            )
            .unwrap();
    }
    package.try_add_part(Box::new(main)).unwrap();
    if !matches!(mutation, Mutation::MissingFootnotePart) {
        let mut footnotes = BlobPart::new(
            PackURI::new("/word/footnotes.xml").unwrap(),
            if matches!(mutation, Mutation::WrongFootnoteContentType) {
                ct::WML_ENDNOTES.to_owned()
            } else {
                ct::WML_FOOTNOTES.to_owned()
            },
            compact_xml(footnotes_xml),
        );
        if matches!(mutation, Mutation::SecondaryOwnsStoryRelationship) {
            let relationship_type = match dialect {
                Dialect::Transitional => rt::HEADER,
                Dialect::Strict => STRICT_HEADER,
            };
            footnotes
                .rels_mut()
                .try_add_relationship(
                    relationship_type.to_owned(),
                    "header1.xml".to_owned(),
                    "rNestedHeader".to_owned(),
                    TargetMode::Internal,
                )
                .unwrap();
        }
        package.try_add_part(Box::new(footnotes)).unwrap();
    }
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new("/word/endnotes.xml").unwrap(),
            ct::WML_ENDNOTES.to_owned(),
            compact_xml(endnotes_xml),
        )))
        .unwrap();
    if matches!(mutation, Mutation::SecondaryOwnsStoryRelationship) {
        package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new("/word/header1.xml").unwrap(),
                ct::WML_HEADER.to_owned(),
                compact_xml(format!(r#"<w:hdr xmlns:w="{word}"/>"#)),
            )))
            .unwrap();
    }
    if matches!(mutation, Mutation::MainStrictStoryRelationship) {
        package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new("/word/header1.xml").unwrap(),
                ct::WML_HEADER.to_owned(),
                compact_xml(format!(r#"<w:hdr xmlns:w="{STRICT_W}"/>"#)),
            )))
            .unwrap();
    }
    if matches!(
        mutation,
        Mutation::HeaderRootDialectMismatch | Mutation::HeaderOwnsStoryRelationship
    ) {
        let header_word = if matches!(mutation, Mutation::HeaderRootDialectMismatch) {
            dialect.opposite_word()
        } else {
            word
        };
        let mut header = BlobPart::new(
            PackURI::new("/word/header1.xml").unwrap(),
            ct::WML_HEADER.to_owned(),
            compact_xml(format!(r#"<w:hdr xmlns:w="{header_word}"/>"#)),
        );
        if matches!(mutation, Mutation::HeaderOwnsStoryRelationship) {
            header
                .rels_mut()
                .try_add_relationship(
                    dialect.footnotes().to_owned(),
                    "footnotes.xml".to_owned(),
                    "rNestedFootnotes".to_owned(),
                    TargetMode::Internal,
                )
                .unwrap();
        }
        package.try_add_part(Box::new(header)).unwrap();
    }
    if matches!(
        mutation,
        Mutation::MainStrictStoryRelationship
            | Mutation::FooterRootDialectMismatch
            | Mutation::FooterOwnsStoryRelationship
    ) {
        let footer_word = if matches!(mutation, Mutation::FooterRootDialectMismatch) {
            dialect.opposite_word()
        } else {
            word
        };
        let mut footer = BlobPart::new(
            PackURI::new("/word/footer1.xml").unwrap(),
            ct::WML_FOOTER.to_owned(),
            compact_xml(format!(r#"<w:ftr xmlns:w="{footer_word}"/>"#)),
        );
        if matches!(mutation, Mutation::FooterOwnsStoryRelationship) {
            footer
                .rels_mut()
                .try_add_relationship(
                    dialect.endnotes().to_owned(),
                    "endnotes.xml".to_owned(),
                    "rNestedEndnotes".to_owned(),
                    TargetMode::Internal,
                )
                .unwrap();
        }
        package.try_add_part(Box::new(footer)).unwrap();
    }
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new("/word/comments.xml").unwrap(),
            ct::WML_COMMENTS.to_owned(),
            compact_xml(comments_xml),
        )))
        .unwrap();
    if matches!(mutation, Mutation::SecondFootnoteRelationship) {
        package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new("/word/footnotes2.xml").unwrap(),
                ct::WML_FOOTNOTES.to_owned(),
                compact_xml(format!(r#"<w:footnotes xmlns:w="{word}"/>"#)),
            )))
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
    PackageWriter::to_bytes(&package).unwrap()
}

fn open(bytes: &[u8]) -> source_backed::Package {
    source_backed::Package::from_read_at(Arc::new(OwnedSource::new(bytes.to_vec()))).unwrap()
}

fn part_bytes(bytes: &[u8], name: &str) -> Vec<u8> {
    let package = OpcPackage::from_bytes(bytes).unwrap();
    package
        .get_part(&PackURI::new(name).unwrap())
        .unwrap()
        .blob()
        .to_vec()
}

fn changed_part(source: &[u8], name: &str, from: &str, to: &str) -> Vec<u8> {
    let mut expected = part_bytes(source, name);
    let from = from.as_bytes();
    let index = expected
        .windows(from.len())
        .position(|window| window == from)
        .expect("fixture text");
    expected.splice(index..index + from.len(), to.as_bytes().iter().copied());
    expected
}

fn exercise_change(
    dialect: Dialect,
    selector: StorySelector,
    part: &str,
    before: &str,
    after: &str,
) {
    let source = fixture(dialect, Mutation::Valid);
    let package = open(&source);
    let snapshot = package.story_text_snapshot(selector.clone()).unwrap();
    assert_eq!(snapshot.paragraph_text(0).unwrap().as_deref(), Some(before));
    let mut edit = snapshot.edit().unwrap();
    edit.replace_paragraph_text(Position::new(0), after)
        .unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(commit.snapshot().extract_text().unwrap(), after);

    let mut published = Vec::new();
    let publication = package
        .publish_story_text_commit_to_stream(&mut published, &commit)
        .unwrap();
    let reopened = open(&published);
    assert_eq!(
        reopened
            .story_text_snapshot(selector.clone())
            .unwrap()
            .paragraph_text(0)
            .unwrap()
            .as_deref(),
        Some(after)
    );
    assert_eq!(
        part_bytes(&published, part),
        changed_part(&source, part, before, after)
    );
    assert_eq!(
        part_bytes(&published, "/word/document.xml"),
        part_bytes(&source, "/word/document.xml")
    );
    assert_eq!(publication.snapshot().selector(), selector);
}

#[test]
fn secondary_stories_select_edit_and_preserve_siblings_root_and_anchors() {
    for dialect in [Dialect::Transitional, Dialect::Strict] {
        exercise_change(
            dialect,
            StorySelector::footnote(41),
            "/word/footnotes.xml",
            "footnote-target",
            "footnote-changed",
        );
        exercise_change(
            dialect,
            StorySelector::endnote(51),
            "/word/endnotes.xml",
            "endnote-target",
            "endnote-changed",
        );
        exercise_change(
            dialect,
            StorySelector::comment(61),
            "/word/comments.xml",
            "comment-target",
            "comment-changed",
        );
    }
}

#[test]
fn secondary_entry_preserves_whitespace_comments_and_opaque_markup() {
    let source = fixture(Dialect::Transitional, Mutation::Valid);
    let package = open(&source);
    let snapshot = package
        .story_text_snapshot(StorySelector::footnote(41))
        .unwrap();
    let mut edit = snapshot.edit().unwrap();
    edit.replace_paragraph_text(Position::new(0), "footnote-preserved")
        .unwrap();
    let commit = edit.commit().unwrap();
    let mut published = Vec::new();
    package
        .publish_story_text_commit_to_stream(&mut published, &commit)
        .unwrap();

    let before_part = part_bytes(&source, "/word/footnotes.xml");
    let after_part = part_bytes(&published, "/word/footnotes.xml");
    let old_text = b"footnote-target";
    let new_text = b"footnote-preserved";
    let old_offset = before_part
        .windows(old_text.len())
        .position(|window| window == old_text)
        .unwrap();
    let new_offset = after_part
        .windows(new_text.len())
        .position(|window| window == new_text)
        .unwrap();
    assert_eq!(&before_part[..old_offset], &after_part[..new_offset]);
    assert_eq!(
        &before_part[old_offset + old_text.len()..],
        &after_part[new_offset + new_text.len()..]
    );
    assert!(
        after_part
            .windows(b"selected-footnote-comment".len())
            .any(|window| { window == b"selected-footnote-comment" })
    );
    assert!(
        after_part
            .windows(b"<x:opaque/>".len())
            .any(|window| window == b"<x:opaque/>")
    );
}

#[test]
fn secondary_noop_is_exact_and_inverse_restores_the_source() {
    let source = fixture(Dialect::Transitional, Mutation::Valid);
    let noop_package = open(&source);
    let noop_snapshot = noop_package
        .story_text_snapshot(StorySelector::comment(61))
        .unwrap();
    let mut noop = noop_snapshot.edit().unwrap();
    noop.replace_paragraph_text(Position::new(0), "comment-target")
        .unwrap();
    let noop_commit = noop.commit().unwrap();
    let mut exact = Vec::new();
    noop_package
        .publish_story_text_commit_to_stream(&mut exact, &noop_commit)
        .unwrap();
    assert_eq!(exact, source);

    let changed_package = open(&source);
    let changed_snapshot = changed_package
        .story_text_snapshot(StorySelector::comment(61))
        .unwrap();
    let mut edit = changed_snapshot.edit().unwrap();
    edit.replace_paragraph_text(Position::new(0), "comment-changed")
        .unwrap();
    let commit = edit.commit().unwrap();
    let mut published = Vec::new();
    let publication = changed_package
        .publish_story_text_commit_to_stream(&mut published, &commit)
        .unwrap();
    let mut restored = Vec::new();
    open(&published)
        .publish_story_text_inverse_to_stream(&mut restored, &publication)
        .unwrap();
    assert_eq!(restored, source);
    assert!(matches!(
        open(&source).publish_story_text_inverse_to_stream(Vec::new(), &publication),
        Err(source_backed::StoryTextError::ArtifactConflict)
    ));
}

#[test]
fn signed_secondary_change_is_refused_but_signed_noop_is_exact() {
    let source = fixture(Dialect::Transitional, Mutation::Signed);
    let noop_package = open(&source);
    let noop_snapshot = noop_package
        .story_text_snapshot(StorySelector::footnote(41))
        .unwrap();

    let mut noop = noop_snapshot.edit().unwrap();
    noop.replace_paragraph_text(Position::new(0), "footnote-target")
        .unwrap();
    let noop_commit = noop.commit().unwrap();
    let mut exact = Vec::new();
    noop_package
        .publish_story_text_commit_to_stream(&mut exact, &noop_commit)
        .unwrap();
    assert_eq!(exact, source);

    let changed_package = open(&source);
    let changed_snapshot = changed_package
        .story_text_snapshot(StorySelector::footnote(41))
        .unwrap();
    let mut edit = changed_snapshot.edit().unwrap();
    edit.replace_paragraph_text(Position::new(0), "signed-change")
        .unwrap();
    let commit = edit.commit().unwrap();
    assert!(
        changed_package
            .publish_story_text_commit_to_stream(Vec::new(), &commit)
            .is_err()
    );
}

#[test]
fn secondary_ids_cover_missing_reserved_and_duplicate_values() {
    let package = open(&fixture(Dialect::Transitional, Mutation::Valid));
    for selector in [
        StorySelector::footnote(99),
        StorySelector::footnote(0),
        StorySelector::footnote(u32::MAX),
        StorySelector::endnote(0),
        StorySelector::comment(99),
    ] {
        assert!(package.story_text_snapshot(selector).is_err());
    }
    assert!(
        open(&fixture(Dialect::Transitional, Mutation::DuplicateFootnote))
            .story_text_snapshot(StorySelector::footnote(41))
            .is_err()
    );
    assert!(
        open(&fixture(Dialect::Transitional, Mutation::DuplicateComment))
            .story_text_snapshot(StorySelector::comment(61))
            .is_err()
    );
}

#[test]
fn secondary_relationship_topology_and_content_type_are_checked() {
    for mutation in [
        Mutation::WrongFootnoteContentType,
        Mutation::ExternalFootnoteRelationship,
        Mutation::MissingFootnotePart,
        Mutation::SecondFootnoteRelationship,
    ] {
        assert!(
            open(&fixture(Dialect::Transitional, mutation))
                .story_text_snapshot(StorySelector::footnote(41))
                .is_err()
        );
    }
}

#[test]
fn secondary_relationship_dialect_mismatch_is_refused() {
    for dialect in [Dialect::Transitional, Dialect::Strict] {
        assert!(
            open(&fixture(
                dialect,
                Mutation::FootnoteRelationshipDialectMismatch
            ))
            .story_text_snapshot(StorySelector::footnote(41))
            .is_err()
        );
    }
}

#[test]
fn transitional_main_rejects_strict_header_relationship_metadata() {
    let source = fixture(Dialect::Transitional, Mutation::MainStrictStoryRelationship);
    let package = open(&source);
    for selector in [
        StorySelector::footnote(41),
        StorySelector::endnote(51),
        StorySelector::comment(61),
    ] {
        assert!(package.story_text_snapshot(selector).is_err());
    }
}

#[test]
fn package_wide_mixed_dialect_rejects_main_header_and_footer_capture() {
    let package = open(&fixture(
        Dialect::Transitional,
        Mutation::MainStrictStoryRelationship,
    ));
    for selector in [
        StorySelector::main(),
        StorySelector::header(0),
        StorySelector::footer(0),
    ] {
        assert!(package.story_text_snapshot(selector).is_err());
    }
}

#[test]
fn header_and_footer_root_dialect_mismatches_are_refused() {
    for dialect in [Dialect::Transitional, Dialect::Strict] {
        assert!(
            open(&fixture(dialect, Mutation::HeaderRootDialectMismatch))
                .story_text_snapshot(StorySelector::header(0))
                .is_err()
        );
        assert!(
            open(&fixture(dialect, Mutation::FooterRootDialectMismatch))
                .story_text_snapshot(StorySelector::footer(0))
                .is_err()
        );
    }
}

#[test]
fn header_and_footer_stories_cannot_own_word_story_relationships() {
    for dialect in [Dialect::Transitional, Dialect::Strict] {
        assert!(
            open(&fixture(dialect, Mutation::HeaderOwnsStoryRelationship))
                .story_text_snapshot(StorySelector::header(0))
                .is_err()
        );
        assert!(
            open(&fixture(dialect, Mutation::FooterOwnsStoryRelationship))
                .story_text_snapshot(StorySelector::footer(0))
                .is_err()
        );
    }
}

#[test]
fn secondary_story_cannot_own_a_word_story_relationship() {
    for dialect in [Dialect::Transitional, Dialect::Strict] {
        assert!(
            open(&fixture(dialect, Mutation::SecondaryOwnsStoryRelationship))
                .story_text_snapshot(StorySelector::footnote(41))
                .is_err()
        );
    }
}

#[test]
fn secondary_patches_reject_stale_and_foreign_snapshots() {
    let source = fixture(Dialect::Transitional, Mutation::Valid);
    let package = open(&source);
    let snapshot = package
        .story_text_snapshot(StorySelector::footnote(41))
        .unwrap();
    let mut edit = snapshot.edit().unwrap();
    edit.replace_paragraph_text(Position::new(0), "changed")
        .unwrap();
    let commit = edit.commit().unwrap();
    assert!(matches!(
        commit.patch().apply(commit.snapshot()),
        Err(source_backed::StoryTextError::StaleSource)
    ));
    let foreign = open(&source)
        .story_text_snapshot(StorySelector::footnote(41))
        .unwrap();
    assert!(matches!(
        commit.patch().apply(&foreign),
        Err(source_backed::StoryTextError::ForeignSource)
    ));
}

fn managed_open(bytes: Vec<u8>) -> (Budget, CancellationSource, source_backed::Package) {
    let memory = (bytes.len() as u64).saturating_mul(4).max(1);
    let budget = Budget::root(
        "source-backed-secondary-story-test",
        CoreLimits::new(memory, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
    );
    let (cancellation_source, cancellation) = CancellationSource::pair();
    let execution_limits = ExecutionLimits::new(
        NonZeroUsize::MIN,
        NonZeroUsize::MIN,
        NonZeroU64::new(memory).unwrap(),
        0,
    )
    .unwrap();
    let package = source_backed::Package::from_read_at_with_execution_context(
        Arc::new(OwnedSource::new(bytes)),
        ReadLimits::default(),
        ExecutionContext::new(budget.clone(), cancellation, execution_limits),
    )
    .unwrap();
    (budget, cancellation_source, package)
}

#[test]
fn managed_secondary_edits_refuse_and_cancellation_stops_selection() {
    let (_budget, cancellation, package) =
        managed_open(fixture(Dialect::Transitional, Mutation::Valid));
    let snapshot = package
        .story_text_snapshot(StorySelector::footnote(41))
        .unwrap();
    assert!(matches!(
        snapshot.edit(),
        Err(source_backed::StoryTextError::Document(
            Error::UnsafeEdit { .. }
        ))
    ));
    cancellation.cancel();
    assert!(
        package
            .story_text_snapshot(StorySelector::comment(61))
            .is_err()
    );
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

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(
            251,
            self.revision.load(Ordering::SeqCst),
        ))
    }
}

#[test]
fn secondary_publication_rejects_a_source_revision_change() {
    let source = fixture(Dialect::Transitional, Mutation::Valid);
    let versioned = Arc::new(VersionedSource {
        bytes: source,
        revision: AtomicU64::new(0),
    });
    let package = source_backed::Package::from_read_at(versioned.clone()).unwrap();
    let snapshot = package
        .story_text_snapshot(StorySelector::endnote(51))
        .unwrap();
    let mut edit = snapshot.edit().unwrap();
    edit.replace_paragraph_text(Position::new(0), "revision-change")
        .unwrap();
    let commit = edit.commit().unwrap();
    versioned.revision.fetch_add(1, Ordering::SeqCst);
    assert!(matches!(
        package.publish_story_text_commit_to_stream(Vec::new(), &commit),
        Err(source_backed::StoryTextError::Document(Error::Opc(
            litchi_opc::OpcError::SourceChanged { .. }
        )))
    ));
}

#[test]
fn secondary_limits_bound_xml_paragraphs_and_output() {
    let source = fixture(Dialect::Transitional, Mutation::Valid);
    let invalid_limits = source_backed::StoryTextLimits::new(4096, 0, 10_000, 64, 4096, 4096, 4096);
    assert!(invalid_limits.is_err());
    let paragraph_limit =
        source_backed::StoryTextLimits::new(4096, 1, 10_000, 64, 4096, 4096, 4096)
            .unwrap()
            .with_secondary_entry_limits(8, 4096, 256)
            .unwrap();
    assert!(matches!(
        open(&fixture(
            Dialect::Transitional,
            Mutation::TwoParagraphFootnote
        ))
        .story_text_snapshot_with_limits(StorySelector::footnote(41), paragraph_limit),
        Err(source_backed::StoryTextError::Limit {
            resource: "paragraphs",
            ..
        })
    ));
    let xml_limit =
        source_backed::StoryTextLimits::new(1, 10, 10_000, 64, 4096, 4096, 4096).unwrap();
    assert!(matches!(
        open(&source).story_text_snapshot_with_limits(StorySelector::footnote(41), xml_limit),
        Err(source_backed::StoryTextError::Limit {
            resource: "XML bytes",
            ..
        })
    ));
    let output_limit =
        source_backed::StoryTextLimits::new(4096, 10, 10_000, 64, 1, 4096, 4096).unwrap();
    let snapshot = open(&source)
        .story_text_snapshot_with_limits(StorySelector::footnote(41), output_limit)
        .unwrap();
    assert!(matches!(
        snapshot.extract_text(),
        Err(source_backed::StoryTextError::Limit { .. })
    ));
}

#[test]
fn deep_large_namespace_input_hits_a_finite_namespace_limit() {
    let limits =
        source_backed::StoryTextLimits::new(1024 * 1024, 10, 100_000, 64, 4096, 4096, 4096)
            .unwrap()
            .with_secondary_entry_limits(64, 1024 * 1024, 8)
            .unwrap();
    assert!(
        open(&fixture(
            Dialect::Transitional,
            Mutation::DeepNamespaceFootnote
        ))
        .story_text_snapshot_with_limits(StorySelector::footnote(41), limits)
        .is_err()
    );
}

#[test]
fn long_namespace_prefix_and_value_hit_namespace_byte_and_count_limits() {
    let namespace_limits =
        source_backed::StoryTextLimits::new(64 * 1024, 10, 100_000, 64, 4096, 4096, 4096)
            .unwrap()
            .with_secondary_entry_limits(64, 64 * 1024, 256)
            .unwrap()
            .with_max_namespace_bytes(2048)
            .unwrap();
    assert!(matches!(
        open(&fixture(
            Dialect::Transitional,
            Mutation::LongNamespaceFootnote,
        ))
        .story_text_snapshot_with_limits(StorySelector::footnote(41), namespace_limits),
        Err(source_backed::StoryTextError::Limit {
            resource: "namespace bytes",
            ..
        })
    ));

    let count_limits =
        source_backed::StoryTextLimits::new(1024 * 1024, 10, 100_000, 64, 4096, 4096, 4096)
            .unwrap()
            .with_secondary_entry_limits(64, 1024 * 1024, 8)
            .unwrap();
    assert!(
        open(&fixture(
            Dialect::Transitional,
            Mutation::DeepNamespaceFootnote,
        ))
        .story_text_snapshot_with_limits(StorySelector::footnote(41), count_limits)
        .is_err()
    );
}

#[test]
fn large_text_event_is_bounded_by_xml_and_decoded_output_limits() {
    let output_limits =
        source_backed::StoryTextLimits::new(64 * 1024, 10, 100_000, 64, 1024, 4096, 4096)
            .unwrap()
            .with_secondary_entry_limits(8, 64 * 1024, 256)
            .unwrap();
    let snapshot = open(&fixture(Dialect::Transitional, Mutation::LargeTextFootnote))
        .story_text_snapshot_with_limits(StorySelector::footnote(41), output_limits)
        .unwrap();
    assert!(matches!(
        snapshot.extract_text(),
        Err(source_backed::StoryTextError::Limit { .. })
    ));

    let xml_limits =
        source_backed::StoryTextLimits::new(1024, 10, 100_000, 64, 64 * 1024, 4096, 4096)
            .unwrap()
            .with_secondary_entry_limits(8, 1024, 256)
            .unwrap();
    assert!(
        open(&fixture(Dialect::Transitional, Mutation::LargeTextFootnote,))
            .story_text_snapshot_with_limits(StorySelector::footnote(41), xml_limits)
            .is_err()
    );
}

#[test]
fn allowed_references_stream_under_caller_bounds_and_unknown_named_refs_refuse() {
    let limits = source_backed::StoryTextLimits::new(32 * 1024, 10, 100_000, 64, 4096, 4096, 4096)
        .unwrap()
        .with_secondary_entry_limits(8, 32 * 1024, 256)
        .unwrap();
    let snapshot = open(&fixture(
        Dialect::Transitional,
        Mutation::EscapedTextFootnote,
    ))
    .story_text_snapshot_with_limits(StorySelector::footnote(41), limits)
    .unwrap();

    let expected = "A&<>\"'".repeat(64);
    let mut complete = Vec::new();
    let report = snapshot
        .write_text_to(
            &mut complete,
            litchi_core::TextOutputOptions::new("", "", expected.len() as u64, 1),
        )
        .unwrap();
    assert_eq!(complete, expected.as_bytes());
    assert_eq!(report.bytes_written(), expected.len() as u64);

    let mut output = Vec::new();
    let error = snapshot
        .write_text_to(
            &mut output,
            litchi_core::TextOutputOptions::new("", "", 8, 1),
        )
        .unwrap_err();
    assert!(matches!(
        &error,
        litchi_core::TextOutputError::Limit { limit, .. }
            if limit.kind() == litchi_core::TextOutputLimitKind::OutputBytes
    ));
    assert_eq!(error.progress().bytes_written(), 0);
    assert!(output.len() <= 8);

    assert!(
        open(&fixture(
            Dialect::Transitional,
            Mutation::UnknownNamedEntityFootnote,
        ))
        .story_text_snapshot(StorySelector::footnote(41))
        .is_err()
    );

    let instruction_snapshot = open(&fixture(
        Dialect::Transitional,
        Mutation::InstrTextReferenceFootnote,
    ))
    .story_text_snapshot(StorySelector::footnote(41))
    .unwrap();
    assert_eq!(instruction_snapshot.extract_text().unwrap(), "visible");
}

#[test]
fn secondary_edit_refuses_when_synthetic_wrapper_exceeds_output_limit() {
    let limits = source_backed::StoryTextLimits::new(4096, 10, 100_000, 64, 1, 4096, 4096)
        .unwrap()
        .with_secondary_entry_limits(8, 4096, 256)
        .unwrap();
    let snapshot = open(&fixture(Dialect::Transitional, Mutation::Valid))
        .story_text_snapshot_with_limits(StorySelector::footnote(41), limits)
        .unwrap();
    assert!(snapshot.edit().is_err());
}

#[test]
fn markup_heavier_than_output_bound_can_decode_to_visible_text() {
    let limits = source_backed::StoryTextLimits::new(64 * 1024, 10, 100_000, 64, 4, 4096, 4096)
        .unwrap()
        .with_secondary_entry_limits(8, 64 * 1024, 256)
        .unwrap();
    let snapshot = open(&fixture(
        Dialect::Transitional,
        Mutation::MarkupHeavyFootnote,
    ))
    .story_text_snapshot_with_limits(StorySelector::footnote(41), limits)
    .unwrap();
    assert_eq!(snapshot.paragraph_text(0).unwrap().as_deref(), Some("fit"));
}

#[test]
fn secondary_output_accumulation_and_borrowed_replacement_are_bounded() {
    let aggregate_limits =
        source_backed::StoryTextLimits::new(4096, 10, 100_000, 64, 20, 4096, 4096)
            .unwrap()
            .with_secondary_entry_limits(8, 4096, 256)
            .unwrap();
    let aggregate = open(&fixture(
        Dialect::Transitional,
        Mutation::TwoParagraphFootnote,
    ))
    .story_text_snapshot_with_limits(StorySelector::footnote(41), aggregate_limits)
    .unwrap();
    assert!(matches!(
        aggregate.extract_text(),
        Err(source_backed::StoryTextError::Limit { .. })
    ));

    let stream_limits =
        source_backed::StoryTextLimits::new(4096, 10, 100_000, 64, 4096, 4096, 4096)
            .unwrap()
            .with_secondary_entry_limits(8, 4096, 256)
            .unwrap();
    let stream_snapshot = open(&fixture(
        Dialect::Transitional,
        Mutation::TwoParagraphFootnote,
    ))
    .story_text_snapshot_with_limits(StorySelector::footnote(41), stream_limits)
    .unwrap();
    let mut streamed = Vec::new();
    let stream_error = stream_snapshot
        .write_text_to(
            &mut streamed,
            litchi_core::TextOutputOptions::new("", ":", 20, 10),
        )
        .unwrap_err();
    assert!(matches!(
        &stream_error,
        litchi_core::TextOutputError::Limit { limit, .. }
            if limit.kind() == litchi_core::TextOutputLimitKind::OutputBytes
    ));
    assert_eq!(stream_error.limit().unwrap().observed(), 30);
    assert_eq!(stream_error.progress().bytes_written(), 15);
    assert_eq!(stream_error.progress().objects_written(), 1);
    assert_eq!(streamed, b"footnote-target");

    let replacement_limits =
        source_backed::StoryTextLimits::new(4096, 10, 100_000, 64, 4096, 4096, 4)
            .unwrap()
            .with_secondary_entry_limits(8, 4096, 256)
            .unwrap();
    let replacement_snapshot = open(&fixture(Dialect::Transitional, Mutation::Valid))
        .story_text_snapshot_with_limits(StorySelector::footnote(41), replacement_limits)
        .unwrap();
    let mut edit = replacement_snapshot.edit().unwrap();
    let borrowed = "borrowed-replacement-too-large";
    assert!(
        edit.replace_paragraph_text(Position::new(0), borrowed)
            .is_err()
    );
    assert_eq!(
        edit.projected()
            .unwrap()
            .paragraph_text(0)
            .unwrap()
            .as_deref(),
        Some("footnote-target")
    );
}

#[test]
fn xml_expanding_replacement_is_refused_without_changing_projection() {
    let limits = source_backed::StoryTextLimits::new(4096, 10, 100_000, 64, 2048, 4096, 2048)
        .unwrap()
        .with_secondary_entry_limits(8, 4096, 256)
        .unwrap();
    let snapshot = open(&fixture(Dialect::Transitional, Mutation::Valid))
        .story_text_snapshot_with_limits(StorySelector::footnote(41), limits)
        .unwrap();
    let mut edit = snapshot.edit().unwrap();
    let before = edit.projected().unwrap().paragraph_text(0).unwrap();
    let replacement = "&<>\"'".repeat(256);
    assert!(
        edit.replace_paragraph_text(Position::new(0), replacement.as_str())
            .is_err()
    );
    assert_eq!(edit.projected().unwrap().paragraph_text(0).unwrap(), before);
}

struct PartialWriter {
    wrote_prefix: bool,
}

impl Write for PartialWriter {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if bytes.is_empty() {
            Ok(0)
        } else if self.wrote_prefix {
            Err(io::Error::new(
                io::ErrorKind::BrokenPipe,
                "intentional partial sink failure",
            ))
        } else {
            self.wrote_prefix = true;
            Ok(1)
        }
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[test]
fn secondary_publication_rejects_a_partial_sink_without_panicking() {
    let source = fixture(Dialect::Transitional, Mutation::Valid);
    let package = open(&source);
    let snapshot = package
        .story_text_snapshot(StorySelector::endnote(51))
        .unwrap();
    let mut edit = snapshot.edit().unwrap();
    edit.replace_paragraph_text(Position::new(0), "changed")
        .unwrap();
    let commit = edit.commit().unwrap();
    let result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        package.publish_story_text_commit_to_stream(
            PartialWriter {
                wrote_prefix: false,
            },
            &commit,
        )
    }));
    assert!(result.is_ok());
    assert!(result.unwrap().is_err());
}
