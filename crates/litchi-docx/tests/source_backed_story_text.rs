use std::io::{self, Write};
use std::num::{NonZeroU64, NonZeroUsize};
use std::sync::Arc;

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits as CoreLimits,
    OwnedSource, Position, TextOutputOptions,
};
use litchi_docx::source_backed::{self, story_text};
use litchi_docx::{Error, SourceLimits};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcPackage, PackURI, PackageWriter, Part, TargetMode};

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

fn fixture() -> Vec<u8> {
    let mut package = OpcPackage::new();
    let mut main = BlobPart::new(
        PackURI::new("/word/document.xml").unwrap(),
        ct::WML_DOCUMENT_MAIN.to_owned(),
        format!(
            r#"<w:document xmlns:w="{W}"><w:body><w:p><w:r><w:t>main</w:t></w:r></w:p></w:body></w:document>"#
        )
        .into_bytes(),
    );
    main.rels_mut()
        .try_add_relationship(
            rt::HEADER.to_owned(),
            "header1.xml".to_owned(),
            "rHeader".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
    main.rels_mut()
        .try_add_relationship(
            rt::FOOTER.to_owned(),
            "footer1.xml".to_owned(),
            "rFooter".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
    package.try_add_part(Box::new(main)).unwrap();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new("/word/header1.xml").unwrap(),
            ct::WML_HEADER.to_owned(),
            format!(r#"<w:hdr xmlns:w="{W}"><w:p><w:r><w:t>header</w:t></w:r></w:p></w:hdr>"#)
                .into_bytes(),
        )))
        .unwrap();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new("/word/footer1.xml").unwrap(),
            ct::WML_FOOTER.to_owned(),
            format!(r#"<w:ftr xmlns:w="{W}"><w:p><w:r><w:t>footer</w:t></w:r></w:p></w:ftr>"#)
                .into_bytes(),
        )))
        .unwrap();
    package.relate_to("word/document.xml", rt::OFFICE_DOCUMENT);
    PackageWriter::to_bytes(&package).unwrap()
}

fn open(bytes: &[u8]) -> source_backed::Package {
    source_backed::Package::from_read_at(Arc::new(OwnedSource::new(bytes.to_vec()))).unwrap()
}

fn main_only_fixture(document: &str) -> Vec<u8> {
    let mut package = OpcPackage::new();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new("/word/document.xml").unwrap(),
            ct::WML_DOCUMENT_MAIN.to_owned(),
            document.as_bytes().to_vec(),
        )))
        .unwrap();
    package.relate_to("word/document.xml", rt::OFFICE_DOCUMENT);
    PackageWriter::to_bytes(&package).unwrap()
}

fn try_main_only_fixture(document: &str) -> Result<Vec<u8>, String> {
    let mut package = OpcPackage::new();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new("/word/document.xml").map_err(|error| format!("{error:?}"))?,
            ct::WML_DOCUMENT_MAIN.to_owned(),
            document.as_bytes().to_vec(),
        )))
        .map_err(|error| format!("{error:?}"))?;
    package.relate_to("word/document.xml", rt::OFFICE_DOCUMENT);
    PackageWriter::to_bytes(&package).map_err(|error| format!("{error:?}"))
}

fn managed_open(bytes: Vec<u8>) -> (Budget, CancellationSource, source_backed::Package) {
    let memory = (bytes.len() as u64).saturating_mul(4).max(1);
    let budget = Budget::root(
        "source-backed-story-text-test",
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
    let package = source_backed::Package::from_read_at_with_options(
        Arc::new(OwnedSource::new(bytes)),
        litchi_docx::OpenOptions::default()
            .source_limits(SourceLimits::default())
            .execution_context(ExecutionContext::new(
                budget.clone(),
                cancellation,
                execution_limits,
            )),
    )
    .unwrap();
    (budget, cancellation_source, package)
}

struct HostileWriter;

impl Write for HostileWriter {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        Ok(bytes.len().saturating_add(1))
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[test]
fn selected_story_snapshot_reads_main_header_and_footer() {
    let package = open(&fixture());
    let main = package
        .story_text_snapshot(story_text::Selector::Main)
        .unwrap();
    assert_eq!(main.paragraph_count().unwrap(), 1);
    assert_eq!(main.paragraph_text(0).unwrap().as_deref(), Some("main"));

    let mut output: Vec<u8> = Vec::new();
    let report = main
        .write_text_to(&mut output, TextOutputOptions::default())
        .unwrap();
    assert_eq!(output, b"main");
    assert_eq!(report.objects_written(), 1);

    let header = package
        .story_text_snapshot(story_text::Selector::Header(0))
        .unwrap();
    assert_eq!(header.paragraph_text(0).unwrap().as_deref(), Some("header"));
    let footer = package
        .story_text_snapshot(story_text::Selector::Footer(0))
        .unwrap();
    assert_eq!(footer.paragraph_text(0).unwrap().as_deref(), Some("footer"));
    assert!(
        package
            .story_text_snapshot(story_text::Selector::Header(1))
            .is_err_and(|error| {
                matches!(
                    error,
                    story_text::Error::Document(Error::OutOfBounds {
                        object: "header",
                        index: 1,
                        ..
                    })
                )
            })
    );
}

#[test]
fn selected_story_edit_publishes_one_part_and_restores_exact_source() {
    let source = fixture();
    let package = open(&source);
    let snapshot = package
        .story_text_snapshot(story_text::Selector::Header(0))
        .unwrap();
    let mut edit = snapshot.raw_tree(snapshot.limits_value()).unwrap();
    edit.replace_paragraph_text(Position::new(0), "changed")
        .unwrap();
    let patch = edit.patch().unwrap();
    assert_eq!(
        patch.target().paragraph_text(0).unwrap().as_deref(),
        Some("changed")
    );

    let mut published = Vec::new();
    let publication = package
        .apply_story_text(&patch.clone().into())
        .and_then(|committed| {
            committed.package().write_to(&mut published)?;
            Ok(committed)
        })
        .unwrap();
    let reopened = open(&published);
    assert_eq!(
        reopened
            .story_text_snapshot(story_text::Selector::Header(0))
            .unwrap()
            .paragraph_text(0)
            .unwrap()
            .as_deref(),
        Some("changed")
    );
    assert_eq!(
        reopened
            .story_text_snapshot(story_text::Selector::Main)
            .unwrap()
            .paragraph_text(0)
            .unwrap()
            .as_deref(),
        Some("main")
    );
    assert_eq!(
        reopened
            .story_text_snapshot(story_text::Selector::Footer(0))
            .unwrap()
            .paragraph_text(0)
            .unwrap()
            .as_deref(),
        Some("footer")
    );

    let inverse_result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        open(&published)
            .apply_story_text(&publication.patch().inverse())
            .and_then(|committed| {
                committed.package().write_to(HostileWriter)?;
                Ok(committed)
            })
    }));
    assert!(inverse_result.is_ok());
    assert!(inverse_result.unwrap().is_err());

    let mut restored = Vec::new();
    open(&published)
        .apply_story_text(&publication.patch().inverse())
        .and_then(|committed| Ok(committed.package().write_to(&mut restored)?))
        .unwrap();
    assert_eq!(restored, source);
}

#[test]
fn selected_story_replays_on_independent_matching_copies_and_chains_edits() {
    let source = fixture();
    let package = open(&source);
    let snapshot = package
        .story_text_snapshot(story_text::Selector::Header(0))
        .unwrap();
    let mut edit = snapshot.raw_tree(snapshot.limits_value()).unwrap();
    edit.replace_paragraph_text(Position::new(0), "first change")
        .unwrap();
    let patch = edit.patch().unwrap();

    // The patch records the selected story's content, so a separately opened
    // package with the same bytes is an eligible recipient.
    let first = open(&source)
        .apply_story_text(&patch.clone().into())
        .unwrap();
    let mut first_bytes = Vec::new();
    first.package().write_to(&mut first_bytes).unwrap();
    assert_eq!(
        open(&first_bytes)
            .story_text_snapshot(story_text::Selector::Header(0))
            .unwrap()
            .paragraph_text(0)
            .unwrap()
            .as_deref(),
        Some("first change")
    );

    // A second raw edit can be authored from that independently published
    // copy, and its inverse restores the exact first publication bytes.
    let second_source = open(&first_bytes);
    let second_snapshot = second_source
        .story_text_snapshot(story_text::Selector::Header(0))
        .unwrap();
    let mut second_edit = second_snapshot
        .raw_tree(second_snapshot.limits_value())
        .unwrap();
    second_edit
        .replace_paragraph_text(Position::new(0), "second change")
        .unwrap();
    let second = second_source
        .apply_story_text(&second_edit.patch().unwrap().into())
        .unwrap();
    let mut second_bytes = Vec::new();
    second.package().write_to(&mut second_bytes).unwrap();
    assert_eq!(
        open(&second_bytes)
            .story_text_snapshot(story_text::Selector::Header(0))
            .unwrap()
            .paragraph_text(0)
            .unwrap()
            .as_deref(),
        Some("second change")
    );

    let mut restored = Vec::new();
    open(&second_bytes)
        .apply_story_text(&second.patch().inverse())
        .unwrap()
        .package()
        .write_to(&mut restored)
        .unwrap();
    assert_eq!(restored, first_bytes);
}

#[test]
fn selected_story_noop_copies_the_complete_source_byte_for_byte() {
    let source = fixture();
    let package = open(&source);
    let snapshot = package
        .story_text_snapshot(story_text::Selector::Footer(0))
        .unwrap();
    let mut edit = snapshot.raw_tree(snapshot.limits_value()).unwrap();
    edit.replace_paragraph_text(Position::new(0), "footer")
        .unwrap();
    let patch = edit.patch().unwrap();
    let mut output = Vec::new();
    package
        .apply_story_text(&patch.clone().into())
        .and_then(|committed| Ok(committed.package().write_to(&mut output)?))
        .unwrap();
    assert_eq!(output, source);
}

#[test]
fn raw_story_tree_caches_projection_and_invalidates_after_mutation() {
    let package = open(&fixture());
    let snapshot = package
        .story_text_snapshot(story_text::Selector::Main)
        .unwrap();
    let mut tree = snapshot.raw_tree(snapshot.limits_value()).unwrap();

    let first = tree.projected().unwrap();
    let second = tree.projected().unwrap();
    assert_eq!(first.part_fingerprint(), second.part_fingerprint());
    assert_eq!(
        first.paragraph_text(0).unwrap(),
        second.paragraph_text(0).unwrap()
    );

    tree.replace_paragraph_text(Position::new(0), "cached mutation")
        .unwrap();
    let changed = tree.projected().unwrap();
    assert_ne!(first.part_fingerprint(), changed.part_fingerprint());
    assert_eq!(
        changed.paragraph_text(0).unwrap().as_deref(),
        Some("cached mutation")
    );
    let patch = tree.patch().unwrap();
    assert_eq!(
        patch.target().part_fingerprint(),
        changed.part_fingerprint()
    );
}

#[test]
fn raw_story_tree_intersects_requested_output_policy() {
    let package = open(&fixture());
    let snapshot = package
        .story_text_snapshot(story_text::Selector::Header(0))
        .unwrap();
    let requested = story_text::Limits::new(4096, 10, 10_000, 64, 1, 4096, 4096).unwrap();
    assert!(matches!(
        snapshot.raw_tree(requested),
        Err(story_text::Error::Limit { .. })
    ));
}

#[test]
fn managed_story_reads_are_bounded_and_edits_are_typed_refusals() {
    let (budget, _cancellation, package) = managed_open(fixture());
    let snapshot = package
        .story_text_snapshot(story_text::Selector::Main)
        .unwrap();
    assert_eq!(snapshot.paragraph_text(0).unwrap().as_deref(), Some("main"));
    assert!(matches!(
            snapshot.raw_tree(snapshot.limits_value()),
            Err(story_text::Error::Document(Error::Refused(refusal))) if refusal.operation() == "source-backed story text edit"
    ));
    drop(snapshot);
    drop(package);
    assert_eq!(budget.used(litchi_core::Resource::Memory), 0);
}

#[test]
fn story_scan_enforces_paragraph_depth_and_output_limits() {
    let document = format!(
        r#"<w:document xmlns:w="{W}"><w:body><w:p><w:r><w:t>one</w:t></w:r></w:p><w:p><w:r><w:t>two</w:t></w:r></w:p></w:body></w:document>"#
    );
    let source = main_only_fixture(&document);
    let paragraph_limited =
        story_text::Limits::new(4096, 1, 10_000, 64, 4096, 4_096, 4096).unwrap();
    assert!(matches!(
        open(&source)
            .with_semantic_limits(
                litchi_docx::SemanticLimits::default().story_text(paragraph_limited)
            )
            .story_text_snapshot(story_text::Selector::Main),
        Err(story_text::Error::Limit {
            kind: story_text::LimitKind::Paragraphs,
            ..
        })
    ));

    let depth_limited = story_text::Limits::new(4096, 10, 10_000, 2, 4096, 4_096, 4096).unwrap();
    assert!(matches!(
        open(&source)
            .with_semantic_limits(litchi_docx::SemanticLimits::default().story_text(depth_limited))
            .story_text_snapshot(story_text::Selector::Main),
        Err(story_text::Error::Limit {
            kind: story_text::LimitKind::XmlDepth,
            ..
        })
    ));

    let empty_depth_document =
        format!(r#"<w:document xmlns:w="{W}"><w:body><w:p/></w:body></w:document>"#);
    let empty_depth = story_text::Limits::new(4096, 10, 10_000, 2, 4096, 4_096, 4096).unwrap();
    assert!(matches!(
        open(&main_only_fixture(&empty_depth_document))
            .with_semantic_limits(litchi_docx::SemanticLimits::default().story_text(empty_depth))
            .story_text_snapshot(story_text::Selector::Main),
        Err(story_text::Error::Limit {
            kind: story_text::LimitKind::XmlDepth,
            ..
        })
    ));

    let output_limited = story_text::Limits::new(4096, 10, 10_000, 64, 1, 4_096, 4096).unwrap();
    let snapshot = open(&main_only_fixture(&format!(
        r#"<w:document xmlns:w="{W}"><w:body><w:p><w:r><w:t>long</w:t></w:r></w:p></w:body></w:document>"#
    )))
    .with_semantic_limits(litchi_docx::SemanticLimits::default().story_text(output_limited)).story_text_snapshot(story_text::Selector::Main)
    .unwrap();
    assert!(matches!(
        snapshot.extract_text(),
        Err(story_text::Error::Limit {
            kind: story_text::LimitKind::ParagraphText,
            ..
        })
    ));

    let aggregate_limits = story_text::Limits::new(4096, 10, 10_000, 64, 6, 4_096, 4096).unwrap();
    let aggregate_snapshot = open(&main_only_fixture(&format!(
        r#"<w:document xmlns:w="{W}"><w:body><w:p><w:r><w:t>one</w:t></w:r></w:p><w:p><w:r><w:t>two</w:t></w:r></w:p></w:body></w:document>"#
    )))
    .with_semantic_limits(litchi_docx::SemanticLimits::default().story_text(aggregate_limits)).story_text_snapshot(story_text::Selector::Main)
    .unwrap();
    let mut aggregate_output = Vec::new();
    let aggregate_error = aggregate_snapshot
        .write_text_to(
            &mut aggregate_output,
            TextOutputOptions::new("|", "::", u64::MAX, 10),
        )
        .unwrap_err();
    assert!(matches!(
        aggregate_error,
        litchi_core::TextOutputError::Limit { limit, .. }
            if limit.kind() == litchi_core::TextOutputLimitKind::OutputBytes
    ));
    assert!(aggregate_output.len() <= 6);
    assert!(aggregate_output.starts_with(b"one"));
}

#[test]
fn hostile_publication_writer_is_rejected_without_panicking() {
    let source = fixture();
    let package = open(&source);
    let snapshot = package
        .story_text_snapshot(story_text::Selector::Header(0))
        .unwrap();
    let mut edit = snapshot.raw_tree(snapshot.limits_value()).unwrap();
    edit.replace_paragraph_text(Position::new(0), "changed")
        .unwrap();
    let patch = edit.patch().unwrap();
    let mut writer = HostileWriter;
    assert!(
        package
            .apply_story_text(&patch.clone().into())
            .and_then(|committed| {
                committed.package().write_to(&mut writer)?;
                Ok(committed)
            })
            .is_err()
    );
}

#[test]
fn trailing_source_bytes_are_preserved_by_an_exact_noop() {
    let mut source = fixture();
    source.extend_from_slice(b"opaque trailing source bytes");
    let package = open(&source);
    let snapshot = package
        .story_text_snapshot(story_text::Selector::Footer(0))
        .unwrap();
    let mut edit = snapshot.raw_tree(snapshot.limits_value()).unwrap();
    edit.replace_paragraph_text(Position::new(0), "footer")
        .unwrap();
    let patch = edit.patch().unwrap();
    let mut output = Vec::new();
    package
        .apply_story_text(&patch.clone().into())
        .and_then(|committed| Ok(committed.package().write_to(&mut output)?))
        .unwrap();
    assert_eq!(output, source);
}

#[test]
fn namespace_aliases_are_resolved_and_wrong_bindings_are_refused() {
    let alias = main_only_fixture(&format!(
        r#"<x:document xmlns:x="{W}"><x:body><x:p><x:r><x:t>alias</x:t></x:r></x:p></x:body></x:document>"#
    ));
    assert_eq!(
        open(&alias)
            .story_text_snapshot(story_text::Selector::Main)
            .unwrap()
            .paragraph_text(0)
            .unwrap()
            .as_deref(),
        Some("alias")
    );

    let escaped_word_namespace = main_only_fixture(
        r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/mai&#x6e;"><w:body><w:p><w:r><w:t>escaped</w:t></w:r></w:p></w:body></w:document>"#,
    );
    let escaped_snapshot = open(&escaped_word_namespace)
        .story_text_snapshot(story_text::Selector::Main)
        .unwrap();
    assert_eq!(
        escaped_snapshot.paragraph_text(0).unwrap().as_deref(),
        Some("escaped")
    );
    assert!(
        escaped_snapshot
            .raw_tree(escaped_snapshot.limits_value())
            .is_err()
    );

    let wrong = main_only_fixture(
        r#"<x:document xmlns:x="urn:not-word"><x:body><x:p><x:r><x:t>wrong</x:t></x:r></x:p></x:body></x:document>"#,
    );
    assert!(matches!(
        open(&wrong).story_text_snapshot(story_text::Selector::Main),
        Err(story_text::Error::Document(Error::Refused(_)))
    ));
}

#[test]
fn nested_mce_and_processing_instructions_are_refused() {
    let mce = main_only_fixture(&format!(
        r#"<w:document xmlns:w="{W}" xmlns:m="http://schemas.openxmlformats.org/markup-compatibility/2006"><w:body><m:AlternateContent/></w:body></w:document>"#
    ));
    assert!(
        open(&mce)
            .story_text_snapshot(story_text::Selector::Main)
            .is_err()
    );

    let escaped_mce = main_only_fixture(&format!(
        r#"<w:document xmlns:w="{W}" xmlns:m="http://schemas.openxmlformats.org/markup-compatibility/200&#x36;"><w:body><m:AlternateContent/></w:body></w:document>"#
    ));
    assert!(
        open(&escaped_mce)
            .story_text_snapshot(story_text::Selector::Main)
            .is_err()
    );

    let processing_instruction = main_only_fixture(&format!(
        r#"<?not-supported?><w:document xmlns:w="{W}"><w:body/></w:document>"#
    ));
    assert!(
        open(&processing_instruction)
            .story_text_snapshot(story_text::Selector::Main)
            .is_err()
    );
}

#[test]
fn strict_xml_rejects_duplicate_attributes_and_mismatched_end_names() {
    let duplicate_namespace = try_main_only_fixture(&format!(
        r#"<w:document xmlns:w="{W}" xmlns:w="{W}"><w:body/></w:document>"#
    ));
    if let Ok(source) = duplicate_namespace {
        assert!(matches!(
            open(&source).story_text_snapshot(story_text::Selector::Main),
            Err(story_text::Error::Document(_))
        ));
    }

    let duplicate_attribute = try_main_only_fixture(&format!(
        r#"<w:document xmlns:w="{W}" w:dup="one" w:dup="two"><w:body/></w:document>"#
    ));
    if let Ok(source) = duplicate_attribute {
        assert!(matches!(
            open(&source).story_text_snapshot(story_text::Selector::Main),
            Err(story_text::Error::Document(_))
        ));
    }

    let mismatched_end = try_main_only_fixture(&format!(
        r#"<w:document xmlns:w="{W}"><w:body></w:document>"#
    ));
    if let Ok(source) = mismatched_end {
        assert!(matches!(
            open(&source).story_text_snapshot(story_text::Selector::Main),
            Err(story_text::Error::Document(_))
        ));
    }
}

#[test]
fn patches_reject_foreign_and_stale_snapshots_and_cancel_before_output() {
    let source = fixture();
    let package = open(&source);
    let snapshot = package
        .story_text_snapshot(story_text::Selector::Main)
        .unwrap();
    let mut edit = snapshot.raw_tree(snapshot.limits_value()).unwrap();
    edit.replace_paragraph_text(Position::new(0), "changed")
        .unwrap();
    let patch = edit.patch().unwrap();
    assert!(matches!(
        patch.apply(patch.target()),
        Err(story_text::Error::PatchConflict)
    ));

    let foreign = open(&source)
        .story_text_snapshot(story_text::Selector::Main)
        .unwrap();
    assert!(matches!(
        patch.apply(&foreign),
        Err(story_text::Error::ForeignSource)
    ));

    let (_budget, cancellation, package) = managed_open(source);
    cancellation.cancel();
    let output: Vec<u8> = Vec::new();
    assert!(
        package
            .story_text_snapshot(story_text::Selector::Main)
            .is_err()
    );
    assert!(output.is_empty());
}
