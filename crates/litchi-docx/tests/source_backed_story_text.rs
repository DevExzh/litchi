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

fn open_with_source_limits(bytes: &[u8], limits: SourceLimits) -> source_backed::Package {
    source_backed::Package::from_read_at_with_options(
        Arc::new(OwnedSource::new(bytes.to_vec())),
        litchi_docx::OpenOptions::default().source_limits(limits),
    )
    .unwrap()
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
fn patches_accept_matching_copy_and_reject_stale_snapshots_and_cancel_before_output() {
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
    let foreign_applied = patch.apply(&foreign).unwrap();
    assert_eq!(
        foreign_applied.paragraph_text(0).unwrap().as_deref(),
        Some("changed")
    );

    let tight_limits = story_text::Limits::new(16 * 1024, 128, 10_000, 64, 1, 16, 4096).unwrap();
    assert!(matches!(
        snapshot.raw_tree(tight_limits),
        Err(story_text::Error::Limit {
            kind: story_text::LimitKind::TransactionXml,
            ..
        })
    ));

    let changed_package = open(&source)
        .apply_story_text(&patch.clone().into())
        .unwrap();
    let mut changed_bytes = Vec::new();
    changed_package
        .package()
        .write_to(&mut changed_bytes)
        .unwrap();
    let changed = open(&changed_bytes)
        .story_text_snapshot(story_text::Selector::Main)
        .unwrap();
    assert!(matches!(
        patch.apply(&changed),
        Err(story_text::Error::PatchConflict)
    ));

    let (_budget, cancellation, package) = managed_open(source.clone());
    let cancelled_snapshot = package
        .story_text_snapshot(story_text::Selector::Main)
        .unwrap();
    cancellation.cancel();
    assert!(matches!(
        patch.apply(&cancelled_snapshot),
        Err(story_text::Error::Document(Error::Opc(
            litchi_opc::OpcError::Cancelled
        )))
    ));
    let noop = snapshot
        .raw_tree(snapshot.limits_value())
        .unwrap()
        .patch()
        .unwrap();
    assert!(noop.is_noop());
    assert!(matches!(
        noop.apply(&cancelled_snapshot),
        Err(story_text::Error::Document(Error::Opc(
            litchi_opc::OpcError::Cancelled
        )))
    ));
    let output: Vec<u8> = Vec::new();
    assert!(
        package
            .story_text_snapshot(story_text::Selector::Main)
            .is_err()
    );
    assert!(output.is_empty());
}

#[test]
fn raw_story_replay_reopens_matches_and_chains_under_recipient_limits() {
    let source = fixture();
    let author = open(&source);
    let snapshot = author
        .story_text_snapshot(story_text::Selector::Header(0))
        .unwrap();
    let mut tree = snapshot.raw_tree(snapshot.limits_value()).unwrap();
    tree.replace_paragraph_text(Position::new(0), "first raw change")
        .unwrap();
    let patch = tree.patch().unwrap();

    let recipient_limits =
        story_text::Limits::new(16 * 1024, 128, 10_000, 64, 16 * 1024, 16, 4096).unwrap();
    let recipient = open(&source)
        .with_semantic_limits(litchi_docx::SemanticLimits::default().story_text(recipient_limits))
        .story_text_snapshot(story_text::Selector::Header(0))
        .unwrap();
    assert_ne!(snapshot.limits_value(), recipient.limits_value());

    let first = patch.apply(&recipient).unwrap();
    assert_eq!(first.limits_value(), recipient.limits_value());
    assert_eq!(
        first.paragraph_text(0).unwrap().as_deref(),
        Some("first raw change")
    );

    let restrictive_limits =
        story_text::Limits::new(16 * 1024, 128, 10_000, 64, 1, 16, 4096).unwrap();
    let restrictive = open(&source)
        .with_semantic_limits(litchi_docx::SemanticLimits::default().story_text(restrictive_limits))
        .story_text_snapshot(story_text::Selector::Header(0))
        .unwrap();
    assert!(matches!(
        patch.apply(&restrictive),
        Err(story_text::Error::Limit {
            kind: story_text::LimitKind::OutputBytes,
            ..
        })
    ));

    let restored = patch.inverse().apply(&first).unwrap();
    assert_eq!(restored.part_fingerprint(), recipient.part_fingerprint());
    assert_eq!(
        restored.paragraph_text(0).unwrap().as_deref(),
        Some("header")
    );

    let mut second_tree = first.raw_tree(first.limits_value()).unwrap();
    second_tree
        .replace_paragraph_text(Position::new(0), "second raw change")
        .unwrap();
    let second_patch = second_tree.patch().unwrap();
    let second = second_patch.apply(&first).unwrap();
    assert_eq!(
        second.paragraph_text(0).unwrap().as_deref(),
        Some("second raw change")
    );
    let second_restored = second_patch.inverse().apply(&second).unwrap();
    assert_eq!(
        second_restored.paragraph_text(0).unwrap().as_deref(),
        Some("first raw change")
    );
}

#[test]
fn story_replay_uses_recipient_limits_for_broad_raw_patches() {
    let source = fixture();
    let producer_limits =
        story_text::Limits::new(128 * 1024, 128, 10_000, 64, 128 * 1024, 16, 16_384).unwrap();
    let producer = open(&source)
        .with_semantic_limits(litchi_docx::SemanticLimits::default().story_text(producer_limits));
    let snapshot = producer
        .story_text_snapshot(story_text::Selector::Main)
        .unwrap();
    let mut tree = snapshot.raw_tree(snapshot.limits_value()).unwrap();
    tree.replace_paragraph_text(Position::new(0), "broad producer patch")
        .unwrap();
    let patch = tree.patch().unwrap();

    let accepted_limits =
        story_text::Limits::new(16 * 1024, 128, 10_000, 64, 1024, 16, 4096).unwrap();
    let accepted = open(&source)
        .with_semantic_limits(litchi_docx::SemanticLimits::default().story_text(accepted_limits))
        .story_text_snapshot(story_text::Selector::Main)
        .unwrap();
    let accepted_target = patch.apply(&accepted).unwrap();
    assert_eq!(accepted_target.limits_value(), accepted_limits);

    let narrow_limits = story_text::Limits::new(16 * 1024, 128, 10_000, 64, 1, 16, 4096).unwrap();
    let narrow = open(&source)
        .with_semantic_limits(litchi_docx::SemanticLimits::default().story_text(narrow_limits));
    let narrow_snapshot = narrow
        .story_text_snapshot(story_text::Selector::Main)
        .unwrap();
    assert!(matches!(
        patch.apply(&narrow_snapshot),
        Err(story_text::Error::Limit {
            kind: story_text::LimitKind::OutputBytes,
            ..
        })
    ));
    let package_patch = patch.clone().into();
    assert!(matches!(
        narrow.apply_story_text(&package_patch),
        Err(story_text::Error::Limit {
            kind: story_text::LimitKind::OutputBytes,
            ..
        })
    ));

    // A package-family commit carries the full-artifact restore fast path.
    // A recipient with a narrower semantic policy must use the family
    // publisher instead, so the recipient's output ceiling is enforced rather
    // than inheriting the producer's admitted policy.
    let mut owner_edit = producer
        .edit_story_text(story_text::Selector::Main)
        .unwrap();
    let broad_owner_text = (0..2_200)
        .map(|index| format!("{index:04x}"))
        .collect::<String>();
    owner_edit
        .replace_paragraph_text(Position::new(0), &broad_owner_text)
        .unwrap();
    let owner_commit = owner_edit.commit().unwrap();

    // An owner commit is deferred: replay it onto an independently opened
    // equal-content package and verify that the recipient writes exactly the
    // producer's bytes, including the inverse. This exercises the bounded
    // artifact-stream restore path rather than the same-owner shortcut.
    let mut producer_bytes = Vec::new();
    owner_commit
        .package()
        .write_to(&mut producer_bytes)
        .unwrap();
    let independent = open(&source)
        .apply_story_text(owner_commit.patch())
        .unwrap();
    let mut independent_bytes = Vec::new();
    independent
        .package()
        .write_to(&mut independent_bytes)
        .unwrap();
    assert_eq!(independent_bytes, producer_bytes);
    let mut independent_inverse = Vec::new();
    independent
        .package()
        .apply_story_text(&owner_commit.patch().inverse())
        .unwrap()
        .package()
        .write_to(&mut independent_inverse)
        .unwrap();
    assert_eq!(independent_inverse, source);

    // The source ceiling admits the recipient's original archive exactly.
    // Changed overlays are checked against their final member and aggregate
    // dimensions; a compact ZIP stream can remain within this input ceiling.
    assert!(producer_bytes.len() > source.len());
    let exact_source_limits = SourceLimits::builder()
        .max_input_bytes(source.len() as u64)
        .unwrap()
        .build()
        .unwrap();
    let narrow_source = open_with_source_limits(&source, exact_source_limits);
    let narrow_result = narrow_source
        .apply_story_text(owner_commit.patch())
        .and_then(|committed| {
            committed.package().write_to(Vec::new())?;
            Ok(committed)
        });
    let narrow_diagnostic = format!("{narrow_result:?}");
    assert!(narrow_result.is_ok(), "result={narrow_diagnostic}");

    // A recipient may admit the original archive while refusing the larger
    // changed main part. The artifact-stream census must enforce the full
    // recipient profile before accepting any output, rather than checking
    // only the archive's total input length.
    let part_tight_limits = SourceLimits::builder()
        .max_input_bytes(source.len() as u64)
        .unwrap()
        .max_part_bytes(1024)
        .unwrap()
        .build()
        .unwrap();
    let part_tight = open_with_source_limits(&source, part_tight_limits);
    let part_result = part_tight
        .apply_story_text(owner_commit.patch())
        .and_then(|committed| {
            committed.package().write_to(Vec::new())?;
            Ok(committed)
        });
    assert!(matches!(
        part_result,
        Err(story_text::Error::Document(Error::Opc(
            litchi_opc::OpcError::Limit {
                kind: litchi_opc::SourceLimitKind::PartBytes,
                ..
            }
        )))
    ));

    // The exact member count remains a valid recipient profile for an
    // unchanged cross-owner replay. One below that admitted count is refused
    // at source admission, before a replay can publish anything.
    let admitted_part_count = open(&source)
        .source_backed_package()
        .unwrap()
        .iter_parts()
        .count();
    let exact_parts = SourceLimits::builder()
        .max_input_bytes(source.len() as u64)
        .unwrap()
        .max_parts(admitted_part_count)
        .unwrap()
        .build()
        .unwrap();
    let exact_part_recipient = open_with_source_limits(&source, exact_parts);
    let exact_part_replay = exact_part_recipient
        .apply_story_text(owner_commit.patch())
        .unwrap();
    let mut exact_part_output = Vec::new();
    exact_part_replay
        .package()
        .write_to(&mut exact_part_output)
        .unwrap();
    assert_eq!(exact_part_output, producer_bytes);

    let below_parts = SourceLimits::builder()
        .max_input_bytes(source.len() as u64)
        .unwrap()
        .max_parts(admitted_part_count.saturating_sub(1))
        .unwrap()
        .build()
        .unwrap();
    assert!(
        source_backed::Package::from_read_at_with_options(
            Arc::new(OwnedSource::new(source.clone())),
            litchi_docx::OpenOptions::default().source_limits(below_parts),
        )
        .is_err()
    );

    assert!(matches!(
        narrow.apply_story_text(owner_commit.patch()),
        Err(story_text::Error::Limit {
            kind: story_text::LimitKind::OutputBytes,
            ..
        })
    ));
}

#[test]
fn ignorable_metadata_preserves_source_and_hidden_word_content_still_refuses() {
    const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    for namespace in [
        MC,
        "http://schemas.openxmlformats.org/markup-compatibility/200&#x36;",
    ] {
        let xml = format!(
            r#"<w:document xmlns:w="{W}" xmlns:compat="{namespace}" xmlns:x="urn:opaque" compat:Ignorable="x"><w:body><w:p x:identity="keep"><x:opaque x:value="retained"/><w:r><w:t>before</w:t></w:r></w:p></w:body></w:document>"#
        );
        let original = main_only_fixture(&xml);
        let package = open(&original);
        let source = package
            .story_text_snapshot(story_text::Selector::Main)
            .unwrap();
        assert_eq!(source.paragraph_text(0).unwrap().as_deref(), Some("before"));
        let mut edit = package.edit_story_text(story_text::Selector::Main).unwrap();
        edit.replace_paragraph_text(Position::new(0), "after")
            .unwrap();
        let commit = edit.commit().unwrap();
        let mut output = Vec::new();
        commit.package().write_to(&mut output).unwrap();
        let opc = OpcPackage::from_bytes(&output, SourceLimits::default()).unwrap();
        assert_eq!(
            opc.get_part(&PackURI::new("/word/document.xml").unwrap())
                .unwrap()
                .blob(),
            xml.replace("before", "after").as_bytes()
        );
    }
    for (directive, body) in [
        (
            "compat:Ignorable=\"unbound\"",
            "<w:p><w:r><w:t>visible</w:t></w:r></w:p>",
        ),
        (
            "compat:Ignorable=\"x\"",
            "<x:hidden><w:p><w:r><w:t>hidden</w:t></w:r></w:p></x:hidden>",
        ),
        (
            "compat:Ignorable=\"x\" compat:ProcessContent=\"x:hidden\"",
            "<x:hidden><w:p/></x:hidden>",
        ),
        (
            "compat:Ignorable=\"x\"",
            "<w:p compat:Ignorable=\"y\"><w:r><w:t>visible</w:t></w:r></w:p>",
        ),
    ] {
        let xml = format!(
            r#"<w:document xmlns:w="{W}" xmlns:compat="{MC}" xmlns:x="urn:opaque" xmlns:y="urn:nested" {directive}><w:body>{body}</w:body></w:document>"#
        );
        let original = main_only_fixture(&xml);
        let package = open(&original);
        assert!(
            matches!(package.story_text_snapshot(story_text::Selector::Main),
            Err(story_text::Error::Document(Error::Refused(r))) if r.kind()==litchi_core::edit::RefusalKind::UnsupportedContent)
        );
        let mut output = Vec::new();
        package.write_to(&mut output).unwrap();
        assert_eq!(output, original);
    }
}
