//! The source-backed package's memos of the visible main document and its
//! body block index.

use super::*;
use litchi_core::Position;

use crate::parts::memo::MEMO_ENTRY_OVERHEAD_BYTES;
use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits, OwnedSource,
};
use litchi_opc::{BlobPart, OpcPackage, PackURI, PackageWriter, SourceLimits};
use std::num::{NonZeroU64, NonZeroUsize};

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

fn docx(document: String) -> Vec<u8> {
    let mut package = OpcPackage::new();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new("/word/document.xml").unwrap(),
            ct::WML_DOCUMENT_MAIN.to_owned(),
            document.into_bytes(),
        )))
        .unwrap();
    package.relate_to("word/document.xml", rt::OFFICE_DOCUMENT);
    PackageWriter::to_bytes(&package).unwrap()
}

fn word_document(namespaces: &str, body: &str) -> String {
    format!(r#"<w:document xmlns:w="{W}"{namespaces}><w:body>{body}</w:body></w:document>"#)
}

fn owned(xml: &DocumentPayload) -> &Arc<Vec<u8>> {
    let DocumentPayload::Owned(xml) = xml else {
        panic!("unmanaged views own their visible XML")
    };
    xml
}

#[test]
fn unmanaged_views_share_one_projection_without_rereading_the_part() {
    let bytes = docx(word_document(
        &format!(r#" xmlns:mc="{MC}""#),
        "<w:p><w:r><w:t>alpha</w:t></w:r></w:p><w:p><w:r><w:t>beta</w:t></w:r></w:p>",
    ));
    let package = Package::from_read_at(Arc::new(OwnedSource::new(bytes))).unwrap();
    let first = package.document().unwrap();
    let after_first = package.physical_package().unwrap().cache_diagnostics();
    let second = package.document().unwrap();
    let after_second = package.physical_package().unwrap().cache_diagnostics();
    assert_eq!(
        (after_second.hits, after_second.successful_loads),
        (after_first.hits, after_first.successful_loads),
        "a memo hit reads no payload"
    );
    assert!(Arc::ptr_eq(owned(&first.xml), owned(&second.xml)));
    let visible = package.paragraph_index_cache.visible_diagnostics();
    assert_eq!((visible.builds, visible.hits), (1, 1));

    let first_body = first.body_blocks().unwrap();
    assert!(first_body.shares_index_with(&second.body_blocks().unwrap()));
    assert_eq!(package.paragraph_index_cache.body_diagnostics().builds, 1);
    assert_eq!(
        second
            .body_paragraph(Position::new(1))
            .unwrap()
            .unwrap()
            .text()
            .unwrap(),
        "beta"
    );
    assert_eq!(first.extract_text().unwrap(), "alphabeta");

    drop((first, second, first_body));
    package.trim_document_memos();
    let visible = package.paragraph_index_cache.visible_diagnostics();
    assert_eq!(visible.clean_evictions, 1);
    assert_eq!(visible.clean_entries, 0);
    let third = package.document().unwrap();
    assert_eq!(third.extract_text().unwrap(), "alphabeta");
    assert_eq!(
        package.paragraph_index_cache.visible_diagnostics().builds,
        2
    );
}

#[test]
fn a_part_without_markup_compatibility_shares_the_source_payload() {
    let bytes = docx(word_document("", "<w:p><w:r><w:t>plain</w:t></w:r></w:p>"));
    let package = Package::from_read_at(Arc::new(OwnedSource::new(bytes))).unwrap();
    let document = package.document().unwrap();
    let raw = package
        .physical_package()
        .unwrap()
        .main_document_part()
        .unwrap()
        .data()
        .unwrap()
        .into_arc()
        .unwrap();
    assert!(Arc::ptr_eq(owned(&document.xml), &raw));
    let visible = package.paragraph_index_cache.visible_diagnostics();
    assert_eq!(visible.retained_weight, MEMO_ENTRY_OVERHEAD_BYTES);
    // The verdict is reused; the payload comes from the source cache.
    let again = package.document().unwrap();
    assert!(Arc::ptr_eq(owned(&again.xml), &raw));
    assert_eq!(package.paragraph_index_cache.visible_diagnostics().hits, 1);
}

#[test]
fn body_queries_refuse_managed_views_and_managed_views_build_no_projection() {
    let bytes = docx(word_document(
        "",
        "<w:p><w:r><w:t>managed</w:t></w:r></w:p>",
    ));
    let memory = 1 << 22;
    let budget = Budget::root(
        "docx-document-memo-test",
        Limits::new(memory, 1 << 26, 1 << 26, 1_000_000, 1024, 1 << 30),
    );
    let (_cancellation_source, cancellation) = CancellationSource::pair();
    let limits = ExecutionLimits::new(
        NonZeroUsize::MIN,
        NonZeroUsize::MIN,
        NonZeroU64::new(memory).unwrap(),
        0,
    )
    .unwrap();
    let package = Package::from_read_at_with_options(
        Arc::new(OwnedSource::new(bytes)),
        crate::OpenOptions::default()
            .source_limits(SourceLimits::default())
            .execution_context(ExecutionContext::new(budget, cancellation, limits)),
    )
    .unwrap();
    let document = package.document().unwrap();
    assert_eq!(document.paragraph_count().unwrap(), 1);
    assert!(matches!(
            document.body_blocks(),
            Err(Error::Refused(refusal)) if refusal.operation() == "document body blocks"
    ));
    assert!(document.body_paragraph(Position::new(0)).is_err());
    let visible = package.paragraph_index_cache.visible_diagnostics();
    assert_eq!(visible.misses, 0);
}

#[test]
fn body_indexes_follow_effective_revisions_and_keep_old_values() {
    let bytes = docx(word_document("", "<w:p><w:r><w:t>before</w:t></w:r></w:p>"));
    let original = Package::from_read_at(Arc::new(OwnedSource::new(bytes))).unwrap();
    let old = original.document().unwrap().body_blocks().unwrap();
    let mut edit = original.edit().unwrap();
    edit.replace_paragraph_text(Position::new(0), "after")
        .unwrap();
    let current = edit.commit().unwrap().into_package();
    let first = current.document().unwrap().body_blocks().unwrap();
    let second = current.document().unwrap().body_blocks().unwrap();
    assert!(first.shares_index_with(&second));
    assert!(!old.shares_index_with(&first));
    assert_eq!(
        old.paragraph(Position::new(0)).unwrap().text().unwrap(),
        "before"
    );
    assert_eq!(
        first.paragraph(Position::new(0)).unwrap().text().unwrap(),
        "after"
    );
    let mut edit = current.edit().unwrap();
    edit.replace_paragraph_text(Position::new(0), "next")
        .unwrap();
    let next = edit.commit().unwrap().into_package();
    assert_eq!(
        next.document()
            .unwrap()
            .body_paragraph(Position::new(0))
            .unwrap()
            .unwrap()
            .text()
            .unwrap(),
        "next"
    );
    assert_eq!(
        first.paragraph(Position::new(0)).unwrap().text().unwrap(),
        "after"
    );
}

/// One source revision's unmanaged transaction snapshot is scanned on first
/// use, shared by every edit and commit started from it (record 0305), and an
/// edit from the memo publishes what an edit from a fresh scan publishes.
#[test]
fn unmanaged_edits_share_one_transaction_snapshot() {
    let bytes = docx(word_document(
        "",
        "<w:p><w:r><w:t>alpha</w:t></w:r></w:p><w:p><w:r><w:t>beta</w:t></w:r></w:p>",
    ));
    let package = Package::from_read_at(Arc::new(OwnedSource::new(bytes.clone()))).unwrap();
    let first = package.edit().unwrap();
    assert!(package.main_snapshot.lock().unwrap().is_none());
    let _ = first.source().unwrap();
    let held = package.main_snapshot.lock().unwrap().clone();
    let held = held.expect("the first projection memoizes its snapshot");
    let raw = held.xml_bytes().to_vec();
    let lineage = package.physical_package().unwrap().source_lineage();
    let version = package
        .physical_package()
        .unwrap()
        .source_version()
        .unwrap();
    let partname = PackURI::new("/word/document.xml").unwrap();
    assert!(
        package
            .memoized_main_snapshot(&raw, &lineage, version, &partname)
            .is_some(),
        "the memo answers for the same revision and bytes"
    );
    let mut other = raw.clone();
    other[0] = b' ';
    assert!(
        package
            .memoized_main_snapshot(&other, &lineage, version, &partname)
            .is_none(),
        "other bytes are never served from the memo"
    );
    let second = package.edit().unwrap();
    assert_eq!(
        first.source().unwrap().xml_bytes(),
        second.source().unwrap().xml_bytes()
    );
    let mut edit = second;
    edit.replace_paragraph_text(Position::new(1), "gamma")
        .unwrap();
    let mut memoized = Vec::new();
    edit.commit()
        .unwrap()
        .into_package()
        .write_to(&mut memoized)
        .unwrap();

    let fresh = Package::from_read_at(Arc::new(OwnedSource::new(bytes))).unwrap();
    let mut edit = fresh.edit().unwrap();
    edit.replace_paragraph_text(Position::new(1), "gamma")
        .unwrap();
    let mut expected = Vec::new();
    edit.commit()
        .unwrap()
        .into_package()
        .write_to(&mut expected)
        .unwrap();
    assert_eq!(memoized, expected);
    drop(first);
}

/// A handle committed from a source-backed owner edits its effective
/// revision through the same memo (record 0311): the revision is scanned
/// once for the edit and its commit, never served for other bytes, and the
/// chain publishes what a chain of fresh scans publishes.
#[test]
fn committed_handles_share_one_effective_transaction_snapshot() {
    let bytes = docx(word_document(
        "",
        "<w:p><w:r><w:t>alpha</w:t></w:r></w:p><w:p><w:r><w:t>beta</w:t></w:r></w:p>",
    ));
    let chain = |forget: bool| {
        let package = Package::from_read_at(Arc::new(OwnedSource::new(bytes.clone()))).unwrap();
        let mut edit = package.edit().unwrap();
        edit.replace_paragraph_text(Position::new(0), "first")
            .unwrap();
        let committed = edit.commit().unwrap().into_package();
        assert!(committed.effective.is_some());
        assert!(committed.main_snapshot.lock().unwrap().is_none());
        let mut edit = committed.edit().unwrap();
        let source = edit.source().unwrap();
        let held = committed
            .main_snapshot
            .lock()
            .unwrap()
            .clone()
            .expect("the first effective projection memoizes its snapshot");
        assert_eq!(held.xml_bytes(), source.xml_bytes());
        let topology = committed.effective.as_ref().unwrap();
        let effective = topology.main_document_part().unwrap().data().unwrap();
        assert_eq!(held.xml_bytes(), effective.as_bytes());
        let lineage = topology.physical_source_lineage();
        let version = topology.physical_source_version().unwrap();
        let partname = PackURI::new("/word/document.xml").unwrap();
        let physical = package.physical_package().unwrap();
        let original = physical.main_document_part().unwrap().data().unwrap();
        assert!(
            committed
                .memoized_main_snapshot(original.as_bytes(), &lineage, version, &partname)
                .is_none(),
            "the physical revision's bytes are not the effective revision's"
        );
        if forget {
            *committed.main_snapshot.lock().unwrap() = None;
        }
        edit.replace_paragraph_text(Position::new(1), "second")
            .unwrap();
        let mut published = Vec::new();
        edit.commit()
            .unwrap()
            .into_package()
            .write_to(&mut published)
            .unwrap();
        published
    };
    assert_eq!(chain(false), chain(true));
}

/// Record 0317: the transaction snapshot carries the source XML audit proof
/// the OPC package retained for its exact bytes, so the preserving commit
/// proves its candidate from it instead of scanning the base again, and the
/// candidate's proof keeps boundaries of its own for the next commit. Every
/// chain publishes what one fresh edit of the same paragraphs publishes.
#[test]
fn preserving_commits_prove_their_candidates_from_retained_proofs() {
    let mut body = String::new();
    for index in 0..40 {
        body.push_str(&format!(
            "<w:p><w:r><w:t>Paragraph {index}</w:t></w:r></w:p>"
        ));
    }
    let bytes = docx(word_document("", &body));
    let open = || Package::from_read_at(Arc::new(OwnedSource::new(bytes.clone()))).unwrap();
    let write = |package: &Package| {
        let mut out = Vec::new();
        package.write_to(&mut out).unwrap();
        out
    };
    let source_proof = |package: &Package| {
        let (_, snapshot) = package.main_document_snapshot("edit", true).unwrap();
        snapshot
            .publication_proof()
            .cloned()
            .map(|proof| (proof, snapshot))
    };
    let package = open();
    // Nothing has been audited yet: the first commit takes the pair audit.
    assert!(source_proof(&package).is_none());
    let mut edit = package.edit().unwrap();
    edit.replace_paragraph_text(Position::new(2), "two")
        .unwrap();
    let (committed, patch) = edit.commit().unwrap().into_parts();
    let proof = patch
        .target()
        .publication_proof()
        .expect("the pair audit proves the candidate");
    assert!(proof.index_depth().is_some(), "an indexed pair audit");

    // The physical owner's next transaction starts from the retained proof.
    let (base, snapshot) =
        source_proof(&package).expect("the package retained the original's proof");
    assert!(core::ptr::eq(base.bytes().as_slice(), snapshot.xml_bytes()));
    let mut edit = package.edit().unwrap();
    edit.replace_paragraph_text(Position::new(30), "thirty")
        .unwrap();
    let (second, patch) = edit.commit().unwrap().into_parts();
    assert!(
        patch
            .target()
            .publication_proof()
            .is_some_and(|proof| proof.index_depth().is_some())
    );
    let mut fresh_edit = open().edit().unwrap();
    fresh_edit
        .replace_paragraph_text(Position::new(30), "thirty")
        .unwrap();
    assert_eq!(
        write(&second),
        write(&fresh_edit.commit().unwrap().into_package())
    );

    // The committed revision's transaction shares its proven allocation, and
    // a chain of commits far apart keeps a proof at every step.
    let mut current = committed;
    for (position, text) in [(38, "thirty-eight"), (10, "ten")] {
        let (source, snapshot) = source_proof(&current).expect("the revision keeps its proof");
        let topology = current.effective.as_ref().unwrap();
        let effective = topology.main_document_part().unwrap().data().unwrap();
        assert!(core::ptr::eq(
            source.bytes().as_slice(),
            effective.as_bytes()
        ));
        assert!(core::ptr::eq(snapshot.xml_bytes(), effective.as_bytes()));
        let mut edit = current.edit().unwrap();
        edit.replace_paragraph_text(Position::new(position), text)
            .unwrap();
        let (next, patch) = edit.commit().unwrap().into_parts();
        assert!(patch.target().publication_proof().is_some());
        current = next;
    }
    let mut fresh_edit = open().edit().unwrap();
    for (position, text) in [(2, "two"), (38, "thirty-eight"), (10, "ten")] {
        fresh_edit
            .replace_paragraph_text(Position::new(position), text)
            .unwrap();
    }
    assert_eq!(
        write(&current),
        write(&fresh_edit.commit().unwrap().into_package())
    );
}
