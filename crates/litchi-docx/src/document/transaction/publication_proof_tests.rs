//! Change 0754: the preserving commit route answers its source gate with the
//! pair audit and proves its candidate to the package writer.
//!
//! The route must be the historical one — gate, then paragraph compaction or
//! whole-document compaction — with the same bytes and the same errors, and a
//! proof must exist exactly when the writer's own audit of the candidate would
//! pass, covering exactly the candidate's allocation. Publishing with the
//! proof must emit the archive publishing without it emits.

use super::*;

/// The commit route before change 0754: the source gate first, then the
/// route it chose.
fn historical_route(base: &Snapshot, projected: &Snapshot) -> TransactionResult<Snapshot> {
    if publication_accepts_preserved_xml(base.xml_bytes()) {
        compact_changed_paragraphs(base, projected)
    } else {
        compact_whole_document(projected)
    }
}

fn outcome(result: &TransactionResult<Snapshot>) -> String {
    match result {
        Ok(snapshot) => format!("ok {:?}", snapshot.xml_bytes()),
        Err(error) => format!("err {error} | {error:?}"),
    }
}

/// Compare the route with the historical one for one projection, and check
/// the proof invariant.
fn assert_route(base: &Snapshot, projected: &Snapshot) -> Option<bool> {
    let committed = commit_preserving_unmodified(base, projected);
    let historical = historical_route(base, projected);
    assert_eq!(
        outcome(&committed),
        outcome(&historical),
        "commit route diverged from the historical gate"
    );
    let candidate = committed.ok()?;
    let limits = xml_minifier::audit::Limits::default();
    let writer_accepts = xml_minifier::audit::verify_source(candidate.xml_bytes(), limits).is_ok();
    let gate = publication_accepts_preserved_xml(base.xml_bytes());
    match candidate.publication_proof() {
        Some(proof) => {
            assert!(
                gate && writer_accepts,
                "a proof exists for unaccepted bytes"
            );
            assert!(proof.covers(candidate.xml_bytes(), limits));
            assert!(Arc::ptr_eq(proof.bytes(), &candidate.shared_xml().unwrap()));
            Some(true)
        },
        None => {
            // Only the whole-document fallback and refused candidates go
            // without a proof on an owned snapshot.
            assert!(
                !gate || !writer_accepts,
                "an accepted preserved candidate carries no proof"
            );
            Some(false)
        },
    }
}

fn fixture_documents() -> Vec<Vec<u8>> {
    let mut paths = Vec::new();
    let mut stack = vec![
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("..")
            .join("..")
            .join("test-data"),
    ];
    while let Some(directory) = stack.pop() {
        let Ok(entries) = std::fs::read_dir(&directory) else {
            continue;
        };
        for entry in entries.flatten() {
            let path = entry.path();
            if path.is_dir() {
                stack.push(path);
            } else if path
                .extension()
                .is_some_and(|extension| extension.eq_ignore_ascii_case("docx"))
            {
                paths.push(path);
            }
        }
    }
    paths.sort();
    paths
        .into_iter()
        .filter_map(|path| {
            let bytes = std::fs::read(&path).ok()?;
            let reader = soapberry_zip::office::ArchiveReader::new(&bytes).ok()?;
            reader.read("word/document.xml").ok()
        })
        .collect()
}

/// Every one-paragraph rewrite the edit accepts, for a few positions.
fn projections(base: &Snapshot) -> Vec<Snapshot> {
    let count = base.paragraph_count();
    let mut positions = vec![0, count / 2, count.saturating_sub(1)];
    positions.dedup();
    positions
        .into_iter()
        .filter(|position| *position < count)
        .filter_map(|position| {
            let mut edit = base.edit();
            edit.replace_paragraph_text(Position::new(position), "proof edit <&>")
                .ok()?;
            Some(edit.projected().clone())
        })
        .collect()
}

#[test]
fn the_route_and_proof_follow_the_historical_gate_on_every_fixture() {
    let mut proven = 0usize;
    let mut unproven = 0usize;
    for xml in fixture_documents() {
        let Ok(base) = Snapshot::from_xml(xml) else {
            continue;
        };
        for projected in projections(&base) {
            match assert_route(&base, &projected) {
                Some(true) => proven += 1,
                Some(false) => unproven += 1,
                None => {},
            }
        }
    }
    // Every editable fixture paragraph publishes preserved and proven; the
    // fallback is exercised by the mutated sources and the refused candidate
    // below.
    assert!(proven > 40, "{proven} {unproven}");
}

#[test]
fn the_route_and_proof_follow_the_historical_gate_on_mutated_sources() {
    // Sources the scanner admits but the publication audit refuses take the
    // whole-document fallback, with no proof.
    let insertions = [
        "<x:undeclared/>",
        "<w:t>]]&gt;</w:t>",
        "<w:t>a]]>b</w:t>",
        "<w:r w:rsidR=\"a\" w:rsidR=\"b\"/>",
        "<w:t xml:space=\"sometimes\">a</w:t>",
        "<w:r xmlns:q=\"\"/>",
        "<!--a--b-->",
    ];
    let mut fallbacks = 0usize;
    for xml in fixture_documents().into_iter().take(20) {
        let Some(body) = xml.windows(8).position(|window| window == b"<w:body>") else {
            continue;
        };
        for insertion in insertions {
            let mut mutated = xml.clone();
            let at = body + b"<w:body>".len();
            mutated.splice(at..at, insertion.bytes());
            let Ok(base) = Snapshot::from_xml(mutated) else {
                continue;
            };
            for projected in projections(&base) {
                if assert_route(&base, &projected) == Some(false)
                    && !publication_accepts_preserved_xml(base.xml_bytes())
                {
                    fallbacks += 1;
                }
            }
        }
    }
    assert!(fallbacks > 20, "{fallbacks}");
}

#[test]
fn a_candidate_the_writer_would_refuse_carries_no_proof() {
    const WORD: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    let base = Snapshot::from_xml(format!(
        r#"<w:document xmlns:w="{WORD}"><w:body><w:p><w:r><w:t>a</w:t></w:r></w:p><w:p/></w:body></w:document>"#
    ))
    .unwrap();
    // The scanner admits an undeclared prefix inside a paragraph; the
    // publication audit does not.
    let projected = Snapshot::from_xml(format!(
        r#"<w:document xmlns:w="{WORD}"><w:body><w:p><w:r><x:y/><w:t>b</w:t></w:r></w:p><w:p/></w:body></w:document>"#
    ))
    .unwrap();
    assert!(publication_accepts_preserved_xml(base.xml_bytes()));
    assert!(!publication_accepts_preserved_xml(projected.xml_bytes()));
    assert_eq!(assert_route(&base, &projected), Some(false));
}

fn generated_package(paragraphs: usize) -> crate::Package {
    let mut package = crate::Package::new().unwrap();
    {
        let document = package.document_mut().unwrap();
        for index in 0..paragraphs {
            document.add_paragraph_with_text(&format!("paragraph {index:05}"));
        }
    }
    let mut output = std::io::Cursor::new(Vec::new());
    package.to_stream(&mut output).unwrap();
    crate::Package::from_reader(std::io::Cursor::new(output.into_inner())).unwrap()
}

fn saved(package: &mut crate::Package) -> Vec<u8> {
    let mut output = std::io::Cursor::new(Vec::new());
    package.to_stream(&mut output).unwrap();
    output.into_inner()
}

#[test]
fn publishing_with_the_proof_emits_what_publishing_without_it_emits() {
    let mut proven_package = generated_package(64);
    let mut edit = proven_package.edit_document().unwrap();
    edit.replace_paragraph_text(Position::new(7), "edited")
        .unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.snapshot().publication_proof().is_some());
    let published = proven_package.publish_document_commit(commit).unwrap();
    // The main part now shares the committed snapshot's allocation.
    assert!(std::ptr::eq(
        proven_package
            .opc_package()
            .main_document_part()
            .unwrap()
            .blob(),
        published.snapshot().xml_bytes()
    ));

    let mut unproven_package = generated_package(64);
    let mut edit = unproven_package.edit_document().unwrap();
    edit.replace_paragraph_text(Position::new(7), "edited")
        .unwrap();
    let mut commit = edit.commit().unwrap();
    commit.snapshot.publication = None;
    commit.patch.after.publication = None;
    assert!(commit.snapshot().publication_proof().is_none());
    let _published = unproven_package.publish_document_commit(commit).unwrap();

    assert_eq!(saved(&mut proven_package), saved(&mut unproven_package));

    // A second edit on the published package, and the inverse of the first,
    // publish the same bytes on both.
    for package in [&mut proven_package, &mut unproven_package] {
        let mut edit = package.edit_document().unwrap();
        assert!(edit.projected().publication_proof().is_none());
        edit.replace_paragraph_text(Position::new(9), "second")
            .unwrap();
        package.publish_document_edit(edit).unwrap();
    }
    assert_eq!(saved(&mut proven_package), saved(&mut unproven_package));
}

#[test]
fn a_source_with_identity_keeps_the_historical_gate_and_builds_no_proof() {
    // The source-backed writer audits its own pair and never reads a proof,
    // so a snapshot bound to a source identity must not pay for one.
    let package = generated_package(16);
    let plain = package.document_snapshot().unwrap();
    let mut archive = std::io::Cursor::new(Vec::new());
    let mut writable = generated_package(16);
    writable.to_stream(&mut archive).unwrap();
    let source_backed = crate::source_backed::Package::from_read_at(Arc::new(
        litchi_core::OwnedSource::new(archive.into_inner()),
    ))
    .unwrap();
    let base = source_backed.edit_document().unwrap().source().clone();
    assert!(base.xml.identity().is_some());
    assert_eq!(base.xml_bytes(), plain.xml_bytes());
    let mut edit = base.edit();
    edit.replace_paragraph_text(Position::new(3), "edited")
        .unwrap();
    let projected = edit.projected().clone();
    let committed = commit_preserving_unmodified(&base, &projected).unwrap();
    assert!(committed.publication_proof().is_none());
    assert_eq!(
        committed.xml_bytes(),
        historical_route(&base, &projected).unwrap().xml_bytes()
    );
    // The same edit on the identity-free snapshot is proven.
    let mut edit = plain.edit();
    edit.replace_paragraph_text(Position::new(3), "edited")
        .unwrap();
    let proven = commit_preserving_unmodified(&plain, edit.projected()).unwrap();
    assert!(proven.publication_proof().is_some());
    assert_eq!(proven.xml_bytes(), committed.xml_bytes());
}

#[test]
fn a_derived_snapshot_never_inherits_a_proof() {
    let package = generated_package(8);
    let mut edit = package.edit_document().unwrap();
    edit.replace_paragraph_text(Position::new(1), "edited")
        .unwrap();
    let commit = edit.commit().unwrap();
    let proven = commit.snapshot();
    assert!(proven.publication_proof().is_some());
    // A clone is the same allocation and keeps it.
    assert!(proven.clone().publication_proof().is_some());
    // Every derivation is another allocation and starts without one.
    let mut edit = proven.edit();
    assert!(edit.projected().publication_proof().is_some());
    edit.replace_paragraph_text(Position::new(2), "again")
        .unwrap();
    assert!(edit.projected().publication_proof().is_none());
    let rescanned = Snapshot::from_xml(proven.xml_bytes().to_vec()).unwrap();
    assert!(rescanned.publication_proof().is_none());
    // The inverse publishes the source snapshot, which never had one.
    assert!(
        commit
            .patch()
            .inverse()
            .target()
            .publication_proof()
            .is_none()
    );
    // A proof attached to other bytes is dropped.
    let other = xml_minifier::audit::VerifiedSource::verify(
        Arc::new(proven.xml_bytes().to_vec()),
        xml_minifier::audit::Limits::default(),
    )
    .unwrap();
    let mut stripped = proven.clone();
    stripped.publication = None;
    assert!(
        stripped
            .with_publication_proof(other)
            .publication_proof()
            .is_none()
    );
}
