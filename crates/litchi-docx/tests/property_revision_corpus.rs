use std::collections::BTreeMap;
use std::path::{Path, PathBuf};

use litchi_docx::Package;
use litchi_docx::revision::{Limits, Revision, RevisionType};
use litchi_docx::source_backed;
use sha2::Digest as _;
use soapberry_zip::office::ArchiveReader;
use tempfile::NamedTempFile;

#[derive(Debug, PartialEq, Eq)]
struct RevisionSummary {
    kind: RevisionType,
    id: String,
    author: Option<String>,
    date: Option<String>,
    original_numbering: Option<String>,
}

fn summarize(records: &[Revision]) -> Vec<RevisionSummary> {
    records
        .iter()
        .map(|record| RevisionSummary {
            kind: record.revision_type(),
            id: record.id().to_owned(),
            author: record.author().map(str::to_owned),
            date: record.date().map(str::to_owned),
            original_numbering: record.original_numbering().map(str::to_owned),
        })
        .collect()
}

fn fixture(relative: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../..")
        .join(relative)
}

fn assert_contains_kind(records: &[Revision], kind: RevisionType) {
    assert!(
        records.iter().any(|record| record.revision_type() == kind),
        "revision corpus fixture did not expose {kind:?}: {records:?}"
    );
}

fn zip_members(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
    let archive = ArchiveReader::new(bytes).expect("parse ZIP corpus");
    archive
        .file_names()
        .map(|name| {
            (
                name.to_owned(),
                archive.read(name).expect("read ZIP corpus member"),
            )
        })
        .collect()
}

#[test]
#[ignore = "requires the ignored local 3rdparty DOCX corpus; run this test explicitly with --ignored"]
fn producer_and_native_revision_corpus_queries_survive_save_and_reopen() {
    // These fixtures are local producer/native-derived DOCX files. Their
    // package SHA-256 values are recorded here so a changed corpus is visible
    // in review without copying binary fixtures into this repository.
    let cases = [
        (
            "3rdparty/libreoffice-core/sw/qa/extras/ooxmlexport/data/tdf157011_ins_del_empty_cols.docx",
            9_954,
            "7c737a45db1ed60fa1a56d8a388e7e4a0d394de909fa9c5ca5d07060a6914794",
            RevisionType::TableGridChange,
        ),
        (
            "3rdparty/libreoffice-core/sw/qa/extras/ooxmlexport/data/n830205.docx",
            33_186,
            "9f27ef4097b11c6384c637bddaa1096e1b123623362d3acd8852eb9f83185cac",
            RevisionType::SectionPropertiesChange,
        ),
        (
            "3rdparty/libreoffice-core/sw/qa/extras/ooxmlexport/data/testTrackChangesInsertedTableCell.docx",
            13_842,
            "6f2d63be32322581fc9e185b3fd1e50993b93c8ba5f843ec6faef35514b11994",
            RevisionType::TableGridChange,
        ),
        (
            "3rdparty/libreoffice-core/sw/qa/extras/ooxmlexport/data/tdf89731.docx",
            44_680,
            "4cd35ae2e80f890f57da8efac0eead7a2caaafdda8abb79f27021d8e32d43977",
            RevisionType::NumberingChange,
        ),
        (
            "3rdparty/Open-XML-SDK/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestDataStorage/O14ISOStrict/Word/sectPr-Previous Section Properties-rsidSelect-004B4C75.docx",
            14_219,
            "ffd9bbadb66fd764e08477e4f9590f89c993b8e855133e0fcf276be4b006cbc3",
            RevisionType::SectionPropertiesChange,
        ),
    ];

    for (relative, expected_size, expected_sha256, expected_kind) in cases {
        let path = fixture(relative);
        assert!(
            path.is_file(),
            "missing local corpus fixture: {}",
            path.display()
        );
        let bytes = std::fs::read(&path).expect("read corpus fixture");
        assert_eq!(
            bytes.len(),
            expected_size,
            "corpus size changed: {relative}"
        );
        let original_members = zip_members(&bytes);
        let actual_sha256: String = sha2::Sha256::digest(&bytes)
            .iter()
            .map(|byte| format!("{byte:02x}"))
            .collect();
        assert_eq!(actual_sha256, expected_sha256, "corpus changed: {relative}");

        let mut package = Package::open(&path).expect("open owning package");
        let document = package.document().expect("read owning document");
        let records = document.revisions().expect("read owning revisions");
        assert_eq!(
            summarize(
                &document
                    .revisions_with_limits(Limits::default())
                    .expect("read owning bounded revisions")
            ),
            summarize(&records),
            "bounded owning query mismatch: {relative}"
        );
        assert_contains_kind(&records, expected_kind);
        if expected_kind == RevisionType::NumberingChange {
            let original: Vec<_> = records
                .iter()
                .filter(|record| record.revision_type() == RevisionType::NumberingChange)
                .map(|record| record.original_numbering())
                .collect();
            assert_eq!(original, [Some("1."), Some("4.")]);
        }
        let owning_summary = summarize(&records);

        let source = source_backed::Package::open(&path).expect("open source-backed package");
        let source_document = source.document().expect("read source-backed document");
        let source_records = source_document
            .revisions()
            .expect("read source-backed revisions");
        assert_eq!(
            summarize(
                &source_document
                    .revisions_with_limits(Limits::default())
                    .expect("read source-backed bounded revisions")
            ),
            owning_summary,
            "bounded source-backed query mismatch: {relative}"
        );
        assert_eq!(
            summarize(&source_records),
            owning_summary,
            "source-backed mismatch: {relative}"
        );

        let saved = NamedTempFile::with_suffix(".docx").expect("create save target");
        package.save(saved.path()).expect("save owning package");
        let saved_bytes = std::fs::read(saved.path()).expect("read saved package");
        assert_eq!(
            zip_members(&saved_bytes),
            original_members,
            "save changed an unprojected ZIP member: {relative}"
        );
        let reopened = Package::open(saved.path()).expect("reopen saved package");
        let reopened_records = reopened
            .document()
            .expect("read reopened document")
            .revisions()
            .expect("read reopened revisions");
        assert_eq!(
            summarize(&reopened_records),
            owning_summary,
            "save/reopen mismatch: {relative}"
        );
    }
}
