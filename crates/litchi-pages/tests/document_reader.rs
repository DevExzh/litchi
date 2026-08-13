use std::path::{Path, PathBuf};
use std::sync::Arc;

use litchi_iwa_archive::{Limits as ArchiveLimits, package::Catalog};
use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
use litchi_iwa_protos::{tp, tsp, tswp};
use litchi_pages::{
    DEFAULT_MAX_TEXT_BYTES, Document, DocumentReadOptions, DocumentSourceLimits,
    MAX_DOCUMENT_PROPERTIES_BYTES, MAX_SECTIONS, Package, PackageError, ReadError, ReadLimitKind,
    SemanticLimits, Stats,
};
use prost::Message as _;

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

const EXPECTED_TEXT: &str =
    "Litchi native Pages fixture\nBuffa lazy-view migration verification\n2026-08-07";
const RETAINED_TEXT_BYTES: usize = EXPECTED_TEXT.len() + "Blank".len();

fn fixture(relative: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork")
        .join(relative)
}

fn assert_send_sync<T: Send + Sync>() {}

fn assert_content_free(error: &ReadError, sentinels: &[&str]) {
    let renderings = [error.to_string(), format!("{error:?}")];
    for sentinel in sentinels {
        for rendered in &renderings {
            assert!(
                !rendered.contains(sentinel),
                "public Pages error leaked sentinel {sentinel:?}: {rendered:?}"
            );
        }
    }

    let mut source = std::error::Error::source(error);
    while let Some(cause) = source {
        let renderings = [cause.to_string(), format!("{cause:?}")];
        for sentinel in sentinels {
            for rendered in &renderings {
                assert!(
                    !rendered.contains(sentinel),
                    "Pages error source leaked sentinel {sentinel:?}: {rendered:?}"
                );
            }
        }
        source = cause.source();
    }
    assert!(std::error::Error::source(error).is_none());
}

fn source_limits(max_input_bytes: u64) -> TestResult<DocumentSourceLimits> {
    let defaults = DocumentSourceLimits::default();
    Ok(DocumentSourceLimits::new(
        max_input_bytes,
        defaults.max_entries(),
        defaults.max_entry_bytes(),
        defaults.max_aggregate_bytes(),
        defaults.max_component_bytes(),
    )?)
}

fn document_options(source: DocumentSourceLimits, semantic: SemanticLimits) -> DocumentReadOptions {
    let options = DocumentReadOptions::new(source, semantic);
    assert_eq!(options.source(), source);
    assert_eq!(options.semantic(), semantic);
    options
}

fn copy_directory_fixture(target: &Path) -> TestResult {
    let source = fixture("directory/pages/basic.pages");
    std::fs::create_dir_all(target.join("Metadata"))?;
    std::fs::copy(source.join("Index.zip"), target.join("Index.zip"))?;
    for name in [
        "Properties.plist",
        "BuildVersionHistory.plist",
        "DocumentIdentifier",
    ] {
        std::fs::copy(
            source.join("Metadata").join(name),
            target.join("Metadata").join(name),
        )?;
    }
    Ok(())
}

fn synthetic_package(root: tp::DocumentArchive, storages: &[&[&str]]) -> TestResult<Vec<u8>> {
    let mut objects = vec![ArchiveObject::new(
        1,
        vec![RawMessage {
            type_: 10_000,
            data: root.encode_to_vec(),
        }],
    )?];
    for (offset, fragments) in storages.iter().enumerate() {
        objects.push(ArchiveObject::new(
            42 + u64::try_from(offset)?,
            vec![RawMessage {
                type_: 2_001,
                data: tswp::StorageArchive {
                    text: fragments.iter().map(|text| (*text).to_owned()).collect(),
                    ..Default::default()
                }
                .encode_to_vec(),
            }],
        )?);
    }
    let stream = SnappyStream::compress(&Archive { objects }.to_bytes()?)?;
    Ok(litchi_iwa_archive::package::to_bytes(
        [("Index/Document.iwa", stream.as_slice())],
        ArchiveLimits::default(),
    )?)
}

fn fallback_package(messages: Vec<RawMessage>) -> TestResult<Vec<u8>> {
    let objects = vec![
        ArchiveObject::new(
            1,
            vec![RawMessage {
                type_: 10_000,
                data: tp::DocumentArchive::default().encode_to_vec(),
            }],
        )?,
        ArchiveObject::new(42, messages)?,
    ];
    let stream = SnappyStream::compress(&Archive { objects }.to_bytes()?)?;
    Ok(litchi_iwa_archive::package::to_bytes(
        [("Index/Document.iwa", stream.as_slice())],
        ArchiveLimits::default(),
    )?)
}

fn cross_component_duplicate_package() -> TestResult<Vec<u8>> {
    let component = |messages| -> TestResult<Vec<u8>> {
        Ok(SnappyStream::compress(
            &Archive {
                objects: vec![ArchiveObject::new(1, messages)?],
            }
            .to_bytes()?,
        )?)
    };
    let document = component(vec![RawMessage {
        type_: 10_000,
        data: tp::DocumentArchive::default().encode_to_vec(),
    }])?;
    let other = component(vec![RawMessage {
        type_: 2_001,
        data: tswp::StorageArchive::default().encode_to_vec(),
    }])?;
    Ok(litchi_iwa_archive::package::to_bytes(
        [
            ("Index/Document.iwa", document.as_slice()),
            ("Index/Other.iwa", other.as_slice()),
        ],
        ArchiveLimits::default(),
    )?)
}

fn storage_message(type_: u32, text: &[&str]) -> RawMessage {
    RawMessage {
        type_,
        data: tswp::StorageArchive {
            text: text.iter().map(|fragment| (*fragment).to_owned()).collect(),
            ..Default::default()
        }
        .encode_to_vec(),
    }
}

#[test]
fn archive_free_document_has_zip_directory_and_package_parity() -> TestResult {
    assert_send_sync::<Document>();
    assert_send_sync::<Stats>();
    assert_send_sync::<DocumentReadOptions>();
    assert_send_sync::<SemanticLimits>();

    let package = Package::open(fixture("pages/basic.pages"))?;
    let zipped = Document::open(fixture("pages/basic.pages"))?;
    let directory = Document::open(fixture("directory/pages/basic.pages"))?;
    let bytes = std::fs::read(fixture("pages/basic.pages"))?;
    let borrowed = Document::from_bytes(&bytes)?;
    let shared_source: Arc<[u8]> = Arc::from(bytes);
    let weak_source = Arc::downgrade(&shared_source);
    let shared = Document::from_shared_bytes(shared_source)?;
    assert!(weak_source.upgrade().is_none());

    zipped.validate()?;
    directory.validate()?;
    assert_eq!(zipped.plain_text(), EXPECTED_TEXT);
    assert_eq!(directory.plain_text(), EXPECTED_TEXT);
    assert_eq!(borrowed.plain_text(), EXPECTED_TEXT);
    assert_eq!(shared.plain_text(), EXPECTED_TEXT);
    let section_summary = |document: &Document| {
        document
            .sections()
            .iter()
            .map(|section| {
                (
                    section.index(),
                    section.name().map(str::to_owned),
                    section.plain_text(),
                )
            })
            .collect::<Vec<_>>()
    };
    assert_eq!(section_summary(&directory), section_summary(&zipped));
    assert_eq!(section_summary(&borrowed), section_summary(&zipped));
    assert_eq!(section_summary(&shared), section_summary(&zipped));
    assert_eq!(
        section_summary(&zipped),
        section_summary(package.semantic_document())
    );
    assert_eq!(zipped.stats(), Some(package.stats()));
    assert_eq!(directory.stats(), zipped.stats());
    assert_eq!(borrowed.stats(), zipped.stats());
    assert_eq!(shared.stats(), zipped.stats());

    let zipped_snapshot = zipped.snapshot();
    let directory_snapshot = directory.snapshot();
    assert!(Arc::ptr_eq(
        &zipped.shared_sections(),
        &zipped_snapshot.shared_sections()
    ));
    assert!(Arc::ptr_eq(
        &directory.shared_sections(),
        &directory_snapshot.shared_sections()
    ));

    let semantic = package.semantic_document();
    assert!(semantic.metadata().is_none());
    assert_eq!(semantic.stats(), None);
    assert!(Arc::ptr_eq(
        &semantic.shared_sections(),
        &semantic.snapshot().shared_sections()
    ));

    let zipped_metadata = zipped
        .metadata()
        .ok_or_else(|| std::io::Error::other("ZIP Pages metadata is missing"))?;
    assert_eq!(zipped_metadata.application.as_deref(), Some("Pages"));
    assert_eq!(
        zipped_metadata.revision.as_deref(),
        Some("0::BEA64FAC-CE32-48F4-A217-1E9D5FB7912E")
    );
    assert_eq!(
        zipped_metadata.content_status.as_deref(),
        Some("Pages Format Version 14.4.1")
    );
    assert_eq!(
        zipped_metadata.identifier.as_deref(),
        Some("28BDDDCE-2DA2-4C07-9308-68C7F148CFF1")
    );
    for metadata in [borrowed.metadata(), shared.metadata()] {
        let metadata = metadata
            .ok_or_else(|| std::io::Error::other("byte-ingress Pages metadata is missing"))?;
        assert_eq!(metadata.revision, zipped_metadata.revision);
        assert_eq!(metadata.content_status, zipped_metadata.content_status);
        assert_eq!(metadata.identifier, zipped_metadata.identifier);
    }

    let directory_metadata = directory
        .metadata()
        .ok_or_else(|| std::io::Error::other("directory Pages metadata is missing"))?;
    assert_eq!(directory_metadata.application.as_deref(), Some("Pages"));
    assert_eq!(
        directory_metadata.revision.as_deref(),
        Some("0::B5E256A8-B968-4724-8245-8E25B36B0CE2")
    );
    assert_eq!(
        directory_metadata.content_status.as_deref(),
        Some("Pages Format Version 14.4.1")
    );
    assert_eq!(
        directory_metadata.identifier.as_deref(),
        Some("DE1FA898-8D15-4D3B-99A9-965B526B2C6E")
    );
    Ok(())
}

#[test]
fn exact_package_reader_refuses_directory_backing() {
    assert!(matches!(
        Package::open(fixture("directory/pages/basic.pages")),
        Err(PackageError::InvalidFormat(_))
    ));
}

#[test]
fn other_iwork_families_are_typed_before_document_publication() {
    for path in [
        fixture("keynote/basic.key"),
        fixture("directory/keynote/basic.key"),
    ] {
        assert!(matches!(Document::open(path), Err(ReadError::NotPages)));
    }
}

#[test]
fn canonical_metadata_authorities_ignore_malformed_near_names() -> TestResult {
    let source = std::fs::read(fixture("pages/basic.pages"))?;
    let catalog = Catalog::from_bytes(&source)?;
    let properties = br#"<?xml version="1.0" encoding="UTF-8"?><plist version="1.0"><dict><key>Title</key><string>Canonical report</string><key>fileFormatVersion</key><string>99</string></dict></plist>"#;
    let build = br#"<?xml version="1.0" encoding="UTF-8"?><plist version="1.0"><array><string>old-build</string><string>canonical-build</string></array></plist>"#;
    let malformed = b"not a plist";
    let mut entries = vec![
        ("A/Properties.plist", malformed.as_slice()),
        ("Properties.plist", malformed.as_slice()),
        ("A/BuildVersionHistory.plist", malformed.as_slice()),
        ("BuildVersionHistory.plist", malformed.as_slice()),
        ("A/DocumentIdentifier", b"near-identifier".as_slice()),
        ("DocumentIdentifier", b"near-identifier".as_slice()),
        ("Metadata/Properties.plist", properties.as_slice()),
        ("Metadata/BuildVersionHistory.plist", build.as_slice()),
        ("Metadata/DocumentIdentifier", b"canonical-id\n".as_slice()),
    ];
    entries.extend(catalog.iter().filter_map(|entry| {
        (!matches!(
            entry.name(),
            "Metadata/Properties.plist"
                | "Metadata/BuildVersionHistory.plist"
                | "Metadata/DocumentIdentifier"
        ))
        .then_some((entry.name(), entry.data()))
    }));
    let candidate = litchi_iwa_archive::package::to_bytes(entries, ArchiveLimits::default())?;
    let temp = tempfile::tempdir()?;
    let path = temp.path().join("canonical.pages");
    std::fs::write(&path, candidate)?;

    let document = Document::open(path)?;
    let metadata = document
        .metadata()
        .ok_or_else(|| std::io::Error::other("canonical Pages metadata is missing"))?;
    assert_eq!(metadata.title.as_deref(), Some("Canonical report"));
    assert_eq!(metadata.revision.as_deref(), Some("canonical-build"));
    assert_eq!(
        metadata.content_status.as_deref(),
        Some("Pages Format Version 99")
    );
    assert_eq!(metadata.identifier.as_deref(), Some("canonical-id"));
    Ok(())
}

#[test]
fn frozen_directory_survives_source_removal_and_limits_are_inclusive() -> TestResult {
    let temp = tempfile::tempdir()?;
    let source = temp.path().join("detached.pages");
    copy_directory_fixture(&source)?;
    std::fs::create_dir(source.join("A"))?;
    std::fs::write(source.join("A/Properties.plist"), b"not a plist")?;
    std::fs::write(source.join("Properties.plist"), b"not a plist")?;

    let exact_input = [
        source.join("Index.zip"),
        source.join("Metadata/Properties.plist"),
        source.join("Metadata/BuildVersionHistory.plist"),
        source.join("Metadata/DocumentIdentifier"),
    ]
    .iter()
    .try_fold(0_u64, |total, path| {
        Ok::<_, std::io::Error>(total.saturating_add(std::fs::metadata(path)?.len()))
    })?;
    let exact_source = source_limits(exact_input)?;
    let exact_semantic = SemanticLimits::new(MAX_SECTIONS, RETAINED_TEXT_BYTES)?;
    let document =
        Document::open_with_options(&source, document_options(exact_source, exact_semantic))?;
    assert_eq!(document.plain_text(), EXPECTED_TEXT);

    let source_error = Document::open_with_options(
        &source,
        document_options(source_limits(exact_input - 1)?, SemanticLimits::default()),
    )
    .expect_err("source max-minus-one must fail before document publication");
    assert!(matches!(
        source_error,
        ReadError::Limit {
            kind: ReadLimitKind::InputBytes,
            observed,
            maximum,
        } if observed == exact_input && maximum == exact_input - 1
    ));

    let semantic_error = Document::open_with_options(
        &source,
        document_options(
            DocumentSourceLimits::default(),
            SemanticLimits::new(MAX_SECTIONS, RETAINED_TEXT_BYTES - 1)?,
        ),
    )
    .expect_err("semantic max-minus-one must fail before document publication");
    assert!(matches!(
        semantic_error,
        ReadError::Limit {
            kind: ReadLimitKind::TextBytes,
            observed,
            maximum,
        } if observed == RETAINED_TEXT_BYTES as u64
            && maximum == (RETAINED_TEXT_BYTES - 1) as u64
    ));

    std::fs::remove_dir_all(&source)?;
    document.validate()?;
    assert_eq!(document.plain_text(), EXPECTED_TEXT);
    assert_eq!(document.stats().map(Stats::section_count), Some(1));
    assert_eq!(
        document
            .metadata()
            .and_then(|metadata| metadata.identifier.as_deref()),
        Some("DE1FA898-8D15-4D3B-99A9-965B526B2C6E")
    );
    Ok(())
}

#[test]
fn canonical_properties_hard_limit_is_exact() -> TestResult {
    let source = std::fs::read(fixture("pages/basic.pages"))?;
    let catalog = Catalog::from_bytes(&source)?;
    let prefix = br#"<?xml version="1.0" encoding="UTF-8"?><plist version="1.0"><dict><key>Title</key><string>Limit</string></dict></plist>"#;
    let mut exact = prefix.to_vec();
    exact.resize(MAX_DOCUMENT_PROPERTIES_BYTES, b' ');

    for (length, should_pass) in [
        (MAX_DOCUMENT_PROPERTIES_BYTES, true),
        (MAX_DOCUMENT_PROPERTIES_BYTES + 1, false),
    ] {
        let mut properties = exact.clone();
        properties.resize(length, b' ');
        let mut entries = vec![("Metadata/Properties.plist", properties.as_slice())];
        entries.extend(catalog.iter().filter_map(|entry| {
            (entry.name() != "Metadata/Properties.plist").then_some((entry.name(), entry.data()))
        }));
        let candidate = litchi_iwa_archive::package::to_bytes(entries, ArchiveLimits::default())?;
        let temp = tempfile::tempdir()?;
        let path = temp.path().join("properties-limit.pages");
        std::fs::write(&path, &candidate)?;
        let result = Document::open(path);
        if should_pass {
            result?.validate()?;
            Document::from_bytes(&candidate)?.validate()?;
            Document::from_shared_bytes(Arc::from(candidate))?.validate()?;
        } else {
            let assert_limit = |result: Result<Document, ReadError>| {
                assert!(matches!(
                    result,
                    Err(ReadError::Limit {
                        kind: ReadLimitKind::PayloadBytes,
                        observed,
                        maximum,
                    }) if observed == length as u64
                        && maximum == MAX_DOCUMENT_PROPERTIES_BYTES as u64
                ));
            };
            assert_limit(result);
            assert_limit(Document::from_bytes(&candidate));
            assert_limit(Document::from_shared_bytes(Arc::from(candidate)));
        }
    }
    Ok(())
}

#[test]
fn rootless_fallback_preserves_storage_boundaries() -> TestResult {
    let bytes = synthetic_package(tp::DocumentArchive::default(), &[&["first", "second"]])?;
    let temp = tempfile::tempdir()?;
    let path = temp.path().join("rootless.pages");
    std::fs::write(&path, bytes)?;
    let exact = "first\nsecond".len();
    let document = Document::open_with_options(
        &path,
        document_options(
            DocumentSourceLimits::default(),
            SemanticLimits::new(MAX_SECTIONS, exact)?,
        ),
    )?;
    assert_eq!(document.plain_text(), "first\nsecond");
    assert_eq!(document.section_count(), 1);

    let error = Document::open_with_options(
        path,
        document_options(
            DocumentSourceLimits::default(),
            SemanticLimits::new(MAX_SECTIONS, exact - 1)?,
        ),
    )
    .expect_err("fallback separator max-minus-one must refuse");
    assert!(matches!(
        error,
        ReadError::Limit {
            kind: ReadLimitKind::TextBytes,
            observed,
            maximum,
        } if observed == exact as u64 && maximum == (exact - 1) as u64
    ));
    Ok(())
}

#[test]
fn rootless_fallback_joins_colocated_storage_messages_and_empty_fragments() -> TestResult {
    let bytes = fallback_package(vec![
        storage_message(2_001, &["first", ""]),
        storage_message(2_006, &["second"]),
    ])?;
    let temp = tempfile::tempdir()?;
    let path = temp.path().join("colocated.pages");
    std::fs::write(&path, bytes)?;
    let document = Document::open(path)?;
    assert_eq!(document.plain_text(), "first\n\nsecond");
    assert_eq!(document.section_count(), 1);
    assert_eq!(document.sections()[0].text_storages().len(), 1);
    Ok(())
}

#[test]
fn fallback_many_messages_enforce_the_aggregate_text_budget() -> TestResult {
    const MESSAGE_COUNT: usize = 64;
    let messages = (0..MESSAGE_COUNT)
        .map(|_| storage_message(2_001, &["x"]))
        .collect::<Vec<_>>();
    let bytes = fallback_package(messages)?;
    let exact = MESSAGE_COUNT * 2 - 1;
    let options = |maximum| {
        document_options(
            DocumentSourceLimits::default(),
            SemanticLimits::new(MAX_SECTIONS, maximum)
                .unwrap_or_else(|error| panic!("fallback limit must be valid: {error}")),
        )
    };

    let exact_document = Document::from_bytes_with_options(&bytes, options(exact))?;
    assert_eq!(exact_document.plain_text().len(), exact);
    assert_eq!(
        exact_document
            .plain_text()
            .bytes()
            .filter(|byte| *byte == b'\n')
            .count(),
        MESSAGE_COUNT - 1
    );

    let error = Document::from_bytes_with_options(&bytes, options(exact - 1))
        .expect_err("many-message aggregate max-minus-one must refuse");
    assert!(matches!(
        error,
        ReadError::Limit {
            kind: ReadLimitKind::TextBytes,
            observed,
            maximum,
        } if observed == exact as u64 && maximum == (exact - 1) as u64
    ));
    Ok(())
}

#[test]
fn rootless_fallback_charges_colocated_messages_before_materialization() -> TestResult {
    let bytes = fallback_package(vec![
        storage_message(2_001, &["four"]),
        storage_message(2_006, &["five"]),
    ])?;
    let exact = "four\nfive".len();
    let exact_document = Document::from_bytes_with_options(
        &bytes,
        document_options(
            DocumentSourceLimits::default(),
            SemanticLimits::new(MAX_SECTIONS, exact)?,
        ),
    )?;
    assert_eq!(exact_document.plain_text(), "four\nfive");

    let error = Document::from_bytes_with_options(
        &bytes,
        document_options(
            DocumentSourceLimits::default(),
            SemanticLimits::new(MAX_SECTIONS, exact - 1)?,
        ),
    )
    .expect_err("aggregate fallback max-minus-one must refuse before publication");
    assert!(matches!(
        error,
        ReadError::Limit {
            kind: ReadLimitKind::TextBytes,
            observed,
            maximum,
        } if observed == exact as u64 && maximum == (exact - 1) as u64
    ));
    Ok(())
}

#[test]
fn fallback_trigger_without_storage_text_remains_empty() -> TestResult {
    let bytes = fallback_package(vec![RawMessage {
        type_: 200,
        data: vec![0x1a, 0x05, b'b', b'o', b'g', b'u', b's'],
    }])?;
    let temp = tempfile::tempdir()?;
    let path = temp.path().join("trigger-only.pages");
    std::fs::write(&path, bytes)?;
    let document = Document::open(path)?;
    assert!(document.is_empty());
    assert_eq!(document.plain_text(), "");
    Ok(())
}

#[test]
fn malformed_root_graph_refuses_without_publishing_a_document() -> TestResult {
    let root = tp::DocumentArchive {
        body_storage: Some(tsp::Reference {
            identifier: 998_244_353,
            ..Default::default()
        }),
        ..Default::default()
    };
    let bytes = synthetic_package(root, &[])?;
    let temp = tempfile::tempdir()?;
    let path = temp.path().join("malformed-secret-marker.pages");
    std::fs::write(&path, bytes)?;
    let error = Document::open(path).expect_err("missing rooted body must refuse");
    assert!(matches!(error, ReadError::InvalidFormat));
    assert_content_free(&error, &["secret-marker", "998244353", "998_244_353"]);
    Ok(())
}

#[test]
fn cross_component_duplicate_object_identifier_refuses_before_publication() -> TestResult {
    let bytes = cross_component_duplicate_package()?;
    assert!(matches!(
        Package::from_bytes(&bytes),
        Err(PackageError::InvalidFormat(_))
    ));
    let error = Document::from_bytes(&bytes)
        .expect_err("cross-component duplicate object identity must refuse");
    assert!(matches!(error, ReadError::InvalidFormat));
    assert_content_free(&error, &["Index/Other.iwa", "object identifier"]);
    Ok(())
}

#[test]
fn public_read_errors_redact_paths_members_and_control_characters() -> TestResult {
    let temp = tempfile::tempdir()?;
    let path_sentinel = "private-path-sentinel-998244353";
    let path_error = Document::open(temp.path().join(path_sentinel))
        .expect_err("missing sentinel path must fail");
    assert!(matches!(path_error, ReadError::Io { .. }));
    assert_content_free(&path_error, &[path_sentinel, "998244353"]);

    let member_sentinel = "private-member-sentinel-776655443";
    let duplicate = litchi_iwa_archive::package::to_bytes(
        [
            ("Index/Document.iwa", member_sentinel.as_bytes()),
            ("Index/Document.iwa", b"duplicate".as_slice()),
        ],
        ArchiveLimits::default(),
    )?;
    let member_error =
        Document::from_bytes(&duplicate).expect_err("duplicate malformed member must fail");
    assert_content_free(&member_error, &[member_sentinel, "776655443"]);

    let control_sentinel = "private-control-sentinel\r\n\u{1b}[31m-112233445";
    let control_error = Document::from_bytes(control_sentinel.as_bytes())
        .expect_err("unrecognized control-character input must fail");
    assert!(matches!(control_error, ReadError::InvalidSource));
    assert_content_free(
        &control_error,
        &[control_sentinel, "private-control-sentinel", "112233445"],
    );
    Ok(())
}

#[test]
fn semantic_limit_constructor_is_checked() {
    assert_eq!(
        SemanticLimits::default().max_text_bytes(),
        DEFAULT_MAX_TEXT_BYTES
    );
    assert!(SemanticLimits::new(0, 1).is_err());
    assert!(SemanticLimits::new(1, 0).is_err());
    assert!(SemanticLimits::new(MAX_SECTIONS + 1, 1).is_err());
    assert!(SemanticLimits::new(1, DEFAULT_MAX_TEXT_BYTES + 1).is_err());

    let defaults = DocumentSourceLimits::default();
    let values = [
        (
            0,
            defaults.max_entries(),
            defaults.max_entry_bytes(),
            defaults.max_aggregate_bytes(),
            defaults.max_component_bytes(),
        ),
        (
            defaults.max_input_bytes(),
            0,
            defaults.max_entry_bytes(),
            defaults.max_aggregate_bytes(),
            defaults.max_component_bytes(),
        ),
        (
            defaults.max_input_bytes(),
            defaults.max_entries(),
            0,
            defaults.max_aggregate_bytes(),
            defaults.max_component_bytes(),
        ),
        (
            defaults.max_input_bytes(),
            defaults.max_entries(),
            defaults.max_entry_bytes(),
            0,
            defaults.max_component_bytes(),
        ),
        (
            defaults.max_input_bytes(),
            defaults.max_entries(),
            defaults.max_entry_bytes(),
            defaults.max_aggregate_bytes(),
            0,
        ),
    ];
    for (input, entries, entry, aggregate, component) in values {
        assert!(DocumentSourceLimits::new(input, entries, entry, aggregate, component).is_err());
    }

    let over_hard = [
        (
            defaults.max_input_bytes() + 1,
            defaults.max_entries(),
            defaults.max_entry_bytes(),
            defaults.max_aggregate_bytes(),
            defaults.max_component_bytes(),
        ),
        (
            defaults.max_input_bytes(),
            defaults.max_entries() + 1,
            defaults.max_entry_bytes(),
            defaults.max_aggregate_bytes(),
            defaults.max_component_bytes(),
        ),
        (
            defaults.max_input_bytes(),
            defaults.max_entries(),
            defaults.max_entry_bytes() + 1,
            defaults.max_aggregate_bytes(),
            defaults.max_component_bytes(),
        ),
        (
            defaults.max_input_bytes(),
            defaults.max_entries(),
            defaults.max_entry_bytes(),
            defaults.max_aggregate_bytes() + 1,
            defaults.max_component_bytes(),
        ),
        (
            defaults.max_input_bytes(),
            defaults.max_entries(),
            defaults.max_entry_bytes(),
            defaults.max_aggregate_bytes(),
            defaults.max_component_bytes() + 1,
        ),
    ];
    for (input, entries, entry, aggregate, component) in over_hard {
        assert!(DocumentSourceLimits::new(input, entries, entry, aggregate, component).is_err());
    }
}
