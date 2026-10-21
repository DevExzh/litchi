use std::path::{Path, PathBuf};
use std::sync::Arc;

use litchi_iwa_archive::{Limits as ArchiveLimits, package::Catalog};
use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
use litchi_iwa_protos::{tp, tsp, tswp};
use litchi_pages::{
    DEFAULT_MAX_TEXT_BYTES, Document, DocumentStats, Error as SemanticError, LimitKind,
    MAX_DOCUMENT_PROPERTIES_BYTES, MAX_SECTIONS, OpenOptions, Package, PackageError,
    SemanticLimitKind, SemanticLimits, SourceLimitKind, SourceLimits, Stats,
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

fn assert_content_free(error: &PackageError, sentinels: &[&str]) {
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
}

fn source_limits(max_input_bytes: u64) -> TestResult<SourceLimits> {
    Ok(SourceLimits::builder()
        .max_input_bytes(max_input_bytes)?
        .build()?)
}

fn text_limits(max_text_bytes: usize) -> TestResult<SemanticLimits> {
    Ok(SemanticLimits::builder()
        .max_sections(MAX_SECTIONS)?
        .max_text_bytes(max_text_bytes)?
        .build()?)
}

fn document_options(source: SourceLimits, semantic: SemanticLimits) -> OpenOptions {
    let options = OpenOptions::new()
        .source_limits(source)
        .semantic_limits(semantic);
    assert_eq!(options.source_limits_value(), source);
    assert_eq!(options.semantic_limits_value(), semantic);
    options
}

/// The projection of a package opened from `path`.
fn open_document(path: impl AsRef<Path>) -> Result<Document, PackageError> {
    Package::open(path).map(|package| package.document())
}

fn open_document_with(
    path: impl AsRef<Path>,
    options: OpenOptions,
) -> Result<Document, PackageError> {
    Package::open_with_options(path, options).map(|package| package.document())
}

fn bytes_document(bytes: &[u8]) -> Result<Document, PackageError> {
    Package::from_bytes(bytes).map(|package| package.document())
}

fn bytes_document_with(bytes: &[u8], options: OpenOptions) -> Result<Document, PackageError> {
    Package::from_bytes_with_options(bytes, options).map(|package| package.document())
}

/// Whether `error` is the semantic text budget refusing `observed` bytes
/// against `maximum`.
fn is_text_limit(error: &PackageError, observed: usize, maximum: usize) -> bool {
    matches!(
        error,
        PackageError::Semantic(SemanticError::TextTooLarge { observed: o, limit })
            if *o == observed && *limit == maximum
    ) || matches!(
        error,
        PackageError::PayloadLimit { observed: o, limit } if *o == observed && *limit == maximum
    )
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
        Default::default(),
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
            Default::default(),
        )?);
    }
    let stream = SnappyStream::compress(
        &Archive { objects }.to_bytes(Default::default())?,
        Default::default(),
    )?;
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
            Default::default(),
        )?,
        ArchiveObject::new(42, messages, Default::default())?,
    ];
    let stream = SnappyStream::compress(
        &Archive { objects }.to_bytes(Default::default())?,
        Default::default(),
    )?;
    Ok(litchi_iwa_archive::package::to_bytes(
        [("Index/Document.iwa", stream.as_slice())],
        ArchiveLimits::default(),
    )?)
}

fn cross_component_duplicate_package() -> TestResult<Vec<u8>> {
    let component = |messages| -> TestResult<Vec<u8>> {
        Ok(SnappyStream::compress(
            &Archive {
                objects: vec![ArchiveObject::new(1, messages, Default::default())?],
            }
            .to_bytes(Default::default())?,
            Default::default(),
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
fn package_projection_has_zip_directory_and_byte_parity() -> TestResult {
    assert_send_sync::<Document>();
    assert_send_sync::<DocumentStats>();
    assert_send_sync::<Stats>();
    assert_send_sync::<OpenOptions>();
    assert_send_sync::<SemanticLimits>();

    let package = Package::open(fixture("pages/basic.pages"))?;
    let zipped = open_document(fixture("pages/basic.pages"))?;
    let directory = open_document(fixture("directory/pages/basic.pages"))?;
    let bytes = std::fs::read(fixture("pages/basic.pages"))?;
    let borrowed = bytes_document(&bytes)?;
    let shared_source: Arc<[u8]> = Arc::from(bytes);
    let shared = Package::from_shared_bytes(Arc::clone(&shared_source))?.document();

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
        section_summary(package.document_ref())
    );
    assert_eq!(
        zipped.stats().map(DocumentStats::section_count),
        Some(package.stats().section_count())
    );
    assert_eq!(directory.stats(), zipped.stats());
    assert_eq!(borrowed.stats(), zipped.stats());
    assert_eq!(shared.stats(), zipped.stats());

    // The owned projection shares the package's decoded sections.
    assert!(Arc::ptr_eq(
        &package.document().shared_sections(),
        &package.document_ref().shared_sections()
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
fn package_directories_convert_to_editable_single_file_packages() -> TestResult {
    let package = Package::open(fixture("directory/pages/basic.pages"))?;
    // The converted package holds every bundle file beside the index members.
    let bytes = package.to_bytes()?;
    let reopened = Package::from_bytes(&bytes)?;
    assert_eq!(reopened.text()?, EXPECTED_TEXT);
    assert_eq!(reopened.to_bytes()?, bytes);
    // Deterministic: a second capture converts to the same bytes.
    assert_eq!(
        Package::open(fixture("directory/pages/basic.pages"))?.to_bytes()?,
        bytes
    );
    // Editable like any exact-source package.
    let commit = package.set_section_text(0_usize, "Directory edit")?;
    assert_eq!(commit.package().text()?, "Directory edit");
    Ok(())
}

#[test]
fn other_iwork_families_are_typed_before_document_publication() {
    for path in [
        fixture("keynote/basic.key"),
        fixture("directory/keynote/basic.key"),
    ] {
        assert!(matches!(
            Package::open(path),
            Err(PackageError::UnexpectedFormat {
                detected: litchi_core::FileFormat::Keynote
            })
        ));
    }
}

#[test]
fn canonical_metadata_authorities_ignore_malformed_near_names() -> TestResult {
    let source = std::fs::read(fixture("pages/basic.pages"))?;
    let catalog = Catalog::from_bytes(&source, Default::default())?;
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

    let document = open_document(path)?;
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
fn captured_directory_survives_source_removal_and_limits_are_inclusive() -> TestResult {
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
        source.join("A/Properties.plist"),
        source.join("Properties.plist"),
    ]
    .iter()
    .try_fold(0_u64, |total, path| {
        Ok::<_, std::io::Error>(total.saturating_add(std::fs::metadata(path)?.len()))
    })?;
    let exact_source = source_limits(exact_input)?;
    let exact_semantic = text_limits(RETAINED_TEXT_BYTES)?;
    let document = open_document_with(&source, document_options(exact_source, exact_semantic))?;
    assert_eq!(document.plain_text(), EXPECTED_TEXT);

    let source_error = open_document_with(
        &source,
        document_options(source_limits(exact_input - 1)?, SemanticLimits::default()),
    )
    .expect_err("source max-minus-one must fail before document publication");
    assert!(matches!(
        source_error,
        PackageError::Limit {
            kind: LimitKind::Source(SourceLimitKind::InputBytes),
            observed,
            maximum,
        } if observed == exact_input && maximum == exact_input - 1
    ));

    let semantic_error = open_document_with(
        &source,
        document_options(
            SourceLimits::default(),
            text_limits(RETAINED_TEXT_BYTES - 1)?,
        ),
    )
    .expect_err("semantic max-minus-one must fail before document publication");
    assert!(
        is_text_limit(
            &semantic_error,
            RETAINED_TEXT_BYTES,
            RETAINED_TEXT_BYTES - 1
        ),
        "{semantic_error:?}"
    );

    std::fs::remove_dir_all(&source)?;
    document.validate()?;
    assert_eq!(document.plain_text(), EXPECTED_TEXT);
    assert_eq!(document.stats().map(DocumentStats::section_count), Some(1));
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
    let catalog = Catalog::from_bytes(&source, Default::default())?;
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
        let result = open_document(path);
        if should_pass {
            result?.validate()?;
            bytes_document(&candidate)?.validate()?;
            Package::from_shared_bytes(Arc::from(candidate))?
                .document()
                .validate()?;
        } else {
            let assert_limit = |result: Result<Document, PackageError>| {
                assert!(matches!(
                    result,
                    Err(PackageError::PayloadLimit { observed, limit })
                        if observed == length && limit == MAX_DOCUMENT_PROPERTIES_BYTES
                ));
            };
            assert_limit(result);
            assert_limit(bytes_document(&candidate));
            assert_limit(Package::from_shared_bytes(Arc::from(candidate)).map(|p| p.document()));
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
    let document = open_document_with(
        &path,
        document_options(SourceLimits::default(), text_limits(exact)?),
    )?;
    assert_eq!(document.plain_text(), "first\nsecond");
    assert_eq!(document.section_count(), 1);

    let error = open_document_with(
        path,
        document_options(SourceLimits::default(), text_limits(exact - 1)?),
    )
    .expect_err("fallback separator max-minus-one must refuse");
    assert!(is_text_limit(&error, exact, exact - 1), "{error:?}");
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
    let document = open_document(path)?;
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
    let options = |maximum| -> TestResult<OpenOptions> {
        Ok(document_options(
            SourceLimits::default(),
            text_limits(maximum)?,
        ))
    };

    let exact_document = bytes_document_with(&bytes, options(exact)?)?;
    assert_eq!(exact_document.plain_text().len(), exact);
    assert_eq!(
        exact_document
            .plain_text()
            .bytes()
            .filter(|byte| *byte == b'\n')
            .count(),
        MESSAGE_COUNT - 1
    );

    let error = bytes_document_with(&bytes, options(exact - 1)?)
        .expect_err("many-message aggregate max-minus-one must refuse");
    assert!(is_text_limit(&error, exact, exact - 1), "{error:?}");
    Ok(())
}

#[test]
fn rootless_fallback_charges_colocated_messages_before_materialization() -> TestResult {
    let bytes = fallback_package(vec![
        storage_message(2_001, &["four"]),
        storage_message(2_006, &["five"]),
    ])?;
    let exact = "four\nfive".len();
    let exact_document = bytes_document_with(
        &bytes,
        document_options(SourceLimits::default(), text_limits(exact)?),
    )?;
    assert_eq!(exact_document.plain_text(), "four\nfive");

    let error = bytes_document_with(
        &bytes,
        document_options(SourceLimits::default(), text_limits(exact - 1)?),
    )
    .expect_err("aggregate fallback max-minus-one must refuse before publication");
    assert!(is_text_limit(&error, exact, exact - 1), "{error:?}");
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
    let document = open_document(path)?;
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
    let error = open_document(path).expect_err("missing rooted body must refuse");
    assert!(matches!(error, PackageError::InvalidFormat(_)));
    assert_content_free(&error, &["secret-marker"]);
    Ok(())
}

#[test]
fn cross_component_duplicate_object_identifier_refuses_before_publication() -> TestResult {
    let bytes = cross_component_duplicate_package()?;
    assert!(matches!(
        Package::from_bytes(&bytes),
        Err(PackageError::InvalidFormat(_))
    ));
    Ok(())
}

#[test]
fn public_read_errors_redact_paths() -> TestResult {
    let temp = tempfile::tempdir()?;
    let path_sentinel = "private-path-sentinel-998244353";
    let path_error = open_document(temp.path().join(path_sentinel))
        .expect_err("missing sentinel path must fail");
    assert!(matches!(path_error, PackageError::Io(_)));
    assert_content_free(&path_error, &[path_sentinel]);

    let control_sentinel = "private-control-sentinel\r\n\u{1b}[31m-112233445";
    let control_error = bytes_document(control_sentinel.as_bytes())
        .expect_err("unrecognized control-character input must fail");
    assert_content_free(
        &control_error,
        &[control_sentinel, "private-control-sentinel", "112233445"],
    );
    Ok(())
}

#[test]
fn limit_builders_are_checked() {
    assert_eq!(
        SemanticLimits::default().max_text_bytes(),
        DEFAULT_MAX_TEXT_BYTES
    );
    for result in [
        SemanticLimits::builder().max_sections(0).map(|_| ()),
        SemanticLimits::builder().max_text_bytes(0).map(|_| ()),
        SemanticLimits::builder()
            .max_sections(MAX_SECTIONS + 1)
            .map(|_| ()),
        SemanticLimits::builder()
            .max_text_bytes(DEFAULT_MAX_TEXT_BYTES + 1)
            .map(|_| ()),
    ] {
        assert!(matches!(
            result,
            Err(PackageError::InvalidLimit {
                kind: LimitKind::Semantic(
                    SemanticLimitKind::Sections | SemanticLimitKind::TextBytes
                ),
                ..
            })
        ));
    }
    let maximum = SourceLimits::MAXIMUM;
    for (kind, result) in [
        (
            SourceLimitKind::InputBytes,
            SourceLimits::builder().max_input_bytes(0).map(|_| ()),
        ),
        (
            SourceLimitKind::Members,
            SourceLimits::builder().max_members(0).map(|_| ()),
        ),
        (
            SourceLimitKind::EntryBytes,
            SourceLimits::builder().max_entry_bytes(0).map(|_| ()),
        ),
        (
            SourceLimitKind::AggregateBytes,
            SourceLimits::builder().max_aggregate_bytes(0).map(|_| ()),
        ),
        (
            SourceLimitKind::ComponentBytes,
            SourceLimits::builder().max_component_bytes(0).map(|_| ()),
        ),
        (
            SourceLimitKind::InputBytes,
            SourceLimits::builder()
                .max_input_bytes(maximum.max_input_bytes() + 1)
                .map(|_| ()),
        ),
        (
            SourceLimitKind::Members,
            SourceLimits::builder()
                .max_members(maximum.max_members() + 1)
                .map(|_| ()),
        ),
        (
            SourceLimitKind::EntryBytes,
            SourceLimits::builder()
                .max_entry_bytes(maximum.max_entry_bytes() + 1)
                .map(|_| ()),
        ),
        (
            SourceLimitKind::AggregateBytes,
            SourceLimits::builder()
                .max_aggregate_bytes(maximum.max_aggregate_bytes() + 1)
                .map(|_| ()),
        ),
        (
            SourceLimitKind::ComponentBytes,
            SourceLimits::builder()
                .max_component_bytes(maximum.max_component_bytes() + 1)
                .map(|_| ()),
        ),
    ] {
        assert!(matches!(
            result,
            Err(PackageError::InvalidLimit { kind: LimitKind::Source(found), .. }) if found == kind
        ));
    }
}
