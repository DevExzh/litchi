//! Selector-first semantic reads for legacy source-built Numbers comments.
//!
//! The migration host can still produce a compact legacy table model whose
//! storage projection is intentionally outside strict `litchi-numbers`
//! document admission.  These tests exercise the private physical handoff
//! that keeps that historical read behavior available while preserving the
//! strict native ingress profile.

use std::{env, fs, io, path::Path};

use litchi_iwa_archive::Limits as ArchiveLimits;
use litchi_numbers::{
    CellPosition, Package, PackageReadOptions, PackageSemanticLimits, SheetSelector,
    TableCellCommentError, TableCellCommentLimitKind, TableCellCommentPath, TableSelector,
};

#[path = "../../litchi-numbers/tests/support/numbers_comment_fixture.rs"]
mod fixture;
use fixture::{
    Corruption, FixtureMode, form_based_source_built_comment_fixture,
    multitable_duplicate_name_fixture, source_built_comment_fixture,
    source_built_comment_fixture_with_metadata_payload,
    source_built_comment_fixture_with_table_style, source_built_comment_fixture_without_metadata,
};

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

const LEGACY_SHEET: &str = "Reply fixture sheet";
const LEGACY_TABLE: &str = "Reply fixture table";
const LEGACY_CELL: CellPosition = CellPosition::new(0, 0);

fn compatibility_root(
    source: &[u8],
    sheet: impl Into<SheetSelector<'static>>,
    table: impl Into<TableSelector<'static>>,
    position: CellPosition,
) -> Result<Option<litchi_numbers::TableCellComment>, TableCellCommentError> {
    Package::__table_cell_comment_from_bytes_for_compatibility(source, sheet, table, position)
}

fn compatibility_replies(
    source: &[u8],
    sheet: impl Into<SheetSelector<'static>>,
    table: impl Into<TableSelector<'static>>,
    position: CellPosition,
) -> Result<Box<[litchi_numbers::TableCellCommentReply]>, TableCellCommentError> {
    Package::__table_cell_comment_replies_from_bytes_for_compatibility(
        source, sheet, table, position,
    )
}

fn assert_legacy_semantics(source: &[u8], row: u32) -> TestResult {
    let root = compatibility_root(
        source,
        SheetSelector::index(0),
        TableSelector::index(0),
        CellPosition::new(row, 0),
    )?
    .ok_or_else(|| io::Error::other("legacy source-built root comment is missing"))?;
    assert_eq!(root.text(), "root comment");
    assert_eq!(root.timestamp().map(|value| value.as_f64()), Some(123.0));
    let author = root
        .author()
        .ok_or_else(|| io::Error::other("legacy root author is missing"))?;
    assert_eq!(author.display_name(), Some("Reply fixture author"));
    assert_eq!(author.public_id(), Some("reply-fixture-author"));

    let named = compatibility_root(
        source,
        SheetSelector::name(LEGACY_SHEET),
        TableSelector::name(LEGACY_TABLE),
        CellPosition::from_a1("A1")?,
    )?
    .ok_or_else(|| io::Error::other("named legacy root comment is missing"))?;
    assert_eq!(named, root);

    let replies = compatibility_replies(
        source,
        SheetSelector::index(0),
        TableSelector::index(0),
        CellPosition::new(row, 0),
    )?;
    assert_eq!(replies.len(), 1);
    assert_eq!(replies[0].text(), "first reply");
    assert_eq!(
        replies[0].timestamp().map(|value| value.as_f64()),
        Some(123.0)
    );
    let reply_author = replies[0]
        .author()
        .ok_or_else(|| io::Error::other("legacy reply author is missing"))?;
    assert_eq!(reply_author.display_name(), Some("Reply fixture author"));
    assert_eq!(reply_author.public_id(), Some("reply-fixture-author"));
    Ok(())
}

fn export_compatibility_fuzz_fixture(name: &str, source: &[u8]) -> TestResult {
    let Ok(directory) = env::var("LITCHI_NUMBERS_COMMENT_COMPAT_FUZZ_DIRECTORY") else {
        return Ok(());
    };
    let directory = Path::new(&directory);
    fs::create_dir_all(directory)?;
    fs::write(directory.join(name), source)?;
    Ok(())
}

fn assert_source_unchanged(source: &[u8], baseline: &[u8]) {
    assert_eq!(
        source, baseline,
        "selector-first compatibility reads must not normalize or rewrite the source"
    );
}

#[test]
fn legacy_source_built_comment_reads_match_semantics_and_preserve_source() -> TestResult {
    let source = source_built_comment_fixture(false)?;
    let baseline = source.clone();

    // This is the compatibility boundary: the source-built table model is
    // intentionally not promoted to strict rooted document admission.
    assert!(Package::from_bytes(&source).is_err());

    assert_legacy_semantics(&source, 0)?;
    assert_source_unchanged(&source, &baseline);
    export_compatibility_fuzz_fixture("source-built-single-root.numbers", &source)?;
    Ok(())
}

#[test]
fn legacy_shared_root_reads_by_each_cell_without_global_key_aliasing() -> TestResult {
    let source = source_built_comment_fixture(true)?;
    let baseline = source.clone();
    assert!(Package::from_bytes(&source).is_err());

    assert_legacy_semantics(&source, 0)?;
    assert_legacy_semantics(&source, 1)?;
    assert_source_unchanged(&source, &baseline);
    export_compatibility_fuzz_fixture("source-built-shared-root.numbers", &source)?;
    Ok(())
}

#[test]
fn legacy_comment_compatibility_selectors_and_bounds_are_typed() -> TestResult {
    let source = source_built_comment_fixture(false)?;

    assert!(matches!(
        compatibility_root(
            &source,
            SheetSelector::index(1),
            TableSelector::index(0),
            LEGACY_CELL,
        ),
        Err(TableCellCommentError::SheetNotFound)
    ));
    assert!(matches!(
        compatibility_root(
            &source,
            SheetSelector::name("missing sheet"),
            TableSelector::index(0),
            LEGACY_CELL,
        ),
        Err(TableCellCommentError::SheetNotFound)
    ));
    assert!(matches!(
        compatibility_root(
            &source,
            SheetSelector::index(0),
            TableSelector::index(1),
            LEGACY_CELL,
        ),
        Err(TableCellCommentError::TableNotFound)
    ));
    assert!(matches!(
        compatibility_root(
            &source,
            SheetSelector::index(0),
            TableSelector::name("missing table"),
            LEGACY_CELL,
        ),
        Err(TableCellCommentError::TableNotFound)
    ));
    assert!(matches!(
        compatibility_root(
            &source,
            SheetSelector::index(0),
            TableSelector::index(0),
            CellPosition::new(99, 0),
        ),
        Err(TableCellCommentError::OutOfBounds { .. })
    ));
    assert!(matches!(
        compatibility_replies(
            &source,
            SheetSelector::index(0),
            TableSelector::index(0),
            CellPosition::new(99, 0),
        ),
        Err(TableCellCommentError::OutOfBounds { .. })
    ));

    // A caller-selected one-table ceiling still admits the inclusive table
    // zero position, while a later position remains a typed miss.
    let one_table = PackageSemanticLimits::new(
        PackageSemanticLimits::MAX_OBJECTS,
        PackageSemanticLimits::MAX_SHEETS,
        1,
        PackageSemanticLimits::MAX_REFERENCES,
    )?;
    let options = PackageReadOptions::new(ArchiveLimits::default(), one_table);
    assert!(
        Package::__table_cell_comment_from_bytes_for_compatibility_with_options(
            &source,
            options,
            SheetSelector::index(0),
            TableSelector::index(0),
            LEGACY_CELL,
        )?
        .is_some()
    );
    assert!(matches!(
        Package::__table_cell_comment_from_bytes_for_compatibility_with_options(
            &source,
            options,
            SheetSelector::index(0),
            TableSelector::index(1),
            LEGACY_CELL,
        ),
        Err(TableCellCommentError::TableNotFound)
    ));
    Ok(())
}

#[test]
fn form_based_legacy_sheet_preserves_index_and_name_selector_semantics() -> TestResult {
    let source = form_based_source_built_comment_fixture(false)?;
    let baseline = source.clone();
    assert!(Package::from_bytes(&source).is_err());

    assert_legacy_semantics(&source, 0)?;
    assert_source_unchanged(&source, &baseline);
    Ok(())
}

#[test]
fn legacy_type_6000_model_with_table_style_edge_still_reads() -> TestResult {
    let source = source_built_comment_fixture_with_table_style(false)?;
    let baseline = source.clone();
    assert!(Package::from_bytes(&source).is_err());

    // The legacy model uses type 6000 and has a nonzero table-style field,
    // which overlaps the native table-info reference probe. It must remain a
    // readable legacy model after payload-based disambiguation.
    assert_legacy_semantics(&source, 0)?;
    assert_source_unchanged(&source, &baseline);
    Ok(())
}

#[test]
fn duplicate_table_names_are_rejected_after_full_scan_and_limits_are_not_truncated() -> TestResult {
    let source = multitable_duplicate_name_fixture()?;
    let baseline = source.clone();

    let first = compatibility_root(
        &source,
        SheetSelector::index(0),
        TableSelector::index(0),
        LEGACY_CELL,
    )?
    .ok_or_else(|| io::Error::other("first duplicate-name table comment is missing"))?;
    assert_eq!(first.text(), "root comment");
    let second = compatibility_root(
        &source,
        SheetSelector::index(0),
        TableSelector::index(1),
        LEGACY_CELL,
    )?
    .ok_or_else(|| io::Error::other("second duplicate-name table comment is missing"))?;
    assert_eq!(second.text(), "second table root");

    assert!(matches!(
        compatibility_root(
            &source,
            SheetSelector::index(0),
            TableSelector::name(LEGACY_TABLE),
            LEGACY_CELL,
        ),
        Err(TableCellCommentError::InvalidSource { .. })
    ));

    let one_table = PackageSemanticLimits::new(
        PackageSemanticLimits::MAX_OBJECTS,
        PackageSemanticLimits::MAX_SHEETS,
        1,
        PackageSemanticLimits::MAX_REFERENCES,
    )?;
    let options = PackageReadOptions::new(ArchiveLimits::default(), one_table);
    for table in [TableSelector::index(0), TableSelector::name(LEGACY_TABLE)] {
        assert!(matches!(
            Package::__table_cell_comment_from_bytes_for_compatibility_with_options(
                &source,
                options,
                SheetSelector::index(0),
                table,
                LEGACY_CELL,
            ),
            Err(TableCellCommentError::LimitExceeded {
                kind: TableCellCommentLimitKind::References,
                observed: 2,
                maximum: 1,
                path: TableCellCommentPath::Package,
            })
        ));
    }
    assert_source_unchanged(&source, &baseline);
    Ok(())
}

#[test]
fn legacy_compatibility_reply_reads_reject_nested_graphs() -> TestResult {
    let source = fixture::fixture(
        FixtureMode::SingleRoot,
        Some(Corruption::NestedReplyReference),
    )?;
    let baseline = source.clone();
    assert!(matches!(
        compatibility_replies(
            &source,
            SheetSelector::index(0),
            TableSelector::index(0),
            LEGACY_CELL,
        ),
        Err(TableCellCommentError::InvalidSource { .. })
    ));
    assert_source_unchanged(&source, &baseline);
    Ok(())
}

#[test]
fn legacy_compatibility_reply_reads_reject_cycles_without_partial_results() -> TestResult {
    let source = fixture::fixture(
        FixtureMode::SingleRoot,
        Some(Corruption::SelfReplyReference),
    )?;
    let baseline = source.clone();
    assert!(matches!(
        compatibility_replies(
            &source,
            SheetSelector::index(0),
            TableSelector::index(0),
            LEGACY_CELL,
        ),
        Err(TableCellCommentError::InvalidSource { .. })
    ));
    assert_source_unchanged(&source, &baseline);
    Ok(())
}

#[test]
fn compatibility_replies_allow_absent_metadata_but_reject_bad_present_metadata() -> TestResult {
    let absent = source_built_comment_fixture_without_metadata(false)?;
    let absent_baseline = absent.clone();
    assert_legacy_semantics_without_metadata(&absent)?;
    assert_source_unchanged(&absent, &absent_baseline);

    let empty_metadata = source_built_comment_fixture_with_metadata_payload(&[])?;
    let empty_baseline = empty_metadata.clone();
    assert!(matches!(
        compatibility_replies(
            &empty_metadata,
            SheetSelector::index(0),
            TableSelector::index(0),
            LEGACY_CELL,
        ),
        Err(TableCellCommentError::InvalidSource { .. })
    ));
    assert_source_unchanged(&empty_metadata, &empty_baseline);

    let malformed_metadata = source_built_comment_fixture_with_metadata_payload(&[0xff])?;
    let malformed_baseline = malformed_metadata.clone();
    let result = compatibility_replies(
        &malformed_metadata,
        SheetSelector::index(0),
        TableSelector::index(0),
        LEGACY_CELL,
    );
    assert!(
        matches!(&result, Err(TableCellCommentError::InvalidSource { .. })),
        "present malformed metadata must remain a strict compatibility error: {result:?}"
    );
    assert_source_unchanged(&malformed_metadata, &malformed_baseline);

    // The strict reader keeps its metadata policy unchanged.  Some package
    // versions reject the missing member at ingress; versions that can still
    // project the table must reject replies when the metadata registry is
    // unavailable.
    let strict_missing =
        fixture::fixture(FixtureMode::SingleRoot, Some(Corruption::MissingMetadata))?;
    if let Ok(package) = Package::from_bytes(&strict_missing) {
        assert!(matches!(
            package.table_cell_comment_replies(
                SheetSelector::index(0),
                TableSelector::index(0),
                LEGACY_CELL,
            ),
            Err(TableCellCommentError::InvalidSource { .. })
        ));
    }
    Ok(())
}

fn assert_legacy_semantics_without_metadata(source: &[u8]) -> TestResult {
    let root = compatibility_root(
        source,
        SheetSelector::index(0),
        TableSelector::index(0),
        LEGACY_CELL,
    )?
    .ok_or_else(|| io::Error::other("legacy root is missing without metadata member"))?;
    assert_eq!(root.text(), "root comment");
    let replies = compatibility_replies(
        source,
        SheetSelector::index(0),
        TableSelector::index(0),
        LEGACY_CELL,
    )?;
    assert_eq!(replies.len(), 1);
    assert_eq!(replies[0].text(), "first reply");
    Ok(())
}

fn options_with_archive_input_limit(input_limit: usize) -> TestResult<PackageReadOptions> {
    let input_limit = u64::try_from(input_limit)?;
    let archive = ArchiveLimits::new(
        input_limit,
        ArchiveLimits::MAX_ENTRIES,
        ArchiveLimits::MAX_ENTRY_BYTES,
        ArchiveLimits::MAX_TOTAL_BYTES,
        ArchiveLimits::MAX_IWA_STREAM_BYTES,
    )?;
    Ok(PackageReadOptions::new(
        archive,
        PackageSemanticLimits::default(),
    ))
}

#[test]
fn legacy_comment_compatibility_honors_archive_object_reference_and_text_limits() -> TestResult {
    let source = source_built_comment_fixture(false)?;

    let archive_limited = options_with_archive_input_limit(source.len().saturating_sub(1))?;
    assert!(matches!(
        Package::__table_cell_comment_from_bytes_for_compatibility_with_options(
            &source,
            archive_limited,
            SheetSelector::index(0),
            TableSelector::index(0),
            LEGACY_CELL,
        ),
        Err(TableCellCommentError::LimitExceeded {
            kind: TableCellCommentLimitKind::InputBytes,
            observed: source_len,
            maximum: maximum_len,
            path: TableCellCommentPath::Package,
        }) if source_len == source.len() && maximum_len == source.len().saturating_sub(1)
    ));

    // The configured archive input ceiling is inclusive at the source
    // boundary.  A successful read here also guards against charging the
    // package's ZIP framing twice.
    let archive_at_source_size = options_with_archive_input_limit(source.len())?;
    assert!(
        Package::__table_cell_comment_from_bytes_for_compatibility_with_options(
            &source,
            archive_at_source_size,
            SheetSelector::index(0),
            TableSelector::index(0),
            LEGACY_CELL,
        )?
        .is_some()
    );

    let objects_limited = PackageSemanticLimits::new(
        1,
        PackageSemanticLimits::MAX_SHEETS,
        PackageSemanticLimits::MAX_TABLES,
        PackageSemanticLimits::MAX_REFERENCES,
    )?;
    let object_options = PackageReadOptions::new(ArchiveLimits::default(), objects_limited);
    let object_result = Package::__table_cell_comment_from_bytes_for_compatibility_with_options(
        &source,
        object_options,
        SheetSelector::index(0),
        TableSelector::index(0),
        LEGACY_CELL,
    );
    assert!(matches!(
        object_result,
        Err(TableCellCommentError::LimitExceeded {
            kind: TableCellCommentLimitKind::References,
            observed: 12,
            maximum: 1,
            path: TableCellCommentPath::Package,
        })
    ));

    let references_limited = PackageSemanticLimits::new(
        PackageSemanticLimits::MAX_OBJECTS,
        PackageSemanticLimits::MAX_SHEETS,
        PackageSemanticLimits::MAX_TABLES,
        1,
    )?;
    let reference_options = PackageReadOptions::new(ArchiveLimits::default(), references_limited);
    let reference_result =
        Package::__table_cell_comment_replies_from_bytes_for_compatibility_with_options(
            &source,
            reference_options,
            SheetSelector::index(0),
            TableSelector::index(0),
            LEGACY_CELL,
        );
    assert!(matches!(
        reference_result,
        Err(TableCellCommentError::LimitExceeded {
            kind: TableCellCommentLimitKind::References,
            observed: 2,
            maximum: 1,
            path: TableCellCommentPath::Cell {
                sheet: 0,
                table: 0,
                row: 0,
                column: 0,
            },
        })
    ));

    let text_limited = PackageSemanticLimits::default().with_projection_limits(1, 4)?;
    let text_options = PackageReadOptions::new(ArchiveLimits::default(), text_limited);
    let text_result = Package::__table_cell_comment_from_bytes_for_compatibility_with_options(
        &source,
        text_options,
        SheetSelector::index(0),
        TableSelector::index(0),
        LEGACY_CELL,
    );
    assert!(matches!(
        text_result,
        Err(TableCellCommentError::LimitExceeded {
            kind: TableCellCommentLimitKind::TextBytes,
            observed: 22,
            maximum: 4,
            path: TableCellCommentPath::Cell {
                sheet: 0,
                table: 0,
                row: 0,
                column: 0,
            },
        })
    ));
    Ok(())
}

#[test]
fn native_comment_ingress_remains_strict_while_compatibility_matches_it() -> TestResult {
    let source = include_bytes!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/iwork/numbers/comment-reader-native-created.numbers"
    ));
    let package = Package::from_bytes(source)?;
    let position = CellPosition::new(1, 1);
    let strict = package
        .table_cell_comment(SheetSelector::index(0), TableSelector::index(0), position)?
        .ok_or_else(|| io::Error::other("native root comment is missing"))?;
    let compatibility = compatibility_root(
        source,
        SheetSelector::index(0),
        TableSelector::index(0),
        position,
    )?
    .ok_or_else(|| io::Error::other("native compatibility root is missing"))?;
    assert_eq!(compatibility, strict);
    assert_eq!(compatibility.text(), "Native root comment");
    assert_eq!(
        compatibility
            .author()
            .and_then(|author| author.display_name()),
        Some("Ryker Zhu")
    );
    assert!(package.table_cell_comment_replies(
        SheetSelector::index(0),
        TableSelector::index(0),
        position,
    )?.is_empty());
    assert!(
        compatibility_replies(
            source,
            SheetSelector::index(0),
            TableSelector::index(0),
            position,
        )?
        .is_empty()
    );
    Ok(())
}

#[test]
fn native_one_table_ceiling_admits_comment_and_empty_replies() -> TestResult {
    let source = include_bytes!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/iwork/numbers/comment-reader-native-created.numbers"
    ));
    let semantic = PackageSemanticLimits::new(
        PackageSemanticLimits::MAX_OBJECTS,
        PackageSemanticLimits::MAX_SHEETS,
        1,
        PackageSemanticLimits::MAX_REFERENCES,
    )?;
    let options = PackageReadOptions::new(ArchiveLimits::default(), semantic);
    let position = CellPosition::new(1, 1);

    // The strict package path and the compatibility selector-first path must
    // agree that the native archive contains one semantic table.  In
    // particular, the native table-info envelope is metadata for table zero,
    // not an additional table that consumes the caller's ceiling.
    let package = Package::from_bytes_with_options(source, options)?;
    let strict_root = package
        .table_cell_comment(SheetSelector::index(0), TableSelector::index(0), position)?
        .ok_or_else(|| io::Error::other("native root comment is missing at one-table limit"))?;
    assert_eq!(strict_root.text(), "Native root comment");
    assert!(package
        .table_cell_comment_replies(
            SheetSelector::index(0),
            TableSelector::index(0),
            position,
        )?
        .is_empty());

    let compatibility = Package::__table_cell_comment_from_bytes_for_compatibility_with_options(
        source,
        options,
        SheetSelector::index(0),
        TableSelector::index(0),
        position,
    )?
    .ok_or_else(|| io::Error::other("native compatibility root is missing at one-table limit"))?;
    assert_eq!(compatibility, strict_root);
    assert!(
        Package::__table_cell_comment_replies_from_bytes_for_compatibility_with_options(
            source,
            options,
            SheetSelector::index(0),
            TableSelector::index(0),
            position,
        )?
        .is_empty()
    );
    Ok(())
}
