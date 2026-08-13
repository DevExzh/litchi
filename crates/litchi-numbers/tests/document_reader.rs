use std::path::{Path, PathBuf};
use std::sync::Arc;

use litchi_iwa_archive::{Limits as ArchiveLimits, package::Catalog};
use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
use litchi_iwa_protos::tsce::ast_node_array_archive::{AstNodeArchive, AstNodeType};
use litchi_iwa_protos::{tn, tsce, tsp, tst, tswp};
use litchi_numbers::cell::Value;
use litchi_numbers::{
    DEFAULT_MAX_TEXT_BYTES, Dimensions, Document, DocumentLimits, DocumentReadError,
    DocumentReadLimitKind, DocumentReadOptions, DocumentSourceLimits, DocumentStats,
    MAX_MATERIALIZED_CELLS, MAX_SHEETS, MAX_TABLES, Package, PackageError, Position, SheetBuilder,
    TableBuilder,
};
use litchi_numbers_wire::BncCell;
use prost::Message as _;

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

const EXPECTED_TEXT: &str = "Sheet 1\nTable 1\nLitchi native Numbers fixture\n42";
const ZIP_IDENTIFIER: &str = "FEE1FC17-F975-41FB-84CF-3EE0B143D36F";
const ZIP_REVISION: &str = "0::7EF7375C-4D56-4B32-876D-FA3417CBFDF0";
const DIRECTORY_IDENTIFIER: &str = "14E88241-88E4-471F-9810-76861B7F4EFE";
const DIRECTORY_REVISION: &str = "0::F97FDBEE-A75C-447C-A238-82CC5A1FB47A";
const EXPECTED_VERSION: &str = "M14.4-7043.0.93-4";
const PROPERTIES: &str = "Metadata/Properties.plist";
const BUILD_HISTORY: &str = "Metadata/BuildVersionHistory.plist";
const DOCUMENT_IDENTIFIER: &str = "Metadata/DocumentIdentifier";

fn fixture(relative: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork")
        .join(relative)
}

fn assert_send_sync<T: Send + Sync>() {}

fn assert_content_free(error: &DocumentReadError, sentinels: &[&str]) {
    let renderings = [error.to_string(), format!("{error:?}")];
    for sentinel in sentinels {
        for rendering in &renderings {
            assert!(
                !rendering.contains(sentinel),
                "public Numbers error leaked sentinel {sentinel:?}: {rendering:?}"
            );
        }
    }
    let mut source = std::error::Error::source(error);
    while let Some(cause) = source {
        let renderings = [cause.to_string(), format!("{cause:?}")];
        for sentinel in sentinels {
            for rendering in &renderings {
                assert!(
                    !rendering.contains(sentinel),
                    "Numbers error source leaked sentinel {sentinel:?}: {rendering:?}"
                );
            }
        }
        source = cause.source();
    }
    assert!(std::error::Error::source(error).is_none());
}

fn copy_directory_fixture(target: &Path) -> TestResult {
    let source = fixture("directory/numbers/basic.numbers");
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

fn options(source: DocumentSourceLimits, semantic: DocumentLimits) -> DocumentReadOptions {
    let options = DocumentReadOptions::new(source, semantic);
    assert_eq!(options.source(), source);
    assert_eq!(options.semantic(), semantic);
    options
}

fn xml_plist(root: &str) -> Vec<u8> {
    format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\
         <!DOCTYPE plist PUBLIC \"-//Apple//DTD PLIST 1.0//EN\" \
         \"http://www.apple.com/DTDs/PropertyList-1.0.dtd\">\
         <plist version=\"1.0\">{root}</plist>"
    )
    .into_bytes()
}

fn native_with_metadata(
    properties: Option<&[u8]>,
    history: Option<&[u8]>,
    identifier: Option<&[u8]>,
) -> TestResult<Vec<u8>> {
    let source = std::fs::read(fixture("numbers/basic.numbers"))?;
    let catalog = Catalog::from_bytes(&source)?;
    let mut entries = catalog
        .iter()
        .filter(|entry| {
            !matches!(
                entry.name(),
                PROPERTIES | BUILD_HISTORY | DOCUMENT_IDENTIFIER
            )
        })
        .map(|entry| (entry.name(), entry.data()))
        .collect::<Vec<_>>();
    entries.extend(
        [
            (PROPERTIES, properties),
            (BUILD_HISTORY, history),
            (DOCUMENT_IDENTIFIER, identifier),
        ]
        .into_iter()
        .filter_map(|(name, data)| data.map(|data| (name, data))),
    );
    Ok(litchi_iwa_archive::package::to_bytes(
        entries,
        ArchiveLimits::default(),
    )?)
}

fn assert_read_limit(
    result: Result<Document, DocumentReadError>,
    kind: DocumentReadLimitKind,
    observed: u64,
    maximum: u64,
) {
    assert!(matches!(
        result,
        Err(DocumentReadError::Limit {
            kind: actual_kind,
            observed: actual_observed,
            maximum: actual_maximum,
        }) if actual_kind == kind && actual_observed == observed && actual_maximum == maximum
    ));
}

fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..Default::default()
    }
}

fn archive_object(identifier: u64, type_: u32, data: Vec<u8>) -> TestResult<ArchiveObject> {
    Ok(ArchiveObject::new(
        identifier,
        vec![RawMessage { type_, data }],
    )?)
}

fn synthetic_formula_source() -> TestResult<Vec<u8>> {
    let formula = tsce::FormulaArchive {
        ast_node_array: tsce::AstNodeArrayArchive {
            ast_node: vec![AstNodeArchive {
                ast_node_type: AstNodeType::NumberNode as i32,
                ast_number_node_number: Some(203.0),
                ..Default::default()
            }],
        },
        ..Default::default()
    };
    let mut lists = ArchiveObject::new(
        5,
        [
            tst::TableDataList {
                list_type: tst::table_data_list::ListType::String as i32,
                next_list_id: 1,
                ..Default::default()
            },
            tst::TableDataList {
                list_type: tst::table_data_list::ListType::Formula as i32,
                next_list_id: 2,
                entries: vec![tst::table_data_list::ListEntry {
                    key: 1,
                    refcount: 1,
                    formula: Some(formula),
                    ..Default::default()
                }],
                ..Default::default()
            },
            tst::TableDataList {
                list_type: tst::table_data_list::ListType::RichTextPayload as i32,
                next_list_id: 2,
                entries: vec![tst::table_data_list::ListEntry {
                    key: 1,
                    refcount: 1,
                    rich_text_payload: Some(reference(7)),
                    ..Default::default()
                }],
                ..Default::default()
            },
        ]
        .into_iter()
        .map(|list| RawMessage {
            type_: 6_005,
            data: list.encode_to_vec(),
        })
        .collect(),
    )?;
    lists.archive_info.message_infos[2]
        .object_references
        .push(7);
    let mut formula_cell = BncCell::minimal();
    formula_cell.set_formula_reference(1);
    let mut rich_cell = BncCell::minimal();
    rich_cell.set_rich_text(1);
    let formula_bytes = formula_cell.encode();
    let rich_bytes = rich_cell.encode();
    let rich_offset = u16::try_from(formula_bytes.len())?.to_le_bytes();
    let mut row_storage = formula_bytes;
    row_storage.extend_from_slice(&rich_bytes);
    let table = tst::TableModelArchive {
        table_id: "formula-table-id".to_owned(),
        table_name: "Formula table".to_owned(),
        number_of_rows: 1,
        number_of_columns: 2,
        base_data_store: tst::DataStore {
            row_headers: tst::HeaderStorage {
                bucket_hash_function: 1,
                ..Default::default()
            },
            column_headers: reference(5),
            tiles: tst::TileStorage {
                tiles: vec![tst::tile_storage::Tile {
                    tileid: 0,
                    tile: reference(6),
                }],
                tile_size: Some(256),
                ..Default::default()
            },
            string_table: reference(5),
            style_table: reference(5),
            formula_table: reference(5),
            rich_text_table: Some(reference(5)),
            format_table_pre_bnc: reference(5),
            next_row_strip_id: 1,
            next_column_strip_id: 1,
            row_tile_tree: tst::TableRbTree::default(),
            column_tile_tree: tst::TableRbTree::default(),
            ..Default::default()
        },
        ..Default::default()
    };
    let tile = tst::Tile {
        numrows: 1,
        row_infos: vec![tst::TileRowInfo {
            tile_row_index: 0,
            cell_count: 2,
            storage_version: Some(5),
            cell_storage_buffer: Some(row_storage),
            cell_offsets: Some(vec![0, 0, rich_offset[0], rich_offset[1]]),
            ..Default::default()
        }],
        ..Default::default()
    };
    let mut rich_payload = archive_object(
        7,
        6_218,
        tst::RichTextPayloadArchive {
            storage: reference(8),
            range: None,
            cellid: tst::CellId {
                packed_data: 1,
                expanded_coord: None,
            },
        }
        .encode_to_vec(),
    )?;
    rich_payload.archive_info.message_infos[0]
        .object_references
        .push(8);
    let objects = vec![
        archive_object(
            1,
            1,
            tn::DocumentArchive {
                sheets: vec![reference(2)],
                ..Default::default()
            }
            .encode_to_vec(),
        )?,
        archive_object(
            2,
            2,
            tn::SheetArchive {
                name: "Formula sheet".to_owned(),
                drawable_infos: vec![reference(3)],
                ..Default::default()
            }
            .encode_to_vec(),
        )?,
        archive_object(
            3,
            6_000,
            tst::TableInfoArchive {
                table_model: reference(4),
                ..Default::default()
            }
            .encode_to_vec(),
        )?,
        archive_object(4, 6_001, table.encode_to_vec())?,
        lists,
        archive_object(6, 6_002, tile.encode_to_vec())?,
        rich_payload,
        archive_object(
            8,
            2_001,
            tswp::StorageArchive {
                kind: Some(tswp::storage_archive::KindType::Cell as i32),
                text: vec!["Rich cell".to_owned()],
                ..Default::default()
            }
            .encode_to_vec(),
        )?,
    ];
    let iwa = SnappyStream::compress(&Archive { objects }.to_bytes()?)?;
    Ok(litchi_iwa_archive::package::to_bytes(
        [("Index/Document.iwa", iwa.as_slice())],
        ArchiveLimits::default(),
    )?)
}

fn assert_native_semantics(document: &Document) -> TestResult {
    document.validate()?;
    assert_eq!(document.sheet_count(), 1);
    let sheet = document
        .sheet(0)?
        .ok_or_else(|| std::io::Error::other("native Numbers fixture has no first sheet"))?;
    assert_eq!(sheet.index(), 0);
    assert_eq!(sheet.name(), "Sheet 1");
    let table = sheet
        .tables()
        .next()
        .ok_or_else(|| std::io::Error::other("native Numbers fixture has no first table"))?;
    assert_eq!(table.name(), "Table 1");
    assert_eq!((table.row_count(), table.column_count()), (22, 7));
    assert_eq!(table.cell_count(), 2);
    assert!(matches!(
        table.get_a1("B2")?,
        Some(Value::Text(value)) if value == "Litchi native Numbers fixture"
    ));
    assert!(matches!(
        table.get_a1("B3")?,
        Some(Value::Number(value)) if value.get().to_bits() == 42.0_f64.to_bits()
    ));
    assert_eq!(
        table
            .iter_cells()
            .filter(|cell| matches!(cell.value(), Value::Formula(_)))
            .count(),
        0
    );
    let text = document.plain_text()?;
    assert_eq!(text, EXPECTED_TEXT);
    assert_eq!(document.text_len(), text.len());
    Ok(())
}

fn assert_native_metadata(
    document: &Document,
    expected_identifier: &str,
    expected_revision: &str,
) -> TestResult {
    let metadata = document
        .metadata()
        .ok_or_else(|| std::io::Error::other("source-backed Numbers metadata is missing"))?;
    assert_eq!(metadata.application.as_deref(), Some("Numbers"));
    assert_eq!(metadata.identifier.as_deref(), Some(expected_identifier));
    assert_eq!(metadata.revision.as_deref(), Some(expected_revision));
    assert_eq!(metadata.version.as_deref(), Some(EXPECTED_VERSION));
    assert_eq!(
        metadata.content_status.as_deref(),
        Some("Numbers Format Version 14.4.1")
    );
    Ok(())
}

fn semantic_document() -> TestResult<Document> {
    let mut table = TableBuilder::new("Types", Dimensions::new(3, 3));
    for (position, value) in [
        (Position::new(0, 0), Value::Text("alpha".to_owned())),
        (Position::new(0, 1), Value::Empty),
        (Position::new(0, 2), Value::Formula("=SUM(A1,2)".to_owned())),
        (Position::new(1, 0), Value::number(42.5)?),
        (Position::new(1, 1), Value::Boolean(true)),
        (Position::new(1, 2), Value::date(12.25)?),
        (Position::new(2, 0), Value::duration(-3.5)?),
        (Position::new(2, 1), Value::Error("bad".to_owned())),
        (Position::new(2, 2), Value::Text(String::new())),
    ] {
        assert!(table.set(position, value).is_ok());
    }
    table.set_column_headers(["hidden-column"])?;
    table.set_row_headers(["hidden-row"])?;
    let table = table.finish()?;
    let mut sheet = SheetBuilder::new("Summary", 0);
    assert!(sheet.push_table(table).is_ok());
    Ok(Document::from_sheets(vec![sheet.finish()])?)
}

#[test]
fn archive_free_document_has_path_byte_directory_and_package_parity() -> TestResult {
    assert_send_sync::<Document>();
    assert_send_sync::<DocumentStats>();
    assert_send_sync::<DocumentReadOptions>();

    let path = fixture("numbers/basic.numbers");
    let package = Package::open(&path)?;
    let zipped = Document::open(&path)?;
    let directory = Document::open(fixture("directory/numbers/basic.numbers"))?;
    let bytes = std::fs::read(path)?;
    let borrowed = Document::from_bytes(&bytes)?;
    let shared_source: Arc<[u8]> = Arc::from(bytes);
    let weak_source = Arc::downgrade(&shared_source);
    let shared = Document::from_shared_bytes(shared_source)?;
    assert!(weak_source.upgrade().is_none());

    for document in [&zipped, &directory, &borrowed, &shared] {
        assert_native_semantics(document)?;
        assert_eq!(
            document.stats(),
            Some(DocumentStats {
                source_record_count: 622,
                sheet_count: 1,
                table_count: 1,
            })
        );
        assert_eq!(document.sheets(), zipped.sheets());
    }
    for document in [&zipped, &borrowed, &shared] {
        assert_native_metadata(document, ZIP_IDENTIFIER, ZIP_REVISION)?;
    }
    assert_native_metadata(&directory, DIRECTORY_IDENTIFIER, DIRECTORY_REVISION)?;

    let snapshot = zipped.snapshot();
    assert!(Arc::ptr_eq(
        &zipped.shared_sheets(),
        &snapshot.shared_sheets()
    ));
    assert_eq!(
        snapshot.metadata().map(|value| value.identifier.as_deref()),
        zipped.metadata().map(|value| value.identifier.as_deref())
    );

    let semantic = package.document();
    assert_eq!(semantic.sheets(), zipped.sheets());
    assert_eq!(semantic.plain_text()?, EXPECTED_TEXT);
    assert_eq!(semantic.text_len(), EXPECTED_TEXT.len());
    assert_eq!(semantic.stats(), None);
    assert!(semantic.metadata().is_none());
    assert!(Arc::ptr_eq(
        &semantic.shared_sheets(),
        &semantic.snapshot().shared_sheets()
    ));
    Ok(())
}

#[test]
fn package_remains_exact_file_only_authority() {
    assert!(matches!(
        Package::open(fixture("directory/numbers/basic.numbers")),
        Err(PackageError::InvalidFormat(_))
    ));
}

#[test]
fn foreign_iwork_families_are_typed_before_publication() {
    for path in [
        fixture("pages/basic.pages"),
        fixture("keynote/basic.key"),
        fixture("directory/pages/basic.pages"),
        fixture("directory/keynote/basic.key"),
    ] {
        assert!(matches!(
            Document::open(path),
            Err(DocumentReadError::NotNumbers)
        ));
    }
}

#[test]
fn frozen_directory_survives_source_deletion_and_source_limits_are_inclusive() -> TestResult {
    let temp = tempfile::tempdir()?;
    let source = temp.path().join("detached.numbers");
    copy_directory_fixture(&source)?;
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
    let document = Document::open_with_options(
        &source,
        options(source_limits(exact_input)?, DocumentLimits::default()),
    )?;
    let error = Document::open_with_options(
        &source,
        options(source_limits(exact_input - 1)?, DocumentLimits::default()),
    )
    .expect_err("source maximum-minus-one must refuse");
    assert!(matches!(
        error,
        DocumentReadError::Limit {
            kind: DocumentReadLimitKind::InputBytes,
            observed,
            maximum,
        } if observed == exact_input && maximum == exact_input - 1
    ));

    std::fs::remove_dir_all(&source)?;
    assert_native_semantics(&document)?;
    assert_native_metadata(&document, DIRECTORY_IDENTIFIER, DIRECTORY_REVISION)?;
    assert_eq!(
        document.stats().map(|stats| stats.source_record_count),
        Some(622)
    );
    Ok(())
}

#[test]
fn zip_input_limit_is_exact_across_path_borrowed_and_shared_sources() -> TestResult {
    let path = fixture("numbers/basic.numbers");
    let bytes = std::fs::read(&path)?;
    let exact = u64::try_from(bytes.len())?;
    let options_for = |maximum| -> TestResult<DocumentReadOptions> {
        Ok(options(source_limits(maximum)?, DocumentLimits::default()))
    };

    Document::open_with_options(&path, options_for(exact)?)?.validate()?;
    Document::from_bytes_with_options(&bytes, options_for(exact)?)?.validate()?;
    Document::from_shared_bytes_with_options(Arc::from(bytes.clone()), options_for(exact)?)?
        .validate()?;

    for result in [
        Document::open_with_options(&path, options_for(exact - 1)?),
        Document::from_bytes_with_options(&bytes, options_for(exact - 1)?),
        Document::from_shared_bytes_with_options(Arc::from(bytes.clone()), options_for(exact - 1)?),
    ] {
        assert_read_limit(result, DocumentReadLimitKind::InputBytes, exact, exact - 1);
    }
    Ok(())
}

#[test]
fn semantic_text_is_deterministic_and_excludes_empty_values_and_headers() -> TestResult {
    let document = semantic_document()?;
    let expected = "Summary\nTypes\nalpha\n=SUM(A1,2)\n42.5\ntrue\n12.25\n-3.5\nERROR: bad";
    let text = document.plain_text()?;
    assert_eq!(text, expected);
    assert_eq!(document.text_len(), expected.len());
    assert!(!text.contains("hidden-column"));
    assert!(!text.contains("hidden-row"));
    assert_eq!(document.stats(), None);
    assert!(document.metadata().is_none());
    Ok(())
}

#[test]
fn rooted_formula_source_is_materialized_once_in_semantic_text() -> TestResult {
    let bytes = synthetic_formula_source()?;
    Package::from_bytes(&bytes)?;
    for document in [
        Document::from_bytes(&bytes)?,
        Document::from_shared_bytes(Arc::from(bytes.clone()))?,
    ] {
        document.validate()?;
        let table = document
            .sheet(0)?
            .and_then(|sheet| sheet.tables().next())
            .ok_or_else(|| std::io::Error::other("synthetic formula table is missing"))?;
        assert!(matches!(table.get_a1("A1")?, Some(Value::Formula(value)) if value == "=203"));
        assert!(matches!(table.get_a1("B1")?, Some(Value::Text(value)) if value == "Rich cell"));
        let expected = "Formula sheet\nFormula table\n=203\nRich cell";
        assert_eq!(document.plain_text()?, expected);
        assert_eq!(document.text_len(), expected.len());
        assert_eq!(document.stats().map(|stats| stats.sheet_count), Some(1));
        assert!(document.metadata().is_some());
    }
    Ok(())
}

#[test]
fn missing_rooted_sheet_refuses_without_partial_publication() -> TestResult {
    let root = archive_object(
        1,
        1,
        tn::DocumentArchive {
            sheets: vec![reference(999)],
            ..Default::default()
        }
        .encode_to_vec(),
    )?;
    let iwa = SnappyStream::compress(
        &Archive {
            objects: vec![root],
        }
        .to_bytes()?,
    )?;
    let bytes = litchi_iwa_archive::package::to_bytes(
        [("Index/Document.iwa", iwa.as_slice())],
        ArchiveLimits::default(),
    )?;
    assert!(matches!(
        Document::from_bytes(&bytes),
        Err(DocumentReadError::InvalidFormat)
    ));
    Ok(())
}

#[cfg(windows)]
#[test]
fn focused_path_ingress_fails_closed_on_windows() {
    assert!(matches!(
        Document::open(fixture("numbers/basic.numbers")),
        Err(DocumentReadError::InvalidSource)
    ));
}

#[test]
fn semantic_limits_are_exact_and_zero_is_never_widened() -> TestResult {
    let bytes = std::fs::read(fixture("numbers/basic.numbers"))?;
    let read = |limits| {
        Document::from_bytes_with_options(&bytes, options(DocumentSourceLimits::default(), limits))
    };

    let exact = DocumentLimits::new(1, 1, 2, EXPECTED_TEXT.len())?;
    assert_eq!(read(exact)?.plain_text()?, EXPECTED_TEXT);
    for (limits, kind, observed) in [
        (
            DocumentLimits::new(0, 1, 2, EXPECTED_TEXT.len())?,
            DocumentReadLimitKind::Sheets,
            1,
        ),
        (
            DocumentLimits::new(1, 0, 2, EXPECTED_TEXT.len())?,
            DocumentReadLimitKind::Tables,
            1,
        ),
        (
            DocumentLimits::new(1, 1, 0, EXPECTED_TEXT.len())?,
            DocumentReadLimitKind::Cells,
            2,
        ),
        (
            DocumentLimits::new(1, 1, 2, 0)?,
            DocumentReadLimitKind::TextBytes,
            "Sheet 1".len(),
        ),
    ] {
        let error = read(limits).expect_err("zero semantic ceiling must refuse the native source");
        assert!(matches!(
            error,
            DocumentReadError::Limit {
                kind: actual_kind,
                observed: actual_observed,
                maximum: 0,
            } if actual_kind == kind && actual_observed == observed as u64
        ));
    }

    let text_error = read(DocumentLimits::new(1, 1, 2, EXPECTED_TEXT.len() - 1)?)
        .expect_err("rendered text maximum-minus-one must refuse");
    assert!(matches!(
        text_error,
        DocumentReadError::Limit {
            kind: DocumentReadLimitKind::TextBytes,
            observed,
            maximum,
        } if observed == EXPECTED_TEXT.len() as u64
            && maximum == (EXPECTED_TEXT.len() - 1) as u64
    ));
    Ok(())
}

#[test]
fn checked_limit_constructors_reject_only_invalid_profiles() {
    assert!(DocumentLimits::new(0, 0, 0, 0).is_ok());
    assert!(DocumentLimits::new(MAX_SHEETS + 1, 0, 0, 0).is_err());
    assert!(DocumentLimits::new(0, MAX_TABLES + 1, 0, 0).is_err());
    assert!(DocumentLimits::new(0, 0, MAX_MATERIALIZED_CELLS + 1, 0).is_err());
    assert!(DocumentLimits::new(0, 0, 0, DEFAULT_MAX_TEXT_BYTES + 1).is_err());

    let defaults = DocumentSourceLimits::default();
    for values in [
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
    ] {
        assert!(
            DocumentSourceLimits::new(values.0, values.1, values.2, values.3, values.4).is_err()
        );
    }
}

#[test]
fn canonical_metadata_authorities_ignore_malformed_near_names() -> TestResult {
    let source = std::fs::read(fixture("numbers/basic.numbers"))?;
    let catalog = Catalog::from_bytes(&source)?;
    let malformed = b"not a plist";
    let mut entries = vec![
        ("A/Properties.plist", malformed.as_slice()),
        ("Properties.plist", malformed.as_slice()),
        ("Metadata/Properties.plist.bak", malformed.as_slice()),
        ("A/BuildVersionHistory.plist", malformed.as_slice()),
        ("BuildVersionHistory.plist", malformed.as_slice()),
        ("Metadata/DocumentIdentifier.bak", b"near-id".as_slice()),
    ];
    entries.extend(catalog.iter().map(|entry| (entry.name(), entry.data())));
    let candidate = litchi_iwa_archive::package::to_bytes(entries, ArchiveLimits::default())?;
    let document = Document::from_bytes(&candidate)?;
    assert_native_metadata(&document, ZIP_IDENTIFIER, ZIP_REVISION)?;
    assert_native_semantics(&document)?;
    Ok(())
}

#[test]
fn metadata_event_depth_and_scalar_limits_are_exact() -> TestResult {
    let pairs = |count: usize| {
        let body = "<key>opaque</key><string>x</string>".repeat(count);
        xml_plist(&format!("<dict>{body}</dict>"))
    };
    Document::from_bytes(&native_with_metadata(Some(&pairs(511)), None, None)?)?.validate()?;
    assert_read_limit(
        Document::from_bytes(&native_with_metadata(Some(&pairs(512)), None, None)?),
        DocumentReadLimitKind::PayloadFields,
        1_025,
        1_024,
    );

    let nested = |count: usize| {
        xml_plist(&format!(
            "<dict><key>opaque</key>{}{}</dict>",
            "<array>".repeat(count),
            "</array>".repeat(count)
        ))
    };
    Document::from_bytes(&native_with_metadata(Some(&nested(15)), None, None)?)?.validate()?;
    assert_read_limit(
        Document::from_bytes(&native_with_metadata(Some(&nested(16)), None, None)?),
        DocumentReadLimitKind::PayloadNesting,
        17,
        16,
    );

    let scalar = |length: usize| {
        xml_plist(&format!(
            "<dict><key>Title</key><string>{}</string></dict>",
            "x".repeat(length)
        ))
    };
    let selected =
        Document::from_bytes(&native_with_metadata(Some(&scalar(16 * 1024)), None, None)?)?;
    assert_eq!(
        selected.metadata().and_then(|value| value.title.as_deref()),
        Some("x".repeat(16 * 1024).as_str())
    );
    assert_read_limit(
        Document::from_bytes(&native_with_metadata(
            Some(&scalar(16 * 1024 + 1)),
            None,
            None,
        )?),
        DocumentReadLimitKind::TextBytes,
        16 * 1024 + 1,
        16 * 1024,
    );
    Ok(())
}

#[test]
fn metadata_history_and_retained_limits_are_exact() -> TestResult {
    let history = |count: usize| {
        xml_plist(&format!(
            "<array>{}</array>",
            "<string>v</string>".repeat(count)
        ))
    };
    let exact_history = history(128);
    let document = Document::from_bytes(&native_with_metadata(None, Some(&exact_history), None)?)?;
    assert_eq!(
        document
            .metadata()
            .and_then(|value| value.version.as_deref()),
        Some("v")
    );
    assert_read_limit(
        Document::from_bytes(&native_with_metadata(None, Some(&history(129)), None)?),
        DocumentReadLimitKind::PayloadFields,
        129,
        128,
    );

    let value = "x".repeat(16 * 1024);
    let properties = xml_plist(&format!(
        "<dict><key>Title</key><string>{value}</string>\
         <key>Author</key><string>{value}</string>\
         <key>Keywords</key><string>{value}</string></dict>"
    ));
    let exact = Document::from_bytes(&native_with_metadata(
        Some(&properties),
        None,
        Some(value.as_bytes()),
    )?)?;
    assert_eq!(
        exact
            .metadata()
            .and_then(|metadata| metadata.title.as_deref()),
        Some(value.as_str())
    );
    let one_more = xml_plist("<array><string>v</string></array>");
    assert_read_limit(
        Document::from_bytes(&native_with_metadata(
            Some(&properties),
            Some(&one_more),
            Some(value.as_bytes()),
        )?),
        DocumentReadLimitKind::TextBytes,
        64 * 1024 + 1,
        64 * 1024,
    );
    Ok(())
}

#[test]
fn metadata_duplicate_keys_refuse_and_history_version_takes_precedence() -> TestResult {
    for root in [
        "<dict><key>Title</key><string>one</string><key>Title</key><string>two</string></dict>",
        "<array><dict><key>Version</key><string>one</string><key>Version</key><string>two</string></dict></array>",
        "<array><dict><key>Build</key><string>one</string><key>Build</key><string>two</string></dict></array>",
    ] {
        let data = xml_plist(root);
        let (properties, history) = if root.starts_with("<dict>") {
            (Some(data.as_slice()), None)
        } else {
            (None, Some(data.as_slice()))
        };
        assert!(matches!(
            Document::from_bytes(&native_with_metadata(properties, history, None)?),
            Err(DocumentReadError::InvalidFormat)
        ));
    }

    let properties = xml_plist("<dict><key>buildVersion</key><string>property</string></dict>");
    let history = xml_plist(
        "<array><dict><key>Build</key><string>build</string>\
         <key>Version</key><string>version</string></dict></array>",
    );
    let document = Document::from_bytes(&native_with_metadata(
        Some(&properties),
        Some(&history),
        None,
    )?)?;
    assert_eq!(
        document
            .metadata()
            .and_then(|value| value.version.as_deref()),
        Some("version")
    );
    Ok(())
}

#[test]
fn canonical_metadata_hard_limit_is_exact_before_publication() -> TestResult {
    const MAX_SIDECAR_BYTES: usize = 64 * 1024;
    let source = std::fs::read(fixture("numbers/basic.numbers"))?;
    let catalog = Catalog::from_bytes(&source)?;
    let prefix = b"<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n\
        <!DOCTYPE plist PUBLIC \"-//Apple//DTD PLIST 1.0//EN\" \
        \"http://www.apple.com/DTDs/PropertyList-1.0.dtd\">\n\
        <plist version=\"1.0\"><dict>\
        <key>fileFormatVersion</key><string>14.4.1</string>";
    let suffix = b"</dict></plist>";
    let mut exact = prefix.to_vec();
    exact.resize(MAX_SIDECAR_BYTES - suffix.len(), b' ');
    exact.extend_from_slice(suffix);
    assert_eq!(exact.len(), MAX_SIDECAR_BYTES);

    for (length, should_pass) in [(MAX_SIDECAR_BYTES, true), (MAX_SIDECAR_BYTES + 1, false)] {
        let mut properties = exact.clone();
        properties.resize(length, b' ');
        let mut entries = vec![("Metadata/Properties.plist", properties.as_slice())];
        entries.extend(catalog.iter().filter_map(|entry| {
            (entry.name() != "Metadata/Properties.plist").then_some((entry.name(), entry.data()))
        }));
        let candidate = litchi_iwa_archive::package::to_bytes(entries, ArchiveLimits::default())?;
        let temp = tempfile::tempdir()?;
        let path = temp.path().join("metadata-limit.numbers");
        std::fs::write(&path, &candidate)?;
        if should_pass {
            Document::open(&path)?.validate()?;
            Document::from_bytes(&candidate)?.validate()?;
            Document::from_shared_bytes(Arc::from(candidate))?.validate()?;
        } else {
            assert_read_limit(
                Document::open(&path),
                DocumentReadLimitKind::PayloadBytes,
                length as u64,
                MAX_SIDECAR_BYTES as u64,
            );
            assert_read_limit(
                Document::from_bytes(&candidate),
                DocumentReadLimitKind::PayloadBytes,
                length as u64,
                MAX_SIDECAR_BYTES as u64,
            );
            assert_read_limit(
                Document::from_shared_bytes(Arc::from(candidate)),
                DocumentReadLimitKind::PayloadBytes,
                length as u64,
                MAX_SIDECAR_BYTES as u64,
            );
        }
    }
    Ok(())
}

#[test]
fn public_read_errors_redact_paths_members_and_control_characters() -> TestResult {
    let temp = tempfile::tempdir()?;
    let path_sentinel = "private-numbers-path-998244353";
    let path_error = Document::open(temp.path().join(path_sentinel))
        .expect_err("missing sentinel path must fail");
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

    let control_sentinel = "private-control\r\n\u{1b}[31m-112233445";
    let control_error = Document::from_bytes(control_sentinel.as_bytes())
        .expect_err("unrecognized control-character input must fail");
    assert_content_free(
        &control_error,
        &[control_sentinel, "private-control", "112233445"],
    );
    Ok(())
}
