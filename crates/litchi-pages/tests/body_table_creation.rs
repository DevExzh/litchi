//! Body-table creation (record 0329): a new, empty table cloned from the
//! structure of the document's first body table and anchored in the body
//! story, read back through the catalog and cell reader, reopened and
//! inverted exactly; documents without a body table refuse, typed.

use std::path::{Path, PathBuf};

use litchi_core::Position;
use litchi_pages::table::creation::{Commit, Error, NewTable};
use litchi_pages::table::read::Value as ReadValue;
use litchi_pages::table::values::Value;
use litchi_pages::{BodyTableSelector, Package};

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

fn pages_dir() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data/iwork/pages")
}

fn corpus() -> Vec<PathBuf> {
    let mut files: Vec<PathBuf> = std::fs::read_dir(pages_dir())
        .expect("corpus directory")
        .map(|entry| entry.expect("entry").path())
        .filter(|path| path.is_file() && path.extension().is_some_and(|ext| ext == "pages"))
        .collect();
    files.sort();
    files
}

fn bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut output = Vec::new();
    package.write_to(&mut output)?;
    Ok(output)
}

/// The UTF-16 length of the first section's text, a position inside the
/// body story.
fn first_section_end(package: &Package) -> usize {
    package
        .sections()
        .first()
        .and_then(|section| section.body_text())
        .map_or(0, |text| text.encode_utf16().count())
}

/// Stage and commit one insertion.
fn insert(package: &Package, anchor: Position, table: NewTable) -> Result<Commit, Error> {
    package
        .edit_body_table_creation()
        .insert(anchor, table)?
        .commit()
}

#[test]
fn the_edit_stages_exactly_one_insertion() -> TestResult {
    let package = Package::open(pages_dir().join("body-table-visible.pages"))?;
    assert!(matches!(
        package.edit_body_table_creation().commit(),
        Err(Error::NoStagedOperation)
    ));
    assert!(matches!(
        package
            .edit_body_table_creation()
            .insert(Position::new(0), NewTable::new(2, 2))?
            .insert(Position::new(0), NewTable::new(2, 2)),
        Err(Error::OperationAlreadyStaged)
    ));
    let commit = insert(&package, Position::new(0), NewTable::new(2, 2))?;
    assert!(commit.diagnostics().changed());
    assert!(!commit.patch().is_noop());
    Ok(())
}

#[test]
fn tables_are_inserted_read_back_reopened_and_inverted() -> TestResult {
    let package = Package::open(pages_dir().join("body-table-visible.pages"))?;
    let source = bytes(&package)?;
    let before = package.body_tables()?.len();
    let commit = insert(
        &package,
        Position::new(0),
        NewTable::new(3, 4)
            .with_name("Budget")
            .with_cell(0, 0, Value::Text("Item".into()))
            .with_cell(2, 3, Value::Number(4.25)),
    )?;
    let edited = commit.package();
    let catalog = edited.body_tables()?;
    assert_eq!(catalog.len(), before + 1);
    let table = commit.table().get();
    let snapshot = catalog.get(table).expect("new table");
    assert_eq!(snapshot.name(), "Budget");
    assert_eq!((snapshot.rows(), snapshot.columns()), (3, 4));
    let reopened = Package::from_bytes(&bytes(edited)?)?;
    let read = reopened.body_table_cells(BodyTableSelector::index(table))?;
    assert_eq!(
        read.get(litchi_iwa_common::table::coordinate::At::new(0, 0)),
        Some(&ReadValue::Text("Item".into()))
    );
    // The template still reads as before.
    let original_index = usize::from(table == 0);
    let original = package.body_table_cells(BodyTableSelector::index(0))?;
    assert_eq!(
        reopened.body_table_cells(BodyTableSelector::index(original_index))?,
        original
    );
    let restored = edited.apply_body_table_creation(&commit.patch().inverse())?;
    assert_eq!(bytes(restored.package())?, source);
    Ok(())
}

#[test]
fn every_corpus_document_inserts_a_table_or_refuses_typed() -> TestResult {
    let (mut created, mut refused) = (0_usize, 0_usize);
    for path in corpus() {
        let package = Package::open(&path)?;
        let tables = match package.body_tables() {
            Ok(catalog) => catalog.len(),
            Err(_) => {
                // The catalog reader refuses the document, and so does
                // creation, typed.
                assert!(
                    matches!(
                        insert(&package, Position::new(0), NewTable::new(2, 2)),
                        Err(Error::Unsupported)
                    ),
                    "{}",
                    path.display()
                );
                refused += 1;
                continue;
            },
        };
        for anchor in [0, first_section_end(&package)] {
            let result = insert(
                &package,
                Position::new(anchor),
                NewTable::new(2, 2).with_cell(1, 0, Value::Text("x".into())),
            );
            let commit = match result {
                Ok(commit) => commit,
                Err(Error::NoTemplate) => {
                    assert_eq!(tables, 0, "{}", path.display());
                    refused += 1;
                    continue;
                },
                Err(error) => panic!("{}: {error:?}", path.display()),
            };
            let reopened = Package::from_bytes(&bytes(commit.package())?)?;
            assert_eq!(reopened.body_tables()?.len(), tables + 1);
            let read = reopened.body_table_cells(BodyTableSelector::index(commit.table().get()))?;
            assert_eq!(
                read.get(litchi_iwa_common::table::coordinate::At::new(1, 0)),
                Some(&ReadValue::Text("x".into())),
                "{}",
                path.display()
            );
            let restored = commit
                .package()
                .apply_body_table_creation(&commit.patch().inverse())?;
            assert_eq!(bytes(restored.package())?, bytes(&package)?);
            created += 1;
        }
    }
    eprintln!("created {created}, refused {refused}");
    assert!(created >= 10);
    Ok(())
}

#[test]
fn invalid_requests_refuse_typed() -> TestResult {
    let package = Package::open(pages_dir().join("body-table-visible.pages"))?;
    assert!(matches!(
        insert(&package, Position::new(0), NewTable::new(0, 2)),
        Err(Error::InvalidSize {
            rows: 0,
            columns: 2
        })
    ));
    assert!(matches!(
        insert(
            &package,
            Position::new(0),
            NewTable::new(2, 2).with_cell(2, 0, Value::Boolean(true))
        ),
        Err(Error::InvalidCell)
    ));
    assert!(matches!(
        insert(&package, Position::new(1_000_000), NewTable::new(2, 2)),
        Err(Error::PositionOutOfBounds)
    ));
    let plain = Package::open(pages_dir().join("basic.pages"))?;
    assert!(matches!(
        insert(&plain, Position::new(0), NewTable::new(2, 2)),
        Err(Error::NoTemplate)
    ));
    Ok(())
}
