//! Body-table creation (record 0329): a new, empty table cloned from the
//! structure of the document's first body table and anchored in the body
//! story, read back through the catalog and cell reader, reopened and
//! inverted exactly; documents without a body table refuse, typed. A
//! duplication (record 0371) copies a populated table under the next free
//! `"<name> copy"` name.

use std::path::{Path, PathBuf};

use litchi_core::Position;
use litchi_pages::table::creation::{Commit, Error, NewTable, Refusal};
use litchi_pages::table::read::{TableRead, Value as ReadValue};
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

/// Stage and commit one duplication.
fn duplicate<'a>(
    package: &Package,
    table: impl Into<BodyTableSelector<'a>>,
    anchor: Position,
) -> Result<Commit, Error> {
    package
        .edit_body_table_creation()
        .duplicate(table, anchor)?
        .commit()
}

#[test]
fn the_edit_stages_exactly_one_insertion() -> TestResult {
    let package = Package::open(pages_dir().join("body-table-visible.pages"))?;
    assert!(matches!(
        package.edit_body_table_creation().commit(),
        Err(Error::Refused(Refusal::NoStagedOperation))
    ));
    assert!(matches!(
        package
            .edit_body_table_creation()
            .insert(Position::new(0), NewTable::new(2, 2))?
            .insert(Position::new(0), NewTable::new(2, 2)),
        Err(Error::Refused(Refusal::OperationAlreadyStaged))
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
                        Err(Error::Refused(Refusal::Unsupported))
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
                Err(Error::Refused(Refusal::NoTemplate)) => {
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
        Err(Error::Refused(Refusal::InvalidSize {
            rows: 0,
            columns: 2
        }))
    ));
    assert!(matches!(
        insert(
            &package,
            Position::new(0),
            NewTable::new(2, 2).with_cell(2, 0, Value::Boolean(true))
        ),
        Err(Error::Refused(Refusal::InvalidCell))
    ));
    assert!(matches!(
        insert(&package, Position::new(1_000_000), NewTable::new(2, 2)),
        Err(Error::Refused(Refusal::PositionOutOfBounds))
    ));
    let plain = Package::open(pages_dir().join("basic.pages"))?;
    assert!(matches!(
        insert(&plain, Position::new(0), NewTable::new(2, 2)),
        Err(Error::Refused(Refusal::NoTemplate))
    ));
    Ok(())
}

/// A table read apart from its name: dimensions, cells and comments.
fn content(read: &TableRead) -> impl PartialEq + std::fmt::Debug + '_ {
    (
        read.table().dimensions(),
        read.table().iter_cells().collect::<Vec<_>>(),
        read.comments().collect::<Vec<_>>(),
    )
}

#[test]
fn tables_are_duplicated_with_their_cells_named_and_inverted() -> TestResult {
    let package = Package::open(pages_dir().join("body-table-cells-native.pages"))?;
    let source = bytes(&package)?;
    let catalog = package.body_tables()?;
    let original = catalog.get(0).expect("a body table");
    let commit = duplicate(&package, BodyTableSelector::index(0), Position::new(0))?;
    assert!(commit.diagnostics().changed());
    assert!(!commit.patch().is_noop());
    let edited = commit.package();
    let after = edited.body_tables()?;
    assert_eq!(after.len(), catalog.len() + 1);
    let copy = after.get(commit.table().get()).expect("the copy");
    assert_eq!(copy.name(), format!("{} copy", original.name()));
    assert_eq!(
        (copy.rows(), copy.columns()),
        (original.rows(), original.columns())
    );
    let reopened = Package::from_bytes(&bytes(edited)?)?;
    let (copied, source_cells) = (
        reopened.body_table_cells(BodyTableSelector::index(commit.table().get()))?,
        package.body_table_cells(BodyTableSelector::index(0))?,
    );
    assert_eq!(copied.table().name(), copy.name());
    assert_eq!(content(&copied), content(&source_cells));
    assert!(copied.table().non_empty_cell_count() > 0);
    // The next copy of the same source takes the next free name.
    let again = duplicate(edited, original.name(), Position::new(0))?;
    assert_eq!(
        again
            .package()
            .body_tables()?
            .get(again.table().get())
            .map(|table| table.name().to_owned()),
        Some(format!("{} copy 2", original.name()))
    );
    let restored = edited.apply_body_table_creation(&commit.patch().inverse())?;
    assert_eq!(bytes(restored.package())?, source);
    Ok(())
}

#[test]
fn duplication_refuses_typed() -> TestResult {
    let package = Package::open(pages_dir().join("body-table-visible.pages"))?;
    let tables = package.body_tables()?.len();
    assert!(matches!(
        duplicate(&package, BodyTableSelector::index(tables), Position::new(0)),
        Err(Error::Refused(Refusal::TableNotFound))
    ));
    assert!(matches!(
        duplicate(&package, "no such table", Position::new(0)),
        Err(Error::Refused(Refusal::TableNotFound))
    ));
    assert!(matches!(
        duplicate(
            &package,
            BodyTableSelector::index(0),
            Position::new(1_000_000)
        ),
        Err(Error::Refused(Refusal::PositionOutOfBounds))
    ));
    assert!(matches!(
        package
            .edit_body_table_creation()
            .duplicate(BodyTableSelector::index(0), Position::new(0))?
            .insert(Position::new(0), NewTable::new(2, 2)),
        Err(Error::Refused(Refusal::OperationAlreadyStaged))
    ));
    assert!(matches!(
        package
            .edit_body_table_creation()
            .insert(Position::new(0), NewTable::new(2, 2))?
            .duplicate(BodyTableSelector::index(0), Position::new(0)),
        Err(Error::Refused(Refusal::OperationAlreadyStaged))
    ));
    let plain = Package::open(pages_dir().join("basic.pages"))?;
    assert!(matches!(
        duplicate(&plain, BodyTableSelector::index(0), Position::new(0)),
        Err(Error::Refused(Refusal::TableNotFound))
    ));
    Ok(())
}

/// Every corpus body table is copied, or refused because its formulas carry
/// dependency state the shared engine does not copy (volatile, spanning,
/// whole-owner, error or spill), as the host refuses it.
#[test]
fn every_corpus_body_table_duplicates_or_refuses_typed() -> TestResult {
    let (mut duplicated, mut refused) = (0_usize, 0_usize);
    for path in corpus() {
        let package = Package::open(&path)?;
        let Ok(catalog) = package.body_tables() else {
            continue;
        };
        for index in 0..catalog.len() {
            let commit =
                match duplicate(&package, BodyTableSelector::index(index), Position::new(0)) {
                    Ok(commit) => commit,
                    Err(Error::Refused(Refusal::UnsupportedDependency)) => {
                        refused += 1;
                        continue;
                    },
                    Err(error) => panic!("{} table {index}: {error:?}", path.display()),
                };
            let reopened = Package::from_bytes(&bytes(commit.package())?)?;
            assert_eq!(
                reopened.body_tables()?.len(),
                catalog.len() + 1,
                "{}",
                path.display()
            );
            let (copied, source_cells) = (
                reopened.body_table_cells(BodyTableSelector::index(commit.table().get()))?,
                package.body_table_cells(BodyTableSelector::index(index))?,
            );
            assert_eq!(
                content(&copied),
                content(&source_cells),
                "{} table {index}",
                path.display()
            );
            let restored = commit
                .package()
                .apply_body_table_creation(&commit.patch().inverse())?;
            assert_eq!(bytes(restored.package())?, bytes(&package)?);
            duplicated += 1;
        }
    }
    eprintln!("duplicated {duplicated}, refused {refused}");
    assert!(duplicated >= 18 && refused > 0);
    Ok(())
}
