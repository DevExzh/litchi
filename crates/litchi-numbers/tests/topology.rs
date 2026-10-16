//! Table row and column insertion and removal (record 0309), on the created
//! workbook and across the Numbers corpus.

use std::path::PathBuf;

use litchi_numbers::cell::Value;
use litchi_numbers::table::cells::Input;
use litchi_numbers::table::topology::{
    ColumnDeletion, ColumnInsertion, Error, RowDeletion, RowInsertion, Unsupported,
};
use litchi_numbers::{CellPosition, Package};

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

fn corpus() -> TestResult<Vec<PathBuf>> {
    let mut paths = Vec::new();
    for directory in [
        "../../test-data/iwork/numbers",
        "../../test-data/iwork/synthetic/numbers",
    ] {
        let root = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join(directory);
        for entry in std::fs::read_dir(root)? {
            let path = entry?.path();
            if path
                .extension()
                .is_some_and(|extension| extension == "numbers")
                && path.is_file()
            {
                paths.push(path);
            }
        }
    }
    paths.sort();
    Ok(paths)
}

fn bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn cells(package: &Package, sheet: usize, table: usize) -> Vec<(u32, u32, Value)> {
    let table = package.sheets()[sheet].tables().nth(table).expect("table");
    let mut cells: Vec<(u32, u32, Value)> = table
        .iter_cells()
        .map(|cell| {
            (
                cell.position().row(),
                cell.position().column(),
                cell.value().clone(),
            )
        })
        .collect();
    cells.sort_by_key(|cell| (cell.0, cell.1));
    cells
}

fn size(package: &Package, sheet: usize, table: usize) -> (u32, u32) {
    let table = package.sheets()[sheet].tables().nth(table).expect("table");
    (table.row_count(), table.column_count())
}

/// A created workbook with a few values in its first table.
fn filled() -> TestResult<Package> {
    let mut package = Package::new();
    for (row, column, input) in [
        (1, 1, Input::Text("b1".into())),
        (2, 1, Input::number(2.0)?),
        (2, 3, Input::Text("d2".into())),
        (4, 2, Input::Boolean(true)),
    ] {
        package = package
            .edit_table_cells(0, 0)?
            .set(CellPosition::new(row, column), input)
            .map_err(|error| format!("stage ({row}, {column}): {error:?}"))?
            .commit()
            .map_err(|error| format!("commit ({row}, {column}): {error:?}"))?
            .into_package();
    }
    Ok(package)
}

#[test]
fn rows_insert_and_remove_on_a_created_workbook() -> TestResult {
    let package = filled()?;
    let source = bytes(&package)?;
    let (rows, columns) = size(&package, 0, 0);
    let header = package.table_header_settings(0, 0)?.header_row_count();
    let before = cells(&package, 0, 0);

    // A body row before the row holding (2, 1) and (2, 3).
    let body_index = 2 - header;
    let commit = package
        .edit_table_topology(0, 0)?
        .insert_row(RowInsertion::body(body_index))?
        .commit()?;
    assert!(commit.diagnostics().changed());
    let inserted = commit.package();
    assert_eq!(size(inserted, 0, 0), (rows + 1, columns));
    let expected: Vec<(u32, u32, Value)> = before
        .iter()
        .map(|(row, column, value)| {
            (
                if *row >= 2 { row + 1 } else { *row },
                *column,
                value.clone(),
            )
        })
        .collect();
    assert_eq!(cells(inserted, 0, 0), expected);
    // The inserted row takes values like any other.
    let written = inserted
        .edit_table_cells(0, 0)?
        .set(CellPosition::new(2, 0), Input::Text("new".into()))?
        .commit()?
        .into_package();
    assert_eq!(
        written.sheets()[0]
            .tables()
            .next()
            .unwrap()
            .get(CellPosition::new(2, 0)),
        Some(&Value::Text("new".into()))
    );
    // The inverse restores the source byte for byte; the patch applies only
    // to its source.
    let restored = inserted.apply_table_topology(&commit.patch().inverse())?;
    assert_eq!(bytes(restored.package())?, source);
    assert!(matches!(
        inserted.apply_table_topology(commit.patch()),
        Err(Error::PatchConflict)
    ));

    // Removing the inserted row gives the cells back.
    let removed = inserted
        .edit_table_topology(0, 0)?
        .remove_row(RowDeletion::body(body_index))?
        .commit()?
        .into_package();
    assert_eq!(size(&removed, 0, 0), (rows, columns));
    assert_eq!(cells(&removed, 0, 0), before);

    // Removing a row with values drops them and shifts the rest up.
    let dropped = package
        .edit_table_topology(0, 0)?
        .remove_row(RowDeletion::body(2 - header))?
        .commit()?
        .into_package();
    assert_eq!(size(&dropped, 0, 0), (rows - 1, columns));
    let expected: Vec<(u32, u32, Value)> = before
        .iter()
        .filter(|(row, _, _)| *row != 2)
        .map(|(row, column, value)| {
            (
                if *row > 2 { row - 1 } else { *row },
                *column,
                value.clone(),
            )
        })
        .collect();
    assert_eq!(cells(&dropped, 0, 0), expected);
    let reopened = Package::from_bytes(&bytes(&dropped)?)?;
    assert_eq!(cells(&reopened, 0, 0), expected);
    Ok(())
}

#[test]
fn columns_insert_and_remove_on_a_created_workbook() -> TestResult {
    let package = filled()?;
    let (rows, columns) = size(&package, 0, 0);
    let header = package.table_header_settings(0, 0)?.header_column_count();
    let before = cells(&package, 0, 0);
    let commit = package
        .edit_table_topology(0, 0)?
        .insert_column(ColumnInsertion::body(2 - header))?
        .insert_column(ColumnInsertion::body(columns as usize + 1 - header))?
        .commit()?;
    let inserted = commit.package();
    assert_eq!(size(inserted, 0, 0), (rows, columns + 2));
    let expected: Vec<(u32, u32, Value)> = before
        .iter()
        .map(|(row, column, value)| {
            (
                *row,
                if *column >= 2 { column + 1 } else { *column },
                value.clone(),
            )
        })
        .collect();
    assert_eq!(cells(inserted, 0, 0), expected);
    let restored = inserted.apply_table_topology(&commit.patch().inverse())?;
    assert_eq!(bytes(restored.package())?, bytes(&package)?);

    let dropped = package
        .edit_table_topology(0, 0)?
        .remove_column(ColumnDeletion::body(1 - header))?
        .commit()?
        .into_package();
    assert_eq!(size(&dropped, 0, 0), (rows, columns - 1));
    let expected: Vec<(u32, u32, Value)> = before
        .iter()
        .filter(|(_, column, _)| *column != 1)
        .map(|(row, column, value)| {
            (
                *row,
                if *column > 1 { column - 1 } else { *column },
                value.clone(),
            )
        })
        .collect();
    assert_eq!(cells(&dropped, 0, 0), expected);
    let reopened = Package::from_bytes(&bytes(&dropped)?)?;
    assert_eq!(cells(&reopened, 0, 0), expected);
    Ok(())
}

#[test]
fn header_sections_grow_and_shrink() -> TestResult {
    let package = filled()?;
    let settings = package.table_header_settings(0, 0)?;
    let (rows, _) = size(&package, 0, 0);
    let grown = package
        .edit_table_topology(0, 0)?
        .insert_row(RowInsertion::header(0))?
        .insert_row(RowInsertion::footer(0))?
        .commit()?
        .into_package();
    let after = grown.table_header_settings(0, 0)?;
    assert_eq!(after.header_row_count(), settings.header_row_count() + 1);
    assert_eq!(after.footer_row_count(), settings.footer_row_count() + 1);
    assert_eq!(size(&grown, 0, 0).0, rows + 2);
    let shrunk = grown
        .edit_table_topology(0, 0)?
        .remove_row(RowDeletion::footer(0))?
        .remove_row(RowDeletion::header(0))?
        .commit()?
        .into_package();
    assert_eq!(shrunk.table_header_settings(0, 0)?, settings);
    assert_eq!(cells(&shrunk, 0, 0), cells(&package, 0, 0));
    Ok(())
}

#[test]
fn refusals_are_typed() -> TestResult {
    let package = filled()?;
    let (rows, _) = size(&package, 0, 0);
    assert!(matches!(
        package.edit_table_topology(0, 0)?.commit(),
        Err(Error::NoStagedOperation)
    ));
    assert!(matches!(
        package
            .edit_table_topology(0, 0)?
            .insert_row(RowInsertion::body(rows as usize + 1)),
        Err(Error::PositionOutOfRange { .. })
    ));
    assert!(matches!(
        package.edit_table_topology(0, 1),
        Err(Error::TableNotFound)
    ));
    assert!(matches!(
        package.edit_table_topology(3, 0),
        Err(Error::SheetNotFound)
    ));
    let mut edit = package.edit_table_topology(0, 0)?;
    for _ in 0..5 {
        match edit.insert_row(RowInsertion::header(0)) {
            Ok(next) => edit = next,
            Err(error) => {
                assert_eq!(error, Error::SectionFull);
                return Ok(());
            },
        }
    }
    panic!("a sixth header row was accepted");
}

/// Every corpus table either takes a row and a column, and gives them back
/// through the inverse byte for byte, or refuses with a typed reason.
#[test]
fn corpus_tables_round_trip_or_refuse_typed() -> TestResult {
    let mut round_trips = 0_usize;
    let mut removals = 0_usize;
    let mut content_removals = 0_usize;
    let mut content_refusals = Vec::new();
    let mut clear_refusals = Vec::new();
    let mut refused = Vec::new();
    for path in corpus()? {
        let package = Package::open(&path)?;
        let source = bytes(&package)?;
        for sheet in 0..package.sheets().len() {
            for table in 0..package.sheets()[sheet].table_count() {
                let before = cells(&package, sheet, table);
                let result = package
                    .edit_table_topology(sheet, table)?
                    .insert_row(RowInsertion::body(0))?
                    .insert_column(ColumnInsertion::body(0))?
                    .commit();
                let commit = match result {
                    Ok(commit) => commit,
                    Err(Error::Unsupported { reason }) => {
                        refused.push((path.clone(), reason));
                        continue;
                    },
                    Err(error) => {
                        panic!(
                            "{} table {table}: untyped refusal {error:?}",
                            path.display()
                        )
                    },
                };
                let edited = commit.package();
                assert_eq!(cells(edited, sheet, table).len(), before.len());
                let restored = edited.apply_table_topology(&commit.patch().inverse())?;
                assert_eq!(bytes(restored.package())?, source, "{}", path.display());
                // Removing them again gives the cells back.
                let removed = edited
                    .edit_table_topology(sheet, table)?
                    .remove_column(ColumnDeletion::body(0))?
                    .remove_row(RowDeletion::body(0))?
                    .commit();
                match removed {
                    Ok(removed) => {
                        assert_eq!(cells(removed.package(), sheet, table), before);
                        removals += 1;
                    },
                    Err(Error::CellClear { error }) => clear_refusals.push(error),
                    Err(error) => panic!("{}: removal {error:?}", path.display()),
                }
                round_trips += 1;
                // Removing a stored row and column drops their cells.
                let dropped = package
                    .edit_table_topology(sheet, table)?
                    .remove_row(RowDeletion::body(0))?
                    .remove_column(ColumnDeletion::body(0))?
                    .commit();
                match dropped {
                    Ok(dropped) => {
                        let restored = dropped
                            .package()
                            .apply_table_topology(&dropped.patch().inverse())?;
                        assert_eq!(bytes(restored.package())?, source);
                        content_removals += 1;
                    },
                    Err(Error::CellClear { error }) => clear_refusals.push(error),
                    Err(Error::Unsupported { reason }) => content_refusals.push(reason),
                    Err(error) => panic!("{}: content removal {error:?}", path.display()),
                }
            }
        }
    }
    eprintln!(
        "round trips: {round_trips}; removals: {removals}; content removals: {content_removals} (refused {content_refusals:?}); clear refusals: {clear_refusals:?}; refused: {refused:?}"
    );
    assert!(round_trips >= 20, "only {round_trips} round trips");
    assert!(removals >= 20, "only {removals} removals");
    assert!(
        content_removals >= 20,
        "only {content_removals} content removals"
    );
    assert!(refused.iter().all(|(_, reason)| matches!(
        reason,
        Unsupported::Formulas
            | Unsupported::ReferencedByOtherOwners
            | Unsupported::Merges
            | Unsupported::Filters
            | Unsupported::Categories
            | Unsupported::StorageLayout
    )));
    Ok(())
}
