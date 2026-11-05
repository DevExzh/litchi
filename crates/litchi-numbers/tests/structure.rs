//! Structural transactions (record 0301): sheet lifecycle and table order,
//! on the created workbook and across the Numbers corpus.

use std::path::PathBuf;

use litchi_numbers::Package;
use litchi_numbers::sheet::lifecycle::{Error as SheetError, InvalidName};
use litchi_numbers::table::order::Error as OrderError;

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

fn corpus() -> TestResult<Vec<PathBuf>> {
    let root = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../test-data/iwork/numbers");
    let mut paths: Vec<PathBuf> = std::fs::read_dir(root)?
        .map(|entry| entry.map(|entry| entry.path()))
        .collect::<Result<_, _>>()?;
    paths.retain(|path| {
        path.extension()
            .is_some_and(|extension| extension == "numbers")
    });
    paths.sort();
    Ok(paths)
}

fn bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn sheet_names(package: &Package) -> Vec<String> {
    package
        .sheets()
        .iter()
        .map(|sheet| sheet.name().to_owned())
        .collect()
}

#[test]
fn a_created_workbook_takes_a_sheet_and_gives_it_back() -> TestResult {
    let package = Package::new();
    let source = bytes(&package)?;
    let commit = package.edit_sheet_lifecycle().add("Second")?.commit()?;
    assert!(commit.diagnostics().changed());
    assert!(!commit.patch().is_noop());
    let added = commit.package();
    assert_eq!(sheet_names(added), ["Sheet 1", "Second"]);
    assert_eq!(added.sheets()[1].table_count(), 0);
    // The added sheet is a full citizen: it can move and be renamed.
    let moved = added
        .edit_sheet_order()
        .move_sheet(1usize, 0)?
        .commit()?
        .into_package();
    assert_eq!(sheet_names(&moved), ["Second", "Sheet 1"]);
    let renamed = added
        .edit_names()
        .rename_sheet(litchi_numbers::SheetSelector::index(1), "Renamed")?
        .commit()?
        .into_package();
    assert_eq!(sheet_names(&renamed), ["Sheet 1", "Renamed"]);

    // The inverse restores the source byte for byte.
    let restored = added.apply_sheet_lifecycle(&commit.patch().inverse())?;
    assert_eq!(bytes(restored.package())?, source);
    // The forward patch applies only to its exact source.
    assert!(matches!(
        added.apply_sheet_lifecycle(commit.patch()),
        Err(SheetError::PatchConflict)
    ));
    let replayed = package.apply_sheet_lifecycle(commit.patch())?;
    assert_eq!(bytes(replayed.package())?, bytes(added)?);

    // Removing the added sheet gives the original sheets back.
    let removed = added
        .edit_sheet_lifecycle()
        .remove("Second")?
        .commit()?
        .into_package();
    assert_eq!(sheet_names(&removed), ["Sheet 1"]);
    let reopened = Package::from_bytes(&bytes(&removed)?)?;
    assert_eq!(sheet_names(&reopened), ["Sheet 1"]);
    Ok(())
}

#[test]
fn sheet_lifecycle_refusals_are_typed() -> TestResult {
    let package = Package::new();
    assert!(matches!(
        package.edit_sheet_lifecycle().add(""),
        Err(SheetError::Refused(
            litchi_numbers::sheet::lifecycle::Refusal::InvalidName {
                reason: InvalidName::Empty
            }
        ))
    ));
    assert!(matches!(
        package.edit_sheet_lifecycle().add("Sheet 1"),
        Err(SheetError::Refused(
            litchi_numbers::sheet::lifecycle::Refusal::InvalidName {
                reason: InvalidName::Duplicate
            }
        ))
    ));
    assert!(matches!(
        package.edit_sheet_lifecycle().add("a\0b"),
        Err(SheetError::Refused(
            litchi_numbers::sheet::lifecycle::Refusal::InvalidName {
                reason: InvalidName::ContainsNul
            }
        ))
    ));
    assert!(matches!(
        package.edit_sheet_lifecycle().remove(0usize),
        Err(SheetError::Refused(
            litchi_numbers::sheet::lifecycle::Refusal::LastSheet
        ))
    ));
    assert!(matches!(
        package.edit_sheet_lifecycle().remove("Missing"),
        Err(SheetError::Refused(
            litchi_numbers::sheet::lifecycle::Refusal::SheetNotFound
        ))
    ));
    assert!(matches!(
        package.edit_sheet_lifecycle().commit(),
        Err(SheetError::Refused(
            litchi_numbers::sheet::lifecycle::Refusal::NoStagedOperation
        ))
    ));
    let two = package
        .edit_sheet_lifecycle()
        .add("Second")?
        .commit()?
        .into_package();
    // The first sheet goes with its table.
    let one = two
        .edit_sheet_lifecycle()
        .remove(0usize)?
        .commit()?
        .into_package();
    assert_eq!(sheet_names(&one), ["Second"]);
    Ok(())
}

#[test]
fn every_corpus_workbook_takes_a_sheet_and_gives_it_back() -> TestResult {
    let mut added = 0;
    for path in corpus()? {
        let package = Package::open(&path)?;
        let names = sheet_names(&package);
        let source = bytes(&package)?;
        let name = "Litchi structural probe";
        let commit = match package.edit_sheet_lifecycle().add(name)?.commit() {
            Ok(commit) => commit,
            Err(SheetError::Refused(
                litchi_numbers::sheet::lifecycle::Refusal::UnsupportedSource,
            )) => continue,
            Err(error) => return Err(format!("{}: {error}", path.display()).into()),
        };
        let mut expected = names.clone();
        expected.push(name.to_owned());
        assert_eq!(
            sheet_names(commit.package()),
            expected,
            "{}",
            path.display()
        );
        let reopened = Package::from_bytes(&bytes(commit.package())?)?;
        assert_eq!(sheet_names(&reopened), expected, "{}", path.display());
        let restored = commit
            .package()
            .apply_sheet_lifecycle(&commit.patch().inverse())?;
        assert_eq!(bytes(restored.package())?, source, "{}", path.display());
        let removed = commit
            .package()
            .edit_sheet_lifecycle()
            .remove(name)?
            .commit()
            .map_err(|error| format!("{}: {error}", path.display()))?;
        assert_eq!(sheet_names(removed.package()), names, "{}", path.display());
        added += 1;
    }
    assert!(added > 30, "only {added} corpus workbooks took a sheet");
    Ok(())
}

#[test]
fn tables_move_within_their_sheet_and_back() -> TestResult {
    let mut moved = 0;
    for path in corpus()? {
        let package = Package::open(&path)?;
        for (sheet_index, sheet) in package.sheets().iter().enumerate() {
            if sheet.table_count() < 2 {
                continue;
            }
            let names: Vec<String> = sheet
                .tables()
                .map(|table| table.name().to_owned())
                .collect();
            let source = bytes(&package)?;
            let commit = package
                .edit_table_order(sheet_index)?
                .move_table(names.len() - 1, 0)?
                .commit()
                .map_err(|error| format!("{}: {error}", path.display()))?;
            let mut expected = names.clone();
            let last = expected.pop().ok_or("no table")?;
            expected.insert(0, last);
            let after: Vec<String> = commit.package().sheets()[sheet_index]
                .tables()
                .map(|table| table.name().to_owned())
                .collect();
            assert_eq!(after, expected, "{}", path.display());
            let restored = commit
                .package()
                .apply_table_order(&commit.patch().inverse())?;
            assert_eq!(bytes(restored.package())?, source, "{}", path.display());
            moved += 1;
        }
    }
    // The created workbook has one table per sheet: a move is out of range.
    let package = Package::new();
    assert!(matches!(
        package.edit_table_order(0usize)?.move_table(0usize, 1),
        Err(OrderError::Refused(
            litchi_numbers::table::order::Refusal::DestinationOutOfRange { .. }
        ))
    ));
    let unchanged = package
        .edit_table_order(0usize)?
        .move_table(0usize, 0)?
        .commit()?;
    assert!(unchanged.patch().is_noop());
    assert!(!unchanged.diagnostics().changed());
    eprintln!("moved tables in {moved} sheets");
    Ok(())
}

#[test]
fn tables_are_removed_with_their_private_graph() -> TestResult {
    use litchi_numbers::table::lifecycle::Error as TableError;
    let package = Package::new();
    let source = bytes(&package)?;
    let commit = package
        .edit_table_lifecycle(0usize)?
        .remove(0usize)?
        .commit()?;
    let emptied = commit.package();
    assert_eq!(emptied.sheets()[0].table_count(), 0);
    let reopened = Package::from_bytes(&bytes(emptied)?)?;
    assert_eq!(reopened.sheets()[0].table_count(), 0);
    let restored = emptied.apply_table_lifecycle(&commit.patch().inverse())?;
    assert_eq!(bytes(restored.package())?, source);
    assert!(matches!(
        emptied.edit_table_lifecycle(0usize)?.remove(0usize),
        Err(TableError::Refused(
            litchi_numbers::table::lifecycle::Refusal::TableNotFound
        ))
    ));
    // An emptied sheet can then go.
    let two = emptied
        .edit_sheet_lifecycle()
        .add("Second")?
        .commit()?
        .into_package();
    let one = two
        .edit_sheet_lifecycle()
        .remove(0usize)?
        .commit()?
        .into_package();
    assert_eq!(sheet_names(&one), ["Second"]);

    let mut removed = 0;
    let mut refused = 0;
    for path in corpus()? {
        let package = Package::open(&path)?;
        let source = bytes(&package)?;
        for (sheet_index, sheet) in package.sheets().iter().enumerate() {
            for table_index in 0..sheet.table_count() {
                let mut names: Vec<String> = sheet
                    .tables()
                    .map(|table| table.name().to_owned())
                    .collect();
                let commit = match package
                    .edit_table_lifecycle(sheet_index)?
                    .remove(table_index)?
                    .commit()
                {
                    Ok(commit) => commit,
                    Err(
                        TableError::Refused(
                            litchi_numbers::table::lifecycle::Refusal::UnsupportedDependency,
                        )
                        | TableError::Refused(
                            litchi_numbers::table::lifecycle::Refusal::UnsupportedSource,
                        ),
                    ) => {
                        refused += 1;
                        continue;
                    },
                    Err(error) => {
                        return Err(format!(
                            "{} {sheet_index}/{table_index}: {error}",
                            path.display()
                        )
                        .into());
                    },
                };
                names.remove(table_index);
                let reopened = Package::from_bytes(&bytes(commit.package())?)?;
                let after: Vec<String> = reopened.sheets()[sheet_index]
                    .tables()
                    .map(|table| table.name().to_owned())
                    .collect();
                assert_eq!(after, names, "{}", path.display());
                let restored = commit
                    .package()
                    .apply_table_lifecycle(&commit.patch().inverse())?;
                assert_eq!(bytes(restored.package())?, source, "{}", path.display());
                removed += 1;
            }
        }
    }
    eprintln!("removed {removed} corpus tables, {refused} refused");
    assert!(removed > 30);
    Ok(())
}

#[test]
fn corpus_sheets_are_removed_with_their_tables() -> TestResult {
    let mut removed = 0;
    let mut refused = 0;
    for path in corpus()? {
        let package = Package::open(&path)?;
        // A one-sheet workbook first takes a second sheet.
        let package = if package.sheets().len() == 1 {
            package
                .edit_sheet_lifecycle()
                .add("Litchi keeper")?
                .commit()?
                .into_package()
        } else {
            package
        };
        let source = bytes(&package)?;
        let mut names = sheet_names(&package);
        let commit = match package.edit_sheet_lifecycle().remove(0usize)?.commit() {
            Ok(commit) => commit,
            Err(
                SheetError::Refused(
                    litchi_numbers::sheet::lifecycle::Refusal::UnsupportedDependency,
                )
                | SheetError::Refused(litchi_numbers::sheet::lifecycle::Refusal::UnsupportedSource)
                | SheetError::Refused(litchi_numbers::sheet::lifecycle::Refusal::SheetNotEmpty),
            ) => {
                refused += 1;
                continue;
            },
            Err(error) => return Err(format!("{}: {error}", path.display()).into()),
        };
        names.remove(0);
        let reopened = Package::from_bytes(&bytes(commit.package())?)?;
        assert_eq!(sheet_names(&reopened), names, "{}", path.display());
        let restored = commit
            .package()
            .apply_sheet_lifecycle(&commit.patch().inverse())?;
        assert_eq!(bytes(restored.package())?, source, "{}", path.display());
        removed += 1;
    }
    eprintln!("removed {removed} corpus sheets, {refused} refused");
    assert!(removed > 20);
    Ok(())
}

#[test]
fn a_new_table_is_native_writable_and_removable() -> TestResult {
    use litchi_numbers::table::cells::Input;
    use litchi_numbers::{SheetSelector, TableSelector};
    let package = Package::new();
    let source = bytes(&package)?;
    let commit = package
        .edit_table_lifecycle(0usize)?
        .add("Added", 4, 3)?
        .commit()?;
    let added = commit.package();
    let sheet = &added.sheets()[0];
    assert_eq!(sheet.table_count(), 2);
    let table = sheet.tables().nth(1).ok_or("no table")?;
    assert_eq!(
        (table.name(), table.row_count(), table.column_count()),
        ("Added", 4, 3)
    );
    // Deterministic: the same edit of the same source is byte-identical.
    let again = package
        .edit_table_lifecycle(0usize)?
        .add("Added", 4, 3)?
        .commit()?;
    assert_eq!(bytes(again.package())?, bytes(added)?);
    // The new table takes values.
    let written = added
        .edit_table_cells(SheetSelector::index(0), TableSelector::index(1))?
        .set_a1("B2", Input::text("hello")?)?
        .set_a1("C4", Input::number(42.0)?)?
        .commit()?
        .into_package();
    let reopened = Package::from_bytes(&bytes(&written)?)?;
    let table = reopened.sheets()[0].tables().nth(1).ok_or("no table")?;
    let values: Vec<String> = table
        .iter_cells()
        .map(|cell| format!("{:?}", cell.value()))
        .collect();
    assert_eq!(values.len(), 2, "{values:?}");
    // The inverse restores the source byte for byte; removal takes it away.
    let restored = added.apply_table_lifecycle(&commit.patch().inverse())?;
    assert_eq!(bytes(restored.package())?, source);
    let removed = written
        .edit_table_lifecycle(0usize)?
        .remove("Added")?
        .commit()?
        .into_package();
    assert_eq!(removed.sheets()[0].table_count(), 1);
    // A sheet added to the workbook takes a table too.
    let two = package
        .edit_sheet_lifecycle()
        .add("Second")?
        .commit()?
        .into_package();
    let with_table = two
        .edit_table_lifecycle(1usize)?
        .add("Data", 2, 2)?
        .commit()?
        .into_package();
    assert_eq!(with_table.sheets()[1].table_count(), 1);
    let renamed = with_table
        .edit_names()
        .rename_table(SheetSelector::index(1), TableSelector::index(0), "Renamed")?
        .commit()?
        .into_package();
    assert_eq!(
        renamed.sheets()[1]
            .tables()
            .next()
            .ok_or("no table")?
            .name(),
        "Renamed"
    );
    Ok(())
}

#[test]
fn every_corpus_workbook_takes_a_table() -> TestResult {
    use litchi_numbers::table::cells::Input;
    use litchi_numbers::table::lifecycle::Error as TableError;
    use litchi_numbers::{SheetSelector, TableSelector};
    let mut added = 0;
    let mut writable = 0;
    let mut removable = 0;
    let mut refused = Vec::new();
    for path in corpus()? {
        let package = Package::open(&path)?;
        let source = bytes(&package)?;
        let tables = package.sheets()[0].table_count();
        let commit = match package
            .edit_table_lifecycle(0usize)?
            .add("Litchi probe", 3, 2)?
            .commit()
        {
            Ok(commit) => commit,
            Err(
                error @ (TableError::Refused(
                    litchi_numbers::table::lifecycle::Refusal::UnsupportedSource,
                )
                | TableError::Refused(litchi_numbers::table::lifecycle::Refusal::NoTemplate)),
            ) => {
                refused.push(format!("{}: {error}", path.display()));
                continue;
            },
            Err(error) => return Err(format!("{}: {error}", path.display()).into()),
        };
        let reopened = Package::from_bytes(&bytes(commit.package())?)?;
        assert_eq!(
            reopened.sheets()[0].table_count(),
            tables + 1,
            "{}",
            path.display()
        );
        // Where the source's own first table takes a value, the new table
        // takes one too (the cell writer's admitted scope is unchanged).
        let template_writable = package
            .edit_table_cells(SheetSelector::index(0), TableSelector::index(0))
            .and_then(|edit| edit.set_a1("B2", Input::text("x")?))
            .and_then(|edit| edit.commit())
            .is_ok();
        if template_writable {
            // The cell writer still needs every formula-bearing table's
            // geometry; a workbook with formulas elsewhere refuses, typed.
            match reopened
                .edit_table_cells(SheetSelector::index(0), TableSelector::index(tables))?
                .set_a1("B2", Input::text("x")?)?
                .commit()
            {
                Ok(_) => writable += 1,
                Err(litchi_numbers::table::cells::Error::Refused(
                    litchi_numbers::table::cells::Refusal::UnsupportedDependency { .. },
                )) => {},
                Err(error) => return Err(format!("{} cell: {error}", path.display()).into()),
            }
        }
        let restored = commit
            .package()
            .apply_table_lifecycle(&commit.patch().inverse())?;
        assert_eq!(bytes(restored.package())?, source, "{}", path.display());
        match commit
            .package()
            .edit_table_lifecycle(0usize)?
            .remove(tables)?
            .commit()
        {
            Ok(removed) => {
                assert_eq!(removed.package().sheets()[0].table_count(), tables);
                removable += 1;
            },
            // A native formula graph the removal phase does not rewrite
            // (a range tile with several from-to ranges) refuses, typed.
            Err(
                TableError::Refused(litchi_numbers::table::lifecycle::Refusal::UnsupportedSource)
                | TableError::Refused(
                    litchi_numbers::table::lifecycle::Refusal::UnsupportedDependency,
                ),
            ) => {},
            Err(error) => return Err(format!("{} remove: {error}", path.display()).into()),
        }
        added += 1;
    }
    eprintln!(
        "added a table to {added} corpus workbooks ({writable} took a value, {removable} removed again); refused: {refused:#?}"
    );
    assert!(
        removable > 25,
        "only {removable} added tables were removed again"
    );
    assert!(added > 30);
    assert!(writable > 15, "only {writable} new tables took a value");
    Ok(())
}
