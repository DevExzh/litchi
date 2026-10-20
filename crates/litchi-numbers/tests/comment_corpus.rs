//! Numbers cell comments over the corpus (record 0318): native packages keep
//! each table's comment list, its tiles and the comment authors in members
//! of their own, so a comment edit rewrites several members. Every corpus
//! table takes a new comment on a stored cell, gives it back through the
//! inverse byte for byte, and takes a further edit and a clear; every
//! natively commented cell is rewritten and cleared the same way; an empty
//! cell in a row its tile does not store takes one too (record 0329).

use std::collections::BTreeMap;
use std::path::PathBuf;

use litchi_numbers::{At, Package};

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

fn text(package: &Package, sheet: usize, table: usize, position: At) -> Option<String> {
    package
        .table_cell_comment(sheet, table, position)
        .ok()
        .flatten()
        .map(|comment| comment.text().to_owned())
}

/// Set, edit again, clear; check each state and the exact inverse.
fn round_trip(
    package: &Package,
    sheet: usize,
    table: usize,
    position: At,
    label: &str,
) -> TestResult {
    let source = bytes(package)?;
    let set = package
        .set_table_cell_comment(sheet, table, position, "record 0318")
        .map_err(|error| format!("{label}: set {error:?}"))?;
    let reopened = Package::from_bytes(&bytes(set.package())?)?;
    assert_eq!(
        text(&reopened, sheet, table, position).as_deref(),
        Some("record 0318"),
        "{label}"
    );
    assert!(
        reopened
            .table_cell_comments(sheet, table)?
            .iter()
            .any(|(at, comment)| *at == position && comment.text() == "record 0318"),
        "{label}"
    );
    let restored = set
        .package()
        .apply_table_cell_comment(&set.patch().inverse())?;
    assert_eq!(bytes(restored.package())?, source, "{label}: inverse");
    // The published package takes further comment edits.
    let edited = reopened
        .set_table_cell_comment(sheet, table, position, "edited again")
        .map_err(|error| format!("{label}: second set {error:?}"))?;
    assert_eq!(
        text(edited.package(), sheet, table, position).as_deref(),
        Some("edited again"),
        "{label}"
    );
    let cleared = edited
        .package()
        .clear_table_cell_comment(sheet, table, position)
        .map_err(|error| format!("{label}: clear {error:?}"))?;
    assert_eq!(
        text(cleared.package(), sheet, table, position),
        None,
        "{label}"
    );
    let restored = cleared
        .package()
        .apply_table_cell_comment(&cleared.patch().inverse())?;
    assert_eq!(
        text(restored.package(), sheet, table, position).as_deref(),
        Some("edited again"),
        "{label}"
    );
    Ok(())
}

#[test]
fn every_corpus_table_takes_and_gives_back_a_cell_comment() -> TestResult {
    let mut summary = BTreeMap::<&str, usize>::new();
    for path in corpus()? {
        let name = path
            .file_name()
            .and_then(|name| name.to_str())
            .unwrap_or("")
            .to_owned();
        let package = Package::open(&path)?;
        for sheet in 0..package.sheets().len() {
            for (table, semantic) in package.sheets()[sheet].tables().enumerate() {
                let label = format!("{name} sheet {sheet} table {table}");
                let mut commented = Vec::new();
                let mut plain = None;
                for cell in semantic.iter_cells() {
                    let position = cell.position();
                    match package.table_cell_comment(sheet, table, position) {
                        Ok(Some(_)) => commented.push(position),
                        Ok(None) => {
                            plain.get_or_insert(position);
                        },
                        Err(error) => panic!("{label}: read {error:?}"),
                    }
                }
                // The table read lists exactly the cells that read a comment.
                let listed: Vec<At> = package
                    .table_cell_comments(sheet, table)?
                    .into_iter()
                    .map(|(position, _)| position)
                    .collect();
                assert_eq!(listed, commented, "{label}");
                if let Some(position) = plain {
                    round_trip(&package, sheet, table, position, &label)?;
                    *summary.entry("created").or_default() += 1;
                }
                // Record 0329: an empty cell, in a row the tile does not
                // store (the table's last cell), takes a comment too.
                let stored: Vec<At> = semantic.iter_cells().map(|cell| cell.position()).collect();
                let last = At::new(
                    semantic.row_count().saturating_sub(1),
                    semantic.column_count().saturating_sub(1),
                );
                if !stored.contains(&last) {
                    round_trip(&package, sheet, table, last, &format!("{label} empty"))?;
                    *summary.entry("created on an empty cell").or_default() += 1;
                    // The other cells keep their values.
                    let after = package
                        .set_table_cell_comment(sheet, table, last, "x")?
                        .into_package();
                    let values = |package: &Package| {
                        package.sheets()[sheet].tables().nth(table).map(|table| {
                            table
                                .iter_cells()
                                .filter(|cell| cell.position() != last)
                                .map(|cell| (cell.position(), cell.value().clone()))
                                .collect::<Vec<_>>()
                        })
                    };
                    assert_eq!(values(&after), values(&package), "{label}: values");
                }
                for position in commented {
                    round_trip(&package, sheet, table, position, &label)?;
                    *summary.entry("native comment rewritten").or_default() += 1;
                }
            }
        }
    }
    eprintln!("{summary:?}");
    assert!(
        summary.get("created").copied().unwrap_or(0) >= 35,
        "{summary:?}"
    );
    assert!(
        summary
            .get("native comment rewritten")
            .copied()
            .unwrap_or(0)
            >= 6,
        "{summary:?}"
    );
    assert!(
        summary
            .get("created on an empty cell")
            .copied()
            .unwrap_or(0)
            >= 25,
        "{summary:?}"
    );
    Ok(())
}

/// A created workbook, and a table added to it, take comments too.
#[test]
fn created_workbooks_and_added_tables_take_cell_comments() -> TestResult {
    use litchi_numbers::table::cells::Input;
    let package = Package::new()
        .edit_table_lifecycle(0)?
        .add("Added", 6, 4)?
        .commit()?
        .into_package();
    let mut package = package;
    for table in 0..2 {
        package = package
            .edit_table_cells(0, table)?
            .set_a1("B2", Input::number(1.0)?)?
            .commit()?
            .into_package();
    }
    for table in 0..2 {
        round_trip(
            &package,
            0,
            table,
            At::new(1, 1),
            &format!("created table {table}"),
        )?;
    }
    Ok(())
}
