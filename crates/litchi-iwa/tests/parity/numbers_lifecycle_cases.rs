//! Numbers sheet and table lifecycle parity (record 0370e): the host
//! editor's `add_empty_sheet`, `remove_sheet`, `add_empty_table` and
//! `remove_table` against the leaf's `edit_sheet_lifecycle` and
//! `edit_table_lifecycle`.

use litchi_iwa::numbers;

use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier, fixture};
use crate::leaf_bytes;

#[derive(Clone, Copy, Debug)]
enum Operation {
    AddSheet,
    RemoveSheet(usize),
    AddTable(usize),
    RemoveTable(usize, usize),
}

fn host_route(source: &[u8], operation: Operation) -> TestResult<Vec<u8>> {
    let mut editor = numbers::Draft::from_bytes(source)?;
    match operation {
        Operation::AddSheet => {
            editor.add_empty_sheet("Added Sheet")?;
        },
        Operation::RemoveSheet(sheet) => {
            editor.remove_sheet(litchi_numbers::SheetSelector::index(sheet))?;
        },
        Operation::AddTable(sheet) => {
            editor.add_empty_table(
                litchi_numbers::SheetSelector::index(sheet),
                "Added Table",
                4,
                3,
            )?;
        },
        Operation::RemoveTable(_sheet, table) => {
            // The host's selector counts tables across the workbook; the
            // cases remove a table of the first sheet.
            editor.remove_table(litchi_numbers::TableSelector::index(table))?;
        },
    }
    Ok(editor.to_bytes()?)
}

fn leaf_route(source: &[u8], operation: Operation) -> TestResult<LeafRun> {
    let package = litchi_numbers::Package::from_bytes(source)?;
    let (output, restored) = match operation {
        Operation::AddSheet | Operation::RemoveSheet(_) => {
            let edit = package.edit_sheet_lifecycle();
            let edit = match operation {
                Operation::AddSheet => edit.add("Added Sheet")?,
                Operation::RemoveSheet(sheet) => edit.remove(sheet)?,
                _ => unreachable!(),
            };
            let commit = edit.commit()?;
            let restored = commit
                .package()
                .apply_sheet_lifecycle(&commit.patch().inverse())?;
            (
                leaf_bytes!(commit.package()),
                leaf_bytes!(restored.package()),
            )
        },
        Operation::AddTable(sheet) | Operation::RemoveTable(sheet, _) => {
            let edit = package.edit_table_lifecycle(sheet)?;
            let edit = match operation {
                Operation::AddTable(_) => edit.add("Added Table", 4, 3)?,
                Operation::RemoveTable(_, table) => edit.remove(table)?,
                _ => unreachable!(),
            };
            let commit = edit.commit()?;
            let restored = commit
                .package()
                .apply_table_lifecycle(&commit.patch().inverse())?;
            (
                leaf_bytes!(commit.package()),
                leaf_bytes!(restored.package()),
            )
        },
    };
    Ok(LeafRun::with_inverse(output, restored))
}

fn leaf_shape(bytes: &[u8]) -> TestResult<String> {
    let package = litchi_numbers::Package::from_bytes(bytes)?;
    Ok(format!("{:?}", package.sheets()))
}

fn run(name: &'static str, tier: Tier, input: &Fixture, operation: Operation) -> TestResult {
    Case::new(name, Format::Numbers, tier, Evidence::Queued, input)
        .leaf_structural_metadata()
        .host(|source| host_route(source, operation))
        .leaf(|source| leaf_route(source, operation))
        .read("leaf shape", leaf_shape)
        .run()?;
    Ok(())
}

/// Every lifecycle operation on the corpus (exploratory: the failures list
/// the remaining divergences, record 0370e §3).
#[test]
#[ignore = "exploratory: sheet and table add/remove are not at parity yet (record 0370e)"]
fn numbers_lifecycle_corpus_parity() -> TestResult {
    let _ = Evidence::Queued;
    let mut failures = Vec::new();
    let mut cases = 0;
    for path in crate::numbers_cases::NUMBERS_CORPUS {
        let input = fixture(path)?;
        for operation in [
            Operation::AddSheet,
            Operation::AddTable(0),
            Operation::RemoveTable(0, 0),
            Operation::RemoveSheet(0),
        ] {
            let host = host_route(&input.bytes, operation);
            let leaf = leaf_route(&input.bytes, operation);
            match (host, leaf) {
                (Ok(_), Ok(_)) => {
                    cases += 1;
                    if let Err(error) = run(
                        "numbers_lifecycle_corpus_parity",
                        Tier::B,
                        &input,
                        operation,
                    ) {
                        failures.push(format!("{path} {operation:?}: {error}"));
                    }
                },
                (Ok(_), Err(error)) => {
                    failures.push(format!("{path} {operation:?}: leaf refuses: {error}"));
                },
                (Err(host), Ok(_)) => eprintln!("{path} {operation:?}: host refuses ({host})"),
                (Err(_), Err(_)) => {},
            }
        }
    }
    eprintln!("lifecycle routes compared: {cases}");
    if failures.is_empty() {
        Ok(())
    } else {
        Err(failures.join("\n").into())
    }
}

/// One resize route: grow or shrink the first table by `rows` × `columns`.
fn resize_case(
    name: &'static str,
    tier: Tier,
    input: &Fixture,
    delta: (i64, i64),
) -> TestResult<bool> {
    let package = litchi_numbers::Package::from_bytes(&input.bytes)?;
    let Some(table) = package.workbook_ref().sheets()[0].tables().next() else {
        return Ok(false);
    };
    let rows = u32::try_from(i64::from(table.row_count()) + delta.0)?;
    let columns = u32::try_from(i64::from(table.column_count()) + delta.1)?;
    let host = |source: &[u8]| -> TestResult<Vec<u8>> {
        let mut editor = numbers::Draft::from_bytes(source)?;
        editor.resize_table(
            litchi_numbers::TableSelector::index(0),
            rows as usize,
            columns as usize,
        )?;
        Ok(editor.to_bytes()?)
    };
    let leaf = |source: &[u8]| -> TestResult<LeafRun> {
        let package = litchi_numbers::Package::from_bytes(source)?;
        let commit = package
            .edit_table_resize(0usize, 0usize)?
            .resize(rows, columns)?
            .commit()?;
        let restored = commit
            .package()
            .apply_table_resize(&commit.patch().inverse())?;
        Ok(LeafRun::with_inverse(
            leaf_bytes!(commit.package()),
            leaf_bytes!(restored.package()),
        ))
    };
    match (host(&input.bytes), leaf(&input.bytes)) {
        (Ok(_), Ok(_)) => {},
        (Err(_), Err(_)) => return Ok(false),
        (Ok(_), Err(error)) => return Err(format!("leaf refuses: {error}").into()),
        (Err(error), Ok(_)) => {
            eprintln!(
                "{} {delta:?}: host refuses ({error}); the leaf edits",
                input.path
            );
            return Ok(false);
        },
    }
    Case::new(name, Format::Numbers, tier, Evidence::Queued, input)
        .host(host)
        .leaf(leaf)
        .read("leaf shape", leaf_shape)
        .run()?;
    Ok(true)
}

/// Growing and shrinking the first table on `basic.numbers`.
#[test]
fn numbers_table_resize_parity() -> TestResult {
    let _ = Evidence::Queued;
    let input = fixture("iwork/numbers/basic.numbers")?;
    for delta in [(3, 0), (0, 2), (5, 4), (-1, 0)] {
        resize_case("numbers_table_resize_parity", Tier::A, &input, delta)
            .map_err(|error| format!("{delta:?}: {error}"))?;
    }
    Ok(())
}

/// Resizing every corpus fixture's first table.
#[test]
fn numbers_table_resize_corpus_parity() -> TestResult {
    let _ = Evidence::Queued;
    let mut failures = Vec::new();
    let mut compared = 0;
    for path in crate::numbers_cases::NUMBERS_CORPUS {
        let input = fixture(path)?;
        for delta in [(2, 0), (0, 1), (4, 3), (-1, 0), (0, -1)] {
            match resize_case("numbers_table_resize_corpus_parity", Tier::B, &input, delta) {
                Ok(true) => compared += 1,
                Ok(false) => {},
                Err(error) => failures.push(format!("{path} {delta:?}: {error}")),
            }
        }
    }
    eprintln!("resize routes compared: {compared}");
    if failures.is_empty() {
        assert!(compared >= 60, "only {compared} routes compared");
        Ok(())
    } else {
        Err(failures.join("\n").into())
    }
}
