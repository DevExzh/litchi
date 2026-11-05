//! Body-table cell values (record 0309) across the Pages corpus.

use std::path::PathBuf;

use litchi_iwa_common::table::cell::value::Value as Read;
use litchi_iwa_common::table::coordinate::At;
use litchi_pages::table::values::{Error, Value};
use litchi_pages::{BodyTableSelector, Package};

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

fn bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut output = Vec::new();
    package.write_to(&mut output)?;
    Ok(output)
}

#[test]
fn corpus_body_tables_take_values_or_refuse_typed() -> TestResult {
    let root = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../test-data/iwork/pages");
    let mut paths: Vec<PathBuf> = std::fs::read_dir(root)?
        .map(|entry| entry.map(|entry| entry.path()))
        .collect::<Result<_, _>>()?;
    paths.retain(|path| {
        path.extension()
            .is_some_and(|extension| extension == "pages")
    });
    paths.sort();
    let mut written = 0_usize;
    let mut refused = Vec::new();
    for path in paths {
        let Ok(package) = Package::open(&path) else {
            continue;
        };
        let Ok(read) = package.body_table_cells(BodyTableSelector::index(0)) else {
            continue;
        };
        let dimensions = read.table().dimensions();
        let (rows, columns) = (dimensions.rows(), dimensions.columns());
        let source = bytes(&package)?;
        let mut edit = package.edit_body_table_values(BodyTableSelector::index(0))?;
        edit.set(rows - 1, columns - 1, Value::Text("written".into()))?
            .set(rows - 1, 0, Value::Number(7.25))?
            .set(0, columns - 1, Value::Boolean(true))?;
        match edit.commit() {
            Ok(commit) => {
                let reopened = Package::from_bytes(&bytes(commit.package())?)?;
                let read = reopened.body_table_cells(BodyTableSelector::index(0))?;
                assert_eq!(
                    read.table().get(At::new(rows - 1, columns - 1)),
                    Some(&Read::Text("written".into())),
                    "{}",
                    path.display()
                );
                let restored = commit
                    .package()
                    .apply_body_table_values(&commit.patch().inverse())?;
                assert_eq!(bytes(restored.package())?, source);
                // Clearing a written cell.
                let mut clear = commit
                    .package()
                    .edit_body_table_values(BodyTableSelector::index(0))?;
                clear.set(rows - 1, columns - 1, Value::Empty)?;
                let cleared = clear.commit()?.into_package();
                let read = cleared.body_table_cells(BodyTableSelector::index(0))?;
                assert!(!matches!(
                    read.table().get(At::new(rows - 1, columns - 1)),
                    Some(Read::Text(_))
                ));
                written += 1;
            },
            Err(
                error @ (Error::Refused(litchi_pages::table::values::Refusal::Formulas)
                | Error::Refused(litchi_pages::table::values::Refusal::UnsupportedStorage)),
            ) => {
                refused.push((path, error));
            },
            Err(error) => panic!("{}: {error:?}", path.display()),
        }
    }
    eprintln!("written {written}; refused {refused:?}");
    assert!(written >= 5, "only {written} tables written");
    Ok(())
}

#[test]
fn refusals_are_typed() -> TestResult {
    let path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/pages/body-table-cells-native.pages");
    let package = Package::open(path)?;
    let mut edit = package.edit_body_table_values(BodyTableSelector::index(0))?;
    assert!(matches!(
        edit.set(1_000, 0, Value::Empty),
        Err(Error::Refused(
            litchi_pages::table::values::Refusal::InvalidCell
        ))
    ));
    assert!(matches!(
        edit.set(0, 0, Value::Number(f64::NAN)),
        Err(Error::Refused(
            litchi_pages::table::values::Refusal::InvalidCell
        ))
    ));
    let mut missing = package.edit_body_table_values(BodyTableSelector::index(40))?;
    assert!(matches!(
        missing.set(0, 0, Value::Empty),
        Err(Error::Refused(
            litchi_pages::table::values::Refusal::TableNotFound
        ))
    ));
    assert!(matches!(
        package
            .edit_body_table_values(BodyTableSelector::index(0))?
            .commit(),
        Err(Error::Refused(
            litchi_pages::table::values::Refusal::NoStagedOperation
        ))
    ));
    Ok(())
}

/// A formula authored into a body table reads back as its source and is
/// removed again by a value write (record 0376a).
#[test]
fn body_table_formulas_author_and_clear() -> TestResult {
    use litchi_pages::table::formula::{
        FormulaBinaryOperator, FormulaCachedValue, FormulaCellReference, FormulaExpression,
    };
    let path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/pages/body-table-catalog-native.pages");
    let package = Package::open(&path)?;
    let mut edit = package.edit_body_table_values(BodyTableSelector::index(0))?;
    edit.set_formula_cached(
        2,
        2,
        FormulaExpression::binary(
            FormulaBinaryOperator::Add,
            FormulaExpression::Cell(FormulaCellReference::relative(2, 1)),
            FormulaExpression::Number(1.0),
        ),
        FormulaCachedValue::number(3.0)?,
    )?;
    let authored = edit.commit()?.into_package();
    let reopened = Package::from_bytes(&bytes(&authored)?)?;
    let read = reopened.body_table_cells(BodyTableSelector::index(0))?;
    assert_eq!(
        read.table().get(At::new(2, 2)),
        Some(&Read::Formula("=(B3+1)".into()))
    );
    let mut edit = reopened.edit_body_table_values(BodyTableSelector::index(0))?;
    edit.set(2, 2, Value::Empty)?;
    let cleared = edit.commit()?.into_package();
    let read = cleared.body_table_cells(BodyTableSelector::index(0))?;
    assert!(!matches!(
        read.table().get(At::new(2, 2)),
        Some(Read::Formula(_))
    ));
    // A date cache is refused at staging.
    let mut edit = cleared.edit_body_table_values(BodyTableSelector::index(0))?;
    assert!(matches!(
        edit.set_formula_cached(
            2,
            2,
            FormulaExpression::Number(1.0),
            FormulaCachedValue::date(1.0)?,
        ),
        Err(Error::Refused(
            litchi_pages::table::values::Refusal::Formulas
        ))
    ));
    Ok(())
}
