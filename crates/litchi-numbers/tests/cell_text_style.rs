//! Cell character properties (record 0309): bold, italic, size, font and
//! colour of cell text, on the created workbook and across the corpus.

use std::path::PathBuf;

use litchi_numbers::cell::text_style::{CharacterProperties, Error};
use litchi_numbers::table::cells::Input;
use litchi_numbers::{At, Package};

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

fn bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut output = Vec::new();
    package.write_to(&mut output)?;
    Ok(output)
}

fn bold() -> CharacterProperties {
    let mut properties = CharacterProperties::new();
    properties.bold = Some(true);
    properties
}

#[test]
fn cells_take_text_styles_on_a_created_workbook() -> TestResult {
    let package = Package::new()
        .edit_table_cells(0, 0)?
        .set(At::new(1, 1), Input::Text("b".into()))?
        .set(At::new(2, 1), Input::Text("c".into()))?
        .set(At::new(2, 2), Input::number(4.0)?)?
        .commit()?
        .into_package();
    let source = bytes(&package)?;
    let base = package.table_cell_text_style(0, 0, At::new(1, 1))?;
    let mut red = CharacterProperties::new();
    red.color = Some([200, 0, 0]);
    red.size = Some(18.0);
    red.font_name = Some("Helvetica".into());
    let commit = package
        .edit_table_cell_text_style(0, 0)?
        .set(At::new(1, 1), bold())?
        .set(At::new(2, 1), bold())?
        .set(At::new(2, 2), red.clone())?
        .set(At::new(2, 2), {
            let mut italic = CharacterProperties::new();
            italic.italic = Some(true);
            italic
        })?
        .commit()?;
    let styled = commit.package();
    let reopened = Package::from_bytes(&bytes(styled)?)?;
    for candidate in [styled, &reopened] {
        let b = candidate.table_cell_text_style(0, 0, At::new(1, 1))?;
        assert_eq!(b.bold, Some(true));
        // Unset properties stay inherited from the base style.
        assert_eq!(b.size, base.size);
        assert_eq!(b.font_name, base.font_name);
        let c = candidate.table_cell_text_style(0, 0, At::new(2, 2))?;
        assert_eq!(c.color, Some([200, 0, 0]));
        assert_eq!(c.size, Some(18.0));
        assert_eq!(c.italic, Some(true));
        assert_eq!(c.font_name.as_deref(), Some("Helvetica"));
        // Values are untouched.
        let table = candidate.sheets()[0].tables().next().expect("table");
        assert_eq!(table.iter_cells().count(), 3);
    }
    // Restyling a styled cell builds on its style and releases the old key.
    let mut unbold = CharacterProperties::new();
    unbold.bold = Some(false);
    let again = styled
        .edit_table_cell_text_style(0, 0)?
        .set(At::new(1, 1), unbold)?
        .commit()?
        .into_package();
    assert_eq!(
        again.table_cell_text_style(0, 0, At::new(1, 1))?.bold,
        Some(false)
    );
    assert_eq!(
        again.table_cell_text_style(0, 0, At::new(2, 1))?.bold,
        Some(true)
    );
    let restored = styled.apply_table_cell_text_style(&commit.patch().inverse())?;
    assert_eq!(bytes(restored.package())?, source);
    // A value write keeps the style.
    let rewritten = again
        .edit_table_cells(0, 0)?
        .set(At::new(2, 1), Input::Text("changed".into()))?
        .commit()?
        .into_package();
    assert_eq!(
        rewritten.table_cell_text_style(0, 0, At::new(2, 1))?.bold,
        Some(true)
    );
    Ok(())
}

#[test]
fn exclusive_existing_style_batch_rewrites_each_cell() -> TestResult {
    let package = Package::new()
        .edit_table_cells(0, 0)?
        .set(At::new(1, 1), Input::Text("b".into()))?
        .commit()?
        .into_package();
    let package = package
        .edit_table_cells(0, 0)?
        .set(At::new(2, 1), Input::Text("c".into()))?
        .commit()?
        .into_package();
    let mut bold = CharacterProperties::new();
    bold.bold = Some(true);
    let styled = package
        .edit_table_cell_text_style(0, 0)?
        .set(At::new(1, 1), bold)?
        .commit()?
        .into_package();
    let mut italic = CharacterProperties::new();
    italic.italic = Some(true);
    let styled = styled
        .edit_table_cell_text_style(0, 0)?
        .set(At::new(2, 1), italic)?
        .commit()?
        .into_package();
    let mut second = CharacterProperties::new();
    second.italic = Some(true);
    let mut first = CharacterProperties::new();
    first.bold = Some(true);
    let commit = styled
        .edit_table_cell_text_style(0, 0)?
        .set(At::new(1, 1), second)?
        .set(At::new(2, 1), first)?
        .commit()?;
    assert_eq!(
        commit
            .package()
            .table_cell_text_style(0, 0, At::new(1, 1))?
            .italic,
        Some(true)
    );
    assert_eq!(
        commit
            .package()
            .table_cell_text_style(0, 0, At::new(2, 1))?
            .bold,
        Some(true)
    );
    let restored = commit
        .package()
        .apply_table_cell_text_style(&commit.patch().inverse())?;
    assert_eq!(bytes(restored.package())?, bytes(&styled)?);
    Ok(())
}

#[test]
fn refusals_are_typed() -> TestResult {
    let package = Package::new();
    assert!(matches!(
        package
            .edit_table_cell_text_style(0, 0)?
            .set(At::new(0, 0), CharacterProperties::new()),
        Err(Error::Refused(
            litchi_numbers::cell::text_style::Refusal::EmptyStyle
        ))
    ));
    assert!(matches!(
        package
            .edit_table_cell_text_style(0, 0)?
            .set(At::new(99, 0), bold()),
        Err(Error::Refused(
            litchi_numbers::cell::text_style::Refusal::CellNotFound
        ))
    ));
    // An empty cell takes a style-only cell (record 0370).
    let commit = package
        .edit_table_cell_text_style(0, 0)?
        .set(At::new(3, 3), bold())?
        .commit()?;
    assert_eq!(
        commit
            .package()
            .table_cell_text_style(0, 0, At::new(3, 3))?
            .bold,
        Some(true)
    );
    assert!(matches!(
        package.edit_table_cell_text_style(0, 0)?.commit(),
        Err(Error::Refused(
            litchi_numbers::cell::text_style::Refusal::NoStagedOperation
        ))
    ));
    Ok(())
}

/// Every stored cell of the corpus takes bold, or refuses typed.
#[test]
fn corpus_cells_take_bold_or_refuse_typed() -> TestResult {
    let root = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../test-data/iwork/numbers");
    let mut styled = 0_usize;
    let mut refused = Vec::new();
    let mut paths: Vec<PathBuf> = std::fs::read_dir(root)?
        .map(|entry| entry.map(|entry| entry.path()))
        .collect::<Result<_, _>>()?;
    paths.retain(|path| {
        path.extension()
            .is_some_and(|extension| extension == "numbers")
    });
    paths.sort();
    for path in paths {
        let package = Package::open(&path)?;
        let source = bytes(&package)?;
        let Some(cell) = package.sheets()[0]
            .tables()
            .next()
            .and_then(|table| table.iter_cells().next())
            .map(|cell| cell.position())
        else {
            continue;
        };
        match package
            .edit_table_cell_text_style(0, 0)?
            .set(cell, bold())?
            .commit()
        {
            Ok(commit) => {
                assert_eq!(
                    commit.package().table_cell_text_style(0, 0, cell)?.bold,
                    Some(true),
                    "{}",
                    path.display()
                );
                let restored = commit
                    .package()
                    .apply_table_cell_text_style(&commit.patch().inverse())?;
                assert_eq!(bytes(restored.package())?, source);
                styled += 1;
            },
            Err(
                error @ Error::Refused(
                    litchi_numbers::cell::text_style::Refusal::UnsupportedDependency,
                ),
            ) => {
                refused.push((path, error));
            },
            Err(error) => panic!("{}: {error:?}", path.display()),
        }
    }
    eprintln!("styled {styled}; refused {refused:?}");
    assert!(styled >= 25, "only {styled} styled");
    Ok(())
}

/// A created workbook authors formulas (request 18, record 0309): its
/// table owner carries the native table extents.
#[test]
fn formulas_work_on_a_created_workbook() -> TestResult {
    use litchi_numbers::cell::Value as CellValue;
    use litchi_numbers::formula::Expression;
    let package = Package::new()
        .edit_table_cells(0, 0)?
        .set(At::new(1, 1), Input::number(2.0)?)?
        .commit()?
        .into_package();
    let commit = package
        .edit_table_cells(0, 0)?
        .set_formula(At::new(2, 1), Expression::parse("B2*3")?)?
        .commit()?;
    let reopened = Package::from_bytes(&bytes(commit.package())?)?;
    let table = reopened.sheets()[0].tables().next().expect("table");
    assert!(matches!(
        table.get(At::new(2, 1)),
        Some(CellValue::Formula(_))
    ));
    Ok(())
}
