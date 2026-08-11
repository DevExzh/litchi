//! Create a Numbers spreadsheet without an input document or template.
use litchi_numbers::table::headers::{Count as HeaderCount, Settings as HeaderSettings};

use litchi_iwa::numbers::{
    FormulaCachedValue, FormulaCellReference, FormulaExpression, NumbersDocumentBuilder,
    NumbersEditor,
};
use litchi_iwa::text::{Font, TextStyle};
use litchi_numbers::table::CellPosition;
use litchi_numbers::table::cells::{Change, Input};
use litchi_numbers::table::topology::RowInsertion;
use litchi_numbers::{Package, SheetSelector, TableSelector};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = std::env::args()
        .nth(1)
        .ok_or("usage: create_numbers <output.numbers>")?;
    let mut editor = NumbersDocumentBuilder::new()
        .table_name("Forecast")
        .table_dimensions(4, 3)
        .build()?;
    let table_id = editor.tables()?.remove(0).id();
    let table = TableSelector::index(0);
    editor.insert_table_row(table, RowInsertion::body(2))?;
    let mut changes = Vec::new();
    for (column, heading) in ["Region", "Q1", "Q2"].into_iter().enumerate() {
        changes.push(Change::set(
            CellPosition::new(0, u32::try_from(column)?),
            Input::text(heading)?,
        ));
    }
    for (row, region, q1, q2) in [(1, "North", 120.0, 145.0), (2, "South", 98.0, 132.0)] {
        changes.extend([
            Change::set(CellPosition::new(row, 0), Input::text(region)?),
            Change::set(CellPosition::new(row, 1), Input::number(q1)?),
            Change::set(CellPosition::new(row, 2), Input::number(q2)?),
        ]);
    }
    changes.extend([
        Change::set(CellPosition::new(3, 0), Input::text("Central")?),
        Change::set(CellPosition::new(3, 1), Input::number(105.0)?),
        Change::set(CellPosition::new(3, 2), Input::number(139.0)?),
    ]);
    editor = set_focused_table_cells(editor, changes)?;
    editor.set_table_cell_text_style(table_id, 1, 0, TextStyle::default().with_bold(true))?;
    editor.set_table_cell_text_font(table_id, 1, 0, Font::named("CourierNewPSMT")?)?;
    editor = set_focused_table_headers(
        editor,
        HeaderSettings {
            header_rows: Some(HeaderCount::ONE),
            header_columns: Some(HeaderCount::ONE),
            footer_rows: Some(HeaderCount::ONE),
            ..Default::default()
        },
    )?;
    editor.set_formula_with_cached_value(
        table_id,
        4,
        1,
        FormulaExpression::function(
            "SUM",
            [FormulaExpression::range(
                FormulaCellReference::relative(1, 1),
                FormulaCellReference::relative(2, 1),
            )],
        ),
        FormulaCachedValue::Number(218.0.try_into()?),
    )?;
    editor.save(output)?;
    Ok(())
}

fn set_focused_table_headers(
    editor: NumbersEditor,
    settings: HeaderSettings,
) -> Result<NumbersEditor, Box<dyn std::error::Error>> {
    let package = Package::from_bytes(&editor.to_bytes()?)?;
    let commit = package
        .edit_table_headers(SheetSelector::index(0), TableSelector::index(0))?
        .set(settings)
        .commit()?;
    let mut bytes = Vec::new();
    commit.package().write_to(&mut bytes)?;
    Ok(NumbersEditor::from_bytes(&bytes)?)
}

fn set_focused_table_cells(
    editor: NumbersEditor,
    changes: impl IntoIterator<Item = Change>,
) -> Result<NumbersEditor, Box<dyn std::error::Error>> {
    let package = Package::from_bytes(&editor.to_bytes()?)?;
    let commit = package
        .edit_table_cells(SheetSelector::index(0), TableSelector::index(0))?
        .extend(changes)?
        .commit()?;
    let mut bytes = Vec::new();
    commit.package().write_to(&mut bytes)?;
    Ok(NumbersEditor::from_bytes(&bytes)?)
}
