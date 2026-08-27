use std::env;

use litchi_iwa::keynote::{
    KeynoteDocumentBuilder, KeynoteEditor, KeynoteTableCellUpdate, KeynoteTableCellValue,
    KeynoteTableTitleSettings,
};
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};
use litchi_keynote::slide::table::dimension::{Dimension, Size};
use litchi_keynote::slide::table::formula::{
    FormulaCachedValue, FormulaCellReference, FormulaExpression,
};
use litchi_keynote::slide::table::headers::{
    Count as KeynoteHeaderCount, Settings as KeynoteHeaderSettings,
};
use litchi_numbers::table::topology::{ColumnInsertion, RowInsertion};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = env::args()
        .nth(1)
        .unwrap_or_else(|| "keynote-table.key".to_owned());
    let mut editor = KeynoteDocumentBuilder::new()
        .title("Regional outlook")
        .subtitle("Created entirely by litchi-iwa")
        .build()?;
    let table = editor.add_slide_table(
        0,
        "Forecast",
        4,
        3,
        DrawablePoint { x: 350.0, y: 390.0 },
        DrawableSize {
            width: 1_220.0,
            height: 430.0,
        },
    )?;
    let mut updates = Vec::new();
    for (row, values) in [
        ["Region", "Q1", "Q2"],
        ["North", "120", "145"],
        ["South", "98", "132"],
        ["West", "110", "127"],
    ]
    .into_iter()
    .enumerate()
    {
        for (column, value) in values.into_iter().enumerate() {
            let value = if row == 0 || column == 0 {
                KeynoteTableCellValue::Text(value.to_owned())
            } else {
                KeynoteTableCellValue::number(value.parse()?)?
            };
            updates.push(KeynoteTableCellUpdate::new(row, column, value));
        }
    }
    editor.set_slide_table_cells(0, table.model_object_id, updates)?;
    editor = set_focused_keynote_table_headers(
        editor,
        KeynoteHeaderSettings {
            header_rows: Some(KeynoteHeaderCount::ONE),
            header_columns: Some(KeynoteHeaderCount::ONE),
            footer_rows: Some(KeynoteHeaderCount::ONE),
            ..Default::default()
        },
    )?;
    editor = set_focused_keynote_table_title(
        editor,
        KeynoteTableTitleSettings::new(Some(true), Some(true)),
    )?;
    for (column, width) in [440.0, 390.0, 390.0].into_iter().enumerate() {
        editor = set_focused_keynote_table_dimension(
            editor,
            Dimension::Column(column),
            Size::points(width)?,
        )?;
    }
    for (row, height) in [90.0, 100.0, 110.0, 130.0].into_iter().enumerate() {
        editor = set_focused_keynote_table_dimension(
            editor,
            Dimension::Row(row),
            Size::points(height)?,
        )?;
    }
    editor.set_slide_table_formula(
        0,
        table.model_object_id,
        3,
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
    editor.insert_slide_table_row(0, table.model_object_id, RowInsertion::body(2))?;
    editor.insert_slide_table_column(0, table.model_object_id, ColumnInsertion::body(1))?;
    editor.set_slide_table_cells(
        0,
        table.model_object_id,
        [
            KeynoteTableCellUpdate::new(3, 0, KeynoteTableCellValue::Text("Central".to_owned())),
            KeynoteTableCellUpdate::new(3, 1, KeynoteTableCellValue::number(105.0)?),
            KeynoteTableCellUpdate::new(3, 2, KeynoteTableCellValue::Text("review".to_owned())),
            KeynoteTableCellUpdate::new(3, 3, KeynoteTableCellValue::number(139.0)?),
            KeynoteTableCellUpdate::new(0, 2, KeynoteTableCellValue::Text("Status".to_owned())),
        ],
    )?;
    editor.save(&output)?;
    println!("created {output}");
    Ok(())
}

fn set_focused_keynote_table_headers(
    editor: KeynoteEditor,
    settings: KeynoteHeaderSettings,
) -> Result<KeynoteEditor, Box<dyn std::error::Error>> {
    let package = litchi_keynote::Package::from_bytes(&editor.to_bytes()?)?;
    let commit = package
        .edit_slide_table_headers(
            litchi_keynote::SlideSelector::index(0),
            litchi_keynote::TableSelector::index(0),
        )?
        .set(settings)
        .commit()?;
    let mut bytes = Vec::new();
    commit.package().write_to(&mut bytes)?;
    Ok(KeynoteEditor::from_bytes(&bytes)?)
}

fn set_focused_keynote_table_dimension(
    editor: KeynoteEditor,
    dimension: Dimension,
    size: Size,
) -> Result<KeynoteEditor, Box<dyn std::error::Error>> {
    let source = editor.to_bytes()?;
    let package = litchi_keynote::Package::from_bytes(&source)?;
    let result = package
        .edit_slide_table_dimension_size(
            litchi_keynote::SlideSelector::index(0),
            litchi_keynote::TableSelector::index(0),
            dimension,
        )?
        .set(size)
        .commit();
    let commit = match result {
        Ok(commit) => commit,
        Err(litchi_keynote::SlideTableDimensionError::UnsupportedDependency) => {
            // KeynoteDocumentBuilder snapshots intentionally do not satisfy
            // the focused owner's exact-source authority proof. Keep the
            // builder's private physical sizing and do not fall back to the
            // retired raw-ID writer.
            if editor.to_bytes()? != source {
                return Err("focused Keynote dimension rejection mutated the builder".into());
            }
            return Ok(editor);
        },
        Err(error) => return Err(error.into()),
    };
    let mut bytes = Vec::new();
    commit.package().write_to(&mut bytes)?;
    let reopened = KeynoteEditor::from_bytes(&bytes)?;
    let package = litchi_keynote::Package::from_bytes(&bytes)?;
    if package.slide_table_dimension_size(
        litchi_keynote::SlideSelector::index(0),
        litchi_keynote::TableSelector::index(0),
        dimension,
    )? != size
    {
        return Err("focused Keynote dimension failed round-trip validation".into());
    }
    Ok(reopened)
}

fn set_focused_keynote_table_title(
    editor: KeynoteEditor,
    settings: KeynoteTableTitleSettings,
) -> Result<KeynoteEditor, Box<dyn std::error::Error>> {
    let package = litchi_keynote::Package::from_bytes(&editor.to_bytes()?)?;
    let commit = package
        .edit_slide_table_title(
            litchi_keynote::SlideSelector::index(0),
            litchi_keynote::TableSelector::index(0),
        )?
        .set(settings)
        .commit()?;
    let mut bytes = Vec::new();
    commit.package().write_to(&mut bytes)?;
    Ok(KeynoteEditor::from_bytes(&bytes)?)
}
