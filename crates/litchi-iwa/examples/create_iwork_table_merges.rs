//! Create Numbers, Pages, and Keynote files with native merged table cells.

use std::path::{Path, PathBuf};

use litchi_iwa::keynote::KeynoteDocumentBuilder;
use litchi_iwa::numbers::{FormulaCachedValue, FormulaExpression, NumbersDocumentBuilder};
use litchi_iwa::pages::PagesDocumentBuilder;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};
use litchi_numbers::TableSelector;
use litchi_numbers::table::merge::Region;
use litchi_numbers::table::topology::{ColumnDeletion, ColumnInsertion, RowDeletion, RowInsertion};

const TABLE_ROWS: usize = 4;
const TABLE_COLUMNS: usize = 5;
const MERGED_FORMULA_RESULT: f64 = 3.0;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = PathBuf::from(
        std::env::args()
            .nth(1)
            .ok_or("usage: create_iwork_table_merges <output-directory>")?,
    );
    std::fs::create_dir_all(&output)?;
    let region = Region::new(1, 1, 2, 3)?;
    create_numbers(&output, region)?;
    create_pages(&output, region)?;
    create_keynote(&output, region)?;
    Ok(())
}

fn create_numbers(output: &Path, region: Region) -> Result<(), Box<dyn std::error::Error>> {
    let mut editor = NumbersDocumentBuilder::new()
        .table_name("Merged Cells")
        .table_dimensions(TABLE_ROWS, TABLE_COLUMNS)
        .build()?;
    let table_id = editor.tables()?.remove(0).id();
    let table = TableSelector::index(0);
    editor.set_formula_with_cached_value(
        table_id,
        region.row() as usize,
        region.column() as usize,
        FormulaExpression::Number(MERGED_FORMULA_RESULT),
        FormulaCachedValue::Number(MERGED_FORMULA_RESULT.try_into()?),
    )?;
    editor.merge_cells(table_id, region)?;
    editor.save(output.join("merged-cells.numbers"))?;
    editor.unmerge_cells(table_id, region)?;
    editor.save(output.join("unmerged-cells.numbers"))?;
    editor.merge_cells(table_id, region)?;
    editor.insert_table_row(table, RowInsertion::body(1))?;
    editor.insert_table_column(table, ColumnInsertion::body(0))?;
    editor.save(output.join("merged-cells-after-axis-insertions.numbers"))?;
    editor.remove_table_row(table, RowDeletion::body(0))?;
    editor.remove_table_column(table, ColumnDeletion::body(1))?;
    editor.save(output.join("merged-cells-after-axis-deletions.numbers"))?;
    Ok(())
}

fn create_pages(output: &Path, region: Region) -> Result<(), Box<dyn std::error::Error>> {
    let mut editor = PagesDocumentBuilder::new()
        .body_text("Native merged-cell CRUD\n")
        .body_table("Merged Cells", TABLE_ROWS, TABLE_COLUMNS)
        .build()?;
    let table_id = editor.tables()?.remove(0).model_object_id;
    editor.set_table_formula(
        table_id,
        region.row() as usize,
        region.column() as usize,
        FormulaExpression::Number(MERGED_FORMULA_RESULT),
        FormulaCachedValue::Number(MERGED_FORMULA_RESULT.try_into()?),
    )?;
    editor.merge_table_cells(table_id, region)?;
    editor.save(output.join("merged-cells.pages"))?;
    editor.unmerge_table_cells(table_id, region)?;
    editor.save(output.join("unmerged-cells.pages"))?;
    editor.merge_table_cells(table_id, region)?;
    editor.insert_table_row(table_id, RowInsertion::body(1))?;
    editor.insert_table_column(table_id, ColumnInsertion::body(0))?;
    editor.save(output.join("merged-cells-after-axis-insertions.pages"))?;
    editor.remove_table_row(table_id, RowDeletion::body(0))?;
    editor.remove_table_column(table_id, ColumnDeletion::body(1))?;
    editor.save(output.join("merged-cells-after-axis-deletions.pages"))?;
    Ok(())
}

fn create_keynote(output: &Path, region: Region) -> Result<(), Box<dyn std::error::Error>> {
    let mut editor = KeynoteDocumentBuilder::new()
        .title("Native merged-cell CRUD")
        .build()?;
    let table = editor.add_slide_table(
        0,
        "Merged Cells",
        TABLE_ROWS,
        TABLE_COLUMNS,
        DrawablePoint { x: 320.0, y: 360.0 },
        DrawableSize {
            width: 1_280.0,
            height: 480.0,
        },
    )?;
    editor.set_slide_table_formula(
        0,
        table.model_object_id,
        region.row() as usize,
        region.column() as usize,
        FormulaExpression::Number(MERGED_FORMULA_RESULT),
        FormulaCachedValue::Number(MERGED_FORMULA_RESULT.try_into()?),
    )?;
    editor.merge_slide_table_cells(0, table.model_object_id, region)?;
    editor.save(output.join("merged-cells.key"))?;
    editor.unmerge_slide_table_cells(0, table.model_object_id, region)?;
    editor.save(output.join("unmerged-cells.key"))?;
    editor.merge_slide_table_cells(0, table.model_object_id, region)?;
    editor.insert_slide_table_row(0, table.model_object_id, RowInsertion::body(1))?;
    editor.insert_slide_table_column(0, table.model_object_id, ColumnInsertion::body(0))?;
    editor.save(output.join("merged-cells-after-axis-insertions.key"))?;
    editor.remove_slide_table_row(0, table.model_object_id, RowDeletion::body(0))?;
    editor.remove_slide_table_column(0, table.model_object_id, ColumnDeletion::body(1))?;
    editor.save(output.join("merged-cells-after-axis-deletions.key"))?;
    Ok(())
}
