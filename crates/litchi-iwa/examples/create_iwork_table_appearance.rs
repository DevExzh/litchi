//! Create Pages, Numbers, and Keynote files with native table appearance overrides.

use std::path::PathBuf;

use litchi_iwa::keynote::{
    KeynoteDocumentBuilder, KeynoteEditor, KeynoteTableHeaderCount, KeynoteTableHeaderSettings,
};
use litchi_iwa::numbers::{
    NumbersDocumentBuilder, NumbersEditor, NumbersTableHeaderCount, NumbersTableHeaderSettings,
};
use litchi_iwa::pages::{
    PagesDocumentBuilder, PagesEditor, PagesTableHeaderCount, PagesTableHeaderSettings,
};
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};
use litchi_iwa::table_appearance::{
    TableAppearance, TableGridlineVisibility, TableGridlines, TableRowBanding, TableRowSizing,
};

const APPEARANCE: TableAppearance = TableAppearance {
    row_banding: TableRowBanding::Enabled,
    row_sizing: TableRowSizing::FitCellContents,
    gridlines: TableGridlines {
        body_horizontal: TableGridlineVisibility::Hidden,
        header_columns_horizontal: TableGridlineVisibility::Visible,
        body_vertical: TableGridlineVisibility::Hidden,
        header_rows_vertical: TableGridlineVisibility::Visible,
        footer_rows_vertical: TableGridlineVisibility::Hidden,
    },
};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = PathBuf::from(
        std::env::args()
            .nth(1)
            .ok_or("usage: create_iwork_table_appearance <output-directory>")?,
    );
    std::fs::create_dir_all(&output)?;

    let numbers_path = output.join("table-appearance.numbers");
    let mut numbers = NumbersDocumentBuilder::new()
        .table_name("Appearance")
        .table_dimensions(6, 3)
        .build()?;
    let numbers_table = numbers.tables()?.remove(0);
    numbers.set_table_header_settings(
        numbers_table.object_id,
        NumbersTableHeaderSettings {
            header_rows: Some(NumbersTableHeaderCount::ONE),
            header_columns: Some(NumbersTableHeaderCount::ONE),
            footer_rows: Some(NumbersTableHeaderCount::ONE),
            ..Default::default()
        },
    )?;
    numbers.set_table_appearance(numbers_table.object_id, APPEARANCE)?;
    numbers.save(&numbers_path)?;
    assert_eq!(
        NumbersEditor::open(&numbers_path)?.table_appearance(numbers_table.object_id)?,
        APPEARANCE
    );

    let pages_path = output.join("table-appearance.pages");
    let mut pages = PagesDocumentBuilder::new()
        .body_text("Created from scratch with litchi-iwa.\n")
        .body_table("Appearance", 6, 3)
        .build()?;
    let pages_table = pages.tables()?.remove(0);
    pages.set_table_header_settings(
        pages_table.model_object_id,
        PagesTableHeaderSettings {
            header_rows: Some(PagesTableHeaderCount::ONE),
            header_columns: Some(PagesTableHeaderCount::ONE),
            footer_rows: Some(PagesTableHeaderCount::ONE),
            ..Default::default()
        },
    )?;
    pages.set_body_table_appearance(pages_table.model_object_id, APPEARANCE)?;
    pages.save(&pages_path)?;
    assert_eq!(
        PagesEditor::open(&pages_path)?.body_table_appearance(pages_table.model_object_id)?,
        APPEARANCE
    );

    let keynote_path = output.join("table-appearance.key");
    let mut keynote = KeynoteDocumentBuilder::new()
        .title("Native table appearance")
        .build()?;
    let keynote_table = keynote.add_slide_table(
        0,
        "Appearance",
        6,
        3,
        DrawablePoint { x: 320.0, y: 300.0 },
        DrawableSize {
            width: 1_280.0,
            height: 600.0,
        },
    )?;
    keynote.set_slide_table_header_settings(
        0,
        keynote_table.model_object_id,
        KeynoteTableHeaderSettings {
            header_rows: Some(KeynoteTableHeaderCount::ONE),
            header_columns: Some(KeynoteTableHeaderCount::ONE),
            footer_rows: Some(KeynoteTableHeaderCount::ONE),
            ..Default::default()
        },
    )?;
    keynote.set_slide_table_appearance(0, keynote_table.model_object_id, APPEARANCE)?;
    keynote.save(&keynote_path)?;
    assert_eq!(
        KeynoteEditor::open(&keynote_path)?
            .slide_table_appearance(0, keynote_table.model_object_id)?,
        APPEARANCE
    );
    Ok(())
}
