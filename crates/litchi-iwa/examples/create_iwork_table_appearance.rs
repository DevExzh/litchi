//! Create Pages, Numbers, and Keynote files with native table appearance overrides.
use litchi_numbers::table::headers::{Count as HeaderCount, Settings as HeaderSettings};

use std::path::PathBuf;

use litchi_iwa::keynote::{KeynoteDocumentBuilder, KeynoteEditor};
use litchi_iwa::numbers::{NumbersDocumentBuilder, NumbersEditor};
use litchi_iwa::pages::{PagesDocumentBuilder, PagesEditor};
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};
use litchi_iwa::table_appearance::{
    TableAppearance, TableGridlineVisibility, TableGridlines, TableRowBanding, TableRowSizing,
};
use litchi_numbers::{Package, SheetSelector, TableSelector};

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
    let numbers_table = TableSelector::index(0);
    numbers = set_focused_table_headers(
        numbers,
        HeaderSettings {
            header_rows: Some(HeaderCount::ONE),
            header_columns: Some(HeaderCount::ONE),
            footer_rows: Some(HeaderCount::ONE),
            ..Default::default()
        },
    )?;
    numbers.set_table_appearance(numbers_table, APPEARANCE)?;
    numbers.save(&numbers_path)?;
    assert_eq!(
        NumbersEditor::open(&numbers_path)?.table_appearance(numbers_table)?,
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
        HeaderSettings {
            header_rows: Some(HeaderCount::ONE),
            header_columns: Some(HeaderCount::ONE),
            footer_rows: Some(HeaderCount::ONE),
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
        HeaderSettings {
            header_rows: Some(HeaderCount::ONE),
            header_columns: Some(HeaderCount::ONE),
            footer_rows: Some(HeaderCount::ONE),
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
