//! Create Pages, Numbers, and Keynote files with native table appearance overrides.
use litchi_numbers::table::headers::{Count as HeaderCount, Settings as HeaderSettings};

use std::path::PathBuf;

use litchi_iwa::keynote::{KeynoteDocumentBuilder, KeynoteEditor};
use litchi_iwa::numbers::{NumbersDocumentBuilder, NumbersEditor};
use litchi_iwa::pages::{PagesDocumentBuilder, PagesEditor};
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};
use litchi_keynote::slide::table::headers::{
    Count as KeynoteHeaderCount, Settings as KeynoteHeaderSettings,
};
use litchi_numbers::{
    Appearance, Banding, GridlineVisibility, Gridlines, Package, RowSizing, SheetSelector,
    TableSelector,
};
use litchi_pages::table::headers::{Count as PagesHeaderCount, Settings as PagesHeaderSettings};
use litchi_pages::{BodyTableSelector, Package as PagesPackage};

const APPEARANCE: Appearance = Appearance {
    row_banding: Banding::Enabled,
    row_sizing: RowSizing::FitCellContents,
    gridlines: Gridlines {
        body_horizontal: GridlineVisibility::Hidden,
        header_columns_horizontal: GridlineVisibility::Visible,
        body_vertical: GridlineVisibility::Hidden,
        header_rows_vertical: GridlineVisibility::Visible,
        footer_rows_vertical: GridlineVisibility::Hidden,
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
    numbers = set_focused_numbers_table_appearance(numbers, APPEARANCE)?;
    numbers.save(&numbers_path)?;
    let numbers_package = Package::from_bytes(&numbers.to_bytes()?)?;
    let read_focused_appearance = Package::table_appearance;
    assert_eq!(
        read_focused_appearance(&numbers_package, SheetSelector::index(0), numbers_table)?,
        APPEARANCE
    );

    let pages_path = output.join("table-appearance.pages");
    let mut pages = PagesDocumentBuilder::new()
        .body_text("Created from scratch with litchi-iwa.\n")
        .body_table("Appearance", 6, 3)
        .build()?;
    set_focused_pages_table_headers(
        &mut pages,
        PagesHeaderSettings {
            header_rows: Some(PagesHeaderCount::ONE),
            header_columns: Some(PagesHeaderCount::ONE),
            footer_rows: Some(PagesHeaderCount::ONE),
            ..Default::default()
        },
    )?;
    let pages_table = BodyTableSelector::index(0);
    pages = set_focused_pages_table_appearance(pages, pages_table, APPEARANCE)?;
    pages.save(&pages_path)?;
    let pages_package = PagesPackage::from_bytes(&PagesEditor::open(&pages_path)?.to_bytes()?)?;
    assert_eq!(
        pages_package.body_table_appearance(pages_table)?,
        APPEARANCE
    );

    let keynote_path = output.join("table-appearance.key");
    let mut keynote = KeynoteDocumentBuilder::new()
        .title("Native table appearance")
        .build()?;
    let _keynote_table = keynote.add_slide_table(
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
    keynote = set_focused_keynote_table_headers(
        keynote,
        KeynoteHeaderSettings {
            header_rows: Some(KeynoteHeaderCount::ONE),
            header_columns: Some(KeynoteHeaderCount::ONE),
            footer_rows: Some(KeynoteHeaderCount::ONE),
            ..Default::default()
        },
    )?;
    keynote = set_focused_keynote_table_appearance(keynote, APPEARANCE)?;
    keynote.save(&keynote_path)?;
    let keynote_package =
        litchi_keynote::Package::from_bytes(&KeynoteEditor::open(&keynote_path)?.to_bytes()?)?;
    assert_eq!(
        keynote_package.slide_table_appearance(
            litchi_keynote::SlideSelector::index(0),
            litchi_keynote::TableSelector::index(0),
        )?,
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

fn set_focused_numbers_table_appearance(
    editor: NumbersEditor,
    appearance: Appearance,
) -> Result<NumbersEditor, Box<dyn std::error::Error>> {
    let package = Package::from_bytes(&editor.to_bytes()?)?;
    let commit = package
        .edit_table_appearance(SheetSelector::index(0), TableSelector::index(0))?
        .set(appearance)
        .commit()?;
    let mut bytes = Vec::new();
    commit.package().write_to(&mut bytes)?;
    Ok(NumbersEditor::from_bytes(&bytes)?)
}

fn set_focused_pages_table_headers(
    editor: &mut PagesEditor,
    settings: PagesHeaderSettings,
) -> Result<(), Box<dyn std::error::Error>> {
    let package = PagesPackage::from_bytes(&editor.to_bytes()?)?;
    let commit = package
        .edit_body_table_header_settings(BodyTableSelector::index(0))?
        .set(settings)
        .commit()?;
    let mut bytes = Vec::new();
    commit.package().write_to(&mut bytes)?;
    *editor = PagesEditor::from_bytes(&bytes)?;
    Ok(())
}

fn set_focused_pages_table_appearance(
    editor: PagesEditor,
    selector: BodyTableSelector<'static>,
    appearance: Appearance,
) -> Result<PagesEditor, Box<dyn std::error::Error>> {
    let package = PagesPackage::from_bytes(&editor.to_bytes()?)?;
    let commit = package
        .edit_body_table_appearance(selector)?
        .set(appearance)
        .commit()?;
    let mut bytes = Vec::new();
    commit.package().write_to(&mut bytes)?;
    Ok(PagesEditor::from_bytes(&bytes)?)
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

fn set_focused_keynote_table_appearance(
    editor: KeynoteEditor,
    appearance: Appearance,
) -> Result<KeynoteEditor, Box<dyn std::error::Error>> {
    let package = litchi_keynote::Package::from_bytes(&editor.to_bytes()?)?;
    let commit = package
        .edit_slide_table_appearance(
            litchi_keynote::SlideSelector::index(0),
            litchi_keynote::TableSelector::index(0),
        )?
        .set(appearance)
        .commit()?;
    let mut bytes = Vec::new();
    commit.package().write_to(&mut bytes)?;
    Ok(KeynoteEditor::from_bytes(&bytes)?)
}
