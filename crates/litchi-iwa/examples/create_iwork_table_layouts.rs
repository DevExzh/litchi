//! Create Pages, Numbers, and Keynote files with native table-cell text formatting.

use std::path::{Path, PathBuf};

use litchi_iwa::keynote::{KeynoteDocumentBuilder, KeynoteEditor};
use litchi_iwa::numbers::{CellValue, NumbersDocumentBuilder, NumbersEditor};
use litchi_iwa::pages::{PagesDocumentBuilder, PagesEditor};
use litchi_iwa::shapes::{DrawablePoint, DrawableSize, RgbColorSpace, RgbaColor};
use litchi_iwa::table_cell_layout::{
    TableCellInset, TableCellInsets, TableCellLayout, TableCellTextWrap, TableCellVerticalAlignment,
};
use litchi_iwa::text::{
    ParagraphIndentPoints, ParagraphIndents, ParagraphLineSpacing, ParagraphLineSpacingMultiple,
    ParagraphLineSpacingPoints, ParagraphSpacing, ParagraphSpacingPoints, ParagraphTabAlignment,
    ParagraphTabLeader, ParagraphTabPosition, ParagraphTabStop, ParagraphTabStops, TextAlignment,
    TextBackground, TextBaselineShift, TextCapitalization, TextCharacterSpacing, TextDecorations,
    TextFont, TextLigatures, TextOutline, TextPointSize, TextScript, TextShadow, TextStrikethrough,
    TextStyle, TextUnderline,
};

const ROW: usize = 1;
const COLUMN: usize = 1;
const INSET_POINTS: f32 = 8.0;
const NUMBERS_TEXT_POINTS: f32 = 18.0;
const PAGES_TEXT_POINTS: f32 = 17.0;
const KEYNOTE_TEXT_POINTS: f32 = 19.0;
const NUMBERS_FONT_NAME: &str = "CourierNewPSMT";
const PAGES_FONT_NAME: &str = "AvenirNext-Regular";
const KEYNOTE_FONT_NAME: &str = "Menlo-Regular";
const CELL_TEXT: &str = "Wrapped text\nwith an 8 pt inset";
const OPAQUE: f32 = 1.0;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let arguments = std::env::args().skip(1).collect::<Vec<_>>();
    let output = PathBuf::from(
        arguments
            .first()
            .ok_or("usage: create_iwork_table_layouts <output-directory> [--verify-only]")?,
    );
    if arguments.get(1).map(String::as_str) != Some("--verify-only") {
        std::fs::create_dir_all(&output)?;
        create_numbers(&output.join("table-layouts.numbers"))?;
        create_pages(&output.join("table-layouts.pages"))?;
        create_keynote(&output.join("table-layouts.key"))?;
    }
    verify(&output)?;
    Ok(())
}

fn verify(output: &Path) -> Result<(), Box<dyn std::error::Error>> {
    let numbers = NumbersEditor::open(output.join("table-layouts.numbers"))?;
    let numbers_table = numbers.tables()?.remove(0);
    assert_eq!(
        numbers.table_cell_text_alignment(numbers_table.object_id, ROW, COLUMN)?,
        TextAlignment::Center
    );
    assert_eq!(
        numbers.table_cell_text_style(numbers_table.object_id, ROW, COLUMN)?,
        numbers_text_style()?
    );
    assert_eq!(
        numbers.table_cell_text_font(numbers_table.object_id, ROW, COLUMN)?,
        numbers_text_font()?
    );
    assert_eq!(
        numbers.table_cell_text_color(numbers_table.object_id, ROW, COLUMN)?,
        numbers_text_color()?
    );
    assert_eq!(
        numbers.table_cell_text_decorations(numbers_table.object_id, ROW, COLUMN)?,
        numbers_text_decorations()
    );
    assert_eq!(
        numbers.table_cell_text_baseline_shift(numbers_table.object_id, ROW, COLUMN)?,
        numbers_baseline_shift()?
    );
    assert_eq!(
        numbers.table_cell_text_capitalization(numbers_table.object_id, ROW, COLUMN)?,
        TextCapitalization::AllCaps
    );
    assert_eq!(
        numbers.table_cell_text_character_spacing(numbers_table.object_id, ROW, COLUMN)?,
        numbers_character_spacing()?
    );
    assert_eq!(
        numbers.table_cell_text_ligatures(numbers_table.object_id, ROW, COLUMN)?,
        TextLigatures::RequiredOnly
    );
    assert_eq!(
        numbers.table_cell_text_background(numbers_table.object_id, ROW, COLUMN)?,
        numbers_text_background()?
    );
    assert_eq!(
        numbers.table_cell_text_outline(numbers_table.object_id, ROW, COLUMN)?,
        TextOutline::standard()
    );
    assert_eq!(
        numbers.table_cell_text_script(numbers_table.object_id, ROW, COLUMN)?,
        TextScript::Superscript
    );
    assert_eq!(
        numbers.table_cell_text_shadow(numbers_table.object_id, ROW, COLUMN)?,
        TextShadow::standard()
    );
    assert_eq!(
        numbers.table_cell_paragraph_line_spacing(numbers_table.object_id, ROW, COLUMN)?,
        numbers_paragraph_line_spacing()
    );
    assert_eq!(
        numbers.table_cell_paragraph_spacing(numbers_table.object_id, ROW, COLUMN)?,
        numbers_paragraph_spacing()?
    );
    assert_eq!(
        numbers.table_cell_paragraph_indents(numbers_table.object_id, ROW, COLUMN)?,
        numbers_paragraph_indents()?
    );
    assert_eq!(
        numbers.table_cell_paragraph_tab_stops(numbers_table.object_id, ROW, COLUMN)?,
        numbers_paragraph_tab_stops()?
    );

    let pages = PagesEditor::open(output.join("table-layouts.pages"))?;
    let pages_table = pages.tables()?.remove(0);
    assert_eq!(
        pages.table_cell_text_alignment(pages_table.model_object_id, ROW, COLUMN)?,
        TextAlignment::Right
    );
    assert_eq!(
        pages.table_cell_text_style(pages_table.model_object_id, ROW, COLUMN)?,
        pages_text_style()?
    );
    assert_eq!(
        pages.table_cell_text_font(pages_table.model_object_id, ROW, COLUMN)?,
        pages_text_font()?
    );
    assert_eq!(
        pages.table_cell_text_color(pages_table.model_object_id, ROW, COLUMN)?,
        pages_text_color()?
    );
    assert_eq!(
        pages.table_cell_text_decorations(pages_table.model_object_id, ROW, COLUMN)?,
        pages_text_decorations()
    );
    assert_eq!(
        pages.table_cell_text_baseline_shift(pages_table.model_object_id, ROW, COLUMN)?,
        pages_baseline_shift()?
    );
    assert_eq!(
        pages.table_cell_text_capitalization(pages_table.model_object_id, ROW, COLUMN)?,
        TextCapitalization::SmallCaps
    );
    assert_eq!(
        pages.table_cell_text_character_spacing(pages_table.model_object_id, ROW, COLUMN)?,
        pages_character_spacing()?
    );
    assert_eq!(
        pages.table_cell_text_ligatures(pages_table.model_object_id, ROW, COLUMN)?,
        TextLigatures::All
    );
    assert_eq!(
        pages.table_cell_text_background(pages_table.model_object_id, ROW, COLUMN)?,
        pages_text_background()?
    );
    assert_eq!(
        pages.table_cell_text_outline(pages_table.model_object_id, ROW, COLUMN)?,
        TextOutline::standard()
    );
    assert_eq!(
        pages.table_cell_text_script(pages_table.model_object_id, ROW, COLUMN)?,
        TextScript::Subscript
    );
    assert_eq!(
        pages.table_cell_text_shadow(pages_table.model_object_id, ROW, COLUMN)?,
        TextShadow::standard()
    );
    assert_eq!(
        pages.table_cell_paragraph_line_spacing(pages_table.model_object_id, ROW, COLUMN)?,
        pages_paragraph_line_spacing()?
    );
    assert_eq!(
        pages.table_cell_paragraph_spacing(pages_table.model_object_id, ROW, COLUMN)?,
        pages_paragraph_spacing()?
    );
    assert_eq!(
        pages.table_cell_paragraph_indents(pages_table.model_object_id, ROW, COLUMN)?,
        pages_paragraph_indents()?
    );
    assert_eq!(
        pages.table_cell_paragraph_tab_stops(pages_table.model_object_id, ROW, COLUMN)?,
        pages_paragraph_tab_stops()?
    );

    let keynote = KeynoteEditor::open(output.join("table-layouts.key"))?;
    let keynote_table = keynote.slide_tables(0)?.remove(0);
    assert_eq!(
        keynote.slide_table_cell_text_alignment(0, keynote_table.model_object_id, ROW, COLUMN)?,
        TextAlignment::Justified
    );
    assert_eq!(
        keynote.slide_table_cell_text_style(0, keynote_table.model_object_id, ROW, COLUMN)?,
        keynote_text_style()?
    );
    assert_eq!(
        keynote.slide_table_cell_text_font(0, keynote_table.model_object_id, ROW, COLUMN)?,
        keynote_text_font()?
    );
    assert_eq!(
        keynote.slide_table_cell_text_color(0, keynote_table.model_object_id, ROW, COLUMN)?,
        keynote_text_color()?
    );
    assert_eq!(
        keynote.slide_table_cell_text_decorations(0, keynote_table.model_object_id, ROW, COLUMN)?,
        keynote_text_decorations()
    );
    assert_eq!(
        keynote.slide_table_cell_text_baseline_shift(
            0,
            keynote_table.model_object_id,
            ROW,
            COLUMN,
        )?,
        keynote_baseline_shift()?
    );
    assert_eq!(
        keynote.slide_table_cell_text_capitalization(
            0,
            keynote_table.model_object_id,
            ROW,
            COLUMN,
        )?,
        TextCapitalization::AllCaps
    );
    assert_eq!(
        keynote.slide_table_cell_text_character_spacing(
            0,
            keynote_table.model_object_id,
            ROW,
            COLUMN,
        )?,
        keynote_character_spacing()?
    );
    assert_eq!(
        keynote.slide_table_cell_text_ligatures(0, keynote_table.model_object_id, ROW, COLUMN,)?,
        TextLigatures::RequiredOnly
    );
    assert_eq!(
        keynote.slide_table_cell_text_background(0, keynote_table.model_object_id, ROW, COLUMN,)?,
        keynote_text_background()?
    );
    assert_eq!(
        keynote.slide_table_cell_text_outline(0, keynote_table.model_object_id, ROW, COLUMN)?,
        TextOutline::standard()
    );
    assert_eq!(
        keynote.slide_table_cell_text_script(0, keynote_table.model_object_id, ROW, COLUMN)?,
        TextScript::Superscript
    );
    assert_eq!(
        keynote.slide_table_cell_text_shadow(0, keynote_table.model_object_id, ROW, COLUMN)?,
        TextShadow::standard()
    );
    assert_eq!(
        keynote.slide_table_cell_paragraph_line_spacing(
            0,
            keynote_table.model_object_id,
            ROW,
            COLUMN,
        )?,
        keynote_paragraph_line_spacing()?
    );
    assert_eq!(
        keynote.slide_table_cell_paragraph_spacing(
            0,
            keynote_table.model_object_id,
            ROW,
            COLUMN,
        )?,
        keynote_paragraph_spacing()?
    );
    assert_eq!(
        keynote.slide_table_cell_paragraph_indents(
            0,
            keynote_table.model_object_id,
            ROW,
            COLUMN,
        )?,
        keynote_paragraph_indents()?
    );
    assert_eq!(
        keynote.slide_table_cell_paragraph_tab_stops(
            0,
            keynote_table.model_object_id,
            ROW,
            COLUMN,
        )?,
        keynote_paragraph_tab_stops()?
    );
    Ok(())
}

fn layout(alignment: TableCellVerticalAlignment) -> Result<TableCellLayout, litchi_iwa::Error> {
    Ok(TableCellLayout::default()
        .with_text_wrap(TableCellTextWrap::Wrapped)
        .with_vertical_alignment(alignment)
        .with_insets(TableCellInsets::uniform(TableCellInset::from_points(
            INSET_POINTS,
        )?)))
}

fn numbers_text_style() -> Result<TextStyle, litchi_iwa::Error> {
    Ok(TextStyle::new(TextPointSize::from_points(NUMBERS_TEXT_POINTS)?).with_bold(true))
}

fn pages_text_style() -> Result<TextStyle, litchi_iwa::Error> {
    Ok(TextStyle::new(TextPointSize::from_points(PAGES_TEXT_POINTS)?).with_italic(true))
}

fn keynote_text_style() -> Result<TextStyle, litchi_iwa::Error> {
    Ok(
        TextStyle::new(TextPointSize::from_points(KEYNOTE_TEXT_POINTS)?)
            .with_bold(true)
            .with_italic(true),
    )
}

fn numbers_text_font() -> Result<TextFont, litchi_iwa::Error> {
    TextFont::named(NUMBERS_FONT_NAME)
}

fn pages_text_font() -> Result<TextFont, litchi_iwa::Error> {
    TextFont::named(PAGES_FONT_NAME)
}

fn keynote_text_font() -> Result<TextFont, litchi_iwa::Error> {
    TextFont::named(KEYNOTE_FONT_NAME)
}

fn numbers_text_color() -> Result<RgbaColor, litchi_iwa::Error> {
    const RED: f32 = 0.72;
    const GREEN: f32 = 0.10;
    const BLUE: f32 = 0.14;
    RgbaColor::new(RED, GREEN, BLUE, OPAQUE, RgbColorSpace::Srgb)
}

fn pages_text_color() -> Result<RgbaColor, litchi_iwa::Error> {
    const RED: f32 = 0.10;
    const GREEN: f32 = 0.32;
    const BLUE: f32 = 0.78;
    RgbaColor::new(RED, GREEN, BLUE, OPAQUE, RgbColorSpace::Srgb)
}

fn keynote_text_color() -> Result<RgbaColor, litchi_iwa::Error> {
    const RED: f32 = 0.08;
    const GREEN: f32 = 0.55;
    const BLUE: f32 = 0.28;
    RgbaColor::new(RED, GREEN, BLUE, OPAQUE, RgbColorSpace::Srgb)
}

fn numbers_text_background() -> Result<TextBackground, litchi_iwa::Error> {
    const RED: f32 = 0.96;
    const GREEN: f32 = 0.86;
    const BLUE: f32 = 0.25;
    Ok(TextBackground::Color(RgbaColor::new(
        RED,
        GREEN,
        BLUE,
        OPAQUE,
        RgbColorSpace::Srgb,
    )?))
}

fn pages_text_background() -> Result<TextBackground, litchi_iwa::Error> {
    const RED: f32 = 0.70;
    const GREEN: f32 = 0.92;
    const BLUE: f32 = 0.96;
    Ok(TextBackground::Color(RgbaColor::new(
        RED,
        GREEN,
        BLUE,
        OPAQUE,
        RgbColorSpace::Srgb,
    )?))
}

fn keynote_text_background() -> Result<TextBackground, litchi_iwa::Error> {
    const RED: f32 = 0.98;
    const GREEN: f32 = 0.70;
    const BLUE: f32 = 0.28;
    Ok(TextBackground::Color(RgbaColor::new(
        RED,
        GREEN,
        BLUE,
        OPAQUE,
        RgbColorSpace::Srgb,
    )?))
}

const fn numbers_text_decorations() -> TextDecorations {
    TextDecorations::new(TextUnderline::Single, TextStrikethrough::Single)
}

const fn pages_text_decorations() -> TextDecorations {
    TextDecorations::new(TextUnderline::Single, TextStrikethrough::None)
}

const fn keynote_text_decorations() -> TextDecorations {
    TextDecorations::new(TextUnderline::None, TextStrikethrough::Single)
}

fn numbers_baseline_shift() -> Result<TextBaselineShift, litchi_iwa::Error> {
    const POINTS: f32 = 1.0;
    TextBaselineShift::from_points(POINTS)
}

fn pages_baseline_shift() -> Result<TextBaselineShift, litchi_iwa::Error> {
    const POINTS: f32 = -1.0;
    TextBaselineShift::from_points(POINTS)
}

fn keynote_baseline_shift() -> Result<TextBaselineShift, litchi_iwa::Error> {
    const POINTS: f32 = 2.0;
    TextBaselineShift::from_points(POINTS)
}

fn numbers_character_spacing() -> Result<TextCharacterSpacing, litchi_iwa::Error> {
    const PERCENT: f32 = 10.0;
    TextCharacterSpacing::from_percent(PERCENT)
}

fn pages_character_spacing() -> Result<TextCharacterSpacing, litchi_iwa::Error> {
    const PERCENT: f32 = 6.0;
    TextCharacterSpacing::from_percent(PERCENT)
}

fn keynote_character_spacing() -> Result<TextCharacterSpacing, litchi_iwa::Error> {
    const PERCENT: f32 = 12.0;
    TextCharacterSpacing::from_percent(PERCENT)
}

const fn numbers_paragraph_line_spacing() -> ParagraphLineSpacing {
    ParagraphLineSpacing::Relative(ParagraphLineSpacingMultiple::ONE_POINT_FIVE)
}

fn pages_paragraph_line_spacing() -> Result<ParagraphLineSpacing, litchi_iwa::Error> {
    const POINTS: f32 = 24.0;
    Ok(ParagraphLineSpacing::Exactly(
        ParagraphLineSpacingPoints::from_points(POINTS)?,
    ))
}

fn keynote_paragraph_line_spacing() -> Result<ParagraphLineSpacing, litchi_iwa::Error> {
    const POINTS: f32 = 7.0;
    Ok(ParagraphLineSpacing::Between(
        ParagraphLineSpacingPoints::from_points(POINTS)?,
    ))
}

fn paragraph_spacing(
    before_points: f32,
    after_points: f32,
) -> Result<ParagraphSpacing, litchi_iwa::Error> {
    Ok(ParagraphSpacing::new(
        ParagraphSpacingPoints::from_points(before_points)?,
        ParagraphSpacingPoints::from_points(after_points)?,
    ))
}

fn numbers_paragraph_spacing() -> Result<ParagraphSpacing, litchi_iwa::Error> {
    const BEFORE_POINTS: f32 = 6.0;
    const AFTER_POINTS: f32 = 9.0;
    paragraph_spacing(BEFORE_POINTS, AFTER_POINTS)
}

fn pages_paragraph_spacing() -> Result<ParagraphSpacing, litchi_iwa::Error> {
    const BEFORE_POINTS: f32 = 5.0;
    const AFTER_POINTS: f32 = 8.0;
    paragraph_spacing(BEFORE_POINTS, AFTER_POINTS)
}

fn keynote_paragraph_spacing() -> Result<ParagraphSpacing, litchi_iwa::Error> {
    const BEFORE_POINTS: f32 = 4.0;
    const AFTER_POINTS: f32 = 7.0;
    paragraph_spacing(BEFORE_POINTS, AFTER_POINTS)
}

fn paragraph_indents(
    first_line_points: f32,
    left_points: f32,
    right_points: f32,
) -> Result<ParagraphIndents, litchi_iwa::Error> {
    Ok(ParagraphIndents::new(
        ParagraphIndentPoints::from_points(first_line_points)?,
        ParagraphIndentPoints::from_points(left_points)?,
        ParagraphIndentPoints::from_points(right_points)?,
    ))
}

fn numbers_paragraph_indents() -> Result<ParagraphIndents, litchi_iwa::Error> {
    const FIRST_LINE_POINTS: f32 = 4.0;
    const LEFT_POINTS: f32 = 8.0;
    const RIGHT_POINTS: f32 = 6.0;
    paragraph_indents(FIRST_LINE_POINTS, LEFT_POINTS, RIGHT_POINTS)
}

fn pages_paragraph_indents() -> Result<ParagraphIndents, litchi_iwa::Error> {
    const FIRST_LINE_POINTS: f32 = 6.0;
    const LEFT_POINTS: f32 = 12.0;
    const RIGHT_POINTS: f32 = 8.0;
    paragraph_indents(FIRST_LINE_POINTS, LEFT_POINTS, RIGHT_POINTS)
}

fn keynote_paragraph_indents() -> Result<ParagraphIndents, litchi_iwa::Error> {
    const FIRST_LINE_POINTS: f32 = 8.0;
    const LEFT_POINTS: f32 = 16.0;
    const RIGHT_POINTS: f32 = 10.0;
    paragraph_indents(FIRST_LINE_POINTS, LEFT_POINTS, RIGHT_POINTS)
}

fn numbers_paragraph_tab_stops() -> Result<ParagraphTabStops, litchi_iwa::Error> {
    const CENTER_POINTS: f32 = 36.0;
    const DECIMAL_POINTS: f32 = 72.0;
    ParagraphTabStops::new(vec![
        ParagraphTabStop::new(
            ParagraphTabPosition::from_points(CENTER_POINTS)?,
            ParagraphTabAlignment::Center,
        ),
        ParagraphTabStop::new(
            ParagraphTabPosition::from_points(DECIMAL_POINTS)?,
            ParagraphTabAlignment::Decimal,
        )
        .with_leader(ParagraphTabLeader::new(".")?),
    ])
}

fn pages_paragraph_tab_stops() -> Result<ParagraphTabStops, litchi_iwa::Error> {
    const LEFT_POINTS: f32 = 42.0;
    const RIGHT_POINTS: f32 = 78.0;
    ParagraphTabStops::new(vec![
        ParagraphTabStop::new(
            ParagraphTabPosition::from_points(LEFT_POINTS)?,
            ParagraphTabAlignment::Left,
        ),
        ParagraphTabStop::new(
            ParagraphTabPosition::from_points(RIGHT_POINTS)?,
            ParagraphTabAlignment::Right,
        )
        .with_leader(ParagraphTabLeader::new("-")?),
    ])
}

fn keynote_paragraph_tab_stops() -> Result<ParagraphTabStops, litchi_iwa::Error> {
    const CENTER_POINTS: f32 = 48.0;
    const RIGHT_POINTS: f32 = 84.0;
    ParagraphTabStops::new(vec![
        ParagraphTabStop::new(
            ParagraphTabPosition::from_points(CENTER_POINTS)?,
            ParagraphTabAlignment::Center,
        ),
        ParagraphTabStop::new(
            ParagraphTabPosition::from_points(RIGHT_POINTS)?,
            ParagraphTabAlignment::Right,
        ),
    ])
}

fn create_numbers(output: &Path) -> Result<(), Box<dyn std::error::Error>> {
    let mut editor = NumbersDocumentBuilder::new()
        .table_name("Layouts")
        .table_dimensions(3, 3)
        .build()?;
    let table_id = editor.tables()?.remove(0).object_id;
    editor.set_cell(table_id, ROW, COLUMN, CellValue::Text(CELL_TEXT.to_owned()))?;
    editor.set_table_cell_layout(
        table_id,
        ROW,
        COLUMN,
        layout(TableCellVerticalAlignment::Middle)?,
    )?;
    editor.set_table_cell_text_alignment(table_id, ROW, COLUMN, TextAlignment::Center)?;
    editor.set_table_cell_text_style(table_id, ROW, COLUMN, numbers_text_style()?)?;
    editor.set_table_cell_text_font(table_id, ROW, COLUMN, numbers_text_font()?)?;
    editor.set_table_cell_text_color(table_id, ROW, COLUMN, numbers_text_color()?)?;
    editor.set_table_cell_text_decorations(table_id, ROW, COLUMN, numbers_text_decorations())?;
    editor.set_table_cell_text_baseline_shift(table_id, ROW, COLUMN, numbers_baseline_shift()?)?;
    editor.set_table_cell_text_capitalization(
        table_id,
        ROW,
        COLUMN,
        TextCapitalization::AllCaps,
    )?;
    editor.set_table_cell_text_character_spacing(
        table_id,
        ROW,
        COLUMN,
        numbers_character_spacing()?,
    )?;
    editor.set_table_cell_text_ligatures(table_id, ROW, COLUMN, TextLigatures::RequiredOnly)?;
    editor.set_table_cell_text_background(table_id, ROW, COLUMN, numbers_text_background()?)?;
    editor.set_table_cell_text_outline(table_id, ROW, COLUMN, TextOutline::standard())?;
    editor.set_table_cell_text_script(table_id, ROW, COLUMN, TextScript::Superscript)?;
    editor.set_table_cell_text_shadow(table_id, ROW, COLUMN, TextShadow::standard())?;
    editor.set_table_cell_paragraph_line_spacing(
        table_id,
        ROW,
        COLUMN,
        numbers_paragraph_line_spacing(),
    )?;
    editor.set_table_cell_paragraph_spacing(table_id, ROW, COLUMN, numbers_paragraph_spacing()?)?;
    editor.set_table_cell_paragraph_indents(table_id, ROW, COLUMN, numbers_paragraph_indents()?)?;
    editor.set_table_cell_paragraph_tab_stops(
        table_id,
        ROW,
        COLUMN,
        numbers_paragraph_tab_stops()?,
    )?;
    editor.save(output)?;
    Ok(())
}

fn create_pages(output: &Path) -> Result<(), Box<dyn std::error::Error>> {
    let mut editor = PagesDocumentBuilder::new()
        .body_text("Created from scratch with native table-cell text layout.\n")
        .body_table("Layouts", 3, 3)
        .build()?;
    let table_id = editor.tables()?.remove(0).model_object_id;
    editor.set_table_cell(table_id, ROW, COLUMN, CellValue::Text(CELL_TEXT.to_owned()))?;
    editor.set_table_cell_layout(
        table_id,
        ROW,
        COLUMN,
        layout(TableCellVerticalAlignment::Bottom)?,
    )?;
    editor.set_table_cell_text_alignment(table_id, ROW, COLUMN, TextAlignment::Right)?;
    editor.set_table_cell_text_style(table_id, ROW, COLUMN, pages_text_style()?)?;
    editor.set_table_cell_text_font(table_id, ROW, COLUMN, pages_text_font()?)?;
    editor.set_table_cell_text_color(table_id, ROW, COLUMN, pages_text_color()?)?;
    editor.set_table_cell_text_decorations(table_id, ROW, COLUMN, pages_text_decorations())?;
    editor.set_table_cell_text_baseline_shift(table_id, ROW, COLUMN, pages_baseline_shift()?)?;
    editor.set_table_cell_text_capitalization(
        table_id,
        ROW,
        COLUMN,
        TextCapitalization::SmallCaps,
    )?;
    editor.set_table_cell_text_character_spacing(
        table_id,
        ROW,
        COLUMN,
        pages_character_spacing()?,
    )?;
    editor.set_table_cell_text_ligatures(table_id, ROW, COLUMN, TextLigatures::All)?;
    editor.set_table_cell_text_background(table_id, ROW, COLUMN, pages_text_background()?)?;
    editor.set_table_cell_text_outline(table_id, ROW, COLUMN, TextOutline::standard())?;
    editor.set_table_cell_text_script(table_id, ROW, COLUMN, TextScript::Subscript)?;
    editor.set_table_cell_text_shadow(table_id, ROW, COLUMN, TextShadow::standard())?;
    editor.set_table_cell_paragraph_line_spacing(
        table_id,
        ROW,
        COLUMN,
        pages_paragraph_line_spacing()?,
    )?;
    editor.set_table_cell_paragraph_spacing(table_id, ROW, COLUMN, pages_paragraph_spacing()?)?;
    editor.set_table_cell_paragraph_indents(table_id, ROW, COLUMN, pages_paragraph_indents()?)?;
    editor.set_table_cell_paragraph_tab_stops(
        table_id,
        ROW,
        COLUMN,
        pages_paragraph_tab_stops()?,
    )?;
    editor.save(output)?;
    Ok(())
}

fn create_keynote(output: &Path) -> Result<(), Box<dyn std::error::Error>> {
    let mut editor = KeynoteDocumentBuilder::new()
        .title("Native table-cell text layouts")
        .build()?;
    let table = editor.add_slide_table(
        0,
        "Layouts",
        3,
        3,
        DrawablePoint { x: 320.0, y: 360.0 },
        DrawableSize {
            width: 1_280.0,
            height: 480.0,
        },
    )?;
    editor.set_slide_table_cell(
        0,
        table.model_object_id,
        ROW,
        COLUMN,
        CellValue::Text(CELL_TEXT.to_owned()),
    )?;
    editor.set_slide_table_cell_layout(
        0,
        table.model_object_id,
        ROW,
        COLUMN,
        layout(TableCellVerticalAlignment::Middle)?,
    )?;
    editor.set_slide_table_cell_text_alignment(
        0,
        table.model_object_id,
        ROW,
        COLUMN,
        TextAlignment::Justified,
    )?;
    editor.set_slide_table_cell_text_style(
        0,
        table.model_object_id,
        ROW,
        COLUMN,
        keynote_text_style()?,
    )?;
    editor.set_slide_table_cell_text_font(
        0,
        table.model_object_id,
        ROW,
        COLUMN,
        keynote_text_font()?,
    )?;
    editor.set_slide_table_cell_text_color(
        0,
        table.model_object_id,
        ROW,
        COLUMN,
        keynote_text_color()?,
    )?;
    editor.set_slide_table_cell_text_decorations(
        0,
        table.model_object_id,
        ROW,
        COLUMN,
        keynote_text_decorations(),
    )?;
    editor.set_slide_table_cell_text_baseline_shift(
        0,
        table.model_object_id,
        ROW,
        COLUMN,
        keynote_baseline_shift()?,
    )?;
    editor.set_slide_table_cell_text_capitalization(
        0,
        table.model_object_id,
        ROW,
        COLUMN,
        TextCapitalization::AllCaps,
    )?;
    editor.set_slide_table_cell_text_character_spacing(
        0,
        table.model_object_id,
        ROW,
        COLUMN,
        keynote_character_spacing()?,
    )?;
    editor.set_slide_table_cell_text_ligatures(
        0,
        table.model_object_id,
        ROW,
        COLUMN,
        TextLigatures::RequiredOnly,
    )?;
    editor.set_slide_table_cell_text_background(
        0,
        table.model_object_id,
        ROW,
        COLUMN,
        keynote_text_background()?,
    )?;
    editor.set_slide_table_cell_text_outline(
        0,
        table.model_object_id,
        ROW,
        COLUMN,
        TextOutline::standard(),
    )?;
    editor.set_slide_table_cell_text_script(
        0,
        table.model_object_id,
        ROW,
        COLUMN,
        TextScript::Superscript,
    )?;
    editor.set_slide_table_cell_text_shadow(
        0,
        table.model_object_id,
        ROW,
        COLUMN,
        TextShadow::standard(),
    )?;
    editor.set_slide_table_cell_paragraph_line_spacing(
        0,
        table.model_object_id,
        ROW,
        COLUMN,
        keynote_paragraph_line_spacing()?,
    )?;
    editor.set_slide_table_cell_paragraph_spacing(
        0,
        table.model_object_id,
        ROW,
        COLUMN,
        keynote_paragraph_spacing()?,
    )?;
    editor.set_slide_table_cell_paragraph_indents(
        0,
        table.model_object_id,
        ROW,
        COLUMN,
        keynote_paragraph_indents()?,
    )?;
    editor.set_slide_table_cell_paragraph_tab_stops(
        0,
        table.model_object_id,
        ROW,
        COLUMN,
        keynote_paragraph_tab_stops()?,
    )?;
    editor.save(output)?;
    Ok(())
}
