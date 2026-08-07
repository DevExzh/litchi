//! Package-boundary adapters for Numbers snapshot edits.
//!
//! These helpers keep archive traversal, wire updates, and package graph
//! mutations behind the semantic editor facade. They are crate-visible only;
//! consumers operate through the typed [`super::NumbersEditor`] API.

use super::table::cell::Borders;
use super::*;
use crate::text::{Alignment, Indents, LineSpacing, Spacing};
use litchi_iwa_common::shape::stroke::Stroke;
use litchi_iwa_common::table::cell::{BorderSide, layout::Layout};
use litchi_numbers::table::dimension::{Dimension, Size};

pub(crate) fn set_table_cell_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    value: CellValue,
) -> Result<()> {
    model::set_attached_cell_in_package(package, table_id, row, column, value)?;
    formula_cache::refresh_formula_caches_after_cell_write(package, table_id, row, column)?;
    Ok(())
}

pub(crate) fn table_cell_borders_in_package(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Borders> {
    stroke_layers::cell_borders(package, table_id, row, column)
}

pub(crate) fn table_cell_fill_in_package(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<crate::shapes::ShapeFill> {
    cell_fill::cell_fill(package, table_id, row, column)
}

pub(crate) fn table_cell_layout_in_package(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Layout> {
    cell_layout::cell_layout(package, table_id, row, column)
}

pub(crate) fn table_cell_text_alignment_in_package(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Alignment> {
    cell_paragraph_style::alignment(package, table_id, row, column)
}

pub(crate) fn set_table_cell_text_alignment_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    alignment: Alignment,
) -> Result<()> {
    cell_paragraph_style::set_alignment(package, table_id, row, column, alignment)
}

pub(crate) fn reset_table_cell_text_alignment_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    cell_paragraph_style::reset_alignment(package, table_id, row, column)
}

pub(crate) fn table_cell_paragraph_line_spacing_in_package(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<LineSpacing> {
    cell_paragraph_style::line_spacing(package, table_id, row, column)
}

pub(crate) fn set_table_cell_paragraph_line_spacing_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    spacing: LineSpacing,
) -> Result<()> {
    cell_paragraph_style::set_line_spacing(package, table_id, row, column, spacing)
}

pub(crate) fn reset_table_cell_paragraph_line_spacing_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    cell_paragraph_style::reset_line_spacing(package, table_id, row, column)
}

pub(crate) fn table_cell_paragraph_spacing_in_package(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Spacing> {
    cell_paragraph_style::spacing(package, table_id, row, column)
}

pub(crate) fn set_table_cell_paragraph_spacing_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    spacing: Spacing,
) -> Result<()> {
    cell_paragraph_style::set_spacing(package, table_id, row, column, spacing)
}

pub(crate) fn reset_table_cell_paragraph_spacing_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    cell_paragraph_style::reset_spacing(package, table_id, row, column)
}

pub(crate) fn table_cell_paragraph_list_in_package(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<ParagraphList> {
    cell_paragraph_list::paragraph_list(package, table_id, row, column)
}

pub(crate) fn set_table_cell_paragraph_list_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    list: ParagraphList,
) -> Result<()> {
    cell_paragraph_list::set_paragraph_list(package, table_id, row, column, list)
}

pub(crate) fn reset_table_cell_paragraph_list_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    cell_paragraph_list::reset_paragraph_list(package, table_id, row, column)
}

pub(crate) fn table_cell_paragraph_lists_in_package(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Vec<ParagraphListPlacement>> {
    cell_paragraph_list::paragraph_lists(package, table_id, row, column)
}

pub(crate) fn set_table_cell_paragraph_lists_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    placements: &[ParagraphListPlacement],
) -> Result<()> {
    cell_paragraph_list::set_paragraph_lists(package, table_id, row, column, placements)
}

pub(crate) fn table_cell_paragraph_list_levels_in_package(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Vec<ParagraphListLevelPlacement>> {
    cell_paragraph_list::paragraph_list_levels(package, table_id, row, column)
}

pub(crate) fn set_table_cell_paragraph_list_level_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    paragraph: TextPosition,
    level: ParagraphListLevel,
) -> Result<()> {
    cell_paragraph_list::set_paragraph_list_level(package, table_id, row, column, paragraph, level)
}

pub(crate) fn reset_table_cell_paragraph_list_level_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    paragraph: TextPosition,
) -> Result<bool> {
    cell_paragraph_list::reset_paragraph_list_level(package, table_id, row, column, paragraph)
}

pub(crate) fn table_cell_paragraph_list_numbering_in_package(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    paragraph: TextPosition,
) -> Result<ParagraphListNumbering> {
    cell_paragraph_list::paragraph_list_numbering(package, table_id, row, column, paragraph)
}

pub(crate) fn set_table_cell_paragraph_list_numbering_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    paragraph: TextPosition,
    numbering: ParagraphListNumbering,
) -> Result<()> {
    cell_paragraph_list::set_paragraph_list_numbering(
        package, table_id, row, column, paragraph, numbering,
    )
}

pub(crate) fn table_cell_paragraph_list_number_format_in_package(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    paragraph: TextPosition,
) -> Result<ParagraphListNumberFormat> {
    cell_paragraph_list::paragraph_list_number_format(package, table_id, row, column, paragraph)
}

pub(crate) fn set_table_cell_paragraph_list_number_format_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    paragraph: TextPosition,
    format: ParagraphListNumberFormat,
) -> Result<()> {
    cell_paragraph_list::set_paragraph_list_number_format(
        package, table_id, row, column, paragraph, format,
    )
}

pub(crate) fn reset_table_cell_paragraph_list_number_format_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    paragraph: TextPosition,
) -> Result<bool> {
    cell_paragraph_list::reset_paragraph_list_number_format(
        package, table_id, row, column, paragraph,
    )
}

pub(crate) fn table_cell_paragraph_list_number_tiering_in_package(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    paragraph: TextPosition,
) -> Result<ParagraphListNumberTiering> {
    cell_paragraph_list::paragraph_list_number_tiering(package, table_id, row, column, paragraph)
}

pub(crate) fn set_table_cell_paragraph_list_number_tiering_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    paragraph: TextPosition,
    tiering: ParagraphListNumberTiering,
) -> Result<()> {
    cell_paragraph_list::set_paragraph_list_number_tiering(
        package, table_id, row, column, paragraph, tiering,
    )
}

pub(crate) fn reset_table_cell_paragraph_list_number_tiering_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    paragraph: TextPosition,
) -> Result<bool> {
    cell_paragraph_list::reset_paragraph_list_number_tiering(
        package, table_id, row, column, paragraph,
    )
}

pub(crate) fn table_cell_paragraph_list_number_scale_in_package(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    paragraph: TextPosition,
) -> Result<ParagraphListNumberScale> {
    cell_paragraph_list::paragraph_list_number_scale(package, table_id, row, column, paragraph)
}

pub(crate) fn set_table_cell_paragraph_list_number_scale_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    paragraph: TextPosition,
    scale: ParagraphListNumberScale,
) -> Result<()> {
    cell_paragraph_list::set_paragraph_list_number_scale(
        package, table_id, row, column, paragraph, scale,
    )
}

pub(crate) fn reset_table_cell_paragraph_list_number_scale_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    paragraph: TextPosition,
) -> Result<bool> {
    cell_paragraph_list::reset_paragraph_list_number_scale(
        package, table_id, row, column, paragraph,
    )
}

pub(crate) fn table_cell_paragraph_list_bullet_in_package(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    paragraph: TextPosition,
) -> Result<ParagraphListBullet> {
    cell_paragraph_list::paragraph_list_bullet(package, table_id, row, column, paragraph)
}

pub(crate) fn set_table_cell_paragraph_list_bullet_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    paragraph: TextPosition,
    bullet: &ParagraphListBullet,
) -> Result<()> {
    cell_paragraph_list::set_paragraph_list_bullet(
        package, table_id, row, column, paragraph, bullet,
    )
}

pub(crate) fn reset_table_cell_paragraph_list_bullet_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    paragraph: TextPosition,
) -> Result<bool> {
    cell_paragraph_list::reset_paragraph_list_bullet(package, table_id, row, column, paragraph)
}

pub(crate) fn table_cell_paragraph_list_bullet_geometry_in_package(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    paragraph: TextPosition,
) -> Result<ParagraphListBulletGeometry> {
    cell_paragraph_list::paragraph_list_bullet_geometry(package, table_id, row, column, paragraph)
}

pub(crate) fn set_table_cell_paragraph_list_bullet_geometry_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    paragraph: TextPosition,
    geometry: ParagraphListBulletGeometry,
) -> Result<()> {
    cell_paragraph_list::set_paragraph_list_bullet_geometry(
        package, table_id, row, column, paragraph, geometry,
    )
}

pub(crate) fn reset_table_cell_paragraph_list_bullet_geometry_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    paragraph: TextPosition,
) -> Result<bool> {
    cell_paragraph_list::reset_paragraph_list_bullet_geometry(
        package, table_id, row, column, paragraph,
    )
}

pub(crate) fn table_cell_paragraph_list_indentation_in_package(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    paragraph: TextPosition,
) -> Result<ParagraphListIndentation> {
    cell_paragraph_list::paragraph_list_indentation(package, table_id, row, column, paragraph)
}

pub(crate) fn set_table_cell_paragraph_list_indentation_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    paragraph: TextPosition,
    indentation: ParagraphListIndentation,
) -> Result<()> {
    cell_paragraph_list::set_paragraph_list_indentation(
        package,
        table_id,
        row,
        column,
        paragraph,
        indentation,
    )
}

pub(crate) fn reset_table_cell_paragraph_list_indentation_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    paragraph: TextPosition,
) -> Result<bool> {
    cell_paragraph_list::reset_paragraph_list_indentation(package, table_id, row, column, paragraph)
}

pub(crate) fn table_cell_paragraph_list_label_color_in_package(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    paragraph: TextPosition,
) -> Result<ParagraphListLabelColor> {
    cell_paragraph_list::paragraph_list_label_color(package, table_id, row, column, paragraph)
}

pub(crate) fn set_table_cell_paragraph_list_label_color_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    paragraph: TextPosition,
    color: ParagraphListLabelColor,
) -> Result<()> {
    cell_paragraph_list::set_paragraph_list_label_color(
        package, table_id, row, column, paragraph, color,
    )
}

pub(crate) fn reset_table_cell_paragraph_list_label_color_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    paragraph: TextPosition,
) -> Result<bool> {
    cell_paragraph_list::reset_paragraph_list_label_color(package, table_id, row, column, paragraph)
}

pub(crate) fn table_cell_paragraph_indents_in_package(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Indents> {
    cell_paragraph_style::indents(package, table_id, row, column)
}

pub(crate) fn set_table_cell_paragraph_indents_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    indents: Indents,
) -> Result<()> {
    cell_paragraph_style::set_indents(package, table_id, row, column, indents)
}

pub(crate) fn reset_table_cell_paragraph_indents_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    cell_paragraph_style::reset_indents(package, table_id, row, column)
}

pub(crate) fn table_cell_paragraph_tab_stops_in_package(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<ParagraphTabStops> {
    cell_paragraph_style::tab_stops(package, table_id, row, column)
}

pub(crate) fn set_table_cell_paragraph_tab_stops_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    stops: ParagraphTabStops,
) -> Result<()> {
    cell_paragraph_style::set_tab_stops(package, table_id, row, column, stops)
}

pub(crate) fn reset_table_cell_paragraph_tab_stops_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    cell_paragraph_style::reset_tab_stops(package, table_id, row, column)
}

pub(crate) fn table_cell_text_background_in_package(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Background> {
    cell_paragraph_style::background(package, table_id, row, column)
}

pub(crate) fn set_table_cell_text_background_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    background: Background,
) -> Result<()> {
    cell_paragraph_style::set_background(package, table_id, row, column, background)
}

pub(crate) fn reset_table_cell_text_background_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    cell_paragraph_style::reset_background(package, table_id, row, column)
}

pub(crate) fn table_cell_text_baseline_shift_in_package(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<TextBaselineShift> {
    cell_paragraph_style::baseline_shift(package, table_id, row, column)
}

pub(crate) fn set_table_cell_text_baseline_shift_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    shift: TextBaselineShift,
) -> Result<()> {
    cell_paragraph_style::set_baseline_shift(package, table_id, row, column, shift)
}

pub(crate) fn reset_table_cell_text_baseline_shift_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    cell_paragraph_style::reset_baseline_shift(package, table_id, row, column)
}

pub(crate) fn table_cell_text_capitalization_in_package(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<TextCapitalization> {
    cell_paragraph_style::capitalization(package, table_id, row, column)
}

pub(crate) fn set_table_cell_text_capitalization_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    capitalization: TextCapitalization,
) -> Result<()> {
    cell_paragraph_style::set_capitalization(package, table_id, row, column, capitalization)
}

pub(crate) fn reset_table_cell_text_capitalization_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    cell_paragraph_style::reset_capitalization(package, table_id, row, column)
}

pub(crate) fn table_cell_text_character_spacing_in_package(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<TextCharacterSpacing> {
    cell_paragraph_style::character_spacing(package, table_id, row, column)
}

pub(crate) fn set_table_cell_text_character_spacing_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    spacing: TextCharacterSpacing,
) -> Result<()> {
    cell_paragraph_style::set_character_spacing(package, table_id, row, column, spacing)
}

pub(crate) fn reset_table_cell_text_character_spacing_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    cell_paragraph_style::reset_character_spacing(package, table_id, row, column)
}

pub(crate) fn table_cell_text_color_in_package(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<RgbaColor> {
    cell_paragraph_style::text_color(package, table_id, row, column)
}

pub(crate) fn set_table_cell_text_color_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    color: RgbaColor,
) -> Result<()> {
    cell_paragraph_style::set_text_color(package, table_id, row, column, color)
}

pub(crate) fn reset_table_cell_text_color_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    cell_paragraph_style::reset_text_color(package, table_id, row, column)
}

pub(crate) fn table_cell_text_decorations_in_package(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<TextDecorations> {
    cell_paragraph_style::decorations(package, table_id, row, column)
}

pub(crate) fn set_table_cell_text_decorations_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    decorations: TextDecorations,
) -> Result<()> {
    cell_paragraph_style::set_decorations(package, table_id, row, column, decorations)
}

pub(crate) fn reset_table_cell_text_decorations_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    cell_paragraph_style::reset_decorations(package, table_id, row, column)
}

pub(crate) fn table_cell_text_font_in_package(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<TextFont> {
    cell_paragraph_style::font(package, table_id, row, column)
}

pub(crate) fn set_table_cell_text_font_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    font: TextFont,
) -> Result<()> {
    cell_paragraph_style::set_font(package, table_id, row, column, font)
}

pub(crate) fn reset_table_cell_text_font_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    cell_paragraph_style::reset_font(package, table_id, row, column)
}

pub(crate) fn table_cell_text_ligatures_in_package(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<TextLigatures> {
    cell_paragraph_style::ligatures(package, table_id, row, column)
}

pub(crate) fn set_table_cell_text_ligatures_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    ligatures: TextLigatures,
) -> Result<()> {
    cell_paragraph_style::set_ligatures(package, table_id, row, column, ligatures)
}

pub(crate) fn reset_table_cell_text_ligatures_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    cell_paragraph_style::reset_ligatures(package, table_id, row, column)
}

pub(crate) fn table_cell_text_outline_in_package(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Outline> {
    cell_paragraph_style::outline(package, table_id, row, column)
}

pub(crate) fn set_table_cell_text_outline_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    outline: Outline,
) -> Result<()> {
    cell_paragraph_style::set_outline(package, table_id, row, column, outline)
}

pub(crate) fn reset_table_cell_text_outline_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    cell_paragraph_style::reset_outline(package, table_id, row, column)
}

pub(crate) fn table_cell_text_script_in_package(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<TextScript> {
    cell_paragraph_style::script(package, table_id, row, column)
}

pub(crate) fn set_table_cell_text_script_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    script: TextScript,
) -> Result<()> {
    cell_paragraph_style::set_script(package, table_id, row, column, script)
}

pub(crate) fn reset_table_cell_text_script_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    cell_paragraph_style::reset_script(package, table_id, row, column)
}

pub(crate) fn table_cell_text_shadow_in_package(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Shadow> {
    cell_paragraph_style::shadow(package, table_id, row, column)
}

pub(crate) fn set_table_cell_text_shadow_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    shadow: Shadow,
) -> Result<()> {
    cell_paragraph_style::set_shadow(package, table_id, row, column, shadow)
}

pub(crate) fn reset_table_cell_text_shadow_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    cell_paragraph_style::reset_shadow(package, table_id, row, column)
}

pub(crate) fn table_cell_text_style_in_package(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<TextStyle> {
    cell_paragraph_style::text_style(package, table_id, row, column)
}

pub(crate) fn set_table_cell_text_style_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    style: TextStyle,
) -> Result<()> {
    cell_paragraph_style::set_text_style(package, table_id, row, column, style)
}

pub(crate) fn reset_table_cell_text_style_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    cell_paragraph_style::reset_text_style(package, table_id, row, column)
}

pub(crate) fn table_cell_number_format_in_package(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Option<Number>> {
    cell_data_format::cell_number_format(package, table_id, row, column)
}

pub(crate) fn set_table_cell_number_format_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    format: Number,
) -> Result<()> {
    cell_data_format::set_cell_number_format(package, table_id, row, column, format)
}

pub(crate) fn common_table_cell_number_format_in_package(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Option<Number>> {
    table_cell_number_format_in_package(package, table_id, row, column)
}

pub(crate) fn set_common_table_cell_number_format_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    format: Number,
) -> Result<()> {
    set_table_cell_number_format_in_package(package, table_id, row, column, format)
}

pub(crate) fn reset_table_cell_number_format_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    cell_data_format::reset_cell_number_format(package, table_id, row, column)
}

pub(crate) fn table_cell_text_format_in_package(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Option<Text>> {
    cell_data_format::cell_text_format(package, table_id, row, column)
}

pub(crate) fn reset_table_cell_text_format_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    cell_data_format::reset_cell_text_format(package, table_id, row, column)
}

pub(crate) fn table_cell_custom_format_in_package(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Option<Custom>> {
    cell_data_format::cell_custom_format(package, table_id, row, column)
}

pub(crate) fn reset_table_cell_custom_format_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    cell_data_format::reset_cell_custom_format(package, table_id, row, column)
}

pub(crate) fn table_cell_currency_format_in_package(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Option<Currency>> {
    cell_data_format::cell_currency_format(package, table_id, row, column)
}

pub(crate) fn reset_table_cell_currency_format_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    cell_data_format::reset_cell_currency_format(package, table_id, row, column)
}

pub(crate) fn table_cell_data_format_in_package(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<DataFormat> {
    cell_data_format::cell_data_format(package, table_id, row, column)
}

pub(crate) fn set_table_cell_data_format_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    format: &DataFormat,
) -> Result<()> {
    cell_data_format::set_cell_data_format(package, table_id, row, column, format)
}

pub(crate) fn table_cell_percentage_format_in_package(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Option<Percentage>> {
    cell_data_format::cell_percentage_format(package, table_id, row, column)
}

pub(crate) fn reset_table_cell_percentage_format_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    cell_data_format::reset_cell_percentage_format(package, table_id, row, column)
}

pub(crate) fn table_cell_scientific_format_in_package(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Option<Scientific>> {
    cell_data_format::cell_scientific_format(package, table_id, row, column)
}

pub(crate) fn reset_table_cell_scientific_format_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    cell_data_format::reset_cell_scientific_format(package, table_id, row, column)
}

pub(crate) fn table_cell_fraction_format_in_package(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Option<Fraction>> {
    cell_data_format::cell_fraction_format(package, table_id, row, column)
}

pub(crate) fn reset_table_cell_fraction_format_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    cell_data_format::reset_cell_fraction_format(package, table_id, row, column)
}

pub(crate) fn table_cell_numeral_system_format_in_package(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Option<NumeralSystem>> {
    cell_data_format::cell_numeral_system_format(package, table_id, row, column)
}

pub(crate) fn reset_table_cell_numeral_system_format_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    cell_data_format::reset_cell_numeral_system_format(package, table_id, row, column)
}

pub(crate) fn table_cell_date_time_format_in_package(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Option<DateTime>> {
    cell_data_format::cell_date_time_format(package, table_id, row, column)
}

pub(crate) fn reset_table_cell_date_time_format_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    cell_data_format::reset_cell_date_time_format(package, table_id, row, column)
}

pub(crate) fn table_cell_duration_format_in_package(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Option<Duration>> {
    cell_data_format::cell_duration_format(package, table_id, row, column)
}

pub(crate) fn reset_table_cell_duration_format_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    cell_data_format::reset_cell_duration_format(package, table_id, row, column)
}

pub(crate) fn table_cell_checkbox_format_in_package(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Option<Checkbox>> {
    cell_data_format::cell_checkbox_format(package, table_id, row, column)
}

pub(crate) fn reset_table_cell_checkbox_format_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    cell_data_format::reset_cell_checkbox_format(package, table_id, row, column)
}

pub(crate) fn table_cell_star_rating_format_in_package(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Option<StarRating>> {
    cell_data_format::cell_star_rating_format(package, table_id, row, column)
}

pub(crate) fn reset_table_cell_star_rating_format_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    cell_data_format::reset_cell_star_rating_format(package, table_id, row, column)
}

pub(crate) fn table_cell_slider_format_in_package(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Option<Slider>> {
    cell_data_format::cell_slider_format(package, table_id, row, column)
}

pub(crate) fn reset_table_cell_slider_format_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    cell_data_format::reset_cell_slider_format(package, table_id, row, column)
}

pub(crate) fn table_cell_stepper_format_in_package(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Option<Stepper>> {
    cell_data_format::cell_stepper_format(package, table_id, row, column)
}

pub(crate) fn reset_table_cell_stepper_format_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    cell_data_format::reset_cell_stepper_format(package, table_id, row, column)
}

pub(crate) fn table_cell_pop_up_menu_format_in_package(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Option<PopUpMenu>> {
    cell_data_format::cell_pop_up_menu_format(package, table_id, row, column)
}

pub(crate) fn reset_table_cell_pop_up_menu_format_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    cell_data_format::reset_cell_pop_up_menu_format(package, table_id, row, column)
}

pub(crate) fn set_table_cell_layout_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    layout: Layout,
) -> Result<()> {
    cell_layout::set_cell_layout(package, table_id, row, column, layout)
}

pub(crate) fn reset_table_cell_layout_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    cell_layout::reset_cell_layout(package, table_id, row, column)
}

pub(crate) fn set_table_cell_fill_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    fill: &crate::shapes::ShapeFill,
) -> Result<()> {
    cell_fill::set_cell_fill(package, table_id, row, column, fill)
}

pub(crate) fn reset_table_cell_fill_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    cell_fill::reset_cell_fill(package, table_id, row, column)
}

pub(crate) fn set_table_cell_border_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    side: BorderSide,
    stroke: Option<Stroke>,
) -> Result<()> {
    stroke_layers::set_cell_border(package, table_id, row, column, side, stroke)
}

pub(crate) fn table_cell_merges_in_package(
    package: &IWorkPackage,
    table_id: u64,
) -> Result<Vec<Region>> {
    cell_merge::regions_in_package(package, table_id)
}

pub(crate) fn merge_table_cells_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    region: Region,
) -> Result<()> {
    cell_merge::merge_in_package(package, table_id, region)
}

pub(crate) fn unmerge_table_cells_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    region: Region,
) -> Result<bool> {
    cell_merge::unmerge_in_package(package, table_id, region)
}

pub(crate) fn table_cell_comment_in_package(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Option<TableCellComment>> {
    model::attached_cell_comment_in_package(package, table_id, row, column)
}

pub(crate) fn table_cell_conditional_highlighting_in_package(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Option<TableCellConditionalHighlightInfo>> {
    conditional_highlight::attached_info_in_package(package, table_id, row, column)
}

pub(crate) fn table_cell_conditional_highlight_rules_in_package(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Option<Vec<Rule>>> {
    conditional_highlight::attached_rules_in_package(package, table_id, row, column)
}

pub(crate) fn clear_table_cell_conditional_highlighting_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<()> {
    conditional_highlight::clear_attached_in_package(package, table_id, row, column)
}

pub(crate) fn set_table_cell_conditional_highlighting_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    rules: &[Rule],
) -> Result<()> {
    conditional_highlight::set_attached_in_package(package, table_id, row, column, rules)
}

pub(crate) fn set_table_cell_comment_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    text: String,
) -> Result<()> {
    model::set_attached_cell_comment_in_package(package, table_id, row, column, text)
}

pub(crate) fn clear_table_cell_comment_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<()> {
    model::clear_attached_cell_comment_in_package(package, table_id, row, column)
}

pub(crate) fn table_cell_comment_replies_in_package(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Vec<TableCellReply>> {
    model::attached_cell_comment_replies_in_package(package, table_id, row, column)
}

pub(crate) fn add_table_cell_comment_reply_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    text: String,
) -> Result<u64> {
    model::add_attached_cell_comment_reply_in_package(package, table_id, row, column, text)
}

pub(crate) fn set_table_cell_comment_reply_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    reply_storage_object_id: u64,
    text: String,
) -> Result<u64> {
    model::set_attached_cell_comment_reply_in_package(
        package,
        table_id,
        row,
        column,
        reply_storage_object_id,
        text,
    )
}

pub(crate) fn remove_table_cell_comment_reply_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    reply_storage_object_id: u64,
) -> Result<()> {
    model::remove_attached_cell_comment_reply_in_package(
        package,
        table_id,
        row,
        column,
        reply_storage_object_id,
    )
}

pub(crate) fn set_table_formula_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    expression: FormulaExpression,
    cached_value: FormulaCachedValue,
) -> Result<()> {
    table_formula::set_attached_table_formula(
        package,
        table_id,
        row,
        column,
        expression,
        Some(cached_value),
    )
}

pub(crate) fn rename_table_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    name: &str,
) -> Result<()> {
    model::rename_attached_table_in_package(package, table_id, name)
}

pub(crate) fn resize_table_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    rows: usize,
    columns: usize,
) -> Result<()> {
    model::resize_attached_table_in_package(package, table_id, rows, columns)
}

pub(crate) fn table_dimensions_in_package(
    package: &IWorkPackage,
    table_id: u64,
) -> Result<(usize, usize)> {
    let descriptor = model::attached_table_descriptor(package, table_id)?;
    Ok((
        descriptor.model.number_of_rows as usize,
        descriptor.model.number_of_columns as usize,
    ))
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum TableTopologyMutation {
    InsertRow(RowInsertion),
    InsertColumn(ColumnInsertion),
    RemoveRow(RowDeletion),
    RemoveColumn(ColumnDeletion),
}

impl TableTopologyMutation {
    pub(crate) fn apply(self, package: &mut IWorkPackage, table_id: u64) -> Result<(usize, usize)> {
        let dimensions = table_dimensions_in_package(package, table_id)?;
        match self {
            Self::InsertRow(row) => Ok((
                insert_table_row_in_package(package, table_id, row)?,
                dimensions.1,
            )),
            Self::InsertColumn(column) => Ok((
                dimensions.0,
                insert_table_column_in_package(package, table_id, column)?,
            )),
            Self::RemoveRow(row) => remove_table_row_in_package(package, table_id, row),
            Self::RemoveColumn(column) => remove_table_column_in_package(package, table_id, column),
        }
    }
}

pub(crate) fn insert_table_row_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    insertion: RowInsertion,
) -> Result<usize> {
    row_insert::insert_attached_table_row(package, table_id, insertion)
}

pub(crate) fn insert_table_column_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    insertion: ColumnInsertion,
) -> Result<usize> {
    column_insert::insert_attached_table_column(package, table_id, insertion)
}

pub(crate) fn remove_table_row_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    deletion: RowDeletion,
) -> Result<(usize, usize)> {
    table_delete::remove_attached_table_row(package, table_id, deletion)
}

pub(crate) fn remove_table_column_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    deletion: ColumnDeletion,
) -> Result<(usize, usize)> {
    table_delete::remove_attached_table_column(package, table_id, deletion)
}

pub(crate) fn set_table_dimension_size_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    dimension: Dimension,
    size: Size,
) -> Result<()> {
    table_dimension::set_attached_table_dimension_size(package, table_id, dimension, size)
}

pub(crate) fn table_dimension_size_in_package(
    package: &IWorkPackage,
    table_id: u64,
    dimension: Dimension,
) -> Result<Size> {
    table_dimension::read_attached_table_dimension_size(package, table_id, dimension)
}

pub(crate) fn table_size_points_in_package(
    package: &IWorkPackage,
    table_id: u64,
) -> Result<(f32, f32)> {
    table_dimension::attached_table_size_points(package, table_id)
}

pub(crate) fn table_header_settings_in_package(
    package: &IWorkPackage,
    table_id: u64,
) -> Result<HeaderSettings> {
    table_headers::read_attached_table_header_settings(package, table_id)
}

pub(crate) fn set_table_header_settings_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    settings: HeaderSettings,
) -> Result<()> {
    table_headers::set_attached_table_header_settings(package, table_id, settings)
}

pub(crate) fn table_owned_object_ids_in_package(
    package: &IWorkPackage,
    table_id: u64,
) -> Result<Vec<u64>> {
    let descriptor = model::attached_table_descriptor(package, table_id)?;
    let locations = object_locations(package)?;
    Ok(table_owned_graph(package, &locations, &descriptor.model)?
        .into_keys()
        .collect())
}

pub(crate) fn remove_table_formula_graph_in_package(
    package: &mut IWorkPackage,
    table_context_ids: &[u64],
) -> Result<Vec<u64>> {
    formula_clone::remove_table_formula_graph_for_contexts(package, table_context_ids)
}

pub(crate) fn create_empty_table_graph_in_package(
    package: &mut IWorkPackage,
    template_info_id: u64,
    template_model_id: u64,
    parent_id: u64,
    name: &str,
    rows: usize,
    columns: usize,
) -> Result<(u64, u64)> {
    let graph = table_create::create_empty_table_graph(
        package,
        template_info_id,
        template_model_id,
        parent_id,
        parent_id,
        name,
        rows,
        columns,
        None,
    )?;
    Ok((graph.info_object_id, graph.model_object_id))
}
