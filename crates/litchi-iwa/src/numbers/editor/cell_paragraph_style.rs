//! Copy-on-write paragraph-style properties for native table cells.

mod property;

use prost::Message;

use super::*;
use crate::text::paragraph_alignment::native::{
    ParagraphStyleOverrides, locate_style, parent_style_id, replace_variation, stylesheet_id,
    variation_object,
};
use crate::text::style_registry::{
    register_private_style, register_style_reference, unregister_owner_reference_if_unused,
    unregister_private_style,
};
use property::{CellParagraphProperty, CellParagraphPropertyKind};

const PARAGRAPH_STYLE_MESSAGE_TYPE: u32 = 2_022;

pub(super) fn style_context(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<(u64, u64)> {
    let descriptor = model::attached_table_descriptor(package, table_id)?;
    validate_coordinate(&descriptor, row, column)?;
    let locations = storage::object_locations(package)?;
    let style_id = local_style_id(package, &descriptor, &locations, row, column)?
        .unwrap_or_else(|| base_style_id(&descriptor, row, column));
    let style = locate_style(package, style_id)?;
    Ok((style_id, stylesheet_id(&style.style, style_id)?))
}

pub(super) fn alignment(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<TextAlignment> {
    match property(
        package,
        table_id,
        row,
        column,
        CellParagraphPropertyKind::Alignment,
    )? {
        CellParagraphProperty::Alignment(value) => Ok(value),
        _ => Err(Error::InvalidFormat(
            "iWork table-cell alignment resolved as another paragraph property".to_owned(),
        )),
    }
}

pub(super) fn set_alignment(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    value: TextAlignment,
) -> Result<()> {
    set_property(
        package,
        table_id,
        row,
        column,
        CellParagraphProperty::Alignment(value),
    )
}

pub(super) fn reset_alignment(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    reset_property(
        package,
        table_id,
        row,
        column,
        CellParagraphPropertyKind::Alignment,
    )
}

pub(super) fn line_spacing(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<ParagraphLineSpacing> {
    match property(
        package,
        table_id,
        row,
        column,
        CellParagraphPropertyKind::LineSpacing,
    )? {
        CellParagraphProperty::LineSpacing(value) => Ok(value),
        _ => Err(Error::InvalidFormat(
            "iWork table-cell line spacing resolved as another paragraph property".to_owned(),
        )),
    }
}

pub(super) fn set_line_spacing(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    value: ParagraphLineSpacing,
) -> Result<()> {
    set_property(
        package,
        table_id,
        row,
        column,
        CellParagraphProperty::LineSpacing(value),
    )
}

pub(super) fn reset_line_spacing(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    reset_property(
        package,
        table_id,
        row,
        column,
        CellParagraphPropertyKind::LineSpacing,
    )
}

pub(super) fn spacing(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<ParagraphSpacing> {
    match property(
        package,
        table_id,
        row,
        column,
        CellParagraphPropertyKind::Spacing,
    )? {
        CellParagraphProperty::Spacing(value) => Ok(value),
        _ => Err(Error::InvalidFormat(
            "iWork table-cell paragraph spacing resolved as another paragraph property".to_owned(),
        )),
    }
}

pub(super) fn set_spacing(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    value: ParagraphSpacing,
) -> Result<()> {
    set_property(
        package,
        table_id,
        row,
        column,
        CellParagraphProperty::Spacing(value),
    )
}

pub(super) fn reset_spacing(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    reset_property(
        package,
        table_id,
        row,
        column,
        CellParagraphPropertyKind::Spacing,
    )
}

pub(super) fn indents(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<ParagraphIndents> {
    match property(
        package,
        table_id,
        row,
        column,
        CellParagraphPropertyKind::Indents,
    )? {
        CellParagraphProperty::Indents(value) => Ok(value),
        _ => Err(Error::InvalidFormat(
            "iWork table-cell paragraph indents resolved as another paragraph property".to_owned(),
        )),
    }
}

pub(super) fn set_indents(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    value: ParagraphIndents,
) -> Result<()> {
    set_property(
        package,
        table_id,
        row,
        column,
        CellParagraphProperty::Indents(value),
    )
}

pub(super) fn reset_indents(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    reset_property(
        package,
        table_id,
        row,
        column,
        CellParagraphPropertyKind::Indents,
    )
}

pub(super) fn tab_stops(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<ParagraphTabStops> {
    match property(
        package,
        table_id,
        row,
        column,
        CellParagraphPropertyKind::TabStops,
    )? {
        CellParagraphProperty::TabStops(value) => Ok(value),
        _ => Err(Error::InvalidFormat(
            "iWork table-cell paragraph tab stops resolved as another paragraph property"
                .to_owned(),
        )),
    }
}

pub(super) fn set_tab_stops(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    value: ParagraphTabStops,
) -> Result<()> {
    set_property(
        package,
        table_id,
        row,
        column,
        CellParagraphProperty::TabStops(value),
    )
}

pub(super) fn reset_tab_stops(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    reset_property(
        package,
        table_id,
        row,
        column,
        CellParagraphPropertyKind::TabStops,
    )
}

pub(super) fn background(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<TextBackground> {
    match property(
        package,
        table_id,
        row,
        column,
        CellParagraphPropertyKind::Background,
    )? {
        CellParagraphProperty::Background(value) => Ok(value),
        _ => Err(Error::InvalidFormat(
            "iWork table-cell text background resolved as another paragraph property".to_owned(),
        )),
    }
}

pub(super) fn set_background(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    value: TextBackground,
) -> Result<()> {
    set_property(
        package,
        table_id,
        row,
        column,
        CellParagraphProperty::Background(value),
    )
}

pub(super) fn reset_background(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    reset_property(
        package,
        table_id,
        row,
        column,
        CellParagraphPropertyKind::Background,
    )
}

pub(super) fn baseline_shift(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<TextBaselineShift> {
    match property(
        package,
        table_id,
        row,
        column,
        CellParagraphPropertyKind::BaselineShift,
    )? {
        CellParagraphProperty::BaselineShift(value) => Ok(value),
        _ => Err(Error::InvalidFormat(
            "iWork table-cell baseline shift resolved as another paragraph property".to_owned(),
        )),
    }
}

pub(super) fn set_baseline_shift(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    value: TextBaselineShift,
) -> Result<()> {
    set_property(
        package,
        table_id,
        row,
        column,
        CellParagraphProperty::BaselineShift(value),
    )
}

pub(super) fn reset_baseline_shift(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    reset_property(
        package,
        table_id,
        row,
        column,
        CellParagraphPropertyKind::BaselineShift,
    )
}

pub(super) fn capitalization(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<TextCapitalization> {
    match property(
        package,
        table_id,
        row,
        column,
        CellParagraphPropertyKind::Capitalization,
    )? {
        CellParagraphProperty::Capitalization(value) => Ok(value),
        _ => Err(Error::InvalidFormat(
            "iWork table-cell capitalization resolved as another paragraph property".to_owned(),
        )),
    }
}

pub(super) fn set_capitalization(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    value: TextCapitalization,
) -> Result<()> {
    set_property(
        package,
        table_id,
        row,
        column,
        CellParagraphProperty::Capitalization(value),
    )
}

pub(super) fn reset_capitalization(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    reset_property(
        package,
        table_id,
        row,
        column,
        CellParagraphPropertyKind::Capitalization,
    )
}

pub(super) fn character_spacing(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<TextCharacterSpacing> {
    match property(
        package,
        table_id,
        row,
        column,
        CellParagraphPropertyKind::CharacterSpacing,
    )? {
        CellParagraphProperty::CharacterSpacing(value) => Ok(value),
        _ => Err(Error::InvalidFormat(
            "iWork table-cell character spacing resolved as another paragraph property".to_owned(),
        )),
    }
}

pub(super) fn set_character_spacing(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    value: TextCharacterSpacing,
) -> Result<()> {
    set_property(
        package,
        table_id,
        row,
        column,
        CellParagraphProperty::CharacterSpacing(value),
    )
}

pub(super) fn reset_character_spacing(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    reset_property(
        package,
        table_id,
        row,
        column,
        CellParagraphPropertyKind::CharacterSpacing,
    )
}

pub(super) fn text_color(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<RgbaColor> {
    match property(
        package,
        table_id,
        row,
        column,
        CellParagraphPropertyKind::Color,
    )? {
        CellParagraphProperty::Color(value) => Ok(value),
        _ => Err(Error::InvalidFormat(
            "iWork table-cell text color resolved as another paragraph property".to_owned(),
        )),
    }
}

pub(super) fn set_text_color(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    value: RgbaColor,
) -> Result<()> {
    set_property(
        package,
        table_id,
        row,
        column,
        CellParagraphProperty::Color(value),
    )
}

pub(super) fn reset_text_color(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    reset_property(
        package,
        table_id,
        row,
        column,
        CellParagraphPropertyKind::Color,
    )
}

pub(super) fn decorations(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<TextDecorations> {
    match property(
        package,
        table_id,
        row,
        column,
        CellParagraphPropertyKind::Decorations,
    )? {
        CellParagraphProperty::Decorations(value) => Ok(value),
        _ => Err(Error::InvalidFormat(
            "iWork table-cell text decorations resolved as another paragraph property".to_owned(),
        )),
    }
}

pub(super) fn set_decorations(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    value: TextDecorations,
) -> Result<()> {
    set_property(
        package,
        table_id,
        row,
        column,
        CellParagraphProperty::Decorations(value),
    )
}

pub(super) fn reset_decorations(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    reset_property(
        package,
        table_id,
        row,
        column,
        CellParagraphPropertyKind::Decorations,
    )
}

pub(super) fn font(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<TextFont> {
    match property(
        package,
        table_id,
        row,
        column,
        CellParagraphPropertyKind::Font,
    )? {
        CellParagraphProperty::Font(value) => Ok(value),
        _ => Err(Error::InvalidFormat(
            "iWork table-cell font resolved as another paragraph property".to_owned(),
        )),
    }
}

pub(super) fn set_font(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    value: TextFont,
) -> Result<()> {
    set_property(
        package,
        table_id,
        row,
        column,
        CellParagraphProperty::Font(value),
    )
}

pub(super) fn reset_font(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    reset_property(
        package,
        table_id,
        row,
        column,
        CellParagraphPropertyKind::Font,
    )
}

pub(super) fn ligatures(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<TextLigatures> {
    match property(
        package,
        table_id,
        row,
        column,
        CellParagraphPropertyKind::Ligatures,
    )? {
        CellParagraphProperty::Ligatures(value) => Ok(value),
        _ => Err(Error::InvalidFormat(
            "iWork table-cell ligatures resolved as another paragraph property".to_owned(),
        )),
    }
}

pub(super) fn set_ligatures(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    value: TextLigatures,
) -> Result<()> {
    set_property(
        package,
        table_id,
        row,
        column,
        CellParagraphProperty::Ligatures(value),
    )
}

pub(super) fn reset_ligatures(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    reset_property(
        package,
        table_id,
        row,
        column,
        CellParagraphPropertyKind::Ligatures,
    )
}

pub(super) fn outline(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<TextOutline> {
    match property(
        package,
        table_id,
        row,
        column,
        CellParagraphPropertyKind::Outline,
    )? {
        CellParagraphProperty::Outline(value) => Ok(value),
        _ => Err(Error::InvalidFormat(
            "iWork table-cell text outline resolved as another paragraph property".to_owned(),
        )),
    }
}

pub(super) fn set_outline(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    value: TextOutline,
) -> Result<()> {
    set_property(
        package,
        table_id,
        row,
        column,
        CellParagraphProperty::Outline(value),
    )
}

pub(super) fn reset_outline(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    reset_property(
        package,
        table_id,
        row,
        column,
        CellParagraphPropertyKind::Outline,
    )
}

pub(super) fn script(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<TextScript> {
    match property(
        package,
        table_id,
        row,
        column,
        CellParagraphPropertyKind::Script,
    )? {
        CellParagraphProperty::Script(value) => Ok(value),
        _ => Err(Error::InvalidFormat(
            "iWork table-cell baseline script resolved as another paragraph property".to_owned(),
        )),
    }
}

pub(super) fn set_script(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    value: TextScript,
) -> Result<()> {
    set_property(
        package,
        table_id,
        row,
        column,
        CellParagraphProperty::Script(value),
    )
}

pub(super) fn reset_script(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    reset_property(
        package,
        table_id,
        row,
        column,
        CellParagraphPropertyKind::Script,
    )
}

pub(super) fn shadow(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<TextShadow> {
    match property(
        package,
        table_id,
        row,
        column,
        CellParagraphPropertyKind::Shadow,
    )? {
        CellParagraphProperty::Shadow(value) => Ok(value),
        _ => Err(Error::InvalidFormat(
            "iWork table-cell text shadow resolved as another paragraph property".to_owned(),
        )),
    }
}

pub(super) fn set_shadow(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    value: TextShadow,
) -> Result<()> {
    set_property(
        package,
        table_id,
        row,
        column,
        CellParagraphProperty::Shadow(value),
    )
}

pub(super) fn reset_shadow(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    reset_property(
        package,
        table_id,
        row,
        column,
        CellParagraphPropertyKind::Shadow,
    )
}

pub(super) fn text_style(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<TextStyle> {
    match property(
        package,
        table_id,
        row,
        column,
        CellParagraphPropertyKind::TextStyle,
    )? {
        CellParagraphProperty::TextStyle(value) => Ok(value),
        _ => Err(Error::InvalidFormat(
            "iWork table-cell character formatting resolved as another paragraph property"
                .to_owned(),
        )),
    }
}

pub(super) fn set_text_style(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    value: TextStyle,
) -> Result<()> {
    set_property(
        package,
        table_id,
        row,
        column,
        CellParagraphProperty::TextStyle(value),
    )
}

pub(super) fn reset_text_style(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    reset_property(
        package,
        table_id,
        row,
        column,
        CellParagraphPropertyKind::TextStyle,
    )
}

fn property(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    kind: CellParagraphPropertyKind,
) -> Result<CellParagraphProperty> {
    let descriptor = model::attached_table_descriptor(package, table_id)?;
    validate_coordinate(&descriptor, row, column)?;
    let locations = storage::object_locations(package)?;
    let style_id = local_style_id(package, &descriptor, &locations, row, column)?
        .unwrap_or_else(|| base_style_id(&descriptor, row, column));
    CellParagraphProperty::inherited(package, style_id, kind)
}

fn set_property(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    value: CellParagraphProperty,
) -> Result<()> {
    let kind = value.kind();
    if property(package, table_id, row, column, kind)? == value {
        return Ok(());
    }
    let mut staged = package.clone();
    table_sparse_storage::ensure_attached_cell_storage(&mut staged, table_id, row, column)?;
    let location = model::locate_attached_cell(&staged, table_id, row, column)?;
    let locations = location.object_locations.clone();
    let style_table_id = location
        .descriptor
        .model
        .base_data_store
        .style_table
        .identifier;
    let old_key = read_bnc(&staged, &location, column)?.text_style_identifier();
    let resolved = storage::resolve_table_data_list(
        &staged,
        &locations,
        style_table_id,
        tst::table_data_list::ListType::Style,
    )?;
    let old_entry = style_entry(&resolved, old_key)?;
    let current_style_id = old_entry
        .and_then(|entry| entry.entry.reference.as_ref())
        .map_or_else(
            || base_style_id(&location.descriptor, row, column),
            |reference| reference.identifier,
        );
    let current_style = locate_style(&staged, current_style_id)?;
    let current_value = CellParagraphProperty::inherited(&staged, current_style_id, kind)?;

    if let Some(entry) = old_entry
        && entry.entry.refcount == 1
        && style_is_exclusive_to_list(
            &staged,
            current_style_id,
            stylesheet_id(&current_style.style, current_style_id)?,
            entry_owner_id(&resolved, entry),
        )?
        && let Some(mut overrides) = crate::text::paragraph_alignment::native::direct_overrides(
            &current_style.style,
            &current_style.message.data,
        )?
    {
        let parent_id = parent_style_id(&current_style.style, current_style_id)?;
        let inherited = CellParagraphProperty::inherited(&staged, parent_id, kind)?;
        if value == inherited {
            drop(staged);
            if !reset_property(package, table_id, row, column, kind)? {
                return Err(Error::InvalidFormat(format!(
                    "iWork table-cell {} could not restore its inherited value",
                    kind.name()
                )));
            }
            return Ok(());
        }
        value.apply_to(&mut overrides, &inherited)?;
        let replacement = variation_object(
            current_style_id,
            parent_id,
            stylesheet_id(&current_style.style, current_style_id)?,
            overrides,
        )?;
        replace_variation(&mut staged, &current_style, replacement)?;
        verify_property(&staged, table_id, row, column, &value)?;
        *package = staged;
        return Ok(());
    }

    let new_style_id = next_object_identifier(&staged)?;
    let stylesheet_id = current_style
        .style
        .super_
        .stylesheet
        .as_ref()
        .map(|reference| reference.identifier)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork paragraph style {current_style_id} has no stylesheet"
            ))
        })?;
    let mut overrides = ParagraphStyleOverrides::default();
    value.apply_to(&mut overrides, &current_value)?;
    let variation = variation_object(new_style_id, current_style_id, stylesheet_id, overrides)?;
    crate::shapes::insert_style_variation(
        &mut staged,
        &current_style.archive_name,
        stylesheet_id,
        current_style_id,
        new_style_id,
        variation,
    )?;
    let new_key = insert_style_entry(&mut staged, &locations, style_table_id, new_style_id)?;
    register_private_style(
        &mut staged,
        &resolved.table_archive,
        &current_style.archive_name,
        new_style_id,
    )?;
    write_text_style_key(&mut staged, &location, row, column, Some(new_key))?;
    if let Some(entry) = old_entry {
        let removed = storage::decrement_table_data_list_entry(
            &mut staged,
            &locations,
            &resolved,
            entry,
            tst::table_data_list::ListType::Style,
        )?;
        if removed {
            unregister_owner_reference_if_unused(
                &mut staged,
                &resolved.table_archive,
                &current_style.archive_name,
                current_style_id,
            )?;
        }
    }
    set_package_last_object_identifier(&mut staged, new_style_id)?;
    verify_property(&staged, table_id, row, column, &value)?;
    *package = staged;
    Ok(())
}

fn reset_property(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    kind: CellParagraphPropertyKind,
) -> Result<bool> {
    let descriptor = model::attached_table_descriptor(package, table_id)?;
    validate_coordinate(&descriptor, row, column)?;
    let locations = storage::object_locations(package)?;
    let Some(key) = text_style_key(package, &descriptor, row, column)? else {
        return Ok(false);
    };
    let style_table_id = descriptor.model.base_data_store.style_table.identifier;
    let resolved = storage::resolve_table_data_list(
        package,
        &locations,
        style_table_id,
        tst::table_data_list::ListType::Style,
    )?;
    let entry = style_entry(&resolved, Some(key))?
        .ok_or_else(|| Error::InvalidFormat(format!("iWork text style table has no key {key}")))?;
    let style_id = entry
        .entry
        .reference
        .as_ref()
        .map(|reference| reference.identifier)
        .ok_or_else(|| {
            Error::InvalidFormat(format!("iWork text style key {key} has no reference"))
        })?;
    let style = locate_style(package, style_id)?;
    let Some(mut overrides) = crate::text::paragraph_alignment::native::direct_overrides(
        &style.style,
        &style.message.data,
    )?
    else {
        return Ok(false);
    };
    if !kind.has_direct(&overrides) {
        return Ok(false);
    }
    let parent_id = parent_style_id(&style.style, style_id)?;
    let inherited = CellParagraphProperty::inherited(package, parent_id, kind)?;

    if entry.entry.refcount == 1
        && style_is_exclusive_to_list(
            package,
            style_id,
            stylesheet_id(&style.style, style_id)?,
            entry_owner_id(&resolved, entry),
        )?
    {
        let mut staged = package.clone();
        kind.clear(&mut overrides);
        if overrides.is_empty() {
            let location = model::locate_attached_cell(&staged, table_id, row, column)?;
            let base_id = base_style_id(&descriptor, row, column);
            if parent_id == base_id {
                write_text_style_key(&mut staged, &location, row, column, None)?;
            } else {
                let parent_key =
                    attach_style_entry(&mut staged, &locations, style_table_id, parent_id)?;
                write_text_style_key(&mut staged, &location, row, column, Some(parent_key))?;
            }
            let removed = storage::decrement_table_data_list_entry(
                &mut staged,
                &locations,
                &resolved,
                entry,
                tst::table_data_list::ListType::Style,
            )?;
            if removed && !style_has_children(&staged, style_id)? {
                crate::shapes::remove_style_variation(
                    &mut staged,
                    &style.archive_name,
                    stylesheet_id(&style.style, style_id)?,
                    parent_id,
                    style_id,
                )?;
                unregister_private_style(
                    &mut staged,
                    &resolved.table_archive,
                    &style.archive_name,
                    style_id,
                    Some(parent_id),
                )?;
                release_package_identifier_suffix(&mut staged, &[style_id])?;
            }
        } else {
            let replacement = variation_object(
                style_id,
                parent_id,
                stylesheet_id(&style.style, style_id)?,
                overrides,
            )?;
            replace_variation(&mut staged, &style, replacement)?;
        }
        verify_property(&staged, table_id, row, column, &inherited)?;
        *package = staged;
        return Ok(true);
    }

    let mut staged = package.clone();
    let location = model::locate_attached_cell(&staged, table_id, row, column)?;
    kind.clear(&mut overrides);
    if overrides.is_empty() {
        let base_id = base_style_id(&descriptor, row, column);
        if parent_id == base_id {
            write_text_style_key(&mut staged, &location, row, column, None)?;
        } else {
            let parent_key =
                attach_style_entry(&mut staged, &locations, style_table_id, parent_id)?;
            write_text_style_key(&mut staged, &location, row, column, Some(parent_key))?;
        }
    } else {
        let new_style_id = next_object_identifier(&staged)?;
        let parent = locate_style(&staged, parent_id)?;
        let stylesheet_id = parent
            .style
            .super_
            .stylesheet
            .as_ref()
            .map(|reference| reference.identifier)
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "iWork paragraph style {parent_id} has no stylesheet"
                ))
            })?;
        let variation = variation_object(new_style_id, parent_id, stylesheet_id, overrides)?;
        crate::shapes::insert_style_variation(
            &mut staged,
            &parent.archive_name,
            stylesheet_id,
            parent_id,
            new_style_id,
            variation,
        )?;
        let new_key = insert_style_entry(&mut staged, &locations, style_table_id, new_style_id)?;
        register_private_style(
            &mut staged,
            &resolved.table_archive,
            &parent.archive_name,
            new_style_id,
        )?;
        write_text_style_key(&mut staged, &location, row, column, Some(new_key))?;
        set_package_last_object_identifier(&mut staged, new_style_id)?;
    }
    let removed = storage::decrement_table_data_list_entry(
        &mut staged,
        &locations,
        &resolved,
        entry,
        tst::table_data_list::ListType::Style,
    )?;
    if removed {
        unregister_owner_reference_if_unused(
            &mut staged,
            &resolved.table_archive,
            &style.archive_name,
            style_id,
        )?;
    }
    verify_property(&staged, table_id, row, column, &inherited)?;
    *package = staged;
    Ok(true)
}

fn validate_coordinate(
    descriptor: &model::TableDescriptor,
    row: usize,
    column: usize,
) -> Result<()> {
    if row >= descriptor.model.number_of_rows as usize
        || column >= descriptor.model.number_of_columns as usize
    {
        return Err(Error::ParseError(format!(
            "Cell ({row}, {column}) is outside iWork table {:?} dimensions {}x{}",
            descriptor.model.table_name,
            descriptor.model.number_of_rows,
            descriptor.model.number_of_columns
        )));
    }
    Ok(())
}

fn base_style_id(descriptor: &model::TableDescriptor, row: usize, column: usize) -> u64 {
    let model = &descriptor.model;
    if row < model.number_of_header_rows.unwrap_or(0) as usize {
        model.header_row_text_style.identifier
    } else if row
        >= model
            .number_of_rows
            .saturating_sub(model.number_of_footer_rows.unwrap_or(0)) as usize
    {
        model.footer_row_text_style.identifier
    } else if column < model.number_of_header_columns.unwrap_or(0) as usize {
        model.header_column_text_style.identifier
    } else {
        model.body_text_style.identifier
    }
}

fn local_style_id(
    package: &IWorkPackage,
    descriptor: &model::TableDescriptor,
    locations: &HashMap<u64, String>,
    row: usize,
    column: usize,
) -> Result<Option<u64>> {
    let Some(key) = text_style_key(package, descriptor, row, column)? else {
        return Ok(None);
    };
    let resolved = storage::resolve_table_data_list(
        package,
        locations,
        descriptor.model.base_data_store.style_table.identifier,
        tst::table_data_list::ListType::Style,
    )?;
    let entry = style_entry(&resolved, Some(key))?
        .ok_or_else(|| Error::InvalidFormat(format!("iWork text style table has no key {key}")))?;
    entry
        .entry
        .reference
        .as_ref()
        .map(|reference| Some(reference.identifier))
        .ok_or_else(|| Error::InvalidFormat(format!("iWork text style key {key} has no reference")))
}

fn text_style_key(
    package: &IWorkPackage,
    descriptor: &model::TableDescriptor,
    row: usize,
    column: usize,
) -> Result<Option<u32>> {
    let location = model::locate_attached_cell(package, descriptor.object_id, row, column)?;
    storage::read_tile_cell(
        package,
        &location.tile_archive,
        location.tile_id,
        location.tile_row,
        column,
    )?
    .map(|data| BncCell::parse(&data).map(|cell| cell.text_style_identifier()))
    .transpose()
    .map(Option::flatten)
    .map_err(Into::into)
}

fn style_entry(
    resolved: &storage::ResolvedTableDataList,
    key: Option<u32>,
) -> Result<Option<&storage::LocatedTableDataListEntry>> {
    let Some(key) = key else {
        return Ok(None);
    };
    let mut matches = resolved
        .entries
        .iter()
        .filter(|entry| entry.entry.key == key);
    let entry = matches.next();
    if matches.next().is_some() {
        return Err(Error::InvalidFormat(format!(
            "iWork text style table repeats key {key}"
        )));
    }
    Ok(entry)
}

fn insert_style_entry(
    package: &mut IWorkPackage,
    locations: &HashMap<u64, String>,
    table_id: u64,
    style_id: u64,
) -> Result<u32> {
    let resolved = storage::resolve_table_data_list(
        package,
        locations,
        table_id,
        tst::table_data_list::ListType::Style,
    )?;
    let key = storage::next_table_data_list_key(&resolved.list, &resolved.entries)?;
    package.update_archive(&resolved.table_archive, |archive| {
        let object = archive.object_mut(table_id).ok_or_else(|| {
            Error::InvalidFormat(format!("iWork style table {table_id} is missing"))
        })?;
        let index = storage::table_data_list_message_index(
            object,
            tst::table_data_list::ListType::Style,
        )
        .ok_or_else(|| Error::InvalidFormat(format!("Object {table_id} has no style list")))?;
        let previous = TableDataList::decode(object.messages[index].data.as_slice())?;
        let mut current = previous.clone();
        current.next_list_id = key
            .checked_add(1)
            .ok_or_else(|| Error::ParseError("iWork text style key overflow".to_owned()))?;
        current.entries.push(tst::table_data_list::ListEntry {
            key,
            refcount: 1,
            reference: Some(tsp::Reference {
                identifier: style_id,
                ..Default::default()
            }),
            ..Default::default()
        });
        let data = storage::rewrite_table_data_list_wire(
            object.messages[index].data.as_slice(),
            &previous,
            &current,
        )?;
        let message_type = object.messages[index].type_;
        object.replace_message(
            index,
            RawMessage {
                type_: message_type,
                data,
            },
        )?;
        storage::add_message_object_reference(object, index, style_id, style_id);
        Ok(())
    })?;
    Ok(key)
}

fn attach_style_entry(
    package: &mut IWorkPackage,
    locations: &HashMap<u64, String>,
    table_id: u64,
    style_id: u64,
) -> Result<u32> {
    let resolved = storage::resolve_table_data_list(
        package,
        locations,
        table_id,
        tst::table_data_list::ListType::Style,
    )?;
    if let Some(entry) = resolved.entries.iter().find(|entry| {
        entry
            .entry
            .reference
            .as_ref()
            .is_some_and(|reference| reference.identifier == style_id)
    }) {
        let key = entry.entry.key;
        storage::increment_table_data_list_entry(
            package,
            locations,
            &resolved,
            entry,
            tst::table_data_list::ListType::Style,
        )?;
        return Ok(key);
    }
    let key = insert_style_entry(package, locations, table_id, style_id)?;
    let style = locate_style(package, style_id)?;
    register_style_reference(
        package,
        &resolved.table_archive,
        &style.archive_name,
        style_id,
    )?;
    Ok(key)
}

fn read_bnc(
    package: &IWorkPackage,
    location: &model::CellLocation,
    column: usize,
) -> Result<BncCell> {
    storage::read_tile_cell(
        package,
        &location.tile_archive,
        location.tile_id,
        location.tile_row,
        column,
    )?
    .map_or_else(|| Ok(BncCell::minimal()), |data| BncCell::parse(&data))
    .map_err(Into::into)
}

fn write_text_style_key(
    package: &mut IWorkPackage,
    location: &model::CellLocation,
    row: usize,
    column: usize,
    key: Option<u32>,
) -> Result<()> {
    let mut cell = read_bnc(package, location, column)?;
    cell.set_text_style_identifier(key);
    storage::set_encoded_cell_value(
        package,
        location.descriptor.object_id,
        row,
        column,
        EncodedValue::Raw(cell.encode()),
    )
}

fn entry_owner_id(
    resolved: &storage::ResolvedTableDataList,
    entry: &storage::LocatedTableDataListEntry,
) -> u64 {
    match &entry.owner {
        storage::TableDataListEntryOwner::Root => resolved.table_id,
        storage::TableDataListEntryOwner::Segment { object_id, .. } => *object_id,
    }
}

fn style_is_exclusive_to_list(
    package: &IWorkPackage,
    style_id: u64,
    stylesheet_id: u64,
    list_owner_id: u64,
) -> Result<bool> {
    for archive_name in package.iwa_entry_names() {
        for object in package.archive(archive_name)?.objects {
            let identifier = object.archive_info.identifier.unwrap_or_default();
            if matches!(identifier, id if id == stylesheet_id || id == list_owner_id || id == style_id)
            {
                continue;
            }
            if object.archive_info.message_infos.iter().any(|info| {
                info.object_references.contains(&style_id)
                    || info
                        .field_infos
                        .iter()
                        .any(|field| field.object_references.contains(&style_id))
            }) {
                return Ok(false);
            }
        }
    }
    Ok(true)
}

fn style_has_children(package: &IWorkPackage, style_id: u64) -> Result<bool> {
    for archive_name in package.iwa_entry_names() {
        if package.archive(archive_name)?.objects.iter().any(|object| {
            object.messages.iter().any(|message| {
                message.type_ == PARAGRAPH_STYLE_MESSAGE_TYPE
                    && tswp::ParagraphStyleArchive::decode(message.data.as_slice()).is_ok_and(
                        |style| {
                            style
                                .super_
                                .parent
                                .as_ref()
                                .is_some_and(|parent| parent.identifier == style_id)
                        },
                    )
            })
        }) {
            return Ok(true);
        }
    }
    Ok(false)
}

fn verify_property(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    expected: &CellParagraphProperty,
) -> Result<()> {
    if property(package, table_id, row, column, expected.kind())? != *expected {
        return Err(Error::InvalidFormat(format!(
            "iWork table-cell {} failed validation",
            expected.kind().name()
        )));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::keynote::KeynoteDocumentBuilder;
    use crate::numbers::{CellValue, NumbersDocumentBuilder};
    use crate::pages::PagesDocumentBuilder;
    use crate::shapes::{DrawablePoint, DrawableSize, RgbColorSpace, RgbaColor};
    use crate::text::{
        ParagraphIndentPoints, ParagraphIndents, ParagraphLineSpacing,
        ParagraphLineSpacingMultiple, ParagraphSpacing, ParagraphSpacingPoints,
        ParagraphTabAlignment, ParagraphTabLeader, ParagraphTabPosition, ParagraphTabStop,
        ParagraphTabStops, TextBackground, TextBaselineShift, TextCapitalization,
        TextCharacterSpacing, TextDecorations, TextFont, TextLigatures, TextOutline, TextPointSize,
        TextScript, TextShadow, TextStrikethrough, TextUnderline,
    };

    fn test_color() -> RgbaColor {
        const RED: f32 = 0.72;
        const GREEN: f32 = 0.10;
        const BLUE: f32 = 0.14;
        const ALPHA: f32 = 1.0;
        RgbaColor::new(RED, GREEN, BLUE, ALPHA, RgbColorSpace::Srgb).unwrap()
    }

    fn test_background() -> TextBackground {
        const RED: f32 = 0.95;
        const GREEN: f32 = 0.82;
        const BLUE: f32 = 0.20;
        const ALPHA: f32 = 1.0;
        TextBackground::Color(RgbaColor::new(RED, GREEN, BLUE, ALPHA, RgbColorSpace::Srgb).unwrap())
    }

    fn test_line_spacing() -> ParagraphLineSpacing {
        ParagraphLineSpacing::Relative(ParagraphLineSpacingMultiple::ONE_POINT_FIVE)
    }

    fn test_paragraph_spacing() -> ParagraphSpacing {
        const BEFORE_POINTS: f32 = 6.0;
        const AFTER_POINTS: f32 = 9.0;
        ParagraphSpacing::new(
            ParagraphSpacingPoints::from_points(BEFORE_POINTS).unwrap(),
            ParagraphSpacingPoints::from_points(AFTER_POINTS).unwrap(),
        )
    }

    fn test_indents() -> ParagraphIndents {
        const FIRST_LINE_POINTS: f32 = 4.0;
        const LEFT_POINTS: f32 = 8.0;
        const RIGHT_POINTS: f32 = 6.0;
        ParagraphIndents::new(
            ParagraphIndentPoints::from_points(FIRST_LINE_POINTS).unwrap(),
            ParagraphIndentPoints::from_points(LEFT_POINTS).unwrap(),
            ParagraphIndentPoints::from_points(RIGHT_POINTS).unwrap(),
        )
    }

    fn test_tab_stops() -> ParagraphTabStops {
        const CENTER_POINTS: f32 = 36.0;
        const DECIMAL_POINTS: f32 = 72.0;
        ParagraphTabStops::new(vec![
            ParagraphTabStop::new(
                ParagraphTabPosition::from_points(CENTER_POINTS).unwrap(),
                ParagraphTabAlignment::Center,
            ),
            ParagraphTabStop::new(
                ParagraphTabPosition::from_points(DECIMAL_POINTS).unwrap(),
                ParagraphTabAlignment::Decimal,
            )
            .with_leader(ParagraphTabLeader::new(".").unwrap()),
        ])
        .unwrap()
    }

    fn explicit_style_id(editor: &NumbersEditor, table_id: u64, row: usize, column: usize) -> u64 {
        let descriptor = model::attached_table_descriptor(&editor.package, table_id).unwrap();
        let locations = storage::object_locations(&editor.package).unwrap();
        local_style_id(&editor.package, &descriptor, &locations, row, column)
            .unwrap()
            .unwrap()
    }

    #[test]
    fn scratch_alignment_reuses_and_reclaims_private_style() {
        let mut editor = NumbersDocumentBuilder::new()
            .table_dimensions(3, 3)
            .build()
            .unwrap();
        let table_id = editor.tables().unwrap()[0].object_id;
        editor
            .set_cell(table_id, 1, 1, CellValue::Text("Aligned".to_owned()))
            .unwrap();
        assert_eq!(
            editor.table_cell_text_alignment(table_id, 1, 1).unwrap(),
            TextAlignment::Natural
        );

        editor
            .set_table_cell_text_alignment(table_id, 1, 1, TextAlignment::Center)
            .unwrap();
        let style_id = explicit_style_id(&editor, table_id, 1, 1);
        let location = model::locate_attached_cell(&editor.package, table_id, 1, 1).unwrap();
        let cell = read_bnc(&editor.package, &location, 1).unwrap();
        assert!(cell.text_style_identifier().is_some());
        assert!(cell.style_identifier().is_none());

        editor
            .set_table_cell_text_alignment(table_id, 1, 1, TextAlignment::Right)
            .unwrap();
        assert_eq!(explicit_style_id(&editor, table_id, 1, 1), style_id);
        assert_eq!(
            editor.table_cell_text_alignment(table_id, 1, 1).unwrap(),
            TextAlignment::Right
        );

        assert!(
            editor
                .reset_table_cell_text_alignment(table_id, 1, 1)
                .unwrap()
        );
        assert_eq!(
            editor.table_cell_text_alignment(table_id, 1, 1).unwrap(),
            TextAlignment::Natural
        );
        let location = model::locate_attached_cell(&editor.package, table_id, 1, 1).unwrap();
        assert!(
            read_bnc(&editor.package, &location, 1)
                .unwrap()
                .text_style_identifier()
                .is_none()
        );
        assert!(editor.package.iwa_entry_names().all(|archive_name| {
            editor
                .package
                .archive(archive_name)
                .unwrap()
                .object(style_id)
                .is_none()
        }));
        assert!(
            !editor
                .reset_table_cell_text_alignment(table_id, 1, 1)
                .unwrap()
        );
    }

    #[test]
    fn shared_alignment_is_copy_on_write_and_reset_is_idempotent() {
        let mut editor = NumbersDocumentBuilder::new()
            .table_dimensions(2, 3)
            .build()
            .unwrap();
        let table_id = editor.tables().unwrap()[0].object_id;
        for column in 1..=2 {
            editor
                .set_cell(
                    table_id,
                    1,
                    column,
                    CellValue::Text(format!("Column {column}")),
                )
                .unwrap();
        }
        editor
            .set_table_cell_text_alignment(table_id, 1, 1, TextAlignment::Center)
            .unwrap();
        let descriptor = model::attached_table_descriptor(&editor.package, table_id).unwrap();
        let locations = storage::object_locations(&editor.package).unwrap();
        let key = text_style_key(&editor.package, &descriptor, 1, 1)
            .unwrap()
            .unwrap();
        let resolved = storage::resolve_table_data_list(
            &editor.package,
            &locations,
            descriptor.model.base_data_store.style_table.identifier,
            tst::table_data_list::ListType::Style,
        )
        .unwrap();
        let entry = style_entry(&resolved, Some(key)).unwrap().unwrap();
        storage::increment_table_data_list_entry(
            &mut editor.package,
            &locations,
            &resolved,
            entry,
            tst::table_data_list::ListType::Style,
        )
        .unwrap();
        let target = model::locate_attached_cell(&editor.package, table_id, 1, 2).unwrap();
        write_text_style_key(&mut editor.package, &target, 1, 2, Some(key)).unwrap();
        let shared_style = explicit_style_id(&editor, table_id, 1, 1);

        editor
            .set_table_cell_text_alignment(table_id, 1, 1, TextAlignment::Right)
            .unwrap();

        assert_ne!(explicit_style_id(&editor, table_id, 1, 1), shared_style);
        assert_eq!(
            editor.table_cell_text_alignment(table_id, 1, 1).unwrap(),
            TextAlignment::Right
        );
        assert_eq!(
            editor.table_cell_text_alignment(table_id, 1, 2).unwrap(),
            TextAlignment::Center
        );

        assert!(
            editor
                .reset_table_cell_text_alignment(table_id, 1, 1)
                .unwrap()
        );
        assert_eq!(
            editor.table_cell_text_alignment(table_id, 1, 1).unwrap(),
            TextAlignment::Center
        );
        assert_eq!(
            editor.table_cell_text_alignment(table_id, 1, 2).unwrap(),
            TextAlignment::Center
        );
        assert!(
            editor
                .reset_table_cell_text_alignment(table_id, 1, 1)
                .unwrap()
        );
        assert_eq!(
            editor.table_cell_text_alignment(table_id, 1, 1).unwrap(),
            TextAlignment::Natural
        );
        assert!(
            !editor
                .reset_table_cell_text_alignment(table_id, 1, 1)
                .unwrap()
        );
    }

    #[test]
    fn paragraph_properties_compose_with_alignment_and_reclaim_independently() {
        let mut editor = NumbersDocumentBuilder::new()
            .table_dimensions(2, 2)
            .build()
            .unwrap();
        let table_id = editor.tables().unwrap()[0].object_id;
        editor
            .set_cell(table_id, 1, 1, CellValue::Text("Styled".to_owned()))
            .unwrap();
        editor
            .set_table_cell_text_alignment(table_id, 1, 1, TextAlignment::Center)
            .unwrap();
        let style_id = explicit_style_id(&editor, table_id, 1, 1);
        let styled = TextStyle::new(TextPointSize::from_points(18.0).unwrap())
            .with_bold(true)
            .with_italic(true);
        let font = TextFont::named("CourierNewPSMT").unwrap();
        let color = test_color();
        let decorations = TextDecorations::new(TextUnderline::Double, TextStrikethrough::Single);
        let baseline_shift = TextBaselineShift::from_points(2.0).unwrap();
        let capitalization = TextCapitalization::TitleCase;
        let character_spacing = TextCharacterSpacing::from_percent(12.0).unwrap();
        let ligatures = TextLigatures::RequiredOnly;
        let background = test_background();
        let outline = TextOutline::standard();
        let script = TextScript::Superscript;
        let shadow = TextShadow::standard();
        let line_spacing = test_line_spacing();
        let paragraph_spacing = test_paragraph_spacing();
        let indents = test_indents();
        let tab_stops = test_tab_stops();

        editor
            .set_table_cell_text_style(table_id, 1, 1, styled)
            .unwrap();
        editor
            .set_table_cell_text_font(table_id, 1, 1, font.clone())
            .unwrap();
        editor
            .set_table_cell_text_color(table_id, 1, 1, color)
            .unwrap();
        editor
            .set_table_cell_text_decorations(table_id, 1, 1, decorations)
            .unwrap();
        editor
            .set_table_cell_text_baseline_shift(table_id, 1, 1, baseline_shift)
            .unwrap();
        editor
            .set_table_cell_text_capitalization(table_id, 1, 1, capitalization)
            .unwrap();
        editor
            .set_table_cell_text_character_spacing(table_id, 1, 1, character_spacing)
            .unwrap();
        editor
            .set_table_cell_text_ligatures(table_id, 1, 1, ligatures)
            .unwrap();
        editor
            .set_table_cell_text_background(table_id, 1, 1, background)
            .unwrap();
        editor
            .set_table_cell_text_outline(table_id, 1, 1, outline)
            .unwrap();
        editor
            .set_table_cell_text_script(table_id, 1, 1, script)
            .unwrap();
        editor
            .set_table_cell_text_shadow(table_id, 1, 1, shadow)
            .unwrap();
        editor
            .set_table_cell_paragraph_line_spacing(table_id, 1, 1, line_spacing)
            .unwrap();
        editor
            .set_table_cell_paragraph_spacing(table_id, 1, 1, paragraph_spacing)
            .unwrap();
        editor
            .set_table_cell_paragraph_indents(table_id, 1, 1, indents)
            .unwrap();
        editor
            .set_table_cell_paragraph_tab_stops(table_id, 1, 1, tab_stops.clone())
            .unwrap();
        assert_eq!(explicit_style_id(&editor, table_id, 1, 1), style_id);
        assert_eq!(
            editor.table_cell_text_style(table_id, 1, 1).unwrap(),
            styled
        );
        assert_eq!(
            editor.table_cell_text_alignment(table_id, 1, 1).unwrap(),
            TextAlignment::Center
        );
        assert_eq!(editor.table_cell_text_font(table_id, 1, 1).unwrap(), font);
        assert_eq!(editor.table_cell_text_color(table_id, 1, 1).unwrap(), color);
        assert_eq!(
            editor.table_cell_text_decorations(table_id, 1, 1).unwrap(),
            decorations
        );
        assert_eq!(
            editor
                .table_cell_text_baseline_shift(table_id, 1, 1)
                .unwrap(),
            baseline_shift
        );
        assert_eq!(
            editor
                .table_cell_text_capitalization(table_id, 1, 1)
                .unwrap(),
            capitalization
        );
        assert_eq!(
            editor
                .table_cell_text_character_spacing(table_id, 1, 1)
                .unwrap(),
            character_spacing
        );
        assert_eq!(
            editor.table_cell_text_ligatures(table_id, 1, 1).unwrap(),
            ligatures
        );
        assert_eq!(
            editor.table_cell_text_background(table_id, 1, 1).unwrap(),
            background
        );
        assert_eq!(
            editor.table_cell_text_outline(table_id, 1, 1).unwrap(),
            outline
        );
        assert_eq!(
            editor.table_cell_text_script(table_id, 1, 1).unwrap(),
            script
        );
        assert_eq!(
            editor.table_cell_text_shadow(table_id, 1, 1).unwrap(),
            shadow
        );
        assert_eq!(
            editor
                .table_cell_paragraph_line_spacing(table_id, 1, 1)
                .unwrap(),
            line_spacing
        );
        assert_eq!(
            editor.table_cell_paragraph_spacing(table_id, 1, 1).unwrap(),
            paragraph_spacing
        );
        assert_eq!(
            editor.table_cell_paragraph_indents(table_id, 1, 1).unwrap(),
            indents
        );
        assert_eq!(
            editor
                .table_cell_paragraph_tab_stops(table_id, 1, 1)
                .unwrap(),
            tab_stops
        );

        assert!(
            editor
                .reset_table_cell_text_capitalization(table_id, 1, 1)
                .unwrap()
        );
        assert_eq!(
            editor
                .table_cell_text_capitalization(table_id, 1, 1)
                .unwrap(),
            TextCapitalization::None
        );
        assert_eq!(
            editor
                .table_cell_text_character_spacing(table_id, 1, 1)
                .unwrap(),
            character_spacing
        );
        assert!(
            editor
                .reset_table_cell_text_character_spacing(table_id, 1, 1)
                .unwrap()
        );
        assert!(editor.reset_table_cell_text_script(table_id, 1, 1).unwrap());
        assert!(
            editor
                .reset_table_cell_text_baseline_shift(table_id, 1, 1)
                .unwrap()
        );
        assert!(
            editor
                .reset_table_cell_text_ligatures(table_id, 1, 1)
                .unwrap()
        );
        assert!(
            editor
                .reset_table_cell_text_background(table_id, 1, 1)
                .unwrap()
        );
        assert_eq!(
            editor.table_cell_text_outline(table_id, 1, 1).unwrap(),
            outline
        );
        assert!(
            editor
                .reset_table_cell_text_outline(table_id, 1, 1)
                .unwrap()
        );
        assert_eq!(
            editor.table_cell_text_shadow(table_id, 1, 1).unwrap(),
            shadow
        );
        assert!(editor.reset_table_cell_text_shadow(table_id, 1, 1).unwrap());
        assert_eq!(
            editor
                .table_cell_paragraph_line_spacing(table_id, 1, 1)
                .unwrap(),
            line_spacing
        );
        assert!(
            editor
                .reset_table_cell_paragraph_spacing(table_id, 1, 1)
                .unwrap()
        );
        assert!(
            editor
                .reset_table_cell_paragraph_line_spacing(table_id, 1, 1)
                .unwrap()
        );
        assert_eq!(
            editor
                .table_cell_paragraph_tab_stops(table_id, 1, 1)
                .unwrap(),
            tab_stops
        );
        assert!(
            editor
                .reset_table_cell_paragraph_indents(table_id, 1, 1)
                .unwrap()
        );
        assert!(
            editor
                .reset_table_cell_paragraph_tab_stops(table_id, 1, 1)
                .unwrap()
        );

        assert!(
            editor
                .reset_table_cell_text_decorations(table_id, 1, 1)
                .unwrap()
        );
        assert_eq!(
            editor.table_cell_text_decorations(table_id, 1, 1).unwrap(),
            TextDecorations::NONE
        );
        assert_eq!(editor.table_cell_text_color(table_id, 1, 1).unwrap(), color);
        assert!(editor.reset_table_cell_text_color(table_id, 1, 1).unwrap());
        assert_eq!(
            editor.table_cell_text_color(table_id, 1, 1).unwrap(),
            RgbaColor::black()
        );
        assert_eq!(editor.table_cell_text_font(table_id, 1, 1).unwrap(), font);

        editor
            .set_table_cell_text_style(table_id, 1, 1, TextStyle::default())
            .unwrap();
        assert_eq!(
            editor.table_cell_text_style(table_id, 1, 1).unwrap(),
            TextStyle::default()
        );
        assert!(!editor.reset_table_cell_text_style(table_id, 1, 1).unwrap());
        editor
            .set_table_cell_text_font(table_id, 1, 1, TextFont::default())
            .unwrap();
        assert_eq!(
            editor.table_cell_text_font(table_id, 1, 1).unwrap(),
            TextFont::default()
        );
        assert!(!editor.reset_table_cell_text_font(table_id, 1, 1).unwrap());
        editor
            .set_table_cell_text_font(table_id, 1, 1, font)
            .unwrap();
        assert!(editor.reset_table_cell_text_font(table_id, 1, 1).unwrap());
        editor
            .set_table_cell_text_style(table_id, 1, 1, styled)
            .unwrap();
        assert!(editor.reset_table_cell_text_style(table_id, 1, 1).unwrap());
        assert_eq!(
            editor.table_cell_text_style(table_id, 1, 1).unwrap(),
            TextStyle::default()
        );
        assert_eq!(explicit_style_id(&editor, table_id, 1, 1), style_id);
        assert!(!editor.reset_table_cell_text_style(table_id, 1, 1).unwrap());
        assert!(
            editor
                .reset_table_cell_text_alignment(table_id, 1, 1)
                .unwrap()
        );
        assert!(editor.package.iwa_entry_names().all(|archive_name| {
            editor
                .package
                .archive(archive_name)
                .unwrap()
                .object(style_id)
                .is_none()
        }));
    }

    #[test]
    fn shared_character_properties_use_copy_on_write() {
        let mut editor = NumbersDocumentBuilder::new()
            .table_dimensions(2, 3)
            .build()
            .unwrap();
        let table_id = editor.tables().unwrap()[0].object_id;
        for column in 1..=2 {
            editor
                .set_cell(
                    table_id,
                    1,
                    column,
                    CellValue::Text(format!("Column {column}")),
                )
                .unwrap();
        }
        editor
            .set_table_cell_text_alignment(table_id, 1, 1, TextAlignment::Center)
            .unwrap();
        let descriptor = model::attached_table_descriptor(&editor.package, table_id).unwrap();
        let locations = storage::object_locations(&editor.package).unwrap();
        let key = text_style_key(&editor.package, &descriptor, 1, 1)
            .unwrap()
            .unwrap();
        let resolved = storage::resolve_table_data_list(
            &editor.package,
            &locations,
            descriptor.model.base_data_store.style_table.identifier,
            tst::table_data_list::ListType::Style,
        )
        .unwrap();
        let entry = style_entry(&resolved, Some(key)).unwrap().unwrap();
        storage::increment_table_data_list_entry(
            &mut editor.package,
            &locations,
            &resolved,
            entry,
            tst::table_data_list::ListType::Style,
        )
        .unwrap();
        let target = model::locate_attached_cell(&editor.package, table_id, 1, 2).unwrap();
        write_text_style_key(&mut editor.package, &target, 1, 2, Some(key)).unwrap();
        let shared_style_id = explicit_style_id(&editor, table_id, 1, 1);
        let font = TextFont::named("CourierNewPSMT").unwrap();
        let color = test_color();
        let decorations = TextDecorations::new(TextUnderline::Single, TextStrikethrough::Double);

        editor
            .set_table_cell_text_capitalization(table_id, 1, 1, TextCapitalization::TitleCase)
            .unwrap();
        assert_ne!(explicit_style_id(&editor, table_id, 1, 1), shared_style_id);
        assert_eq!(
            editor
                .table_cell_text_capitalization(table_id, 1, 1)
                .unwrap(),
            TextCapitalization::TitleCase
        );
        assert_eq!(
            editor
                .table_cell_text_capitalization(table_id, 1, 2)
                .unwrap(),
            TextCapitalization::None
        );
        assert!(
            editor
                .reset_table_cell_text_capitalization(table_id, 1, 1)
                .unwrap()
        );
        assert_eq!(explicit_style_id(&editor, table_id, 1, 1), shared_style_id);

        let line_spacing = test_line_spacing();
        editor
            .set_table_cell_paragraph_line_spacing(table_id, 1, 1, line_spacing)
            .unwrap();
        assert_ne!(explicit_style_id(&editor, table_id, 1, 1), shared_style_id);
        assert_eq!(
            editor
                .table_cell_paragraph_line_spacing(table_id, 1, 1)
                .unwrap(),
            line_spacing
        );
        assert_eq!(
            editor
                .table_cell_paragraph_line_spacing(table_id, 1, 2)
                .unwrap(),
            ParagraphLineSpacing::default()
        );
        assert!(
            editor
                .reset_table_cell_paragraph_line_spacing(table_id, 1, 1)
                .unwrap()
        );
        assert_eq!(explicit_style_id(&editor, table_id, 1, 1), shared_style_id);

        let paragraph_spacing = test_paragraph_spacing();
        editor
            .set_table_cell_paragraph_spacing(table_id, 1, 1, paragraph_spacing)
            .unwrap();
        assert_ne!(explicit_style_id(&editor, table_id, 1, 1), shared_style_id);
        assert_eq!(
            editor.table_cell_paragraph_spacing(table_id, 1, 1).unwrap(),
            paragraph_spacing
        );
        assert_eq!(
            editor.table_cell_paragraph_spacing(table_id, 1, 2).unwrap(),
            ParagraphSpacing::NONE
        );
        assert!(
            editor
                .reset_table_cell_paragraph_spacing(table_id, 1, 1)
                .unwrap()
        );
        assert_eq!(explicit_style_id(&editor, table_id, 1, 1), shared_style_id);

        let indents = test_indents();
        editor
            .set_table_cell_paragraph_indents(table_id, 1, 1, indents)
            .unwrap();
        assert_ne!(explicit_style_id(&editor, table_id, 1, 1), shared_style_id);
        assert_eq!(
            editor.table_cell_paragraph_indents(table_id, 1, 1).unwrap(),
            indents
        );
        assert_eq!(
            editor.table_cell_paragraph_indents(table_id, 1, 2).unwrap(),
            ParagraphIndents::NONE
        );
        assert!(
            editor
                .reset_table_cell_paragraph_indents(table_id, 1, 1)
                .unwrap()
        );
        assert_eq!(explicit_style_id(&editor, table_id, 1, 1), shared_style_id);

        let tab_stops = test_tab_stops();
        editor
            .set_table_cell_paragraph_tab_stops(table_id, 1, 1, tab_stops.clone())
            .unwrap();
        assert_ne!(explicit_style_id(&editor, table_id, 1, 1), shared_style_id);
        assert_eq!(
            editor
                .table_cell_paragraph_tab_stops(table_id, 1, 1)
                .unwrap(),
            tab_stops
        );
        assert!(
            editor
                .table_cell_paragraph_tab_stops(table_id, 1, 2)
                .unwrap()
                .is_empty()
        );
        assert!(
            editor
                .reset_table_cell_paragraph_tab_stops(table_id, 1, 1)
                .unwrap()
        );
        assert_eq!(explicit_style_id(&editor, table_id, 1, 1), shared_style_id);

        let background = test_background();
        editor
            .set_table_cell_text_background(table_id, 1, 1, background)
            .unwrap();
        assert_ne!(explicit_style_id(&editor, table_id, 1, 1), shared_style_id);
        assert_eq!(
            editor.table_cell_text_background(table_id, 1, 1).unwrap(),
            background
        );
        assert_eq!(
            editor.table_cell_text_background(table_id, 1, 2).unwrap(),
            TextBackground::None
        );
        assert!(
            editor
                .reset_table_cell_text_background(table_id, 1, 1)
                .unwrap()
        );
        assert_eq!(explicit_style_id(&editor, table_id, 1, 1), shared_style_id);

        editor
            .set_table_cell_text_decorations(table_id, 1, 1, decorations)
            .unwrap();
        assert_ne!(explicit_style_id(&editor, table_id, 1, 1), shared_style_id);
        assert_eq!(
            editor.table_cell_text_decorations(table_id, 1, 1).unwrap(),
            decorations
        );
        assert_eq!(
            editor.table_cell_text_decorations(table_id, 1, 2).unwrap(),
            TextDecorations::NONE
        );
        assert!(
            editor
                .reset_table_cell_text_decorations(table_id, 1, 1)
                .unwrap()
        );
        assert_eq!(explicit_style_id(&editor, table_id, 1, 1), shared_style_id);

        editor
            .set_table_cell_text_color(table_id, 1, 1, color)
            .unwrap();
        assert_ne!(explicit_style_id(&editor, table_id, 1, 1), shared_style_id);
        assert_eq!(editor.table_cell_text_color(table_id, 1, 1).unwrap(), color);
        assert_eq!(
            editor.table_cell_text_color(table_id, 1, 2).unwrap(),
            RgbaColor::black()
        );
        assert!(editor.reset_table_cell_text_color(table_id, 1, 1).unwrap());
        assert_eq!(explicit_style_id(&editor, table_id, 1, 1), shared_style_id);

        editor
            .set_table_cell_text_font(table_id, 1, 1, font.clone())
            .unwrap();
        assert_ne!(explicit_style_id(&editor, table_id, 1, 1), shared_style_id);
        assert_eq!(editor.table_cell_text_font(table_id, 1, 1).unwrap(), font);
        assert_eq!(
            editor.table_cell_text_font(table_id, 1, 2).unwrap(),
            TextFont::default()
        );
        assert_eq!(
            editor.table_cell_text_alignment(table_id, 1, 2).unwrap(),
            TextAlignment::Center
        );

        assert!(editor.reset_table_cell_text_font(table_id, 1, 1).unwrap());
        assert_eq!(explicit_style_id(&editor, table_id, 1, 1), shared_style_id);
        assert!(!editor.reset_table_cell_text_font(table_id, 1, 1).unwrap());
    }

    #[test]
    fn shared_text_style_uses_copy_on_write() {
        let mut editor = NumbersDocumentBuilder::new()
            .table_dimensions(2, 3)
            .build()
            .unwrap();
        let table_id = editor.tables().unwrap()[0].object_id;
        for column in 1..=2 {
            editor
                .set_cell(
                    table_id,
                    1,
                    column,
                    CellValue::Text(format!("Column {column}")),
                )
                .unwrap();
        }
        editor
            .set_table_cell_text_alignment(table_id, 1, 1, TextAlignment::Center)
            .unwrap();
        let descriptor = model::attached_table_descriptor(&editor.package, table_id).unwrap();
        let locations = storage::object_locations(&editor.package).unwrap();
        let key = text_style_key(&editor.package, &descriptor, 1, 1)
            .unwrap()
            .unwrap();
        let resolved = storage::resolve_table_data_list(
            &editor.package,
            &locations,
            descriptor.model.base_data_store.style_table.identifier,
            tst::table_data_list::ListType::Style,
        )
        .unwrap();
        let entry = style_entry(&resolved, Some(key)).unwrap().unwrap();
        storage::increment_table_data_list_entry(
            &mut editor.package,
            &locations,
            &resolved,
            entry,
            tst::table_data_list::ListType::Style,
        )
        .unwrap();
        let target = model::locate_attached_cell(&editor.package, table_id, 1, 2).unwrap();
        write_text_style_key(&mut editor.package, &target, 1, 2, Some(key)).unwrap();
        let shared_style_id = explicit_style_id(&editor, table_id, 1, 1);
        let styled = TextStyle::new(TextPointSize::from_points(20.0).unwrap()).with_bold(true);

        editor
            .set_table_cell_text_style(table_id, 1, 1, styled)
            .unwrap();
        assert_ne!(explicit_style_id(&editor, table_id, 1, 1), shared_style_id);
        assert_eq!(
            editor.table_cell_text_style(table_id, 1, 1).unwrap(),
            styled
        );
        assert_eq!(
            editor.table_cell_text_style(table_id, 1, 2).unwrap(),
            TextStyle::default()
        );
        assert_eq!(
            editor.table_cell_text_alignment(table_id, 1, 2).unwrap(),
            TextAlignment::Center
        );

        assert!(editor.reset_table_cell_text_style(table_id, 1, 1).unwrap());
        assert_eq!(explicit_style_id(&editor, table_id, 1, 1), shared_style_id);
        assert!(!editor.reset_table_cell_text_style(table_id, 1, 1).unwrap());
    }

    #[test]
    fn scratch_paragraph_styles_round_trip_in_pages_and_keynote() {
        let pages_style =
            TextStyle::new(TextPointSize::from_points(17.0).unwrap()).with_italic(true);
        let pages_font = TextFont::named("AvenirNext-Regular").unwrap();
        let pages_color = test_color();
        let pages_decorations =
            TextDecorations::new(TextUnderline::Single, TextStrikethrough::None);
        let pages_baseline_shift = TextBaselineShift::from_points(-1.5).unwrap();
        let pages_spacing = TextCharacterSpacing::from_percent(8.0).unwrap();
        let pages_background = test_background();
        let pages_outline = TextOutline::standard();
        let pages_shadow = TextShadow::standard();
        let pages_line_spacing = test_line_spacing();
        let pages_paragraph_spacing = test_paragraph_spacing();
        let pages_indents = test_indents();
        let pages_tab_stops = test_tab_stops();
        let mut pages = PagesDocumentBuilder::new()
            .body_table("Aligned", 2, 2)
            .build()
            .unwrap();
        let pages_table = pages.tables().unwrap()[0].model_object_id;
        pages
            .set_table_cell_text_alignment(pages_table, 1, 1, TextAlignment::Justified)
            .unwrap();
        pages
            .set_table_cell_text_style(pages_table, 1, 1, pages_style)
            .unwrap();
        pages
            .set_table_cell_text_font(pages_table, 1, 1, pages_font.clone())
            .unwrap();
        pages
            .set_table_cell_text_color(pages_table, 1, 1, pages_color)
            .unwrap();
        pages
            .set_table_cell_text_decorations(pages_table, 1, 1, pages_decorations)
            .unwrap();
        pages
            .set_table_cell_text_baseline_shift(pages_table, 1, 1, pages_baseline_shift)
            .unwrap();
        pages
            .set_table_cell_text_capitalization(pages_table, 1, 1, TextCapitalization::SmallCaps)
            .unwrap();
        pages
            .set_table_cell_text_character_spacing(pages_table, 1, 1, pages_spacing)
            .unwrap();
        pages
            .set_table_cell_text_ligatures(pages_table, 1, 1, TextLigatures::All)
            .unwrap();
        pages
            .set_table_cell_text_background(pages_table, 1, 1, pages_background)
            .unwrap();
        pages
            .set_table_cell_text_outline(pages_table, 1, 1, pages_outline)
            .unwrap();
        pages
            .set_table_cell_text_script(pages_table, 1, 1, TextScript::Subscript)
            .unwrap();
        pages
            .set_table_cell_text_shadow(pages_table, 1, 1, pages_shadow)
            .unwrap();
        pages
            .set_table_cell_paragraph_line_spacing(pages_table, 1, 1, pages_line_spacing)
            .unwrap();
        pages
            .set_table_cell_paragraph_spacing(pages_table, 1, 1, pages_paragraph_spacing)
            .unwrap();
        pages
            .set_table_cell_paragraph_indents(pages_table, 1, 1, pages_indents)
            .unwrap();
        pages
            .set_table_cell_paragraph_tab_stops(pages_table, 1, 1, pages_tab_stops.clone())
            .unwrap();
        let mut pages = crate::pages::PagesEditor::from_bytes(&pages.to_bytes().unwrap()).unwrap();
        assert_eq!(
            pages.table_cell_text_alignment(pages_table, 1, 1).unwrap(),
            TextAlignment::Justified
        );
        assert_eq!(
            pages.table_cell_text_style(pages_table, 1, 1).unwrap(),
            pages_style
        );
        assert_eq!(
            pages.table_cell_text_font(pages_table, 1, 1).unwrap(),
            pages_font
        );
        assert_eq!(
            pages.table_cell_text_color(pages_table, 1, 1).unwrap(),
            pages_color
        );
        assert_eq!(
            pages
                .table_cell_text_decorations(pages_table, 1, 1)
                .unwrap(),
            pages_decorations
        );
        assert_eq!(
            pages
                .table_cell_text_baseline_shift(pages_table, 1, 1)
                .unwrap(),
            pages_baseline_shift
        );
        assert_eq!(
            pages
                .table_cell_text_capitalization(pages_table, 1, 1)
                .unwrap(),
            TextCapitalization::SmallCaps
        );
        assert_eq!(
            pages
                .table_cell_text_character_spacing(pages_table, 1, 1)
                .unwrap(),
            pages_spacing
        );
        assert_eq!(
            pages.table_cell_text_ligatures(pages_table, 1, 1).unwrap(),
            TextLigatures::All
        );
        assert_eq!(
            pages.table_cell_text_background(pages_table, 1, 1).unwrap(),
            pages_background
        );
        assert_eq!(
            pages.table_cell_text_outline(pages_table, 1, 1).unwrap(),
            pages_outline
        );
        assert_eq!(
            pages.table_cell_text_script(pages_table, 1, 1).unwrap(),
            TextScript::Subscript
        );
        assert_eq!(
            pages.table_cell_text_shadow(pages_table, 1, 1).unwrap(),
            pages_shadow
        );
        assert_eq!(
            pages
                .table_cell_paragraph_line_spacing(pages_table, 1, 1)
                .unwrap(),
            pages_line_spacing
        );
        assert_eq!(
            pages
                .table_cell_paragraph_spacing(pages_table, 1, 1)
                .unwrap(),
            pages_paragraph_spacing
        );
        assert_eq!(
            pages
                .table_cell_paragraph_indents(pages_table, 1, 1)
                .unwrap(),
            pages_indents
        );
        assert_eq!(
            pages
                .table_cell_paragraph_tab_stops(pages_table, 1, 1)
                .unwrap(),
            pages_tab_stops
        );
        assert!(
            pages
                .reset_table_cell_text_capitalization(pages_table, 1, 1)
                .unwrap()
        );
        assert!(
            pages
                .reset_table_cell_text_character_spacing(pages_table, 1, 1)
                .unwrap()
        );
        assert!(
            pages
                .reset_table_cell_text_baseline_shift(pages_table, 1, 1)
                .unwrap()
        );
        assert!(
            pages
                .reset_table_cell_text_ligatures(pages_table, 1, 1)
                .unwrap()
        );
        assert!(
            pages
                .reset_table_cell_text_background(pages_table, 1, 1)
                .unwrap()
        );
        assert!(
            pages
                .reset_table_cell_text_outline(pages_table, 1, 1)
                .unwrap()
        );
        assert!(
            pages
                .reset_table_cell_text_script(pages_table, 1, 1)
                .unwrap()
        );
        assert!(
            pages
                .reset_table_cell_text_shadow(pages_table, 1, 1)
                .unwrap()
        );
        assert!(
            pages
                .reset_table_cell_paragraph_spacing(pages_table, 1, 1)
                .unwrap()
        );
        assert!(
            pages
                .reset_table_cell_paragraph_line_spacing(pages_table, 1, 1)
                .unwrap()
        );
        assert!(
            pages
                .reset_table_cell_paragraph_indents(pages_table, 1, 1)
                .unwrap()
        );
        assert!(
            pages
                .reset_table_cell_paragraph_tab_stops(pages_table, 1, 1)
                .unwrap()
        );
        assert!(
            pages
                .reset_table_cell_text_decorations(pages_table, 1, 1)
                .unwrap()
        );
        assert!(
            pages
                .reset_table_cell_text_color(pages_table, 1, 1)
                .unwrap()
        );
        assert!(pages.reset_table_cell_text_font(pages_table, 1, 1).unwrap());
        assert!(
            pages
                .reset_table_cell_text_style(pages_table, 1, 1)
                .unwrap()
        );

        let keynote_style =
            TextStyle::new(TextPointSize::from_points(19.0).unwrap()).with_bold(true);
        let keynote_font = TextFont::named("Menlo-Regular").unwrap();
        let keynote_color = test_color();
        let keynote_decorations =
            TextDecorations::new(TextUnderline::None, TextStrikethrough::Single);
        let keynote_baseline_shift = TextBaselineShift::from_points(2.0).unwrap();
        let keynote_spacing = TextCharacterSpacing::from_percent(12.0).unwrap();
        let keynote_background = test_background();
        let keynote_outline = TextOutline::standard();
        let keynote_shadow = TextShadow::standard();
        let keynote_line_spacing = test_line_spacing();
        let keynote_paragraph_spacing = test_paragraph_spacing();
        let keynote_indents = test_indents();
        let keynote_tab_stops = test_tab_stops();
        let mut keynote = KeynoteDocumentBuilder::new()
            .title("Aligned")
            .build()
            .unwrap();
        let table = keynote
            .add_slide_table(
                0,
                "Aligned",
                2,
                2,
                DrawablePoint { x: 100.0, y: 100.0 },
                DrawableSize {
                    width: 400.0,
                    height: 200.0,
                },
            )
            .unwrap();
        keynote
            .set_slide_table_cell_text_alignment(
                0,
                table.model_object_id,
                1,
                1,
                TextAlignment::Left,
            )
            .unwrap();
        keynote
            .set_slide_table_cell_text_style(0, table.model_object_id, 1, 1, keynote_style)
            .unwrap();
        keynote
            .set_slide_table_cell_text_font(0, table.model_object_id, 1, 1, keynote_font.clone())
            .unwrap();
        keynote
            .set_slide_table_cell_text_color(0, table.model_object_id, 1, 1, keynote_color)
            .unwrap();
        keynote
            .set_slide_table_cell_text_decorations(
                0,
                table.model_object_id,
                1,
                1,
                keynote_decorations,
            )
            .unwrap();
        keynote
            .set_slide_table_cell_text_baseline_shift(
                0,
                table.model_object_id,
                1,
                1,
                keynote_baseline_shift,
            )
            .unwrap();
        keynote
            .set_slide_table_cell_text_capitalization(
                0,
                table.model_object_id,
                1,
                1,
                TextCapitalization::AllCaps,
            )
            .unwrap();
        keynote
            .set_slide_table_cell_text_character_spacing(
                0,
                table.model_object_id,
                1,
                1,
                keynote_spacing,
            )
            .unwrap();
        keynote
            .set_slide_table_cell_text_ligatures(
                0,
                table.model_object_id,
                1,
                1,
                TextLigatures::RequiredOnly,
            )
            .unwrap();
        keynote
            .set_slide_table_cell_text_background(
                0,
                table.model_object_id,
                1,
                1,
                keynote_background,
            )
            .unwrap();
        keynote
            .set_slide_table_cell_text_outline(0, table.model_object_id, 1, 1, keynote_outline)
            .unwrap();
        keynote
            .set_slide_table_cell_text_script(
                0,
                table.model_object_id,
                1,
                1,
                TextScript::Superscript,
            )
            .unwrap();
        keynote
            .set_slide_table_cell_text_shadow(0, table.model_object_id, 1, 1, keynote_shadow)
            .unwrap();
        keynote
            .set_slide_table_cell_paragraph_line_spacing(
                0,
                table.model_object_id,
                1,
                1,
                keynote_line_spacing,
            )
            .unwrap();
        keynote
            .set_slide_table_cell_paragraph_spacing(
                0,
                table.model_object_id,
                1,
                1,
                keynote_paragraph_spacing,
            )
            .unwrap();
        keynote
            .set_slide_table_cell_paragraph_indents(0, table.model_object_id, 1, 1, keynote_indents)
            .unwrap();
        keynote
            .set_slide_table_cell_paragraph_tab_stops(
                0,
                table.model_object_id,
                1,
                1,
                keynote_tab_stops.clone(),
            )
            .unwrap();
        let mut keynote =
            crate::keynote::KeynoteEditor::from_bytes(&keynote.to_bytes().unwrap()).unwrap();
        assert_eq!(
            keynote
                .slide_table_cell_text_alignment(0, table.model_object_id, 1, 1)
                .unwrap(),
            TextAlignment::Left
        );
        assert_eq!(
            keynote
                .slide_table_cell_text_style(0, table.model_object_id, 1, 1)
                .unwrap(),
            keynote_style
        );
        assert_eq!(
            keynote
                .slide_table_cell_text_font(0, table.model_object_id, 1, 1)
                .unwrap(),
            keynote_font
        );
        assert_eq!(
            keynote
                .slide_table_cell_text_color(0, table.model_object_id, 1, 1)
                .unwrap(),
            keynote_color
        );
        assert_eq!(
            keynote
                .slide_table_cell_text_decorations(0, table.model_object_id, 1, 1)
                .unwrap(),
            keynote_decorations
        );
        assert_eq!(
            keynote
                .slide_table_cell_text_baseline_shift(0, table.model_object_id, 1, 1)
                .unwrap(),
            keynote_baseline_shift
        );
        assert_eq!(
            keynote
                .slide_table_cell_text_capitalization(0, table.model_object_id, 1, 1)
                .unwrap(),
            TextCapitalization::AllCaps
        );
        assert_eq!(
            keynote
                .slide_table_cell_text_character_spacing(0, table.model_object_id, 1, 1)
                .unwrap(),
            keynote_spacing
        );
        assert_eq!(
            keynote
                .slide_table_cell_text_ligatures(0, table.model_object_id, 1, 1)
                .unwrap(),
            TextLigatures::RequiredOnly
        );
        assert_eq!(
            keynote
                .slide_table_cell_text_background(0, table.model_object_id, 1, 1)
                .unwrap(),
            keynote_background
        );
        assert_eq!(
            keynote
                .slide_table_cell_text_outline(0, table.model_object_id, 1, 1)
                .unwrap(),
            keynote_outline
        );
        assert_eq!(
            keynote
                .slide_table_cell_text_script(0, table.model_object_id, 1, 1)
                .unwrap(),
            TextScript::Superscript
        );
        assert_eq!(
            keynote
                .slide_table_cell_text_shadow(0, table.model_object_id, 1, 1)
                .unwrap(),
            keynote_shadow
        );
        assert_eq!(
            keynote
                .slide_table_cell_paragraph_line_spacing(0, table.model_object_id, 1, 1)
                .unwrap(),
            keynote_line_spacing
        );
        assert_eq!(
            keynote
                .slide_table_cell_paragraph_spacing(0, table.model_object_id, 1, 1)
                .unwrap(),
            keynote_paragraph_spacing
        );
        assert_eq!(
            keynote
                .slide_table_cell_paragraph_indents(0, table.model_object_id, 1, 1)
                .unwrap(),
            keynote_indents
        );
        assert_eq!(
            keynote
                .slide_table_cell_paragraph_tab_stops(0, table.model_object_id, 1, 1)
                .unwrap(),
            keynote_tab_stops
        );
        assert!(
            keynote
                .reset_slide_table_cell_text_capitalization(0, table.model_object_id, 1, 1)
                .unwrap()
        );
        assert!(
            keynote
                .reset_slide_table_cell_text_character_spacing(0, table.model_object_id, 1, 1)
                .unwrap()
        );
        assert!(
            keynote
                .reset_slide_table_cell_text_baseline_shift(0, table.model_object_id, 1, 1)
                .unwrap()
        );
        assert!(
            keynote
                .reset_slide_table_cell_text_ligatures(0, table.model_object_id, 1, 1)
                .unwrap()
        );
        assert!(
            keynote
                .reset_slide_table_cell_text_background(0, table.model_object_id, 1, 1)
                .unwrap()
        );
        assert!(
            keynote
                .reset_slide_table_cell_text_outline(0, table.model_object_id, 1, 1)
                .unwrap()
        );
        assert!(
            keynote
                .reset_slide_table_cell_text_script(0, table.model_object_id, 1, 1)
                .unwrap()
        );
        assert!(
            keynote
                .reset_slide_table_cell_text_shadow(0, table.model_object_id, 1, 1)
                .unwrap()
        );
        assert!(
            keynote
                .reset_slide_table_cell_paragraph_spacing(0, table.model_object_id, 1, 1)
                .unwrap()
        );
        assert!(
            keynote
                .reset_slide_table_cell_paragraph_line_spacing(0, table.model_object_id, 1, 1)
                .unwrap()
        );
        assert!(
            keynote
                .reset_slide_table_cell_paragraph_indents(0, table.model_object_id, 1, 1)
                .unwrap()
        );
        assert!(
            keynote
                .reset_slide_table_cell_paragraph_tab_stops(0, table.model_object_id, 1, 1)
                .unwrap()
        );
        assert!(
            keynote
                .reset_slide_table_cell_text_decorations(0, table.model_object_id, 1, 1)
                .unwrap()
        );
        assert!(
            keynote
                .reset_slide_table_cell_text_color(0, table.model_object_id, 1, 1)
                .unwrap()
        );
        assert!(
            keynote
                .reset_slide_table_cell_text_font(0, table.model_object_id, 1, 1)
                .unwrap()
        );
        assert!(
            keynote
                .reset_slide_table_cell_text_style(0, table.model_object_id, 1, 1)
                .unwrap()
        );
    }

    #[test]
    fn invalid_paragraph_style_coordinate_is_transactional() {
        let mut editor = NumbersDocumentBuilder::new()
            .table_dimensions(2, 2)
            .build()
            .unwrap();
        let table_id = editor.tables().unwrap()[0].object_id;
        let before = editor.to_bytes().unwrap();
        assert!(
            editor
                .set_table_cell_text_alignment(table_id, 2, 1, TextAlignment::Center)
                .is_err()
        );
        assert!(
            editor
                .set_table_cell_text_style(table_id, 1, 2, TextStyle::default())
                .is_err()
        );
        assert!(
            editor
                .set_table_cell_text_font(
                    table_id,
                    1,
                    2,
                    TextFont::named("CourierNewPSMT").unwrap(),
                )
                .is_err()
        );
        assert!(
            editor
                .set_table_cell_text_color(table_id, 1, 2, test_color())
                .is_err()
        );
        assert!(
            editor
                .set_table_cell_text_decorations(
                    table_id,
                    1,
                    2,
                    TextDecorations::new(TextUnderline::Single, TextStrikethrough::Single),
                )
                .is_err()
        );
        assert!(
            editor
                .set_table_cell_text_baseline_shift(table_id, 1, 2, TextBaselineShift::ZERO)
                .is_err()
        );
        assert!(
            editor
                .set_table_cell_text_capitalization(table_id, 1, 2, TextCapitalization::AllCaps,)
                .is_err()
        );
        assert!(
            editor
                .set_table_cell_text_character_spacing(
                    table_id,
                    1,
                    2,
                    TextCharacterSpacing::NORMAL,
                )
                .is_err()
        );
        assert!(
            editor
                .set_table_cell_text_ligatures(table_id, 1, 2, TextLigatures::All)
                .is_err()
        );
        assert!(
            editor
                .set_table_cell_paragraph_line_spacing(table_id, 1, 2, test_line_spacing())
                .is_err()
        );
        assert!(
            editor
                .set_table_cell_paragraph_spacing(table_id, 1, 2, test_paragraph_spacing())
                .is_err()
        );
        assert!(
            editor
                .set_table_cell_paragraph_indents(table_id, 1, 2, test_indents())
                .is_err()
        );
        assert!(
            editor
                .set_table_cell_paragraph_tab_stops(table_id, 1, 2, test_tab_stops())
                .is_err()
        );
        assert!(
            editor
                .set_table_cell_text_background(table_id, 1, 2, test_background())
                .is_err()
        );
        assert!(
            editor
                .set_table_cell_text_outline(table_id, 1, 2, TextOutline::standard())
                .is_err()
        );
        assert!(
            editor
                .set_table_cell_text_script(table_id, 1, 2, TextScript::Superscript)
                .is_err()
        );
        assert!(
            editor
                .set_table_cell_text_shadow(table_id, 1, 2, TextShadow::standard())
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), before);
    }
}
