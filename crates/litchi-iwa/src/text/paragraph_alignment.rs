//! Composable native paragraph-style CRUD shared by Pages, Numbers, and Keynote.

mod native;
mod storage;

use std::borrow::Cow;

use crate::package_metadata::{
    next_object_identifier, release_package_identifier_suffix, set_package_last_object_identifier,
};
use crate::shapes::{RgbaColor, insert_style_variation, remove_style_variation};
use crate::{Error, IWorkPackage, Result};

use self::native::ParagraphStyleOverrides;
use super::font::TextFont;
use super::paragraph_tabs::ParagraphTabStops;
use super::style::{
    ParagraphIndents, ParagraphLineSpacing, ParagraphSpacing, TextAlignment, TextBackground,
    TextBaselineShift, TextCapitalization, TextCharacterSpacing, TextDecorations, TextLigatures,
    TextOutline, TextScript, TextShadow, TextStyle,
};
use super::style_registry::{
    object_archive_name, register_private_style, unregister_private_style,
};

#[derive(Debug, Clone)]
enum ParagraphProperty<'a> {
    TextStyle(TextStyle),
    TextFont(TextFont),
    TextDecorations(TextDecorations),
    TextColor(RgbaColor),
    TextCapitalization(TextCapitalization),
    TextScript(TextScript),
    TextBaselineShift(TextBaselineShift),
    TextCharacterSpacing(TextCharacterSpacing),
    TextLigatures(TextLigatures),
    TextOutline(TextOutline),
    TextShadow(TextShadow),
    TextBackground(TextBackground),
    Alignment(TextAlignment),
    LineSpacing(ParagraphLineSpacing),
    Spacing(ParagraphSpacing),
    Indents(ParagraphIndents),
    TabStops(Cow<'a, ParagraphTabStops>),
}

#[derive(Debug, Clone, Copy)]
enum ParagraphPropertyKind {
    TextStyle,
    TextFont,
    TextDecorations,
    TextColor,
    TextCapitalization,
    TextScript,
    TextBaselineShift,
    TextCharacterSpacing,
    TextLigatures,
    TextOutline,
    TextShadow,
    TextBackground,
    Alignment,
    LineSpacing,
    Spacing,
    Indents,
    TabStops,
}

#[derive(Debug, Clone)]
enum InheritedCharacterProperty {
    None,
    TextStyle(TextStyle),
    TextFont(TextFont),
    TextDecorations(TextDecorations),
    TextColor(RgbaColor),
    TextCapitalization(TextCapitalization),
    TextScript(TextScript),
    TextBaselineShift(TextBaselineShift),
    TextCharacterSpacing(TextCharacterSpacing),
    TextLigatures(TextLigatures),
    TextOutline(TextOutline),
    TextShadow(TextShadow),
    TextBackground(TextBackground),
}

pub(super) fn text_style(package: &IWorkPackage, storage_id: u64) -> Result<TextStyle> {
    let storage = storage::locate(package, storage_id)?;
    native::inherited_text_style(package, storage.style_id)
}

pub(super) fn set_text_style(
    package: &mut IWorkPackage,
    storage_id: u64,
    style: TextStyle,
) -> Result<()> {
    if text_style(package, storage_id)? == style {
        return Ok(());
    }
    set_property(package, storage_id, ParagraphProperty::TextStyle(style))
}

pub(super) fn reset_text_style(package: &mut IWorkPackage, storage_id: u64) -> Result<bool> {
    reset_property(package, storage_id, ParagraphPropertyKind::TextStyle)
}

pub(super) fn text_font(package: &IWorkPackage, storage_id: u64) -> Result<TextFont> {
    let storage = storage::locate(package, storage_id)?;
    native::inherited_text_font(package, storage.style_id)
}

pub(super) fn set_text_font(
    package: &mut IWorkPackage,
    storage_id: u64,
    font: TextFont,
) -> Result<()> {
    if text_font(package, storage_id)? == font {
        return Ok(());
    }
    set_property(package, storage_id, ParagraphProperty::TextFont(font))
}

pub(super) fn reset_text_font(package: &mut IWorkPackage, storage_id: u64) -> Result<bool> {
    reset_property(package, storage_id, ParagraphPropertyKind::TextFont)
}

pub(super) fn text_decorations(package: &IWorkPackage, storage_id: u64) -> Result<TextDecorations> {
    let storage = storage::locate(package, storage_id)?;
    native::inherited_text_decorations(package, storage.style_id)
}

pub(super) fn set_text_decorations(
    package: &mut IWorkPackage,
    storage_id: u64,
    decorations: TextDecorations,
) -> Result<()> {
    if text_decorations(package, storage_id)? == decorations {
        return Ok(());
    }
    set_property(
        package,
        storage_id,
        ParagraphProperty::TextDecorations(decorations),
    )
}

pub(super) fn reset_text_decorations(package: &mut IWorkPackage, storage_id: u64) -> Result<bool> {
    reset_property(package, storage_id, ParagraphPropertyKind::TextDecorations)
}

pub(super) fn text_color(package: &IWorkPackage, storage_id: u64) -> Result<RgbaColor> {
    let storage = storage::locate(package, storage_id)?;
    native::inherited_text_color(package, storage.style_id)
}

pub(super) fn set_text_color(
    package: &mut IWorkPackage,
    storage_id: u64,
    color: RgbaColor,
) -> Result<()> {
    if text_color(package, storage_id)? == color {
        return Ok(());
    }
    set_property(package, storage_id, ParagraphProperty::TextColor(color))
}

pub(super) fn reset_text_color(package: &mut IWorkPackage, storage_id: u64) -> Result<bool> {
    reset_property(package, storage_id, ParagraphPropertyKind::TextColor)
}

pub(super) fn text_capitalization(
    package: &IWorkPackage,
    storage_id: u64,
) -> Result<TextCapitalization> {
    let storage = storage::locate(package, storage_id)?;
    native::inherited_text_capitalization(package, storage.style_id)
}

pub(super) fn set_text_capitalization(
    package: &mut IWorkPackage,
    storage_id: u64,
    capitalization: TextCapitalization,
) -> Result<()> {
    if text_capitalization(package, storage_id)? == capitalization {
        return Ok(());
    }
    set_property(
        package,
        storage_id,
        ParagraphProperty::TextCapitalization(capitalization),
    )
}

pub(super) fn reset_text_capitalization(
    package: &mut IWorkPackage,
    storage_id: u64,
) -> Result<bool> {
    reset_property(
        package,
        storage_id,
        ParagraphPropertyKind::TextCapitalization,
    )
}

pub(super) fn text_script(package: &IWorkPackage, storage_id: u64) -> Result<TextScript> {
    let storage = storage::locate(package, storage_id)?;
    native::inherited_text_script(package, storage.style_id)
}

pub(super) fn set_text_script(
    package: &mut IWorkPackage,
    storage_id: u64,
    script: TextScript,
) -> Result<()> {
    if text_script(package, storage_id)? == script {
        return Ok(());
    }
    set_property(package, storage_id, ParagraphProperty::TextScript(script))
}

pub(super) fn reset_text_script(package: &mut IWorkPackage, storage_id: u64) -> Result<bool> {
    reset_property(package, storage_id, ParagraphPropertyKind::TextScript)
}

pub(super) fn text_baseline_shift(
    package: &IWorkPackage,
    storage_id: u64,
) -> Result<TextBaselineShift> {
    let storage = storage::locate(package, storage_id)?;
    native::inherited_text_baseline_shift(package, storage.style_id)
}

pub(super) fn set_text_baseline_shift(
    package: &mut IWorkPackage,
    storage_id: u64,
    shift: TextBaselineShift,
) -> Result<()> {
    if text_baseline_shift(package, storage_id)? == shift {
        return Ok(());
    }
    set_property(
        package,
        storage_id,
        ParagraphProperty::TextBaselineShift(shift),
    )
}

pub(super) fn reset_text_baseline_shift(
    package: &mut IWorkPackage,
    storage_id: u64,
) -> Result<bool> {
    reset_property(
        package,
        storage_id,
        ParagraphPropertyKind::TextBaselineShift,
    )
}

pub(super) fn text_character_spacing(
    package: &IWorkPackage,
    storage_id: u64,
) -> Result<TextCharacterSpacing> {
    let storage = storage::locate(package, storage_id)?;
    native::inherited_text_character_spacing(package, storage.style_id)
}

pub(super) fn set_text_character_spacing(
    package: &mut IWorkPackage,
    storage_id: u64,
    spacing: TextCharacterSpacing,
) -> Result<()> {
    if text_character_spacing(package, storage_id)? == spacing {
        return Ok(());
    }
    set_property(
        package,
        storage_id,
        ParagraphProperty::TextCharacterSpacing(spacing),
    )
}

pub(super) fn reset_text_character_spacing(
    package: &mut IWorkPackage,
    storage_id: u64,
) -> Result<bool> {
    reset_property(
        package,
        storage_id,
        ParagraphPropertyKind::TextCharacterSpacing,
    )
}

pub(super) fn text_ligatures(package: &IWorkPackage, storage_id: u64) -> Result<TextLigatures> {
    let storage = storage::locate(package, storage_id)?;
    native::inherited_text_ligatures(package, storage.style_id)
}

pub(super) fn set_text_ligatures(
    package: &mut IWorkPackage,
    storage_id: u64,
    ligatures: TextLigatures,
) -> Result<()> {
    if text_ligatures(package, storage_id)? == ligatures {
        return Ok(());
    }
    set_property(
        package,
        storage_id,
        ParagraphProperty::TextLigatures(ligatures),
    )
}

pub(super) fn reset_text_ligatures(package: &mut IWorkPackage, storage_id: u64) -> Result<bool> {
    reset_property(package, storage_id, ParagraphPropertyKind::TextLigatures)
}

pub(super) fn text_outline(package: &IWorkPackage, storage_id: u64) -> Result<TextOutline> {
    let storage = storage::locate(package, storage_id)?;
    native::inherited_text_outline(package, storage.style_id)
}

pub(super) fn set_text_outline(
    package: &mut IWorkPackage,
    storage_id: u64,
    outline: TextOutline,
) -> Result<()> {
    if text_outline(package, storage_id)? == outline {
        return Ok(());
    }
    set_property(package, storage_id, ParagraphProperty::TextOutline(outline))
}

pub(super) fn reset_text_outline(package: &mut IWorkPackage, storage_id: u64) -> Result<bool> {
    reset_property(package, storage_id, ParagraphPropertyKind::TextOutline)
}

pub(super) fn text_shadow(package: &IWorkPackage, storage_id: u64) -> Result<TextShadow> {
    let storage = storage::locate(package, storage_id)?;
    native::inherited_text_shadow(package, storage.style_id)
}

pub(super) fn set_text_shadow(
    package: &mut IWorkPackage,
    storage_id: u64,
    shadow: TextShadow,
) -> Result<()> {
    if text_shadow(package, storage_id)? == shadow {
        return Ok(());
    }
    set_property(package, storage_id, ParagraphProperty::TextShadow(shadow))
}

pub(super) fn reset_text_shadow(package: &mut IWorkPackage, storage_id: u64) -> Result<bool> {
    reset_property(package, storage_id, ParagraphPropertyKind::TextShadow)
}

pub(super) fn text_background(package: &IWorkPackage, storage_id: u64) -> Result<TextBackground> {
    let storage = storage::locate(package, storage_id)?;
    native::inherited_text_background(package, storage.style_id)
}

pub(super) fn set_text_background(
    package: &mut IWorkPackage,
    storage_id: u64,
    background: TextBackground,
) -> Result<()> {
    if text_background(package, storage_id)? == background {
        return Ok(());
    }
    set_property(
        package,
        storage_id,
        ParagraphProperty::TextBackground(background),
    )
}

pub(super) fn reset_text_background(package: &mut IWorkPackage, storage_id: u64) -> Result<bool> {
    reset_property(package, storage_id, ParagraphPropertyKind::TextBackground)
}

pub(super) fn paragraph_alignment(
    package: &IWorkPackage,
    storage_id: u64,
) -> Result<TextAlignment> {
    let storage = storage::locate(package, storage_id)?;
    native::inherited_alignment(package, storage.style_id)
}

pub(super) fn set_paragraph_alignment(
    package: &mut IWorkPackage,
    storage_id: u64,
    alignment: TextAlignment,
) -> Result<()> {
    if paragraph_alignment(package, storage_id)? == alignment {
        return Ok(());
    }
    set_property(package, storage_id, ParagraphProperty::Alignment(alignment))
}

pub(super) fn reset_paragraph_alignment(
    package: &mut IWorkPackage,
    storage_id: u64,
) -> Result<bool> {
    reset_property(package, storage_id, ParagraphPropertyKind::Alignment)
}

pub(super) fn paragraph_line_spacing(
    package: &IWorkPackage,
    storage_id: u64,
) -> Result<ParagraphLineSpacing> {
    let storage = storage::locate(package, storage_id)?;
    native::inherited_line_spacing(package, storage.style_id)
}

pub(super) fn set_paragraph_line_spacing(
    package: &mut IWorkPackage,
    storage_id: u64,
    spacing: ParagraphLineSpacing,
) -> Result<()> {
    if paragraph_line_spacing(package, storage_id)? == spacing {
        return Ok(());
    }
    set_property(package, storage_id, ParagraphProperty::LineSpacing(spacing))
}

pub(super) fn reset_paragraph_line_spacing(
    package: &mut IWorkPackage,
    storage_id: u64,
) -> Result<bool> {
    reset_property(package, storage_id, ParagraphPropertyKind::LineSpacing)
}

pub(super) fn paragraph_spacing(
    package: &IWorkPackage,
    storage_id: u64,
) -> Result<ParagraphSpacing> {
    let storage = storage::locate(package, storage_id)?;
    native::inherited_spacing(package, storage.style_id)
}

pub(super) fn set_paragraph_spacing(
    package: &mut IWorkPackage,
    storage_id: u64,
    spacing: ParagraphSpacing,
) -> Result<()> {
    if paragraph_spacing(package, storage_id)? == spacing {
        return Ok(());
    }
    set_property(package, storage_id, ParagraphProperty::Spacing(spacing))
}

pub(super) fn reset_paragraph_spacing(package: &mut IWorkPackage, storage_id: u64) -> Result<bool> {
    reset_property(package, storage_id, ParagraphPropertyKind::Spacing)
}

pub(super) fn paragraph_indents(
    package: &IWorkPackage,
    storage_id: u64,
) -> Result<ParagraphIndents> {
    let storage = storage::locate(package, storage_id)?;
    native::inherited_indents(package, storage.style_id)
}

pub(super) fn set_paragraph_indents(
    package: &mut IWorkPackage,
    storage_id: u64,
    indents: ParagraphIndents,
) -> Result<()> {
    if paragraph_indents(package, storage_id)? == indents {
        return Ok(());
    }
    set_property(package, storage_id, ParagraphProperty::Indents(indents))
}

pub(super) fn reset_paragraph_indents(package: &mut IWorkPackage, storage_id: u64) -> Result<bool> {
    reset_property(package, storage_id, ParagraphPropertyKind::Indents)
}

pub(super) fn paragraph_tab_stops(
    package: &IWorkPackage,
    storage_id: u64,
) -> Result<ParagraphTabStops> {
    let storage = storage::locate(package, storage_id)?;
    native::inherited_tab_stops(package, storage.style_id)
}

pub(super) fn set_paragraph_tab_stops(
    package: &mut IWorkPackage,
    storage_id: u64,
    stops: &ParagraphTabStops,
) -> Result<()> {
    if paragraph_tab_stops(package, storage_id)?.as_slice() == stops.as_slice() {
        return Ok(());
    }
    set_property(
        package,
        storage_id,
        ParagraphProperty::TabStops(Cow::Borrowed(stops)),
    )
}

pub(super) fn reset_paragraph_tab_stops(
    package: &mut IWorkPackage,
    storage_id: u64,
) -> Result<bool> {
    reset_property(package, storage_id, ParagraphPropertyKind::TabStops)
}

fn set_property(
    package: &mut IWorkPackage,
    storage_id: u64,
    property: ParagraphProperty<'_>,
) -> Result<()> {
    let storage = storage::locate(package, storage_id)?;
    let style = native::locate_style(package, storage.style_id)?;
    let stylesheet_id = native::stylesheet_id(&style.style, storage.style_id)?;
    let stylesheet_archive_name = object_archive_name(package, stylesheet_id)?;
    if stylesheet_archive_name != style.archive_name {
        return Err(Error::InvalidFormat(format!(
            "iWork paragraph style {} is not stored with stylesheet {stylesheet_id}",
            storage.style_id
        )));
    }

    if let Some(mut overrides) = native::direct_overrides(&style.style, &style.message.data)?
        && native::is_exclusive(package, storage.style_id)?
    {
        let parent_style_id = native::parent_style_id(&style.style, storage.style_id)?;
        let inherited = inherited_character_property(package, parent_style_id, &property)?;
        apply_property(&mut overrides, &property, inherited)?;
        if overrides.is_empty() {
            let mut staged = package.clone();
            storage::patch_style_reference(
                &mut staged,
                &storage.archive_name,
                storage_id,
                storage.style_id,
                parent_style_id,
            )?;
            remove_style_variation(
                &mut staged,
                &style.archive_name,
                stylesheet_id,
                parent_style_id,
                storage.style_id,
            )?;
            unregister_private_style(
                &mut staged,
                &storage.archive_name,
                &style.archive_name,
                storage.style_id,
                Some(parent_style_id),
            )?;
            release_package_identifier_suffix(&mut staged, &[storage.style_id])?;
            validate_property(&staged, storage_id, property)?;
            *package = staged;
            return Ok(());
        }
        let replacement =
            native::variation_object(storage.style_id, parent_style_id, stylesheet_id, overrides)?;
        let mut staged = package.clone();
        native::replace_variation(
            &mut staged,
            &style.archive_name,
            storage.style_id,
            replacement,
        )?;
        validate_property(&staged, storage_id, property)?;
        *package = staged;
        return Ok(());
    }

    let new_style_id = next_object_identifier(package)?;
    let mut overrides = ParagraphStyleOverrides::default();
    let inherited = inherited_character_property(package, storage.style_id, &property)?;
    apply_property(&mut overrides, &property, inherited)?;
    let new_style =
        native::variation_object(new_style_id, storage.style_id, stylesheet_id, overrides)?;
    let mut staged = package.clone();
    storage::patch_style_reference(
        &mut staged,
        &storage.archive_name,
        storage_id,
        storage.style_id,
        new_style_id,
    )?;
    insert_style_variation(
        &mut staged,
        &style.archive_name,
        stylesheet_id,
        storage.style_id,
        new_style_id,
        new_style,
    )?;
    register_private_style(
        &mut staged,
        &storage.archive_name,
        &style.archive_name,
        new_style_id,
    )?;
    set_package_last_object_identifier(&mut staged, new_style_id)?;
    validate_property(&staged, storage_id, property)?;
    *package = staged;
    Ok(())
}

fn reset_property(
    package: &mut IWorkPackage,
    storage_id: u64,
    kind: ParagraphPropertyKind,
) -> Result<bool> {
    let storage = storage::locate(package, storage_id)?;
    let style = native::locate_style(package, storage.style_id)?;
    let Some(mut overrides) = native::direct_overrides(&style.style, &style.message.data)? else {
        return Ok(false);
    };
    if !has_property(&overrides, kind) || !native::is_exclusive(package, storage.style_id)? {
        return Ok(false);
    }
    clear_property(&mut overrides, kind);
    let parent_style_id = native::parent_style_id(&style.style, storage.style_id)?;
    let stylesheet_id = native::stylesheet_id(&style.style, storage.style_id)?;
    let expected = inherited_property(package, parent_style_id, kind)?;
    let mut staged = package.clone();
    if overrides.is_empty() {
        storage::patch_style_reference(
            &mut staged,
            &storage.archive_name,
            storage_id,
            storage.style_id,
            parent_style_id,
        )?;
        remove_style_variation(
            &mut staged,
            &style.archive_name,
            stylesheet_id,
            parent_style_id,
            storage.style_id,
        )?;
        unregister_private_style(
            &mut staged,
            &storage.archive_name,
            &style.archive_name,
            storage.style_id,
            Some(parent_style_id),
        )?;
        release_package_identifier_suffix(&mut staged, &[storage.style_id])?;
    } else {
        let replacement =
            native::variation_object(storage.style_id, parent_style_id, stylesheet_id, overrides)?;
        native::replace_variation(
            &mut staged,
            &style.archive_name,
            storage.style_id,
            replacement,
        )?;
    }
    validate_expected_property(&staged, storage_id, expected)?;
    *package = staged;
    Ok(true)
}

fn inherited_character_property(
    package: &IWorkPackage,
    parent_style_id: u64,
    property: &ParagraphProperty<'_>,
) -> Result<InheritedCharacterProperty> {
    match property {
        ParagraphProperty::TextStyle(_) => native::inherited_text_style(package, parent_style_id)
            .map(InheritedCharacterProperty::TextStyle),
        ParagraphProperty::TextFont(_) => native::inherited_text_font(package, parent_style_id)
            .map(InheritedCharacterProperty::TextFont),
        ParagraphProperty::TextDecorations(_) => {
            native::inherited_text_decorations(package, parent_style_id)
                .map(InheritedCharacterProperty::TextDecorations)
        },
        ParagraphProperty::TextColor(_) => native::inherited_text_color(package, parent_style_id)
            .map(InheritedCharacterProperty::TextColor),
        ParagraphProperty::TextCapitalization(_) => {
            native::inherited_text_capitalization(package, parent_style_id)
                .map(InheritedCharacterProperty::TextCapitalization)
        },
        ParagraphProperty::TextScript(_) => native::inherited_text_script(package, parent_style_id)
            .map(InheritedCharacterProperty::TextScript),
        ParagraphProperty::TextBaselineShift(_) => {
            native::inherited_text_baseline_shift(package, parent_style_id)
                .map(InheritedCharacterProperty::TextBaselineShift)
        },
        ParagraphProperty::TextCharacterSpacing(_) => {
            native::inherited_text_character_spacing(package, parent_style_id)
                .map(InheritedCharacterProperty::TextCharacterSpacing)
        },
        ParagraphProperty::TextLigatures(_) => {
            native::inherited_text_ligatures(package, parent_style_id)
                .map(InheritedCharacterProperty::TextLigatures)
        },
        ParagraphProperty::TextOutline(_) => {
            native::inherited_text_outline(package, parent_style_id)
                .map(InheritedCharacterProperty::TextOutline)
        },
        ParagraphProperty::TextShadow(_) => native::inherited_text_shadow(package, parent_style_id)
            .map(InheritedCharacterProperty::TextShadow),
        ParagraphProperty::TextBackground(_) => {
            native::inherited_text_background(package, parent_style_id)
                .map(InheritedCharacterProperty::TextBackground)
        },
        _ => Ok(InheritedCharacterProperty::None),
    }
}

fn apply_property(
    overrides: &mut ParagraphStyleOverrides,
    property: &ParagraphProperty<'_>,
    inherited: InheritedCharacterProperty,
) -> Result<()> {
    match property {
        ParagraphProperty::TextStyle(style) => {
            let InheritedCharacterProperty::TextStyle(inherited) = inherited else {
                return Err(Error::InvalidFormat(
                    "text-style mutation has no inherited character formatting".to_owned(),
                ));
            };
            overrides.point_size =
                (style.point_size != inherited.point_size).then_some(style.point_size);
            overrides.bold = (style.bold != inherited.bold).then_some(style.bold);
            overrides.italic = (style.italic != inherited.italic).then_some(style.italic);
        },
        ParagraphProperty::TextFont(font) => {
            let InheritedCharacterProperty::TextFont(inherited) = inherited else {
                return Err(Error::InvalidFormat(
                    "text-font mutation has no inherited character font".to_owned(),
                ));
            };
            overrides.font = (font != &inherited).then(|| font.clone());
        },
        ParagraphProperty::TextDecorations(decorations) => {
            let InheritedCharacterProperty::TextDecorations(inherited) = inherited else {
                return Err(Error::InvalidFormat(
                    "text-decoration mutation has no inherited character formatting".to_owned(),
                ));
            };
            overrides.underline =
                (decorations.underline != inherited.underline).then_some(decorations.underline);
            overrides.strikethrough = (decorations.strikethrough != inherited.strikethrough)
                .then_some(decorations.strikethrough);
        },
        ParagraphProperty::TextColor(color) => {
            let InheritedCharacterProperty::TextColor(inherited) = inherited else {
                return Err(Error::InvalidFormat(
                    "text-color mutation has no inherited character color".to_owned(),
                ));
            };
            overrides.font_color = (*color != inherited).then_some(*color);
        },
        ParagraphProperty::TextCapitalization(capitalization) => {
            let InheritedCharacterProperty::TextCapitalization(inherited) = inherited else {
                return Err(Error::InvalidFormat(
                    "text-capitalization mutation has no inherited character formatting".to_owned(),
                ));
            };
            overrides.capitalization = (*capitalization != inherited).then_some(*capitalization);
        },
        ParagraphProperty::TextScript(script) => {
            let InheritedCharacterProperty::TextScript(inherited) = inherited else {
                return Err(Error::InvalidFormat(
                    "text-script mutation has no inherited character formatting".to_owned(),
                ));
            };
            overrides.script = (*script != inherited).then_some(*script);
        },
        ParagraphProperty::TextBaselineShift(shift) => {
            let InheritedCharacterProperty::TextBaselineShift(inherited) = inherited else {
                return Err(Error::InvalidFormat(
                    "text baseline-shift mutation has no inherited character formatting".to_owned(),
                ));
            };
            overrides.baseline_shift = (*shift != inherited).then_some(*shift);
        },
        ParagraphProperty::TextCharacterSpacing(spacing) => {
            let InheritedCharacterProperty::TextCharacterSpacing(inherited) = inherited else {
                return Err(Error::InvalidFormat(
                    "text character-spacing mutation has no inherited character formatting"
                        .to_owned(),
                ));
            };
            overrides.character_spacing = (*spacing != inherited).then_some(*spacing);
        },
        ParagraphProperty::TextLigatures(ligatures) => {
            let InheritedCharacterProperty::TextLigatures(inherited) = inherited else {
                return Err(Error::InvalidFormat(
                    "text ligature mutation has no inherited character formatting".to_owned(),
                ));
            };
            overrides.ligatures = (*ligatures != inherited).then_some(*ligatures);
        },
        ParagraphProperty::TextOutline(outline) => {
            let InheritedCharacterProperty::TextOutline(inherited) = inherited else {
                return Err(Error::InvalidFormat(
                    "text outline mutation has no inherited character formatting".to_owned(),
                ));
            };
            overrides.outline = (*outline != inherited).then_some(*outline);
        },
        ParagraphProperty::TextShadow(shadow) => {
            let InheritedCharacterProperty::TextShadow(inherited) = inherited else {
                return Err(Error::InvalidFormat(
                    "text shadow mutation has no inherited character formatting".to_owned(),
                ));
            };
            overrides.shadow = (*shadow != inherited).then_some(*shadow);
        },
        ParagraphProperty::TextBackground(background) => {
            let InheritedCharacterProperty::TextBackground(inherited) = inherited else {
                return Err(Error::InvalidFormat(
                    "text background mutation has no inherited character formatting".to_owned(),
                ));
            };
            overrides.background = (*background != inherited).then_some(*background);
        },
        ParagraphProperty::Alignment(alignment) => overrides.alignment = Some(*alignment),
        ParagraphProperty::LineSpacing(spacing) => overrides.line_spacing = Some(*spacing),
        ParagraphProperty::Spacing(spacing) => {
            overrides.space_before = Some(spacing.before);
            overrides.space_after = Some(spacing.after);
        },
        ParagraphProperty::Indents(indents) => {
            overrides.first_line_indent = Some(indents.first_line);
            overrides.left_indent = Some(indents.left);
            overrides.right_indent = Some(indents.right);
        },
        ParagraphProperty::TabStops(stops) => {
            overrides.tab_stops = Some(stops.as_ref().clone());
        },
    }
    Ok(())
}

fn has_property(overrides: &ParagraphStyleOverrides, kind: ParagraphPropertyKind) -> bool {
    match kind {
        ParagraphPropertyKind::TextStyle => {
            overrides.point_size.is_some() || overrides.bold.is_some() || overrides.italic.is_some()
        },
        ParagraphPropertyKind::TextFont => overrides.font.is_some(),
        ParagraphPropertyKind::TextDecorations => {
            overrides.underline.is_some() || overrides.strikethrough.is_some()
        },
        ParagraphPropertyKind::TextColor => overrides.font_color.is_some(),
        ParagraphPropertyKind::TextCapitalization => overrides.capitalization.is_some(),
        ParagraphPropertyKind::TextScript => overrides.script.is_some(),
        ParagraphPropertyKind::TextBaselineShift => overrides.baseline_shift.is_some(),
        ParagraphPropertyKind::TextCharacterSpacing => overrides.character_spacing.is_some(),
        ParagraphPropertyKind::TextLigatures => overrides.ligatures.is_some(),
        ParagraphPropertyKind::TextOutline => overrides.outline.is_some(),
        ParagraphPropertyKind::TextShadow => overrides.shadow.is_some(),
        ParagraphPropertyKind::TextBackground => overrides.background.is_some(),
        ParagraphPropertyKind::Alignment => overrides.alignment.is_some(),
        ParagraphPropertyKind::LineSpacing => overrides.line_spacing.is_some(),
        ParagraphPropertyKind::Spacing => {
            overrides.space_before.is_some() || overrides.space_after.is_some()
        },
        ParagraphPropertyKind::Indents => {
            overrides.first_line_indent.is_some()
                || overrides.left_indent.is_some()
                || overrides.right_indent.is_some()
        },
        ParagraphPropertyKind::TabStops => overrides.tab_stops.is_some(),
    }
}

fn clear_property(overrides: &mut ParagraphStyleOverrides, kind: ParagraphPropertyKind) {
    match kind {
        ParagraphPropertyKind::TextStyle => {
            overrides.point_size = None;
            overrides.bold = None;
            overrides.italic = None;
        },
        ParagraphPropertyKind::TextFont => overrides.font = None,
        ParagraphPropertyKind::TextDecorations => {
            overrides.underline = None;
            overrides.strikethrough = None;
        },
        ParagraphPropertyKind::TextColor => overrides.font_color = None,
        ParagraphPropertyKind::TextCapitalization => overrides.capitalization = None,
        ParagraphPropertyKind::TextScript => overrides.script = None,
        ParagraphPropertyKind::TextBaselineShift => overrides.baseline_shift = None,
        ParagraphPropertyKind::TextCharacterSpacing => overrides.character_spacing = None,
        ParagraphPropertyKind::TextLigatures => overrides.ligatures = None,
        ParagraphPropertyKind::TextOutline => overrides.outline = None,
        ParagraphPropertyKind::TextShadow => overrides.shadow = None,
        ParagraphPropertyKind::TextBackground => overrides.background = None,
        ParagraphPropertyKind::Alignment => overrides.alignment = None,
        ParagraphPropertyKind::LineSpacing => overrides.line_spacing = None,
        ParagraphPropertyKind::Spacing => {
            overrides.space_before = None;
            overrides.space_after = None;
        },
        ParagraphPropertyKind::Indents => {
            overrides.first_line_indent = None;
            overrides.left_indent = None;
            overrides.right_indent = None;
        },
        ParagraphPropertyKind::TabStops => overrides.tab_stops = None,
    }
}

fn inherited_property(
    package: &IWorkPackage,
    style_id: u64,
    kind: ParagraphPropertyKind,
) -> Result<ParagraphProperty<'static>> {
    match kind {
        ParagraphPropertyKind::TextStyle => Ok(ParagraphProperty::TextStyle(
            native::inherited_text_style(package, style_id)?,
        )),
        ParagraphPropertyKind::TextFont => Ok(ParagraphProperty::TextFont(
            native::inherited_text_font(package, style_id)?,
        )),
        ParagraphPropertyKind::TextDecorations => Ok(ParagraphProperty::TextDecorations(
            native::inherited_text_decorations(package, style_id)?,
        )),
        ParagraphPropertyKind::TextColor => Ok(ParagraphProperty::TextColor(
            native::inherited_text_color(package, style_id)?,
        )),
        ParagraphPropertyKind::TextCapitalization => Ok(ParagraphProperty::TextCapitalization(
            native::inherited_text_capitalization(package, style_id)?,
        )),
        ParagraphPropertyKind::TextScript => Ok(ParagraphProperty::TextScript(
            native::inherited_text_script(package, style_id)?,
        )),
        ParagraphPropertyKind::TextBaselineShift => Ok(ParagraphProperty::TextBaselineShift(
            native::inherited_text_baseline_shift(package, style_id)?,
        )),
        ParagraphPropertyKind::TextCharacterSpacing => Ok(ParagraphProperty::TextCharacterSpacing(
            native::inherited_text_character_spacing(package, style_id)?,
        )),
        ParagraphPropertyKind::TextLigatures => Ok(ParagraphProperty::TextLigatures(
            native::inherited_text_ligatures(package, style_id)?,
        )),
        ParagraphPropertyKind::TextOutline => Ok(ParagraphProperty::TextOutline(
            native::inherited_text_outline(package, style_id)?,
        )),
        ParagraphPropertyKind::TextShadow => Ok(ParagraphProperty::TextShadow(
            native::inherited_text_shadow(package, style_id)?,
        )),
        ParagraphPropertyKind::TextBackground => Ok(ParagraphProperty::TextBackground(
            native::inherited_text_background(package, style_id)?,
        )),
        ParagraphPropertyKind::Alignment => Ok(ParagraphProperty::Alignment(
            native::inherited_alignment(package, style_id)?,
        )),
        ParagraphPropertyKind::LineSpacing => Ok(ParagraphProperty::LineSpacing(
            native::inherited_line_spacing(package, style_id)?,
        )),
        ParagraphPropertyKind::Spacing => Ok(ParagraphProperty::Spacing(
            native::inherited_spacing(package, style_id)?,
        )),
        ParagraphPropertyKind::Indents => Ok(ParagraphProperty::Indents(
            native::inherited_indents(package, style_id)?,
        )),
        ParagraphPropertyKind::TabStops => Ok(ParagraphProperty::TabStops(Cow::Owned(
            native::inherited_tab_stops(package, style_id)?,
        ))),
    }
}

fn validate_property(
    package: &IWorkPackage,
    storage_id: u64,
    expected: ParagraphProperty<'_>,
) -> Result<()> {
    validate_expected_property(package, storage_id, expected).map_err(|_| {
        Error::InvalidFormat("iWork paragraph-style update failed validation".to_owned())
    })
}

fn validate_expected_property(
    package: &IWorkPackage,
    storage_id: u64,
    expected: ParagraphProperty<'_>,
) -> Result<()> {
    let matches = match expected {
        ParagraphProperty::TextStyle(style) => text_style(package, storage_id)? == style,
        ParagraphProperty::TextFont(font) => text_font(package, storage_id)? == font,
        ParagraphProperty::TextDecorations(decorations) => {
            text_decorations(package, storage_id)? == decorations
        },
        ParagraphProperty::TextColor(color) => text_color(package, storage_id)? == color,
        ParagraphProperty::TextCapitalization(capitalization) => {
            text_capitalization(package, storage_id)? == capitalization
        },
        ParagraphProperty::TextScript(script) => text_script(package, storage_id)? == script,
        ParagraphProperty::TextBaselineShift(shift) => {
            text_baseline_shift(package, storage_id)? == shift
        },
        ParagraphProperty::TextCharacterSpacing(spacing) => {
            text_character_spacing(package, storage_id)? == spacing
        },
        ParagraphProperty::TextLigatures(ligatures) => {
            text_ligatures(package, storage_id)? == ligatures
        },
        ParagraphProperty::TextOutline(outline) => text_outline(package, storage_id)? == outline,
        ParagraphProperty::TextShadow(shadow) => text_shadow(package, storage_id)? == shadow,
        ParagraphProperty::TextBackground(background) => {
            text_background(package, storage_id)? == background
        },
        ParagraphProperty::Alignment(alignment) => {
            paragraph_alignment(package, storage_id)? == alignment
        },
        ParagraphProperty::LineSpacing(spacing) => {
            paragraph_line_spacing(package, storage_id)? == spacing
        },
        ParagraphProperty::Spacing(spacing) => paragraph_spacing(package, storage_id)? == spacing,
        ParagraphProperty::Indents(indents) => paragraph_indents(package, storage_id)? == indents,
        ParagraphProperty::TabStops(stops) => {
            paragraph_tab_stops(package, storage_id)?.as_slice() == stops.as_ref().as_slice()
        },
    };
    if matches {
        Ok(())
    } else {
        Err(Error::InvalidFormat(
            "iWork paragraph-style property does not match its expected value".to_owned(),
        ))
    }
}

#[cfg(test)]
mod tests;
