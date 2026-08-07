//! Typed paragraph-property inheritance with cycle and depth guards.

use std::collections::HashSet;

use crate::protobuf::tswp;
use crate::shapes::RgbaColor;
use crate::text::font::TextFont;
use crate::text::paragraph_direction::{
    ParagraphWritingDirection, from_native as writing_direction_from_native,
};
use crate::text::paragraph_flow::{ParagraphFlow, hyphenation_from_native};
use crate::text::paragraph_tabs::{
    ParagraphDecimalTabCharacter, ParagraphDefaultTabInterval, ParagraphTabStops,
    decimal_character_from_native,
};
use litchi_iwa_text::appearance::{Background, Outline, ParagraphBackground, Shadow};
use litchi_iwa_text::character::{
    TextBaselineShift, TextCapitalization, TextCharacterSpacing, TextDecorations, TextLigatures,
    TextPointSize, TextScript, TextStrikethrough, TextStyle, TextUnderline,
};
use litchi_iwa_text::paragraph::format::{
    Alignment, Borders, IndentPoints, Indents, LineSpacing, Spacing, SpacingPoints,
};
use crate::{Error, IWorkPackage, Result};
use litchi_iwa_text::paragraph::style::{ParagraphFollowingStyle, raw::from_native_id};

use super::super::{NativeTextCharacterSpacing, NativeTextValue};
use super::{
    capitalization_from_character, line_spacing_from_archive, locate_style,
    paragraph_background_from_properties, paragraph_borders_from_properties, tabs,
    text_background_from_character, text_color_from_character, text_font_from_character,
    text_outline_from_character, text_shadow_from_character,
};

const MAX_STYLE_INHERITANCE_DEPTH: usize = 64;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum InheritanceControl {
    Continue,
    Complete,
}

pub(super) fn text_style(package: &IWorkPackage, first_style_id: u64) -> Result<TextStyle> {
    let (point_size, bold, italic) = walk(
        package,
        first_style_id,
        (None, None, None),
        |(point_size, bold, italic), style| {
            if let Some(properties) = style.char_properties.as_ref() {
                if point_size.is_none() {
                    *point_size = properties
                        .font_size
                        .map(TextPointSize::from_points)
                        .transpose()?;
                }
                if bold.is_none() {
                    *bold = properties.bold;
                }
                if italic.is_none() {
                    *italic = properties.italic;
                }
            }
            Ok(
                if point_size.is_some() && bold.is_some() && italic.is_some() {
                    InheritanceControl::Complete
                } else {
                    InheritanceControl::Continue
                },
            )
        },
    )?;
    Ok(TextStyle::new(point_size.unwrap_or_default())
        .with_bold(bold.unwrap_or(false))
        .with_italic(italic.unwrap_or(false)))
}

pub(super) fn text_font(package: &IWorkPackage, first_style_id: u64) -> Result<TextFont> {
    let value = walk(package, first_style_id, None, |value, style| {
        let Some(properties) = style.char_properties.as_ref() else {
            return Ok(InheritanceControl::Continue);
        };
        let Some(font) = text_font_from_character(properties)? else {
            return Ok(InheritanceControl::Continue);
        };
        *value = Some(font);
        Ok(InheritanceControl::Complete)
    })?;
    Ok(value.unwrap_or_default())
}

pub(super) fn text_decorations(
    package: &IWorkPackage,
    first_style_id: u64,
) -> Result<TextDecorations> {
    let (underline, strikethrough) = walk(
        package,
        first_style_id,
        (None, None),
        |(underline, strikethrough), style| {
            if let Some(properties) = style.char_properties.as_ref() {
                if underline.is_none() {
                    *underline = properties
                        .underline
                        .map(TextUnderline::from_native_value)
                        .transpose()?;
                }
                if strikethrough.is_none() {
                    *strikethrough = properties
                        .strikethru
                        .map(TextStrikethrough::from_native_value)
                        .transpose()?;
                }
            }
            Ok(if underline.is_some() && strikethrough.is_some() {
                InheritanceControl::Complete
            } else {
                InheritanceControl::Continue
            })
        },
    )?;
    Ok(TextDecorations::new(
        underline.unwrap_or_default(),
        strikethrough.unwrap_or_default(),
    ))
}

pub(super) fn text_color(package: &IWorkPackage, first_style_id: u64) -> Result<RgbaColor> {
    let value = walk(package, first_style_id, None, |value, style| {
        let Some(properties) = style.char_properties.as_ref() else {
            return Ok(InheritanceControl::Continue);
        };
        let Some(color) = text_color_from_character(properties)? else {
            return Ok(InheritanceControl::Continue);
        };
        *value = Some(color);
        Ok(InheritanceControl::Complete)
    })?;
    Ok(value.unwrap_or_else(RgbaColor::black))
}

pub(super) fn text_capitalization(
    package: &IWorkPackage,
    first_style_id: u64,
) -> Result<TextCapitalization> {
    let value = walk(package, first_style_id, None, |value, style| {
        let Some(properties) = style.char_properties.as_ref() else {
            return Ok(InheritanceControl::Continue);
        };
        let Some(capitalization) = capitalization_from_character(properties)? else {
            return Ok(InheritanceControl::Continue);
        };
        *value = Some(capitalization);
        Ok(InheritanceControl::Complete)
    })?;
    Ok(value.unwrap_or_default())
}

pub(super) fn text_script(package: &IWorkPackage, first_style_id: u64) -> Result<TextScript> {
    let value = walk(package, first_style_id, None, |value, style| {
        let Some(script) = style
            .char_properties
            .as_ref()
            .and_then(|properties| properties.superscript)
        else {
            return Ok(InheritanceControl::Continue);
        };
        *value = Some(TextScript::from_native_value(script)?);
        Ok(InheritanceControl::Complete)
    })?;
    Ok(value.unwrap_or_default())
}

pub(super) fn text_baseline_shift(
    package: &IWorkPackage,
    first_style_id: u64,
) -> Result<TextBaselineShift> {
    let value = walk(package, first_style_id, None, |value, style| {
        let Some(shift) = style
            .char_properties
            .as_ref()
            .and_then(|properties| properties.baseline_shift)
        else {
            return Ok(InheritanceControl::Continue);
        };
        *value = Some(TextBaselineShift::from_points(shift)?);
        Ok(InheritanceControl::Complete)
    })?;
    Ok(value.unwrap_or_default())
}

pub(super) fn text_character_spacing(
    package: &IWorkPackage,
    first_style_id: u64,
) -> Result<TextCharacterSpacing> {
    let value = walk(package, first_style_id, None, |value, style| {
        let Some(spacing) = style
            .char_properties
            .as_ref()
            .and_then(|properties| properties.tracking)
        else {
            return Ok(InheritanceControl::Continue);
        };
        *value = Some(TextCharacterSpacing::from_native_ratio(spacing)?);
        Ok(InheritanceControl::Complete)
    })?;
    Ok(value.unwrap_or_default())
}

pub(super) fn text_ligatures(package: &IWorkPackage, first_style_id: u64) -> Result<TextLigatures> {
    let value = walk(package, first_style_id, None, |value, style| {
        let Some(ligatures) = style
            .char_properties
            .as_ref()
            .and_then(|properties| properties.ligatures)
        else {
            return Ok(InheritanceControl::Continue);
        };
        *value = Some(TextLigatures::from_native_value(ligatures)?);
        Ok(InheritanceControl::Complete)
    })?;
    Ok(value.unwrap_or_default())
}

pub(super) fn text_outline(package: &IWorkPackage, first_style_id: u64) -> Result<Outline> {
    let value = walk(package, first_style_id, None, |value, style| {
        let Some(properties) = style.char_properties.as_ref() else {
            return Ok(InheritanceControl::Continue);
        };
        let Some(outline) = text_outline_from_character(properties)? else {
            return Ok(InheritanceControl::Continue);
        };
        *value = Some(outline);
        Ok(InheritanceControl::Complete)
    })?;
    Ok(value.unwrap_or_default())
}

pub(super) fn text_shadow(package: &IWorkPackage, first_style_id: u64) -> Result<Shadow> {
    let value = walk(package, first_style_id, None, |value, style| {
        let Some(properties) = style.char_properties.as_ref() else {
            return Ok(InheritanceControl::Continue);
        };
        let Some(shadow) = text_shadow_from_character(properties)? else {
            return Ok(InheritanceControl::Continue);
        };
        *value = Some(shadow);
        Ok(InheritanceControl::Complete)
    })?;
    Ok(value.unwrap_or_default())
}

pub(super) fn text_background(
    package: &IWorkPackage,
    first_style_id: u64,
) -> Result<Background> {
    let value = walk(package, first_style_id, None, |value, style| {
        let Some(properties) = style.char_properties.as_ref() else {
            return Ok(InheritanceControl::Continue);
        };
        let Some(background) = text_background_from_character(properties)? else {
            return Ok(InheritanceControl::Continue);
        };
        *value = Some(background);
        Ok(InheritanceControl::Complete)
    })?;
    Ok(value.unwrap_or_default())
}

pub(super) fn paragraph_background(
    package: &IWorkPackage,
    first_style_id: u64,
) -> Result<ParagraphBackground> {
    let value = walk(package, first_style_id, None, |value, style| {
        let Some(properties) = style.para_properties.as_ref() else {
            return Ok(InheritanceControl::Continue);
        };
        let Some(background) = paragraph_background_from_properties(properties)? else {
            return Ok(InheritanceControl::Continue);
        };
        *value = Some(background);
        Ok(InheritanceControl::Complete)
    })?;
    Ok(value.unwrap_or_default())
}

pub(super) fn paragraph_borders(
    package: &IWorkPackage,
    first_style_id: u64,
) -> Result<Borders> {
    let value = walk(package, first_style_id, None, |value, style| {
        let Some(properties) = style.para_properties.as_ref() else {
            return Ok(InheritanceControl::Continue);
        };
        let Some(borders) = paragraph_borders_from_properties(properties)? else {
            return Ok(InheritanceControl::Continue);
        };
        *value = Some(borders);
        Ok(InheritanceControl::Complete)
    })?;
    Ok(value.unwrap_or_default())
}

pub(super) fn paragraph_flow(package: &IWorkPackage, first_style_id: u64) -> Result<ParagraphFlow> {
    let (hyphenation, keep_lines, keep_next, new_page, widow_orphan) = walk(
        package,
        first_style_id,
        (None, None, None, None, None),
        |(hyphenation, keep_lines, keep_next, new_page, widow_orphan), style| {
            if let Some(properties) = style.para_properties.as_ref() {
                if hyphenation.is_none() {
                    *hyphenation = properties.hyphenate.map(hyphenation_from_native);
                }
                if keep_lines.is_none() {
                    *keep_lines = properties.keep_lines_together;
                }
                if keep_next.is_none() {
                    *keep_next = properties.keep_with_next;
                }
                if new_page.is_none() {
                    *new_page = properties.page_break_before;
                }
                if widow_orphan.is_none() {
                    *widow_orphan = properties.widow_control;
                }
            }
            Ok(
                if hyphenation.is_some()
                    && keep_lines.is_some()
                    && keep_next.is_some()
                    && new_page.is_some()
                    && widow_orphan.is_some()
                {
                    InheritanceControl::Complete
                } else {
                    InheritanceControl::Continue
                },
            )
        },
    )?;
    let defaults = ParagraphFlow::default();
    Ok(ParagraphFlow::new()
        .with_hyphenation(hyphenation.unwrap_or(defaults.hyphenation()))
        .with_keep_lines_together(keep_lines.unwrap_or(defaults.keeps_lines_together()))
        .with_keep_with_next(keep_next.unwrap_or(defaults.keeps_with_next()))
        .with_start_on_new_page(new_page.unwrap_or(defaults.starts_on_new_page()))
        .with_prevent_widow_orphan_lines(
            widow_orphan.unwrap_or(defaults.prevents_widow_orphan_lines()),
        ))
}

pub(super) fn paragraph_writing_direction(
    package: &IWorkPackage,
    first_style_id: u64,
) -> Result<ParagraphWritingDirection> {
    let value = walk(package, first_style_id, None, |value, style| {
        let Some(direction) = style
            .char_properties
            .as_ref()
            .and_then(|properties| properties.writing_direction)
        else {
            return Ok(InheritanceControl::Continue);
        };
        *value = Some(writing_direction_from_native(direction)?);
        Ok(InheritanceControl::Complete)
    })?;
    Ok(value.unwrap_or_default())
}

pub(super) fn alignment(package: &IWorkPackage, first_style_id: u64) -> Result<Alignment> {
    let value = walk(package, first_style_id, None, |value, style| {
        let Some(alignment) = style
            .para_properties
            .as_ref()
            .and_then(|properties| properties.alignment)
        else {
            return Ok(InheritanceControl::Continue);
        };
        *value = Some(super::alignment_from_native(alignment)?);
        Ok(InheritanceControl::Complete)
    })?;
    Ok(value.unwrap_or(Alignment::Natural))
}

pub(super) fn line_spacing(
    package: &IWorkPackage,
    first_style_id: u64,
) -> Result<LineSpacing> {
    let value = walk(package, first_style_id, None, |value, style| {
        let Some(properties) = style.para_properties.as_ref() else {
            return Ok(InheritanceControl::Continue);
        };
        if properties.line_spacing_null == Some(true) {
            *value = Some(LineSpacing::default());
            return Ok(InheritanceControl::Complete);
        }
        let Some(spacing) = properties.line_spacing.as_ref() else {
            return Ok(InheritanceControl::Continue);
        };
        *value = Some(line_spacing_from_archive(spacing)?);
        Ok(InheritanceControl::Complete)
    })?;
    Ok(value.unwrap_or_default())
}

pub(super) fn spacing(package: &IWorkPackage, first_style_id: u64) -> Result<Spacing> {
    let (before, after) = walk(
        package,
        first_style_id,
        (None, None),
        |(before, after), style| {
            if let Some(properties) = style.para_properties.as_ref() {
                if before.is_none() {
                    *before = properties
                        .space_before
                        .map(SpacingPoints::from_points)
                        .transpose()?;
                }
                if after.is_none() {
                    *after = properties
                        .space_after
                        .map(SpacingPoints::from_points)
                        .transpose()?;
                }
            }
            Ok(if before.is_some() && after.is_some() {
                InheritanceControl::Complete
            } else {
                InheritanceControl::Continue
            })
        },
    )?;
    Ok(Spacing::new(
        before.unwrap_or(SpacingPoints::ZERO),
        after.unwrap_or(SpacingPoints::ZERO),
    ))
}

pub(super) fn indents(package: &IWorkPackage, first_style_id: u64) -> Result<Indents> {
    let (first_line, left, right) = walk(
        package,
        first_style_id,
        (None, None, None),
        |(first_line, left, right), style| {
            if let Some(properties) = style.para_properties.as_ref() {
                if first_line.is_none() {
                    *first_line = properties
                        .first_line_indent
                        .map(IndentPoints::from_points)
                        .transpose()?;
                }
                if left.is_none() {
                    *left = properties
                        .left_indent
                        .map(IndentPoints::from_points)
                        .transpose()?;
                }
                if right.is_none() {
                    *right = properties
                        .right_indent
                        .map(IndentPoints::from_points)
                        .transpose()?;
                }
            }
            Ok(
                if first_line.is_some() && left.is_some() && right.is_some() {
                    InheritanceControl::Complete
                } else {
                    InheritanceControl::Continue
                },
            )
        },
    )?;
    Ok(Indents::new(
        first_line.unwrap_or(IndentPoints::ZERO),
        left.unwrap_or(IndentPoints::ZERO),
        right.unwrap_or(IndentPoints::ZERO),
    ))
}

pub(super) fn tab_stops(package: &IWorkPackage, first_style_id: u64) -> Result<ParagraphTabStops> {
    let value = walk(package, first_style_id, None, |value, style| {
        let Some(properties) = style.para_properties.as_ref() else {
            return Ok(InheritanceControl::Continue);
        };
        if properties.tabs_null == Some(true) {
            if properties.tabs.is_some() {
                return Err(Error::InvalidFormat(
                    "native iWork paragraph tabs are both null and populated".to_owned(),
                ));
            }
            *value = Some(ParagraphTabStops::default());
            return Ok(InheritanceControl::Complete);
        }
        let Some(archive) = properties.tabs.as_ref() else {
            return Ok(InheritanceControl::Continue);
        };
        *value = Some(tabs::from_archive(archive)?);
        Ok(InheritanceControl::Complete)
    })?;
    Ok(value.unwrap_or_default())
}

pub(super) fn decimal_tab_character(
    package: &IWorkPackage,
    first_style_id: u64,
) -> Result<ParagraphDecimalTabCharacter> {
    let value = walk(package, first_style_id, None, |value, style| {
        let Some(properties) = style.para_properties.as_ref() else {
            return Ok(InheritanceControl::Continue);
        };
        if properties.decimal_tab_null == Some(true) {
            if properties.decimal_tab.is_some() {
                return Err(Error::InvalidFormat(
                    "native iWork decimal tab is both null and populated".to_owned(),
                ));
            }
            *value = Some(ParagraphDecimalTabCharacter::default());
            return Ok(InheritanceControl::Complete);
        }
        let Some(character) = properties.decimal_tab.as_deref() else {
            return Ok(InheritanceControl::Continue);
        };
        *value = Some(decimal_character_from_native(character)?);
        Ok(InheritanceControl::Complete)
    })?;
    Ok(value.unwrap_or_default())
}

pub(super) fn default_tab_interval(
    package: &IWorkPackage,
    first_style_id: u64,
) -> Result<ParagraphDefaultTabInterval> {
    let value = walk(package, first_style_id, None, |value, style| {
        let Some(points) = style
            .para_properties
            .as_ref()
            .and_then(|properties| properties.default_tab_stops)
        else {
            return Ok(InheritanceControl::Continue);
        };
        *value = Some(ParagraphDefaultTabInterval::from_points(points)?);
        Ok(InheritanceControl::Complete)
    })?;
    Ok(value.unwrap_or_default())
}

pub(super) fn following_style(
    package: &IWorkPackage,
    first_style_id: u64,
) -> Result<ParagraphFollowingStyle> {
    let value = walk(package, first_style_id, None, |value, style| {
        let Some(properties) = style.para_properties.as_ref() else {
            return Ok(InheritanceControl::Continue);
        };
        if properties.following_style_null == Some(true) {
            if properties.following_style.is_some() {
                return Err(Error::InvalidFormat(
                    "native iWork following paragraph style is both null and populated".to_owned(),
                ));
            }
            *value = Some(ParagraphFollowingStyle::Same);
            return Ok(InheritanceControl::Complete);
        }
        let Some(reference) = properties.following_style else {
            return Ok(InheritanceControl::Continue);
        };
        *value = Some(ParagraphFollowingStyle::Named(from_native_id(
            reference.identifier,
        )?));
        Ok(InheritanceControl::Complete)
    })?;
    Ok(value.unwrap_or_default())
}

fn walk<T, F>(package: &IWorkPackage, first_style_id: u64, mut state: T, mut visit: F) -> Result<T>
where
    F: FnMut(&mut T, &tswp::ParagraphStyleArchive) -> Result<InheritanceControl>,
{
    let mut visited = HashSet::new();
    let mut style_id = Some(first_style_id);
    for _ in 0..MAX_STYLE_INHERITANCE_DEPTH {
        let Some(identifier) = style_id else {
            return Ok(state);
        };
        if !visited.insert(identifier) {
            return Err(Error::InvalidFormat(format!(
                "iWork paragraph style inheritance cycles at {identifier}"
            )));
        }
        let location = locate_style(package, identifier)?;
        if visit(&mut state, &location.style)? == InheritanceControl::Complete {
            return Ok(state);
        }
        style_id = location.style.super_.parent.map(|parent| parent.identifier);
    }
    Err(Error::InvalidFormat(format!(
        "iWork paragraph style inheritance exceeds {MAX_STYLE_INHERITANCE_DEPTH} levels"
    )))
}
