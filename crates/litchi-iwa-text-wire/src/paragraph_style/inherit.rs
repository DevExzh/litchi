//! Effective property values through a paragraph style's parent chain.

use std::collections::HashSet;

use litchi_iwa_common::color::Rgba;
use litchi_iwa_text::character::{TextBaselineShift, TextDecorations, TextPointSize, TextStyle};
use litchi_iwa_text::paragraph::flow::Flow;
use litchi_iwa_text::paragraph::format::{
    Alignment, IndentPoints, Indents, LineSpacing, Spacing, SpacingPoints,
};
use litchi_iwa_text::paragraph::style::{ParagraphFollowingStyle, raw};
use litchi_iwa_text::paragraph::tabs::{DecimalCharacter, DefaultInterval, Stops};

use super::{Property, PropertyKind, Styles, convert};
use crate::native::{self, Budget, NativeError, NativeResult};
use crate::values;

/// The host's inheritance depth ceiling.
pub(super) const MAX_DEPTH: usize = 64;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Control {
    Continue,
    Complete,
}

fn transpose<T, E>(value: Option<Result<T, E>>, reason: &'static str) -> NativeResult<Option<T>> {
    value.transpose().map_err(|_| NativeError::Invalid(reason))
}

/// Decode the paragraph style `identifier` from `styles`.
pub(super) fn locate<'s>(
    styles: &'s (impl Styles + ?Sized),
    identifier: u64,
    budget: &mut Budget,
) -> NativeResult<(native::ParagraphStyle, &'s [u8])> {
    let Some(payload) = styles.paragraph_style(identifier) else {
        return Err(NativeError::Invalid("iWork paragraph style is missing"));
    };
    Ok((native::ParagraphStyle::decode(payload, budget)?, payload))
}

fn walk<T>(
    styles: &(impl Styles + ?Sized),
    first: u64,
    budget: &mut Budget,
    mut state: T,
    mut visit: impl FnMut(&mut T, &native::ParagraphStyle) -> NativeResult<Control>,
) -> NativeResult<T> {
    let mut visited = HashSet::new();
    let mut style = Some(first);
    for _ in 0..MAX_DEPTH {
        let Some(identifier) = style else {
            return Ok(state);
        };
        if !visited.insert(identifier) {
            return Err(NativeError::Invalid(
                "iWork paragraph style inheritance cycles",
            ));
        }
        let (decoded, _) = locate(styles, identifier, budget)?;
        if visit(&mut state, &decoded)? == Control::Complete {
            return Ok(state);
        }
        style = decoded.base.parent.map(|parent| parent.identifier);
    }
    Err(NativeError::Invalid(
        "iWork paragraph style inheritance exceeds 64 levels",
    ))
}

/// The first value `read` finds on the chain, or `None`.
fn first<T>(
    styles: &(impl Styles + ?Sized),
    style: u64,
    budget: &mut Budget,
    mut read: impl FnMut(&native::ParagraphStyle) -> NativeResult<Option<T>>,
) -> NativeResult<Option<T>> {
    walk(styles, style, budget, None, |value, decoded| {
        let Some(found) = read(decoded)? else {
            return Ok(Control::Continue);
        };
        *value = Some(found);
        Ok(Control::Complete)
    })
}

fn character<T>(
    read: impl Fn(&native::CharacterProperties) -> NativeResult<Option<T>>,
) -> impl FnMut(&native::ParagraphStyle) -> NativeResult<Option<T>> {
    move |style| style.character.as_ref().map_or(Ok(None), &read)
}

fn paragraph<T>(
    read: impl Fn(&native::ParagraphProperties) -> NativeResult<Option<T>>,
) -> impl FnMut(&native::ParagraphStyle) -> NativeResult<Option<T>> {
    move |style| style.paragraph.as_ref().map_or(Ok(None), &read)
}

#[allow(clippy::too_many_lines, reason = "one arm per property family")]
pub(super) fn effective(
    styles: &(impl Styles + ?Sized),
    style: u64,
    kind: PropertyKind,
    budget: &mut Budget,
) -> NativeResult<Property> {
    Ok(match kind {
        PropertyKind::TextStyle => {
            let (point_size, bold, italic) = walk(
                styles,
                style,
                budget,
                (None, None, None),
                |(point_size, bold, italic), decoded| {
                    if let Some(properties) = decoded.character.as_ref() {
                        if point_size.is_none() {
                            *point_size = transpose(
                                properties.font_size.map(TextPointSize::from_points),
                                "native iWork point size",
                            )?;
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
                            Control::Complete
                        } else {
                            Control::Continue
                        },
                    )
                },
            )?;
            Property::TextStyle(
                TextStyle::new(point_size.unwrap_or_default())
                    .with_bold(bold.unwrap_or(false))
                    .with_italic(italic.unwrap_or(false)),
            )
        },
        PropertyKind::TextFont => Property::TextFont(
            first(styles, style, budget, character(convert::text_font))?.unwrap_or_default(),
        ),
        PropertyKind::TextDecorations => {
            let (underline, strikethrough) = walk(
                styles,
                style,
                budget,
                (None, None),
                |(underline, strikethrough), decoded| {
                    if let Some(properties) = decoded.character.as_ref() {
                        if underline.is_none() {
                            *underline = properties
                                .underline
                                .map(values::underline_from_native)
                                .transpose()?;
                        }
                        if strikethrough.is_none() {
                            *strikethrough = properties
                                .strikethru
                                .map(values::strikethrough_from_native)
                                .transpose()?;
                        }
                    }
                    Ok(if underline.is_some() && strikethrough.is_some() {
                        Control::Complete
                    } else {
                        Control::Continue
                    })
                },
            )?;
            Property::TextDecorations(TextDecorations::new(
                underline.unwrap_or_default(),
                strikethrough.unwrap_or_default(),
            ))
        },
        PropertyKind::TextColor => Property::TextColor(
            first(styles, style, budget, character(convert::text_color))?
                .unwrap_or_else(Rgba::black),
        ),
        PropertyKind::TextCapitalization => Property::TextCapitalization(
            first(styles, style, budget, character(convert::capitalization))?.unwrap_or_default(),
        ),
        PropertyKind::TextScript => Property::TextScript(
            first(
                styles,
                style,
                budget,
                character(|properties| {
                    properties
                        .superscript
                        .map(values::script_from_native)
                        .transpose()
                }),
            )?
            .unwrap_or_default(),
        ),
        PropertyKind::TextBaselineShift => Property::TextBaselineShift(
            first(
                styles,
                style,
                budget,
                character(|properties| {
                    transpose(
                        properties
                            .baseline_shift
                            .map(TextBaselineShift::from_points),
                        "native iWork baseline shift",
                    )
                }),
            )?
            .unwrap_or_default(),
        ),
        PropertyKind::TextCharacterSpacing => Property::TextCharacterSpacing(
            first(
                styles,
                style,
                budget,
                character(|properties| {
                    properties
                        .tracking
                        .map(values::character_spacing_from_native)
                        .transpose()
                }),
            )?
            .unwrap_or_default(),
        ),
        PropertyKind::TextLigatures => Property::TextLigatures(
            first(
                styles,
                style,
                budget,
                character(|properties| {
                    properties
                        .ligatures
                        .map(values::ligatures_from_native)
                        .transpose()
                }),
            )?
            .unwrap_or_default(),
        ),
        PropertyKind::TextOutline => Property::TextOutline(
            first(styles, style, budget, character(convert::text_outline))?.unwrap_or_default(),
        ),
        PropertyKind::TextShadow => Property::TextShadow(
            first(styles, style, budget, character(convert::text_shadow))?.unwrap_or_default(),
        ),
        PropertyKind::TextBackground => Property::TextBackground(
            first(styles, style, budget, character(convert::text_background))?.unwrap_or_default(),
        ),
        PropertyKind::ParagraphBackground => Property::ParagraphBackground(
            first(
                styles,
                style,
                budget,
                paragraph(convert::paragraph_background),
            )?
            .unwrap_or_default(),
        ),
        PropertyKind::Borders => Property::Borders(
            first(styles, style, budget, paragraph(convert::paragraph_borders))?
                .unwrap_or_default(),
        ),
        PropertyKind::Flow => {
            let (hyphenation, keep_lines, keep_next, new_page, widow_orphan) = walk(
                styles,
                style,
                budget,
                (None, None, None, None, None),
                |(hyphenation, keep_lines, keep_next, new_page, widow_orphan), decoded| {
                    if let Some(properties) = decoded.paragraph.as_ref() {
                        if hyphenation.is_none() {
                            *hyphenation =
                                properties.hyphenate.map(values::hyphenation_from_native);
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
                            Control::Complete
                        } else {
                            Control::Continue
                        },
                    )
                },
            )?;
            let defaults = Flow::default();
            Property::Flow(
                Flow::new()
                    .with_hyphenation(hyphenation.unwrap_or(defaults.hyphenation()))
                    .with_keep_lines_together(keep_lines.unwrap_or(defaults.keeps_lines_together()))
                    .with_keep_with_next(keep_next.unwrap_or(defaults.keeps_with_next()))
                    .with_start_on_new_page(new_page.unwrap_or(defaults.starts_on_new_page()))
                    .with_prevent_widow_orphan_lines(
                        widow_orphan.unwrap_or(defaults.prevents_widow_orphan_lines()),
                    ),
            )
        },
        PropertyKind::WritingDirection => Property::WritingDirection(
            first(
                styles,
                style,
                budget,
                character(|properties| {
                    properties
                        .writing_direction
                        .map(values::writing_direction_from_native)
                        .transpose()
                }),
            )?
            .unwrap_or_default(),
        ),
        PropertyKind::Alignment => Property::Alignment(
            first(
                styles,
                style,
                budget,
                paragraph(|properties| {
                    properties
                        .alignment
                        .map(values::alignment_from_native)
                        .transpose()
                }),
            )?
            .unwrap_or(Alignment::Natural),
        ),
        PropertyKind::LineSpacing => Property::LineSpacing(
            first(
                styles,
                style,
                budget,
                paragraph(|properties| {
                    if properties.line_spacing_null == Some(true) {
                        return Ok(Some(LineSpacing::default()));
                    }
                    properties
                        .line_spacing
                        .as_ref()
                        .map(values::line_spacing_from_native)
                        .transpose()
                }),
            )?
            .unwrap_or_default(),
        ),
        PropertyKind::Spacing => {
            let (before, after) = walk(
                styles,
                style,
                budget,
                (None, None),
                |(before, after), decoded| {
                    if let Some(properties) = decoded.paragraph.as_ref() {
                        if before.is_none() {
                            *before = transpose(
                                properties.space_before.map(SpacingPoints::from_points),
                                "native iWork paragraph spacing",
                            )?;
                        }
                        if after.is_none() {
                            *after = transpose(
                                properties.space_after.map(SpacingPoints::from_points),
                                "native iWork paragraph spacing",
                            )?;
                        }
                    }
                    Ok(if before.is_some() && after.is_some() {
                        Control::Complete
                    } else {
                        Control::Continue
                    })
                },
            )?;
            Property::Spacing(Spacing::new(
                before.unwrap_or(SpacingPoints::ZERO),
                after.unwrap_or(SpacingPoints::ZERO),
            ))
        },
        PropertyKind::Indents => {
            let (first_line, left, right) = walk(
                styles,
                style,
                budget,
                (None, None, None),
                |(first_line, left, right), decoded| {
                    if let Some(properties) = decoded.paragraph.as_ref() {
                        let indent = |value: Option<f32>| {
                            transpose(
                                value.map(IndentPoints::from_points),
                                "native iWork paragraph indent",
                            )
                        };
                        if first_line.is_none() {
                            *first_line = indent(properties.first_line_indent)?;
                        }
                        if left.is_none() {
                            *left = indent(properties.left_indent)?;
                        }
                        if right.is_none() {
                            *right = indent(properties.right_indent)?;
                        }
                    }
                    Ok(
                        if first_line.is_some() && left.is_some() && right.is_some() {
                            Control::Complete
                        } else {
                            Control::Continue
                        },
                    )
                },
            )?;
            Property::Indents(Indents::new(
                first_line.unwrap_or(IndentPoints::ZERO),
                left.unwrap_or(IndentPoints::ZERO),
                right.unwrap_or(IndentPoints::ZERO),
            ))
        },
        PropertyKind::TabStops => Property::TabStops(
            first(
                styles,
                style,
                budget,
                paragraph(|properties| {
                    if properties.tabs_null == Some(true) {
                        if properties.tabs.is_some() {
                            return Err(NativeError::Invalid(
                                "native iWork paragraph tabs are both null and populated",
                            ));
                        }
                        return Ok(Some(Stops::default()));
                    }
                    properties
                        .tabs
                        .as_ref()
                        .map(values::tabs_from_native)
                        .transpose()
                }),
            )?
            .unwrap_or_default(),
        ),
        PropertyKind::DecimalTabCharacter => Property::DecimalTabCharacter(
            first(
                styles,
                style,
                budget,
                paragraph(|properties| {
                    if properties.decimal_tab_null == Some(true) {
                        if properties.decimal_tab.is_some() {
                            return Err(NativeError::Invalid(
                                "native iWork decimal tab is both null and populated",
                            ));
                        }
                        return Ok(Some(DecimalCharacter::default()));
                    }
                    properties
                        .decimal_tab
                        .as_deref()
                        .map(values::decimal_character_from_native)
                        .transpose()
                }),
            )?
            .unwrap_or_default(),
        ),
        PropertyKind::DefaultTabInterval => Property::DefaultTabInterval(
            first(
                styles,
                style,
                budget,
                paragraph(|properties| {
                    transpose(
                        properties
                            .default_tab_stops
                            .map(DefaultInterval::from_points),
                        "native iWork default tab interval",
                    )
                }),
            )?
            .unwrap_or_default(),
        ),
        PropertyKind::FollowingStyle => Property::FollowingStyle(
            first(
                styles,
                style,
                budget,
                paragraph(|properties| {
                    if properties.following_style_null == Some(true) {
                        if properties.following_style.is_some() {
                            return Err(NativeError::Invalid(
                                "native iWork following paragraph style is both null and populated",
                            ));
                        }
                        return Ok(Some(ParagraphFollowingStyle::Same));
                    }
                    properties
                        .following_style
                        .map(|reference| {
                            raw::from_native_id(reference.identifier)
                                .map(ParagraphFollowingStyle::Named)
                                .map_err(|_| {
                                    NativeError::Invalid("native iWork following paragraph style")
                                })
                        })
                        .transpose()
                }),
            )?
            .unwrap_or_default(),
        ),
    })
}
