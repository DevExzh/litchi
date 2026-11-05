//! Typed DOCX property setters used by the unified write registry.
//!
//! The setters deliberately operate on the leaf's public direct-property
//! models.  The transaction owns the XML splice and preservation rules;
//! this module only performs the scalar-to-model conversion that the
//! registry has already admitted.

use litchi_docx::color::Theme;
use litchi_docx::paragraph::{LineSpacingRule, RunUnderline, RunUnderlineColor};
use litchi_docx::properties::{self as direct, Fonts, Indentation, Language};

use crate::error::{Code, Error, Result};
use crate::value::{Base, Length, LengthContext, Scalar, Scheme, parse_number};

/// A leaf-supported DOCX property operation.
///
/// `ParagraphText` and `RunText` are write-only registry metadata.  They are
/// applied by the text replacement methods on `document::Edit`; all other
/// variants are applied inside `Edit::format_paragraph` or `Edit::format_run`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum Setter {
    ParagraphText,
    /// A hyperlink's display text (record 0251a).
    HyperlinkText,
    /// A table cell's text, replacing its content with one paragraph.
    CellText,
    /// A new cell's preferred width, `w:tcW` (record 0345).
    CellWidth,
    /// A grid column's width in twips (record 0345).
    ColumnWidth,
    /// A new column's cell text (record 0345).
    ColumnText,
    /// A hyperlink's external target (record 0251b).
    HyperlinkUrl,
    /// A new hyperlink's bookmark target.
    HyperlinkAnchor,
    /// A hyperlink's screen tip, `w:tooltip` (record 0345).
    HyperlinkTooltip,
    /// A new table's row count.
    TableRows,
    /// A new table's column count.
    TableColumns,
    /// A new table's style identifier.
    TableStyle,
    /// A new table's comma-separated grid widths, in twips.
    TableColumnWidths,
    /// A new table's cell text: rows separated by `;`, cells by `,`.
    TableData,
    /// A new row's cell count.
    RowColumns,
    /// Every table border but the diagonals: `border.all` (record 0342).
    TableBorderAll,
    /// A table's top border.
    TableBorderTop,
    /// A table's bottom border.
    TableBorderBottom,
    /// A table's left (leading) border.
    TableBorderLeft,
    /// A table's right (trailing) border.
    TableBorderRight,
    /// A table's inside horizontal border, between rows.
    TableBorderHorizontal,
    /// A table's inside vertical border, between columns.
    TableBorderVertical,
    /// A row's height, without a rule (`auto`, record 0342).
    RowHeight,
    /// A row's exact height.
    RowHeightExact,
    /// A row's minimum height.
    RowHeightAtLeast,
    /// Whether a row repeats as the table header.
    RowHeader,
    /// Whether a row stays on one page.
    RowCantSplit,
    /// Whether a row's end mark is hidden.
    RowHidden,
    /// A paragraph's list kind (record 0251e); applied through the
    /// numbering part, not the paragraph's properties.
    ParagraphListStyle,
    /// A paragraph list's restart value (record 0251e).
    ParagraphListStart,
    /// A new picture's image source (record 0251d).
    PictureSrc,
    /// A new picture's width.
    PictureWidth,
    /// A new picture's height.
    PictureHeight,
    /// A new picture's alternative text.
    PictureAlt,
    /// A new picture's name.
    PictureName,
    /// A new comment's author (record 0251c).
    CommentAuthor,
    /// A new comment's author initials.
    CommentInitials,
    /// A new comment's date.
    CommentDate,
    /// A new comment's text.
    CommentText,
    /// A comment's resolved state, `w15:done` (record 0345).
    CommentDone,
    /// A new comment's reply parent, by comment id (record 0345).
    CommentParent,
    /// A new comment's range start: after this 1-based run, 0 for the
    /// paragraph start (record 0345).
    CommentRunStart,
    /// Whether a new comment has range markers (record 0345).
    CommentRange,
    ParagraphStyle,
    ParagraphAlign,
    ParagraphSpaceBefore,
    ParagraphSpaceAfter,
    ParagraphLineSpacing,
    ParagraphLineRule,
    ParagraphSpaceBeforeLines,
    ParagraphSpaceAfterLines,
    ParagraphSpaceBeforeAuto,
    ParagraphSpaceAfterAuto,
    ParagraphFirstLineIndent,
    ParagraphHangingIndent,
    ParagraphIndent,
    ParagraphRightIndent,
    ParagraphFirstLineChars,
    ParagraphHangingChars,
    ParagraphLeftChars,
    ParagraphRightChars,
    ParagraphKeepNext,
    ParagraphKeepLines,
    ParagraphPageBreakBefore,
    ParagraphWidowControl,
    ParagraphDirection,
    ParagraphContextualSpacing,
    ParagraphNumId,
    ParagraphNumLevel,
    ParagraphOutlineLevel,
    RunText,
    RunFontLatin,
    RunFontAscii,
    RunFontHAnsi,
    RunFontEastAsia,
    RunFontAsciiTheme,
    RunFontHAnsiTheme,
    RunFontEastAsiaTheme,
    RunFontComplexScriptTheme,
    RunFontHint,
    RunLanguageLatin,
    RunLanguageEastAsia,
    RunLanguageComplexScript,
    RunSize,
    RunBold,
    RunItalic,
    RunFontComplexScript,
    RunSizeComplexScript,
    RunBoldComplexScript,
    RunItalicComplexScript,
    RunColor,
    RunUnderline,
    RunUnderlineColor,
    RunStrike,
    RunHighlight,
    RunCaps,
    RunSmallCaps,
    RunDoubleStrike,
    RunVanish,
    RunOutline,
    RunShadow,
    RunEmboss,
    RunImprint,
    RunNoProof,
    RunDirection,
    RunSuperscript,
    RunSubscript,
    RunPosition,
    RunCharacterSpacing,
    RunStyle,
}

impl Setter {
    /// Whether this setter is handled by a text replacement operation rather
    /// than a direct-property formatting closure.
    pub(crate) const fn is_text(self) -> bool {
        matches!(
            self,
            Self::ParagraphText | Self::RunText | Self::HyperlinkText | Self::CellText
        )
    }

    /// Apply a paragraph property through the leaf's typed formatting model.
    pub(crate) fn apply_paragraph(
        self,
        paragraph: &mut direct::Paragraph,
        value: Option<Scalar<'_>>,
    ) -> Result<()> {
        match self {
            Self::ParagraphText | Self::RunText | Self::HyperlinkText | Self::CellText => {
                Err(unsupported_text())
            },
            Self::ParagraphStyle => {
                paragraph.style = text(value)?.map(str::to_owned);
                Ok(())
            },
            Self::ParagraphAlign => {
                paragraph.justification = token(value)?.map(parse_justification).transpose()?;
                Ok(())
            },
            Self::ParagraphSpaceBefore => {
                let value = value
                    .map(length)
                    .transpose()?
                    .map(twips)
                    .transpose()?
                    .map(|value| value.0);
                if let Some(spacing) = paragraph.spacing.as_mut() {
                    spacing.before = value.map(u64::try_from).transpose().map_err(|_| {
                        invalid("paragraph spacing before is outside Word's integer range")
                    })?;
                    clean_spacing(paragraph);
                } else if let Some(value) = value {
                    paragraph.spacing =
                        Some(litchi_docx::paragraph::ParagraphSpacing {
                            before: Some(u64::try_from(value).map_err(|_| {
                                invalid("paragraph spacing before cannot be negative")
                            })?),
                            ..Default::default()
                        });
                }
                Ok(())
            },
            Self::ParagraphSpaceAfter => {
                let value = value
                    .map(length)
                    .transpose()?
                    .map(twips)
                    .transpose()?
                    .map(|value| value.0);
                if let Some(spacing) = paragraph.spacing.as_mut() {
                    spacing.after = value.map(u64::try_from).transpose().map_err(|_| {
                        invalid("paragraph spacing after is outside Word's integer range")
                    })?;
                    clean_spacing(paragraph);
                } else if let Some(value) = value {
                    paragraph.spacing =
                        Some(litchi_docx::paragraph::ParagraphSpacing {
                            after: Some(u64::try_from(value).map_err(|_| {
                                invalid("paragraph spacing after cannot be negative")
                            })?),
                            ..Default::default()
                        });
                }
                Ok(())
            },
            Self::ParagraphLineSpacing => apply_line_spacing(paragraph, value),
            Self::ParagraphLineRule => {
                let value = token(value)?.map(parse_line_rule).transpose()?;
                if let Some(spacing) = paragraph.spacing.as_mut() {
                    spacing.line_rule = value;
                    clean_spacing(paragraph);
                } else if let Some(value) = value {
                    paragraph.spacing = Some(litchi_docx::paragraph::ParagraphSpacing {
                        line_rule: Some(value),
                        ..Default::default()
                    });
                }
                Ok(())
            },
            Self::ParagraphSpaceBeforeLines => {
                set_spacing_i32(paragraph, value, |spacing, value| {
                    spacing.before_lines = value
                })
            },
            Self::ParagraphSpaceAfterLines => {
                set_spacing_i32(paragraph, value, |spacing, value| {
                    spacing.after_lines = value
                })
            },
            Self::ParagraphSpaceBeforeAuto => {
                set_spacing_bool(paragraph, value, |spacing, value| {
                    spacing.before_auto_spacing = value
                })
            },
            Self::ParagraphSpaceAfterAuto => {
                set_spacing_bool(paragraph, value, |spacing, value| {
                    spacing.after_auto_spacing = value
                })
            },
            Self::ParagraphFirstLineIndent => {
                set_indent_length(paragraph, value, |indentation, value| {
                    indentation.first_line = value
                })
            },
            Self::ParagraphHangingIndent => {
                set_indent_length(paragraph, value, |indentation, value| {
                    indentation.hanging = value
                })
            },
            Self::ParagraphIndent => {
                let value = value.map(length).transpose()?.map(twips).transpose()?;
                let indentation = paragraph.indentation.get_or_insert_default();
                indentation.left = value;
                indentation.start = None;
                if value.is_none() {
                    clean_indentation(paragraph);
                }
                Ok(())
            },
            Self::ParagraphRightIndent => {
                let value = value.map(length).transpose()?.map(twips).transpose()?;
                let indentation = paragraph.indentation.get_or_insert_default();
                indentation.right = value;
                indentation.end = None;
                if value.is_none() {
                    clean_indentation(paragraph);
                }
                Ok(())
            },
            Self::ParagraphFirstLineChars => {
                set_indent_i64(paragraph, value, |indentation, value| {
                    indentation.first_line_chars = value
                })
            },
            Self::ParagraphHangingChars => {
                set_indent_i64(paragraph, value, |indentation, value| {
                    indentation.hanging_chars = value
                })
            },
            Self::ParagraphLeftChars => {
                let value = value.map(int).transpose()?;
                let indentation = paragraph.indentation.get_or_insert_default();
                indentation.left_chars = value;
                indentation.start_chars = None;
                if value.is_none() {
                    clean_indentation(paragraph);
                }
                Ok(())
            },
            Self::ParagraphRightChars => {
                let value = value.map(int).transpose()?;
                let indentation = paragraph.indentation.get_or_insert_default();
                indentation.right_chars = value;
                indentation.end_chars = None;
                if value.is_none() {
                    clean_indentation(paragraph);
                }
                Ok(())
            },
            Self::ParagraphKeepNext => set_bool(&mut paragraph.keep_next, value),
            Self::ParagraphKeepLines => set_bool(&mut paragraph.keep_lines, value),
            Self::ParagraphPageBreakBefore => set_bool(&mut paragraph.page_break_before, value),
            Self::ParagraphWidowControl => set_bool(&mut paragraph.widow_control, value),
            Self::ParagraphDirection => {
                paragraph.bidi = token(value)?.map(parse_direction).transpose()?;
                Ok(())
            },
            Self::ParagraphContextualSpacing => set_bool(&mut paragraph.contextual_spacing, value),
            // Planned by `super::lists`; admission only checks the value.
            Self::ParagraphListStyle | Self::ParagraphListStart => Ok(()),
            Self::ParagraphNumId => {
                let value = text(value)?;
                let value = value
                    .map(|value| {
                        value.trim().parse::<u32>().map_err(|_| {
                            invalid("paragraph numId must be an unsigned 32-bit integer")
                        })
                    })
                    .transpose()?;
                let numbering = paragraph.numbering.get_or_insert_default();
                numbering.id = value;
                if value.is_none() && numbering.level.is_none() {
                    paragraph.numbering = None;
                }
                Ok(())
            },
            Self::ParagraphNumLevel => {
                let value = text(value)?;
                let value = value
                    .map(|value| {
                        value.trim().parse::<u8>().map_err(|_| {
                            invalid("paragraph numLevel must be an unsigned 8-bit integer")
                        })
                    })
                    .transpose()?;
                let numbering = paragraph.numbering.get_or_insert_default();
                numbering.level = value;
                if value.is_none() && numbering.id.is_none() {
                    paragraph.numbering = None;
                }
                Ok(())
            },
            Self::ParagraphOutlineLevel => {
                paragraph.outline_level = value
                    .map(int)
                    .transpose()?
                    .map(|value| {
                        u8::try_from(value)
                            .map_err(|_| invalid("paragraph outlineLvl must be between 0 and 9"))
                    })
                    .transpose()?;
                Ok(())
            },
            _ => Err(unsupported(
                "the setter does not target paragraph properties",
            )),
        }
    }

    /// Apply a run property through the leaf's typed formatting model.
    pub(crate) fn apply_run(self, run: &mut direct::Run, value: Option<Scalar<'_>>) -> Result<()> {
        match self {
            Self::ParagraphText | Self::RunText | Self::HyperlinkText | Self::CellText => {
                Err(unsupported_text())
            },
            Self::RunFontLatin => {
                let value = text(value)?.map(str::to_owned);
                let fonts = run.fonts.get_or_insert_default();
                fonts.ascii = value.clone();
                fonts.h_ansi = value;
                clean_fonts(run);
                Ok(())
            },
            Self::RunFontAscii => set_font_text(run, value, |fonts, value| fonts.ascii = value),
            Self::RunFontHAnsi => set_font_text(run, value, |fonts, value| fonts.h_ansi = value),
            Self::RunFontEastAsia => {
                set_font_text(run, value, |fonts, value| fonts.east_asia = value)
            },
            Self::RunFontComplexScript => {
                set_font_text(run, value, |fonts, value| fonts.complex_script = value)
            },
            Self::RunFontAsciiTheme => {
                set_font_theme(run, value, |fonts, value| fonts.ascii_theme = value)
            },
            Self::RunFontHAnsiTheme => {
                set_font_theme(run, value, |fonts, value| fonts.h_ansi_theme = value)
            },
            Self::RunFontEastAsiaTheme => {
                set_font_theme(run, value, |fonts, value| fonts.east_asia_theme = value)
            },
            Self::RunFontComplexScriptTheme => set_font_theme(run, value, |fonts, value| {
                fonts.complex_script_theme = value
            }),
            Self::RunFontHint => set_font_hint(run, value),
            Self::RunLanguageLatin => {
                set_language_text(run, value, |language, value| language.latin = value)
            },
            Self::RunLanguageEastAsia => {
                set_language_text(run, value, |language, value| language.east_asia = value)
            },
            Self::RunLanguageComplexScript => {
                set_language_text(run, value, |language, value| language.bidi = value)
            },
            Self::RunSize => {
                run.size = value
                    .map(length)
                    .transpose()?
                    .map(half_points)
                    .transpose()?;
                Ok(())
            },
            Self::RunBold => set_bool(&mut run.bold, value),
            Self::RunItalic => set_bool(&mut run.italic, value),
            Self::RunSizeComplexScript => {
                run.size_complex_script = value
                    .map(length)
                    .transpose()?
                    .map(half_points)
                    .transpose()?;
                Ok(())
            },
            Self::RunBoldComplexScript => set_bool(&mut run.bold_complex_script, value),
            Self::RunItalicComplexScript => set_bool(&mut run.italic_complex_script, value),
            Self::RunColor => {
                run.color = value.map(color).transpose()?;
                Ok(())
            },
            Self::RunUnderline => {
                run.underline = token(value)?
                    .map(parse_underline)
                    .transpose()?
                    .map(|style| RunUnderline {
                        style,
                        color: None,
                        theme_color: None,
                        theme_tint: None,
                        theme_shade: None,
                    });
                Ok(())
            },
            Self::RunUnderlineColor => set_underline_color(run, value),
            Self::RunStrike => set_bool(&mut run.strike, value),
            Self::RunHighlight => {
                run.highlight = token(value)?.map(parse_highlight).transpose()?;
                Ok(())
            },
            Self::RunCaps => set_bool(&mut run.caps, value),
            Self::RunSmallCaps => set_bool(&mut run.small_caps, value),
            Self::RunDoubleStrike => set_bool(&mut run.double_strike, value),
            Self::RunVanish => set_bool(&mut run.vanish, value),
            Self::RunOutline => set_bool(&mut run.outline, value),
            Self::RunShadow => set_bool(&mut run.shadow, value),
            Self::RunEmboss => set_bool(&mut run.emboss, value),
            Self::RunImprint => set_bool(&mut run.imprint, value),
            Self::RunNoProof => set_bool(&mut run.no_proof, value),
            Self::RunDirection => {
                run.right_to_left = token(value)?.map(parse_direction).transpose()?;
                Ok(())
            },
            Self::RunSuperscript => set_vertical(run, value, direct::VerticalAlign::Superscript),
            Self::RunSubscript => set_vertical(run, value, direct::VerticalAlign::Subscript),
            Self::RunPosition => {
                run.position = value
                    .map(length)
                    .transpose()?
                    .map(half_points)
                    .transpose()?;
                Ok(())
            },
            Self::RunCharacterSpacing => {
                run.character_spacing = value.map(length).transpose()?.map(twips).transpose()?;
                Ok(())
            },
            Self::RunStyle => {
                run.style = text(value)?.map(str::to_owned);
                Ok(())
            },
            _ => Err(unsupported("the setter does not target run properties")),
        }
    }
}

/// One side of a table's borders.
#[derive(Clone, Copy)]
enum Side {
    Top,
    Bottom,
    Left,
    Right,
    Horizontal,
    Vertical,
}

impl Setter {
    /// Apply a table property (record 0342): the style and the borders.
    pub(crate) fn apply_table(
        self,
        table: &mut direct::Table,
        value: Option<Scalar<'_>>,
    ) -> Result<()> {
        let sides: &[Side] = match self {
            Self::TableStyle => {
                table.style = text(value)?.map(str::to_owned);
                return Ok(());
            },
            Self::TableBorderAll => &[
                Side::Top,
                Side::Bottom,
                Side::Left,
                Side::Right,
                Side::Horizontal,
                Side::Vertical,
            ],
            Self::TableBorderTop => &[Side::Top],
            Self::TableBorderBottom => &[Side::Bottom],
            Self::TableBorderLeft => &[Side::Left],
            Self::TableBorderRight => &[Side::Right],
            Self::TableBorderHorizontal => &[Side::Horizontal],
            Self::TableBorderVertical => &[Side::Vertical],
            _ => return Err(unsupported("the setter does not target table properties")),
        };
        let border = text(value)?.map(parse_border).transpose()?;
        let borders = table.borders.get_or_insert_default();
        for side in sides {
            // A side written with its Strict spelling keeps it.
            let slot = match side {
                Side::Top => &mut borders.top,
                Side::Bottom => &mut borders.bottom,
                Side::Left if borders.left.is_none() && borders.start.is_some() => {
                    &mut borders.start
                },
                Side::Left => &mut borders.left,
                Side::Right if borders.right.is_none() && borders.end.is_some() => &mut borders.end,
                Side::Right => &mut borders.right,
                Side::Horizontal => &mut borders.inside_horizontal,
                Side::Vertical => &mut borders.inside_vertical,
            };
            *slot = border;
        }
        if table.borders == Some(direct::TableBorders::default()) {
            table.borders = None;
        }
        Ok(())
    }

    /// Apply a row property (record 0342): the height and the header,
    /// keep-together and hidden flags. A flag set to `false` is removed, as
    /// OfficeCLI removes it.
    pub(crate) fn apply_row(self, row: &mut direct::Row, value: Option<Scalar<'_>>) -> Result<()> {
        let flag = |value: Option<Scalar<'_>>| Ok::<_, Error>(bool_value(value)?.filter(|on| *on));
        match self {
            Self::RowHeight | Self::RowHeightExact | Self::RowHeightAtLeast => {
                let rule = match self {
                    Self::RowHeightExact => Some(direct::HeightRule::Exact),
                    Self::RowHeightAtLeast => Some(direct::HeightRule::AtLeast),
                    _ => None,
                };
                row.height = text(value)?
                    .map(|text| {
                        let length = Length::parse_in(text.trim(), LengthContext::WordSpacing)
                            .map_err(|_| {
                                invalid("a row height is twips or a length such as 1cm")
                            })?;
                        Ok::<_, Error>(direct::Height::new(twips(length)?, rule))
                    })
                    .transpose()?;
                Ok(())
            },
            Self::RowHeader => {
                row.header = flag(value)?;
                Ok(())
            },
            Self::RowCantSplit => {
                row.cant_split = flag(value)?;
                Ok(())
            },
            Self::RowHidden => {
                row.hidden = flag(value)?;
                Ok(())
            },
            _ => Err(unsupported("the setter does not target row properties")),
        }
    }
}

/// OfficeCLI's border value (`ParseBorderValue`): `STYLE[;SIZE[;COLOR[;SPACE]]]`
/// or the CSS-like `SIZE STYLE COLOR`. A bare size is eighths of a point,
/// `pt`, `cm` and `in` are converted; a missing size is 4 (12 for `thick`,
/// none for `nil`).
fn parse_border(value: &str) -> Result<direct::Border> {
    let value = value.trim();
    let (style, size, color, space) = if value.contains(';') {
        let mut parts = value.split(';').map(str::trim);
        let style = parts.next().unwrap_or_default();
        let size = parts.next().filter(|part| !part.is_empty());
        let color = parts.next().filter(|part| !part.is_empty());
        let space = parts.next().filter(|part| !part.is_empty());
        (style, size, color, space)
    } else {
        let (mut style, mut size, mut color) = (None, None, None);
        for token in value.split_whitespace() {
            if size.is_none() && border_size(token).is_ok() {
                size = Some(token);
            } else if color.is_none() && style.is_some() && hex(token).is_ok() {
                color = Some(token);
            } else if style.is_none() {
                style = Some(token);
            } else {
                return Err(invalid("a border is STYLE;SIZE;COLOR or SIZE STYLE COLOR"));
            }
        }
        (style.unwrap_or_default(), size, color, None)
    };
    let style = border_style(style)?;
    let mut border = direct::Border::new(style);
    border.size = match size {
        Some(size) => Some(direct::EighthPoints(border_size(size)?)),
        None if style.as_str() == "nil" => None,
        None if style.as_str() == "thick" => Some(direct::EighthPoints(12)),
        None => Some(direct::EighthPoints(4)),
    };
    border.color = color.map(hex).transpose()?;
    border.space = space
        .map(|space| {
            space
                .parse::<u32>()
                .map(|space| direct::Points(i64::from(space)))
                .map_err(|_| invalid("a border space is a whole number of points"))
        })
        .transpose()?;
    Ok(border)
}

/// A border style, ASCII case-insensitively, with OfficeCLI's aliases.
fn border_style(token: &str) -> Result<direct::BorderStyle> {
    const LINES: [&str; 27] = [
        "single",
        "thick",
        "double",
        "dotted",
        "dashed",
        "dotDash",
        "dotDotDash",
        "triple",
        "thinThickSmallGap",
        "thickThinSmallGap",
        "thinThickThinSmallGap",
        "thinThickMediumGap",
        "thickThinMediumGap",
        "thinThickThinMediumGap",
        "thinThickLargeGap",
        "thickThinLargeGap",
        "thinThickThinLargeGap",
        "wave",
        "doubleWave",
        "dashSmallGap",
        "dashDotStroked",
        "threeDEmboss",
        "threeDEngrave",
        "outset",
        "inset",
        "none",
        "nil",
    ];
    let lower = token.to_ascii_lowercase();
    let canonical = match lower.as_str() {
        "thin" => "single",
        "medium" => "thick",
        "wavy" => "wave",
        "wavydouble" | "wavyheavy" => "doubleWave",
        "3demboss" => "threeDEmboss",
        "3dengrave" => "threeDEngrave",
        _ => LINES
            .iter()
            .find(|line| line.eq_ignore_ascii_case(token))
            .copied()
            .unwrap_or(token),
    };
    direct::BorderStyle::from_xml(canonical).ok_or_else(|| invalid("unknown border style"))
}

/// A border width in eighths of a point.
fn border_size(token: &str) -> Result<i64> {
    if let Ok(eighths) = token.parse::<u32>() {
        return Ok(i64::from(eighths));
    }
    let lower = token.to_ascii_lowercase();
    let (number, per_unit) = if let Some(number) = lower.strip_suffix("pt") {
        (number, 8.0)
    } else if let Some(number) = lower.strip_suffix("cm") {
        (number, 72.0 / 2.54 * 8.0)
    } else if let Some(number) = lower.strip_suffix("in") {
        (number, 72.0 * 8.0)
    } else {
        return Err(invalid(
            "a border size is eighths of a point or a length such as 1pt",
        ));
    };
    let number: f64 = number
        .trim()
        .parse()
        .map_err(|_| invalid("a border size is eighths of a point or a length such as 1pt"))?;
    let eighths = (number * per_unit).round();
    if !(0.0..=f64::from(u32::MAX)).contains(&eighths) {
        return Err(invalid("a border size is out of range"));
    }
    #[expect(
        clippy::cast_possible_truncation,
        reason = "checked to be a whole number in u32 range above"
    )]
    Ok(eighths as i64)
}

/// A border colour: `auto` or six hexadecimal digits, with an optional `#`.
fn hex(token: &str) -> Result<direct::HexColor> {
    if token.eq_ignore_ascii_case("auto") {
        return Ok(direct::HexColor::Auto);
    }
    let digits = token.strip_prefix('#').unwrap_or(token);
    let byte = |index: usize| {
        digits
            .get(index..index + 2)
            .and_then(|pair| u8::from_str_radix(pair, 16).ok())
    };
    match (digits.len(), byte(0), byte(2), byte(4)) {
        (6, Some(red), Some(green), Some(blue)) => Ok(direct::HexColor::Rgb([red, green, blue])),
        _ => Err(invalid("a border colour is six hexadecimal digits or auto")),
    }
}

fn unsupported_text() -> Error {
    unsupported("text is applied by the leaf text replacement operation")
}

fn unsupported(message: &'static str) -> Error {
    Error::new(Code::UnsupportedProperty, message)
}

fn invalid(message: &'static str) -> Error {
    Error::new(Code::InvalidValue, message)
}

fn bool_value(value: Option<Scalar<'_>>) -> Result<Option<bool>> {
    match value {
        None => Ok(None),
        Some(Scalar::Bool(value)) => Ok(Some(value)),
        Some(_) => Err(invalid("expected a boolean value")),
    }
}

fn int(value: Scalar<'_>) -> Result<i64> {
    match value {
        Scalar::Int(value) => Ok(value),
        _ => Err(invalid("expected an integer value")),
    }
}

fn length(value: Scalar<'_>) -> Result<Length> {
    match value {
        Scalar::Length(value) => Ok(value),
        _ => Err(invalid("expected a length value")),
    }
}

fn text(value: Option<Scalar<'_>>) -> Result<Option<&str>> {
    match value {
        None => Ok(None),
        Some(Scalar::Text(value)) => Ok(Some(value)),
        Some(_) => Err(invalid("expected a text value")),
    }
}

fn token(value: Option<Scalar<'_>>) -> Result<Option<&'static str>> {
    match value {
        None => Ok(None),
        Some(Scalar::Token(value)) => Ok(Some(value)),
        Some(_) => Err(invalid("expected an enumeration value")),
    }
}

fn rounded_units(value: Length, unit: i64) -> Result<i64> {
    let emu = value.emu();
    let adjusted = if emu >= 0 {
        emu.checked_add(unit / 2)
    } else {
        emu.checked_sub(unit / 2)
    }
    .ok_or_else(|| invalid("the length is outside the DOCX integer range"))?;
    adjusted
        .checked_div(unit)
        .ok_or_else(|| invalid("the length is outside the DOCX integer range"))
}

pub(super) fn twips(value: Length) -> Result<direct::Twips> {
    Ok(direct::Twips(rounded_units(value, 635)?))
}

fn half_points(value: Length) -> Result<direct::HalfPoints> {
    Ok(direct::HalfPoints(rounded_units(value, 6_350)?))
}

fn set_bool(slot: &mut Option<bool>, value: Option<Scalar<'_>>) -> Result<()> {
    *slot = bool_value(value)?;
    Ok(())
}

fn set_spacing_i32(
    paragraph: &mut direct::Paragraph,
    value: Option<Scalar<'_>>,
    set: impl FnOnce(&mut litchi_docx::paragraph::ParagraphSpacing, Option<i32>),
) -> Result<()> {
    let value = value
        .map(int)
        .transpose()?
        .map(|value| {
            i32::try_from(value).map_err(|_| invalid("paragraph spacing lines exceed i32"))
        })
        .transpose()?;
    let spacing = paragraph.spacing.get_or_insert_default();
    set(spacing, value);
    if value.is_none() {
        clean_spacing(paragraph);
    }
    Ok(())
}

fn set_spacing_bool(
    paragraph: &mut direct::Paragraph,
    value: Option<Scalar<'_>>,
    set: impl FnOnce(&mut litchi_docx::paragraph::ParagraphSpacing, Option<bool>),
) -> Result<()> {
    let value = bool_value(value)?;
    let spacing = paragraph.spacing.get_or_insert_default();
    set(spacing, value);
    if value.is_none() {
        clean_spacing(paragraph);
    }
    Ok(())
}

fn set_indent_length(
    paragraph: &mut direct::Paragraph,
    value: Option<Scalar<'_>>,
    set: impl FnOnce(&mut Indentation, Option<direct::Twips>),
) -> Result<()> {
    let value = value.map(length).transpose()?.map(twips).transpose()?;
    let indentation = paragraph.indentation.get_or_insert_default();
    set(indentation, value);
    if value.is_none() {
        clean_indentation(paragraph);
    }
    Ok(())
}

fn set_indent_i64(
    paragraph: &mut direct::Paragraph,
    value: Option<Scalar<'_>>,
    set: impl FnOnce(&mut Indentation, Option<i64>),
) -> Result<()> {
    let value = value.map(int).transpose()?;
    let indentation = paragraph.indentation.get_or_insert_default();
    set(indentation, value);
    if value.is_none() {
        clean_indentation(paragraph);
    }
    Ok(())
}

fn set_font_text(
    run: &mut direct::Run,
    value: Option<Scalar<'_>>,
    set: impl FnOnce(&mut Fonts, Option<String>),
) -> Result<()> {
    let value = text(value)?.map(str::to_owned);
    let fonts = run.fonts.get_or_insert_default();
    set(fonts, value);
    clean_fonts(run);
    Ok(())
}

fn set_font_theme(
    run: &mut direct::Run,
    value: Option<Scalar<'_>>,
    set: impl FnOnce(&mut Fonts, Option<direct::ThemeFont>),
) -> Result<()> {
    let value = token(value)?.map(parse_theme_font).transpose()?;
    let fonts = run.fonts.get_or_insert_default();
    set(fonts, value);
    clean_fonts(run);
    Ok(())
}

fn set_font_hint(run: &mut direct::Run, value: Option<Scalar<'_>>) -> Result<()> {
    let value = token(value)?.map(parse_font_hint).transpose()?;
    let fonts = run.fonts.get_or_insert_default();
    fonts.hint = value;
    clean_fonts(run);
    Ok(())
}

fn set_language_text(
    run: &mut direct::Run,
    value: Option<Scalar<'_>>,
    set: impl FnOnce(&mut Language, Option<String>),
) -> Result<()> {
    let value = text(value)?.map(str::to_owned);
    let language = run.language.get_or_insert_default();
    set(language, value);
    clean_language(run);
    Ok(())
}

fn set_underline_color(run: &mut direct::Run, value: Option<Scalar<'_>>) -> Result<()> {
    let value = value.map(color_with_theme).transpose()?;
    let Some((color, theme)) = value else {
        if let Some(underline) = run.underline.as_mut() {
            underline.color = None;
            underline.theme_color = None;
            underline.theme_tint = None;
            underline.theme_shade = None;
        }
        return Ok(());
    };
    let underline = run.underline.get_or_insert(RunUnderline {
        style: litchi_docx::UnderlineStyle::Single,
        color: None,
        theme_color: None,
        theme_tint: None,
        theme_shade: None,
    });
    underline.color = Some(color);
    underline.theme_color = theme;
    underline.theme_tint = None;
    underline.theme_shade = None;
    Ok(())
}

fn set_vertical(
    run: &mut direct::Run,
    value: Option<Scalar<'_>>,
    placement: direct::VerticalAlign,
) -> Result<()> {
    let value = bool_value(value)?;
    match value {
        Some(true) => run.vertical_align = Some(placement),
        Some(false) | None if run.vertical_align == Some(placement) => run.vertical_align = None,
        Some(false) | None => {},
    }
    Ok(())
}

fn apply_line_spacing(paragraph: &mut direct::Paragraph, value: Option<Scalar<'_>>) -> Result<()> {
    let Some(value) = value else {
        if let Some(spacing) = paragraph.spacing.as_mut() {
            spacing.line = None;
            clean_spacing(paragraph);
        }
        return Ok(());
    };
    let Scalar::Text(value) = value else {
        return Err(invalid(
            "lineSpacing expects a value such as '1.15x' or '12pt'",
        ));
    };
    let value = value.trim();
    let (line, rule) = if let Some(number) = strip_suffix_ci(value, "x") {
        (line_auto(number, false)?, LineSpacingRule::Auto)
    } else if let Some(number) = value.strip_suffix('%') {
        (line_auto(number, true)?, LineSpacingRule::Auto)
    } else if has_suffix_ci(value, "pt") || has_suffix_ci(value, "cm") || has_suffix_ci(value, "in")
    {
        let length = Length::parse_in(value, LengthContext::WordSpacing)
            .map_err(|_| invalid("lineSpacing expects a non-negative Word length"))?;
        let line = i32::try_from(twips(length)?.0)
            .map_err(|_| invalid("lineSpacing is outside the DOCX integer range"))?;
        let rule = paragraph
            .spacing
            .and_then(|spacing| spacing.line_rule)
            .filter(|rule| matches!(rule, LineSpacingRule::AtLeast))
            .unwrap_or(LineSpacingRule::Exact);
        (line, rule)
    } else {
        (line_auto(value, false)?, LineSpacingRule::Auto)
    };
    let spacing = paragraph.spacing.get_or_insert_default();
    spacing.line = Some(line);
    spacing.line_rule = Some(rule);
    Ok(())
}

fn line_auto(value: &str, percent: bool) -> Result<i32> {
    let number = parse_number(value.trim())
        .map_err(|_| invalid("lineSpacing must contain a finite number"))?;
    if number == 0.0 {
        return Err(invalid("lineSpacing must not be zero"));
    }
    let units = if percent {
        number / 100.0 * 240.0
    } else {
        number * 240.0
    };
    let rounded = units.round_ties_even();
    if !rounded.is_finite() || !(i32::MIN as f64..=i32::MAX as f64).contains(&rounded) {
        return Err(invalid("lineSpacing is outside the DOCX integer range"));
    }
    Ok(rounded as i32)
}

fn strip_suffix_ci<'a>(value: &'a str, suffix: &str) -> Option<&'a str> {
    let value = value.trim();
    let start = value.len().checked_sub(suffix.len())?;
    value
        .get(start..)
        .filter(|tail| tail.eq_ignore_ascii_case(suffix))
        .map(|_| &value[..start])
}

fn has_suffix_ci(value: &str, suffix: &str) -> bool {
    strip_suffix_ci(value, suffix).is_some()
}

fn color(value: Scalar<'_>) -> Result<direct::Color> {
    let Scalar::Color(input) = value else {
        return Err(invalid("expected a Word color"));
    };
    let (wire, theme) = match input.base() {
        Base::Auto => (direct::HexColor::Auto, None),
        Base::Rgb(rgb) if rgb.is_opaque() => (
            direct::HexColor::Rgb([rgb.red(), rgb.green(), rgb.blue()]),
            None,
        ),
        Base::Scheme(scheme) => (direct::HexColor::Auto, Some(word_theme(scheme)?)),
        Base::Rgb(_) => return Err(invalid("Word colors do not support alpha")),
        Base::Theme(_) | Base::System(_) => {
            return Err(invalid("the color is outside Word's color domain"));
        },
    };
    if !input.transforms().is_empty() {
        return Err(invalid("Word colors do not support color transforms"));
    }
    Ok(direct::Color {
        value: wire,
        theme,
        theme_tint: None,
        theme_shade: None,
    })
}

fn color_with_theme(value: Scalar<'_>) -> Result<(RunUnderlineColor, Option<Theme>)> {
    let Scalar::Color(value) = value else {
        return Err(invalid("expected a Word color"));
    };
    if !value.transforms().is_empty() {
        return Err(invalid("Word underline colors do not support transforms"));
    }
    Ok(match value.base() {
        Base::Auto => (RunUnderlineColor::Auto, None),
        Base::Rgb(rgb) if rgb.is_opaque() => (
            RunUnderlineColor::Rgb([rgb.red(), rgb.green(), rgb.blue()]),
            None,
        ),
        Base::Scheme(scheme) => (RunUnderlineColor::Auto, Some(word_theme(scheme)?)),
        Base::Rgb(_) => return Err(invalid("Word colors do not support alpha")),
        Base::Theme(_) | Base::System(_) => {
            return Err(invalid("the color is outside Word's color domain"));
        },
    })
}

fn word_theme(value: Scheme) -> Result<Theme> {
    Ok(match value {
        Scheme::Dark1 => Theme::Dark1,
        Scheme::Light1 => Theme::Light1,
        Scheme::Dark2 => Theme::Dark2,
        Scheme::Light2 => Theme::Light2,
        Scheme::Accent1 => Theme::Accent1,
        Scheme::Accent2 => Theme::Accent2,
        Scheme::Accent3 => Theme::Accent3,
        Scheme::Accent4 => Theme::Accent4,
        Scheme::Accent5 => Theme::Accent5,
        Scheme::Accent6 => Theme::Accent6,
        Scheme::Hyperlink => Theme::Hyperlink,
        Scheme::FollowedHyperlink => Theme::FollowedHyperlink,
        Scheme::Background1 => Theme::Background1,
        Scheme::Text1 => Theme::Text1,
        Scheme::Background2 => Theme::Background2,
        Scheme::Text2 => Theme::Text2,
        Scheme::Placeholder => return Err(invalid("phClr is not a Word theme color")),
    })
}

fn parse_justification(value: &'static str) -> Result<direct::Justification> {
    if value == "justify" {
        return Ok(direct::Justification::Both);
    }
    direct::Justification::from_xml(value)
        .ok_or_else(|| invalid("the alignment token is outside Word's domain"))
}

fn parse_line_rule(value: &'static str) -> Result<LineSpacingRule> {
    LineSpacingRule::from_xml(value)
        .ok_or_else(|| invalid("the lineRule token is outside Word's domain"))
}

fn parse_theme_font(value: &'static str) -> Result<direct::ThemeFont> {
    direct::ThemeFont::from_xml(value)
        .ok_or_else(|| invalid("the theme font token is outside Word's domain"))
}

fn parse_font_hint(value: &'static str) -> Result<direct::FontHint> {
    direct::FontHint::from_xml(value)
        .ok_or_else(|| invalid("the font hint token is outside Word's domain"))
}

fn parse_underline(value: &'static str) -> Result<litchi_docx::UnderlineStyle> {
    litchi_docx::UnderlineStyle::from_xml(value)
        .ok_or_else(|| invalid("the underline token is outside Word's domain"))
}

fn parse_highlight(value: &'static str) -> Result<direct::Highlight> {
    direct::Highlight::from_xml(value)
        .ok_or_else(|| invalid("the highlight token is outside Word's domain"))
}

fn parse_direction(value: &'static str) -> Result<bool> {
    match value {
        "rtl" => Ok(true),
        "ltr" => Ok(false),
        _ => Err(invalid("the direction token is outside Word's domain")),
    }
}

fn clean_spacing(paragraph: &mut direct::Paragraph) {
    if paragraph.spacing.is_some_and(|spacing| {
        spacing.before.is_none()
            && spacing.before_lines.is_none()
            && spacing.before_auto_spacing.is_none()
            && spacing.after.is_none()
            && spacing.after_lines.is_none()
            && spacing.after_auto_spacing.is_none()
            && spacing.line.is_none()
            && spacing.line_rule.is_none()
    }) {
        paragraph.spacing = None;
    }
}

fn clean_indentation(paragraph: &mut direct::Paragraph) {
    if paragraph.indentation.is_some_and(|indentation| {
        indentation.start.is_none()
            && indentation.start_chars.is_none()
            && indentation.end.is_none()
            && indentation.end_chars.is_none()
            && indentation.left.is_none()
            && indentation.left_chars.is_none()
            && indentation.right.is_none()
            && indentation.right_chars.is_none()
            && indentation.hanging.is_none()
            && indentation.hanging_chars.is_none()
            && indentation.first_line.is_none()
            && indentation.first_line_chars.is_none()
    }) {
        paragraph.indentation = None;
    }
}

fn clean_fonts(run: &mut direct::Run) {
    if run.fonts.as_ref().is_some_and(|fonts| {
        fonts.hint.is_none()
            && fonts.ascii.is_none()
            && fonts.h_ansi.is_none()
            && fonts.east_asia.is_none()
            && fonts.complex_script.is_none()
            && fonts.ascii_theme.is_none()
            && fonts.h_ansi_theme.is_none()
            && fonts.east_asia_theme.is_none()
            && fonts.complex_script_theme.is_none()
    }) {
        run.fonts = None;
    }
}

fn clean_language(run: &mut direct::Run) {
    if run.language.as_ref().is_some_and(|language| {
        language.latin.is_none() && language.east_asia.is_none() && language.bidi.is_none()
    }) {
        run.language = None;
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn line_spacing_matches_word_units_and_domains() {
        let mut paragraph = direct::Paragraph::default();
        Setter::ParagraphLineSpacing
            .apply_paragraph(&mut paragraph, Some(Scalar::Text("1.5X")))
            .unwrap();
        assert_eq!(paragraph.spacing.as_ref().unwrap().line, Some(360));
        assert_eq!(
            paragraph.spacing.as_ref().unwrap().line_rule,
            Some(LineSpacingRule::Auto)
        );

        Setter::ParagraphLineSpacing
            .apply_paragraph(&mut paragraph, Some(Scalar::Text("150%")))
            .unwrap();
        assert_eq!(paragraph.spacing.as_ref().unwrap().line, Some(360));

        Setter::ParagraphLineSpacing
            .apply_paragraph(&mut paragraph, Some(Scalar::Text("12PT")))
            .unwrap();
        assert_eq!(paragraph.spacing.as_ref().unwrap().line, Some(240));
        assert_eq!(
            paragraph.spacing.as_ref().unwrap().line_rule,
            Some(LineSpacingRule::Exact)
        );

        assert!(
            Setter::ParagraphLineSpacing
                .apply_paragraph(&mut paragraph, Some(Scalar::Text("-1PT")))
                .is_err()
        );
        assert!(
            Setter::ParagraphLineSpacing
                .apply_paragraph(&mut paragraph, Some(Scalar::Text("0X")))
                .is_err()
        );
    }

    #[test]
    fn vertical_false_only_clears_the_selected_placement() {
        let mut run = direct::Run::default();
        run.vertical_align = Some(direct::VerticalAlign::Subscript);
        set_vertical(
            &mut run,
            Some(Scalar::Bool(false)),
            direct::VerticalAlign::Superscript,
        )
        .unwrap();
        assert_eq!(run.vertical_align, Some(direct::VerticalAlign::Subscript));
    }
}
