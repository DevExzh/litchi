//! Values of a shape's text and tables, in OfficeCLI's spellings (record
//! 0010 §2.3; `PowerPointHandler.NodeBuilder.cs`).
//!
//! Text values come from the leaf's DrawingML text-body reader
//! (L-PPTX-R7), which reports local properties only; table values come from
//! the leaf's table reader over the graphic frame. Every function here
//! borrows the leaf value and allocates only where OfficeCLI's text is
//! composed (`1.5x`, `33%`, a cell's text).

use std::borrow::Cow;

use litchi_drawingml::color::{
    self as drawing, Base as DrawingBase, Transform as DrawingTransform,
};
use litchi_drawingml::text::body::{
    Body, Paragraph, ParagraphProperties, Run, RunKind, RunProperties, Spacing,
};
use litchi_drawingml::text::{Alignment, Caps, Strike, Underline};
use litchi_pptx::shape::owned::Handle as ShapeHandle;
use litchi_pptx::table::{Properties as TableProperties, Table};

use super::node::TableItem;
use super::spec::{Kind, ParagraphKey, RunKey, ShapeKey};
use crate::error::Result;
use crate::node::Value;
use crate::value::{Color, Rgba, Scheme, System, Transform, TransformKind};

/// What a listed run is: a text run (`run[N]`) or a line break
/// (`linebreak[N]`). Fields are not listed, as in OfficeCLI.
pub(crate) fn listed_run(run: &Run) -> Option<Kind> {
    match run.kind {
        RunKind::Text => Some(Kind::Run),
        RunKind::Break => Some(Kind::LineBreak),
        _ => None,
    }
}

/// A paragraph's text: its runs, line breaks (`\n`) and field results, as
/// the leaf joins them; borrowed when one run holds it all.
pub(crate) fn paragraph_text(paragraph: &Paragraph) -> Cow<'_, str> {
    let mut texts = paragraph.runs.iter().filter(|run| !run.text.is_empty());
    match (texts.next(), texts.next()) {
        (None, _) => Cow::Borrowed(""),
        (Some(only), None) => Cow::Borrowed(only.text.as_str()),
        (Some(_), Some(_)) => Cow::Owned(paragraph.text()),
    }
}

/// The table a graphic frame holds.
fn table(handle: &ShapeHandle) -> Result<Table<'_>> {
    Ok(handle.table()?)
}

/// A table's row count and its first row's cell count.
pub(crate) fn table_shape(handle: &ShapeHandle) -> Result<(usize, usize)> {
    let table = table(handle)?;
    let rows = table.rows()?;
    let cols = rows.first().map_or(0, |row| row.cell_count());
    Ok((rows.len(), cols))
}

/// A table's `a:tblPr` switches and style.
pub(crate) fn table_properties(handle: &ShapeHandle) -> Result<Option<TableProperties>> {
    Ok(table(handle)?.properties()?)
}

/// The number of cells in row `row` (zero-based) of a table.
pub(crate) fn row_cells(handle: &ShapeHandle, row: u32) -> Result<usize> {
    Ok(table(handle)?
        .rows()?
        .get(row as usize)
        .map_or(0, |row| row.cell_count()))
}

/// The text of a table cell, paragraphs separated by line feeds.
pub(crate) fn cell_text(item: &TableItem) -> Result<String> {
    let table = table(&item.table.handle)?;
    let Some(row) = table.rows()?.get(item.row as usize).copied() else {
        return Ok(String::new());
    };
    match row.cells()?.get(item.column as usize) {
        Some(cell) => Ok(cell.text()?),
        None => Ok(String::new()),
    }
}

/// A table row's height (`a:tr@h`, record 0333).
pub(crate) fn row_height(item: &TableItem) -> Result<Option<i64>> {
    let table = table(&item.table.handle)?;
    match table.rows()?.get(item.row as usize) {
        Some(row) => Ok(row.height()?),
        None => Ok(None),
    }
}

/// A table cell's text body with its local properties (record 0333);
/// `None` for a missing cell or a cell without a body.
pub(crate) fn cell_body(item: &TableItem) -> Result<Option<Body>> {
    let table = table(&item.table.handle)?;
    match table.cell(item.row as usize, item.column as usize)? {
        Some(cell) => Ok(cell.text_body()?),
        None => Ok(None),
    }
}

/// A table cell's local properties: fill, vertical anchor and insets
/// (record 0333).
pub(crate) fn cell_properties(item: &TableItem) -> Result<litchi_pptx::table::CellProperties> {
    let table = table(&item.table.handle)?;
    match table.cell(item.row as usize, item.column as usize)? {
        Some(cell) => Ok(cell.properties()?),
        None => Ok(litchi_pptx::table::CellProperties::default()),
    }
}

/// The value of a table cell property.
pub(crate) fn cell_value(
    properties: &litchi_pptx::table::CellProperties,
    key: super::spec::CellKey,
) -> Option<Value<'static>> {
    use super::spec::CellKey;
    use litchi_drawingml::text::Anchor;
    use litchi_pptx::opened::CellFill;
    let inset = |index: usize| {
        properties.margins[index].map(|emu| Value::Length(crate::value::Length::from_emu(emu)))
    };
    match key {
        CellKey::Fill => match properties.fill? {
            CellFill::Rgb(rgb) => Some(Value::Color(Color::rgb(rgb.0[0], rgb.0[1], rgb.0[2]))),
            CellFill::Scheme(scheme) => Some(Value::Color(Color::scheme(scheme_of(scheme)))),
            // An explicit no-fill has no colour to report.
            _ => None,
        },
        CellKey::Valign => match properties.anchor? {
            Anchor::Top => Some(Value::Token("top")),
            Anchor::Center => Some(Value::Token("center")),
            Anchor::Bottom => Some(Value::Token("bottom")),
            Anchor::Justified | Anchor::Distributed => None,
        },
        CellKey::PaddingLeft => inset(0),
        CellKey::PaddingRight => inset(1),
        CellKey::PaddingTop => inset(2),
        CellKey::PaddingBottom => inset(3),
    }
}

/// The text of every cell of a table, row by row.
pub(crate) fn table_cells(handle: &ShapeHandle) -> Result<Vec<String>> {
    let table = table(handle)?;
    let mut cells = Vec::new();
    for row in table.rows()? {
        for cell in row.cells()? {
            cells.push(cell.text()?);
        }
    }
    Ok(cells)
}

// ---------------------------------------------------------------------------
// Shape-level text facts
// ---------------------------------------------------------------------------

/// What a shape's text says at shape level (`ShapeToNode`): the first text
/// run's character properties, unless the runs disagree on size, colour,
/// bold or italic; the first paragraph's mark properties when it has no
/// run; and the first paragraph's properties.
pub(crate) struct Facts<'a> {
    first: Option<&'a RunProperties>,
    /// The first paragraph's `a:endParaRPr`, for a shape without runs.
    end: Option<&'a RunProperties>,
    mixed_size: bool,
    mixed_color: bool,
    mixed_bold: bool,
    mixed_italic: bool,
    paragraph: Option<&'a ParagraphProperties>,
}

impl<'a> Facts<'a> {
    pub(crate) fn of(body: &'a Body) -> Self {
        let mut runs = body
            .paragraphs
            .iter()
            .flat_map(|paragraph| paragraph.runs.iter())
            .filter(|run| run.is_text())
            .map(|run| &run.properties);
        let first = runs.next();
        let (mut mixed_size, mut mixed_color, mut mixed_bold, mut mixed_italic) =
            (false, false, false, false);
        if let Some(first) = first {
            for run in runs {
                mixed_size |= run.size != first.size;
                mixed_color |= run.color() != first.color();
                mixed_bold |= run.bold != first.bold;
                mixed_italic |= run.italic != first.italic;
            }
        }
        let paragraph = body.paragraphs.first();
        Self {
            first,
            end: if first.is_none() {
                paragraph.and_then(|paragraph| paragraph.end_properties.as_deref())
            } else {
                None
            },
            mixed_size,
            mixed_color,
            mixed_bold,
            mixed_italic,
            paragraph: paragraph.map(|paragraph| &paragraph.properties),
        }
    }

    /// The value of a shape key read from the text.
    pub(crate) fn value(&self, key: ShapeKey) -> Option<Value<'a>> {
        if let Some(paragraph) = self.paragraph {
            let paragraph_key = match key {
                ShapeKey::Align => Some(ParagraphKey::Align),
                ShapeKey::LineSpacing => Some(ParagraphKey::LineSpacing),
                ShapeKey::SpaceBefore => Some(ParagraphKey::SpaceBefore),
                ShapeKey::SpaceAfter => Some(ParagraphKey::SpaceAfter),
                ShapeKey::Indent => Some(ParagraphKey::Indent),
                ShapeKey::MarginLeft => Some(ParagraphKey::MarginLeft),
                ShapeKey::MarginRight => Some(ParagraphKey::MarginRight),
                ShapeKey::Direction => Some(ParagraphKey::Direction),
                _ => None,
            };
            if let Some(paragraph_key) = paragraph_key {
                return paragraph_properties_value(paragraph, paragraph_key);
            }
        }
        if let Some(first) = self.first {
            return match key {
                ShapeKey::Font | ShapeKey::FontLatin => run_value(first, RunKey::FontLatin),
                ShapeKey::FontEa => run_value(first, RunKey::FontEa),
                ShapeKey::FontCs => run_value(first, RunKey::FontCs),
                ShapeKey::Size if !self.mixed_size => run_value(first, RunKey::Size),
                ShapeKey::Bold if !self.mixed_bold => run_value(first, RunKey::Bold),
                ShapeKey::Italic if !self.mixed_italic => run_value(first, RunKey::Italic),
                ShapeKey::Cap => run_value(first, RunKey::Cap),
                ShapeKey::Underline => run_value(first, RunKey::Underline),
                ShapeKey::Strike => run_value(first, RunKey::Strike),
                ShapeKey::Highlight => run_value(first, RunKey::Highlight),
                ShapeKey::Spacing => run_value(first, RunKey::Spacing),
                ShapeKey::Color if !self.mixed_color => run_value(first, RunKey::Color),
                _ => None,
            };
        }
        let end = self.end?;
        match key {
            ShapeKey::Font | ShapeKey::FontLatin => run_value(end, RunKey::FontLatin),
            ShapeKey::FontEa => run_value(end, RunKey::FontEa),
            ShapeKey::FontCs => run_value(end, RunKey::FontCs),
            ShapeKey::Size => run_value(end, RunKey::Size),
            ShapeKey::Bold => run_value(end, RunKey::Bold),
            ShapeKey::Italic => run_value(end, RunKey::Italic),
            ShapeKey::Cap => run_value(end, RunKey::Cap),
            ShapeKey::Underline => run_value(end, RunKey::Underline),
            ShapeKey::Strike => run_value(end, RunKey::Strike),
            ShapeKey::Color => run_value(end, RunKey::Color),
            _ => None,
        }
    }
}

// ---------------------------------------------------------------------------
// Paragraph and run values
// ---------------------------------------------------------------------------

/// The value of a paragraph key.
pub(crate) fn paragraph_value(paragraph: &Paragraph, key: ParagraphKey) -> Option<Value<'_>> {
    if key == ParagraphKey::Size {
        // An empty paragraph's mark size: a spacer line's height (OfficeCLI).
        if paragraph.runs.iter().any(Run::is_text) {
            return None;
        }
        return paragraph
            .end_properties
            .as_deref()
            .and_then(|end| run_value(end, RunKey::Size));
    }
    paragraph_properties_value(&paragraph.properties, key)
}

fn paragraph_properties_value(
    properties: &ParagraphProperties,
    key: ParagraphKey,
) -> Option<Value<'static>> {
    match key {
        ParagraphKey::Align => properties.alignment.map(|alignment| {
            Value::Token(match alignment {
                Alignment::Left => "left",
                Alignment::Center => "center",
                Alignment::Right => "right",
                Alignment::Justified => "justify",
                Alignment::JustifiedLow => "justLow",
                Alignment::Distributed => "dist",
                Alignment::ThaiDistributed => "thaiDist",
            })
        }),
        ParagraphKey::Level => properties.level.map(|level| Value::Int(level.into())),
        ParagraphKey::Indent => properties.indent.map(emu),
        ParagraphKey::MarginLeft => properties.margin_left.map(emu),
        ParagraphKey::MarginRight => properties.margin_right.map(emu),
        ParagraphKey::Direction => properties
            .right_to_left
            .map(|rtl| Value::Token(if rtl { "rtl" } else { "ltr" })),
        ParagraphKey::LineSpacing => properties.line_spacing.map(|spacing| match spacing {
            Spacing::Percent(percent) => Value::Text(Cow::Owned(format!(
                "{}x",
                decimal(f64::from(percent) / 100_000.0, 5)
            ))),
            Spacing::Points(points) => Value::Text(Cow::Owned(format!(
                "{}pt",
                decimal(f64::from(points) / 100.0, 2)
            ))),
            _ => Value::Text(Cow::Borrowed("")),
        }),
        ParagraphKey::SpaceBefore => points_spacing(properties.space_before),
        ParagraphKey::SpaceAfter => points_spacing(properties.space_after),
        ParagraphKey::Size => None,
    }
}

/// OfficeCLI reads exact spacing (`a:spcPts`) only.
fn points_spacing(spacing: Option<Spacing>) -> Option<Value<'static>> {
    match spacing? {
        Spacing::Points(points) => Some(centipoints(i64::from(points))),
        _ => None,
    }
}

/// The value of a run key.
pub(crate) fn run_value(properties: &RunProperties, key: RunKey) -> Option<Value<'_>> {
    fn font(font: Option<&litchi_drawingml::text::body::Font>) -> Option<Value<'_>> {
        font.map(|font| Value::Text(Cow::Borrowed(font.typeface.as_str())))
    }
    match key {
        RunKey::FontLatin => font(properties.latin.as_ref()),
        RunKey::FontEa => font(properties.east_asian.as_ref()),
        RunKey::FontCs => font(properties.complex_script.as_ref()),
        RunKey::Size => properties
            .size
            .map(|size| centipoints(i64::from(size.get()))),
        RunKey::Bold => properties.bold.map(Value::Bool),
        RunKey::Italic => properties.italic.map(Value::Bool),
        RunKey::Underline => properties
            .underline
            .and_then(underline_token)
            .map(Value::Token),
        RunKey::Strike => properties.strike.map(|strike| {
            Value::Token(match strike {
                Strike::None => "none",
                Strike::Single => "single",
                Strike::Double => "double",
            })
        }),
        RunKey::Spacing => properties
            .spacing
            .and_then(|spacing| Value::number(f64::from(spacing.get()) / 100.0)),
        RunKey::Baseline => properties
            .baseline
            .filter(|baseline| *baseline != 0)
            .map(|baseline| {
                Value::Text(Cow::Owned(format!(
                    "{}%",
                    decimal(f64::from(baseline) / 1_000.0, 2)
                )))
            }),
        RunKey::Color => properties.color().and_then(drawing_color).map(Value::Color),
        RunKey::Highlight => properties
            .highlight
            .as_ref()
            .and_then(drawing_color)
            .map(Value::Color),
        RunKey::Cap => properties.caps.and_then(|caps| match caps {
            Caps::All => Some(Value::Token("all")),
            Caps::Small => Some(Value::Token("small")),
            Caps::None => None,
        }),
        RunKey::Kern => properties
            .kerning
            .map(|kerning| Value::Int(i64::from(kerning.get()))),
        // Resolved through the slide's relationships by the node layer.
        RunKey::Link | RunKey::Tooltip => None,
        RunKey::Lang => properties
            .language
            .as_deref()
            .map(|language| Value::Text(Cow::Borrowed(language))),
    }
}

/// OfficeCLI's underline names: `single` and `double` for `sng` and `dbl`,
/// the token as written for the others; no value for `none`.
fn underline_token(underline: Underline) -> Option<&'static str> {
    Some(match underline {
        Underline::None => return None,
        Underline::Words => "words",
        Underline::Single => "single",
        Underline::Double => "double",
        Underline::Heavy => "heavy",
        Underline::Dotted => "dotted",
        Underline::DottedHeavy => "dottedHeavy",
        Underline::Dash => "dash",
        Underline::DashHeavy => "dashHeavy",
        Underline::DashLong => "dashLong",
        Underline::DashLongHeavy => "dashLongHeavy",
        Underline::DotDash => "dotDash",
        Underline::DotDashHeavy => "dotDashHeavy",
        Underline::DotDotDash => "dotDotDash",
        Underline::DotDotDashHeavy => "dotDotDashHeavy",
        Underline::Wavy => "wavy",
        Underline::WavyHeavy => "wavyHeavy",
        Underline::WavyDouble => "wavyDbl",
    })
}

fn emu(value: i32) -> Value<'static> {
    Value::Length(crate::value::Length::from_emu(i64::from(value)))
}

/// A length of hundredths of a point (127 EMU each).
fn centipoints(value: i64) -> Value<'static> {
    Value::Length(crate::value::Length::from_emu(value.saturating_mul(127)))
}

/// `value` with at most `places` decimals and no trailing zeros, as .NET's
/// `0.##` format writes it.
pub(crate) fn decimal(value: f64, places: usize) -> String {
    let mut text = format!("{value:.places$}");
    if text.contains('.') {
        while text.ends_with('0') {
            text.pop();
        }
        if text.ends_with('.') {
            text.pop();
        }
    }
    if text == "-0" {
        text.remove(0);
    }
    text
}

// ---------------------------------------------------------------------------
// Colours
// ---------------------------------------------------------------------------

/// A DrawingML colour as the node view writes it: `#RRGGBB` (with alpha,
/// `#RRGGBBAA`), a scheme name, a system or preset colour, each with the
/// transforms OfficeCLI encodes (`accent1+lumMod75`). Colours OfficeCLI does
/// not print (scRGB, HSL, unknown extensions) have no value.
pub(crate) fn drawing_color(value: &drawing::Value) -> Option<Color> {
    match value {
        drawing::Value::Rgb(rgb) => Some(Color::rgb(rgb.red(), rgb.green(), rgb.blue())),
        drawing::Value::Scheme(scheme) => Some(Color::scheme(scheme_of(*scheme))),
        drawing::Value::System(system) => System::from_token(system.token()).map(Color::system),
        drawing::Value::Preset(preset) => {
            Rgba::preset(preset.token()).map(|rgba| Color::new(crate::value::Base::Rgb(rgba)))
        },
        drawing::Value::Transformed(transformed) => {
            let mut color = match transformed.base() {
                DrawingBase::Rgb(rgb) => Color::rgb(rgb.red(), rgb.green(), rgb.blue()),
                DrawingBase::Scheme(scheme) => Color::scheme(scheme_of(*scheme)),
                DrawingBase::System(system) => Color::system(System::from_token(system.token())?),
                DrawingBase::Preset(preset) => {
                    Color::new(crate::value::Base::Rgb(Rgba::preset(preset.token())?))
                },
                _ => return None,
            };
            for transform in transformed.transforms() {
                let (kind, value) = match *transform {
                    DrawingTransform::LumMod(value) => (TransformKind::LumMod, value.value()),
                    DrawingTransform::LumOff(value) => (TransformKind::LumOff, value.value()),
                    DrawingTransform::SatMod(value) => (TransformKind::SatMod, value.value()),
                    DrawingTransform::SatOff(value) => (TransformKind::SatOff, value.value()),
                    DrawingTransform::HueMod(value) => {
                        (TransformKind::HueMod, i32::try_from(value.value()).ok()?)
                    },
                    DrawingTransform::HueOff(value) => (TransformKind::HueOff, value.value()),
                    DrawingTransform::Shade(value) => {
                        (TransformKind::Shade, i32::try_from(value.value()).ok()?)
                    },
                    DrawingTransform::Tint(value) => {
                        (TransformKind::Tint, i32::try_from(value.value()).ok()?)
                    },
                    DrawingTransform::Alpha(value) => {
                        let alpha = i32::try_from(value.value()).ok()?;
                        // OfficeCLI writes an RGB colour's alpha as `#RRGGBBAA`.
                        if let crate::value::Base::Rgb(rgba) = color.base() {
                            if color.transforms().is_empty() {
                                let channel = (f64::from(alpha) * 255.0 / 100_000.0).round();
                                #[allow(
                                    clippy::cast_possible_truncation,
                                    clippy::cast_sign_loss,
                                    reason = "the alpha percentage is 0..=100000, so the channel is 0..=255"
                                )]
                                let channel = channel.clamp(0.0, 255.0) as u8;
                                color =
                                    Color::new(crate::value::Base::Rgb(rgba.with_alpha(channel)));
                                continue;
                            }
                        }
                        (TransformKind::Alpha, alpha)
                    },
                    // OfficeCLI encodes only the transforms above.
                    _ => continue,
                };
                if let Some(transform) = Transform::new(kind, value) {
                    color = color.with_transform(transform)?;
                }
            }
            Some(color)
        },
        _ => None,
    }
}

fn scheme_of(scheme: drawing::Scheme) -> Scheme {
    match scheme {
        drawing::Scheme::Background => Scheme::Background1,
        drawing::Scheme::Text => Scheme::Text1,
        drawing::Scheme::Background2 => Scheme::Background2,
        drawing::Scheme::Text2 => Scheme::Text2,
        drawing::Scheme::Accent1 => Scheme::Accent1,
        drawing::Scheme::Accent2 => Scheme::Accent2,
        drawing::Scheme::Accent3 => Scheme::Accent3,
        drawing::Scheme::Accent4 => Scheme::Accent4,
        drawing::Scheme::Accent5 => Scheme::Accent5,
        drawing::Scheme::Accent6 => Scheme::Accent6,
        drawing::Scheme::Hyperlink => Scheme::Hyperlink,
        drawing::Scheme::FollowedHyperlink => Scheme::FollowedHyperlink,
        drawing::Scheme::Dark1 => Scheme::Dark1,
        drawing::Scheme::Light1 => Scheme::Light1,
        drawing::Scheme::Dark2 => Scheme::Dark2,
        drawing::Scheme::Light2 => Scheme::Light2,
        drawing::Scheme::Placeholder => Scheme::Placeholder,
    }
}

#[cfg(test)]
mod tests {
    use super::decimal;

    #[test]
    fn decimals_print_like_dotnet() {
        assert_eq!(decimal(1.5, 5), "1.5");
        assert_eq!(decimal(1.0, 5), "1");
        assert_eq!(decimal(0.001, 5), "0.001");
        assert_eq!(decimal(33.0, 2), "33");
        assert_eq!(decimal(-25.5, 2), "-25.5");
        assert_eq!(decimal(-0.0001, 2), "0");
    }
}
