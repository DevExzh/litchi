//! XLS cell formatting: the read projection of a cell's resolved XF and the
//! unified property names of the native `FormatChange` writer.

use std::borrow::Cow;
use std::sync::Arc;

use litchi_xls::{
    BorderStyle, Color as XlsColor, FillPattern, Font, FontUnderline, Formatting,
    HorizontalAlignment, Palette, StyleTables, VerticalAlignment,
};

use super::spec::CellKey;
use crate::node::Value;
use crate::value::{Base, Color, Scalar};

/// Where a cell's formatting tables live: the eager reader, or the
/// positional owner's retained XF table plus its font and palette tables.
pub(crate) enum Styles<'a> {
    Eager(&'a litchi_xls::raw::Reader<std::io::Cursor<Vec<u8>>>),
    Source {
        formatting: Arc<Formatting>,
        tables: Arc<StyleTables>,
    },
}

impl Styles<'_> {
    fn formatting(&self) -> &Formatting {
        match self {
            Self::Eager(reader) => reader.formatting(),
            Self::Source { formatting, .. } => formatting,
        }
    }

    fn font(&self, index: u16) -> Option<&Font> {
        match self {
            Self::Eager(reader) => reader.font(index),
            Self::Source { tables, .. } => tables.font(index),
        }
    }

    fn palette(&self) -> &Palette {
        match self {
            Self::Eager(reader) => reader.palette(),
            Self::Source { tables, .. } => tables.palette(),
        }
    }
}

fn color(color: XlsColor) -> Value<'static> {
    Value::Color(Color::rgb(color.red(), color.green(), color.blue()))
}

const fn horizontal_token(value: HorizontalAlignment) -> Option<&'static str> {
    Some(match value {
        HorizontalAlignment::General => return None,
        HorizontalAlignment::Left => "left",
        HorizontalAlignment::Center => "center",
        HorizontalAlignment::Right => "right",
        HorizontalAlignment::Fill => "fill",
        HorizontalAlignment::Justify => "justify",
        HorizontalAlignment::CenterAcrossSelection => "centerContinuous",
        HorizontalAlignment::Distributed => "distributed",
    })
}

const fn vertical_token(value: VerticalAlignment) -> Option<&'static str> {
    Some(match value {
        VerticalAlignment::Bottom => return None,
        VerticalAlignment::Top => "top",
        VerticalAlignment::Center => "center",
        VerticalAlignment::Justify => "justify",
        VerticalAlignment::Distributed => "distributed",
    })
}

fn border_token(value: BorderStyle) -> Option<&'static str> {
    Some(match value {
        BorderStyle::Thin => "thin",
        BorderStyle::Medium => "medium",
        BorderStyle::Dashed => "dashed",
        BorderStyle::Dotted => "dotted",
        BorderStyle::Thick => "thick",
        BorderStyle::Double => "double",
        BorderStyle::Hair => "hair",
        BorderStyle::MediumDashed => "mediumDashed",
        BorderStyle::DashDot => "dashDot",
        BorderStyle::MediumDashDot => "mediumDashDot",
        BorderStyle::DashDotDot => "dashDotDot",
        BorderStyle::MediumDashDotDot => "mediumDashDotDot",
        BorderStyle::SlantedDashDot => "slantDashDot",
        _ => return None,
    })
}

const FILL_PATTERNS: [(FillPattern, &str); 19] = [
    (FillPattern::None, "none"),
    (FillPattern::Solid, "solid"),
    (FillPattern::MediumGray, "mediumGray"),
    (FillPattern::DarkGray, "darkGray"),
    (FillPattern::LightGray, "lightGray"),
    (FillPattern::DarkHorizontal, "darkHorizontal"),
    (FillPattern::DarkVertical, "darkVertical"),
    (FillPattern::DarkDown, "darkDown"),
    (FillPattern::DarkUp, "darkUp"),
    (FillPattern::DarkGrid, "darkGrid"),
    (FillPattern::DarkTrellis, "darkTrellis"),
    (FillPattern::LightHorizontal, "lightHorizontal"),
    (FillPattern::LightVertical, "lightVertical"),
    (FillPattern::LightDown, "lightDown"),
    (FillPattern::LightUp, "lightUp"),
    (FillPattern::LightGrid, "lightGrid"),
    (FillPattern::LightTrellis, "lightTrellis"),
    (FillPattern::Gray125, "gray125"),
    (FillPattern::Gray0625, "gray0625"),
];

/// One formatting key of a cell whose record names XF `xf`; non-default
/// values only, mirroring the XLSX projection's omissions.
pub(crate) fn value(styles: &Styles<'_>, xf: u16, key: CellKey) -> Option<Value<'static>> {
    let formatting = styles.formatting();
    let format = formatting.resolved_cell_format(xf)?;
    let font_index = format.font_index();
    let font = styles.font(font_index);
    let custom_font = font_index != 0;
    let palette = styles.palette();
    let alignment = format.alignment();
    let borders = format.borders();
    let fill = format.fill();
    match key {
        CellKey::Bold => font?.is_bold().then_some(Value::Bool(true)),
        CellKey::Italic => font?.is_italic().then_some(Value::Bool(true)),
        CellKey::Strike => font?.is_struck_out().then_some(Value::Bool(true)),
        CellKey::Underline => match font?.underline() {
            FontUnderline::None => None,
            FontUnderline::Single => Some(Value::Token("single")),
            FontUnderline::Double => Some(Value::Token("double")),
            FontUnderline::SingleAccounting => Some(Value::Token("singleAccounting")),
            FontUnderline::DoubleAccounting => Some(Value::Token("doubleAccounting")),
        },
        CellKey::FontColor => palette.color(font?.color_index()).map(color),
        CellKey::FontSize => custom_font.then(|| {
            Value::Length(crate::value::Length::from_emu(
                i64::from(font.map_or(200, Font::height_twips)) * 635,
            ))
        }),
        CellKey::FontName => custom_font
            .then(|| font.map(|font| Value::Text(Cow::Owned(font.name().to_owned()))))
            .flatten(),
        CellKey::Fill => match fill.pattern() {
            FillPattern::None => None,
            _ => fill.foreground_color(palette).map(color),
        },
        CellKey::FillPattern => match fill.pattern() {
            FillPattern::None | FillPattern::Solid => None,
            pattern => FILL_PATTERNS
                .iter()
                .find(|(candidate, _)| *candidate == pattern)
                .map(|(_, name)| Value::Token(name)),
        },
        CellKey::FillBg => match fill.pattern() {
            FillPattern::None | FillPattern::Solid => None,
            _ => fill.background_color(palette).map(color),
        },
        CellKey::BorderLeft => border_token(borders.left().style()).map(Value::Token),
        CellKey::BorderRight => border_token(borders.right().style()).map(Value::Token),
        CellKey::BorderTop => border_token(borders.top().style()).map(Value::Token),
        CellKey::BorderBottom => border_token(borders.bottom().style()).map(Value::Token),
        CellKey::WrapText => alignment.wraps_text().then_some(Value::Bool(true)),
        CellKey::Horizontal => horizontal_token(alignment.horizontal()).map(Value::Token),
        CellKey::Vertical => vertical_token(alignment.vertical()).map(Value::Token),
        CellKey::Indent => {
            (alignment.indent() > 0).then_some(Value::Int(i64::from(alignment.indent())))
        },
        CellKey::ShrinkToFit => alignment.shrinks_to_fit().then_some(Value::Bool(true)),
        CellKey::Locked => (!format.locked()).then_some(Value::Bool(false)),
        CellKey::FormulaHidden => format.hidden().then_some(Value::Bool(true)),
        CellKey::NumFmtId => {
            let id = format.number_format_id();
            (id != 0).then_some(Value::Int(i64::from(id)))
        },
        CellKey::NumberFormat => formatting
            .number_format(format.number_format_id())
            .filter(|_| format.number_format_id() != 0)
            .map(|format| Value::Text(Cow::Owned(format.code().to_owned()))),
        CellKey::Type | CellKey::Formula | CellKey::CachedValue | CellKey::Empty => None,
    }
}

/// The native horizontal alignment a unified token names.
pub(crate) fn horizontal(token: &str) -> Option<HorizontalAlignment> {
    Some(match token {
        "general" => HorizontalAlignment::General,
        "left" => HorizontalAlignment::Left,
        "center" => HorizontalAlignment::Center,
        "right" => HorizontalAlignment::Right,
        "fill" => HorizontalAlignment::Fill,
        "justify" => HorizontalAlignment::Justify,
        "centerContinuous" => HorizontalAlignment::CenterAcrossSelection,
        "distributed" => HorizontalAlignment::Distributed,
        _ => return None,
    })
}

/// The native vertical alignment a unified token names.
pub(crate) fn vertical(token: &str) -> Option<VerticalAlignment> {
    Some(match token {
        "top" => VerticalAlignment::Top,
        "center" => VerticalAlignment::Center,
        "bottom" => VerticalAlignment::Bottom,
        "justify" => VerticalAlignment::Justify,
        "distributed" => VerticalAlignment::Distributed,
        _ => return None,
    })
}

/// The native border line style a unified token names; `None` clears.
pub(crate) fn border(token: Option<&str>) -> Option<BorderStyle> {
    let Some(token) = token else {
        return Some(BorderStyle::None);
    };
    Some(match token {
        "thin" => BorderStyle::Thin,
        "medium" => BorderStyle::Medium,
        "dashed" => BorderStyle::Dashed,
        "dotted" => BorderStyle::Dotted,
        "thick" => BorderStyle::Thick,
        "double" => BorderStyle::Double,
        "hair" => BorderStyle::Hair,
        "mediumDashed" => BorderStyle::MediumDashed,
        "dashDot" => BorderStyle::DashDot,
        "mediumDashDot" => BorderStyle::MediumDashDot,
        "dashDotDot" => BorderStyle::DashDotDot,
        "mediumDashDotDot" => BorderStyle::MediumDashDotDot,
        "slantDashDot" => BorderStyle::SlantedDashDot,
        _ => return None,
    })
}

/// The native fill pattern a unified token names.
pub(crate) fn fill_pattern(token: &str) -> Option<FillPattern> {
    FILL_PATTERNS
        .iter()
        .find(|(_, name)| *name == token)
        .map(|(pattern, _)| *pattern)
}

/// An opaque RGB color for a palette lookup.
pub(crate) fn rgb(value: Option<Scalar<'_>>) -> Option<XlsColor> {
    let Some(Scalar::Color(color)) = value else {
        return None;
    };
    if !color.transforms().is_empty() || color.spreadsheet_tint().is_some() {
        return None;
    }
    match color.base() {
        Base::Rgb(value) if value.alpha() == 0xff => {
            Some(XlsColor::new(value.red(), value.green(), value.blue()))
        },
        _ => None,
    }
}
