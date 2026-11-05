//! Local formatting read from automatic styles, and the typed edits that
//! derive new automatic styles.

use std::borrow::Cow;

use litchi_core::{Error, Result};

use super::xml::{self, Ns, Prefixes};

/// Character properties a style sets (`style:text-properties`), as written.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct TextProperties<'a> {
    font_weight: Option<Cow<'a, str>>,
    font_style: Option<Cow<'a, str>>,
    underline: Option<Cow<'a, str>>,
    line_through: Option<Cow<'a, str>>,
    font_size: Option<Cow<'a, str>>,
    color: Option<Cow<'a, str>>,
    font_name: Option<Cow<'a, str>>,
    font_family: Option<Cow<'a, str>>,
    background: Option<Cow<'a, str>>,
    position: Option<Cow<'a, str>>,
    transform: Option<Cow<'a, str>>,
    variant: Option<Cow<'a, str>>,
}

impl<'a> TextProperties<'a> {
    pub(crate) fn read(tag: &'a str, prefixes: &Prefixes) -> Self {
        let mut properties = Self::default();
        for (name, value) in xml::attributes(tag) {
            let Some((ns, local)) = prefixes.resolve(name.as_bytes()) else {
                continue;
            };
            let slot = match (ns, local) {
                (Ns::Fo, b"font-weight") => &mut properties.font_weight,
                (Ns::Fo, b"font-style") => &mut properties.font_style,
                (Ns::Style, b"text-underline-style") => &mut properties.underline,
                (Ns::Style, b"text-line-through-style") => &mut properties.line_through,
                (Ns::Fo, b"font-size") => &mut properties.font_size,
                (Ns::Fo, b"color") => &mut properties.color,
                (Ns::Style, b"font-name") => &mut properties.font_name,
                (Ns::Fo, b"font-family") => &mut properties.font_family,
                (Ns::Fo, b"background-color") => &mut properties.background,
                (Ns::Style, b"text-position") => &mut properties.position,
                (Ns::Fo, b"text-transform") => &mut properties.transform,
                (Ns::Fo, b"font-variant") => &mut properties.variant,
                _ => continue,
            };
            *slot = Some(xml::unescape(value));
        }
        properties
    }

    /// Whether the weight is bold (`fo:font-weight` `bold` or 600 and up).
    #[must_use]
    pub fn bold(&self) -> Option<bool> {
        self.font_weight.as_deref().map(|weight| match weight {
            "bold" => true,
            "normal" => false,
            numeric => numeric.parse::<u16>().is_ok_and(|weight| weight >= 600),
        })
    }

    /// Whether the posture is italic or oblique.
    #[must_use]
    pub fn italic(&self) -> Option<bool> {
        self.font_style
            .as_deref()
            .map(|style| matches!(style, "italic" | "oblique"))
    }

    /// Whether text is underlined (`style:text-underline-style` other than
    /// `none`).
    #[must_use]
    pub fn underline(&self) -> Option<bool> {
        self.underline.as_deref().map(|style| style != "none")
    }

    /// The raw `style:text-underline-style`.
    #[must_use]
    pub fn underline_style(&self) -> Option<&str> {
        self.underline.as_deref()
    }

    /// Whether text is struck through.
    #[must_use]
    pub fn strike(&self) -> Option<bool> {
        self.line_through.as_deref().map(|style| style != "none")
    }

    /// `fo:font-size` as written (`12pt`, `120%`).
    #[must_use]
    pub fn size(&self) -> Option<&str> {
        self.font_size.as_deref()
    }

    /// `fo:color` as written (`#rrggbb`).
    #[must_use]
    pub fn color(&self) -> Option<&str> {
        self.color.as_deref()
    }

    /// The font: `style:font-name`, else `fo:font-family` without quotes.
    #[must_use]
    pub fn font(&self) -> Option<&str> {
        self.font_name.as_deref().or_else(|| {
            self.font_family
                .as_deref()
                .map(|family| family.trim_matches(|c| c == '\'' || c == '"'))
        })
    }

    /// `fo:background-color` as written (`#rrggbb` or `transparent`).
    #[must_use]
    pub fn highlight(&self) -> Option<&str> {
        self.background.as_deref()
    }

    /// Whether the text is raised (`style:text-position` `super` or a
    /// positive offset).
    #[must_use]
    pub fn superscript(&self) -> Option<bool> {
        self.position.as_deref().map(|position| {
            let offset = position.split_whitespace().next().unwrap_or("0%");
            offset == "super"
                || offset
                    .trim_end_matches('%')
                    .parse::<f64>()
                    .is_ok_and(|v| v > 0.0)
        })
    }

    /// Whether the text is lowered (`style:text-position` `sub` or a negative
    /// offset).
    #[must_use]
    pub fn subscript(&self) -> Option<bool> {
        self.position.as_deref().map(|position| {
            let offset = position.split_whitespace().next().unwrap_or("0%");
            offset == "sub"
                || offset
                    .trim_end_matches('%')
                    .parse::<f64>()
                    .is_ok_and(|v| v < 0.0)
        })
    }

    /// Whether letters are capitalized (`fo:text-transform` `uppercase`).
    #[must_use]
    pub fn caps(&self) -> Option<bool> {
        self.transform.as_deref().map(|value| value == "uppercase")
    }

    /// Whether letters are small capitals (`fo:font-variant`).
    #[must_use]
    pub fn small_caps(&self) -> Option<bool> {
        self.variant.as_deref().map(|value| value == "small-caps")
    }
}

/// Paragraph properties a style sets (`style:paragraph-properties`).
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct ParagraphProperties<'a> {
    align: Option<Cow<'a, str>>,
    margin_top: Option<Cow<'a, str>>,
    margin_bottom: Option<Cow<'a, str>>,
    margin_left: Option<Cow<'a, str>>,
    margin_right: Option<Cow<'a, str>>,
    text_indent: Option<Cow<'a, str>>,
    line_height: Option<Cow<'a, str>>,
    keep_with_next: Option<Cow<'a, str>>,
    break_before: Option<Cow<'a, str>>,
    background: Option<Cow<'a, str>>,
}

impl<'a> ParagraphProperties<'a> {
    pub(crate) fn read(tag: &'a str, prefixes: &Prefixes) -> Self {
        let mut properties = Self::default();
        for (name, value) in xml::attributes(tag) {
            let Some((ns, local)) = prefixes.resolve(name.as_bytes()) else {
                continue;
            };
            let slot = match (ns, local) {
                (Ns::Fo, b"text-align") => &mut properties.align,
                (Ns::Fo, b"margin-top") => &mut properties.margin_top,
                (Ns::Fo, b"margin-bottom") => &mut properties.margin_bottom,
                (Ns::Fo, b"margin-left") => &mut properties.margin_left,
                (Ns::Fo, b"margin-right") => &mut properties.margin_right,
                (Ns::Fo, b"text-indent") => &mut properties.text_indent,
                (Ns::Fo, b"line-height") => &mut properties.line_height,
                (Ns::Fo, b"keep-with-next") => &mut properties.keep_with_next,
                (Ns::Fo, b"break-before") => &mut properties.break_before,
                (Ns::Fo, b"background-color") => &mut properties.background,
                _ => continue,
            };
            *slot = Some(xml::unescape(value));
        }
        properties
    }

    /// The alignment, normalized to `left`, `right`, `center` or `justify`
    /// (`start` reads as `left` and `end` as `right`).
    #[must_use]
    pub fn align(&self) -> Option<&str> {
        self.align.as_deref().map(|align| match align {
            "start" | "left" => "left",
            "end" | "right" => "right",
            "center" => "center",
            "justify" => "justify",
            other => other,
        })
    }

    /// `fo:margin-top` as written.
    #[must_use]
    pub fn space_before(&self) -> Option<&str> {
        self.margin_top.as_deref()
    }

    /// `fo:margin-bottom` as written.
    #[must_use]
    pub fn space_after(&self) -> Option<&str> {
        self.margin_bottom.as_deref()
    }

    /// `fo:margin-left` as written.
    #[must_use]
    pub fn indent(&self) -> Option<&str> {
        self.margin_left.as_deref()
    }

    /// `fo:margin-right` as written.
    #[must_use]
    pub fn right_indent(&self) -> Option<&str> {
        self.margin_right.as_deref()
    }

    /// `fo:text-indent` as written.
    #[must_use]
    pub fn first_line_indent(&self) -> Option<&str> {
        self.text_indent.as_deref()
    }

    /// `fo:line-height` as written (`115%`, `0.5cm`).
    #[must_use]
    pub fn line_height(&self) -> Option<&str> {
        self.line_height.as_deref()
    }

    /// Whether the paragraph keeps with the next (`fo:keep-with-next`
    /// `always`).
    #[must_use]
    pub fn keep_next(&self) -> Option<bool> {
        self.keep_with_next
            .as_deref()
            .map(|value| value == "always")
    }

    /// Whether a page break precedes the paragraph (`fo:break-before`
    /// `page`).
    #[must_use]
    pub fn page_break_before(&self) -> Option<bool> {
        self.break_before.as_deref().map(|value| value == "page")
    }

    /// `fo:background-color` as written.
    #[must_use]
    pub fn shading(&self) -> Option<&str> {
        self.background.as_deref()
    }
}

/// Cell properties a style sets (`style:table-cell-properties`).
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct CellProperties<'a> {
    background: Option<Cow<'a, str>>,
    vertical_align: Option<Cow<'a, str>>,
}

impl<'a> CellProperties<'a> {
    pub(crate) fn read(tag: &'a str, prefixes: &Prefixes) -> Self {
        let mut properties = Self::default();
        for (name, value) in xml::attributes(tag) {
            match prefixes.resolve(name.as_bytes()) {
                Some((Ns::Fo, b"background-color")) => {
                    properties.background = Some(xml::unescape(value));
                },
                Some((Ns::Style, b"vertical-align")) => {
                    properties.vertical_align = Some(xml::unescape(value));
                },
                _ => {},
            }
        }
        properties
    }

    /// `fo:background-color` as written.
    #[must_use]
    pub fn shading(&self) -> Option<&str> {
        self.background.as_deref()
    }

    /// `style:vertical-align` as written.
    #[must_use]
    pub fn vertical_align(&self) -> Option<&str> {
        self.vertical_align.as_deref()
    }
}

/// Character properties to set on a run, paragraph or heading. Each field
/// left `None` keeps what the node's style had.
#[derive(Clone, Debug, Default, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub struct TextPropertyUpdate {
    /// Bold (`fo:font-weight`, with its Asian and complex twins).
    pub bold: Option<bool>,
    /// Italic (`fo:font-style`, with its twins).
    pub italic: Option<bool>,
    /// Single underline, or none.
    pub underline: Option<bool>,
    /// Single strike-through, or none.
    pub strike: Option<bool>,
    /// The font size, an ODF length such as `12pt`.
    pub size: Option<String>,
    /// The colour, `#rrggbb`.
    pub color: Option<String>,
    /// The font family name.
    pub font: Option<String>,
    /// The highlight (`#rrggbb` or `transparent`).
    pub highlight: Option<String>,
    /// Superscript (`true`), or normal position (`false`).
    pub superscript: Option<bool>,
    /// Subscript (`true`), or normal position (`false`).
    pub subscript: Option<bool>,
    /// All capitals, or none.
    pub caps: Option<bool>,
    /// Small capitals, or normal.
    pub small_caps: Option<bool>,
}

impl TextPropertyUpdate {
    /// An edit that changes nothing.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Whether the edit sets nothing.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        *self == Self::default()
    }

    /// Merge `later` over this edit: its set fields win.
    pub fn merge(&mut self, later: &Self) {
        macro_rules! take {
            ($($field:ident),*) => {$(
                if later.$field.is_some() {
                    self.$field.clone_from(&later.$field);
                }
            )*};
        }
        take!(
            bold,
            italic,
            underline,
            strike,
            size,
            color,
            font,
            highlight,
            superscript,
            subscript,
            caps,
            small_caps
        );
    }

    pub(crate) fn validate(&self) -> Result<()> {
        if let Some(size) = &self.size {
            check_length(size, "font size", false)?;
        }
        if let Some(color) = &self.color {
            check_color(color, false)?;
        }
        if let Some(color) = &self.highlight {
            check_color(color, true)?;
        }
        if let Some(font) = &self.font {
            if font.is_empty() || font.len() > 256 {
                return Err(Error::InvalidFormat(
                    "ODT font names hold 1 to 256 bytes".to_string(),
                ));
            }
            xml::check_text(font, "font name")?;
        }
        Ok(())
    }

    /// The `(namespace, local, value)` attributes this edit writes.
    pub(crate) fn attributes(&self) -> Vec<(Ns, &'static str, String)> {
        let mut out = Vec::new();
        if let Some(bold) = self.bold {
            let value = if bold { "bold" } else { "normal" };
            out.push((Ns::Fo, "font-weight", value.to_string()));
            out.push((Ns::Style, "font-weight-asian", value.to_string()));
            out.push((Ns::Style, "font-weight-complex", value.to_string()));
        }
        if let Some(italic) = self.italic {
            let value = if italic { "italic" } else { "normal" };
            out.push((Ns::Fo, "font-style", value.to_string()));
            out.push((Ns::Style, "font-style-asian", value.to_string()));
            out.push((Ns::Style, "font-style-complex", value.to_string()));
        }
        if let Some(underline) = self.underline {
            if underline {
                out.push((Ns::Style, "text-underline-style", "solid".to_string()));
                out.push((Ns::Style, "text-underline-width", "auto".to_string()));
                out.push((Ns::Style, "text-underline-color", "font-color".to_string()));
            } else {
                out.push((Ns::Style, "text-underline-style", "none".to_string()));
            }
        }
        if let Some(strike) = self.strike {
            let value = if strike { "solid" } else { "none" };
            out.push((Ns::Style, "text-line-through-style", value.to_string()));
        }
        if let Some(size) = &self.size {
            out.push((Ns::Fo, "font-size", size.clone()));
            out.push((Ns::Style, "font-size-asian", size.clone()));
            out.push((Ns::Style, "font-size-complex", size.clone()));
        }
        if let Some(color) = &self.color {
            out.push((Ns::Fo, "color", color.to_ascii_lowercase()));
        }
        if let Some(font) = &self.font {
            let quoted = if font.contains(' ') {
                format!("'{font}'")
            } else {
                font.clone()
            };
            out.push((Ns::Fo, "font-family", quoted));
        }
        if let Some(highlight) = &self.highlight {
            out.push((Ns::Fo, "background-color", highlight.to_ascii_lowercase()));
        }
        match (self.superscript, self.subscript) {
            (Some(true), _) => out.push((Ns::Style, "text-position", "super 58%".to_string())),
            (_, Some(true)) => out.push((Ns::Style, "text-position", "sub 58%".to_string())),
            (Some(false), _) | (_, Some(false)) => {
                out.push((Ns::Style, "text-position", "0% 100%".to_string()));
            },
            (None, None) => {},
        }
        if let Some(caps) = self.caps {
            let value = if caps { "uppercase" } else { "none" };
            out.push((Ns::Fo, "text-transform", value.to_string()));
        }
        if let Some(small_caps) = self.small_caps {
            let value = if small_caps { "small-caps" } else { "normal" };
            out.push((Ns::Fo, "font-variant", value.to_string()));
        }
        out
    }
}

/// Paragraph properties to set on a paragraph or heading.
#[derive(Clone, Debug, Default, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub struct ParagraphPropertyUpdate {
    /// `left`, `right`, `center` or `justify`.
    pub align: Option<String>,
    /// Space above, an ODF length.
    pub space_before: Option<String>,
    /// Space below, an ODF length.
    pub space_after: Option<String>,
    /// Left indent, an ODF length (may be negative).
    pub indent: Option<String>,
    /// Right indent, an ODF length (may be negative).
    pub right_indent: Option<String>,
    /// First-line indent, an ODF length (may be negative).
    pub first_line_indent: Option<String>,
    /// Line height: a percentage (`115%`) or a length.
    pub line_height: Option<String>,
    /// Keep with the next paragraph.
    pub keep_next: Option<bool>,
    /// Start on a new page.
    pub page_break_before: Option<bool>,
    /// Background colour, `#rrggbb` or `transparent`.
    pub shading: Option<String>,
    /// The paragraph's tab stops (`style:tab-stops`), replacing the ones
    /// its automatic style declares; an empty list declares none (record
    /// 0328).
    pub tab_stops: Option<Vec<TabStopSpec>>,
}

/// One tab stop a [`ParagraphPropertyUpdate`] declares (`style:tab-stop`).
#[derive(Clone, Debug, Default, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub struct TabStopSpec {
    /// `style:position`, an ODF length from the paragraph's start indent.
    pub position: String,
    /// `style:type`: `left` (the default), `center`, `right` or `char`.
    pub kind: Option<String>,
    /// `style:leader-style`: `none`, `solid`, `dotted`, `dash`,
    /// `long-dash`, `dot-dash`, `dot-dot-dash` or `wave`.
    pub leader_style: Option<String>,
    /// `style:leader-text`: one leader character.
    pub leader_text: Option<String>,
    /// `style:char`: the alignment character of a `char` stop.
    pub character: Option<String>,
}

impl TabStopSpec {
    /// A left tab stop at `position` (an ODF length).
    #[must_use]
    pub fn new(position: impl Into<String>) -> Self {
        Self {
            position: position.into(),
            ..Self::default()
        }
    }

    fn validate(&self) -> Result<()> {
        check_length(&self.position, "tab stop position", false)?;
        if let Some(kind) = &self.kind {
            if !matches!(kind.as_str(), "left" | "center" | "right" | "char") {
                return Err(Error::InvalidFormat(
                    "ODT tab stop type is not left, center, right or char".to_owned(),
                ));
            }
        }
        if let Some(style) = &self.leader_style {
            if !matches!(
                style.as_str(),
                "none"
                    | "solid"
                    | "dotted"
                    | "dash"
                    | "long-dash"
                    | "dot-dash"
                    | "dot-dot-dash"
                    | "wave"
            ) {
                return Err(Error::InvalidFormat(
                    "ODT tab stop leader style is not an ODF line style".to_owned(),
                ));
            }
        }
        for (value, field) in [
            (&self.leader_text, "leader text"),
            (&self.character, "alignment character"),
        ] {
            if let Some(value) = value {
                let mut chars = value.chars();
                if chars.next().is_none_or(char::is_control) || chars.next().is_some() {
                    return Err(Error::InvalidFormat(format!(
                        "ODT tab stop {field} is not one character"
                    )));
                }
            }
        }
        if self.kind.as_deref() == Some("char") && self.character.is_none() {
            return Err(Error::InvalidFormat(
                "an ODT char tab stop needs its alignment character".to_owned(),
            ));
        }
        Ok(())
    }

    fn push_xml(&self, out: &mut String, prefixes: &Prefixes) {
        out.push('<');
        prefixes.push_name(out, Ns::Style, "tab-stop");
        for (local, value) in [
            ("position", Some(&self.position)),
            ("type", self.kind.as_ref()),
            ("char", self.character.as_ref()),
            ("leader-style", self.leader_style.as_ref()),
            ("leader-text", self.leader_text.as_ref()),
        ] {
            if let Some(value) = value {
                out.push(' ');
                prefixes.push_name(out, Ns::Style, local);
                out.push_str("=\"");
                xml::escape_into(out, value);
                out.push('"');
            }
        }
        out.push_str("/>");
    }
}

impl ParagraphPropertyUpdate {
    /// An edit that changes nothing.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Whether the edit sets nothing.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        *self == Self::default()
    }

    /// Merge `later` over this edit: its set fields win.
    pub fn merge(&mut self, later: &Self) {
        macro_rules! take {
            ($($field:ident),*) => {$(
                if later.$field.is_some() {
                    self.$field.clone_from(&later.$field);
                }
            )*};
        }
        take!(
            align,
            space_before,
            space_after,
            indent,
            right_indent,
            first_line_indent,
            line_height,
            keep_next,
            page_break_before,
            shading,
            tab_stops
        );
    }

    pub(crate) fn validate(&self) -> Result<()> {
        if let Some(align) = &self.align {
            if !matches!(align.as_str(), "left" | "right" | "center" | "justify") {
                return Err(Error::InvalidFormat(
                    "ODT paragraph alignment is not left, right, center or justify".to_owned(),
                ));
            }
        }
        for (value, field, signed) in [
            (&self.space_before, "space before", false),
            (&self.space_after, "space after", false),
            (&self.indent, "indent", true),
            (&self.right_indent, "right indent", true),
            (&self.first_line_indent, "first-line indent", true),
        ] {
            if let Some(value) = value {
                check_length(value, field, signed)?;
            }
        }
        if let Some(height) = &self.line_height {
            if !(height.ends_with('%')
                && height
                    .trim_end_matches('%')
                    .parse::<f64>()
                    .is_ok_and(|value| value.is_finite() && value > 0.0))
            {
                check_length(height, "line height", false)?;
            }
        }
        if let Some(color) = &self.shading {
            check_color(color, true)?;
        }
        if let Some(stops) = &self.tab_stops {
            if stops.len() > 256 {
                return Err(Error::InvalidFormat(
                    "an ODT paragraph declares at most 256 tab stops".to_owned(),
                ));
            }
            for stop in stops {
                stop.validate()?;
            }
        }
        Ok(())
    }

    /// The `style:tab-stops` element the edit declares, if it sets tab
    /// stops.
    pub(crate) fn tab_stops_xml(&self, prefixes: &Prefixes) -> Option<String> {
        let stops = self.tab_stops.as_ref()?;
        let mut out = String::from("<");
        prefixes.push_name(&mut out, Ns::Style, "tab-stops");
        if stops.is_empty() {
            out.push_str("/>");
            return Some(out);
        }
        out.push('>');
        for stop in stops {
            stop.push_xml(&mut out, prefixes);
        }
        out.push_str("</");
        prefixes.push_name(&mut out, Ns::Style, "tab-stops");
        out.push('>');
        Some(out)
    }

    pub(crate) fn attributes(&self) -> Vec<(Ns, &'static str, String)> {
        let mut out = Vec::new();
        if let Some(align) = &self.align {
            let value = match align.as_str() {
                "left" => "start",
                "right" => "end",
                other => other,
            };
            out.push((Ns::Fo, "text-align", value.to_string()));
        }
        for (value, local) in [
            (&self.space_before, "margin-top"),
            (&self.space_after, "margin-bottom"),
            (&self.indent, "margin-left"),
            (&self.right_indent, "margin-right"),
            (&self.first_line_indent, "text-indent"),
            (&self.line_height, "line-height"),
        ] {
            if let Some(value) = value {
                out.push((Ns::Fo, local, value.clone()));
            }
        }
        if let Some(keep) = self.keep_next {
            let value = if keep { "always" } else { "auto" };
            out.push((Ns::Fo, "keep-with-next", value.to_string()));
        }
        if let Some(page) = self.page_break_before {
            let value = if page { "page" } else { "auto" };
            out.push((Ns::Fo, "break-before", value.to_string()));
        }
        if let Some(color) = &self.shading {
            out.push((Ns::Fo, "background-color", color.to_ascii_lowercase()));
        }
        out
    }
}

/// An ODF length: a decimal number and one of the units ODF admits.
pub(crate) fn check_length(value: &str, field: &str, signed: bool) -> Result<()> {
    let unit_start = value
        .find(|c: char| c.is_ascii_alphabetic())
        .unwrap_or(value.len());
    let (number, unit) = value.split_at(unit_start);
    let ok = matches!(unit, "pt" | "pc" | "cm" | "mm" | "in" | "px")
        && number.parse::<f64>().is_ok_and(|number| {
            number.is_finite() && (signed || number >= 0.0) && number.abs() < 1.0e6
        })
        && !number.starts_with('+');
    if ok {
        Ok(())
    } else {
        Err(Error::InvalidFormat(format!(
            "ODT {field} is not an ODF length"
        )))
    }
}

fn check_color(value: &str, transparent: bool) -> Result<()> {
    let hex = value
        .strip_prefix('#')
        .is_some_and(|hex| hex.len() == 6 && hex.bytes().all(|byte| byte.is_ascii_hexdigit()));
    if hex || (transparent && value == "transparent") {
        Ok(())
    } else {
        Err(Error::InvalidFormat("ODT colour is not #rrggbb".to_owned()))
    }
}
