//! XLSX theme part (`/xl/theme/theme1.xml`, ECMA-376 DrawingML theme).
//!
//! The theme defines the color scheme that theme-indexed colors in fonts,
//! fills, borders, tab colors, and charts resolve against. The format scheme
//! and any extension content are preserved verbatim as inert XML; nothing is
//! rendered.

use quick_xml::{
    events::Event,
    name::{Namespace, ResolveResult},
    reader::NsReader,
};

use crate::error::{Error, Result as SheetResult};
use litchi_ooxml_common::xml::unqualified_attribute_value;

const DRAWINGML: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/main";

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

/// One of the twelve theme color slots, in `a:clrScheme` order.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ThemeColorSlot {
    /// Dark 1 (`dk1`), usually text.
    Dk1,
    /// Light 1 (`lt1`), usually background.
    Lt1,
    /// Dark 2 (`dk2`).
    Dk2,
    /// Light 2 (`lt2`).
    Lt2,
    /// Accent 1 (`accent1`).
    Accent1,
    /// Accent 2 (`accent2`).
    Accent2,
    /// Accent 3 (`accent3`).
    Accent3,
    /// Accent 4 (`accent4`).
    Accent4,
    /// Accent 5 (`accent5`).
    Accent5,
    /// Accent 6 (`accent6`).
    Accent6,
    /// Hyperlink (`hlink`).
    Hyperlink,
    /// Followed hyperlink (`folHlink`).
    FollowedHyperlink,
}

impl ThemeColorSlot {
    /// All twelve slots in scheme order.
    pub const ALL: [Self; 12] = [
        Self::Dk1,
        Self::Lt1,
        Self::Dk2,
        Self::Lt2,
        Self::Accent1,
        Self::Accent2,
        Self::Accent3,
        Self::Accent4,
        Self::Accent5,
        Self::Accent6,
        Self::Hyperlink,
        Self::FollowedHyperlink,
    ];

    /// Zero-based slot index in the scheme.
    pub const fn index(self) -> usize {
        match self {
            Self::Dk1 => 0,
            Self::Lt1 => 1,
            Self::Dk2 => 2,
            Self::Lt2 => 3,
            Self::Accent1 => 4,
            Self::Accent2 => 5,
            Self::Accent3 => 6,
            Self::Accent4 => 7,
            Self::Accent5 => 8,
            Self::Accent6 => 9,
            Self::Hyperlink => 10,
            Self::FollowedHyperlink => 11,
        }
    }

    /// The `a:clrScheme` child element name of this slot.
    pub const fn element_name(self) -> &'static str {
        match self {
            Self::Dk1 => "dk1",
            Self::Lt1 => "lt1",
            Self::Dk2 => "dk2",
            Self::Lt2 => "lt2",
            Self::Accent1 => "accent1",
            Self::Accent2 => "accent2",
            Self::Accent3 => "accent3",
            Self::Accent4 => "accent4",
            Self::Accent5 => "accent5",
            Self::Accent6 => "accent6",
            Self::Hyperlink => "hlink",
            Self::FollowedHyperlink => "folHlink",
        }
    }
}

/// The value of one theme color slot.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ThemeColorValue {
    /// `a:srgbClr`: an explicit RGB triple.
    Srgb([u8; 3]),
    /// `a:sysClr`: a system color name with its cached RGB triple.
    System {
        /// System color name (`val`, e.g. `windowText`).
        name: String,
        /// Cached RGB triple (`lastClr`).
        last_rgb: [u8; 3],
    },
}

fn parse_hex_rgb(value: &str, context: &str) -> SheetResult<[u8; 3]> {
    if value.len() != 6 || !value.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        return Err(invalid(format!("{context} is not a 6-digit hex RGB value")));
    }
    let byte = |offset: usize| {
        u8::from_str_radix(&value[offset..offset + 2], 16)
            .map_err(|_| invalid(format!("{context} is not a 6-digit hex RGB value")))
    };
    Ok([byte(0)?, byte(2)?, byte(4)?])
}

/// A parsed XLSX theme part.
#[derive(Debug, Clone)]
pub struct Theme {
    name: Option<String>,
    color_scheme_name: String,
    colors: [ThemeColorValue; 12],
    major_font: Option<String>,
    minor_font: Option<String>,
    /// Verbatim `a:fmtScheme` XML, preserved inert.
    format_scheme_xml: String,
}

impl Theme {
    /// The theme name (`a:theme name`), when present.
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    /// The color scheme name (`a:clrScheme name`).
    pub fn color_scheme_name(&self) -> &str {
        &self.color_scheme_name
    }

    /// The major (heading) latin typeface, when declared.
    pub fn major_font(&self) -> Option<&str> {
        self.major_font.as_deref()
    }

    /// The minor (body) latin typeface, when declared.
    pub fn minor_font(&self) -> Option<&str> {
        self.minor_font.as_deref()
    }

    /// The value of one color slot.
    pub fn color(&self, slot: ThemeColorSlot) -> &ThemeColorValue {
        &self.colors[slot.index()]
    }

    /// The effective RGB triple of one color slot (the cached value for
    /// system colors).
    pub fn rgb(&self, slot: ThemeColorSlot) -> [u8; 3] {
        match self.color(slot) {
            ThemeColorValue::Srgb(rgb) => *rgb,
            ThemeColorValue::System { last_rgb, .. } => *last_rgb,
        }
    }

    /// Verbatim `a:fmtScheme` XML, preserved inert.
    pub fn format_scheme_xml(&self) -> &str {
        &self.format_scheme_xml
    }

    /// Parse a theme part.
    pub fn parse(xml: &str) -> SheetResult<Self> {
        let mut reader = NsReader::from_str(xml);
        let mut buffer = Vec::new();
        let mut name = None;
        let mut color_scheme_name = String::new();
        let mut colors: [Option<ThemeColorValue>; 12] = std::array::from_fn(|_| None);
        let mut major_font = None;
        let mut minor_font = None;
        let mut format_scheme_xml = String::new();
        let mut slot_index = None::<usize>;
        let mut font_kind = None::<bool>; // true = major, false = minor
        let mut fmt_start = None::<usize>;
        let mut fmt_depth = 0usize;

        loop {
            let decoder = reader.decoder();
            let event_start = reader.buffer_position() as usize;
            let (namespace, event) = reader
                .read_resolved_event_into(&mut buffer)
                .map_err(|error| invalid(format!("theme XML error: {error}")))?;
            let drawingml =
                matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == DRAWINGML);
            let event = event.into_owned();
            let event_end = reader.buffer_position() as usize;
            match event {
                Event::Start(element) => {
                    let local_name = element.local_name();
                    let local = local_name.as_ref();
                    if drawingml && local == b"theme" {
                        name = unqualified_attribute_value(&element, b"name", decoder)?;
                    } else if drawingml && local == b"clrScheme" {
                        color_scheme_name =
                            unqualified_attribute_value(&element, b"name", decoder)?
                                .unwrap_or_default();
                        slot_index = Some(0);
                    } else if drawingml && local == b"fmtScheme" {
                        fmt_start = Some(event_start);
                        fmt_depth = 1;
                    } else if fmt_start.is_some() {
                        fmt_depth += 1;
                    } else if let Some(index) = slot_index.as_mut() {
                        let expected = ThemeColorSlot::ALL[*index];
                        if !(drawingml && local == expected.element_name().as_bytes()) {
                            return Err(invalid(format!(
                                "a:clrScheme slot {} must be <a:{}>",
                                index,
                                expected.element_name()
                            )));
                        }
                    } else if drawingml && local == b"majorFont" {
                        font_kind = Some(true);
                    } else if drawingml && local == b"minorFont" {
                        font_kind = Some(false);
                    } else if drawingml && local == b"latin" {
                        let typeface = unqualified_attribute_value(&element, b"typeface", decoder)?;
                        match font_kind {
                            Some(true) => major_font = typeface,
                            Some(false) => minor_font = typeface,
                            None => {},
                        }
                    }
                },
                Event::Empty(element) => {
                    let local_name = element.local_name();
                    let local = local_name.as_ref();
                    if drawingml && local == b"latin" {
                        let typeface = unqualified_attribute_value(&element, b"typeface", decoder)?;
                        match font_kind {
                            Some(true) => major_font = typeface,
                            Some(false) => minor_font = typeface,
                            None => {},
                        }
                    } else if let Some(index) = slot_index.as_mut() {
                        if drawingml && local == b"srgbClr" {
                            let value = unqualified_attribute_value(&element, b"val", decoder)?
                                .ok_or_else(|| invalid("a:srgbClr is missing its val attribute"))?;
                            colors[*index] = Some(ThemeColorValue::Srgb(parse_hex_rgb(
                                &value,
                                "a:srgbClr val",
                            )?));
                            *index += 1;
                        } else if drawingml && local == b"sysClr" {
                            let color_name =
                                unqualified_attribute_value(&element, b"val", decoder)?
                                    .ok_or_else(|| {
                                        invalid("a:sysClr is missing its val attribute")
                                    })?;
                            let last = unqualified_attribute_value(&element, b"lastClr", decoder)?
                                .ok_or_else(|| {
                                    invalid("a:sysClr is missing its lastClr attribute")
                                })?;
                            colors[*index] = Some(ThemeColorValue::System {
                                name: color_name,
                                last_rgb: parse_hex_rgb(&last, "a:sysClr lastClr")?,
                            });
                            *index += 1;
                        }
                    }
                },
                Event::End(element) => {
                    let local_name = element.local_name();
                    let local = local_name.as_ref();
                    if fmt_depth > 0 {
                        fmt_depth -= 1;
                        if fmt_depth == 0 && drawingml && local == b"fmtScheme" {
                            format_scheme_xml =
                                xml[fmt_start.expect("fmt start")..event_end].to_string();
                        }
                    } else if drawingml && local == b"clrScheme" {
                        slot_index = None;
                    } else if drawingml && (local == b"majorFont" || local == b"minorFont") {
                        font_kind = None;
                    }
                },
                Event::Eof => break,
                _ => {},
            }
            buffer.clear();
        }

        if let Some(index) = slot_index {
            if index != 12 {
                return Err(invalid(format!(
                    "a:clrScheme declares {index} of the required 12 color slots"
                )));
            }
            return Err(invalid("a:clrScheme is unterminated"));
        }
        let colors = colors
            .into_iter()
            .enumerate()
            .map(|(index, value)| {
                value.ok_or_else(|| {
                    invalid(format!(
                        "a:clrScheme is missing the <a:{}> slot",
                        ThemeColorSlot::ALL[index].element_name()
                    ))
                })
            })
            .collect::<SheetResult<Vec<_>>>()?;
        let colors: [ThemeColorValue; 12] = colors
            .try_into()
            .map_err(|_| invalid("a:clrScheme must contain exactly 12 slots"))?;
        Ok(Self {
            name,
            color_scheme_name,
            colors,
            major_font,
            minor_font,
            format_scheme_xml,
        })
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    const THEME: &str = concat!(
        r#"<?xml version="1.0"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office">"#,
        r#"<a:themeElements>"#,
        r#"<a:clrScheme name="Office">"#,
        r#"<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>"#,
        r#"<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>"#,
        r#"<a:dk2><a:srgbClr val="1F497D"/></a:dk2>"#,
        r#"<a:lt2><a:srgbClr val="EEECE1"/></a:lt2>"#,
        r#"<a:accent1><a:srgbClr val="4F81BD"/></a:accent1>"#,
        r#"<a:accent2><a:srgbClr val="C0504D"/></a:accent2>"#,
        r#"<a:accent3><a:srgbClr val="9BBB59"/></a:accent3>"#,
        r#"<a:accent4><a:srgbClr val="8064A2"/></a:accent4>"#,
        r#"<a:accent5><a:srgbClr val="4BACC6"/></a:accent5>"#,
        r#"<a:accent6><a:srgbClr val="F79646"/></a:accent6>"#,
        r#"<a:hlink><a:srgbClr val="0000FF"/></a:hlink>"#,
        r#"<a:folHlink><a:srgbClr val="800080"/></a:folHlink>"#,
        r#"</a:clrScheme>"#,
        r#"<a:fontScheme name="Office">"#,
        r#"<a:majorFont><a:latin typeface="Cambria"/></a:majorFont>"#,
        r#"<a:minorFont><a:latin typeface="Calibri"/></a:minorFont>"#,
        r#"</a:fontScheme>"#,
        r#"<a:fmtScheme name="Office"><a:fillStyleLst/></a:fmtScheme>"#,
        r#"</a:themeElements>"#,
        r#"</a:theme>"#,
    );

    #[test]
    fn parses_theme_colors_and_fonts() {
        let theme = Theme::parse(THEME).unwrap();
        assert_eq!(theme.name(), Some("Office"));
        assert_eq!(theme.color_scheme_name(), "Office");
        assert_eq!(theme.major_font(), Some("Cambria"));
        assert_eq!(theme.minor_font(), Some("Calibri"));
        assert_eq!(theme.rgb(ThemeColorSlot::Dk1), [0, 0, 0]);
        assert_eq!(theme.rgb(ThemeColorSlot::Accent1), [0x4F, 0x81, 0xBD]);
        assert_eq!(
            theme.rgb(ThemeColorSlot::FollowedHyperlink),
            [0x80, 0, 0x80]
        );
        assert_eq!(
            theme.color(ThemeColorSlot::Dk1),
            &ThemeColorValue::System {
                name: "windowText".to_string(),
                last_rgb: [0, 0, 0]
            }
        );
        assert_eq!(
            theme.color(ThemeColorSlot::Lt2),
            &ThemeColorValue::Srgb([0xEE, 0xEC, 0xE1])
        );
        assert_eq!(
            theme.format_scheme_xml(),
            "<a:fmtScheme name=\"Office\"><a:fillStyleLst/></a:fmtScheme>"
        );
    }

    #[test]
    fn rejects_malformed_themes() {
        // Wrong slot order.
        let bad = THEME
            .replace("<a:lt1>", "<a:dk2>")
            .replace("</a:lt1>", "</a:dk2>");
        assert!(Theme::parse(&bad).is_err());
        // Missing a slot.
        let bad = THEME.replace("<a:folHlink><a:srgbClr val=\"800080\"/></a:folHlink>", "");
        assert!(Theme::parse(&bad).is_err());
        // Bad hex.
        let bad = THEME.replace("4F81BD", "4F81BZ");
        assert!(Theme::parse(&bad).is_err());
    }
}
