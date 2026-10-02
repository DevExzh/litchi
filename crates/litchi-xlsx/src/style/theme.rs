//! XLSX theme part (`/xl/theme/theme1.xml`, ECMA-376 DrawingML theme).
//!
//! The theme defines the color scheme that theme-indexed colors in fonts,
//! fills, borders, tab colors, and charts resolve against. The format scheme
//! and any extension content are preserved verbatim as inert XML; nothing is
//! rendered.

use litchi_core::xml::ReaderOrigin;
use quick_xml::{
    events::Event,
    name::{Namespace, ResolveResult},
    reader::NsReader,
};

use crate::error::{Error, Result as SheetResult};
use crate::leniency::{Location, ThemeDefect, ToleranceReport};
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
    #[must_use]
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

    /// The slot named by an `a:clrScheme` child element, if any.
    #[must_use]
    pub fn from_element_name(name: &[u8]) -> Option<Self> {
        Self::ALL
            .into_iter()
            .find(|slot| slot.element_name().as_bytes() == name)
    }

    /// The Office theme's value for this slot, which a repaired theme
    /// substitutes for a missing or malformed slot (leniency finding 4.11).
    #[must_use]
    pub fn default_value(self) -> ThemeColorValue {
        let rgb = match self {
            Self::Dk1 => {
                return ThemeColorValue::System {
                    name: "windowText".to_owned(),
                    last_rgb: [0, 0, 0],
                };
            },
            Self::Lt1 => {
                return ThemeColorValue::System {
                    name: "window".to_owned(),
                    last_rgb: [0xFF, 0xFF, 0xFF],
                };
            },
            Self::Dk2 => [0x1F, 0x49, 0x7D],
            Self::Lt2 => [0xEE, 0xEC, 0xE1],
            Self::Accent1 => [0x4F, 0x81, 0xBD],
            Self::Accent2 => [0xC0, 0x50, 0x4D],
            Self::Accent3 => [0x9B, 0xBB, 0x59],
            Self::Accent4 => [0x80, 0x64, 0xA2],
            Self::Accent5 => [0x4B, 0xAC, 0xC6],
            Self::Accent6 => [0xF7, 0x96, 0x46],
            Self::Hyperlink => [0x00, 0x00, 0xFF],
            Self::FollowedHyperlink => [0x80, 0x00, 0x80],
        };
        ThemeColorValue::Srgb(rgb)
    }

    /// The `a:clrScheme` child element name of this slot.
    #[must_use]
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

/// The RGB value of a system color by name, for an `a:sysClr` whose
/// optional `lastClr` cache is absent or malformed: the two names Office
/// themes use, else black.
fn system_rgb(name: &str) -> [u8; 3] {
    match name {
        "window" | "highlightText" | "btnHighlight" => [0xFF, 0xFF, 0xFF],
        _ => [0, 0, 0],
    }
}

fn parse_hex_rgb(value: &str, context: &str) -> SheetResult<[u8; 3]> {
    if value.len() != 6 || !value.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        return Err(invalid(format!("{context} is not a 6-digit hex RGB value")));
    }
    let byte = |offset: usize| {
        u8::from_str_radix(&value[offset..offset + 2], 16)
            .map_err(|_source| invalid(format!("{context} is not a 6-digit hex RGB value")))
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
    /// Color-scheme repairs (leniency finding 4.11).
    tolerance: ToleranceReport,
}

impl Theme {
    /// The theme name (`a:theme name`), when present.
    #[must_use]
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    /// The color scheme name (`a:clrScheme name`).
    #[must_use]
    pub fn color_scheme_name(&self) -> &str {
        &self.color_scheme_name
    }

    /// The major (heading) latin typeface, when declared.
    #[must_use]
    pub fn major_font(&self) -> Option<&str> {
        self.major_font.as_deref()
    }

    /// The minor (body) latin typeface, when declared.
    #[must_use]
    pub fn minor_font(&self) -> Option<&str> {
        self.minor_font.as_deref()
    }

    /// The value of one color slot.
    #[must_use]
    pub fn color(&self, slot: ThemeColorSlot) -> &ThemeColorValue {
        &self.colors[slot.index()]
    }

    /// The effective RGB triple of one color slot (the cached value for
    /// system colors).
    #[must_use]
    pub fn rgb(&self, slot: ThemeColorSlot) -> [u8; 3] {
        match self.color(slot) {
            ThemeColorValue::Srgb(rgb) => *rgb,
            ThemeColorValue::System { last_rgb, .. } => *last_rgb,
        }
    }

    /// Verbatim `a:fmtScheme` XML, preserved inert.
    #[must_use]
    pub fn format_scheme_xml(&self) -> &str {
        &self.format_scheme_xml
    }

    /// The color-scheme repairs made while this theme was parsed, each a
    /// [`ThemeDefect`] at [`Location::ColorSchemeSlot`]; empty for a
    /// conformant scheme. See [`crate::leniency`].
    #[must_use]
    pub fn tolerance_report(&self) -> &ToleranceReport {
        &self.tolerance
    }

    /// Parse a theme part.
    ///
    /// The color scheme is read the way LibreOffice reads it (audit finding
    /// 4.11): slots are matched by element name in any order (the first
    /// occurrence wins), a missing slot or one whose color is malformed or
    /// not an `a:srgbClr`/`a:sysClr` takes the Office theme's value
    /// ([`ThemeColorSlot::default_value`]), and an `a:sysClr` without its
    /// optional `lastClr` cache resolves by name. Each repair is recorded in
    /// [`Theme::tolerance_report`]. Ill-formed XML, a theme without an
    /// `a:clrScheme`, and an unterminated scheme stay errors.
    pub fn parse(xml: &str) -> SheetResult<Self> {
        let mut reader = NsReader::from_str(xml);
        // `format_scheme_xml` is sliced from `xml` at reader positions, which
        // exclude a leading byte-order mark.
        let origin = ReaderOrigin::of(xml.as_bytes());
        let mut buffer = Vec::new();
        let mut name = None;
        let mut color_scheme_name = String::new();
        let mut colors: [Option<ThemeColorValue>; 12] = std::array::from_fn(|_| None);
        let mut tolerance = ToleranceReport::default();
        // Inside `a:clrScheme`: the highest slot seen so far, and the slot
        // whose element is open.
        let mut in_scheme = false;
        let mut seen_scheme = false;
        let mut highest_slot = None::<usize>;
        let mut open_slot = None::<ThemeColorSlot>;
        let mut scheme_depth = 0usize;
        let mut major_font = None;
        let mut minor_font = None;
        let mut format_scheme_xml = String::new();
        let mut font_kind = None::<bool>; // true = major, false = minor
        let mut fmt_start = None::<usize>;
        let mut fmt_depth = 0usize;

        loop {
            let decoder = reader.decoder();
            let event_start = origin
                .offset(reader.buffer_position())
                .ok_or_else(|| invalid("theme XML offset exceeds platform size"))?;
            let (namespace, event) = reader
                .read_resolved_event_into(&mut buffer)
                .map_err(|error| invalid(format!("theme XML error: {error}")))?;
            let drawingml =
                matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == DRAWINGML);
            let event = event.into_owned();
            let event_end = origin
                .offset(reader.buffer_position())
                .ok_or_else(|| invalid("theme XML offset exceeds platform size"))?;
            match event {
                Event::Start(element) => {
                    let local_name = element.local_name();
                    let local = local_name.as_ref();
                    if drawingml && local == b"theme" {
                        name = unqualified_attribute_value(&element, b"name", decoder)?;
                    } else if drawingml && local == b"clrScheme" && !in_scheme {
                        color_scheme_name =
                            unqualified_attribute_value(&element, b"name", decoder)?
                                .unwrap_or_default();
                        in_scheme = true;
                        seen_scheme = true;
                    } else if drawingml && local == b"fmtScheme" {
                        fmt_start = Some(event_start);
                        fmt_depth = 1;
                    } else if fmt_start.is_some() {
                        fmt_depth += 1;
                    } else if in_scheme {
                        scheme_depth += 1;
                        if scheme_depth == 1 {
                            open_slot = color_scheme_slot(
                                drawingml,
                                local,
                                &mut highest_slot,
                                &colors,
                                &mut tolerance,
                            );
                        } else if scheme_depth == 2
                            && let Some(slot) = open_slot
                        {
                            read_slot_color(
                                drawingml,
                                local,
                                &element,
                                decoder,
                                slot,
                                &mut colors,
                                &mut tolerance,
                            )?;
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
                    } else if in_scheme && fmt_start.is_none() {
                        if scheme_depth == 0 {
                            // An empty slot element names a slot without a
                            // color; it is defaulted below.
                            color_scheme_slot(
                                drawingml,
                                local,
                                &mut highest_slot,
                                &colors,
                                &mut tolerance,
                            );
                        } else if scheme_depth == 1
                            && let Some(slot) = open_slot
                        {
                            read_slot_color(
                                drawingml,
                                local,
                                &element,
                                decoder,
                                slot,
                                &mut colors,
                                &mut tolerance,
                            )?;
                        }
                    }
                },
                Event::End(element) => {
                    let local_name = element.local_name();
                    let local = local_name.as_ref();
                    if fmt_depth > 0 {
                        fmt_depth -= 1;
                        if fmt_depth == 0 && drawingml && local == b"fmtScheme" {
                            format_scheme_xml = xml[fmt_start.unwrap_or_else(|| {
                                crate::error::panic_missing_invariant("fmt start")
                            })..event_end]
                                .to_string();
                        }
                    } else if in_scheme && scheme_depth > 0 {
                        scheme_depth -= 1;
                        if scheme_depth == 0 {
                            open_slot = None;
                        }
                    } else if in_scheme && drawingml && local == b"clrScheme" {
                        in_scheme = false;
                    } else if drawingml && (local == b"majorFont" || local == b"minorFont") {
                        font_kind = None;
                    }
                },
                Event::Eof => break,
                Event::Text(_)
                | Event::CData(_)
                | Event::Comment(_)
                | Event::Decl(_)
                | Event::PI(_)
                | Event::DocType(_)
                | Event::GeneralRef(_) => {},
            }
            buffer.clear();
        }

        if in_scheme {
            return Err(invalid("a:clrScheme is unterminated"));
        }
        if !seen_scheme {
            return Err(invalid("theme is missing its a:clrScheme"));
        }
        let colors: [ThemeColorValue; 12] = std::array::from_fn(|index| {
            let slot = ThemeColorSlot::ALL[index];
            colors[index].take().unwrap_or_else(|| {
                tolerance.record(ThemeDefect::MissingSlot, slot_location(slot));
                slot.default_value()
            })
        });
        Ok(Self {
            name,
            color_scheme_name,
            colors,
            major_font,
            minor_font,
            format_scheme_xml,
            tolerance,
        })
    }
}

fn slot_location(slot: ThemeColorSlot) -> Location {
    Location::ColorSchemeSlot(u8::try_from(slot.index()).unwrap_or(u8::MAX))
}

/// Classify one `a:clrScheme` child: the slot it names (recording a slot
/// that follows a later schema slot, or a repeat, as misordered), or `None`
/// for a repeat or an element that names no slot (such as `a:extLst`).
fn color_scheme_slot(
    drawingml: bool,
    local: &[u8],
    highest_slot: &mut Option<usize>,
    colors: &[Option<ThemeColorValue>; 12],
    tolerance: &mut ToleranceReport,
) -> Option<ThemeColorSlot> {
    let slot = drawingml
        .then(|| ThemeColorSlot::from_element_name(local))
        .flatten()?;
    let repeated = colors[slot.index()].is_some()
        || highest_slot.is_some_and(|highest| highest == slot.index());
    if repeated || highest_slot.is_some_and(|highest| slot.index() < highest) {
        tolerance.record(ThemeDefect::MisorderedSlot, slot_location(slot));
    }
    *highest_slot = Some(highest_slot.map_or(slot.index(), |highest| highest.max(slot.index())));
    // A repeated slot: the first occurrence wins.
    (!repeated).then_some(slot)
}

/// Read one slot's color element. A malformed color takes the slot default
/// (recorded); an `a:sysClr` without `lastClr` resolves by name (recorded);
/// any other color element leaves the slot to be defaulted.
fn read_slot_color(
    drawingml: bool,
    local: &[u8],
    element: &quick_xml::events::BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    slot: ThemeColorSlot,
    colors: &mut [Option<ThemeColorValue>; 12],
    tolerance: &mut ToleranceReport,
) -> SheetResult<()> {
    if !drawingml || colors[slot.index()].is_some() {
        return Ok(());
    }
    let value = if local == b"srgbClr" {
        match unqualified_attribute_value(element, b"val", decoder)?
            .as_deref()
            .map(|value| parse_hex_rgb(value, "a:srgbClr val"))
        {
            Some(Ok(rgb)) => ThemeColorValue::Srgb(rgb),
            Some(Err(_)) | None => {
                tolerance.record(ThemeDefect::MalformedColor, slot_location(slot));
                slot.default_value()
            },
        }
    } else if local == b"sysClr" {
        let Some(name) = unqualified_attribute_value(element, b"val", decoder)? else {
            tolerance.record(ThemeDefect::MalformedColor, slot_location(slot));
            colors[slot.index()] = Some(slot.default_value());
            return Ok(());
        };
        let last_rgb = match unqualified_attribute_value(element, b"lastClr", decoder)? {
            Some(last) => match parse_hex_rgb(&last, "a:sysClr lastClr") {
                Ok(rgb) => rgb,
                Err(_) => {
                    tolerance.record(ThemeDefect::MalformedColor, slot_location(slot));
                    system_rgb(&name)
                },
            },
            None => {
                tolerance.record(ThemeDefect::MissingLastColor, slot_location(slot));
                system_rgb(&name)
            },
        };
        ThemeColorValue::System { name, last_rgb }
    } else {
        return Ok(());
    };
    colors[slot.index()] = Some(value);
    Ok(())
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

    fn repairs(theme: &Theme) -> Vec<(ThemeDefect, Location)> {
        theme
            .tolerance_report()
            .iter()
            .map(|entry| match entry.defect {
                crate::leniency::Defect::Theme(defect) => (defect, entry.location.clone()),
                other => panic!("unexpected defect {other:?}"),
            })
            .collect()
    }

    fn at(slot: ThemeColorSlot) -> Location {
        slot_location(slot)
    }

    #[test]
    fn a_conformant_scheme_repairs_nothing() {
        let theme = Theme::parse(THEME).unwrap();
        assert!(theme.tolerance_report().is_clean());
        // An extension list after the slots and color transforms inside a
        // slot are schema-valid and read without repair.
        let extended = THEME
            .replace(
                "</a:clrScheme>",
                "<a:extLst><a:ext uri=\"urn:x\"/></a:extLst></a:clrScheme>",
            )
            .replace(
                r#"<a:srgbClr val="1F497D"/>"#,
                r#"<a:srgbClr val="1F497D"><a:lumMod val="75000"/></a:srgbClr>"#,
            );
        let theme = Theme::parse(&extended).unwrap();
        assert!(theme.tolerance_report().is_clean());
        assert_eq!(theme.rgb(ThemeColorSlot::Dk2), [0x1F, 0x49, 0x7D]);
    }

    /// Audit finding 4.11: slots match by name, missing or malformed slots
    /// take the Office default, a `sysClr` without `lastClr` resolves by
    /// name; each repair is recorded at the slot.
    #[test]
    fn repairs_and_records_color_scheme_deviations() {
        // Slots out of order are matched by name.
        let swapped = THEME.replace(
            r#"<a:dk2><a:srgbClr val="1F497D"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2>"#,
            r#"<a:lt2><a:srgbClr val="EEECE1"/></a:lt2><a:dk2><a:srgbClr val="1F497D"/></a:dk2>"#,
        );
        let theme = Theme::parse(&swapped).unwrap();
        assert_eq!(theme.rgb(ThemeColorSlot::Dk2), [0x1F, 0x49, 0x7D]);
        assert_eq!(theme.rgb(ThemeColorSlot::Lt2), [0xEE, 0xEC, 0xE1]);
        assert_eq!(
            repairs(&theme),
            [(ThemeDefect::MisorderedSlot, at(ThemeColorSlot::Dk2))]
        );

        // A missing slot takes its default.
        let missing = THEME.replace(r#"<a:folHlink><a:srgbClr val="800080"/></a:folHlink>"#, "");
        let theme = Theme::parse(&missing).unwrap();
        assert_eq!(
            theme.color(ThemeColorSlot::FollowedHyperlink),
            &ThemeColorSlot::FollowedHyperlink.default_value()
        );
        assert_eq!(
            repairs(&theme),
            [(
                ThemeDefect::MissingSlot,
                at(ThemeColorSlot::FollowedHyperlink)
            )]
        );

        // Malformed hex takes the slot default.
        let malformed = THEME.replace(
            r#"<a:srgbClr val="9BBB59"/>"#,
            r#"<a:srgbClr val="9BBB5Z"/>"#,
        );
        let theme = Theme::parse(&malformed).unwrap();
        assert_eq!(theme.rgb(ThemeColorSlot::Accent3), [0x9B, 0xBB, 0x59]);
        assert_eq!(
            repairs(&theme),
            [(ThemeDefect::MalformedColor, at(ThemeColorSlot::Accent3))]
        );

        // `lastClr` is optional (CT_SystemColor): resolve by name.
        let no_last = THEME.replace(
            r#"<a:sysClr val="window" lastClr="FFFFFF"/>"#,
            r#"<a:sysClr val="window"/>"#,
        );
        let theme = Theme::parse(&no_last).unwrap();
        assert_eq!(theme.rgb(ThemeColorSlot::Lt1), [0xFF, 0xFF, 0xFF]);
        assert_eq!(
            repairs(&theme),
            [(ThemeDefect::MissingLastColor, at(ThemeColorSlot::Lt1))]
        );

        // A color form this reader does not model leaves the slot defaulted.
        let preset = THEME.replace(r#"<a:srgbClr val="0000FF"/>"#, r#"<a:prstClr val="blue"/>"#);
        let theme = Theme::parse(&preset).unwrap();
        assert_eq!(theme.rgb(ThemeColorSlot::Hyperlink), [0, 0, 0xFF]);
        assert_eq!(
            repairs(&theme),
            [(ThemeDefect::MissingSlot, at(ThemeColorSlot::Hyperlink))]
        );
    }

    #[test]
    fn structural_theme_defects_stay_fatal() {
        // Ill-formed XML.
        assert!(Theme::parse(&THEME.replace("</a:clrScheme>", "</a:clrSchemeX>")).is_err());
        // No color scheme at all.
        let start = THEME.find("<a:clrScheme").unwrap();
        let end = THEME.find("</a:clrScheme>").unwrap() + "</a:clrScheme>".len();
        let without = format!("{}{}", &THEME[..start], &THEME[end..]);
        assert!(Theme::parse(&without).is_err());
    }
}
