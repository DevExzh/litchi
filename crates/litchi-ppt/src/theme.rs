//! The `PowerPoint` 2007 theme a legacy presentation carries (record 0323).
//!
//! `PowerPoint` 2007 and later write each main master's DrawingML theme as an
//! embedded OPC package in a `RoundTripThemeAtom` ([MS-PPT] 2.3.?): the same
//! `a:theme` part a PPTX file holds. This module reads its color scheme and
//! font scheme; the package bytes stay exactly as stored.

use litchi_ole_common::xml_attributes::BytesStartExt as _;
use litchi_opc::OpcPackage;
use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};

use crate::package::{Error, Result};

/// The twelve DrawingML color-scheme slots, in `a:clrScheme` order.
pub const COLOR_SLOTS: [&str; 12] = [
    "dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
    "hlink", "folHlink",
];

/// A theme's typefaces for one script family.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Fonts {
    pub latin: Option<String>,
    pub east_asian: Option<String>,
    pub complex: Option<String>,
}

/// A presentation theme: its color scheme and major and minor fonts.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Theme {
    /// The color scheme's name (`a:clrScheme/@name`).
    pub name: Option<String>,
    /// Each slot of [`COLOR_SLOTS`] as `RRGGBB`; a system color reads as its
    /// last computed color.
    pub colors: [Option<[u8; 3]>; 12],
    /// The heading fonts (`a:majorFont`).
    pub major: Fonts,
    /// The body fonts (`a:minorFont`).
    pub minor: Fonts,
}

impl Theme {
    /// The color of one slot of [`COLOR_SLOTS`].
    #[must_use]
    pub fn color(&self, slot: &str) -> Option<[u8; 3]> {
        COLOR_SLOTS
            .iter()
            .position(|name| *name == slot)
            .and_then(|index| self.colors[index])
    }
}

fn attribute(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    name: &[u8],
) -> Result<Option<String>> {
    for attribute in element.checked_attributes() {
        let attribute =
            attribute.map_err(|error| Error::Corrupted(format!("theme XML attribute: {error}")))?;
        if attribute.key.local_name().as_ref() == name {
            let value = attribute
                .decoded_and_normalized_value(quick_xml::XmlVersion::Explicit1_0, decoder)
                .map_err(|error| Error::Corrupted(format!("theme XML attribute: {error}")))?;
            return Ok(Some(value.into_owned()));
        }
    }
    Ok(None)
}

fn rgb(value: &str) -> Option<[u8; 3]> {
    if value.len() != 6 {
        return None;
    }
    let byte = |index: usize| u8::from_str_radix(value.get(index..index + 2)?, 16).ok();
    Some([byte(0)?, byte(2)?, byte(4)?])
}

/// Parse a DrawingML `a:theme` part.
///
/// # Errors
///
/// Returns [`Error::Corrupted`] for malformed XML.
pub fn parse(xml: &[u8]) -> Result<Theme> {
    let mut reader = Reader::from_reader(xml);
    let mut theme = Theme::default();
    let mut stack: Vec<Vec<u8>> = Vec::new();
    let mut buffer = Vec::new();
    loop {
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(|error| Error::Corrupted(format!("theme XML: {error}")))?;
        let decoder = reader.decoder();
        let (element, empty) = match &event {
            Event::Start(element) => (Some(element.clone()), false),
            Event::Empty(element) => (Some(element.clone()), true),
            Event::End(_) => {
                stack.pop();
                (None, false)
            },
            Event::Eof => break,
            _ => (None, false),
        };
        if let Some(element) = element {
            let local = element.local_name().as_ref().to_vec();
            let parent = stack.last().map(Vec::as_slice);
            let in_colors = stack.iter().any(|name| name == b"clrScheme");
            match local.as_slice() {
                b"clrScheme" => theme.name = attribute(&element, decoder, b"name")?,
                b"srgbClr" | b"sysClr" if in_colors => {
                    let value = if local == b"srgbClr" {
                        attribute(&element, decoder, b"val")?
                    } else {
                        attribute(&element, decoder, b"lastClr")?
                    };
                    if let Some(slot) = parent.and_then(|parent| {
                        COLOR_SLOTS
                            .iter()
                            .position(|name| name.as_bytes() == parent)
                    }) {
                        theme.colors[slot] = value.as_deref().and_then(rgb);
                    }
                },
                b"latin" | b"ea" | b"cs" => {
                    let fonts = if stack.iter().any(|name| name == b"majorFont") {
                        Some(&mut theme.major)
                    } else if stack.iter().any(|name| name == b"minorFont") {
                        Some(&mut theme.minor)
                    } else {
                        None
                    };
                    if let Some(fonts) = fonts {
                        let typeface = attribute(&element, decoder, b"typeface")?
                            .filter(|face| !face.is_empty());
                        match local.as_slice() {
                            b"latin" => fonts.latin = typeface,
                            b"ea" => fonts.east_asian = typeface,
                            _ => fonts.complex = typeface,
                        }
                    }
                },
                _ => {},
            }
            if !empty {
                stack.push(local);
            }
        }
        buffer.clear();
    }
    Ok(theme)
}

/// Read the theme of a `RoundTripThemeAtom` package.
///
/// # Errors
///
/// Returns [`Error::Corrupted`] when the package or its theme part is
/// malformed.
pub fn from_package(package: &crate::slide_round_trip::ThemePackage) -> Result<Theme> {
    let opc = OpcPackage::from_bytes(&package.data, litchi_opc::SourceLimits::default())
        .map_err(|error| Error::Corrupted(format!("theme package: {error}")))?;
    for part in opc.try_iter_parts() {
        let part =
            part.map_err(|error| Error::Corrupted(format!("theme package part: {error}")))?;
        if part.partname().to_string() == package.theme_part_name {
            return parse(part.blob());
        }
    }
    Err(Error::Corrupted("theme package has no theme part".into()))
}
