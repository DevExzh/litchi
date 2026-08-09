//! Package-free `PresentationML` background XML codec.
//!
//! This layer serializes and parses the semantic background model without
//! resolving package relationships or loading image parts.

use super::model::{GradientStop, GradientType, PatternType, PictureStyle, SlideBackground};
use crate::error::{Error, Result};

impl SlideBackground {
    /// Parse a package-free `<p:bg>` fragment from slide XML.
    ///
    /// The codec reads solid, gradient, and pattern fills. Picture fills are
    /// intentionally returned as `None`: their relationship and image bytes
    /// belong to the package-aware host adapter.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn from_xml(xml: &[u8]) -> Result<Option<Self>> {
        use quick_xml::Reader;
        use quick_xml::events::Event;

        let mut reader = Reader::from_reader(xml);
        reader.config_mut().trim_text(true);

        let mut in_bg = false;
        let mut background: Option<SlideBackground> = None;

        loop {
            match reader.read_event() {
                Ok(Event::Start(ref e) | Event::Empty(ref e)) => {
                    let tag_name = e.local_name();

                    if tag_name.as_ref() == b"bg" {
                        in_bg = true;
                    }

                    if in_bg {
                        // Parse solid fill
                        if tag_name.as_ref() == b"solidFill" {
                            // Look for color
                            if let Some(color) = Self::parse_color(&mut reader)? {
                                background = Some(SlideBackground::Solid { color });
                            }
                        }
                        // Parse gradient fill
                        else if tag_name.as_ref() == b"gradFill" {
                            if let Some(gradient) = Self::parse_gradient(&mut reader)? {
                                background = Some(gradient);
                            }
                        }
                        // Parse picture fill (blipFill)
                        else if tag_name.as_ref() == b"blipFill" {
                            if let Some(picture) = Self::parse_picture(&mut reader)? {
                                background = Some(picture);
                            }
                        }
                        // Parse pattern fill
                        else if tag_name.as_ref() == b"pattFill" {
                            let pattern_type = e
                                .attributes()
                                .flatten()
                                .find_map(|attr| {
                                    (attr.key.as_ref() == b"prst")
                                        .then(|| {
                                            std::str::from_utf8(&attr.value)
                                                .ok()
                                                .map(Self::pattern_from_xml)
                                        })
                                        .flatten()
                                })
                                .unwrap_or(PatternType::Pct50);
                            if let Some(pattern) = Self::parse_pattern(&mut reader, pattern_type)? {
                                background = Some(pattern);
                            }
                        }
                    }
                },
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"bg" => {
                    in_bg = false;
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(Error::Xml(e.to_string())),
                _ => {},
            }
        }

        Ok(background)
    }

    /// Parse color from XML.
    fn parse_color(reader: &mut quick_xml::Reader<&[u8]>) -> Result<Option<String>> {
        use quick_xml::events::Event;

        loop {
            match reader.read_event() {
                Ok(Event::Empty(ref e) | Event::Start(ref e)) => {
                    let tag_name = e.local_name();
                    if tag_name.as_ref() == b"srgbClr" {
                        // RGB color
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"val" {
                                return Ok(Some(
                                    std::str::from_utf8(&attr.value).unwrap_or("").to_string(),
                                ));
                            }
                        }
                    } else if tag_name.as_ref() == b"schemeClr" {
                        // Scheme color - return the scheme name
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"val" {
                                return Ok(Some(
                                    std::str::from_utf8(&attr.value).unwrap_or("").to_string(),
                                ));
                            }
                        }
                    }
                },
                Ok(Event::End(_)) => break,
                Ok(Event::Eof) => break,
                _ => {},
            }
        }
        Ok(None)
    }

    /// Parse gradient fill from XML.
    fn parse_gradient(reader: &mut quick_xml::Reader<&[u8]>) -> Result<Option<SlideBackground>> {
        use quick_xml::events::Event;

        let mut gradient_type = GradientType::Linear;
        let mut angle: Option<f64> = None;
        let mut stops = Vec::new();

        loop {
            match reader.read_event() {
                Ok(Event::Start(ref e)) => {
                    let tag_name = e.local_name();

                    // Parse gradient stops
                    if tag_name.as_ref() == b"gs" {
                        let mut position = 0.0;
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"pos" {
                                // Position is in percentage * 1000
                                if let Ok(pos_str) = std::str::from_utf8(&attr.value)
                                    && let Ok(pos_val) = pos_str.parse::<f64>()
                                {
                                    position = pos_val / 100_000.0;
                                }
                            }
                        }

                        // Parse color for this stop
                        if let Some(color) = Self::parse_color(reader)? {
                            stops.push(GradientStop { position, color });
                        }
                    } else if matches!(tag_name.as_ref(), b"lin" | b"path") {
                        Self::parse_gradient_descriptor(e, &mut gradient_type, &mut angle);
                    }
                },
                Ok(Event::Empty(ref e)) => {
                    if matches!(e.local_name().as_ref(), b"lin" | b"path") {
                        Self::parse_gradient_descriptor(e, &mut gradient_type, &mut angle);
                    }
                },
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"gradFill" => break,
                Ok(Event::Eof) => break,
                Err(e) => return Err(Error::Xml(e.to_string())),
                _ => {},
            }
        }

        if stops.is_empty() {
            Ok(None)
        } else {
            Ok(Some(SlideBackground::Gradient {
                gradient_type,
                angle,
                stops,
            }))
        }
    }

    fn parse_gradient_descriptor(
        element: &quick_xml::events::BytesStart<'_>,
        gradient_type: &mut GradientType,
        angle: &mut Option<f64>,
    ) {
        match element.local_name().as_ref() {
            b"lin" => {
                *gradient_type = GradientType::Linear;
                for attr in element.attributes().flatten() {
                    if attr.key.as_ref() == b"ang"
                        && let Ok(value) = std::str::from_utf8(&attr.value)
                        && let Ok(value) = value.parse::<f64>()
                    {
                        // DrawingML fixed angles are measured in 60,000ths
                        // of a degree.
                        *angle = Some(value / 60000.0);
                    }
                }
            },
            b"path" => {
                for attr in element.attributes().flatten() {
                    if attr.key.as_ref() == b"path"
                        && let Ok(value) = std::str::from_utf8(&attr.value)
                    {
                        *gradient_type = match value {
                            "circle" => GradientType::Radial,
                            "rect" => GradientType::Rectangular,
                            "shape" => GradientType::Path,
                            _ => GradientType::Linear,
                        };
                    }
                }
            },
            _ => {},
        }
    }

    /// Parse picture fill from XML.
    fn parse_picture(_reader: &mut quick_xml::Reader<&[u8]>) -> Result<Option<SlideBackground>> {
        // Note: Picture fills in backgrounds require accessing the image data
        // from the package relationships, which is not available in this parsing context.
        // The image data and relationship ID should be handled at a higher level
        // when the full package context is available.
        // For now, we return None to indicate we can't fully parse picture backgrounds here.
        Ok(None)
    }

    /// Parse pattern fill from XML.
    fn parse_pattern(
        reader: &mut quick_xml::Reader<&[u8]>,
        pattern_type: PatternType,
    ) -> Result<Option<SlideBackground>> {
        use quick_xml::events::Event;

        let mut fg_color = String::from("000000");
        let mut bg_color = String::from("FFFFFF");

        loop {
            match reader.read_event() {
                Ok(Event::Start(ref e)) => {
                    let tag_name = e.local_name();

                    if tag_name.as_ref() == b"fgClr" {
                        if let Some(color) = Self::parse_color(reader)? {
                            fg_color = color;
                        }
                    } else if tag_name.as_ref() == b"bgClr"
                        && let Some(color) = Self::parse_color(reader)?
                    {
                        bg_color = color;
                    }
                },
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"pattFill" => break,
                Ok(Event::Eof) => break,
                Err(e) => return Err(Error::Xml(e.to_string())),
                _ => {},
            }
        }

        Ok(Some(SlideBackground::Pattern {
            pattern_type,
            fg_color,
            bg_color,
        }))
    }

    fn pattern_from_xml(value: &str) -> PatternType {
        match value {
            "pct5" => PatternType::Pct5,
            "pct10" => PatternType::Pct10,
            "pct20" => PatternType::Pct20,
            "pct25" => PatternType::Pct25,
            "pct30" => PatternType::Pct30,
            "pct40" => PatternType::Pct40,
            "pct50" => PatternType::Pct50,
            "pct60" => PatternType::Pct60,
            "pct70" => PatternType::Pct70,
            "pct75" => PatternType::Pct75,
            "pct80" => PatternType::Pct80,
            "pct90" => PatternType::Pct90,
            "horz" => PatternType::Horizontal,
            "vert" => PatternType::Vertical,
            "ltHorz" => PatternType::LightHorizontal,
            "ltVert" => PatternType::LightVertical,
            "dkHorz" => PatternType::DarkHorizontal,
            "dkVert" => PatternType::DarkVertical,
            "narHorz" => PatternType::NarrowHorizontal,
            "narVert" => PatternType::NarrowVertical,
            "dashHorz" => PatternType::DashedHorizontal,
            "dashVert" => PatternType::DashedVertical,
            "dnDiag" => PatternType::DownDiagonal,
            "upDiag" => PatternType::UpDiagonal,
            "ltDnDiag" => PatternType::LightDownDiagonal,
            "ltUpDiag" => PatternType::LightUpDiagonal,
            "dkDnDiag" => PatternType::DarkDownDiagonal,
            "dkUpDiag" => PatternType::DarkUpDiagonal,
            "wdDnDiag" => PatternType::WideDownDiagonal,
            "wdUpDiag" => PatternType::WideUpDiagonal,
            "dashDnDiag" => PatternType::DashedDownDiagonal,
            "dashUpDiag" => PatternType::DashedUpDiagonal,
            "cross" => PatternType::Cross,
            "diagCross" => PatternType::DiagonalCross,
            "smCheck" => PatternType::SmallCheck,
            "lgCheck" => PatternType::LargeCheck,
            "smGrid" => PatternType::SmallGrid,
            "lgGrid" => PatternType::LargeGrid,
            "dotGrid" => PatternType::DottedGrid,
            "smConfetti" => PatternType::SmallConfetti,
            "lgConfetti" => PatternType::LargeConfetti,
            "horzBrick" => PatternType::HorizontalBrick,
            "diagBrick" => PatternType::DiagonalBrick,
            "solidDmnd" => PatternType::SolidDiamond,
            "openDmnd" => PatternType::OpenDiamond,
            "dotDmnd" => PatternType::DottedDiamond,
            "plaid" => PatternType::Plaid,
            "sphere" => PatternType::Sphere,
            "weave" => PatternType::Weave,
            "divot" => PatternType::Divot,
            "shingle" => PatternType::Shingle,
            "wave" => PatternType::Wave,
            "trellis" => PatternType::Trellis,
            "zigZag" => PatternType::ZigZag,
            _ => PatternType::Pct50,
        }
    }

    /// Encode this background as a package-free `<p:bg>` fragment.
    ///
    /// `rel_id` is supplied by the package host after it resolves a picture
    /// relationship. The owner never interprets or allocates relationship
    /// identifiers.
    ///
    /// # Errors
    ///
    /// Returns an error if the output cannot be encoded or written.
    pub fn to_xml(&self, rel_id: Option<&str>) -> Result<String> {
        let mut xml = String::with_capacity(512);

        match self {
            SlideBackground::None => {
                // No background element
            },
            SlideBackground::Solid { color } => {
                xml.push_str("<p:bg>");
                xml.push_str("<p:bgPr>");
                xml.push_str("<a:solidFill>");
                xml.push_str("<a:srgbClr val=\"");
                xml.push_str(color);
                xml.push_str("\"/>");
                xml.push_str("</a:solidFill>");
                xml.push_str("<a:effectLst/>");
                xml.push_str("</p:bgPr>");
                xml.push_str("</p:bg>");
            },
            SlideBackground::Gradient {
                gradient_type,
                angle,
                stops,
            } => {
                xml.push_str("<p:bg>");
                xml.push_str("<p:bgPr>");
                xml.push_str("<a:gradFill");

                match gradient_type {
                    GradientType::Linear => {
                        xml.push_str(" rotWithShape=\"1\">");
                        xml.push_str("<a:gsLst>");
                        for stop in stops {
                            let pos = (stop.position * 100_000.0) as i32;
                            xml.push_str(&format!("<a:gs pos=\"{pos}\">"));
                            xml.push_str("<a:srgbClr val=\"");
                            xml.push_str(&stop.color);
                            xml.push_str("\"/>");
                            xml.push_str("</a:gs>");
                        }
                        xml.push_str("</a:gsLst>");

                        if let Some(ang) = angle {
                            let ang_val = (ang * 60000.0) as i32;
                            xml.push_str(&format!("<a:lin ang=\"{ang_val}\" scaled=\"0\"/>"));
                        }
                    },
                    GradientType::Radial => {
                        xml.push('>');
                        xml.push_str("<a:gsLst>");
                        for stop in stops {
                            let pos = (stop.position * 100_000.0) as i32;
                            xml.push_str(&format!("<a:gs pos=\"{pos}\">"));
                            xml.push_str("<a:srgbClr val=\"");
                            xml.push_str(&stop.color);
                            xml.push_str("\"/>");
                            xml.push_str("</a:gs>");
                        }
                        xml.push_str("</a:gsLst>");
                        xml.push_str("<a:path path=\"circle\"><a:fillToRect l=\"50000\" t=\"50000\" r=\"50000\" b=\"50000\"/></a:path>");
                    },
                    _ => {
                        xml.push('>');
                        xml.push_str("<a:gsLst>");
                        for stop in stops {
                            let pos = (stop.position * 100_000.0) as i32;
                            xml.push_str(&format!("<a:gs pos=\"{pos}\">"));
                            xml.push_str("<a:srgbClr val=\"");
                            xml.push_str(&stop.color);
                            xml.push_str("\"/>");
                            xml.push_str("</a:gs>");
                        }
                        xml.push_str("</a:gsLst>");
                    },
                }

                xml.push_str("</a:gradFill>");
                xml.push_str("<a:effectLst/>");
                xml.push_str("</p:bgPr>");
                xml.push_str("</p:bg>");
            },
            SlideBackground::Picture { style, .. } => {
                xml.push_str("<p:bg>");
                xml.push_str("<p:bgPr>");
                xml.push_str("<a:blipFill");

                match style {
                    PictureStyle::Stretch => xml.push_str(" dpi=\"0\" rotWithShape=\"1\">"),
                    PictureStyle::Tile => xml.push_str(" dpi=\"0\" rotWithShape=\"1\">"),
                    _ => xml.push('>'),
                }

                if let Some(rid) = rel_id {
                    xml.push_str("<a:blip r:embed=\"");
                    xml.push_str(rid);
                    xml.push_str("\"/>");
                } else {
                    xml.push_str("<a:blip/>");
                }

                match style {
                    PictureStyle::Stretch => {
                        xml.push_str("<a:stretch><a:fillRect/></a:stretch>");
                    },
                    PictureStyle::Tile => {
                        xml.push_str("<a:tile/>");
                    },
                    _ => {
                        xml.push_str("<a:stretch><a:fillRect/></a:stretch>");
                    },
                }

                xml.push_str("</a:blipFill>");
                xml.push_str("<a:effectLst/>");
                xml.push_str("</p:bgPr>");
                xml.push_str("</p:bg>");
            },
            SlideBackground::Pattern {
                pattern_type,
                fg_color,
                bg_color,
            } => {
                xml.push_str("<p:bg>");
                xml.push_str("<p:bgPr>");
                xml.push_str("<a:pattFill prst=\"");
                xml.push_str(Self::pattern_to_xml(*pattern_type));
                xml.push_str("\">");
                xml.push_str("<a:fgClr><a:srgbClr val=\"");
                xml.push_str(fg_color);
                xml.push_str("\"/></a:fgClr>");
                xml.push_str("<a:bgClr><a:srgbClr val=\"");
                xml.push_str(bg_color);
                xml.push_str("\"/></a:bgClr>");
                xml.push_str("</a:pattFill>");
                xml.push_str("<a:effectLst/>");
                xml.push_str("</p:bgPr>");
                xml.push_str("</p:bg>");
            },
        }

        Ok(xml)
    }

    /// Convert pattern type to XML value.
    fn pattern_to_xml(pattern: PatternType) -> &'static str {
        match pattern {
            PatternType::Pct5 => "pct5",
            PatternType::Pct10 => "pct10",
            PatternType::Pct20 => "pct20",
            PatternType::Pct25 => "pct25",
            PatternType::Pct30 => "pct30",
            PatternType::Pct40 => "pct40",
            PatternType::Pct50 => "pct50",
            PatternType::Pct60 => "pct60",
            PatternType::Pct70 => "pct70",
            PatternType::Pct75 => "pct75",
            PatternType::Pct80 => "pct80",
            PatternType::Pct90 => "pct90",
            PatternType::Horizontal => "horz",
            PatternType::Vertical => "vert",
            PatternType::LightHorizontal => "ltHorz",
            PatternType::LightVertical => "ltVert",
            PatternType::DarkHorizontal => "dkHorz",
            PatternType::DarkVertical => "dkVert",
            PatternType::NarrowHorizontal => "narHorz",
            PatternType::NarrowVertical => "narVert",
            PatternType::DashedHorizontal => "dashHorz",
            PatternType::DashedVertical => "dashVert",
            PatternType::DownDiagonal => "dnDiag",
            PatternType::UpDiagonal => "upDiag",
            PatternType::LightDownDiagonal => "ltDnDiag",
            PatternType::LightUpDiagonal => "ltUpDiag",
            PatternType::DarkDownDiagonal => "dkDnDiag",
            PatternType::DarkUpDiagonal => "dkUpDiag",
            PatternType::WideDownDiagonal => "wdDnDiag",
            PatternType::WideUpDiagonal => "wdUpDiag",
            PatternType::DashedDownDiagonal => "dashDnDiag",
            PatternType::DashedUpDiagonal => "dashUpDiag",
            PatternType::Cross => "cross",
            PatternType::DiagonalCross => "diagCross",
            PatternType::SmallCheck => "smCheck",
            PatternType::LargeCheck => "lgCheck",
            PatternType::SmallGrid => "smGrid",
            PatternType::LargeGrid => "lgGrid",
            PatternType::DottedGrid => "dotGrid",
            PatternType::SmallConfetti => "smConfetti",
            PatternType::LargeConfetti => "lgConfetti",
            PatternType::HorizontalBrick => "horzBrick",
            PatternType::DiagonalBrick => "diagBrick",
            PatternType::SolidDiamond => "solidDmnd",
            PatternType::OpenDiamond => "openDmnd",
            PatternType::DottedDiamond => "dotDmnd",
            PatternType::Plaid => "plaid",
            PatternType::Sphere => "sphere",
            PatternType::Weave => "weave",
            PatternType::Divot => "divot",
            PatternType::Shingle => "shingle",
            PatternType::Wave => "wave",
            PatternType::Trellis => "trellis",
            PatternType::ZigZag => "zigZag",
        }
    }
}
