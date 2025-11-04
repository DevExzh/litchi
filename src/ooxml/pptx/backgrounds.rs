//! Slide background support for PowerPoint presentations.
//!
//! This module provides types for working with slide backgrounds, including
//! solid colors, gradients, images, and patterns.

use crate::ooxml::error::{OoxmlError, Result};

/// Slide background configuration.
///
/// Represents the background fill for a slide.
#[derive(Debug, Clone, PartialEq)]
pub enum SlideBackground {
    /// No background (transparent)
    None,
    /// Solid color background
    Solid {
        /// RGB color in hexadecimal format (e.g., "FFFFFF" for white)
        color: String,
    },
    /// Gradient background
    Gradient {
        /// Gradient type (linear or radial)
        gradient_type: GradientType,
        /// Gradient angle in degrees (0-360, for linear gradients)
        angle: Option<f64>,
        /// Gradient stops (position 0.0-1.0, color pairs)
        stops: Vec<GradientStop>,
    },
    /// Picture/image background
    Picture {
        /// Image data (will be embedded)
        image_data: Vec<u8>,
        /// Image format
        format: super::format::ImageFormat,
        /// How the image should be displayed
        style: PictureStyle,
    },
    /// Pattern background
    Pattern {
        /// Pattern type
        pattern_type: PatternType,
        /// Foreground color
        fg_color: String,
        /// Background color
        bg_color: String,
    },
}

/// Gradient type for gradient backgrounds.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum GradientType {
    /// Linear gradient
    Linear,
    /// Radial gradient (from center)
    Radial,
    /// Rectangular gradient
    Rectangular,
    /// Path gradient
    Path,
}

/// A gradient stop (position and color).
#[derive(Debug, Clone, PartialEq)]
pub struct GradientStop {
    /// Position from 0.0 to 1.0
    pub position: f64,
    /// RGB color in hex format
    pub color: String,
}

/// Picture display style for picture backgrounds.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PictureStyle {
    /// Stretch to fill the slide
    Stretch,
    /// Fit within the slide (maintain aspect ratio)
    Fit,
    /// Tile the image
    Tile,
    /// Center the image
    Center,
}

/// Pattern type for pattern backgrounds.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PatternType {
    /// 5% dotted pattern
    Pct5,
    /// 10% dotted pattern
    Pct10,
    /// 20% dotted pattern
    Pct20,
    /// 25% dotted pattern
    Pct25,
    /// 30% dotted pattern
    Pct30,
    /// 40% dotted pattern
    Pct40,
    /// 50% dotted pattern
    Pct50,
    /// 60% dotted pattern
    Pct60,
    /// 70% dotted pattern
    Pct70,
    /// 75% dotted pattern
    Pct75,
    /// 80% dotted pattern
    Pct80,
    /// 90% dotted pattern
    Pct90,
    /// Horizontal lines
    Horizontal,
    /// Vertical lines
    Vertical,
    /// Light horizontal lines
    LightHorizontal,
    /// Light vertical lines
    LightVertical,
    /// Dark horizontal lines
    DarkHorizontal,
    /// Dark vertical lines
    DarkVertical,
    /// Narrow horizontal lines
    NarrowHorizontal,
    /// Narrow vertical lines
    NarrowVertical,
    /// Dashed horizontal lines
    DashedHorizontal,
    /// Dashed vertical lines
    DashedVertical,
    /// Diagonal lines (top-left to bottom-right)
    DownDiagonal,
    /// Diagonal lines (bottom-left to top-right)
    UpDiagonal,
    /// Light down diagonal
    LightDownDiagonal,
    /// Light up diagonal
    LightUpDiagonal,
    /// Dark down diagonal
    DarkDownDiagonal,
    /// Dark up diagonal
    DarkUpDiagonal,
    /// Wide down diagonal
    WideDownDiagonal,
    /// Wide up diagonal
    WideUpDiagonal,
    /// Dashed down diagonal
    DashedDownDiagonal,
    /// Dashed up diagonal
    DashedUpDiagonal,
    /// Crosshatch
    Cross,
    /// Diagonal crosshatch
    DiagonalCross,
    /// Small check pattern
    SmallCheck,
    /// Large check pattern
    LargeCheck,
    /// Small grid
    SmallGrid,
    /// Large grid
    LargeGrid,
    /// Dotted grid
    DottedGrid,
    /// Small confetti
    SmallConfetti,
    /// Large confetti
    LargeConfetti,
    /// Horizontal brick
    HorizontalBrick,
    /// Diagonal brick
    DiagonalBrick,
    /// Solid diamond
    SolidDiamond,
    /// Open diamond
    OpenDiamond,
    /// Dotted diamond
    DottedDiamond,
    /// Plaid pattern
    Plaid,
    /// Sphere pattern
    Sphere,
    /// Weave pattern
    Weave,
    /// Divot pattern
    Divot,
    /// Shingle pattern
    Shingle,
    /// Wave pattern
    Wave,
    /// Trellis pattern
    Trellis,
    /// Zig-zag pattern
    ZigZag,
}

impl SlideBackground {
    /// Create a solid color background.
    ///
    /// # Arguments
    /// * `color` - RGB color in hex format (e.g., "FF0000" for red)
    ///
    /// # Examples
    ///
    /// ```rust
    /// use litchi::ooxml::pptx::SlideBackground;
    ///
    /// let bg = SlideBackground::solid("4472C4"); // Blue background
    /// ```
    pub fn solid(color: impl Into<String>) -> Self {
        SlideBackground::Solid {
            color: color.into(),
        }
    }

    /// Create a linear gradient background.
    ///
    /// # Arguments
    /// * `angle` - Gradient angle in degrees (0-360)
    /// * `stops` - Vector of gradient stops
    ///
    /// # Examples
    ///
    /// ```rust
    /// use litchi::ooxml::pptx::{SlideBackground, GradientStop};
    ///
    /// let stops = vec![
    ///     GradientStop { position: 0.0, color: "4472C4".to_string() },
    ///     GradientStop { position: 1.0, color: "FFFFFF".to_string() },
    /// ];
    /// let bg = SlideBackground::linear_gradient(90.0, stops);
    /// ```
    pub fn linear_gradient(angle: f64, stops: Vec<GradientStop>) -> Self {
        SlideBackground::Gradient {
            gradient_type: GradientType::Linear,
            angle: Some(angle),
            stops,
        }
    }

    /// Create a radial gradient background.
    ///
    /// # Arguments
    /// * `stops` - Vector of gradient stops
    pub fn radial_gradient(stops: Vec<GradientStop>) -> Self {
        SlideBackground::Gradient {
            gradient_type: GradientType::Radial,
            angle: None,
            stops,
        }
    }

    /// Create a picture background from image data.
    ///
    /// # Arguments
    /// * `image_data` - Image bytes
    /// * `format` - Image format
    /// * `style` - How to display the image
    pub fn picture(
        image_data: Vec<u8>,
        format: super::format::ImageFormat,
        style: PictureStyle,
    ) -> Self {
        SlideBackground::Picture {
            image_data,
            format,
            style,
        }
    }

    /// Create a pattern background.
    ///
    /// # Arguments
    /// * `pattern_type` - Type of pattern
    /// * `fg_color` - Foreground color in hex
    /// * `bg_color` - Background color in hex
    pub fn pattern(pattern_type: PatternType, fg_color: String, bg_color: String) -> Self {
        SlideBackground::Pattern {
            pattern_type,
            fg_color,
            bg_color,
        }
    }

    /// Parse background from slide XML.
    pub(crate) fn from_xml(xml: &[u8]) -> Result<Option<Self>> {
        use quick_xml::Reader;
        use quick_xml::events::Event;

        let mut reader = Reader::from_reader(xml);
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();

        let mut in_bg = false;
        let mut background: Option<SlideBackground> = None;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                    let tag_name = e.local_name();

                    if tag_name.as_ref() == b"bg" {
                        in_bg = true;
                    }

                    if in_bg {
                        // Parse solid fill
                        if tag_name.as_ref() == b"solidFill" {
                            // Look for color
                            if let Some(color) = Self::parse_color(&mut reader, &mut buf)? {
                                background = Some(SlideBackground::Solid { color });
                            }
                        }
                        // Parse gradient fill
                        else if tag_name.as_ref() == b"gradFill" {
                            if let Some(gradient) = Self::parse_gradient(&mut reader, &mut buf)? {
                                background = Some(gradient);
                            }
                        }
                        // Parse picture fill (blipFill)
                        else if tag_name.as_ref() == b"blipFill" {
                            if let Some(picture) = Self::parse_picture(&mut reader, &mut buf)? {
                                background = Some(picture);
                            }
                        }
                        // Parse pattern fill
                        else if tag_name.as_ref() == b"pattFill"
                            && let Some(pattern) = Self::parse_pattern(&mut reader, &mut buf)?
                        {
                            background = Some(pattern);
                        }
                    }
                },
                Ok(Event::End(ref e)) => {
                    if e.local_name().as_ref() == b"bg" {
                        in_bg = false;
                    }
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
            buf.clear();
        }

        Ok(background)
    }

    /// Parse color from XML.
    fn parse_color(
        reader: &mut quick_xml::Reader<&[u8]>,
        buf: &mut Vec<u8>,
    ) -> Result<Option<String>> {
        use quick_xml::events::Event;

        loop {
            buf.clear();
            match reader.read_event_into(buf) {
                Ok(Event::Empty(ref e)) | Ok(Event::Start(ref e)) => {
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
    fn parse_gradient(
        reader: &mut quick_xml::Reader<&[u8]>,
        buf: &mut Vec<u8>,
    ) -> Result<Option<SlideBackground>> {
        use quick_xml::events::Event;

        let mut gradient_type = GradientType::Linear;
        let mut angle: Option<f64> = None;
        let mut stops = Vec::new();
        let mut depth = 0;

        loop {
            buf.clear();
            match reader.read_event_into(buf) {
                Ok(Event::Start(ref e)) => {
                    depth += 1;
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
                                    position = pos_val / 100000.0;
                                }
                            }
                        }

                        // Parse color for this stop
                        if let Some(color) = Self::parse_color(reader, buf)? {
                            stops.push(GradientStop { position, color });
                        }
                    }
                    // Parse linear gradient angle
                    else if tag_name.as_ref() == b"lin" {
                        gradient_type = GradientType::Linear;
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"ang" {
                                // Angle is in 60,000ths of a degree
                                if let Ok(ang_str) = std::str::from_utf8(&attr.value)
                                    && let Ok(ang_val) = ang_str.parse::<f64>()
                                {
                                    angle = Some(ang_val / 60000.0);
                                }
                            }
                        }
                    }
                    // Parse radial/path gradient
                    else if tag_name.as_ref() == b"path" {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"path"
                                && let Ok(path_str) = std::str::from_utf8(&attr.value)
                            {
                                gradient_type = match path_str {
                                    "circle" => GradientType::Radial,
                                    "rect" => GradientType::Rectangular,
                                    "shape" => GradientType::Path,
                                    _ => GradientType::Linear,
                                };
                            }
                        }
                    }
                },
                Ok(Event::End(ref e)) => {
                    if e.local_name().as_ref() == b"gradFill" {
                        depth -= 1;
                        if depth == 0 {
                            break;
                        }
                    }
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
        }

        if !stops.is_empty() {
            Ok(Some(SlideBackground::Gradient {
                gradient_type,
                angle,
                stops,
            }))
        } else {
            Ok(None)
        }
    }

    /// Parse picture fill from XML.
    fn parse_picture(
        _reader: &mut quick_xml::Reader<&[u8]>,
        _buf: &mut Vec<u8>,
    ) -> Result<Option<SlideBackground>> {
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
        buf: &mut Vec<u8>,
    ) -> Result<Option<SlideBackground>> {
        use quick_xml::events::Event;

        let mut pattern_type = PatternType::Pct50;
        let mut fg_color = String::from("000000");
        let mut bg_color = String::from("FFFFFF");
        let mut depth = 0;
        let mut in_fg = false;
        let mut in_bg = false;

        loop {
            buf.clear();
            match reader.read_event_into(buf) {
                Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                    let tag_name = e.local_name();

                    if tag_name.as_ref() == b"pattFill" {
                        depth += 1;
                        // Parse pattern type
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"prst"
                                && let Ok(prst_str) = std::str::from_utf8(&attr.value)
                            {
                                pattern_type = match prst_str {
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
                                    "cross" => PatternType::Cross,
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
                                };
                            }
                        }
                    } else if tag_name.as_ref() == b"fgClr" {
                        in_fg = true;
                    } else if tag_name.as_ref() == b"bgClr" {
                        in_bg = true;
                    } else if in_fg && tag_name.as_ref() == b"srgbClr" {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"val" {
                                fg_color = std::str::from_utf8(&attr.value)
                                    .unwrap_or("000000")
                                    .to_string();
                            }
                        }
                    } else if in_bg && tag_name.as_ref() == b"srgbClr" {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"val" {
                                bg_color = std::str::from_utf8(&attr.value)
                                    .unwrap_or("FFFFFF")
                                    .to_string();
                            }
                        }
                    }
                },
                Ok(Event::End(ref e)) => {
                    let tag_name = e.local_name();
                    if tag_name.as_ref() == b"pattFill" {
                        depth -= 1;
                        if depth == 0 {
                            break;
                        }
                    } else if tag_name.as_ref() == b"fgClr" {
                        in_fg = false;
                    } else if tag_name.as_ref() == b"bgClr" {
                        in_bg = false;
                    }
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
        }

        Ok(Some(SlideBackground::Pattern {
            pattern_type,
            fg_color,
            bg_color,
        }))
    }

    /// Generate XML for this background.
    pub(crate) fn to_xml(&self, rel_id: Option<&str>) -> Result<String> {
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
                            let pos = (stop.position * 100000.0) as i32;
                            xml.push_str(&format!("<a:gs pos=\"{}\">", pos));
                            xml.push_str("<a:srgbClr val=\"");
                            xml.push_str(&stop.color);
                            xml.push_str("\"/>");
                            xml.push_str("</a:gs>");
                        }
                        xml.push_str("</a:gsLst>");

                        if let Some(ang) = angle {
                            let ang_val = (ang * 60000.0) as i32;
                            xml.push_str(&format!("<a:lin ang=\"{}\" scaled=\"0\"/>", ang_val));
                        }
                    },
                    GradientType::Radial => {
                        xml.push('>');
                        xml.push_str("<a:gsLst>");
                        for stop in stops {
                            let pos = (stop.position * 100000.0) as i32;
                            xml.push_str(&format!("<a:gs pos=\"{}\">", pos));
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
                            let pos = (stop.position * 100000.0) as i32;
                            xml.push_str(&format!("<a:gs pos=\"{}\">", pos));
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

    /// Get image data if this is a picture background.
    pub(crate) fn get_image_data(&self) -> Option<(&[u8], &super::format::ImageFormat)> {
        match self {
            SlideBackground::Picture {
                image_data, format, ..
            } => Some((image_data, format)),
            _ => None,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_solid_background() {
        let bg = SlideBackground::solid("FF0000");
        assert!(matches!(bg, SlideBackground::Solid { .. }));
    }

    #[test]
    fn test_gradient_background() {
        let stops = vec![
            GradientStop {
                position: 0.0,
                color: "FF0000".to_string(),
            },
            GradientStop {
                position: 1.0,
                color: "0000FF".to_string(),
            },
        ];
        let bg = SlideBackground::linear_gradient(90.0, stops);
        assert!(matches!(bg, SlideBackground::Gradient { .. }));
    }

    #[test]
    fn test_solid_background_xml() {
        let bg = SlideBackground::solid("FF0000");
        let xml = bg.to_xml(None).unwrap();
        assert!(xml.contains("FF0000"));
        assert!(xml.contains("<a:solidFill>"));
    }
}
