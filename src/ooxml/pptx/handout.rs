//! Handout master support for PowerPoint presentations.
//!
//! Handout masters define the layout for printed handouts that show
//! multiple slides per page.

use crate::common::id::generate_guid_braced;
use crate::ooxml::error::{OoxmlError, Result};
use quick_xml::Reader;
use quick_xml::events::Event;
use std::str::FromStr;

/// Number of slides per handout page.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum HandoutLayout {
    /// 1 slide per page
    #[default]
    OneSlide,
    /// 2 slides per page
    TwoSlides,
    /// 3 slides per page (with lines for notes)
    ThreeSlides,
    /// 4 slides per page
    FourSlides,
    /// 6 slides per page
    SixSlides,
    /// 9 slides per page
    NineSlides,
    /// Outline view
    Outline,
}

impl HandoutLayout {
    /// Get the layout type string for handout master XML.
    pub fn as_str(&self) -> &'static str {
        match self {
            HandoutLayout::OneSlide => "handout1",
            HandoutLayout::TwoSlides => "handout2",
            HandoutLayout::ThreeSlides => "handout3",
            HandoutLayout::FourSlides => "handout4",
            HandoutLayout::SixSlides => "handout6",
            HandoutLayout::NineSlides => "handout9",
            HandoutLayout::Outline => "handoutOutline",
        }
    }

    /// Get the print property value for presProps.xml (prnWhat attribute).
    /// Per OOXML spec ST_PrintWhat: slides, handouts1, handouts2, handouts3, handouts4, handouts6, handouts9, notes, outline
    pub fn print_what(&self) -> &'static str {
        match self {
            HandoutLayout::OneSlide => "handouts1",
            HandoutLayout::TwoSlides => "handouts2",
            HandoutLayout::ThreeSlides => "handouts3",
            HandoutLayout::FourSlides => "handouts4",
            HandoutLayout::SixSlides => "handouts6",
            HandoutLayout::NineSlides => "handouts9",
            HandoutLayout::Outline => "outline",
        }
    }
}

impl FromStr for HandoutLayout {
    type Err = ();

    fn from_str(s: &str) -> std::result::Result<Self, Self::Err> {
        Ok(match s {
            "handout1" => HandoutLayout::OneSlide,
            "handout2" => HandoutLayout::TwoSlides,
            "handout3" => HandoutLayout::ThreeSlides,
            "handout4" => HandoutLayout::FourSlides,
            "handout6" => HandoutLayout::SixSlides,
            "handout9" => HandoutLayout::NineSlides,
            "handoutOutline" => HandoutLayout::Outline,
            _ => HandoutLayout::OneSlide,
        })
    }
}

/// Header/Footer configuration for handouts.
#[derive(Debug, Clone, Default)]
pub struct HandoutHeaderFooter {
    /// Show header text
    pub show_header: bool,
    /// Header text
    pub header_text: Option<String>,
    /// Show footer text
    pub show_footer: bool,
    /// Footer text
    pub footer_text: Option<String>,
    /// Show slide number
    pub show_slide_number: bool,
    /// Show date/time
    pub show_date_time: bool,
    /// Date/time format (if fixed)
    pub date_time_text: Option<String>,
    /// Use automatic date
    pub auto_date: bool,
}

/// Handout master for a presentation.
#[derive(Debug, Clone)]
pub struct HandoutMaster {
    /// Handout layout
    pub layout: HandoutLayout,
    /// Header/footer settings
    pub header_footer: HandoutHeaderFooter,
    /// Background color (hex RGB)
    pub background_color: Option<String>,
    /// Whether to show slide images
    pub show_slide_images: bool,
}

impl Default for HandoutMaster {
    fn default() -> Self {
        Self {
            layout: HandoutLayout::default(),
            header_footer: HandoutHeaderFooter::default(),
            background_color: None,
            show_slide_images: true,
        }
    }
}

impl HandoutMaster {
    /// Create a new handout master with default settings.
    pub fn new() -> Self {
        Self::default()
    }

    /// Set the handout layout.
    pub fn with_layout(mut self, layout: HandoutLayout) -> Self {
        self.layout = layout;
        self
    }

    /// Set the header text.
    pub fn with_header(mut self, text: impl Into<String>) -> Self {
        self.header_footer.show_header = true;
        self.header_footer.header_text = Some(text.into());
        self
    }

    /// Set the footer text.
    pub fn with_footer(mut self, text: impl Into<String>) -> Self {
        self.header_footer.show_footer = true;
        self.header_footer.footer_text = Some(text.into());
        self
    }

    /// Enable slide numbers.
    pub fn with_slide_numbers(mut self) -> Self {
        self.header_footer.show_slide_number = true;
        self
    }

    /// Enable automatic date/time display.
    pub fn with_date_time(mut self) -> Self {
        self.header_footer.show_date_time = true;
        self.header_footer.auto_date = true;
        self
    }

    /// Set a fixed date text (disables auto date).
    pub fn with_fixed_date(mut self, date_text: impl Into<String>) -> Self {
        self.header_footer.show_date_time = true;
        self.header_footer.auto_date = false;
        self.header_footer.date_time_text = Some(date_text.into());
        self
    }

    /// Set the background color (hex RGB without #).
    pub fn with_background_color(mut self, color: impl Into<String>) -> Self {
        self.background_color = Some(color.into());
        self
    }

    /// Parse handout master XML.
    pub fn parse_xml(xml: &str) -> Result<Self> {
        let mut master = Self::default();
        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);

        loop {
            match reader.read_event() {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => match e.local_name().as_ref() {
                    b"hf" => {
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"hdr" => {
                                    master.header_footer.show_header = attr.value.as_ref() == b"1"
                                },
                                b"ftr" => {
                                    master.header_footer.show_footer = attr.value.as_ref() == b"1"
                                },
                                b"sldNum" => {
                                    master.header_footer.show_slide_number =
                                        attr.value.as_ref() == b"1"
                                },
                                b"dt" => {
                                    master.header_footer.show_date_time =
                                        attr.value.as_ref() == b"1"
                                },
                                _ => {},
                            }
                        }
                    },
                    b"srgbClr" => {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"val"
                                && let Ok(color) = std::str::from_utf8(&attr.value)
                            {
                                master.background_color = Some(color.to_string());
                            }
                        }
                    },
                    _ => {},
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
        }

        Ok(master)
    }

    /// Generate handout master XML.
    /// Structure matches Apache POI PowerPoint-created files.
    pub fn to_xml(&self) -> String {
        let mut xml = String::with_capacity(8192);

        xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
        xml.push_str(r#"<p:handoutMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">"#);

        // Common slide data
        xml.push_str("<p:cSld>");

        // Background - use bgPr with solidFill like PowerPoint
        xml.push_str(r#"<p:bg><p:bgPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>"#);

        // Shape tree with ALL 4 placeholders
        xml.push_str("<p:spTree>");
        xml.push_str(
            r#"<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>"#,
        );
        xml.push_str(r#"<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>"#);

        // Header placeholder (top left)
        xml.push_str(r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Header Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="hdr" sz="quarter"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2971800" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0"/><a:lstStyle><a:lvl1pPr algn="l"><a:defRPr sz="1200"/></a:lvl1pPr></a:lstStyle><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>"#);

        // Date placeholder (top right) - with field for auto date
        xml.push_str(&format!(
            r#"<p:sp><p:nvSpPr><p:cNvPr id="3" name="Date Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="dt" sz="quarter" idx="1"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="3884613" y="0"/><a:ext cx="2971800" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0"/><a:lstStyle><a:lvl1pPr algn="r"><a:defRPr sz="1200"/></a:lvl1pPr></a:lstStyle><a:p><a:fld id="{}" type="datetime1"><a:rPr lang="en-US"/><a:pPr/><a:t>1/1/2000</a:t></a:fld><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>"#,
            generate_guid_braced()
        ));

        // Footer placeholder (bottom left)
        xml.push_str(r#"<p:sp><p:nvSpPr><p:cNvPr id="4" name="Footer Placeholder 3"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="ftr" sz="quarter" idx="2"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="8685213"/><a:ext cx="2971800" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="b"/><a:lstStyle><a:lvl1pPr algn="l"><a:defRPr sz="1200"/></a:lvl1pPr></a:lstStyle><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>"#);

        // Slide number placeholder (bottom right) - with field
        xml.push_str(&format!(
            r#"<p:sp><p:nvSpPr><p:cNvPr id="5" name="Slide Number Placeholder 4"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldNum" sz="quarter" idx="3"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="3884613" y="8685213"/><a:ext cx="2971800" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="b"/><a:lstStyle><a:lvl1pPr algn="r"><a:defRPr sz="1200"/></a:lvl1pPr></a:lstStyle><a:p><a:fld id="{}" type="slidenum"><a:rPr lang="en-US"/><a:pPr/><a:t>‹#›</a:t></a:fld><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>"#,
            generate_guid_braced()
        ));

        xml.push_str("</p:spTree>");
        xml.push_str("</p:cSld>");

        // Color map (required element)
        xml.push_str(r#"<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>"#);

        xml.push_str("</p:handoutMaster>");

        xml
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_handout_layout() {
        assert_eq!(HandoutLayout::SixSlides.as_str(), "handout6");
        assert_eq!(
            HandoutLayout::from_str("handout6").unwrap(),
            HandoutLayout::SixSlides
        );
    }

    #[test]
    fn test_handout_master_builder() {
        let master = HandoutMaster::new()
            .with_layout(HandoutLayout::ThreeSlides)
            .with_header("My Presentation")
            .with_footer("Confidential")
            .with_slide_numbers();

        assert_eq!(master.layout, HandoutLayout::ThreeSlides);
        assert!(master.header_footer.show_header);
        assert!(master.header_footer.show_slide_number);
    }
}
