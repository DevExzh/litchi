//! Handout master support for PowerPoint presentations.
//!
//! Handout masters define the layout for printed handouts that show
//! multiple slides per page.

use super::model::*;
use crate::presentation_properties::metadata::new_guid;
use crate::{Error, Result};
use quick_xml::Reader;
use quick_xml::events::Event;

impl Master {
    /// Parse handout master XML.
    pub fn parse_xml(xml: &str) -> Result<Self> {
        let mut master = Self::default();
        let xml = litchi_ooxml_common::mce::process_str(xml)?;
        let mut reader = Reader::from_str(xml.as_ref());
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
                Err(e) => return Err(Error::Xml(e.to_string())),
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
            new_guid()
        ));

        // Footer placeholder (bottom left)
        xml.push_str(r#"<p:sp><p:nvSpPr><p:cNvPr id="4" name="Footer Placeholder 3"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="ftr" sz="quarter" idx="2"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="8685213"/><a:ext cx="2971800" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="b"/><a:lstStyle><a:lvl1pPr algn="l"><a:defRPr sz="1200"/></a:lvl1pPr></a:lstStyle><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>"#);

        // Slide number placeholder (bottom right) - with field
        xml.push_str(&format!(
            r#"<p:sp><p:nvSpPr><p:cNvPr id="5" name="Slide Number Placeholder 4"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldNum" sz="quarter" idx="3"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="3884613" y="8685213"/><a:ext cx="2971800" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="b"/><a:lstStyle><a:lvl1pPr algn="r"><a:defRPr sz="1200"/></a:lvl1pPr></a:lstStyle><a:p><a:fld id="{}" type="slidenum"><a:rPr lang="en-US"/><a:pPr/><a:t>‹#›</a:t></a:fld><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>"#,
            new_guid()
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
    use std::str::FromStr;

    #[test]
    fn test_handout_layout() {
        assert_eq!(Layout::SixSlides.as_str(), "handout6");
        assert_eq!(Layout::from_str("handout6").unwrap(), Layout::SixSlides);
    }

    #[test]
    fn test_handout_master_builder() {
        let master = Master::new()
            .with_layout(Layout::ThreeSlides)
            .with_header("My Presentation")
            .with_footer("Confidential")
            .with_slide_numbers();

        assert_eq!(master.layout, Layout::ThreeSlides);
        assert!(master.header_footer.show_header);
        assert!(master.header_footer.show_slide_number);
    }
}
