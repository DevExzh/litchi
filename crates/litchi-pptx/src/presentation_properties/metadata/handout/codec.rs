//! Handout master support for `PowerPoint` presentations.
//!
//! Handout masters define the layout for printed handouts that show
//! multiple slides per page.

use super::super::{BoundedString, Limits, process_str};
use super::model::Master;
use crate::presentation_properties::metadata::new_guid;
use crate::{Error, Result};
use litchi_ooxml_common::xml::attributes::{count_up_to, first_wins};
use quick_xml::Reader;
use quick_xml::events::Event;

impl Master {
    /// Parse handout master XML.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn parse_xml(xml: &str, limits: Limits) -> Result<Self> {
        if xml.len() > limits.max_input_bytes() {
            return Err(Error::Limit {
                kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::InputBytes),
                observed: xml.len(),
                maximum: limits.max_input_bytes(),
            });
        }
        let mut master = Self::default();
        let xml = process_str(xml, limits)?;
        if xml.len() > limits.max_output_bytes() {
            return Err(Error::Limit {
                kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::OutputBytes),
                observed: xml.len(),
                maximum: limits.max_output_bytes(),
            });
        }
        let mut reader = Reader::from_str(xml.as_ref());
        reader.config_mut().trim_text(true);
        let mut nodes = 0usize;
        let mut depth = 0usize;

        loop {
            match reader.read_event() {
                Ok(Event::Start(element)) => {
                    depth = depth.checked_add(1).ok_or_else(|| Error::Limit {
                        kind: crate::opened::LimitKind::Codec(
                            crate::opened::CodecLimitKind::XmlDepth,
                        ),
                        observed: usize::MAX,
                        maximum: limits.max_depth(),
                    })?;
                    if depth > limits.max_depth() {
                        return Err(Error::Limit {
                            kind: crate::opened::LimitKind::Codec(
                                crate::opened::CodecLimitKind::XmlDepth,
                            ),
                            observed: depth,
                            maximum: limits.max_depth(),
                        });
                    }
                    parse_element(&element, reader.decoder(), &mut master, &mut nodes, limits)?;
                },
                Ok(Event::Empty(element)) => {
                    check_empty_depth(depth, limits)?;
                    parse_element(&element, reader.decoder(), &mut master, &mut nodes, limits)?
                },
                Ok(Event::End(_)) => {
                    depth = depth.checked_sub(1).ok_or_else(|| {
                        Error::Invalid("handout-master XML has an unmatched end tag".into())
                    })?;
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(Error::Xml(e.to_string())),
                _ => {},
            }
        }

        if depth != 0 {
            return Err(Error::Invalid(
                "handout-master XML has an unterminated element".into(),
            ));
        }

        Ok(master)
    }

    /// Generate handout master XML.
    /// Structure matches Apache POI PowerPoint-created files.
    pub fn to_xml(&self, limits: Limits) -> Result<String> {
        let mut xml = BoundedString::new(limits, "serialized handout-master XML");

        xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#)?;
        xml.push_str(r#"<p:handoutMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">"#)?;

        // Common slide data
        xml.push_str("<p:cSld>")?;

        // Background - use bgPr with solidFill like PowerPoint
        xml.push_str(r#"<p:bg><p:bgPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>"#)?;

        // Shape tree with ALL 4 placeholders
        xml.push_str("<p:spTree>")?;
        xml.push_str(
            r#"<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>"#,
        )?;
        xml.push_str(r#"<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>"#)?;

        // Header placeholder (top left)
        xml.push_str(r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Header Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="hdr" sz="quarter"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2971800" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0"/><a:lstStyle><a:lvl1pPr algn="l"><a:defRPr sz="1200"/></a:lvl1pPr></a:lstStyle><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>"#)?;

        // Date placeholder (top right) - with field for auto date
        xml.write_fmt(format_args!(
            r#"<p:sp><p:nvSpPr><p:cNvPr id="3" name="Date Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="dt" sz="quarter" idx="1"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="3884613" y="0"/><a:ext cx="2971800" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0"/><a:lstStyle><a:lvl1pPr algn="r"><a:defRPr sz="1200"/></a:lvl1pPr></a:lstStyle><a:p><a:fld id="{}" type="datetime1"><a:rPr lang="en-US"/><a:pPr/><a:t>1/1/2000</a:t></a:fld><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>"#,
            new_guid()
        ))?;

        // Footer placeholder (bottom left)
        xml.push_str(r#"<p:sp><p:nvSpPr><p:cNvPr id="4" name="Footer Placeholder 3"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="ftr" sz="quarter" idx="2"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="8685213"/><a:ext cx="2971800" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="b"/><a:lstStyle><a:lvl1pPr algn="l"><a:defRPr sz="1200"/></a:lvl1pPr></a:lstStyle><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>"#)?;

        // Slide number placeholder (bottom right) - with field
        xml.write_fmt(format_args!(
            r#"<p:sp><p:nvSpPr><p:cNvPr id="5" name="Slide Number Placeholder 4"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldNum" sz="quarter" idx="3"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="3884613" y="8685213"/><a:ext cx="2971800" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="b"/><a:lstStyle><a:lvl1pPr algn="r"><a:defRPr sz="1200"/></a:lvl1pPr></a:lstStyle><a:p><a:fld id="{}" type="slidenum"><a:rPr lang="en-US"/><a:pPr/><a:t>‹#›</a:t></a:fld><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>"#,
            new_guid()
        ))?;

        xml.push_str("</p:spTree>")?;
        xml.push_str("</p:cSld>")?;

        // Color map (required element)
        xml.push_str(r#"<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>"#)?;

        xml.push_str("</p:handoutMaster>")?;
        Ok(xml.finish())
    }
}

fn parse_element(
    element: &quick_xml::events::BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    master: &mut Master,
    nodes: &mut usize,
    limits: Limits,
) -> Result<()> {
    *nodes = nodes.checked_add(1).ok_or_else(|| Error::Limit {
        kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::XmlNodes),
        observed: usize::MAX,
        maximum: limits.max_nodes(),
    })?;
    if *nodes > limits.max_nodes() {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::XmlNodes),
            observed: *nodes,
            maximum: limits.max_nodes(),
        });
    }
    check_attributes(element, decoder, limits)?;
    match element.local_name().as_ref() {
        b"hf" => {
            for attr in first_wins(element).flatten() {
                match attr.key.as_ref() {
                    b"hdr" => master.header_footer.show_header = attr.value.as_ref() == b"1",
                    b"ftr" => master.header_footer.show_footer = attr.value.as_ref() == b"1",
                    b"sldNum" => {
                        master.header_footer.show_slide_number = attr.value.as_ref() == b"1";
                    },
                    b"dt" => master.header_footer.show_date_time = attr.value.as_ref() == b"1",
                    _ => {},
                }
            }
        },
        b"srgbClr" => {
            for attr in first_wins(element).flatten() {
                if attr.key.as_ref() == b"val"
                    && let Ok(color) = std::str::from_utf8(&attr.value)
                {
                    master.background_color = Some(color.to_string());
                }
            }
        },
        _ => {},
    }
    Ok(())
}

fn check_attributes(
    element: &quick_xml::events::BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    limits: Limits,
) -> Result<()> {
    let count = count_up_to(element, limits.max_attributes());
    if count > limits.max_attributes() {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::XmlAttributes),
            observed: count,
            maximum: limits.max_attributes(),
        });
    }
    // `first_wins` is the established semantic policy for these metadata
    // elements: duplicate names are accepted and the first value is used.
    // The physical count above includes every duplicate; only the first value
    // is retained by the model and checked for semantic string bounds below.
    for attribute in first_wins(element) {
        let attribute = match attribute {
            Ok(attribute) => attribute,
            Err(quick_xml::events::attributes::AttrError::Duplicated(..)) => continue,
            Err(error) => return Err(Error::Xml(error.to_string())),
        };
        if attribute.key.as_ref().len() > limits.max_string_bytes() {
            return Err(Error::Limit {
                kind: crate::opened::LimitKind::Codec(
                    crate::opened::CodecLimitKind::XmlAttributeBytes,
                ),
                observed: attribute.key.as_ref().len(),
                maximum: limits.max_string_bytes(),
            });
        }
        if attribute.value.as_ref().len() > limits.max_string_bytes() {
            return Err(Error::Limit {
                kind: crate::opened::LimitKind::Codec(
                    crate::opened::CodecLimitKind::XmlAttributeBytes,
                ),
                observed: attribute.value.as_ref().len(),
                maximum: limits.max_string_bytes(),
            });
        }
        let value = attribute
            .decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, decoder)
            .map_err(|error| Error::Xml(error.to_string()))?;
        if value.len() > limits.max_string_bytes() {
            return Err(Error::Limit {
                kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::StringBytes),
                observed: value.len(),
                maximum: limits.max_string_bytes(),
            });
        }
    }
    Ok(())
}

fn check_empty_depth(depth: usize, limits: Limits) -> Result<()> {
    let observed = depth.checked_add(1).unwrap_or(usize::MAX);
    if observed > limits.max_depth() {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::XmlDepth),
            observed,
            maximum: limits.max_depth(),
        });
    }
    Ok(())
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
mod tests {
    use super::super::model::Layout;
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

    /// `19_999` distinct filler attributes, then `name="first"`, then `20_000`
    /// repeats of `name="repeat"`: a checked attribute iterator that skips its
    /// duplicate errors scans every earlier name for each repeat.
    fn crowded(name: &str, first: &str, repeat: &str) -> String {
        const DISTINCT: usize = 20_000;
        let mut attributes = String::new();
        for index in 0..DISTINCT - 1 {
            attributes.push_str(&format!(" n{index}=\"{index}\""));
        }
        attributes.push_str(&format!(" {name}=\"{first}\""));
        let repeated = format!(" {name}=\"{repeat}\"");
        for _ in 0..DISTINCT {
            attributes.push_str(&repeated);
        }
        attributes
    }

    #[test]
    fn parsed_handout_attributes_keep_their_first_occurrence() {
        let single = Master::parse_xml(
            r#"<p:handoutMaster><p:hf hdr="0" ftr="1" sldNum="1" dt="0"/><a:srgbClr val="112233"/></p:handoutMaster>"#,
            Limits::DEFAULT,
        )
        .unwrap();
        assert!(!single.header_footer.show_header);
        assert!(single.header_footer.show_footer);
        assert!(single.header_footer.show_slide_number);
        assert!(!single.header_footer.show_date_time);
        assert_eq!(single.background_color.as_deref(), Some("112233"));

        let repeated = Master::parse_xml(
            r#"<p:handoutMaster><p:hf hdr="0" ftr="1" hdr="1" ftr="0" sldNum="1" dt="0" dt="1" sldNum="0"/><a:srgbClr val="112233" val="445566"/></p:handoutMaster>"#,
            Limits::DEFAULT,
        )
        .unwrap();
        assert!(!repeated.header_footer.show_header);
        assert!(repeated.header_footer.show_footer);
        assert!(repeated.header_footer.show_slide_number);
        assert!(!repeated.header_footer.show_date_time);
        assert_eq!(repeated.background_color.as_deref(), Some("112233"));
    }

    #[test]
    fn crowded_handout_attributes_are_read_first_wins() {
        let xml = format!(
            "<p:handoutMaster><p:hf{} ftr=\"1\"/><a:srgbClr{}/></p:handoutMaster>",
            crowded("hdr", "1", "0"),
            crowded("val", "112233", "445566"),
        );
        let master = Master::parse_xml(&xml, Limits::DEFAULT).unwrap();
        assert!(master.header_footer.show_header);
        assert!(master.header_footer.show_footer);
        assert!(!master.header_footer.show_slide_number);
        assert_eq!(master.background_color.as_deref(), Some("112233"));
    }

    #[test]
    fn duplicate_attributes_count_against_the_selected_physical_limit() {
        let xml = r#"<p:handoutMaster><p:hf hdr="1" hdr="0" ftr="1"/></p:handoutMaster>"#;
        let exact = Limits::DEFAULT.with_attributes(3).unwrap();
        Master::parse_xml(xml, exact).expect("three physical attributes fit");

        let one_under = Limits::DEFAULT.with_attributes(2).unwrap();
        assert!(matches!(
            Master::parse_xml(xml, one_under),
            Err(Error::Limit {
                kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::XmlAttributes),
                observed: 3,
                maximum: 2,
            })
        ));
    }

    #[test]
    fn empty_handout_depth_is_checked_before_attributes() {
        let limits = Limits::DEFAULT.with_depth(1).expect("depth limit");
        let exact = r#"<p:handoutMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>"#;
        Master::parse_xml(exact, limits).expect("depth-one empty root");

        let one_over = r#"<p:handoutMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:hf/></p:handoutMaster>"#;
        assert_xml_depth_limit(
            Master::parse_xml(one_over, limits).expect_err("depth-two empty element"),
            2,
            1,
        );
        assert!(matches!(
            Master::parse_xml(
                "<p:handoutMaster xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">",
                Limits::DEFAULT,
            ),
            Err(Error::Invalid(_))
        ));
    }

    fn assert_xml_depth_limit(error: Error, observed: usize, maximum: usize) {
        match error {
            Error::Limit {
                kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::XmlDepth),
                observed: actual,
                maximum: configured,
            } => {
                assert_eq!(actual, observed);
                assert_eq!(configured, maximum);
            },
            other => panic!("expected XML depth limit, got {other:?}"),
        }
    }
}
