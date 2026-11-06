//! Handout master support for `PowerPoint` presentations.
//!
//! Handout masters define the layout for printed handouts that show
//! multiple slides per page.

use super::super::{BoundedString, Limits};
use super::model::Master;
use crate::presentation_properties::metadata::new_guid;
use crate::{Error, Result};
use litchi_ooxml_common::xml::attributes::{count_up_to, first_wins};

impl Master {
    /// Read direct header/footer placeholders, their visibility, and a direct
    /// solid RGB background from handout master XML. Other shapes, styling and
    /// theme fills are outside this projection.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn parse_xml(xml: &str, limits: Limits) -> Result<Self> {
        super::reader::parse(xml, limits)
    }

    /// Generate a fresh handout master from the represented text, flags and
    /// background. This does not preserve a source shape tree. The print
    /// layout belongs in `presProps.xml`, not in this part. The `layout` and
    /// `show_slide_images` authoring hints are not serialized.
    pub fn to_xml(&self, limits: Limits) -> Result<String> {
        for value in [
            self.header_footer.header_text.as_deref(),
            self.header_footer.footer_text.as_deref(),
            self.header_footer.date_time_text.as_deref(),
        ]
        .into_iter()
        .flatten()
        {
            validate_text(value, limits)?;
        }
        if let Some(color) = &self.background_color {
            super::super::check_string(color, limits, "handout background")?;
            validate_color(color)?;
        }
        let mut xml = BoundedString::new(limits, "serialized handout-master XML");
        xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#)?;
        xml.push_str(r#"<p:handoutMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:bg><p:bgPr><a:solidFill>"#)?;
        if let Some(color) = &self.background_color {
            xml.write_fmt(format_args!(r#"<a:srgbClr val="{color}"/>"#))?;
        } else {
            xml.push_str(r#"<a:schemeClr val="bg1"/>"#)?;
        }
        xml.push_str(r#"</a:solidFill><a:effectLst/></p:bgPr></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>"#)?;
        write_placeholder(&mut xml, 0, self.header_footer.header_text.as_deref(), None)?;
        write_placeholder(
            &mut xml,
            1,
            self.header_footer.date_time_text.as_deref(),
            self.header_footer.auto_date.then_some("datetime1"),
        )?;
        write_placeholder(&mut xml, 2, self.header_footer.footer_text.as_deref(), None)?;
        write_placeholder(&mut xml, 3, Some("‹#›"), Some("slidenum"))?;
        xml.push_str(r#"</p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>"#)?;
        xml.write_fmt(format_args!(
            r#"<p:hf hdr="{}" ftr="{}" dt="{}" sldNum="{}"/></p:handoutMaster>"#,
            u8::from(self.header_footer.show_header),
            u8::from(self.header_footer.show_footer),
            u8::from(self.header_footer.show_date_time),
            u8::from(self.header_footer.show_slide_number)
        ))?;
        Ok(xml.finish())
    }
}

fn write_placeholder(
    xml: &mut BoundedString,
    index: usize,
    text: Option<&str>,
    field: Option<&str>,
) -> Result<()> {
    let (kind, name, x, y) = match index {
        0 => ("hdr", "Header Placeholder 1", 0, 0),
        1 => ("dt", "Date Placeholder 2", 3884613, 0),
        2 => ("ftr", "Footer Placeholder 3", 0, 8685213),
        _ => ("sldNum", "Slide Number Placeholder 4", 3884613, 8685213),
    };
    let alignment = if index.is_multiple_of(2) { "l" } else { "r" };
    let anchor = if index >= 2 { r#" anchor="b""# } else { "" };
    xml.write_fmt(format_args!(r#"<p:sp><p:nvSpPr><p:cNvPr id="{}" name="{name}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="{kind}" sz="quarter" idx="{index}"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="2971800" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0"{anchor}/><a:lstStyle><a:lvl1pPr algn="{alignment}"><a:defRPr sz="1200"/></a:lvl1pPr></a:lstStyle><a:p>"#, index+2))?;
    if let Some(field) = field {
        xml.write_fmt(format_args!(
            r#"<a:fld id="{}" type="{field}"><a:rPr lang="en-US"/><a:pPr/>"#,
            new_guid()
        ))?;
        if let Some(text) = text {
            write_text(xml, text)?;
        }
        xml.push_str("</a:fld>")?;
    } else if let Some(text) = text {
        for (index, line) in text.split('\n').enumerate() {
            if index != 0 {
                xml.push_str("<a:br/>")?;
            }
            xml.push_str("<a:r>")?;
            write_text(xml, line)?;
            xml.push_str("</a:r>")?;
        }
    }
    xml.push_str(r#"<a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>"#)
}

fn write_text(xml: &mut BoundedString, text: &str) -> Result<()> {
    xml.push_str(r#"<a:t xml:space="preserve">"#)?;
    xml.push_escaped_xml(text)?;
    xml.push_str("</a:t>")
}

pub(super) fn validate_color(value: &str) -> Result<()> {
    if value.len() != 6 || !value.bytes().all(|v| v.is_ascii_hexdigit()) {
        return Err(Error::Invalid(
            "handout background must be a six-digit RGB value".into(),
        ));
    }
    Ok(())
}

pub(super) fn validate_text(value: &str, limits: Limits) -> Result<()> {
    super::super::check_string(value, limits, "handout placeholder text")?;
    if value.chars().any(|c| !matches!(c, '\u{9}' | '\u{a}' | '\u{d}' | '\u{20}'..='\u{d7ff}' | '\u{e000}'..='\u{fffd}' | '\u{10000}'..='\u{10ffff}')) {
        return Err(Error::Invalid("handout text contains an invalid XML character".into()));
    }
    Ok(())
}

pub(super) fn check_element(
    element: &quick_xml::events::BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
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

pub(super) fn check_empty_depth(depth: usize, limits: Limits) -> Result<()> {
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
            r#"<p:handoutMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="112233"/></a:solidFill></p:bgPr></p:bg></p:cSld><p:hf hdr="0" ftr="1" sldNum="1" dt="0"/></p:handoutMaster>"#,
            Limits::DEFAULT,
        )
        .unwrap();
        assert!(!single.header_footer.show_header);
        assert!(single.header_footer.show_footer);
        assert!(single.header_footer.show_slide_number);
        assert!(!single.header_footer.show_date_time);
        assert_eq!(single.background_color.as_deref(), Some("112233"));

        let repeated = Master::parse_xml(
            r#"<p:handoutMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="112233" val="445566"/></a:solidFill></p:bgPr></p:bg></p:cSld><p:hf hdr="0" ftr="1" hdr="1" ftr="0" sldNum="1" dt="0" dt="1" sldNum="0"/></p:handoutMaster>"#,
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
            r#"<p:handoutMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr{}/></a:solidFill></p:bgPr></p:bg></p:cSld><p:hf{} ftr="1"/></p:handoutMaster>"#,
            crowded("val", "112233", "445566"),
            crowded("hdr", "1", "0"),
        );
        let master = Master::parse_xml(&xml, Limits::DEFAULT).unwrap();
        assert!(master.header_footer.show_header);
        assert!(master.header_footer.show_footer);
        assert!(master.header_footer.show_slide_number); // Omitted hf attribute defaults to true.
        assert_eq!(master.background_color.as_deref(), Some("112233"));
    }

    #[test]
    fn duplicate_attributes_count_against_the_selected_physical_limit() {
        let xml = r#"<p:handoutMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:hf hdr="1" hdr="0" ftr="1"/></p:handoutMaster>"#;
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
