/// Slide parts and related types.
///
/// This module contains parts for slides, slide layouts, and slide masters.
use crate::error::{OoxmlError, Result};
use crate::pptx::namespace::{
    is_presentationml_name, presentation_name, relationship_attribute_value,
    scan_presentationml_element_ranges,
};
use crate::pptx::shapes::base::{BaseShape, ShapeType};
use crate::pptx::shapes::textframe::extract_drawingml_text;
use litchi_opc::part::Part;
use quick_xml::events::Event;
use quick_xml::reader::NsReader;

/// A slide part.
///
/// Corresponds to `/ppt/slides/slideN.xml` in the package.
pub struct SlidePart<'a> {
    /// The underlying OPC part
    part: &'a dyn Part,
}

impl<'a> SlidePart<'a> {
    /// Create a SlidePart from an OPC Part.
    pub fn from_part(part: &'a dyn Part) -> Result<Self> {
        Ok(Self { part })
    }

    /// Get the XML bytes of the slide.
    #[inline]
    fn xml_bytes(&self) -> &[u8] {
        self.part.blob()
    }

    /// Get the slide name.
    ///
    /// Returns the name attribute from the <p:cSld> element.
    pub fn name(&self) -> Result<String> {
        presentation_name(self.xml_bytes())
    }

    /// Extract all text content from the slide.
    ///
    /// This extracts text from all `<a:t>` elements in the slide (DrawingML text).
    pub fn extract_text(&self) -> Result<String> {
        extract_drawingml_text(self.xml_bytes(), Some('\n'))
    }

    /// Get the underlying OPC part.
    #[inline]
    pub fn part(&self) -> &'a dyn Part {
        self.part
    }

    /// Parse and return all shapes on this slide.
    ///
    /// Returns a vector of BaseShape objects that can be checked for type
    /// and converted to specific shape types.
    pub fn shapes(&self) -> Result<Vec<BaseShape>> {
        let mut shapes = Vec::new();
        const TARGETS: &[&[u8]] = &[b"sp", b"pic", b"graphicFrame", b"grpSp", b"cxnSp"];
        const TYPES: &[ShapeType] = &[
            ShapeType::Shape,
            ShapeType::Picture,
            ShapeType::GraphicFrame,
            ShapeType::GroupShape,
            ShapeType::Connector,
        ];
        scan_presentationml_element_ranges(self.xml_bytes(), TARGETS, |target, start, length| {
            let start = usize::try_from(start).map_err(|_| {
                OoxmlError::InvalidFormat("shape offset does not fit usize".to_string())
            })?;
            let length = usize::try_from(length).map_err(|_| {
                OoxmlError::InvalidFormat("shape length does not fit usize".to_string())
            })?;
            let end = start.checked_add(length).ok_or_else(|| {
                OoxmlError::InvalidFormat("shape byte range overflow".to_string())
            })?;
            let xml = self.xml_bytes().get(start..end).ok_or_else(|| {
                OoxmlError::InvalidFormat("shape byte range is outside slide XML".to_string())
            })?;
            let shape_type = TYPES.get(target).ok_or_else(|| {
                OoxmlError::InvalidFormat("invalid shape range target".to_string())
            })?;
            shapes.push(BaseShape::new(xml.to_vec(), shape_type.clone()));
            Ok(())
        })?;
        Ok(shapes)
    }

    /// Get the transition effect for this slide.
    ///
    /// Parses the `<p:transition>` element from the slide XML.
    /// Returns `None` if no transition is defined.
    pub fn transition(&self) -> Result<Option<crate::pptx::transitions::SlideTransition>> {
        crate::pptx::transitions::SlideTransition::from_xml(self.xml_bytes())
    }

    /// Get the background for this slide.
    ///
    /// Parses the `<p:bg>` element from the slide XML.
    /// Returns `None` if no background is defined.
    pub fn background(&self) -> Result<Option<crate::pptx::backgrounds::SlideBackground>> {
        crate::pptx::backgrounds::SlideBackground::from_xml(self.xml_bytes())
    }
}

/// A slide layout part.
///
/// Corresponds to `/ppt/slideLayouts/slideLayoutN.xml` in the package.
pub struct SlideLayoutPart<'a> {
    /// The underlying OPC part
    part: &'a dyn Part,
}

impl<'a> SlideLayoutPart<'a> {
    /// Create a SlideLayoutPart from an OPC Part.
    pub fn from_part(part: &'a dyn Part) -> Result<Self> {
        Ok(Self { part })
    }

    /// Get the XML bytes of the layout.
    #[inline]
    fn xml_bytes(&self) -> &[u8] {
        self.part.blob()
    }

    /// Get the layout name.
    pub fn name(&self) -> Result<String> {
        presentation_name(self.xml_bytes())
    }

    /// Get the underlying OPC part.
    #[inline]
    pub fn part(&self) -> &'a dyn Part {
        self.part
    }
}

/// A slide master part.
///
/// Corresponds to `/ppt/slideMasters/slideMasterN.xml` in the package.
pub struct SlideMasterPart<'a> {
    /// The underlying OPC part
    part: &'a dyn Part,
}

impl<'a> SlideMasterPart<'a> {
    /// Create a SlideMasterPart from an OPC Part.
    pub fn from_part(part: &'a dyn Part) -> Result<Self> {
        Ok(Self { part })
    }

    /// Get the XML bytes of the master.
    #[inline]
    fn xml_bytes(&self) -> &[u8] {
        self.part.blob()
    }

    /// Get the master name.
    pub fn name(&self) -> Result<String> {
        presentation_name(self.xml_bytes())
    }

    /// Get the relationship IDs of all slide layouts in this master.
    pub fn slide_layout_rids(&self) -> Result<Vec<String>> {
        let mut reader = NsReader::from_reader(self.xml_bytes());

        let mut rids = Vec::new();

        loop {
            let decoder = reader.decoder();
            let event = reader
                .read_event()
                .map_err(|error| OoxmlError::Xml(error.to_string()))?
                .into_owned();
            let resolver = reader.resolver().clone();
            let (namespace, event) = resolver.resolve_event(event);
            match event {
                Event::Start(element) | Event::Empty(element)
                    if is_presentationml_name(&namespace, element.name(), b"sldLayoutId") =>
                {
                    if let Some(rid) =
                        relationship_attribute_value(&element, b"id", decoder, &resolver)?
                    {
                        if rid.is_empty() {
                            return Err(OoxmlError::InvalidFormat(
                                "empty slide-layout relationship ID".to_string(),
                            ));
                        }
                        rids.push(rid);
                    }
                },
                Event::Eof => break,
                _ => {},
            }
        }

        Ok(rids)
    }

    /// Get the underlying OPC part.
    #[inline]
    pub fn part(&self) -> &'a dyn Part {
        self.part
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::packuri::PackURI;
    use litchi_opc::part::BlobPart;

    const P: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
    const A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
    const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    fn part(path: &str, xml: impl Into<Vec<u8>>) -> BlobPart {
        BlobPart::new(
            PackURI::new(path).unwrap(),
            "application/xml".to_string(),
            xml.into(),
        )
    }

    #[test]
    fn slide_metadata_and_text_resolve_namespaces() {
        let xml = format!(
            r#"<q:sld xmlns:q="{P}" xmlns:d="{A}" xmlns:f="urn:foreign">
                <f:cSld name="Spoof"/><q:cSld name="A &amp; B"><q:spTree>
                    <q:sp><q:txBody><d:p><d:r><d:t xml:space="preserve"> One &amp; <![CDATA[Two]]></d:t></d:r></d:p>
                        <d:p><d:r><d:t>Three</d:t></d:r></d:p></q:txBody></q:sp>
                    <f:t>Ignored</f:t>
                </q:spTree></q:cSld></q:sld>"#
        );
        let blob = part("/ppt/slides/slide1.xml", xml);
        let slide = SlidePart::from_part(&blob).unwrap();
        assert_eq!(slide.name().unwrap(), "A & B");
        assert_eq!(slide.extract_text().unwrap(), " One & Two\nThree");
    }

    #[test]
    fn shapes_are_namespace_filtered_and_preserve_source_xml() {
        let xml = format!(
            r#"<p:sld xmlns:p="{P}" xmlns:a="{A}" xmlns:f="urn:foreign"><p:cSld><p:spTree>
                <f:sp><f:cNvPr name="Spoof"/></f:sp>
                <p:sp custom="kept"><p:nvSpPr><p:cNvPr name="Real &amp; Name"/></p:nvSpPr>
                    <p:txBody><a:p><a:r><a:t><![CDATA[A < B]]></a:t></a:r></a:p></p:txBody>
                    <!--keep-comment--><p:extLst><f:data key="value"/></p:extLst>
                </p:sp>
                <p:pic/><p:graphicFrame/><p:cxnSp/>
            </p:spTree></p:cSld></p:sld>"#
        );
        let blob = part("/ppt/slides/slide1.xml", xml);
        let slide = SlidePart::from_part(&blob).unwrap();
        let mut shapes = slide.shapes().unwrap();
        assert_eq!(shapes.len(), 4);
        assert_eq!(shapes[0].shape_type(), &ShapeType::Shape);
        assert_eq!(shapes[0].name().unwrap(), "Real & Name");
        let raw = std::str::from_utf8(shapes[0].xml_bytes()).unwrap();
        assert!(raw.starts_with("<p:sp custom=\"kept\">"));
        assert!(raw.contains("<![CDATA[A < B]]>"));
        assert!(raw.contains("<!--keep-comment-->"));
        assert!(raw.ends_with("</p:sp>"));
        assert_eq!(shapes[1].shape_type(), &ShapeType::Picture);
        assert_eq!(shapes[2].shape_type(), &ShapeType::GraphicFrame);
        assert_eq!(shapes[3].shape_type(), &ShapeType::Connector);
    }

    #[test]
    fn strict_aliases_and_relationship_aliases_are_supported() {
        let xml = r#"<x:sldMaster xmlns:x="http://purl.oclc.org/ooxml/presentationml/main"
            xmlns:rel="http://purl.oclc.org/ooxml/officeDocument/relationships"
            xmlns:f="urn:foreign"><x:cSld name="Strict"/>
            <x:sldLayoutIdLst><f:sldLayoutId rel:id="spoof"/>
                <x:sldLayoutId f:id="wrong" rel:id="layout-alpha"/>
            </x:sldLayoutIdLst></x:sldMaster>"#;
        let blob = part("/ppt/slideMasters/slideMaster1.xml", xml);
        let master = SlideMasterPart::from_part(&blob).unwrap();
        assert_eq!(master.name().unwrap(), "Strict");
        assert_eq!(master.slide_layout_rids().unwrap(), ["layout-alpha"]);
    }

    #[test]
    fn malformed_slide_xml_is_reported() {
        let xml = format!(r#"<p:sld xmlns:p="{P}"><p:sp>"#);
        let blob = part("/ppt/slides/slide1.xml", xml);
        let slide = SlidePart::from_part(&blob).unwrap();
        assert!(slide.shapes().is_err());
    }

    #[test]
    fn duplicate_relationship_attributes_are_rejected() {
        let xml = format!(
            r#"<p:sldMaster xmlns:p="{P}" xmlns:r="{R}" xmlns:q="{R}">
                <p:sldLayoutId r:id="one" q:id="two"/></p:sldMaster>"#
        );
        let blob = part("/ppt/slideMasters/slideMaster1.xml", xml);
        let master = SlideMasterPart::from_part(&blob).unwrap();
        assert!(master.slide_layout_rids().is_err());
    }
}
