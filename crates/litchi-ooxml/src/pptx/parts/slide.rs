/// Slide parts and related types.
///
/// This module contains parts for slides, slide layouts, and slide masters.
use crate::common::xml::unqualified_attribute_value;
use crate::error::{OoxmlError, Result};
use crate::pptx::namespace::{
    is_presentationml_name, presentation_name, relationship_attribute_value,
    scan_presentationml_element_ranges,
};
use crate::pptx::shapes::base::{BaseShape, ShapeType};
use crate::pptx::shapes::textframe::extract_drawingml_text;
use litchi_opc::part::Part;
use quick_xml::events::{BytesStart, Event};
use quick_xml::reader::NsReader;
use std::sync::Arc;

fn processed(part: &dyn Part) -> Result<Arc<Vec<u8>>> {
    Ok(match crate::common::mce::process_ooxml(part.blob())? {
        std::borrow::Cow::Borrowed(_) => part.blob_arc(),
        std::borrow::Cow::Owned(v) => Arc::new(v),
    })
}

fn parse_shapes(xml: &[u8]) -> Result<Vec<BaseShape>> {
    let mut shapes = Vec::new();
    const TARGETS: &[&[u8]] = &[b"sp", b"pic", b"graphicFrame", b"grpSp", b"cxnSp"];
    const TYPES: &[ShapeType] = &[
        ShapeType::Shape,
        ShapeType::Picture,
        ShapeType::GraphicFrame,
        ShapeType::GroupShape,
        ShapeType::Connector,
    ];
    scan_presentationml_element_ranges(xml, TARGETS, |target, start, length| {
        let start = usize::try_from(start).map_err(|_| {
            OoxmlError::InvalidFormat("shape offset does not fit usize".to_string())
        })?;
        let length = usize::try_from(length).map_err(|_| {
            OoxmlError::InvalidFormat("shape length does not fit usize".to_string())
        })?;
        let end = start
            .checked_add(length)
            .ok_or_else(|| OoxmlError::InvalidFormat("shape byte range overflow".to_string()))?;
        let xml = xml.get(start..end).ok_or_else(|| {
            OoxmlError::InvalidFormat("shape byte range is outside slide XML".to_string())
        })?;
        let shape_type = TYPES
            .get(target)
            .ok_or_else(|| OoxmlError::InvalidFormat("invalid shape range target".to_string()))?;
        shapes.push(BaseShape::new(xml.to_vec(), shape_type.clone()));
        Ok(())
    })?;
    Ok(shapes)
}

fn filter_placeholders(shapes: Vec<BaseShape>) -> Vec<BaseShape> {
    shapes
        .into_iter()
        .filter(BaseShape::is_placeholder)
        .collect()
}

/// Master-content visibility settings declared by a slide or slide layout.
///
/// PresentationML defaults both settings to true when their corresponding
/// attributes are omitted.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct MasterVisibility {
    show_master_shapes: bool,
    show_master_placeholder_animations: bool,
}

impl MasterVisibility {
    /// Whether shapes from the slide master are shown.
    #[inline]
    pub const fn shows_master_shapes(&self) -> bool {
        self.show_master_shapes
    }

    /// Whether animations on placeholders from the slide master are shown.
    #[inline]
    pub const fn shows_master_placeholder_animations(&self) -> bool {
        self.show_master_placeholder_animations
    }

    fn from_element(
        element: &BytesStart<'_>,
        decoder: quick_xml::encoding::Decoder,
        root_label: &str,
    ) -> Result<Self> {
        Ok(Self {
            show_master_shapes: parse_boolean_attribute(
                element,
                b"showMasterSp",
                decoder,
                root_label,
                true,
            )?,
            show_master_placeholder_animations: parse_boolean_attribute(
                element,
                b"showMasterPhAnim",
                decoder,
                root_label,
                true,
            )?,
        })
    }
}

impl Default for MasterVisibility {
    fn default() -> Self {
        Self {
            show_master_shapes: true,
            show_master_placeholder_animations: true,
        }
    }
}

/// Root-level metadata declared by a slide layout.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SlideLayoutMetadata {
    matching_name: String,
    layout_type: String,
    preserve: bool,
    user_drawn: bool,
}

impl SlideLayoutMetadata {
    /// Return the name used to match this layout during template changes.
    ///
    /// This is empty when the matchingName attribute is omitted.
    #[inline]
    pub fn matching_name(&self) -> &str {
        &self.matching_name
    }

    /// Return the layout type token.
    ///
    /// This is cust when the type attribute is omitted.
    #[inline]
    pub fn layout_type(&self) -> &str {
        &self.layout_type
    }

    /// Whether the layout is retained after its dependent slides are removed.
    #[inline]
    pub const fn is_preserved(&self) -> bool {
        self.preserve
    }

    /// Whether the layout is marked as user-drawn.
    #[inline]
    pub const fn is_user_drawn(&self) -> bool {
        self.user_drawn
    }

    fn from_element(
        element: &BytesStart<'_>,
        decoder: quick_xml::encoding::Decoder,
    ) -> Result<Self> {
        Ok(Self {
            matching_name: unqualified_attribute_value(element, b"matchingName", decoder)?
                .unwrap_or_default(),
            layout_type: unqualified_attribute_value(element, b"type", decoder)?
                .unwrap_or_else(|| "cust".to_string()),
            preserve: parse_boolean_attribute(
                element,
                b"preserve",
                decoder,
                "slide layout",
                false,
            )?,
            user_drawn: parse_boolean_attribute(
                element,
                b"userDrawn",
                decoder,
                "slide layout",
                false,
            )?,
        })
    }
}

impl Default for SlideLayoutMetadata {
    fn default() -> Self {
        Self {
            matching_name: String::new(),
            layout_type: "cust".to_string(),
            preserve: false,
            user_drawn: false,
        }
    }
}

/// Header and footer placeholder visibility declared by a master or layout.
///
/// PresentationML defaults all four settings to true when their corresponding
/// attributes are omitted.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SlideHeaderFooterVisibility {
    show_date_time: bool,
    show_footer: bool,
    show_header: bool,
    show_slide_number: bool,
}

impl SlideHeaderFooterVisibility {
    /// Whether the date and time placeholder is shown.
    #[inline]
    pub const fn shows_date_time(&self) -> bool {
        self.show_date_time
    }

    /// Whether the footer placeholder is shown.
    #[inline]
    pub const fn shows_footer(&self) -> bool {
        self.show_footer
    }

    /// Whether the header placeholder is shown.
    #[inline]
    pub const fn shows_header(&self) -> bool {
        self.show_header
    }

    /// Whether the slide-number placeholder is shown.
    #[inline]
    pub const fn shows_slide_number(&self) -> bool {
        self.show_slide_number
    }

    fn from_element(
        element: &BytesStart<'_>,
        decoder: quick_xml::encoding::Decoder,
        root_label: &str,
    ) -> Result<Self> {
        Ok(Self {
            show_date_time: parse_boolean_attribute(element, b"dt", decoder, root_label, true)?,
            show_footer: parse_boolean_attribute(element, b"ftr", decoder, root_label, true)?,
            show_header: parse_boolean_attribute(element, b"hdr", decoder, root_label, true)?,
            show_slide_number: parse_boolean_attribute(
                element, b"sldNum", decoder, root_label, true,
            )?,
        })
    }
}

impl Default for SlideHeaderFooterVisibility {
    fn default() -> Self {
        Self {
            show_date_time: true,
            show_footer: true,
            show_header: true,
            show_slide_number: true,
        }
    }
}

fn parse_master_visibility(
    xml: &[u8],
    root_name: &[u8],
    root_label: &str,
) -> Result<MasterVisibility> {
    let (element, decoder) = read_root_element(xml, root_name, root_label)?;
    MasterVisibility::from_element(&element, decoder, root_label)
}

fn parse_slide_show(xml: &[u8]) -> Result<bool> {
    let (element, decoder) = read_root_element(xml, b"sld", "slide")?;
    parse_boolean_attribute(&element, b"show", decoder, "slide", true)
}

fn parse_slide_layout_metadata(xml: &[u8]) -> Result<SlideLayoutMetadata> {
    let (element, decoder) = read_root_element(xml, b"sldLayout", "slide layout")?;
    SlideLayoutMetadata::from_element(&element, decoder)
}

fn parse_slide_master_preserve(xml: &[u8]) -> Result<bool> {
    let (element, decoder) = read_root_element(xml, b"sldMaster", "slide master")?;
    parse_boolean_attribute(&element, b"preserve", decoder, "slide master", false)
}

fn parse_header_footer_visibility(
    xml: &[u8],
    root_name: &[u8],
    root_label: &str,
) -> Result<Option<SlideHeaderFooterVisibility>> {
    let mut reader = NsReader::from_reader(xml);
    let mut depth = 0usize;
    let mut saw_root = false;
    let mut header_footer = None;

    loop {
        let decoder = reader.decoder();
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| OoxmlError::Xml(error.to_string()))?;
        match event {
            Event::Start(element) => {
                depth = depth.checked_add(1).ok_or_else(|| {
                    OoxmlError::InvalidFormat(format!("{root_label} XML nesting is too deep"))
                })?;
                if depth == 1 {
                    if saw_root {
                        return Err(OoxmlError::InvalidFormat(format!(
                            "{root_label} XML has multiple roots"
                        )));
                    }
                    require_presentationml_root(&namespace, &element, root_name, root_label)?;
                    saw_root = true;
                } else if depth == 2 && is_presentationml_name(&namespace, element.name(), b"hf") {
                    store_header_footer_visibility(
                        &mut header_footer,
                        SlideHeaderFooterVisibility::from_element(&element, decoder, root_label)?,
                        root_label,
                    )?;
                }
            },
            Event::Empty(element) => {
                if depth == 0 {
                    if saw_root {
                        return Err(OoxmlError::InvalidFormat(format!(
                            "{root_label} XML has multiple roots"
                        )));
                    }
                    require_presentationml_root(&namespace, &element, root_name, root_label)?;
                    saw_root = true;
                } else if depth == 1 && is_presentationml_name(&namespace, element.name(), b"hf") {
                    store_header_footer_visibility(
                        &mut header_footer,
                        SlideHeaderFooterVisibility::from_element(&element, decoder, root_label)?,
                        root_label,
                    )?;
                }
            },
            Event::End(_) => {
                depth = depth.checked_sub(1).ok_or_else(|| {
                    OoxmlError::InvalidFormat(format!("invalid {root_label} XML nesting"))
                })?;
            },
            Event::Eof => break,
            _ => {},
        }
    }

    if depth != 0 || !saw_root {
        return Err(OoxmlError::InvalidFormat(format!(
            "unterminated {root_label} XML"
        )));
    }
    Ok(header_footer)
}

fn store_header_footer_visibility(
    slot: &mut Option<SlideHeaderFooterVisibility>,
    value: SlideHeaderFooterVisibility,
    root_label: &str,
) -> Result<()> {
    if slot.replace(value).is_some() {
        return Err(OoxmlError::InvalidFormat(format!(
            "{root_label} has multiple header/footer elements"
        )));
    }
    Ok(())
}

fn read_root_element(
    xml: &[u8],
    root_name: &[u8],
    root_label: &str,
) -> Result<(BytesStart<'static>, quick_xml::encoding::Decoder)> {
    let mut reader = NsReader::from_reader(xml);
    let mut depth = 0usize;
    let mut root = None;

    loop {
        let decoder = reader.decoder();
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| OoxmlError::Xml(error.to_string()))?;
        match event {
            Event::Start(element) => {
                depth = depth.checked_add(1).ok_or_else(|| {
                    OoxmlError::InvalidFormat(format!("{root_label} XML nesting is too deep"))
                })?;
                if depth == 1 {
                    if root.is_some() {
                        return Err(OoxmlError::InvalidFormat(format!(
                            "{root_label} XML has multiple roots"
                        )));
                    }
                    require_presentationml_root(&namespace, &element, root_name, root_label)?;
                    root = Some((element.into_owned(), decoder));
                }
            },
            Event::Empty(element) if depth == 0 => {
                if root.is_some() {
                    return Err(OoxmlError::InvalidFormat(format!(
                        "{root_label} XML has multiple roots"
                    )));
                }
                require_presentationml_root(&namespace, &element, root_name, root_label)?;
                root = Some((element.into_owned(), decoder));
            },
            Event::End(_) => {
                depth = depth.checked_sub(1).ok_or_else(|| {
                    OoxmlError::InvalidFormat(format!("invalid {root_label} XML nesting"))
                })?;
            },
            Event::Eof => break,
            _ => {},
        }
    }

    if depth != 0 {
        return Err(OoxmlError::InvalidFormat(format!(
            "unterminated {root_label} XML"
        )));
    }
    root.ok_or_else(|| OoxmlError::InvalidFormat(format!("{root_label} XML has no root element")))
}

fn require_presentationml_root(
    namespace: &quick_xml::name::ResolveResult<'_>,
    element: &BytesStart<'_>,
    root_name: &[u8],
    root_label: &str,
) -> Result<()> {
    if is_presentationml_name(namespace, element.name(), root_name) {
        Ok(())
    } else {
        Err(OoxmlError::InvalidFormat(format!(
            "{root_label} XML must have a PresentationML {root_label} root"
        )))
    }
}

fn parse_boolean_attribute(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: quick_xml::encoding::Decoder,
    root_label: &str,
    default: bool,
) -> Result<bool> {
    let Some(value) = unqualified_attribute_value(element, name, decoder)? else {
        return Ok(default);
    };
    match value.as_str() {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => Err(OoxmlError::InvalidFormat(format!(
            "{root_label} has invalid {} value '{value}'",
            String::from_utf8_lossy(name)
        ))),
    }
}

/// A slide part.
///
/// Corresponds to `/ppt/slides/slideN.xml` in the package.
pub struct SlidePart<'a> {
    /// The underlying OPC part
    part: &'a dyn Part,
    xml: Arc<Vec<u8>>,
}

impl<'a> SlidePart<'a> {
    /// Create a SlidePart from an OPC Part.
    pub fn from_part(part: &'a dyn Part) -> Result<Self> {
        let xml = processed(part)?;
        Ok(Self { part, xml })
    }

    /// Get the XML bytes of the slide.
    #[inline]
    fn xml_bytes(&self) -> &[u8] {
        self.xml.as_slice()
    }

    /// Get the slide name.
    ///
    /// Returns the name attribute from the <p:cSld> element.
    pub fn name(&self) -> Result<String> {
        presentation_name(self.xml_bytes())
    }

    /// Whether this slide is hidden during a slide show.
    ///
    /// The PresentationML show attribute defaults to true, so slides without
    /// an explicit value are not hidden.
    pub fn is_hidden(&self) -> Result<bool> {
        Ok(!parse_slide_show(self.xml_bytes())?)
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
        parse_shapes(self.xml_bytes())
    }

    /// Parse and return all placeholder shapes on this slide.
    pub fn placeholders(&self) -> Result<Vec<BaseShape>> {
        Ok(filter_placeholders(self.shapes()?))
    }

    /// Get the flags controlling whether master content is shown on this slide.
    pub fn master_visibility(&self) -> Result<MasterVisibility> {
        parse_master_visibility(self.xml_bytes(), b"sld", "slide")
    }

    /// Get the color-map override declared by this slide.
    pub fn color_map_override(&self) -> Result<Option<crate::pptx::color_map::ColorMapOverride>> {
        crate::pptx::color_map::parse_color_map_override(self.xml_bytes(), b"sld", "slide")
    }

    /// Get the transition effect for this slide.
    ///
    /// Parses the `<p:transition>` element from the slide XML.
    /// Returns `None` if no transition is defined.
    pub fn transition(&self) -> Result<Option<crate::pptx::transitions::SlideTransition>> {
        crate::pptx::transitions::SlideTransition::from_xml(self.part.blob())
    }

    /// Parse the simple shape-animation metadata in this slide's timing tree.
    pub fn animations(&self) -> Result<crate::pptx::animations::AnimationSequence> {
        crate::pptx::animations::AnimationSequence::parse_slide_xml(self.xml_bytes())
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
    xml: Arc<Vec<u8>>,
}

impl<'a> SlideLayoutPart<'a> {
    /// Create a SlideLayoutPart from an OPC Part.
    pub fn from_part(part: &'a dyn Part) -> Result<Self> {
        let xml = processed(part)?;
        Ok(Self { part, xml })
    }

    /// Get the XML bytes of the layout.
    #[inline]
    fn xml_bytes(&self) -> &[u8] {
        self.xml.as_slice()
    }

    /// Get the layout name.
    pub fn name(&self) -> Result<String> {
        presentation_name(self.xml_bytes())
    }

    /// Get the root-level metadata declared by this slide layout.
    pub fn metadata(&self) -> Result<SlideLayoutMetadata> {
        parse_slide_layout_metadata(self.xml_bytes())
    }

    /// Get local header and footer placeholder visibility for this layout.
    pub fn header_footer(&self) -> Result<Option<SlideHeaderFooterVisibility>> {
        parse_header_footer_visibility(self.xml_bytes(), b"sldLayout", "slide layout")
    }

    /// Get all shapes defined by this layout.
    pub fn shapes(&self) -> Result<Vec<BaseShape>> {
        parse_shapes(self.xml_bytes())
    }

    /// Get all placeholder shapes defined by this layout.
    pub fn placeholders(&self) -> Result<Vec<BaseShape>> {
        Ok(filter_placeholders(self.shapes()?))
    }

    /// Get the flags controlling whether master content is shown on this layout.
    pub fn master_visibility(&self) -> Result<MasterVisibility> {
        parse_master_visibility(self.xml_bytes(), b"sldLayout", "slide layout")
    }

    /// Get the color-map override declared by this layout.
    pub fn color_map_override(&self) -> Result<Option<crate::pptx::color_map::ColorMapOverride>> {
        crate::pptx::color_map::parse_color_map_override(
            self.xml_bytes(),
            b"sldLayout",
            "slide layout",
        )
    }

    /// Parse the timing metadata declared by this slide layout.
    pub fn animations(&self) -> Result<crate::pptx::animations::AnimationSequence> {
        crate::pptx::animations::AnimationSequence::parse_slide_xml(self.xml_bytes())
    }

    /// Get the transition effect inherited from this slide layout.
    ///
    /// Parses the `<p:transition>` element from the layout XML.
    /// Returns `None` if the layout has no transition.
    pub fn transition(&self) -> Result<Option<crate::pptx::transitions::SlideTransition>> {
        crate::pptx::transitions::SlideTransition::from_xml(self.part.blob())
    }

    /// Get the background defined by this slide layout.
    ///
    /// Parses the p:bg element from the layout XML. Returns `None` when the
    /// layout has no local background.
    pub fn background(&self) -> Result<Option<crate::pptx::backgrounds::SlideBackground>> {
        crate::pptx::backgrounds::SlideBackground::from_xml(self.xml_bytes())
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
    xml: Arc<Vec<u8>>,
}

impl<'a> SlideMasterPart<'a> {
    /// Create a SlideMasterPart from an OPC Part.
    pub fn from_part(part: &'a dyn Part) -> Result<Self> {
        let xml = processed(part)?;
        Ok(Self { part, xml })
    }

    /// Get the XML bytes of the master.
    #[inline]
    fn xml_bytes(&self) -> &[u8] {
        self.xml.as_slice()
    }

    /// Get the master name.
    pub fn name(&self) -> Result<String> {
        presentation_name(self.xml_bytes())
    }

    /// Whether this slide master is retained after its dependent slides are removed.
    pub fn is_preserved(&self) -> Result<bool> {
        parse_slide_master_preserve(self.xml_bytes())
    }

    /// Get local header and footer placeholder visibility for this master.
    pub fn header_footer(&self) -> Result<Option<SlideHeaderFooterVisibility>> {
        parse_header_footer_visibility(self.xml_bytes(), b"sldMaster", "slide master")
    }

    /// Get all shapes defined by this master.
    pub fn shapes(&self) -> Result<Vec<BaseShape>> {
        parse_shapes(self.xml_bytes())
    }

    /// Get all placeholder shapes defined by this master.
    pub fn placeholders(&self) -> Result<Vec<BaseShape>> {
        Ok(filter_placeholders(self.shapes()?))
    }

    /// Get the color map defined by this master.
    pub fn color_map(&self) -> Result<crate::pptx::color_map::ColorMap> {
        crate::pptx::color_map::parse_master_color_map(self.xml_bytes())
    }

    /// Parse the timing metadata declared by this slide master.
    pub fn animations(&self) -> Result<crate::pptx::animations::AnimationSequence> {
        crate::pptx::animations::AnimationSequence::parse_slide_xml(self.xml_bytes())
    }

    /// Get the transition effect inherited from this slide master.
    ///
    /// Parses the `<p:transition>` element from the master XML.
    /// Returns `None` if the master has no transition.
    pub fn transition(&self) -> Result<Option<crate::pptx::transitions::SlideTransition>> {
        crate::pptx::transitions::SlideTransition::from_xml(self.part.blob())
    }

    /// Get the background defined by this slide master.
    ///
    /// Parses the p:bg element from the master XML. Returns `None` when the
    /// master has no local background.
    pub fn background(&self) -> Result<Option<crate::pptx::backgrounds::SlideBackground>> {
        crate::pptx::backgrounds::SlideBackground::from_xml(self.xml_bytes())
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
    fn slide_part_exposes_animation_metadata() {
        let xml = format!(
            r#"<p:sld xmlns:p="{P}"><p:cSld><p:spTree><p:sp><p:nvSpPr>
            <p:cNvPr id="3" name="Animated"/></p:nvSpPr></p:sp></p:spTree></p:cSld>
            <p:timing><p:tnLst><p:par><p:cTn><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst>
            <p:childTnLst><p:par><p:cTn><p:stCondLst><p:cond delay="25"/></p:stCondLst>
            <p:childTnLst><p:par><p:cTn presetID="10" presetClass="entr" nodeType="clickEffect" dur="500">
            <p:childTnLst><p:set><p:cBhvr><p:tgtEl><p:spTgt spid="3"/></p:tgtEl></p:cBhvr></p:set>
            </p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>
            </p:tnLst></p:timing></p:sld>"#
        );
        let blob = part("/ppt/slides/slide1.xml", xml);
        let animations = SlidePart::from_part(&blob).unwrap().animations().unwrap();
        assert_eq!(animations.len(), 1);
        assert_eq!(animations.animations[0].shape_id, 3);
        assert_eq!(animations.animations[0].duration, 500);
        assert_eq!(animations.animations[0].delay, 25);
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

    #[test]
    fn master_visibility_defaults_to_true_and_supports_strict_namespaces() {
        let default_slide = format!(r#"<p:sld xmlns:p="{P}"><p:cSld/></p:sld>"#);
        let visibility =
            parse_master_visibility(default_slide.as_bytes(), b"sld", "slide").unwrap();
        assert_eq!(visibility, MasterVisibility::default());
        assert!(visibility.shows_master_shapes());
        assert!(visibility.shows_master_placeholder_animations());

        let strict_layout = r#"<q:sldLayout
            xmlns:q="http://purl.oclc.org/ooxml/presentationml/main"
            showMasterSp="0" showMasterPhAnim="1"><q:cSld/></q:sldLayout>"#;
        let visibility =
            parse_master_visibility(strict_layout.as_bytes(), b"sldLayout", "slide layout")
                .unwrap();
        assert!(!visibility.shows_master_shapes());
        assert!(visibility.shows_master_placeholder_animations());
    }

    #[test]
    fn master_visibility_rejects_invalid_values_and_roots() {
        let invalid_value =
            format!(r#"<p:sld xmlns:p="{P}" showMasterSp="sometimes"><p:cSld/></p:sld>"#);
        assert!(parse_master_visibility(invalid_value.as_bytes(), b"sld", "slide").is_err());

        let wrong_root = format!(r#"<p:sldLayout xmlns:p="{P}"><p:cSld/></p:sldLayout>"#);
        assert!(parse_master_visibility(wrong_root.as_bytes(), b"sld", "slide").is_err());
    }

    #[test]
    fn slide_show_flag_defaults_to_true_and_supports_strict_namespaces() {
        let default_slide = format!(r#"<p:sld xmlns:p="{P}"><p:cSld/></p:sld>"#);
        assert!(parse_slide_show(default_slide.as_bytes()).unwrap());

        let strict_slide = r#"<q:sld xmlns:q="http://purl.oclc.org/ooxml/presentationml/main"
            show="0"><q:cSld/></q:sld>"#;
        assert!(!parse_slide_show(strict_slide.as_bytes()).unwrap());
    }

    #[test]
    fn slide_show_flag_rejects_invalid_values_and_roots() {
        let invalid_value = format!(r#"<p:sld xmlns:p="{P}" show="sometimes"><p:cSld/></p:sld>"#);
        assert!(parse_slide_show(invalid_value.as_bytes()).is_err());

        let wrong_root = format!(r#"<p:sldLayout xmlns:p="{P}"><p:cSld/></p:sldLayout>"#);
        assert!(parse_slide_show(wrong_root.as_bytes()).is_err());
    }

    #[test]
    fn slide_layout_metadata_reports_values_and_defaults() {
        let defined = format!(
            r#"<p:sldLayout xmlns:p="{P}" matchingName="Picture Caption" type="picTx"
                preserve="1" userDrawn="false"><p:cSld/></p:sldLayout>"#
        );
        let metadata = parse_slide_layout_metadata(defined.as_bytes()).unwrap();
        assert_eq!(metadata.matching_name(), "Picture Caption");
        assert_eq!(metadata.layout_type(), "picTx");
        assert!(metadata.is_preserved());
        assert!(!metadata.is_user_drawn());

        let strict_default = r#"<q:sldLayout
            xmlns:q="http://purl.oclc.org/ooxml/presentationml/main"><q:cSld/></q:sldLayout>"#;
        let metadata = parse_slide_layout_metadata(strict_default.as_bytes()).unwrap();
        assert_eq!(metadata, SlideLayoutMetadata::default());
        assert_eq!(metadata.matching_name(), "");
        assert_eq!(metadata.layout_type(), "cust");
        assert!(!metadata.is_preserved());
        assert!(!metadata.is_user_drawn());
    }

    #[test]
    fn slide_layout_metadata_rejects_invalid_boolean_values_and_roots() {
        let invalid_value =
            format!(r#"<p:sldLayout xmlns:p="{P}" preserve="sometimes"><p:cSld/></p:sldLayout>"#);
        assert!(parse_slide_layout_metadata(invalid_value.as_bytes()).is_err());

        let wrong_root = format!(r#"<p:sld xmlns:p="{P}"><p:cSld/></p:sld>"#);
        assert!(parse_slide_layout_metadata(wrong_root.as_bytes()).is_err());
    }

    #[test]
    fn slide_master_preserve_defaults_to_false_and_supports_strict_namespaces() {
        let default_master = format!(r#"<p:sldMaster xmlns:p="{P}"><p:cSld/></p:sldMaster>"#);
        assert!(!parse_slide_master_preserve(default_master.as_bytes()).unwrap());

        let strict_master = r#"<q:sldMaster
            xmlns:q="http://purl.oclc.org/ooxml/presentationml/main"
            preserve="true"><q:cSld/></q:sldMaster>"#;
        assert!(parse_slide_master_preserve(strict_master.as_bytes()).unwrap());
    }

    #[test]
    fn slide_master_preserve_rejects_invalid_values_and_roots() {
        let invalid_value =
            format!(r#"<p:sldMaster xmlns:p="{P}" preserve="sometimes"><p:cSld/></p:sldMaster>"#);
        assert!(parse_slide_master_preserve(invalid_value.as_bytes()).is_err());

        let wrong_root = format!(r#"<p:sld xmlns:p="{P}"><p:cSld/></p:sld>"#);
        assert!(parse_slide_master_preserve(wrong_root.as_bytes()).is_err());
    }

    #[test]
    fn header_footer_visibility_defaults_to_true_and_supports_strict_namespaces() {
        let master = format!(
            r#"<p:sldMaster xmlns:p="{P}"><p:cSld/><p:hf
                dt="0" ftr="true" hdr="0" sldNum="false"/></p:sldMaster>"#
        );
        let visibility =
            parse_header_footer_visibility(master.as_bytes(), b"sldMaster", "slide master")
                .unwrap()
                .unwrap();
        assert!(!visibility.shows_date_time());
        assert!(visibility.shows_footer());
        assert!(!visibility.shows_header());
        assert!(!visibility.shows_slide_number());

        let strict_layout = r#"<q:sldLayout
            xmlns:q="http://purl.oclc.org/ooxml/presentationml/main"><q:cSld/><q:hf/></q:sldLayout>"#;
        let default_visibility =
            parse_header_footer_visibility(strict_layout.as_bytes(), b"sldLayout", "slide layout")
                .unwrap()
                .unwrap();
        assert_eq!(default_visibility, SlideHeaderFooterVisibility::default());
        assert!(default_visibility.shows_date_time());
        assert!(default_visibility.shows_footer());
        assert!(default_visibility.shows_header());
        assert!(default_visibility.shows_slide_number());

        let no_header_footer = format!(r#"<p:sldLayout xmlns:p="{P}"><p:cSld/></p:sldLayout>"#);
        assert_eq!(
            parse_header_footer_visibility(
                no_header_footer.as_bytes(),
                b"sldLayout",
                "slide layout"
            )
            .unwrap(),
            None
        );
    }

    #[test]
    fn header_footer_visibility_rejects_invalid_and_duplicate_elements() {
        let invalid_value =
            format!(r#"<p:sldLayout xmlns:p="{P}"><p:cSld/><p:hf dt="sometimes"/></p:sldLayout>"#);
        assert!(
            parse_header_footer_visibility(invalid_value.as_bytes(), b"sldLayout", "slide layout")
                .is_err()
        );

        let duplicate =
            format!(r#"<p:sldMaster xmlns:p="{P}"><p:cSld/><p:hf/><p:hf/></p:sldMaster>"#);
        assert!(
            parse_header_footer_visibility(duplicate.as_bytes(), b"sldMaster", "slide master")
                .is_err()
        );
    }
}
