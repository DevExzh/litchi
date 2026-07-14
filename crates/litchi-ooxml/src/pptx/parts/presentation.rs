/// Presentation part - the main part in a .pptx package.
///
/// Corresponds to `/ppt/presentation.xml` in the package.
use crate::common::xml::unqualified_attribute_value;
use crate::error::{OoxmlError, Result};
use crate::pptx::namespace::{is_presentationml_name, relationship_attribute_value};
use litchi_opc::part::Part;
use quick_xml::encoding::Decoder;
use quick_xml::events::BytesStart;
use quick_xml::events::Event;
use quick_xml::reader::NsReader;

/// The main presentation part.
///
/// This part contains the presentation-level properties and references to slides,
/// slide masters, and other presentation resources.
///
/// # Example
///
/// ```rust,ignore
/// let pres_part = PresentationPart::from_part(opc_part)?;
/// let slide_count = pres_part.slide_count()?;
/// ```
pub struct PresentationPart<'a> {
    /// The underlying OPC part
    part: &'a dyn Part,
}

impl<'a> PresentationPart<'a> {
    /// Create a PresentationPart from an OPC Part.
    ///
    /// # Arguments
    ///
    /// * `part` - The underlying OPC part
    ///
    /// # Example
    ///
    /// ```rust,ignore
    /// let pres_part = PresentationPart::from_part(opc_part)?;
    /// ```
    pub fn from_part(part: &'a dyn Part) -> Result<Self> {
        Ok(Self { part })
    }

    /// Get the XML bytes of the presentation.
    #[inline]
    fn xml_bytes(&self) -> &[u8] {
        self.part.blob()
    }

    /// Get the number of slides in the presentation.
    ///
    /// This counts the `<p:sldId>` elements in the presentation.
    ///
    /// # Example
    ///
    /// ```rust,ignore
    /// let count = pres_part.slide_count()?;
    /// println!("Presentation has {} slides", count);
    /// ```
    pub fn slide_count(&self) -> Result<usize> {
        Ok(PresentationInfo::parse(self.xml_bytes())?.slides.len())
    }

    /// Get the slide width in EMUs (English Metric Units).
    ///
    /// Returns None if the slide size is not defined.
    ///
    /// # Example
    ///
    /// ```rust,ignore
    /// if let Some(width) = pres_part.slide_width()? {
    ///     println!("Slide width: {} EMUs", width);
    /// }
    /// ```
    pub fn slide_width(&self) -> Result<Option<i64>> {
        Ok(PresentationInfo::parse(self.xml_bytes())?
            .slide_size
            .map(|(width, _)| width))
    }

    /// Get the slide height in EMUs (English Metric Units).
    ///
    /// Returns None if the slide size is not defined.
    ///
    /// # Example
    ///
    /// ```rust,ignore
    /// if let Some(height) = pres_part.slide_height()? {
    ///     println!("Slide height: {} EMUs", height);
    /// }
    /// ```
    pub fn slide_height(&self) -> Result<Option<i64>> {
        Ok(PresentationInfo::parse(self.xml_bytes())?
            .slide_size
            .map(|(_, height)| height))
    }

    /// Get the relationship IDs of all slides in presentation order.
    ///
    /// Returns a vector of relationship IDs that can be used to access
    /// the actual slide parts.
    ///
    /// # Example
    ///
    /// ```rust,ignore
    /// let slide_rids = pres_part.slide_rids()?;
    /// for rid in slide_rids {
    ///     // Use rid to get slide part
    /// }
    /// ```
    pub fn slide_rids(&self) -> Result<Vec<String>> {
        Ok(PresentationInfo::parse(self.xml_bytes())?
            .slides
            .into_iter()
            .map(|(_, relationship_id)| relationship_id)
            .collect())
    }

    /// Get the relationship IDs of all slide masters.
    ///
    /// # Example
    ///
    /// ```rust,ignore
    /// let master_rids = pres_part.slide_master_rids()?;
    /// ```
    pub fn slide_master_rids(&self) -> Result<Vec<String>> {
        Ok(PresentationInfo::parse(self.xml_bytes())?
            .masters
            .into_iter()
            .map(|(_, relationship_id)| relationship_id)
            .collect())
    }

    /// Get the underlying OPC part.
    #[inline]
    pub fn part(&self) -> &'a dyn Part {
        self.part
    }
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum PresentationContext {
    Presentation,
    SlideList,
    MasterList,
    Other,
}

#[derive(Default)]
struct PresentationInfo {
    slides: Vec<(u32, String)>,
    masters: Vec<(u32, String)>,
    slide_size: Option<(i64, i64)>,
    seen_slide_list: bool,
    seen_master_list: bool,
}

impl PresentationInfo {
    fn parse(xml: &[u8]) -> Result<Self> {
        let mut reader = NsReader::from_reader(xml);
        let mut info = Self::default();
        let mut stack = Vec::new();
        let mut closed_root = false;

        loop {
            let decoder = reader.decoder();
            let event = reader
                .read_event()
                .map_err(|error| OoxmlError::Xml(error.to_string()))?
                .into_owned();
            let resolver = reader.resolver().clone();
            let (namespace, event) = resolver.resolve_event(event);
            match event {
                Event::Start(element) => {
                    if stack.is_empty() {
                        if closed_root
                            || !is_presentationml_name(&namespace, element.name(), b"presentation")
                        {
                            return Err(OoxmlError::InvalidFormat(
                                "presentation XML must have one PresentationML presentation root"
                                    .to_string(),
                            ));
                        }
                        stack.push(PresentationContext::Presentation);
                        continue;
                    }
                    info.process_element(
                        *stack.last().ok_or_else(|| {
                            OoxmlError::InvalidFormat(
                                "missing PowerPoint presentation context".to_string(),
                            )
                        })?,
                        &namespace,
                        &element,
                        decoder,
                        &resolver,
                    )?;
                    stack.push(info.child_context(
                        *stack.last().ok_or_else(|| {
                            OoxmlError::InvalidFormat(
                                "missing PowerPoint presentation context".to_string(),
                            )
                        })?,
                        &namespace,
                        &element,
                    )?);
                },
                Event::Empty(element) => {
                    let parent = *stack.last().ok_or_else(|| {
                        OoxmlError::InvalidFormat(
                            "presentation XML has an empty or missing root".to_string(),
                        )
                    })?;
                    info.process_element(parent, &namespace, &element, decoder, &resolver)?;
                    info.observe_empty_container(parent, &namespace, &element)?;
                },
                Event::End(element) => {
                    let context = stack.pop().ok_or_else(|| {
                        OoxmlError::InvalidFormat(
                            "invalid PowerPoint presentation nesting".to_string(),
                        )
                    })?;
                    if stack.is_empty() {
                        if context != PresentationContext::Presentation
                            || !is_presentationml_name(&namespace, element.name(), b"presentation")
                        {
                            return Err(OoxmlError::InvalidFormat(
                                "invalid PowerPoint presentation root closure".to_string(),
                            ));
                        }
                        closed_root = true;
                    }
                },
                Event::Eof if !closed_root || !stack.is_empty() => {
                    return Err(OoxmlError::InvalidFormat(
                        "unterminated PowerPoint presentation XML".to_string(),
                    ));
                },
                Event::Eof => break,
                _ => {},
            }
        }
        Ok(info)
    }

    fn child_context(
        &mut self,
        parent: PresentationContext,
        namespace: &quick_xml::name::ResolveResult<'_>,
        element: &BytesStart<'_>,
    ) -> Result<PresentationContext> {
        if parent == PresentationContext::Presentation
            && is_presentationml_name(namespace, element.name(), b"sldIdLst")
        {
            if self.seen_slide_list {
                return Err(OoxmlError::InvalidFormat(
                    "duplicate PowerPoint slide ID list".to_string(),
                ));
            }
            self.seen_slide_list = true;
            Ok(PresentationContext::SlideList)
        } else if parent == PresentationContext::Presentation
            && is_presentationml_name(namespace, element.name(), b"sldMasterIdLst")
        {
            if self.seen_master_list {
                return Err(OoxmlError::InvalidFormat(
                    "duplicate PowerPoint slide-master ID list".to_string(),
                ));
            }
            self.seen_master_list = true;
            Ok(PresentationContext::MasterList)
        } else {
            Ok(PresentationContext::Other)
        }
    }

    fn observe_empty_container(
        &mut self,
        parent: PresentationContext,
        namespace: &quick_xml::name::ResolveResult<'_>,
        element: &BytesStart<'_>,
    ) -> Result<()> {
        if parent == PresentationContext::Presentation
            && is_presentationml_name(namespace, element.name(), b"sldIdLst")
        {
            if self.seen_slide_list {
                return Err(OoxmlError::InvalidFormat(
                    "duplicate PowerPoint slide ID list".to_string(),
                ));
            }
            self.seen_slide_list = true;
        } else if parent == PresentationContext::Presentation
            && is_presentationml_name(namespace, element.name(), b"sldMasterIdLst")
        {
            if self.seen_master_list {
                return Err(OoxmlError::InvalidFormat(
                    "duplicate PowerPoint slide-master ID list".to_string(),
                ));
            }
            self.seen_master_list = true;
        }
        Ok(())
    }

    fn process_element(
        &mut self,
        parent: PresentationContext,
        namespace: &quick_xml::name::ResolveResult<'_>,
        element: &BytesStart<'_>,
        decoder: Decoder,
        resolver: &quick_xml::name::NamespaceResolver,
    ) -> Result<()> {
        if parent == PresentationContext::Presentation
            && is_presentationml_name(namespace, element.name(), b"sldSz")
        {
            if self.slide_size.is_some() {
                return Err(OoxmlError::InvalidFormat(
                    "duplicate PowerPoint slide size".to_string(),
                ));
            }
            let width = required_positive_i64(element, b"cx", decoder, "slide width")?;
            let height = required_positive_i64(element, b"cy", decoder, "slide height")?;
            self.slide_size = Some((width, height));
        } else if parent == PresentationContext::SlideList
            && is_presentationml_name(namespace, element.name(), b"sldId")
        {
            let id = required_u32(element, b"id", decoder, "slide ID")?;
            if id < 256 {
                return Err(OoxmlError::InvalidFormat(format!(
                    "PowerPoint slide ID {id} is below 256"
                )));
            }
            let relationship_id = required_relationship_id(element, decoder, resolver, "slide")?;
            push_unique_reference(&mut self.slides, id, relationship_id, "slide")?;
        } else if parent == PresentationContext::MasterList
            && is_presentationml_name(namespace, element.name(), b"sldMasterId")
        {
            let id = required_u32(element, b"id", decoder, "slide-master ID")?;
            if id < 2_147_483_648 {
                return Err(OoxmlError::InvalidFormat(format!(
                    "PowerPoint slide-master ID {id} is below 2147483648"
                )));
            }
            let relationship_id =
                required_relationship_id(element, decoder, resolver, "slide master")?;
            push_unique_reference(&mut self.masters, id, relationship_id, "slide master")?;
        }
        Ok(())
    }
}

fn required_relationship_id(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &quick_xml::name::NamespaceResolver,
    description: &str,
) -> Result<String> {
    let value =
        relationship_attribute_value(element, b"id", decoder, resolver)?.ok_or_else(|| {
            OoxmlError::InvalidFormat(format!("PowerPoint {description} is missing r:id"))
        })?;
    if value.is_empty() {
        return Err(OoxmlError::InvalidFormat(format!(
            "PowerPoint {description} has an empty relationship ID"
        )));
    }
    Ok(value)
}

fn required_u32(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> Result<u32> {
    let value = unqualified_attribute_value(element, name, decoder)?
        .ok_or_else(|| OoxmlError::InvalidFormat(format!("missing {description} attribute")))?;
    value
        .parse::<u32>()
        .map_err(|_| OoxmlError::InvalidFormat(format!("invalid {description} value '{value}'")))
}

fn required_positive_i64(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> Result<i64> {
    let value = unqualified_attribute_value(element, name, decoder)?
        .ok_or_else(|| OoxmlError::InvalidFormat(format!("missing {description} attribute")))?;
    let parsed = value
        .parse::<i64>()
        .map_err(|_| OoxmlError::InvalidFormat(format!("invalid {description} value '{value}'")))?;
    if parsed <= 0 {
        return Err(OoxmlError::InvalidFormat(format!(
            "{description} must be positive"
        )));
    }
    Ok(parsed)
}

fn push_unique_reference(
    references: &mut Vec<(u32, String)>,
    id: u32,
    relationship_id: String,
    description: &str,
) -> Result<()> {
    if references
        .iter()
        .any(|(existing_id, existing_rid)| *existing_id == id || *existing_rid == relationship_id)
    {
        return Err(OoxmlError::InvalidFormat(format!(
            "duplicate PowerPoint {description} ID or relationship"
        )));
    }
    references.push((id, relationship_id));
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::packuri::PackURI;
    use litchi_opc::part::BlobPart;

    const P: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
    const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    fn part(xml: impl Into<Vec<u8>>) -> BlobPart {
        BlobPart::new(
            PackURI::new("/ppt/presentation.xml").unwrap(),
            "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
                .to_string(),
            xml.into(),
        )
    }

    #[test]
    fn parses_ordered_references_and_dimensions_by_namespace() {
        let xml = format!(
            r#"<q:presentation xmlns:q="{P}" xmlns:rel="{R}" xmlns:f="urn:foreign">
                <q:sldMasterIdLst><f:sldMasterId id="4294967295" rel:id="spoof"/>
                    <q:sldMasterId id="2147483648" f:id="7" rel:id="master-alpha"/></q:sldMasterIdLst>
                <q:sldIdLst><q:sldId id="256" f:id="1" rel:id="slide-alpha"/>
                    <q:sldId id="257" rel:id="slide-beta"/><f:sldId id="258" rel:id="spoof"/></q:sldIdLst>
                <f:sldSz cx="1" cy="1"/><q:sldSz cx="9144000" cy="5143500"/>
            </q:presentation>"#
        );
        let blob = part(xml);
        let presentation = PresentationPart::from_part(&blob).unwrap();
        assert_eq!(presentation.slide_count().unwrap(), 2);
        assert_eq!(
            presentation.slide_rids().unwrap(),
            ["slide-alpha", "slide-beta"]
        );
        assert_eq!(presentation.slide_master_rids().unwrap(), ["master-alpha"]);
        assert_eq!(presentation.slide_width().unwrap(), Some(9_144_000));
        assert_eq!(presentation.slide_height().unwrap(), Some(5_143_500));
    }

    #[test]
    fn accepts_strict_relationship_aliases() {
        let xml = r#"<x:presentation xmlns:x="http://purl.oclc.org/ooxml/presentationml/main"
            xmlns:z="http://purl.oclc.org/ooxml/officeDocument/relationships">
            <x:sldMasterIdLst><x:sldMasterId id="2147483648" z:id="m"/></x:sldMasterIdLst>
            <x:sldIdLst><x:sldId id="256" z:id="s"/></x:sldIdLst>
            <x:sldSz cx="1" cy="2"/></x:presentation>"#;
        let blob = part(xml);
        let presentation = PresentationPart::from_part(&blob).unwrap();
        assert_eq!(presentation.slide_rids().unwrap(), ["s"]);
        assert_eq!(presentation.slide_master_rids().unwrap(), ["m"]);
    }

    #[test]
    fn ignores_nested_and_foreign_reference_lookalikes() {
        let xml = format!(
            r#"<p:presentation xmlns:p="{P}" xmlns:r="{R}" xmlns:f="urn:foreign">
                <p:sldIdLst><f:wrapper><p:sldId id="256" r:id="nested"/></f:wrapper>
                    <p:sldId id="257" r:id="real"/></p:sldIdLst>
                <p:extLst><p:sldIdLst><p:sldId id="258" r:id="extension"/></p:sldIdLst></p:extLst>
            </p:presentation>"#
        );
        let blob = part(xml);
        let presentation = PresentationPart::from_part(&blob).unwrap();
        assert_eq!(presentation.slide_rids().unwrap(), ["real"]);
    }

    #[test]
    fn rejects_malformed_presentation_metadata() {
        let invalid = [
            format!(
                r#"<p:presentation xmlns:p="{P}"><p:sldIdLst><p:sldId id="255"/></p:sldIdLst></p:presentation>"#
            ),
            format!(
                r#"<p:presentation xmlns:p="{P}" xmlns:r="{R}"><p:sldIdLst><p:sldId id="256" r:id=""/></p:sldIdLst></p:presentation>"#
            ),
            format!(
                r#"<p:presentation xmlns:p="{P}" xmlns:r="{R}"><p:sldIdLst><p:sldId id="256" r:id="a"/><p:sldId id="256" r:id="b"/></p:sldIdLst></p:presentation>"#
            ),
            format!(r#"<p:presentation xmlns:p="{P}"><p:sldSz cx="0" cy="1"/></p:presentation>"#),
            format!(r#"<p:presentation xmlns:p="{P}"><p:sldSz cx="1"/></p:presentation>"#),
            format!(r#"<p:presentation xmlns:p="{P}"><p:sldIdLst/><p:sldIdLst/></p:presentation>"#),
            format!(r#"<p:presentation xmlns:p="{P}"><p:sldSz cx="1" cy="2"/>"#),
        ];
        for xml in invalid {
            let blob = part(xml);
            assert!(
                PresentationPart::from_part(&blob)
                    .unwrap()
                    .slide_count()
                    .is_err()
            );
        }
    }
}
