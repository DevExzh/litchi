//! Borrowed presentation graph facade.

use litchi_opc::OpcPackage;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use quick_xml::events::Event;
use quick_xml::reader::NsReader;

use crate::parts::{PresentationPart, SlideMasterPart, SlidePart, SlideReference};
use crate::slide::{Key, Slide, SlideLayout, SlideMaster};
use crate::{Error, Result};

/// Layered semantic services that attach embedded resources to a
/// PresentationML graph without interpreting executable payloads.
#[path = "presentation/embedded/mod.rs"]
pub mod embedded;

/// Move-first authoring values for slide audio and video.
#[path = "presentation/media.rs"]
pub mod media;

/// Borrowed semantic view of one PresentationML package graph.
pub struct Presentation<'a> {
    package: &'a OpcPackage,
    part: PresentationPart<'a>,
}

impl<'a> Presentation<'a> {
    /// Construct a view from a validated main part and its package.
    pub fn new(part: PresentationPart<'a>, package: &'a OpcPackage) -> Self {
        Self { package, part }
    }

    /// The underlying OPC package.
    #[inline]
    pub fn package(&self) -> &'a OpcPackage {
        self.package
    }

    /// The borrowed main-document part view.
    #[inline]
    pub fn part(&self) -> &PresentationPart<'a> {
        &self.part
    }

    /// The ordered, low-level slide references.
    pub fn slide_references(&self) -> Result<Vec<SlideReference>> {
        self.part.slide_references()
    }

    /// Number of slides in the ordered presentation graph.
    pub fn slide_count(&self) -> Result<usize> {
        Ok(self.part.slide_references()?.len())
    }

    /// Presentation slide size in EMUs.
    pub fn slide_size(&self) -> Result<(i64, i64)> {
        self.part.slide_size()
    }

    /// Load the optional DrawingML table-style catalog owned by this
    /// presentation. The catalog remains a detached value so callers can
    /// inspect it without holding a mutable package borrow.
    pub fn styles(&self) -> Result<Option<crate::table::style::List>> {
        crate::table::style::load(self.package)
    }

    /// Load the complete inert speaker-notes graph, when present.
    pub fn notes(&self) -> Result<Option<crate::notes::Graph>> {
        crate::notes::load(self.package, self.part.part().partname())
    }

    /// Discover the optional opaque VBA project relationship owned by this
    /// presentation. The binary payload is never decoded or executed.
    pub fn vba(&self) -> Result<Option<embedded::vba::Project>> {
        embedded::vba::discover(self.package, self.part.part().partname())
    }

    /// Discover inert hyperlinks owned by the presentation's slides.
    ///
    /// Each result contains the zero-based slide position and a typed target.
    /// Relationship targets and inline actions are parsed as values only;
    /// they are never followed, opened, or executed.
    pub fn hyperlinks(&self) -> Result<Vec<(usize, crate::hyperlinks::Hyperlink)>> {
        let mut hyperlinks = Vec::new();
        for (slide_index, slide) in self.slides()?.into_iter().enumerate() {
            for relationship in slide.part().part().rels().iter().filter(|relationship| {
                matches!(relationship.reltype(), rt::HYPERLINK | rt::STRICT_HYPERLINK)
            }) {
                let target = relationship.target_ref();
                if target.is_empty() {
                    return Err(Error::Invalid(format!(
                        "hyperlink relationship '{}' on slide {slide_index} has an empty target",
                        relationship.r_id()
                    )));
                }
                hyperlinks.push((
                    slide_index,
                    crate::hyperlinks::Hyperlink::from_xml(target, None)?,
                ));
            }
            hyperlinks.extend(
                Self::parse_inline_hyperlinks(slide.part().part().blob())?
                    .into_iter()
                    .map(|value| (slide_index, value)),
            );
        }
        Ok(hyperlinks)
    }

    /// Resolve one ordered slide by zero-based index.
    pub fn slide(&self, index: usize) -> Result<Option<Slide<'a>>> {
        let references = self.part.slide_references()?;
        let Some(reference) = references.get(index) else {
            return Ok(None);
        };
        let part = self.slide_part(reference)?;
        Ok(Some(Slide::new(self.package, part)))
    }

    /// Resolve a slide by checked index or exact producer-visible name.
    pub fn find_slide<'k>(&self, key: impl Into<Key<'k>>) -> Result<Option<Slide<'a>>> {
        match key.into() {
            Key::Index(index) => self.slide(index),
            Key::Name(name) => {
                let mut found = None;
                let mut matches = 0usize;
                for slide in self.slides()? {
                    if slide.name()? == name {
                        matches = matches.saturating_add(1);
                        found = Some(slide);
                    }
                }
                if matches > 1 {
                    return Err(Error::AmbiguousSlideName {
                        name: name.to_string(),
                        matches,
                    });
                }
                Ok(found)
            },
        }
    }

    /// Resolve all slides in presentation order.
    pub fn slides(&self) -> Result<Vec<Slide<'a>>> {
        let references = self.part.slide_references()?;
        let mut slides = Vec::with_capacity(references.len());
        for reference in &references {
            slides.push(Slide::new(self.package, self.slide_part(reference)?));
        }
        Ok(slides)
    }

    /// Resolve the slide masters declared by `p:sldMasterIdLst` in XML order.
    pub fn slide_masters(&self) -> Result<Vec<SlideMaster<'a>>> {
        let relationship_ids = self.part.slide_master_references()?;
        let mut masters = Vec::with_capacity(relationship_ids.len());
        for relationship_id in relationship_ids {
            let relationship = self
                .part
                .part()
                .rels()
                .get(&relationship_id)
                .ok_or_else(|| {
                    Error::Relationship(format!(
                        "presentation master reference is missing relationship '{relationship_id}'"
                    ))
                })?;
            if relationship.is_external() {
                return Err(Error::Relationship(
                    "slide-master relationship must be internal".to_string(),
                ));
            }
            if !crate::parts::is_relationship_type(
                relationship.reltype(),
                rt::SLIDE_MASTER,
                "slideMaster",
            ) {
                return Err(Error::Relationship(format!(
                    "relationship '{relationship_id}' is not a slide-master relationship"
                )));
            }
            let target = relationship.target_partname()?;
            let part = self.package.get_part(&target)?;
            if part.content_type() != ct::PML_SLIDE_MASTER {
                return Err(Error::ContentType {
                    expected: ct::PML_SLIDE_MASTER.to_string(),
                    actual: part.content_type().to_string(),
                });
            }
            masters.push(SlideMaster::new(
                self.package,
                SlideMasterPart::from_part(part)?,
            ));
        }
        Ok(masters)
    }

    /// Resolve all layouts reachable from all presentation masters.
    pub fn slide_layouts(&self) -> Result<Vec<SlideLayout<'a>>> {
        let mut layouts = Vec::new();
        for master in self.slide_masters()? {
            layouts.extend(master.layouts()?);
        }
        Ok(layouts)
    }

    /// Flatten all slide text in presentation order.
    pub fn text(&self) -> Result<String> {
        let mut text = String::new();
        for slide in self.slides()? {
            let value = slide.text()?;
            if !value.is_empty() {
                if !text.is_empty() {
                    text.push('\n');
                }
                text.push_str(&value);
            }
        }
        Ok(text)
    }

    /// Load the slide-library synchronization metadata reachable from this
    /// presentation's slide graph.
    pub fn slide_sync(
        &self,
    ) -> Result<Vec<crate::presentation_properties::metadata::slide_sync::Part>> {
        crate::presentation_properties::metadata::slide_sync::load(self.package)
    }

    fn slide_part(&self, reference: &SlideReference) -> Result<SlidePart<'a>> {
        let part = crate::parts::find_related_part(
            self.package,
            self.part.part(),
            reference.relationship_id(),
            rt::SLIDE,
            "slide",
            ct::PML_SLIDE,
        )?;
        SlidePart::from_part(part)
    }

    fn parse_inline_hyperlinks(xml: &[u8]) -> Result<Vec<crate::hyperlinks::Hyperlink>> {
        let processed = litchi_ooxml_common::mce::process_ooxml(xml)?;
        let mut reader = NsReader::from_reader(processed.as_ref());
        reader.config_mut().trim_text(true);
        let mut hyperlinks = Vec::new();
        loop {
            let decoder = reader.decoder();
            let (namespace, event) = reader.read_resolved_event()?;
            match event {
                Event::Start(element) | Event::Empty(element)
                    if litchi_ooxml_common::xml::is_drawingml_name(
                        &namespace,
                        element.name(),
                        b"hlinkClick",
                    ) =>
                {
                    let action = litchi_ooxml_common::xml::unqualified_attribute_value(
                        &element, b"action", decoder,
                    )?;
                    let tooltip = litchi_ooxml_common::xml::unqualified_attribute_value(
                        &element, b"tooltip", decoder,
                    )?;
                    if let Some(action) = action {
                        if action.is_empty() {
                            return Err(Error::Invalid(
                                "inline hyperlink action cannot be empty".into(),
                            ));
                        }
                        hyperlinks.push(crate::hyperlinks::Hyperlink::from_xml(&action, tooltip)?);
                    }
                },
                Event::Eof => break,
                _ => {},
            }
        }
        Ok(hyperlinks)
    }
}
