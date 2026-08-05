//! Borrowed presentation graph facade.

use litchi_opc::OpcPackage;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};

use crate::parts::{PresentationPart, SlideMasterPart, SlidePart, SlideReference};
use crate::slide::{Key, Slide, SlideLayout, SlideMaster};
use crate::{Error, Result};

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

    /// Resolve all slide masters in relationship-ID order.
    pub fn slide_masters(&self) -> Result<Vec<SlideMaster<'a>>> {
        let mut relationships: Vec<_> = self
            .part
            .part()
            .rels()
            .iter()
            .filter(|relationship| {
                crate::parts::is_relationship_type(
                    relationship.reltype(),
                    rt::SLIDE_MASTER,
                    "slideMaster",
                )
            })
            .collect();
        relationships.sort_unstable_by(|left, right| left.r_id().cmp(right.r_id()));
        let mut masters = Vec::with_capacity(relationships.len());
        for relationship in relationships {
            if relationship.is_external() {
                return Err(Error::Relationship(
                    "slide-master relationship cannot be external".to_string(),
                ));
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
}
