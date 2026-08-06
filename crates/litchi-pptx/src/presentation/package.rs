//! PresentationML package-graph adapters.
//!
//! This layer resolves slides, masters, companion parts, and relationships
//! without changing the borrowed typed facade or interpreting XML payloads.

use litchi_opc::OpcPackage;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};

use crate::parts::{PresentationPart, SlideMasterPart, SlidePart, SlideReference};
use crate::slide::{Key, Slide, SlideLayout, SlideMaster};
use crate::{Error, Result};

use super::codec;
use super::embedded;

pub(super) fn slide_references(part: &PresentationPart<'_>) -> Result<Vec<SlideReference>> {
    part.slide_references()
}

pub(super) fn slide_count(part: &PresentationPart<'_>) -> Result<usize> {
    Ok(part.slide_references()?.len())
}

pub(super) fn slide_size(part: &PresentationPart<'_>) -> Result<(i64, i64)> {
    part.slide_size()
}

pub(super) fn styles(package: &OpcPackage) -> Result<Option<crate::table::style::List>> {
    crate::table::style::load(package)
}

pub(super) fn notes(
    package: &OpcPackage,
    presentation: &PresentationPart<'_>,
) -> Result<Option<crate::notes::Graph>> {
    crate::notes::load(package, presentation.part().partname())
}

pub(super) fn vba(
    package: &OpcPackage,
    presentation: &PresentationPart<'_>,
) -> Result<Option<embedded::vba::Project>> {
    embedded::vba::discover(package, presentation.part().partname())
}

pub(super) fn content_parts(
    package: &OpcPackage,
    presentation: &PresentationPart<'_>,
) -> Result<Vec<embedded::content_parts::ContentPart>> {
    let mut limits = embedded::content_parts::Limits::default();
    let mut content_parts = Vec::new();
    for (slide_index, slide) in slides(package, presentation)?.into_iter().enumerate() {
        content_parts.extend(embedded::content_parts::load_slide(
            package,
            slide_index,
            slide.part().part(),
            &mut limits,
        )?);
    }
    Ok(content_parts)
}

pub(super) fn hyperlinks(
    package: &OpcPackage,
    presentation: &PresentationPart<'_>,
) -> Result<Vec<(usize, crate::hyperlinks::Hyperlink)>> {
    let mut hyperlinks = Vec::new();
    for (slide_index, slide) in slides(package, presentation)?.into_iter().enumerate() {
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
            codec::parse_inline_hyperlinks(slide.part().part().blob())?
                .into_iter()
                .map(|value| (slide_index, value)),
        );
    }
    Ok(hyperlinks)
}

pub(super) fn slide<'a>(
    package: &'a OpcPackage,
    presentation: &PresentationPart<'a>,
    index: usize,
) -> Result<Option<Slide<'a>>> {
    let references = presentation.slide_references()?;
    let Some(reference) = references.get(index) else {
        return Ok(None);
    };
    let part = slide_part(package, presentation, reference)?;
    Ok(Some(Slide::new(package, part)))
}

pub(super) fn find_slide<'a, 'k>(
    package: &'a OpcPackage,
    presentation: &PresentationPart<'a>,
    key: Key<'k>,
) -> Result<Option<Slide<'a>>> {
    match key {
        Key::Index(index) => slide(package, presentation, index),
        Key::Name(name) => {
            let mut found = None;
            let mut matches = 0usize;
            for slide in slides(package, presentation)? {
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

pub(super) fn slides<'a>(
    package: &'a OpcPackage,
    presentation: &PresentationPart<'a>,
) -> Result<Vec<Slide<'a>>> {
    let references = presentation.slide_references()?;
    let mut slides = Vec::with_capacity(references.len());
    for reference in &references {
        slides.push(Slide::new(
            package,
            slide_part(package, presentation, reference)?,
        ));
    }
    Ok(slides)
}

pub(super) fn slide_masters<'a>(
    package: &'a OpcPackage,
    presentation: &PresentationPart<'a>,
) -> Result<Vec<SlideMaster<'a>>> {
    let relationship_ids = presentation.slide_master_references()?;
    let mut masters = Vec::with_capacity(relationship_ids.len());
    for relationship_id in relationship_ids {
        let relationship = presentation
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
        let part = package.get_part(&target)?;
        if part.content_type() != ct::PML_SLIDE_MASTER {
            return Err(Error::ContentType {
                expected: ct::PML_SLIDE_MASTER.to_string(),
                actual: part.content_type().to_string(),
            });
        }
        masters.push(SlideMaster::new(package, SlideMasterPart::from_part(part)?));
    }
    Ok(masters)
}

pub(super) fn slide_layouts<'a>(
    package: &'a OpcPackage,
    presentation: &PresentationPart<'a>,
) -> Result<Vec<SlideLayout<'a>>> {
    let mut layouts = Vec::new();
    for master in slide_masters(package, presentation)? {
        layouts.extend(master.layouts()?);
    }
    Ok(layouts)
}

pub(super) fn text<'a>(
    package: &'a OpcPackage,
    presentation: &PresentationPart<'a>,
) -> Result<String> {
    let mut text = String::new();
    for slide in slides(package, presentation)? {
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

pub(super) fn slide_sync(
    package: &OpcPackage,
) -> Result<Vec<crate::presentation_properties::metadata::slide_sync::Part>> {
    crate::presentation_properties::metadata::slide_sync::load(package)
}

fn slide_part<'a>(
    package: &'a OpcPackage,
    presentation: &PresentationPart<'a>,
    reference: &SlideReference,
) -> Result<SlidePart<'a>> {
    let part = crate::parts::find_related_part(
        package,
        presentation.part(),
        reference.relationship_id(),
        rt::SLIDE,
        "slide",
        ct::PML_SLIDE,
    )?;
    SlidePart::from_part(part)
}
