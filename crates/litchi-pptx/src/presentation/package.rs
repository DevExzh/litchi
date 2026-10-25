//! `PresentationML` package-graph adapters.
//!
//! This layer resolves slides, masters, companion parts, and relationships
//! without changing the borrowed typed facade or interpreting XML payloads.

use litchi_core::{
    SequentialTextWriter, TextObjectKind, TextOutputError, TextOutputOptions, TextOutputReport,
};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{OpcPackage, SourceLimits};
use std::io::Write;

use crate::notes::SlideRootProof;
use crate::parts::{MceCapture, PresentationPart, SlideMasterPart, SlidePart, SlideReference};
use crate::slide::{Selector, Slide, SlideLayout, SlideMaster};
use crate::{Error, Result};

use super::Presentation;
use super::codec;
use super::embedded;

pub(crate) struct CaptureSlide<'a> {
    pub(crate) slide: Slide<'a>,
    pub(crate) name: Option<String>,
}

pub(crate) struct CaptureSlides<'a> {
    pub(crate) slides: Vec<CaptureSlide<'a>>,
    pub(crate) first_name_error: Option<(usize, Error)>,
    pub(crate) slide_root_proofs: Option<Vec<SlideRootProof<'a>>>,
}

/// Validate and wrap a slide after an earlier producer-name projection has
/// failed. The caller deliberately keeps this root-only fallback so the first
/// name error retains its deferred ordering, but physical source admission
/// must still precede the MCE parse.
pub(crate) fn root_only_slide_with_source_limits<'a>(
    part: &'a dyn litchi_opc::Part,
    source_limits: Option<SourceLimits>,
) -> Result<SlidePart<'a>> {
    if let Some(source_limits) = source_limits {
        let source = crate::parts::raw_xml(part)?;
        crate::parts::validate_source_part_xml(source, source_limits)?;
    }
    SlidePart::from_part(part)
}

pub(super) fn validate_slide_catalog(
    package: &OpcPackage,
    presentation: &PresentationPart<'_>,
    references: &[SlideReference],
) -> Result<()> {
    for reference in references {
        crate::parts::validate_slide_relationship(
            presentation.part().rels().get(reference.relationship_id()),
            reference.relationship_id(),
            |target| Ok(package.get_part(target)?),
            |part| part.content_type(),
        )?;
    }
    Ok(())
}

#[cfg(test)]
pub(crate) fn capture_slides<'a>(presentation: &Presentation<'a>) -> Result<CaptureSlides<'a>> {
    capture_slides_inner(presentation, None, None, None)
}

/// Capture every slide, reusing MCE projections the parent snapshot retained
/// and notes-root classifications `proved_roots` holds for the exact payload
/// allocations this capture reads.
pub(crate) fn capture_slides_with_mce<'a, 'parent>(
    presentation: &Presentation<'a>,
    capture: &mut MceCapture<'a, 'parent>,
    proved_roots: Option<&crate::notes::SlideRootMemo>,
    source_limits: SourceLimits,
) -> Result<CaptureSlides<'a>> {
    capture_slides_inner(
        presentation,
        Some(capture),
        proved_roots,
        Some(source_limits),
    )
}

fn capture_slides_inner<'a, 'parent>(
    presentation: &Presentation<'a>,
    mut capture: Option<&mut MceCapture<'a, 'parent>>,
    proved_roots: Option<&crate::notes::SlideRootMemo>,
    source_limits: Option<SourceLimits>,
) -> Result<CaptureSlides<'a>> {
    let package = presentation.package();
    let references = presentation.catalog()?;
    let mut slides = Vec::new();
    slides
        .try_reserve_exact(references.len())
        .map_err(|source| Error::Allocation {
            resource: "opened-presentation capture slides",
            source,
        })?;
    let mut proofs = Vec::new();
    let mut collect_notes_proofs = proofs.try_reserve_exact(references.len()).is_ok();
    let mut slide_root_proofs = collect_notes_proofs.then_some(proofs);
    let mut first_name_error = None;

    for (index, reference) in references.iter().enumerate() {
        let (_, _, part) = crate::parts::validate_slide_relationship(
            presentation
                .part()
                .part()
                .rels()
                .get(reference.relationship_id()),
            reference.relationship_id(),
            |target| Ok(package.get_part(target)?),
            |part| part.content_type(),
        )?;
        let (slide_part, name) = if first_name_error.is_none() {
            let (slide_part, name, notes_proof) = match capture.as_deref_mut() {
                Some(capture) => {
                    let Some(source_limits) = source_limits else {
                        return Err(Error::Invalid(
                            "captured slides lack source limits".to_owned(),
                        ));
                    };
                    SlidePart::from_part_with_name_with_capture(
                        part,
                        collect_notes_proofs,
                        capture,
                        proved_roots,
                        source_limits,
                    )?
                },
                None => SlidePart::from_part_with_name(part, collect_notes_proofs)?,
            };
            if let Some(proof) = notes_proof {
                if let Some(proofs) = slide_root_proofs.as_mut() {
                    proofs.push(proof);
                }
                if !proof.is_valid() {
                    collect_notes_proofs = false;
                }
            }
            let name = match name {
                Ok(name) => Some(name),
                Err(error) => {
                    first_name_error = Some((index, error));
                    None
                },
            };
            (slide_part, name)
        } else {
            (
                root_only_slide_with_source_limits(part, source_limits)?,
                None,
            )
        };
        slides.push(CaptureSlide {
            slide: Slide::new_with_shape_and_semantic_limits(
                package,
                slide_part,
                presentation.shape_limits(),
                presentation.semantic_limits.clone(),
            ),
            name,
        });
    }

    Ok(CaptureSlides {
        slides,
        first_name_error,
        slide_root_proofs,
    })
}

pub(super) fn write_text_to<W: Write + ?Sized>(
    presentation: &Presentation<'_>,
    output: &mut W,
    options: TextOutputOptions<'_>,
) -> std::result::Result<TextOutputReport, TextOutputError<Error>> {
    let package = presentation.package();
    let paragraph_separator = options.paragraph_separator();
    let mut writer = SequentialTextWriter::new(output, options);
    // This preflight catalog is bounded relationship metadata only. It never
    // contains slide XML or decoded slide text; the payload remains one slide
    // at a time below.
    let references = match presentation.catalog() {
        Ok(references) => references,
        Err(source) => return Err(writer.document_error(source)),
    };
    if let Err(source) = validate_slide_catalog(package, presentation.part(), references) {
        return Err(writer.document_error(source));
    }

    for reference in references {
        let (_, _, part) = match crate::parts::validate_slide_relationship(
            presentation
                .part()
                .part()
                .rels()
                .get(reference.relationship_id()),
            reference.relationship_id(),
            |target| Ok(package.get_part(target)?),
            |part| part.content_type(),
        ) {
            Ok(value) => value,
            Err(source) => return Err(writer.document_error(source)),
        };
        let text = match SlidePart::semantic_text_from_part(part, paragraph_separator) {
            Ok(text) => text,
            Err(source) => return Err(writer.document_error(source)),
        };
        writer.write_object::<Error>(TextObjectKind::Slide, &text)?;
    }

    Ok(writer.finish())
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
    limits: crate::notes::Limits,
) -> Result<Option<crate::notes::Graph>> {
    crate::notes::load_with_limits(package, presentation.part().partname(), limits)
}

#[cfg(feature = "vba-inspection")]
pub(super) fn vba(
    package: &OpcPackage,
    presentation: &PresentationPart<'_>,
) -> Result<Option<embedded::vba::Project>> {
    embedded::vba::discover(package, presentation.part().partname())
}

pub(super) fn content_parts(
    presentation: &Presentation<'_>,
) -> Result<Vec<embedded::content_parts::ContentPart>> {
    let package = presentation.package();
    let mut limits = embedded::content_parts::Limits::default();
    let mut content_parts = Vec::new();
    for (slide_index, slide) in slides(presentation)?.into_iter().enumerate() {
        content_parts.extend(embedded::content_parts::load_slide(
            package,
            slide_index,
            slide.part().part(),
            &mut limits,
        )?);
    }
    Ok(content_parts)
}

pub(super) fn ink_actions(
    presentation: &Presentation<'_>,
) -> Result<Vec<embedded::ink_actions::Snapshot>> {
    ink_actions_with_limits(presentation, embedded::ink_actions::Limits::default())
}

pub(super) fn ink_actions_with_limits(
    presentation: &Presentation<'_>,
    limits: embedded::ink_actions::Limits,
) -> Result<Vec<embedded::ink_actions::Snapshot>> {
    let package = presentation.package();
    let slide_parts = slides(presentation)?
        .into_iter()
        .enumerate()
        .map(|(slide_index, slide)| (slide_index, slide.part().part()));
    embedded::ink_actions::load_snapshots(package, slide_parts, &limits)
}

pub(super) fn hyperlinks(
    presentation: &Presentation<'_>,
) -> Result<Vec<(usize, crate::hyperlinks::Hyperlink)>> {
    let mut hyperlinks = Vec::new();
    for (slide_index, slide) in slides(presentation)?.into_iter().enumerate() {
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
    presentation: &Presentation<'a>,
    index: usize,
) -> Result<Option<Slide<'a>>> {
    let package = presentation.package();
    let references = presentation.catalog()?;
    let Some(reference) = references.get(index) else {
        return Ok(None);
    };
    let part = slide_part(package, presentation.part(), reference)?;
    Ok(Some(Slide::new_with_shape_and_semantic_limits(
        package,
        part,
        presentation.shape_limits(),
        presentation.semantic_limits.clone(),
    )))
}

pub(super) fn find_slide<'a>(
    presentation: &Presentation<'a>,
    key: Selector<'_>,
) -> Result<Option<Slide<'a>>> {
    match key {
        Selector::Position(position) => slide(presentation, position.get()),
        Selector::Name(name) => {
            let mut found = None;
            let mut matches = 0usize;
            for slide in slides(presentation)? {
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

pub(super) fn slides<'a>(presentation: &Presentation<'a>) -> Result<Vec<Slide<'a>>> {
    let package = presentation.package();
    let references = presentation.catalog()?;
    let mut slides = Vec::with_capacity(references.len());
    for reference in references {
        slides.push(Slide::new_with_shape_and_semantic_limits(
            package,
            slide_part(package, presentation.part(), reference)?,
            presentation.shape_limits(),
            // The semantic profile is retained by every slide projection.
            // Keep this call in the same constructor family as `slide` so a
            // projection never silently falls back to the default policy.
            presentation.semantic_limits.clone(),
        ));
    }
    Ok(slides)
}

pub(super) fn slide_masters<'a>(
    package: &'a OpcPackage,
    presentation: &PresentationPart<'a>,
    shape_limits: crate::shape::Limits,
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
        masters.push(SlideMaster::new_with_shape_limits(
            package,
            SlideMasterPart::from_part(part)?,
            shape_limits,
        ));
    }
    Ok(masters)
}

pub(super) fn slide_layouts<'a>(
    package: &'a OpcPackage,
    presentation: &PresentationPart<'a>,
    shape_limits: crate::shape::Limits,
) -> Result<Vec<SlideLayout<'a>>> {
    let mut layouts = Vec::new();
    for master in slide_masters(package, presentation, shape_limits)? {
        layouts.extend(master.layouts()?);
    }
    Ok(layouts)
}

pub(super) fn text(presentation: &Presentation<'_>) -> Result<String> {
    let mut text = String::new();
    for slide in slides(presentation)? {
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
    limits: crate::presentation_properties::metadata::slide_sync::Limits,
) -> Result<Vec<crate::presentation_properties::metadata::slide_sync::Part>> {
    crate::presentation_properties::metadata::slide_sync::load(
        package,
        limits.constrained_by_source(package.read_limits()),
    )
}

fn slide_part<'a>(
    package: &'a OpcPackage,
    presentation: &PresentationPart<'a>,
    reference: &SlideReference,
) -> Result<SlidePart<'a>> {
    let (_, _, part) = crate::parts::validate_slide_relationship(
        presentation.part().rels().get(reference.relationship_id()),
        reference.relationship_id(),
        |target| Ok(package.get_part(target)?),
        |part| part.content_type(),
    )?;
    SlidePart::from_part(part)
}

#[cfg(test)]
mod tests {
    use super::root_only_slide_with_source_limits;
    use crate::Error;
    use litchi_opc::{BlobPart, OpcError, PackURI, SourceLimitKind, SourceLimits};

    #[test]
    fn root_only_fallback_admits_source_before_mce() {
        let xml = br#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:x/></p:sld>"#;
        let part = BlobPart::new(
            PackURI::new("/ppt/slides/fallback.xml").expect("valid part name"),
            crate::notes::SLIDE_CT.to_owned(),
            xml.to_vec(),
        );
        let limits = SourceLimits::builder()
            .max_events(2)
            .expect("valid event limit")
            .build()
            .expect("valid source profile");
        assert!(matches!(
            root_only_slide_with_source_limits(&part, Some(limits)),
            Err(Error::Opc(OpcError::Limit {
                kind: SourceLimitKind::Events,
                observed: 3,
                maximum: 2,
            }))
        ));
    }
}
