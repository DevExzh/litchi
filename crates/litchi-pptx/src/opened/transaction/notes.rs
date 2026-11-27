//! Creating a slide's notes page and rewriting its text.

use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, PackURI, Part as _};

use super::{Edit, invalid};
use crate::{Error, Result};

fn refuse(reason: &'static str) -> Error {
    Error::refused(
        litchi_core::edit::RefusalKind::UnsupportedTarget,
        "set_notes_text",
        reason,
    )
}

impl Edit {
    /// Stage a new notes slide for the slide part `slide` holding `text`,
    /// related to the presentation's notes master exactly as `PowerPoint`
    /// wires one: notes slide → slide and → notes master, slide → notes
    /// slide.
    pub(super) fn create_notes_slide(&mut self, slide: &PackURI, text: &str) -> Result<()> {
        let (presentation_name, max_parts) = {
            let source = self.source_ref()?;
            (source.presentation_name.clone(), source.limits.max_parts())
        };
        let master = {
            let presentation = self.working.get_part(&presentation_name)?;
            let mut masters = presentation.rels().iter().filter(|relationship| {
                !relationship.is_external()
                    && (relationship.reltype() == rt::NOTES_MASTER
                        || relationship.reltype() == rt::STRICT_NOTES_MASTER)
            });
            let master = masters
                .next()
                .ok_or_else(|| refuse("the presentation has no notes master"))?
                .target_partname()?;
            if masters.next().is_some() {
                return Err(refuse("the presentation has several notes masters"));
            }
            master
        };
        let owner = self.working.get_part(slide)?;
        if crate::media_parts::document_conformance_with_limits(
            owner.blob(),
            self.semantic_limits()
                .media_parts()
                .constrained_by_source(self.working.read_limits()),
        )? != crate::media_parts::Conformance::Transitional
        {
            return Err(refuse(
                "notes creation is implemented for transitional presentations only",
            ));
        }
        if owner
            .rels()
            .iter()
            .any(|relationship| relationship.reltype() == rt::NOTES_SLIDE)
        {
            return Err(refuse("the slide already relates a notes page"));
        }
        if self.working.part_count() >= max_parts {
            return Err(Error::Limit {
                kind: crate::opened::LimitKind::Parts,
                observed: self
                    .working
                    .part_count()
                    .checked_add(1)
                    .unwrap_or(usize::MAX),
                maximum: max_parts,
            });
        }
        let name = self.free_notes_name()?;
        let mut part = BlobPart::new(
            name.clone(),
            ct::PML_NOTES_SLIDE.to_owned(),
            crate::notes::write_text(text, self.notes_limits())?,
        );
        part.relate_to(&slide.relative_ref(name.base_uri()), rt::SLIDE);
        part.relate_to(&master.relative_ref(name.base_uri()), rt::NOTES_MASTER);
        let mut candidate = self.working.clone();
        candidate.try_add_part(Box::new(part))?;
        candidate
            .get_part_mut(slide)?
            .relate_to(&name.relative_ref(slide.base_uri()), rt::NOTES_SLIDE);
        self.working = candidate;
        Ok(())
    }

    /// The first `/ppt/notesSlides/notesSlideN.xml` that names no part and
    /// no relationship target.
    fn free_notes_name(&self) -> Result<PackURI> {
        let targeted = |name: &PackURI| {
            std::iter::once(self.working.rels())
                .chain(self.working.iter_parts().map(|part| part.rels()))
                .flat_map(|relationships| relationships.iter())
                .filter(|relationship| !relationship.is_external())
                .any(|relationship| {
                    relationship
                        .target_partname()
                        .is_ok_and(|target| &target == name)
                })
        };
        let bound = self.working.part_count().saturating_add(1);
        for index in 1..=bound {
            let name = PackURI::new(format!("/ppt/notesSlides/notesSlide{index}.xml"))
                .map_err(Error::Invalid)?;
            if self.working.validate_new_part_name(&name).is_ok() && !targeted(&name) {
                return Ok(name);
            }
        }
        Err(refuse("no free notes slide part name"))
    }
}

/// The refusal reason for a notes text write whose body holds an equation.
const EQUATION_NOTES_REASON: &str =
    "the notes hold an Office Math equation, which text edits cannot change";

/// The notes XML with the text of its body placeholder replaced by `text`,
/// or `None` when that text is already `text`. `operation` names the
/// refusals.
///
/// The notes text is the text of the one top-level body placeholder, as the
/// presentation reader projects it; only that shape's `DrawingML` text runs
/// are rewritten, so the slide number, header, footer and date placeholders
/// and any other shape keep their bytes. An empty body first takes one empty
/// run in its first paragraph.
pub(crate) fn rewrite_notes_body(
    xml: &[u8],
    text: &str,
    operation: &'static str,
    shape_limits: crate::shape::Limits,
    max_text_bytes: usize,
) -> Result<Option<Vec<u8>>> {
    let refusal = |kind, reason| Error::refused(kind, operation, reason);
    let scene = crate::shape::Scene::read(xml, shape_limits)?;
    let mut bodies = scene.roots().filter(|shape| {
        matches!(shape.placeholder(), Some(placeholder) if placeholder.kind() == Some("body"))
    });
    let body = bodies.next();
    if bodies.next().is_some() {
        return Err(Error::Invalid(
            "notes slide has multiple body placeholders".into(),
        ));
    }
    let current = body
        .and_then(|body| body.common().text())
        .filter(|current| !current.is_empty());
    if current == (!text.is_empty()).then_some(text) {
        return Ok(None);
    }
    let Some(body) = body else {
        return Err(refusal(
            litchi_core::edit::RefusalKind::UnsupportedTarget,
            "the notes page has no body placeholder",
        ));
    };
    if body.common().text().is_none() {
        return Err(refusal(
            litchi_core::edit::RefusalKind::UnsupportedTarget,
            "the notes body placeholder has no text body",
        ));
    }
    let span = crate::tag::shape::selected_raw_span_for_shape(xml, &scene, body)?;
    super::refuse_equation_shape(xml, &scene, &span, operation, EQUATION_NOTES_REASON)?;
    let refusals = crate::opened::xml::TextRefusals::notes(operation);
    match crate::opened::xml::insert_run_into_empty_text(xml, span.clone())? {
        Some((with_run, span)) => crate::opened::xml::rewrite_shape_text_with_limits(
            &with_run,
            span,
            text,
            max_text_bytes,
            shape_limits,
            refusals,
        ),
        None => crate::opened::xml::rewrite_shape_text_with_limits(
            xml,
            span,
            text,
            max_text_bytes,
            shape_limits,
            refusals,
        ),
    }
    .map(Some)
}

/// Read the notes of `slide` in `package` back as the presentation reader
/// does and require `text`.
pub(crate) fn verify_notes_text(
    package: &litchi_opc::OpcPackage,
    presentation_name: &PackURI,
    slide: &PackURI,
    text: &str,
    shape_limits: crate::shape::Limits,
    notes_limits: crate::notes::Limits,
) -> Result<()> {
    let staged = crate::notes::load_snapshot_with_limits(package, presentation_name, notes_limits)?
        .ok_or_else(|| invalid("opened-presentation staged notes graph disappeared"))?;
    if !staged
        .slides()
        .iter()
        .any(|note| note.owner() == slide.as_str())
    {
        return Err(invalid(
            "opened-presentation staged notes owner disappeared",
        ));
    }
    let slide_part = crate::parts::SlidePart::from_part(package.get_part(slide)?)?;
    let notes = crate::presentation::source::related::load_eager_notes(
        package,
        &slide_part,
        shape_limits,
        notes_limits,
    )?
    .ok_or_else(|| invalid("opened-presentation staged notes owner disappeared"))?;
    if notes.text() != (!text.is_empty()).then_some(text) {
        return Err(invalid(
            "opened-presentation notes text did not round-trip semantically",
        ));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use crate::opened::xml::insert_run_into_empty_text;

    const A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";

    /// `body` inside a shape, with the shape's range; the inserted XML of the
    /// shape, or `None`.
    fn insert(declaration: &str, body: &str) -> Option<String> {
        let xml = format!(r#"<root xmlns:p="p"{declaration}><p:sp>{body}</p:sp><p:sp/></root>"#);
        let start = xml.find("<p:sp>").unwrap();
        let end = xml.find("</p:sp>").unwrap() + "</p:sp>".len();
        insert_run_into_empty_text(xml.as_bytes(), start..end)
            .unwrap()
            .map(|(output, owner)| {
                let output = String::from_utf8(output).unwrap();
                assert!(output.ends_with("<p:sp/></root>"));
                output[owner].to_owned()
            })
    }

    #[test]
    fn an_empty_first_paragraph_takes_one_run() {
        let a = format!(r#" xmlns:a="{A}""#);
        assert_eq!(
            insert(
                &a,
                r#"<p:txBody><a:p><a:endParaRPr lang="en-US"/></a:p><a:p/></p:txBody>"#
            )
            .as_deref(),
            Some(
                r#"<p:sp><p:txBody><a:p><a:r><a:rPr lang="en-US"/><a:t></a:t></a:r><a:endParaRPr lang="en-US"/></a:p><a:p/></p:txBody></p:sp>"#
            )
        );
        assert_eq!(
            insert(&a, r#"<p:txBody><a:p/></p:txBody>"#).as_deref(),
            Some(r#"<p:sp><p:txBody><a:p><a:r><a:t></a:t></a:r></a:p></p:txBody></p:sp>"#)
        );
        assert_eq!(
            insert(&a, r#"<p:txBody><a:p><a:pPr/></a:p></p:txBody>"#).as_deref(),
            Some(r#"<p:sp><p:txBody><a:p><a:pPr/><a:r><a:t></a:t></a:r></a:p></p:txBody></p:sp>"#)
        );
        // End-paragraph properties with children are copied whole.
        assert_eq!(
            insert(
                &a,
                r#"<p:txBody><a:p><a:endParaRPr sz="1200"><a:latin typeface="X"/></a:endParaRPr></a:p></p:txBody>"#
            )
            .as_deref(),
            Some(
                r#"<p:sp><p:txBody><a:p><a:r><a:rPr sz="1200"><a:latin typeface="X"/></a:rPr><a:t></a:t></a:r><a:endParaRPr sz="1200"><a:latin typeface="X"/></a:endParaRPr></a:p></p:txBody></p:sp>"#
            )
        );
        // A default DrawingML namespace takes unprefixed names.
        assert_eq!(
            insert(
                "",
                &format!(r#"<p:txBody><p xmlns="{A}"><endParaRPr/></p></p:txBody>"#)
            )
            .as_deref(),
            Some(
                format!(
                    r#"<p:sp><p:txBody><p xmlns="{A}"><r><rPr/><t></t></r><endParaRPr/></p></p:txBody></p:sp>"#
                )
                .as_str()
            )
        );
    }

    #[test]
    fn a_body_with_any_text_element_is_left_alone() {
        let a = format!(r#" xmlns:a="{A}""#);
        assert!(
            insert(
                &a,
                r#"<p:txBody><a:p/><a:p><a:fld id="x"><a:t>1</a:t></a:fld></a:p></p:txBody>"#
            )
            .is_none()
        );
        assert!(insert(&a, r#"<p:txBody><a:p><a:r><a:t/></a:r></a:p></p:txBody>"#).is_none());
        // A paragraph outside DrawingML is not a text paragraph.
        assert!(insert(&a, r#"<p:txBody><x:p xmlns:x="urn:x"/></p:txBody>"#).is_none());
    }
}
