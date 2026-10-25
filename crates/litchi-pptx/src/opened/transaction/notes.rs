//! Creating a slide's notes page (record 0252).

use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, PackURI, Part as _};

use super::Edit;
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
