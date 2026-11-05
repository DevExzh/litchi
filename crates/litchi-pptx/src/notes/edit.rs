//! Immutable speaker-notes edits of a [`Package`] (record 0351 §2.3).
//!
//! [`Package::edit_notes`] starts an [`Edit`] over the package's notes graph:
//! [`Edit::set`] replaces the complete detached graph, the `set_*` verbs
//! replace one notes slide's text or XML or the notes master or theme XML,
//! [`Edit::remove`] removes one slide's notes and [`Edit::clear`] every
//! slide's notes. [`Edit::commit`] returns a [`Commit`] with the new
//! package, the reversible [`Patch`] and the [`Diagnostics`];
//! [`Package::apply_notes`] replays a patch onto an equal package. Opaque
//! XML and inert links are never interpreted. An exact no-op keeps every
//! byte, and a changed edit of a signed package is refused with
//! [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy).

use super::{Draft, Graph};
use crate::part_edit::{Staged, family};
use crate::parts::PresentationPart;
use crate::slide::Selector;
use crate::{Package, Result};

family!("notes");

/// A staged speaker-notes edit of one [`Package`]. Dropping it discards it.
#[derive(Debug)]
pub struct Edit<'a> {
    staged: Staged<'a>,
}

impl Package {
    /// Start a speaker-notes edit.
    ///
    /// # Errors
    ///
    /// Returns the encryption-policy error for a package whose typed
    /// mutations are refused.
    pub fn edit_notes(&self) -> Result<Edit<'_>> {
        Ok(Edit {
            staged: Staged::new(self, "edit_notes")?,
        })
    }

    /// Apply a speaker-notes patch to a new package; this one is unchanged.
    /// An empty patch returns an equal package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state, and the signature refusal for a changed patch of a
    /// signed package.
    pub fn apply_notes(&self, patch: &Patch) -> Result<Commit> {
        let package = patch.0.apply_with_replay_validation(
            self,
            "apply_notes",
            |package, target, patch| {
                let limits = package
                    .semantic_limits()
                    .notes()
                    .constrained_by_source(target.read_limits());
                patch.validate_changed_parts(target, |_, part| {
                    super::validate_replay_part(part, limits)
                })?;
                if patch.touches_notes(&package.opc, target)? {
                    let index = patch.replay_index_after(package, target)?;
                    super::validate_replay_total(&index, limits)?;
                }
                Ok(())
            },
        )?;
        Ok(commit_of((package, patch.0.clone())))
    }
}

impl Edit<'_> {
    fn limits(&self) -> super::Limits {
        self.staged
            .semantic_limits()
            .notes()
            .constrained_by_source(self.staged.opc().read_limits())
    }

    fn snapshot(&self) -> Result<Option<super::Snapshot>> {
        let opc = self.staged.opc();
        let presentation = PresentationPart::from_package(opc)?;
        super::load_snapshot_with_limits(opc, presentation.part().partname(), self.limits())
    }

    fn stage(&mut self, change: impl FnOnce(&mut Draft) -> Result<bool>) -> Result<bool> {
        let source = self
            .snapshot()?
            .ok_or_else(|| super::invalid("notes graph is absent"))?;
        let mut draft = source.edit();
        if !change(&mut draft)? {
            return Ok(false);
        }
        let prepared = draft.commit()?;
        if !prepared.diagnostics().changed() {
            return Ok(false);
        }
        self.staged
            .update(|opc| super::apply_commit(opc, prepared).map(drop))?;
        Ok(true)
    }

    /// The staged notes graph, when the presentation has one.
    ///
    /// # Errors
    ///
    /// Returns an error when the notes graph is malformed or exceeds its
    /// resource limits.
    pub fn graph(&self) -> Result<Option<Graph>> {
        Ok(self.snapshot()?.map(|snapshot| snapshot.graph().clone()))
    }

    /// Replace the complete detached notes graph without changing physical
    /// owners. Returns whether anything changed.
    ///
    /// # Errors
    ///
    /// Returns an error for an absent or malformed graph, a replacement that
    /// retargets a relationship identity, or a changed edit of a signed
    /// package.
    pub fn set(&mut self, graph: Graph) -> Result<bool> {
        self.stage(|draft| draft.replace(graph))
    }

    /// Replace the inert text of the selected slide's notes while preserving
    /// their XML. Returns whether anything changed.
    ///
    /// # Errors
    ///
    /// Returns an error for an absent graph, a missing, ambiguous or
    /// out-of-range slide, a slide without notes, or a changed edit of a
    /// signed package.
    pub fn set_text<'s>(
        &mut self,
        slide: impl Into<Selector<'s>>,
        text: impl AsRef<str>,
    ) -> Result<bool> {
        let index = self.notes_index(slide.into())?;
        self.stage(|draft| draft.set_text(index, text))
    }

    /// Replace the XML of the selected slide's notes slide after bounded
    /// conformance validation. Returns whether anything changed.
    ///
    /// # Errors
    ///
    /// As [`Self::set_text`], and for malformed or nonconforming XML.
    pub fn set_slide_xml<'s>(
        &mut self,
        slide: impl Into<Selector<'s>>,
        xml: Vec<u8>,
    ) -> Result<bool> {
        let index = self.notes_index(slide.into())?;
        self.stage(|draft| draft.replace_slide_xml(index, xml))
    }

    /// Replace the notes master XML. Returns whether anything changed.
    ///
    /// # Errors
    ///
    /// As [`Self::set_slide_xml`].
    pub fn set_master_xml(&mut self, xml: Vec<u8>) -> Result<bool> {
        self.stage(|draft| draft.replace_master_xml(xml))
    }

    /// Replace the notes theme XML. Returns whether anything changed.
    ///
    /// # Errors
    ///
    /// As [`Self::set_slide_xml`].
    pub fn set_theme_xml(&mut self, xml: Vec<u8>) -> Result<bool> {
        self.stage(|draft| draft.replace_theme_xml(xml))
    }

    /// Remove the notes of the selected slide. Slide-owned comments are
    /// retained; descendants of the removed notes slide are collected only
    /// when package-wide inbound-edge validation proves exclusive ownership.
    /// Returns whether anything changed.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing, ambiguous or out-of-range slide,
    /// malformed or shared ownership, or a changed edit of a signed package.
    pub fn remove<'s>(&mut self, slide: impl Into<Selector<'s>>) -> Result<bool> {
        let slide = crate::package::resolve_slide(self.staged.opc(), slide.into())?;
        let Some(source) = self.snapshot()? else {
            return Ok(false);
        };
        let limits = self.limits();
        self.staged
            .update(|opc| super::remove_checked_with_limits(opc, &source, &slide, limits))
    }

    /// Remove the notes of every slide while retaining shared notes
    /// resources. Returns the number of notes slides removed.
    ///
    /// # Errors
    ///
    /// Returns an error when the graph is malformed, descendants have shared
    /// ownership, or for a changed edit of a signed package.
    pub fn clear(&mut self) -> Result<usize> {
        let Some(source) = self.snapshot()? else {
            return Ok(0);
        };
        let limits = self.limits();
        self.staged
            .update(|opc| super::clear_checked_with_limits(opc, &source, limits))
    }

    /// The position, in the staged notes graph, of the selected slide's
    /// notes slide.
    fn notes_index(&self, slide: Selector<'_>) -> Result<usize> {
        let owner = crate::package::resolve_slide(self.staged.opc(), slide)?;
        let graph = self.graph()?.ok_or(Error::refused(
            litchi_core::edit::RefusalKind::NotFound,
            "notes",
            "the presentation has no notes",
        ))?;
        graph
            .slides()
            .iter()
            .position(|notes| notes.owner() == owner.as_str())
            .ok_or(Error::refused(
                litchi_core::edit::RefusalKind::NotFound,
                "notes",
                "the slide has no notes",
            ))
    }

    /// Discard every staged change.
    pub fn reset(&mut self) {
        self.staged.reset();
    }

    /// Publish the staged changes to a new package.
    ///
    /// # Errors
    ///
    /// Returns the part-count limit of the patch.
    pub fn commit(self) -> Result<Commit> {
        Ok(commit_of(self.staged.commit()?))
    }
}
