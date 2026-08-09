//! Detached composition of ordering, shape-text, and notes-text edits.

use litchi_opc::{OpcPackage, PackURI};

use super::model::{Slide, Snapshot, capture, invalid};
use super::patch::{Delta, Patch};
use crate::{Error, Result};

/// One failure-atomic edit rooted in an immutable opened-package snapshot.
#[derive(Clone)]
pub struct Transaction {
    source: Snapshot,
    working: OpcPackage,
    slides: Vec<Slide>,
    touched: Vec<PackURI>,
}

impl std::fmt::Debug for Transaction {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Transaction")
            .field("source", &self.source)
            .field("slides", &self.slides)
            .field("touched", &self.touched)
            .finish_non_exhaustive()
    }
}

impl Transaction {
    pub(crate) fn new(source: Snapshot) -> Self {
        Self {
            working: source.package.as_ref().clone(),
            slides: source.slides.clone(),
            source,
            touched: Vec::new(),
        }
    }

    /// Immutable source root for this transaction.
    #[must_use]
    pub const fn source(&self) -> &Snapshot {
        &self.source
    }

    /// Slides in the transaction's currently staged order.
    #[must_use]
    pub fn slides(&self) -> &[Slide] {
        &self.slides
    }

    /// Whether any managed resource differs from the source root.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        !self.touched.is_empty()
    }

    /// Move one slide between checked zero-based positions.
    ///
    /// # Errors
    ///
    /// Returns an error for an out-of-range position or an unsupported
    /// slide-list dependency in the raw presentation XML.
    pub fn move_slide(&mut self, from: usize, to: usize) -> Result<bool> {
        let length = self.slides.len();
        if from >= length {
            return Err(Error::SlideIndexOutOfBounds {
                index: from,
                len: length,
            });
        }
        if to >= length {
            return Err(Error::SlideIndexOutOfBounds {
                index: to,
                len: length,
            });
        }
        if from == to {
            return Ok(false);
        }
        let mut order: Vec<_> = self.slides.iter().map(Slide::id).collect();
        let id = order.remove(from);
        order.insert(to, id);
        self.reorder_slides(&order)
    }

    /// Replace the complete slide order by stable slide IDs.
    ///
    /// # Errors
    ///
    /// Returns an error unless the IDs are an exact permutation of the
    /// currently staged slide identities.
    pub fn reorder_slides(&mut self, ordered_ids: &[u32]) -> Result<bool> {
        if ordered_ids == self.slides.iter().map(Slide::id).collect::<Vec<_>>() {
            return Ok(false);
        }
        let main = self.working.get_part(&self.source.presentation_name)?;
        let xml = super::xml::reorder_slides(main.blob(), &self.slides, ordered_ids)?;
        let mut reordered = Vec::new();
        reordered
            .try_reserve_exact(self.slides.len())
            .map_err(|source| Error::Allocation {
                resource: "opened-presentation reordered identities",
                source,
            })?;
        for id in ordered_ids {
            reordered.push(
                self.slides
                    .iter()
                    .find(|slide| slide.id == *id)
                    .cloned()
                    .ok_or_else(|| invalid("opened-presentation slide order lost an identity"))?,
            );
        }
        self.register(self.source.presentation_name.clone())?;
        self.working
            .get_part_mut(&self.source.presentation_name)?
            .set_blob(xml);
        self.slides = reordered;
        Ok(true)
    }

    /// Replace all visible text runs in one existing semantic shape.
    ///
    /// The first existing `a:t` run receives the escaped replacement and
    /// later runs become empty. Shape structure, formatting, unknown XML, and
    /// relationships remain exact.
    ///
    /// # Errors
    ///
    /// Returns an error for missing or ambiguous selectors, shapes without a
    /// text body, malformed raw XML, invalid characters, or exceeded bounds.
    pub fn set_shape_text<'s, 'k>(
        &mut self,
        slide: impl Into<crate::slide::Key<'s>>,
        shape: impl Into<crate::shape::Key<'k>>,
        text: impl AsRef<str>,
    ) -> Result<bool> {
        let selected_slide = self.resolve_slide(slide.into())?;
        let key = shape.into();
        let text = text.as_ref();
        let owner = self.working.get_part(&selected_slide.part_name)?;
        crate::parts::validate_content_type(owner, litchi_opc::constants::content_type::PML_SLIDE)?;
        let scene = crate::shape::Scene::read(owner.blob())?;
        let selected_shape = scene.shape(key)?;
        if selected_shape.common().text() == Some(text) {
            return Ok(false);
        }
        if selected_shape.common().text().is_none() {
            return Err(invalid(
                "opened-presentation selected shape has no text body",
            ));
        }
        let span = crate::tag::shape::selected_raw_span(owner.blob(), key)?;
        let xml = super::xml::rewrite_shape_text(
            owner.blob(),
            span,
            text,
            self.source.limits.max_text_bytes(),
        )?;
        let staged = crate::shape::Scene::read(&xml)?;
        let staged_shape = staged.shape(key)?;
        if staged_shape.common().text() != Some(text) || staged.len() != scene.len() {
            return Err(invalid(
                "opened-presentation shape text did not round-trip semantically",
            ));
        }
        self.register(selected_slide.part_name.clone())?;
        self.working
            .get_part_mut(&selected_slide.part_name)?
            .set_blob(xml);
        Ok(true)
    }

    /// Replace the existing notes text owned by one checked slide.
    ///
    /// # Errors
    ///
    /// Returns an error when the slide has no notes, the notes graph has an
    /// unsupported dependency, or the replacement is malformed or unbounded.
    pub fn set_notes_text<'s>(
        &mut self,
        slide: impl Into<crate::slide::Key<'s>>,
        text: impl AsRef<str>,
    ) -> Result<bool> {
        let selected_slide = self.resolve_slide(slide.into())?;
        let source = crate::notes::load_snapshot(&self.working, &self.source.presentation_name)?
            .ok_or_else(|| invalid("opened-presentation notes graph is absent"))?;
        let note = source
            .slides()
            .iter()
            .find(|note| note.owner() == selected_slide.part_name.as_str())
            .ok_or_else(|| invalid("opened-presentation selected slide has no notes"))?;
        let text = text.as_ref();
        if note.text()?.as_deref() == (!text.is_empty()).then_some(text) {
            return Ok(false);
        }
        if text.len() > self.source.limits.max_text_bytes() {
            return Err(Error::Limit {
                resource: "opened-presentation notes text bytes",
                limit: self.source.limits.max_text_bytes(),
            });
        }
        let part_name = PackURI::new(note.part()).map_err(Error::Invalid)?;
        let xml = crate::notes::rewrite_text(note.xml(), text)?;
        self.register(part_name.clone())?;
        self.working.get_part_mut(&part_name)?.set_blob(xml);
        let staged = crate::notes::load_snapshot(&self.working, &self.source.presentation_name)?
            .ok_or_else(|| invalid("opened-presentation staged notes graph disappeared"))?;
        let staged_note = staged
            .slides()
            .iter()
            .find(|note| note.owner() == selected_slide.part_name.as_str())
            .ok_or_else(|| invalid("opened-presentation staged notes owner disappeared"))?;
        let actual = staged_note.text()?;
        if actual.as_deref() != (!text.is_empty()).then_some(text) {
            return Err(invalid(format!(
                "opened-presentation notes text did not round-trip semantically: requested {text:?}, read {actual:?}"
            )));
        }
        Ok(true)
    }

    /// Validate and consume all staged edits into one atomic commit.
    ///
    /// # Errors
    ///
    /// Returns an error if a dependency changed, the patch exceeds bounds, or
    /// the complete staged package cannot be captured as a coherent snapshot.
    pub fn commit(self) -> Result<Commit> {
        let mut deltas = Vec::new();
        deltas
            .try_reserve_exact(self.touched.len())
            .map_err(|source| Error::Allocation {
                resource: "opened-presentation commit deltas",
                source,
            })?;
        for name in self.touched {
            if let Some(delta) = Delta::capture(self.source.package.as_ref(), &self.working, name)?
            {
                deltas.push(delta);
            }
        }
        let patch = Patch::new(
            self.source.presentation_name.clone(),
            deltas,
            self.source.limits,
        )?;
        if patch.is_empty() {
            return Ok(Commit {
                snapshot: self.source,
                patch,
            });
        }
        let mut working = self.working;
        working.unsign();
        let snapshot = capture(&working, self.source.limits)?;
        Ok(Commit { snapshot, patch })
    }

    /// Discard all staged edits and recover the immutable source root.
    #[must_use]
    pub fn rollback(self) -> Snapshot {
        self.source
    }

    fn resolve_slide(&self, key: crate::slide::Key<'_>) -> Result<Slide> {
        match key {
            crate::slide::Key::Index(index) => {
                self.slides
                    .get(index)
                    .cloned()
                    .ok_or(Error::SlideIndexOutOfBounds {
                        index,
                        len: self.slides.len(),
                    })
            },
            crate::slide::Key::Name(name) => {
                let mut matches = self.slides.iter().filter(|slide| slide.name == name);
                let selected = matches.next().cloned();
                if matches.next().is_some() {
                    return Err(Error::AmbiguousSlideName {
                        name: name.to_owned(),
                        matches: self
                            .slides
                            .iter()
                            .filter(|slide| slide.name == name)
                            .count(),
                    });
                }
                selected.ok_or_else(|| Error::SlideNameNotFound(name.to_owned()))
            },
        }
    }

    fn register(&mut self, name: PackURI) -> Result<()> {
        if !self.touched.iter().any(|existing| existing == &name)
            && self.touched.len() == self.source.limits.max_parts()
        {
            return Err(Error::Limit {
                resource: "opened-presentation transaction parts",
                limit: self.source.limits.max_parts(),
            });
        }
        if !self.touched.iter().any(|existing| existing == &name) {
            self.touched.push(name);
        }
        Ok(())
    }
}

/// Validated result of one opened-presentation transaction.
#[derive(Debug, Clone)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
}

impl Commit {
    /// Candidate snapshot after atomic publication.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Durable exact-source patch for this commit.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Whether publication changes any part bytes.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        !self.patch.is_empty()
    }

    /// Consume the commit into its candidate snapshot and patch.
    #[must_use]
    pub fn into_parts(self) -> (Snapshot, Patch) {
        (self.snapshot, self.patch)
    }

    /// Consume the commit into the durable patch.
    #[must_use]
    pub fn into_patch(self) -> Patch {
        self.patch
    }
}
