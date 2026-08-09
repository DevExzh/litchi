use super::codec::encode_record;
use super::package::validate_unrelated_streams;
use super::{Facet, Font, FontCollections, FontEmbeddingFlags, Patch, Scope, Snapshot};
use crate::package::{Error, Result};
use crate::records::Record;
use std::sync::Arc;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChangeKind {
    Replace,
    Append,
    SetFacet,
    RemoveFacet,
    EmbeddingFlags,
}

/// Compact mutation descriptor; large EOT payloads are not duplicated here.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Change {
    kind: ChangeKind,
    scope: Option<Scope>,
    index: Option<u16>,
    facet: Option<Facet>,
}

impl Change {
    #[must_use]
    pub const fn kind(&self) -> ChangeKind {
        self.kind
    }
    #[must_use]
    pub const fn scope(&self) -> Option<Scope> {
        self.scope
    }
    #[must_use]
    pub const fn index(&self) -> Option<u16> {
        self.index
    }
    #[must_use]
    pub const fn facet(&self) -> Option<Facet> {
        self.facet
    }
}

#[derive(Debug, Clone)]
pub struct Transaction {
    source: Snapshot,
    document: Arc<Record>,
    fonts: FontCollections,
    changes: Vec<Change>,
}

impl Transaction {
    pub(crate) fn new(source: Snapshot) -> Self {
        Self {
            fonts: source.fonts.clone(),
            document: source.document_record.clone(),
            source,
            changes: Vec::new(),
        }
    }

    #[must_use]
    pub const fn source(&self) -> &Snapshot {
        &self.source
    }
    #[must_use]
    pub const fn fonts(&self) -> &FontCollections {
        &self.fonts
    }
    #[must_use]
    pub fn changes(&self) -> &[Change] {
        &self.changes
    }
    #[must_use]
    pub fn is_changed(&self) -> bool {
        !self.changes.is_empty()
    }

    /// Stage a replacement of the font at `index` in the collection for
    /// `scope`.
    ///
    /// # Errors
    ///
    /// Returns an error if the collection is absent, `index` is unknown, or
    /// the result violates the source limits.
    pub fn replace_font(&mut self, scope: Scope, index: u16, font: Font) -> Result<()> {
        if self
            .fonts
            .collection(scope)
            .and_then(|collection| collection.get(index))
            .is_some_and(|old| old == &font)
        {
            return Ok(());
        }
        let mut candidate = self.fonts.clone();
        candidate
            .collection_mut(scope)
            .ok_or_else(|| missing_collection(scope))?
            .replace(index, font)?;
        candidate.validate_with_limits(self.source.limits.fonts)?;
        self.fonts = candidate;
        self.changes.push(Change {
            kind: ChangeKind::Replace,
            scope: Some(scope),
            index: Some(index),
            facet: None,
        });
        Ok(())
    }

    /// Stage an append of `font` to the collection for `scope`, returning its
    /// new ordinal.
    ///
    /// # Errors
    ///
    /// Returns an error if the collection is absent, the format font limit
    /// would be exceeded, or the result violates the source limits.
    pub fn append_font(&mut self, scope: Scope, font: Font) -> Result<u16> {
        let mut candidate = self.fonts.clone();
        let index = candidate
            .collection_mut(scope)
            .ok_or_else(|| missing_collection(scope))?
            .try_push(font)?;
        candidate.validate_with_limits(self.source.limits.fonts)?;
        self.fonts = candidate;
        self.changes.push(Change {
            kind: ChangeKind::Append,
            scope: Some(scope),
            index: Some(index),
            facet: None,
        });
        Ok(index)
    }

    /// Stage a validated EOT facet on the font at `index` in the collection
    /// for `scope`.
    ///
    /// # Errors
    ///
    /// Returns an error if the collection is absent, `index` is unknown,
    /// `data` fails EOT validation, or the result violates the source limits.
    pub fn set_facet(
        &mut self,
        scope: Scope,
        index: u16,
        facet: Facet,
        data: impl Into<super::SharedFontData>,
    ) -> Result<()> {
        let payload = data.into();
        if self
            .fonts
            .collection(scope)
            .ok_or_else(|| missing_collection(scope))?
            .get(index)
            .and_then(|font| font.facet(facet))
            .is_some_and(|old| old.bytes() == payload.as_ref())
        {
            return Ok(());
        }
        let mut candidate = self.fonts.clone();
        candidate
            .collection_mut(scope)
            .ok_or_else(|| missing_collection(scope))?
            .set_facet(index, facet, payload)?;
        candidate.validate_with_limits(self.source.limits.fonts)?;
        self.fonts = candidate;
        self.changes.push(Change {
            kind: ChangeKind::SetFacet,
            scope: Some(scope),
            index: Some(index),
            facet: Some(facet),
        });
        Ok(())
    }

    /// Prepare and stage one inert EOT facet without loading or executing it.
    ///
    /// The `PowerPoint` facet is derived from `font.style`. License, intent, and
    /// EOT-limit failures happen before the transaction model is mutated.
    ///
    /// # Errors
    ///
    /// Returns an error if the collection or font ordinal is absent, the
    /// prepared font cannot be encoded or rolled back, or the result violates
    /// the source limits.
    #[cfg(feature = "fonts")]
    pub fn set_prepared_facet(
        &mut self,
        scope: Scope,
        index: u16,
        font: &mut super::PreparedFont,
        intent: super::EotIntent,
        limits: super::EotLimits,
    ) -> Result<()> {
        let facet = super::prepared::facet_for_style(font.style);
        let subsetted = font.subsetted;
        self.fonts.validate_with_limits(self.source.limits.fonts)?;
        let current = self
            .fonts
            .collection(scope)
            .and_then(|collection| collection.get(index))
            .ok_or_else(|| {
                Error::InvalidFormat(format!("unknown {scope:?} font ordinal {index}"))
            })?;
        if current
            .embedded_fonts
            .iter()
            .any(|embedded| embedded.style != facet as u8)
            && current.embedded_subset != subsetted
        {
            return Err(Error::InvalidFormat(
                "one PowerPoint face cannot mix subsetted and complete embedded facets".into(),
            ));
        }

        let data = litchi_fonts::embedding::powerpoint::encode(font, intent, limits)
            .map_err(|error| Error::InvalidFormat(format!("font preparation failed: {error}")))?;
        if data.len() > self.source.limits.fonts.max_facet_bytes {
            super::prepared::restore_encoded(font, data, limits).map_err(|error| {
                Error::Corrupted(format!("prepared-font rollback invariant failed: {error}"))
            })?;
            return Err(Error::ResourceLimit(
                "prepared embedded font facet exceeds its byte limit".into(),
            ));
        }

        let mut candidate = self.fonts.clone();
        let collection = candidate
            .collection_mut(scope)
            .ok_or_else(|| Error::InvalidFormat(format!("{scope:?} font collection is absent")))?;
        super::prepared::stage_facet(collection, index, facet, data);
        let staged = collection.get_mut(index).ok_or_else(|| {
            Error::InvalidFormat(format!("unknown {scope:?} font ordinal {index}"))
        })?;
        staged.embedded_subset = subsetted;
        if subsetted {
            staged.font_flags |= 1;
        } else {
            staged.font_flags &= !1;
        }
        if let Err(error) = candidate.validate_with_limits(self.source.limits.fonts) {
            super::prepared::restore_staged(font, &mut candidate, scope, index, facet, limits)
                .map_err(|restore| {
                    Error::Corrupted(format!(
                        "prepared-font rollback invariant failed: {restore}"
                    ))
                })?;
            return Err(error);
        }
        let staged_collection = candidate
            .collection(scope)
            .ok_or_else(|| Error::InvalidFormat(format!("{scope:?} font collection is absent")))?;
        let serialized = staged_collection.to_record_bytes_with_limits(self.source.limits.fonts);
        if let Err(error) = serialized {
            super::prepared::restore_staged(font, &mut candidate, scope, index, facet, limits)
                .map_err(|restore| {
                    Error::Corrupted(format!(
                        "prepared-font rollback invariant failed: {restore}"
                    ))
                })?;
            return Err(error);
        }
        if candidate == self.fonts {
            super::prepared::restore_staged(font, &mut candidate, scope, index, facet, limits)
                .map_err(|restore| {
                    Error::Corrupted(format!(
                        "prepared-font rollback invariant failed: {restore}"
                    ))
                })?;
            return Ok(());
        }
        self.fonts = candidate;
        self.changes.push(Change {
            kind: ChangeKind::SetFacet,
            scope: Some(scope),
            index: Some(index),
            facet: Some(facet),
        });
        Ok(())
    }

    /// Stage removal of `facet` from the font at `index` in the collection
    /// for `scope`, returning whether a facet was removed.
    ///
    /// # Errors
    ///
    /// Returns an error if the collection is absent, `index` is unknown, or
    /// the result violates the source limits.
    pub fn remove_facet(&mut self, scope: Scope, index: u16, facet: Facet) -> Result<bool> {
        let mut candidate = self.fonts.clone();
        let removed = candidate
            .collection_mut(scope)
            .ok_or_else(|| missing_collection(scope))?
            .remove_facet(index, facet)?
            .is_some();
        if removed {
            let font = candidate
                .collection_mut(scope)
                .and_then(|collection| collection.get_mut(index))
                .ok_or_else(|| {
                    Error::InvalidFormat(format!("unknown {scope:?} font ordinal {index}"))
                })?;
            if font.embedded_fonts.is_empty() {
                font.embedded_subset = false;
                font.font_flags &= !1;
            }
            candidate.validate_with_limits(self.source.limits.fonts)?;
            self.fonts = candidate;
            self.changes.push(Change {
                kind: ChangeKind::RemoveFacet,
                scope: Some(scope),
                index: Some(index),
                facet: Some(facet),
            });
        }
        Ok(removed)
    }

    /// Stage document-wide embedding flags.
    ///
    /// # Errors
    ///
    /// Returns an error if the result violates the source limits.
    pub fn set_embedding_flags(&mut self, flags: Option<FontEmbeddingFlags>) -> Result<()> {
        if self.fonts.embedding_flags == flags {
            return Ok(());
        }
        let mut candidate = self.fonts.clone();
        candidate.embedding_flags = flags;
        candidate.validate_with_limits(self.source.limits.fonts)?;
        self.fonts = candidate;
        self.changes.push(Change {
            kind: ChangeKind::EmbeddingFlags,
            scope: None,
            index: None,
            facet: None,
        });
        Ok(())
    }

    /// Removal is refused until every base and PP10 reference can be proven and remapped.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn remove_font(&mut self, _scope: Scope, _index: u16) -> Result<Font> {
        Err(Error::InvalidFormat(
            "font removal requires a complete base and PP10 reference remap".into(),
        ))
    }

    /// Reordering is refused until every base and PP10 reference can be proven and remapped.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn reorder_fonts(&mut self, _scope: Scope, _order: &[u16]) -> Result<()> {
        Err(Error::InvalidFormat(
            "font reordering requires a complete base and PP10 reference remap".into(),
        ))
    }

    /// Publish the staged changes as an incremental patch and return the
    /// resulting commit.
    ///
    /// # Errors
    ///
    /// Returns an error if the publication budget is exceeded, the live owner
    /// changed during staging, serialization fails, or the published candidate
    /// fails its semantic reopen.
    pub fn commit(mut self) -> Result<Commit> {
        if self.changes.is_empty() {
            let revision = self.source.revision();
            let patch = Patch {
                base: revision,
                target: revision,
                before: self.source.bytes.clone(),
                after: self.source.bytes.clone(),
                changes: Vec::new(),
                limits: self.source.limits,
            };
            return Ok(Commit {
                snapshot: self.source,
                patch,
            });
        }

        preflight_publication_budget(
            self.source.bytes.len(),
            self.source.document.len(),
            self.source.limits.max_source_bytes,
        )?;
        // Gate protected input and validate the live owner before cloning the
        // normalized record tree or allocating a serialization candidate.
        super::package::require_stream_only_cfb(self.source.bytes())?;
        let mut editor = crate::embedded::object::Editor::open_records_arc_with_limit(
            self.source.bytes.clone(),
            self.source.limits.max_source_bytes,
        )?;
        let live = editor.persisted_record(self.source.document_persist_id)?;
        if live.as_slice() != self.source.document.as_ref() {
            return Err(Error::InvalidFormat(
                "live font owner changed during staging".into(),
            ));
        }

        let document = Arc::make_mut(&mut self.document);
        self.fonts
            .materialize_to_document(document, self.source.limits.fonts)?;
        synchronize_save_with_fonts(document, collections_have_embedded_fonts(&self.fonts))?;
        let target_document = encode_record(document, self.source.limits.fonts.records)?;
        if target_document == self.source.document.as_ref() {
            let revision = self.source.revision();
            let patch = Patch {
                base: revision,
                target: revision,
                before: self.source.bytes.clone(),
                after: self.source.bytes.clone(),
                changes: Vec::new(),
                limits: self.source.limits,
            };
            return Ok(Commit {
                snapshot: self.source,
                patch,
            });
        }

        editor
            .replace_persisted_record(self.source.document_persist_id, target_document.clone())?;
        let bytes = editor.finish()?;
        validate_unrelated_streams(self.source.bytes(), &bytes)?;
        let snapshot = Snapshot::from_arc(Arc::from(bytes), self.source.limits)?;
        if snapshot.document_persist_id != self.source.document_persist_id
            || snapshot.document.as_ref() != target_document
            || snapshot.fonts != self.fonts
        {
            return Err(Error::Corrupted(
                "published font candidate failed semantic reopen".into(),
            ));
        }
        let patch = Patch {
            base: self.source.revision(),
            target: snapshot.revision(),
            before: self.source.bytes,
            after: snapshot.bytes.clone(),
            changes: self.changes,
            limits: snapshot.limits,
        };
        Ok(Commit { snapshot, patch })
    }

    /// Consume the transaction, publishing its staged changes.
    ///
    /// # Errors
    ///
    /// Returns an error under the same conditions as [`Self::commit`].
    pub fn finish(self) -> Result<Commit> {
        self.commit()
    }
    #[must_use]
    pub fn rollback(self) -> Snapshot {
        self.source
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
}

impl Commit {
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }
    #[must_use]
    pub const fn fonts(&self) -> &FontCollections {
        self.snapshot.fonts()
    }
    /// Revert the committed patch against an exact current snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error if `current` is neither this commit's snapshot nor
    /// its exact base source.
    pub fn undo(&self, current: &Snapshot) -> Result<Snapshot> {
        self.patch.undo(current)
    }
    /// Re-apply the committed patch against an exact current snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error if `current` is neither the base source nor the exact
    /// already-applied target of the patch.
    pub fn redo(&self, current: &Snapshot) -> Result<Snapshot> {
        self.patch.redo(current)
    }
    #[must_use]
    pub fn into_parts(self) -> (Snapshot, Patch) {
        (self.snapshot, self.patch)
    }
}

fn missing_collection(scope: Scope) -> Error {
    Error::InvalidFormat(format!(
        "{scope:?} font collection is absent; owner creation is not losslessly provable"
    ))
}

fn preflight_publication_budget(
    source_bytes: usize,
    document_bytes: usize,
    maximum: usize,
) -> Result<()> {
    let projected = source_bytes
        .checked_add(document_bytes)
        .and_then(|value| value.checked_add(64))
        .ok_or_else(|| Error::ResourceLimit("PowerPoint publication size overflows".into()))?;
    if projected > maximum {
        return Err(Error::ResourceLimit(format!(
            "PowerPoint incremental publication requires at least {projected} bytes, exceeding the {maximum}-byte source limit"
        )));
    }
    Ok(())
}

fn collections_have_embedded_fonts(fonts: &FontCollections) -> bool {
    fonts
        .base
        .iter()
        .chain(fonts.international.iter())
        .flat_map(|collection| collection.fonts.iter())
        .any(|font| !font.embedded_fonts.is_empty())
}

fn synchronize_save_with_fonts(document: &mut Record, expected: bool) -> Result<()> {
    let mut document_atom_position = None;
    for (index, child) in document.children.iter().enumerate() {
        if child.record_type != crate::RecordType::DocumentAtom {
            continue;
        }
        if document_atom_position.replace(index).is_some() {
            return Err(Error::Corrupted(
                "live DocumentContainer contains multiple DocumentAtom records".into(),
            ));
        }
    }
    let Some(position) = document_atom_position else {
        return Err(Error::Corrupted(
            "live DocumentContainer is missing its DocumentAtom".into(),
        ));
    };
    let record = document.children.get_mut(position).ok_or_else(|| {
        Error::Corrupted("live DocumentContainer is missing its DocumentAtom".into())
    })?;
    let atom = crate::document_atom::DocumentAtom::parse(record)?;
    if atom.save_with_fonts != expected {
        record.data[36] = u8::from(expected);
    }
    Ok(())
}
