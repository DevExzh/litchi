//! Implementation of the public [`crate::sheet::drawable::language`]
//! transaction: the text-language runs of the storage a sheet text box or
//! shape owns, on the shared language engine
//! ([`litchi_iwa_text_wire::language`], record 0374), record 0378.
//!
//! Each change rewrites the storage's language table in place; every other
//! byte of the storage and every other object stays as it is. This is the
//! behaviour of the host editor's `set_sheet_text_box_text_language`,
//! `remove_sheet_text_box_text_language_boundary` and
//! `reset_sheet_text_box_text_languages` routes.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "the focused transaction keeps its public types beside the staged rewrite"
)]

use std::fmt;

use litchi_iwa_common::WireLimits;
use litchi_iwa_core::RawMessage;
use litchi_iwa_text::TextPosition;
use litchi_iwa_text::language::{TextLanguage, TextLanguageRun};
use litchi_iwa_text_wire::language::{self as engine, Change};
use thiserror::Error as ThisError;

use super::Package;
use super::sheet_drawable_graph::{self as graph, SHAPE_INFO_MESSAGE_TYPE};
use super::structure::{self, Fault, Published, references};
use crate::selector::SheetSelector;
use crate::sheet::drawable::{DrawableSelector, Kind};

structure::structural_family!("text-language");

/// `TSWP.ShapeInfoArchive.owned_storage`.
const OWNED_STORAGE_FIELD: u32 = 4;
/// `TSWP.StorageArchive`.
const STORAGE_MESSAGE_TYPE: u32 = 2_001;

/// A typed, content-redacted text-language failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies
    /// it (record 0351 §2.4).
    #[error(transparent)]
    Refused(Refusal),
    /// The rooted drawable graph or its language table is malformed.
    #[error("rooted text-language source is invalid")]
    InvalidSource,
    /// A finite transaction limit was exceeded.
    #[error("{kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        /// Resource category.
        kind: LimitKind,
        /// Observed amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded allocation failed.
    #[error("could not allocate {amount} text-language items or bytes")]
    Allocation {
        /// Requested amount.
        amount: usize,
    },
    /// The reopened candidate did not reproduce the requested runs.
    #[error("text-language candidate verification failed")]
    Verification,
    /// The patch was applied to an artifact other than its exact source.
    #[error("text-language patch conflicts with this package")]
    PatchConflict,
}

/// A well-formed request this family's route refuses, classified by
/// [`Refusal::kind`] (record 0351 §2.4).
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    /// The sheet selector did not resolve one sheet.
    #[error("sheet selector did not resolve")]
    SheetNotFound,
    /// No drawable matches the selector.
    #[error("the sheet has no drawable at {selector:?}")]
    DrawableNotFound {
        /// The selector that did not resolve.
        selector: DrawableSelector,
    },
    /// The drawable owns no text (it is not a text box or shape).
    #[error("a Numbers sheet {kind} owns no text")]
    NoText {
        /// The selected kind.
        kind: &'static str,
    },
    /// The position is past the text or splits a surrogate pair.
    #[error("the position is not a text boundary")]
    NotABoundary,
    /// The index-zero boundary goes only with [`Edit::reset`].
    #[error("the first language boundary cannot be removed")]
    FirstBoundary,
    /// Commit was requested before staging a change.
    #[error("no text-language change is staged")]
    NoStagedOperation,
    /// A surviving object depends on the rewritten graph.
    #[error("a surviving object depends on the text")]
    UnsupportedDependency,
    /// The source cannot publish an exact changed artifact.
    #[error("source representation does not support text-language edits")]
    UnsupportedSource,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::SheetNotFound | Self::DrawableNotFound { .. } => RefusalKind::NotFound,
            Self::NoText { .. } => RefusalKind::UnsupportedTarget,
            Self::NotABoundary | Self::FirstBoundary => RefusalKind::InvalidValue,
            Self::NoStagedOperation => RefusalKind::UnsupportedOperation,
            Self::UnsupportedDependency | Self::UnsupportedSource => {
                RefusalKind::UnsupportedContent
            },
        }
    }
}

impl From<Refusal> for Error {
    fn from(refusal: Refusal) -> Self {
        Self::Refused(refusal)
    }
}

fn from_engine(error: engine::Error) -> Error {
    match error {
        engine::Error::Refused(engine::Refusal::FirstBoundary) => Refusal::FirstBoundary.into(),
        engine::Error::Refused(_) => Refusal::NotABoundary.into(),
        engine::Error::Allocation { amount } => Error::Allocation { amount },
        _ => Error::InvalidSource,
    }
}

/// Language changes staged against the text of one sheet drawable.
pub struct Edit<'a> {
    source: &'a Package,
    sheet: usize,
    selector: DrawableSelector,
    changes: Vec<Change>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("sheet", &self.sheet)
            .field("selector", &self.selector)
            .field("changes", &self.changes.len())
            .finish_non_exhaustive()
    }
}

impl Edit<'_> {
    fn stage(&mut self, change: Change) -> Result<&mut Self, Error> {
        structure::reserve(&mut self.changes, 1)?;
        self.changes.push(change);
        Ok(self)
    }

    /// Stage `language` effective from `position` to the next boundary.
    ///
    /// # Errors
    ///
    /// [`Error::Allocation`]; boundary refusals surface at commit.
    pub fn set(
        &mut self,
        position: TextPosition,
        language: TextLanguage,
    ) -> Result<&mut Self, Error> {
        self.stage(Change::Set { position, language })
    }

    /// Stage removing the boundary at `position` (not zero), so the text
    /// inherits the preceding run.
    ///
    /// # Errors
    ///
    /// [`Error::Allocation`].
    pub fn remove_boundary(&mut self, position: TextPosition) -> Result<&mut Self, Error> {
        self.stage(Change::RemoveBoundary(position))
    }

    /// Stage removing every explicit language.
    ///
    /// # Errors
    ///
    /// [`Error::Allocation`].
    pub fn reset(&mut self) -> Result<&mut Self, Error> {
        self.stage(Change::Reset)
    }

    /// Validate and publish the staged changes in order; an edit that
    /// changes nothing is an exact no-op.
    ///
    /// # Errors
    ///
    /// [`Refusal::NoStagedOperation`], [`Refusal::NotABoundary`],
    /// [`Refusal::FirstBoundary`], source and verification errors.
    pub fn commit(self) -> Result<Commit, Error> {
        if self.changes.is_empty() {
            return Err(Refusal::NoStagedOperation.into());
        }
        let (storage, index) = storage_target(self.source, self.sheet, self.selector)?;
        let object = graph::object(self.source, storage)?;
        let mut payload = object.messages[index].data.clone();
        let mut changed = false;
        for change in &self.changes {
            if let Some(rewritten) =
                engine::rewrite(&payload, change, WireLimits::DEFAULT).map_err(from_engine)?
            {
                payload = rewritten;
                changed = true;
            }
        }
        if !changed || payload == object.messages[index].data {
            return Commit::unchanged(self.source);
        }
        let expected = engine::runs(&payload, WireLimits::DEFAULT).map_err(from_engine)?;
        let mut data = Some(payload);
        let mut replace = |object: &mut litchi_iwa_core::ArchiveObject,
                           limits: litchi_iwa_core::Limits| {
            object
                .replace_message_preserving_header(
                    index,
                    RawMessage {
                        type_: STORAGE_MESSAGE_TYPE,
                        data: data.take().ok_or(Fault::Invalid)?,
                    },
                    limits,
                )
                .map(drop)
                .map_err(|error| structure::map_core(&error))
        };
        let published = graph::publish_in_place(self.source, &mut [(storage, &mut replace)])?;
        if runs(&published.package, self.sheet, self.selector)? != expected {
            return Err(Error::Verification);
        }
        Commit::new(
            self.source,
            Published {
                package: published.package,
                touched_components: published.touched_components,
                deleted_previews: 0,
            },
        )
    }
}

fn storage_target(
    package: &Package,
    sheet: usize,
    selector: DrawableSelector,
) -> Result<(u64, usize), Error> {
    let drawables = graph::sheet_drawables(package, sheet)?;
    let entry = drawables
        .select(selector)
        .ok_or(Refusal::DrawableNotFound { selector })?;
    if !matches!(entry.kind, Kind::TextBox | Kind::Shape) {
        return Err(Refusal::NoText {
            kind: entry.kind.name(),
        }
        .into());
    }
    let (index, message_type) = entry.message.ok_or(Error::InvalidSource)?;
    if message_type != SHAPE_INFO_MESSAGE_TYPE {
        return Err(Error::InvalidSource);
    }
    let object = graph::object(package, entry.identifier)?;
    let records = references::records(&object.messages[index].data, &[OWNED_STORAGE_FIELD])?;
    let [record] = records.as_slice() else {
        return Err(Error::InvalidSource);
    };
    let storage = references::reference_identifier(record)?;
    let storage_object = graph::object(package, storage)?;
    let index = structure::unique_message(storage_object, &[STORAGE_MESSAGE_TYPE])?;
    Ok((storage, index))
}

fn runs(
    package: &Package,
    sheet: usize,
    selector: DrawableSelector,
) -> Result<Vec<TextLanguageRun>, Error> {
    let (storage, index) = storage_target(package, sheet, selector)?;
    let object = graph::object(package, storage)?;
    engine::runs(&object.messages[index].data, WireLimits::DEFAULT).map_err(from_engine)
}

fn resolve_sheet(package: &Package, selector: SheetSelector<'_>) -> Result<usize, Error> {
    graph::sheet_position(package, selector)?.ok_or_else(|| Refusal::SheetNotFound.into())
}

impl Package {
    /// The explicit language runs of the text a sheet text box or shape
    /// owns, in text order (empty when the text uses automatic language
    /// selection).
    ///
    /// # Errors
    ///
    /// Selector errors, [`Refusal::NoText`] and source errors.
    pub fn sheet_drawable_text_languages<'sheet>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        drawable: impl Into<DrawableSelector>,
    ) -> Result<Vec<TextLanguageRun>, Error> {
        let sheet = resolve_sheet(self, sheet.into())?;
        runs(self, sheet, drawable.into())
    }

    /// The language effective at `position` of the text a sheet text box or
    /// shape owns.
    ///
    /// # Errors
    ///
    /// As [`Self::sheet_drawable_text_languages`], and
    /// [`Refusal::NotABoundary`].
    pub fn sheet_drawable_text_language<'sheet>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        drawable: impl Into<DrawableSelector>,
        position: TextPosition,
    ) -> Result<TextLanguage, Error> {
        let sheet = resolve_sheet(self, sheet.into())?;
        let (storage, index) = storage_target(self, sheet, drawable.into())?;
        let object = graph::object(self, storage)?;
        engine::at(&object.messages[index].data, position, WireLimits::DEFAULT).map_err(from_engine)
    }

    /// Start an immutable language edit of the text of one sheet text box
    /// or shape.
    ///
    /// ```no_run
    /// use litchi_numbers::Package;
    /// use litchi_numbers::sheet::drawable::DrawableSelector;
    ///
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let package = Package::open("input.numbers")?;
    /// let mut edit = package.edit_sheet_drawable_text_languages(0, DrawableSelector::text_box(0))?;
    /// edit.reset()?;
    /// let commit = edit.commit()?;
    /// # let _ = commit;
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// # Errors
    ///
    /// Selector errors, [`Refusal::NoText`] and source errors.
    pub fn edit_sheet_drawable_text_languages<'sheet>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        drawable: impl Into<DrawableSelector>,
    ) -> Result<Edit<'_>, Error> {
        let sheet = resolve_sheet(self, sheet.into())?;
        let selector = drawable.into();
        storage_target(self, sheet, selector)?;
        Ok(Edit {
            source: self,
            sheet,
            selector,
            changes: Vec::new(),
        })
    }

    /// Apply a text-language patch to its exact source package.
    ///
    /// # Errors
    ///
    /// [`Error::PatchConflict`] when this package is not the patch's exact
    /// source.
    pub fn apply_sheet_drawable_text_languages(&self, patch: &Patch) -> Result<Commit, Error> {
        apply_patch(self, patch)
    }
}
