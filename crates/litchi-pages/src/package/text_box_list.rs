//! Implementation of the public [`crate::drawable::list`] transactions:
//! paragraph lists of the text a body text box or shape owns (record 0413,
//! the port of the Numbers family of record 0378).
//!
//! The list change runs W1-D's engine ([`litchi_iwa_text_wire::list`],
//! record 0374) and is published by the shared storage-list publisher
//! (`litchi_iwa_text_wire::list::publish`, the record 0370b publisher of the
//! Numbers cell-list family, lifted into the text crate by record 0413);
//! `litchi_iwa_table::structure::list_workspace` is its working-set side.
//! Host `set_x`/`reset_x` pairs collapse into one property verb
//! ([`Edit::set`]/[`Edit::reset`]); `reset_text_box_paragraph_list`
//! and `reset_text_box_paragraph_list_level` are [`Edit::set_list`]
//! with [`ParagraphList::None`] and [`Edit::set_level`] with
//! [`ParagraphListLevel::ZERO`].

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "the focused transaction keeps its public types beside the staged rewrite"
)]

use std::fmt;

use litchi_iwa_common::WireLimits;
use litchi_iwa_text::paragraph::list::{
    ParagraphList, ParagraphListLevel, ParagraphListLevelPlacement, ParagraphListNumbering,
    ParagraphListPlacement,
};
use litchi_iwa_text::position::TextPosition;
use litchi_iwa_text_wire::list::{self, Change, Property, PropertyKind};
use thiserror::Error as ThisError;

use litchi_iwa_table::structure::list_workspace::{ListWorkspace, ReadStyles};
use litchi_iwa_text_wire::list::publish::{self, Failure, Publisher};

use super::Package;
use super::text_box_structure::{
    self as structure, Fault, Published, Resolved, Working,
    metadata::{self, Transition},
};
use crate::selector::DrawableSelector;

structure::text_box_family!("text-box list", Wire(litchi_iwa_common::LimitKind));

/// A typed, content-redacted text-box list failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies
    /// it (record 0351 §2.4).
    #[error(transparent)]
    Refused(Refusal),
    /// The rooted source is malformed or ambiguous.
    #[error("rooted text-box list source is invalid")]
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
    #[error("could not allocate {amount} text-box list items or bytes")]
    Allocation {
        /// Requested amount.
        amount: usize,
    },
    /// The reopened candidate did not reproduce the requested state.
    #[error("text-box list candidate verification failed")]
    Verification,
    /// The patch was applied to an artifact other than its exact source.
    #[error("text-box list patch conflicts with this package")]
    PatchConflict,
}

/// A well-formed request this family's route refuses, classified by
/// [`Refusal::kind`] (record 0351 §2.4).
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    /// No drawable matches the selector.
    #[error("the body has no drawable at {selector:?}")]
    DrawableNotFound {
        /// The selector that did not resolve.
        selector: DrawableSelector,
    },
    /// The drawable owns no text (it is not a text box or shape).
    #[error("a Pages body {kind} owns no text")]
    NoText {
        /// The selected kind.
        kind: &'static str,
    },
    /// Commit was requested before staging a change.
    #[error("no text-box list change is staged")]
    NoStagedOperation,
    /// The position is not the start of a paragraph.
    #[error("position is not a paragraph start")]
    NotAParagraphStart,
    /// The paragraph's list kind has no such property.
    #[error("the paragraph's list has no such property")]
    NotApplicable,
    /// Placements must start at zero and strictly increase.
    #[error("list placements must start at zero and strictly increase")]
    InvalidPlacements,
    /// The paragraphs do not share one list.
    #[error("the paragraphs do not share one list")]
    NotUniform,
    /// The text's styles live outside what this transaction rewrites.
    #[error("a surviving object depends on the text's list graph")]
    UnsupportedDependency,
    /// The source cannot publish an exact changed artifact.
    #[error("source representation does not support text-box lists")]
    UnsupportedSource,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::DrawableNotFound { .. } => RefusalKind::NotFound,
            Self::NoText { .. } => RefusalKind::UnsupportedTarget,
            Self::NoStagedOperation | Self::NotApplicable | Self::NotUniform => {
                RefusalKind::UnsupportedOperation
            },
            Self::NotAParagraphStart | Self::InvalidPlacements => RefusalKind::InvalidValue,
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

fn from_list(error: list::Error) -> Error {
    match error {
        list::Error::Refused(list::Refusal::NotAParagraphStart) => {
            Refusal::NotAParagraphStart.into()
        },
        list::Error::Refused(list::Refusal::NotApplicable) => Refusal::NotApplicable.into(),
        list::Error::Refused(list::Refusal::InvalidPlacements) => Refusal::InvalidPlacements.into(),
        list::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: LimitKind::Wire(kind),
            observed: observed as u64,
            maximum: maximum as u64,
        },
        list::Error::Allocation { amount } => Error::Allocation { amount },
        _ => Error::InvalidSource,
    }
}

impl From<Failure<Fault>> for Error {
    fn from(failure: Failure<Fault>) -> Self {
        match failure {
            Failure::Workspace(fault) => fault.into(),
            Failure::List(error) => from_list(error),
        }
    }
}

/// One staged change.
#[derive(Debug, Clone, PartialEq)]
enum Staged {
    List(ParagraphList),
    Lists(Vec<ParagraphListPlacement>),
    Level(TextPosition, ParagraphListLevel),
    Numbering(TextPosition, ParagraphListNumbering),
    Set(TextPosition, Property),
    Reset(TextPosition, PropertyKind),
}

/// List changes staged against the text of one body text box or shape.
pub struct Edit<'a> {
    source: &'a Package,
    selector: DrawableSelector,
    changes: Vec<Staged>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("selector", &self.selector)
            .field("changes", &self.changes.len())
            .finish_non_exhaustive()
    }
}

impl Edit<'_> {
    fn stage(&mut self, change: Staged) -> Result<&mut Self, Error> {
        structure::reserve(&mut self.changes, 1)?;
        self.changes.push(change);
        Ok(self)
    }

    /// Stage giving every paragraph the list preset `list`
    /// ([`ParagraphList::None`] removes the lists).
    ///
    /// # Errors
    ///
    /// [`Error::Allocation`].
    pub fn set_list(&mut self, list: ParagraphList) -> Result<&mut Self, Error> {
        self.stage(Staged::List(list))
    }

    /// Stage giving the paragraphs the presets of `placements`: each
    /// placement's list runs from its paragraph to the next placement.
    ///
    /// # Errors
    ///
    /// [`Error::Allocation`].
    pub fn set_lists(
        &mut self,
        placements: impl IntoIterator<Item = ParagraphListPlacement>,
    ) -> Result<&mut Self, Error> {
        let mut staged = Vec::new();
        for placement in placements {
            structure::reserve(&mut staged, 1)?;
            staged.push(placement);
        }
        self.stage(Staged::Lists(staged))
    }

    /// Stage the list level of the paragraph starting at `paragraph`
    /// ([`ParagraphListLevel::ZERO`] resets it).
    ///
    /// # Errors
    ///
    /// [`Error::Allocation`].
    pub fn set_level(
        &mut self,
        paragraph: TextPosition,
        level: ParagraphListLevel,
    ) -> Result<&mut Self, Error> {
        self.stage(Staged::Level(paragraph, level))
    }

    /// Stage whether the paragraph starting at `paragraph` continues or
    /// restarts its list's numbering.
    ///
    /// # Errors
    ///
    /// [`Error::Allocation`].
    pub fn set_numbering(
        &mut self,
        paragraph: TextPosition,
        numbering: ParagraphListNumbering,
    ) -> Result<&mut Self, Error> {
        self.stage(Staged::Numbering(paragraph, numbering))
    }

    /// Stage `property` as the effective list property of the paragraph
    /// starting at `paragraph`.
    ///
    /// # Errors
    ///
    /// [`Error::Allocation`].
    pub fn set(&mut self, paragraph: TextPosition, property: Property) -> Result<&mut Self, Error> {
        self.stage(Staged::Set(paragraph, property))
    }

    /// Stage restoring the list's own value of `kind` for the paragraph
    /// starting at `paragraph`.
    ///
    /// # Errors
    ///
    /// [`Error::Allocation`].
    pub fn reset(
        &mut self,
        paragraph: TextPosition,
        kind: PropertyKind,
    ) -> Result<&mut Self, Error> {
        self.stage(Staged::Reset(paragraph, kind))
    }

    /// Validate and atomically publish the staged changes, in order. A
    /// change that leaves the text as it is is skipped; an edit whose
    /// changes are all skipped is an exact no-op.
    ///
    /// # Errors
    ///
    /// [`Refusal::NoStagedOperation`]; the list refusals; source, resource
    /// and verification errors.
    pub fn commit(self) -> Result<Commit, Error> {
        if self.changes.is_empty() {
            return Err(Refusal::NoStagedOperation.into());
        }
        let mut current: Option<Package> = None;
        let mut touched = 0_usize;
        for change in &self.changes {
            let base = current.as_ref().unwrap_or(self.source);
            if let Some(published) = apply_one(base, self.selector, change)? {
                verify(&published.package, self.selector, change)?;
                touched = touched.saturating_add(published.touched_components);
                current = Some(published.package);
            }
        }
        let Some(package) = current else {
            return Ok(Commit::unchanged(self.source));
        };
        Ok(Commit::new(
            self.source,
            Published {
                package,
                touched_components: touched,
                deleted_previews: 0,
            },
        ))
    }
}

fn storage_of(package: &Package, selector: DrawableSelector) -> Result<u64, Error> {
    match structure::text_target(package, selector)? {
        Resolved::Text(target) => Ok(target.storage),
        Resolved::NoText => Err(Refusal::NoText {
            kind: super::body_drawable_graph::kind_name(selector),
        }
        .into()),
        Resolved::NotFound => Err(Refusal::DrawableNotFound { selector }.into()),
    }
}

fn wire_limits(working: &Working<'_>) -> Result<WireLimits, Error> {
    let maximum = working.archive_limits().max_message_bytes();
    WireLimits::default()
        .with_input_bytes(maximum.min(WireLimits::MAX_INPUT_BYTES))
        .and_then(|limits| limits.with_output_bytes(maximum.min(WireLimits::MAX_OUTPUT_BYTES)))
        .map_err(|_error| Error::InvalidSource)
}

/// Reads of one drawable's text lists.
struct Reader<'p> {
    working: Working<'p>,
    storage: u64,
    limits: WireLimits,
}

impl<'p> Reader<'p> {
    fn new(package: &'p Package, selector: DrawableSelector) -> Result<Self, Error> {
        let storage = storage_of(package, selector)?;
        let working = Working::new(package)?;
        let limits = wire_limits(&working)?;
        Ok(Self {
            working,
            storage,
            limits,
        })
    }

    fn payload(&self) -> Result<Vec<u8>, Error> {
        let object = self.working.object(self.storage)?;
        let index = structure::unique_message(object, &[publish::STORAGE_MESSAGE_TYPE])?;
        Ok(structure::copy_bytes(&object.messages[index].data)?)
    }

    fn list(&self) -> Result<ParagraphList, Error> {
        let style =
            list::uniform_style(&self.payload()?, self.limits).map_err(|error| match error {
                list::Error::InvalidSource { .. } => Refusal::NotUniform.into(),
                other => from_list(other),
            })?;
        list::preset(&ReadStyles(&self.working), style, self.limits).map_err(from_list)
    }

    fn lists(&self) -> Result<Vec<ParagraphListPlacement>, Error> {
        list::placements(&self.payload()?, &ReadStyles(&self.working), self.limits)
            .map_err(from_list)
    }

    fn levels(&self) -> Result<Vec<ParagraphListLevelPlacement>, Error> {
        list::levels(&self.payload()?, self.limits).map_err(from_list)
    }

    fn level(&self, paragraph: TextPosition) -> Result<ParagraphListLevel, Error> {
        list::level(&self.payload()?, paragraph, self.limits).map_err(from_list)
    }

    fn numbering(&self, paragraph: TextPosition) -> Result<ParagraphListNumbering, Error> {
        list::numbering(&self.payload()?, paragraph, self.limits).map_err(from_list)
    }

    fn property(&self, paragraph: TextPosition, kind: PropertyKind) -> Result<Property, Error> {
        list::property(
            &self.payload()?,
            &ReadStyles(&self.working),
            paragraph,
            kind,
            self.limits,
        )
        .map_err(from_list)
    }
}

fn verify(package: &Package, selector: DrawableSelector, change: &Staged) -> Result<(), Error> {
    let reader = Reader::new(package, selector)?;
    let holds = match change {
        Staged::List(preset) => reader.list()? == *preset,
        Staged::Lists(_) | Staged::Reset(..) => true,
        Staged::Level(paragraph, level) => reader.level(*paragraph)? == *level,
        Staged::Numbering(paragraph, numbering) => reader.numbering(*paragraph)? == *numbering,
        Staged::Set(paragraph, property) => {
            reader.property(*paragraph, property.kind())? == *property
        },
    };
    if holds {
        Ok(())
    } else {
        Err(Error::Verification)
    }
}

fn apply_one(
    source: &Package,
    selector: DrawableSelector,
    change: &Staged,
) -> Result<Option<Published>, Error> {
    let storage = storage_of(source, selector)?;
    let mut working = Working::new(source)?;
    working.keep_previews();
    let limits = wire_limits(&working)?;
    let registry =
        metadata::registry(&working)?.ok_or(Error::Refused(Refusal::UnsupportedSource))?;
    let mut transition = Transition::default();
    let mut workspace = ListWorkspace {
        working: &mut working,
        registry: &registry,
        transition: &mut transition,
    };
    let mut publisher = Publisher::new(&mut workspace, limits);
    let changed = match change {
        Staged::List(preset) => publisher.set_list(storage, *preset)?,
        Staged::Lists(placements) => publisher.set_lists(storage, placements)?,
        Staged::Level(paragraph, level) => publisher.set_level(storage, *paragraph, *level)?,
        Staged::Numbering(paragraph, numbering) => {
            publisher.set_numbering(storage, *paragraph, *numbering)?
        },
        Staged::Set(paragraph, property) => {
            publisher.property(storage, *paragraph, &Change::Set(property.clone()))?
        },
        Staged::Reset(paragraph, kind) => {
            publisher.property(storage, *paragraph, &Change::Reset(*kind))?
        },
    };
    if !changed {
        return Ok(None);
    }
    Ok(Some(working.publish(&transition)?))
}

impl Package {
    /// The list preset every paragraph of one body text box or shape shares.
    ///
    /// # Errors
    ///
    /// Selector errors, [`Refusal::NoText`], [`Refusal::NotUniform`], source
    /// errors.
    pub fn body_text_box_list(
        &self,
        drawable: impl Into<DrawableSelector>,
    ) -> Result<ParagraphList, Error> {
        Reader::new(self, drawable.into())?.list()
    }

    /// The list preset of each run of paragraphs of one body text box or shape.
    ///
    /// # Errors
    ///
    /// As [`Self::body_text_box_list`], without [`Refusal::NotUniform`].
    pub fn body_text_box_lists(
        &self,
        drawable: impl Into<DrawableSelector>,
    ) -> Result<Vec<ParagraphListPlacement>, Error> {
        Reader::new(self, drawable.into())?.lists()
    }

    /// The list level of each run of paragraphs of one body text box or shape.
    ///
    /// # Errors
    ///
    /// As [`Self::body_text_box_lists`].
    pub fn body_text_box_list_levels(
        &self,
        drawable: impl Into<DrawableSelector>,
    ) -> Result<Vec<ParagraphListLevelPlacement>, Error> {
        Reader::new(self, drawable.into())?.levels()
    }

    /// The list level of the paragraph starting at `paragraph`.
    ///
    /// # Errors
    ///
    /// As [`Self::body_text_box_lists`], [`Refusal::NotAParagraphStart`].
    pub fn body_text_box_list_level(
        &self,
        drawable: impl Into<DrawableSelector>,
        paragraph: TextPosition,
    ) -> Result<ParagraphListLevel, Error> {
        Reader::new(self, drawable.into())?.level(paragraph)
    }

    /// Whether the paragraph starting at `paragraph` continues or restarts
    /// its list's numbering.
    ///
    /// # Errors
    ///
    /// As [`Self::body_text_box_list_level`].
    pub fn body_text_box_list_numbering(
        &self,
        drawable: impl Into<DrawableSelector>,
        paragraph: TextPosition,
    ) -> Result<ParagraphListNumbering, Error> {
        Reader::new(self, drawable.into())?.numbering(paragraph)
    }

    /// The effective list property `kind` of the paragraph starting at
    /// `paragraph`.
    ///
    /// # Errors
    ///
    /// As [`Self::body_text_box_list_level`], [`Refusal::NotApplicable`].
    pub fn body_text_box_list_property(
        &self,
        drawable: impl Into<DrawableSelector>,
        paragraph: TextPosition,
        kind: PropertyKind,
    ) -> Result<Property, Error> {
        Reader::new(self, drawable.into())?.property(paragraph, kind)
    }

    /// Start an immutable paragraph-list edit of the text of one body text
    /// box or shape.
    ///
    /// ```no_run
    /// use litchi_pages::drawable::list::ParagraphList;
    /// use litchi_pages::{DrawableSelector, Package};
    ///
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let package = Package::open("input.pages")?;
    /// let mut edit = package.edit_body_text_box_lists(DrawableSelector::text_box(0))?;
    /// edit.set_list(ParagraphList::Bullet)?;
    /// let commit = edit.commit()?;
    /// # let _ = commit;
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// # Errors
    ///
    /// [`Refusal::DrawableNotFound`], [`Refusal::NoText`], source errors.
    pub fn edit_body_text_box_lists(
        &self,
        drawable: impl Into<DrawableSelector>,
    ) -> Result<Edit<'_>, Error> {
        let selector = drawable.into();
        storage_of(self, selector)?;
        Ok(Edit {
            source: self,
            selector,
            changes: Vec::new(),
        })
    }

    /// Apply a text-box list patch to its exact source package.
    ///
    /// # Errors
    ///
    /// [`Error::PatchConflict`] when this package is not the patch's exact
    /// source.
    pub fn apply_body_text_box_lists(&self, patch: &Patch) -> Result<Commit, Error> {
        apply_patch(self, patch)
    }
}
