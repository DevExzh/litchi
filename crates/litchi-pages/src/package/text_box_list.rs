//! Implementation of the public [`crate::drawable::list`] transactions:
//! paragraph lists of the text a body text box or shape owns (record 0413,
//! the port of the Numbers family of record 0378).
//!
//! The list change runs W1-D's engine ([`litchi_iwa_text_wire::list`],
//! record 0374) and is published by the shared storage-list publisher
//! (`litchi_iwa_text_wire::list::publish`, the record 0370b publisher of the
//! Numbers cell-list family, lifted into the text crate by record 0413);
//! [`ListWorkspace`] is its Pages working-set side.
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

use litchi_iwa_core::{ArchiveObject, RawMessage};
use litchi_iwa_text_wire::list::publish::{self, Failure, Publisher, Workspace};

use super::Package;
use super::text_box_structure::{
    self as structure, Fault, FaultResult, Published, Resolved, Working,
    metadata::{self, Registry, Transition},
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

/// The list styles of a working set, as the read planners see them.
struct ReadStyles<'w, 'p>(&'w Working<'p>);

impl list::Styles for ReadStyles<'_, '_> {
    fn list_style(&self, identifier: u64) -> Option<&[u8]> {
        let object = self.0.object(identifier).ok()?;
        object
            .messages
            .iter()
            .find(|message| message.type_ == list::MESSAGE_TYPE)
            .map(|message| message.data.as_slice())
    }
}

/// The Pages side of the shared list publisher: the structural working
/// set, the source registry and the registry transition it accumulates.
struct ListWorkspace<'a, 'p> {
    working: &'a mut Working<'p>,
    registry: &'a Registry,
    transition: &'a mut Transition,
}

impl ListWorkspace<'_, '_> {
    fn info(
        &mut self,
        identifier: u64,
        index: usize,
    ) -> FaultResult<&mut litchi_iwa_core::MessageInfo> {
        self.working
            .object_mut(identifier)?
            .archive_info
            .message_infos
            .get_mut(index)
            .ok_or(Fault::Invalid)
    }
}

/// The host's `replace_reference_sequence` over one reference list.
fn sequence(references: &mut Vec<u64>, old: &[u64], new: &[u64]) -> FaultResult<()> {
    let positions = references
        .iter()
        .filter(|reference| old.contains(reference))
        .count();
    if positions != old.len() {
        return Err(Fault::Invalid);
    }
    let insertion = references
        .iter()
        .position(|reference| old.contains(reference))
        .unwrap_or(references.len());
    references.retain(|reference| !old.contains(reference));
    structure::reserve(references, new.len())?;
    references.splice(insertion..insertion, new.iter().copied());
    Ok(())
}

impl Workspace for ListWorkspace<'_, '_> {
    type Error = Fault;

    fn entry(&self, identifier: u64) -> FaultResult<String> {
        self.working.locate(identifier)?.ok_or(Fault::Invalid)
    }

    fn message(&self, identifier: u64, message_type: u32) -> FaultResult<(usize, &[u8])> {
        let object = self.working.object(identifier)?;
        let index = structure::unique_message(object, &[message_type])?;
        Ok((index, object.messages[index].data.as_slice()))
    }

    fn replace(&mut self, identifier: u64, index: usize, payload: Vec<u8>) -> FaultResult<()> {
        let limits = self.working.archive_limits();
        let object = self.working.object_mut(identifier)?;
        let message_type = object.messages.get(index).ok_or(Fault::Invalid)?.type_;
        object
            .replace_message_preserving_header(
                index,
                RawMessage {
                    type_: message_type,
                    data: payload,
                },
                limits,
            )
            .map(drop)
            .map_err(|error| structure::map_core(&error))
    }

    fn replace_reference_sequence(
        &mut self,
        identifier: u64,
        index: usize,
        old: &[u64],
        new: &[u64],
    ) -> FaultResult<()> {
        let info = self.info(identifier, index)?;
        sequence(&mut info.object_references, old, new)?;
        for field in &mut info.field_infos {
            if field
                .object_references
                .iter()
                .any(|reference| old.contains(reference))
            {
                sequence(&mut field.object_references, old, new)?;
            }
        }
        Ok(())
    }

    fn append_reference(
        &mut self,
        identifier: u64,
        index: usize,
        reference: u64,
    ) -> FaultResult<()> {
        let info = self.info(identifier, index)?;
        if info.object_references.contains(&reference) {
            return Err(Fault::Invalid);
        }
        structure::reserve(&mut info.object_references, 1)?;
        info.object_references.push(reference);
        Ok(())
    }

    fn remove_reference(
        &mut self,
        identifier: u64,
        index: usize,
        reference: u64,
    ) -> FaultResult<()> {
        let info = self.info(identifier, index)?;
        info.object_references.retain(|item| *item != reference);
        for field in &mut info.field_infos {
            field.object_references.retain(|item| *item != reference);
        }
        Ok(())
    }

    fn allocate(&mut self) -> FaultResult<u64> {
        self.working
            .allocate(1)?
            .first()
            .copied()
            .ok_or(Fault::Invalid)
    }

    fn insert(
        &mut self,
        entry: &str,
        identifier: u64,
        message_type: u32,
        version: &[u32],
        payload: Vec<u8>,
        references: Vec<u64>,
    ) -> FaultResult<()> {
        let mut object: ArchiveObject = structure::new_object(
            identifier,
            None,
            message_type,
            payload,
            references,
            self.working.archive_limits(),
        )?;
        object.archive_info.message_infos[0].versions = version.to_vec().into();
        self.working.insert(entry, object)
    }

    fn remove(&mut self, identifier: u64) -> FaultResult<()> {
        self.working.remove(identifier).map(drop)
    }

    fn entry_messages(&self, entry: &str, message_type: u32) -> FaultResult<Vec<(u64, &[u8])>> {
        let archive = self.working.archive(entry)?;
        Ok(archive
            .objects
            .iter()
            .filter_map(|object| {
                let identifier = object.archive_info.identifier?;
                object
                    .messages
                    .iter()
                    .find(|message| message.type_ == message_type)
                    .map(|message| (identifier, message.data.as_slice()))
            })
            .collect())
    }

    fn entry_references(&self, entry: &str, reference: u64) -> FaultResult<bool> {
        let archive = self.working.archive(entry)?;
        Ok(archive.objects.iter().any(|object| {
            object
                .archive_info
                .message_infos
                .iter()
                .any(|info| info.object_references.contains(&reference))
        }))
    }

    fn all_messages(&self, message_type: u32) -> FaultResult<Vec<&[u8]>> {
        let mut output = Vec::new();
        for name in self.working.component_names() {
            let archive = self.working.archive(&name)?;
            for object in &archive.objects {
                for message in &object.messages {
                    if message.type_ == message_type {
                        structure::reserve(&mut output, 1)?;
                        output.push(message.data.as_slice());
                    }
                }
            }
        }
        Ok(output)
    }

    fn component(&self, entry: &str) -> Option<u64> {
        self.registry
            .component_for_entry(entry)
            .ok()
            .map(|component| component.identifier)
    }

    fn add_external_reference(&mut self, owner: u64, target: u64, object: u64) -> FaultResult<()> {
        self.transition
            .reference(self.registry, owner, target, object)
    }

    fn remove_external_reference(
        &mut self,
        owner: u64,
        target: u64,
        object: u64,
    ) -> FaultResult<()> {
        self.transition
            .remove_reference(self.registry, owner, target, object)
    }

    fn register_uuid(&mut self, component: u64, object: u64) -> FaultResult<()> {
        let seed = metadata::identity_seed(self.working, object)?;
        self.transition
            .register_uuid(self.registry, component, object, seed)
    }

    fn unregister(&mut self, object: u64) -> FaultResult<()> {
        self.transition.remove_objects(self.registry, &[object])?;
        self.transition.release_suffix(&[object])
    }

    fn invalid() -> Fault {
        Fault::Invalid
    }
}
