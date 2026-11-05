//! Implementation of the public [`crate::drawable::date_time`]
//! transaction: the date-time field of the text a body text box or shape owns,
//! on W1-D's ranged smart-field engine
//! ([`litchi_iwa_text_wire::{ranged, smart_field}`](litchi_iwa_text_wire::ranged),
//! record 0374), record 0378.
//!
//! A date-time field is a `TSWP.DateTimeSmartFieldArchive` object stored beside the
//! storage and named by a range of the storage's smart-field table. Adding
//! one stores the object, its table entry and the storage's header
//! reference and raises the identifier watermark; moving or retargeting
//! rewrites the entry and the object in place; removing drops both and
//! forgets the object's registrations, as the host text editor does.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "the focused transaction keeps its public types beside the staged rewrite"
)]

use std::fmt;

use litchi_iwa_common::WireLimits;
use litchi_iwa_core::{ArchiveObject, RawMessage};
use litchi_iwa_text::date_time::{DisplayText, Settings};
use litchi_iwa_text::position::{TextPosition, TextRange};
use litchi_iwa_text_wire::ranged::{self, Table};
use litchi_iwa_text_wire::smart_field::{self, TextAttributeUuid};
use litchi_iwa_text_wire::{RewriteBehavior, RewriteLimits, StorageRewriteExecutionLimits};
use thiserror::Error as ThisError;

use super::Package;
use super::text_box_structure::{
    self as structure, Fault, FaultResult, Published, Resolved, Working,
    metadata::{self, Transition},
};
use crate::selector::DrawableSelector;

structure::text_box_family!("text date-time field", Wire(litchi_iwa_common::LimitKind));

/// `TSWP.StorageArchive`.
const STORAGE_MESSAGE_TYPE: u32 = 2_001;

/// A typed, content-redacted date-time field failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies
    /// it (record 0351 §2.4).
    #[error(transparent)]
    Refused(Refusal),
    /// The rooted source is malformed or ambiguous.
    #[error("rooted date-time field source is invalid")]
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
    #[error("could not allocate {amount} date-time field items or bytes")]
    Allocation {
        /// Requested amount.
        amount: usize,
    },
    /// The reopened candidate did not reproduce the requested state.
    #[error("date-time field candidate verification failed")]
    Verification,
    /// The patch was applied to an artifact other than its exact source.
    #[error("date-time field patch conflicts with this package")]
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
    /// No date-time field is at the selected position.
    #[error("the text has no date-time field at position {position}")]
    FieldNotFound {
        /// The selected zero-based position in text order.
        position: usize,
    },
    /// A range bound is past the text or splits a surrogate pair.
    #[error("the range is not on text boundaries")]
    NotABoundary,
    /// The range overlaps another field's range.
    #[error("the range overlaps another field")]
    Occupied,
    /// Commit was requested before staging a change.
    #[error("no date-time field change is staged")]
    NoStagedOperation,
    /// A surviving object depends on the rewritten graph.
    #[error("a surviving object depends on the date-time field")]
    UnsupportedDependency,
    /// The source cannot publish an exact changed artifact.
    #[error("source representation does not support date-time field edits")]
    UnsupportedSource,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::DrawableNotFound { .. } | Self::FieldNotFound { .. } => RefusalKind::NotFound,
            Self::NoText { .. } => RefusalKind::UnsupportedTarget,
            Self::NotABoundary | Self::Occupied => RefusalKind::InvalidValue,
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

fn from_ranged(error: ranged::Error) -> Error {
    match error {
        ranged::Error::Refused(ranged::Refusal::Occupied) => Refusal::Occupied.into(),
        ranged::Error::Refused(ranged::Refusal::NotABoundary) => Refusal::NotABoundary.into(),
        ranged::Error::Refused(_) => Refusal::UnsupportedDependency.into(),
        ranged::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: LimitKind::Wire(kind),
            observed: observed as u64,
            maximum: maximum as u64,
        },
        ranged::Error::Allocation { amount } => Error::Allocation { amount },
        _ => Error::InvalidSource,
    }
}

fn from_field(error: smart_field::Error) -> Error {
    match error {
        smart_field::Error::Allocation { amount } => Error::Allocation { amount },
        _ => Error::InvalidSource,
    }
}

/// One date-time field of a drawable's text.
#[derive(Debug, Clone, PartialEq)]
pub struct DateTimeField {
    range: TextRange,
    settings: Settings,
}

impl DateTimeField {
    /// The linked text range.
    #[must_use]
    pub const fn range(&self) -> TextRange {
        self.range
    }

    /// The link settings.
    #[must_use]
    pub const fn settings(&self) -> &Settings {
        &self.settings
    }
}

#[derive(Debug, Clone, PartialEq)]
enum Staged {
    Add(TextRange, Settings),
    Insert(TextPosition, DisplayText, Settings),
    Update(usize, TextRange, Settings),
    Remove(usize),
}

/// Date & Time field changes staged against the text of one body text box or shape.
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

    /// Stage a new date-time field over `range`.
    ///
    /// # Errors
    ///
    /// [`Error::Allocation`]; range refusals surface at commit.
    pub fn add(&mut self, range: TextRange, settings: Settings) -> Result<&mut Self, Error> {
        self.stage(Staged::Add(range, settings))
    }

    /// Stage inserting `display` at `position` and a new field over it.
    ///
    /// # Errors
    ///
    /// [`Error::Allocation`].
    pub fn insert(
        &mut self,
        position: TextPosition,
        display: DisplayText,
        settings: Settings,
    ) -> Result<&mut Self, Error> {
        self.stage(Staged::Insert(position, display, settings))
    }

    /// Stage moving the date-time field at zero-based `position` (in text order,
    /// as [`crate::Package::body_drawable_date_time_fields`] lists them, after
    /// the earlier changes) to `range` with `settings`.
    ///
    /// # Errors
    ///
    /// [`Error::Allocation`].
    pub fn update(
        &mut self,
        position: usize,
        range: TextRange,
        settings: Settings,
    ) -> Result<&mut Self, Error> {
        self.stage(Staged::Update(position, range, settings))
    }

    /// Stage removing the date-time field at zero-based `position`.
    ///
    /// # Errors
    ///
    /// [`Error::Allocation`].
    pub fn remove(&mut self, position: usize) -> Result<&mut Self, Error> {
        self.stage(Staged::Remove(position))
    }

    /// Validate and atomically publish the staged changes, in order.
    ///
    /// # Errors
    ///
    /// [`Refusal::NoStagedOperation`], the range refusals, source and
    /// verification errors.
    pub fn commit(self) -> Result<Commit, Error> {
        if self.changes.is_empty() {
            return Err(Refusal::NoStagedOperation.into());
        }
        let mut current: Option<Package> = None;
        let mut touched = 0_usize;
        for change in &self.changes {
            let base = current.as_ref().unwrap_or(self.source);
            if let Some(published) = apply_one(base, self.selector, change)? {
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

/// The storage's date-time field with their objects, in text order.
fn fields(package: &Package, storage: u64) -> Result<Vec<(u64, DateTimeField)>, Error> {
    let object = structure::object(package, storage)?;
    let payload =
        &object.messages[structure::unique_message(object, &[STORAGE_MESSAGE_TYPE])?].data;
    let mut found = Vec::new();
    for entry in
        ranged::entries(payload, Table::SmartField, WireLimits::DEFAULT).map_err(from_ranged)?
    {
        let Ok(field) = structure::object(package, entry.object()) else {
            continue;
        };
        let Some(message) = field
            .messages
            .iter()
            .find(|message| message.type_ == smart_field::DATE_TIME_MESSAGE_TYPE)
        else {
            continue;
        };
        let (_, settings) =
            smart_field::date_time(&message.data, WireLimits::DEFAULT).map_err(from_field)?;
        found.push((
            entry.object(),
            DateTimeField {
                range: entry.range(),
                settings,
            },
        ));
    }
    Ok(found)
}

/// Replace the storage message canonically and add or remove one header
/// reference (the host's ranged-table bookkeeping).
fn replace_storage(
    object: &mut ArchiveObject,
    index: usize,
    data: Vec<u8>,
    added: Option<u64>,
    removed: Option<u64>,
    limits: litchi_iwa_core::Limits,
) -> FaultResult<()> {
    let mut staged = object.clone();
    staged
        .replace_message(
            index,
            RawMessage {
                type_: STORAGE_MESSAGE_TYPE,
                data,
            },
            limits,
        )
        .map_err(|error| structure::map_core(&error))?;
    let info = staged
        .archive_info
        .message_infos
        .get_mut(index)
        .ok_or(Fault::Invalid)?;
    if let Some(identifier) = added {
        if info.object_references.contains(&identifier) {
            return Err(Fault::Invalid);
        }
        structure::reserve(&mut info.object_references, 1)?;
        info.object_references.push(identifier);
    }
    if let Some(identifier) = removed {
        if info
            .object_references
            .iter()
            .filter(|candidate| **candidate == identifier)
            .count()
            != 1
        {
            return Err(Fault::Invalid);
        }
        info.object_references
            .retain(|candidate| *candidate != identifier);
        for field in &mut info.field_infos {
            field
                .object_references
                .retain(|candidate| *candidate != identifier);
        }
    }
    staged
        .validate(limits)
        .map_err(|error| structure::map_core(&error))?;
    *object = staged;
    Ok(())
}

/// A deterministic hyphenated text-attribute UUID for a new field.
fn field_uuid(working: &Working<'_>, object: u64) -> Result<TextAttributeUuid, Error> {
    let seed = metadata::identity_seed(working, object)?;
    let low = metadata::mix(seed ^ object);
    let high = metadata::mix(low ^ 0x6c69_7463_6869_d378);
    let text = format!(
        "{:08x}-{:04x}-4{:03x}-{:04x}-{:012x}",
        high >> 32,
        (high >> 16) & 0xffff,
        high & 0x0fff,
        0x8000 | ((low >> 48) & 0x3fff),
        low & 0xffff_ffff_ffff
    );
    TextAttributeUuid::new(&text).map_err(from_field)
}

/// Store a new field over `range` of the storage payload `payload`.
fn add_field(
    working: &mut Working<'_>,
    storage: u64,
    index: usize,
    member: &str,
    payload: Vec<u8>,
    range: TextRange,
    settings: &Settings,
) -> Result<(), Error> {
    let limits = WireLimits::DEFAULT;
    let archive_limits = working.archive_limits();
    let object = *working.allocate(1)?.first().ok_or(Error::InvalidSource)?;
    let rewrite = ranged::rewrite(
        &payload,
        Table::SmartField,
        &ranged::Change::Add { object, range },
        limits,
    )
    .map_err(from_ranged)?;
    let (added, removed) = (rewrite.added(), rewrite.removed());
    replace_storage(
        working.object_mut(storage)?,
        index,
        rewrite.into_payload(),
        added,
        removed,
        archive_limits,
    )?;
    let uuid = field_uuid(working, object)?;
    let field = smart_field::encode_date_time(&uuid, settings, limits).map_err(from_field)?;
    let created = ArchiveObject::new(
        object,
        vec![RawMessage {
            type_: smart_field::DATE_TIME_MESSAGE_TYPE,
            data: field,
        }],
        archive_limits,
    )
    .map_err(|error| structure::map_core(&error))?;
    working.insert(member, created)?;
    Ok(())
}

fn apply_one(
    source: &Package,
    selector: DrawableSelector,
    change: &Staged,
) -> Result<Option<Published>, Error> {
    let storage = storage_of(source, selector)?;
    let existing = fields(source, storage)?;
    let mut working = Working::new(source)?;
    working.keep_previews();
    let limits = WireLimits::DEFAULT;
    let archive_limits = working.archive_limits();
    let member = working.locate(storage)?.ok_or(Error::InvalidSource)?;
    let index = structure::unique_message(working.object(storage)?, &[STORAGE_MESSAGE_TYPE])?;
    let payload = working.object(storage)?.messages[index].data.clone();
    let mut transition = Transition::default();
    let resolve = |position: usize| {
        existing
            .get(position)
            .map(|(object, link)| (*object, link.clone()))
            .ok_or(Error::Refused(Refusal::FieldNotFound { position }))
    };
    match change {
        Staged::Add(range, settings) => {
            add_field(
                &mut working,
                storage,
                index,
                &member,
                payload,
                *range,
                settings,
            )?;
        },
        Staged::Insert(position, display, settings) => {
            let start =
                usize::try_from(position.utf16_index()).map_err(|_error| Error::InvalidSource)?;
            let end = start + display.as_str().encode_utf16().count();
            let prepared = litchi_iwa_text_wire::prepare_storage_text_rewrite_with_behavior(
                &payload,
                start..start,
                display.as_str(),
                RewriteBehavior::ReplaceSelection,
                RewriteLimits::default(),
            )
            .map_err(|_error| Error::Refused(Refusal::NotABoundary))?;
            let requirements = prepared.execution_requirements();
            let inserted = prepared
                .execute(StorageRewriteExecutionLimits {
                    max_output_bytes: requirements.output_bytes(),
                    max_retained_elements: requirements.retained_elements(),
                    max_retained_bytes: requirements.retained_bytes(),
                    max_peak_scratch_bytes: requirements.peak_scratch_bytes(),
                    max_allocations: requirements.allocations(),
                    max_work: requirements.work(),
                })
                .map_err(|_error| Error::InvalidSource)?;
            if !inserted.removed_object_references().is_empty() {
                return Err(Refusal::UnsupportedDependency.into());
            }
            let range =
                TextRange::from_utf16_indexes(start, end).map_err(|_error| Error::InvalidSource)?;
            add_field(
                &mut working,
                storage,
                index,
                &member,
                inserted.into_bytes(),
                range,
                settings,
            )?;
        },
        Staged::Update(position, range, settings) => {
            let (object, current) = resolve(*position)?;
            if current.range == *range && current.settings == *settings {
                return Ok(None);
            }
            let rewrite = ranged::rewrite(
                &payload,
                Table::SmartField,
                &ranged::Change::Move {
                    object,
                    range: *range,
                },
                limits,
            )
            .map_err(from_ranged)?;
            replace_storage(
                working.object_mut(storage)?,
                index,
                rewrite.into_payload(),
                None,
                None,
                archive_limits,
            )?;
            let field = working.object(object)?;
            let field_index =
                structure::unique_message(field, &[smart_field::DATE_TIME_MESSAGE_TYPE])?;
            let data = smart_field::date_time_with_settings(
                &field.messages[field_index].data,
                settings,
                limits,
            )
            .map_err(from_field)?;
            let mut staged = working.object(object)?.clone();
            staged
                .replace_message(
                    field_index,
                    RawMessage {
                        type_: smart_field::DATE_TIME_MESSAGE_TYPE,
                        data,
                    },
                    archive_limits,
                )
                .map_err(|error| Error::from(structure::map_core(&error)))?;
            *working.object_mut(object)? = staged;
        },
        Staged::Remove(position) => {
            let (object, _) = resolve(*position)?;
            let rewrite = ranged::rewrite(
                &payload,
                Table::SmartField,
                &ranged::Change::Remove { object },
                limits,
            )
            .map_err(from_ranged)?;
            let removed = rewrite.removed();
            replace_storage(
                working.object_mut(storage)?,
                index,
                rewrite.into_payload(),
                None,
                removed,
                archive_limits,
            )?;
            if let Some(registry) = metadata::registry(&working)? {
                transition.remove_objects(&registry, &[object])?;
            }
            transition.release_suffix(&[object])?;
            working.remove(object)?;
        },
    }
    let published = working.publish(&transition)?;
    Ok(Some(published))
}

impl Package {
    /// The date-time field of the text a body text box or shape owns, in text
    /// order.
    ///
    /// # Errors
    ///
    /// Selector errors, [`Refusal::NoText`], source errors.
    pub fn body_drawable_date_time_fields(
        &self,
        drawable: impl Into<DrawableSelector>,
    ) -> Result<Vec<DateTimeField>, Error> {
        let storage = storage_of(self, drawable.into())?;
        Ok(fields(self, storage)?
            .into_iter()
            .map(|(_, link)| link)
            .collect())
    }

    /// Start an immutable date-time field edit of the text of one body text box
    /// or shape.
    ///
    /// # Errors
    ///
    /// Selector errors, [`Refusal::NoText`], source errors.
    pub fn edit_body_drawable_date_time_fields(
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

    /// Apply a date-time field patch to its exact source package.
    ///
    /// # Errors
    ///
    /// [`Error::PatchConflict`] when this package is not the patch's exact
    /// source.
    pub fn apply_body_drawable_date_time_fields(&self, patch: &Patch) -> Result<Commit, Error> {
        apply_patch(self, patch)
    }
}
