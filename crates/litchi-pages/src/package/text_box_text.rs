//! Implementation of the public [`crate::drawable::text_box_text`]
//! transaction: one character or paragraph property of the whole text a
//! body text box (or ordinary shape) owns (record 0413, the port of the
//! Numbers family of record 0378).
//!
//! The drawable's storage is uniformly styled by one paragraph style. Each
//! staged change runs the shared paragraph-style engine
//! ([`litchi_iwa_text_wire::paragraph_style`], record 0374) over that style,
//! counted exclusive when nothing but the storage references it and no
//! paragraph style names it as parent. The engine's plan is published here:
//!
//! - **Rewrite**: the storage's own variation takes the new payload in place;
//! - **Derive**: a new child variation of the storage's style is stored
//!   beside its stylesheet and registered there, and the storage points at
//!   it (in place in its paragraph-style table and its archive header);
//! - **Collapse**: the storage's variation no longer overrides anything, so
//!   the storage points back at the parent and the variation is removed with
//!   its stylesheet registration.
//!
//! This is the behaviour of the host editor's `set_text_box_text_*`,
//! `reset_text_box_text_*`, `set_text_box_paragraph_*` and
//! `reset_text_box_paragraph_*` families, collapsed into one property
//! verb.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "the focused transaction keeps its public types beside the staged rewrite"
)]

use std::fmt;

use litchi_iwa_common::WireLimits;
use litchi_iwa_core::{
    ArchiveObject, FieldInfo, FieldPath, FieldType, RawMessage, UnknownFieldRule,
};
use litchi_iwa_text_wire::paragraph_style::{
    self as engine, Change, ManagedFields, Plan, Property, PropertyKind, Styles, Target,
};
use thiserror::Error as ThisError;

use super::Package;
use super::text_box_structure::{
    self as structure, Fault, FaultResult, Published, Resolved, Working,
    metadata::{self, Transition},
    references,
};
use crate::selector::DrawableSelector;

structure::text_box_family!("text-box text", Wire(litchi_iwa_common::LimitKind));

/// `TSWP.ParagraphStyleArchive`.
const PARAGRAPH_STYLE_MESSAGE_TYPE: u32 = engine::MESSAGE_TYPE;
/// `TSS.StylesheetArchive`.
const STYLESHEET_MESSAGE_TYPE: u32 = 401;
/// `TSWP.StorageArchive`.
const STORAGE_MESSAGE_TYPE: u32 = 2_001;

/// A typed, content-redacted text-box text failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies
    /// it (record 0351 §2.4).
    #[error(transparent)]
    Refused(Refusal),
    /// The rooted source is malformed or ambiguous.
    #[error("rooted text-box text source is invalid")]
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
    #[error("could not allocate {amount} text-box text items or bytes")]
    Allocation {
        /// Requested amount.
        amount: usize,
    },
    /// The reopened candidate did not reproduce the requested state.
    #[error("text-box text candidate verification failed")]
    Verification,
    /// The patch was applied to an artifact other than its exact source.
    #[error("text-box text patch conflicts with this package")]
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
    /// The text is not uniformly styled by one paragraph style.
    #[error("the drawable's text has several paragraph styles")]
    MixedParagraphStyles,
    /// Commit was requested before staging a change.
    #[error("no text-box text change is staged")]
    NoStagedOperation,
    /// The style lives outside what this transaction rewrites (no
    /// stylesheet, a style in another component).
    #[error("a surviving object depends on the text's style graph")]
    UnsupportedDependency,
    /// The source cannot publish an exact changed artifact.
    #[error("source representation does not support text-box text properties")]
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
            Self::NoStagedOperation => RefusalKind::UnsupportedOperation,
            Self::MixedParagraphStyles | Self::UnsupportedDependency | Self::UnsupportedSource => {
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
        engine::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: LimitKind::Wire(kind),
            observed: observed as u64,
            maximum: maximum as u64,
        },
        engine::Error::Allocation { amount } => Error::Allocation { amount },
        _ => Error::InvalidSource,
    }
}

/// One staged change.
#[derive(Debug, Clone, PartialEq)]
enum Staged {
    Set(Property),
    Reset(PropertyKind),
}

/// Property changes staged against the text of one body text box or shape.
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

    /// Stage `property` as the effective value of the whole text. Changes
    /// apply in staging order.
    ///
    /// # Errors
    ///
    /// [`Error::Allocation`].
    pub fn set(&mut self, property: Property) -> Result<&mut Self, Error> {
        self.stage(Staged::Set(property))
    }

    /// Stage removing the text's own override of `kind`, so the value is
    /// inherited again.
    ///
    /// # Errors
    ///
    /// [`Error::Allocation`].
    pub fn reset(&mut self, kind: PropertyKind) -> Result<&mut Self, Error> {
        self.stage(Staged::Reset(kind))
    }

    /// Validate and atomically publish the staged changes, in order. A
    /// change that leaves a property as it is is skipped; an edit whose
    /// changes are all skipped is an exact no-op.
    ///
    /// # Errors
    ///
    /// [`Refusal::NoStagedOperation`]; [`Refusal::UnsupportedDependency`]
    /// for a style graph this transaction does not rewrite; source,
    /// resource and verification errors. Failure never publishes a partial
    /// package.
    pub fn commit(self) -> Result<Commit, Error> {
        if self.changes.is_empty() {
            return Err(Refusal::NoStagedOperation.into());
        }
        let mut current: Option<Package> = None;
        let mut touched = 0_usize;
        for change in &self.changes {
            let base = current.as_ref().unwrap_or(self.source);
            let change = match change {
                Staged::Set(property) => Change::Set(property.clone()),
                Staged::Reset(kind) => Change::Reset(*kind),
            };
            if let Some(published) = apply_one(base, self.selector, &change)? {
                touched = touched.saturating_add(published.touched_components);
                current = Some(published.package);
            }
        }
        let Some(package) = current else {
            return Ok(Commit::unchanged(self.source));
        };
        for (position, change) in self.changes.iter().enumerate() {
            if let Staged::Set(property) = change
                && !self.changes[position + 1..]
                    .iter()
                    .any(|later| match later {
                        Staged::Set(other) => other.kind() == property.kind(),
                        Staged::Reset(kind) => *kind == property.kind(),
                    })
                && &read(&package, self.selector, property.kind())? != property
            {
                return Err(Error::Verification);
            }
        }
        let commit = Commit::new(
            self.source,
            Published {
                package,
                touched_components: touched,
                deleted_previews: 0,
            },
        );
        Ok(commit)
    }
}

/// The drawable's storage: its object and storage message.
struct StorageTarget {
    storage: u64,
    index: usize,
}

fn storage_target(package: &Package, selector: DrawableSelector) -> Result<StorageTarget, Error> {
    let target = match structure::text_target(package, selector)? {
        Resolved::Text(target) => target,
        Resolved::NotFound => return Err(Refusal::DrawableNotFound { selector }.into()),
        Resolved::NoText => {
            return Err(Refusal::NoText {
                kind: super::body_drawable_graph::kind_name(selector),
            }
            .into());
        },
    };
    let storage_object = structure::object(package, target.storage)?;
    let index = structure::unique_message(storage_object, &[STORAGE_MESSAGE_TYPE])?;
    Ok(StorageTarget {
        storage: target.storage,
        index,
    })
}

/// The package's paragraph styles, as the engine reads them.
struct WorkingStyles<'w, 'p>(&'w Working<'p>);

impl Styles for WorkingStyles<'_, '_> {
    fn paragraph_style(&self, identifier: u64) -> Option<&[u8]> {
        let object = self.0.object(identifier).ok()?;
        object
            .messages
            .iter()
            .find(|message| message.type_ == PARAGRAPH_STYLE_MESSAGE_TYPE)
            .map(|message| message.data.as_slice())
    }
}

/// Styles read straight from a package (for reads).
struct PackageStyles<'p>(&'p Package);

impl Styles for PackageStyles<'_> {
    fn paragraph_style(&self, identifier: u64) -> Option<&[u8]> {
        let object = structure::object(self.0, identifier).ok()?;
        object
            .messages
            .iter()
            .find(|message| message.type_ == PARAGRAPH_STYLE_MESSAGE_TYPE)
            .map(|message| message.data.as_slice())
    }
}

fn wire_limits(maximum: usize) -> Result<WireLimits, Error> {
    WireLimits::default()
        .with_input_bytes(maximum.min(WireLimits::MAX_INPUT_BYTES))
        .and_then(|limits| limits.with_output_bytes(maximum.min(WireLimits::MAX_OUTPUT_BYTES)))
        .map_err(|_error| Error::InvalidSource)
}

fn uniform(payload: &[u8], limits: WireLimits) -> Result<u64, Error> {
    engine::uniform_style(payload, limits).map_err(|error| match error {
        engine::Error::InvalidSource { .. } => Refusal::MixedParagraphStyles.into(),
        other => from_engine(other),
    })
}

/// The effective `kind` of the drawable's text.
fn read(
    package: &Package,
    selector: DrawableSelector,
    kind: PropertyKind,
) -> Result<Property, Error> {
    let target = storage_target(package, selector)?;
    let limits = wire_limits(
        structure::exact_source(package)?
            .limits()
            .effective_archive_limits()
            .map_err(|error| structure::map_archive(&error))?
            .max_message_bytes(),
    )?;
    let object = structure::object(package, target.storage)?;
    let style = uniform(&object.messages[target.index].data, limits)?;
    engine::effective(&PackageStyles(package), style, kind, limits).map_err(from_engine)
}

/// The stylesheet a paragraph style names (`TSS.StyleArchive.stylesheet`).
fn stylesheet_of(working: &Working<'_>, style: u64, limits: WireLimits) -> Result<u64, Error> {
    let object = working.object(style)?;
    let index = structure::unique_message(object, &[PARAGRAPH_STYLE_MESSAGE_TYPE])?;
    let view = litchi_iwa_common::wire::WireView::parse(&object.messages[index].data, limits)?;
    let base = structure::unique_field(&view, 1, 2)?.ok_or(Error::InvalidSource)?;
    let records = references::records(base.payload(), &[5])?;
    match records.as_slice() {
        [record] => Ok(references::reference_identifier(record)?),
        [] => Err(Refusal::UnsupportedDependency.into()),
        _ => Err(Error::InvalidSource),
    }
}

impl From<litchi_iwa_common::Error> for Error {
    fn from(error: litchi_iwa_common::Error) -> Self {
        structure::map_wire(&error).into()
    }
}

/// Whether nothing but `storage`, the stylesheet and the style itself
/// references `style`, and no paragraph style names it as parent (the host
/// editor's storage-count and child checks).
fn exclusive(
    working: &Working<'_>,
    style: u64,
    stylesheet: u64,
    storage: u64,
    limits: WireLimits,
) -> bool {
    for object in working
        .source()
        .state
        .source
        .components()
        .iter()
        .flat_map(|component| component.archive().objects.iter())
    {
        let identifier = object.archive_info.identifier.unwrap_or_default();
        if identifier == style || identifier == stylesheet || identifier == storage {
            continue;
        }
        if object.archive_info.message_infos.iter().any(|info| {
            info.object_references.contains(&style)
                || info
                    .field_infos
                    .iter()
                    .any(|field| field.object_references.contains(&style))
        }) {
            return false;
        }
        for message in &object.messages {
            if message.type_ == PARAGRAPH_STYLE_MESSAGE_TYPE
                && litchi_iwa_text_wire::character_style_parent(&message.data, limits)
                    .ok()
                    .flatten()
                    == Some(style)
            {
                return false;
            }
        }
    }
    true
}

/// Bring `infos` in line with the variation's managed paths (the host
/// editor's `sync_managed_field_info`).
pub(super) fn sync_managed(
    infos: &mut Vec<FieldInfo>,
    managed: ManagedFields,
    limits: litchi_iwa_core::Limits,
) -> FaultResult<()> {
    for (present, path) in [
        (managed.text_fill(), [11, 46]),
        (managed.text_stroke(), [11, 44]),
        (managed.text_shadow(), [11, 21]),
        (managed.text_background(), [11, 26]),
    ] {
        let is_path = |field: &FieldInfo| field.path.as_slice() == path;
        if present {
            if !infos.iter().any(is_path) {
                let mut info = FieldInfo::new(
                    FieldPath::new(path.to_vec(), limits).map_err(|_| Fault::Invalid)?,
                    limits,
                )
                .map_err(|_| Fault::Invalid)?;
                info.r#type = Some(FieldType::Message);
                info.unknown_field_rule = Some(UnknownFieldRule::IgnoreAndPreserve);
                structure::reserve(infos, 1)?;
                infos.push(info);
            }
        } else {
            infos.retain(|field| !is_path(field));
        }
    }
    Ok(())
}

/// Replace message `index` of `object`, rewriting every header reference to
/// `from` into `to` in place (the storage's style repoint; the header is
/// re-encoded canonically, as the host editor's repoint does).
pub(super) fn replace_repointing(
    object: &mut ArchiveObject,
    index: usize,
    data: Vec<u8>,
    from: u64,
    to: u64,
    limits: litchi_iwa_core::Limits,
) -> FaultResult<()> {
    let info = object
        .archive_info
        .message_infos
        .get(index)
        .ok_or(Fault::Invalid)?;
    if info
        .object_references
        .iter()
        .filter(|value| **value == from)
        .count()
        != 1
        || info.object_references.contains(&to)
    {
        return Err(Fault::Invalid);
    }
    let message_type = object.messages.get(index).ok_or(Fault::Invalid)?.type_;
    let mut repointed = object.clone();
    repointed
        .replace_message(
            index,
            RawMessage {
                type_: message_type,
                data,
            },
            limits,
        )
        .map_err(|error| structure::map_core(&error))?;
    let info = &mut repointed.archive_info.message_infos[index];
    for reference in &mut info.object_references {
        if *reference == from {
            *reference = to;
        }
    }
    for field in &mut info.field_infos {
        for reference in &mut field.object_references {
            if *reference == from {
                *reference = to;
            }
        }
    }
    repointed
        .validate(limits)
        .map_err(|error| structure::map_core(&error))?;
    *object = repointed;
    Ok(())
}

/// Replace message `index` of object `identifier`, removing and appending
/// aggregate header references.
fn replace_with_references(
    working: &mut Working<'_>,
    identifier: u64,
    index: usize,
    data: Vec<u8>,
    removed: &[u64],
    appended: &[u64],
) -> FaultResult<()> {
    let limits = working.archive_limits();
    let object = working.object_mut(identifier)?;
    references::replace(object, index, data, removed, appended, &[], limits)
}

#[allow(
    clippy::too_many_lines,
    reason = "one staged change: plan, then the rewrite, derive or collapse publication"
)]
fn apply_one(
    source: &Package,
    selector: DrawableSelector,
    change: &Change,
) -> Result<Option<Published>, Error> {
    let target = storage_target(source, selector)?;
    let mut working = Working::new(source)?;
    working.keep_previews();
    let wire = wire_limits(working.archive_limits().max_message_bytes())?;
    let storage_payload = working.object(target.storage)?.messages[target.index]
        .data
        .clone();
    let style = uniform(&storage_payload, wire)?;
    let stylesheet = stylesheet_of(&working, style, wire)?;
    let is_exclusive = exclusive(&working, style, stylesheet, target.storage, wire);
    let plan = engine::plan(
        &WorkingStyles(&working),
        Target::new(style, is_exclusive),
        change,
        wire,
    )
    .map_err(from_engine)?;
    let registry =
        metadata::registry(&working)?.ok_or(Error::Refused(Refusal::UnsupportedSource))?;
    let storage_member = working
        .locate(target.storage)?
        .ok_or(Error::InvalidSource)?;
    let mut transition = Transition::default();
    match plan {
        Plan::Unchanged => return Ok(None),
        Plan::Rewrite(variation) => {
            let index =
                structure::unique_message(working.object(style)?, &[PARAGRAPH_STYLE_MESSAGE_TYPE])?;
            let before = working.object(style)?.archive_info.message_infos[index]
                .object_references
                .clone();
            let after = variation.references().to_vec();
            // The header keeps every reference it had, as the host editor's
            // in-place rewrite does: a following style replaced by "same
            // style" leaves its named style referenced (record 0413 §3).
            let removed: [u64; 0] = [];
            let appended: Vec<u64> = after
                .iter()
                .copied()
                .filter(|reference| !before.contains(reference))
                .collect();
            let managed = variation.managed_fields();
            replace_with_references(
                &mut working,
                style,
                index,
                variation.into_payload(),
                &removed,
                &appended,
            )?;
            let limits = working.archive_limits();
            let object = working.object_mut(style)?;
            sync_managed(
                &mut object.archive_info.message_infos[index].field_infos,
                managed,
                limits,
            )?;
        },
        Plan::Derive(variation) => {
            let stylesheet_member = working.locate(stylesheet)?.ok_or(Error::InvalidSource)?;
            if working.locate(style)?.as_deref() != Some(stylesheet_member.as_str()) {
                return Err(Refusal::UnsupportedDependency.into());
            }
            let identifier = *working.allocate(1)?.first().ok_or(Error::InvalidSource)?;
            let managed = variation.managed_fields();
            let references = variation.references().to_vec();
            let mut object = structure::new_object(
                identifier,
                None,
                PARAGRAPH_STYLE_MESSAGE_TYPE,
                variation.into_payload(),
                references,
                working.archive_limits(),
            )?;
            object.archive_info.message_infos[0].versions = engine::MESSAGE_VERSION.to_vec().into();
            sync_managed(
                &mut object.archive_info.message_infos[0].field_infos,
                managed,
                working.archive_limits(),
            )?;
            working.insert(&stylesheet_member, object)?;
            let sheet_object = working.object(stylesheet)?;
            let sheet_index = structure::unique_message(sheet_object, &[STYLESHEET_MESSAGE_TYPE])?;
            let sheet_data = engine::stylesheet_with_variation(
                &sheet_object.messages[sheet_index].data,
                style,
                identifier,
                wire,
            )
            .map_err(from_engine)?;
            replace_with_references(
                &mut working,
                stylesheet,
                sheet_index,
                sheet_data,
                &[],
                &[identifier],
            )?;
            let repointed =
                engine::repoint(&storage_payload, style, identifier, wire).map_err(from_engine)?;
            let limits = working.archive_limits();
            replace_repointing(
                working.object_mut(target.storage)?,
                target.index,
                repointed,
                style,
                identifier,
                limits,
            )?;
            let stylesheet_component = registry.component_for_entry(&stylesheet_member)?.identifier;
            let storage_component = registry.component_for_entry(&storage_member)?.identifier;
            let seed = metadata::identity_seed(&working, identifier)?;
            transition.register_uuid(&registry, stylesheet_component, identifier, seed)?;
            transition.reference(
                &registry,
                storage_component,
                stylesheet_component,
                identifier,
            )?;
        },
        Plan::Collapse { parent, stylesheet } => {
            let repointed =
                engine::repoint(&storage_payload, style, parent, wire).map_err(from_engine)?;
            let limits = working.archive_limits();
            replace_repointing(
                working.object_mut(target.storage)?,
                target.index,
                repointed,
                style,
                parent,
                limits,
            )?;
            let sheet_object = working.object(stylesheet)?;
            let sheet_index = structure::unique_message(sheet_object, &[STYLESHEET_MESSAGE_TYPE])?;
            let sheet_data = engine::stylesheet_without_variation(
                &sheet_object.messages[sheet_index].data,
                parent,
                style,
                wire,
            )
            .map_err(from_engine)?;
            replace_with_references(
                &mut working,
                stylesheet,
                sheet_index,
                sheet_data,
                &[style],
                &[],
            )?;
            let parent_member = working.locate(parent)?.ok_or(Error::InvalidSource)?;
            let storage_component = registry.component_for_entry(&storage_member)?.identifier;
            let parent_component = registry.component_for_entry(&parent_member)?.identifier;
            transition.remove_objects(&registry, &[style])?;
            transition.reference(&registry, storage_component, parent_component, parent)?;
            transition.release_suffix(&[style])?;
            working.remove(style)?;
        },
        _ => return Err(Error::InvalidSource),
    }
    Ok(Some(working.publish(&transition)?))
}

impl Package {
    /// The effective value of `kind` for the whole text of one body text
    /// box or shape (its paragraph style and that style's parents).
    ///
    /// # Errors
    ///
    /// [`Refusal::DrawableNotFound`], [`Refusal::NoText`],
    /// [`Refusal::MixedParagraphStyles`] and source errors.
    pub fn body_text_box_text(
        &self,
        drawable: impl Into<DrawableSelector>,
        kind: PropertyKind,
    ) -> Result<Property, Error> {
        read(self, drawable.into(), kind)
    }

    /// Start an immutable character and paragraph property edit of the
    /// text of one body text box or shape.
    ///
    /// ```no_run
    /// use litchi_pages::{DrawableSelector, Package};
    /// use litchi_pages::drawable::text_box_text::PropertyKind;
    ///
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let package = Package::open("input.pages")?;
    /// let mut edit = package.edit_body_text_box_text(DrawableSelector::text_box(0))?;
    /// edit.reset(PropertyKind::Alignment)?;
    /// let commit = edit.commit()?;
    /// let restored = commit.package().apply_body_text_box_text(&commit.patch().inverse())?;
    /// # let _ = restored;
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// # Errors
    ///
    /// [`Refusal::DrawableNotFound`], [`Refusal::NoText`] and source errors.
    pub fn edit_body_text_box_text(
        &self,
        drawable: impl Into<DrawableSelector>,
    ) -> Result<Edit<'_>, Error> {
        let selector = drawable.into();
        storage_target(self, selector)?;
        Ok(Edit {
            source: self,
            selector,
            changes: Vec::new(),
        })
    }

    /// Apply a text-box text patch to its exact source package.
    ///
    /// # Errors
    ///
    /// [`Error::PatchConflict`] when this package is not the patch's exact
    /// source.
    pub fn apply_body_text_box_text(&self, patch: &Patch) -> Result<Commit, Error> {
        apply_patch(self, patch)
    }
}
