//! Implementation of the public [`crate::sheet::drawable::named_style`]
//! transaction: the named paragraph styles of the text a sheet text box or
//! shape owns, on W1-D's engine
//! ([`litchi_iwa_text_wire::paragraph_style::named`], record 0374), record
//! 0378.
//!
//! The catalog is the union of the paragraph presets of the themes whose
//! document stylesheet is the storage's. The engine plans each change; this
//! module publishes it as the host text editor does:
//!
//! - **apply** repoints the storage to the named style and removes its
//!   exclusive override chain (with each variation's stylesheet entry, UUID
//!   and references), or registers the new cross-component reference and
//!   releases the old one when the storage's component no longer uses it;
//! - **create** copies a named style under a new name beside its stylesheet,
//!   registers it in the stylesheet and every listing theme, and registers its
//!   UUID and the themes' references;
//! - **rename** rewrites the style's name;
//! - **delete** unregisters a style nothing but its stylesheet and themes
//!   names, everywhere;
//! - **delete applied** applies a replacement and deletes the style;
//! - **redefine** folds the storage's overrides into its named style before
//!   applying it.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "the focused transaction keeps its public types beside the staged rewrite"
)]

use std::collections::BTreeMap;
use std::fmt;

use litchi_iwa_common::WireLimits;
use litchi_iwa_core::{ArchiveObject, RawMessage};
use litchi_iwa_text::paragraph::style::raw::native_id;
use litchi_iwa_text::paragraph::style::{
    NamedParagraphStyle, ParagraphStyleId, ParagraphStyleName,
};
use litchi_iwa_text_wire::paragraph_style::named::{self, ApplyPlan, Catalog};
use litchi_iwa_text_wire::paragraph_style::{self as engine, Refusal as EngineRefusal};
use thiserror::Error as ThisError;

use super::Package;
use super::sheet_drawable_graph::{self as graph, TextTarget};
use super::sheet_drawable_text::{replace_repointing, sync_managed};
use super::structure::{
    self, Fault, FaultResult, Published, Working,
    metadata::{self, Registry, Transition},
    references,
};
use crate::selector::SheetSelector;
use crate::sheet::drawable::DrawableSelector;

structure::structural_family!("named-style", Wire(litchi_iwa_common::LimitKind));

/// `TSWP.StorageArchive`.
const STORAGE_MESSAGE_TYPE: u32 = 2_001;
/// `TSS.StylesheetArchive`.
const STYLESHEET_MESSAGE_TYPE: u32 = 401;
/// `TSWP.StorageArchive.table_para_style`.
const STORAGE_PARAGRAPH_TABLE: u32 = 5;

/// A typed, content-redacted named-style failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies
    /// it (record 0351 §2.4).
    #[error(transparent)]
    Refused(Refusal),
    /// The rooted source is malformed or ambiguous.
    #[error("rooted named-style source is invalid")]
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
    #[error("could not allocate {amount} named-style items or bytes")]
    Allocation {
        /// Requested amount.
        amount: usize,
    },
    /// The reopened candidate did not reproduce the requested state.
    #[error("named-style candidate verification failed")]
    Verification,
    /// The patch was applied to an artifact other than its exact source.
    #[error("named-style patch conflicts with this package")]
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
    /// The text's paragraphs do not share one paragraph style.
    #[error("the text's paragraphs use more than one paragraph style")]
    MixedParagraphStyles,
    /// The style is not a named style of the text's stylesheet.
    #[error("the style is not a named style of the text's stylesheet")]
    StyleNotFound,
    /// Another named style already has the requested name.
    #[error("another named style already has that name")]
    NameTaken,
    /// The themes must keep at least one named paragraph style.
    #[error("the themes must keep at least one named paragraph style")]
    LastStyle,
    /// The text has no overrides on its applied named style.
    #[error("the text has no overrides on its applied named style")]
    NoOverrides,
    /// The text does not use the style.
    #[error("the text does not use the named style")]
    NotApplied,
    /// The replacement is the deleted style.
    #[error("the replacement is the deleted style")]
    SameReplacement,
    /// Another object still uses the style.
    #[error("another object still uses the named style")]
    StyleInUse,
    /// Commit was requested before staging a change.
    #[error("no named-style change is staged")]
    NoStagedOperation,
    /// A surviving object depends on the rewritten graph.
    #[error("a surviving object depends on the named-style graph")]
    UnsupportedDependency,
    /// The source cannot publish an exact changed artifact.
    #[error("source representation does not support named-style edits")]
    UnsupportedSource,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::SheetNotFound | Self::DrawableNotFound { .. } | Self::StyleNotFound => {
                RefusalKind::NotFound
            },
            Self::NoText { .. } => RefusalKind::UnsupportedTarget,
            Self::NameTaken | Self::SameReplacement | Self::NotApplied | Self::NoOverrides => {
                RefusalKind::InvalidValue
            },
            Self::NoStagedOperation | Self::LastStyle => RefusalKind::UnsupportedOperation,
            Self::MixedParagraphStyles
            | Self::StyleInUse
            | Self::UnsupportedDependency
            | Self::UnsupportedSource => RefusalKind::UnsupportedContent,
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
        engine::Error::Refused(refusal) => match refusal {
            EngineRefusal::NotNamed => Refusal::StyleNotFound,
            EngineRefusal::NameTaken => Refusal::NameTaken,
            EngineRefusal::LastStyle => Refusal::LastStyle,
            EngineRefusal::NoOverrides => Refusal::NoOverrides,
            EngineRefusal::NotApplied => Refusal::NotApplied,
            EngineRefusal::SameReplacement => Refusal::SameReplacement,
            _ => Refusal::UnsupportedDependency,
        }
        .into(),
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

impl From<litchi_iwa_common::Error> for Error {
    fn from(error: litchi_iwa_common::Error) -> Self {
        structure::map_wire(&error).into()
    }
}

/// One staged named-style change.
#[derive(Debug, Clone)]
enum Change {
    Apply(ParagraphStyleId),
    Create(ParagraphStyleId, ParagraphStyleName),
    Rename(ParagraphStyleId, ParagraphStyleName),
    Delete(ParagraphStyleId),
    DeleteApplied(ParagraphStyleId, ParagraphStyleId),
    Redefine,
}

/// Named-style changes staged against the text of one sheet drawable.
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

    /// Stage applying named style `style` to the whole text, clearing its
    /// direct paragraph overrides.
    ///
    /// # Errors
    ///
    /// [`Error::Allocation`].
    pub fn apply(&mut self, style: ParagraphStyleId) -> Result<&mut Self, Error> {
        self.stage(Change::Apply(style))
    }

    /// Stage creating a named style called `name` as a copy of `source`.
    ///
    /// # Errors
    ///
    /// [`Error::Allocation`].
    pub fn create(
        &mut self,
        source: ParagraphStyleId,
        name: ParagraphStyleName,
    ) -> Result<&mut Self, Error> {
        self.stage(Change::Create(source, name))
    }

    /// Stage renaming named style `style` to `name`.
    ///
    /// # Errors
    ///
    /// [`Error::Allocation`].
    pub fn rename_style(
        &mut self,
        style: ParagraphStyleId,
        name: ParagraphStyleName,
    ) -> Result<&mut Self, Error> {
        self.stage(Change::Rename(style, name))
    }

    /// Stage deleting named style `style`, which nothing may still use.
    ///
    /// # Errors
    ///
    /// [`Error::Allocation`].
    pub fn delete(&mut self, style: ParagraphStyleId) -> Result<&mut Self, Error> {
        self.stage(Change::Delete(style))
    }

    /// Stage applying `replacement` to the text and deleting its applied
    /// named style `style`.
    ///
    /// # Errors
    ///
    /// [`Error::Allocation`].
    pub fn delete_applied(
        &mut self,
        style: ParagraphStyleId,
        replacement: ParagraphStyleId,
    ) -> Result<&mut Self, Error> {
        self.stage(Change::DeleteApplied(style, replacement))
    }

    /// Stage folding the text's direct overrides into its applied named
    /// style.
    ///
    /// # Errors
    ///
    /// [`Error::Allocation`].
    pub fn redefine(&mut self) -> Result<&mut Self, Error> {
        self.stage(Change::Redefine)
    }

    /// Validate and atomically publish the staged changes, in order; an
    /// edit that changes nothing is an exact no-op.
    ///
    /// # Errors
    ///
    /// [`Refusal::NoStagedOperation`], the catalog refusals
    /// ([`Refusal::StyleNotFound`], [`Refusal::NameTaken`],
    /// [`Refusal::LastStyle`], [`Refusal::NoOverrides`],
    /// [`Refusal::NotApplied`], [`Refusal::SameReplacement`],
    /// [`Refusal::StyleInUse`]), source and verification errors.
    pub fn commit(self) -> Result<Commit, Error> {
        if self.changes.is_empty() {
            return Err(Refusal::NoStagedOperation.into());
        }
        let mut current: Option<Package> = None;
        let mut touched = 0_usize;
        for change in &self.changes {
            let base = current.as_ref().unwrap_or(self.source);
            let storage = storage_of(base, self.sheet, self.selector)?;
            if let Some(published) = apply_one(base, storage, change)? {
                verify(&published.package, storage, change)?;
                touched = touched.saturating_add(published.touched_components);
                current = Some(published.package);
            }
        }
        let Some(package) = current else {
            return Commit::unchanged(self.source);
        };
        Commit::new(
            self.source,
            Published {
                package,
                touched_components: touched,
                deleted_previews: 0,
            },
        )
    }
}

/// The host editor's post-change validation of `change`.
fn verify(package: &Package, storage: u64, change: &Change) -> Result<(), Error> {
    let working = Working::new(package)?;
    let context = Context::load(&working, storage)?;
    let holds = match change {
        Change::Apply(style) => {
            let applied = context.applied()?;
            applied.style().id() == *style && !applied.has_overrides()
        },
        Change::Redefine => !context.applied()?.has_overrides(),
        Change::Create(_, name) => {
            context
                .catalog
                .styles()
                .iter()
                .filter(|style| style.name().as_str() == name.as_str())
                .count()
                == 1
        },
        Change::Rename(style, name) => context
            .catalog
            .get(*style)
            .is_some_and(|found| found.name().as_str() == name.as_str()),
        Change::Delete(style) => context.catalog.get(*style).is_none(),
        Change::DeleteApplied(style, replacement) => {
            context.catalog.get(*style).is_none() && context.applied()?.style().id() == *replacement
        },
    };
    if holds {
        Ok(())
    } else {
        Err(Error::Verification)
    }
}

fn storage_of(package: &Package, sheet: usize, selector: DrawableSelector) -> Result<u64, Error> {
    match graph::text_storage(package, sheet, selector)? {
        TextTarget::Storage { storage } => Ok(storage),
        TextTarget::NoText(kind) => Err(Refusal::NoText { kind: kind.name() }.into()),
        TextTarget::NotFound => Err(Refusal::DrawableNotFound { selector }.into()),
    }
}

/// A listing theme: (object, message index, message type, payload, presets).
type Theme = (u64, usize, u32, Vec<u8>, Vec<u64>);

/// The storage's named-style catalog and the graph it is drawn from.
struct Context {
    first: u64,
    styles: BTreeMap<u64, Vec<u8>>,
    catalog: Catalog,
    themes: Vec<Theme>,
}

impl Context {
    fn load(working: &Working<'_>, storage: u64) -> Result<Self, Error> {
        let limits = WireLimits::DEFAULT;
        let payload = storage_payload(working, storage)?.1;
        let first = engine::uniform_style(&payload, limits).map_err(|error| match error {
            engine::Error::InvalidSource { .. } => Refusal::MixedParagraphStyles.into(),
            other => from_engine(other),
        })?;
        let mut styles = BTreeMap::new();
        let mut candidates = Vec::new();
        for object in working.source().state.components.iter_objects() {
            let identifier = object.archive_info.identifier.ok_or(Error::InvalidSource)?;
            let object = working.object(identifier)?;
            for (index, message) in object.messages.iter().enumerate() {
                if message.type_ == engine::MESSAGE_TYPE {
                    styles.insert(identifier, message.data.clone());
                } else if named::THEME_MESSAGE_TYPES.contains(&message.type_) {
                    candidates.push((identifier, index, message.type_, message.data.clone()));
                }
            }
        }
        let stylesheet = named::stylesheet(&styles, first, limits).map_err(from_engine)?;
        let sheet = sheet_payload(working, stylesheet)?.1;
        let mut themes = Vec::new();
        let mut presets = Vec::new();
        for (identifier, index, message_type, payload) in candidates {
            if let Some(listed) =
                named::theme_presets(&payload, stylesheet, limits).map_err(from_engine)?
            {
                presets.extend_from_slice(&listed);
                themes.push((identifier, index, message_type, payload, listed));
            }
        }
        let catalog =
            named::catalog(&styles, stylesheet, &sheet, &presets, limits).map_err(from_engine)?;
        Ok(Self {
            first,
            styles,
            catalog,
            themes,
        })
    }

    fn applied(&self) -> Result<named::Applied, Error> {
        named::applied(&self.styles, &self.catalog, self.first, WireLimits::DEFAULT)
            .map_err(from_engine)
    }

    /// The themes that list `style`.
    fn themes_with(&self, style: u64) -> Result<Vec<&Theme>, Error> {
        let found = self
            .themes
            .iter()
            .filter(|theme| theme.4.contains(&style))
            .collect::<Vec<_>>();
        if found.is_empty() {
            return Err(Refusal::UnsupportedDependency.into());
        }
        Ok(found)
    }
}

/// The storage message: (index, payload).
fn storage_payload(working: &Working<'_>, storage: u64) -> Result<(usize, Vec<u8>), Error> {
    let object = working.object(storage)?;
    let index = structure::unique_message(object, &[STORAGE_MESSAGE_TYPE])?;
    Ok((index, object.messages[index].data.clone()))
}

/// The stylesheet message: (index, payload).
fn sheet_payload(working: &Working<'_>, stylesheet: u64) -> Result<(usize, Vec<u8>), Error> {
    let object = working.object(stylesheet)?;
    let index = structure::unique_message(object, &[STYLESHEET_MESSAGE_TYPE])?;
    Ok((index, object.messages[index].data.clone()))
}

/// The paragraph style message index of `style`.
fn style_index(working: &Working<'_>, style: u64) -> Result<usize, Error> {
    Ok(structure::unique_message(
        working.object(style)?,
        &[engine::MESSAGE_TYPE],
    )?)
}

/// Every storage's paragraph-table style references, by storage.
fn storage_tables(working: &Working<'_>) -> Result<Vec<(u64, Vec<u64>)>, Error> {
    let mut tables = Vec::new();
    for object in working.source().state.components.iter_objects() {
        let identifier = object.archive_info.identifier.ok_or(Error::InvalidSource)?;
        for message in &working.object(identifier)?.messages {
            if message.type_ != STORAGE_MESSAGE_TYPE {
                continue;
            }
            let mut styles = Vec::new();
            for table in references::records(&message.data, &[STORAGE_PARAGRAPH_TABLE])? {
                for entry in references::records(table, &[1])? {
                    for object in references::records(entry, &[2])? {
                        styles.push(references::reference_identifier(object)?);
                    }
                }
            }
            tables.push((identifier, styles));
        }
    }
    Ok(tables)
}

/// Replace message `index` of `object` canonically and adjust its header
/// references with `adjust`.
fn replace_adjusting(
    working: &mut Working<'_>,
    object: u64,
    index: usize,
    data: Vec<u8>,
    adjust: impl FnOnce(&mut litchi_iwa_core::MessageInfo) -> FaultResult<()>,
) -> FaultResult<()> {
    let limits = working.archive_limits();
    let object = working.object_mut(object)?;
    let message_type = object.messages.get(index).ok_or(Fault::Invalid)?.type_;
    let mut staged = object.clone();
    staged
        .replace_message(
            index,
            RawMessage {
                type_: message_type,
                data,
            },
            limits,
        )
        .map_err(|error| structure::map_core(&error))?;
    adjust(
        staged
            .archive_info
            .message_infos
            .get_mut(index)
            .ok_or(Fault::Invalid)?,
    )?;
    staged
        .validate(limits)
        .map_err(|error| structure::map_core(&error))?;
    *object = staged;
    Ok(())
}

fn drop_reference(info: &mut litchi_iwa_core::MessageInfo, target: u64) {
    info.object_references
        .retain(|reference| *reference != target);
    for field in &mut info.field_infos {
        field
            .object_references
            .retain(|reference| *reference != target);
    }
}

fn component(working: &Working<'_>, registry: &Registry, object: u64) -> Result<u64, Error> {
    let member = working.locate(object)?.ok_or(Error::InvalidSource)?;
    Ok(registry.component_for_entry(&member)?.identifier)
}

/// Drop the storage component's cross-component reference to `style` when
/// no object of the storage's member still references it (the host's
/// `unregister_owner_reference_if_unused`).
fn release_owner_reference(
    working: &Working<'_>,
    registry: &Registry,
    transition: &mut Transition,
    storage: u64,
    style: u64,
) -> Result<(), Error> {
    let storage_member = working.locate(storage)?.ok_or(Error::InvalidSource)?;
    let owner = registry.component_for_entry(&storage_member)?.identifier;
    let target = component(working, registry, style)?;
    if owner == target {
        return Ok(());
    }
    let archive = working.archive(&storage_member)?;
    if archive.objects.iter().any(|object| {
        object
            .archive_info
            .message_infos
            .iter()
            .any(|info| info.object_references.contains(&style))
    }) {
        return Ok(());
    }
    transition.remove_reference(registry, owner, target, style)?;
    Ok(())
}

/// Apply `target` to `storage` on `working`.
fn apply(
    working: &mut Working<'_>,
    registry: &Registry,
    transition: &mut Transition,
    storage: u64,
    target: ParagraphStyleId,
) -> Result<bool, Error> {
    let limits = WireLimits::DEFAULT;
    let context = Context::load(working, storage)?;
    let tables = storage_tables(working)?;
    let mut parents = Vec::new();
    for (style, payload) in &context.styles {
        if let Some(parent) = litchi_iwa_text_wire::character_style_parent(payload, limits)
            .map_err(|_error| Error::InvalidSource)?
        {
            parents.push((*style, parent));
        }
    }
    let is_exclusive = tables
        .iter()
        .flat_map(|(_, styles)| styles)
        .filter(|style| **style == context.first)
        .count()
        == 1
        && !parents.iter().any(|(_, parent)| *parent == context.first);
    let plan = named::plan_apply(
        &context.styles,
        &context.catalog,
        context.first,
        target,
        is_exclusive,
        |style, removed| {
            !tables
                .iter()
                .any(|(owner, styles)| *owner != storage && styles.contains(&style))
                && !parents
                    .iter()
                    .any(|(child, parent)| *parent == style && !removed.contains(child))
        },
        limits,
    )
    .map_err(from_engine)?;
    let (index, payload) = storage_payload(working, storage)?;
    let archive_limits = working.archive_limits();
    let owner = component(working, registry, storage)?;
    match plan {
        ApplyPlan::Unchanged => return Ok(false),
        ApplyPlan::Reference { from, to } => {
            let repointed = engine::repoint(&payload, from, to, limits).map_err(from_engine)?;
            replace_repointing(
                working.object_mut(storage)?,
                index,
                repointed,
                from,
                to,
                archive_limits,
            )?;
            let to_component = component(working, registry, to)?;
            transition.reference(registry, owner, to_component, to)?;
            release_owner_reference(working, registry, transition, storage, from)?;
        },
        ApplyPlan::Replace { from, to, removals } => {
            let repointed = engine::repoint(&payload, from, to, limits).map_err(from_engine)?;
            replace_repointing(
                working.object_mut(storage)?,
                index,
                repointed,
                from,
                to,
                archive_limits,
            )?;
            for removal in removals {
                let style_component = component(working, registry, removal.style)?;
                let (sheet_index, sheet) = sheet_payload(working, removal.stylesheet)?;
                let data = engine::stylesheet_without_variation(
                    &sheet,
                    removal.parent,
                    removal.style,
                    limits,
                )
                .map_err(from_engine)?;
                replace_adjusting(working, removal.stylesheet, sheet_index, data, |info| {
                    drop_reference(info, removal.style);
                    Ok(())
                })?;
                working.remove(removal.style)?;
                transition.remove_objects(registry, &[removal.style])?;
                transition.reference(registry, owner, style_component, to)?;
                transition.release_suffix(&[removal.style])?;
            }
        },
        _ => return Err(Error::InvalidSource),
    }
    Ok(true)
}

fn create(
    working: &mut Working<'_>,
    registry: &Registry,
    transition: &mut Transition,
    storage: u64,
    source: ParagraphStyleId,
    name: &ParagraphStyleName,
) -> Result<(), Error> {
    let limits = WireLimits::DEFAULT;
    let context = Context::load(working, storage)?;
    context
        .catalog
        .plan_create(source, name)
        .map_err(from_engine)?;
    let source_id = native_id(source);
    let stylesheet = context.catalog.stylesheet();
    let sheet_member = working.locate(stylesheet)?.ok_or(Error::InvalidSource)?;
    if working.locate(source_id)?.as_deref() != Some(sheet_member.as_str()) {
        return Err(Refusal::UnsupportedDependency.into());
    }
    let themes = context.themes_with(source_id)?;
    let created = *working.allocate(1)?.first().ok_or(Error::InvalidSource)?;
    let identifier = named::generated_identifier(created);
    let payload = named::named_copy(
        context.styles.get(&source_id).ok_or(Error::InvalidSource)?,
        name,
        &identifier,
        limits,
    )
    .map_err(from_engine)?;
    let archive_limits = working.archive_limits();
    let original = working.object(source_id)?.clone();
    let source_index = style_index(working, source_id)?;
    let mut object = ArchiveObject::new(created, original.messages.clone(), archive_limits)
        .map_err(|error| structure::map_core(&error))?;
    object.archive_info.message_infos = original.archive_info.message_infos.clone();
    object.archive_info.should_merge = original.archive_info.should_merge;
    object
        .replace_message(
            source_index,
            RawMessage {
                type_: engine::MESSAGE_TYPE,
                data: payload,
            },
            archive_limits,
        )
        .map_err(|error| structure::map_core(&error))?;
    let (sheet_index, sheet) = sheet_payload(working, stylesheet)?;
    let sheet_data =
        named::stylesheet_with_named(&sheet, created, &identifier, limits).map_err(from_engine)?;
    replace_adjusting(working, stylesheet, sheet_index, sheet_data, |info| {
        structure::reserve(&mut info.object_references, 1)?;
        info.object_references.push(created);
        Ok(())
    })?;
    working.insert(&sheet_member, object)?;
    let style_component = registry.component_for_entry(&sheet_member)?.identifier;
    for (theme, index, _, payload, _) in &themes {
        let rewritten =
            named::theme_with_preset(payload, stylesheet, created, limits).map_err(from_engine)?;
        replace_adjusting(working, *theme, *index, rewritten, |info| {
            structure::reserve(&mut info.object_references, 1)?;
            info.object_references.push(created);
            Ok(())
        })?;
    }
    let seed = metadata::identity_seed(working, created)?;
    transition.register_uuid(registry, style_component, created, seed)?;
    for (theme, ..) in &themes {
        let theme_component = component(working, registry, *theme)?;
        transition.reference(registry, theme_component, style_component, created)?;
    }
    Ok(())
}

fn rename_style(
    working: &mut Working<'_>,
    storage: u64,
    target: ParagraphStyleId,
    name: &ParagraphStyleName,
) -> Result<bool, Error> {
    let context = Context::load(working, storage)?;
    if context
        .catalog
        .plan_rename(target, name)
        .map_err(from_engine)?
        .is_none()
    {
        return Ok(false);
    }
    let id = native_id(target);
    let index = style_index(working, id)?;
    let payload = named::renamed(
        &working.object(id)?.messages[index].data,
        name,
        WireLimits::DEFAULT,
    )
    .map_err(from_engine)?;
    replace_adjusting(working, id, index, payload, |_info| Ok(()))?;
    Ok(true)
}

fn delete(
    working: &mut Working<'_>,
    registry: &Registry,
    transition: &mut Transition,
    storage: u64,
    target: ParagraphStyleId,
) -> Result<(), Error> {
    let limits = WireLimits::DEFAULT;
    let context = Context::load(working, storage)?;
    context.catalog.plan_delete(target).map_err(from_engine)?;
    let id = native_id(target);
    let stylesheet = context.catalog.stylesheet();
    let sheet_member = working.locate(stylesheet)?.ok_or(Error::InvalidSource)?;
    if working.locate(id)?.as_deref() != Some(sheet_member.as_str()) {
        return Err(Refusal::UnsupportedDependency.into());
    }
    let themes = context.themes_with(id)?;
    // Nothing but the stylesheet and the listing themes may still name it.
    for object in working.source().state.components.iter_objects() {
        let object_id = object.archive_info.identifier.ok_or(Error::InvalidSource)?;
        let object = working.object(object_id)?;
        for (message, info) in object
            .messages
            .iter()
            .zip(&object.archive_info.message_infos)
        {
            if !info.object_references.contains(&id) {
                continue;
            }
            let is_sheet = object_id == stylesheet && message.type_ == STYLESHEET_MESSAGE_TYPE;
            let is_theme = themes
                .iter()
                .any(|theme| theme.0 == object_id && theme.2 == message.type_);
            if !is_sheet && !is_theme {
                return Err(Refusal::StyleInUse.into());
            }
        }
    }
    for (theme, index, _, payload, _) in &themes {
        let rewritten =
            named::theme_without_preset(payload, stylesheet, id, limits).map_err(from_engine)?;
        replace_adjusting(working, *theme, *index, rewritten, |info| {
            drop_reference(info, id);
            Ok(())
        })?;
    }
    let (sheet_index, sheet) = sheet_payload(working, stylesheet)?;
    let sheet_data = named::stylesheet_without_named(&sheet, id, limits).map_err(from_engine)?;
    replace_adjusting(working, stylesheet, sheet_index, sheet_data, |info| {
        drop_reference(info, id);
        Ok(())
    })?;
    working.remove(id)?;
    transition.remove_objects(registry, &[id])?;
    transition.release_suffix(&[id])?;
    Ok(())
}

fn redefine(
    working: &mut Working<'_>,
    registry: &Registry,
    transition: &mut Transition,
    storage: u64,
) -> Result<(), Error> {
    let context = Context::load(working, storage)?;
    let applied = context.applied()?;
    let redefinition = named::redefine(
        &context.styles,
        &context.catalog,
        context.first,
        WireLimits::DEFAULT,
    )
    .map_err(from_engine)?;
    let style = redefinition.style();
    let index = style_index(working, style)?;
    let managed = redefinition.managed_fields();
    let (previous, following) = (redefinition.previous_following(), redefinition.following());
    let archive_limits = working.archive_limits();
    replace_adjusting(
        working,
        style,
        index,
        redefinition.payload().to_vec(),
        |info| {
            if let Some(previous) = previous {
                info.object_references
                    .retain(|reference| *reference != previous);
            }
            if let Some(following) = following
                && !info.object_references.contains(&following)
            {
                structure::reserve(&mut info.object_references, 1)?;
                info.object_references.push(following);
            }
            sync_managed(&mut info.field_infos, managed, archive_limits)
        },
    )?;
    apply(working, registry, transition, storage, applied.style().id())?;
    Ok(())
}

fn apply_one(source: &Package, storage: u64, change: &Change) -> Result<Option<Published>, Error> {
    let mut working = Working::new(source)?;
    working.keep_previews();
    let registry =
        metadata::registry(&working)?.ok_or(Error::Refused(Refusal::UnsupportedSource))?;
    let mut transition = Transition::default();
    let changed = match change {
        Change::Apply(target) => apply(&mut working, &registry, &mut transition, storage, *target)?,
        Change::Create(source, name) => {
            create(
                &mut working,
                &registry,
                &mut transition,
                storage,
                *source,
                name,
            )?;
            true
        },
        Change::Rename(target, name) => rename_style(&mut working, storage, *target, name)?,
        Change::Delete(target) => {
            delete(&mut working, &registry, &mut transition, storage, *target)?;
            true
        },
        Change::DeleteApplied(target, replacement) => {
            let context = Context::load(&working, storage)?;
            context
                .catalog
                .plan_replacement(&context.applied()?, *target, *replacement)
                .map_err(from_engine)?;
            apply(
                &mut working,
                &registry,
                &mut transition,
                storage,
                *replacement,
            )?;
            delete(&mut working, &registry, &mut transition, storage, *target)?;
            true
        },
        Change::Redefine => {
            redefine(&mut working, &registry, &mut transition, storage)?;
            true
        },
    };
    if !changed {
        return Ok(None);
    }
    Ok(Some(working.publish(&transition)?))
}

fn resolve_sheet(package: &Package, selector: SheetSelector<'_>) -> Result<usize, Error> {
    graph::sheet_position(package, selector)?.ok_or_else(|| Refusal::SheetNotFound.into())
}

impl Package {
    /// The named paragraph styles the text of one sheet text box or shape can
    /// use, in theme preset order.
    ///
    /// # Errors
    ///
    /// Selector errors, [`Refusal::NoText`],
    /// [`Refusal::MixedParagraphStyles`], source errors.
    pub fn sheet_text_box_named_styles<'sheet>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        drawable: impl Into<DrawableSelector>,
    ) -> Result<Vec<NamedParagraphStyle>, Error> {
        let sheet = resolve_sheet(self, sheet.into())?;
        let storage = storage_of(self, sheet, drawable.into())?;
        let working = Working::new(self)?;
        Ok(Context::load(&working, storage)?.catalog.styles().to_vec())
    }

    /// The named paragraph style the text of one sheet text box or shape
    /// uses, and whether the text overrides it.
    ///
    /// # Errors
    ///
    /// As [`Self::sheet_text_box_named_styles`].
    pub fn sheet_text_box_applied_named_style<'sheet>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        drawable: impl Into<DrawableSelector>,
    ) -> Result<(NamedParagraphStyle, bool), Error> {
        let sheet = resolve_sheet(self, sheet.into())?;
        let storage = storage_of(self, sheet, drawable.into())?;
        let working = Working::new(self)?;
        let applied = Context::load(&working, storage)?.applied()?;
        Ok((applied.style().clone(), applied.has_overrides()))
    }

    /// Start an immutable named-style edit of the text of one sheet text box
    /// or shape.
    ///
    /// # Errors
    ///
    /// Selector errors, [`Refusal::NoText`], source errors.
    pub fn edit_sheet_text_box_named_styles<'sheet>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        drawable: impl Into<DrawableSelector>,
    ) -> Result<Edit<'_>, Error> {
        let sheet = resolve_sheet(self, sheet.into())?;
        let selector = drawable.into();
        storage_of(self, sheet, selector)?;
        Ok(Edit {
            source: self,
            sheet,
            selector,
            changes: Vec::new(),
        })
    }

    /// Apply a named-style patch to its exact source package.
    ///
    /// # Errors
    ///
    /// [`Error::PatchConflict`] when this package is not the patch's exact
    /// source.
    pub fn apply_sheet_text_box_named_styles(&self, patch: &Patch) -> Result<Commit, Error> {
        apply_patch(self, patch)
    }
}
