//! Implementation of the public [`crate::sheet::drawable::drop_cap`]
//! transaction: plain-text Drop Caps of the paragraphs of the text a sheet
//! text box or shape owns, on W1-D's engine
//! ([`litchi_iwa_text_wire::drop_cap`], record 0374), record 0378.
//!
//! The engine plans each change; this module publishes the plan as the host
//! text editor does:
//!
//! - **Rewrite**: the paragraph's own variation takes the new payload;
//! - **Derive**: a new variation of the stylesheet's Drop Cap base style is
//!   stored beside the base and registered in the stylesheet, the paragraph's
//!   Drop Cap entry names it, and its UUID and cross-component reference are
//!   registered; a previous style the storage's component no longer uses
//!   loses its cross-component reference;
//! - **Detach**: the entry names no style; an exclusive variation is removed
//!   with its stylesheet registration, UUID and references.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "the focused transaction keeps its public types beside the staged rewrite"
)]

use std::collections::BTreeMap;
use std::fmt;

use litchi_iwa_common::WireLimits;
use litchi_iwa_common::wire::WireView;
use litchi_iwa_core::{ArchiveObject, RawMessage};
use litchi_iwa_text::paragraph::drop_cap::DropCap;
use litchi_iwa_text::position::TextPosition;
use litchi_iwa_text_wire::drop_cap::{self as engine, Change, Plan};
use litchi_iwa_text_wire::paragraph_style;
use thiserror::Error as ThisError;

use super::Package;
use super::sheet_drawable_graph::{self as graph, TextTarget};
use super::structure::{
    self, Fault, FaultResult, Published, Working,
    metadata::{self, Registry, Transition},
    references,
};
use crate::selector::SheetSelector;
use crate::sheet::drawable::DrawableSelector;

structure::structural_family!("drop-cap", Wire(litchi_iwa_common::LimitKind));

/// `TSWP.StorageArchive`.
const STORAGE_MESSAGE_TYPE: u32 = 2_001;
/// `TSS.StylesheetArchive`.
const STYLESHEET_MESSAGE_TYPE: u32 = 401;
/// `TN.ThemeArchive`.
const THEME_MESSAGE_TYPE: u32 = 12_009;
/// `TSWP.ThemePresetsArchive.dropcap_style_presets`.
const DROP_CAP_PRESETS_FIELD: u32 = 8;

/// A typed, content-redacted Drop Cap failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies
    /// it (record 0351 §2.4).
    #[error(transparent)]
    Refused(Refusal),
    /// The rooted source is malformed or ambiguous.
    #[error("rooted Drop Cap source is invalid")]
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
    #[error("could not allocate {amount} Drop Cap items or bytes")]
    Allocation {
        /// Requested amount.
        amount: usize,
    },
    /// The reopened candidate did not reproduce the requested state.
    #[error("Drop Cap candidate verification failed")]
    Verification,
    /// The patch was applied to an artifact other than its exact source.
    #[error("Drop Cap patch conflicts with this package")]
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
    /// The position is not the start of a paragraph.
    #[error("position is not a paragraph start")]
    NotAParagraphStart,
    /// The stylesheet has no Drop Cap base style (the theme lists none).
    #[error("the document theme has no Drop Cap base style")]
    MissingBaseStyle,
    /// Commit was requested before staging a change.
    #[error("no Drop Cap change is staged")]
    NoStagedOperation,
    /// A surviving object depends on the rewritten graph.
    #[error("a surviving object depends on the Drop Cap style graph")]
    UnsupportedDependency,
    /// The source cannot publish an exact changed artifact.
    #[error("source representation does not support Drop Cap edits")]
    UnsupportedSource,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::SheetNotFound | Self::DrawableNotFound { .. } => RefusalKind::NotFound,
            Self::NoText { .. } | Self::MissingBaseStyle => RefusalKind::UnsupportedTarget,
            Self::NotAParagraphStart => RefusalKind::InvalidValue,
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
        engine::Error::Refused(_) => Refusal::NotAParagraphStart.into(),
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

/// Drop Cap changes staged against the text of one sheet drawable.
pub struct Edit<'a> {
    source: &'a Package,
    sheet: usize,
    selector: DrawableSelector,
    changes: Vec<(TextPosition, Change)>,
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
    /// Stage `drop_cap` on the paragraph starting at `paragraph`.
    ///
    /// # Errors
    ///
    /// [`Error::Allocation`].
    pub fn set(&mut self, paragraph: TextPosition, drop_cap: DropCap) -> Result<&mut Self, Error> {
        structure::reserve(&mut self.changes, 1)?;
        self.changes.push((paragraph, Change::Set(drop_cap)));
        Ok(self)
    }

    /// Stage removing the Drop Cap of the paragraph starting at
    /// `paragraph`.
    ///
    /// # Errors
    ///
    /// [`Error::Allocation`].
    pub fn remove(&mut self, paragraph: TextPosition) -> Result<&mut Self, Error> {
        structure::reserve(&mut self.changes, 1)?;
        self.changes.push((paragraph, Change::Remove));
        Ok(self)
    }

    /// Validate and atomically publish the staged changes, in order; an
    /// edit that changes nothing is an exact no-op.
    ///
    /// # Errors
    ///
    /// [`Refusal::NoStagedOperation`], [`Refusal::NotAParagraphStart`],
    /// [`Refusal::MissingBaseStyle`], source and verification errors.
    pub fn commit(self) -> Result<Commit, Error> {
        if self.changes.is_empty() {
            return Err(Refusal::NoStagedOperation.into());
        }
        let mut current: Option<Package> = None;
        let mut touched = 0_usize;
        for (paragraph, change) in &self.changes {
            let base = current.as_ref().unwrap_or(self.source);
            if let Some(published) =
                apply_one(base, self.sheet, self.selector, *paragraph, *change)?
            {
                let after = drop_caps(&published.package, self.sheet, self.selector)?;
                let holds = match change {
                    Change::Set(value) => after
                        .iter()
                        .any(|(at, cap)| at == paragraph && cap == value),
                    _ => !after.iter().any(|(at, _)| at == paragraph),
                };
                if !holds {
                    return Err(Error::Verification);
                }
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

fn storage_of(package: &Package, sheet: usize, selector: DrawableSelector) -> Result<u64, Error> {
    match graph::text_storage(package, sheet, selector)? {
        TextTarget::Storage { storage } => Ok(storage),
        TextTarget::NoText(kind) => Err(Refusal::NoText { kind: kind.name() }.into()),
        TextTarget::NotFound => Err(Refusal::DrawableNotFound { selector }.into()),
    }
}

/// The package's Drop Cap styles.
struct Styles {
    payloads: BTreeMap<u64, Vec<u8>>,
    exclusive: BTreeMap<u64, bool>,
}

/// The object references of a storage's Drop Cap table entries.
fn table_styles(storage: &[u8]) -> FaultResult<Vec<u64>> {
    let mut styles = Vec::new();
    for table in references::records(storage, &[engine::TABLE_FIELD])? {
        for entry in references::records(table, &[1])? {
            for object in references::records(entry, &[2])? {
                styles.push(references::reference_identifier(object)?);
            }
        }
    }
    Ok(styles)
}

fn styles(working: &Working<'_>) -> Result<Styles, Error> {
    let mut payloads = BTreeMap::new();
    let mut parents = Vec::new();
    let mut counts = BTreeMap::<u64, usize>::new();
    for object in working.source().state.components.iter_objects() {
        let identifier = object.archive_info.identifier.ok_or(Error::InvalidSource)?;
        for message in &object.messages {
            if message.type_ == engine::MESSAGE_TYPE {
                if let Some(parent) =
                    litchi_iwa_text_wire::character_style_parent(&message.data, WireLimits::DEFAULT)
                        .map_err(|_error| Error::InvalidSource)?
                {
                    parents.push(parent);
                }
                payloads.insert(identifier, message.data.clone());
            } else if message.type_ == STORAGE_MESSAGE_TYPE {
                for style in table_styles(&message.data)? {
                    *counts.entry(style).or_default() += 1;
                }
            }
        }
    }
    let exclusive = payloads
        .keys()
        .map(|style| {
            (
                *style,
                counts.get(style) == Some(&1) && !parents.contains(style),
            )
        })
        .collect();
    Ok(Styles {
        payloads,
        exclusive,
    })
}

/// A singular field of a `TSS.StyleArchive` (`super` of a style message).
fn style_field(style: &[u8], number: u32) -> Result<Option<Vec<u8>>, Error> {
    let view = WireView::parse(style, WireLimits::DEFAULT)?;
    let Some(base) = graph::unique_field(&view, 1, 2)? else {
        return Ok(None);
    };
    let base = WireView::parse(base.payload(), WireLimits::DEFAULT)?;
    Ok(base
        .fields()
        .filter(|field| field.number() == number)
        .last()
        .map(|field| field.payload().to_vec()))
}

/// The stylesheet's Drop Cap base: the theme preset that is a root style of
/// the stylesheet (the host's `base_style`).
fn base(package: &Package, styles: &Styles, stylesheet: u64) -> Result<u64, Error> {
    let document = graph::object(package, 1)?;
    let document = &document.messages[structure::document_message(document)?].data;
    let theme = references::records(document, &[6])?;
    let [theme] = theme.as_slice() else {
        return Err(Error::InvalidSource);
    };
    let theme = graph::object(package, references::reference_identifier(theme)?)?;
    let theme = &theme.messages[structure::unique_message(theme, &[THEME_MESSAGE_TYPE])?].data;
    let mut presets = Vec::new();
    for super_ in references::records(theme, &[1])? {
        for text in references::records(super_, &[110])? {
            for preset in references::records(text, &[DROP_CAP_PRESETS_FIELD])? {
                presets.push(references::reference_identifier(preset)?);
            }
        }
    }
    let mut found = Vec::new();
    for preset in presets {
        let Some(payload) = styles.payloads.get(&preset) else {
            continue;
        };
        let parent = style_field(payload, 3)?;
        let variation = style_field(payload, 4)?;
        let sheet = style_field(payload, 5)?
            .map(|reference| references::reference_identifier(&reference))
            .transpose()?;
        let variation = match variation {
            Some(raw) => {
                litchi_iwa_common::decode_varint_from_bytes(&raw)
                    .map_err(|_error| Error::InvalidSource)?
                    .0
                    != 0
            },
            None => false,
        };
        if parent.is_none() && !variation && sheet == Some(stylesheet) && !found.contains(&preset) {
            found.push(preset);
        }
    }
    match found.as_slice() {
        [base] => Ok(*base),
        [] => Err(Refusal::MissingBaseStyle.into()),
        _ => Err(Error::InvalidSource),
    }
}

/// The Drop Caps of the drawable's paragraphs.
fn drop_caps(
    package: &Package,
    sheet: usize,
    selector: DrawableSelector,
) -> Result<Vec<(TextPosition, DropCap)>, Error> {
    let storage = storage_of(package, sheet, selector)?;
    let object = graph::object(package, storage)?;
    let payload =
        &object.messages[structure::unique_message(object, &[STORAGE_MESSAGE_TYPE])?].data;
    let placements = engine::placements(payload, WireLimits::DEFAULT).map_err(from_engine)?;
    if placements
        .entries()
        .iter()
        .all(|(_, style)| style.is_none())
    {
        return Ok(Vec::new());
    }
    let working = Working::new(package)?;
    let styles = styles(&working)?;
    let base = base(package, &styles, placements.stylesheet())?;
    let mut found = Vec::new();
    for (paragraph, style) in placements.entries() {
        let Some(style) = style.filter(|style| *style != base) else {
            continue;
        };
        let payload = styles.payloads.get(&style).ok_or(Error::InvalidSource)?;
        found.push((
            *paragraph,
            engine::model(payload, WireLimits::DEFAULT)
                .map_err(from_engine)?
                .drop_cap(),
        ));
    }
    Ok(found)
}

/// Replace message `index` of `object` canonically and adjust its header
/// references with `adjust`.
fn replace_adjusting(
    object: &mut ArchiveObject,
    index: usize,
    data: Vec<u8>,
    limits: litchi_iwa_core::Limits,
    adjust: impl FnOnce(&mut litchi_iwa_core::MessageInfo) -> FaultResult<()>,
) -> FaultResult<()> {
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

/// The host's Drop Cap reference bookkeeping on the storage message.
fn table_references(
    info: &mut litchi_iwa_core::MessageInfo,
    added: Option<u64>,
    removed: Option<u64>,
) -> FaultResult<()> {
    if let Some(old) = removed {
        if info
            .object_references
            .iter()
            .filter(|reference| **reference == old)
            .count()
            != 1
        {
            return Err(Fault::Invalid);
        }
        info.object_references.retain(|reference| *reference != old);
        for field in &mut info.field_infos {
            if field.path.as_slice().first() == Some(&engine::TABLE_FIELD) {
                field
                    .object_references
                    .retain(|reference| *reference != old);
            }
        }
    }
    if let Some(new) = added {
        if info.object_references.contains(&new) {
            return Err(Fault::Invalid);
        }
        info.object_references.push(new);
        for field in &mut info.field_infos {
            if field.path.as_slice().first() == Some(&engine::TABLE_FIELD)
                && !field.object_references.contains(&new)
            {
                field.object_references.push(new);
            }
        }
    }
    Ok(())
}

/// Drop the storage component's cross-component reference to `style` when
/// no object of that component still references it (the host's
/// `unregister_owner_reference_if_unused`).
fn release_owner_reference(
    working: &Working<'_>,
    registry: &Registry,
    transition: &mut Transition,
    storage_member: &str,
    style: u64,
) -> Result<(), Error> {
    let style_member = working.locate(style)?.ok_or(Error::InvalidSource)?;
    let storage_component = registry.component_for_entry(storage_member)?.identifier;
    let style_component = registry.component_for_entry(&style_member)?.identifier;
    if storage_component == style_component {
        return Ok(());
    }
    let archive = working.archive(storage_member)?;
    if archive.objects.iter().any(|object| {
        object
            .archive_info
            .message_infos
            .iter()
            .any(|info| info.object_references.contains(&style))
    }) {
        return Ok(());
    }
    for external in &registry.externals {
        if external.source == storage_component
            && external.target == style_component
            && external.object == style
        {
            structure::reserve(&mut transition.external_removals, 1)?;
            transition.external_removals.push(*external);
        }
    }
    Ok(())
}

#[allow(
    clippy::too_many_lines,
    reason = "one staged change: plan, then the rewrite, derive or detach publication"
)]
fn apply_one(
    source: &Package,
    sheet: usize,
    selector: DrawableSelector,
    paragraph: TextPosition,
    change: Change,
) -> Result<Option<Published>, Error> {
    let storage = storage_of(source, sheet, selector)?;
    let mut working = Working::new(source)?;
    working.keep_previews();
    let limits = WireLimits::DEFAULT;
    let storage_member = working.locate(storage)?.ok_or(Error::InvalidSource)?;
    let storage_index =
        structure::unique_message(working.object(storage)?, &[STORAGE_MESSAGE_TYPE])?;
    let payload = working.object(storage)?.messages[storage_index]
        .data
        .clone();
    let placements = engine::placements(&payload, limits).map_err(from_engine)?;
    let stylesheet = placements.stylesheet();
    let styles = styles(&working)?;
    let base = base(source, &styles, stylesheet)?;
    let plan = engine::plan(
        &payload,
        paragraph,
        base,
        |style| styles.payloads.get(&style).map(Vec::as_slice),
        |style| styles.exclusive.get(&style).copied().unwrap_or(false),
        change,
        limits,
    )
    .map_err(from_engine)?;
    let registry =
        metadata::registry(&working)?.ok_or(Error::Refused(Refusal::UnsupportedSource))?;
    let mut transition = Transition::default();
    let archive_limits = working.archive_limits();
    match plan {
        Plan::Unchanged => return Ok(None),
        Plan::Rewrite { style, payload } => {
            let index = structure::unique_message(working.object(style)?, &[engine::MESSAGE_TYPE])?;
            replace_adjusting(
                working.object_mut(style)?,
                index,
                payload,
                archive_limits,
                |_info| Ok(()),
            )?;
        },
        Plan::Derive { previous, payload } => {
            let base_member = working.locate(base)?.ok_or(Error::InvalidSource)?;
            let created = *working.allocate(1)?.first().ok_or(Error::InvalidSource)?;
            let mut object = structure::new_object(
                created,
                None,
                engine::MESSAGE_TYPE,
                payload,
                vec![base],
                archive_limits,
            )?;
            object.archive_info.message_infos[0].versions = engine::MESSAGE_VERSION.to_vec().into();
            let sheet_index =
                structure::unique_message(working.object(stylesheet)?, &[STYLESHEET_MESSAGE_TYPE])?;
            let sheet_data = paragraph_style::stylesheet_with_variation(
                &working.object(stylesheet)?.messages[sheet_index].data,
                base,
                created,
                limits,
            )
            .map_err(|_error| Error::InvalidSource)?;
            replace_adjusting(
                working.object_mut(stylesheet)?,
                sheet_index,
                sheet_data,
                archive_limits,
                |info| {
                    structure::reserve(&mut info.object_references, 1)?;
                    info.object_references.push(created);
                    Ok(())
                },
            )?;
            working.insert(&base_member, object)?;
            let rewrite = engine::point(
                &payload_of(&working, storage, storage_index)?,
                paragraph,
                previous,
                Some(created),
                limits,
            )
            .map_err(from_engine)?;
            let (added, removed) = (rewrite.added(), rewrite.removed());
            replace_adjusting(
                working.object_mut(storage)?,
                storage_index,
                rewrite.into_payload(),
                archive_limits,
                |info| table_references(info, added, removed),
            )?;
            let style_component = registry.component_for_entry(&base_member)?.identifier;
            let storage_component = registry.component_for_entry(&storage_member)?.identifier;
            let seed = metadata::identity_seed(&working, created)?;
            transition.register_uuid(&registry, style_component, created, seed)?;
            transition.reference(&registry, storage_component, style_component, created)?;
            if let Some(previous) = previous {
                release_owner_reference(
                    &working,
                    &registry,
                    &mut transition,
                    &storage_member,
                    previous,
                )?;
            }
        },
        Plan::Detach {
            style,
            parent,
            exclusive,
        } => {
            let rewrite = engine::point(&payload, paragraph, Some(style), None, limits)
                .map_err(from_engine)?;
            let (added, removed) = (rewrite.added(), rewrite.removed());
            replace_adjusting(
                working.object_mut(storage)?,
                storage_index,
                rewrite.into_payload(),
                archive_limits,
                |info| table_references(info, added, removed),
            )?;
            if exclusive {
                let sheet_index = structure::unique_message(
                    working.object(stylesheet)?,
                    &[STYLESHEET_MESSAGE_TYPE],
                )?;
                let sheet_data = paragraph_style::stylesheet_without_variation(
                    &working.object(stylesheet)?.messages[sheet_index].data,
                    parent,
                    style,
                    limits,
                )
                .map_err(|_error| Error::InvalidSource)?;
                replace_adjusting(
                    working.object_mut(stylesheet)?,
                    sheet_index,
                    sheet_data,
                    archive_limits,
                    |info| {
                        info.object_references
                            .retain(|reference| *reference != style);
                        for field in &mut info.field_infos {
                            field
                                .object_references
                                .retain(|reference| *reference != style);
                        }
                        Ok(())
                    },
                )?;
                transition.remove_objects(&registry, &[style])?;
                transition.release_suffix(&[style])?;
                working.remove(style)?;
            } else {
                release_owner_reference(
                    &working,
                    &registry,
                    &mut transition,
                    &storage_member,
                    style,
                )?;
            }
        },
        _ => return Err(Error::InvalidSource),
    }
    Ok(Some(working.publish(&transition)?))
}

fn payload_of(working: &Working<'_>, storage: u64, index: usize) -> Result<Vec<u8>, Error> {
    Ok(working.object(storage)?.messages[index].data.clone())
}

fn resolve_sheet(package: &Package, selector: SheetSelector<'_>) -> Result<usize, Error> {
    graph::sheet_position(package, selector)?.ok_or_else(|| Refusal::SheetNotFound.into())
}

impl Package {
    /// The Drop Caps of the paragraphs of the text a sheet text box or shape
    /// owns, by paragraph start.
    ///
    /// # Errors
    ///
    /// Selector errors, [`Refusal::NoText`], source errors.
    pub fn sheet_drawable_drop_caps<'sheet>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        drawable: impl Into<DrawableSelector>,
    ) -> Result<Vec<(TextPosition, DropCap)>, Error> {
        let sheet = resolve_sheet(self, sheet.into())?;
        drop_caps(self, sheet, drawable.into())
    }

    /// Start an immutable Drop Cap edit of the text of one sheet text box or
    /// shape.
    ///
    /// # Errors
    ///
    /// Selector errors, [`Refusal::NoText`], source errors.
    pub fn edit_sheet_drawable_drop_caps<'sheet>(
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

    /// Apply a Drop Cap patch to its exact source package.
    ///
    /// # Errors
    ///
    /// [`Error::PatchConflict`] when this package is not the patch's exact
    /// source.
    pub fn apply_sheet_drawable_drop_caps(&self, patch: &Patch) -> Result<Commit, Error> {
        apply_patch(self, patch)
    }
}
