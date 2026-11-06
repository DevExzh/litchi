//! Package-level master-metadata edits (record 0357).
//!
//! [`Package::edit_master_metadata`] stages the bounded metadata of one
//! master record (its name, main-master design name and notes-master text
//! styles), [`Edit::commit`] publishes a new [`Package`] through one
//! append-only user edit that rewrites only that master's persisted record,
//! and [`Package::apply_master_metadata`] replays a [`Patch`] onto the exact
//! package it was made from. Dropping an [`Edit`] discards it.

use litchi_core::Position;

use super::{
    Change, Commit, Context, Error, Name, Patch, RecordEdit, Result, Snapshot, notes_styles,
    template,
};
use crate::Package;
use crate::master::{self, Kind};
use crate::master_layout;
use crate::records::Record;
use crate::scoped::{self, Base, Failure, Lazy};
use crate::slide_order::{self, records::Owner};

const OPERATION: &str = "edit_master_metadata";

fn layout_limits(base: &Base) -> master_layout::Limits {
    let source = base.open_options().source_limits_value();
    master_layout::Limits {
        max_depth: source.max_depth(),
        max_records: source.max_events(),
        max_bytes: source.max_record_bytes(),
    }
}

/// A staged metadata edit of one master over one immutable [`Package`].
pub struct Edit {
    base: Base,
    kind: Kind,
    position: Position,
    state: Lazy<State>,
}

struct State {
    native: slide_order::Snapshot,
    owner: Owner,
    before: Vec<u8>,
    inner: RecordEdit,
}

impl std::fmt::Debug for Edit {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("kind", &self.kind)
            .field("position", &self.position)
            .field("initialized", &self.state.get().is_some())
            .finish_non_exhaustive()
    }
}

impl Edit {
    fn new(package: &Package, kind: Kind, position: Position) -> Result<Self> {
        let base = Base::capture(package)?;
        Ok(Self {
            base,
            kind,
            position,
            state: Lazy::new(),
        })
    }

    fn initialize_state(&self) -> std::result::Result<State, Failure> {
        let native = self.base.native()?;
        let located =
            master::locate::locate(&self.base, &native, self.kind, self.position, OPERATION)?;
        let limits = layout_limits(&self.base);
        let inner = Snapshot::parse(located.context, &located.record, &limits)
            .map_err(scoped::publication(OPERATION))?
            .edit();
        Ok(State {
            native,
            owner: located.owner,
            before: located.record,
            inner,
        })
    }

    fn state(&self) -> std::result::Result<&State, Failure> {
        self.state.get_or_try_init(|| self.initialize_state())
    }

    fn state_mut(&mut self) -> std::result::Result<&mut State, Failure> {
        let base = self.base.clone();
        let kind = self.kind;
        let position = self.position;
        self.state.get_or_try_init_mut(|| {
            let native = base.native()?;
            let located = master::locate::locate(&base, &native, kind, position, OPERATION)?;
            let limits = layout_limits(&base);
            let inner = Snapshot::parse(located.context, &located.record, &limits)
                .map_err(scoped::publication(OPERATION))?
                .edit();
            Ok::<State, Failure>(State {
                native,
                owner: located.owner,
                before: located.record,
                inner,
            })
        })
    }

    /// The kind of the edited master.
    #[must_use]
    pub const fn kind(&self) -> Kind {
        self.kind
    }

    /// The position of the edited master among the masters of its kind.
    #[must_use]
    pub const fn position(&self) -> Position {
        self.position
    }

    /// The contextual layout kind of the edited master record.
    #[must_use = "handle the fallible master context result"]
    pub fn context(&self) -> Result<Context> {
        Ok(self.state()?.inner.context())
    }

    /// The staged master name.
    ///
    /// # Errors
    ///
    /// Returns [`super::Error::Refused`] when the staged name atom cannot be
    /// read.
    pub fn name(&self) -> Result<Option<Name>> {
        Ok(self
            .state()?
            .inner
            .name()
            .map_err(scoped::publication("master_metadata.name"))?)
    }

    /// The staged main-master design name.
    ///
    /// # Errors
    ///
    /// Returns [`super::Error::Refused`] when the staged design-name atom
    /// cannot be read.
    pub fn template_name(&self) -> Result<Option<template::Name>> {
        Ok(self
            .state()?
            .inner
            .template_name()
            .map_err(scoped::publication("master_metadata.template_name"))?)
    }

    /// The staged notes-master `txStyles` round-trip package.
    ///
    /// # Errors
    ///
    /// Returns [`super::Error::Refused`] when the staged package cannot be
    /// read.
    pub fn notes_styles(&self) -> Result<Option<notes_styles::Styles>> {
        Ok(self
            .state()?
            .inner
            .notes_styles()
            .map_err(scoped::publication("master_metadata.notes_styles"))?)
    }

    /// The staged master record tree.
    #[must_use = "handle the fallible master record result"]
    pub fn record(&self) -> Result<&Record> {
        Ok(self.state()?.inner.record())
    }

    /// The staged structural changes, in order.
    #[must_use = "handle the fallible master changes result"]
    pub fn changes(&self) -> Result<&[Change]> {
        Ok(self.state()?.inner.changes())
    }

    /// Set or replace the master name.
    ///
    /// # Errors
    ///
    /// Returns [`super::Error::Refused`] when the record grammar refuses the
    /// value, leaving the edit unchanged.
    pub fn set_name(&mut self, value: impl Into<String>) -> Result<()> {
        Ok(self
            .state_mut()?
            .inner
            .set_name(value)
            .map_err(scoped::staging("master_metadata.set_name"))?)
    }

    /// Set or replace the main-master design name.
    ///
    /// # Errors
    ///
    /// Returns [`super::Error::Refused`] when the master is not a main
    /// master or the record grammar refuses the value, leaving the edit
    /// unchanged.
    pub fn set_template_name(&mut self, value: impl Into<String>) -> Result<()> {
        Ok(self
            .state_mut()?
            .inner
            .set_template_name(value)
            .map_err(scoped::staging("master_metadata.set_template_name"))?)
    }

    /// Set or replace the notes-master `txStyles` package.
    ///
    /// # Errors
    ///
    /// Returns [`super::Error::Refused`] when the master is not the notes
    /// master or the record grammar refuses the package, leaving the edit
    /// unchanged.
    pub fn set_notes_styles(&mut self, styles: notes_styles::Styles) -> Result<()> {
        Ok(self
            .state_mut()?
            .inner
            .set_notes_styles(styles)
            .map_err(scoped::staging("master_metadata.set_notes_styles"))?)
    }

    /// Build and set a notes-master `txStyles` package from XML.
    ///
    /// # Errors
    ///
    /// Returns a package validation error for malformed XML or embedded-package
    /// content, or [`super::Error::Refused`] for a value the master cannot carry.
    /// A failure leaves the edit unchanged.
    pub fn set_notes_styles_xml(&mut self, xml: impl AsRef<[u8]>) -> Result<()> {
        Ok(self
            .state_mut()?
            .inner
            .set_notes_styles_xml(xml)
            .map_err(scoped::staging("master_metadata.set_notes_styles_xml"))?)
    }

    /// Remove the master name, returning whether one was present.
    ///
    /// # Errors
    ///
    /// Returns [`super::Error::Refused`] when the master's records cannot be
    /// edited, leaving the edit unchanged.
    pub fn clear_name(&mut self) -> Result<bool> {
        Ok(self
            .state_mut()?
            .inner
            .clear_name()
            .map_err(scoped::staging("master_metadata.clear_name"))?)
    }

    /// Remove the main-master design name, returning whether one was
    /// present.
    ///
    /// # Errors
    ///
    /// Returns [`super::Error::Refused`] when the master's records cannot be
    /// edited, leaving the edit unchanged.
    pub fn clear_template_name(&mut self) -> Result<bool> {
        Ok(self
            .state_mut()?
            .inner
            .clear_template_name()
            .map_err(scoped::staging("master_metadata.clear_template_name"))?)
    }

    /// Remove the notes-master `txStyles` package, returning whether one was
    /// present.
    ///
    /// # Errors
    ///
    /// Returns [`super::Error::Refused`] when the master's records cannot be
    /// edited, leaving the edit unchanged.
    pub fn clear_notes_styles(&mut self) -> Result<bool> {
        Ok(self
            .state_mut()?
            .inner
            .clear_notes_styles()
            .map_err(scoped::staging("master_metadata.clear_notes_styles"))?)
    }

    /// Publish the staged metadata as a new package. An edit with no net
    /// change publishes the source package unchanged.
    ///
    /// # Errors
    ///
    /// Returns [`super::Error::Refused`] when the staged master is not valid,
    /// [`super::Error::Signed`] for a changed write to a signed package and
    /// [`super::Error::Package`] when the result cannot be republished.
    pub fn commit(self) -> Result<Commit> {
        self.state
            .get_or_try_init(|| self.initialize_state())
            .map_err(Error::from)?;
        let Self { base, state, .. } = self;
        let State {
            native,
            owner,
            before,
            inner,
        } = state.into_inner().ok_or_else(|| {
            Failure::Package(crate::package::Error::Corrupted(
                "PPT master-metadata edit state was not initialized".into(),
            ))
        })?;
        if inner.changes().is_empty() {
            return Ok(super::published(base.unchanged()?));
        }
        let (target, _) = inner
            .commit()
            .map_err(master::locate::candidate("master_metadata.commit"))?
            .into_parts();
        Ok(super::published(base.publish_owner(
            &native,
            owner,
            before,
            target.bytes().to_vec(),
        )?))
    }
}

impl Package {
    /// Start a metadata edit of one master: the master of `kind` at a
    /// zero-based `position` among the masters of that kind, in
    /// [`crate::Presentation::masters`] order. Notes and handout masters
    /// accept only position 0.
    ///
    /// # Errors
    ///
    /// An encrypted presentation is refused immediately. A missing master or
    /// unreadable record is reported by the first projection, staging
    /// operation, or commit.
    pub fn edit_master_metadata(&self, kind: Kind, position: Position) -> Result<Edit> {
        Edit::new(self, kind, position)
    }

    /// Apply a master-metadata patch to the exact package it was made from.
    ///
    /// # Errors
    ///
    /// Returns [`super::Error::PatchConflict`] when this package is not the
    /// patch's source.
    pub fn apply_master_metadata(&self, patch: &Patch) -> Result<Commit> {
        super::apply_patch(self, patch)
    }
}
