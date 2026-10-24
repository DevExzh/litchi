//! Package-level master-layout edits (record 0357).
//!
//! [`Package::edit_master_layout`] stages record-tree edits of one master
//! record, [`Edit::commit`] validates the candidate and publishes a new
//! [`Package`] through one append-only user edit that rewrites only that
//! master's persisted record, and [`Package::apply_master_layout`] replays a
//! [`Patch`] onto the exact package it was made from. Dropping an [`Edit`]
//! discards it.

use litchi_core::Position;

use super::model::Limits;
use super::transaction::RecordEdit;
use super::{Change, Commit, Context, Error, Patch, Path, Result, Revision, Snapshot, View};
use crate::Package;
use crate::master::{self, Kind};
use crate::records::Record;
use crate::scoped::{self, Base, Failure, Lazy};
use crate::slide_order::{self, records::Owner};

const OPERATION: &str = "edit_master_layout";

/// A staged record-tree edit of one master over one immutable [`Package`].
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
        let inner =
            Snapshot::parse_with_limits(located.context, &located.record, Limits::default())
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
            let inner =
                Snapshot::parse_with_limits(located.context, &located.record, Limits::default())
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
    #[must_use = "handle the fallible master-layout context result"]
    pub fn context(&self) -> Result<Context> {
        Ok(self.state()?.inner.context())
    }

    /// A borrowed read view of the staged record tree.
    #[must_use = "handle the fallible master-layout view result"]
    pub fn view(&self) -> Result<View<'_>> {
        Ok(self.state()?.inner.view())
    }

    /// The staged master record tree.
    #[must_use = "handle the fallible master-layout record result"]
    pub fn record(&self) -> Result<&Record> {
        Ok(self.state()?.inner.record())
    }

    /// The content revision of the staged record tree.
    #[must_use = "handle the fallible master-layout revision result"]
    pub fn revision(&self) -> Result<Revision> {
        Ok(self.state()?.inner.revision())
    }

    /// The staged structural changes, in order.
    #[must_use = "handle the fallible master-layout changes result"]
    pub fn changes(&self) -> Result<&[Change]> {
        Ok(self.state()?.inner.changes())
    }

    /// Insert one child at a checked position under `parent`.
    ///
    /// # Errors
    ///
    /// Returns [`super::Error::Refused`] when the path, index or record is
    /// refused, leaving the edit unchanged.
    pub fn add(&mut self, parent: impl Into<Path>, index: usize, record: Record) -> Result<()> {
        Ok(self
            .state_mut()?
            .inner
            .add(parent, index, record)
            .map_err(scoped::staging("master_layout.add"))?)
    }

    /// Append one child to a container.
    ///
    /// # Errors
    ///
    /// Returns [`super::Error::Refused`] when the path or record is refused,
    /// leaving the edit unchanged.
    pub fn append(&mut self, parent: impl Into<Path>, record: Record) -> Result<()> {
        Ok(self
            .state_mut()?
            .inner
            .append(parent, record)
            .map_err(scoped::staging("master_layout.append"))?)
    }

    /// Remove one child, returning it. A candidate that no longer forms a
    /// valid master is refused at commit.
    ///
    /// # Errors
    ///
    /// Returns [`super::Error::Refused`] when the path is refused, leaving
    /// the edit unchanged.
    pub fn remove(&mut self, path: impl Into<Path>) -> Result<Record> {
        Ok(self
            .state_mut()?
            .inner
            .remove(path)
            .map_err(scoped::staging("master_layout.remove"))?)
    }

    /// Replace one record in place, including the root when `path` is
    /// empty, returning the replaced record.
    ///
    /// # Errors
    ///
    /// Returns [`super::Error::Refused`] when the path or record is refused,
    /// leaving the edit unchanged.
    pub fn replace(&mut self, path: impl Into<Path>, record: Record) -> Result<Record> {
        Ok(self
            .state_mut()?
            .inner
            .replace(path, record)
            .map_err(scoped::staging("master_layout.replace"))?)
    }

    /// Validate the staged master and publish it as a new package. An edit
    /// with no net change publishes the source package unchanged.
    ///
    /// # Errors
    ///
    /// Returns [`super::Error::Refused`] when the staged record tree is not a
    /// valid master of its context, [`super::Error::Signed`] for a changed
    /// write to a signed package and [`super::Error::Package`] when the
    /// result cannot be republished.
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
                "PPT master-layout edit state was not initialized".into(),
            ))
        })?;
        if inner.changes().is_empty() {
            return Ok(super::published(base.unchanged()?));
        }
        let (target, _) = inner
            .commit()
            .map_err(master::locate::candidate("master_layout.commit"))?
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
    /// Start a record-tree edit of one master: the master of `kind` at a
    /// zero-based `position` among the masters of that kind, in
    /// [`crate::Presentation::masters`] order. Notes and handout masters
    /// accept only position 0.
    ///
    /// # Errors
    ///
    /// An encrypted presentation is refused immediately. A missing master or
    /// unreadable record is reported by the first projection, staging
    /// operation, or commit.
    pub fn edit_master_layout(&self, kind: Kind, position: Position) -> Result<Edit> {
        Edit::new(self, kind, position)
    }

    /// Apply a master-layout patch to the exact package it was made from.
    ///
    /// # Errors
    ///
    /// Returns [`super::Error::PatchConflict`] when this package is not the
    /// patch's source.
    pub fn apply_master_layout(&self, patch: &Patch) -> Result<Commit> {
        super::apply_patch(self, patch)
    }
}
