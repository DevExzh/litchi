//! Immutable, shareable object-package state.
//!
//! A snapshot is the read side of the object owner.  It keeps the package
//! topology and stream allocations shared across clones while leaving all
//! mutation inside [`super::Editor`].  Format crates can therefore retain a
//! stable view between semantic reads without holding a mutable editor.

use super::codec::Package;
use super::editor::Editor;
use super::link::Link;
use super::model::Objects;
use super::target::Targets;
use litchi_cfb::OleError;
use std::sync::Arc;

#[derive(Debug, Clone)]
struct State {
    targets: Targets,
    limits: super::model::Limits,
    original: Arc<Vec<u8>>,
    package: Package,
    objects: Objects,
    changed: bool,
}

/// Immutable, cheap-to-clone view of a captured OLE object package.
///
/// Cloning a snapshot shares the original package bytes and every captured
/// stream allocation.  Only small topology vectors are cloned, so callers can
/// pass read state between format-specific semantic layers without copying
/// payloads.  Use [`Snapshot::edit`] to start an independent transactional
/// edit.
#[derive(Debug, Clone)]
pub struct Snapshot {
    state: Arc<State>,
}

impl Snapshot {
    pub(crate) fn new(
        targets: Targets,
        limits: super::model::Limits,
        original: Arc<Vec<u8>>,
        package: Package,
        objects: Objects,
        changed: bool,
    ) -> Self {
        Self {
            state: Arc::new(State {
                targets,
                limits,
                original,
                package,
                objects,
                changed,
            }),
        }
    }

    /// Opens a package as an immutable snapshot.
    ///
    /// # Errors
    ///
    /// Returns the same bounded CFB and target-validation errors as
    /// [`super::Editor::open`].
    pub fn open(
        bytes: Vec<u8>,
        targets: Targets,
        limits: super::model::Limits,
    ) -> Result<Self, OleError> {
        Ok(Editor::open(bytes, targets, limits)?.snapshot())
    }

    /// The target catalog used by this snapshot.
    #[must_use]
    pub fn targets(&self) -> &Targets {
        &self.state.targets
    }

    /// The current target-selected object catalog.
    #[must_use]
    pub fn objects(&self) -> &Objects {
        &self.state.objects
    }

    /// Whether the snapshot was produced after a committed edit.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.state.changed
    }

    /// Borrows an opaque package stream without copying it.
    #[must_use]
    pub fn stream(&self, path: &[String]) -> Option<&[u8]> {
        self.state.package.stream(path)
    }

    /// Returns shared ownership of an opaque package stream allocation.
    #[must_use]
    pub fn stream_shared(&self, path: &[String]) -> Option<Arc<[u8]>> {
        self.state.package.stream_shared(path)
    }

    /// Parses one selected object's OLEDS `\x01Ole` metadata, when present.
    ///
    /// The returned link is inert and retains unknown wire bytes.
    ///
    /// # Errors
    ///
    /// Returns an error when `key` is absent or its OLEDS stream is malformed
    /// or exceeds the metadata limit.
    pub fn link(&self, key: &str) -> Result<Option<Link>, OleError> {
        self.state
            .objects
            .get(key)
            .ok_or_else(|| OleError::InvalidFormat(format!("object target {key:?} not found")))?
            .link()
    }

    /// Creates an independent transactional editor from this snapshot.
    ///
    /// The editor shares captured stream allocations until a write replaces
    /// one of them, so creating a transaction does not copy large payloads.
    #[must_use]
    pub fn edit(&self) -> Editor {
        Editor::from_snapshot(self)
    }

    /// Renders the snapshot, preserving the original bytes for a true no-op.
    ///
    /// # Errors
    ///
    /// Returns an error when the captured package cannot be rendered.
    pub fn finish(&self) -> Result<Vec<u8>, OleError> {
        if self.state.changed {
            self.state.package.render()
        } else {
            Ok(self.state.original.as_ref().clone())
        }
    }

    pub(crate) fn limits(&self) -> super::model::Limits {
        self.state.limits
    }

    pub(crate) fn original(&self) -> Arc<Vec<u8>> {
        Arc::clone(&self.state.original)
    }

    pub(crate) fn package(&self) -> Package {
        self.state.package.clone()
    }

    pub(crate) fn objects_clone(&self) -> Objects {
        self.state.objects.clone()
    }

    pub(crate) fn changed(&self) -> bool {
        self.state.changed
    }
}
