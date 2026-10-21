//! Committed source-backed workbooks (record 0360, the shape of DOCX
//! record 0355a §1).
//!
//! A source-backed family commit returns a new [`SourceBackedWorkbook`]
//! without writing anything. Its output is the family's byte-preserving
//! source splice, run when the workbook is written
//! ([`SourceBackedWorkbook::write_to`], [`SourceBackedWorkbook::save`]), so a
//! commit followed by a save streams exactly as the former
//! `publish_*_to_stream` routes did: untouched members are raw-copied from
//! the source. A read of the committed workbook (or another edit of it)
//! writes that output into memory once and serves every later read from the
//! reopened artifact.

use std::io::Write;
use std::sync::{Arc, OnceLock};

use litchi_opc::SourceBackedPackage;

use super::SourceBackedWorkbook;
use crate::error::Result;

/// Writes a committed artifact from the base package to a sink.
pub(crate) type Publish = dyn Fn(&SourceBackedPackage, &mut dyn Write) -> Result<()> + Send + Sync;

/// The pending output of a committed workbook.
pub(crate) struct Deferred {
    publish: Box<Publish>,
    materialized: OnceLock<SourceBackedWorkbook>,
}

impl Deferred {
    /// The reopened output, once a read has materialized it.
    pub(super) fn materialized(&self) -> Option<&SourceBackedWorkbook> {
        self.materialized.get()
    }
}

impl SourceBackedWorkbook {
    /// A committed workbook over `self` whose artifact `publish` writes from
    /// `self`'s package.
    ///
    /// A changed write of a signed or encrypted source is refused here,
    /// before any output exists, as the stream writer would refuse it.
    pub(crate) fn deferred(
        &self,
        publish: impl Fn(&SourceBackedPackage, &mut dyn Write) -> Result<()> + Send + Sync + 'static,
    ) -> Result<Self> {
        let base = self.view()?;
        let package = &base.inner.package;
        if package.has_signature_infrastructure() {
            return Err(litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy.into());
        }
        if package.has_encrypted_entries() {
            return Err(litchi_opc::OpcError::SourceBackedOverlayUnavailable {
                reason: "source XML publication refuses encrypted ZIP members".into(),
            }
            .into());
        }
        Ok(Self {
            inner: Arc::clone(&base.inner),
            pending: Some(Arc::new(Deferred {
                publish: Box::new(publish),
                materialized: OnceLock::new(),
            })),
        })
    }

    /// The workbook every read and edit observes: `self` when opened, the
    /// reopened output (written into memory once) when committed.
    pub(crate) fn view(&self) -> Result<&Self> {
        let Some(pending) = &self.pending else {
            return Ok(self);
        };
        if let Some(view) = pending.materialized.get() {
            return Ok(view);
        }
        let mut bytes = Vec::new();
        (pending.publish)(&self.inner.package, &mut bytes)?;
        let view = Self::from_read_at_with_options(
            Arc::new(litchi_core::OwnedSource::new(bytes)),
            self.inner.options.clone(),
        )?;
        Ok(pending.materialized.get_or_init(|| view))
    }

    /// Stream this workbook's complete archive.
    ///
    /// An opened workbook writes its source archive; a committed one writes
    /// its family's source splice (or the bytes it already materialized for
    /// a read). Repeated writes of one workbook are byte-identical. A source
    /// that changed after it was opened is refused before any byte is
    /// written.
    ///
    /// # Errors
    ///
    /// Returns the OPC writer's error, a source-version error, or the
    /// family's publication error (for example a patch conflict).
    pub fn write_to(&self, mut writer: impl Write) -> Result<()> {
        match &self.pending {
            None => Ok(self
                .inner
                .package
                .clone()
                .write_part_overlays_shared_to_stream(&mut writer, Vec::new())?),
            Some(pending) => match pending.materialized.get() {
                Some(view) => view.write_to(writer),
                None => (pending.publish)(&self.inner.package, &mut writer),
            },
        }
    }

    /// Serialize this workbook's complete archive into owned bytes.
    ///
    /// # Errors
    ///
    /// As [`Self::write_to`].
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        let mut bytes = Vec::new();
        self.write_to(&mut bytes)?;
        Ok(bytes)
    }

    /// Atomically save this workbook's archive at a path.
    ///
    /// # Errors
    ///
    /// As [`Self::write_to`], and the file-system errors of the atomic
    /// replacement.
    pub fn save(&self, path: impl AsRef<std::path::Path>) -> Result<()> {
        self.save_with_durability(path, litchi_core::Durability::Full)
    }

    /// Atomically save at a caller-chosen
    /// [`Durability`](litchi_core::Durability).
    ///
    /// # Errors
    ///
    /// As [`Self::save`].
    pub fn save_with_durability(
        &self,
        path: impl AsRef<std::path::Path>,
        durability: litchi_core::Durability,
    ) -> Result<()> {
        litchi_opc::atomic::replace_with_durability::<crate::Error>(
            path.as_ref(),
            durability,
            |sink| self.write_to(sink),
        )
    }
}

impl SourceBackedWorkbook {
    /// The indexed package every source-backed edit reads: the reopened
    /// output of a committed workbook.
    pub(crate) fn edit_package(&self) -> Result<&SourceBackedPackage> {
        let view = self.view()?;
        view.inner.package.check_execution()?;
        Ok(&view.inner.package)
    }
}

/// One source-backed family's replayable change: its detached record patch,
/// which knows its exact before state and how the workbook's output carries
/// its after state.
pub(crate) trait Change: Clone + Send + Sync + 'static {
    /// Whether the change leaves the workbook unchanged.
    fn is_noop(&self) -> bool;
    /// The change that restores the before state.
    #[must_use]
    fn inverse(&self) -> Self;
    /// Refuse a package that does not hold the change's before state.
    fn check(&self, package: &SourceBackedPackage) -> Result<()>;
    /// Check `package` again and write the after artifact: the selected
    /// parts replaced, every other member raw-copied.
    fn write(&self, package: &SourceBackedPackage, writer: &mut dyn Write) -> Result<()>;
}

/// The committed workbook of `change` over `workbook`: `workbook` itself
/// for a no-op, otherwise a deferred workbook writing the change's splice.
pub(crate) fn publish<C: Change>(
    workbook: &SourceBackedWorkbook,
    change: &C,
) -> Result<SourceBackedWorkbook> {
    if change.is_noop() {
        return Ok(workbook.clone());
    }
    let change = change.clone();
    workbook.deferred(move |package, writer| change.write(package, writer))
}

/// Replay `change` onto `workbook`: refuse a workbook without the before
/// state (`PatchConflict`), then publish.
pub(crate) fn replay<C: Change>(
    workbook: &SourceBackedWorkbook,
    change: &C,
) -> Result<SourceBackedWorkbook> {
    change.check(workbook.edit_package()?)?;
    publish(workbook, change)
}

/// Declare a source-backed family's `Patch`, `Commit`, `Diagnostics` and
/// `Error` over its record [`Change`], and its commit and replay helpers.
macro_rules! source_family {
    // A family bounded only by the workbook's source limits.
    ($what:literal, $change:ty) => {
        crate::workbook::source::deferred::source_family!($what, $change, own_limits);

        #[doc = concat!(
                    "The configured budgets of the source-backed ", $what, " family. It has ",
                    "none of its own: the workbook's source limits bound every part it reads."
                )]
        #[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
        #[non_exhaustive]
        pub enum LimitKind {}

        impl std::fmt::Display for LimitKind {
            fn fmt(&self, _formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
                match *self {}
            }
        }
    };
    // A family module that declares (or re-exports) its own `LimitKind`.
    ($what:literal, $change:ty, own_limits) => {
        pub use crate::Error;
        pub use litchi_core::edit::Diagnostics;

        #[doc = concat!(
                    "A replayable ", $what, " change of a source-backed workbook: the exact ",
                    "before and after state of the part it changed. Replay it with the ",
                    "family's `SourceBackedWorkbook::apply_*` onto a workbook holding the ",
                    "before state; [`Patch::inverse`] restores it."
                )]
        #[derive(Clone, Debug)]
        pub struct Patch(pub(crate) $change);

        impl Patch {
            /// Whether applying the patch changes nothing.
            #[must_use]
            pub fn is_noop(&self) -> bool {
                crate::workbook::source::deferred::Change::is_noop(&self.0)
            }

            /// The patch that restores the before state.
            #[must_use]
            pub fn inverse(&self) -> Self {
                Self(crate::workbook::source::deferred::Change::inverse(&self.0))
            }
        }

        #[doc = concat!(
                    "A committed ", $what, " change: the new source-backed workbook, its ",
                    "patch and diagnostics."
                )]
        pub type Commit = litchi_core::edit::Commit<crate::SourceBackedWorkbook, Patch>;

        #[allow(dead_code, reason = "not every family commits from a staged edit")]
        fn commit_change(
            workbook: &crate::SourceBackedWorkbook,
            change: $change,
        ) -> crate::Result<Commit> {
            let next = crate::workbook::source::deferred::publish(workbook, &change)?;
            let changed = !crate::workbook::source::deferred::Change::is_noop(&change);
            Ok(Commit::new(
                next,
                Patch(change),
                litchi_core::edit::Diagnostics::new(changed),
            ))
        }

        fn apply_change(
            workbook: &crate::SourceBackedWorkbook,
            patch: &Patch,
        ) -> crate::Result<Commit> {
            let next = crate::workbook::source::deferred::replay(workbook, &patch.0)?;
            Ok(Commit::new(
                next,
                patch.clone(),
                litchi_core::edit::Diagnostics::new(!patch.is_noop()),
            ))
        }
    };
}

pub(crate) use source_family;
