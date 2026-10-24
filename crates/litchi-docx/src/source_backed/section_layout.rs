//! The section-layout family of a source-backed package (record 0355a):
//! page size, margins, start and columns of one existing section.
//!
//! [`Package::edit_section_layout`](crate::source_backed::Package::edit_section_layout)
//! stages an [`Edit`] whose commit returns a new source-backed package;
//! writing it rewrites only the main-document part.
//! [`Package::apply_section_layout`](crate::source_backed::Package::apply_section_layout)
//! replays a [`Patch`].

use std::cell::{RefCell, RefMut};
use std::sync::Arc;

use crate::Result;
use crate::section::layout as record;
use crate::section::{Columns, Margins, PageSize, Selector, Start};
use crate::source_backed::Package;
use crate::source_backed::deferred::{Change, source_family};

/// The detached main-document patch the family publishes, with the commit
/// that made it while its retained base can still be reused.
#[derive(Clone)]
pub(crate) struct LayoutChange {
    patch: record::Patch,
    commit: Option<Arc<record::Commit>>,
}

impl Change for LayoutChange {
    type Error = crate::Error;

    fn is_noop(&self) -> bool {
        self.patch.is_noop()
    }

    fn inverse(&self) -> Self {
        Self {
            patch: self.patch.inverse(),
            commit: None,
        }
    }

    fn preflight_foreign(&self, package: &Package) -> Result<()> {
        if self.patch.is_noop() {
            return Ok(());
        }
        package.section_layout_snapshot().map(|_| ())
    }

    fn publish(&self, package: &Package) -> Result<Package> {
        let (main, target, fingerprint) = package.section_layout_target(
            &self.patch,
            self.commit.as_deref(),
            "commit_section_layout",
        )?;
        if self.is_noop() {
            return Ok(package.clone());
        }
        let base = package.clone();
        let stream_base = base.clone();
        let stream_main = main.clone();
        let stream_target = Arc::clone(&target);
        package.deferred(
            move |writer, policy| {
                // The committed target was derived from this complete artifact.
                if base.source_artifact_fingerprint()? != fingerprint {
                    return Err(crate::Error::conflict(
                        "commit_section_layout",
                        "section patch artifact fingerprint differs",
                    ));
                }
                base.write_part_overlay(writer, &main, &target, policy)
            },
            move |target_sink| {
                if stream_base.source_artifact_fingerprint()? != fingerprint {
                    return Err(crate::Error::conflict(
                        "commit_section_layout",
                        "section patch artifact fingerprint differs",
                    ));
                }
                stream_base.stream_part_overlay(target_sink, &stream_main, &stream_target)
            },
        )
    }
}

source_family!("section-layout", LayoutChange);

impl From<record::raw::Patch> for Patch {
    /// Replay a detached raw section-layout patch onto a package.  The raw
    /// patch is XML-content scoped; the family wrapper binds the publication
    /// to the complete artifact when `Package::apply_section_layout` is
    /// called, while retaining the old offline authoring route.
    fn from(patch: record::raw::Patch) -> Self {
        Self {
            change: LayoutChange {
                patch: patch.into_record(),
                commit: None,
            },
            restore: None,
        }
    }
}

/// The budgets of this family, reported by
/// [`crate::Error::SectionInventoryLimit`].
pub use crate::section::LimitKind;

/// A staged layout change of one section of a source-backed package. The
/// source is never changed: [`Self::commit`] returns a new package, and
/// dropping the edit discards it.
#[derive(Clone)]
pub struct Edit {
    package: Package,
    selector: Selector,
    limits: crate::section::Limits,
    inner: RefCell<Option<record::Edit>>,
    max_retained_base_bytes: usize,
}

impl std::fmt::Debug for Edit {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.debug_struct("Edit").finish_non_exhaustive()
    }
}

macro_rules! property {
    ($get:ident, $set:ident, $clear:ident, $ty:ty, $what:literal) => {
        #[doc = concat!("The section's ", $what, ".")]
        ///
        /// # Errors
        ///
        /// Returns an error for a value the section cannot represent.
        pub fn $get(&self) -> Result<Option<$ty>> {
            self.inner_ref(|inner| inner.$get())?
        }

        #[doc = concat!("Set or remove the section's ", $what, ".")]
        ///
        /// # Errors
        ///
        /// Returns an error for a value Word refuses.
        pub fn $set(&mut self, value: Option<$ty>) -> Result<&mut Self> {
            self.inner_mut()?.$set(value)?;
            Ok(self)
        }

        #[doc = concat!("Remove the section's ", $what, ".")]
        ///
        /// # Errors
        ///
        /// Returns an error when the section cannot be rewritten.
        pub fn $clear(&mut self) -> Result<&mut Self> {
            self.inner_mut()?.$clear()?;
            Ok(self)
        }
    };
}

impl Edit {
    fn materialize(&self) -> Result<record::Edit> {
        self.package
            .section_layout_snapshot_bounded(&self.limits)?
            .edit(self.selector)
            .map(|edit| edit.with_max_retained_base_bytes(self.max_retained_base_bytes))
    }

    fn ensure_inner(&self) -> Result<()> {
        let mut inner = self.inner.borrow_mut();
        if inner.is_none() {
            *inner = Some(self.materialize()?);
        }
        Ok(())
    }

    fn inner_ref<R>(&self, read: impl FnOnce(&record::Edit) -> R) -> Result<R> {
        self.ensure_inner()?;
        let inner = self.inner.borrow();
        let inner = inner.as_ref().ok_or_else(|| {
            crate::Error::internal(
                "section_layout.edit",
                "section-layout edit materializer is unavailable",
            )
        })?;
        Ok(read(inner))
    }

    fn inner_mut(&self) -> Result<RefMut<'_, record::Edit>> {
        self.ensure_inner()?;
        RefMut::filter_map(self.inner.borrow_mut(), Option::as_mut).map_err(|_| {
            crate::Error::internal(
                "section_layout.edit",
                "section-layout edit materializer is unavailable",
            )
        })
    }

    fn into_inner(mut self) -> Result<record::Edit> {
        if self.inner.get_mut().is_none() {
            let inner = self.materialize()?;
            *self.inner.get_mut() = Some(inner);
        }
        self.inner.into_inner().ok_or_else(|| {
            crate::Error::internal(
                "section_layout.edit",
                "section-layout edit materializer is unavailable",
            )
        })
    }

    /// The package the edit is staged over.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    property!(
        page_size,
        set_page_size,
        clear_page_size,
        PageSize,
        "page size"
    );
    property!(margins, set_margins, clear_margins, Margins, "page margins");
    property!(start, set_start, clear_start, Start, "section start");
    property!(columns, set_columns, clear_columns, Columns, "columns");

    /// Retain at most `maximum` bytes of the base inventory for the commit
    /// (the reuse that avoids a second main-document scan).
    #[must_use]
    pub fn with_max_retained_base_bytes(mut self, maximum: usize) -> Self {
        self.max_retained_base_bytes = maximum;
        if let Some(inner) = self.inner.get_mut().take() {
            *self.inner.get_mut() = Some(inner.with_max_retained_base_bytes(maximum));
        }
        self
    }

    /// Commit the staged layout as a new source-backed package.
    ///
    /// # Errors
    ///
    /// Returns a typed source, protection, signature, XML or limit error
    /// before any output.
    pub fn commit(self) -> Result<Commit> {
        let this = self;
        let package = this.package.clone();
        let commit = this.into_inner()?.commit()?;
        commit_change(
            &package,
            LayoutChange {
                patch: commit.patch().clone(),
                commit: Some(Arc::new(commit)),
            },
        )
    }
}

impl Package {
    /// Stage a layout change of one existing locally authored section.
    ///
    /// # Errors
    ///
    /// The section inventory and selector are admitted on the first getter,
    /// mutation or commit, under the package's retained section limits.
    pub fn edit_section_layout(&self, selector: impl Into<Selector>) -> Result<Edit> {
        let selector = selector.into();
        Ok(Edit {
            package: self.clone(),
            selector,
            limits: self.semantic_limits.sections_value().clone(),
            inner: RefCell::new(None),
            max_retained_base_bytes: 1024 * 1024,
        })
    }

    /// Replay a section-layout patch onto this package, which must hold the
    /// patch's before state.
    ///
    /// # Errors
    ///
    /// Returns a typed [`crate::Error::PatchConflict`] for another
    /// main-document state, and the errors of [`Edit::commit`].
    pub fn apply_section_layout(&self, patch: &Patch) -> Result<Commit> {
        apply_change(self, patch)
    }
}
