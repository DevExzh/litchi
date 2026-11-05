//! Inert DDE declaration edits of a [`SourceBackedSpreadsheet`] (record
//! 0358a): formula DDE links and worksheet-local DDE sources, published by
//! replacing only `content.xml`. No DDE source is resolved, refreshed or
//! contacted.
//!
//! [`SourceBackedSpreadsheet::edit_dde`] (or
//! [`SourceBackedSpreadsheet::edit_dde_with`] under explicit limits and an
//! execution context) starts an [`Edit`]; [`Edit::commit`] returns a
//! [`Commit`] holding the new spreadsheet, the reversible [`Patch`] and the
//! [`Diagnostics`]; [`SourceBackedSpreadsheet::apply_dde`] replays a patch.
//! [`SourceBackedSpreadsheet::dde`] reads the inventory as a [`Snapshot`].
//! A changed edit of a signed source is refused at commit.

use std::{fmt, sync::Arc};

use litchi_core::{ExecutionContext, Position, SourceVersion};

use crate::SourceBackedSpreadsheet;
use crate::dde::{
    Edit as Model, Limits, Link, LinkSpec, SheetSelector, SheetSource, Snapshot as Inner, Source,
};

pub use litchi_core::edit::Diagnostics;
pub use litchi_odf_common::{Error, LimitKind};

super::content_patch!("a DDE declaration edit");

/// The result of a DDE edit's `commit`: the new source-backed spreadsheet,
/// the patch and the diagnostics.
pub type Commit = litchi_core::edit::Commit<SourceBackedSpreadsheet, Patch, Diagnostics>;

type Result<T> = std::result::Result<T, Error>;

/// The inert DDE inventory of one [`SourceBackedSpreadsheet`].
///
/// The retained source is parsed lazily: constructing this snapshot shares
/// the source Arc and execution policy, while a read or staged operation
/// initializes the bounded model inventory.
pub struct Snapshot<'source> {
    owner: &'source SourceBackedSpreadsheet,
    source: Arc<str>,
    inner: std::sync::OnceLock<Inner>,
    limits: Limits,
    context: ExecutionContext,
    enforce_context_lineage: bool,
    source_version: SourceVersion,
}

impl fmt::Debug for Snapshot<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Snapshot")
            .field("source_version", &self.source_version)
            .field("source_bytes", &self.source.len())
            .finish_non_exhaustive()
    }
}

impl<'source> Clone for Snapshot<'source> {
    fn clone(&self) -> Self {
        let inner = std::sync::OnceLock::new();
        if let Some(value) = self.inner.get() {
            let _ = inner.set(value.clone());
        }
        Self {
            owner: self.owner,
            source: Arc::clone(&self.source),
            inner,
            limits: self.limits,
            context: self.context.clone(),
            enforce_context_lineage: self.enforce_context_lineage,
            source_version: self.source_version,
        }
    }
}

impl<'source> Snapshot<'source> {
    /// Capture the DDE inventory from one live source owner.
    ///
    /// `enforce_context_lineage` is `false` for the convenience entrypoint,
    /// whose context is created internally, and `true` for an explicit
    /// caller-owned execution context.  The model parser retains its own
    /// finite limits and execution reservations in the resulting snapshot.
    pub(crate) fn from_owner(
        owner: &'source SourceBackedSpreadsheet,
        limits: Limits,
        context: &ExecutionContext,
        enforce_context_lineage: bool,
    ) -> Result<Self> {
        let snapshot = Self::from_owner_lazy(owner, limits, context, enforce_context_lineage)?;
        snapshot.parsed()?;
        Ok(snapshot)
    }

    fn from_owner_lazy(
        owner: &'source SourceBackedSpreadsheet,
        limits: Limits,
        context: &ExecutionContext,
        enforce_context_lineage: bool,
    ) -> Result<Self> {
        limits
            .validate()
            .map_err(litchi_core::Error::from)
            .map_err(Error::from)?;
        context
            .check()
            .map_err(crate::dde::map_execution)
            .map_err(litchi_core::Error::from)
            .map_err(Error::from)?;
        owner.check_source()?;
        let source_version = owner.source_version()?;
        let source = owner.content_xml_arc()?;
        Ok(Self {
            owner,
            source,
            inner: std::sync::OnceLock::new(),
            limits,
            context: context.clone(),
            enforce_context_lineage,
            source_version,
        })
    }

    /// Exact source revision captured by this snapshot.
    #[must_use]
    pub const fn source_version(&self) -> SourceVersion {
        self.source_version
    }

    /// Borrow the exact retained `content.xml` bytes.
    pub fn source_xml(&self) -> &str {
        &self.source
    }

    /// Read formula DDE links in source order.
    pub fn links(&self) -> Result<&[Link]> {
        let value = self.parsed()?.links();
        self.check_source()?;
        Ok(value)
    }

    /// Read worksheet-local DDE declarations in source order.
    pub fn sheet_sources(&self) -> Result<&[SheetSource]> {
        let value = self.parsed()?.sheet_sources();
        self.check_source()?;
        Ok(value)
    }

    /// Read spreadsheet table names in source order when available.
    pub fn table_names(&self) -> Result<&[Option<String>]> {
        let value = self.parsed()?.table_names();
        self.check_source()?;
        Ok(value)
    }

    /// An isolated edit of this inventory, committed under `context`.
    fn edit(&self, context: ExecutionContext) -> Result<Edit<'source>> {
        self.check_source()?;
        context
            .check()
            .map_err(crate::dde::map_execution)
            .map_err(litchi_core::Error::from)
            .map_err(Error::from)?;
        self.check_source()?;
        Ok(Edit {
            before: self.clone(),
            inner: None,
            context,
        })
    }

    fn check_source(&self) -> Result<()> {
        let observed = self.owner.source_version()?;
        if observed != self.source_version {
            return Err(litchi_core::Error::SourceChanged {
                expected: self.source_version,
                observed,
            }
            .into());
        }
        self.context
            .check()
            .map_err(crate::dde::map_execution)
            .map_err(litchi_core::Error::from)
            .map_err(Error::from)
    }

    fn parsed(&self) -> Result<&Inner> {
        self.check_source()?;
        if let Some(inner) = self.inner.get() {
            return Ok(inner);
        }
        // Parse the already retained source Arc directly. This preserves
        // zero-copy link/cache views and avoids a second full `content.xml`
        // allocation in the source adapter. The parse is intentionally
        // deferred until a read or staged operation needs the model.
        let parsed = Inner::parse_shared_with_context(
            Arc::clone(&self.source),
            self.limits,
            &self.context,
            self.enforce_context_lineage,
        )
        .map_err(litchi_core::Error::from)
        .map_err(Error::from)?;
        let _ = self.inner.set(parsed);
        self.check_source()?;
        self.inner.get().ok_or_else(|| {
            Error::from(litchi_core::Error::Other(
                "ODS DDE snapshot initialization failed".to_string(),
            ))
        })
    }

    fn is_signed(&self) -> Result<bool> {
        Ok(self.owner.sheet_metadata_signed()?)
    }
}

/// A staged DDE declaration edit of one [`SourceBackedSpreadsheet`]. Dropping
/// it discards it.
pub struct Edit<'source> {
    before: Snapshot<'source>,
    inner: Option<Model>,
    context: ExecutionContext,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("source_version", &self.before.source_version)
            .finish_non_exhaustive()
    }
}

impl<'source> Edit<'source> {
    /// Snapshot from which this edit was started.
    #[must_use]
    pub const fn before(&self) -> &Snapshot<'source> {
        &self.before
    }

    /// Number of staged formula links.
    pub fn link_count(&self) -> Result<usize> {
        self.before.check_source()?;
        let value = match &self.inner {
            Some(inner) => inner.link_count(),
            None => self.before.links()?.len(),
        };
        self.before.check_source()?;
        Ok(value)
    }

    /// Number of staged worksheet-local source declarations.
    pub fn sheet_source_count(&self) -> Result<usize> {
        self.before.check_source()?;
        let value = match &self.inner {
            Some(inner) => inner.sheet_source_count(),
            None => self.before.sheet_sources()?.len(),
        };
        self.before.check_source()?;
        Ok(value)
    }

    /// Borrow one staged worksheet-local declaration by source order.
    pub fn sheet_source(&self, index: usize) -> Result<Option<(&str, &Source)>> {
        self.before.check_source()?;
        let value = match &self.inner {
            Some(inner) => inner.sheet_source(index),
            None => self
                .before
                .sheet_sources()?
                .get(index)
                .map(|entry| (entry.sheet(), entry.source())),
        };
        self.before.check_source()?;
        Ok(value)
    }

    /// Set or add one worksheet-local source declaration.
    pub fn set_sheet_source<'a>(
        &mut self,
        selector: impl Into<SheetSelector<'a>>,
        source: Source,
    ) -> Result<()> {
        self.stage(|inner| inner.set_sheet_source(selector, source))
    }

    /// Add one worksheet-local source declaration.
    pub fn add_sheet_source<'a>(
        &mut self,
        selector: impl Into<SheetSelector<'a>>,
        source: Source,
    ) -> Result<()> {
        self.stage(|inner| inner.add_sheet_source(selector, source))
    }

    /// Replace an existing worksheet-local source declaration.
    pub fn replace_sheet_source<'a>(
        &mut self,
        selector: impl Into<SheetSelector<'a>>,
        source: Source,
    ) -> Result<()> {
        self.stage(|inner| inner.replace_sheet_source(selector, source))
    }

    /// Remove one worksheet-local source declaration.
    pub fn remove_sheet_source<'a>(
        &mut self,
        selector: impl Into<SheetSelector<'a>>,
    ) -> Result<Source> {
        self.stage(|inner| inner.remove_sheet_source(selector))
    }

    /// Append one formula link.
    pub fn add_link(&mut self, link: LinkSpec) -> Result<()> {
        self.stage(|inner| inner.add_link(link))
    }

    /// Insert one formula link at a source-order position.
    pub fn insert_link(&mut self, index: usize, link: LinkSpec) -> Result<()> {
        self.stage(|inner| inner.insert_link(index, link))
    }

    /// Replace one formula link at a source-order position.
    pub fn replace_link(&mut self, index: impl Into<Position>, link: LinkSpec) -> Result<()> {
        self.stage(|inner| inner.replace_link(index, link))
    }

    /// Replace one formula link at a source-order position.
    pub fn replace_link_at(&mut self, index: impl Into<Position>, link: LinkSpec) -> Result<()> {
        self.stage(|inner| inner.replace_link_at(index, link))
    }

    /// Replace only one formula link's source declaration, preserving its
    /// existing cached table and opaque cache markup.
    pub fn replace_link_source(
        &mut self,
        index: impl Into<Position>,
        source: Source,
    ) -> Result<()> {
        self.stage(|inner| inner.replace_link_source(index, source))
    }

    /// Replace only a uniquely named formula link's source declaration.
    pub fn replace_link_source_named(&mut self, name: &str, source: Source) -> Result<()> {
        self.stage(|inner| inner.replace_link_source_named(name, source))
    }

    /// Remove one formula link at a source-order position.
    pub fn remove_link(&mut self, index: impl Into<Position>) -> Result<()> {
        self.stage(|inner| inner.remove_link(index))
    }

    /// Remove one formula link at a source-order position.
    pub fn remove_link_at(&mut self, index: impl Into<Position>) -> Result<()> {
        self.stage(|inner| inner.remove_link_at(index))
    }

    /// Move one formula link to a source-order position.
    pub fn move_link(&mut self, from: impl Into<Position>, to: impl Into<Position>) -> Result<()> {
        self.stage(|inner| inner.move_link(from, to))
    }

    /// Replace a uniquely named formula link.
    pub fn replace_link_named(&mut self, name: &str, link: LinkSpec) -> Result<()> {
        self.stage(|inner| inner.replace_link_named(name, link))
    }

    /// Remove a uniquely named formula link.
    pub fn remove_link_named(&mut self, name: &str) -> Result<()> {
        self.stage(|inner| inner.remove_link_named(name))
    }

    /// Whether this edit currently retains the exact source bytes.
    pub fn is_noop(&self) -> Result<bool> {
        self.before.check_source()?;
        let value = self.inner.as_ref().is_none_or(Model::is_noop);
        self.before.check_source()?;
        Ok(value)
    }

    /// Discard every staged change.
    pub fn reset(&mut self) {
        self.inner = None;
    }

    /// Validate the staged changes under the edit's execution context and
    /// commit them to a new spreadsheet. Nothing is written.
    ///
    /// # Errors
    ///
    /// Refuses a changed edit of a signed source
    /// ([`RefusalKind::Protected`](litchi_core::edit::RefusalKind::Protected));
    /// returns the validation, limit or execution failure, or a source
    /// change.
    pub fn commit(mut self) -> Result<Commit> {
        self.before.check_source()?;
        let mut inner = match self.inner.take() {
            Some(inner) => inner,
            None => self.before.parsed()?.edit(),
        };
        let commit_result = inner.commit(&self.context);
        self.before.check_source()?;
        let commit = commit_result?;
        if commit.changed() && self.before.is_signed()? {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::Protected,
                "ods.source_backed.dde",
                "signed-source refusal: changed ODS DDE metadata requires explicit unsign/resign policy",
            ));
        }
        self.before.check_source()?;
        let owner = self.before.owner;
        let change = super::Change::new(owner, Arc::clone(commit.snapshot().content_arc()));
        let changed = !change.is_noop();
        let package = if changed {
            owner.with_content(Arc::clone(commit.snapshot().content_arc()))
        } else {
            owner.share()
        };
        Ok(Commit::new(
            package,
            Patch { change },
            Diagnostics::new(changed),
        ))
    }

    fn stage<T>(
        &mut self,
        operation: impl FnOnce(&mut Model) -> litchi_core::Result<T>,
    ) -> Result<T> {
        self.before.check_source()?;
        let result = operation(self.model_mut()?);
        self.before.check_source()?;
        Ok(result?)
    }

    fn model_mut(&mut self) -> Result<&mut Model> {
        if self.inner.is_none() {
            let inner = self.before.parsed()?.edit();
            self.inner = Some(inner);
        }
        self.inner.as_mut().ok_or_else(|| {
            Error::from(litchi_core::Error::Other(
                "ODS DDE edit initialization failed".to_string(),
            ))
        })
    }
}

impl SourceBackedSpreadsheet {
    /// Read the inert DDE inventory under finite default limits.
    ///
    /// # Errors
    ///
    /// Returns the inventory's parse, limit or execution failure, or a
    /// source change.
    pub fn dde(&self) -> Result<Snapshot<'_>> {
        Snapshot::from_owner(
            self.effective()?,
            Limits::default(),
            &crate::dde::default_context(),
            false,
        )
    }

    /// Read the inert DDE inventory under explicit limits and execution
    /// context.
    ///
    /// # Errors
    ///
    /// As [`Self::dde`].
    pub fn dde_with(&self, limits: Limits, context: &ExecutionContext) -> Result<Snapshot<'_>> {
        Snapshot::from_owner(self.effective()?, limits, context, true)
    }

    /// Start a DDE declaration edit under finite default limits.
    ///
    /// # Errors
    ///
    /// As [`Self::dde`].
    pub fn edit_dde(&self) -> Result<Edit<'_>> {
        Snapshot::from_owner_lazy(
            self.effective()?,
            Limits::default(),
            &crate::dde::default_context(),
            false,
        )?
        .edit(crate::dde::default_context())
    }

    /// Start a DDE declaration edit under explicit limits; the inventory is
    /// read, and the edit committed, under `context`.
    ///
    /// # Errors
    ///
    /// As [`Self::dde`].
    pub fn edit_dde_with(&self, limits: Limits, context: &ExecutionContext) -> Result<Edit<'_>> {
        Snapshot::from_owner_lazy(self.effective()?, limits, context, true)?.edit(context.clone())
    }

    /// [`Self::apply_dde`] admitting the patch's target `content.xml` under
    /// explicit destination limits and execution context: its output
    /// ceiling and a full parse under that policy.
    ///
    /// # Errors
    ///
    /// As [`Self::apply_dde`], and the destination policy's limit or execution
    /// failure.
    pub fn apply_dde_with(
        &self,
        patch: &Patch,
        limits: Limits,
        context: &ExecutionContext,
    ) -> Result<Commit> {
        if !patch.is_noop() {
            crate::dde::admit_target(patch.change.after_arc(), limits, context)?;
        }
        self.apply_dde(patch)
    }

    /// Apply a DDE declaration patch to a new spreadsheet; this one is
    /// unchanged.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this spreadsheet is over another
    /// source or does not hold the patch's before `content.xml`, and refuses
    /// a changed patch of a signed source.
    pub fn apply_dde(&self, patch: &Patch) -> Result<Commit> {
        if !patch.is_noop() && self.sheet_metadata_signed()? {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::Protected,
                "ods.source_backed.dde",
                "signed-source refusal: changed ODS DDE patch cannot be applied",
            ));
        }
        let package = patch.change.publish(self, "ods.source_backed.dde.apply")?;
        Ok(Commit::new(
            package,
            patch.clone(),
            Diagnostics::new(!patch.is_noop()),
        ))
    }
}

#[cfg(test)]
mod tests {
    use std::num::{NonZeroU64, NonZeroUsize};

    use super::Limits;
    use litchi_core::{
        Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits as BudgetLimits,
        Profile,
    };

    fn cancelled_context() -> (CancellationSource, ExecutionContext) {
        let (cancellation, token) = CancellationSource::pair();
        let context = ExecutionContext::new(
            Budget::root(
                "ods-dde-lazy-entry",
                BudgetLimits::for_profile(Profile::TrustedBatch),
            ),
            token,
            ExecutionLimits::new(
                NonZeroUsize::new(1).expect("worker"),
                NonZeroUsize::new(1).expect("task"),
                NonZeroU64::new(1024).expect("bytes"),
                0,
            )
            .expect("execution limits"),
        );
        (cancellation, context)
    }

    #[test]
    fn edit_entry_defers_scan_and_drop_keeps_the_owner_unchanged() {
        let owner = super::super::test_support::owner();
        let context = crate::dde::default_context();
        let edit = owner
            .edit_dde_with(Limits::default().with_input_bytes(1), &context)
            .expect("DDE edit construction is source-Arc-only");
        assert!(
            edit.commit().is_err(),
            "the first commit must enforce input limits"
        );

        let edit = owner.edit_dde().expect("DDE edit");
        let source = edit.before().source_xml().to_owned();
        drop(edit);
        assert_eq!(owner.content_xml().expect("content.xml"), source);
    }

    #[test]
    fn edit_entry_admits_limits_and_cancellation_before_parsing() {
        let owner = super::super::test_support::owner();
        assert!(
            owner
                .edit_dde_with(
                    Limits::default().with_links(usize::MAX),
                    &crate::dde::default_context()
                )
                .is_err()
        );

        let (cancellation, context) = cancelled_context();
        cancellation.cancel();
        assert!(owner.edit_dde_with(Limits::default(), &context).is_err());
    }

    #[test]
    fn no_op_commit_preserves_deferred_source_publication() {
        let owner = super::super::test_support::owner();
        let commit = owner.edit_dde().expect("DDE edit").commit().expect("no-op");
        assert!(!commit.diagnostics().changed());
        assert!(commit.patch().is_noop());
        let mut output = Vec::new();
        commit
            .package()
            .write_to(&mut output)
            .expect("source-backed write");
        let mut original = Vec::new();
        owner.write_to(&mut original).expect("source-backed write");
        assert_eq!(output, original);
    }
}
