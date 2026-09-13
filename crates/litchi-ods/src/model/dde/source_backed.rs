//! Source-bound ODS DDE metadata transactions.
//!
//! The model transaction types operate on an immutable `content.xml` string.
//! This module binds those transactions to a live positional ODS owner.  The
//! binding is deliberately explicit: every operation verifies the owner's
//! source version, and a patch also requires the original owner and exact
//! source bytes.  Publication remains a sequential content-only operation;
//! no DDE source is resolved, refreshed, or contacted.

use std::{fmt, io::Write};

use litchi_core::{Error, ExecutionContext, Result, SourceVersion};
use litchi_odf_common::core::{
    SourceContentPublicationError, SourceContentPublicationOptions, SourceContentPublicationReport,
};

use super::{Edit, Limits, Link, LinkSpec, Patch, SheetSelector, SheetSource, Snapshot, Source};
use litchi_core::Position;

/// An immutable DDE inventory tied to one live positional ODS source.
pub struct SourceSnapshot<'source> {
    owner: &'source crate::facade::SourceBackedSpreadsheet,
    inner: Snapshot,
    source_version: SourceVersion,
}

impl fmt::Debug for SourceSnapshot<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SourceSnapshot")
            .field("source_version", &self.source_version)
            .field("source_bytes", &self.inner.source_xml().len())
            .finish_non_exhaustive()
    }
}

impl<'source> Clone for SourceSnapshot<'source> {
    fn clone(&self) -> Self {
        Self {
            owner: self.owner,
            inner: self.inner.clone(),
            source_version: self.source_version,
        }
    }
}

impl<'source> SourceSnapshot<'source> {
    /// Capture the DDE inventory from one live source owner.
    ///
    /// `enforce_context_lineage` is `false` for the convenience entrypoint,
    /// whose context is created internally, and `true` for an explicit
    /// caller-owned execution context.  The model parser retains its own
    /// finite limits and execution reservations in the resulting snapshot.
    pub(crate) fn from_owner(
        owner: &'source crate::facade::SourceBackedSpreadsheet,
        limits: Limits,
        context: &ExecutionContext,
        enforce_context_lineage: bool,
    ) -> Result<Self> {
        owner.check_source()?;
        let source_version = owner.source_version()?;
        let source = owner.content_xml_arc()?;
        // Parse the already retained source Arc directly. This preserves
        // zero-copy link/cache views and avoids a second full `content.xml`
        // allocation in the source adapter.
        let inner =
            Snapshot::parse_shared_with_context(source, limits, context, enforce_context_lineage)
                .map_err(Error::from)?;
        owner.check_source()?;
        Ok(Self {
            owner,
            inner,
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
        self.inner.source_xml()
    }

    /// Read formula DDE links in source order.
    pub fn links(&self) -> Result<&[Link]> {
        self.check_source()?;
        let value = self.inner.links();
        self.check_source()?;
        Ok(value)
    }

    /// Read worksheet-local DDE declarations in source order.
    pub fn sheet_sources(&self) -> Result<&[SheetSource]> {
        self.check_source()?;
        let value = self.inner.sheet_sources();
        self.check_source()?;
        Ok(value)
    }

    /// Read spreadsheet table names in source order when available.
    pub fn table_names(&self) -> Result<&[Option<String>]> {
        self.check_source()?;
        let value = self.inner.table_names();
        self.check_source()?;
        Ok(value)
    }

    /// Start an isolated source-bound edit.
    pub fn edit(&self) -> Result<SourceEdit<'source>> {
        self.check_source()?;
        let inner = self.inner.edit();
        self.check_source()?;
        Ok(SourceEdit {
            before: self.clone(),
            inner,
        })
    }

    fn check_source(&self) -> Result<()> {
        let observed = self.owner.source_version()?;
        if observed != self.source_version {
            return Err(Error::SourceChanged {
                expected: self.source_version,
                observed,
            });
        }
        self.inner
            .context()
            .check()
            .map_err(crate::dde::map_execution)
            .map_err(Error::from)
    }

    fn is_signed(&self) -> Result<bool> {
        self.owner.sheet_metadata_signed()
    }

    fn same_source(&self, other: &Self) -> bool {
        std::ptr::eq(self.owner, other.owner) && self.source_version == other.source_version
    }
}

/// A source-bound DDE edit that forwards operations to the model transaction.
pub struct SourceEdit<'source> {
    before: SourceSnapshot<'source>,
    inner: Edit,
}

impl fmt::Debug for SourceEdit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SourceEdit")
            .field("source_version", &self.before.source_version)
            .finish_non_exhaustive()
    }
}

impl<'source> SourceEdit<'source> {
    /// Snapshot from which this edit was started.
    #[must_use]
    pub const fn before(&self) -> &SourceSnapshot<'source> {
        &self.before
    }

    /// Number of staged formula links.
    pub fn link_count(&self) -> Result<usize> {
        self.before.check_source()?;
        let value = self.inner.link_count();
        self.before.check_source()?;
        Ok(value)
    }

    /// Number of staged worksheet-local source declarations.
    pub fn sheet_source_count(&self) -> Result<usize> {
        self.before.check_source()?;
        let value = self.inner.sheet_source_count();
        self.before.check_source()?;
        Ok(value)
    }

    /// Borrow one staged worksheet-local declaration by source order.
    pub fn sheet_source(&self, index: usize) -> Result<Option<(&str, &Source)>> {
        self.before.check_source()?;
        let value = self.inner.sheet_source(index);
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
        let value = self.inner.is_noop();
        self.before.check_source()?;
        Ok(value)
    }

    /// Commit the staged DDE changes under the retained execution policy.
    pub fn commit(&mut self, context: &ExecutionContext) -> Result<SourceCommit<'source>> {
        self.before.check_source()?;
        let commit_result = self.inner.commit(context);
        self.before.check_source()?;
        let commit = commit_result?;
        if commit.changed() && self.before.is_signed()? {
            return Err(Error::Unsupported(
                "signed-source refusal: changed ODS DDE metadata requires explicit unsign/resign policy"
                    .to_string(),
            ));
        }
        self.before.check_source()?;
        let target = SourceSnapshot {
            owner: self.before.owner,
            inner: commit.snapshot().clone(),
            source_version: self.before.source_version,
        };
        Ok(SourceCommit {
            snapshot: target.clone(),
            patch: SourcePatch {
                before: self.before.clone(),
                target,
                inner: commit.patch().clone(),
            },
            changed: commit.changed(),
        })
    }

    fn stage<T>(&mut self, operation: impl FnOnce(&mut Edit) -> Result<T>) -> Result<T> {
        self.before.check_source()?;
        let result = operation(&mut self.inner);
        self.before.check_source()?;
        result
    }
}

/// A source-bound accepted DDE transaction.
pub struct SourceCommit<'source> {
    snapshot: SourceSnapshot<'source>,
    patch: SourcePatch<'source>,
    changed: bool,
}

impl fmt::Debug for SourceCommit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SourceCommit")
            .field("changed", &self.changed)
            .field("source_bytes", &self.snapshot.inner.source_xml().len())
            .finish_non_exhaustive()
    }
}

impl<'source> SourceCommit<'source> {
    /// Candidate source-bound DDE snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &SourceSnapshot<'source> {
        &self.snapshot
    }

    /// Exact reversible source-bound patch.
    #[must_use]
    pub const fn patch(&self) -> &SourcePatch<'source> {
        &self.patch
    }

    /// Whether the candidate changes `content.xml`.
    #[must_use]
    pub const fn changed(&self) -> bool {
        self.changed
    }

    /// Publish the accepted candidate to a sequential sink.
    pub fn write_to<W: Write>(
        &self,
        writer: W,
        options: SourceContentPublicationOptions,
    ) -> std::result::Result<SourceContentPublicationReport, SourceContentPublicationError> {
        self.snapshot
            .check_source()
            .map_err(publication_source_error)?;
        if self.changed
            && self
                .snapshot
                .is_signed()
                .map_err(SourceContentPublicationError::Core)?
        {
            return Err(SourceContentPublicationError::Core(Error::Unsupported(
                "signed-source refusal: changed ODS DDE metadata cannot be published".to_string(),
            )));
        }
        self.snapshot.owner.write_sheet_metadata_content(
            writer,
            self.snapshot.inner.source_xml().as_bytes(),
            options,
        )
    }

    /// Publish with the default finite source-content policy.
    pub fn write_to_default<W: Write>(
        &self,
        writer: W,
    ) -> std::result::Result<SourceContentPublicationReport, SourceContentPublicationError> {
        self.write_to(writer, SourceContentPublicationOptions::new())
    }
}

/// An exact source-bound DDE patch.
pub struct SourcePatch<'source> {
    before: SourceSnapshot<'source>,
    target: SourceSnapshot<'source>,
    inner: Patch,
}

impl fmt::Debug for SourcePatch<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SourcePatch")
            .field("changed", &self.changed())
            .finish_non_exhaustive()
    }
}

impl<'source> Clone for SourcePatch<'source> {
    fn clone(&self) -> Self {
        Self {
            before: self.before.clone(),
            target: self.target.clone(),
            inner: self.inner.clone(),
        }
    }
}

impl<'source> SourcePatch<'source> {
    /// Exact source snapshot authorized by this patch.
    #[must_use]
    pub const fn source(&self) -> &SourceSnapshot<'source> {
        &self.before
    }

    /// Exact target snapshot produced by this patch.
    #[must_use]
    pub const fn target(&self) -> &SourceSnapshot<'source> {
        &self.target
    }

    /// Exact source `content.xml` bytes authorized by this patch.
    pub fn source_xml(&self) -> &str {
        self.before.source_xml()
    }

    /// Exact target `content.xml` bytes produced by this patch.
    #[must_use]
    pub fn target_xml(&self) -> &str {
        self.target.inner.source_xml()
    }

    /// Whether target bytes differ from source bytes.
    #[must_use]
    pub fn changed(&self) -> bool {
        self.inner.changed()
    }

    /// Whether this patch is an exact no-op.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.inner.is_empty()
    }

    /// Return the exact inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.target.clone(),
            target: self.before.clone(),
            inner: self.inner.inverse(),
        }
    }

    /// Apply this patch to an exact live source snapshot.
    ///
    /// The bare model patch is applied after destination source checks.  That
    /// call reopens the target under the destination snapshot's limits and
    /// execution context, so source-backed patch replay cannot bypass the
    /// destination budget profile.
    pub fn apply(&self, snapshot: &SourceSnapshot<'source>) -> Result<SourceCommit<'source>> {
        snapshot.check_source()?;
        snapshot
            .inner
            .context()
            .check()
            .map_err(crate::dde::map_execution)?;
        if !snapshot.same_source(&self.before) {
            return Err(Error::InvalidFormat(
                "ODS source DDE patch source snapshot does not match".to_string(),
            ));
        }
        if self.changed() && snapshot.is_signed()? {
            return Err(Error::Unsupported(
                "signed-source refusal: changed ODS DDE patch cannot be applied".to_string(),
            ));
        }
        if !self.changed() {
            snapshot.check_source()?;
            return Ok(SourceCommit {
                snapshot: snapshot.clone(),
                patch: self.clone(),
                changed: false,
            });
        }
        let commit = self.inner.apply(&snapshot.inner)?;
        snapshot.check_source()?;
        let target = SourceSnapshot {
            owner: snapshot.owner,
            inner: commit.snapshot().clone(),
            source_version: snapshot.source_version,
        };
        Ok(SourceCommit {
            snapshot: target.clone(),
            patch: SourcePatch {
                before: snapshot.clone(),
                target,
                inner: commit.patch().clone(),
            },
            changed: commit.changed(),
        })
    }
}

fn publication_source_error(error: Error) -> SourceContentPublicationError {
    match error {
        Error::SourceChanged { expected, observed } => {
            SourceContentPublicationError::SourceChanged {
                expected,
                observed,
                progress: litchi_odf_common::core::SourceContentPublicationProgress::Untouched,
            }
        },
        other => SourceContentPublicationError::Core(other),
    }
}
