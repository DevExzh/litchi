//! Explicit immutable transactions and process-local reversible patches.
//!
//! DOCX main-document operations resolve against preceding staged results.
//! Atomic batches and closure edits group consecutive direct-path formatting
//! sets in one native transaction. Other commands and staged reads close a run.

use std::fmt;
use std::sync::Arc;

use crate::path::{IntoPath, Path};
use crate::{Code, File, IntoNode, IntoPlacement, IntoProps, IntoTarget, Result};

mod limits;
pub use limits::{Error, LimitKind, SemanticLimits, SemanticLimitsBuilder};

/// Content-free facts about a complete explicit transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Diagnostics {
    changed: bool,
    operations: usize,
    commits: usize,
    leaf_operations: usize,
    snapshot_bytes: usize,
}
impl Diagnostics {
    /// Whether the final logical artifact differs from the input.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }
    /// Number of successful mutation calls, including exact no-ops.
    #[must_use]
    pub const fn operations(self) -> usize {
        self.operations
    }
    /// Number of retained native commits. Compatible formatting calls can
    /// share one commit without changing their operation count.
    #[must_use]
    pub const fn commits(self) -> usize {
        self.commits
    }
    /// Sum of semantic operations reported by native commits, saturating on overflow.
    #[must_use]
    pub const fn leaf_operations(self) -> usize {
        self.leaf_operations
    }
    /// Conservative retained native snapshot charge, including repeated shared snapshots.
    #[must_use]
    pub const fn snapshot_bytes(self) -> usize {
        self.snapshot_bytes
    }
}

/// The complete result of [`File::edit`], including the closure's value.
///
/// The fields may be destructured to take every result without cloning.
/// `Debug` omits the closure value and all authored document content.
#[non_exhaustive]
pub struct Commit<T = ()> {
    /// New immutable file. The original remains unchanged.
    pub file: File,
    /// Process-local, owner-bound reversible patch.
    pub patch: Patch,
    /// Content-free transaction facts.
    pub diagnostics: Diagnostics,
    /// Value returned by the closure.
    pub value: T,
}
impl<T> fmt::Debug for Commit<T> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Commit")
            .field("diagnostics", &self.diagnostics)
            .finish_non_exhaustive()
    }
}

/// Opaque, process-local history of validated native commits.
///
/// Applying a patch accepts its original input handle or a cheap clone. A
/// separately opened file is a foreign snapshot even if its document text or
/// bytes match. Native owners, source guards, origin and policy sidecars are
/// retained; applying reuses the already-validated endpoint without a ZIP
/// round trip. Source freshness is checked in both directions.
///
/// Signature consumption precedes editing. Inverse replay restores the
/// authorized unsigned logical source and never recreates a signed state.
/// This is not a durable or cross-process patch format.
#[derive(Clone)]
pub struct Patch {
    inner: Arc<History>,
    inverse: bool,
}
struct History {
    source: File,
    target: File,
    steps: Vec<Step>,
    changed: bool,
}
struct Step {
    before: File,
    after: File,
    leaf: Leaf,
    leaf_operations: usize,
    snapshot_bytes: usize,
}
enum Leaf {
    #[cfg(feature = "doc")]
    Doc(Box<litchi_doc::Patch>),
    #[cfg(feature = "xlsb")]
    Xlsb(Box<crate::format::xlsb::write::Patch>),
    #[cfg(feature = "xls")]
    Xls(Box<litchi_xls::Patch>),
    #[cfg(feature = "pptx")]
    Pptx {
        source: crate::format::pptx::Editable,
        target: crate::format::pptx::Editable,
        _proof: Box<crate::format::pptx::write::Proof>,
    },
    #[cfg(feature = "docx")]
    Docx(Box<litchi_docx::document::Patch>),
    #[cfg(feature = "xlsx")]
    Xlsx {
        source: litchi_xlsx::Workbook,
        target: litchi_xlsx::Workbook,
        // Retain the complete native proof even though process-local replay
        // reuses a validated owner-bound endpoint.
        _patch: Box<litchi_xlsx::workbook::Patch>,
    },
}
impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("steps", &self.inner.steps.len())
            .field("inverse", &self.inverse)
            .finish_non_exhaustive()
    }
}
impl Patch {
    /// Whether the complete transaction preserves the input logical artifact.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        !self.inner.changed
    }
    /// Reverse the ordered history, retaining the same immutable endpoints.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            inner: Arc::clone(&self.inner),
            inverse: !self.inverse,
        }
    }
    /// Apply to the exact bound input (or a cheap clone), returning the retained
    /// validated native endpoint. The input is unchanged on success or failure.
    ///
    /// # Errors
    /// Returns a typed [`Error::PatchConflict`] for another snapshot, protected
    /// state errors for signed/locked input, or source freshness failures.
    pub fn apply(&self, file: &File) -> Result<File> {
        let result = (|| {
            check_source(file)?;
            self.check_sources()?;
            let (source, target) = if self.inverse {
                (&self.inner.target, &self.inner.source)
            } else {
                (&self.inner.source, &self.inner.target)
            };
            if !same_snapshot(file, source) {
                return Err(Error::PatchConflict.into());
            }
            let target = target.clone();
            self.check_sources()?;
            Ok(target)
        })();
        check_source(file)?;
        result
    }
    fn check_sources(&self) -> Result<()> {
        check_source(&self.inner.source)?;
        // Ordered retained endpoints are already validated by their leaf
        // commits. Check every source guard, including a source promoted to
        // an eager unsigned owner; the last endpoint alone is insufficient.
        if self.inverse {
            for step in self.inner.steps.iter().rev() {
                check_source(&step.after)?;
                check_source(&step.before)?;
            }
        } else {
            for step in &self.inner.steps {
                check_source(&step.before)?;
                check_source(&step.after)?;
            }
        }
        check_source(&self.inner.target)
    }
}

/// Atomic transaction proxy, created only by [`File::edit`].
///
/// Every mutation resolves against the preceding immutable result. The first
/// failed mutation poisons the transaction even if its error is caught by the
/// closure; no partial commit can escape. Dropping or unwinding discards it.
/// Reads observe successful changes staged so far.
pub struct Edit {
    identity: Arc<()>,
    current: File,
    steps: Vec<Step>,
    limits: SemanticLimits,
    failure: Option<crate::Error>,
    snapshot_bytes: usize,
    operations: usize,
    #[cfg(feature = "doc")]
    doc: Option<crate::format::doc::write::Run>,
    #[cfg(feature = "docx")]
    formatting: Option<crate::format::docx::write::coalesce::Run>,
    #[cfg(feature = "xlsx")]
    cells: Option<crate::format::xlsx::write::coalesce::Run>,
    #[cfg(feature = "xlsb")]
    xlsb_cells: Option<crate::format::xlsb::write::coalesce::Run>,
    #[cfg(feature = "xlsb")]
    xlsb_sheets: Option<crate::format::xlsb::write::sheet_coalesce::Run>,
    #[cfg(feature = "pptx")]
    shapes: Option<crate::format::pptx::write::coalesce::Run>,
}
impl fmt::Debug for Edit {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("operations", &self.operations)
            .finish_non_exhaustive()
    }
}

impl File {
    /// Edit atomically through a closure, retaining a reversible patch and
    /// diagnostics. Integrates DOCX main-document writes, supported XLSX/XLSB
    /// cell and worksheet writes, immutable XLS cell writes, plus PPTX shape text,
    /// slide/text-box creation, slide ordering, top-level shape removal and
    /// dependency-free slide removal.
    ///
    /// # Errors
    /// Returns the first failed mutation with its zero-based item index, the
    /// closure error, source/protection failures, or a resource-limit error.
    /// The original handle is unchanged, including when the closure panics.
    ///
    /// ```
    /// # #[cfg(feature = "docx")]
    /// # {
    /// use litchi_core::FileFormat;
    /// use litchi_ops::{NewNode, create, props};
    /// let file = create(FileFormat::Docx)?;
    /// let commit = file.edit(|tx| {
    ///     let path = tx.add("/body", NewNode::new("paragraph").prop("text", "Hello"))?;
    ///     tx.set(&path, props! { "align" => "center" })?;
    ///     Ok(path)
    /// })?;
    /// assert!(commit.diagnostics.changed());
    /// let restored = commit.patch.inverse().apply(&commit.file)?;
    /// assert_eq!(restored.to_bytes()?, file.to_bytes()?);
    /// # }
    /// # Ok::<(), litchi_ops::Error>(())
    /// ```
    pub fn edit<T>(&self, f: impl FnOnce(&mut Edit) -> Result<T>) -> Result<Commit<T>> {
        self.edit_with_limits(&SemanticLimits::default(), f)
    }

    /// [`Self::edit`] with caller-selected operation, retained-snapshot and
    /// XLSX/PPTX target-expansion and staging limits.
    ///
    /// # Errors
    /// As [`Self::edit`]. A failed limit check discards the complete candidate.
    pub fn edit_with_limits<T>(
        &self,
        limits: &SemanticLimits,
        f: impl FnOnce(&mut Edit) -> Result<T>,
    ) -> Result<Commit<T>> {
        let result = (|| {
            admit(self)?;
            let identity = Arc::new(());
            let mut edit = Edit {
                identity: Arc::clone(&identity),
                current: self.clone(),
                steps: Vec::new(),
                limits: *limits,
                failure: None,
                snapshot_bytes: 0,
                operations: 0,
                #[cfg(feature = "doc")]
                doc: None,
                #[cfg(feature = "docx")]
                formatting: None,
                #[cfg(feature = "xlsx")]
                cells: None,
                #[cfg(feature = "xlsb")]
                xlsb_cells: None,
                #[cfg(feature = "xlsb")]
                xlsb_sheets: None,
                #[cfg(feature = "pptx")]
                shapes: None,
            };
            let value = f(&mut edit);
            if let Some(error) = edit.failure {
                return Err(error);
            }
            let value = value?;
            if !Arc::ptr_eq(&identity, &edit.identity) {
                return Err(Error::PatchConflict.into());
            }
            edit.flush_pending()?;
            let changed = changed(&edit.steps)?;
            // Canonical no-op collapse: public replay uses these endpoints,
            // never the last step owner. Keep steps only as guarded history.
            let file = if changed { edit.current } else { self.clone() };
            let diagnostics = Diagnostics {
                changed,
                operations: edit.operations,
                commits: edit.steps.len(),
                leaf_operations: edit
                    .steps
                    .iter()
                    .fold(0usize, |sum, step| sum.saturating_add(step.leaf_operations)),
                snapshot_bytes: edit.snapshot_bytes,
            };
            let patch = Patch {
                inner: Arc::new(History {
                    source: self.clone(),
                    target: file.clone(),
                    steps: edit.steps,
                    changed,
                }),
                inverse: false,
            };
            patch.check_sources()?;
            Ok(Commit {
                file,
                patch,
                diagnostics,
                value,
            })
        })();
        check_source(self)?;
        result
    }
}

impl Edit {
    /// Read a required node from the current candidate.
    ///
    /// # Errors
    /// As [`File::get`], or the first previously failed mutation.
    pub fn get<'p>(&mut self, path: impl IntoPath<'p>) -> Result<crate::node::Node<'_>> {
        self.check()?;
        self.flush_pending()?;
        self.current.get(path)
    }
    /// Read an optional node from the current candidate.
    ///
    /// # Errors
    /// As [`File::find`], or the first previously failed mutation.
    pub fn find<'p>(&mut self, path: impl IntoPath<'p>) -> Result<Option<crate::node::Node<'_>>> {
        self.check()?;
        self.flush_pending()?;
        self.current.find(path)
    }
    /// Set properties atomically within this transaction.
    /// Consecutive direct paragraph/run formatting sets share a native edit.
    /// XLS assignments use one isolated native edit per call and retain its
    /// exact reversible patch in the transaction history.
    /// Distinct direct XLSX cell payload assignments also share a native edit.
    /// Distinct direct XLSB cell scalar assignments also share a native edit.
    /// Consecutive direct XLSB sheet visibility assignments to distinct tabs
    /// with one exact state also share one native structural edit. Hiding the
    /// active tab closes its run before the next command so intermediate view
    /// validation retains sequential semantics.
    /// Distinct direct PPTX shape text/geometry sets also share a native edit.
    /// Repeated PPTX shapes close their run before resolving partial properties.
    /// Reads and selectors close pending runs; repeated XLSX cells and font
    /// assignments close the payload run.
    /// A successful call stages its assignments; native commit errors can surface
    /// at the next barrier or when the outer closure finishes, with the original
    /// mutation index. Any mutation or flush failure poisons the transaction.
    ///
    /// # Errors
    /// As [`File::set`], with the mutation's item index and transaction limits.
    pub fn set<'p, 'v>(
        &mut self,
        _target: impl IntoTarget<'p>,
        _props: impl IntoProps<'v>,
    ) -> Result<()> {
        self.set_with_policy(_target, _props, false)
    }

    pub(crate) fn set_command<'p, 'v>(
        &mut self,
        target: impl IntoTarget<'p>,
        props: impl IntoProps<'v>,
    ) -> Result<()> {
        self.set_with_policy(target, props, true)
    }

    fn set_with_policy<'p, 'v>(
        &mut self,
        _target: impl IntoTarget<'p>,
        _props: impl IntoProps<'v>,
        _command_text: bool,
    ) -> Result<()> {
        self.check()?;
        let index = self.operations;
        let result = (|| {
            if index >= self.limits.max_operations() {
                return Err(Error::Limit {
                    kind: LimitKind::Operations,
                    observed: index.saturating_add(1),
                    maximum: self.limits.max_operations(),
                }
                .into());
            }
            let target = match _target.into_target() {
                Ok(target) => target,
                Err(error) => {
                    self.flush_pending()?;
                    return Err(error);
                },
            };
            let props = match crate::Props::from_props(_props) {
                Ok(props) => props,
                Err(error) => {
                    self.flush_pending()?;
                    return Err(error);
                },
            };
            if self
                .coalesce_parsed(&target, &props, index, None)?
                .is_some()
            {
                return Ok(());
            }
            #[cfg(any(
                feature = "doc",
                feature = "xlsx",
                feature = "pptx",
                feature = "xlsb",
                feature = "xls"
            ))]
            let limits = self.limits;
            self.execute(|file| match file {
                #[cfg(feature = "doc")]
                File::Doc(handle) => {
                    wrap_doc(handle, handle.set_with_limits(target, props, &limits)?, ())
                },
                #[cfg(feature = "docx")]
                File::Docx(handle) => {
                    let commit = if _command_text {
                        handle.set_command(target, props)?
                    } else {
                        handle.set(target, props)?
                    };
                    wrap(handle, commit, ())
                },
                #[cfg(feature = "xlsx")]
                File::Xlsx(handle) => {
                    wrap_xlsx(handle, handle.set_with_limits(target, props, &limits)?, ())
                },
                #[cfg(feature = "xlsb")]
                File::Xlsb(handle) => {
                    wrap_xlsb(handle, handle.set_with_limits(target, props, &limits)?, ())
                },
                #[cfg(feature = "xls")]
                File::Xls(handle) => {
                    wrap_xls(handle, handle.set_with_limits(target, props, &limits)?, ())
                },
                #[cfg(feature = "pptx")]
                File::Pptx(handle) => {
                    wrap_pptx(handle, handle.set_with_limits(target, props, &limits)?, ())
                },
                #[allow(unreachable_patterns)]
                _ => Err(pending()),
            })
        })();
        result.map_err(|error| self.poison(error, index))
    }
    /// Add an element and return its current canonical path.
    /// Later structural edits can change positional paths; the path is not a
    /// stable element identity.
    ///
    /// # Errors
    /// As [`File::add`], with the mutation's item index and transaction limits.
    pub fn add<'p, 'n>(
        &mut self,
        _parent: impl IntoPath<'p>,
        _node: impl IntoNode<'n>,
    ) -> Result<Path<'static>> {
        #[cfg(any(feature = "pptx", feature = "xlsb"))]
        let limits = self.limits;
        self.execute(|file| match file {
            #[cfg(feature = "docx")]
            File::Docx(handle) => {
                let (commit, path) = handle.add(_parent.into_path()?, _node.into_node()?)?;
                wrap(handle, commit, path)
            },
            #[cfg(feature = "xlsx")]
            File::Xlsx(handle) => {
                let (commit, path) = handle.add(_parent.into_path()?, _node.into_node()?)?;
                wrap_xlsx(handle, commit, path)
            },
            #[cfg(feature = "pptx")]
            File::Pptx(handle) => {
                let (commit, path) =
                    handle.add_with_limits(_parent.into_path()?, _node.into_node()?, &limits)?;
                wrap_pptx(handle, commit, path)
            },
            #[cfg(feature = "xlsb")]
            File::Xlsb(handle) => {
                let (commit, path) =
                    handle.add_with_limits(_parent.into_path()?, _node.into_node()?, &limits)?;
                wrap_xlsb(handle, commit, path)
            },
            #[allow(unreachable_patterns)]
            _ => Err(pending()),
        })
    }
    pub(crate) fn add_command<'p, 'n>(
        &mut self,
        _parent: impl IntoPath<'p>,
        _node: impl IntoNode<'n>,
    ) -> Result<Path<'static>> {
        #[cfg(any(feature = "pptx", feature = "xlsb"))]
        let limits = self.limits;
        self.execute(|file| match file {
            #[cfg(feature = "docx")]
            File::Docx(handle) => {
                let (commit, path) =
                    handle.add_command(_parent.into_path()?, _node.into_node()?)?;
                wrap(handle, commit, path)
            },
            #[cfg(feature = "xlsx")]
            File::Xlsx(handle) => {
                let (commit, path) = handle.add(_parent.into_path()?, _node.into_node()?)?;
                wrap_xlsx(handle, commit, path)
            },
            #[cfg(feature = "pptx")]
            File::Pptx(handle) => {
                let (commit, path) =
                    handle.add_with_limits(_parent.into_path()?, _node.into_node()?, &limits)?;
                wrap_pptx(handle, commit, path)
            },
            #[cfg(feature = "xlsb")]
            File::Xlsb(handle) => {
                let (commit, path) =
                    handle.add_with_limits(_parent.into_path()?, _node.into_node()?, &limits)?;
                wrap_xlsb(handle, commit, path)
            },
            #[allow(unreachable_patterns)]
            _ => Err(pending()),
        })
    }

    /// Remove selected elements from the current candidate.
    ///
    /// Consecutive direct XLSB cell removals on one worksheet share one
    /// native cell-values edit.  Complete-catalog sheet paths and homogeneous
    /// sheet queries use one native structural edit after any pending direct
    /// cell run is flushed; ranges, heterogeneous queries, and other targets
    /// retain the format handler's precise admission rules.
    ///
    /// # Errors
    /// As [`File::remove`], with the mutation's item index and transaction limits.
    pub fn remove<'p>(&mut self, _target: impl IntoTarget<'p>) -> Result<()> {
        self.check()?;
        let index = self.operations;
        let result = (|| {
            if index >= self.limits.max_operations() {
                return Err(Error::Limit {
                    kind: LimitKind::Operations,
                    observed: index.saturating_add(1),
                    maximum: self.limits.max_operations(),
                }
                .into());
            }
            let target = match _target.into_target() {
                Ok(target) => target,
                Err(error) => {
                    self.flush_pending()?;
                    return Err(error);
                },
            };
            if self.coalesce_remove_parsed(&target, index)?.is_some() {
                return Ok(());
            }
            #[cfg(any(feature = "xlsx", feature = "xlsb", feature = "xls", feature = "pptx"))]
            let limits = self.limits;
            self.execute(|file| match file {
                #[cfg(feature = "docx")]
                File::Docx(handle) => wrap(handle, handle.remove(target.clone())?, ()),
                #[cfg(feature = "xlsx")]
                File::Xlsx(handle) => wrap_xlsx(
                    handle,
                    handle.remove_with_limits(target.clone(), &limits)?,
                    (),
                ),
                #[cfg(feature = "xlsb")]
                File::Xlsb(handle) => wrap_xlsb(
                    handle,
                    handle.remove_with_limits(target.clone(), &limits)?,
                    (),
                ),
                #[cfg(feature = "xls")]
                File::Xls(handle) => wrap_xls(
                    handle,
                    handle.remove_with_limits(target.clone(), &limits)?,
                    (),
                ),
                #[cfg(feature = "pptx")]
                File::Pptx(handle) => wrap_pptx(
                    handle,
                    handle.remove_with_limits(target.clone(), &limits)?,
                    (),
                ),
                #[allow(unreachable_patterns)]
                _ => Err(pending()),
            })
        })();
        result.map_err(|error| self.poison(error, index))
    }

    /// Try to stage a direct XLSB cell removal in the pending native cell
    /// edit.  Range and query removals close the run and use the full
    /// worksheet-grouping path, while consecutive direct removals on one
    /// worksheet share one native transaction.
    fn coalesce_remove_parsed(
        &mut self,
        target: &crate::TargetArg<'_>,
        index: usize,
    ) -> Result<Option<bool>> {
        #[cfg(feature = "xlsb")]
        if let File::Xlsb(_) = &self.current {
            if self.xlsb_sheets.is_some() {
                self.flush_pending()?;
            }
            let planned = match &self.xlsb_cells {
                Some(run) => run.remove_plan(target, &self.limits),
                None => {
                    let File::Xlsb(handle) = &self.current else {
                        return Err(pending());
                    };
                    handle.remove_plan(target, &self.limits)
                },
            };
            if planned.is_err() {
                // An earlier deferred native failure takes precedence over
                // admission failure in the following command.
                self.flush_pending()?;
            }
            let Some(plan) = planned? else {
                return Ok(None);
            };
            if self
                .xlsb_cells
                .as_ref()
                .is_some_and(|run| run.requires_barrier(&plan, &self.limits))
            {
                self.flush_pending()?;
            }
            if self.xlsb_cells.is_none() {
                let File::Xlsb(handle) = &self.current else {
                    return Err(pending());
                };
                self.xlsb_cells = Some(crate::format::xlsb::write::coalesce::Run::new(
                    handle,
                    plan.target.sheet,
                    &self.limits,
                    self.limits
                        .max_snapshot_bytes()
                        .saturating_sub(self.snapshot_bytes),
                )?);
            }
            let run = self
                .xlsb_cells
                .as_mut()
                .ok_or_else(|| crate::Error::new(Code::InternalError, "XLSB cell run missing"))?;
            let changed = run.stage(plan, index)?;
            self.operations += 1;
            return Ok(Some(changed));
        }
        let _ = (target, index);
        Ok(None)
    }
    /// Move an element, resolving its path against the current candidate.
    ///
    /// # Errors
    /// As [`File::move_to`], with the mutation's item index and transaction limits.
    pub fn move_to<'p, 'q>(
        &mut self,
        _path: impl IntoPath<'p>,
        _to: impl IntoPlacement<'q>,
    ) -> Result<()> {
        #[cfg(any(feature = "pptx", feature = "xlsb"))]
        let limits = self.limits;
        self.execute(|file| match file {
            #[cfg(feature = "docx")]
            File::Docx(handle) => wrap(
                handle,
                handle.move_to(_path.into_path()?, _to.into_placement()?)?,
                (),
            ),
            #[cfg(feature = "xlsx")]
            File::Xlsx(handle) => wrap_xlsx(
                handle,
                handle.move_to(_path.into_path()?, _to.into_placement()?)?,
                (),
            ),
            #[cfg(feature = "xlsb")]
            File::Xlsb(handle) => wrap_xlsb(
                handle,
                handle.move_to_with_limits(_path.into_path()?, _to.into_placement()?, &limits)?,
                (),
            ),
            #[cfg(feature = "pptx")]
            File::Pptx(handle) => wrap_pptx(
                handle,
                handle.move_to_with_limits(_path.into_path()?, _to.into_placement()?, &limits)?,
                (),
            ),
            #[allow(unreachable_patterns)]
            _ => Err(pending()),
        })
    }
    /// Swap two elements of the current candidate.
    ///
    /// # Errors
    /// As [`File::swap`], with the mutation's item index and transaction limits.
    pub fn swap<'p, 'q>(&mut self, _a: impl IntoPath<'p>, _b: impl IntoPath<'q>) -> Result<()> {
        #[cfg(any(feature = "pptx", feature = "xlsb"))]
        let limits = self.limits;
        self.execute(|file| match file {
            #[cfg(feature = "docx")]
            File::Docx(handle) => wrap(handle, handle.swap(_a.into_path()?, _b.into_path()?)?, ()),
            #[cfg(feature = "xlsx")]
            File::Xlsx(handle) => {
                wrap_xlsx(handle, handle.swap(_a.into_path()?, _b.into_path()?)?, ())
            },
            #[cfg(feature = "xlsb")]
            File::Xlsb(handle) => wrap_xlsb(
                handle,
                handle.swap_with_limits(_a.into_path()?, _b.into_path()?, &limits)?,
                (),
            ),
            #[cfg(feature = "pptx")]
            File::Pptx(handle) => wrap_pptx(
                handle,
                handle.swap_with_limits(_a.into_path()?, _b.into_path()?, &limits)?,
                (),
            ),
            #[allow(unreachable_patterns)]
            _ => Err(pending()),
        })
    }
    /// Command-channel entry to the shared write coordinator. Only proven
    /// operation families join; other commands first close the pending run.
    pub(crate) fn coalesce_set(
        &mut self,
        _target: &str,
        _props: &crate::Props<'_>,
        _index: usize,
    ) -> Result<Option<bool>> {
        self.check()?;
        let target = match _target.into_target() {
            Ok(target) => target,
            Err(error) => {
                self.flush_pending()?;
                return Err(error);
            },
        };
        self.coalesce_parsed(&target, _props, _index, Some(_target))
    }

    /// Command-channel entry to the shared XLSB cell-removal coordinator.
    /// Only a direct cell path joins the pending native run; ranges, queries,
    /// and other formats return `None` so the ordinary dispatcher can handle
    /// them after closing the run.
    pub(crate) fn coalesce_remove(&mut self, _target: &str, _index: usize) -> Result<Option<bool>> {
        self.check()?;
        let target = match _target.into_target() {
            Ok(target) => target,
            Err(error) => {
                self.flush_pending()?;
                return Err(error);
            },
        };
        self.coalesce_remove_parsed(&target, _index)
    }

    fn coalesce_parsed(
        &mut self,
        _target: &crate::TargetArg<'_>,
        _props: &crate::Props<'_>,
        _index: usize,
        _spelling: Option<&str>,
    ) -> Result<Option<bool>> {
        #[cfg(feature = "doc")]
        if let File::Doc(handle) = self.current.clone() {
            // A body paragraph's ordinal is stable across text replacement,
            // so direct `text` assignments can share one variable-width leaf
            // edit. Queries and non-text properties close the run before the
            // ordinary setter resolves them against the published candidate.
            let planned = match self.doc.as_ref() {
                Some(run) => run.plan(_spelling, _target, _props),
                None => handle.paragraph_plan(_target, _props),
            };
            if planned.is_err() {
                self.flush_pending()?;
            }
            let Some(plan) = planned? else {
                if self.doc.is_some() {
                    self.flush_pending()?;
                }
                return Ok(None);
            };
            if self.operations >= self.limits.max_operations() {
                return Err(Error::Limit {
                    kind: LimitKind::Operations,
                    observed: self.operations.saturating_add(1),
                    maximum: self.limits.max_operations(),
                }
                .into());
            }
            if self.doc.is_none() {
                self.doc = Some(crate::format::doc::write::Run::new(&handle, &self.limits)?);
            }
            let run = self
                .doc
                .as_mut()
                .ok_or_else(|| crate::Error::new(Code::InternalError, "DOC body run missing"))?;
            let changed = run.stage(plan, _index, _spelling)?;
            let bytes = self.snapshot_bytes.saturating_add(run.snapshot_bytes());
            if bytes > self.limits.max_snapshot_bytes() {
                return Err(Error::Limit {
                    kind: LimitKind::SnapshotBytes,
                    observed: bytes,
                    maximum: self.limits.max_snapshot_bytes(),
                }
                .into());
            }
            self.operations += 1;
            return Ok(Some(changed));
        }
        #[cfg(feature = "docx")]
        if let File::Docx(handle) = &self.current {
            // The shared paragraph/run `text` key always closes this family.
            // Let its ordinary handler resolve and admit once, including mixed
            // text/formatting calls. The native setter check remains defensive
            // for any future text alias that is not recognized by this hint.
            if _props
                .iter()
                .any(|(key, _)| key.eq_ignore_ascii_case("text"))
            {
                return Ok(None);
            }
            let planned = match (&self.formatting, _spelling) {
                (Some(run), Some(spelling)) => run.formatting_plan(spelling, _target, _props),
                _ => handle.formatting_plan(_target, _props),
            };
            let Some(plan) = planned? else {
                return Ok(None);
            };
            if self.operations >= self.limits.max_operations() {
                return Err(Error::Limit {
                    kind: LimitKind::Operations,
                    observed: self.operations.saturating_add(1),
                    maximum: self.limits.max_operations(),
                }
                .into());
            }
            if self
                .formatting
                .as_ref()
                .is_some_and(|run| plan.operation_bound() > run.remaining_operations())
            {
                self.flush_pending()?;
            }
            if self.formatting.is_none() {
                let File::Docx(handle) = &self.current else {
                    return Err(pending());
                };
                self.formatting = Some(crate::format::docx::write::coalesce::Run::new(handle)?);
            }
            let run = self
                .formatting
                .as_mut()
                .ok_or_else(|| crate::Error::new(Code::InternalError, "formatting run missing"))?;
            let changed = run.stage(plan, _index, _spelling)?;
            let bytes = self.snapshot_bytes.saturating_add(run.snapshot_bytes());
            if bytes > self.limits.max_snapshot_bytes() {
                return Err(Error::Limit {
                    kind: LimitKind::SnapshotBytes,
                    observed: bytes,
                    maximum: self.limits.max_snapshot_bytes(),
                }
                .into());
            }
            self.operations += 1;
            return Ok(Some(changed));
        }
        #[cfg(feature = "xlsb")]
        if let File::Xlsb(_) = &self.current {
            let sheet_shape = crate::format::xlsb::write::sheet_state_shape(_target, _props);
            if sheet_shape && self.xlsb_cells.is_some() {
                // Sheet resolution must see the candidate after a pending
                // cell run; the two native owners cannot share a transaction.
                self.flush_pending()?;
            }
            if sheet_shape {
                let mut planned = {
                    let File::Xlsb(handle) = &self.current else {
                        return Err(pending());
                    };
                    match self.xlsb_sheets.as_ref() {
                        Some(run) => run.sheet_plan(_target, _props, &self.limits),
                        None => handle.sheet_plan(_target, _props, &self.limits),
                    }
                };
                if planned.is_err() {
                    // An earlier deferred native failure takes precedence over
                    // this command's setter admission error.
                    self.flush_pending()?;
                    let File::Xlsb(handle) = &self.current else {
                        return Err(pending());
                    };
                    planned = handle.sheet_plan(_target, _props, &self.limits);
                }
                let Some(mut plan) = planned? else {
                    return Ok(None);
                };
                if self.operations >= self.limits.max_operations() {
                    return Err(Error::Limit {
                        kind: LimitKind::Operations,
                        observed: self.operations.saturating_add(1),
                        maximum: self.limits.max_operations(),
                    }
                    .into());
                }
                let barrier = (|| -> Result<Option<(bool, bool)>> {
                    let Some(run) = self.xlsb_sheets.as_ref() else {
                        return Ok(None);
                    };
                    let requires_barrier = run.requires_barrier(&plan, &self.limits);
                    let hides_last_visible = (!requires_barrier)
                        .then(|| run.would_hide_last_visible(&plan))
                        .transpose()?
                        .unwrap_or(false);
                    Ok(Some((requires_barrier, hides_last_visible)))
                })();
                if barrier.is_err() {
                    self.flush_pending()?;
                }
                let barrier = barrier?;
                if let Some((requires_barrier, hides_last_visible)) = barrier {
                    if requires_barrier {
                        self.flush_pending()?;
                        let File::Xlsb(handle) = &self.current else {
                            return Err(pending());
                        };
                        plan = handle
                            .sheet_plan(_target, _props, &self.limits)?
                            .ok_or_else(|| {
                                crate::Error::new(
                                    Code::InternalError,
                                    "XLSB sheet visibility plan disappeared after flush",
                                )
                                .with_format(litchi_core::FileFormat::Xlsb)
                            })?;
                    } else if hides_last_visible {
                        let error = crate::Error::new(
                            Code::InvalidArgument,
                            "a workbook must retain a visible sheet",
                        )
                        .with_format(litchi_core::FileFormat::Xlsb);
                        // Flush first so a deferred native view/source failure
                        // from an earlier item wins over this admission error.
                        self.flush_pending()?;
                        return Err(error);
                    }
                }
                if self.xlsb_sheets.is_none() {
                    let File::Xlsb(handle) = &self.current else {
                        return Err(pending());
                    };
                    self.xlsb_sheets = Some(crate::format::xlsb::write::sheet_coalesce::Run::new(
                        handle,
                        plan.state,
                        &self.limits,
                        self.limits
                            .max_snapshot_bytes()
                            .saturating_sub(self.snapshot_bytes),
                    )?);
                }
                let (had_prior, staged) = {
                    let run = self.xlsb_sheets.as_mut().ok_or_else(|| {
                        crate::Error::new(Code::InternalError, "XLSB sheet run missing")
                    })?;
                    let had_prior = run.first_index().is_some();
                    (had_prior, run.stage(plan, _index))
                };
                let changed = match staged {
                    Ok(changed) => changed,
                    Err(error) => {
                        if had_prior {
                            self.flush_pending()?;
                        }
                        return Err(error);
                    },
                };
                self.operations += 1;
                return Ok(Some(changed));
            }

            if self.xlsb_sheets.is_some() {
                // A cell plan must resolve against the package published by
                // the sheet run, never against its stale pre-edit snapshot.
                self.flush_pending()?;
            }
            let planned = match &self.xlsb_cells {
                Some(run) => run.cell_plan(_target, _props, &self.limits),
                None => {
                    let File::Xlsb(handle) = &self.current else {
                        return Err(pending());
                    };
                    handle.cell_plan(_target, _props, &self.limits)
                },
            };
            if planned.is_err() {
                // An earlier deferred native failure takes precedence over
                // admission failure in the following command.
                self.flush_pending()?;
            }
            let Some(plan) = planned? else {
                return Ok(None);
            };
            if self.operations >= self.limits.max_operations() {
                return Err(Error::Limit {
                    kind: LimitKind::Operations,
                    observed: self.operations.saturating_add(1),
                    maximum: self.limits.max_operations(),
                }
                .into());
            }
            if self
                .xlsb_cells
                .as_ref()
                .is_some_and(|run| run.requires_barrier(&plan, &self.limits))
            {
                self.flush_pending()?;
            }
            if self.xlsb_cells.is_none() {
                let File::Xlsb(handle) = &self.current else {
                    return Err(pending());
                };
                self.xlsb_cells = Some(crate::format::xlsb::write::coalesce::Run::new(
                    handle,
                    plan.target.sheet,
                    &self.limits,
                    self.limits
                        .max_snapshot_bytes()
                        .saturating_sub(self.snapshot_bytes),
                )?);
            }
            let run = self
                .xlsb_cells
                .as_mut()
                .ok_or_else(|| crate::Error::new(Code::InternalError, "XLSB cell run missing"))?;
            let changed = run.stage(plan, _index)?;
            self.operations += 1;
            return Ok(Some(changed));
        }
        #[cfg(feature = "xlsx")]
        if let File::Xlsx(handle) = &self.current {
            let planned = match &self.cells {
                Some(run) => run.cell_plan(_target, _props, &self.limits),
                None => handle.cell_plan(_target, _props, &self.limits),
            };
            if planned.is_err() {
                // An earlier deferred native failure takes precedence over
                // admission failure in the following command.
                self.flush_pending()?;
            }
            let Some(plan) = planned? else {
                return Ok(None);
            };
            if self.operations >= self.limits.max_operations() {
                return Err(Error::Limit {
                    kind: LimitKind::Operations,
                    observed: self.operations.saturating_add(1),
                    maximum: self.limits.max_operations(),
                }
                .into());
            }
            if self
                .cells
                .as_ref()
                .is_some_and(|run| run.requires_barrier(&plan, &self.limits))
            {
                self.flush_pending()?;
            }
            if self.cells.is_none() {
                let File::Xlsx(handle) = &self.current else {
                    return Err(pending());
                };
                self.cells = Some(crate::format::xlsx::write::coalesce::Run::new(
                    handle,
                    &self.limits,
                )?);
            }
            let run = self
                .cells
                .as_mut()
                .ok_or_else(|| crate::Error::new(Code::InternalError, "cell run missing"))?;
            let changed = run.stage(plan, _index)?;
            self.operations += 1;
            return Ok(Some(changed));
        }
        #[cfg(feature = "pptx")]
        if let File::Pptx(handle) = &self.current {
            let planned = match &mut self.shapes {
                Some(run) => run.shape_plan(_target, _props, &self.limits),
                None => handle.shape_plan(_target, _props, &self.limits),
            };
            if planned.is_err() {
                // Preserve an earlier deferred native failure's wire index.
                self.flush_pending()?;
            }
            let Some(mut plan) = planned? else {
                return Ok(None);
            };
            if self.operations >= self.limits.max_operations() {
                return Err(Error::Limit {
                    kind: LimitKind::Operations,
                    observed: self.operations.saturating_add(1),
                    maximum: self.limits.max_operations(),
                }
                .into());
            }
            if self
                .shapes
                .as_ref()
                .is_some_and(|run| run.requires_barrier(&plan, &self.limits))
            {
                self.flush_pending()?;
                // A repeated shape may merge partial geometry with earlier
                // assignments. Resolve it again against the published run.
                let File::Pptx(handle) = &self.current else {
                    return Err(pending());
                };
                let Some(replanned) = handle.shape_plan(_target, _props, &self.limits)? else {
                    return Ok(None);
                };
                plan = replanned;
            }
            if self.shapes.is_none() {
                let File::Pptx(handle) = &self.current else {
                    return Err(pending());
                };
                self.shapes = Some(crate::format::pptx::write::coalesce::Run::new(
                    handle,
                    &self.limits,
                )?);
            }
            let run = self
                .shapes
                .as_mut()
                .ok_or_else(|| crate::Error::new(Code::InternalError, "shape run missing"))?;
            let changed = run.stage(plan, _index)?;
            self.operations += 1;
            return Ok(Some(changed));
        }
        Ok(None)
    }

    pub(crate) fn flush_pending(&mut self) -> Result<()> {
        #[cfg(feature = "doc")]
        if let Some(run) = self.doc.take() {
            let index = run.first_index();
            let result = (|| -> Result<()> {
                let commit = run.commit()?;
                let (step, ()) = wrap_doc(&commit.before.clone(), commit, ())?;
                let bytes = self.snapshot_bytes.saturating_add(step.snapshot_bytes);
                if bytes > self.limits.max_snapshot_bytes() {
                    return Err(Error::Limit {
                        kind: LimitKind::SnapshotBytes,
                        observed: bytes,
                        maximum: self.limits.max_snapshot_bytes(),
                    }
                    .into());
                }
                self.steps
                    .try_reserve(1)
                    .map_err(|source| Error::Allocation { source })?;
                self.current = step.after.clone();
                self.steps.push(step);
                self.snapshot_bytes = bytes;
                Ok(())
            })();
            result.map_err(|error| self.poison(error, index.unwrap_or(self.operations)))?;
        }
        #[cfg(feature = "docx")]
        if let Some(run) = self.formatting.take() {
            let index = run.first_index();
            let result = (|| -> Result<()> {
                let (base, commit) = run.commit()?;
                let (step, ()) = wrap(&base, commit, ())?;
                let bytes = self.snapshot_bytes.saturating_add(step.snapshot_bytes);
                if bytes > self.limits.max_snapshot_bytes() {
                    return Err(Error::Limit {
                        kind: LimitKind::SnapshotBytes,
                        observed: bytes,
                        maximum: self.limits.max_snapshot_bytes(),
                    }
                    .into());
                }
                self.steps
                    .try_reserve(1)
                    .map_err(|source| Error::Allocation { source })?;
                self.current = step.after.clone();
                self.steps.push(step);
                self.snapshot_bytes = bytes;
                Ok(())
            })();
            result.map_err(|error| self.poison(error, index.unwrap_or(self.operations)))?;
        }
        #[cfg(feature = "xlsb")]
        if let Some(run) = self.xlsb_sheets.take() {
            let index = run.first_index();
            let result = (|| -> Result<()> {
                let (base, commit) = run.commit()?;
                let (step, ()) = wrap_xlsb(&base, commit, ())?;
                let bytes = self.snapshot_bytes.saturating_add(step.snapshot_bytes);
                if bytes > self.limits.max_snapshot_bytes() {
                    return Err(Error::Limit {
                        kind: LimitKind::SnapshotBytes,
                        observed: bytes,
                        maximum: self.limits.max_snapshot_bytes(),
                    }
                    .into());
                }
                self.steps
                    .try_reserve(1)
                    .map_err(|source| Error::Allocation { source })?;
                self.current = step.after.clone();
                self.steps.push(step);
                self.snapshot_bytes = bytes;
                Ok(())
            })();
            result.map_err(|error| self.poison(error, index.unwrap_or(self.operations)))?;
        }
        #[cfg(feature = "xlsb")]
        if let Some(run) = self.xlsb_cells.take() {
            let index = run.first_index();
            let result = (|| -> Result<()> {
                let (base, commit) = run.commit()?;
                let (step, ()) = wrap_xlsb(&base, commit, ())?;
                let bytes = self.snapshot_bytes.saturating_add(step.snapshot_bytes);
                if bytes > self.limits.max_snapshot_bytes() {
                    return Err(Error::Limit {
                        kind: LimitKind::SnapshotBytes,
                        observed: bytes,
                        maximum: self.limits.max_snapshot_bytes(),
                    }
                    .into());
                }
                self.steps
                    .try_reserve(1)
                    .map_err(|source| Error::Allocation { source })?;
                self.current = step.after.clone();
                self.steps.push(step);
                self.snapshot_bytes = bytes;
                Ok(())
            })();
            result.map_err(|error| self.poison(error, index.unwrap_or(self.operations)))?;
        }
        #[cfg(feature = "xlsx")]
        if let Some(run) = self.cells.take() {
            let index = run.first_index();
            let result = (|| -> Result<()> {
                let (base, commit) = run.commit()?;
                let (step, ()) = wrap_xlsx(&base, commit, ())?;
                let bytes = self.snapshot_bytes.saturating_add(step.snapshot_bytes);
                if bytes > self.limits.max_snapshot_bytes() {
                    return Err(Error::Limit {
                        kind: LimitKind::SnapshotBytes,
                        observed: bytes,
                        maximum: self.limits.max_snapshot_bytes(),
                    }
                    .into());
                }
                self.steps
                    .try_reserve(1)
                    .map_err(|source| Error::Allocation { source })?;
                self.current = step.after.clone();
                self.steps.push(step);
                self.snapshot_bytes = bytes;
                Ok(())
            })();
            result.map_err(|error| self.poison(error, index.unwrap_or(self.operations)))?;
        }
        #[cfg(feature = "pptx")]
        if let Some(run) = self.shapes.take() {
            let index = run.first_index();
            let result = (|| -> Result<()> {
                let (base, commit) = run.commit()?;
                let (step, ()) = wrap_pptx(&base, commit, ())?;
                let bytes = self.snapshot_bytes.saturating_add(step.snapshot_bytes);
                if bytes > self.limits.max_snapshot_bytes() {
                    return Err(Error::Limit {
                        kind: LimitKind::SnapshotBytes,
                        observed: bytes,
                        maximum: self.limits.max_snapshot_bytes(),
                    }
                    .into());
                }
                self.steps
                    .try_reserve(1)
                    .map_err(|source| Error::Allocation { source })?;
                self.current = step.after.clone();
                self.steps.push(step);
                self.snapshot_bytes = bytes;
                Ok(())
            })();
            result.map_err(|error| self.poison(error, index.unwrap_or(self.operations)))?;
        }
        Ok(())
    }

    /// Isolate one command without allowing its poison or candidate to escape.
    /// The outer closure still owns the complete batch and its final source fence.
    pub(crate) fn isolated<T>(
        &mut self,
        operation: impl FnOnce(&mut Self) -> Result<T>,
    ) -> Result<(T, bool)> {
        self.check()?;
        self.flush_pending()?;
        let start = self.steps.len();
        let operations = self.operations;
        let before = self.current.clone();
        let bytes = self.snapshot_bytes;
        let result = operation(self).and_then(|value| {
            self.check()?;
            self.flush_pending()?;
            Ok((value, changed(&self.steps[start..])?))
        });
        let result = match self.failure.take() {
            Some(error) => Err(error),
            None => result,
        };
        match result {
            Ok(value) => Ok(value),
            Err(error) => {
                #[cfg(feature = "doc")]
                {
                    self.doc = None;
                }
                #[cfg(feature = "docx")]
                {
                    self.formatting = None;
                }
                #[cfg(feature = "xlsx")]
                {
                    self.cells = None;
                }
                #[cfg(feature = "xlsb")]
                {
                    self.xlsb_cells = None;
                    self.xlsb_sheets = None;
                }
                #[cfg(feature = "pptx")]
                {
                    self.shapes = None;
                }
                self.current = before;
                self.steps.truncate(start);
                self.snapshot_bytes = bytes;
                self.operations = operations;
                Err(error)
            },
        }
    }

    fn poison(&mut self, error: crate::Error, index: usize) -> crate::Error {
        let error = self.failure.clone().unwrap_or_else(|| {
            if error.item().is_some() {
                error
            } else {
                error.with_item(index)
            }
        });
        self.failure = Some(error.clone());
        error
    }

    fn check(&self) -> Result<()> {
        self.failure.clone().map_or(Ok(()), Err)
    }
    fn execute<T>(&mut self, operation: impl FnOnce(&File) -> Result<(Step, T)>) -> Result<T> {
        self.check()?;
        self.flush_pending()?;
        let index = self.operations;
        let result = (|| {
            if index >= self.limits.max_operations() {
                return Err(Error::Limit {
                    kind: LimitKind::Operations,
                    observed: index.saturating_add(1),
                    maximum: self.limits.max_operations(),
                }
                .into());
            }
            let (step, value) = operation(&self.current)?;
            let bytes = self.snapshot_bytes.saturating_add(step.snapshot_bytes);
            if bytes > self.limits.max_snapshot_bytes() {
                return Err(Error::Limit {
                    kind: LimitKind::SnapshotBytes,
                    observed: bytes,
                    maximum: self.limits.max_snapshot_bytes(),
                }
                .into());
            }
            self.steps
                .try_reserve(1)
                .map_err(|source| Error::Allocation { source })?;
            self.current = step.after.clone();
            self.steps.push(step);
            self.snapshot_bytes = bytes;
            self.operations += 1;
            Ok(value)
        })();
        result.map_err(|error| self.poison(error, index))
    }
}

#[cfg(feature = "docx")]
fn wrap<T>(
    before: &crate::format::docx::Docx,
    commit: crate::format::docx::write::Commit,
    value: T,
) -> Result<(Step, T)> {
    let (after, main, diagnostics) = commit.into_parts();
    let snapshot_bytes = main
        .source()
        .xml_bytes()
        .len()
        .saturating_add(main.target().xml_bytes().len());
    Ok((
        Step {
            before: File::Docx(before.clone()),
            after: File::Docx(after),
            leaf: Leaf::Docx(Box::new(main)),
            leaf_operations: diagnostics.operations(),
            snapshot_bytes,
        },
        value,
    ))
}

#[cfg(feature = "doc")]
fn wrap_doc<T>(
    _before: &crate::format::doc::Doc,
    commit: crate::format::doc::WriteCommit,
    value: T,
) -> Result<(Step, T)> {
    let crate::format::doc::WriteCommit {
        before: committed_before,
        after,
        patch,
        operations,
        snapshot_bytes,
    } = commit;
    Ok((
        Step {
            before: File::Doc(committed_before),
            after: File::Doc(after),
            leaf: Leaf::Doc(Box::new(patch)),
            leaf_operations: operations,
            snapshot_bytes,
        },
        value,
    ))
}

#[cfg(feature = "xlsb")]
fn wrap_xlsb<T>(
    before: &crate::format::xlsb::Xlsb,
    commit: crate::format::xlsb::write::Commit,
    value: T,
) -> Result<(Step, T)> {
    let (after, patch) = commit.into_parts();
    let leaf_operations = patch.operation_count();
    // Structural and cell-value patches both retain exact package images.
    // Cell selectors spanning worksheets retain every per-worksheet image;
    // structural patches carry the checked native operation charge.
    let snapshot_bytes = patch.retained_bytes();
    Ok((
        Step {
            before: File::Xlsb(before.clone()),
            after: File::Xlsb(after),
            leaf: Leaf::Xlsb(Box::new(patch)),
            leaf_operations,
            snapshot_bytes,
        },
        value,
    ))
}

#[cfg(feature = "xls")]
fn wrap_xls<T>(
    _before: &crate::format::xls::Xls,
    commit: crate::format::xls::write::WriteCommit,
    value: T,
) -> Result<(Step, T)> {
    let crate::format::xls::write::WriteCommit {
        before,
        after,
        patch,
        operations,
        snapshot_bytes,
    } = commit;
    Ok((
        Step {
            before: File::Xls(before),
            after: File::Xls(after),
            leaf: Leaf::Xls(Box::new(patch)),
            leaf_operations: operations,
            snapshot_bytes,
        },
        value,
    ))
}

#[cfg(feature = "xlsx")]
fn wrap_xlsx<T>(
    before: &crate::format::xlsx::Xlsx,
    commit: crate::format::xlsx::write::Commit,
    value: T,
) -> Result<(Step, T)> {
    let (after, source, target, patch) = commit.into_parts();
    let snapshot_bytes = patch.retained_bytes();
    let leaf_operations = patch.len();
    Ok((
        Step {
            before: File::Xlsx(before.clone()),
            after: File::Xlsx(after),
            leaf: Leaf::Xlsx {
                source,
                target,
                _patch: Box::new(patch),
            },
            leaf_operations,
            snapshot_bytes,
        },
        value,
    ))
}

#[cfg(feature = "pptx")]
fn wrap_pptx<T>(
    before: &crate::format::pptx::Pptx,
    commit: crate::format::pptx::write::Commit,
    value: T,
) -> Result<(Step, T)> {
    let snapshot_bytes = commit.proof.retained_bytes();
    Ok((
        Step {
            before: File::Pptx(before.clone()),
            after: File::Pptx(commit.after),
            leaf: Leaf::Pptx {
                source: commit.source,
                target: commit.target,
                _proof: Box::new(commit.proof),
            },
            leaf_operations: commit.operations,
            snapshot_bytes,
        },
        value,
    ))
}

fn changed(steps: &[Step]) -> Result<bool> {
    let (Some(first), Some(last)) = (steps.first(), steps.last()) else {
        return Ok(false);
    };
    match (&first.leaf, &last.leaf) {
        #[cfg(feature = "doc")]
        (Leaf::Doc(first), Leaf::Doc(last)) => Ok(first.before_bytes() != last.after_bytes()),
        #[cfg(feature = "xlsb")]
        (Leaf::Xlsb(first), Leaf::Xlsb(last)) => Ok(first.before() != last.after()),
        #[cfg(feature = "xls")]
        (Leaf::Xls(first), Leaf::Xls(last)) => Ok(first.before() != last.after()),
        #[cfg(feature = "pptx")]
        (Leaf::Pptx { source, .. }, Leaf::Pptx { target, .. }) => Ok(!source.same_content(target)?),
        #[cfg(feature = "docx")]
        (Leaf::Docx(first), Leaf::Docx(last)) => {
            Ok(first.source().xml_bytes() != last.target().xml_bytes())
        },
        #[cfg(feature = "xlsx")]
        (Leaf::Xlsx { source, .. }, Leaf::Xlsx { target, .. }) => Ok(!source.same_content(target)?),
        #[allow(unreachable_patterns)]
        _ => Err(Error::PatchConflict.into()),
    }
}

fn same_snapshot(left: &File, right: &File) -> bool {
    match (left, right) {
        #[cfg(feature = "doc")]
        (File::Doc(left), File::Doc(right)) => left.same_snapshot(right),
        #[cfg(feature = "xlsb")]
        (File::Xlsb(left), File::Xlsb(right)) => left.same_snapshot(right),
        #[cfg(feature = "xls")]
        (File::Xls(left), File::Xls(right)) => left.same_snapshot(right),
        #[cfg(feature = "pptx")]
        (File::Pptx(left), File::Pptx(right)) => left.same_snapshot(right),
        #[cfg(feature = "docx")]
        (File::Docx(left), File::Docx(right)) => left.same_snapshot(right),
        #[cfg(feature = "xlsx")]
        (File::Xlsx(left), File::Xlsx(right)) => left.same_snapshot(right),
        _ => false,
    }
}
fn admit(file: &File) -> Result<()> {
    match file {
        #[cfg(feature = "doc")]
        File::Doc(handle) => {
            if handle.signature() == crate::format::signature::Signature::Signed {
                return Err(crate::Error::signed(litchi_core::FileFormat::Doc));
            }
            match handle.editable_package()? {
                Some(package) => {
                    drop(package.edit()?);
                    Ok(())
                },
                None => Err(crate::Error::new(
                    Code::UnsupportedCapability,
                    "explicit transactions for DOC require retained source bytes",
                )
                .with_format(litchi_core::FileFormat::Doc)),
            }
        },
        #[cfg(feature = "pptx")]
        File::Pptx(handle) => {
            drop(handle.editable()?);
            Ok(())
        },
        #[cfg(feature = "docx")]
        File::Docx(handle) => {
            drop(handle.edit()?);
            Ok(())
        },
        #[cfg(feature = "xlsx")]
        File::Xlsx(handle) => handle.admit_edit(),
        #[cfg(feature = "xlsb")]
        File::Xlsb(handle) => handle.admit_edit(),
        #[cfg(feature = "xls")]
        File::Xls(handle) => handle.admit_edit(),
        _ => check_source(file),
    }
}
fn check_source(file: &File) -> Result<()> {
    match file {
        #[cfg(feature = "doc")]
        File::Doc(handle) => {
            handle.check_source()?;
            if handle.signature() == crate::format::signature::Signature::Signed {
                return Err(crate::Error::signed(litchi_core::FileFormat::Doc));
            }
            Ok(())
        },
        #[cfg(feature = "pptx")]
        File::Pptx(handle) => handle.check_source(),
        #[cfg(feature = "docx")]
        File::Docx(handle) => {
            handle.owner().check_source()?;
            if handle.signature() == crate::format::signature::Signature::Signed {
                return Err(crate::Error::signed(litchi_core::FileFormat::Docx));
            }
            Ok(())
        },
        #[cfg(feature = "xlsx")]
        File::Xlsx(handle) => handle.check_source(),
        #[cfg(feature = "xlsb")]
        File::Xlsb(handle) => handle.check_source(),
        #[cfg(feature = "xls")]
        File::Xls(handle) => handle.check_source(),
        File::Signed(signed) => Err(crate::Error::signed(signed.format())),
        File::Locked(locked) => Err(locked.read_error()),
        #[allow(unreachable_patterns)]
        _ => Err(pending()),
    }
}
fn pending() -> crate::Error {
    crate::Error::new(
        Code::UnsupportedCapability,
        "explicit transactions for this format are not integrated",
    )
}
