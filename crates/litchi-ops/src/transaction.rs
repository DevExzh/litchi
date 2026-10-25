//! Explicit immutable transactions and process-local reversible patches.
//!
//! DOCX main-document operations resolve against preceding staged results.
//! Atomic batches and closure edits group consecutive direct-path formatting
//! sets in one native transaction. Other commands and staged reads close a run.

use std::fmt;
use std::io::{self, Write};
use std::sync::Arc;

use sha2::{Digest as _, Sha256};

use crate::path::{IntoPath, Path};
use crate::{Code, File, IntoNode, IntoPlacement, IntoProps, IntoTarget, Result};

mod import;
mod limits;
mod merge;
mod options;
mod refresh;
pub use limits::{Error, LimitKind, SemanticLimits, SemanticLimitsBuilder};
pub use options::EditOptions;

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
/// The accessors follow the commit triad of every format crate
/// (`litchi_core::edit::Commit`): [`Self::file`] is the facade's spelling of
/// `package`. [`Self::into_parts`] takes every result without cloning.
/// `Debug` omits the closure value and all authored document content.
pub struct Commit<T = ()> {
    file: File,
    patch: Patch,
    diagnostics: Diagnostics,
    value: T,
}
impl<T> Commit<T> {
    pub(crate) const fn new(file: File, patch: Patch, diagnostics: Diagnostics, value: T) -> Self {
        Self {
            file,
            patch,
            diagnostics,
            value,
        }
    }
    /// The new immutable file. The original remains unchanged.
    #[must_use]
    pub const fn file(&self) -> &File {
        &self.file
    }
    /// Take the new file, dropping the patch and the value.
    #[must_use]
    pub fn into_file(self) -> File {
        self.file
    }
    /// The process-local reversible patch. The exact owner/clone path uses
    /// retained identity; a same-variant reopened copy is admitted only after
    /// its published base passes the bounded content proof.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }
    /// Content-free transaction facts.
    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }
    /// The value the closure (or the single command) returned.
    #[must_use]
    pub const fn value(&self) -> &T {
        &self.value
    }
    /// Take the value, dropping the file and the patch.
    #[must_use]
    pub fn into_value(self) -> T {
        self.value
    }
    /// Take the file, the patch and the value; the diagnostics are `Copy`
    /// and stay readable through [`Self::diagnostics`] before the call.
    #[must_use]
    pub fn into_parts(self) -> (File, Patch, T) {
        (self.file, self.patch, self.value)
    }
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
/// Applying a patch accepts its original input handle, a cheap clone, or a
/// separately opened file of the same variant whose published bytes equal the
/// patch's source. The exact-owner path is constant work; a reopened candidate
/// is checked with a bounded-memory length-and-digest proof (record 0362m).
/// A file whose bytes differ is refused in every format. Native owners, source
/// guards, origin and policy sidecars are retained; applying reuses the
/// already-validated endpoint without a ZIP round trip. Source freshness is
/// checked in both directions.
///
/// Signature consumption precedes changed editing. An empty or exact no-op
/// transaction may retain a signed publication, and inverse replay restores
/// the authorized unsigned logical source without recreating a signed state.
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
    Doc(Box<crate::format::doc::StepPatch>),
    #[cfg(feature = "xlsb")]
    Xlsb(Box<crate::format::xlsb::write::Patch>),
    #[cfg(feature = "xls")]
    Xls(Box<crate::format::xls::write::StepPatch>),
    #[cfg(feature = "ppt")]
    Ppt(Box<crate::format::ppt::write::StepPatch>),
    // The two-endpoint variants are boxed whole, so `Leaf` stays one
    // pointer wide in every feature combination (record 0304).
    #[cfg(feature = "pptx")]
    Pptx(Box<PptxLeaf>),
    #[cfg(feature = "docx")]
    Docx(Box<litchi_docx::document::Patch>),
    #[cfg(feature = "rtf")]
    Rtf(Box<litchi_rtf::edit::Patch>),
    #[cfg(feature = "odt")]
    Odt(Box<crate::format::odt::write::Patch>),
    #[cfg(feature = "ods")]
    Ods(Box<crate::format::ods::write::Patch>),
    #[cfg(feature = "odp")]
    Odp(Box<crate::format::odp::write::Patch>),
    #[cfg(odf_objects)]
    Object(Box<crate::format::odf_tree::write::Patch>),
    #[cfg(feature = "iwork-adapter")]
    Iwork(Box<crate::format::iwork::write::Patch>),
    #[cfg(feature = "xlsx")]
    Xlsx(Box<XlsxLeaf>),
}

// Every variant is one box: a discriminant and a pointer.
const _: () = assert!(size_of::<Leaf>() <= 16);
#[cfg(feature = "pptx")]
struct PptxLeaf {
    source: crate::format::pptx::Editable,
    target: crate::format::pptx::Editable,
    _proof: crate::format::pptx::write::Proof,
}
#[cfg(feature = "xlsx")]
struct XlsxLeaf {
    source: litchi_xlsx::Package,
    target: litchi_xlsx::Package,
    // Retain the complete native proof even though process-local replay
    // reuses a validated endpoint after a content proof for reopened copies.
    _patch: litchi_xlsx::workbook::Patch,
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
    /// Apply to the bound input, a cheap clone, or a separately opened
    /// byte-identical copy of the same file variant, returning the retained
    /// validated native endpoint. The input is unchanged on success or
    /// failure. A reopened candidate is checked by streaming its publication
    /// through a bounded-memory content proof.
    ///
    /// # Errors
    /// Returns a typed [`Error::PatchConflict`] for changed or mismatched
    /// content, protected state errors for changed signed/locked input, or
    /// source freshness failures. An exact no-op may replay onto a signed
    /// recipient after its published bytes match the recorded base.
    pub fn apply(&self, file: &File) -> Result<File> {
        let exact_noop = self.is_noop();
        let result = (|| {
            if exact_noop {
                check_source_for_replay(file)?;
            } else {
                check_source(file)?;
            }
            self.check_sources()?;
            let (source, target) = if self.inverse {
                (&self.inner.target, &self.inner.source)
            } else {
                (&self.inner.source, &self.inner.target)
            };
            let matches_recorded_base = if exact_noop && (file.is_signed() || source.is_signed()) {
                // A signed recipient must always go through the byte proof,
                // even when its leaf owner happens to be the same Arc as the
                // patch source. The caller's published bytes are the replay
                // admission fact; owner identity alone is insufficient.
                same_content_for_replay(file, source, true)?
            } else {
                same_snapshot(file, source) || same_content_for_replay(file, source, exact_noop)?
            };
            if !matches_recorded_base {
                return Err(Error::PatchConflict.into());
            }
            let target = if exact_noop && file.is_signed() {
                // Retain the recipient wrapper and signature bytes. The
                // fingerprint above proves it is the patch base; returning a
                // different retained endpoint could silently discard them.
                file.clone()
            } else {
                target.clone()
            };
            self.check_sources()?;
            Ok(target)
        })();
        if exact_noop {
            check_source_for_replay(file)?;
        } else {
            check_source(file)?;
        }
        result
    }
    fn check_sources(&self) -> Result<()> {
        let check = if self.is_noop() {
            check_source_for_replay
        } else {
            check_source
        };
        check(&self.inner.source)?;
        // Ordered retained endpoints are already validated by their leaf
        // commits. Check every source guard, including a source promoted to
        // an eager unsigned owner; the last endpoint alone is insufficient.
        if self.inverse {
            for step in self.inner.steps.iter().rev() {
                check(&step.after)?;
                check(&step.before)?;
            }
        } else {
            for step in &self.inner.steps {
                check(&step.before)?;
                check(&step.after)?;
            }
        }
        check(&self.inner.target)
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
    #[cfg(feature = "rtf")]
    rtf: Option<crate::format::rtf::write::coalesce::Run>,
    /// Direct RTF picture removals published as one native removal
    /// (record 0280).
    #[cfg(feature = "rtf")]
    rtf_pictures: Option<crate::format::rtf::write::removal::PictureRun>,
    #[cfg(feature = "odt")]
    odt: Option<crate::format::odt::write::coalesce::Run>,
    #[cfg(feature = "xlsx")]
    cells: Option<crate::format::xlsx::write::coalesce::Run>,
    #[cfg(feature = "xlsb")]
    xlsb_cells: Option<crate::format::xlsb::write::coalesce::Run>,
    #[cfg(feature = "xls")]
    xls_cells: Option<crate::format::xls::write::coalesce::Run>,
    #[cfg(feature = "xlsb")]
    xlsb_sheets: Option<crate::format::xlsb::write::sheet_coalesce::Run>,
    #[cfg(feature = "pptx")]
    shapes: Option<crate::format::pptx::write::coalesce::Run>,
    #[cfg(feature = "ppt")]
    ppt_shapes: Option<crate::format::ppt::write::coalesce::Run>,
    #[cfg(feature = "odp")]
    odp: Option<crate::format::odp::write::coalesce::Run>,
    #[cfg(feature = "ods")]
    ods_cells: Option<crate::format::ods::write::coalesce::Run>,
    /// Coalesced object-format sets (record 0288).
    #[cfg(odf_objects)]
    odf_tree: Option<crate::format::odf_tree::coalesce::Run>,
    /// Coalesced Numbers cell values and formulas (record 0297).
    #[cfg(feature = "numbers")]
    numbers_cells: Option<crate::format::iwork::numbers::coalesce::Run>,
    /// Coalesced Pages paragraph text sets (record 0299).
    #[cfg(feature = "pages")]
    pages: Option<crate::format::iwork::pages::Run>,
    /// Coalesced Keynote sets, new drawables and new slides (record 0298).
    #[cfg(feature = "keynote")]
    keynote: Option<crate::format::iwork::keynote::Run>,
    /// The match count of the last `find=` set (record 0273).
    matched: Option<usize>,
    /// Coalesced items whose change a native run reported at its commit,
    /// by item index (record 0348): their staging returned `false`.
    reported: Vec<(usize, bool)>,
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
    /// Batch inputs are an externally supplied stream.  A signed file is a
    /// protected write target, so reject it before asking the caller for even
    /// one item.  Ordinary `edit` retains its exact no-op exception because a
    /// closure can prove that it made no change; a batch cannot do that before
    /// consuming its stream.
    pub(crate) fn preflight_batch(&self) -> Result<()> {
        if !self.is_signed() {
            return Ok(());
        }
        check_source_for_replay(self)?;
        Err(crate::Error::signed(self.format().ok_or_else(pending)?))
    }

    /// Edit atomically through a closure, retaining a reversible patch and
    /// diagnostics. Integrates DOCX main-document writes, supported XLSX/XLSB
    /// cell and worksheet writes, immutable XLS cell writes, plus PPTX shape text,
    /// slide/text-box creation, slide ordering, top-level shape removal and
    /// dependency-free slide removal. Legacy PPT supports native shape text/anchor
    /// and slide visibility setters.
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
    /// assert!(commit.diagnostics().changed());
    /// let restored = commit.patch().inverse().apply(commit.file())?;
    /// assert_eq!(restored.to_bytes()?, file.to_bytes()?);
    /// # }
    /// # Ok::<(), litchi_ops::Error>(())
    /// ```
    pub fn edit<T>(&self, f: impl FnOnce(&mut Edit) -> Result<T>) -> Result<Commit<T>> {
        self.edit_with_options(EditOptions::default(), f)
    }

    /// [`Self::edit`] under caller-selected [`EditOptions`]: operation,
    /// retained-snapshot and format-specific target-expansion and staging
    /// limits.
    ///
    /// ```
    /// # #[cfg(feature = "docx")]
    /// # {
    /// use litchi_core::FileFormat;
    /// use litchi_ops::transaction::SemanticLimits;
    /// use litchi_ops::{Code, EditOptions, NewNode, create};
    ///
    /// let file = create(FileFormat::Docx)?;
    /// let options = EditOptions::new()
    ///     .semantic_limits(SemanticLimits::builder().max_operations(1)?.build()?);
    /// let error = file
    ///     .edit_with_options(options, |tx| {
    ///         tx.add("/body", NewNode::new("paragraph").prop("text", "one"))?;
    ///         tx.add("/body", NewNode::new("paragraph").prop("text", "two"))?;
    ///         Ok(())
    ///     })
    ///     .unwrap_err();
    /// assert_eq!(error.code(), Code::LimitExceeded);
    /// # }
    /// # Ok::<(), litchi_ops::Error>(())
    /// ```
    ///
    /// # Errors
    /// As [`Self::edit`]. A failed limit check discards the complete candidate.
    pub fn edit_with_options<T>(
        &self,
        options: EditOptions,
        f: impl FnOnce(&mut Edit) -> Result<T>,
    ) -> Result<Commit<T>> {
        let limits = options.semantic_limits_value();
        let result = (|| {
            admit(self)?;
            let identity = Arc::new(());
            let mut edit = Edit {
                identity: Arc::clone(&identity),
                // Coalescers operate on the format leaf.  A signed facade
                // keeps the same signed disposition on that leaf, while the
                // outer wrapper remains the immutable transaction endpoint.
                current: transaction_owner(self),
                steps: Vec::new(),
                limits,
                failure: None,
                snapshot_bytes: 0,
                operations: 0,
                #[cfg(feature = "doc")]
                doc: None,
                #[cfg(feature = "docx")]
                formatting: None,
                #[cfg(feature = "rtf")]
                rtf: None,
                #[cfg(feature = "rtf")]
                rtf_pictures: None,
                #[cfg(feature = "odt")]
                odt: None,
                #[cfg(feature = "xlsx")]
                cells: None,
                #[cfg(feature = "xlsb")]
                xlsb_cells: None,
                #[cfg(feature = "xls")]
                xls_cells: None,
                #[cfg(feature = "xlsb")]
                xlsb_sheets: None,
                #[cfg(feature = "pptx")]
                shapes: None,
                #[cfg(feature = "ppt")]
                ppt_shapes: None,
                #[cfg(feature = "odp")]
                odp: None,
                #[cfg(feature = "ods")]
                ods_cells: None,
                #[cfg(odf_objects)]
                odf_tree: None,
                #[cfg(feature = "numbers")]
                numbers_cells: None,
                #[cfg(feature = "pages")]
                pages: None,
                #[cfg(feature = "keynote")]
                keynote: None,
                matched: None,
                reported: Vec::new(),
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
            if changed && self.is_signed() {
                return Err(crate::Error::signed(self.format().ok_or_else(pending)?));
            }
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
            Ok(Commit::new(file, patch, diagnostics, value))
        })();
        if self.is_signed() {
            check_source_for_replay(self)?;
        } else {
            check_source(self)?;
        }
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
    /// Distinct direct XLS scalar sets, clears and removals share one native
    /// package edit. Repeated coordinates close that run and resolve again.
    /// Consecutive direct XLSB sheet visibility assignments to distinct tabs
    /// with one exact state also share one native structural edit. Hiding the
    /// active tab closes its run before the next command so intermediate view
    /// validation retains sequential semantics.
    /// Distinct direct PPTX shape text/geometry sets also share a native edit.
    /// Repeated PPTX shapes close their run before resolving partial properties.
    /// Distinct direct legacy PPT shape text/geometry sets share a native edit.
    /// Repeated targets and native operation limits close the run before replanning.
    /// Reads and selectors close pending runs; repeated XLSX cells and font
    /// assignments close the payload run.
    /// A successful call stages its assignments; native commit errors can surface
    /// at the next barrier or when the outer closure finishes, with the original
    /// mutation index. Any mutation or flush failure poisons the transaction.
    ///
    /// A `find` property searches the target's subtree, as [`File::set`]
    /// describes, in one native commit that never joins a coalesced run.
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
                // A deferred native coalescer failure belongs to the earlier
                // wire item and must surface before this admission error.
                self.flush_pending()?;
                return Err(Error::Limit {
                    kind: LimitKind::Operations,
                    observed: index.checked_add(1).unwrap_or(usize::MAX),
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
            if crate::find::requested(&props) {
                self.matched = Some(self.find_replace(&target, &props)?);
                return Ok(());
            }
            if self
                .coalesce_parsed(&target, &props, index, None)?
                .is_some()
            {
                return Ok(());
            }
            // The direct path below is normally flushed by `execute`, but the
            // PPT planner needs the post-flush snapshot remainder before it
            // captures its native before image.
            self.flush_pending()?;
            #[cfg(feature = "ppt")]
            let remaining_snapshot = self
                .limits
                .max_snapshot_bytes()
                .saturating_sub(self.snapshot_bytes);
            #[cfg(any(
                feature = "doc",
                feature = "odp",
                feature = "docx",
                feature = "xlsx",
                feature = "xlsb",
                feature = "xls",
                feature = "ppt",
                feature = "pptx",
                feature = "rtf",
                feature = "ods",
                feature = "odt",
                odf_objects,
                feature = "iwork-adapter"
            ))]
            let operation_limits = self.limits.with_snapshot_budget(
                self.limits
                    .max_snapshot_bytes()
                    .saturating_sub(self.snapshot_bytes),
            );
            self.execute(|file| match file {
                #[cfg(feature = "doc")]
                File::Doc(handle) => wrap_doc(
                    handle,
                    handle.set_with_limits(target, props, &operation_limits)?,
                    (),
                ),
                #[cfg(feature = "odp")]
                File::Odp(handle) => wrap_odp(
                    handle,
                    handle.set_with_limits(target, props, &operation_limits)?,
                    (),
                ),
                #[cfg(feature = "docx")]
                File::Docx(handle) => {
                    let commit = if _command_text {
                        handle.set_command_with_limits(&target, props, &operation_limits)?
                    } else {
                        handle.set_with_limits(&target, props, &operation_limits)?
                    };
                    wrap(handle, commit, ())
                },
                #[cfg(feature = "xlsx")]
                File::Xlsx(handle) => wrap_xlsx(
                    handle,
                    handle.set_with_limits(target, props, &operation_limits)?,
                    (),
                ),
                #[cfg(feature = "xlsb")]
                File::Xlsb(handle) => wrap_xlsb(
                    handle,
                    handle.set_with_limits(target, props, &operation_limits)?,
                    (),
                ),
                #[cfg(feature = "xls")]
                File::Xls(handle) => wrap_xls(
                    handle,
                    handle.set_with_limits(target, props, &operation_limits)?,
                    (),
                ),
                #[cfg(feature = "ppt")]
                File::Ppt(handle) => wrap_ppt(
                    handle,
                    handle.set_with_snapshot_budget(
                        target,
                        props,
                        &operation_limits,
                        remaining_snapshot,
                    )?,
                    (),
                ),
                #[cfg(feature = "pptx")]
                File::Pptx(handle) => wrap_pptx(
                    handle,
                    handle.set_with_limits(target, props, &operation_limits)?,
                    (),
                ),
                #[cfg(feature = "rtf")]
                File::Rtf(handle) => wrap_rtf(
                    handle,
                    handle.set_with_limits(target, props, &operation_limits)?,
                    (),
                ),
                #[cfg(feature = "ods")]
                File::Ods(handle) => wrap_ods(
                    handle,
                    handle.set_with_limits(target, props, &operation_limits)?,
                    (),
                ),
                #[cfg(feature = "odt")]
                File::Odt(handle) | File::Oth(handle) | File::Odm(handle) => wrap_odt(
                    handle,
                    handle.set_with_limits(target, props, &operation_limits)?,
                    (),
                ),

                #[cfg(odf_objects)]
                File::Odg(handle)
                | File::Odc(handle)
                | File::Odi(handle)
                | File::Odb(handle)
                | File::Odf(handle) => wrap_object(
                    handle,
                    handle.set_with_limits(target, props, &operation_limits)?,
                    (),
                ),
                #[cfg(feature = "iwork-adapter")]
                File::Pages(handle) | File::Keynote(handle) | File::Numbers(handle) => wrap_iwork(
                    handle,
                    handle.set_with_limits(target, props, &operation_limits)?,
                    (),
                ),
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
        #[cfg(feature = "keynote")]
        if matches!(self.current, File::Keynote(_)) {
            return self.add_keynote(_parent, _node);
        }
        // The leaf must receive the snapshot remainder after any coalesced
        // run has been published.  Passing the transaction-wide maximum here
        // lets a native add allocate its complete before/after capture and
        // only then fail in `execute`, which is too late for bounded-memory
        // admission.
        self.check()?;
        self.flush_pending()?;
        #[cfg(any(
            feature = "doc",
            feature = "odp",
            feature = "docx",
            feature = "xlsx",
            feature = "xlsb",
            feature = "xls",
            feature = "ppt",
            feature = "pptx",
            feature = "rtf",
            feature = "ods",
            feature = "odt",
            odf_objects,
            feature = "iwork-adapter"
        ))]
        let operation_limits = self.limits.with_snapshot_budget(
            self.limits
                .max_snapshot_bytes()
                .saturating_sub(self.snapshot_bytes),
        );
        self.execute(|file| match file {
            #[cfg(feature = "docx")]
            File::Docx(handle) => {
                let (commit, path) = handle.add_with_limits(
                    _parent.into_path()?,
                    _node.into_node()?,
                    &operation_limits,
                )?;
                wrap(handle, commit, path)
            },
            #[cfg(feature = "xls")]
            File::Xls(handle) => {
                let (commit, path) = handle.add_with_limits(
                    _parent.into_path()?,
                    _node.into_node()?,
                    &operation_limits,
                )?;
                wrap_xls(handle, commit, path)
            },
            #[cfg(feature = "xlsx")]
            File::Xlsx(handle) => {
                let (commit, path) = handle.add_with_limits(
                    _parent.into_path()?,
                    _node.into_node()?,
                    &operation_limits,
                )?;
                wrap_xlsx(handle, commit, path)
            },
            #[cfg(feature = "pptx")]
            File::Pptx(handle) => {
                let (commit, path) = handle.add_with_limits(
                    _parent.into_path()?,
                    _node.into_node()?,
                    &operation_limits,
                )?;
                wrap_pptx(handle, commit, path)
            },
            #[cfg(feature = "xlsb")]
            File::Xlsb(handle) => {
                let (commit, path) = handle.add_with_limits(
                    _parent.into_path()?,
                    _node.into_node()?,
                    &operation_limits,
                )?;
                wrap_xlsb(handle, commit, path)
            },
            #[cfg(feature = "rtf")]
            File::Rtf(handle) => {
                let (commit, path) = handle.add_with_limits(
                    _parent.into_path()?,
                    _node.into_node()?,
                    &operation_limits,
                )?;
                wrap_rtf(handle, commit, path)
            },
            #[cfg(feature = "ods")]
            File::Ods(handle) => {
                let (commit, path) = handle.add_with_limits(
                    _parent.into_path()?,
                    _node.into_node()?,
                    &operation_limits,
                )?;
                wrap_ods(handle, commit, path)
            },
            #[cfg(feature = "odt")]
            File::Odt(handle) | File::Oth(handle) | File::Odm(handle) => {
                let (commit, path) = handle.add_with_limits(
                    _parent.into_path()?,
                    _node.into_node()?,
                    &operation_limits,
                )?;
                wrap_odt(handle, commit, path)
            },

            #[cfg(odf_objects)]
            File::Odg(handle)
            | File::Odc(handle)
            | File::Odi(handle)
            | File::Odb(handle)
            | File::Odf(handle) => {
                let (commit, path) = handle.add_with_limits(
                    _parent.into_path()?,
                    _node.into_node()?,
                    &operation_limits,
                )?;
                wrap_object(handle, commit, path)
            },
            #[cfg(feature = "iwork-adapter")]
            File::Pages(handle) | File::Keynote(handle) | File::Numbers(handle) => {
                let (commit, path) = handle.add_with_limits(
                    _parent.into_path()?,
                    _node.into_node()?,
                    &operation_limits,
                )?;
                wrap_iwork(handle, commit, path)
            },
            #[cfg(feature = "doc")]
            File::Doc(handle) => {
                let (commit, path) = handle.add_with_limits(
                    _parent.into_path()?,
                    _node.into_node()?,
                    &operation_limits,
                )?;
                wrap_doc(handle, commit, path)
            },
            #[cfg(feature = "ppt")]
            File::Ppt(handle) => {
                let (commit, path) = handle.add_with_limits(
                    _parent.into_path()?,
                    _node.into_node()?,
                    &operation_limits,
                )?;
                wrap_ppt(handle, commit, path)
            },
            #[cfg(feature = "odp")]
            File::Odp(handle) => {
                let (commit, path) = handle.add_with_limits(
                    _parent.into_path()?,
                    _node.into_node()?,
                    &operation_limits,
                )?;
                wrap_odp(handle, commit, path)
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
        // Command batches share the same transaction budget as the public
        // route.  Flush a preceding coalesced step before deriving the
        // remainder so native add capture cannot temporarily exceed the
        // outer limit.
        self.check()?;
        self.flush_pending()?;
        #[cfg(any(
            feature = "doc",
            feature = "odp",
            feature = "docx",
            feature = "xlsx",
            feature = "xlsb",
            feature = "xls",
            feature = "ppt",
            feature = "pptx",
            feature = "rtf",
            feature = "ods",
            feature = "odt",
            odf_objects,
            feature = "iwork-adapter"
        ))]
        let operation_limits = self.limits.with_snapshot_budget(
            self.limits
                .max_snapshot_bytes()
                .saturating_sub(self.snapshot_bytes),
        );
        self.execute(|file| match file {
            #[cfg(feature = "docx")]
            File::Docx(handle) => {
                let (commit, path) = handle.add_command_with_limits(
                    _parent.into_path()?,
                    _node.into_node()?,
                    &operation_limits,
                )?;
                wrap(handle, commit, path)
            },
            #[cfg(feature = "xls")]
            File::Xls(handle) => {
                let (commit, path) = handle.add_with_limits(
                    _parent.into_path()?,
                    _node.into_node()?,
                    &operation_limits,
                )?;
                wrap_xls(handle, commit, path)
            },
            #[cfg(feature = "xlsx")]
            File::Xlsx(handle) => {
                let (commit, path) = handle.add_with_limits(
                    _parent.into_path()?,
                    _node.into_node()?,
                    &operation_limits,
                )?;
                wrap_xlsx(handle, commit, path)
            },
            #[cfg(feature = "pptx")]
            File::Pptx(handle) => {
                let (commit, path) = handle.add_with_limits(
                    _parent.into_path()?,
                    _node.into_node()?,
                    &operation_limits,
                )?;
                wrap_pptx(handle, commit, path)
            },
            #[cfg(feature = "xlsb")]
            File::Xlsb(handle) => {
                let (commit, path) = handle.add_with_limits(
                    _parent.into_path()?,
                    _node.into_node()?,
                    &operation_limits,
                )?;
                wrap_xlsb(handle, commit, path)
            },
            #[cfg(feature = "rtf")]
            File::Rtf(handle) => {
                let (commit, path) = handle.add_with_limits(
                    _parent.into_path()?,
                    _node.into_node()?,
                    &operation_limits,
                )?;
                wrap_rtf(handle, commit, path)
            },
            #[cfg(feature = "ods")]
            File::Ods(handle) => {
                let (commit, path) = handle.add_with_limits(
                    _parent.into_path()?,
                    _node.into_node()?,
                    &operation_limits,
                )?;
                wrap_ods(handle, commit, path)
            },
            #[cfg(feature = "odt")]
            File::Odt(handle) | File::Oth(handle) | File::Odm(handle) => {
                let (commit, path) = handle.add_with_limits(
                    _parent.into_path()?,
                    _node.into_node()?,
                    &operation_limits,
                )?;
                wrap_odt(handle, commit, path)
            },

            #[cfg(odf_objects)]
            File::Odg(handle)
            | File::Odc(handle)
            | File::Odi(handle)
            | File::Odb(handle)
            | File::Odf(handle) => {
                let (commit, path) = handle.add_with_limits(
                    _parent.into_path()?,
                    _node.into_node()?,
                    &operation_limits,
                )?;
                wrap_object(handle, commit, path)
            },
            #[cfg(feature = "iwork-adapter")]
            File::Pages(handle) | File::Keynote(handle) | File::Numbers(handle) => {
                let (commit, path) = handle.add_with_limits(
                    _parent.into_path()?,
                    _node.into_node()?,
                    &operation_limits,
                )?;
                wrap_iwork(handle, commit, path)
            },
            #[cfg(feature = "doc")]
            File::Doc(handle) => {
                let (commit, path) = handle.add_with_limits(
                    _parent.into_path()?,
                    _node.into_node()?,
                    &operation_limits,
                )?;
                wrap_doc(handle, commit, path)
            },
            #[cfg(feature = "ppt")]
            File::Ppt(handle) => {
                let (commit, path) = handle.add_with_limits(
                    _parent.into_path()?,
                    _node.into_node()?,
                    &operation_limits,
                )?;
                wrap_ppt(handle, commit, path)
            },
            #[cfg(feature = "odp")]
            File::Odp(handle) => {
                let (commit, path) = handle.add_with_limits(
                    _parent.into_path()?,
                    _node.into_node()?,
                    &operation_limits,
                )?;
                wrap_odp(handle, commit, path)
            },
            #[allow(unreachable_patterns)]
            _ => Err(pending()),
        })
    }

    /// Remove selected elements from the current candidate.
    ///
    /// Distinct direct XLS removals share a native package edit with scalar
    /// setters. Repeated coordinates, ranges, queries and reads close the run.
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
                    observed: index.checked_add(1).unwrap_or(usize::MAX),
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
            // The ordinary path below is flushed again by `execute`; do it
            // here so the leaf sees the post-flush snapshot remainder while
            // it preflights its native before image.
            self.flush_pending()?;
            #[cfg(any(
                feature = "doc",
                feature = "odp",
                feature = "docx",
                feature = "xlsx",
                feature = "xlsb",
                feature = "xls",
                feature = "ppt",
                feature = "pptx",
                feature = "rtf",
                feature = "ods",
                feature = "odt",
                odf_objects,
                feature = "iwork-adapter"
            ))]
            let operation_limits = self.limits.with_snapshot_budget(
                self.limits
                    .max_snapshot_bytes()
                    .saturating_sub(self.snapshot_bytes),
            );
            self.execute(|file| match file {
                #[cfg(feature = "docx")]
                File::Docx(handle) => wrap(
                    handle,
                    handle.remove_with_limits(target.clone(), &operation_limits)?,
                    (),
                ),
                #[cfg(feature = "xlsx")]
                File::Xlsx(handle) => wrap_xlsx(
                    handle,
                    handle.remove_with_limits(target.clone(), &operation_limits)?,
                    (),
                ),
                #[cfg(feature = "xlsb")]
                File::Xlsb(handle) => wrap_xlsb(
                    handle,
                    handle.remove_with_limits(target.clone(), &operation_limits)?,
                    (),
                ),
                #[cfg(feature = "xls")]
                File::Xls(handle) => wrap_xls(
                    handle,
                    handle.remove_with_limits(target.clone(), &operation_limits)?,
                    (),
                ),
                #[cfg(feature = "ppt")]
                File::Ppt(handle) => wrap_ppt(
                    handle,
                    handle.remove_with_limits(target.clone(), &operation_limits)?,
                    (),
                ),
                #[cfg(feature = "pptx")]
                File::Pptx(handle) => wrap_pptx(
                    handle,
                    handle.remove_with_limits(target.clone(), &operation_limits)?,
                    (),
                ),
                #[cfg(feature = "rtf")]
                File::Rtf(handle) => wrap_rtf(
                    handle,
                    handle.remove_with_limits(target.clone(), &operation_limits)?,
                    (),
                ),
                #[cfg(feature = "ods")]
                File::Ods(handle) => wrap_ods(
                    handle,
                    handle.remove_with_limits(target.clone(), &operation_limits)?,
                    (),
                ),
                #[cfg(feature = "odt")]
                File::Odt(handle) | File::Oth(handle) | File::Odm(handle) => wrap_odt(
                    handle,
                    handle.remove_with_limits(target.clone(), &operation_limits)?,
                    (),
                ),

                #[cfg(odf_objects)]
                File::Odg(handle)
                | File::Odc(handle)
                | File::Odi(handle)
                | File::Odb(handle)
                | File::Odf(handle) => wrap_object(
                    handle,
                    handle.remove_with_limits(target.clone(), &operation_limits)?,
                    (),
                ),
                #[cfg(feature = "iwork-adapter")]
                File::Pages(handle) | File::Keynote(handle) | File::Numbers(handle) => wrap_iwork(
                    handle,
                    handle.remove_with_limits(target.clone(), &operation_limits)?,
                    (),
                ),
                #[cfg(feature = "doc")]
                File::Doc(handle) => wrap_doc(
                    handle,
                    handle.remove_with_limits(target.clone(), &operation_limits)?,
                    (),
                ),
                #[cfg(feature = "odp")]
                File::Odp(handle) => wrap_odp(
                    handle,
                    handle.remove_with_limits(target.clone(), &operation_limits)?,
                    (),
                ),
                #[allow(unreachable_patterns)]
                _ => Err(pending()),
            })
        })();
        result.map_err(|error| self.poison(error, index))
    }

    /// Try to stage a direct XLS or XLSB cell removal in the pending native cell
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
        #[cfg(feature = "xls")]
        if let File::Xls(_) = &self.current {
            let planned = match &self.xls_cells {
                Some(run) => run.remove_plan(target, &self.limits),
                None => {
                    let File::Xls(handle) = &self.current else {
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
            let Some(mut plan) = planned? else {
                return Ok(None);
            };
            if self.operations >= self.limits.max_operations() {
                return Err(Error::Limit {
                    kind: LimitKind::Operations,
                    observed: self.operations.checked_add(1).unwrap_or(usize::MAX),
                    maximum: self.limits.max_operations(),
                }
                .into());
            }
            if self
                .xls_cells
                .as_ref()
                .is_some_and(|run| run.requires_barrier(&plan, &self.limits))
            {
                self.flush_pending()?;
                let File::Xls(handle) = &self.current else {
                    return Err(pending());
                };
                let Some(replanned) = handle.remove_plan(target, &self.limits)? else {
                    return Ok(None);
                };
                plan = replanned;
            }
            if self.xls_cells.is_none() {
                let File::Xls(handle) = &self.current else {
                    return Err(pending());
                };
                self.xls_cells = Some(crate::format::xls::write::coalesce::Run::new(
                    handle,
                    &self.limits,
                    self.limits
                        .max_snapshot_bytes()
                        .saturating_sub(self.snapshot_bytes),
                )?);
            }
            let (had_prior, staged) = {
                let run = self.xls_cells.as_mut().ok_or_else(|| {
                    crate::Error::new(Code::InternalError, "XLS cell run missing")
                })?;
                (run.first_index().is_some(), run.stage(plan, index))
            };
            let changed = match staged {
                Ok(changed) => changed,
                Err(error) => {
                    // Native staging is failure-atomic. Publish the earlier
                    // prefix first so its deferred error keeps wire precedence.
                    if had_prior {
                        self.flush_pending()?;
                        if xls_staging_full(&error) {
                            return self.coalesce_remove_parsed(target, index);
                        }
                    }
                    return Err(error);
                },
            };
            self.operations += 1;
            return Ok(Some(changed));
        }
        #[cfg(feature = "rtf")]
        if let File::Rtf(handle) = &self.current {
            return self.coalesce_rtf_picture_removal(&handle.clone(), target, index);
        }
        let _ = (target, index);
        Ok(None)
    }

    /// Stages a direct RTF picture removal (`/body/p[N]/picture[M]`) in the
    /// pending picture run, so consecutive removals publish one native
    /// removal (record 0280). Any other target closes the run and takes the
    /// ordinary path.
    #[cfg(feature = "rtf")]
    fn coalesce_rtf_picture_removal(
        &mut self,
        handle: &crate::format::rtf::Rtf,
        target: &crate::TargetArg<'_>,
        index: usize,
    ) -> Result<Option<bool>> {
        use crate::format::rtf::write::removal::PictureRun;
        if self.rtf.is_some() {
            self.flush_pending()?;
        }
        if self.rtf_pictures.as_ref().is_some_and(PictureRun::is_full) {
            self.flush_pending()?;
        }
        let planned = match &self.rtf_pictures {
            Some(run) => run.plan(target),
            None => PictureRun::new(handle).plan(target),
        };
        let Some(picture) = planned else {
            // Not a remaining direct picture: publish the staged removals so
            // the ordinary path resolves against the current candidate.
            self.flush_pending()?;
            return Ok(None);
        };
        let File::Rtf(current) = &self.current else {
            return Err(pending());
        };
        let run = self
            .rtf_pictures
            .get_or_insert_with(|| PictureRun::new(current));
        run.stage(picture, index);
        self.operations += 1;
        Ok(Some(true))
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
        self.check()?;
        self.flush_pending()?;
        #[cfg(any(
            feature = "doc",
            feature = "odp",
            feature = "docx",
            feature = "xlsx",
            feature = "xlsb",
            feature = "xls",
            feature = "ppt",
            feature = "pptx",
            feature = "rtf",
            feature = "ods",
            feature = "odt",
            odf_objects,
            feature = "iwork-adapter"
        ))]
        let remaining_snapshot = self
            .limits
            .max_snapshot_bytes()
            .saturating_sub(self.snapshot_bytes);
        #[cfg(any(
            feature = "doc",
            feature = "odp",
            feature = "docx",
            feature = "xlsx",
            feature = "xlsb",
            feature = "xls",
            feature = "ppt",
            feature = "pptx",
            feature = "rtf",
            feature = "ods",
            feature = "odt",
            odf_objects,
            feature = "iwork-adapter"
        ))]
        let operation_limits = self.limits.with_snapshot_budget(remaining_snapshot);
        self.execute(|file| match file {
            #[cfg(feature = "docx")]
            File::Docx(handle) => wrap(
                handle,
                handle.move_to_with_limits(
                    _path.into_path()?,
                    _to.into_placement()?,
                    &operation_limits,
                )?,
                (),
            ),
            #[cfg(feature = "xls")]
            File::Xls(handle) => wrap_xls(
                handle,
                handle.move_to_with_limits(
                    _path.into_path()?,
                    _to.into_placement()?,
                    &operation_limits,
                )?,
                (),
            ),
            #[cfg(feature = "xlsx")]
            File::Xlsx(handle) => wrap_xlsx(
                handle,
                handle.move_to_with_limits(
                    _path.into_path()?,
                    _to.into_placement()?,
                    &operation_limits,
                )?,
                (),
            ),
            #[cfg(feature = "xlsb")]
            File::Xlsb(handle) => wrap_xlsb(
                handle,
                handle.move_to_with_limits(
                    _path.into_path()?,
                    _to.into_placement()?,
                    &operation_limits,
                )?,
                (),
            ),
            #[cfg(feature = "ppt")]
            File::Ppt(handle) => wrap_ppt(
                handle,
                handle.move_to_with_limits(
                    _path.into_path()?,
                    _to.into_placement()?,
                    &operation_limits,
                )?,
                (),
            ),
            #[cfg(feature = "pptx")]
            File::Pptx(handle) => wrap_pptx(
                handle,
                handle.move_to_with_limits(
                    _path.into_path()?,
                    _to.into_placement()?,
                    &operation_limits,
                )?,
                (),
            ),
            #[cfg(feature = "rtf")]
            File::Rtf(handle) => wrap_rtf(
                handle,
                handle.move_to_with_limits(
                    _path.into_path()?,
                    _to.into_placement()?,
                    &operation_limits,
                )?,
                (),
            ),
            #[cfg(feature = "ods")]
            File::Ods(handle) => wrap_ods(
                handle,
                handle.move_to_with_limits(
                    _path.into_path()?,
                    _to.into_placement()?,
                    &operation_limits,
                )?,
                (),
            ),
            #[cfg(feature = "odt")]
            File::Odt(handle) | File::Oth(handle) | File::Odm(handle) => wrap_odt(
                handle,
                handle.move_to_with_limits(
                    _path.into_path()?,
                    _to.into_placement()?,
                    &operation_limits,
                )?,
                (),
            ),

            #[cfg(odf_objects)]
            File::Odg(handle)
            | File::Odc(handle)
            | File::Odi(handle)
            | File::Odb(handle)
            | File::Odf(handle) => wrap_object(
                handle,
                handle.move_to_with_limits(
                    _path.into_path()?,
                    _to.into_placement()?,
                    &operation_limits,
                )?,
                (),
            ),
            #[cfg(feature = "iwork-adapter")]
            File::Pages(handle) | File::Keynote(handle) | File::Numbers(handle) => wrap_iwork(
                handle,
                handle.move_to_with_limits(
                    _path.into_path()?,
                    _to.into_placement()?,
                    &operation_limits,
                )?,
                (),
            ),
            #[cfg(feature = "doc")]
            File::Doc(handle) => wrap_doc(
                handle,
                handle.move_to_with_limits(
                    _path.into_path()?,
                    _to.into_placement()?,
                    &operation_limits,
                )?,
                (),
            ),
            #[cfg(feature = "odp")]
            File::Odp(handle) => wrap_odp(
                handle,
                handle.move_to_with_limits(
                    _path.into_path()?,
                    _to.into_placement()?,
                    &operation_limits,
                )?,
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
        self.check()?;
        self.flush_pending()?;
        #[cfg(any(
            feature = "doc",
            feature = "odp",
            feature = "docx",
            feature = "xlsx",
            feature = "xlsb",
            feature = "xls",
            feature = "ppt",
            feature = "pptx",
            feature = "rtf",
            feature = "ods",
            feature = "odt",
            odf_objects,
            feature = "iwork-adapter"
        ))]
        let remaining_snapshot = self
            .limits
            .max_snapshot_bytes()
            .saturating_sub(self.snapshot_bytes);
        #[cfg(any(
            feature = "doc",
            feature = "odp",
            feature = "docx",
            feature = "xlsx",
            feature = "xlsb",
            feature = "xls",
            feature = "ppt",
            feature = "pptx",
            feature = "rtf",
            feature = "ods",
            feature = "odt",
            odf_objects,
            feature = "iwork-adapter"
        ))]
        let operation_limits = self.limits.with_snapshot_budget(remaining_snapshot);
        self.execute(|file| match file {
            #[cfg(feature = "docx")]
            File::Docx(handle) => wrap(
                handle,
                handle.swap_with_limits(_a.into_path()?, _b.into_path()?, &operation_limits)?,
                (),
            ),
            #[cfg(feature = "xls")]
            File::Xls(handle) => wrap_xls(
                handle,
                handle.swap_with_limits(_a.into_path()?, _b.into_path()?, &operation_limits)?,
                (),
            ),
            #[cfg(feature = "xlsx")]
            File::Xlsx(handle) => wrap_xlsx(
                handle,
                handle.swap_with_limits(_a.into_path()?, _b.into_path()?, &operation_limits)?,
                (),
            ),
            #[cfg(feature = "xlsb")]
            File::Xlsb(handle) => wrap_xlsb(
                handle,
                handle.swap_with_limits(_a.into_path()?, _b.into_path()?, &operation_limits)?,
                (),
            ),
            #[cfg(feature = "ppt")]
            File::Ppt(handle) => wrap_ppt(
                handle,
                handle.swap_with_limits(_a.into_path()?, _b.into_path()?, &operation_limits)?,
                (),
            ),
            #[cfg(feature = "pptx")]
            File::Pptx(handle) => wrap_pptx(
                handle,
                handle.swap_with_limits(_a.into_path()?, _b.into_path()?, &operation_limits)?,
                (),
            ),
            #[cfg(feature = "rtf")]
            File::Rtf(handle) => wrap_rtf(
                handle,
                handle.swap_with_limits(_a.into_path()?, _b.into_path()?, &operation_limits)?,
                (),
            ),
            #[cfg(feature = "ods")]
            File::Ods(handle) => wrap_ods(
                handle,
                handle.swap_with_limits(_a.into_path()?, _b.into_path()?, &operation_limits)?,
                (),
            ),
            #[cfg(feature = "odt")]
            File::Odt(handle) | File::Oth(handle) | File::Odm(handle) => wrap_odt(
                handle,
                handle.swap_with_limits(_a.into_path()?, _b.into_path()?, &operation_limits)?,
                (),
            ),

            #[cfg(odf_objects)]
            File::Odg(handle)
            | File::Odc(handle)
            | File::Odi(handle)
            | File::Odb(handle)
            | File::Odf(handle) => wrap_object(
                handle,
                handle.swap_with_limits(_a.into_path()?, _b.into_path()?, &operation_limits)?,
                (),
            ),
            #[cfg(feature = "iwork-adapter")]
            File::Pages(handle) | File::Keynote(handle) | File::Numbers(handle) => wrap_iwork(
                handle,
                handle.swap_with_limits(_a.into_path()?, _b.into_path()?, &operation_limits)?,
                (),
            ),
            #[cfg(feature = "doc")]
            File::Doc(handle) => wrap_doc(
                handle,
                handle.swap_with_limits(_a.into_path()?, _b.into_path()?, &operation_limits)?,
                (),
            ),
            #[cfg(feature = "odp")]
            File::Odp(handle) => wrap_odp(
                handle,
                handle.swap_with_limits(_a.into_path()?, _b.into_path()?, &operation_limits)?,
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
        if crate::find::requested(_props) {
            // A find resolves its whole scope: it never joins a native run.
            self.flush_pending()?;
            return Ok(None);
        }
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
                    observed: self.operations.checked_add(1).unwrap_or(usize::MAX),
                    maximum: self.limits.max_operations(),
                }
                .into());
            }
            if self.doc.is_none() {
                self.doc = Some(crate::format::doc::write::Run::new(
                    &handle,
                    &self.limits,
                    self.limits
                        .max_snapshot_bytes()
                        .saturating_sub(self.snapshot_bytes),
                )?);
            }
            let run = self
                .doc
                .as_mut()
                .ok_or_else(|| crate::Error::new(Code::InternalError, "DOC body run missing"))?;
            let changed = run.stage(plan, _index, _spelling)?;
            let bytes = self.snapshot_bytes.saturating_add(run.snapshot_bytes()?);
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
            //
            // The exception is `text` alone on a direct body paragraph: its
            // ordinal is stable across text replacement, so it joins the
            // run (record 0355a). A run that has replaced text closes
            // before any target it cannot prove against its base.
            let text = _props
                .iter()
                .any(|(key, _)| key.eq_ignore_ascii_case("text"));
            let planned = match (&self.formatting, _spelling, text) {
                (Some(run), _, true) => run.text_plan(_target, _props),
                (None, _, true) => handle.text_plan(_target, _props),
                (Some(run), Some(spelling), false) => {
                    run.formatting_plan(spelling, _target, _props)
                },
                (_, _, false) => handle.formatting_plan(_target, _props),
            };
            if self.formatting.as_ref().is_some_and(|run| run.has_text())
                && planned.as_ref().map_or(true, |plan| {
                    plan.as_ref().is_none_or(|plan| {
                        !self.formatting.as_ref().is_some_and(|run| run.admits(plan))
                    })
                })
            {
                // Resolve against the published candidate instead.
                self.flush_pending()?;
                return Ok(None);
            }
            let Some(plan) = planned? else {
                return Ok(None);
            };
            if self.operations >= self.limits.max_operations() {
                return Err(Error::Limit {
                    kind: LimitKind::Operations,
                    observed: self.operations.checked_add(1).unwrap_or(usize::MAX),
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
                self.formatting = Some(
                    crate::format::docx::write::coalesce::Run::new_with_snapshot_budget(
                        handle,
                        self.limits
                            .max_snapshot_bytes()
                            .saturating_sub(self.snapshot_bytes),
                    )?,
                );
            }
            let run = self
                .formatting
                .as_mut()
                .ok_or_else(|| crate::Error::new(Code::InternalError, "formatting run missing"))?;
            let changed = run.stage(plan, _index, _spelling)?;
            let bytes = self.snapshot_bytes.saturating_add(run.snapshot_bytes()?);
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
        // A set closes pending picture removals (record 0280).
        #[cfg(feature = "rtf")]
        if self.rtf_pictures.is_some() {
            self.flush_pending()?;
        }
        #[cfg(feature = "rtf")]
        if let File::Rtf(handle) = self.current.clone() {
            if _props.is_empty() {
                // Preserve the direct adapter's exact no-op path without
                // constructing a whole-document coalescer index.
                if self.rtf.is_some() {
                    self.flush_pending()?;
                }
                return Ok(None);
            }
            let mut planned = if let Some(run) = self.rtf.as_ref() {
                run.plan(_target, _props)
            } else if crate::format::rtf::write::coalesce::direct_candidate(_target, _props)? {
                let Some(run) =
                    crate::format::rtf::write::coalesce::Run::new(&handle, &self.limits)?
                else {
                    return Ok(None);
                };
                let plan = run.plan(_target, _props)?;
                self.rtf = Some(run);
                Ok(plan)
            } else {
                Ok(None)
            };
            if let Err(error) = &planned {
                let error = error.clone();
                // Resolve admission errors only after publishing an earlier
                // deferred native failure, preserving wire attribution.
                self.flush_pending()?;
                // A structural edit can create a positional target. A full
                // coalesced staging window can also admit the next mutation
                // after publication; an individually oversized plan still
                // fails when retried against the fresh window.
                if !matches!(error.code(), Code::PathNotFound | Code::LimitExceeded) {
                    return Err(error);
                }
                let File::Rtf(handle) = self.current.clone() else {
                    return Err(pending());
                };
                planned = if crate::format::rtf::write::coalesce::direct_candidate(_target, _props)?
                {
                    let Some(run) =
                        crate::format::rtf::write::coalesce::Run::new(&handle, &self.limits)?
                    else {
                        return Ok(None);
                    };
                    let plan = run.plan(_target, _props)?;
                    self.rtf = Some(run);
                    Ok(plan)
                } else {
                    Ok(None)
                };
            }
            let Some(mut plan) = planned? else {
                return Ok(None);
            };
            if self.operations >= self.limits.max_operations() {
                return Err(Error::Limit {
                    kind: LimitKind::Operations,
                    observed: self.operations.checked_add(1).unwrap_or(usize::MAX),
                    maximum: self.limits.max_operations(),
                }
                .into());
            }
            if self
                .rtf
                .as_ref()
                .is_some_and(|run| run.requires_barrier(&plan))
            {
                self.flush_pending()?;
                let File::Rtf(handle) = self.current.clone() else {
                    return Err(pending());
                };
                if !crate::format::rtf::write::coalesce::direct_candidate(_target, _props)? {
                    return Ok(None);
                }
                let Some(run) =
                    crate::format::rtf::write::coalesce::Run::new(&handle, &self.limits)?
                else {
                    return Ok(None);
                };
                plan = run.plan(_target, _props)?.ok_or_else(|| {
                    crate::Error::new(Code::InternalError, "RTF set plan disappeared")
                })?;
                self.rtf = Some(run);
            }
            if self.rtf.is_none() {
                let File::Rtf(handle) = self.current.clone() else {
                    return Err(pending());
                };
                let Some(run) =
                    crate::format::rtf::write::coalesce::Run::new(&handle, &self.limits)?
                else {
                    return Ok(None);
                };
                plan = run.plan(_target, _props)?.ok_or_else(|| {
                    crate::Error::new(Code::InternalError, "RTF set plan disappeared")
                })?;
                self.rtf = Some(run);
            }
            let run = self
                .rtf
                .as_mut()
                .ok_or_else(|| crate::Error::new(Code::InternalError, "RTF run missing"))?;
            let changed = match run.stage(plan, _index) {
                Ok(changed) => changed,
                Err(error) => {
                    self.flush_pending()?;
                    return Err(error);
                },
            };
            self.operations += 1;
            return Ok(Some(changed));
        }
        #[cfg(feature = "pages")]
        if let File::Pages(handle) = &self.current {
            // Paragraph `text` sets keep every paragraph's position, so a
            // run of them shares one leaf paragraph transaction per section.
            let plan = match crate::format::iwork::pages::plan(handle, _target, _props) {
                Ok(Some(plan)) => plan,
                Ok(None) => {
                    if self.pages.is_some() {
                        self.flush_pending()?;
                    }
                    return Ok(None);
                },
                Err(error) => {
                    self.flush_pending()?;
                    return Err(error);
                },
            };
            if self.operations >= self.limits.max_operations() {
                return Err(Error::Limit {
                    kind: LimitKind::Operations,
                    observed: self.operations.checked_add(1).unwrap_or(usize::MAX),
                    maximum: self.limits.max_operations(),
                }
                .into());
            }
            let File::Pages(handle) = &self.current else {
                return Err(pending());
            };
            let run = self
                .pages
                .get_or_insert_with(|| crate::format::iwork::pages::Run::new(handle, &self.limits));
            let changed = run.stage(plan, _index)?;
            self.operations += 1;
            return Ok(Some(changed));
        }
        #[cfg(feature = "keynote")]
        if let File::Keynote(handle) = self.current.clone() {
            // Direct text, paragraph, frame and visibility sets join one
            // run: one slide-drawable commit plus one commit per text
            // (record 0298). A run holding new slides publishes first.
            let plan = match self.keynote.as_mut() {
                // Record 0314: the run memoizes the paths it resolved.
                Some(run) if run.accepts_sets() => run.plan_set(_target, _props),
                Some(_) => None,
                None => crate::format::iwork::keynote::coalesce::plan_set(&handle, _target, _props),
            };
            let Some(mut node) = plan else {
                if self.keynote.is_some() {
                    self.flush_pending()?;
                }
                return Ok(None);
            };
            if self.operations >= self.limits.max_operations() {
                return Err(Error::Limit {
                    kind: LimitKind::Operations,
                    observed: self.operations.checked_add(1).unwrap_or(usize::MAX),
                    maximum: self.limits.max_operations(),
                }
                .into());
            }
            let mut handle = handle;
            if self
                .keynote
                .as_mut()
                .is_some_and(|run| run.requires_barrier(node, _props))
            {
                // Record 0314: publish the run, then plan on its result.
                self.flush_pending()?;
                let File::Keynote(current) = self.current.clone() else {
                    return Err(pending());
                };
                let Some(replanned) =
                    crate::format::iwork::keynote::coalesce::plan_set(&current, _target, _props)
                else {
                    return Ok(None);
                };
                (node, handle) = (replanned, current);
            }
            let run = self.keynote.get_or_insert_with(|| {
                crate::format::iwork::keynote::Run::new(&handle, &self.limits)
            });
            run.stage_set(node, _props, _index)?;
            self.operations += 1;
            return Ok(Some(true));
        }
        #[cfg(feature = "odt")]
        if let Some(handle) = self.current.odt_family().cloned() {
            // A run plans every set against its base handle (the current
            // candidate until the run is flushed). A set the run cannot plan
            // runs on its own after the run is published.
            let plan = match crate::format::odt::write::plan(&handle, _target, _props) {
                Ok(Some(plan)) => plan,
                Ok(None) => return Ok(None),
                Err(_) => {
                    // Resolve and report against the published candidate.
                    self.flush_pending()?;
                    return Ok(None);
                },
            };
            if self.operations >= self.limits.max_operations() {
                return Err(Error::Limit {
                    kind: LimitKind::Operations,
                    observed: self.operations.checked_add(1).unwrap_or(usize::MAX),
                    maximum: self.limits.max_operations(),
                }
                .into());
            }
            let plan = if self
                .odt
                .as_ref()
                .is_some_and(|run| run.requires_barrier(&plan))
            {
                self.flush_pending()?;
                let Some(handle) = self.current.odt_family().cloned() else {
                    return Err(pending());
                };
                match crate::format::odt::write::plan(&handle, _target, _props) {
                    Ok(Some(plan)) => plan,
                    Ok(None) | Err(_) => return Ok(None),
                }
            } else {
                plan
            };
            if self.odt.is_none() {
                let Some(handle) = self.current.odt_family() else {
                    return Err(pending());
                };
                self.odt = Some(crate::format::odt::write::coalesce::Run::new(
                    handle,
                    &self.limits,
                )?);
            }
            let run = self
                .odt
                .as_mut()
                .ok_or_else(|| crate::Error::new(Code::InternalError, "ODT run missing"))?;
            let changed = match run.stage(plan, _index) {
                Ok(changed) => changed,
                Err(error) => {
                    self.flush_pending()?;
                    return Err(error);
                },
            };
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
                        observed: self.operations.checked_add(1).unwrap_or(usize::MAX),
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
                    observed: self.operations.checked_add(1).unwrap_or(usize::MAX),
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
        #[cfg(feature = "xls")]
        if let File::Xls(_) = &self.current {
            let planned = match &self.xls_cells {
                Some(run) => run.cell_plan(_target, _props, &self.limits),
                None => {
                    let File::Xls(handle) = &self.current else {
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
            let Some(mut plan) = planned? else {
                return Ok(None);
            };
            if self.operations >= self.limits.max_operations() {
                return Err(Error::Limit {
                    kind: LimitKind::Operations,
                    observed: self.operations.checked_add(1).unwrap_or(usize::MAX),
                    maximum: self.limits.max_operations(),
                }
                .into());
            }
            if self
                .xls_cells
                .as_ref()
                .is_some_and(|run| run.requires_barrier(&plan, &self.limits))
            {
                self.flush_pending()?;
                let File::Xls(handle) = &self.current else {
                    return Err(pending());
                };
                let Some(replanned) = handle.cell_plan(_target, _props, &self.limits)? else {
                    return Ok(None);
                };
                plan = replanned;
            }
            if self.xls_cells.is_none() {
                let File::Xls(handle) = &self.current else {
                    return Err(pending());
                };
                self.xls_cells = Some(crate::format::xls::write::coalesce::Run::new(
                    handle,
                    &self.limits,
                    self.limits
                        .max_snapshot_bytes()
                        .saturating_sub(self.snapshot_bytes),
                )?);
            }
            let (had_prior, staged) = {
                let run = self.xls_cells.as_mut().ok_or_else(|| {
                    crate::Error::new(Code::InternalError, "XLS cell run missing")
                })?;
                (run.first_index().is_some(), run.stage(plan, _index))
            };
            let changed = match staged {
                Ok(changed) => changed,
                Err(error) => {
                    // Native staging is failure-atomic. Publish the earlier
                    // prefix first so its deferred error keeps wire precedence.
                    if had_prior {
                        self.flush_pending()?;
                        if xls_staging_full(&error) {
                            return self.coalesce_parsed(_target, _props, _index, _spelling);
                        }
                    }
                    return Err(error);
                },
            };
            self.operations += 1;
            return Ok(Some(changed));
        }
        #[cfg(feature = "numbers")]
        if let File::Numbers(handle) = &self.current {
            let planned = match &mut self.numbers_cells {
                Some(run) => run.cell_plan(_target, _props),
                None => handle.numbers_cell_plan(_target, _props, None),
            };
            if planned.is_err() {
                self.flush_pending()?;
            }
            let Some(mut plan) = planned? else {
                if self.numbers_cells.is_some() {
                    self.flush_pending()?;
                }
                return Ok(None);
            };
            if self.operations >= self.limits.max_operations() {
                return Err(Error::Limit {
                    kind: LimitKind::Operations,
                    observed: self.operations.checked_add(1).unwrap_or(usize::MAX),
                    maximum: self.limits.max_operations(),
                }
                .into());
            }
            if self
                .numbers_cells
                .as_ref()
                .is_some_and(|run| run.requires_barrier(&plan))
            {
                self.flush_pending()?;
                let File::Numbers(handle) = &self.current else {
                    return Err(pending());
                };
                let Some(replanned) = handle.numbers_cell_plan(_target, _props, None)? else {
                    return Ok(None);
                };
                plan = replanned;
            }
            if self.numbers_cells.is_none() {
                let File::Numbers(handle) = &self.current else {
                    return Err(pending());
                };
                self.numbers_cells = Some(crate::format::iwork::numbers::coalesce::Run::new(
                    handle,
                    &self.limits,
                )?);
            }
            let run = self.numbers_cells.as_mut().ok_or_else(|| {
                crate::Error::new(Code::InternalError, "Numbers cell run missing")
            })?;
            let changed = run.stage(plan, _index)?;
            self.operations += 1;
            return Ok(Some(changed));
        }
        #[cfg(feature = "ods")]
        if let File::Ods(_) = &self.current {
            let planned = match &self.ods_cells {
                Some(run) => run.cell_plan(_target, _props, &self.limits),
                None => {
                    let File::Ods(handle) = &self.current else {
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
            let Some(mut plan) = planned? else {
                return Ok(None);
            };
            if self.operations >= self.limits.max_operations() {
                return Err(Error::Limit {
                    kind: LimitKind::Operations,
                    observed: self.operations.checked_add(1).unwrap_or(usize::MAX),
                    maximum: self.limits.max_operations(),
                }
                .into());
            }
            if self
                .ods_cells
                .as_ref()
                .is_some_and(|run| run.requires_barrier(&plan, &self.limits))
            {
                self.flush_pending()?;
                let File::Ods(handle) = &self.current else {
                    return Err(pending());
                };
                let Some(replanned) = handle.cell_plan(_target, _props, &self.limits)? else {
                    return Ok(None);
                };
                plan = replanned;
            }
            if self.ods_cells.is_none() {
                let File::Ods(handle) = &self.current else {
                    return Err(pending());
                };
                self.ods_cells = Some(crate::format::ods::write::coalesce::Run::new(
                    handle,
                    &self.limits,
                )?);
            }
            let run = self
                .ods_cells
                .as_mut()
                .ok_or_else(|| crate::Error::new(Code::InternalError, "ODS cell run missing"))?;
            let changed = run.stage(plan, _index)?;
            self.operations += 1;
            return Ok(Some(changed));
        }
        #[cfg(feature = "xlsx")]
        if let File::Xlsx(handle) = &self.current {
            let planned = match &self.cells {
                Some(run) => run.cell_plan(_target, _props, &self.limits),
                None => handle.cell_plan_with_snapshot_limit(
                    _target,
                    _props,
                    &self.limits,
                    self.limits
                        .max_snapshot_bytes()
                        .saturating_sub(self.snapshot_bytes),
                ),
            };
            if planned.is_err() {
                // An earlier deferred native failure takes precedence over
                // admission failure in the following command.
                self.flush_pending()?;
            }
            let Some(mut plan) = planned? else {
                return Ok(None);
            };
            if self.operations >= self.limits.max_operations() {
                return Err(Error::Limit {
                    kind: LimitKind::Operations,
                    observed: self.operations.checked_add(1).unwrap_or(usize::MAX),
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
                let promoted_source = plan.take_promoted_source();
                self.cells = Some(crate::format::xlsx::write::coalesce::Run::new(
                    handle,
                    &self.limits,
                    self.limits
                        .max_snapshot_bytes()
                        .saturating_sub(self.snapshot_bytes),
                    promoted_source,
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
                    observed: self.operations.checked_add(1).unwrap_or(usize::MAX),
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
                    self.limits
                        .max_snapshot_bytes()
                        .saturating_sub(self.snapshot_bytes),
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
        #[cfg(feature = "ppt")]
        if let File::Ppt(handle) = &self.current {
            let planned = match &mut self.ppt_shapes {
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
                    observed: self.operations.checked_add(1).unwrap_or(usize::MAX),
                    maximum: self.limits.max_operations(),
                }
                .into());
            }
            if self
                .ppt_shapes
                .as_ref()
                .is_some_and(|run| run.requires_barrier(&plan, &self.limits))
            {
                self.flush_pending()?;
                // A repeated shape may merge partial geometry with earlier
                // assignments. Resolve it again against the published run.
                let File::Ppt(handle) = &self.current else {
                    return Err(pending());
                };
                let Some(replanned) = handle.shape_plan(_target, _props, &self.limits)? else {
                    return Ok(None);
                };
                plan = replanned;
            }
            if self.ppt_shapes.is_none() {
                let File::Ppt(handle) = &self.current else {
                    return Err(pending());
                };
                self.ppt_shapes = Some(crate::format::ppt::write::coalesce::Run::new(
                    handle,
                    &self.limits,
                    self.limits
                        .max_snapshot_bytes()
                        .saturating_sub(self.snapshot_bytes),
                )?);
            }
            let run = self
                .ppt_shapes
                .as_mut()
                .ok_or_else(|| crate::Error::new(Code::InternalError, "shape run missing"))?;
            let changed = run.stage(plan, _index)?;
            self.operations += 1;
            return Ok(Some(changed));
        }
        #[cfg(feature = "odp")]
        if let File::Odp(handle) = &self.current {
            // Sets resolve against the pending run's base: the run holds
            // only sets that do not change any positional address.
            let planned = match &self.odp {
                Some(run) => run.base().set_plan(_target, _props, &self.limits),
                None => handle.set_plan(_target, _props, &self.limits),
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
                    observed: self.operations.checked_add(1).unwrap_or(usize::MAX),
                    maximum: self.limits.max_operations(),
                }
                .into());
            }
            if self
                .odp
                .as_ref()
                .is_some_and(|run| run.requires_barrier(&plan, &self.limits))
            {
                self.flush_pending()?;
                let File::Odp(handle) = &self.current else {
                    return Err(pending());
                };
                let Some(replanned) = handle.set_plan(_target, _props, &self.limits)? else {
                    return Ok(None);
                };
                plan = replanned;
            }
            if self.odp.is_none() {
                let File::Odp(handle) = &self.current else {
                    return Err(pending());
                };
                self.odp = Some(crate::format::odp::write::coalesce::Run::new(handle));
            }
            let run = self
                .odp
                .as_mut()
                .ok_or_else(|| crate::Error::new(Code::InternalError, "ODP run missing"))?;
            run.stage(plan, _index);
            self.operations += 1;
            return Ok(Some(true));
        }
        #[cfg(odf_objects)]
        if let Some(handle) = self.current.odf_tree() {
            // Sets resolve against the pending run's base: the run holds
            // only sets that change no address and read no pending value.
            let planned = match &self.odf_tree {
                Some(run) => run.base().set_plan(_target, _props),
                None => handle.set_plan(_target, _props),
            };
            if planned.is_err() {
                self.flush_pending()?;
            }
            let Some(mut node) = planned? else {
                self.flush_pending()?;
                return Ok(None);
            };
            if self.operations >= self.limits.max_operations() {
                return Err(Error::Limit {
                    kind: LimitKind::Operations,
                    observed: self.operations.checked_add(1).unwrap_or(usize::MAX),
                    maximum: self.limits.max_operations(),
                }
                .into());
            }
            if self
                .odf_tree
                .as_ref()
                .is_some_and(|run| run.requires_barrier(node, _props, &self.limits))
            {
                self.flush_pending()?;
                let Some(handle) = self.current.odf_tree() else {
                    return Err(pending());
                };
                let Some(replanned) = handle.set_plan(_target, _props)? else {
                    return Ok(None);
                };
                node = replanned;
            }
            if self.odf_tree.is_none() {
                let Some(handle) = self.current.odf_tree() else {
                    return Err(pending());
                };
                self.odf_tree = Some(crate::format::odf_tree::coalesce::Run::new(handle));
            }
            let run = self
                .odf_tree
                .as_mut()
                .ok_or_else(|| crate::Error::new(Code::InternalError, "object run missing"))?;
            run.stage(node, _props, _index);
            self.operations += 1;
            return Ok(Some(true));
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
        #[cfg(feature = "rtf")]
        if let Some(run) = self.rtf.take() {
            let index = run.first_index();
            let result = (|| -> Result<()> {
                let (base, commit) = run.commit()?;
                let (step, ()) = wrap_rtf(&base, commit, ())?;
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
        #[cfg(feature = "rtf")]
        if let Some(run) = self.rtf_pictures.take() {
            let index = run.first_index();
            let result = (|| -> Result<()> {
                let (base, commit) = run.commit(&self.limits)?;
                let (step, ()) = wrap_rtf(&base, commit, ())?;
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
        #[cfg(feature = "pages")]
        if let Some(run) = self.pages.take() {
            let index = run.first_index();
            let result = (|| -> Result<()> {
                let (base, commit) = run.commit()?;
                let (step, ()) = wrap_iwork(&base, commit, ())?;
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
        #[cfg(feature = "keynote")]
        if let Some(run) = self.keynote.take() {
            let first = run.first_index();
            let result = (|| -> std::result::Result<(), (crate::Error, Option<usize>)> {
                let (base, commit) = run
                    .commit()
                    .map_err(|(error, index)| (error, Some(index)))?;
                let (step, ()) = wrap_iwork(&base, commit, ()).map_err(|error| (error, first))?;
                let bytes = self.snapshot_bytes.saturating_add(step.snapshot_bytes);
                if bytes > self.limits.max_snapshot_bytes() {
                    return Err((
                        Error::Limit {
                            kind: LimitKind::SnapshotBytes,
                            observed: bytes,
                            maximum: self.limits.max_snapshot_bytes(),
                        }
                        .into(),
                        first,
                    ));
                }
                self.steps
                    .try_reserve(1)
                    .map_err(|source| (Error::Allocation { source }.into(), first))?;
                self.current = step.after.clone();
                self.steps.push(step);
                self.snapshot_bytes = bytes;
                Ok(())
            })();
            result
                .map_err(|(error, index)| self.poison(error, index.unwrap_or(self.operations)))?;
        }
        #[cfg(feature = "odt")]
        if let Some(run) = self.odt.take() {
            let index = run.first_index();
            let result = (|| -> Result<()> {
                let (base, commit) = run.commit()?;
                let (step, ()) = wrap_odt(&base, commit, ())?;
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
        #[cfg(feature = "xls")]
        if let Some(run) = self.xls_cells.take() {
            let index = run.first_index();
            let result = (|| -> Result<()> {
                let (base, commit) = run.commit()?;
                let (step, ()) = wrap_xls(&base, commit, ())?;
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
                let (base, commit, reported) = run.commit()?;
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
                self.reported
                    .try_reserve(reported.len())
                    .map_err(|source| Error::Allocation { source })?;
                self.reported.extend(reported);
                self.current = step.after.clone();
                self.steps.push(step);
                self.snapshot_bytes = bytes;
                Ok(())
            })();
            result.map_err(|error| self.poison(error, index.unwrap_or(self.operations)))?;
        }
        #[cfg(feature = "ods")]
        if let Some(run) = self.ods_cells.take() {
            let index = run.first_index();
            let result = (|| -> Result<()> {
                let (base, commit) = run.commit()?;
                let (step, ()) = wrap_ods(&base, commit, ())?;
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
        #[cfg(feature = "ppt")]
        if let Some(run) = self.ppt_shapes.take() {
            let index = run.first_index();
            let result = (|| -> Result<()> {
                let (base, commit) = run.commit()?;
                let (step, ()) = wrap_ppt(&base, commit, ())?;
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
        #[cfg(feature = "numbers")]
        if let Some(run) = self.numbers_cells.take() {
            let index = run.first_index();
            let result = (|| -> Result<()> {
                let (base, commit) = run.commit()?;
                let (step, ()) = wrap_iwork(&base, commit, ())?;
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
        #[cfg(odf_objects)]
        if let Some(run) = self.odf_tree.take() {
            let index = run.first_index();
            let limits = self.limits;
            let result = (|| -> Result<()> {
                let (base, commit) = run.commit(&limits)?;
                let (step, ()) = wrap_object(&base, commit, ())?;
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
        #[cfg(feature = "odp")]
        if let Some(run) = self.odp.take() {
            let index = run.first_index();
            let limits = self.limits;
            let result = (|| -> Result<()> {
                let (base, commit) = run.commit(&limits)?;
                let (step, ()) = wrap_odp(&base, commit, ())?;
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
        let reported = self.reported.len();
        let result = operation(self).and_then(|value| {
            self.check()?;
            self.flush_pending()?;
            Ok((value, changed(&self.steps[start..])?))
        });
        // The isolated command's change is its steps' (record 0348).
        self.reported.truncate(reported);
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
                #[cfg(feature = "rtf")]
                {
                    self.rtf = None;
                }
                #[cfg(feature = "odt")]
                {
                    self.odt = None;
                }
                #[cfg(feature = "pages")]
                {
                    self.pages = None;
                }
                #[cfg(feature = "keynote")]
                {
                    self.keynote = None;
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
                #[cfg(feature = "xls")]
                {
                    self.xls_cells = None;
                }
                #[cfg(feature = "pptx")]
                {
                    self.shapes = None;
                }
                #[cfg(feature = "ppt")]
                {
                    self.ppt_shapes = None;
                }
                #[cfg(feature = "ods")]
                {
                    self.ods_cells = None;
                }
                #[cfg(feature = "numbers")]
                {
                    self.numbers_cells = None;
                }
                #[cfg(feature = "odp")]
                {
                    self.odp = None;
                }
                self.current = before;
                self.steps.truncate(start);
                self.snapshot_bytes = bytes;
                self.operations = operations;
                Err(error)
            },
        }
    }

    /// The match count of the last `find=` set, taken (record 0273).
    pub(crate) fn take_matched(&mut self) -> Option<usize> {
        self.matched.take()
    }

    /// The changes native runs reported at their commits, by item index,
    /// taken (record 0348): the items' staging returned `false`. Flush the
    /// pending run first to include it.
    pub(crate) fn take_reported(&mut self) -> Vec<(usize, bool)> {
        std::mem::take(&mut self.reported)
    }

    /// A `set` whose properties name `find` (record 0273): one native commit
    /// over the target's subtree, returning the match count.
    fn find_replace(
        &mut self,
        target: &crate::TargetArg<'_>,
        props: &crate::Props<'_>,
    ) -> Result<usize> {
        let spec = match crate::find::Spec::admit(props) {
            Ok(spec) => spec,
            Err(error) => {
                self.flush_pending()?;
                return Err(error);
            },
        };
        // Native find adapters capture their source before returning a step;
        // hand them the post-flush remainder so that capture is rejected
        // before a private editor is allocated.
        self.flush_pending()?;
        let limits = self.limits.with_snapshot_budget(
            self.limits
                .max_snapshot_bytes()
                .saturating_sub(self.snapshot_bytes),
        );
        let _ = (target, &spec, &limits);
        #[cfg(feature = "xls")]
        if matches!(self.current, File::Xls(_)) {
            return self.execute_steps(|file| match file {
                File::Xls(handle) => {
                    let (commits, matched) = crate::format::xls::write::find::find_replace(
                        handle, target, &spec, &limits,
                    )?;
                    Ok((xls_steps(commits)?, matched))
                },
                _ => Err(crate::find::unsupported(file)),
            });
        }
        self.execute(|file| match file {
            #[cfg(feature = "docx")]
            File::Docx(handle) => {
                let (commit, matched) = handle.find_replace(target, &spec, &limits)?;
                wrap(handle, commit, matched)
            },
            #[cfg(feature = "xlsx")]
            File::Xlsx(handle) => {
                let (commit, matched) = handle.find_replace(target, &spec, &limits)?;
                wrap_xlsx(handle, commit, matched)
            },
            #[cfg(feature = "pptx")]
            File::Pptx(handle) => {
                let (commit, matched) = handle.find_replace(target, &spec, &limits)?;
                wrap_pptx(handle, commit, matched)
            },
            #[cfg(feature = "rtf")]
            File::Rtf(handle) => {
                let (commit, matched) = handle.find_replace(target, &spec, &limits)?;
                wrap_rtf(handle, commit, matched)
            },
            #[cfg(feature = "doc")]
            File::Doc(handle) => {
                let (commit, matched) = handle.find_replace(target, &spec, &limits)?;
                wrap_doc(handle, commit, matched)
            },
            #[cfg(feature = "xlsb")]
            File::Xlsb(handle) => {
                let (commit, matched) =
                    crate::format::xlsb::write::find::find_replace(handle, target, &spec, &limits)?;
                wrap_xlsb(handle, commit, matched)
            },
            #[cfg(feature = "ods")]
            File::Ods(handle) => {
                let (commit, matched) =
                    crate::format::ods::write::find::find_replace(handle, target, &spec, &limits)?;
                wrap_ods(handle, commit, matched)
            },
            #[cfg(feature = "odt")]
            File::Odt(handle) | File::Oth(handle) | File::Odm(handle) => {
                let (commit, matched) =
                    crate::format::odt::write::find::find_replace(handle, target, &spec, &limits)?;
                wrap_odt(handle, commit, matched)
            },
            #[cfg(feature = "odp")]
            File::Odp(handle) => {
                let (commit, matched) =
                    crate::format::odp::write::find::find_replace(handle, target, &spec, &limits)?;
                wrap_odp(handle, commit, matched)
            },
            #[cfg(feature = "ppt")]
            File::Ppt(handle) => {
                let (commit, matched) =
                    crate::format::ppt::write::find::find_replace(handle, target, &spec, &limits)?;
                wrap_ppt(handle, commit, matched)
            },
            _ => Err(crate::find::unsupported(file)),
        })
    }

    /// Add to a Keynote presentation: a new slide, text box, shape or
    /// picture joins the transaction's Keynote run (record 0298); anything
    /// else publishes the run and runs alone.
    #[cfg(feature = "keynote")]
    fn add_keynote<'p, 'n>(
        &mut self,
        parent: impl IntoPath<'p>,
        node: impl IntoNode<'n>,
    ) -> Result<Path<'static>> {
        self.check()?;
        let index = self.operations;
        let staged = (|| {
            let (parent, node) = match (parent.into_path(), node.into_node()) {
                (Ok(parent), Ok(node)) => (parent, node),
                (Err(error), _) | (_, Err(error)) => {
                    self.flush_pending()?;
                    return Err(error);
                },
            };
            let File::Keynote(handle) = self.current.clone() else {
                return Err(pending());
            };
            if index >= self.limits.max_operations() {
                self.flush_pending()?;
                return Err(Error::Limit {
                    kind: LimitKind::Operations,
                    observed: index.checked_add(1).unwrap_or(usize::MAX),
                    maximum: self.limits.max_operations(),
                }
                .into());
            }
            let staged = match self.keynote.as_mut() {
                Some(run) => run.stage_add(&parent, &node, index),
                None => {
                    let mut run = crate::format::iwork::keynote::Run::new(&handle, &self.limits);
                    let staged = run.stage_add(&parent, &node, index);
                    if matches!(staged, Ok(Some(_))) {
                        self.keynote = Some(run);
                    }
                    staged
                },
            };
            match staged {
                Ok(Some(path)) => {
                    self.operations += 1;
                    Ok(Ok(path))
                },
                Ok(None) => Ok(Err((parent, node))),
                Err(error) => {
                    self.flush_pending()?;
                    Err(error)
                },
            }
        })();
        let (parent, node) = match staged {
            Ok(Ok(path)) => return Ok(path),
            Ok(Err(request)) => request,
            Err(error) => return Err(self.poison(error, index)),
        };
        // A native fallback is still one operation in this transaction.  The
        // staged Keynote run may already have published earlier work, so its
        // native before/after capture must receive the post-flush remainder
        // rather than the transaction-wide ceiling.
        self.flush_pending()?;
        let limits = self.limits.with_snapshot_budget(
            self.limits
                .max_snapshot_bytes()
                .saturating_sub(self.snapshot_bytes),
        );
        self.execute(|file| match file {
            File::Keynote(handle) => {
                let (commit, path) = handle.add_with_limits(parent, node, &limits)?;
                wrap_iwork(handle, commit, path)
            },
            _ => Err(pending()),
        })
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
        self.execute_steps(|file| {
            let (step, value) = operation(file)?;
            Ok((vec![step], value))
        })
    }

    /// [`Self::execute`] for one mutation that commits several native steps,
    /// each on the previous one's result (XLS chunks, record 0295).
    fn execute_steps<T>(
        &mut self,
        operation: impl FnOnce(&File) -> Result<(Vec<Step>, T)>,
    ) -> Result<T> {
        self.check()?;
        self.flush_pending()?;
        let index = self.operations;
        let result = (|| {
            if index >= self.limits.max_operations() {
                return Err(Error::Limit {
                    kind: LimitKind::Operations,
                    observed: index.checked_add(1).unwrap_or(usize::MAX),
                    maximum: self.limits.max_operations(),
                }
                .into());
            }
            let (steps, value) = operation(&self.current)?;
            let bytes = steps.iter().fold(self.snapshot_bytes, |bytes, step| {
                bytes.saturating_add(step.snapshot_bytes)
            });
            if bytes > self.limits.max_snapshot_bytes() {
                return Err(Error::Limit {
                    kind: LimitKind::SnapshotBytes,
                    observed: bytes,
                    maximum: self.limits.max_snapshot_bytes(),
                }
                .into());
            }
            self.steps
                .try_reserve(steps.len())
                .map_err(|source| Error::Allocation { source })?;
            for step in steps {
                self.current = step.after.clone();
                self.steps.push(step);
            }
            self.snapshot_bytes = bytes;
            self.operations += 1;
            Ok(value)
        })();
        result.map_err(|error| self.poison(error, index))
    }
}

/// The steps of chained XLS commits, each on the previous one's result.
#[cfg(feature = "xls")]
fn xls_steps(commits: Vec<crate::format::xls::WriteCommit>) -> Result<Vec<Step>> {
    commits
        .into_iter()
        .map(|commit| {
            let before = commit.before.clone();
            Ok(wrap_xls(&before, commit, ())?.0)
        })
        .collect()
}

#[cfg(feature = "docx")]
fn wrap<T>(
    before: &crate::format::docx::Docx,
    commit: crate::format::docx::write::Commit,
    value: T,
) -> Result<(Step, T)> {
    let leaf_operations = commit.operations();
    let (after, main, _diagnostics) = commit.into_parts();
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
            leaf_operations,
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
    let leaf_operations = commit.operations();
    let snapshot_bytes = commit.snapshot_bytes();
    let (after, patch) = commit.into_parts();
    // Structural and cell-value patches both retain exact package images.
    // Cell selectors spanning worksheets retain every per-worksheet image;
    // structural patches carry the checked native operation charge.
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

#[cfg(feature = "ppt")]
fn wrap_ppt<T>(
    _before: &crate::format::ppt::Ppt,
    commit: crate::format::ppt::write::WriteCommit,
    value: T,
) -> Result<(Step, T)> {
    let crate::format::ppt::write::WriteCommit {
        before,
        after,
        patch,
        operations,
        snapshot_bytes,
    } = commit;
    Ok((
        Step {
            before: File::Ppt(before),
            after: File::Ppt(after),
            leaf: Leaf::Ppt(Box::new(patch)),
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
    let leaf_operations = commit.operations();
    let snapshot_bytes = commit.snapshot_bytes();
    let (after, source, target, patch) = commit.into_parts();
    Ok((
        Step {
            before: File::Xlsx(before.clone()),
            after: File::Xlsx(after),
            leaf: Leaf::Xlsx(Box::new(XlsxLeaf {
                source,
                target,
                _patch: patch,
            })),
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
    let snapshot_bytes = commit.snapshot_bytes();
    Ok((
        Step {
            before: File::Pptx(before.clone()),
            after: File::Pptx(commit.after),
            leaf: Leaf::Pptx(Box::new(PptxLeaf {
                source: commit.source,
                target: commit.target,
                _proof: commit.proof,
            })),
            leaf_operations: commit.operations,
            snapshot_bytes,
        },
        value,
    ))
}

#[cfg(feature = "rtf")]
fn wrap_rtf<T>(
    _before: &crate::format::rtf::Rtf,
    commit: crate::format::rtf::write::Commit,
    value: T,
) -> Result<(Step, T)> {
    let crate::format::rtf::write::Commit {
        before,
        after,
        patch,
        operations,
        snapshot_bytes,
    } = commit;
    Ok((
        Step {
            before: File::Rtf(before),
            after: File::Rtf(after),
            leaf: Leaf::Rtf(Box::new(patch)),
            leaf_operations: operations,
            snapshot_bytes,
        },
        value,
    ))
}

#[cfg(feature = "ods")]
fn wrap_ods<T>(
    _before: &crate::format::ods::Ods,
    commit: crate::format::ods::WriteCommit,
    value: T,
) -> Result<(Step, T)> {
    let crate::format::ods::WriteCommit {
        before,
        after,
        patch,
        operations,
        snapshot_bytes,
    } = commit;
    Ok((
        Step {
            before: File::Ods(before),
            after: File::Ods(after),
            leaf: Leaf::Ods(Box::new(patch)),
            leaf_operations: operations,
            snapshot_bytes,
        },
        value,
    ))
}

#[cfg(odf_objects)]
fn wrap_object<T>(
    _before: &crate::format::odf_tree::OdfTree,
    commit: crate::format::odf_tree::write::Commit,
    value: T,
) -> Result<(Step, T)> {
    let crate::format::odf_tree::write::Commit {
        before,
        after,
        patch,
        operations,
        snapshot_bytes,
    } = commit;
    Ok((
        Step {
            before: before.into_file(),
            after: after.into_file(),
            leaf: Leaf::Object(Box::new(patch)),
            leaf_operations: operations,
            snapshot_bytes,
        },
        value,
    ))
}

#[cfg(feature = "iwork-adapter")]
fn wrap_iwork<T>(
    _before: &crate::format::iwork::Iwork,
    commit: crate::format::iwork::write::Commit,
    value: T,
) -> Result<(Step, T)> {
    let crate::format::iwork::write::Commit {
        before,
        after,
        patch,
        operations,
        snapshot_bytes,
    } = commit;
    Ok((
        Step {
            before: before.into_file(),
            after: after.into_file(),
            leaf: Leaf::Iwork(Box::new(patch)),
            leaf_operations: operations,
            snapshot_bytes,
        },
        value,
    ))
}

#[cfg(feature = "odt")]
fn wrap_odt<T>(
    _before: &crate::format::odt::Odt,
    commit: crate::format::odt::write::Commit,
    value: T,
) -> Result<(Step, T)> {
    let crate::format::odt::write::Commit {
        before,
        after,
        patch,
        operations,
        snapshot_bytes,
    } = commit;
    Ok((
        Step {
            before: before.into_file(),
            after: after.into_file(),
            leaf: Leaf::Odt(Box::new(patch)),
            leaf_operations: operations,
            snapshot_bytes,
        },
        value,
    ))
}

#[cfg(feature = "odp")]
fn wrap_odp<T>(
    _before: &crate::format::odp::Odp,
    commit: crate::format::odp::write::WriteCommit,
    value: T,
) -> Result<(Step, T)> {
    let crate::format::odp::write::WriteCommit {
        before,
        after,
        patch,
        operations,
        snapshot_bytes,
    } = commit;
    Ok((
        Step {
            before: File::Odp(before),
            after: File::Odp(after),
            leaf: Leaf::Odp(Box::new(patch)),
            leaf_operations: operations,
            snapshot_bytes,
        },
        value,
    ))
}

fn changed(steps: &[Step]) -> Result<bool> {
    let (Some(first), Some(last)) = (steps.first(), steps.last()) else {
        return Ok(false);
    };
    #[cfg(feature = "ods")]
    if let (File::Ods(before), File::Ods(after)) = (&first.before, &last.after) {
        if let [
            Step {
                leaf: Leaf::Ods(patch),
                ..
            },
        ] = steps
        {
            return Ok(!patch.is_noop());
        }
        return Ok(!before.same_content(after)?);
    }
    #[cfg(feature = "odp")]
    if let (File::Odp(before), File::Odp(after)) = (&first.before, &last.after) {
        if let [
            Step {
                leaf: Leaf::Odp(patch),
                ..
            },
        ] = steps
        {
            return Ok(!patch.is_noop());
        }
        return Ok(!before.same_content(after)?);
    }
    #[cfg(feature = "rtf")]
    if let (File::Rtf(before), File::Rtf(after)) = (&first.before, &last.after) {
        if let [
            Step {
                leaf: Leaf::Rtf(patch),
                ..
            },
        ] = steps
        {
            return Ok(!patch.is_noop());
        }
        return Ok(!before.same_content(after)?);
    }
    #[cfg(odf_objects)]
    if let (Some(before), Some(after)) = (first.before.odf_tree(), last.after.odf_tree()) {
        if let [
            Step {
                leaf: Leaf::Object(patch),
                ..
            },
        ] = steps
        {
            return Ok(!patch.is_noop());
        }
        return Ok(!before.same_content(after));
    }
    #[cfg(feature = "iwork-adapter")]
    if let (Some(before), Some(after)) = (first.before.iwork(), last.after.iwork()) {
        if let [
            Step {
                leaf: Leaf::Iwork(patch),
                ..
            },
        ] = steps
        {
            return Ok(!patch.is_noop());
        }
        return before.same_content(after).map(|same| !same);
    }
    #[cfg(feature = "odt")]
    if let (Some(before), Some(after)) = (first.before.odt_family(), last.after.odt_family()) {
        if let [
            Step {
                leaf: Leaf::Odt(patch),
                ..
            },
        ] = steps
        {
            return Ok(!patch.is_noop());
        }
        return Ok(!before.same_content(after)?);
    }
    match (&first.leaf, &last.leaf) {
        #[cfg(feature = "doc")]
        (Leaf::Doc(first), Leaf::Doc(last)) => Ok(first.before() != last.after()),
        #[cfg(feature = "xlsb")]
        (Leaf::Xlsb(first), Leaf::Xlsb(last)) => Ok(first.before() != last.after()),
        #[cfg(feature = "xls")]
        (Leaf::Xls(first), Leaf::Xls(last)) => Ok(first.before() != last.after()),
        #[cfg(feature = "ppt")]
        (Leaf::Ppt(first), Leaf::Ppt(last)) => Ok(first.before() != last.after()),
        #[cfg(feature = "pptx")]
        (Leaf::Pptx(first), Leaf::Pptx(last)) => Ok(!first.source.same_content(&last.target)?),
        #[cfg(feature = "docx")]
        (Leaf::Docx(first_patch), Leaf::Docx(last_patch)) => {
            // A template merge can change only a header or other story
            // part: a package-graph change is a change (record 0274).
            Ok(
                first_patch.source().xml_bytes() != last_patch.target().xml_bytes()
                || steps.iter().any(|step| match &step.leaf {
                    Leaf::Docx(patch) => patch.operations().iter().any(|operation| {
                        matches!(
                            operation,
                            litchi_docx::document::Operation::ChangeGraph { .. }
                        )
                    }),
                    #[allow(unreachable_patterns)]
                    _ => false,
                })
                // A metadata step changes only the core-properties part
                // (record 0338).
                || docx_metadata_changed(&first.before, &last.after)?,
            )
        },
        #[cfg(feature = "xlsx")]
        (Leaf::Xlsx(first), Leaf::Xlsx(last)) => Ok(!first.source.same_content(&last.target)?),
        #[allow(unreachable_patterns)]
        _ => Err(Error::PatchConflict.into()),
    }
}

/// Whether two DOCX endpoints carry different core properties (record
/// 0338), or different application or custom properties (record 0342).
#[cfg(feature = "docx")]
fn docx_metadata_changed(before: &File, after: &File) -> Result<bool> {
    match (before, after) {
        (File::Docx(before), File::Docx(after)) => {
            if before.same_snapshot(after) {
                return Ok(false);
            }
            if before.metadata()? != after.metadata()? {
                return Ok(true);
            }
            Ok(
                match (before.property_part_bytes()?, after.property_part_bytes()?) {
                    (Some(before), Some(after)) => before != after,
                    // Unreadable parts prove no property change (a metadata
                    // write always publishes a readable eager owner).
                    _ => false,
                },
            )
        },
        #[allow(unreachable_patterns)]
        _ => Ok(false),
    }
}

fn same_snapshot(left: &File, right: &File) -> bool {
    match (left, right) {
        #[cfg(feature = "doc")]
        (File::Doc(left), File::Doc(right)) => left.same_snapshot(right),
        #[cfg(feature = "xlsb")]
        (File::Xlsb(left), File::Xlsb(right)) => left.same_snapshot(right),
        #[cfg(feature = "odt")]
        (File::Odt(left), File::Odt(right))
        | (File::Oth(left), File::Oth(right))
        | (File::Odm(left), File::Odm(right)) => left.same_snapshot(right),
        #[cfg(odf_objects)]
        (File::Odg(left), File::Odg(right))
        | (File::Odc(left), File::Odc(right))
        | (File::Odi(left), File::Odi(right))
        | (File::Odb(left), File::Odb(right))
        | (File::Odf(left), File::Odf(right)) => left.same_snapshot(right),
        #[cfg(feature = "iwork-adapter")]
        (File::Pages(left), File::Pages(right))
        | (File::Keynote(left), File::Keynote(right))
        | (File::Numbers(left), File::Numbers(right)) => left.same_snapshot(right),
        #[cfg(feature = "xls")]
        (File::Xls(left), File::Xls(right)) => left.same_snapshot(right),
        #[cfg(feature = "ppt")]
        (File::Ppt(left), File::Ppt(right)) => left.same_snapshot(right),
        #[cfg(feature = "pptx")]
        (File::Pptx(left), File::Pptx(right)) => left.same_snapshot(right),
        #[cfg(feature = "docx")]
        (File::Docx(left), File::Docx(right)) => left.same_snapshot(right),
        #[cfg(feature = "xlsx")]
        (File::Xlsx(left), File::Xlsx(right)) => left.same_snapshot(right),
        (File::Signed(left), File::Signed(right)) => same_snapshot(left.file(), right.file()),
        #[cfg(feature = "rtf")]
        (File::Rtf(left), File::Rtf(right)) => left.same_snapshot(right),
        #[cfg(feature = "ods")]
        (File::Ods(left), File::Ods(right)) => left.same_snapshot(right),
        #[cfg(feature = "odp")]
        (File::Odp(left), File::Odp(right)) => left.same_snapshot(right),
        _ => false,
    }
}

/// Compare two separately opened plain files without retaining either
/// publication in a second `Vec`. The exact-owner path is checked by
/// [`same_snapshot`] before this function is called. A length is compared in
/// addition to SHA-256 so the proof cannot equate a digest of a prefix with a
/// different-length publication. This is a process-local replay proof; it is
/// never serialized into a durable patch envelope.
fn same_content(left: &File, right: &File) -> Result<bool> {
    if std::mem::discriminant(left) != std::mem::discriminant(right)
        || !same_security_disposition(left, right)
        || matches!(left, File::Signed(_) | File::Locked(_))
        || matches!(right, File::Signed(_) | File::Locked(_))
    {
        return Ok(false);
    }
    Ok(stream_fingerprint(left)? == stream_fingerprint(right)?)
}

/// Compare a replay recipient with the recorded base. Exact no-op replay
/// compares each endpoint's actual published bytes, including signature
/// parts; changed patches keep the stricter [`same_content`] policy.
pub(crate) fn same_content_for_replay(left: &File, right: &File, exact_noop: bool) -> Result<bool> {
    if !exact_noop || (!left.is_signed() && !right.is_signed()) {
        return same_content(left, right);
    }
    if matches!(left, File::Locked(_))
        || matches!(right, File::Locked(_))
        || left.format() != right.format()
        || !same_security_disposition(left, right)
    {
        return Ok(false);
    }
    Ok(stream_fingerprint(left)? == stream_fingerprint(right)?)
}

/// Content equality cannot transfer a signature or encryption disposition.
/// This helper keeps an endpoint whose publication happens to be equal from
/// matching a differently protected endpoint.
fn same_security_disposition(left: &File, right: &File) -> bool {
    if left.is_signed() != right.is_signed() {
        return false;
    }
    match (left, right) {
        #[cfg(all(feature = "pptx", feature = "encryption"))]
        (File::Pptx(left), File::Pptx(right)) => {
            left.requires_reencryption() == right.requires_reencryption()
        },
        #[cfg(all(feature = "xlsx", feature = "encryption"))]
        (File::Xlsx(left), File::Xlsx(right)) => {
            left.requires_reencryption() == right.requires_reencryption()
        },
        #[cfg(feature = "xls")]
        (File::Xls(left), File::Xls(right)) => left.is_encrypted() == right.is_encrypted(),
        #[cfg(feature = "ppt")]
        (File::Ppt(left), File::Ppt(right)) => left.is_encrypted() == right.is_encrypted(),
        _ => true,
    }
}

#[derive(Debug, PartialEq, Eq)]
struct StreamFingerprint {
    bytes: u64,
    digest: [u8; 32],
}

fn stream_fingerprint(file: &File) -> Result<StreamFingerprint> {
    let mut sink = DigestSink::default();
    let bytes = file.write_to(&mut sink)?;
    debug_assert_eq!(bytes, sink.bytes);
    Ok(StreamFingerprint {
        bytes,
        digest: sink.hasher.finalize().into(),
    })
}

#[derive(Default)]
struct DigestSink {
    hasher: Sha256,
    bytes: u64,
}

impl Write for DigestSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        self.hasher.update(bytes);
        self.bytes = self
            .bytes
            .checked_add(bytes.len() as u64)
            .ok_or_else(|| io::Error::other("content length exceeds u64"))?;
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

fn admit(file: &File) -> Result<()> {
    // A signed source may establish an exact no-op patch.  The closure is
    // still allowed to run so that the transaction can distinguish that
    // source-only case from an attempted mutation; every mutation path below
    // keeps the signed refusal and the final changed check is a second fence.
    if file.is_signed() {
        return check_source_for_replay(file);
    }
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
        #[cfg(feature = "ppt")]
        File::Ppt(handle) => handle.admit_edit(),
        #[cfg(feature = "rtf")]
        File::Rtf(handle) => handle.admit_edit(),
        #[cfg(feature = "ods")]
        File::Ods(handle) => handle.admit_edit(),
        #[cfg(feature = "odt")]
        File::Odt(handle) | File::Oth(handle) | File::Odm(handle) => handle.admit_edit(),
        #[cfg(odf_objects)]
        File::Odg(handle)
        | File::Odc(handle)
        | File::Odi(handle)
        | File::Odb(handle)
        | File::Odf(handle) => handle.admit_edit(),
        #[cfg(feature = "iwork-adapter")]
        File::Pages(handle) | File::Keynote(handle) | File::Numbers(handle) => handle.admit_edit(),
        #[cfg(feature = "odp")]
        File::Odp(handle) => handle.admit_edit(),
        _ => check_source(file),
    }
}

/// Return the format owner used while staging an explicit transaction.
///
/// The signed facade is a read-only wrapper around a format leaf.  The leaf
/// keeps its signed disposition, so it still enforces source freshness and
/// the final changed-publication fence.  Staging against that leaf lets the
/// format coalescer see a cancellation (change followed by restoration),
/// while [`File::edit`] returns the original facade when the complete
/// transaction is a no-op.
fn transaction_owner(file: &File) -> File {
    match file {
        File::Signed(signed) => signed.file().clone(),
        _ => file.clone(),
    }
}

fn check_source(file: &File) -> Result<()> {
    // Source freshness wins over the protected-state classification.  The
    // leaf checks below deliberately validate the retained owner without
    // consuming the signed no-op exception; this fence adds the refusal only
    // to changed edits and replay patches.
    check_source_for_replay(file)?;
    if file.is_signed() {
        return Err(crate::Error::signed(file.format().ok_or_else(pending)?));
    }
    Ok(())
}

/// Check retained source freshness while leaving the signature disposition
/// to the exact byte proof in [`same_content_for_replay`]. This is used only
/// by no-op replay; edit admission and changed patches continue through
/// [`check_source`].
fn check_source_for_replay(file: &File) -> Result<()> {
    match file {
        #[cfg(feature = "doc")]
        File::Doc(handle) => handle.check_source(),
        #[cfg(feature = "pptx")]
        File::Pptx(handle) => handle.check_source_for_replay(),
        #[cfg(feature = "docx")]
        File::Docx(handle) => handle.owner().check_source(),
        #[cfg(feature = "xlsx")]
        File::Xlsx(handle) => handle.check_source(),
        #[cfg(feature = "xlsb")]
        File::Xlsb(handle) => handle.check_source_for_replay(),
        #[cfg(feature = "xls")]
        File::Xls(handle) => handle.check_source(),
        #[cfg(feature = "ppt")]
        File::Ppt(handle) => handle.check_source(),
        #[cfg(feature = "rtf")]
        File::Rtf(_) => Ok(()),
        #[cfg(feature = "ods")]
        File::Ods(handle) => handle.check_source(),
        #[cfg(feature = "odt")]
        File::Odt(handle) | File::Oth(handle) | File::Odm(handle) => handle.check_source(),
        #[cfg(odf_objects)]
        File::Odg(handle)
        | File::Odc(handle)
        | File::Odi(handle)
        | File::Odb(handle)
        | File::Odf(handle) => handle.check_source(),
        #[cfg(feature = "iwork-adapter")]
        File::Pages(handle) | File::Keynote(handle) | File::Numbers(handle) => {
            handle.check_source()
        },
        #[cfg(feature = "odp")]
        File::Odp(handle) => handle.check_source(),
        File::Signed(signed) => check_source_for_replay(signed.file()),
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

// A native staging boundary closes a run; caller limits and other native
// errors remain failures. Retrying on an empty run is deliberately excluded.
#[cfg(feature = "xls")]
fn xls_staging_full(error: &crate::Error) -> bool {
    matches!(
        error.leaf(),
        crate::error::Source::XlsEdit(litchi_xls::package::edit::Error::Limit {
            kind: litchi_xls::package::edit::LimitKind::Operations,
            ..
        })
    )
}
