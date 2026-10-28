//! Immutable XLSX cell writes (record 0097).
//!
//! This adapter resolves direct cells, ranges and scoped selectors, admits
//! every property through the XLSX write registry, converts authored scalar
//! values to the native cell model, and commits one isolated native workbook
//! edit. The native worksheet API is address-oriented, so a bounded range or
//! selector is expanded before that one edit and all addresses are staged in
//! the same native transaction.

use litchi_core::FileFormat;
use litchi_xlsx::{self as xlsx, Address, Rect};
use std::borrow::Cow;
use std::collections::{HashSet, TryReserveError};
use std::mem::size_of;

use super::node::{Node, SheetHandle};
use super::spec::{CELL_WRITE_TABLE, CellSetter};
use super::{Nav, Xlsx};
use crate::error::{Code, Error, Result};
use crate::props::{IntoProps, PropValue, Props};
use crate::registry::Operation;
use crate::target::TargetArg;
use crate::transaction::{Error as TransactionError, LimitKind, SemanticLimits};
use crate::value::Scalar;

/// Complete result of one XLSX cell operation.
///
/// The native source and target snapshots remain alongside the native patch;
/// the transaction coordinator uses them for source-bound replay and final
/// semantic equality without serializing and reopening the workbook.
pub(crate) struct Commit {
    after_xlsx: Xlsx,
    source: xlsx::Package,
    target: xlsx::Package,
    native_patch: xlsx::workbook::Patch,
    operations: usize,
    snapshot_bytes: usize,
}

impl Commit {
    /// Adapt a native commit while enforcing the caller's snapshot budget
    /// before a signed changed patch can be discarded for canonical replay.
    pub(super) fn from_native_with_limit(
        base: &Xlsx,
        source: xlsx::Package,
        native: xlsx::workbook::Commit,
        max_snapshot_bytes: usize,
    ) -> Result<Self> {
        base.check_source_for_replay()?;
        let result = (|| {
            let (target, native_patch) = native.into_parts();
            let native_operations = native_patch.operation_count();
            // Signed writes may discard this private probe for canonical
            // replay. Unsigned writes retain the native proof directly so
            // every route has the same nonzero snapshot charge.
            let signed = base.signature() == super::Signature::Signed;
            let native_snapshot_bytes = native_patch.retained_bytes();
            if native_snapshot_bytes > max_snapshot_bytes {
                return Err(snapshot_limit(native_snapshot_bytes, max_snapshot_bytes));
            }
            let changed = !source.same_content(&target).map_err(native_error)?;
            let (source, target, native_patch, canonical_snapshot_bytes) = if !changed && signed {
                // Writes stage through an unsigned private workbook. A
                // no-op patch must instead be rooted in the retained
                // signed package so native replay and File use identical
                // physical base bytes.
                // Release the staged package and proof before promoting the
                // signed owner. The final charge intentionally remains
                // max(staged, canonical), but the two endpoint graphs must
                // never be live together while the canonical probe is built.
                drop((source, target, native_patch));
                let original = base.original_package_for_noop()?;
                let (target, native_patch) =
                    edit_with_snapshot_limit(&original, max_snapshot_bytes)?
                        .commit()
                        .map_err(native_error)?
                        .into_parts();
                let canonical_snapshot_bytes = native_patch.retained_bytes();
                (original, target, native_patch, canonical_snapshot_bytes)
            } else {
                (source, target, native_patch, native_snapshot_bytes)
            };
            let snapshot_bytes = native_snapshot_bytes.max(canonical_snapshot_bytes);
            if snapshot_bytes > max_snapshot_bytes {
                return Err(snapshot_limit(snapshot_bytes, max_snapshot_bytes));
            }
            let after_xlsx = if changed {
                base.adopt_workbook(target.clone())
            } else {
                base.clone()
            };
            let commit = Self {
                after_xlsx,
                source,
                target,
                native_patch,
                operations: native_operations,
                snapshot_bytes,
            };
            Ok(commit)
        })();
        base.check_source_for_replay()?;
        result
    }

    /// Adapt a metadata commit (record 0253). The workbook patch is empty:
    /// the leaf metadata patch changes only the `docProps` parts, which the
    /// retained endpoints' content comparison observes.
    pub(super) fn from_metadata(
        base: &Xlsx,
        after_xlsx: Xlsx,
        source: xlsx::Package,
        target: xlsx::Package,
        changed: bool,
        authored_operations: usize,
        native_patch_bytes: usize,
        max_snapshot_bytes: usize,
    ) -> Result<Self> {
        if changed && base.signature() == super::Signature::Signed {
            return Err(Error::signed(FileFormat::Xlsx));
        }
        if !changed && base.signature() == super::Signature::Signed {
            // Feature and metadata families stage through a private unsigned
            // package. Charge both real endpoint owners before dropping that
            // probe; the canonical identity patch below may retain more.
            let staged_snapshot_bytes = source
                .retained_bytes()
                .saturating_add(target.retained_bytes())
                .saturating_add(native_patch_bytes);
            // Metadata/features use the leaf metadata adapter rather than the
            // primary workbook editor. Rebuild a genuine identity workbook
            // patch over the retained signed owner instead of publishing the
            // unsigned private probe as the replay base.
            drop((source, target));
            let original = base.original_package_for_noop()?;
            let (target, native_patch) = edit_with_snapshot_limit(&original, max_snapshot_bytes)?
                .commit()
                .map_err(native_error)?
                .into_parts();
            let canonical_snapshot_bytes = native_patch.retained_bytes();
            let snapshot_bytes = staged_snapshot_bytes.max(canonical_snapshot_bytes);
            if snapshot_bytes > max_snapshot_bytes {
                return Err(snapshot_limit(snapshot_bytes, max_snapshot_bytes));
            }
            let commit = Self {
                after_xlsx: base.clone(),
                source: original,
                target,
                native_patch,
                operations: authored_operations,
                snapshot_bytes,
            };
            return Ok(commit);
        }
        let snapshot_bytes = source
            .retained_bytes()
            .saturating_add(target.retained_bytes())
            .saturating_add(native_patch_bytes);
        let commit = Self {
            after_xlsx,
            source,
            target,
            native_patch: xlsx::workbook::Patch::default(),
            operations: authored_operations,
            snapshot_bytes,
        };
        if commit.snapshot_bytes > max_snapshot_bytes {
            return Err(snapshot_limit(commit.snapshot_bytes, max_snapshot_bytes));
        }
        Ok(commit)
    }

    /// Consume the adapter result for explicit transaction history.
    pub(crate) fn into_parts(self) -> (Xlsx, xlsx::Package, xlsx::Package, xlsx::workbook::Patch) {
        (self.after_xlsx, self.source, self.target, self.native_patch)
    }

    /// Consume the result for the format-neutral one-shot setter.
    pub(crate) fn into_file(self) -> Xlsx {
        self.after_xlsx
    }

    pub(crate) fn operations(&self) -> usize {
        self.operations
    }

    pub(crate) const fn snapshot_bytes(&self) -> usize {
        self.snapshot_bytes
    }

    pub(super) fn is_noop_for(&self, base: &Xlsx) -> bool {
        self.after_xlsx.same_snapshot(base)
    }

    /// The handle after the operation.
    pub(super) fn file(&self) -> &Xlsx {
        &self.after_xlsx
    }
}

impl std::fmt::Debug for Commit {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Commit")
            .field("native_patch", &self.native_patch)
            .finish_non_exhaustive()
    }
}

/// Start every XLSX native edit with the transaction's remaining snapshot
/// ceiling.  Native worksheet actions perform their own capture preflight;
/// passing the remainder here makes that guard cover rich, feature and
/// structural families as well as the coalesced scalar-cell path.
pub(super) fn edit_with_snapshot_limit(
    source: &xlsx::Package,
    maximum: usize,
) -> Result<xlsx::workbook::Edit> {
    source
        .edit_with_options(
            xlsx::workbook::edit::EditOptions::default().with_max_retained_bytes(maximum),
        )
        .map_err(native_error)
}

#[derive(Debug, Clone)]
enum CellOperation {
    Set(xlsx::cell::Content),
    Clear,
    /// Rich-text runs, written as the cell's inline rich string (record
    /// 0338).
    Rich(xlsx::RichText),
    /// `type` without a value: each target's stored value is converted
    /// (record 0338).
    Retype(CellType),
}

/// A cell `type` (record 0338).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum CellType {
    String,
    Number,
    Boolean,
    Date,
    Error,
    RichText,
}

/// Validated authored assignments; absent and explicit clear stay distinct.
#[derive(Debug, Default)]
struct CellAssignments {
    payload: Option<CellOperation>,
    cell_type: Option<CellType>,
    runs: Option<xlsx::RichText>,
    bold: Option<Option<bool>>,
    italic: Option<Option<bool>>,
    strike: Option<Option<bool>>,
    outline: Option<Option<bool>>,
    shadow: Option<Option<bool>>,
    condense: Option<Option<bool>>,
    extend: Option<Option<bool>>,
    underline: Option<Option<xlsx::style::Underline>>,
    font_name: Option<Option<Box<str>>>,
    font_size: Option<Option<f64>>,
    script: Option<Option<xlsx::style::stylesheet::font::Script>>,
    font_color: Option<Option<xlsx::style::stylesheet::font::FontColor>>,
    font_charset: Option<Option<u32>>,
    font_family: Option<Option<u32>>,
    font_scheme: Option<Option<xlsx::style::Scheme>>,
    alignment_horizontal: Option<Option<xlsx::style::stylesheet::alignment::Horizontal>>,
    alignment_vertical: Option<Option<xlsx::style::stylesheet::alignment::Vertical>>,
    alignment_wrap_text: Option<Option<bool>>,
    alignment_text_rotation: Option<Option<xlsx::style::stylesheet::alignment::Rotation>>,
    alignment_indent: Option<Option<xlsx::style::stylesheet::alignment::Indent>>,
    alignment_shrink_to_fit: Option<Option<bool>>,
    alignment_reading_order: Option<Option<xlsx::style::stylesheet::alignment::Reading>>,
    alignment_relative_indent: Option<Option<i32>>,
    alignment_justify_last_line: Option<Option<bool>>,
    protection_locked: Option<Option<bool>>,
    protection_hidden: Option<Option<bool>>,
    quote_prefix: Option<Option<bool>>,
    number_format: Option<Option<Box<str>>>,
    fill: Option<Option<xlsx::style::stylesheet::border::Color>>,
    /// A complete native gradient replacement admitted from `fill` text.
    /// The replacement owns one Arc-backed native fill; target clones share
    /// it rather than cloning the stop collection.
    fill_gradient: Option<xlsx::style::FillChange>,
    fill_gradient_bytes: usize,
    fill_pattern: Option<Option<xlsx::style::cell_format::Pattern>>,
    fill_bg: Option<Option<xlsx::style::stylesheet::border::Color>>,
    border_all: Option<Option<xlsx::style::cell_format::Line>>,
    border_all_color: Option<Option<xlsx::style::stylesheet::border::Color>>,
    border_left: Option<Option<xlsx::style::cell_format::Line>>,
    border_left_color: Option<Option<xlsx::style::stylesheet::border::Color>>,
    border_right: Option<Option<xlsx::style::cell_format::Line>>,
    border_right_color: Option<Option<xlsx::style::stylesheet::border::Color>>,
    border_top: Option<Option<xlsx::style::cell_format::Line>>,
    border_top_color: Option<Option<xlsx::style::stylesheet::border::Color>>,
    border_bottom: Option<Option<xlsx::style::cell_format::Line>>,
    border_bottom_color: Option<Option<xlsx::style::stylesheet::border::Color>>,
    border_diagonal: Option<Option<xlsx::style::cell_format::Line>>,
    border_diagonal_color: Option<Option<xlsx::style::stylesheet::border::Color>>,
    border_diagonal_up: Option<Option<bool>>,
    border_diagonal_down: Option<Option<bool>>,
}

const WRITE_ENTRY_OVERHEAD: usize = 64;

impl CellAssignments {
    fn has_border(&self) -> bool {
        self.border_all.is_some()
            || self.border_all_color.is_some()
            || self.border_left.is_some()
            || self.border_left_color.is_some()
            || self.border_right.is_some()
            || self.border_right_color.is_some()
            || self.border_top.is_some()
            || self.border_top_color.is_some()
            || self.border_bottom.is_some()
            || self.border_bottom_color.is_some()
            || self.border_diagonal.is_some()
            || self.border_diagonal_color.is_some()
            || self.border_diagonal_up.is_some()
            || self.border_diagonal_down.is_some()
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
struct CellTarget {
    sheet: usize,
    address: Address,
}

impl Xlsx {
    /// Assign cell payloads and supported formatting to cells, a range or
    /// scoped matches.
    ///
    /// Admission happens before the native editor is built.  The full target,
    /// property, edit and commit sequence is source-fenced; if the retained
    /// source changes while native work is in progress, that freshness error
    /// wins and no candidate handle is published.
    pub(crate) fn set<'p, 'v>(
        &self,
        target: TargetArg<'p>,
        props: impl IntoProps<'v>,
    ) -> Result<Commit> {
        self.set_with_limits(target, props, &SemanticLimits::default())
    }

    /// Set cells, sheets or dimensions under an explicit expansion budget.
    pub(crate) fn set_with_limits<'p, 'v>(
        &self,
        target: TargetArg<'p>,
        props: impl IntoProps<'v>,
        limits: &SemanticLimits,
    ) -> Result<Commit> {
        if self.signature() == super::Signature::Signed {
            self.check_source_for_replay()?;
        } else {
            self.check_source_for_replay()?;
        }
        let result = self.set_inner(target, props, limits);
        if self.signature() == super::Signature::Signed {
            self.check_source_for_replay()?;
        } else {
            self.check_source_for_replay()?;
        }
        result.and_then(|commit| {
            let snapshot_bytes = commit.snapshot_bytes();
            if snapshot_bytes > limits.max_snapshot_bytes() {
                return Err(snapshot_limit(snapshot_bytes, limits.max_snapshot_bytes()));
            }
            if self.signature() == super::Signature::Signed {
                if commit.is_noop_for(self) {
                    Ok(commit)
                } else {
                    Err(Error::signed(FileFormat::Xlsx))
                }
            } else {
                Ok(commit)
            }
        })
    }

    fn set_inner<'p, 'v>(
        &self,
        target: TargetArg<'p>,
        props: impl IntoProps<'v>,
        limits: &SemanticLimits,
    ) -> Result<Commit> {
        let props = Props::from_props(props)?;
        if let Some(commit) = self.try_set_metadata(&target, &props, limits)? {
            return Ok(commit);
        }
        // Rich-text runs (record 0338).
        if let Some(commit) = self.try_set_run(&target, &props, limits)? {
            return Ok(commit);
        }
        if let Some(commit) = self.try_set_chart(&target, &props, limits)? {
            return Ok(commit);
        }
        if let Some(commit) = self.try_set_feature(&target, &props, limits)? {
            return Ok(commit);
        }
        // Cell hyperlinks (`link`, `tooltip`, `display`) share the cell edit
        // (record 0320).
        let (link, props) = super::feature_write::split_cell_link(&props)?;
        if let Some(link) = link {
            return self.set_cell_link(&target, &props, &link, limits);
        }
        if let Some(commit) = self.try_set_merge(&target, &props, limits)? {
            return Ok(commit);
        }
        if let Some(commit) = self.try_set_dimensions(&target, &props, limits)? {
            return Ok(commit);
        }
        if let Some(commit) = self.try_set_sheet(&target, &props, limits)? {
            return Ok(commit);
        }
        // Even an empty property list must honor signed, retained-source and
        // encrypted edit policy.  Native empty commits return the original
        // unified handle below, preserving exact source bytes on a no-op.
        let source = if self.signature() == super::Signature::Signed {
            self.editor_for_write_with_snapshot_limit(limits.max_snapshot_bytes())?
        } else {
            self.editor_with_snapshot_limit(limits.max_snapshot_bytes())?
        };
        // Promote/admit the native owner before resolving cell state and
        // cloning rich/style payloads in `PreparedCells`.
        let cells = PreparedCells::new(self, &target, &props, limits)?;
        let mut edit = edit_with_snapshot_limit(&source, limits.max_snapshot_bytes())?;
        cells.apply(&mut edit)?;
        let native = edit.commit().map_err(native_error)?;
        Commit::from_native_with_limit(self, source, native, limits.max_snapshot_bytes())
    }
}

impl Xlsx {
    /// A cell `set` carrying hyperlink properties: the cell assignments, if
    /// any, and the hyperlink change in one native edit.
    fn set_cell_link(
        &self,
        target: &TargetArg<'_>,
        props: &Props<'_>,
        link: &super::feature_write::CellLink,
        limits: &SemanticLimits,
    ) -> Result<Commit> {
        let node = match target.path() {
            Some(path) => crate::path::resolve::get(&Nav::new(self), path)?,
            None => {
                return Err(Error::new(
                    Code::UnsupportedCapability,
                    "cell hyperlinks are set through one cell path",
                )
                .with_property("link")
                .with_format(FileFormat::Xlsx));
            },
        };
        let Some((sheet, at)) = node.write_cell_target() else {
            return Err(Error::new(
                Code::UnsupportedCapability,
                "cell hyperlinks are set on one cell",
            )
            .with_property("link")
            .with_format(FileFormat::Xlsx));
        };
        let source = if self.signature() == super::Signature::Signed {
            self.editor_for_write_with_snapshot_limit(limits.max_snapshot_bytes())?
        } else {
            self.editor_with_snapshot_limit(limits.max_snapshot_bytes())?
        };
        let cells = if props.is_empty() {
            None
        } else {
            Some(PreparedCells::new(self, target, props, limits)?)
        };
        let mut edit = edit_with_snapshot_limit(&source, limits.max_snapshot_bytes())?;
        if let Some(cells) = &cells {
            cells.apply(&mut edit)?;
        }
        self.stage_cell_link(&mut edit, sheet, at, link)?;
        let native = edit.commit().map_err(native_error)?;
        Commit::from_native_with_limit(self, source, native, limits.max_snapshot_bytes())
    }
}

/// Admitted cell assignments that can share a native edit with range actions.
pub(super) struct PreparedCells {
    operation: CellAssignments,
    targets: Vec<CellTarget>,
    charge: usize,
    fill_changes: Vec<xlsx::style::FillChange>,
    border_change: xlsx::style::BorderChange,
}

impl PreparedCells {
    pub(super) fn new(
        file: &Xlsx,
        target: &TargetArg<'_>,
        props: &Props<'_>,
        limits: &SemanticLimits,
    ) -> Result<Self> {
        let operation = admit_properties_with_limits(props, limits)?;
        let per_target = per_target_charge(&operation, limits)?;
        let targets = admit_targets(file, target, per_target, limits)?;
        let charge = targets.len().saturating_mul(per_target);
        let fill_changes = prepare_fill_changes(file, &targets, &operation)?;
        let border_change = prepare_border_change(file, &targets, &operation)?;
        Ok(Self {
            operation,
            targets,
            charge,
            fill_changes,
            border_change,
        })
    }

    pub(super) fn target_count(&self) -> usize {
        self.targets.len()
    }
    pub(super) fn staged_bytes(&self) -> usize {
        self.charge
    }
    pub(super) fn apply(&self, edit: &mut xlsx::workbook::Edit) -> Result<()> {
        apply_operation(
            edit,
            &self.targets,
            &self.operation,
            &self.fill_changes,
            &self.border_change,
        )
    }
}

fn admit_targets(
    file: &Xlsx,
    target: &TargetArg<'_>,
    per_target_charge: usize,
    limits: &SemanticLimits,
) -> Result<Vec<CellTarget>> {
    // Resolve only the syntactic A1 range before constructing `Nav`.  The
    // normal range resolver pins a worksheet grid for the returned node; a
    // caller-selected staged-byte limit must reject a huge direct range
    // before that read-side allocation occurs.
    preflight_direct_range(target, per_target_charge, limits)?;
    let nav = Nav::new(file);
    let mut targets = Vec::new();
    let mut seen = HashSet::new();
    let mut matched = false;
    if let Some(path) = target.path() {
        let node = crate::path::resolve::get(&nav, path)?;
        matched = true;
        append_node(&mut targets, &mut seen, node, per_target_charge, limits)?;
    } else if let Some(query) = target.query() {
        let walk = crate::query::run(nav, query)?;
        // A mutation must reject an unsupported predicate. Treating it as an
        // absent property could silently widen a cell write.
        if !walk.plan().diagnostics().is_empty() {
            return Err(Error::new(
                Code::InvalidSelector,
                "the mutation selector contains an unsupported predicate",
            )
            .with_format(FileFormat::Xlsx));
        }
        // Consume matches as they arrive; only compact sheet/address pairs
        // survive. Deduplication is performed before the caller-selected
        // expansion limit, so a selector union cannot count one cell twice.
        for node in walk {
            matched = true;
            append_node(&mut targets, &mut seen, node?, per_target_charge, limits)?;
        }
    } else {
        return Err(
            Error::new(Code::InvalidArgument, "the mutation target is empty")
                .with_format(FileFormat::Xlsx),
        );
    };

    file.check_source_for_replay()?;
    if !matched {
        return Err(Error::new(
            Code::ZeroMatches,
            "the mutation selector matched no element",
        )
        .with_format(FileFormat::Xlsx));
    }
    Ok(targets)
}

fn preflight_direct_range(
    target: &TargetArg<'_>,
    per_target_charge: usize,
    limits: &SemanticLimits,
) -> Result<()> {
    let Some(path) = target.path() else {
        return Ok(());
    };
    let Some(step) = path.path().steps().last() else {
        return Ok(());
    };
    let Some(cells) = step.cell_range() else {
        return Ok(());
    };
    let rect = super::nav::rect(cells)?;
    check_range_area(range_cells(rect), per_target_charge, limits)
}

fn range_cells(rect: Rect) -> usize {
    u64::from(rect.rows())
        .checked_mul(u64::from(rect.columns()))
        .and_then(|cells| usize::try_from(cells).ok())
        .unwrap_or(usize::MAX)
}

fn append_node(
    targets: &mut Vec<CellTarget>,
    seen: &mut HashSet<CellTarget>,
    node: Node<'_>,
    per_target_charge: usize,
    limits: &SemanticLimits,
) -> Result<()> {
    if let Some((sheet, address)) = node.write_cell_target() {
        append_cell_target(targets, seen, sheet, address, per_target_charge, limits)
    } else if let Some((sheet, rect)) = node.write_range_target() {
        append_range_targets(targets, seen, sheet, rect, per_target_charge, limits)
    } else {
        Err(Error::new(
            Code::UnsupportedElement,
            "the XLSX write target must be a cell or range",
        )
        .with_format(FileFormat::Xlsx))
    }
}

fn append_cell_target(
    targets: &mut Vec<CellTarget>,
    seen: &mut HashSet<CellTarget>,
    sheet: usize,
    address: Address,
    per_target_charge: usize,
    limits: &SemanticLimits,
) -> Result<()> {
    let target = CellTarget { sheet, address };
    if seen.contains(&target) {
        return Ok(());
    }
    ensure_target_capacity(targets, seen, 1, per_target_charge, limits)?;
    seen.insert(target);
    targets.push(target);
    Ok(())
}

fn append_range_targets(
    targets: &mut Vec<CellTarget>,
    seen: &mut HashSet<CellTarget>,
    sheet: usize,
    rect: Rect,
    per_target_charge: usize,
    limits: &SemanticLimits,
) -> Result<()> {
    // Every rectangle contains this many distinct coordinates even if it
    // overlaps earlier matches. Refuse oversized areas before allocation.
    let cells = range_cells(rect);
    check_range_area(cells, per_target_charge, limits)?;
    let start = rect.start();
    let (end_row, end_column) = rect.end();
    for row in start.row().get()..end_row {
        for column in start.column().get()..end_column {
            let address = Address::at(row, column)
                .map_err(xlsx::Error::from)
                .map_err(|error| native_error(error).caused_by_argument(Code::OutOfRange))?;
            append_cell_target(targets, seen, sheet, address, per_target_charge, limits)?;
        }
    }
    Ok(())
}

fn ensure_target_capacity(
    targets: &mut Vec<CellTarget>,
    seen: &mut HashSet<CellTarget>,
    additional: usize,
    per_target_charge: usize,
    limits: &SemanticLimits,
) -> Result<()> {
    let total = targets
        .len()
        .checked_add(additional)
        .ok_or_else(|| expanded_target_limit(usize::MAX, limits.max_expanded_targets()))?;
    if total > limits.max_expanded_targets() {
        return Err(expanded_target_limit(total, limits.max_expanded_targets()));
    }
    check_staged_growth(targets.len(), additional, per_target_charge, limits)?;
    targets
        .try_reserve(additional)
        .map_err(|source| allocation("XLSX cell write targets", source))?;
    seen.try_reserve(additional)
        .map_err(|source| allocation("XLSX cell write target identities", source))?;
    Ok(())
}

fn expanded_target_limit(observed: usize, maximum: usize) -> Error {
    Error::from(TransactionError::Limit {
        kind: LimitKind::ExpandedTargets,
        observed,
        maximum,
    })
    .with_format(FileFormat::Xlsx)
}

fn snapshot_limit(observed: usize, maximum: usize) -> Error {
    Error::from(TransactionError::Limit {
        kind: LimitKind::SnapshotBytes,
        observed,
        maximum,
    })
    .with_format(FileFormat::Xlsx)
}

fn per_target_charge(operation: &CellAssignments, limits: &SemanticLimits) -> Result<usize> {
    operation
        .payload
        .as_ref()
        .map(operation_payload_bytes)
        .unwrap_or(0)
        .checked_add(
            operation
                .font_name
                .as_ref()
                .and_then(Option::as_ref)
                .map_or(0, |value| value.len()),
        )
        .and_then(|bytes| {
            bytes.checked_add(
                operation
                    .number_format
                    .as_ref()
                    .and_then(Option::as_ref)
                    .map_or(0, |value| value.len()),
            )
        })
        .and_then(|bytes| {
            let has_fill = operation.fill.is_some()
                || operation.fill_gradient.is_some()
                || operation.fill_pattern.is_some()
                || operation.fill_bg.is_some();
            let fill_bytes = if has_fill {
                size_of::<xlsx::style::FillChange>()
            } else {
                0
            };
            let gradient_bytes = operation.fill_gradient_bytes;
            bytes
                .checked_add(fill_bytes)
                .and_then(|bytes| bytes.checked_add(gradient_bytes))
        })
        .and_then(|bytes| {
            let has_border = operation.has_border();
            let border_bytes = if has_border {
                size_of::<xlsx::style::BorderChange>()
            } else {
                0
            };
            bytes.checked_add(border_bytes)
        })
        .and_then(|bytes| bytes.checked_add(WRITE_ENTRY_OVERHEAD))
        .ok_or_else(|| staged_bytes_limit(usize::MAX, limits.max_staged_bytes()))
}

fn gradient_retained_bytes(gradient: &xlsx::style::cell_format::GradientFill) -> usize {
    gradient_retained_bytes_for_stops(gradient.stops().len())
}

fn gradient_retained_bytes_for_stops(stops: usize) -> usize {
    size_of::<xlsx::style::cell_format::GradientFill>()
        .saturating_add(stops.saturating_mul(size_of::<xlsx::style::cell_format::GradientStop>()))
}

fn check_gradient_staging(
    limits: &SemanticLimits,
    stops: usize,
    separator_count: usize,
    key: &'static str,
) -> Result<()> {
    if stops < 2 {
        return Err(invalid_property(
            key,
            "gradient fill must contain at least two colors",
        ));
    }
    xlsx::style::cell_format::GradientFill::validate_stop_count(stops).map_err(native_error)?;
    // Parse only after the complete stop collection and the temporary
    // separator scan fit the caller's budget. The raw spelling is checked
    // separately by `check_format_text_staging`.
    let observed = gradient_retained_bytes_for_stops(stops)
        .checked_add(size_of::<xlsx::style::FillChange>())
        .and_then(|bytes| bytes.checked_add(WRITE_ENTRY_OVERHEAD))
        .and_then(|bytes| bytes.checked_add(separator_count.saturating_mul(size_of::<usize>())))
        .ok_or_else(|| staged_bytes_limit(usize::MAX, limits.max_staged_bytes()))?;
    if observed > limits.max_staged_bytes() {
        return Err(staged_bytes_limit(observed, limits.max_staged_bytes()));
    }
    Ok(())
}

fn check_range_area(cells: usize, per_target_charge: usize, limits: &SemanticLimits) -> Result<()> {
    if cells > limits.max_expanded_targets() {
        return Err(expanded_target_limit(cells, limits.max_expanded_targets()));
    }
    let staged = cells.saturating_mul(per_target_charge);
    if staged > limits.max_staged_bytes() {
        return Err(staged_bytes_limit(staged, limits.max_staged_bytes()));
    }
    Ok(())
}

fn check_staged_growth(
    existing: usize,
    additional: usize,
    per_target_charge: usize,
    limits: &SemanticLimits,
) -> Result<()> {
    let total = existing.saturating_add(additional);
    let staged = total.saturating_mul(per_target_charge);
    if staged > limits.max_staged_bytes() {
        return Err(staged_bytes_limit(staged, limits.max_staged_bytes()));
    }
    Ok(())
}

fn operation_payload_bytes(operation: &CellOperation) -> usize {
    match operation {
        CellOperation::Clear => 0,
        CellOperation::Set(xlsx::cell::Content::Value(value)) => match value {
            xlsx::cell::Value::Bool(_) => 0,
            xlsx::cell::Value::Number(number) => number.as_str().len(),
            xlsx::cell::Value::Text(text) => text.as_str().len(),
            xlsx::cell::Value::Date(date) => date.as_str().len(),
            _ => 0,
        },
        CellOperation::Set(xlsx::cell::Content::Formula(formula)) => formula.text().len(),
        CellOperation::Rich(rich) => rich.retained_bytes(),
        // Retyping first projects the existing cell state before it can lower
        // the converted value. Reserve a semantic cell slot before that
        // source read; the concrete setter replaces this coarse charge with
        // the final authored payload once the conversion is known.
        CellOperation::Retype(_) => size_of::<xlsx::cell::Content>(),
        _ => 0,
    }
}

fn check_format_text_staging(bytes: usize, limits: &SemanticLimits) -> Result<()> {
    let observed = bytes
        .checked_add(WRITE_ENTRY_OVERHEAD)
        .ok_or_else(|| staged_bytes_limit(usize::MAX, limits.max_staged_bytes()))?;
    if observed > limits.max_staged_bytes() {
        return Err(staged_bytes_limit(observed, limits.max_staged_bytes()));
    }
    Ok(())
}

fn number_format_alias(value: &str) -> Option<&'static str> {
    let value = value.trim();
    if value.eq_ignore_ascii_case("general") {
        Some("General")
    } else if value.eq_ignore_ascii_case("number") || value.eq_ignore_ascii_case("comma") {
        Some("#,##0.00")
    } else if value.eq_ignore_ascii_case("currency") {
        Some("\"$\"#,##0.00")
    } else if value.eq_ignore_ascii_case("accounting") {
        Some("_(\"$\"* #,##0.00_);_(\"$\"* \\(#,##0.00\\);_(\"$\"* \"-\"??_);_(@_)")
    } else if value.eq_ignore_ascii_case("percent") || value.eq_ignore_ascii_case("percentage") {
        Some("0.00%")
    } else if value.eq_ignore_ascii_case("scientific") {
        Some("0.00E+00")
    } else if value.eq_ignore_ascii_case("text") {
        Some("@")
    } else if value.eq_ignore_ascii_case("date") {
        Some("yyyy-mm-dd")
    } else if value.eq_ignore_ascii_case("time") {
        Some("h:mm:ss")
    } else if value.eq_ignore_ascii_case("datetime") {
        Some("yyyy-mm-dd h:mm:ss")
    } else {
        None
    }
}

fn admit_number_format(value: &str, key: &str, limits: &SemanticLimits) -> Result<Box<str>> {
    let normalized = number_format_alias(value).unwrap_or(value);
    check_format_text_staging(normalized.len(), limits)?;
    xlsx::style::stylesheet::NumberFormat::validate_code(normalized).map_err(|_| {
        invalid_property(key, "number format is invalid for SpreadsheetML authoring")
    })?;
    Ok(normalized.to_owned().into_boxed_str())
}

fn normalize_fill_pattern_value(value: &PropValue<'_>) -> Option<PropValue<'static>> {
    let PropValue::Text(value) = value else {
        return None;
    };
    let canonical = canonical_fill_pattern(value)?;
    Some(PropValue::Text(Cow::Borrowed(canonical)))
}

fn canonical_fill_pattern(value: &str) -> Option<&'static str> {
    let value = value.trim();
    for pattern in xlsx::style::FillPattern::ALL {
        if compact_fill_pattern_eq(value, pattern.as_str()) {
            return Some(pattern.as_str());
        }
    }
    if compact_fill_pattern_eq(value, "gray25") {
        Some("lightGray")
    } else if compact_fill_pattern_eq(value, "gray50") {
        Some("mediumGray")
    } else if compact_fill_pattern_eq(value, "gray75") {
        Some("darkGray")
    } else {
        None
    }
}

fn compact_fill_pattern_eq(left: &str, right: &str) -> bool {
    let mut left = left
        .chars()
        .filter(|character| *character != '-' && *character != ' ')
        .map(|character| character.to_ascii_lowercase());
    let mut right = right
        .chars()
        .map(|character| character.to_ascii_lowercase());
    loop {
        match (left.next(), right.next()) {
            (Some(left), Some(right)) if left == right => {},
            (None, None) => return true,
            _ => return false,
        }
    }
}

fn staged_bytes_limit(observed: usize, maximum: usize) -> Error {
    Error::from(TransactionError::Limit {
        kind: LimitKind::StagedBytes,
        observed,
        maximum,
    })
    .with_format(FileFormat::Xlsx)
}

fn allocation(resource: &'static str, source: TryReserveError) -> Error {
    Error::from(xlsx::Error::Allocation { resource, source }).with_format(FileFormat::Xlsx)
}

fn admit_properties_with_limits(
    props: &Props<'_>,
    limits: &SemanticLimits,
) -> Result<CellAssignments> {
    admit_properties_inner(props, limits)
}

#[cfg(test)]
fn admit_properties(props: &Props<'_>) -> Result<CellAssignments> {
    admit_properties_with_limits(props, &SemanticLimits::default())
}

fn admit_properties_inner(props: &Props<'_>, limits: &SemanticLimits) -> Result<CellAssignments> {
    let mut operation = CellAssignments::default();
    let mut payload_setter = None;
    let mut script_setter = None;
    for (key, value) in props.iter() {
        let (_, entry) = CELL_WRITE_TABLE
            .resolve(key.as_str())
            .map_err(|error| property_error(error, key))?;
        let entry = *entry;
        let normalized_fill_pattern = (entry.spec.name == "fillPattern")
            .then(|| normalize_fill_pattern_value(value))
            .flatten();
        let prepared_value = normalized_fill_pattern.as_ref().unwrap_or(value);
        entry
            .prepare(Operation::Set, prepared_value, |current, value| {
                match current {
                    CellSetter::Bold
                    | CellSetter::Italic
                    | CellSetter::Strike
                    | CellSetter::FontOutline
                    | CellSetter::FontShadow
                    | CellSetter::FontCondense
                    | CellSetter::FontExtend => {
                        let font = match current {
                            CellSetter::Bold => &mut operation.bold,
                            CellSetter::Italic => &mut operation.italic,
                            CellSetter::Strike => &mut operation.strike,
                            CellSetter::FontOutline => &mut operation.outline,
                            CellSetter::FontShadow => &mut operation.shadow,
                            CellSetter::FontCondense => &mut operation.condense,
                            CellSetter::FontExtend => &mut operation.extend,
                            _ => unreachable!(),
                        };
                        if font.is_some() {
                            return Err(Error::new(
                                Code::InvalidArgument,
                                "the same XLSX cell property was supplied more than once",
                            ));
                        }
                        *font = Some(match value {
                            None => None,
                            Some(Scalar::Bool(value)) => Some(value),
                            _ => {
                                return Err(invalid_property(
                                    entry.spec.name,
                                    "font flags must be boolean or null",
                                ));
                            },
                        });
                        return Ok(());
                    },
                    CellSetter::Underline => {
                        if operation.underline.is_some() {
                            return Err(Error::new(
                                Code::InvalidArgument,
                                "the same XLSX cell property was supplied more than once",
                            ));
                        }
                        operation.underline = Some(match value {
                            None => None,
                            Some(Scalar::Token(value)) => Some(match value {
                                "single" => xlsx::style::Underline::Single,
                                "double" => xlsx::style::Underline::Double,
                                "singleAccounting" => xlsx::style::Underline::SingleAccounting,
                                "doubleAccounting" => xlsx::style::Underline::DoubleAccounting,
                                "none" => xlsx::style::Underline::None,
                                _ => {
                                    return Err(invalid_property(
                                        entry.spec.name,
                                        "underline has an unsupported token",
                                    ));
                                },
                            }),
                            _ => {
                                return Err(invalid_property(
                                    entry.spec.name,
                                    "underline must be a token or null",
                                ));
                            },
                        });
                        return Ok(());
                    },
                    CellSetter::FontName => {
                        if operation.font_name.is_some() {
                            return Err(Error::new(
                                Code::InvalidArgument,
                                "the same XLSX cell property was supplied more than once",
                            ));
                        }
                        operation.font_name = Some(match value {
                            None => None,
                            Some(Scalar::Text(value)) => {
                                check_format_text_staging(value.len(), limits)?;
                                validate_font_name(value, entry.spec.name)?;
                                Some(value.to_owned().into_boxed_str())
                            },
                            _ => {
                                return Err(invalid_property(
                                    entry.spec.name,
                                    "font name must be non-empty text or null",
                                ));
                            },
                        });
                        return Ok(());
                    },
                    CellSetter::FontSize => {
                        if operation.font_size.is_some() {
                            return Err(Error::new(
                                Code::InvalidArgument,
                                "the same XLSX cell property was supplied more than once",
                            ));
                        }
                        operation.font_size = Some(match value {
                            None => None,
                            Some(Scalar::Length(value)) => Some(value.points()),
                            _ => {
                                return Err(invalid_property(
                                    entry.spec.name,
                                    "font size must be a length or null",
                                ));
                            },
                        });
                        return Ok(());
                    },
                    CellSetter::Superscript | CellSetter::Subscript => {
                        if let Some(previous) = script_setter {
                            if previous == current {
                                return Err(Error::new(
                                    Code::InvalidArgument,
                                    "the same XLSX cell property was supplied more than once",
                                ));
                            }
                            return Err(Error::new(
                                Code::InvalidArgument,
                                "superscript and subscript cannot be supplied together",
                            ));
                        }
                        script_setter = Some(current);
                        let script = match current {
                            CellSetter::Superscript => {
                                xlsx::style::stylesheet::font::Script::Superscript
                            },
                            CellSetter::Subscript => {
                                xlsx::style::stylesheet::font::Script::Subscript
                            },
                            _ => unreachable!(),
                        };
                        operation.script = Some(match value {
                            None => None,
                            Some(Scalar::Bool(value)) => Some(if value {
                                script
                            } else {
                                xlsx::style::stylesheet::font::Script::Baseline
                            }),
                            _ => {
                                return Err(invalid_property(
                                    entry.spec.name,
                                    "script properties must be boolean or null",
                                ));
                            },
                        });
                        return Ok(());
                    },
                    CellSetter::FontColor => {
                        if operation.font_color.is_some() {
                            return Err(Error::new(
                                Code::InvalidArgument,
                                "the same XLSX cell property was supplied more than once",
                            ));
                        }
                        operation.font_color = Some(match value {
                            None => None,
                            Some(Scalar::Color(value)) => {
                                Some(native_font_color(value, entry.spec.name)?)
                            },
                            _ => {
                                return Err(invalid_property(
                                    entry.spec.name,
                                    "font color must be a SpreadsheetML color or null",
                                ));
                            },
                        });
                        return Ok(());
                    },
                    CellSetter::FontCharset | CellSetter::FontFamily => {
                        let font = match current {
                            CellSetter::FontCharset => &mut operation.font_charset,
                            CellSetter::FontFamily => &mut operation.font_family,
                            _ => unreachable!(),
                        };
                        if font.is_some() {
                            return Err(Error::new(
                                Code::InvalidArgument,
                                "the same XLSX cell property was supplied more than once",
                            ));
                        }
                        *font = Some(match value {
                            None => None,
                            Some(Scalar::Int(value)) => {
                                Some(u32::try_from(value).map_err(|_| {
                                    invalid_property(
                                        entry.spec.name,
                                        "font integer is outside its supported domain",
                                    )
                                })?)
                            },
                            _ => {
                                return Err(invalid_property(
                                    entry.spec.name,
                                    "font integer must be an integer or null",
                                ));
                            },
                        });
                        return Ok(());
                    },
                    CellSetter::FontScheme => {
                        if operation.font_scheme.is_some() {
                            return Err(Error::new(
                                Code::InvalidArgument,
                                "the same XLSX cell property was supplied more than once",
                            ));
                        }
                        operation.font_scheme = Some(match value {
                            None => None,
                            Some(Scalar::Token(value)) => Some(match value {
                                "none" => xlsx::style::Scheme::None,
                                "major" => xlsx::style::Scheme::Major,
                                "minor" => xlsx::style::Scheme::Minor,
                                _ => {
                                    return Err(invalid_property(
                                        entry.spec.name,
                                        "font scheme has an unsupported token",
                                    ));
                                },
                            }),
                            _ => {
                                return Err(invalid_property(
                                    entry.spec.name,
                                    "font scheme must be a token or null",
                                ));
                            },
                        });
                        return Ok(());
                    },
                    CellSetter::AlignmentWrapText | CellSetter::AlignmentShrinkToFit => {
                        let alignment = match current {
                            CellSetter::AlignmentWrapText => &mut operation.alignment_wrap_text,
                            CellSetter::AlignmentShrinkToFit => {
                                &mut operation.alignment_shrink_to_fit
                            },
                            _ => unreachable!(),
                        };
                        if alignment.is_some() {
                            return Err(Error::new(
                                Code::InvalidArgument,
                                "the same XLSX cell property was supplied more than once",
                            ));
                        }
                        *alignment = Some(match value {
                            None => None,
                            Some(Scalar::Bool(value)) => Some(value),
                            _ => {
                                return Err(invalid_property(
                                    entry.spec.name,
                                    "alignment flags must be boolean or null",
                                ));
                            },
                        });
                        return Ok(());
                    },
                    CellSetter::AlignmentHorizontal => {
                        if operation.alignment_horizontal.is_some() {
                            return Err(Error::new(
                                Code::InvalidArgument,
                                "the same XLSX cell property was supplied more than once",
                            ));
                        }
                        operation.alignment_horizontal = Some(match value {
                            None => None,
                            Some(Scalar::Token(value)) => Some(match value {
                                "general" => {
                                    xlsx::style::stylesheet::alignment::Horizontal::General
                                },
                                "left" => xlsx::style::stylesheet::alignment::Horizontal::Left,
                                "center" => xlsx::style::stylesheet::alignment::Horizontal::Center,
                                "right" => xlsx::style::stylesheet::alignment::Horizontal::Right,
                                "fill" => xlsx::style::stylesheet::alignment::Horizontal::Fill,
                                "justify" => {
                                    xlsx::style::stylesheet::alignment::Horizontal::Justify
                                },
                                "centerContinuous" => {
                                    xlsx::style::stylesheet::alignment::Horizontal::CenterContinuous
                                },
                                "distributed" => {
                                    xlsx::style::stylesheet::alignment::Horizontal::Distributed
                                },
                                _ => {
                                    return Err(invalid_property(
                                        entry.spec.name,
                                        "horizontal alignment has an unsupported token",
                                    ));
                                },
                            }),
                            _ => {
                                return Err(invalid_property(
                                    entry.spec.name,
                                    "horizontal alignment must be a token or null",
                                ));
                            },
                        });
                        return Ok(());
                    },
                    CellSetter::AlignmentVertical => {
                        if operation.alignment_vertical.is_some() {
                            return Err(Error::new(
                                Code::InvalidArgument,
                                "the same XLSX cell property was supplied more than once",
                            ));
                        }
                        operation.alignment_vertical = Some(match value {
                            None => None,
                            Some(Scalar::Token(value)) => Some(match value {
                                "top" => xlsx::style::stylesheet::alignment::Vertical::Top,
                                "center" => xlsx::style::stylesheet::alignment::Vertical::Center,
                                "bottom" => xlsx::style::stylesheet::alignment::Vertical::Bottom,
                                "justify" => xlsx::style::stylesheet::alignment::Vertical::Justify,
                                "distributed" => {
                                    xlsx::style::stylesheet::alignment::Vertical::Distributed
                                },
                                _ => {
                                    return Err(invalid_property(
                                        entry.spec.name,
                                        "vertical alignment has an unsupported token",
                                    ));
                                },
                            }),
                            _ => {
                                return Err(invalid_property(
                                    entry.spec.name,
                                    "vertical alignment must be a token or null",
                                ));
                            },
                        });
                        return Ok(());
                    },
                    CellSetter::AlignmentReadingOrder => {
                        if operation.alignment_reading_order.is_some() {
                            return Err(Error::new(
                                Code::InvalidArgument,
                                "the same XLSX cell property was supplied more than once",
                            ));
                        }
                        operation.alignment_reading_order = Some(match value {
                            None => None,
                            Some(Scalar::Token(value)) => match value {
                                // OfficeCLI treats the schema-default context
                                // value as an omitted attribute.  Preserve
                                // that canonical no-op distinction instead of
                                // authoring an explicit readingOrder="0".
                                "context" => None,
                                "ltr" => {
                                    Some(xlsx::style::stylesheet::alignment::Reading::LeftToRight)
                                },
                                "rtl" => {
                                    Some(xlsx::style::stylesheet::alignment::Reading::RightToLeft)
                                },
                                _ => {
                                    return Err(invalid_property(
                                        entry.spec.name,
                                        "reading order has an unsupported token",
                                    ));
                                },
                            },
                            _ => {
                                return Err(invalid_property(
                                    entry.spec.name,
                                    "reading order must be a token or null",
                                ));
                            },
                        });
                        return Ok(());
                    },
                    CellSetter::AlignmentTextRotation => {
                        if operation.alignment_text_rotation.is_some() {
                            return Err(Error::new(
                                Code::InvalidArgument,
                                "the same XLSX cell property was supplied more than once",
                            ));
                        }
                        operation.alignment_text_rotation = Some(match value {
                            None => None,
                            Some(Scalar::Int(value)) => {
                                let value = u32::try_from(value).map_err(|_| {
                                    invalid_property(
                                        entry.spec.name,
                                        "text rotation is outside its supported domain",
                                    )
                                })?;
                                if value == 254 {
                                    return Err(invalid_property(
                                        entry.spec.name,
                                        "contextual text rotation is not supported",
                                    ));
                                }
                                Some(
                                    xlsx::style::stylesheet::alignment::Rotation::try_from(value)
                                        .map_err(|_| {
                                        invalid_property(
                                            entry.spec.name,
                                            "text rotation must be 0..=180 or 255",
                                        )
                                    })?,
                                )
                            },
                            _ => {
                                return Err(invalid_property(
                                    entry.spec.name,
                                    "text rotation must be an integer or null",
                                ));
                            },
                        });
                        return Ok(());
                    },
                    CellSetter::AlignmentIndent => {
                        if operation.alignment_indent.is_some() {
                            return Err(Error::new(
                                Code::InvalidArgument,
                                "the same XLSX cell property was supplied more than once",
                            ));
                        }
                        operation.alignment_indent = Some(match value {
                            None => None,
                            Some(Scalar::Int(value)) => {
                                let value = u32::try_from(value).map_err(|_| {
                                    invalid_property(
                                        entry.spec.name,
                                        "indent is outside its supported domain",
                                    )
                                })?;
                                Some(
                                    xlsx::style::stylesheet::alignment::Indent::try_from(value)
                                        .map_err(|_| {
                                            invalid_property(
                                                entry.spec.name,
                                                "indent must be between 0 and 255",
                                            )
                                        })?,
                                )
                            },
                            _ => {
                                return Err(invalid_property(
                                    entry.spec.name,
                                    "indent must be an integer or null",
                                ));
                            },
                        });
                        return Ok(());
                    },
                    CellSetter::AlignmentRelativeIndent => {
                        if operation.alignment_relative_indent.is_some() {
                            return Err(Error::new(
                                Code::InvalidArgument,
                                "the same XLSX cell property was supplied more than once",
                            ));
                        }
                        operation.alignment_relative_indent = Some(match value {
                            None => None,
                            Some(Scalar::Int(value)) => {
                                Some(i32::try_from(value).map_err(|_| {
                                    invalid_property(
                                        entry.spec.name,
                                        "relative indent is outside the signed 32-bit domain",
                                    )
                                })?)
                            },
                            _ => {
                                return Err(invalid_property(
                                    entry.spec.name,
                                    "relative indent must be an integer or null",
                                ));
                            },
                        });
                        return Ok(());
                    },
                    CellSetter::AlignmentJustifyLastLine
                    | CellSetter::ProtectionLocked
                    | CellSetter::ProtectionHidden => {
                        let property = match current {
                            CellSetter::AlignmentJustifyLastLine => {
                                &mut operation.alignment_justify_last_line
                            },
                            CellSetter::ProtectionLocked => &mut operation.protection_locked,
                            CellSetter::ProtectionHidden => &mut operation.protection_hidden,
                            _ => unreachable!(),
                        };
                        if property.is_some() {
                            return Err(Error::new(
                                Code::InvalidArgument,
                                "the same XLSX cell property was supplied more than once",
                            ));
                        }
                        *property = Some(match value {
                            None => None,
                            Some(Scalar::Bool(value)) => Some(value),
                            _ => {
                                return Err(invalid_property(
                                    entry.spec.name,
                                    "the property must be boolean or null",
                                ));
                            },
                        });
                        return Ok(());
                    },
                    CellSetter::QuotePrefix => {
                        if operation.quote_prefix.is_some() {
                            return Err(Error::new(
                                Code::InvalidArgument,
                                "the same XLSX cell property was supplied more than once",
                            ));
                        }
                        operation.quote_prefix = Some(match value {
                            None => None,
                            Some(Scalar::Bool(value)) => Some(value),
                            _ => {
                                return Err(invalid_property(
                                    entry.spec.name,
                                    "quote prefix must be boolean or null",
                                ));
                            },
                        });
                        return Ok(());
                    },
                    CellSetter::NumberFormat => {
                        if operation.number_format.is_some() {
                            return Err(Error::new(
                                Code::InvalidArgument,
                                "the same XLSX cell property was supplied more than once",
                            ));
                        }
                        operation.number_format = Some(match value {
                            None => None,
                            Some(Scalar::Text(value)) => {
                                Some(admit_number_format(value, entry.spec.name, limits)?)
                            },
                            _ => {
                                return Err(invalid_property(
                                    entry.spec.name,
                                    "number format must be text or null",
                                ));
                            },
                        });
                        return Ok(());
                    },
                    CellSetter::Fill | CellSetter::FillBg => {
                        if (current == CellSetter::Fill
                            && (operation.fill.is_some() || operation.fill_gradient.is_some()))
                            || (current == CellSetter::FillBg && operation.fill_bg.is_some())
                        {
                            return Err(Error::new(
                                Code::InvalidArgument,
                                "the same XLSX cell property was supplied more than once",
                            ));
                        }
                        match current {
                            CellSetter::Fill => match value {
                                None => operation.fill = Some(None),
                                Some(Scalar::Text(value)) => {
                                    check_format_text_staging(value.len(), limits)?;
                                    match admit_fill_value(value, entry.spec.name, limits)? {
                                        AdmittedFill::Solid(color) => {
                                            operation.fill = Some(Some(color));
                                        },
                                        AdmittedFill::Gradient(gradient) => {
                                            operation.fill_gradient_bytes =
                                                gradient_retained_bytes(&gradient);
                                            operation.fill_gradient = Some(
                                                xlsx::style::FillChange::replace(
                                                    xlsx::style::cell_format::Fill::Gradient(gradient),
                                                ),
                                            );
                                        },
                                    }
                                },
                                _ => {
                                    return Err(invalid_property(
                                        entry.spec.name,
                                        "fill must be a SpreadsheetML color, gradient text, or null",
                                    ));
                                },
                            },
                            CellSetter::FillBg => {
                                operation.fill_bg = Some(match value {
                                    None => None,
                                    Some(Scalar::Color(value)) => {
                                        Some(native_fill_color(value, entry.spec.name)?)
                                    },
                                    _ => {
                                        return Err(invalid_property(
                                            entry.spec.name,
                                            "fill background must be a SpreadsheetML color or null",
                                        ));
                                    },
                                });
                            },
                            _ => unreachable!(),
                        }
                        return Ok(());
                    },
                    CellSetter::FillPattern => {
                        if operation.fill_pattern.is_some() {
                            return Err(Error::new(
                                Code::InvalidArgument,
                                "the same XLSX cell property was supplied more than once",
                            ));
                        }
                        operation.fill_pattern = Some(match value {
                            None => None,
                            Some(Scalar::Token(value)) => Some(
                                xlsx::style::FillPattern::from_token(value).ok_or_else(|| {
                                    invalid_property(
                                        entry.spec.name,
                                        "fill pattern has an unsupported token",
                                    )
                                })?,
                            ),
                            _ => {
                                return Err(invalid_property(
                                    entry.spec.name,
                                    "fill pattern must be a token or null",
                                ));
                            },
                        });
                        return Ok(());
                    },
                    CellSetter::BorderAll
                    | CellSetter::BorderLeft
                    | CellSetter::BorderRight
                    | CellSetter::BorderTop
                    | CellSetter::BorderBottom
                    | CellSetter::BorderDiagonal => {
                        let line = match current {
                            CellSetter::BorderAll => &mut operation.border_all,
                            CellSetter::BorderLeft => &mut operation.border_left,
                            CellSetter::BorderRight => &mut operation.border_right,
                            CellSetter::BorderTop => &mut operation.border_top,
                            CellSetter::BorderBottom => &mut operation.border_bottom,
                            CellSetter::BorderDiagonal => &mut operation.border_diagonal,
                            _ => unreachable!(),
                        };
                        if line.is_some() {
                            return Err(duplicate_property(entry.spec.name));
                        }
                        *line = Some(admit_border_line(value, entry.spec.name)?);
                        return Ok(());
                    },
                    CellSetter::BorderAllColor
                    | CellSetter::BorderLeftColor
                    | CellSetter::BorderRightColor
                    | CellSetter::BorderTopColor
                    | CellSetter::BorderBottomColor
                    | CellSetter::BorderDiagonalColor => {
                        let color = match current {
                            CellSetter::BorderAllColor => &mut operation.border_all_color,
                            CellSetter::BorderLeftColor => &mut operation.border_left_color,
                            CellSetter::BorderRightColor => &mut operation.border_right_color,
                            CellSetter::BorderTopColor => &mut operation.border_top_color,
                            CellSetter::BorderBottomColor => &mut operation.border_bottom_color,
                            CellSetter::BorderDiagonalColor => &mut operation.border_diagonal_color,
                            _ => unreachable!(),
                        };
                        if color.is_some() {
                            return Err(duplicate_property(entry.spec.name));
                        }
                        *color = Some(admit_border_color(value, entry.spec.name)?);
                        return Ok(());
                    },
                    CellSetter::BorderDiagonalUp | CellSetter::BorderDiagonalDown => {
                        let direction = match current {
                            CellSetter::BorderDiagonalUp => &mut operation.border_diagonal_up,
                            CellSetter::BorderDiagonalDown => &mut operation.border_diagonal_down,
                            _ => unreachable!(),
                        };
                        if direction.is_some() {
                            return Err(duplicate_property(entry.spec.name));
                        }
                        *direction = Some(admit_border_bool(value, entry.spec.name)?);
                        return Ok(());
                    },
                    CellSetter::Type => {
                        if operation.cell_type.is_some() {
                            return Err(duplicate_property(entry.spec.name));
                        }
                        operation.cell_type = Some(match value {
                            Some(Scalar::Token("string")) => CellType::String,
                            Some(Scalar::Token("number")) => CellType::Number,
                            Some(Scalar::Token("boolean")) => CellType::Boolean,
                            Some(Scalar::Token("date")) => CellType::Date,
                            Some(Scalar::Token("error")) => CellType::Error,
                            Some(Scalar::Token("richtext")) => CellType::RichText,
                            _ => {
                                return Err(invalid_property(
                                    entry.spec.name,
                                    "the cell type is string, number, boolean, date, error or richtext",
                                ));
                            },
                        });
                        return Ok(());
                    },
                    CellSetter::Runs => {
                        if operation.runs.is_some() {
                            return Err(duplicate_property(entry.spec.name));
                        }
                        let Some(Scalar::Text(text)) = value else {
                            return Err(invalid_property(
                                entry.spec.name,
                                "runs is a JSON array of run objects",
                            ));
                        };
                        check_format_text_staging(text.len(), limits)?;
                        operation.runs = Some(super::runs::parse_runs(text, entry.spec.name)?);
                        return Ok(());
                    },
                    CellSetter::Ref => {
                        return Err(invalid_property(
                            entry.spec.name,
                            "ref names the cell an add creates",
                        ));
                    },
                    _ => {},
                }
                if let Some(previous) = payload_setter {
                    if previous != current {
                        return Err(Error::new(
                            Code::InvalidArgument,
                            "cell properties value, formula and clear conflict",
                        ));
                    }
                    return Err(Error::new(
                        Code::InvalidArgument,
                        "the same XLSX cell property was supplied more than once",
                    ));
                }
                payload_setter = Some(current);
                operation.payload = Some(authored_operation(current, value, entry.spec.name)?);
                Ok(())
            })
            .map_err(|error| property_error(error, key))?;
    }

    if operation.fill_gradient.is_some()
        && (operation.fill_pattern.is_some() || operation.fill_bg.is_some())
    {
        return Err(invalid_property(
            "fill",
            "a gradient fill cannot be combined with fillPattern or fillBg",
        ));
    }

    validate_border_assignments(&operation)?;
    resolve_cell_type(&mut operation)?;

    Ok(operation)
}

/// Combine `type` and `runs` with the payload (record 0338), as OfficeCLI's
/// `Set`/`AddCell` do: `runs` (with or without `type=richtext`) writes the
/// runs; `type=richtext` with a text value writes one unformatted run; any
/// other type converts the value, or each target's stored value when no
/// value is given.
fn resolve_cell_type(operation: &mut CellAssignments) -> Result<()> {
    let conflict = |message: &'static str| {
        Error::new(Code::InvalidArgument, message)
            .with_property("type")
            .with_format(FileFormat::Xlsx)
    };
    if let Some(runs) = operation.runs.take() {
        if !matches!(operation.cell_type, None | Some(CellType::RichText)) {
            return Err(conflict("runs is valid only with type=richtext"));
        }
        if operation.payload.is_some() {
            return Err(conflict("runs replaces the cell value: give one of them"));
        }
        operation.payload = Some(CellOperation::Rich(runs));
        return Ok(());
    }
    let Some(cell_type) = operation.cell_type else {
        return Ok(());
    };
    operation.payload = Some(match operation.payload.take() {
        None if cell_type == CellType::RichText => {
            return Err(Error::new(
                Code::MissingProperty,
                "type=richtext needs runs, or a text value",
            )
            .with_property("runs")
            .with_format(FileFormat::Xlsx));
        },
        None => CellOperation::Retype(cell_type),
        Some(CellOperation::Set(xlsx::cell::Content::Value(value))) => {
            if cell_type == CellType::RichText {
                let text = value_text(&value)
                    .ok_or_else(|| invalid_property("type", "an error value has no runs"))?;
                CellOperation::Rich(xlsx::RichText::plain(text))
            } else {
                CellOperation::Set(xlsx::cell::Content::Value(typed_value(value, cell_type)?))
            }
        },
        Some(CellOperation::Set(xlsx::cell::Content::Formula(_))) => {
            return Err(conflict("type applies to a value, not a formula"));
        },
        Some(_) => return Err(conflict("type applies to a value, not clear")),
    });
    Ok(())
}

/// The text of a value, for a string or rich-text retype.
fn value_text(value: &xlsx::cell::Value) -> Option<String> {
    Some(match value {
        xlsx::cell::Value::Text(text) => text.as_str().to_owned(),
        xlsx::cell::Value::Number(number) => number.as_str().to_owned(),
        xlsx::cell::Value::Bool(value) => if *value { "TRUE" } else { "FALSE" }.to_owned(),
        xlsx::cell::Value::Date(date) => date.as_str().to_owned(),
        _ => return None,
    })
}

/// `value` stored as `cell_type` (OfficeCLI's cell type switch): a refusal
/// when the value has no such form.
fn typed_value(value: xlsx::cell::Value, cell_type: CellType) -> Result<xlsx::cell::Value> {
    use xlsx::cell::{ErrorValue, Value};
    let refused = |message: &'static str| invalid_property("type", message);
    Ok(match (cell_type, value) {
        (CellType::String | CellType::RichText, value) => Value::text(
            value_text(&value).ok_or_else(|| refused("an error value has no text form"))?,
        ),
        (CellType::Number, Value::Number(number)) => Value::Number(number),
        (CellType::Number, Value::Text(text)) => {
            let number = xlsx::cell::Number::new(text.as_str().trim())
                .map_err(|_| refused("the value is not a number"))?;
            Value::Number(number)
        },
        (CellType::Number, _) => return Err(refused("the value is not a number")),
        (CellType::Boolean, Value::Bool(value)) => Value::Bool(value),
        (CellType::Boolean, Value::Text(text)) => {
            match text.as_str().trim().to_ascii_lowercase().as_str() {
                "true" | "yes" | "1" => Value::Bool(true),
                "false" | "no" | "0" => Value::Bool(false),
                _ => return Err(refused("a boolean is true/false, yes/no or 1/0")),
            }
        },
        (CellType::Boolean, Value::Number(number)) => match number.as_str() {
            "1" => Value::Bool(true),
            "0" => Value::Bool(false),
            _ => return Err(refused("a boolean is true/false, yes/no or 1/0")),
        },
        (CellType::Boolean, _) => return Err(refused("a boolean is true/false, yes/no or 1/0")),
        (CellType::Date, Value::Date(date)) => Value::Date(date),
        (CellType::Date, Value::Number(number)) => Value::Number(number),
        (CellType::Date, Value::Text(text)) => Value::date(text.as_str().trim())
            .map_err(|_| refused("a date is an ISO 8601 date or date-time"))?,
        (CellType::Date, _) => return Err(refused("a date is an ISO 8601 date or date-time")),
        (CellType::Error, Value::Error(error)) => Value::Error(error),
        (CellType::Error, Value::Text(text)) => Value::Error(match text.as_str().trim() {
            "#NULL!" => ErrorValue::Null,
            "#DIV/0!" => ErrorValue::DivZero,
            "#VALUE!" => ErrorValue::Value,
            "#REF!" => ErrorValue::Ref,
            "#NAME?" => ErrorValue::Name,
            "#NUM!" => ErrorValue::Num,
            "#N/A" => ErrorValue::NotAvailable,
            "#GETTING_DATA" => ErrorValue::GettingData,
            _ => {
                return Err(refused(
                    "an error value is #NULL!, #DIV/0!, #VALUE!, #REF!, #NAME?, #NUM!, #N/A or #GETTING_DATA",
                ));
            },
        }),
        (CellType::Error, _) => return Err(refused("the value is not an error value")),
    })
}

fn authored_operation(
    setter: CellSetter,
    value: Option<Scalar<'_>>,
    key: &'static str,
) -> Result<CellOperation> {
    match setter {
        CellSetter::Value => authored_value(value, key),
        CellSetter::Formula => authored_formula(value, key),
        CellSetter::Clear => authored_clear(value, key),
        CellSetter::Bold
        | CellSetter::Italic
        | CellSetter::Strike
        | CellSetter::Underline
        | CellSetter::FontOutline
        | CellSetter::FontShadow
        | CellSetter::FontCondense
        | CellSetter::FontExtend
        | CellSetter::FontCharset
        | CellSetter::FontFamily
        | CellSetter::FontScheme
        | CellSetter::AlignmentHorizontal
        | CellSetter::AlignmentVertical
        | CellSetter::AlignmentWrapText
        | CellSetter::AlignmentTextRotation
        | CellSetter::AlignmentIndent
        | CellSetter::AlignmentShrinkToFit
        | CellSetter::AlignmentReadingOrder
        | CellSetter::AlignmentRelativeIndent
        | CellSetter::AlignmentJustifyLastLine
        | CellSetter::ProtectionLocked
        | CellSetter::ProtectionHidden
        | CellSetter::QuotePrefix
        | CellSetter::NumberFormat
        | CellSetter::Fill
        | CellSetter::FillPattern
        | CellSetter::FillBg
        | CellSetter::BorderAll
        | CellSetter::BorderAllColor
        | CellSetter::BorderLeft
        | CellSetter::BorderLeftColor
        | CellSetter::BorderRight
        | CellSetter::BorderRightColor
        | CellSetter::BorderTop
        | CellSetter::BorderTopColor
        | CellSetter::BorderBottom
        | CellSetter::BorderBottomColor
        | CellSetter::BorderDiagonal
        | CellSetter::BorderDiagonalColor
        | CellSetter::BorderDiagonalUp
        | CellSetter::BorderDiagonalDown => Err(invalid_property(
            key,
            "format properties are not cell payloads",
        )),
        CellSetter::FontName
        | CellSetter::FontSize
        | CellSetter::Superscript
        | CellSetter::Subscript
        | CellSetter::FontColor => Err(invalid_property(
            key,
            "font properties are not cell payloads",
        )),
        CellSetter::Type | CellSetter::Runs | CellSetter::Ref => Err(invalid_property(
            key,
            "type, runs and ref are combined with the payload",
        )),
    }
}

fn authored_value(value: Option<Scalar<'_>>, key: &str) -> Result<CellOperation> {
    let content = match value {
        None => return Ok(CellOperation::Clear),
        Some(Scalar::Bool(value)) => xlsx::cell::Content::from(value),
        Some(Scalar::Int(value)) => xlsx::cell::Content::Value(authored_number(value, key)?),
        Some(Scalar::Number(value)) => {
            let number = xlsx::cell::Number::try_from(value)
                .map_err(|error| native_property_error(error, key))?;
            xlsx::cell::Content::Value(xlsx::cell::Value::from(number))
        },
        Some(Scalar::Text(value)) => {
            validate_text(value, key)?;
            xlsx::cell::Content::Value(xlsx::cell::Value::text(value))
        },
        Some(Scalar::Token(_) | Scalar::Length(_) | Scalar::Color(_) | Scalar::List(_)) => {
            return Err(invalid_property(
                key,
                "cell values support only text, bool, integer and finite number scalars",
            ));
        },
    };
    Ok(CellOperation::Set(content))
}

fn authored_number<T: std::fmt::Display>(value: T, key: &str) -> Result<xlsx::cell::Value> {
    let lexical = value.to_string();
    if lexical.len() > 32_767 {
        return Err(invalid_property(key, "the cell number is too long"));
    }
    let number = xlsx::cell::Number::new(lexical.into_boxed_str())
        .map_err(|error| native_property_error(error, key))?;
    Ok(xlsx::cell::Value::from(number))
}

fn authored_formula(value: Option<Scalar<'_>>, key: &str) -> Result<CellOperation> {
    let Some(value) = value else {
        return Ok(CellOperation::Clear);
    };
    let Scalar::Text(value) = value else {
        return Err(invalid_property(key, "formula must be text or null"));
    };
    validate_xml_text(value, key)?;
    let formula = xlsx::Formula::new(value.to_owned().into_boxed_str())
        .map_err(|error| native_property_error(error, key))?;
    Ok(CellOperation::Set(xlsx::cell::Content::from(formula)))
}

fn authored_clear(value: Option<Scalar<'_>>, key: &str) -> Result<CellOperation> {
    match value {
        Some(Scalar::Bool(true)) => Ok(CellOperation::Clear),
        Some(Scalar::Bool(false)) => Err(invalid_property(key, "clear must be true")),
        _ => Err(invalid_property(key, "clear must be a boolean")),
    }
}

pub(super) fn validate_text(value: &str, key: &str) -> Result<()> {
    if value.chars().count() > 32_767 {
        return Err(invalid_property(key, "the cell text is too long"));
    }
    validate_xml_text(value, key)
}

fn validate_font_name(value: &str, key: &str) -> Result<()> {
    if value.is_empty() {
        return Err(invalid_property(key, "font name must be non-empty text"));
    }
    validate_xml_text(value, key)
}

pub(super) fn native_font_color(
    value: crate::value::Color,
    key: &str,
) -> Result<xlsx::style::stylesheet::font::FontColor> {
    if !value.transforms().is_empty() {
        return Err(invalid_property(
            key,
            "font color transforms are not supported in SpreadsheetML",
        ));
    }
    let mut color = match value.base() {
        crate::value::Base::Rgb(value) => xlsx::style::stylesheet::font::FontColor::rgb(
            xlsx::Rgb::argb(value.alpha(), value.red(), value.green(), value.blue()),
        ),
        crate::value::Base::Theme(value) => {
            xlsx::style::stylesheet::font::FontColor::theme(value.index())
        },
        _ => {
            return Err(invalid_property(
                key,
                "font color must use an RGB or spreadsheet theme base",
            ));
        },
    };
    if let Some(tint) = value.spreadsheet_tint() {
        let tint = xlsx::style::stylesheet::border::Tint::new(tint).map_err(|_| {
            invalid_property(key, "font color tint is outside the SpreadsheetML range")
        })?;
        color = color.with_tint(tint);
    }
    Ok(color)
}

fn native_fill_color(
    value: crate::value::Color,
    key: &str,
) -> Result<xlsx::style::stylesheet::border::Color> {
    if !value.transforms().is_empty() {
        return Err(invalid_property(
            key,
            "fill color transforms are not supported in SpreadsheetML",
        ));
    }
    let mut color = match value.base() {
        crate::value::Base::Rgb(value) => xlsx::style::stylesheet::border::Color::argb(
            value.alpha(),
            value.red(),
            value.green(),
            value.blue(),
        ),
        crate::value::Base::Theme(value) => {
            xlsx::style::stylesheet::border::Color::theme(value.index())
        },
        _ => {
            return Err(invalid_property(
                key,
                "fill color must use an RGB or spreadsheet theme base",
            ));
        },
    };
    if let Some(tint) = value.spreadsheet_tint() {
        let tint = xlsx::style::stylesheet::border::Tint::new(tint).map_err(|_| {
            invalid_property(key, "fill color tint is outside the SpreadsheetML range")
        })?;
        color = color.with_tint(tint);
    }
    Ok(color)
}

enum AdmittedFill {
    Solid(xlsx::style::stylesheet::border::Color),
    Gradient(xlsx::style::cell_format::GradientFill),
}

fn admit_fill_value(
    value: &str,
    key: &'static str,
    limits: &SemanticLimits,
) -> Result<AdmittedFill> {
    if let Ok(color) = crate::value::Color::parse(value, crate::value::ColorContext::Spreadsheet) {
        return Ok(AdmittedFill::Solid(native_fill_color(color, key)?));
    }
    if looks_like_gradient(value) {
        return Ok(AdmittedFill::Gradient(parse_gradient_fill(
            value, key, limits,
        )?));
    }
    Err(invalid_property(
        key,
        "fill must be a SpreadsheetML color or a supported gradient",
    ))
}

fn looks_like_gradient(value: &str) -> bool {
    let value = value.trim();
    strip_prefix_ci(value, "gradient;").is_some()
        || strip_prefix_ci(value, "radial:").is_some()
        || value.contains('-')
        || value.contains(';')
}

fn strip_prefix_ci<'a>(value: &'a str, prefix: &str) -> Option<&'a str> {
    value
        .get(..prefix.len())
        .filter(|head| head.eq_ignore_ascii_case(prefix))
        .map(|_| &value[prefix.len()..])
}

fn parse_gradient_fill(
    value: &str,
    key: &'static str,
    limits: &SemanticLimits,
) -> Result<xlsx::style::cell_format::GradientFill> {
    let value = value.trim();
    if let Some(body) = strip_prefix_ci(value, "gradient;") {
        return parse_semicolon_gradient(body, key, limits);
    }

    if let Some(body) = strip_prefix_ci(value, "radial:") {
        return parse_dash_gradient(body, true, key, limits);
    }

    parse_dash_gradient(value, false, key, limits)
}

fn parsed_gradient_color(value: &str) -> Option<crate::value::Color> {
    crate::value::Color::parse(value.trim(), crate::value::ColorContext::Spreadsheet).ok()
}

fn parse_semicolon_gradient(
    value: &str,
    key: &'static str,
    limits: &SemanticLimits,
) -> Result<xlsx::style::cell_format::GradientFill> {
    let field_count = value.split(';').count();
    if field_count < 3 {
        return Err(invalid_property(
            key,
            "gradient fill must have at least two colors and a degree",
        ));
    }
    let color_count = field_count - 1;
    let degree_text = value.rsplit(';').next().unwrap_or_default();
    let degree = parse_gradient_degree(degree_text, key)?;
    check_gradient_staging(limits, color_count, 0, key)?;

    let mut stops = Vec::new();
    stops
        .try_reserve_exact(color_count)
        .map_err(|source| allocation("XLSX gradient stops", source))?;
    for (index, color) in value.split(';').take(color_count).enumerate() {
        #[allow(
            clippy::cast_precision_loss,
            reason = "gradient stop counts stay below the exact f64 integer range"
        )]
        let position = index as f64 / (color_count - 1) as f64;
        stops.push(gradient_stop_from_text(color, position, key)?);
    }
    let stops = stops.into_boxed_slice();
    xlsx::style::cell_format::GradientFill::linear(degree, stops).map_err(|_| {
        invalid_property(
            key,
            "gradient fill geometry is outside the supported domain",
        )
    })
}

fn parse_dash_gradient(
    value: &str,
    radial: bool,
    key: &'static str,
    limits: &SemanticLimits,
) -> Result<xlsx::style::cell_format::GradientFill> {
    let (separator_count, last_separator) = dash_separator_stats(value)?;
    let Some(last_separator) = last_separator else {
        return Err(invalid_property(
            key,
            "gradient fill must contain at least two colors",
        ));
    };
    let tail = &value[last_separator + 1..];
    let tail = tail.trim();
    let tail_is_color = parsed_gradient_color(tail).is_some();
    let numeric_angle = (!radial && !tail.is_empty() && (tail.len() <= 3 || !tail_is_color))
        .then(|| tail.parse::<f64>().ok())
        .flatten();
    let (color_end, degree, stop_count) = if let Some(angle) = numeric_angle {
        if !angle.is_finite() || !(0.0..=360.0).contains(&angle) {
            return Err(invalid_property(
                key,
                "gradient degree must be finite and between 0 and 360",
            ));
        }
        (last_separator, Some(angle), separator_count)
    } else if !radial && !tail_is_color && tail.len() <= 3 && !tail.is_empty() {
        return Err(invalid_property(
            key,
            "gradient degree must be finite and between 0 and 360",
        ));
    } else {
        (value.len(), None, separator_count + 1)
    };
    if stop_count < 2 {
        return Err(invalid_property(
            key,
            "gradient fill must contain at least two colors",
        ));
    }
    check_gradient_staging(limits, stop_count, separator_count, key)?;

    let mut separators = Vec::new();
    separators
        .try_reserve_exact(separator_count)
        .map_err(|source| allocation("XLSX gradient separators", source))?;
    scan_dash_separators(value, |index| separators.push(index));
    debug_assert_eq!(separators.len(), separator_count);

    let mut stops = Vec::new();
    stops
        .try_reserve_exact(stop_count)
        .map_err(|source| allocation("XLSX gradient stops", source))?;
    for index in 0..stop_count {
        let start = if index == 0 {
            0
        } else {
            separators[index - 1] + 1
        };
        let end = if index + 1 < stop_count {
            separators[index]
        } else {
            color_end
        };
        #[allow(
            clippy::cast_precision_loss,
            reason = "gradient stop counts stay below the exact f64 integer range"
        )]
        let position = index as f64 / (stop_count - 1) as f64;
        stops.push(gradient_stop_from_text(&value[start..end], position, key)?);
    }
    let stops = stops.into_boxed_slice();
    if radial {
        xlsx::style::cell_format::GradientFill::path(0.0, 0.0, 0.0, 0.0, stops).map_err(|_| {
            invalid_property(key, "radial fill geometry is outside the supported domain")
        })
    } else {
        xlsx::style::cell_format::GradientFill::linear(degree.unwrap_or(90.0), stops).map_err(
            |_| {
                invalid_property(
                    key,
                    "gradient fill geometry is outside the supported domain",
                )
            },
        )
    }
}

fn dash_separator_stats(value: &str) -> Result<(usize, Option<usize>)> {
    let mut count = 0usize;
    let mut last = None;
    scan_dash_separators(value, |index| {
        if let Some(next) = count.checked_add(1) {
            count = next;
            last = Some(index);
        }
    });
    if count == usize::MAX {
        return Err(staged_bytes_limit(usize::MAX, usize::MAX));
    }
    xlsx::style::cell_format::GradientFill::validate_stop_count(count).map_err(native_error)?;
    Ok((count, last))
}

fn scan_dash_separators(value: &str, mut visit: impl FnMut(usize)) {
    let mut index = 0;
    while index < value.len() {
        let Some(character) = value[index..].chars().next() else {
            break;
        };
        if character == '-' {
            if let Some(end) = tint_number_end(value, index) {
                index = end;
                continue;
            }
            visit(index);
        }
        index += character.len_utf8();
    }
}

fn tint_number_end(value: &str, index: usize) -> Option<usize> {
    let mut suffix = value[..index].trim_end();
    if let Some(before_equals) = suffix.strip_suffix('=') {
        suffix = before_equals.trim_end();
    }
    let name_start = suffix.len().checked_sub(4)?;
    if !suffix.get(name_start..)?.eq_ignore_ascii_case("tint")
        || !suffix.get(..name_start)?.trim_end().ends_with('+')
    {
        return None;
    }
    let rest = &value[index + 1..];
    let number_len = float_prefix_len(rest)?;
    let number_end = index + 1 + number_len;
    let trailing = &value[number_end..];
    let trimmed = trailing.trim_start();
    let end = number_end + trailing.len() - trimmed.len();
    if !trimmed.is_empty() && !trimmed.starts_with('-') {
        return None;
    }
    Some(end)
}

fn float_prefix_len(value: &str) -> Option<usize> {
    let bytes = value.as_bytes();
    let mut index = 0;
    let mut digits = 0;
    while index < bytes.len() && bytes[index].is_ascii_digit() {
        index += 1;
        digits += 1;
    }
    if index < bytes.len() && bytes[index] == b'.' {
        index += 1;
        while index < bytes.len() && bytes[index].is_ascii_digit() {
            index += 1;
            digits += 1;
        }
    }
    if digits == 0 {
        return None;
    }
    if index < bytes.len() && matches!(bytes[index], b'e' | b'E') {
        let exponent = index;
        index += 1;
        if index < bytes.len() && matches!(bytes[index], b'+' | b'-') {
            index += 1;
        }
        let exponent_start = index;
        while index < bytes.len() && bytes[index].is_ascii_digit() {
            index += 1;
        }
        if exponent_start == index {
            index = exponent;
        }
    }
    value[..index].parse::<f64>().ok().map(|_| index)
}

fn parse_gradient_degree(value: &str, key: &'static str) -> Result<f64> {
    value
        .trim()
        .parse::<f64>()
        .ok()
        .filter(|degree| degree.is_finite() && (0.0..=360.0).contains(degree))
        .ok_or_else(|| {
            invalid_property(key, "gradient degree must be finite and between 0 and 360")
        })
}

fn gradient_stop_from_text(
    value: &str,
    position: f64,
    key: &'static str,
) -> Result<xlsx::style::cell_format::GradientStop> {
    let color = parsed_gradient_color(value).ok_or_else(|| {
        invalid_property(key, "gradient fill contains an invalid SpreadsheetML color")
    })?;
    let color = native_fill_color(color, key)?;
    xlsx::style::cell_format::GradientStop::new(position, color)
        .map_err(|_| invalid_property(key, "gradient stop is outside the supported domain"))
}

#[derive(Clone, Copy)]
enum ExistingFill {
    Pattern {
        pattern: Option<xlsx::style::FillPattern>,
    },
    Gradient,
}

fn prepare_fill_changes(
    file: &Xlsx,
    targets: &[CellTarget],
    operation: &CellAssignments,
) -> Result<Vec<xlsx::style::FillChange>> {
    if operation.fill.is_none()
        && operation.fill_gradient.is_none()
        && operation.fill_pattern.is_none()
        && operation.fill_bg.is_none()
    {
        return Ok(Vec::new());
    }
    let mut changes = Vec::new();
    changes
        .try_reserve_exact(targets.len())
        .map_err(|source| allocation("XLSX cell fill changes", source))?;

    // `FillChange::clone` shares the native replacement's Arc.  Build one
    // gradient replacement for the whole target set instead of cloning its
    // boxed stops once per cell.
    let gradient_change = operation.fill_gradient.clone();
    if let Some(gradient_change) = gradient_change {
        changes.extend((0..targets.len()).map(|_| gradient_change.clone()));
        return Ok(changes);
    }

    let mut sheet_position = None;
    let mut sheet = None;
    for target in targets {
        if sheet_position != Some(target.sheet) {
            sheet = SheetHandle::at(file, target.sheet)?;
            sheet_position = Some(target.sheet);
        }
        let sheet = sheet.as_ref().ok_or_else(|| {
            Error::new(
                Code::OutOfRange,
                "the XLSX worksheet position is out of range",
            )
            .with_format(FileFormat::Xlsx)
        })?;
        changes.push(fill_change_for_target(sheet, target.address, operation)?);
    }
    Ok(changes)
}

fn existing_fill(sheet: &SheetHandle, address: Address) -> Result<ExistingFill> {
    let Some(format) = sheet.effective_format(address)? else {
        return Ok(ExistingFill::Pattern { pattern: None });
    };
    let Some(fill) = format.fill() else {
        return Ok(ExistingFill::Pattern { pattern: None });
    };
    Ok(match fill {
        xlsx::style::cell_format::Fill::Pattern(pattern) => ExistingFill::Pattern {
            pattern: pattern.pattern(),
        },
        xlsx::style::cell_format::Fill::Gradient(_) => ExistingFill::Gradient,
        _ => {
            return Err(unsupported_fill(
                "fill",
                "the native fill kind is not supported",
            ));
        },
    })
}

fn prepare_border_change(
    file: &Xlsx,
    targets: &[CellTarget],
    operation: &CellAssignments,
) -> Result<xlsx::style::BorderChange> {
    if !operation.has_border() {
        return Ok(xlsx::style::BorderChange::new());
    }

    // A non-null color must have a typed line either in this operation or in
    // the effective source format.  The native granular change preserves the
    // source side, but deliberately does not invent a side for color-only
    // requests; doing this check here gives the facade a stable, attributed
    // refusal before the native edit is staged.
    if border_source_validation_needed(operation) {
        let mut sheet_position = None;
        let mut sheet = None;
        for target in targets {
            if sheet_position != Some(target.sheet) {
                sheet = SheetHandle::at(file, target.sheet)?;
                sheet_position = Some(target.sheet);
            }
            let sheet = sheet.as_ref().ok_or_else(|| {
                Error::new(
                    Code::OutOfRange,
                    "the XLSX worksheet position is out of range",
                )
                .with_format(FileFormat::Xlsx)
            })?;
            validate_border_target(sheet, target.address, operation)?;
        }
    }

    // Every admitted target uses the same native change; retain one value and
    // let the leaf stage it, without a second per-cell clone array.
    Ok(border_change_for_operation(operation))
}

fn border_source_validation_needed(operation: &CellAssignments) -> bool {
    [
        (
            operation.border_left.or(operation.border_all),
            operation.border_left_color.or(operation.border_all_color),
        ),
        (
            operation.border_right.or(operation.border_all),
            operation.border_right_color.or(operation.border_all_color),
        ),
        (
            operation.border_top.or(operation.border_all),
            operation.border_top_color.or(operation.border_all_color),
        ),
        (
            operation.border_bottom.or(operation.border_all),
            operation.border_bottom_color.or(operation.border_all_color),
        ),
        (operation.border_diagonal, operation.border_diagonal_color),
    ]
    .into_iter()
    .any(|(line, color)| color.is_some_and(|value| value.is_some()) && line.is_none())
}

fn validate_border_target(
    sheet: &SheetHandle,
    address: Address,
    operation: &CellAssignments,
) -> Result<()> {
    let format = sheet.effective_format(address)?;
    let border = format.as_ref().and_then(|format| format.border());
    for (line, color, color_key, current) in [
        (
            operation.border_left.or(operation.border_all),
            operation.border_left_color.or(operation.border_all_color),
            "border.left.color",
            border.and_then(|border| border.left.as_ref()),
        ),
        (
            operation.border_right.or(operation.border_all),
            operation.border_right_color.or(operation.border_all_color),
            "border.right.color",
            border.and_then(|border| border.right.as_ref()),
        ),
        (
            operation.border_top.or(operation.border_all),
            operation.border_top_color.or(operation.border_all_color),
            "border.top.color",
            border.and_then(|border| border.top.as_ref()),
        ),
        (
            operation.border_bottom.or(operation.border_all),
            operation.border_bottom_color.or(operation.border_all_color),
            "border.bottom.color",
            border.and_then(|border| border.bottom.as_ref()),
        ),
        (
            operation.border_diagonal,
            operation.border_diagonal_color,
            "border.diagonal.color",
            border
                .and_then(|border| border.diagonal.as_ref())
                .and_then(|value| value.side()),
        ),
    ] {
        if color.is_some_and(|value| value.is_some()) && line.is_none() && current.is_none() {
            return Err(invalid_property(
                color_key,
                "a border color requires an existing typed line or a line in the same operation",
            ));
        }
    }
    Ok(())
}

fn border_change_for_operation(operation: &CellAssignments) -> xlsx::style::BorderChange {
    let mut change = xlsx::style::BorderChange::new();

    if let Some(value) = operation.border_all {
        change = change
            .left_line(value)
            .right_line(value)
            .top_line(value)
            .bottom_line(value);
    }
    if let Some(value) = operation.border_left {
        change = change.left_line(value);
    }
    if let Some(value) = operation.border_right {
        change = change.right_line(value);
    }
    if let Some(value) = operation.border_top {
        change = change.top_line(value);
    }
    if let Some(value) = operation.border_bottom {
        change = change.bottom_line(value);
    }
    if let Some(value) = operation.border_diagonal {
        change = change.diagonal_line(value);
    }

    if let Some(value) = operation.border_all_color {
        change = change
            .left_color(value)
            .right_color(value)
            .top_color(value)
            .bottom_color(value);
    }
    if let Some(value) = operation.border_left_color {
        change = change.left_color(value);
    }
    if let Some(value) = operation.border_right_color {
        change = change.right_color(value);
    }
    if let Some(value) = operation.border_top_color {
        change = change.top_color(value);
    }
    if let Some(value) = operation.border_bottom_color {
        change = change.bottom_color(value);
    }
    if let Some(value) = operation.border_diagonal_color {
        change = change.diagonal_color(value);
    }
    if let Some(value) = operation.border_diagonal_up {
        change = change.diagonal_up(value);
    }
    if let Some(value) = operation.border_diagonal_down {
        change = change.diagonal_down(value);
    }
    change
}

fn fill_change_for_target(
    sheet: &SheetHandle,
    address: Address,
    operation: &CellAssignments,
) -> Result<xlsx::style::FillChange> {
    if let Some(gradient) = operation.fill_gradient.as_ref() {
        if operation.fill_pattern.is_some() || operation.fill_bg.is_some() {
            return Err(invalid_property(
                "fill",
                "a gradient fill cannot be combined with fillPattern or fillBg",
            ));
        }
        return Ok(gradient.clone());
    }
    let existing = existing_fill(sheet, address)?;
    let explicit_pattern = operation.fill_pattern.is_some();
    let explicit_foreground = operation.fill.is_some();
    let explicit_background = operation.fill_bg.is_some();

    if let Some(pattern) = operation.fill_pattern {
        if pattern.is_none()
            && (operation.fill.is_some_and(|value| value.is_some())
                || operation.fill_bg.is_some_and(|value| value.is_some()))
        {
            return Err(invalid_property(
                fill_error_key(operation),
                "fill colors cannot accompany a null fillPattern",
            ));
        }
        if matches!(existing, ExistingFill::Gradient) {
            if pattern.is_none() {
                // A gradient has no authored pattern attribute to remove.
                // Keep this explicit null a no-op unless `fill=null` also
                // requests the whole gradient be cleared.
                if operation.fill.is_some_and(|value| value.is_none()) {
                    return Ok(xlsx::style::FillChange::replace(
                        xlsx::style::cell_format::Fill::Pattern(
                            xlsx::style::cell_format::PatternFill::new(None, None, None),
                        ),
                    ));
                }
                return Ok(xlsx::style::FillChange::new());
            }
            // A pattern request is the explicit escape from a source
            // gradient.  Replace the complete fill so no gradient children
            // remain hidden behind a granular pattern edit.
            return Ok(xlsx::style::FillChange::replace(
                xlsx::style::cell_format::Fill::Pattern(
                    xlsx::style::cell_format::PatternFill::new(
                        pattern,
                        operation.fill.unwrap_or(None),
                        operation.fill_bg.unwrap_or(None),
                    ),
                ),
            ));
        }
        if pattern.is_none() {
            // A null pattern removes only the authored pattern attribute.
            // Explicit color nulls still clear their selected fields, while
            // omitted colors remain untouched.
            let mut change = xlsx::style::FillChange::new().pattern(None);
            if let Some(color) = operation.fill {
                change = change.foreground(color);
            }
            if let Some(color) = operation.fill_bg {
                change = change.background(color);
            }
            return Ok(change);
        }
        let mut change = xlsx::style::FillChange::new().pattern(pattern);
        // A requested non-null pattern is a replacement operation. Unspecified
        // colors are explicitly cleared, matching OfficeCLI's pattern-fill
        // route.
        change = change.foreground(operation.fill.unwrap_or(None));
        change = change.background(operation.fill_bg.unwrap_or(None));
        return Ok(change);
    }

    if matches!(existing, ExistingFill::Gradient) {
        if let Some(foreground) = operation.fill {
            if operation.fill_bg.is_some_and(|value| value.is_some()) {
                return Err(unsupported_fill(
                    "fillBg",
                    "fillBg cannot accompany a solid replacement of a gradient",
                ));
            }
            let pattern = foreground.map(|_| xlsx::style::FillPattern::Solid);
            return Ok(xlsx::style::FillChange::replace(
                xlsx::style::cell_format::Fill::Pattern(
                    xlsx::style::cell_format::PatternFill::new(pattern, foreground, None),
                ),
            ));
        }
        if operation.fill_bg.is_some() {
            if operation.fill_bg.is_some_and(|value| value.is_none()) {
                // A gradient has no authored background attribute.  A null
                // fillBg therefore has the same no-op meaning as clearing an
                // absent pattern background.
                return Ok(xlsx::style::FillChange::new());
            }
            return Err(unsupported_fill(
                "fillBg",
                "fillBg cannot modify a gradient fill without a pattern or solid replacement",
            ));
        }
    }

    let ExistingFill::Pattern { pattern: current } = existing else {
        unreachable!();
    };
    let non_solid = current.is_some_and(|pattern| {
        !matches!(
            pattern,
            xlsx::style::FillPattern::None | xlsx::style::FillPattern::Solid
        )
    });

    if explicit_background && operation.fill_bg.is_some_and(|value| value.is_some()) && !non_solid {
        return Err(unsupported_fill(
            "fillBg",
            "fillBg requires an existing non-solid pattern or an explicit fillPattern",
        ));
    }

    let mut change = xlsx::style::FillChange::new();
    if let Some(foreground) = operation.fill {
        if foreground.is_some() && !non_solid {
            change = change.pattern(Some(xlsx::style::FillPattern::Solid));
        } else if foreground.is_some() && non_solid {
            // Keep the current non-solid pattern and its background.
            change = change.pattern(current);
        }
        change = change.foreground(foreground);
    }
    if let Some(background) = operation.fill_bg {
        if background.is_some() {
            // The non-solid check above has already proved this is safe.
            change = change.pattern(current);
        }
        change = change.background(background);
    }
    debug_assert!(explicit_pattern || explicit_foreground || explicit_background);
    Ok(change)
}

fn fill_error_key(operation: &CellAssignments) -> &'static str {
    if operation.fill_pattern.is_some() {
        "fillPattern"
    } else if operation.fill_bg.is_some() {
        "fillBg"
    } else {
        "fill"
    }
}

fn unsupported_fill(key: &'static str, message: &'static str) -> Error {
    Error::new(Code::UnsupportedCapability, message)
        .with_format(FileFormat::Xlsx)
        .with_property(key.to_owned())
}

fn validate_xml_text(value: &str, key: &str) -> Result<()> {
    if value.chars().any(|character| {
        !matches!(character, '\u{9}' | '\u{A}' | '\u{D}'
            | '\u{20}'..='\u{D7FF}'
            | '\u{E000}'..='\u{FFFD}'
            | '\u{10000}'..='\u{10FFFF}')
    }) {
        return Err(invalid_property(
            key,
            "the text contains an XML-forbidden character",
        ));
    }
    Ok(())
}

fn apply_operation(
    edit: &mut xlsx::workbook::Edit,
    targets: &[CellTarget],
    operation: &CellAssignments,
    fill_changes: &[xlsx::style::FillChange],
    border_change: &xlsx::style::BorderChange,
) -> Result<()> {
    let mut font = xlsx::style::FontChange::new();
    if let Some(value) = operation.bold {
        font = font.bold(value);
    }
    if let Some(value) = operation.italic {
        font = font.italic(value);
    }
    if let Some(value) = operation.strike {
        font = font.strike(value);
    }
    if let Some(value) = operation.outline {
        font = font.outline(value);
    }
    if let Some(value) = operation.shadow {
        font = font.shadow(value);
    }
    if let Some(value) = operation.condense {
        font = font.condense(value);
    }
    if let Some(value) = operation.extend {
        font = font.extend(value);
    }
    if let Some(value) = operation.underline {
        font = font.underline(value);
    }
    if let Some(value) = operation.font_name.as_ref() {
        font = font.name(value.as_deref()).map_err(native_error)?;
    }
    if let Some(value) = operation.font_size {
        font = font.size(value).map_err(native_error)?;
    }
    if let Some(value) = operation.script {
        font = font.script(value);
    }
    if let Some(value) = operation.font_color.as_ref() {
        font = font.color(value.clone());
    }
    if let Some(value) = operation.font_charset {
        font = font.charset(value).map_err(native_error)?;
    }
    if let Some(value) = operation.font_family {
        font = font.family(value).map_err(native_error)?;
    }
    if let Some(value) = operation.font_scheme {
        font = font.scheme(value);
    }
    let mut alignment = xlsx::style::AlignmentChange::new();
    if let Some(value) = operation.alignment_horizontal {
        alignment = alignment.horizontal(value);
    }
    if let Some(value) = operation.alignment_vertical {
        alignment = alignment.vertical(value);
    }
    if let Some(value) = operation.alignment_wrap_text {
        alignment = alignment.wrap_text(value);
    }
    if let Some(value) = operation.alignment_text_rotation {
        alignment = alignment.text_rotation(value);
    }
    if let Some(value) = operation.alignment_indent {
        alignment = alignment.indent(value);
    }
    if let Some(value) = operation.alignment_shrink_to_fit {
        alignment = alignment.shrink_to_fit(value);
    }
    if let Some(value) = operation.alignment_reading_order {
        alignment = alignment.reading_order(value);
    }
    if let Some(value) = operation.alignment_relative_indent {
        alignment = alignment.relative_indent(value);
    }
    if let Some(value) = operation.alignment_justify_last_line {
        alignment = alignment.justify_last_line(value);
    }
    let mut protection = xlsx::style::ProtectionChange::new();
    if let Some(value) = operation.protection_locked {
        protection = protection.locked(value);
    }
    if let Some(value) = operation.protection_hidden {
        protection = protection.hidden(value);
    }
    // Keep one worksheet editor alive for each contiguous sheet run. This is
    // still one native workbook transaction; there is no per-cell facade
    // setter or serialize/reopen boundary.
    let mut index = 0;
    while index < targets.len() {
        let sheet = targets[index].sheet;
        if let Some(payload) = operation.payload.as_ref() {
            let mut reserve_index = index;
            while let Some(target) = targets
                .get(reserve_index)
                .filter(|target| target.sheet == sheet)
            {
                edit.reserve_cell_capture(sheet, target.address, operation_payload_bytes(payload))
                    .map_err(native_error)?;
                reserve_index += 1;
            }
        }
        let mut worksheet = edit.sheet(sheet).map_err(native_error)?.ok_or_else(|| {
            Error::new(
                Code::OutOfRange,
                "the XLSX worksheet position is out of range",
            )
            .with_format(FileFormat::Xlsx)
        })?;
        while let Some(target) = targets.get(index).filter(|target| target.sheet == sheet) {
            match &operation.payload {
                Some(CellOperation::Set(content)) => {
                    worksheet
                        .set(target.address, content.clone())
                        .map_err(native_error)?;
                },
                Some(CellOperation::Clear) => {
                    worksheet.clear(target.address).map_err(native_error)?;
                },
                Some(CellOperation::Rich(rich)) => {
                    worksheet
                        .set_rich_text(target.address, rich.clone())
                        .map_err(native_error)?;
                },
                Some(CellOperation::Retype(cell_type)) => {
                    retype(&mut worksheet, target.address, *cell_type)?;
                },
                None => {},
            };
            if !font.is_empty() {
                worksheet
                    .font(target.address, &font)
                    .map_err(native_error)?;
            }
            if !alignment.is_empty() {
                worksheet
                    .alignment(target.address, &alignment)
                    .map_err(native_error)?;
            }
            if !protection.is_empty() {
                worksheet
                    .protection(target.address, &protection)
                    .map_err(native_error)?;
            }
            if let Some(value) = operation.quote_prefix {
                worksheet
                    .quote_prefix(target.address, value)
                    .map_err(native_error)?;
            }
            if let Some(value) = operation.number_format.as_ref() {
                worksheet
                    .number_format(target.address, value.as_deref())
                    .map_err(native_error)?;
            }
            if !fill_changes.is_empty() {
                worksheet
                    .fill(target.address, &fill_changes[index])
                    .map_err(native_error)?;
            }
            if !border_change.is_empty() {
                worksheet
                    .border(target.address, border_change)
                    .map_err(native_error)?;
            }
            index += 1;
        }
    }
    Ok(())
}

/// Convert the value stored at `address` to `cell_type` (record 0338): an
/// empty coordinate is left alone, and a formula cell is refused.
fn retype(
    worksheet: &mut xlsx::workbook::edit::Sheet<'_>,
    address: Address,
    cell_type: CellType,
) -> Result<()> {
    let state = worksheet.state(address).map_err(native_error)?;
    let xlsx::State::Cell { content, .. } = state else {
        return Ok(());
    };
    match content {
        xlsx::Cell::Value(value) => {
            let content = if cell_type == CellType::RichText {
                let text = value_text(&value)
                    .ok_or_else(|| invalid_property("type", "an error value has no runs"))?;
                worksheet
                    .set_rich_text(address, xlsx::RichText::plain(text))
                    .map_err(native_error)?;
                return Ok(());
            } else {
                typed_value(value, cell_type)?
            };
            worksheet
                .set(address, xlsx::cell::Content::Value(content))
                .map_err(native_error)?;
            Ok(())
        },
        _ => Err(Error::new(
            Code::InvalidArgument,
            "type converts a stored value, not a formula",
        )
        .with_property("type")
        .with_format(FileFormat::Xlsx)),
    }
}

pub(super) fn invalid_property(key: &str, message: &'static str) -> Error {
    Error::new(Code::InvalidValue, message)
        .with_format(FileFormat::Xlsx)
        .with_property(key.to_owned())
}

pub(super) fn native_error(error: xlsx::Error) -> Error {
    Error::from(error).with_format(FileFormat::Xlsx)
}

pub(super) fn native_property_error(error: xlsx::Error, key: &str) -> Error {
    native_error(error)
        .caused_by_argument(Code::InvalidValue)
        .with_property(key.to_owned())
}

fn duplicate_property(key: &str) -> Error {
    Error::new(
        Code::InvalidArgument,
        "the same XLSX cell property was supplied more than once",
    )
    .with_property(key.to_owned())
}

fn admit_border_line(
    value: Option<Scalar<'_>>,
    key: &'static str,
) -> Result<Option<xlsx::style::cell_format::Line>> {
    match value {
        None => Ok(None),
        Some(Scalar::Token("none")) => Ok(None),
        Some(Scalar::Token(value)) => value.parse().map(Some).map_err(|_| {
            invalid_property(key, "border style has an unsupported SpreadsheetML token")
        }),
        _ => Err(invalid_property(
            key,
            "border style must be a token or null",
        )),
    }
}

fn admit_border_color(
    value: Option<Scalar<'_>>,
    key: &'static str,
) -> Result<Option<xlsx::style::stylesheet::border::Color>> {
    match value {
        None => Ok(None),
        Some(Scalar::Color(value)) => Ok(Some(native_fill_color(value, key)?)),
        _ => Err(invalid_property(
            key,
            "border color must be a SpreadsheetML color or null",
        )),
    }
}

fn admit_border_bool(value: Option<Scalar<'_>>, key: &'static str) -> Result<Option<bool>> {
    match value {
        None => Ok(None),
        Some(Scalar::Bool(value)) => Ok(Some(value)),
        _ => Err(invalid_property(
            key,
            "border direction must be boolean or null",
        )),
    }
}

fn validate_border_assignments(operation: &CellAssignments) -> Result<()> {
    for (line, color, color_key) in [
        (
            operation.border_left.or(operation.border_all),
            operation.border_left_color.or(operation.border_all_color),
            "border.left.color",
        ),
        (
            operation.border_right.or(operation.border_all),
            operation.border_right_color.or(operation.border_all_color),
            "border.right.color",
        ),
        (
            operation.border_top.or(operation.border_all),
            operation.border_top_color.or(operation.border_all_color),
            "border.top.color",
        ),
        (
            operation.border_bottom.or(operation.border_all),
            operation.border_bottom_color.or(operation.border_all_color),
            "border.bottom.color",
        ),
        (
            operation.border_diagonal,
            operation.border_diagonal_color,
            "border.diagonal.color",
        ),
    ] {
        if line == Some(None) && color.is_some_and(|value| value.is_some()) {
            return Err(invalid_property(
                color_key,
                "a border color cannot accompany a side with no line",
            ));
        }
    }
    Ok(())
}

fn property_error(error: Error, key: &crate::props::Key<'_>) -> Error {
    error
        .with_format(FileFormat::Xlsx)
        .with_property(key.as_str().to_owned())
}

/// A bounded run of disjoint direct cell payload writes.
///
/// This is deliberately kept behind the format adapter.  The transaction
/// coordinator owns the run lifecycle, while this module owns target
/// admission, native projected-state comparison, and the native commit.  A
/// run never shadows cell state in the unified layer: the native
/// `xlsx::workbook::edit::Sheet::state` view is the authority for each staged before/after
/// pair.
/// CSV/TSV import through one native edit (record 0266).
mod find;
pub(crate) mod import;
pub(crate) mod merge;
#[cfg(feature = "eval")]
pub(crate) mod refresh;

pub(crate) mod coalesce {
    use super::{
        CellOperation, CellTarget, Commit, Xlsx, admit_properties_with_limits,
        edit_with_snapshot_limit, native_error, snapshot_limit,
    };
    use crate::error::{Code, Error, Result};
    use crate::format::signature::Signature;
    use crate::props::Props;
    use crate::target::TargetArg;
    use crate::transaction::{Error as TransactionError, SemanticLimits};
    use litchi_core::FileFormat;
    use litchi_xlsx as xlsx;
    use std::collections::HashMap;

    /// One admitted direct cell payload assignment.
    pub(crate) struct Plan {
        target: CellTarget,
        operation: CellOperation,
        staged_bytes: usize,
        /// The package promoted while resolving the first plan for a run.
        /// The coordinator hands this owner to `Run::new` so the same source
        /// is not promoted a second time before native staging begins.
        promoted_source: Option<xlsx::Package>,
    }

    impl Plan {
        fn charge(&self) -> usize {
            self.staged_bytes
        }

        fn key(&self) -> CellTarget {
            self.target
        }

        pub(crate) fn take_promoted_source(&mut self) -> Option<xlsx::Package> {
            self.promoted_source.take()
        }
    }

    struct Item {
        plan: Plan,
        index: usize,
    }

    /// Native transaction for one eligible XLSX set run.
    ///
    /// Distinct direct cell targets are admitted for ordinary runs. Signed
    /// replay runs may repeat a target to prove an exact cancellation; its
    /// authored charge is replaced by the final target charge. Selectors,
    /// ranges, font assignments, or a staged-byte boundary remain barriers
    /// owned by the outer transaction.
    pub(crate) struct Run {
        base: Xlsx,
        resolution: Xlsx,
        source: xlsx::Package,
        edit: xlsx::workbook::Edit,
        items: Vec<Item>,
        /// Each staged target's charge; a key is present exactly when the
        /// target was staged in this run (record 0377: one map, not a set
        /// and a map, per cell).
        charges: HashMap<CellTarget, usize>,
        initial: Option<HashMap<CellTarget, xlsx::State>>,
        final_states: Option<HashMap<CellTarget, xlsx::State>>,
        /// The last sheet whose worksheet before-image was admitted. A
        /// semantic no-op can release that admission, so `stage_native`
        /// clears this after an unchanged operation before the next read.
        capture_sheet: Option<usize>,
        staged_bytes: usize,
        max_staged_bytes: usize,
        max_snapshot_bytes: usize,
        allow_repeated: bool,
    }

    impl Xlsx {
        /// Admit the narrow coalesced family.  Returning `None` means that
        /// ordinary XLSX set handling remains authoritative for this target or
        /// property shape; an admission error is preserved for the caller.
        pub(crate) fn cell_plan(
            &self,
            target: &TargetArg<'_>,
            props: &Props<'_>,
            limits: &SemanticLimits,
        ) -> Result<Option<Plan>> {
            self.cell_plan_with_snapshot_limit(target, props, limits, limits.max_snapshot_bytes())
        }

        /// Admit a coalesced plan against the transaction's remaining
        /// snapshot ceiling. The ordinary helper is retained for callers
        /// that have no prior committed snapshots; transaction batches use
        /// this variant so promotion cannot spend the batch's full maximum.
        pub(crate) fn cell_plan_with_snapshot_limit(
            &self,
            target: &TargetArg<'_>,
            props: &Props<'_>,
            limits: &SemanticLimits,
            remaining_snapshot: usize,
        ) -> Result<Option<Plan>> {
            self.check_source_for_replay()?;
            let result = (|| {
                let Some(path) = target.path() else {
                    return Ok(None);
                };
                // Root names such as `/A1` select tabs. Only a scoped,
                // explicit cell address belongs to this coalesced family.
                if path.path().steps().len() < 2
                    || path
                        .path()
                        .steps()
                        .last()
                        .is_none_or(|step| step.cell_ref().is_none())
                {
                    return Ok(None);
                }
                // Merges and cell hyperlinks (`link`, `tooltip`, `display`
                // and their aliases) have their own ordinary set paths; the
                // payload admission below does not know them (record 0330).
                if props.iter().any(|(key, _)| {
                    key.as_str().eq_ignore_ascii_case("merge")
                        || super::super::spec::CELL_LINK_TABLE
                            .lookup(key.as_str())
                            .is_some()
                }) {
                    return Ok(None);
                }
                let assignments = admit_properties_with_limits(props, limits)?;
                // Formatting changes and empty property sets have their own
                // native semantics and therefore close this payload-only run.
                if assignments.bold.is_some()
                    || assignments.italic.is_some()
                    || assignments.strike.is_some()
                    || assignments.outline.is_some()
                    || assignments.shadow.is_some()
                    || assignments.condense.is_some()
                    || assignments.extend.is_some()
                    || assignments.underline.is_some()
                    || assignments.font_name.is_some()
                    || assignments.font_size.is_some()
                    || assignments.script.is_some()
                    || assignments.font_color.is_some()
                    || assignments.font_charset.is_some()
                    || assignments.font_family.is_some()
                    || assignments.font_scheme.is_some()
                    || assignments.alignment_horizontal.is_some()
                    || assignments.alignment_vertical.is_some()
                    || assignments.alignment_wrap_text.is_some()
                    || assignments.alignment_text_rotation.is_some()
                    || assignments.alignment_indent.is_some()
                    || assignments.alignment_shrink_to_fit.is_some()
                    || assignments.alignment_reading_order.is_some()
                    || assignments.alignment_relative_indent.is_some()
                    || assignments.alignment_justify_last_line.is_some()
                    || assignments.protection_locked.is_some()
                    || assignments.protection_hidden.is_some()
                    || assignments.quote_prefix.is_some()
                    || assignments.number_format.is_some()
                    || assignments.fill.is_some()
                    || assignments.fill_gradient.is_some()
                    || assignments.fill_pattern.is_some()
                    || assignments.fill_bg.is_some()
                    || assignments.has_border()
                    || assignments.payload.is_none()
                {
                    return Ok(None);
                }
                let staged_bytes = super::per_target_charge(&assignments, limits)?;
                super::check_range_area(1, staged_bytes, limits)?;
                let Some(operation) = assignments.payload else {
                    return Ok(None);
                };
                // Runs and retypes have their own staging (record 0338).
                if matches!(operation, CellOperation::Rich(_) | CellOperation::Retype(_)) {
                    return Ok(None);
                }
                // Mutation already requires native promotion. Resolve
                // against that owner so a retained source does not stream
                // the worksheet again for each command in this run.
                let promoted = match &self.0.owner {
                    super::super::XlsxOwner::Source(_) if self.signature() == Signature::Signed => {
                        Some(self.adopt_workbook(
                            self.editor_for_write_with_snapshot_limit(remaining_snapshot)?,
                        ))
                    },
                    super::super::XlsxOwner::Source(_) => Some(
                        self.adopt_workbook(self.editor_with_snapshot_limit(remaining_snapshot)?),
                    ),
                    _ => None,
                };
                let resolved = promoted.as_ref().unwrap_or(self);
                let node = crate::path::resolve::get(&super::super::Nav::new(resolved), path)?;
                let Some((sheet, address)) = node.write_cell_target() else {
                    return Ok(None);
                };
                super::check_range_area(1, staged_bytes, limits)?;
                Ok(Some(Plan {
                    target: CellTarget { sheet, address },
                    operation,
                    staged_bytes,
                    promoted_source: promoted.map(|owner| match &owner.0.owner {
                        super::super::XlsxOwner::Eager(package) => package.clone(),
                        super::super::XlsxOwner::Source(_) => {
                            unreachable!("a promoted XLSX plan must own its package")
                        },
                    }),
                }))
            })();
            let freshness = self.check_source_for_replay();
            freshness.and(result)
        }
    }

    impl Run {
        /// Promote/admit the native source once for this run.
        pub(crate) fn new(
            base: &Xlsx,
            limits: &SemanticLimits,
            remaining_snapshot: usize,
            promoted_source: Option<xlsx::Package>,
        ) -> Result<Self> {
            base.check_source_for_replay()?;
            let source = match promoted_source {
                Some(source) => source,
                None if base.signature() == Signature::Signed => {
                    base.editor_for_write_with_snapshot_limit(remaining_snapshot)?
                },
                None => base.editor_with_snapshot_limit(remaining_snapshot)?,
            };
            let edit = source
                .edit_with_options(
                    xlsx::workbook::edit::EditOptions::default()
                        .with_max_retained_bytes(remaining_snapshot),
                )
                .map_err(native_error)?;
            base.check_source_for_replay()?;
            Ok(Self {
                base: base.clone(),
                resolution: base.adopt_workbook(source.clone()),
                source,
                edit,
                items: Vec::new(),
                charges: HashMap::new(),
                initial: (base.signature() == Signature::Signed).then(HashMap::new),
                final_states: (base.signature() == Signature::Signed).then(HashMap::new),
                capture_sheet: None,
                staged_bytes: 0,
                max_staged_bytes: limits.max_staged_bytes(),
                max_snapshot_bytes: remaining_snapshot,
                allow_repeated: base.signature() == Signature::Signed,
            })
        }

        /// Resolve stable direct addresses through the promoted owner while
        /// retaining the original source's freshness fence.
        pub(crate) fn cell_plan(
            &self,
            target: &TargetArg<'_>,
            props: &Props<'_>,
            limits: &SemanticLimits,
        ) -> Result<Option<Plan>> {
            self.base.check_source_for_replay()?;
            let result = self.resolution.cell_plan(target, props, limits);
            self.base.check_source_for_replay()?;
            result
        }

        /// The first wire index in the run, used for a final fallback when a
        /// native aggregate failure has no uniquely isolatable item.
        pub(crate) fn first_index(&self) -> Option<usize> {
            self.items.first().map(|item| item.index)
        }

        /// Whether staging this plan would cross the run's disjoint-target or
        /// cumulative staged-byte boundary.  The plan itself has already
        /// passed the per-cell admission bound.
        pub(crate) fn requires_barrier(&self, plan: &Plan, limits: &SemanticLimits) -> bool {
            let staged = self.charges.get(&plan.key()).copied();
            let previous = staged.unwrap_or(0);
            staged.is_some() && !self.allow_repeated
                || self
                    .staged_bytes
                    .saturating_sub(previous)
                    .checked_add(plan.charge())
                    .is_none_or(|bytes| bytes > limits.max_staged_bytes())
        }

        /// Stage a plan if it remains in this run's narrow family.
        ///
        /// The coordinator must call [`Self::requires_barrier`] first; a
        /// missed barrier is an internal error rather than an ordinary item
        /// refusal, because staging after it would make changed attribution
        /// ambiguous.
        pub(crate) fn stage(&mut self, plan: Plan, index: usize) -> Result<bool> {
            let staged = self.charges.get(&plan.key()).copied();
            let previous = staged.unwrap_or(0);
            if staged.is_some() && !self.allow_repeated
                || self
                    .staged_bytes
                    .saturating_sub(previous)
                    .checked_add(plan.charge())
                    .is_none_or(|bytes| bytes > self.max_staged_bytes)
            {
                return Err(Error::new(
                    Code::InternalError,
                    "XLSX cell coalescing barrier was not flushed",
                )
                .with_format(FileFormat::Xlsx));
            }
            self.base.check_source_for_replay()?;
            self.items
                .try_reserve(1)
                .map_err(|source| TransactionError::Allocation { source })?;
            self.charges
                .try_reserve(1)
                .map_err(|source| TransactionError::Allocation { source })?;

            let result = self.stage_native(&plan).inspect(|_| {
                self.staged_bytes = self
                    .staged_bytes
                    .saturating_sub(previous)
                    .saturating_add(plan.charge());
                self.charges.insert(plan.key(), plan.charge());
                self.items.push(Item { plan, index });
            });
            let freshness = self.base.check_source_for_replay();
            freshness.and(result)
        }

        fn stage_native(&mut self, plan: &Plan) -> Result<bool> {
            // A projected state read can decode a deferred worksheet. Admit
            // the native before-image before that read; the concrete setter
            // below replaces this coarse authored charge with its final
            // semantic payload.
            // Signed runs intentionally permit repeated targets while they
            // prove cancellation endpoints. A repeated setter can release
            // the last capture even when its projected before/after states
            // differ, so keep those runs on the native reservation path.
            if self.allow_repeated || self.capture_sheet != Some(plan.target.sheet) {
                self.edit
                    .reserve_cell_capture(plan.target.sheet, plan.target.address, plan.charge())
                    .map_err(native_error)?;
            }
            let mut worksheet = self
                .edit
                .sheet(plan.target.sheet)
                .map_err(native_error)?
                .ok_or_else(|| {
                    Error::new(
                        Code::OutOfRange,
                        "the XLSX worksheet position is out of range",
                    )
                    .with_format(FileFormat::Xlsx)
                })?;
            let before = worksheet.state(plan.target.address).map_err(native_error)?;
            let key = plan.key();
            if let Some(initial) = self.initial.as_mut() {
                initial.entry(key).or_insert_with(|| before.clone());
            }
            match &plan.operation {
                CellOperation::Set(content) => worksheet
                    .set(plan.target.address, content.clone())
                    .map_err(native_error)?,
                CellOperation::Clear => {
                    worksheet.clear(plan.target.address).map_err(native_error)?
                },
                CellOperation::Rich(_) | CellOperation::Retype(_) => {
                    return Err(not_coalesced());
                },
            };
            let after = worksheet.state(plan.target.address).map_err(native_error)?;
            if let Some(final_states) = self.final_states.as_mut() {
                final_states.insert(key, after.clone());
            }
            let changed = before != after;
            // A bounded setter may release the capture it just admitted when
            // it proves an exact source no-op. Do not carry the sheet cache
            // across that release; the next projected read must re-admit the
            // worksheet before decoding it.
            self.capture_sheet = (!self.allow_repeated && changed).then_some(plan.target.sheet);
            Ok(changed)
        }

        /// Commit one native patch for the run.  Native aggregate failures
        /// are replayed through fresh editors to retain the originating wire
        /// index, following the DOCX coalescer's error-only bisection rule.
        pub(crate) fn commit(self) -> Result<(Xlsx, Commit)> {
            let Self {
                base,
                source,
                edit,
                items,
                initial,
                final_states,
                allow_repeated,
                max_snapshot_bytes,
                ..
            } = self;
            let logical_noop = allow_repeated
                && initial.as_ref().is_some_and(|initial| {
                    !initial.is_empty()
                        && final_states.as_ref().is_some_and(|final_states| {
                            initial
                                .iter()
                                .all(|(key, before)| final_states.get(key) == Some(before))
                        })
                });
            let original = if logical_noop {
                Some(base.original_package_for_noop()?)
            } else {
                None
            };
            base.check_source_for_replay()?;
            // Keep the staged source in an option so a signed logical no-op
            // can explicitly drop the probe and its native proof before the
            // one canonical identity edit is allocated. The failure-only
            // locator still borrows the staged source for ordinary writes.
            let mut staged_source = Some(source);
            let result = edit.commit().map_err(native_error).and_then(|native| {
                if logical_noop {
                    let staged_source = staged_source
                        .take()
                        .ok_or_else(|| Error::signed(FileFormat::Xlsx))?;
                    let (staged_target, staged_patch) = native.into_parts();
                    let staged_snapshot_bytes = staged_patch.retained_bytes();
                    drop((staged_source, staged_target, staged_patch));

                    let original = original
                        .as_ref()
                        .ok_or_else(|| Error::signed(FileFormat::Xlsx))?;
                    let (target, native_patch) =
                        edit_with_snapshot_limit(original, max_snapshot_bytes)?
                            .commit()
                            .map_err(native_error)?
                            .into_parts();
                    let canonical_snapshot_bytes = native_patch.retained_bytes();
                    let snapshot_bytes = staged_snapshot_bytes.max(canonical_snapshot_bytes);
                    if snapshot_bytes > max_snapshot_bytes {
                        return Err(snapshot_limit(snapshot_bytes, max_snapshot_bytes));
                    }
                    Ok((
                        base.clone(),
                        Commit {
                            after_xlsx: base.clone(),
                            source: original.clone(),
                            target,
                            native_patch,
                            operations: items.len(),
                            snapshot_bytes,
                        },
                    ))
                } else {
                    let source = staged_source.as_ref().ok_or_else(|| {
                        Error::new(Code::InternalError, "XLSX staged source missing")
                    })?;
                    make_commit(&base, source, native, items.len(), max_snapshot_bytes)
                }
            });
            let freshness = base.check_source_for_replay();
            match freshness.and(result) {
                Ok(commit) => Ok(commit),
                Err(error) if logical_noop => Err(error),
                Err(error) => Err(locate_failure(
                    &base,
                    staged_source
                        .as_ref()
                        .expect("ordinary XLSX failures retain their staged source"),
                    &items,
                    error,
                    max_snapshot_bytes,
                )),
            }
        }
    }

    fn make_commit(
        base: &Xlsx,
        source: &xlsx::Package,
        native: xlsx::workbook::Commit,
        operations: usize,
        max_snapshot_bytes: usize,
    ) -> Result<(Xlsx, Commit)> {
        let mut commit =
            Commit::from_native_with_limit(base, source.clone(), native, max_snapshot_bytes)?;
        commit.operations = operations;
        let snapshot_bytes = commit.snapshot_bytes();
        if snapshot_bytes > max_snapshot_bytes {
            return Err(snapshot_limit(snapshot_bytes, max_snapshot_bytes));
        }
        Ok((base.clone(), commit))
    }

    fn replay(
        base: &Xlsx,
        source: &xlsx::Package,
        items: &[Item],
        max_snapshot_bytes: usize,
    ) -> Result<xlsx::Package> {
        base.check_source_for_replay()?;
        let mut edit = edit_with_snapshot_limit(source, max_snapshot_bytes)?;
        for item in items {
            stage_replay(&mut edit, &item.plan).map_err(|error| error.with_item(item.index))?;
        }
        let result = edit.commit().map_err(native_error).map(|commit| {
            let (target, _) = commit.into_parts();
            target
        });
        let freshness = base.check_source_for_replay();
        freshness.and(result)
    }

    fn stage_replay(edit: &mut xlsx::workbook::Edit, plan: &Plan) -> Result<()> {
        let mut worksheet = edit
            .sheet(plan.target.sheet)
            .map_err(native_error)?
            .ok_or_else(|| {
                Error::new(
                    Code::OutOfRange,
                    "the XLSX worksheet position is out of range",
                )
                .with_format(FileFormat::Xlsx)
            })?;
        match &plan.operation {
            CellOperation::Set(content) => worksheet
                .set(plan.target.address, content.clone())
                .map_err(native_error)?,
            CellOperation::Clear => worksheet.clear(plan.target.address).map_err(native_error)?,
            CellOperation::Rich(_) | CellOperation::Retype(_) => return Err(not_coalesced()),
        };
        Ok(())
    }

    fn not_coalesced() -> Error {
        Error::new(Code::InternalError, "runs and retypes are not coalesced")
            .with_format(FileFormat::Xlsx)
    }

    fn locate_failure(
        base: &Xlsx,
        source: &xlsx::Package,
        items: &[Item],
        error: Error,
        max_snapshot_bytes: usize,
    ) -> Error {
        locate_with(base, source, items, error, max_snapshot_bytes)
    }

    fn locate_with(
        base: &Xlsx,
        source: &xlsx::Package,
        items: &[Item],
        error: Error,
        max_snapshot_bytes: usize,
    ) -> Error {
        let Some(first) = items.first() else {
            return error;
        };
        if items.len() == 1
            || matches!(
                error.code(),
                Code::IoError
                    | Code::Cancelled
                    | Code::LimitExceeded
                    | Code::InternalError
                    | Code::DocumentProtected
            )
        {
            return error.with_item(first.index);
        }
        let middle = items.len() / 2;
        let prefix = match replay(base, source, &items[..middle], max_snapshot_bytes) {
            Ok(target) => target,
            Err(error) => {
                return locate_with(base, source, &items[..middle], error, max_snapshot_bytes);
            },
        };
        match replay(base, &prefix, &items[middle..], max_snapshot_bytes) {
            Err(error) => locate_with(base, &prefix, &items[middle..], error, max_snapshot_bytes),
            Ok(_) => error.with_item(first.index),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::{
        CellOperation, admit_properties, authored_formula, authored_value, parse_gradient_fill,
        validate_xml_text,
    };
    use crate::error::Code;
    use crate::props::Props;
    use crate::transaction::SemanticLimits;
    use crate::value::Scalar;
    use litchi_xlsx as xlsx;

    #[test]
    fn authored_text_accepts_spreadsheet_controls() {
        assert!(validate_xml_text("a\tb\nc\rd", "value").is_ok());
        assert!(validate_xml_text("a\u{1}b", "value").is_err());
    }

    #[test]
    fn scalar_values_keep_native_kinds() {
        let value = authored_value(Some(Scalar::Int(42)), "value").expect("number");
        assert!(matches!(
            value,
            CellOperation::Set(xlsx::cell::Content::Value(xlsx::cell::Value::Number(number)))
                if number.as_str() == "42"
        ));

        let value = authored_value(Some(Scalar::Bool(true)), "value").expect("bool");
        assert!(matches!(
            value,
            CellOperation::Set(xlsx::cell::Content::Value(xlsx::cell::Value::Bool(true)))
        ));

        let formula =
            authored_formula(Some(Scalar::Text("SUM(A1:A2)")), "formula").expect("formula");
        assert!(matches!(
            formula,
            CellOperation::Set(xlsx::cell::Content::Formula(formula))
                if formula.text() == "SUM(A1:A2)"
        ));
    }

    #[test]
    fn scalar_registry_rejects_wide_integer_values() {
        let props = Props::new().with("value", u128::MAX);
        let error = admit_properties(&props).expect_err("wide scalar");
        assert_eq!(error.code(), Code::InvalidValue);
    }

    #[test]
    fn native_formula_validation_keeps_leaf_source() {
        let error =
            authored_formula(Some(Scalar::Text("=A1")), "formula").expect_err("leading equals");
        assert_eq!(error.code(), Code::InvalidValue);
        assert!(std::error::Error::source(&error).is_some());
    }

    #[test]
    fn cell_property_conflicts_are_admitted_before_native_edit() {
        let props = Props::new().with("value", 1).with("clear", true);
        let error = admit_properties(&props).expect_err("conflict");
        assert_eq!(error.code(), Code::InvalidArgument);

        let props = Props::new().with("unknown", 1);
        let error = admit_properties(&props).expect_err("unknown");
        assert_eq!(error.code(), Code::UnsupportedProperty);
    }

    #[test]
    fn long_tail_font_properties_use_closed_domains_and_null_clears() {
        let props = Props::new()
            .with("font.outline", true)
            .with("font.shadow", false)
            .with("font.condense", true)
            .with("font.extend", false)
            .with("font.charset", 255)
            .with("font.family", 5)
            .with("font.scheme", "major");
        let operation = admit_properties(&props).expect("long-tail font properties");
        assert_eq!(operation.outline, Some(Some(true)));
        assert_eq!(operation.shadow, Some(Some(false)));
        assert_eq!(operation.condense, Some(Some(true)));
        assert_eq!(operation.extend, Some(Some(false)));
        assert_eq!(operation.font_charset, Some(Some(255)));
        assert_eq!(operation.font_family, Some(Some(5)));
        assert_eq!(
            operation.font_scheme,
            Some(Some(xlsx::style::Scheme::Major))
        );

        let clears = Props::new()
            .with("font.outline", crate::props::PropValue::Null)
            .with("font.charset", crate::props::PropValue::Null)
            .with("font.scheme", crate::props::PropValue::Null);
        let operation = admit_properties(&clears).expect("null removes authored children");
        assert_eq!(operation.outline, Some(None));
        assert_eq!(operation.font_charset, Some(None));
        assert_eq!(operation.font_scheme, Some(None));
    }

    #[test]
    fn long_tail_font_integer_bounds_are_checked_before_edit() {
        let error = admit_properties(&Props::new().with("font.charset", 256))
            .expect_err("charset upper bound");
        assert_eq!(error.code(), Code::InvalidValue);

        let error =
            admit_properties(&Props::new().with("font.family", 6)).expect_err("family upper bound");
        assert_eq!(error.code(), Code::InvalidValue);

        let error = admit_properties(&Props::new().with("font.scheme", "custom"))
            .expect_err("scheme token");
        assert_eq!(error.code(), Code::InvalidValue);
    }

    #[test]
    fn alignment_properties_use_native_domains_and_officecli_aliases() {
        let props = Props::new()
            .with("align", "center")
            .with("valign", "bottom")
            .with("wrap", true)
            .with("rotation", 180)
            .with("indent", 255)
            .with("shrink", false)
            .with("direction", "rtl");
        let operation = admit_properties(&props).expect("alignment properties");
        assert_eq!(
            operation.alignment_horizontal,
            Some(Some(xlsx::style::stylesheet::alignment::Horizontal::Center))
        );
        assert_eq!(
            operation.alignment_vertical,
            Some(Some(xlsx::style::stylesheet::alignment::Vertical::Bottom))
        );
        assert_eq!(operation.alignment_wrap_text, Some(Some(true)));
        assert_eq!(
            operation.alignment_text_rotation,
            Some(Some(
                xlsx::style::stylesheet::alignment::Rotation::degrees(180).expect("rotation")
            ))
        );
        assert_eq!(
            operation.alignment_indent,
            Some(Some(xlsx::style::stylesheet::alignment::Indent::new(255)))
        );
        assert_eq!(operation.alignment_shrink_to_fit, Some(Some(false)));
        assert_eq!(
            operation.alignment_reading_order,
            Some(Some(
                xlsx::style::stylesheet::alignment::Reading::RightToLeft
            ))
        );

        let context = admit_properties(&Props::new().with("readingOrder", "context"))
            .expect("context reading order");
        assert_eq!(context.alignment_reading_order, Some(None));
    }

    #[test]
    fn alignment_rejects_transitional_rotation_and_bad_integer_domains() {
        let error = admit_properties(&Props::new().with("textRotation", 254))
            .expect_err("contextual rotation is transitional");
        assert_eq!(error.code(), Code::InvalidValue);

        let error =
            admit_properties(&Props::new().with("textRotation", 181)).expect_err("rotation gap");
        assert_eq!(error.code(), Code::InvalidValue);

        let error =
            admit_properties(&Props::new().with("indent", 256)).expect_err("indent upper bound");
        assert_eq!(error.code(), Code::InvalidValue);

        let error = admit_properties(&Props::new().with("alignment.horizontal", "middle"))
            .expect_err("horizontal token");
        assert_eq!(error.code(), Code::InvalidValue);
    }

    #[test]
    fn relative_indent_uses_the_signed_schema_domain() {
        let operation = admit_properties(
            &Props::new()
                .with("alignment.relativeIndent", i32::MIN)
                .with("alignment.justifyLastLine", true),
        )
        .expect("signed relative indent");
        assert_eq!(operation.alignment_relative_indent, Some(Some(i32::MIN)));
        assert_eq!(operation.alignment_justify_last_line, Some(Some(true)));

        let operation =
            admit_properties(&Props::new().with("alignment.relativeIndent", i64::from(i32::MAX)))
                .expect("maximum signed relative indent");
        assert_eq!(operation.alignment_relative_indent, Some(Some(i32::MAX)));

        let error = admit_properties(
            &Props::new().with("alignment.relativeIndent", i64::from(i32::MAX) + 1),
        )
        .expect_err("values beyond Int32 must be refused");
        assert_eq!(error.code(), Code::InvalidValue);

        let error = admit_properties(
            &Props::new().with("alignment.relativeIndent", i64::from(i32::MIN) - 1),
        )
        .expect_err("values below Int32 must be refused");
        assert_eq!(error.code(), Code::InvalidValue);
    }

    #[test]
    fn protection_and_quote_prefix_support_officecli_aliases_and_clears() {
        let operation = admit_properties(
            &Props::new()
                .with("locked", true)
                .with("formulahidden", false)
                .with("quotePrefix", true),
        )
        .expect("protection and quote-prefix properties");
        assert_eq!(operation.protection_locked, Some(Some(true)));
        assert_eq!(operation.protection_hidden, Some(Some(false)));
        assert_eq!(operation.quote_prefix, Some(Some(true)));

        let clears = Props::new()
            .with("protection.locked", crate::props::PropValue::Null)
            .with("protection.hidden", crate::props::PropValue::Null)
            .with("quotePrefix", crate::props::PropValue::Null)
            .with("alignment.justifyLastLine", crate::props::PropValue::Null);
        let operation = admit_properties(&clears).expect("null removes authored attributes");
        assert_eq!(operation.protection_locked, Some(None));
        assert_eq!(operation.protection_hidden, Some(None));
        assert_eq!(operation.quote_prefix, Some(None));
        assert_eq!(operation.alignment_justify_last_line, Some(None));

        let error = admit_properties(
            &Props::new()
                .with("locked", true)
                .with("protection.locked", false),
        )
        .expect_err("canonical and alias duplicates must be refused");
        assert_eq!(error.code(), Code::InvalidArgument);
    }

    #[test]
    fn number_formats_normalize_friendly_aliases_and_clear() {
        let operation = admit_properties(
            &Props::new()
                .with("numfmt", "  PeRcEnT  ")
                .with("format", "date"),
        )
        .expect_err("two number-format keys must conflict");
        assert_eq!(operation.code(), Code::InvalidArgument);

        let operation = admit_properties(&Props::new().with("numfmt", "  PeRcEnT  "))
            .expect("friendly number-format alias");
        assert!(matches!(
            operation.number_format.as_ref(),
            Some(Some(value)) if value.as_ref() == "0.00%"
        ));

        let operation = admit_properties(&Props::new().with("format", "0.000"))
            .expect("custom number-format code");
        assert!(matches!(
            operation.number_format.as_ref(),
            Some(Some(value)) if value.as_ref() == "0.000"
        ));

        let operation =
            admit_properties(&Props::new().with("numberformat", crate::props::PropValue::Null))
                .expect("number-format clear");
        assert_eq!(operation.number_format, Some(None));
    }

    #[test]
    fn number_formats_reject_malformed_codes_before_native_edit() {
        for code in ["0\"", "0;0;0;0;0", "[Red"] {
            let error = admit_properties(&Props::new().with("numberformat", code))
                .expect_err("malformed number format");
            assert_eq!(error.code(), Code::InvalidValue);
        }

        let error = admit_properties(&Props::new().with("numberformat", "a\u{1}b"))
            .expect_err("XML-forbidden number-format text");
        assert_eq!(error.code(), Code::InvalidValue);

        let error = admit_properties(&Props::new().with("numberformat", "0".repeat(256)))
            .expect_err("number-format length limit");
        assert_eq!(error.code(), Code::InvalidValue);
    }

    #[test]
    fn fill_patterns_normalize_officecli_separator_and_gray_aliases() {
        let operation = admit_properties(&Props::new().with("fillPattern", " d-a rk Horizontal "))
            .expect("separator-tolerant fill pattern");
        assert_eq!(
            operation.fill_pattern,
            Some(Some(xlsx::style::FillPattern::DarkHorizontal))
        );

        let operation =
            admit_properties(&Props::new().with("fillPattern", "GRAY50")).expect("gray alias");
        assert_eq!(
            operation.fill_pattern,
            Some(Some(xlsx::style::FillPattern::MediumGray))
        );
    }

    #[test]
    fn fill_accepts_bounded_gradients_and_preserves_negative_tints_as_colors() {
        let operation = admit_properties(&Props::new().with("fill", "FF0000-0000FF-90"))
            .expect("linear gradient");
        assert!(operation.fill_gradient.is_some());
        let gradient = parse_gradient_fill("FF0000-0000FF-90", "fill", &SemanticLimits::default())
            .expect("parsed gradient");
        assert_eq!(
            gradient.kind(),
            xlsx::style::cell_format::GradientKind::Linear
        );
        assert_eq!(gradient.degree(), 90.0);
        assert_eq!(gradient.stops().len(), 2);

        let gradient = parse_gradient_fill(
            "gradient;FF0000;00FF00;0000FF;180",
            "fill",
            &SemanticLimits::default(),
        )
        .expect("semicolon multi-stop gradient");
        assert_eq!(gradient.degree(), 180.0);
        assert_eq!(gradient.stops().len(), 3);

        let gradient = parse_gradient_fill(
            "accent1+tint-25-#00FF00-#0000FF-90",
            "fill",
            &SemanticLimits::default(),
        )
        .expect("negative tint in a multi-stop gradient");
        assert_eq!(gradient.stops().len(), 3);

        for value in [
            "accent1+tint=-25-#00FF00-#0000FF-90",
            "accent1 + tint-25-#00FF00-#0000FF-90",
            "accent1 + tint-25  -#00FF00-#0000FF-90",
            "accent1+tint-1e-1-#00FF00-#0000FF-90",
        ] {
            let gradient = parse_gradient_fill(value, "fill", &SemanticLimits::default())
                .expect("shared SpreadsheetML tint spelling in a gradient stop");
            assert_eq!(gradient.stops().len(), 3, "{value:?}");
        }

        let integer_angle =
            parse_gradient_fill("FF0000-0000FF-180", "fill", &SemanticLimits::default())
                .expect("three-digit numeric suffix is an angle");
        assert_eq!(integer_angle.degree(), 180.0);
        assert_eq!(integer_angle.stops().len(), 2);
        let fractional =
            parse_gradient_fill("FF0000-0000FF-90.5", "fill", &SemanticLimits::default())
                .expect("fractional linear angle");
        assert_eq!(fractional.degree(), 90.5);
        let decimal_integer =
            parse_gradient_fill("FF0000-0000FF-360.0", "fill", &SemanticLimits::default())
                .expect("decimal integer linear angle");
        assert_eq!(decimal_integer.degree(), 360.0);
        let numeric_rgb =
            parse_gradient_fill("FF0000-0000FF-123456", "fill", &SemanticLimits::default())
                .expect("six-digit numeric RGB stop");
        assert_eq!(numeric_rgb.stops().len(), 3);

        let operation = admit_properties(&Props::new().with("fill", "radial:#FF0000-#0000FF"))
            .expect("radial gradient");
        assert!(operation.fill_gradient.is_some());
        let radial =
            parse_gradient_fill("radial:#FF0000-#0000FF", "fill", &SemanticLimits::default())
                .expect("parsed radial gradient");
        assert_eq!(radial.kind(), xlsx::style::cell_format::GradientKind::Path);

        let operation = admit_properties(&Props::new().with("fill", "accent1+tint-25"))
            .expect("negative tint color");
        assert!(operation.fill.is_some());
        assert!(operation.fill_gradient.is_none());
    }

    #[test]
    fn fill_rejects_ambiguous_gradient_combinations_and_malformed_stops() {
        for value in [
            "gradient;FF0000;0000FF",
            "FF0000-0000FF-90-10",
            "FF0000-0000FF-361",
        ] {
            let error = admit_properties(&Props::new().with("fill", value))
                .expect_err("invalid gradient spelling");
            assert_eq!(error.code(), Code::InvalidValue);
        }

        let error = admit_properties(
            &Props::new()
                .with("fill", "FF0000-0000FF-90")
                .with("fillPattern", "solid"),
        )
        .expect_err("gradient and pattern are mutually exclusive");
        assert_eq!(error.code(), Code::InvalidValue);
    }

    #[test]
    fn borders_admit_shorthands_side_aliases_and_tri_state_values() {
        let operation = admit_properties(
            &Props::new()
                .with("border", "thin")
                .with("border.color", "accent1+tint-25")
                .with("border.left.style", "double")
                .with("border.left.color", "#FF0000")
                .with("border.diagonalUp", true)
                .with("border.diagonalDown", crate::props::PropValue::Null),
        )
        .expect("border assignments");
        assert_eq!(
            operation.border_all,
            Some(Some(xlsx::style::cell_format::Line::Thin))
        );
        assert!(
            operation
                .border_all_color
                .is_some_and(|color| color.is_some())
        );
        assert_eq!(
            operation.border_left,
            Some(Some(xlsx::style::cell_format::Line::Double))
        );
        assert!(
            operation
                .border_left_color
                .is_some_and(|color| color.is_some())
        );
        assert_eq!(operation.border_diagonal_up, Some(Some(true)));
        assert_eq!(operation.border_diagonal_down, Some(None));
    }

    #[test]
    fn borders_reject_duplicate_aliases_and_color_without_a_line() {
        let error = admit_properties(
            &Props::new()
                .with("border.left", "none")
                .with("border.left.color", "#FF0000"),
        )
        .expect_err("color on a cleared side");
        assert_eq!(error.code(), Code::InvalidValue);

        let error = admit_properties(
            &Props::new()
                .with("border.left", "thin")
                .with("border.left.style", "double"),
        )
        .expect_err("duplicate side aliases");
        assert_eq!(error.code(), Code::InvalidArgument);
    }
}
