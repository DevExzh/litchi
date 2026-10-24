//! Immutable XLSB sheet-property writes and publication.
//!
//! The XLSB unified writes admit sheet names, tab visibility and direct cell scalar
//! payloads (`value`, `clear: true`, and null clears). Target resolution and
//! property coercion happen before the native workbook editor is materialized.
//! The native `Package` transaction remains responsible for exact package
//! replay, binary validation, signature policy, and retained patch proofs.

use std::collections::{BTreeMap, HashSet};
use std::io::{self, Write};
use std::mem::size_of;
use std::path::Path as FilePath;

use caseless::Caseless as _;
use litchi_core::selector::{CellRange, Path, Predicate, Step};
use litchi_core::{FileFormat, Position};
use litchi_xlsb::cell_values::structure::Edit as NativeStructureEdit;
pub(super) use litchi_xlsb::cell_values::structure::EditOptions as NativeStructureOptions;
use litchi_xlsb::cell_values::{BorderLine as NativeBorderLine, Format as NativeFormat};
use litchi_xlsb::cell_values::{
    Edit as NativeCellEdit, EditOptions as NativeCellOptions, Patch as NativeCellPatch, Reference,
    Scalar as NativeScalar, TransferLimits, structure::Patch as NativeStructurePatch,
};
use litchi_xlsb::sheet::{Rect, Selector as NativeSelector, State as SheetState};
use unicode_normalization::UnicodeNormalization;

use super::node::{Item, Node};
use super::spec::{CELL_WRITE_TABLE, CellSetter, Kind, SHEET_WRITE_TABLE, SheetSetter};
use super::{Nav, Xlsb, XlsbEditor, XlsbOwner};
use crate::error::{Code, Error, Result};
use crate::new_node::NewNode;
use crate::path::PathArg;
use crate::placement::Placement;
use crate::props::{IntoProps, Props};
use crate::registry::Operation;
use crate::save::{Overwrite, SaveOptions};
use crate::target::TargetArg;
use crate::transaction::{Error as TransactionError, LimitKind, SemanticLimits};
use crate::value::Scalar;

/// The native proof retained by one immutable XLSB write.
///
/// Structural sheet edits and cell-value edits have different native patch
/// owners. Keeping the distinction here lets the transaction coordinator
/// retain exact before/after package images without treating a cell patch as
/// a structural worksheet patch.
#[derive(Debug, Clone)]
pub(crate) enum Patch {
    Structure {
        patch: NativeStructurePatch,
        retained_bytes: usize,
    },
    Cells(Vec<NativeCellPatch>),
    /// A document-metadata patch with its canonical package images (record
    /// 0253).
    Metadata(Box<super::metadata::Images>),
    /// A worksheet-feature patch with its canonical package images (record
    /// 0322).
    Feature(Box<super::feature_write::Images>),
}

impl Patch {
    pub(crate) fn before(&self) -> &[u8] {
        match self {
            Self::Structure { patch, .. } => patch.before(),
            Self::Cells(patches) => patches
                .first()
                .expect("cell patch invariant: at least one worksheet")
                .before(),
            Self::Metadata(images) => images.before(),
            Self::Feature(images) => images.before(),
        }
    }

    pub(crate) fn after(&self) -> &[u8] {
        match self {
            Self::Structure { patch, .. } => patch.after(),
            Self::Cells(patches) => patches
                .last()
                .expect("cell patch invariant: at least one worksheet")
                .after(),
            Self::Metadata(images) => images.after(),
            Self::Feature(images) => images.after(),
        }
    }

    /// Exact native history charge. Structural patches cache the checked
    /// native calculation at commit construction so transaction diagnostics
    /// remain infallible; cell patches retain their native semantic payloads.
    pub(crate) fn retained_bytes(&self) -> usize {
        match self {
            Self::Metadata(images) => images.retained_bytes(),
            Self::Feature(images) => images.retained_bytes(),
            Self::Structure { retained_bytes, .. } => *retained_bytes,
            Self::Cells(patches) => patches
                .iter()
                .fold(size_of::<Vec<NativeCellPatch>>(), |bytes, patch| {
                    bytes.saturating_add(patch.retained_bytes())
                }),
        }
    }

    pub(crate) fn is_noop(&self) -> bool {
        match self {
            Self::Metadata(images) => images.patch.is_noop(),
            Self::Feature(images) => images.is_noop(),
            _ => self.before() == self.after(),
        }
    }

    #[allow(
        dead_code,
        reason = "structure and feature history adapters use their own authored counts"
    )]
    pub(crate) fn operation_count(&self) -> usize {
        match self {
            Self::Metadata(images) => images.patch.changes().len(),
            Self::Feature(images) => images.operation_count(),
            Self::Structure { patch, .. } => patch.operation_count(),
            Self::Cells(patches) => patches.iter().fold(0usize, |count, patch| {
                count.saturating_add(patch.changes().len())
            }),
        }
    }
}

/// A committed immutable XLSB write.
pub(crate) struct Commit {
    after: Xlsb,
    patch: Patch,
    /// Authored leaf operations, kept separately from the canonical native
    /// patch. A signed cancellation is represented by a true empty patch, so
    /// deriving this from `patch` would lose the operation diagnostics.
    operations: usize,
    /// Peak retained native proof while a private changed patch was reduced
    /// to the canonical signed identity patch.
    snapshot_bytes: usize,
}

impl Commit {
    /// Consume the result for the format-neutral one-shot setter.
    pub(crate) fn into_file(self) -> Xlsb {
        self.after
    }

    /// Adapt a metadata commit (record 0253).
    pub(super) fn from_metadata(after: Xlsb, images: super::metadata::Images) -> Self {
        let operations = images.authored_operations;
        let snapshot_bytes = images.retained_bytes();
        Self {
            after,
            patch: Patch::Metadata(Box::new(images)),
            operations,
            snapshot_bytes,
        }
    }

    /// Adapt a worksheet-feature commit (record 0322).
    pub(super) fn from_feature(after: Xlsb, images: super::feature_write::Images) -> Self {
        let operations = images.operation_count();
        let snapshot_bytes = images.retained_bytes();
        Self {
            after,
            patch: Patch::Feature(Box::new(images)),
            operations,
            snapshot_bytes,
        }
    }

    /// The retained native proof.
    #[allow(dead_code, reason = "reserved for format-local patch inspection")]
    pub(super) const fn patch(&self) -> &Patch {
        &self.patch
    }

    pub(super) fn is_noop(&self) -> bool {
        self.patch.is_noop()
    }

    pub(crate) const fn operations(&self) -> usize {
        self.operations
    }

    pub(crate) const fn snapshot_bytes(&self) -> usize {
        self.snapshot_bytes
    }

    /// Consume the result for explicit transaction history.
    pub(crate) fn into_parts(self) -> (Xlsb, Patch) {
        (self.after, self.patch)
    }

    pub(super) fn from_native(
        base: &Xlsb,
        native: litchi_xlsb::cell_values::structure::Commit,
        max_snapshot_bytes: usize,
    ) -> Result<Self> {
        base.check_source_for_replay()?;
        let result = (|| {
            let (package, patch) = native.into_parts();
            let operations = patch.operation_count();
            let signed = base.signature() == super::Signature::Signed;
            // Structure patches already retain their checked native history
            // charge. Keep that charge for unsigned publication as well: the
            // transaction diagnostics and snapshot limit must not silently
            // report a zero-sized retained proof.
            let native_snapshot_bytes = patch.retained_bytes().map_err(native_error)?;
            if native_snapshot_bytes > max_snapshot_bytes {
                return Err(snapshot_limit(native_snapshot_bytes, max_snapshot_bytes));
            }
            let (package, patch, retained_bytes) = if patch.is_noop() && signed {
                // Structural writes are staged through an unsigned private
                // package. Rebuild the identity patch over the retained
                // signed package so the leaf and facade share one base.
                let original = base.original_package_for_noop()?;
                let (package, patch) = original
                    .edit()
                    .map_err(native_error)?
                    .commit()
                    .map_err(native_error)?
                    .into_parts();
                let retained_bytes = patch.retained_bytes().map_err(native_error)?;
                (package, patch, retained_bytes)
            } else {
                (package, patch, native_snapshot_bytes)
            };
            let snapshot_bytes = native_snapshot_bytes.max(retained_bytes);
            if snapshot_bytes > max_snapshot_bytes {
                return Err(snapshot_limit(snapshot_bytes, max_snapshot_bytes));
            }
            // A native edit may retain a semantic operation whose replay is
            // an exact image no-op. Preserve the original immutable handle in
            // that case so source-backed publication never canonicalizes a
            // byte-identical no-op through a new package owner.
            let after = if patch.is_noop() {
                base.clone()
            } else {
                base.adopt_package(package)?
            };
            Ok(Self {
                after,
                patch: Patch::Structure {
                    patch,
                    retained_bytes,
                },
                operations,
                snapshot_bytes,
            })
        })();
        base.check_source_for_replay()?;
        result
    }

    fn from_cell_native(
        base: &Xlsb,
        native: litchi_xlsb::cell_values::Commit,
        sheet: usize,
        limits: TransferLimits,
        max_snapshot_bytes: usize,
        authored_operations: usize,
    ) -> Result<Self> {
        base.check_source_for_replay()?;
        let result = (|| {
            let (package, patch) = native.into_parts();
            let signed = base.signature() == super::Signature::Signed;
            let native_patch_bytes = patch.retained_bytes();
            if native_patch_bytes > limits.bytes() {
                return Err(snapshot_limit(native_patch_bytes, limits.bytes()));
            }
            let native_snapshot_bytes =
                native_patch_bytes.saturating_add(size_of::<Vec<NativeCellPatch>>());
            if native_snapshot_bytes > max_snapshot_bytes {
                return Err(snapshot_limit(native_snapshot_bytes, max_snapshot_bytes));
            }
            let (patch, retained_bytes, package) = if patch.is_noop() && signed {
                // The staged package and proof are discarded before building
                // the canonical identity patch.  The peak is retained as a
                // scalar below, rather than by keeping two package owners
                // alive during canonical construction.
                drop(package);
                drop(patch);
                let patch = canonical_cell_patch(base, sheet, limits)?;
                let retained_bytes = patch.retained_bytes();
                (patch, retained_bytes, None)
            } else {
                (
                    Patch::Cells(vec![patch]),
                    native_snapshot_bytes,
                    Some(package),
                )
            };
            let after = if patch.is_noop() {
                base.clone()
            } else {
                let package = package.ok_or_else(|| {
                    Error::new(
                        Code::InternalError,
                        "a changed XLSB cell patch lost its package owner",
                    )
                    .with_format(FileFormat::Xlsb)
                })?;
                base.adopt_package(package)?
            };
            let snapshot_bytes = native_snapshot_bytes.max(retained_bytes);
            if snapshot_bytes > max_snapshot_bytes {
                return Err(snapshot_limit(snapshot_bytes, max_snapshot_bytes));
            }
            Ok(Self {
                after,
                patch,
                // Native cell patches count physical deltas.  Keep the
                // authored count separately so a same-value set/remove (and
                // the canonical empty patch used for a signed replay) still
                // reports the operation the caller actually wrote.
                operations: authored_operations,
                snapshot_bytes,
            })
        })();
        base.check_source_for_replay()?;
        result
    }

    fn from_cell_parts(
        base: &Xlsb,
        package: litchi_xlsb::Package,
        patches: Vec<NativeCellPatch>,
        canonical: (usize, TransferLimits),
        native_snapshot_bytes: usize,
        max_snapshot_bytes: usize,
        authored_operations: usize,
    ) -> Result<Self> {
        if patches.is_empty() {
            return Err(Error::new(
                Code::InternalError,
                "an XLSB cell commit must retain at least one worksheet patch",
            )
            .with_format(FileFormat::Xlsb));
        }
        base.check_source_for_replay()?;
        let result = (|| {
            let native_patch = Patch::Cells(patches);
            let signed = base.signature() == super::Signature::Signed;
            if native_snapshot_bytes > max_snapshot_bytes {
                return Err(snapshot_limit(native_snapshot_bytes, max_snapshot_bytes));
            }
            let all_noop = match &native_patch {
                Patch::Cells(patches) => patches.iter().all(NativeCellPatch::is_noop),
                _ => false,
            };
            let (patch, retained_bytes, package) = if signed && all_noop {
                // Do not retain the private staged package/proof while the
                // original-owner identity patch is being constructed.
                drop(native_patch);
                drop(package);
                let patch = canonical_cell_patch(base, canonical.0, canonical.1)?;
                let retained_bytes = patch.retained_bytes();
                (patch, retained_bytes, None)
            } else {
                (native_patch, native_snapshot_bytes, Some(package))
            };
            let after = if patch.is_noop() {
                base.clone()
            } else {
                let package = package.ok_or_else(|| {
                    Error::new(
                        Code::InternalError,
                        "a changed XLSB cell patch lost its package owner",
                    )
                    .with_format(FileFormat::Xlsb)
                })?;
                base.adopt_package(package)?
            };
            let snapshot_bytes = native_snapshot_bytes.max(retained_bytes);
            if snapshot_bytes > max_snapshot_bytes {
                return Err(snapshot_limit(snapshot_bytes, max_snapshot_bytes));
            }
            Ok(Self {
                after,
                patch,
                operations: authored_operations,
                snapshot_bytes,
            })
        })();
        base.check_source_for_replay()?;
        result
    }
}

fn canonical_cell_patch(base: &Xlsb, sheet: usize, limits: TransferLimits) -> Result<Patch> {
    let original = base.original_package_for_noop()?;
    let edit = original
        .edit_cell_values_with_options(
            NativeSelector::from(sheet),
            NativeCellOptions::new().transfer_limits(limits),
        )
        .map_err(native_error)?
        .resolve()
        .map_err(native_error)?
        .ok_or_else(|| {
            Error::new(
                Code::UnsupportedCapability,
                "the signed no-op cell patch has no worksheet owner",
            )
            .with_format(FileFormat::Xlsb)
        })?;
    let (_, patch) = edit.commit().map_err(native_error)?.into_parts();
    Ok(Patch::Cells(vec![patch]))
}

#[derive(Debug, Default)]
struct SheetAssignments {
    name: Option<String>,
    state: Option<SheetState>,
}

/// Whether a setter is eligible for the narrow direct-path visibility run.
/// Names, combined assignments, selectors, and empty property sets stay on
/// the ordinary structural path so their existing resolution and policy
/// ordering remains unchanged.
pub(crate) fn sheet_state_shape(target: &TargetArg<'_>, props: &Props<'_>) -> bool {
    let Some(path) = target.path() else {
        return false;
    };
    if !path.path().is_absolute() || path.path().steps().len() != 1 || props.len() != 1 {
        return false;
    }
    let step = &path.path().steps()[0];
    if step.name() != "sheet" || step.predicates().len() != 1 || step.index().is_none() {
        return false;
    }
    let Some((key, _)) = props.iter().next() else {
        return false;
    };
    let Some((_, entry)) = SHEET_WRITE_TABLE.lookup(key.as_str()) else {
        return false;
    };
    matches!(
        entry.access,
        crate::registry::Access::Set(SheetSetter::Hidden | SheetSetter::Visibility)
            | crate::registry::Access::SetAndAdd(SheetSetter::Hidden | SheetSetter::Visibility)
    )
}

#[derive(Clone, Copy)]
struct SheetTarget {
    tab: usize,
}

const WRITE_ENTRY_OVERHEAD: usize = 64;

/// The XLSB one-shot setter entry point.
pub(crate) fn set_with_limits<'p, 'v>(
    file: &Xlsb,
    target: TargetArg<'p>,
    props: impl IntoProps<'v>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    file.admit_edit_for_write()?;
    let props = Props::from_props(props)?;
    let result = (|| {
        if let Some(commit) = file.try_set_metadata(&target, &props, limits)? {
            return Ok(commit);
        }
        if let Some(commit) = file.try_set_feature(&target, &props, limits)? {
            return Ok(commit);
        }
        if let Some(commit) =
            super::merge_write::set_merges_with_limits(file, &target, &props, limits)?
        {
            return Ok(commit);
        }
        if let Some(commit) = super::name_write::try_set_named_range(file, &target, &props, limits)?
        {
            return Ok(commit);
        }
        if let Some(commit) = set_cells_with_limits(file, &target, &props, limits)? {
            return Ok(commit);
        }
        set_sheet_with_limits(file, target, &props, limits)
    })();
    result.and_then(|commit| {
        let snapshot_bytes = commit.snapshot_bytes();
        if snapshot_bytes > limits.max_snapshot_bytes() {
            return Err(snapshot_limit(snapshot_bytes, limits.max_snapshot_bytes()));
        }
        Ok(commit)
    })
}

/// Add one worksheet tab through the native workbook structure editor.
///
/// The facade resolves the parent and placement against the complete tab
/// catalog, admits the shared sheet property registry, and then delegates the
/// package rewrite to the XLSB leaf.  Keeping insertion in the native edit
/// preserves workbook relationships, tab identity and active-view policy.
pub(crate) fn add_with_limits<'p, 'n>(
    file: &Xlsb,
    parent: PathArg<'p>,
    node: NewNode<'n>,
    limits: &SemanticLimits,
) -> Result<(Commit, Path<'static>)> {
    file.admit_edit()?;
    let parent = crate::path::resolve::get(&Nav::new(file), &parent)?;
    if matches!(parent.element(), Kind::Sheet | Kind::Cell | Kind::Range)
        && let Some((kind, implied)) = super::feature_write::feature_kind(node.kind(), node.props())
    {
        return file.add_feature(&parent, kind, implied, &node, limits.max_snapshot_bytes());
    }
    if parent.element() != Kind::Workbook {
        return Err(Error::new(
            Code::UnsupportedType,
            "an XLSB sheet can only be added directly under the workbook",
        )
        .with_format(FileFormat::Xlsb));
    }
    if node.kind().eq_ignore_ascii_case("namedrange") {
        return super::name_write::add_named_range(file, &node, limits);
    }
    if !node.kind().eq_ignore_ascii_case("sheet") {
        return Err(Error::new(
            Code::UnsupportedElement,
            "the XLSB structural verb supports worksheet sheets and defined names only",
        )
        .with_format(FileFormat::Xlsb));
    }

    let mut assignments = admit_sheet_properties(Operation::Add, node.props(), limits)?;
    let count = file.tab_count()?;
    let position = add_position(file, count, node.placement())?;
    let name = match assignments.name.take() {
        Some(name) => name,
        None => default_sheet_name(file)?,
    };
    validate_sheet_name(file, &[], &name)?;
    let state = assignments.state.unwrap_or(SheetState::Visible);
    validate_new_sheet_state(file, state)?;
    // The native operation is one semantic insertion. Include the generated
    // default name in the staged-byte admission just like an authored name.
    assignments.name = Some(name.clone());
    check_limits(1, &assignments, limits)?;

    let package = package_from_editor(file.editor_for_write()?);
    let native_limits = TransferLimits::new(
        limits.max_snapshot_bytes(),
        limits.max_expanded_targets(),
        1,
        limits.max_snapshot_bytes(),
    );
    let mut edit = package
        .edit_with_options(NativeStructureOptions::new().transfer_limits(native_limits))
        .map_err(native_error)?;
    edit.insert_sheet(Position::new(position), name, state)
        .map_err(native_error)?;
    let native = edit.commit().map_err(native_error)?;
    let commit = Commit::from_native(file, native, limits.max_snapshot_bytes())?;
    Ok((commit, sheet_path(position)))
}

fn add_position(file: &Xlsb, count: usize, placement: Option<&Placement<'_>>) -> Result<usize> {
    let position = match placement {
        None | Some(Placement::End) => count,
        Some(Placement::Start) => 0,
        Some(Placement::Position(position)) => position.get(),
        Some(Placement::Before(path)) => resolve_sheet_path(file, path)?,
        Some(Placement::After(path)) => resolve_sheet_path(file, path)?.saturating_add(1),
        Some(Placement::Into(path, position)) => {
            let parent = crate::path::resolve::get(&Nav::new(file), path)?;
            if parent.element() != Kind::Workbook {
                return Err(Error::new(
                    Code::UnsupportedType,
                    "XLSB sheets can only be added within the workbook",
                )
                .with_format(FileFormat::Xlsb));
            }
            position.map_or(count, Position::get)
        },
    };
    if position > count {
        return Err(Error::new(
            Code::OutOfRange,
            "the XLSB sheet insertion position is outside the workbook",
        )
        .with_format(FileFormat::Xlsb));
    }
    Ok(position)
}

fn default_sheet_name(file: &Xlsb) -> Result<String> {
    let catalog = file.catalog()?;
    let mut names = HashSet::new();
    names
        .try_reserve(catalog.len())
        .map_err(|_| allocation("XLSB existing sheet-name keys"))?;
    for position in 0..catalog.len() {
        if let Some(tab) = catalog.get(position) {
            names.insert(caseless_key(tab.name()));
        }
    }
    let mut number = 1usize;
    loop {
        let name = format!("Sheet{number}");
        let key = caseless_key(&name);
        if !names.contains(&key) {
            return Ok(name);
        }
        number = number.checked_add(1).ok_or_else(|| {
            Error::new(
                Code::OutOfRange,
                "the XLSB default sheet name is outside its supported range",
            )
            .with_format(FileFormat::Xlsb)
        })?;
    }
}

fn validate_new_sheet_state(file: &Xlsb, state: SheetState) -> Result<()> {
    if state.is_visible() {
        return Ok(());
    }
    let catalog = file.catalog()?;
    if (0..catalog.len()).any(|position| {
        catalog
            .get(position)
            .is_some_and(|tab| tab.state().is_visible())
    }) {
        return Ok(());
    }
    Err(Error::new(
        Code::InvalidArgument,
        "a workbook must retain a visible sheet",
    )
    .with_format(FileFormat::Xlsb))
}

fn sheet_path(position: usize) -> Path<'static> {
    Path::from_steps([Step::new("sheet").with(Predicate::Index(Position::new(position)))])
}

/// Remove a complete-catalog sheet selected by a direct sheet path or
/// homogeneous sheet query, or stored cells selected by a direct cell path,
/// range, or scoped query. The target expansion and worksheet grouping are
/// shared with cell setters, while the native cell-values owner supplies the
/// deletion semantics: a virtual (unstored) cell is an exact no-op and no row
/// or column shifting is attempted.
pub(crate) fn remove_with_limits(
    file: &Xlsb,
    target: TargetArg<'_>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    // The facade performs the signed admission fence and rejects changed
    // output after native staging.  Keep the native path open long enough to
    // prove a real identity patch for signed no-op removals (for example an
    // unstored cell or an identity sheet selector).
    file.admit_edit_for_write()?;
    if let Some(commit) = file.try_remove_feature(&target, limits)? {
        return Ok(commit);
    }
    if let Some(commit) = super::name_write::try_remove_named_range(file, &target, limits)? {
        return Ok(commit);
    }
    if let Some(targets) = resolve_sheet_removal_targets(file, &target, limits)? {
        return remove_sheets_with_limits(file, targets, limits);
    }
    if let Some(range) = target
        .path()
        .and_then(|path| path.path().steps().last())
        .and_then(|step| step.cell_range())
    {
        check_remove_limits(range_area(range), limits)?;
    }
    let Some(targets) = resolve_cell_targets(file, &target, limits)? else {
        return Err(Error::new(
            Code::UnsupportedElement,
            "an XLSB removal target must be a cell, range, or cell query",
        )
        .with_format(FileFormat::Xlsb));
    };
    check_remove_limits(targets.len(), limits)?;

    // Complete target resolution and grouping happen before the native
    // package owner is materialized.  One native edit covers each worksheet;
    // a cross-worksheet selector publishes those edits in tab order through
    // one immutable package chain.
    let authored_operations = targets.len();
    let mut grouped = BTreeMap::<usize, Vec<CellTarget>>::new();
    for target in targets {
        grouped
            .entry(target.sheet)
            .or_default()
            .try_reserve(1)
            .map_err(|_| allocation("XLSB cell removal worksheet targets"))?;
        grouped
            .get_mut(&target.sheet)
            .expect("entry was just inserted")
            .push(target);
    }
    let editor = file.editor_for_write()?;
    let mut package = package_from_editor(editor);
    let mut patches = Vec::new();
    patches
        .try_reserve(grouped.len())
        .map_err(|_| allocation("XLSB cell removal patches"))?;
    let mut canonical = None;
    // `Patch::Cells` retains one vector owner in addition to each native
    // patch. Reserve that exact container charge before any native editor is
    // materialized; the native patch itself supplies the complete semantic
    // cell-delta charge.
    let mut remaining_snapshot = cell_snapshot_budget(limits.max_snapshot_bytes())?;
    let mut retained_snapshot = cell_patch_container_bytes();
    for (sheet, targets) in grouped {
        let native_limits = cell_transfer_limits(remaining_snapshot, limits)?;
        canonical.get_or_insert((sheet, native_limits));
        let mut edit = package
            .edit_cell_values_with_options(
                NativeSelector::from(sheet),
                NativeCellOptions::new().transfer_limits(native_limits),
            )
            .map_err(native_error)?
            .resolve()
            .map_err(native_error)?
            .ok_or_else(|| {
                Error::new(
                    Code::OutOfRange,
                    "the XLSB worksheet position is out of range",
                )
                .with_format(FileFormat::Xlsb)
            })?;
        for target in targets {
            // The native owner treats an absent or virtual cell as an
            // idempotent no-op.  This is deliberately a cell-value removal;
            // it never shifts neighboring cells or edits row/column layout.
            edit.remove(target.reference).map_err(native_error)?;
        }
        let native = edit.commit().map_err(native_error)?;
        let (next, patch) = native.into_parts();
        let retained = patch.retained_bytes();
        if retained > remaining_snapshot {
            let consumed = limits
                .max_snapshot_bytes()
                .saturating_sub(remaining_snapshot);
            return Err(snapshot_limit(
                consumed.saturating_add(retained),
                limits.max_snapshot_bytes(),
            ));
        }
        remaining_snapshot -= retained;
        retained_snapshot = retained_snapshot
            .checked_add(retained)
            .ok_or_else(|| snapshot_limit(usize::MAX, limits.max_snapshot_bytes()))?;
        package = next;
        patches.push(patch);
    }
    let canonical = canonical.ok_or_else(|| {
        Error::new(
            Code::InternalError,
            "XLSB cell removal lost its worksheet owner",
        )
        .with_format(FileFormat::Xlsb)
    })?;
    let commit = Commit::from_cell_parts(
        file,
        package,
        patches,
        canonical,
        retained_snapshot,
        limits.max_snapshot_bytes(),
        authored_operations,
    )?;
    Ok(commit)
}

/// Remove complete-catalog tabs in one native structural edit.
///
/// The facade resolves every target against the immutable source catalog before
/// creating an editor.  Native selectors therefore keep their source ordinal
/// even when more than one tab is staged, and a failed later selector cannot
/// expose a partially published package.  The complete catalog includes chart,
/// dialog, macro, and unbound tabs; the leaf owns their dependency and view
/// remapping.
fn remove_sheets_with_limits(
    file: &Xlsb,
    targets: Vec<SheetTarget>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    check_remove_limits(targets.len(), limits)?;
    let catalog = file.catalog()?;
    let count = catalog.len();
    if targets.is_empty() {
        return Err(
            Error::new(Code::ZeroMatches, "the mutation selector matched no sheet")
                .with_format(FileFormat::Xlsb),
        );
    }
    if targets.len() >= count {
        return Err(Error::new(
            Code::InvalidArgument,
            "an XLSB workbook must retain at least one sheet",
        )
        .with_format(FileFormat::Xlsb));
    }

    let mut selected = HashSet::new();
    selected
        .try_reserve(targets.len())
        .map_err(|_| allocation("XLSB sheet removal targets"))?;
    selected.extend(targets.iter().map(|target| target.tab));
    let visible = (0..count).any(|position| {
        !selected.contains(&position)
            && catalog
                .get(position)
                .is_some_and(|tab| tab.state().is_visible())
    });
    if !visible {
        return Err(Error::new(
            Code::InvalidArgument,
            "an XLSB workbook must retain a visible sheet",
        )
        .with_format(FileFormat::Xlsb));
    }

    let editor = file.editor_for_write()?;
    let package = package_from_editor(editor);
    let native_limits = TransferLimits::new(
        limits.max_snapshot_bytes(),
        targets.len(),
        1,
        limits.max_snapshot_bytes(),
    );
    let mut edit = package
        .edit_with_options(NativeStructureOptions::new().transfer_limits(native_limits))
        .map_err(native_error)?;
    for target in targets {
        // The native edit resolves every selector against its immutable source
        // catalog.  Do not translate a later target through an earlier removal.
        let _ = edit
            .remove_sheet(NativeSelector::Position(Position::new(target.tab)))
            .map_err(native_error)?;
    }
    let native = edit.commit().map_err(native_error)?;
    let commit = Commit::from_native(file, native, limits.max_snapshot_bytes())?;
    Ok(commit)
}

/// Reorder one complete-catalog XLSB tab with the shared placement vocabulary.
///
/// The public resolver addresses every tab kind in workbook order.  A native
/// edit receives the resulting permutation in one operation, so relationship
/// identity, active-view state, and opaque tab payloads remain owned by the
/// binary leaf rather than being reconstructed by the facade.
pub(crate) fn move_to_with_limits<'p, 'q>(
    file: &Xlsb,
    path: PathArg<'p>,
    placement: Placement<'q>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    file.admit_edit_for_write()?;
    let source = resolve_sheet_path(file, &path)?;
    let count = file.tab_count()?;
    let destination = move_destination(file, source, count, &placement)?;
    admit_order_targets(1, count, limits)?;
    let order = move_permutation(count, source, destination)?;
    reorder_with_limits(file, order, limits)
}

/// Exchange two complete-catalog XLSB tabs in one native structural edit.
pub(crate) fn swap_with_limits<'p, 'q>(
    file: &Xlsb,
    left: PathArg<'p>,
    right: PathArg<'q>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    file.admit_edit_for_write()?;
    let left = resolve_sheet_path(file, &left)?;
    let right = resolve_sheet_path(file, &right)?;
    let count = file.tab_count()?;
    if left >= count || right >= count {
        return Err(
            Error::new(Code::PathNotFound, "the selected XLSB tab is absent")
                .with_format(FileFormat::Xlsb),
        );
    }
    admit_order_targets(usize::from(left != right).saturating_add(1), count, limits)?;
    let mut order = identity_permutation(count)?;
    order.swap(left, right);
    reorder_with_limits(file, order, limits)
}

fn reorder_with_limits(file: &Xlsb, order: Vec<usize>, limits: &SemanticLimits) -> Result<Commit> {
    let editor = file.editor_for_write()?;
    let package = package_from_editor(editor);
    let native_limits = TransferLimits::new(
        limits.max_snapshot_bytes(),
        limits.max_expanded_targets(),
        1,
        limits.max_snapshot_bytes(),
    );
    let mut edit = package
        .edit_with_options(NativeStructureOptions::new().transfer_limits(native_limits))
        .map_err(native_error)?;
    // `reorder_sheets` borrows semantic selectors for the complete
    // permutation. The native owner validates them against its immutable base
    // catalog before staging, and retains exact no-op patches when the
    // permutation is the identity.
    let mut selectors = Vec::new();
    selectors
        .try_reserve_exact(order.len())
        .map_err(|_| allocation("XLSB sheet-order selectors"))?;
    selectors.extend(
        order
            .into_iter()
            .map(|position| NativeSelector::Position(Position::new(position))),
    );
    edit.reorder_sheets(&selectors).map_err(native_error)?;
    let native = edit.commit().map_err(native_error)?;
    let commit = Commit::from_native(file, native, limits.max_snapshot_bytes())?;
    Ok(commit)
}

fn resolve_sheet_path(file: &Xlsb, path: &PathArg<'_>) -> Result<usize> {
    let node = crate::path::resolve::get(&Nav::new(file), path)?;
    if node.element() != Kind::Sheet {
        return Err(Error::new(
            Code::UnsupportedType,
            "the XLSB structural verb requires a sheet tab",
        )
        .with_format(FileFormat::Xlsb));
    }
    Ok(node.sheet_position())
}

fn move_destination(
    file: &Xlsb,
    source: usize,
    count: usize,
    placement: &Placement<'_>,
) -> Result<usize> {
    let last = count.checked_sub(1).ok_or_else(|| {
        Error::new(Code::PathNotFound, "the XLSB workbook has no sheet tabs")
            .with_format(FileFormat::Xlsb)
    })?;
    let position = match placement {
        Placement::Start => 0,
        Placement::End => last,
        Placement::Position(position) => position.get(),
        Placement::Before(path) => {
            let anchor = resolve_sheet_path(file, path)?;
            if source == anchor {
                source
            } else if source < anchor {
                anchor.saturating_sub(1)
            } else {
                anchor
            }
        },
        Placement::After(path) => {
            let anchor = resolve_sheet_path(file, path)?;
            if source <= anchor {
                anchor
            } else {
                anchor.saturating_add(1)
            }
        },
        Placement::Into(path, position) => {
            let parent = crate::path::resolve::get(&Nav::new(file), path)?;
            if parent.element() != Kind::Workbook {
                return Err(Error::new(
                    Code::UnsupportedType,
                    "XLSB sheets can only be reordered within the workbook",
                )
                .with_format(FileFormat::Xlsb));
            }
            position.map_or(last, Position::get)
        },
    };
    if position >= count {
        return Err(Error::new(
            Code::OutOfRange,
            "the XLSB sheet move position is outside the workbook",
        )
        .with_format(FileFormat::Xlsb));
    }
    Ok(position)
}

fn move_permutation(count: usize, source: usize, destination: usize) -> Result<Vec<usize>> {
    if source >= count || destination >= count {
        return Err(Error::new(
            Code::OutOfRange,
            "the XLSB sheet move position is outside the workbook",
        )
        .with_format(FileFormat::Xlsb));
    }
    let mut order = identity_permutation(count)?;
    if source != destination {
        let tab = order.remove(source);
        order.insert(destination, tab);
    }
    Ok(order)
}

fn identity_permutation(count: usize) -> Result<Vec<usize>> {
    let mut order = Vec::new();
    order
        .try_reserve_exact(count)
        .map_err(|_| allocation("XLSB sheet-order permutation"))?;
    order.extend(0..count);
    Ok(order)
}

fn admit_order_targets(targets: usize, catalog_tabs: usize, limits: &SemanticLimits) -> Result<()> {
    if targets > limits.max_expanded_targets() {
        return Err(expanded_limit(targets, limits.max_expanded_targets()));
    }
    // The semantic charge is the selected tab count (one for a move and one
    // or two for a swap), while the native permutation itself is catalog-wide.
    // Account for the complete facade permutation and selector vectors before
    // allocating either vector. Native transfer limits independently bound its
    // encoded patch history.
    let permutation_bytes = catalog_tabs
        .checked_mul(size_of::<usize>())
        .ok_or_else(|| staged_limit(usize::MAX, limits.max_staged_bytes()))?;
    let selector_bytes = catalog_tabs
        .checked_mul(size_of::<NativeSelector<'_>>())
        .ok_or_else(|| staged_limit(usize::MAX, limits.max_staged_bytes()))?;
    let staged = permutation_bytes
        .checked_add(selector_bytes)
        .ok_or_else(|| staged_limit(usize::MAX, limits.max_staged_bytes()))?;
    if staged > limits.max_staged_bytes() {
        return Err(staged_limit(staged, limits.max_staged_bytes()));
    }
    Ok(())
}

fn set_sheet_with_limits(
    file: &Xlsb,
    target: TargetArg<'_>,
    props: &Props<'_>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    let positions = resolve_sheet_targets(file, &target, limits)?;
    let assignments = admit_sheet_properties(Operation::Set, props, limits)?;
    check_limits(positions.len(), &assignments, limits)?;

    if assignments.name.is_some() && positions.len() > 1 {
        return Err(Error::new(
            Code::InvalidArgument,
            "renaming a selector with multiple sheet matches is ambiguous",
        )
        .with_format(FileFormat::Xlsb)
        .with_property("name"));
    }

    if let Some(name) = assignments.name.as_deref() {
        validate_sheet_name(file, &[positions[0].tab], name)?;
    }

    if let Some(state) = assignments.state {
        validate_sheet_state(file, &positions, state)?;
    }

    let catalog = file.catalog()?;
    let name_is_changed = assignments.name.as_deref().is_some_and(|name| {
        catalog
            .get(positions[0].tab)
            .is_none_or(|tab| tab.name() != name)
    });
    let state_is_changed = assignments.state.is_some_and(|state| {
        positions.iter().any(|target| {
            catalog
                .get(target.tab)
                .is_none_or(|tab| tab.state() != state)
        })
    });

    // All target and property admission above is intentionally before this
    // editor construction. Signed and retention policy are still checked by
    // `Xlsb::check_source` at the outer fence and by `editor` here.
    let editor = file.editor_for_write()?;
    let package = package_from_editor(editor);
    let native_limits = TransferLimits::new(
        limits.max_snapshot_bytes(),
        limits.max_expanded_targets().saturating_mul(2),
        1,
        limits.max_snapshot_bytes(),
    );
    let mut edit = package
        .edit_with_options(NativeStructureOptions::new().transfer_limits(native_limits))
        .map_err(native_error)?;
    if name_is_changed && let Some(name) = assignments.name {
        // A name assignment is restricted to one target above, so cloning the
        // owned value does not multiply authored payload retention.
        edit.set_sheet_name(positions[0].tab, name)
            .map_err(rename_error)?
            .ok_or_else(|| {
                Error::new(Code::PathNotFound, "the selected XLSB tab is absent")
                    .with_format(FileFormat::Xlsb)
                    .with_property("name")
            })?;
    }
    if state_is_changed && let Some(state) = assignments.state {
        for target in &positions {
            edit.set_sheet_state(target.tab, state)
                .map_err(native_error)?
                .ok_or_else(|| {
                    Error::new(Code::PathNotFound, "the selected XLSB tab is absent")
                        .with_format(FileFormat::Xlsb)
                })?;
        }
    }
    let native = edit.commit().map_err(native_error)?;
    let commit = Commit::from_native(file, native, limits.max_snapshot_bytes())?;
    Ok(commit)
}

#[derive(Debug, Clone)]
struct CellAssignments {
    value: Option<NativeScalar>,
    formula: Option<String>,
    format: NativeFormat,
    staged_bytes: usize,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub(crate) struct CellTarget {
    pub(crate) sheet: usize,
    pub(crate) reference: Reference,
}

fn set_cells_with_limits(
    file: &Xlsb,
    target: &TargetArg<'_>,
    props: &Props<'_>,
    limits: &SemanticLimits,
) -> Result<Option<Commit>> {
    let admitted = target
        .path()
        .and_then(|path| path.path().steps().last())
        .and_then(|step| step.cell_range())
        .map(|_| admit_cell_properties(props, limits))
        .transpose()?;
    if let Some(assignments) = admitted.as_ref() {
        let cells = target
            .path()
            .and_then(|path| path.path().steps().last())
            .and_then(|step| step.cell_range())
            .map(range_area)
            .unwrap_or(0);
        let staged = cells
            .checked_mul(assignments.staged_bytes)
            .ok_or_else(|| staged_limit(usize::MAX, limits.max_staged_bytes()))?;
        if staged > limits.max_staged_bytes() {
            return Err(staged_limit(staged, limits.max_staged_bytes()));
        }
    }
    let Some(targets) = resolve_cell_targets(file, target, limits)? else {
        return Ok(None);
    };
    if props.is_empty() {
        return Ok(Some(identity_commit(file, limits)?));
    }
    let assignments = match admitted {
        Some(assignments) => assignments,
        None => admit_cell_properties(props, limits)?,
    };
    check_cell_limits(targets.len(), &assignments, limits)?;

    // Complete target grouping is still admission work: it must finish before
    // the native package owner is materialized. Every target in one worksheet
    // is staged in one native cell-values edit; a selector spanning worksheets
    // publishes the edits in workbook order through one exact package chain.
    let authored_operations = targets.len();
    let mut grouped = BTreeMap::<usize, Vec<CellTarget>>::new();
    for target in targets {
        grouped
            .entry(target.sheet)
            .or_default()
            .try_reserve(1)
            .map_err(|_| allocation("XLSB cell write worksheet targets"))?;
        grouped
            .get_mut(&target.sheet)
            .expect("entry was just inserted")
            .push(target);
    }
    let editor = file.editor_for_write()?;
    let mut package = package_from_editor(editor);
    let mut patches = Vec::new();
    patches
        .try_reserve(grouped.len())
        .map_err(|_| allocation("XLSB cell write patches"))?;
    let mut canonical = None;
    let mut remaining_snapshot = cell_snapshot_budget(limits.max_snapshot_bytes())?;
    let mut retained_snapshot = cell_patch_container_bytes();
    for (sheet, targets) in grouped {
        let native_limits = cell_transfer_limits(remaining_snapshot, limits)?;
        canonical.get_or_insert((sheet, native_limits));
        let mut edit = package
            .edit_cell_values_with_options(
                NativeSelector::from(sheet),
                NativeCellOptions::new().transfer_limits(native_limits),
            )
            .map_err(native_error)?
            .resolve()
            .map_err(native_error)?
            .ok_or_else(|| {
                Error::new(
                    Code::OutOfRange,
                    "the XLSB worksheet position is out of range",
                )
                .with_format(FileFormat::Xlsb)
            })?;
        for target in targets {
            apply_cell_assignments(&mut edit, target.reference, &assignments)?;
        }
        let native = edit.commit().map_err(native_error)?;
        let (next, patch) = native.into_parts();
        let retained = patch.retained_bytes();
        if retained > remaining_snapshot {
            let consumed = limits
                .max_snapshot_bytes()
                .saturating_sub(remaining_snapshot);
            return Err(snapshot_limit(
                consumed.saturating_add(retained),
                limits.max_snapshot_bytes(),
            ));
        }
        remaining_snapshot -= retained;
        retained_snapshot = retained_snapshot
            .checked_add(retained)
            .ok_or_else(|| snapshot_limit(usize::MAX, limits.max_snapshot_bytes()))?;
        package = next;
        patches.push(patch);
    }
    let canonical = canonical.ok_or_else(|| {
        Error::new(
            Code::InternalError,
            "XLSB cell write lost its worksheet owner",
        )
        .with_format(FileFormat::Xlsb)
    })?;
    let commit = Commit::from_cell_parts(
        file,
        package,
        patches,
        canonical,
        retained_snapshot,
        limits.max_snapshot_bytes(),
        authored_operations,
    )?;
    Ok(Some(commit))
}

/// Publish an exact native identity patch for an admitted empty-property
/// write.  Keep this on the structural native owner so the transaction
/// reports zero leaf operations while still applying the normal source,
/// signature and snapshot-limit checks.
fn identity_commit(file: &Xlsb, limits: &SemanticLimits) -> Result<Commit> {
    let package = package_from_editor(file.editor_for_write()?);
    let native_limits = TransferLimits::new(
        limits.max_snapshot_bytes(),
        limits.max_expanded_targets(),
        1,
        limits.max_snapshot_bytes(),
    );
    let native = package
        .edit_with_options(NativeStructureOptions::new().transfer_limits(native_limits))
        .map_err(native_error)?
        .commit()
        .map_err(native_error)?;
    Commit::from_native(file, native, limits.max_snapshot_bytes())
}

/// Stage one cell's admitted payload and format in a native cell edit: the
/// value or formula first, then the partial format change.
fn apply_cell_assignments(
    edit: &mut NativeCellEdit,
    reference: Reference,
    assignments: &CellAssignments,
) -> Result<()> {
    if let Some(value) = assignments.value.as_ref() {
        edit.set_scalar(reference, value.clone())
            .map_err(native_error)?;
    }
    if let Some(formula) = assignments.formula.as_deref() {
        edit.set_formula_text(reference, formula)
            .map_err(|error| caller_value_error(error, "formula"))?;
    }
    if !assignments.format.is_empty() {
        edit.set_format(reference, &assignments.format)
            .map_err(native_error)?;
    }
    Ok(())
}

/// Classify a leaf refusal of a caller-supplied value: invalid text is
/// `invalid_value` on the property; other failures keep their class.
fn caller_value_error(error: litchi_xlsb::package::PackageError, property: &'static str) -> Error {
    match error {
        litchi_xlsb::package::PackageError::InvalidFormula(_)
        | litchi_xlsb::package::PackageError::InvalidCellReference(_)
        | litchi_xlsb::package::PackageError::InvalidFormat(_) => Error::new(
            Code::InvalidValue,
            format!("the XLSB {property} value is not valid"),
        )
        .with_format(FileFormat::Xlsb)
        .with_property(property),
        error => native_error(error).with_property(property),
    }
}

fn admit_cell_properties(props: &Props<'_>, limits: &SemanticLimits) -> Result<CellAssignments> {
    let mut assignments = CellAssignments {
        value: None,
        formula: None,
        format: NativeFormat::default(),
        staged_bytes: WRITE_ENTRY_OVERHEAD,
    };
    let mut clear = false;
    let mut seen = Vec::<CellSetter>::new();
    // `border.all` fans out to the four sides; a side named in the same set
    // overrides it. Colors need a line style in the same set.
    let mut all_line: Option<Option<litchi_xlsb::styles::BorderStyle>> = None;
    let mut all_color: Option<u32> = None;
    let mut lines: [Option<Option<litchi_xlsb::styles::BorderStyle>>; 4] = [None; 4];
    let mut colors: [Option<u32>; 4] = [None; 4];
    for (key, value) in props.iter() {
        let (_, entry) = CELL_WRITE_TABLE
            .resolve(key.as_str())
            .map_err(|error| error.with_property(key.as_str().to_owned()))?;
        let entry = *entry;
        entry
            .prepare(Operation::Set, value, |setter, value| {
                if seen.contains(&setter) {
                    return Err(duplicate_cell_property(entry.spec.name));
                }
                seen.try_reserve(1)
                    .map_err(|_| allocation("XLSB cell property set"))?;
                seen.push(setter);
                let property = entry.spec.name;
                let format = &mut assignments.format;
                match setter {
                    CellSetter::Value => {
                        let scalar = match value {
                            Some(value) => native_scalar(value)?,
                            None => NativeScalar::Blank,
                        };
                        let bytes = native_scalar_bytes(&scalar);
                        assignments.staged_bytes = assignments
                            .staged_bytes
                            .checked_add(bytes)
                            .ok_or_else(|| staged_limit(usize::MAX, limits.max_staged_bytes()))?;
                        assignments.value = Some(scalar);
                    },
                    CellSetter::Clear => {
                        let Some(Scalar::Bool(true)) = value else {
                            return Err(Error::new(Code::InvalidValue, "clear must be true")
                                .with_format(FileFormat::Xlsb)
                                .with_property("clear"));
                        };
                        clear = true;
                        assignments.staged_bytes = assignments
                            .staged_bytes
                            .checked_add(1)
                            .ok_or_else(|| staged_limit(usize::MAX, limits.max_staged_bytes()))?;
                    },
                    CellSetter::Formula => {
                        let Some(Scalar::Text(text)) = value else {
                            return Err(invalid_cell_value(property, "formula must be text"));
                        };
                        assignments.staged_bytes = assignments
                            .staged_bytes
                            .checked_add(text.len())
                            .ok_or_else(|| staged_limit(usize::MAX, limits.max_staged_bytes()))?;
                        assignments.formula = Some(text.to_owned());
                    },
                    CellSetter::Bold => format.bold = Some(cell_bool(value, property)?),
                    CellSetter::Italic => format.italic = Some(cell_bool(value, property)?),
                    CellSetter::Strike => format.strike = Some(cell_bool(value, property)?),
                    CellSetter::Underline => {
                        format.underline = Some(matches!(value, Some(Scalar::Token("single"))));
                    },
                    CellSetter::FontName => {
                        let Some(Scalar::Text(name)) = value else {
                            return Err(invalid_cell_value(property, "font name must be text"));
                        };
                        format.font_name = Some(name.to_owned());
                    },
                    CellSetter::FontSize => {
                        let Some(Scalar::Length(size)) = value else {
                            return Err(invalid_cell_value(property, "font size must be a length"));
                        };
                        format.font_size = Some(size.points());
                    },
                    CellSetter::FontColor => format.font_color = Some(cell_color(value, property)?),
                    CellSetter::Fill => format.fill = Some(cell_color(value, property)?),
                    CellSetter::BorderAll => all_line = Some(border_line(value)),
                    CellSetter::BorderAllColor => {
                        all_color = cell_color(value, property)?;
                    },
                    CellSetter::BorderLeft
                    | CellSetter::BorderRight
                    | CellSetter::BorderTop
                    | CellSetter::BorderBottom => {
                        lines[border_side(setter)] = Some(border_line(value));
                    },
                    CellSetter::BorderLeftColor
                    | CellSetter::BorderRightColor
                    | CellSetter::BorderTopColor
                    | CellSetter::BorderBottomColor => {
                        colors[border_side(setter)] = cell_color(value, property)?;
                    },
                    CellSetter::NumberFormat => {
                        let Some(Scalar::Text(code)) = value else {
                            return Err(invalid_cell_value(property, "number format must be text"));
                        };
                        format.number_format = Some(code.to_owned());
                    },
                    CellSetter::Horizontal => {
                        format.horizontal = Some(match value {
                            Some(Scalar::Token(token)) => horizontal_token(token),
                            _ => {
                                return Err(invalid_cell_value(
                                    property,
                                    "horizontal alignment must be a token",
                                ));
                            },
                        });
                    },
                    CellSetter::Vertical => {
                        format.vertical = Some(match value {
                            Some(Scalar::Token(token)) => vertical_token(token),
                            _ => {
                                return Err(invalid_cell_value(
                                    property,
                                    "vertical alignment must be a token",
                                ));
                            },
                        });
                    },
                    CellSetter::WrapText => format.wrap_text = Some(cell_bool(value, property)?),
                    CellSetter::ShrinkToFit => {
                        format.shrink_to_fit = Some(cell_bool(value, property)?);
                    },
                    CellSetter::TextRotation => {
                        let rotation = cell_byte(value, property)?;
                        if rotation > 180 && rotation != 255 {
                            return Err(invalid_cell_value(
                                property,
                                "text rotation must be 0..=180 or 255",
                            ));
                        }
                        format.rotation = Some(rotation);
                    },
                    CellSetter::Indent => format.indent = Some(cell_byte(value, property)?),
                    CellSetter::ReadingOrder => {
                        format.reading_order = Some(match value {
                            Some(Scalar::Token("ltr")) => 1,
                            Some(Scalar::Token("rtl")) => 2,
                            _ => 0,
                        });
                    },
                }
                Ok(())
            })
            .map_err(|error| error.with_format(FileFormat::Xlsb))?;
    }
    for (side, line) in lines.iter_mut().enumerate() {
        if line.is_none() && all_line.is_some() {
            *line = all_line;
        }
        if colors[side].is_none() {
            colors[side] = all_color;
        }
        let color = colors[side];
        let resolved = match *line {
            Some(Some(style)) => Some(Some(NativeBorderLine { style, color })),
            Some(None) => {
                if color.is_some() {
                    return Err(invalid_cell_value(
                        "border.color",
                        "a border color cannot be set on a removed line",
                    ));
                }
                Some(None)
            },
            None => {
                if color.is_some() {
                    return Err(Error::new(
                        Code::InvalidArgument,
                        "an XLSB border color needs its line style in the same set",
                    )
                    .with_format(FileFormat::Xlsb)
                    .with_suggestion(
                        "set `border.<side>` (or `border.all`) together with its color",
                    ));
                }
                None
            },
        };
        let target = &mut assignments.format;
        match side {
            0 => target.border_left = resolved,
            1 => target.border_right = resolved,
            2 => target.border_top = resolved,
            _ => target.border_bottom = resolved,
        }
    }
    if clear {
        if assignments.value.is_some() {
            return Err(Error::new(
                Code::InvalidArgument,
                "cell properties value and clear cannot be supplied together",
            )
            .with_format(FileFormat::Xlsb)
            .with_property("clear"));
        }
        assignments.value = Some(NativeScalar::Blank);
    }
    if assignments.formula.is_some() && assignments.value.is_some() {
        return Err(Error::new(
            Code::InvalidArgument,
            "cell properties formula and value (or clear) cannot be supplied together",
        )
        .with_format(FileFormat::Xlsb)
        .with_property("formula"));
    }
    assignments
        .format
        .validate()
        .map_err(|error| caller_value_error(error, "format"))?;
    if assignments.value.is_none()
        && assignments.formula.is_none()
        && assignments.format.is_empty()
        && !props.is_empty()
    {
        return Err(Error::new(
            Code::InvalidArgument,
            "an XLSB cell write requires value, clear, formula or a format property",
        )
        .with_format(FileFormat::Xlsb));
    }
    Ok(assignments)
}

const fn border_side(setter: CellSetter) -> usize {
    match setter {
        CellSetter::BorderLeft | CellSetter::BorderLeftColor => 0,
        CellSetter::BorderRight | CellSetter::BorderRightColor => 1,
        CellSetter::BorderTop | CellSetter::BorderTopColor => 2,
        _ => 3,
    }
}

fn border_line(value: Option<Scalar<'_>>) -> Option<litchi_xlsb::styles::BorderStyle> {
    use litchi_xlsb::styles::BorderStyle;
    let Some(Scalar::Token(token)) = value else {
        return None;
    };
    Some(match token {
        "thin" => BorderStyle::Thin,
        "medium" => BorderStyle::Medium,
        "dashed" => BorderStyle::Dashed,
        "dotted" => BorderStyle::Dotted,
        "thick" => BorderStyle::Thick,
        "double" => BorderStyle::Double,
        "hair" => BorderStyle::Hair,
        "mediumDashed" => BorderStyle::MediumDashed,
        "dashDot" => BorderStyle::DashDot,
        "mediumDashDot" => BorderStyle::MediumDashDot,
        "dashDotDot" => BorderStyle::DashDotDot,
        "mediumDashDotDot" => BorderStyle::MediumDashDotDot,
        "slantDashDot" => BorderStyle::SlantDashDot,
        _ => return None,
    })
}

fn horizontal_token(token: &str) -> litchi_xlsb::styles::HorizontalAlignment {
    use litchi_xlsb::styles::HorizontalAlignment;
    match token {
        "left" => HorizontalAlignment::Left,
        "center" => HorizontalAlignment::Center,
        "right" => HorizontalAlignment::Right,
        "fill" => HorizontalAlignment::Fill,
        "justify" => HorizontalAlignment::Justify,
        "centerContinuous" => HorizontalAlignment::CenterContinuous,
        "distributed" => HorizontalAlignment::Distributed,
        _ => HorizontalAlignment::General,
    }
}

fn vertical_token(token: &str) -> litchi_xlsb::styles::VerticalAlignment {
    use litchi_xlsb::styles::VerticalAlignment;
    match token {
        "top" => VerticalAlignment::Top,
        "center" => VerticalAlignment::Center,
        "justify" => VerticalAlignment::Justify,
        "distributed" => VerticalAlignment::Distributed,
        _ => VerticalAlignment::Bottom,
    }
}

fn cell_bool(value: Option<Scalar<'_>>, property: &'static str) -> Result<bool> {
    match value {
        Some(Scalar::Bool(value)) => Ok(value),
        _ => Err(invalid_cell_value(property, "the property must be boolean")),
    }
}

fn cell_byte(value: Option<Scalar<'_>>, property: &'static str) -> Result<u8> {
    match value {
        Some(Scalar::Int(value)) => u8::try_from(value)
            .map_err(|_| invalid_cell_value(property, "the integer is outside its field")),
        _ => Err(invalid_cell_value(
            property,
            "the property must be an integer",
        )),
    }
}

/// A direct RGB color, or `None` for a null (automatic color, no fill).
fn cell_color(value: Option<Scalar<'_>>, property: &'static str) -> Result<Option<u32>> {
    let Some(value) = value else {
        return Ok(None);
    };
    let Scalar::Color(color) = value else {
        return Err(invalid_cell_value(property, "the property must be a color"));
    };
    if !color.transforms().is_empty() || color.spreadsheet_tint().is_some() {
        return Err(invalid_cell_value(
            property,
            "XLSB colors take no transforms or tint",
        ));
    }
    match color.base() {
        crate::value::Base::Rgb(rgb) => Ok(Some(u32::from_be_bytes([
            rgb.alpha(),
            rgb.red(),
            rgb.green(),
            rgb.blue(),
        ]))),
        _ => Err(Error::new(
            Code::UnsupportedCapability,
            "XLSB formatting writes direct RGB colors only",
        )
        .with_format(FileFormat::Xlsb)
        .with_property(property)
        .with_suggestion("use an RGB color such as `1F4E79`")),
    }
}

fn invalid_cell_value(property: &'static str, message: &'static str) -> Error {
    Error::new(Code::InvalidValue, message)
        .with_format(FileFormat::Xlsb)
        .with_property(property)
}

fn native_scalar(value: Scalar<'_>) -> Result<NativeScalar> {
    match value {
        Scalar::Bool(value) => Ok(NativeScalar::Boolean(value)),
        Scalar::Text(value) => {
            validate_cell_text(value)?;
            Ok(NativeScalar::Text(value.to_owned()))
        },
        Scalar::Int(value) => {
            let number = value as f64;
            if !number.is_finite() || number as i128 != i128::from(value) {
                return Err(Error::new(
                    Code::InvalidValue,
                    "integer cell values must be exactly representable as XLSB numbers",
                )
                .with_format(FileFormat::Xlsb)
                .with_property("value"));
            }
            Ok(NativeScalar::Number(number))
        },
        Scalar::Number(value) if value.is_normal() || value.to_bits() == 0 => {
            Ok(NativeScalar::Number(value))
        },
        Scalar::Number(_) => Err(Error::new(
            Code::InvalidValue,
            "cell numbers must be finite normalized values or positive zero",
        )
        .with_format(FileFormat::Xlsb)
        .with_property("value")),
        _ => Err(Error::new(
            Code::InvalidValue,
            "XLSB cell values support only text, bool, integer and finite number scalars",
        )
        .with_format(FileFormat::Xlsb)
        .with_property("value")),
    }
}

fn validate_cell_text(value: &str) -> Result<()> {
    if value.encode_utf16().count() > 32_767
        || value
            .chars()
            .any(|character| !is_xml_10_character(character))
    {
        return Err(Error::new(
            Code::InvalidValue,
            "cell text is outside the XLSB string domain",
        )
        .with_format(FileFormat::Xlsb)
        .with_property("value"));
    }
    Ok(())
}

fn native_scalar_bytes(value: &NativeScalar) -> usize {
    match value {
        NativeScalar::Text(value) => value.len(),
        _ => 8,
    }
}

fn check_cell_limits(
    targets: usize,
    assignments: &CellAssignments,
    limits: &SemanticLimits,
) -> Result<()> {
    if limits.max_operations() < 1 {
        return Err(operation_limit(1, limits.max_operations()));
    }
    if targets > limits.max_expanded_targets() {
        return Err(expanded_limit(targets, limits.max_expanded_targets()));
    }
    let staged = targets
        .checked_mul(assignments.staged_bytes)
        .ok_or_else(|| staged_limit(usize::MAX, limits.max_staged_bytes()))?;
    if staged > limits.max_staged_bytes() {
        return Err(staged_limit(staged, limits.max_staged_bytes()));
    }
    Ok(())
}

fn check_remove_limits(targets: usize, limits: &SemanticLimits) -> Result<()> {
    if limits.max_operations() < 1 {
        return Err(operation_limit(1, limits.max_operations()));
    }
    if targets > limits.max_expanded_targets() {
        return Err(expanded_limit(targets, limits.max_expanded_targets()));
    }
    let staged = targets
        .checked_mul(WRITE_ENTRY_OVERHEAD)
        .ok_or_else(|| staged_limit(usize::MAX, limits.max_staged_bytes()))?;
    if staged > limits.max_staged_bytes() {
        return Err(staged_limit(staged, limits.max_staged_bytes()));
    }
    Ok(())
}

/// Reserve the exact owner charge added by the facade's cell-patch vector.
/// NativeCellPatch::retained_bytes already includes package images and every
/// owned cell delta; no per-target estimate is permitted here.
pub(crate) fn cell_snapshot_budget(max_snapshot_bytes: usize) -> Result<usize> {
    let container = cell_patch_container_bytes();
    max_snapshot_bytes
        .checked_sub(container)
        .filter(|budget| *budget > 0)
        .ok_or_else(|| snapshot_limit(max_snapshot_bytes.saturating_add(1), max_snapshot_bytes))
}

pub(crate) const fn cell_patch_container_bytes() -> usize {
    size_of::<Vec<NativeCellPatch>>()
}

pub(crate) fn cell_transfer_limits(
    remaining_snapshot: usize,
    limits: &SemanticLimits,
) -> Result<TransferLimits> {
    if remaining_snapshot == 0 {
        return Err(snapshot_limit(
            limits.max_snapshot_bytes().saturating_add(1),
            limits.max_snapshot_bytes(),
        ));
    }
    Ok(TransferLimits::new(
        remaining_snapshot,
        limits.max_expanded_targets(),
        1,
        remaining_snapshot,
    ))
}

fn duplicate_cell_property(property: &'static str) -> Error {
    Error::new(
        Code::InvalidArgument,
        "the same XLSB cell property was supplied more than once",
    )
    .with_format(FileFormat::Xlsb)
    .with_property(property)
}

pub(super) fn package_from_editor(editor: XlsbEditor<'_>) -> litchi_xlsb::Package {
    match editor {
        XlsbEditor::Eager(package) => package.clone(),
        XlsbEditor::Materialized(package) => *package,
    }
}

/// Resolve cells and ranges while retaining the complete XLSB tab position.
/// `None` means that the target selected only sheet-level nodes, so the sheet
/// setter may handle it. A target that mixes cell and non-cell nodes is
/// refused rather than partially applying a mutation.
fn resolve_cell_targets(
    file: &Xlsb,
    target: &TargetArg<'_>,
    limits: &SemanticLimits,
) -> Result<Option<Vec<CellTarget>>> {
    preflight_direct_cell_range(target, limits)?;
    let mut targets = Vec::new();
    let mut seen = HashSet::new();
    let mut matched = false;
    let mut non_cell = false;
    if let Some(path) = target.path() {
        let nav = if direct_cell_path(target) {
            Nav::for_cell_target(file)
        } else {
            Nav::new(file)
        };
        let node = crate::path::resolve::get(&nav, path)?;
        matched = true;
        append_cell_node(&mut targets, &mut seen, &mut non_cell, node, limits)?;
    } else if let Some(query) = target.query() {
        let walk = crate::query::run(Nav::new(file), query)?;
        if !walk.plan().diagnostics().is_empty() {
            return Err(Error::new(
                Code::InvalidSelector,
                "the mutation selector contains an unsupported predicate",
            )
            .with_format(FileFormat::Xlsb));
        }
        for node in walk {
            matched = true;
            append_cell_node(&mut targets, &mut seen, &mut non_cell, node?, limits)?;
        }
    } else {
        return Err(
            Error::new(Code::InvalidArgument, "the mutation target is empty")
                .with_format(FileFormat::Xlsb),
        );
    }
    file.check_source_for_replay()?;
    if !matched {
        return Err(Error::new(
            Code::ZeroMatches,
            "the mutation selector matched no element",
        )
        .with_format(FileFormat::Xlsb));
    }
    if non_cell && !targets.is_empty() {
        return Err(Error::new(
            Code::UnsupportedElement,
            "an XLSB cell mutation selector cannot mix cells with sheet nodes",
        )
        .with_format(FileFormat::Xlsb));
    }
    if targets.is_empty() {
        Ok(None)
    } else {
        Ok(Some(targets))
    }
}

fn direct_cell_path(target: &TargetArg<'_>) -> bool {
    target.path().is_some_and(|path| {
        // A root-level `/A1` can be a valid sheet name.  Only a scoped,
        // bare-coordinate path is a direct cell mutation target; all other
        // paths retain the normal read-capable navigation mode.
        path.path().steps().len() > 1
            && path
                .path()
                .steps()
                .last()
                .is_some_and(|step| step.cell_ref().is_some())
    })
}

fn preflight_direct_cell_range(target: &TargetArg<'_>, limits: &SemanticLimits) -> Result<()> {
    let Some(path) = target.path() else {
        return Ok(());
    };
    let Some(step) = path.path().steps().last() else {
        return Ok(());
    };
    let Some(range) = step.cell_range() else {
        return Ok(());
    };
    let cells = range_area(range);
    if cells > limits.max_expanded_targets() {
        return Err(expanded_limit(cells, limits.max_expanded_targets()));
    }
    Ok(())
}

fn append_cell_node(
    targets: &mut Vec<CellTarget>,
    seen: &mut HashSet<CellTarget>,
    non_cell: &mut bool,
    node: Node<'_>,
    limits: &SemanticLimits,
) -> Result<()> {
    if node.element() == Kind::Cell {
        let Some(cell) = node.cell_ref() else {
            return Err(
                Error::new(Code::InvalidPath, "the XLSB cell reference is invalid")
                    .with_format(FileFormat::Xlsb),
            );
        };
        let reference = Reference::new(cell.row(), cell.column()).map_err(native_error)?;
        append_cell_target(
            targets,
            seen,
            CellTarget {
                sheet: node.sheet_position(),
                reference,
            },
            limits,
        )?;
        return Ok(());
    }
    if node.element() == Kind::Range {
        let Item::Range(range) = node.item() else {
            return Err(
                Error::new(Code::InvalidPath, "the XLSB range reference is invalid")
                    .with_format(FileFormat::Xlsb),
            );
        };
        append_range_targets(targets, seen, node.sheet_position(), range.rect, limits)?;
        return Ok(());
    }
    *non_cell = true;
    Ok(())
}

fn append_cell_target(
    targets: &mut Vec<CellTarget>,
    seen: &mut HashSet<CellTarget>,
    target: CellTarget,
    limits: &SemanticLimits,
) -> Result<()> {
    if !seen.insert(target) {
        return Ok(());
    }
    let total = targets
        .len()
        .checked_add(1)
        .ok_or_else(|| expanded_limit(usize::MAX, limits.max_expanded_targets()))?;
    if total > limits.max_expanded_targets() {
        return Err(expanded_limit(total, limits.max_expanded_targets()));
    }
    targets
        .try_reserve(1)
        .map_err(|_| allocation("XLSB cell write targets"))?;
    targets.push(target);
    Ok(())
}

fn append_range_targets(
    targets: &mut Vec<CellTarget>,
    seen: &mut HashSet<CellTarget>,
    sheet: usize,
    range: Rect,
    limits: &SemanticLimits,
) -> Result<()> {
    let start = range.start();
    let (end_row, end_column) = range.end();
    let first_row = start.row().get();
    let last_row = end_row.saturating_sub(1);
    let first_column = start.column().get();
    let last_column = end_column.saturating_sub(1);
    let rows = usize::try_from(last_row - first_row + 1).unwrap_or(usize::MAX);
    let columns = usize::try_from(last_column - first_column + 1).unwrap_or(usize::MAX);
    let cells = rows
        .checked_mul(columns)
        .ok_or_else(|| expanded_limit(usize::MAX, limits.max_expanded_targets()))?;
    if cells > limits.max_expanded_targets()
        || targets.len().saturating_add(cells) > limits.max_expanded_targets()
    {
        let observed = targets.len().saturating_add(cells);
        return Err(expanded_limit(observed, limits.max_expanded_targets()));
    }
    targets
        .try_reserve(cells)
        .map_err(|_| allocation("XLSB cell write range targets"))?;
    seen.try_reserve(cells)
        .map_err(|_| allocation("XLSB cell write range identities"))?;
    for row in first_row..=last_row {
        for column in first_column..=last_column {
            let reference = Reference::new(row, column).map_err(native_error)?;
            append_cell_target(targets, seen, CellTarget { sheet, reference }, limits)?;
        }
    }
    Ok(())
}

fn range_area(range: CellRange) -> usize {
    let start = range.start();
    let end = range.end();
    let rows = usize::try_from(end.row().abs_diff(start.row()))
        .ok()
        .and_then(|value| value.checked_add(1))
        .unwrap_or(usize::MAX);
    let columns = usize::try_from(end.column().abs_diff(start.column()))
        .ok()
        .and_then(|value| value.checked_add(1))
        .unwrap_or(usize::MAX);
    rows.saturating_mul(columns)
}

/// Classify a removal target before the cell resolver expands it.  A direct
/// sheet path and a homogeneous sheet query use the complete catalog removal
/// route.  Queries that mix a sheet with any other node are rejected so a
/// caller cannot accidentally delete only one branch of a heterogeneous
/// selection; all-cell and all-non-cell queries remain with the existing cell
/// resolver, which supplies their established precise refusal.
fn resolve_sheet_removal_targets(
    file: &Xlsb,
    target: &TargetArg<'_>,
    limits: &SemanticLimits,
) -> Result<Option<Vec<SheetTarget>>> {
    if let Some(path) = target.path() {
        // Preserve the cell remover's deferred-grid route.  In particular, a
        // cached-only source can still remove a checked cell coordinate even
        // though its full worksheet metadata is intentionally unavailable.
        if path.path().steps().len() > 1
            && path
                .path()
                .steps()
                .last()
                .is_some_and(|step| step.cell_ref().is_some() || step.cell_range().is_some())
        {
            return Ok(None);
        }
        let node = crate::path::resolve::get(&Nav::new(file), path)?;
        if node.element() != Kind::Sheet {
            return Ok(None);
        }
        let mut positions = Vec::new();
        let mut seen = HashSet::new();
        append_sheet(&mut positions, &mut seen, node, limits)?;
        file.check_source_for_replay()?;
        return Ok(Some(positions));
    }

    let Some(query) = target.query() else {
        return Ok(None);
    };
    let walk = crate::query::run(Nav::new(file), query)?;
    if !walk.plan().diagnostics().is_empty() {
        return Err(Error::new(
            Code::InvalidSelector,
            "the mutation selector contains an unsupported predicate",
        )
        .with_format(FileFormat::Xlsb));
    }
    let mut positions = Vec::new();
    let mut seen = HashSet::new();
    let mut matched = false;
    let mut other = false;
    for node in walk {
        matched = true;
        let node = node?;
        if node.element() == Kind::Sheet {
            append_sheet(&mut positions, &mut seen, node, limits)?;
        } else {
            other = true;
        }
    }
    file.check_source_for_replay()?;
    if !matched {
        return Err(Error::new(
            Code::ZeroMatches,
            "the mutation selector matched no element",
        )
        .with_format(FileFormat::Xlsb));
    }
    if other && !positions.is_empty() {
        return Err(Error::new(
            Code::UnsupportedElement,
            "an XLSB sheet removal selector cannot mix sheets with other elements",
        )
        .with_format(FileFormat::Xlsb));
    }
    if positions.is_empty() {
        Ok(None)
    } else {
        Ok(Some(positions))
    }
}

/// Resolve only workbook sheet nodes. A cell, range, table, or any other
/// element is rejected before an XLSB native editor is built.
fn resolve_sheet_targets(
    file: &Xlsb,
    target: &TargetArg<'_>,
    limits: &SemanticLimits,
) -> Result<Vec<SheetTarget>> {
    let mut positions = Vec::new();
    let mut seen = HashSet::new();

    if let Some(path) = target.path() {
        let node = crate::path::resolve::get(&Nav::new(file), path)?;
        append_sheet(&mut positions, &mut seen, node, limits)?;
    } else if let Some(query) = target.query() {
        let walk = crate::query::run(Nav::new(file), query)?;
        if !walk.plan().diagnostics().is_empty() {
            return Err(Error::new(
                Code::InvalidSelector,
                "the mutation selector contains an unsupported predicate",
            )
            .with_format(FileFormat::Xlsb));
        }
        let mut matched = false;
        for node in walk {
            matched = true;
            append_sheet(&mut positions, &mut seen, node?, limits)?;
        }
        file.check_source_for_replay()?;
        if !matched {
            return Err(Error::new(
                Code::ZeroMatches,
                "the mutation selector matched no element",
            )
            .with_format(FileFormat::Xlsb));
        }
    } else {
        return Err(
            Error::new(Code::InvalidArgument, "the mutation target is empty")
                .with_format(FileFormat::Xlsb),
        );
    }

    file.check_source_for_replay()?;
    Ok(positions)
}

fn append_sheet(
    positions: &mut Vec<SheetTarget>,
    seen: &mut HashSet<usize>,
    node: Node<'_>,
    limits: &SemanticLimits,
) -> Result<()> {
    if node.element() != Kind::Sheet {
        return Err(Error::new(
            Code::UnsupportedElement,
            "the XLSB mutation target must be a sheet",
        )
        .with_format(FileFormat::Xlsb)
        .with_suggestion(format!("resolved target kind: {}", node.kind())));
    }
    let position = node.sheet_position();
    if !seen.insert(position) {
        return Ok(());
    }
    let observed = positions
        .len()
        .checked_add(1)
        .ok_or_else(|| expanded_limit(usize::MAX, limits.max_expanded_targets()))?;
    if observed > limits.max_expanded_targets() {
        return Err(expanded_limit(observed, limits.max_expanded_targets()));
    }
    positions
        .try_reserve(1)
        .map_err(|_| allocation("XLSB sheet write targets"))?;
    positions.push(SheetTarget { tab: position });
    Ok(())
}

fn admit_sheet_properties(
    operation: Operation,
    props: &Props<'_>,
    limits: &SemanticLimits,
) -> Result<SheetAssignments> {
    let mut assignments = SheetAssignments::default();
    for (key, value) in props.iter() {
        let (_, entry) = SHEET_WRITE_TABLE
            .resolve(key.as_str())
            .map_err(|error| error.with_property(key.as_str().to_owned()))?;
        let entry = *entry;
        entry
            .prepare(operation, value, |setter, value| match setter {
                SheetSetter::Name => {
                    if assignments.name.is_some() {
                        return Err(duplicate_property());
                    }
                    let Some(Scalar::Text(value)) = value else {
                        return Err(Error::new(Code::InvalidValue, "sheet name must be text")
                            .with_format(FileFormat::Xlsb)
                            .with_property("name"));
                    };
                    let observed = value.len().saturating_add(WRITE_ENTRY_OVERHEAD);
                    if observed > limits.max_staged_bytes() {
                        return Err(staged_limit(observed, limits.max_staged_bytes()));
                    }
                    assignments.name = Some(value.to_owned());
                    Ok(())
                },
                SheetSetter::Hidden | SheetSetter::Visibility => {
                    if assignments.state.is_some() {
                        return Err(Error::new(
                            Code::InvalidArgument,
                            "hidden and visibility assign the same sheet state",
                        ));
                    }
                    assignments.state = Some(match (setter, value) {
                        (SheetSetter::Hidden, Some(Scalar::Bool(true))) => SheetState::Hidden,
                        (SheetSetter::Hidden, Some(Scalar::Bool(false))) => SheetState::Visible,
                        (SheetSetter::Visibility, Some(Scalar::Token("visible"))) => {
                            SheetState::Visible
                        },
                        (SheetSetter::Visibility, Some(Scalar::Token("hidden"))) => {
                            SheetState::Hidden
                        },
                        (SheetSetter::Visibility, Some(Scalar::Token("veryHidden"))) => {
                            SheetState::VeryHidden
                        },
                        _ => return Err(Error::new(Code::InvalidValue, "invalid sheet state")),
                    });
                    Ok(())
                },
            })
            .map_err(|error| error.with_format(FileFormat::Xlsb))?;
    }
    Ok(assignments)
}

fn validate_sheet_state(file: &Xlsb, targets: &[SheetTarget], state: SheetState) -> Result<()> {
    let catalog = file.catalog()?;
    let mut selected = HashSet::new();
    selected
        .try_reserve(targets.len())
        .map_err(|_| allocation("XLSB tab state targets"))?;
    selected.extend(targets.iter().map(|target| target.tab));
    let mut visible = false;
    let mut changed = false;
    for position in 0..catalog.len() {
        let Some(tab) = catalog.get(position) else {
            continue;
        };
        let next = if selected.contains(&position) {
            state
        } else {
            tab.state()
        };
        if selected.contains(&position)
            && tab.kind() == litchi_xlsb::sheet::Kind::Unbound
            && next != SheetState::VeryHidden
        {
            return Err(Error::new(
                Code::UnsupportedCapability,
                "unbound XLSB tabs must remain very hidden",
            )
            .with_format(FileFormat::Xlsb));
        }
        changed |= next != tab.state();
        visible |= next.is_visible();
    }
    if changed && !visible {
        return Err(Error::new(
            Code::InvalidArgument,
            "a workbook must retain a visible sheet",
        )
        .with_format(FileFormat::Xlsb));
    }
    Ok(())
}

fn validate_sheet_name(file: &Xlsb, targets: &[usize], name: &str) -> Result<()> {
    let units = name.encode_utf16().count();
    if units == 0
        || units > 31
        || name.contains([':', '\\', '*', '?', '/', '[', ']'])
        || name
            .chars()
            .any(|character| !is_xml_10_character(character))
        || name.starts_with('\'')
        || name.ends_with('\'')
    {
        return Err(Error::new(Code::InvalidValue, "the sheet name is invalid")
            .with_format(FileFormat::Xlsb)
            .with_property("name"));
    }
    let catalog = file.catalog()?;
    let key = caseless_key(name);
    if (0..catalog.len())
        .filter_map(|index| catalog.get(index))
        .any(|tab| {
            let selected = targets.contains(&tab.position().get());
            !selected && caseless_key(tab.name()) == key
        })
    {
        return Err(
            Error::new(Code::DuplicateName, "a sheet name is already in use")
                .with_format(FileFormat::Xlsb)
                .with_property("name"),
        );
    }
    Ok(())
}

const fn is_xml_10_character(character: char) -> bool {
    matches!(character, '\u{9}' | '\u{a}' | '\u{d}')
        || (character >= '\u{20}' && character <= '\u{d7ff}')
        || (character >= '\u{e000}' && character <= '\u{fffd}')
        || (character >= '\u{10000}' && character <= '\u{10ffff}')
}

fn caseless_key(value: &str) -> String {
    value.nfd().default_case_fold().nfd().collect()
}

fn check_limits(
    targets: usize,
    assignments: &SheetAssignments,
    limits: &SemanticLimits,
) -> Result<()> {
    if limits.max_operations() < 1 {
        return Err(operation_limit(1, limits.max_operations()));
    }
    if targets > limits.max_expanded_targets() {
        return Err(expanded_limit(targets, limits.max_expanded_targets()));
    }
    let per_target = assignments
        .name
        .as_ref()
        .map_or(0, String::len)
        .saturating_add(WRITE_ENTRY_OVERHEAD)
        .saturating_add(usize::from(assignments.state.is_some()) * WRITE_ENTRY_OVERHEAD);
    let staged = targets.saturating_mul(per_target);
    if staged > limits.max_staged_bytes() {
        return Err(staged_limit(staged, limits.max_staged_bytes()));
    }
    Ok(())
}

fn operation_limit(observed: usize, maximum: usize) -> Error {
    transaction_limit(LimitKind::Operations, observed, maximum)
}

fn expanded_limit(observed: usize, maximum: usize) -> Error {
    transaction_limit(LimitKind::ExpandedTargets, observed, maximum)
}

fn staged_limit(observed: usize, maximum: usize) -> Error {
    transaction_limit(LimitKind::StagedBytes, observed, maximum)
}

pub(super) fn snapshot_limit(observed: usize, maximum: usize) -> Error {
    transaction_limit(LimitKind::SnapshotBytes, observed, maximum)
}

fn transaction_limit(kind: LimitKind, observed: usize, maximum: usize) -> Error {
    Error::from(TransactionError::Limit {
        kind,
        observed,
        maximum,
    })
    .with_format(FileFormat::Xlsb)
}

fn duplicate_property() -> Error {
    Error::new(
        Code::InvalidArgument,
        "the same XLSB sheet property was supplied more than once",
    )
    .with_format(FileFormat::Xlsb)
    .with_property("name")
}

fn allocation(resource: &'static str) -> Error {
    Error::new(Code::InternalError, "could not allocate mutation targets")
        .with_format(FileFormat::Xlsb)
        .with_suggestion(resource)
}

pub(super) fn native_error(error: litchi_xlsb::package::PackageError) -> Error {
    Error::from(error).with_format(FileFormat::Xlsb)
}

fn rename_error(error: litchi_xlsb::package::PackageError) -> Error {
    let duplicate = matches!(
        &error,
        litchi_xlsb::package::PackageError::InvalidFormat(message)
            if message.starts_with("duplicate sheet name")
    );
    if duplicate {
        return Error::new(Code::DuplicateName, "a sheet name is already in use")
            .with_format(FileFormat::Xlsb)
            .with_property("name");
    }
    native_error(error).with_property("name")
}

/// Stream a finalized XLSB snapshot. Source-backed owners stream their
/// retained artifact directly; eager owners serialize their validated package.
pub(crate) fn write_to<W: Write>(file: &Xlsb, mut sink: W) -> Result<()> {
    file.0.owner.finish(|| {
        if file.signature() == super::Signature::RemovalAuthorized {
            return file
                .unsigned_package()?
                .write_to(&mut sink)
                .map_err(native_error);
        }
        match &file.0.owner {
            XlsbOwner::Source(workbook) => workbook
                .workbook()
                .write_to(&mut sink)
                .map_err(native_error),
            XlsbOwner::Eager(package) => package.write_to(&mut sink).map_err(native_error),
        }
    })
}

/// Publish an XLSB snapshot using the shared guarded atomic writer. The
/// source freshness fence runs before replacement; no post-replacement fence
/// is run against a path-backed source whose file has just been replaced.
pub(crate) fn save_to(file: &Xlsb, path: &FilePath, options: &SaveOptions) -> Result<u64> {
    #[cfg(feature = "encryption")]
    if options.reencrypt.is_some() {
        return Err(Error::new(
            Code::UnsupportedCapability,
            "XLSB re-encryption is not supported",
        )
        .with_format(FileFormat::Xlsb));
    }
    file.0.owner.ensure_current()?;
    let context = match &file.0.owner {
        XlsbOwner::Source(workbook) => workbook
            .workbook()
            .source_backed_package()
            .execution_context(),
        XlsbOwner::Eager(_) => None,
    };
    let mut bytes = 0_u64;
    let write = |sink: &mut litchi_opc::atomic::AtomicSink<'_>| -> Result<()> {
        let mut counted = Counted { sink, bytes: 0 };
        write_to(file, &mut counted)?;
        bytes = counted.bytes;
        Ok(())
    };
    let fence = || file.0.owner.ensure_current();
    match options.overwrite {
        Overwrite::Replace => litchi_opc::atomic::replace_guarded_with_durability(
            path,
            options.durability,
            context.as_ref(),
            write,
            fence,
        )?,
        Overwrite::Refuse => litchi_opc::atomic::replace_guarded_with_durability_noclobber(
            path,
            options.durability,
            context.as_ref(),
            write,
            fence,
        )?,
    }
    Ok(bytes)
}

/// The narrow direct-sheet visibility run used by explicit transactions.
///
/// A run fixes one exact `SheetState` and admits distinct direct sheet paths.
/// Names, combined assignments, selectors, repeated targets, state changes,
/// reads, and quota boundaries are barriers owned by the transaction
/// coordinator. One native structural edit therefore preserves the order of
/// active-tab relocation and can still replay a deferred native failure to
/// its originating command index.
pub(crate) mod sheet_coalesce {
    use super::{
        Commit, FileFormat, NativeStructureEdit, Result, SheetState, SheetTarget, TargetArg,
        WRITE_ENTRY_OVERHEAD, Xlsb, admit_sheet_properties, check_limits, native_error,
        package_from_editor, resolve_sheet_targets, sheet_state_shape, validate_sheet_state,
    };
    use crate::error::{Code, Error};
    use crate::props::Props;
    use crate::registry::Operation;
    use crate::transaction::Error as TransactionError;
    use litchi_xlsb::cell_values::TransferLimits as NativeTransferLimits;
    use litchi_xlsb::cell_values::structure::EditOptions as NativeStructureOptions;
    use std::collections::HashSet;

    #[derive(Debug, Clone, Copy)]
    pub(crate) struct Plan {
        pub(crate) tab: usize,
        pub(crate) state: SheetState,
        staged_bytes: usize,
    }

    struct Item {
        plan: Plan,
        index: usize,
    }

    pub(crate) struct Run {
        base: Xlsb,
        catalog: std::sync::Arc<crate::format::xlsb::owner::Catalog>,
        source: litchi_xlsb::Package,
        edit: NativeStructureEdit,
        state: SheetState,
        items: Vec<Item>,
        seen: HashSet<usize>,
        visible_count: usize,
        active_catalog: Option<usize>,
        active_hide: bool,
        staged_bytes: usize,
        native_limits: NativeTransferLimits,
        max_staged_bytes: usize,
    }

    impl Xlsb {
        /// Admit one direct sheet visibility assignment for the structural
        /// run. All other property shapes return `None` so the ordinary
        /// one-shot handler retains their existing semantics.
        pub(crate) fn sheet_plan(
            &self,
            target: &TargetArg<'_>,
            props: &Props<'_>,
            limits: &crate::transaction::SemanticLimits,
        ) -> Result<Option<Plan>> {
            self.check_source_for_replay()?;
            let Some((plan, target)) = admit_sheet_state_plan(self, target, props, limits)? else {
                return Ok(None);
            };
            validate_sheet_state(self, std::slice::from_ref(&target), plan.state)?;
            self.check_source_for_replay().and(Ok(Some(plan)))
        }
    }

    /// Admit the direct positional visibility shape without deciding whether
    /// the complete workbook would become invalid.  The latter check is
    /// intentionally supplied by the caller: the first command and every
    /// post-barrier replan use the ordinary full-catalog validator, while a
    /// pending homogeneous run validates only its selected tab here and uses
    /// its maintained visible count for the last-visible policy.
    fn admit_sheet_state_plan(
        file: &Xlsb,
        target: &TargetArg<'_>,
        props: &Props<'_>,
        limits: &crate::transaction::SemanticLimits,
    ) -> Result<Option<(Plan, SheetTarget)>> {
        if !sheet_state_shape(target, props) {
            return Ok(None);
        }
        let assignments = admit_sheet_properties(Operation::Set, props, limits)?;
        let Some(state) = assignments.state else {
            return Ok(None);
        };
        let positions = resolve_sheet_targets(file, target, limits)?;
        if positions.len() != 1 {
            return Ok(None);
        }
        check_limits(positions.len(), &assignments, limits)?;
        let target = positions[0];
        Ok(Some((
            Plan {
                tab: target.tab,
                state,
                staged_bytes: WRITE_ENTRY_OVERHEAD.saturating_mul(2),
            },
            target,
        )))
    }

    impl Run {
        pub(crate) fn new(
            base: &Xlsb,
            state: SheetState,
            limits: &crate::transaction::SemanticLimits,
            remaining_snapshot: usize,
        ) -> Result<Self> {
            if remaining_snapshot == 0 {
                return Err(super::snapshot_limit(
                    limits.max_snapshot_bytes().saturating_add(1),
                    limits.max_snapshot_bytes(),
                ));
            }
            base.check_source_for_replay()?;
            let source = package_from_editor(
                if base.signature() == crate::format::signature::Signature::Signed {
                    base.editor_for_write()?
                } else {
                    base.editor()?
                },
            );
            let native_limits = NativeTransferLimits::new(
                remaining_snapshot,
                limits.max_expanded_targets(),
                1,
                remaining_snapshot,
            );
            let edit = source
                .edit_with_options(NativeStructureOptions::new().transfer_limits(native_limits))
                .map_err(native_error)?;
            let catalog = base.catalog()?;
            let mut visible_count = 0usize;
            for position in 0..catalog.len() {
                if catalog
                    .get(position)
                    .is_some_and(|tab| tab.state().is_visible())
                {
                    visible_count = visible_count.saturating_add(1);
                }
            }
            base.check_source_for_replay()?;
            Ok(Self {
                base: base.clone(),
                catalog,
                source,
                edit,
                state,
                items: Vec::new(),
                seen: HashSet::new(),
                visible_count,
                active_catalog: base
                    .active_catalog_position()?
                    .map(|position| position.get()),
                active_hide: false,
                staged_bytes: 0,
                native_limits,
                max_staged_bytes: limits.max_staged_bytes(),
            })
        }

        pub(crate) fn first_index(&self) -> Option<usize> {
            self.items.first().map(|item| item.index)
        }

        /// Admit a direct visibility command against the still-pending run.
        /// The target and property shape, quotas, and selected-tab kind retain
        /// their ordinary admission rules, but the workbook-wide last-visible
        /// scan is deferred to [`Self::would_hide_last_visible`], which uses
        /// the run's maintained count.  A caller must flush and replan through
        /// [`Xlsb::sheet_plan`] whenever [`Self::requires_barrier`] is true.
        pub(crate) fn sheet_plan(
            &self,
            target: &TargetArg<'_>,
            props: &Props<'_>,
            limits: &crate::transaction::SemanticLimits,
        ) -> Result<Option<Plan>> {
            self.base.check_source_for_replay()?;
            let Some((plan, target)) = admit_sheet_state_plan(&self.base, target, props, limits)?
            else {
                return Ok(None);
            };
            let Some(tab) = self.catalog.get(target.tab) else {
                return Err(
                    Error::new(Code::PathNotFound, "the selected XLSB tab is absent")
                        .with_format(FileFormat::Xlsb),
                );
            };
            if tab.kind() == litchi_xlsb::sheet::Kind::Unbound
                && plan.state != SheetState::VeryHidden
            {
                return Err(Error::new(
                    Code::UnsupportedCapability,
                    "unbound XLSB tabs must remain very hidden",
                )
                .with_format(FileFormat::Xlsb));
            }
            self.base.check_source_for_replay().and(Ok(Some(plan)))
        }

        pub(crate) fn would_hide_last_visible(&self, plan: &Plan) -> Result<bool> {
            if plan.state.is_visible() {
                return Ok(false);
            }
            let Some(tab) = self.catalog.get(plan.tab) else {
                return Err(
                    Error::new(Code::PathNotFound, "the selected XLSB tab is absent")
                        .with_format(FileFormat::Xlsb),
                );
            };
            Ok(tab.state().is_visible() && self.visible_count <= 1)
        }

        pub(crate) fn requires_barrier(
            &self,
            plan: &Plan,
            limits: &crate::transaction::SemanticLimits,
        ) -> bool {
            plan.state != self.state
                || self.seen.contains(&plan.tab)
                || self.active_hide
                || self.items.len() >= limits.max_expanded_targets()
                || self
                    .staged_bytes
                    .checked_add(plan.staged_bytes)
                    .is_none_or(|bytes| bytes > self.max_staged_bytes)
        }

        pub(crate) fn stage(&mut self, plan: Plan, index: usize) -> Result<bool> {
            if plan.state != self.state
                || self.seen.contains(&plan.tab)
                || self.items.len() >= self.native_limits.changes()
                || self
                    .staged_bytes
                    .checked_add(plan.staged_bytes)
                    .is_none_or(|bytes| bytes > self.max_staged_bytes)
            {
                return Err(Error::new(
                    Code::InternalError,
                    "XLSB sheet coalescing barrier was not flushed",
                )
                .with_format(FileFormat::Xlsb));
            }
            if self.would_hide_last_visible(&plan)? {
                return Err(last_visible_error());
            }
            self.seen
                .try_reserve(1)
                .map_err(|source| TransactionError::Allocation { source })?;
            self.items
                .try_reserve(1)
                .map_err(|source| TransactionError::Allocation { source })?;
            self.base.check_source_for_replay()?;
            let changed = self
                .edit
                .set_sheet_state(plan.tab, plan.state)
                .map_err(native_error)?
                .ok_or_else(|| {
                    Error::new(Code::PathNotFound, "the selected XLSB tab is absent")
                        .with_format(FileFormat::Xlsb)
                })?;
            if let Some(tab) = self.catalog.get(plan.tab) {
                if tab.state().is_visible() && !plan.state.is_visible() {
                    self.visible_count = self.visible_count.saturating_sub(1);
                } else if !tab.state().is_visible() && plan.state.is_visible() {
                    self.visible_count = self.visible_count.saturating_add(1);
                }
            }
            if changed && self.active_catalog == Some(plan.tab) && !plan.state.is_visible() {
                self.active_hide = true;
            }
            self.seen.insert(plan.tab);
            self.staged_bytes = self.staged_bytes.saturating_add(plan.staged_bytes);
            self.items.push(Item { plan, index });
            self.base.check_source_for_replay()?;
            Ok(changed)
        }

        pub(crate) fn commit(self) -> Result<(Xlsb, Commit)> {
            let Self {
                base,
                source,
                edit,
                items,
                native_limits,
                ..
            } = self;
            let authored_operations = items.len();
            base.check_source_for_replay()?;
            let result = edit
                .commit()
                .map_err(native_error)
                .and_then(|native| Commit::from_native(&base, native, native_limits.bytes()));
            let freshness = base.check_source_for_replay();
            match freshness.and(result) {
                Ok(mut commit) => {
                    commit.operations = authored_operations;
                    Ok((base, commit))
                },
                Err(error) => Err(locate_failure(
                    &base,
                    &source,
                    items.as_slice(),
                    native_limits,
                    error,
                )),
            }
        }
    }

    fn replay(
        base: &Xlsb,
        source: &litchi_xlsb::Package,
        items: &[Item],
        limits: NativeTransferLimits,
    ) -> Result<litchi_xlsb::Package> {
        base.check_source_for_replay()?;
        if items.is_empty() {
            return Ok(source.clone());
        }
        let mut edit = source
            .edit_with_options(NativeStructureOptions::new().transfer_limits(limits))
            .map_err(native_error)?;
        for item in items {
            edit.set_sheet_state(item.plan.tab, item.plan.state)
                .map_err(|error| native_error(error).with_item(item.index))?;
        }
        let package = edit.commit().map_err(native_error)?.into_package();
        base.check_source_for_replay()?;
        Ok(package)
    }

    fn locate_failure(
        base: &Xlsb,
        source: &litchi_xlsb::Package,
        items: &[Item],
        limits: NativeTransferLimits,
        error: Error,
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
        let prefix = match replay(base, source, &items[..middle], limits) {
            Ok(package) => package,
            Err(error) => return locate_failure(base, source, &items[..middle], limits, error),
        };
        match replay(base, &prefix, &items[middle..], limits) {
            Err(error) => locate_failure(base, &prefix, &items[middle..], limits, error),
            Ok(_) => error.with_item(first.index),
        }
    }

    fn last_visible_error() -> Error {
        Error::new(
            Code::InvalidArgument,
            "a workbook must retain a visible sheet",
        )
        .with_format(FileFormat::Xlsb)
    }
}

/// The narrow direct-cell payload run used by explicit transactions.
///
/// A run fixes one worksheet and admits only distinct direct cell paths. A
/// range, query, repeated cell, sheet property, read, or staged-byte boundary
/// is a barrier owned by the transaction coordinator. This preserves the
/// original mutation index for deferred native failures while letting a
/// sequence of ordinary scalar writes or removals share one package-owned
/// edit.
/// CSV/TSV import through one native edit (record 0266).
pub(crate) mod find;
pub(crate) mod import;
#[cfg(feature = "eval")]
pub(crate) mod refresh;

pub(crate) mod coalesce {
    use super::{
        CellTarget, Commit, FileFormat, NativeCellEdit, NativeScalar, NativeSelector, Result,
        SemanticLimits, TargetArg, WRITE_ENTRY_OVERHEAD, Xlsb, admit_cell_properties,
        check_cell_limits, check_remove_limits, native_error, package_from_editor,
        resolve_cell_targets,
    };
    use crate::error::{Code, Error};
    use crate::props::Props;
    use crate::transaction::Error as TransactionError;
    use litchi_core::sheet::Cell as _;
    use litchi_xlsb::cell_values::{EditOptions as NativeCellOptions, TransferLimits};
    use std::collections::HashSet;

    #[derive(Debug, Clone)]
    pub(crate) enum Action {
        Set(NativeScalar),
        Remove,
    }

    #[derive(Debug, Clone)]
    pub(crate) struct Plan {
        pub(crate) target: CellTarget,
        action: Action,
        staged_bytes: usize,
        original: Option<NativeScalar>,
        original_known: bool,
        present: bool,
    }

    impl Plan {
        fn key(&self) -> CellTarget {
            self.target
        }

        fn charge(&self) -> usize {
            self.staged_bytes
        }
    }

    struct Item {
        plan: Plan,
        index: usize,
    }

    pub(crate) struct Run {
        base: Xlsb,
        source: litchi_xlsb::Package,
        edit: NativeCellEdit,
        sheet: usize,
        items: Vec<Item>,
        seen: HashSet<CellTarget>,
        staged_bytes: usize,
        native_limits: TransferLimits,
        snapshot_budget: usize,
        max_staged_bytes: usize,
        allow_repeated: bool,
    }

    impl Xlsb {
        pub(crate) fn cell_plan(
            &self,
            target: &TargetArg<'_>,
            props: &Props<'_>,
            limits: &SemanticLimits,
        ) -> Result<Option<Plan>> {
            self.check_source_for_replay()?;
            let result = (|| {
                let Some(path) = target.path() else {
                    return Ok(None);
                };
                let Some(step) = path.path().steps().last() else {
                    return Ok(None);
                };
                if step.cell_ref().is_none() {
                    return Ok(None);
                }
                let assignments = match admit_cell_properties(props, limits) {
                    Ok(assignments) => assignments,
                    Err(error) if error.code() == Code::UnsupportedProperty => {
                        return Err(error);
                    },
                    Err(error) => return Err(error),
                };
                if assignments.formula.is_some() || !assignments.format.is_empty() {
                    // Formulas and formats are staged by the ordinary cell
                    // route; the payload run carries scalar values only.
                    return Ok(None);
                }
                let Some(value) = assignments.value.as_ref().cloned() else {
                    // Empty property sets retain XLSX's admitted no-op
                    // behavior, but do not enter a payload run.
                    return Ok(None);
                };
                check_cell_limits(1, &assignments, limits)?;
                let Some(targets) = resolve_cell_targets(self, target, limits)? else {
                    return Ok(None);
                };
                if targets.len() != 1 {
                    return Ok(None);
                }
                let spelling = target;
                let target = targets[0];
                let signed = self.signature() == crate::format::signature::Signature::Signed;
                let (present, original_known, original) = if signed {
                    original_scalar(self, spelling)?
                } else {
                    (false, true, None)
                };
                Ok(Some(Plan {
                    target,
                    action: Action::Set(value),
                    staged_bytes: assignments.staged_bytes,
                    original,
                    original_known,
                    present,
                }))
            })();
            self.check_source_for_replay().and(result)
        }

        pub(crate) fn remove_plan(
            &self,
            target: &TargetArg<'_>,
            limits: &SemanticLimits,
        ) -> Result<Option<Plan>> {
            self.check_source_for_replay()?;
            let result = (|| {
                let Some(path) = target.path() else {
                    return Ok(None);
                };
                let Some(step) = path.path().steps().last() else {
                    return Ok(None);
                };
                if step.cell_ref().is_none() {
                    return Ok(None);
                }
                let Some(targets) = resolve_cell_targets(self, target, limits)? else {
                    return Ok(None);
                };
                if targets.len() != 1 {
                    return Ok(None);
                }
                check_remove_limits(1, limits)?;
                let spelling = target;
                let target = targets[0];
                let signed = self.signature() == crate::format::signature::Signature::Signed;
                let (present, original_known, original) = if signed {
                    original_scalar(self, spelling)?
                } else {
                    (false, true, None)
                };
                Ok(Some(Plan {
                    target,
                    action: Action::Remove,
                    staged_bytes: WRITE_ENTRY_OVERHEAD,
                    original,
                    original_known,
                    present,
                }))
            })();
            self.check_source_for_replay().and(result)
        }
    }

    impl Run {
        pub(crate) fn new(
            base: &Xlsb,
            sheet: usize,
            limits: &SemanticLimits,
            remaining_snapshot: usize,
        ) -> Result<Self> {
            if remaining_snapshot == 0 {
                return Err(super::snapshot_limit(
                    limits.max_snapshot_bytes().saturating_add(1),
                    limits.max_snapshot_bytes(),
                ));
            }
            base.check_source_for_replay()?;
            let source = package_from_editor(
                if base.signature() == crate::format::signature::Signature::Signed {
                    base.editor_for_write()?
                } else {
                    base.editor()?
                },
            );
            let native_budget = super::cell_snapshot_budget(remaining_snapshot)?;
            let native_limits = TransferLimits::new(
                native_budget,
                limits.max_expanded_targets(),
                1,
                native_budget,
            );
            let edit = source
                .edit_cell_values_with_options(
                    NativeSelector::from(sheet),
                    NativeCellOptions::new().transfer_limits(native_limits),
                )
                .map_err(native_error)?
                .resolve()
                .map_err(native_error)?
                .ok_or_else(|| {
                    Error::new(
                        Code::OutOfRange,
                        "the XLSB worksheet position is out of range",
                    )
                    .with_format(FileFormat::Xlsb)
                })?;
            base.check_source_for_replay()?;
            Ok(Self {
                base: base.clone(),
                source,
                edit,
                sheet,
                items: Vec::new(),
                seen: HashSet::new(),
                staged_bytes: 0,
                native_limits,
                snapshot_budget: remaining_snapshot,
                max_staged_bytes: limits.max_staged_bytes(),
                allow_repeated: base.signature() == crate::format::signature::Signature::Signed,
            })
        }

        pub(crate) fn cell_plan(
            &self,
            target: &TargetArg<'_>,
            props: &Props<'_>,
            limits: &SemanticLimits,
        ) -> Result<Option<Plan>> {
            self.base.cell_plan(target, props, limits)
        }

        pub(crate) fn remove_plan(
            &self,
            target: &TargetArg<'_>,
            limits: &SemanticLimits,
        ) -> Result<Option<Plan>> {
            self.base.remove_plan(target, limits)
        }

        pub(crate) fn first_index(&self) -> Option<usize> {
            self.items.first().map(|item| item.index)
        }

        pub(crate) fn requires_barrier(&self, plan: &Plan, limits: &SemanticLimits) -> bool {
            plan.target.sheet != self.sheet
                || (self.seen.contains(&plan.key()) && !(self.allow_repeated && plan.present))
                || self
                    .staged_bytes
                    .checked_add(plan.charge())
                    .is_none_or(|bytes| bytes > limits.max_staged_bytes())
        }

        pub(crate) fn stage(&mut self, plan: Plan, index: usize) -> Result<bool> {
            if plan.target.sheet != self.sheet
                || (self.seen.contains(&plan.key()) && !(self.allow_repeated && plan.present))
                || self
                    .staged_bytes
                    .checked_add(plan.charge())
                    .is_none_or(|bytes| bytes > self.max_staged_bytes)
            {
                return Err(Error::new(
                    Code::InternalError,
                    "XLSB cell coalescing barrier was not flushed",
                )
                .with_format(FileFormat::Xlsb));
            }
            self.base.check_source_for_replay()?;
            self.items
                .try_reserve(1)
                .map_err(|source| TransactionError::Allocation { source })?;
            self.seen
                .try_reserve(1)
                .map_err(|source| TransactionError::Allocation { source })?;
            let changed = match &plan.action {
                Action::Set(value) => self
                    .edit
                    .set_scalar(plan.target.reference, value.clone())
                    .map_err(native_error)?,
                Action::Remove => self
                    .edit
                    .remove(plan.target.reference)
                    .map_err(native_error)?,
            };
            self.seen.insert(plan.key());
            self.staged_bytes = self.staged_bytes.saturating_add(plan.charge());
            self.items.push(Item { plan, index });
            self.base.check_source_for_replay()?;
            Ok(changed)
        }

        pub(crate) fn commit(self) -> Result<(Xlsb, Commit)> {
            let Self {
                base,
                source,
                edit,
                items,
                sheet,
                native_limits,
                snapshot_budget,
                allow_repeated,
                ..
            } = self;
            let authored_operations = items.len();
            let logical_noop = allow_repeated && logical_noop(&items);
            let original = if logical_noop {
                Some(base.original_package_for_noop()?)
            } else {
                None
            };
            base.check_source_for_replay()?;
            let result = edit
                .commit()
                .map_err(native_error)
                .and_then(|native| {
                    Commit::from_cell_native(
                        &base,
                        native,
                        sheet,
                        native_limits,
                        snapshot_budget,
                        authored_operations,
                    )
                })
                .and_then(|commit| {
                    if logical_noop {
                        let staged_snapshot_bytes = commit.snapshot_bytes;
                        drop(commit);
                        let original = original
                            .as_ref()
                            .ok_or_else(|| Error::signed(FileFormat::Xlsb))?;
                        let empty = original
                            .edit_cell_values_with_options(
                                NativeSelector::from(sheet),
                                NativeCellOptions::new().transfer_limits(native_limits),
                            )
                            .map_err(native_error)?
                            .resolve()
                            .map_err(native_error)?
                            .ok_or_else(|| {
                                Error::new(
                                    Code::OutOfRange,
                                    "the XLSB worksheet position is out of range",
                                )
                                .with_format(FileFormat::Xlsb)
                            })?
                            .commit()
                            .map_err(native_error)?;
                        let mut canonical = Commit::from_cell_native(
                            &base,
                            empty,
                            sheet,
                            native_limits,
                            snapshot_budget,
                            authored_operations,
                        )?;
                        // The native changed probe may retain more than the
                        // canonical signed identity patch.  Keep the peak
                        // staged charge when replacing it with that genuine
                        // empty patch, otherwise a signed cancellation would
                        // under-report its snapshot budget.
                        canonical.snapshot_bytes =
                            canonical.snapshot_bytes.max(staged_snapshot_bytes);
                        Ok(canonical)
                    } else {
                        Ok(commit)
                    }
                });
            let freshness = base.check_source_for_replay();
            match freshness.and(result) {
                Ok(mut commit) => {
                    commit.operations = authored_operations;
                    Ok((base, commit))
                },
                Err(error) => Err(locate_failure(
                    &base,
                    &source,
                    items.as_slice(),
                    native_limits,
                    error,
                )),
            }
        }
    }

    fn logical_noop(items: &[Item]) -> bool {
        let mut final_values = std::collections::HashMap::<CellTarget, Option<NativeScalar>>::new();
        let mut originals =
            std::collections::HashMap::<CellTarget, (bool, Option<NativeScalar>)>::new();
        for item in items {
            let plan = &item.plan;
            originals
                .entry(plan.key())
                .or_insert((plan.original_known, plan.original.clone()));
            let final_value = match &plan.action {
                Action::Set(value) => Some(value.clone()),
                Action::Remove if plan.present => None,
                Action::Remove => plan.original.clone(),
            };
            final_values.insert(plan.key(), final_value);
        }
        !final_values.is_empty()
            && final_values.into_iter().all(|(key, value)| {
                originals
                    .get(&key)
                    .is_some_and(|(known, original)| *known && original == &value)
            })
    }

    /// Read the original scalar without constructing a second public owner.
    /// A missing cell is a known `None`; unsupported formula/error payloads are
    /// unknown and therefore cannot prove a signed cancellation.
    fn original_scalar(
        file: &Xlsb,
        spelling: &TargetArg<'_>,
    ) -> Result<(bool, bool, Option<NativeScalar>)> {
        let Some(path) = spelling.path() else {
            return Ok((false, false, None));
        };
        let node = crate::path::resolve::get(&super::super::Nav::new(file), path)?;
        let Some(state) = node.cell_state()? else {
            return Ok((false, false, None));
        };
        let value = match state {
            super::super::node::CellState::Missing | super::super::node::CellState::Covered(_) => {
                return Ok((false, true, None));
            },
            super::super::node::CellState::Stub => return Ok((false, false, None)),
            super::super::node::CellState::Stored(cell) => match cell.value() {
                litchi_core::sheet::CellValue::Empty => Some(NativeScalar::Blank),
                litchi_core::sheet::CellValue::Bool(value) => Some(NativeScalar::Boolean(*value)),
                litchi_core::sheet::CellValue::Int(value) => {
                    Some(NativeScalar::Number(*value as f64))
                },
                litchi_core::sheet::CellValue::Float(value)
                | litchi_core::sheet::CellValue::DateTime(value) => {
                    Some(NativeScalar::Number(*value))
                },
                litchi_core::sheet::CellValue::String(value) => {
                    Some(NativeScalar::Text(value.clone()))
                },
                litchi_core::sheet::CellValue::Error(_)
                | litchi_core::sheet::CellValue::Formula { .. } => return Ok((true, false, None)),
            },
        };
        Ok((true, true, value))
    }

    fn replay(
        base: &Xlsb,
        source: &litchi_xlsb::Package,
        items: &[Item],
        limits: TransferLimits,
    ) -> Result<litchi_xlsb::Package> {
        base.check_source_for_replay()?;
        let Some(first) = items.first() else {
            return Ok(source.clone());
        };
        let mut edit = source
            .edit_cell_values_with_options(
                NativeSelector::from(first.plan.target.sheet),
                NativeCellOptions::new().transfer_limits(limits),
            )
            .map_err(native_error)?
            .resolve()
            .map_err(native_error)?
            .ok_or_else(|| {
                Error::new(
                    Code::OutOfRange,
                    "the XLSB worksheet position is out of range",
                )
                .with_format(FileFormat::Xlsb)
            })?;
        for item in items {
            let result = match &item.plan.action {
                Action::Set(value) => edit.set_scalar(item.plan.target.reference, value.clone()),
                Action::Remove => edit.remove(item.plan.target.reference),
            };
            result.map_err(|error| native_error(error).with_item(item.index))?;
        }
        let package = edit.commit().map_err(native_error)?.into_package();
        base.check_source_for_replay()?;
        Ok(package)
    }

    fn locate_failure(
        base: &Xlsb,
        source: &litchi_xlsb::Package,
        items: &[Item],
        limits: TransferLimits,
        error: Error,
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
        let prefix = match replay(base, source, &items[..middle], limits) {
            Ok(package) => package,
            Err(error) => return locate_failure(base, source, &items[..middle], limits, error),
        };
        match replay(base, &prefix, &items[middle..], limits) {
            Err(error) => locate_failure(base, &prefix, &items[middle..], limits, error),
            Ok(_) => error.with_item(first.index),
        }
    }
}

struct Counted<W> {
    sink: W,
    bytes: u64,
}

impl<W: Write> Write for Counted<W> {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        let count = self.sink.write(bytes)?;
        self.bytes = self
            .bytes
            .checked_add(count as u64)
            .ok_or_else(|| io::Error::other("output byte count exceeds u64"))?;
        Ok(count)
    }

    fn flush(&mut self) -> io::Result<()> {
        self.sink.flush()
    }
}
