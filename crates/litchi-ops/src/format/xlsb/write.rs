//! Immutable XLSB sheet-property writes and publication.
//!
//! The XLSB unified writes admit sheet names, tab visibility and direct cell scalar
//! payloads (`value`, `clear: true`, and null clears). Target resolution and
//! property coercion happen before the native workbook editor is materialized.
//! The native `Package` transaction remains responsible for exact package
//! replay, binary validation, signature policy, and retained patch proofs.

use std::collections::{BTreeMap, HashSet};
use std::io::{self, Write};
use std::path::Path as FilePath;

use caseless::Caseless as _;
use litchi_core::selector::{CellRange, Path, Predicate, Step};
use litchi_core::{FileFormat, Position};
use litchi_xlsb::cell_values::structure::Edit as NativeStructureEdit;
use litchi_xlsb::cell_values::{
    Edit as NativeCellEdit, Patch as NativeCellPatch, Reference, Scalar as NativeScalar,
    TransferLimits, structure::Patch as NativeStructurePatch,
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
    Structure(NativeStructurePatch),
    Cells(Vec<NativeCellPatch>),
}

impl Patch {
    pub(crate) fn before(&self) -> &[u8] {
        match self {
            Self::Structure(patch) => patch.before(),
            Self::Cells(patches) => patches
                .first()
                .expect("cell patch invariant: at least one worksheet")
                .before(),
        }
    }

    pub(crate) fn after(&self) -> &[u8] {
        match self {
            Self::Structure(patch) => patch.after(),
            Self::Cells(patches) => patches
                .last()
                .expect("cell patch invariant: at least one worksheet")
                .after(),
        }
    }

    /// Charge every retained native image and semantic cell delta. A selector
    /// spanning worksheets keeps one exact package patch per worksheet, so
    /// endpoint images alone would undercharge transaction history.
    pub(crate) fn snapshot_bytes(&self) -> usize {
        match self {
            Self::Structure(patch) => patch.before().len().saturating_add(patch.after().len()),
            Self::Cells(patches) => patches
                .iter()
                .fold(size_of::<Vec<NativeCellPatch>>(), |bytes, patch| {
                    bytes.saturating_add(patch.retained_bytes())
                }),
        }
    }

    pub(crate) fn retained_bytes(&self) -> usize {
        self.snapshot_bytes()
            .saturating_add(self.operation_count().saturating_mul(124))
    }

    pub(crate) fn is_noop(&self) -> bool {
        self.before() == self.after()
    }

    pub(crate) fn operation_count(&self) -> usize {
        match self {
            Self::Structure(patch) => patch.operation_count(),
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
}

impl Commit {
    /// Consume the result for the format-neutral one-shot setter.
    pub(crate) fn into_file(self) -> Xlsb {
        self.after
    }

    /// Consume the result for explicit transaction history.
    pub(crate) fn into_parts(self) -> (Xlsb, Patch) {
        (self.after, self.patch)
    }

    fn from_native(
        base: &Xlsb,
        native: litchi_xlsb::cell_values::structure::Commit,
    ) -> Result<Self> {
        base.check_source()?;
        let result = (|| {
            let (package, patch) = native.into_parts();
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
                patch: Patch::Structure(patch),
            })
        })();
        base.check_source()?;
        result
    }

    fn from_cell_native(base: &Xlsb, native: litchi_xlsb::cell_values::Commit) -> Result<Self> {
        base.check_source()?;
        let result = (|| {
            let (package, patch) = native.into_parts();
            let patch = Patch::Cells(vec![patch]);
            let after = if patch.is_noop() {
                base.clone()
            } else {
                base.adopt_package(package)?
            };
            Ok(Self { after, patch })
        })();
        base.check_source()?;
        result
    }

    fn from_cell_parts(
        base: &Xlsb,
        package: litchi_xlsb::Package,
        patches: Vec<NativeCellPatch>,
    ) -> Result<Self> {
        if patches.is_empty() {
            return Err(Error::new(
                Code::InternalError,
                "an XLSB cell commit must retain at least one worksheet patch",
            )
            .with_format(FileFormat::Xlsb));
        }
        base.check_source()?;
        let result = (|| {
            let patch = Patch::Cells(patches);
            let after = if patch.is_noop() {
                base.clone()
            } else {
                base.adopt_package(package)?
            };
            Ok(Self { after, patch })
        })();
        base.check_source()?;
        result
    }
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
    file.admit_edit()?;
    let props = Props::from_props(props)?;
    if let Some(commit) = set_cells_with_limits(file, &target, &props, limits)? {
        return Ok(commit);
    }
    set_sheet_with_limits(file, target, &props, limits)
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
    if parent.element() != Kind::Workbook {
        return Err(Error::new(
            Code::UnsupportedType,
            "an XLSB sheet can only be added directly under the workbook",
        )
        .with_format(FileFormat::Xlsb));
    }
    if !node.kind().eq_ignore_ascii_case("sheet") {
        return Err(Error::new(
            Code::UnsupportedElement,
            "the XLSB structural verb supports worksheet sheets only",
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

    let package = package_from_editor(file.editor()?);
    let native_limits = TransferLimits::new(
        limits.max_snapshot_bytes(),
        limits.max_expanded_targets(),
        1,
        limits.max_snapshot_bytes(),
    );
    let mut edit = package
        .edit_with_limits(native_limits)
        .map_err(native_error)?;
    edit.insert_sheet(Position::new(position), name, state)
        .map_err(native_error)?;
    let native = edit.commit().map_err(native_error)?;
    let commit = Commit::from_native(file, native)?;
    check_snapshot_limits(&commit.patch, limits)?;
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

/// Remove stored cells selected by a direct cell path, range, or scoped
/// query.  The target expansion and worksheet grouping are shared with cell
/// setters, while the native cell-values owner supplies the deletion
/// semantics: a virtual (unstored) cell is an exact no-op and no row or
/// column shifting is attempted.
pub(crate) fn remove_with_limits(
    file: &Xlsb,
    target: TargetArg<'_>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    file.admit_edit()?;
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
    let editor = file.editor()?;
    let mut package = package_from_editor(editor);
    let mut patches = Vec::new();
    patches
        .try_reserve(grouped.len())
        .map_err(|_| allocation("XLSB cell removal patches"))?;
    let mut remaining_snapshot = limits.max_snapshot_bytes();
    for (sheet, targets) in grouped {
        let semantic_charge = cell_semantic_charge(targets.len(), limits)?;
        let image_budget = remaining_snapshot
            .checked_sub(semantic_charge)
            .ok_or_else(|| snapshot_limit(usize::MAX, limits.max_snapshot_bytes()))?;
        if image_budget == 0 {
            return Err(snapshot_limit(
                limits.max_snapshot_bytes().saturating_add(1),
                limits.max_snapshot_bytes(),
            ));
        }
        let native_limits =
            TransferLimits::new(image_budget, limits.max_expanded_targets(), 1, image_budget);
        let mut edit = package
            .edit_cell_values_with_limits(
                NativeSelector::from(sheet),
                Default::default(),
                native_limits,
            )
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
        let patch_bytes = patch.retained_bytes();
        let retained = patch_bytes
            .checked_add(semantic_charge)
            .ok_or_else(|| snapshot_limit(usize::MAX, limits.max_snapshot_bytes()))?;
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
        package = next;
        patches.push(patch);
    }
    let commit = Commit::from_cell_parts(file, package, patches)?;
    check_snapshot_limits(&commit.patch, limits)?;
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
    file.admit_edit()?;
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
    file.admit_edit()?;
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
    let editor = file.editor()?;
    let package = package_from_editor(editor);
    let native_limits = TransferLimits::new(
        limits.max_snapshot_bytes(),
        limits.max_expanded_targets(),
        1,
        limits.max_snapshot_bytes(),
    );
    let mut edit = package
        .edit_with_limits(native_limits)
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
    let commit = Commit::from_native(file, native)?;
    check_snapshot_limits(&commit.patch, limits)?;
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

    // All target and property admission above is intentionally before this
    // editor construction. Signed and retention policy are still checked by
    // `Xlsb::check_source` at the outer fence and by `editor` here.
    let editor = file.editor()?;
    let package = package_from_editor(editor);
    let native_limits = TransferLimits::new(
        limits.max_snapshot_bytes(),
        limits.max_expanded_targets().saturating_mul(2),
        1,
        limits.max_snapshot_bytes(),
    );
    let mut edit = package
        .edit_with_limits(native_limits)
        .map_err(native_error)?;
    if let Some(name) = assignments.name {
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
    if let Some(state) = assignments.state {
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
    let commit = Commit::from_native(file, native)?;
    check_snapshot_limits(&commit.patch, limits)?;
    Ok(commit)
}

#[derive(Debug, Clone)]
struct CellAssignments {
    value: Option<NativeScalar>,
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
    let assignments = match admitted {
        Some(assignments) => assignments,
        None => admit_cell_properties(props, limits)?,
    };
    check_cell_limits(targets.len(), &assignments, limits)?;

    // Complete target grouping is still admission work: it must finish before
    // the native package owner is materialized. Every target in one worksheet
    // is staged in one native cell-values edit; a selector spanning worksheets
    // publishes the edits in workbook order through one exact package chain.
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
    let editor = file.editor()?;
    let mut package = package_from_editor(editor);
    let mut patches = Vec::new();
    patches
        .try_reserve(grouped.len())
        .map_err(|_| allocation("XLSB cell write patches"))?;
    let mut remaining_snapshot = limits.max_snapshot_bytes();
    for (sheet, targets) in grouped {
        let semantic_charge = cell_semantic_charge(targets.len(), limits)?;
        let image_budget = remaining_snapshot
            .checked_sub(semantic_charge)
            .ok_or_else(|| snapshot_limit(usize::MAX, limits.max_snapshot_bytes()))?;
        if image_budget == 0 {
            return Err(snapshot_limit(
                limits.max_snapshot_bytes().saturating_add(1),
                limits.max_snapshot_bytes(),
            ));
        }
        let native_limits =
            TransferLimits::new(image_budget, limits.max_expanded_targets(), 1, image_budget);
        let mut edit = package
            .edit_cell_values_with_limits(
                NativeSelector::from(sheet),
                Default::default(),
                native_limits,
            )
            .map_err(native_error)?
            .ok_or_else(|| {
                Error::new(
                    Code::OutOfRange,
                    "the XLSB worksheet position is out of range",
                )
                .with_format(FileFormat::Xlsb)
            })?;
        for target in targets {
            if let Some(value) = assignments.value.as_ref() {
                edit.set_scalar(target.reference, value.clone())
                    .map_err(native_error)?;
            }
        }
        let native = edit.commit().map_err(native_error)?;
        let (next, patch) = native.into_parts();
        let patch_bytes = patch.retained_bytes();
        let retained = patch_bytes
            .checked_add(semantic_charge)
            .ok_or_else(|| snapshot_limit(usize::MAX, limits.max_snapshot_bytes()))?;
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
        package = next;
        patches.push(patch);
    }
    let commit = Commit::from_cell_parts(file, package, patches)?;
    check_snapshot_limits(&commit.patch, limits)?;
    Ok(Some(commit))
}

fn admit_cell_properties(props: &Props<'_>, limits: &SemanticLimits) -> Result<CellAssignments> {
    let mut assignments = CellAssignments {
        value: None,
        staged_bytes: WRITE_ENTRY_OVERHEAD,
    };
    for (key, value) in props.iter() {
        let (_, entry) = CELL_WRITE_TABLE
            .resolve(key.as_str())
            .map_err(|error| error.with_property(key.as_str().to_owned()))?;
        let entry = *entry;
        entry
            .prepare(Operation::Set, value, |setter, value| match setter {
                CellSetter::Value => {
                    if assignments.value.is_some() {
                        return Err(duplicate_cell_property("value"));
                    }
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
                    Ok(())
                },
                CellSetter::Clear => {
                    if assignments.value.is_some() {
                        return Err(Error::new(
                            Code::InvalidArgument,
                            "cell properties value and clear cannot be supplied together",
                        )
                        .with_format(FileFormat::Xlsb)
                        .with_property("clear"));
                    }
                    let Some(Scalar::Bool(true)) = value else {
                        return Err(Error::new(Code::InvalidValue, "clear must be true")
                            .with_format(FileFormat::Xlsb)
                            .with_property("clear"));
                    };
                    assignments.staged_bytes = assignments
                        .staged_bytes
                        .checked_add(1)
                        .ok_or_else(|| staged_limit(usize::MAX, limits.max_staged_bytes()))?;
                    assignments.value = Some(NativeScalar::Blank);
                    Ok(())
                },
            })
            .map_err(|error| error.with_format(FileFormat::Xlsb))?;
    }
    if assignments.value.is_none() && !props.is_empty() {
        return Err(Error::new(
            Code::InvalidArgument,
            "an XLSB cell write requires value or clear",
        )
        .with_format(FileFormat::Xlsb));
    }
    Ok(assignments)
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

fn cell_semantic_charge(targets: usize, limits: &SemanticLimits) -> Result<usize> {
    let operations = targets
        .checked_mul(124)
        .ok_or_else(|| snapshot_limit(usize::MAX, limits.max_snapshot_bytes()))?;
    Ok(operations)
}

fn duplicate_cell_property(property: &'static str) -> Error {
    Error::new(
        Code::InvalidArgument,
        "the same XLSB cell property was supplied more than once",
    )
    .with_format(FileFormat::Xlsb)
    .with_property(property)
}

fn package_from_editor(editor: XlsbEditor<'_>) -> litchi_xlsb::Package {
    match editor {
        XlsbEditor::Eager(workbook) => workbook.to_package(),
        XlsbEditor::Materialized(workbook) => workbook.into_package(),
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
        let walk = crate::query::run(Nav::new(file), query.clone())?;
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
    file.check_source()?;
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
        let walk = crate::query::run(Nav::new(file), query.clone())?;
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
        file.check_source()?;
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

    file.check_source()?;
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

fn check_snapshot_limits(patch: &Patch, limits: &SemanticLimits) -> Result<()> {
    let observed = patch.snapshot_bytes();
    if observed > limits.max_snapshot_bytes() {
        return Err(snapshot_limit(observed, limits.max_snapshot_bytes()));
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

fn snapshot_limit(observed: usize, maximum: usize) -> Error {
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

fn native_error(error: litchi_xlsb::package::PackageError) -> Error {
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
                .unsigned_workbook()?
                .into_package()
                .write_to(&mut sink)
                .map_err(native_error);
        }
        match &file.0.owner {
            XlsbOwner::Source(workbook) => workbook
                .workbook()
                .source_backed_package()
                .source_artifact()
                .write_to_stream(&mut sink)
                .map_err(Error::from),
            XlsbOwner::Eager(workbook) => workbook
                .to_package()
                .write_to(&mut sink)
                .map_err(native_error),
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
            self.check_source()?;
            let Some((plan, target)) = admit_sheet_state_plan(self, target, props, limits)? else {
                return Ok(None);
            };
            validate_sheet_state(self, std::slice::from_ref(&target), plan.state)?;
            self.check_source().and(Ok(Some(plan)))
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
            base.check_source()?;
            let source = package_from_editor(base.editor()?);
            let native_limits = NativeTransferLimits::new(
                remaining_snapshot,
                limits.max_expanded_targets(),
                1,
                remaining_snapshot,
            );
            let edit = source
                .edit_with_limits(native_limits)
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
            base.check_source()?;
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
            self.base.check_source()?;
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
            self.base.check_source().and(Ok(Some(plan)))
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
            self.base.check_source()?;
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
            self.base.check_source()?;
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
            base.check_source()?;
            let result = edit
                .commit()
                .map_err(native_error)
                .and_then(|native| Commit::from_native(&base, native));
            let freshness = base.check_source();
            match freshness.and(result) {
                Ok(commit) => Ok((base, commit)),
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
        base.check_source()?;
        if items.is_empty() {
            return Ok(source.clone());
        }
        let mut edit = source.edit_with_limits(limits).map_err(native_error)?;
        for item in items {
            edit.set_sheet_state(item.plan.tab, item.plan.state)
                .map_err(|error| native_error(error).with_item(item.index))?;
        }
        let package = edit.commit().map_err(native_error)?.into_package();
        base.check_source()?;
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
    use litchi_xlsb::cell_values::{Limits, TransferLimits};
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
        max_staged_bytes: usize,
    }

    impl Xlsb {
        pub(crate) fn cell_plan(
            &self,
            target: &TargetArg<'_>,
            props: &Props<'_>,
            limits: &SemanticLimits,
        ) -> Result<Option<Plan>> {
            self.check_source()?;
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
                let target = targets[0];
                Ok(Some(Plan {
                    target,
                    action: Action::Set(value),
                    staged_bytes: assignments.staged_bytes,
                }))
            })();
            self.check_source().and(result)
        }

        pub(crate) fn remove_plan(
            &self,
            target: &TargetArg<'_>,
            limits: &SemanticLimits,
        ) -> Result<Option<Plan>> {
            self.check_source()?;
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
                Ok(Some(Plan {
                    target: targets[0],
                    action: Action::Remove,
                    staged_bytes: WRITE_ENTRY_OVERHEAD,
                }))
            })();
            self.check_source().and(result)
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
            base.check_source()?;
            let source = package_from_editor(base.editor()?);
            let native_limits = TransferLimits::new(
                remaining_snapshot,
                limits.max_expanded_targets(),
                1,
                remaining_snapshot,
            );
            let edit = source
                .edit_cell_values_with_limits(
                    NativeSelector::from(sheet),
                    Limits::DEFAULT,
                    native_limits,
                )
                .map_err(native_error)?
                .ok_or_else(|| {
                    Error::new(
                        Code::OutOfRange,
                        "the XLSB worksheet position is out of range",
                    )
                    .with_format(FileFormat::Xlsb)
                })?;
            base.check_source()?;
            Ok(Self {
                base: base.clone(),
                source,
                edit,
                sheet,
                items: Vec::new(),
                seen: HashSet::new(),
                staged_bytes: 0,
                native_limits,
                max_staged_bytes: limits.max_staged_bytes(),
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
                || self.seen.contains(&plan.key())
                || self
                    .staged_bytes
                    .checked_add(plan.charge())
                    .is_none_or(|bytes| bytes > limits.max_staged_bytes())
        }

        pub(crate) fn stage(&mut self, plan: Plan, index: usize) -> Result<bool> {
            if plan.target.sheet != self.sheet
                || self.seen.contains(&plan.key())
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
            self.base.check_source()?;
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
            self.base.check_source()?;
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
            base.check_source()?;
            let result = edit
                .commit()
                .map_err(native_error)
                .and_then(|native| Commit::from_cell_native(&base, native));
            let freshness = base.check_source();
            match freshness.and(result) {
                Ok(commit) => Ok((base, commit)),
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
        limits: TransferLimits,
    ) -> Result<litchi_xlsb::Package> {
        base.check_source()?;
        let Some(first) = items.first() else {
            return Ok(source.clone());
        };
        let mut edit = source
            .edit_cell_values_with_limits(
                NativeSelector::from(first.plan.target.sheet),
                Limits::DEFAULT,
                limits,
            )
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
        base.check_source()?;
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
