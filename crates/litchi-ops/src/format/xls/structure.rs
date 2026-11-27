//! Unified XLS worksheet-tab structure: add, remove, move and swap.
//!
//! The adapter resolves public paths and placements against the complete tab
//! catalog (chart and macro tabs included) and hands one exclusive native
//! `Package::edit` tab change to the leaf.  The leaf rewrites the tab
//! directory and every workbook-global tab address in one closure and refuses
//! changes whose dependencies it cannot close; no BIFF record is rebuilt here.

use litchi_core::selector::{Path, Predicate, Step};
use litchi_core::{FileFormat, Position};
use litchi_xls::cell_values::Selector;

use super::spec::{Kind, SHEET_WRITE_TABLE, SheetSetter};
use super::write::{WriteCommit, finish_commit, native_edit_error};
use super::{Nav, Xls};
use crate::error::{Code, Error, Result};
use crate::new_node::NewNode;
use crate::path::PathArg;
use crate::placement::Placement;
use crate::registry::Operation;
use crate::target::TargetArg;
use crate::transaction::{Error as TransactionError, LimitKind, SemanticLimits};
use crate::value::Scalar;

/// Canonical `/sheet[N]` path of a catalog position.
fn sheet_path(position: usize) -> Path<'static> {
    Path::from_steps([Step::new("sheet").with(Predicate::Index(Position::new(position)))])
}

fn refusal(code: Code, message: &'static str) -> Error {
    Error::new(code, message).with_format(FileFormat::Xls)
}

fn limit(kind: LimitKind, observed: usize, maximum: usize) -> Error {
    Error::from(TransactionError::Limit {
        kind,
        observed,
        maximum,
    })
    .with_format(FileFormat::Xls)
}

/// Admit one tab-structure operation against the caller's semantic budget.
fn admit(targets: usize, staged: usize, limits: &SemanticLimits) -> Result<()> {
    if limits.max_operations() < 1 {
        return Err(limit(LimitKind::Operations, 1, limits.max_operations()));
    }
    if targets > limits.max_expanded_targets() {
        return Err(limit(
            LimitKind::ExpandedTargets,
            targets,
            limits.max_expanded_targets(),
        ));
    }
    if staged > limits.max_staged_bytes() {
        return Err(limit(
            LimitKind::StagedBytes,
            staged,
            limits.max_staged_bytes(),
        ));
    }
    Ok(())
}

/// One permutation selector per catalog tab, charged at its native width.
fn order_charge(count: usize) -> usize {
    count.saturating_mul(size_of::<Selector<'_>>())
}

pub(super) fn tab_count(file: &Xls) -> Result<usize> {
    super::nav::tab_count(file)
}

fn resolve_sheet(file: &Xls, path: &PathArg<'_>) -> Result<usize> {
    let node = crate::path::resolve::get(&Nav::new(file), path)?;
    if node.element() != Kind::Sheet {
        return Err(refusal(
            Code::UnsupportedType,
            "the XLS structural verb requires a sheet tab",
        ));
    }
    Ok(node.sheet_position())
}

fn require_workbook(file: &Xls, path: &PathArg<'_>) -> Result<()> {
    let node = crate::path::resolve::get(&Nav::new(file), path)?;
    if node.element() == Kind::Workbook {
        Ok(())
    } else {
        Err(refusal(
            Code::UnsupportedType,
            "XLS sheet tabs live directly under the workbook",
        ))
    }
}

/// Add one empty worksheet tab and return its canonical path.
pub(crate) fn add_sheet(
    file: &Xls,
    parent: &PathArg<'_>,
    node: &NewNode<'_>,
    limits: &SemanticLimits,
) -> Result<(WriteCommit, Path<'static>)> {
    file.admit_edit_for_write()?;
    require_workbook(file, parent)?;
    let mut name = None;
    for (key, value) in node.props().iter() {
        let (_, entry) = SHEET_WRITE_TABLE.resolve(key.as_str()).map_err(|error| {
            error
                .with_format(FileFormat::Xls)
                .with_property(key.as_str().to_owned())
        })?;
        let entry = *entry;
        entry
            .prepare(Operation::Add, value, |setter, value| match setter {
                SheetSetter::Name => {
                    if name.is_some() {
                        return Err(refusal(
                            Code::InvalidArgument,
                            "the same XLS sheet property was supplied more than once",
                        )
                        .with_property("name"));
                    }
                    let Some(Scalar::Text(value)) = value else {
                        return Err(refusal(
                            Code::InvalidValue,
                            "an XLS worksheet name requires text",
                        )
                        .with_property("name"));
                    };
                    name = Some(value.to_owned());
                    Ok(())
                },
            })
            .map_err(|error| error.with_format(FileFormat::Xls))?;
    }
    let count = tab_count(file)?;
    let position = add_position(file, count, node.placement())?;
    let name = match name {
        Some(name) => name,
        None => default_name(file, count)?,
    };
    validate_sheet_name(file, None, &name)?;
    admit(1, name.len().saturating_add(64), limits)?;
    file.preflight_snapshot_bytes(limits.max_snapshot_bytes(), name.len().saturating_add(64))?;
    let mut edit = file
        .editable_package_with_snapshot_limit(limits.max_snapshot_bytes())?
        .edit()
        .map_err(native_edit_error)?;
    edit.insert_sheet(Position::new(position), &name)
        .map_err(native_edit_error)?;
    let native = edit.commit().map_err(native_edit_error)?;
    Ok((
        finish_commit(file, native, 1, limits)?,
        sheet_path(position),
    ))
}

fn add_position(file: &Xls, count: usize, placement: Option<&Placement<'_>>) -> Result<usize> {
    let position = match placement {
        None | Some(Placement::End) => count,
        Some(Placement::Start) => 0,
        Some(Placement::Position(position)) => position.get(),
        Some(Placement::Before(path)) => resolve_sheet(file, path)?,
        Some(Placement::After(path)) => resolve_sheet(file, path)?.saturating_add(1),
        Some(Placement::Into(path, position)) => {
            require_workbook(file, path)?;
            position.map_or(count, Position::get)
        },
    };
    if position > count {
        return Err(refusal(
            Code::OutOfRange,
            "the XLS sheet insertion position is outside the workbook",
        ));
    }
    Ok(position)
}

/// Refuse an invalid BIFF8 tab name, or one another tab already holds under
/// the registered Unicode caseless identity.  `except` is the tab being
/// renamed, which may keep or recase its own name.
pub(crate) fn validate_sheet_name(file: &Xls, except: Option<usize>, name: &str) -> Result<()> {
    if !litchi_xls::sheet_name::is_valid(name) {
        return Err(
            refusal(Code::InvalidValue, "the XLS sheet name is invalid").with_property("name")
        );
    }
    let workbook = file.owner().workbook();
    for position in 0..tab_count(file)? {
        if Some(position) == except {
            continue;
        }
        if workbook
            .sheet(position)?
            .is_some_and(|tab| litchi_xls::sheet_name::equivalent(tab.name(), name))
        {
            return Err(
                refusal(Code::DuplicateName, "a sheet name is already in use")
                    .with_property("name"),
            );
        }
    }
    Ok(())
}

/// The first `SheetN` name that no catalog tab already holds.
fn default_name(file: &Xls, count: usize) -> Result<String> {
    let workbook = file.owner().workbook();
    let mut existing = Vec::new();
    existing
        .try_reserve_exact(count)
        .map_err(|_| refusal(Code::InternalError, "could not allocate XLS tab names"))?;
    for position in 0..count {
        if let Some(tab) = workbook.sheet(position)? {
            existing.push(tab.name().to_owned());
        }
    }
    (1..=count.saturating_add(1))
        .map(|number| format!("Sheet{number}"))
        .find(|candidate| {
            !existing
                .iter()
                .any(|name| litchi_xls::sheet_name::equivalent(name, candidate))
        })
        .ok_or_else(|| refusal(Code::OutOfRange, "no default XLS sheet name is available"))
}

/// Remove one worksheet tab selected by a sheet path or a query whose every
/// match is that one sheet.  Returns `None` when the target is not a sheet.
pub(crate) fn remove_sheet(
    file: &Xls,
    target: &TargetArg<'_>,
    limits: &SemanticLimits,
) -> Result<Option<WriteCommit>> {
    let Some(position) = sheet_removal_target(file, target)? else {
        return Ok(None);
    };
    admit(1, 64, limits)?;
    file.preflight_snapshot_bytes(limits.max_snapshot_bytes(), 64)?;
    let mut edit = file
        .editable_package_with_snapshot_limit(limits.max_snapshot_bytes())?
        .edit()
        .map_err(native_edit_error)?;
    edit.remove_sheet(Selector::Position(Position::new(position)))
        .map_err(native_edit_error)?;
    let native = edit.commit().map_err(native_edit_error)?;
    finish_commit(file, native, 1, limits).map(Some)
}

fn sheet_removal_target(file: &Xls, target: &TargetArg<'_>) -> Result<Option<usize>> {
    if let Some(path) = target.path() {
        let node = crate::path::resolve::get(&Nav::new(file), path)?;
        return Ok((node.element() == Kind::Sheet).then(|| node.sheet_position()));
    }
    let Some(query) = target.query() else {
        return Ok(None);
    };
    let walk = crate::query::run(Nav::new(file), query)?;
    if !walk.plan().diagnostics().is_empty() {
        return Err(refusal(
            Code::InvalidSelector,
            "the mutation selector contains an unsupported predicate",
        ));
    }
    let mut found = None;
    for node in walk {
        let node = node?;
        if node.element() != Kind::Sheet {
            if found.is_some() {
                return Err(refusal(
                    Code::InvalidArgument,
                    "an XLS removal query must select only sheets or only cells",
                ));
            }
            return Ok(None);
        }
        if found.replace(node.sheet_position()).is_some() {
            return Err(refusal(
                Code::InvalidArgument,
                "XLS removes one worksheet tab per call",
            ));
        }
    }
    file.check_source()?;
    // An empty match set is left to the cell path, where it is an exact
    // no-op like every other virtual removal target.
    Ok(found)
}

/// Move one tab with the shared placement vocabulary.  Positional
/// destinations are measured after detaching the tab.
pub(crate) fn move_sheet(
    file: &Xls,
    path: &PathArg<'_>,
    placement: &Placement<'_>,
    limits: &SemanticLimits,
) -> Result<WriteCommit> {
    file.admit_edit_for_write()?;
    let source = resolve_sheet(file, path)?;
    let count = tab_count(file)?;
    let last = count
        .checked_sub(1)
        .ok_or_else(|| refusal(Code::PathNotFound, "the XLS workbook has no sheet tabs"))?;
    let destination = match placement {
        Placement::Start => 0,
        Placement::End => last,
        Placement::Position(position) => position.get(),
        Placement::Before(anchor) => {
            let anchor = resolve_sheet(file, anchor)?;
            if source < anchor { anchor - 1 } else { anchor }
        },
        Placement::After(anchor) => {
            let anchor = resolve_sheet(file, anchor)?;
            if source <= anchor { anchor } else { anchor + 1 }
        },
        Placement::Into(parent, position) => {
            require_workbook(file, parent)?;
            position.map_or(last, Position::get)
        },
    };
    if destination > last {
        return Err(refusal(
            Code::OutOfRange,
            "the XLS sheet move position is outside the workbook",
        ));
    }
    let mut order: Vec<usize> = (0..count).collect();
    let tab = order.remove(source);
    order.insert(destination, tab);
    reorder(file, &order, 1, limits)
}

/// Exchange two tabs in one native edit.
pub(crate) fn swap_sheets(
    file: &Xls,
    left: &PathArg<'_>,
    right: &PathArg<'_>,
    limits: &SemanticLimits,
) -> Result<WriteCommit> {
    file.admit_edit_for_write()?;
    let left = resolve_sheet(file, left)?;
    let right = resolve_sheet(file, right)?;
    let count = tab_count(file)?;
    let mut order: Vec<usize> = (0..count).collect();
    order.swap(left, right);
    reorder(file, &order, if left == right { 1 } else { 2 }, limits)
}

fn reorder(
    file: &Xls,
    order: &[usize],
    targets: usize,
    limits: &SemanticLimits,
) -> Result<WriteCommit> {
    admit(targets, order_charge(order.len()), limits)?;
    let mut selectors = Vec::new();
    selectors
        .try_reserve_exact(order.len())
        .map_err(|_| refusal(Code::InternalError, "could not allocate the XLS tab order"))?;
    selectors.extend(order.iter().copied().map(Selector::from));
    file.preflight_snapshot_bytes(limits.max_snapshot_bytes(), order_charge(order.len()))?;
    let mut edit = file
        .editable_package_with_snapshot_limit(limits.max_snapshot_bytes())?
        .edit()
        .map_err(native_edit_error)?;
    edit.reorder_sheets(&selectors).map_err(native_edit_error)?;
    let native = edit.commit().map_err(native_edit_error)?;
    finish_commit(file, native, 1, limits)
}

impl Xls {
    /// Add a worksheet tab under the workbook root.
    pub(crate) fn add_with_limits<'p, 'n>(
        &self,
        parent: PathArg<'p>,
        node: NewNode<'n>,
        limits: &SemanticLimits,
    ) -> Result<(WriteCommit, Path<'static>)> {
        self.finish(|| {
            if let Some((kind, implied)) =
                super::feature_write::feature_kind(node.kind(), node.props())
            {
                self.admit_edit_for_write()?;
                let parent_node = crate::path::resolve::get(&Nav::new(self), &parent)?;
                if matches!(
                    parent_node.element(),
                    Kind::Sheet | Kind::Cell | Kind::Range
                ) {
                    return self.add_feature(&parent_node, kind, implied, &node, limits);
                }
            }
            if !node.kind().eq_ignore_ascii_case("sheet") {
                return super::layout::add_axis(self, &parent, &node, limits);
            }
            add_sheet(self, &parent, &node, limits)
        })
    }

    /// Move one worksheet tab.
    pub(crate) fn move_to_with_limits<'p, 'q>(
        &self,
        path: PathArg<'p>,
        placement: Placement<'q>,
        limits: &SemanticLimits,
    ) -> Result<WriteCommit> {
        self.finish(|| move_sheet(self, &path, &placement, limits))
    }

    /// Swap two worksheet tabs.
    pub(crate) fn swap_with_limits<'p, 'q>(
        &self,
        left: PathArg<'p>,
        right: PathArg<'q>,
        limits: &SemanticLimits,
    ) -> Result<WriteCommit> {
        self.finish(|| swap_sheets(self, &left, &right, limits))
    }
}
