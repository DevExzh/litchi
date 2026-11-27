//! Immutable XLSX merged-range writes.
//!
//! SpreadsheetML keeps merged geometry separate from cell records.  The
//! unified write route therefore resolves merge membership from the retained
//! grid first and stages only exact native merge intents.  In particular, an
//! unmerge is never delegated to the native "any cell in a merge" lookup with
//! an unverified range: a broad or partially overlapping range is refused so
//! a caller cannot accidentally remove a different merge.

use std::collections::TryReserveError;
use std::collections::{HashMap, HashSet};

use litchi_core::FileFormat;
use litchi_xlsx::{self as xlsx, Area, Rect};

use super::node::{Grid, Node, SheetHandle};
use super::sheet_write::{SheetAssignments, admit_sheet_properties};
use super::spec::{MERGE_WRITE_TABLE, MergeSetter};
use super::write::{Commit, PreparedCells, edit_with_snapshot_limit};
use super::{Nav, Xlsx};
use crate::error::{Code, Error, Result};
use crate::props::{Key, Props};
use crate::registry::Operation;
use crate::target::TargetArg;
use crate::transaction::{Error as TransactionError, LimitKind, SemanticLimits};
use crate::value::Scalar;

const MERGE_ENTRY_OVERHEAD: usize = 64;

#[derive(Debug)]
struct MergeAssignment {
    mode: MergeMode,
}

#[derive(Debug)]
enum MergeMode {
    Enable,
    Disable,
    Sweep,
    Explicit(Vec<Rect>),
}

impl MergeMode {
    fn adds(&self) -> bool {
        matches!(self, Self::Enable | Self::Explicit(_))
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
struct MergeTarget {
    sheet: usize,
    range: Rect,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum TargetFamily {
    Range,
    Cell,
    Sheet,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
enum SelectedNode {
    Range {
        sheet: usize,
        range: Rect,
    },
    Cell {
        sheet: usize,
        address: xlsx::Address,
    },
    Sheet {
        sheet: usize,
    },
}

/// Try the XLSX merged-range property route.
///
/// `Ok(None)` means that no `merge` property was supplied and another write
/// family should admit the request.  Once the property is present, this route
/// owns admission and target classification; a mixed selector or property is
/// refused rather than silently narrowed to one mutation family.
pub(crate) fn try_set_merge(
    file: &Xlsx,
    target: &TargetArg<'_>,
    props: &Props<'_>,
    limits: &SemanticLimits,
) -> Result<Option<Commit>> {
    if !props
        .iter()
        .any(|(key, _)| key.eq_ignore_ascii_case("merge"))
    {
        return Ok(None);
    }

    file.check_source_for_replay()?;
    let result = (|| {
        let assignment = admit_merge(props, limits)?;
        let resolved = resolve_targets(file, target, &assignment.mode, limits)?;

        let other_props = non_merge_props(props);
        let (prepared_cells, sheet_assignments) = if other_props.is_empty() {
            (None, None)
        } else if resolved.family == Some(TargetFamily::Sheet) {
            (
                None,
                Some(admit_sheet_properties(Operation::Set, &other_props)?),
            )
        } else {
            (
                Some(PreparedCells::new(file, target, &other_props, limits)?),
                None,
            )
        };
        if sheet_assignments
            .as_ref()
            .is_some_and(|assignments| assignments.name.is_some())
            && resolved.sheet_positions.len() > 1
        {
            return Err(Error::new(
                Code::InvalidArgument,
                "renaming a selector with multiple sheet matches is ambiguous",
            )
            .with_format(FileFormat::Xlsx)
            .with_property("name"));
        }
        check_combined_limits(
            resolved.targets.len(),
            prepared_cells.as_ref(),
            sheet_assignments.as_ref(),
            resolved.sheet_positions.len(),
            limits,
        )?;

        // The native editor performs the protected-sheet, unknown-content,
        // formula, and compatibility checks.  Every merge and any companion
        // cell write is staged in that one transaction.
        let source = file.editor_for_write_with_snapshot_limit(limits.max_snapshot_bytes())?;
        let mut edit = edit_with_snapshot_limit(&source, limits.max_snapshot_bytes())?;
        apply_merge_targets(&mut edit, &resolved.targets, assignment.mode.adds())?;
        if let Some(cells) = prepared_cells {
            cells.apply(&mut edit)?;
        }
        if let Some(assignments) = sheet_assignments {
            apply_sheet_assignments(file, &mut edit, &resolved.sheet_positions, &assignments)?;
        }
        let native = edit.commit().map_err(native_error)?;
        Commit::from_native_with_limit(file, source, native, limits.max_snapshot_bytes()).map(Some)
    })();
    file.check_source_for_replay()?;
    result
}

impl Xlsx {
    /// Set or clear the OfficeCLI `merge` property through one native edit.
    pub(crate) fn try_set_merge(
        &self,
        target: &TargetArg<'_>,
        props: &Props<'_>,
        limits: &SemanticLimits,
    ) -> Result<Option<Commit>> {
        try_set_merge(self, target, props, limits)
    }
}

fn admit_merge(props: &Props<'_>, limits: &SemanticLimits) -> Result<MergeAssignment> {
    let mut assignment = None;
    for (key, value) in props.iter() {
        // A merge route can share one native edit with ordinary cell
        // properties.  Those keys belong to PreparedCells; resolving them
        // against the merge table here would reject an otherwise valid mixed
        // assignment before the cell family sees it.
        if !key.eq_ignore_ascii_case("merge") {
            continue;
        }
        let (_, entry) = MERGE_WRITE_TABLE
            .resolve(key.as_str())
            .map_err(|error| property_error(error, key))?;
        let entry = *entry;
        entry
            .prepare(Operation::Set, value, |setter, value| {
                match setter {
                    MergeSetter::Merge => {
                        if assignment.is_some() {
                            return Err(duplicate_property(entry.spec.name));
                        }
                        assignment = Some(MergeAssignment {
                            mode: match value {
                                Some(value) => parse_merge_value(value, limits)?,
                                None => MergeMode::Disable,
                            },
                        });
                    },
                }
                Ok(())
            })
            .map_err(|error| property_error(error, key))?;
    }
    assignment.ok_or_else(|| {
        Error::new(
            Code::InternalError,
            "merge property admission produced no assignment",
        )
        .with_format(FileFormat::Xlsx)
    })
}

fn parse_merge_value(value: Scalar<'_>, limits: &SemanticLimits) -> Result<MergeMode> {
    match value {
        Scalar::Bool(true) => Ok(MergeMode::Enable),
        Scalar::Bool(false) => Ok(MergeMode::Disable),
        Scalar::Text(text) => parse_merge_text(text, limits),
        _ => Err(invalid_property(
            "merge",
            "the merge property must be boolean or a bounded range operation",
        )),
    }
}

fn parse_merge_text(text: &str, limits: &SemanticLimits) -> Result<MergeMode> {
    // Bound ingress work even if a long list deduplicates to one rectangle.
    // Check before trimming so whitespace cannot evade the same byte bound.
    if text.len() > limits.max_staged_bytes() {
        return Err(staged_limit(text.len(), limits.max_staged_bytes()));
    }
    let text = text.trim();
    if text.eq_ignore_ascii_case("true") || text == "1" || text.eq_ignore_ascii_case("yes") {
        return Ok(MergeMode::Enable);
    }
    if text.is_empty()
        || text == "0"
        || text.eq_ignore_ascii_case("no")
        || text.eq_ignore_ascii_case("false")
        || text.eq_ignore_ascii_case("none")
        || text.eq_ignore_ascii_case("empty")
        || text.eq_ignore_ascii_case("unmerge")
    {
        return Ok(MergeMode::Disable);
    }
    if text.eq_ignore_ascii_case("sweep") {
        return Ok(MergeMode::Sweep);
    }
    let mut ranges = Vec::new();
    let mut seen = HashSet::new();
    for item in text.split(',') {
        let item = item.trim();
        if item.is_empty() {
            return Err(invalid_property(
                "merge",
                "the merge range list contains an empty item",
            ));
        }
        let range = Rect::from_a1(item)
            .map_err(|_| invalid_property("merge", "the merge range is invalid"))?;
        if seen.contains(&range) {
            continue;
        }
        seen.try_reserve(1).map_err(allocation)?;
        seen.insert(range);
        let count = ranges
            .len()
            .checked_add(1)
            .ok_or_else(|| expanded_limit(usize::MAX, limits.max_expanded_targets()))?;
        if count > limits.max_expanded_targets() {
            return Err(expanded_limit(count, limits.max_expanded_targets()));
        }
        let staged = count
            .checked_mul(MERGE_ENTRY_OVERHEAD)
            .ok_or_else(|| staged_limit(usize::MAX, limits.max_staged_bytes()))?;
        if staged > limits.max_staged_bytes() {
            return Err(staged_limit(staged, limits.max_staged_bytes()));
        }
        ranges.try_reserve(1).map_err(allocation)?;
        ranges.push(range);
    }
    if ranges.is_empty() {
        return Err(invalid_property("merge", "the merge range list is empty"));
    }
    Ok(MergeMode::Explicit(ranges))
}

fn non_merge_props<'a>(props: &'a Props<'a>) -> Props<'a> {
    let mut result = Props::with_capacity(props.len().saturating_sub(1));
    for (key, value) in props.iter() {
        if !key.eq_ignore_ascii_case("merge") {
            result.push(key.reborrow(), value.reborrow());
        }
    }
    result
}

struct ResolvedTargets<'a> {
    targets: Vec<MergeTarget>,
    seen: HashSet<MergeTarget>,
    selected: HashSet<SelectedNode>,
    sheet_positions: Vec<usize>,
    grids: HashMap<usize, Grid>,
    expanded_lists: HashSet<usize>,
    mode: &'a MergeMode,
    family: Option<TargetFamily>,
    limits: &'a SemanticLimits,
}

fn resolve_targets<'a>(
    file: &Xlsx,
    target: &TargetArg<'_>,
    mode: &'a MergeMode,
    limits: &'a SemanticLimits,
) -> Result<ResolvedTargets<'a>> {
    let mut resolved = ResolvedTargets {
        targets: Vec::new(),
        seen: HashSet::new(),
        selected: HashSet::new(),
        sheet_positions: Vec::new(),
        grids: HashMap::new(),
        expanded_lists: HashSet::new(),
        mode,
        family: None,
        limits,
    };
    if let Some(path) = target.path() {
        let node = crate::path::resolve::get(&Nav::new(file), path)?;
        resolved.consume(node)?;
    } else if let Some(query) = target.query() {
        let walk = crate::query::run(Nav::new(file), query)?;
        if !walk.plan().diagnostics().is_empty() {
            return Err(Error::new(
                Code::InvalidSelector,
                "the mutation selector contains an unsupported predicate",
            )
            .with_format(FileFormat::Xlsx));
        }
        let mut matched = false;
        for node in walk {
            matched = true;
            resolved.consume(node?)?;
        }
        if !matched {
            return Err(Error::new(
                Code::ZeroMatches,
                "the mutation selector matched no element",
            )
            .with_format(FileFormat::Xlsx));
        }
    } else {
        return Err(
            Error::new(Code::InvalidArgument, "the mutation target is empty")
                .with_format(FileFormat::Xlsx),
        );
    }
    file.check_source_for_replay()?;
    resolved.targets.sort_unstable_by_key(|target| {
        let start = target.range.start();
        let end = target.range.end();
        (
            target.sheet,
            start.row().get(),
            start.column().get(),
            end.0,
            end.1,
        )
    });
    Ok(resolved)
}

impl ResolvedTargets<'_> {
    fn consume(&mut self, node: Node<'_>) -> Result<()> {
        let family = match node.element() {
            super::spec::Kind::Range => TargetFamily::Range,
            super::spec::Kind::Cell => TargetFamily::Cell,
            super::spec::Kind::Sheet => TargetFamily::Sheet,
            _ => {
                return Err(unsupported_target());
            },
        };
        if let Some(previous) = self.family {
            if previous != family {
                return Err(mixed_target());
            }
        } else {
            self.family = Some(family);
        }

        let selected = match family {
            TargetFamily::Range => {
                let Some((sheet, range)) = node.write_range_target() else {
                    return Err(unsupported_target());
                };
                SelectedNode::Range { sheet, range }
            },
            TargetFamily::Cell => {
                let Some((sheet, address)) = node.write_cell_target() else {
                    return Err(unsupported_target());
                };
                SelectedNode::Cell { sheet, address }
            },
            TargetFamily::Sheet => SelectedNode::Sheet {
                sheet: node.sheet_position(),
            },
        };
        let inserted = self.record_selected(selected)?;
        if inserted && family == TargetFamily::Sheet {
            self.sheet_positions.try_reserve(1).map_err(allocation)?;
            self.sheet_positions.push(node.sheet_position());
        }

        match family {
            TargetFamily::Range => {
                let Some((sheet, range)) = node.write_range_target() else {
                    return Err(unsupported_target());
                };
                self.consume_range(&node, sheet, range)
            },
            TargetFamily::Cell => {
                let Some((sheet, address)) = node.write_cell_target() else {
                    return Err(unsupported_target());
                };
                self.consume_cell(&node, sheet, address)
            },
            TargetFamily::Sheet => self.consume_sheet(&node, node.sheet_position()),
        }
    }

    fn record_selected(&mut self, selected: SelectedNode) -> Result<bool> {
        if self.selected.contains(&selected) {
            return Ok(false);
        }
        let observed = self
            .selected
            .len()
            .checked_add(1)
            .ok_or_else(|| expanded_limit(usize::MAX, self.limits.max_expanded_targets()))?;
        if observed > self.limits.max_expanded_targets() {
            return Err(expanded_limit(observed, self.limits.max_expanded_targets()));
        }
        self.selected.try_reserve(1).map_err(allocation)?;
        self.selected.insert(selected);
        Ok(true)
    }

    fn consume_range(&mut self, node: &Node<'_>, sheet: usize, range: Rect) -> Result<()> {
        match self.mode {
            MergeMode::Enable => self.add_ranges(node, sheet, std::slice::from_ref(&range)),
            MergeMode::Disable => self.remove_exact(node, sheet, range),
            MergeMode::Sweep => self.remove_sweep(node, sheet, range),
            MergeMode::Explicit(ranges) => {
                if ranges.len() != 1 || ranges[0] != range {
                    return Err(invalid_property(
                        "merge",
                        "a range merge value must match its range target",
                    ));
                }
                self.add_ranges(node, sheet, ranges)
            },
        }
    }

    fn consume_cell(
        &mut self,
        node: &Node<'_>,
        sheet: usize,
        address: xlsx::Address,
    ) -> Result<()> {
        match self.mode {
            MergeMode::Enable => Err(unsupported_target()),
            MergeMode::Disable => {
                // A cell clear addresses the merge anchor.  Passing a
                // follower to native `unmerge` would remove a different
                // logical target, so a follower is a safe no-op.
                let range = self.find_containing(node, sheet, address)?;
                if let Some(range) = range.filter(|range| range.start() == address) {
                    self.push_target(MergeTarget { sheet, range })?;
                }
                Ok(())
            },
            MergeMode::Sweep => Err(unsupported_target()),
            MergeMode::Explicit(ranges) => self.add_ranges(node, sheet, ranges),
        }
    }

    fn consume_sheet(&mut self, node: &Node<'_>, sheet: usize) -> Result<()> {
        match self.mode {
            MergeMode::Enable | MergeMode::Disable => Err(unsupported_target()),
            MergeMode::Sweep => Err(unsupported_target()),
            MergeMode::Explicit(ranges) => self.add_ranges(node, sheet, ranges),
        }
    }

    fn add_ranges(&mut self, node: &Node<'_>, sheet: usize, ranges: &[Rect]) -> Result<()> {
        // An explicit list is constant for this call. Cell selectors may
        // match many cells in one sheet, but they all request identical merge
        // geometry there. Selected-node admission still runs for every match.
        let explicit = matches!(self.mode, MergeMode::Explicit(_));
        if explicit && self.expanded_lists.contains(&sheet) {
            return Ok(());
        }
        for range in ranges {
            if range.rows() == 1 && range.columns() == 1 {
                return Err(invalid_property(
                    "merge",
                    "a merged range must contain at least two cells",
                ));
            }
        }

        // Keep the parsed grid handle, but only retain exact matches from its
        // borrowed merge iterator.  A large worksheet therefore never gets
        // copied into a second catalog for every requested range.
        let mut requested = HashSet::new();
        requested.try_reserve(ranges.len()).map_err(allocation)?;
        for range in ranges {
            requested.insert(*range);
        }
        let existing = {
            let grid = self.catalog(node, sheet)?.clone();
            let mut exact = HashSet::new();
            for candidate in grid.merges()? {
                if requested.contains(&candidate) {
                    exact.try_reserve(1).map_err(allocation)?;
                    exact.insert(candidate);
                }
            }
            exact
        };
        for range in ranges {
            if !existing.contains(range) {
                self.push_target(MergeTarget {
                    sheet,
                    range: *range,
                })?;
            }
        }
        if explicit {
            self.expanded_lists.try_reserve(1).map_err(allocation)?;
            self.expanded_lists.insert(sheet);
        }
        Ok(())
    }

    fn remove_exact(&mut self, node: &Node<'_>, sheet: usize, range: Rect) -> Result<()> {
        let (exact, overlap) = {
            let grid = self.catalog(node, sheet)?.clone();
            let mut exact = false;
            let mut overlap = false;
            for candidate in grid.merges()? {
                if candidate == range {
                    exact = true;
                } else if rectangles_overlap(candidate, range) {
                    overlap = true;
                }
            }
            (exact, overlap)
        };
        if exact {
            self.push_target(MergeTarget { sheet, range })?;
        } else if overlap {
            return Err(merge_not_exact());
        }
        Ok(())
    }

    fn remove_sweep(&mut self, node: &Node<'_>, sheet: usize, range: Rect) -> Result<()> {
        self.select_catalog(node, sheet, |candidate| contains_range(range, candidate))
    }

    fn find_containing(
        &mut self,
        node: &Node<'_>,
        sheet: usize,
        address: xlsx::Address,
    ) -> Result<Option<Rect>> {
        let grid = self.catalog(node, sheet)?.clone();
        Ok(grid.merges()?.find(|candidate| candidate.contains(address)))
    }

    fn select_catalog(
        &mut self,
        node: &Node<'_>,
        sheet: usize,
        matches: impl Fn(Rect) -> bool,
    ) -> Result<()> {
        let grid = self.catalog(node, sheet)?.clone();
        for candidate in grid.merges()? {
            if matches(candidate) {
                self.push_target(MergeTarget {
                    sheet,
                    range: candidate,
                })?;
            }
        }
        Ok(())
    }

    fn catalog(&mut self, node: &Node<'_>, sheet: usize) -> Result<&Grid> {
        if !self.grids.contains_key(&sheet) {
            self.grids.try_reserve(1).map_err(allocation)?;
            let grid = load_catalog(node, sheet)?;
            self.grids.insert(sheet, grid);
        }
        Ok(self.grids.get(&sheet).expect("merge grid inserted"))
    }

    fn push_target(&mut self, target: MergeTarget) -> Result<()> {
        if self.seen.contains(&target) {
            return Ok(());
        }
        let observed = self
            .targets
            .len()
            .checked_add(1)
            .ok_or_else(|| expanded_limit(usize::MAX, self.limits.max_expanded_targets()))?;
        if observed > self.limits.max_expanded_targets() {
            return Err(expanded_limit(observed, self.limits.max_expanded_targets()));
        }
        let staged = observed
            .checked_mul(MERGE_ENTRY_OVERHEAD)
            .ok_or_else(|| staged_limit(usize::MAX, self.limits.max_staged_bytes()))?;
        if staged > self.limits.max_staged_bytes() {
            return Err(staged_limit(staged, self.limits.max_staged_bytes()));
        }
        self.targets.try_reserve(1).map_err(allocation)?;
        self.seen.try_reserve(1).map_err(allocation)?;
        self.seen.insert(target);
        self.targets.push(target);
        Ok(())
    }
}

fn load_catalog(node: &Node<'_>, sheet: usize) -> Result<Grid> {
    let handle = SheetHandle::at(node.file(), sheet)?.ok_or_else(|| {
        Error::new(
            Code::OutOfRange,
            "the XLSX worksheet position is out of range",
        )
    })?;
    handle.grid()
}

fn contains_range(outer: Rect, inner: Rect) -> bool {
    let (outer_end_row, outer_end_column) = outer.end();
    let (inner_end_row, inner_end_column) = inner.end();
    outer.start().row().get() <= inner.start().row().get()
        && outer.start().column().get() <= inner.start().column().get()
        && inner_end_row <= outer_end_row
        && inner_end_column <= outer_end_column
}

fn apply_merge_targets(
    edit: &mut xlsx::workbook::Edit,
    targets: &[MergeTarget],
    enabled: bool,
) -> Result<()> {
    let mut index = 0;
    while index < targets.len() {
        let sheet = targets[index].sheet;
        let mut worksheet = edit.sheet(sheet).map_err(native_error)?.ok_or_else(|| {
            Error::new(
                Code::OutOfRange,
                "the XLSX worksheet position is out of range",
            )
            .with_format(FileFormat::Xlsx)
        })?;
        while let Some(target) = targets.get(index).filter(|target| target.sheet == sheet) {
            if enabled {
                worksheet
                    .merge(Area::from(target.range))
                    .map_err(native_error)?;
            } else {
                worksheet
                    .unmerge(target.range.start())
                    .map_err(native_error)?;
            }
            index += 1;
        }
    }
    Ok(())
}

fn apply_sheet_assignments(
    file: &Xlsx,
    edit: &mut xlsx::workbook::Edit,
    positions: &[usize],
    assignments: &SheetAssignments,
) -> Result<()> {
    let names = file.worksheet_names()?;
    for &position in positions {
        let Some(mut tab) = edit.tab(position).map_err(native_error)? else {
            return Err(Error::new(
                Code::OutOfRange,
                "the XLSX worksheet position is out of range",
            )
            .with_format(FileFormat::Xlsx));
        };
        if let Some(name) = assignments.name.as_ref()
            && names
                .get(position)
                .is_none_or(|current| current.as_str() != name.as_str())
        {
            tab.rename(name.clone()).map_err(native_error)?;
        }
        if let Some(hidden) = assignments.hidden {
            if hidden {
                tab.hide();
            } else {
                tab.show();
            }
        }
    }
    Ok(())
}

fn check_combined_limits(
    merge_count: usize,
    cells: Option<&PreparedCells>,
    sheets: Option<&SheetAssignments>,
    sheet_count: usize,
    limits: &SemanticLimits,
) -> Result<()> {
    let sheet_count = if sheets.is_some() { sheet_count } else { 0 };
    let cell_count = cells.map_or(0, PreparedCells::target_count);
    let count = merge_count
        .checked_add(cell_count)
        .and_then(|count| count.checked_add(sheet_count))
        .ok_or_else(|| expanded_limit(usize::MAX, limits.max_expanded_targets()))?;
    if count > limits.max_expanded_targets() {
        return Err(expanded_limit(count, limits.max_expanded_targets()));
    }
    let merge_bytes = merge_count
        .checked_mul(MERGE_ENTRY_OVERHEAD)
        .ok_or_else(|| staged_limit(usize::MAX, limits.max_staged_bytes()))?;
    let cell_bytes = cells.map_or(0, PreparedCells::staged_bytes);
    let sheet_bytes = sheets.map_or(0, |assignments| {
        sheet_count.saturating_mul(sheet_assignment_charge(assignments))
    });
    let staged = merge_bytes
        .checked_add(cell_bytes)
        .and_then(|staged| staged.checked_add(sheet_bytes))
        .ok_or_else(|| staged_limit(usize::MAX, limits.max_staged_bytes()))?;
    if staged > limits.max_staged_bytes() {
        return Err(staged_limit(staged, limits.max_staged_bytes()));
    }
    Ok(())
}

fn sheet_assignment_charge(assignments: &SheetAssignments) -> usize {
    assignments
        .name
        .as_ref()
        .map_or(0, |name| name.as_str().len())
        .saturating_add(usize::from(assignments.hidden.is_some()))
        .saturating_add(MERGE_ENTRY_OVERHEAD)
}

fn rectangles_overlap(left: Rect, right: Rect) -> bool {
    let (left_end_row, left_end_column) = left.end();
    let (right_end_row, right_end_column) = right.end();
    left.start().row().get() < right_end_row
        && right.start().row().get() < left_end_row
        && left.start().column().get() < right_end_column
        && right.start().column().get() < left_end_column
}

fn unsupported_target() -> Error {
    Error::new(
        Code::UnsupportedElement,
        "the XLSX merge property requires a worksheet range, cell, or sheet",
    )
    .with_format(FileFormat::Xlsx)
    .with_property("merge")
}

fn mixed_target() -> Error {
    Error::new(
        Code::UnsupportedElement,
        "an XLSX merge selector cannot mix ranges, cells, and sheets",
    )
    .with_format(FileFormat::Xlsx)
    .with_property("merge")
}

fn merge_not_exact() -> Error {
    Error::new(
        Code::MergeNotExact,
        "the XLSX merge target does not identify one exact merged range",
    )
    .with_format(FileFormat::Xlsx)
    .with_property("merge")
}

fn duplicate_property(key: &'static str) -> Error {
    Error::new(
        Code::InvalidArgument,
        "the same XLSX merge property was supplied more than once",
    )
    .with_property(key)
}

fn invalid_property(key: &'static str, message: &'static str) -> Error {
    Error::new(Code::InvalidValue, message)
        .with_format(FileFormat::Xlsx)
        .with_property(key)
}

fn property_error(error: Error, key: &Key<'_>) -> Error {
    error
        .with_format(FileFormat::Xlsx)
        .with_property(key.as_str().to_owned())
}

fn native_error(error: xlsx::Error) -> Error {
    match error {
        error @ xlsx::Error::TabEditBlocked {
            reason: xlsx::TabEditBlock::LastVisibleTab,
            ..
        } => Error::from(error)
            .with_format(FileFormat::Xlsx)
            .caused_by_argument(Code::InvalidArgument)
            .with_property("hidden"),
        error => super::sheet_write::sheet_error(error),
    }
}

fn expanded_limit(observed: usize, maximum: usize) -> Error {
    Error::from(TransactionError::Limit {
        kind: LimitKind::ExpandedTargets,
        observed,
        maximum,
    })
    .with_format(FileFormat::Xlsx)
}

fn staged_limit(observed: usize, maximum: usize) -> Error {
    Error::from(TransactionError::Limit {
        kind: LimitKind::StagedBytes,
        observed,
        maximum,
    })
    .with_format(FileFormat::Xlsx)
}

fn allocation(source: TryReserveError) -> Error {
    Error::from(xlsx::Error::Allocation {
        resource: "XLSX merged-range write targets",
        source,
    })
    .with_format(FileFormat::Xlsx)
}
