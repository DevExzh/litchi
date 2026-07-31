//! Isolated worksheet transactions, disjoint joins, and source-checked patches.

use std::collections::{BTreeMap, btree_map::Entry};
use std::fmt;
use std::sync::Arc;

use litchi_opc::{OpcPackage, PackURI, Part, Relationship};
use litchi_sheet::{At, Cell as Address, Column as ColumnIndex, ColumnAt, Row as RowIndex, RowAt};

use super::{Sheet, SheetKind, SheetSelector, Visibility, Workbook};
use crate::cell::{Cell, Content, Stored};
use crate::error::{EditBlock, Error, Result, TabEditBlock, invalid};
use crate::raw;
use crate::raw::worksheet::edit::{Action, ColumnAction, Payload, Plan, RowAction, StyleEffect};
use crate::{Style, StyleKey, StyleState};

/// Cell state recorded before or after one semantic change.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum State {
    Missing,
    Cell { content: Cell, style: StyleState },
}

impl State {
    fn read(value: Option<&Stored>, workbook: &Workbook) -> Self {
        value.map_or(Self::Missing, |stored| Self::Cell {
            content: stored.cell.clone(),
            style: stored.style.map_or(StyleState::Default, |key| {
                StyleState::Shared(StyleKey::new(
                    key,
                    Arc::clone(&workbook.inner.style_lineage),
                ))
            }),
        })
    }

    fn after(before: Option<&Stored>, action: &Action, workbook: &Workbook) -> Self {
        let Action::Update { payload, style } = action else {
            return Self::Missing;
        };
        let exists = before.is_some() || action.creates_missing();
        if !exists {
            return Self::Missing;
        }
        let content = match payload {
            Some(Payload::Set(content)) => content.as_cell(),
            Some(Payload::Clear | Payload::ClearIfPresent) => Cell::Empty,
            None => before.map_or(Cell::Empty, |stored| stored.cell.clone()),
        };
        let style = match style {
            Some(StyleEffect::Set(key)) => StyleState::Shared(StyleKey::new(
                *key,
                Arc::clone(&workbook.inner.style_lineage),
            )),
            Some(StyleEffect::Reset) => StyleState::Default,
            None => before
                .and_then(|stored| stored.style)
                .map_or(StyleState::Default, |key| {
                    StyleState::Shared(StyleKey::new(
                        key,
                        Arc::clone(&workbook.inner.style_lineage),
                    ))
                }),
        };
        Self::Cell { content, style }
    }

    fn rebind_style(&mut self, workbook: &Workbook) {
        if let Self::Cell { style, .. } = self {
            style.rebind(&workbook.inner.style_lineage);
        }
    }

    const fn uses_shared_style(&self) -> bool {
        matches!(
            self,
            Self::Cell {
                style: StyleState::Shared(_),
                ..
            }
        )
    }

    fn calculation_content(&self) -> Option<&Cell> {
        match self {
            Self::Cell { content, .. } if !matches!(content, Cell::Empty) => Some(content),
            Self::Missing | Self::Cell { .. } => None,
        }
    }
}

/// Exact row-record state before or after a visibility edit.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum RowState {
    Missing,
    Stored { hidden: bool },
}

/// Exact effective column-record state before or after a visibility edit.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum ColumnState {
    Missing,
    Stored { hidden: bool },
}

/// Semantic active-tab identity recorded in a patch without native Office IDs.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ActiveTab {
    name: Box<str>,
    position: usize,
}

impl ActiveTab {
    /// Developer-facing sheet name at the source snapshot.
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Checked zero-based workbook position at the source snapshot.
    pub const fn position(&self) -> usize {
        self.position
    }
}

impl ColumnState {
    fn read(value: Option<&crate::column::Stored>) -> Self {
        value.map_or(Self::Missing, |column| Self::Stored {
            hidden: column
                .properties
                .flags
                .contains(crate::column::Flags::HIDDEN),
        })
    }

    fn after(before: Option<&crate::column::Stored>, action: ColumnAction) -> Self {
        match (before, action) {
            (None, ColumnAction::Show) => Self::Missing,
            (_, action) => Self::Stored {
                hidden: action.hidden(),
            },
        }
    }
}

impl RowState {
    fn read(value: Option<&crate::row::Stored>) -> Self {
        value.map_or(Self::Missing, |row| Self::Stored { hidden: row.hidden })
    }

    fn after(before: Option<&crate::row::Stored>, action: RowAction) -> Self {
        match (before, action) {
            (None, RowAction::Show) => Self::Missing,
            (_, action) => Self::Stored {
                hidden: action.hidden(),
            },
        }
    }
}

/// One deterministic semantic change in a reversible patch.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum Change {
    Active {
        before: ActiveTab,
        after: ActiveTab,
    },
    Visibility {
        sheet: Box<str>,
        position: usize,
        before: Visibility,
        after: Visibility,
    },
    Cell {
        sheet: Box<str>,
        address: Address,
        before: State,
        after: State,
    },
    Row {
        sheet: Box<str>,
        row: RowIndex,
        before: RowState,
        after: RowState,
    },
    Column {
        sheet: Box<str>,
        column: ColumnIndex,
        before: ColumnState,
        after: ColumnState,
    },
}

impl Change {
    pub fn sheet(&self) -> &str {
        match self {
            Self::Active { after, .. } => after.name(),
            Self::Visibility { sheet, .. }
            | Self::Cell { sheet, .. }
            | Self::Row { sheet, .. }
            | Self::Column { sheet, .. } => sheet,
        }
    }

    /// Active-tab transition when this is a workbook-view change.
    pub fn active(&self) -> Option<(&ActiveTab, &ActiveTab)> {
        match self {
            Self::Active { before, after } => Some((before, after)),
            Self::Visibility { .. }
            | Self::Cell { .. }
            | Self::Row { .. }
            | Self::Column { .. } => None,
        }
    }

    /// Visibility state tuple when this is a workbook tab change.
    pub fn visibility(&self) -> Option<(usize, &Visibility, &Visibility)> {
        match self {
            Self::Visibility {
                position,
                before,
                after,
                ..
            } => Some((*position, before, after)),
            Self::Active { .. } | Self::Cell { .. } | Self::Row { .. } | Self::Column { .. } => {
                None
            },
        }
    }

    /// Cell state tuple when this is an ordinary cell change.
    pub fn cell(&self) -> Option<(Address, &State, &State)> {
        match self {
            Self::Cell {
                address,
                before,
                after,
                ..
            } => Some((*address, before, after)),
            Self::Active { .. }
            | Self::Visibility { .. }
            | Self::Row { .. }
            | Self::Column { .. } => None,
        }
    }

    /// Row state tuple when this is a row-property change.
    pub fn row(&self) -> Option<(RowIndex, RowState, RowState)> {
        match self {
            Self::Row {
                row, before, after, ..
            } => Some((*row, *before, *after)),
            Self::Active { .. }
            | Self::Visibility { .. }
            | Self::Cell { .. }
            | Self::Column { .. } => None,
        }
    }

    /// Column state tuple when this is a column-property change.
    pub fn column(&self) -> Option<(ColumnIndex, ColumnState, ColumnState)> {
        match self {
            Self::Column {
                column,
                before,
                after,
                ..
            } => Some((*column, *before, *after)),
            Self::Active { .. }
            | Self::Visibility { .. }
            | Self::Cell { .. }
            | Self::Row { .. } => None,
        }
    }

    fn inverse(&self) -> Self {
        match self {
            Self::Active { before, after } => Self::Active {
                before: after.clone(),
                after: before.clone(),
            },
            Self::Visibility {
                sheet,
                position,
                before,
                after,
            } => Self::Visibility {
                sheet: sheet.clone(),
                position: *position,
                before: after.clone(),
                after: before.clone(),
            },
            Self::Cell {
                sheet,
                address,
                before,
                after,
            } => Self::Cell {
                sheet: sheet.clone(),
                address: *address,
                before: after.clone(),
                after: before.clone(),
            },
            Self::Row {
                sheet,
                row,
                before,
                after,
            } => Self::Row {
                sheet: sheet.clone(),
                row: *row,
                before: *after,
                after: *before,
            },
            Self::Column {
                sheet,
                column,
                before,
                after,
            } => Self::Column {
                sheet: sheet.clone(),
                column: *column,
                before: *after,
                after: *before,
            },
        }
    }

    fn rebind_style(&mut self, workbook: &Workbook) {
        if let Self::Cell { before, after, .. } = self {
            before.rebind_style(workbook);
            after.rebind_style(workbook);
        }
    }

    fn uses_shared_style(&self) -> bool {
        matches!(
            self,
            Self::Cell { before, after, .. }
                if before.uses_shared_style() || after.uses_shared_style()
        )
    }
}

/// Overlapping effects on one logical workbook sheet.
#[derive(Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum Conflict {
    Active {
        sheet: Box<str>,
        position: usize,
    },
    Tab {
        sheet: Box<str>,
        position: usize,
    },
    Cells {
        sheet: Box<str>,
        position: usize,
        addresses: Box<[Address]>,
    },
    Rows {
        sheet: Box<str>,
        position: usize,
        rows: Box<[RowIndex]>,
    },
    Columns {
        sheet: Box<str>,
        position: usize,
        columns: Box<[ColumnIndex]>,
    },
}

impl Conflict {
    /// Developer-facing sheet name.
    pub fn sheet(&self) -> &str {
        match self {
            Self::Active { sheet, .. }
            | Self::Tab { sheet, .. }
            | Self::Cells { sheet, .. }
            | Self::Rows { sheet, .. }
            | Self::Columns { sheet, .. } => sheet,
        }
    }

    /// Checked zero-based sheet position in the shared base snapshot.
    pub fn position(&self) -> usize {
        match self {
            Self::Active { position, .. }
            | Self::Tab { position, .. }
            | Self::Cells { position, .. }
            | Self::Rows { position, .. }
            | Self::Columns { position, .. } => *position,
        }
    }

    /// Whether both edits target this sheet tab's visibility.
    pub const fn is_tab(&self) -> bool {
        matches!(self, Self::Tab { .. })
    }

    /// Whether both edits target the workbook's one active-tab facet.
    pub const fn is_active(&self) -> bool {
        matches!(self, Self::Active { .. })
    }

    /// Deterministically ordered cells written by both edits, when applicable.
    pub fn cells(&self) -> Option<&[Address]> {
        match self {
            Self::Cells { addresses, .. } => Some(addresses),
            Self::Active { .. } | Self::Tab { .. } | Self::Rows { .. } | Self::Columns { .. } => {
                None
            },
        }
    }

    /// Deterministically ordered rows written by both edits, when applicable.
    pub fn rows(&self) -> Option<&[RowIndex]> {
        match self {
            Self::Rows { rows, .. } => Some(rows),
            Self::Active { .. } | Self::Tab { .. } | Self::Cells { .. } | Self::Columns { .. } => {
                None
            },
        }
    }

    /// Deterministically ordered columns written by both edits, when
    /// applicable.
    pub fn columns(&self) -> Option<&[ColumnIndex]> {
        match self {
            Self::Columns { columns, .. } => Some(columns),
            Self::Active { .. } | Self::Tab { .. } | Self::Cells { .. } | Self::Rows { .. } => None,
        }
    }

    fn len(&self) -> usize {
        match self {
            Self::Active { .. } | Self::Tab { .. } => 1,
            Self::Cells { addresses, .. } => addresses.len(),
            Self::Rows { rows, .. } => rows.len(),
            Self::Columns { columns, .. } => columns.len(),
        }
    }
}

/// Structured overlap report returned by [`Edit::join`].
#[derive(Debug, PartialEq, Eq)]
pub struct ConflictSet {
    conflicts: Box<[Conflict]>,
}

impl ConflictSet {
    /// Conflicts in deterministic workbook-effect and sheet order.
    pub fn conflicts(&self) -> &[Conflict] {
        &self.conflicts
    }

    /// Number of overlapping effects across the workbook.
    pub fn len(&self) -> usize {
        self.conflicts.iter().map(Conflict::len).sum()
    }

    /// Whether no overlapping effects were found.
    pub fn is_empty(&self) -> bool {
        self.conflicts.is_empty()
    }
}

impl fmt::Display for ConflictSet {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(
            formatter,
            "{} overlapping effect(s) across {} conflict group(s)",
            self.len(),
            self.conflicts.len()
        )
    }
}

/// Why two independently prepared edits could not be joined.
#[derive(Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum JoinFailure {
    /// The edits were not created from the same immutable snapshot lineage.
    DifferentSnapshot,
    /// Both edits write at least one of the same effect facets.
    Overlap(ConflictSet),
}

/// Recoverable join failure that returns ownership of the rejected edit.
pub struct JoinError {
    failure: JoinFailure,
    rejected: Edit,
}

impl fmt::Debug for JoinError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("JoinError")
            .field("failure", &self.failure)
            .field("rejected_effects", &self.rejected.len())
            .finish()
    }
}

impl JoinError {
    /// Structured reason the join was refused.
    pub fn failure(&self) -> &JoinFailure {
        &self.failure
    }

    /// Overlapping effects, or `None` for a lineage mismatch.
    pub fn conflicts(&self) -> Option<&ConflictSet> {
        match &self.failure {
            JoinFailure::Overlap(conflicts) => Some(conflicts),
            JoinFailure::DifferentSnapshot => None,
        }
    }

    /// Borrow the edit that was not merged.
    pub fn rejected(&self) -> &Edit {
        &self.rejected
    }

    /// Recover the edit that was not merged.
    pub fn into_rejected(self) -> Edit {
        self.rejected
    }

    /// Recover both the structured reason and rejected edit.
    pub fn into_parts(self) -> (JoinFailure, Edit) {
        (self.failure, self.rejected)
    }
}

impl fmt::Display for JoinError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match &self.failure {
            JoinFailure::DifferentSnapshot => {
                formatter.write_str("edits belong to different workbook snapshots")
            },
            JoinFailure::Overlap(conflicts) => {
                write!(formatter, "edit effects overlap: {conflicts}")
            },
        }
    }
}

impl std::error::Error for JoinError {}

#[derive(Debug, Clone)]
struct PartChange {
    uri: PackURI,
    before: Arc<Vec<u8>>,
    after: Arc<Vec<u8>>,
}

#[derive(Debug, Clone)]
struct StyleGuard {
    uri: PackURI,
    content: Arc<Vec<u8>>,
}

impl StyleGuard {
    fn validate(&self, workbook: &Workbook) -> Result<()> {
        let matches = workbook.inner.styles_uri.as_ref() == Some(&self.uri)
            && workbook
                .inner
                .package
                .get_part(&self.uri)
                .is_ok_and(|part| part.blob() == self.content.as_slice());
        if matches {
            Ok(())
        } else {
            Err(Error::PatchConflict {
                part: self.uri.to_string(),
            })
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum GraphAction {
    Add,
    Remove,
}

#[derive(Clone)]
struct GraphChange {
    action: GraphAction,
    source: PackURI,
    relationship: Relationship,
    part: Box<dyn Part + Send + Sync>,
}

impl std::fmt::Debug for GraphChange {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("GraphChange")
            .field("action", &self.action)
            .field("source", &self.source)
            .field("relationship", &self.relationship.r_id())
            .field("part", self.part.partname())
            .finish()
    }
}

/// Versioned reversible patch produced by a successful transaction.
///
/// The current representation is an in-memory source-checked patch. Its public
/// changes are semantic and contain no native Office IDs; private deltas retain
/// exact changed-part bytes and relationship fields until the deterministic
/// wire format lands.
#[derive(Debug, Clone, Default)]
pub struct Patch {
    changes: Box<[Change]>,
    parts: Box<[PartChange]>,
    graph: Box<[GraphChange]>,
    style_guard: Option<StyleGuard>,
}

impl Patch {
    pub const VERSION: u16 = 1;

    pub const fn version(&self) -> u16 {
        Self::VERSION
    }

    pub fn changes(&self) -> &[Change] {
        &self.changes
    }

    pub fn len(&self) -> usize {
        self.changes.len()
    }

    pub fn is_empty(&self) -> bool {
        self.changes.is_empty()
    }

    /// Build the inverse without copying part payloads.
    pub fn inverse(&self) -> Self {
        Self {
            changes: self
                .changes
                .iter()
                .map(Change::inverse)
                .collect::<Vec<_>>()
                .into_boxed_slice(),
            parts: self
                .parts
                .iter()
                .map(|part| PartChange {
                    uri: part.uri.clone(),
                    before: Arc::clone(&part.after),
                    after: Arc::clone(&part.before),
                })
                .collect::<Vec<_>>()
                .into_boxed_slice(),
            graph: self
                .graph
                .iter()
                .map(|change| GraphChange {
                    action: match change.action {
                        GraphAction::Add => GraphAction::Remove,
                        GraphAction::Remove => GraphAction::Add,
                    },
                    source: change.source.clone(),
                    relationship: change.relationship.clone(),
                    part: change.part.clone(),
                })
                .collect::<Vec<_>>()
                .into_boxed_slice(),
            style_guard: self.style_guard.clone(),
        }
    }

    pub(super) fn apply_to(&self, workbook: &Workbook) -> Result<Commit> {
        ensure_unsigned(workbook)?;
        if let Some(guard) = &self.style_guard {
            guard.validate(workbook)?;
        }
        if self.parts.is_empty() && self.graph.is_empty() {
            return Ok(Commit {
                workbook: workbook.clone(),
                patch: self.clone(),
            });
        }
        let mut package = workbook.inner.package.clone();
        for change in &self.parts {
            let part = package.get_part(&change.uri)?;
            if part.blob() != change.before.as_slice() {
                return Err(Error::PatchConflict {
                    part: change.uri.to_string(),
                });
            }
        }
        for change in &self.graph {
            change.validate(&package)?;
        }
        for change in &self.parts {
            package
                .get_part_mut(&change.uri)?
                .set_blob_shared(Arc::clone(&change.after));
        }
        for change in &self.graph {
            change.apply(&mut package)?;
        }
        let workbook = Workbook::from_package_with_styles(package, Some(workbook))?;
        let mut patch = self.clone();
        for change in &mut patch.changes {
            change.rebind_style(&workbook);
        }
        Ok(Commit { workbook, patch })
    }
}

/// Successful atomic transaction result.
#[derive(Debug)]
pub struct Commit {
    workbook: Workbook,
    patch: Patch,
}

impl Commit {
    pub fn workbook(&self) -> &Workbook {
        &self.workbook
    }

    pub fn patch(&self) -> &Patch {
        &self.patch
    }

    pub fn into_workbook(self) -> Workbook {
        self.workbook
    }

    pub fn into_parts(self) -> (Workbook, Patch) {
        (self.workbook, self.patch)
    }
}

impl GraphChange {
    fn validate(&self, package: &OpcPackage) -> Result<()> {
        let source = package.get_part(&self.source)?;
        match self.action {
            GraphAction::Remove => {
                let relationship =
                    source.rels().get(self.relationship.r_id()).ok_or_else(|| {
                        Error::PatchConflict {
                            part: self.source.to_string(),
                        }
                    })?;
                if !same_relationship(relationship, &self.relationship) {
                    return Err(Error::PatchConflict {
                        part: self.source.to_string(),
                    });
                }
                let part = package.get_part(self.part.partname())?;
                if !same_part(part, &*self.part) {
                    return Err(Error::PatchConflict {
                        part: self.part.partname().to_string(),
                    });
                }
            },
            GraphAction::Add => {
                if source.rels().get(self.relationship.r_id()).is_some()
                    || package
                        .validate_new_part_name(self.part.partname())
                        .is_err()
                {
                    return Err(Error::PatchConflict {
                        part: self.part.partname().to_string(),
                    });
                }
            },
        }
        Ok(())
    }

    fn apply(&self, package: &mut OpcPackage) -> Result<()> {
        match self.action {
            GraphAction::Remove => {
                package
                    .get_part_mut(&self.source)?
                    .rels_mut()
                    .remove(self.relationship.r_id())
                    .ok_or_else(|| Error::PatchConflict {
                        part: self.source.to_string(),
                    })?;
                if !package.remove_part(self.part.partname()) {
                    return Err(Error::PatchConflict {
                        part: self.part.partname().to_string(),
                    });
                }
            },
            GraphAction::Add => {
                package.try_add_part(self.part.clone())?;
                package
                    .get_part_mut(&self.source)?
                    .rels_mut()
                    .try_add_relationship(
                        self.relationship.reltype().to_owned(),
                        self.relationship.target_ref().to_owned(),
                        self.relationship.r_id().to_owned(),
                        self.relationship.target_mode(),
                    )?;
            },
        }
        Ok(())
    }
}

#[derive(Debug, Default)]
struct SheetActions {
    visibility: Option<TabAction>,
    cells: BTreeMap<Address, Action>,
    rows: BTreeMap<RowIndex, RowAction>,
    columns: BTreeMap<ColumnIndex, ColumnAction>,
}

impl SheetActions {
    fn len(&self) -> usize {
        usize::from(self.visibility.is_some())
            .saturating_add(self.cells.len())
            .saturating_add(self.rows.len())
            .saturating_add(self.columns.len())
    }

    fn is_empty(&self) -> bool {
        self.visibility.is_none()
            && self.cells.is_empty()
            && self.rows.is_empty()
            && self.columns.is_empty()
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum TabAction {
    Show,
    Hide,
    VeryHide,
}

impl TabAction {
    const fn visibility(self) -> Visibility {
        match self {
            Self::Show => Visibility::Visible,
            Self::Hide => Visibility::Hidden,
            Self::VeryHide => Visibility::VeryHidden,
        }
    }

    const fn raw(self) -> raw::catalog_edit::State {
        match self {
            Self::Show => raw::catalog_edit::State::Visible,
            Self::Hide => raw::catalog_edit::State::Hidden,
            Self::VeryHide => raw::catalog_edit::State::VeryHidden,
        }
    }
}

/// Isolated workbook transaction. Dropping it rolls back every pending change.
#[derive(Debug)]
pub struct Edit {
    base: Workbook,
    active: Option<usize>,
    sheets: BTreeMap<usize, SheetActions>,
}

impl Edit {
    pub(super) fn new(base: Workbook) -> Result<Self> {
        ensure_unsigned(&base)?;
        Ok(Self {
            base,
            active: None,
            sheets: BTreeMap::new(),
        })
    }

    /// Select a worksheet for short transaction-scoped operations.
    pub fn sheet<'e, 's>(
        &'e mut self,
        selector: impl Into<SheetSelector<'s>>,
    ) -> Result<Option<SheetEdit<'e>>> {
        let sheet = self.base.sheet(selector)?;
        let Some(sheet) = sheet else {
            return Ok(None);
        };
        if sheet.kind() != SheetKind::Worksheet {
            return Err(Error::NotWorksheet {
                sheet: sheet.name().to_owned(),
            });
        }
        let position = sheet.position();
        Ok(Some(SheetEdit {
            edit: self,
            position,
        }))
    }

    /// Select any workbook sheet tab by its developer-facing name or checked
    /// zero-based position. Unlike [`Self::sheet`], this entry point also
    /// accepts chart, dialog, and macro sheets because visibility belongs to
    /// the workbook catalog rather than worksheet cell storage.
    pub fn tab<'e, 's>(
        &'e mut self,
        selector: impl Into<SheetSelector<'s>>,
    ) -> Result<Option<TabEdit<'e>>> {
        let tab = self.base.sheet(selector)?;
        Ok(tab.map(|tab| TabEdit {
            edit: self,
            position: tab.position(),
        }))
    }

    pub fn len(&self) -> usize {
        self.sheets
            .values()
            .fold(usize::from(self.active.is_some()), |len, actions| {
                len.saturating_add(actions.len())
            })
    }

    pub fn is_empty(&self) -> bool {
        self.active.is_none() && self.sheets.values().all(SheetActions::is_empty)
    }

    /// Join an independently prepared edit when every effect is disjoint.
    ///
    /// Both edits must originate from the same immutable snapshot. On failure
    /// `self` is unchanged and [`JoinError`] returns ownership of `other`, so
    /// callers never lose prepared work while resolving a conflict.
    pub fn join(&mut self, other: Self) -> std::result::Result<&mut Self, JoinError> {
        if !Arc::ptr_eq(&self.base.inner, &other.base.inner) {
            return Err(JoinError {
                failure: JoinFailure::DifferentSnapshot,
                rejected: other,
            });
        }
        let conflicts = self.conflicts_with(&other);
        if !conflicts.is_empty() {
            return Err(JoinError {
                failure: JoinFailure::Overlap(conflicts),
                rejected: other,
            });
        }

        if self.active.is_none() {
            self.active = other.active;
        }
        for (position, actions) in other.sheets {
            match self.sheets.entry(position) {
                Entry::Vacant(entry) => {
                    entry.insert(actions);
                },
                Entry::Occupied(entry) => {
                    let accepted = entry.into_mut();
                    if accepted.visibility.is_none() {
                        accepted.visibility = actions.visibility;
                    }
                    for (address, action) in actions.cells {
                        match accepted.cells.entry(address) {
                            Entry::Vacant(entry) => {
                                entry.insert(action);
                            },
                            Entry::Occupied(mut entry) => {
                                entry.get_mut().merge(action);
                            },
                        }
                    }
                    accepted.rows.extend(actions.rows);
                    accepted.columns.extend(actions.columns);
                },
            }
        }
        Ok(self)
    }

    /// Validate and atomically publish a new immutable snapshot.
    pub fn commit(self) -> Result<Commit> {
        ensure_unsigned(&self.base)?;
        if self.is_empty() {
            return Ok(Commit {
                workbook: self.base,
                patch: Patch::default(),
            });
        }
        let Self {
            base,
            active: requested_active,
            sheets,
        } = self;
        let mut changes = Vec::new();
        let mut parts = Vec::new();
        let mut needs_recalculation = false;

        let mut effective_tabs = Vec::new();
        for (position, requested) in &sheets {
            let Some(action) = requested.visibility else {
                continue;
            };
            let data =
                base.inner.sheets.get(*position).ok_or_else(|| {
                    invalid(format!("edited sheet position {position} disappeared"))
                })?;
            let after = action.visibility();
            if data.visibility == after {
                continue;
            }
            effective_tabs.push((*position, action));
        }

        let final_is_visible = |position: usize| {
            sheets
                .get(&position)
                .and_then(|requested| requested.visibility)
                .map_or_else(
                    || {
                        base.inner
                            .sheets
                            .get(position)
                            .is_some_and(|sheet| sheet.visibility == Visibility::Visible)
                    },
                    |action| action == TabAction::Show,
                )
        };
        if !effective_tabs.is_empty() && !(0..base.inner.sheets.len()).any(&final_is_visible) {
            let (position, _) = effective_tabs
                .iter()
                .find(|(_, action)| *action != TabAction::Show)
                .copied()
                .ok_or_else(|| invalid("tab visibility invariant failed without a hide action"))?;
            let data = base
                .inner
                .sheets
                .get(position)
                .ok_or_else(|| invalid("last visible tab disappeared during edit"))?;
            return Err(Error::TabEditBlocked {
                sheet: data.name.clone(),
                position,
                reason: TabEditBlock::LastVisibleTab,
            });
        }

        let current_active = base.inner.active_sheet;
        let final_active = if let Some(position) = requested_active {
            let data = base
                .inner
                .sheets
                .get(position)
                .ok_or_else(|| invalid("requested active tab disappeared during edit"))?;
            if position > raw::catalog_edit::MAX_ACTIVE_TAB {
                return Err(Error::TabEditBlocked {
                    sheet: data.name.clone(),
                    position,
                    reason: TabEditBlock::ActiveTabLimit,
                });
            }
            if !final_is_visible(position) {
                return Err(Error::TabEditBlocked {
                    sheet: data.name.clone(),
                    position,
                    reason: TabEditBlock::NotVisible,
                });
            }
            Some(position)
        } else if effective_tabs.is_empty() || current_active.is_some_and(final_is_visible) {
            current_active
        } else {
            let len = base.inner.sheets.len();
            if len == 0 {
                None
            } else {
                current_active
                    .filter(|current| *current < len)
                    .and_then(|current| {
                        let remaining = len - current;
                        (1..=len)
                            .map(|offset| {
                                if offset >= remaining {
                                    offset - remaining
                                } else {
                                    current + offset
                                }
                            })
                            .find(|position| final_is_visible(*position))
                    })
                    .or_else(|| (0..len).find(|position| final_is_visible(*position)))
            }
        };
        let active_change = (final_active != current_active)
            .then_some(final_active)
            .flatten();

        if let Some(position) = active_change {
            if position > raw::catalog_edit::MAX_ACTIVE_TAB {
                let data = base
                    .inner
                    .sheets
                    .get(position)
                    .ok_or_else(|| invalid("replacement active tab disappeared during edit"))?;
                return Err(Error::TabEditBlocked {
                    sheet: data.name.clone(),
                    position,
                    reason: TabEditBlock::ActiveTabLimit,
                });
            }
            let before = current_active
                .map(|position| active_tab(&base, position))
                .transpose()?
                .ok_or_else(|| invalid("non-empty workbook has no source active tab"))?;
            let after = active_tab(&base, position)?;
            changes.push(Change::Active { before, after });
        }

        for (position, action) in &effective_tabs {
            let data = base
                .inner
                .sheets
                .get(*position)
                .ok_or_else(|| invalid("effective tab disappeared during edit"))?;
            changes.push(Change::Visibility {
                sheet: data.name.clone().into_boxed_str(),
                position: *position,
                before: data.visibility.clone(),
                after: action.visibility(),
            });
        }

        for (position, requested) in sheets {
            let data =
                base.inner.sheets.get(position).cloned().ok_or_else(|| {
                    invalid(format!("edited sheet position {position} disappeared"))
                })?;
            let SheetActions {
                visibility: _,
                cells,
                rows,
                columns,
            } = requested;
            if cells.is_empty() && rows.is_empty() && columns.is_empty() {
                continue;
            }
            if data.kind != SheetKind::Worksheet {
                return Err(Error::NotWorksheet {
                    sheet: data.name.clone(),
                });
            }
            let sheet = Sheet {
                owner: Arc::clone(&base.inner),
                data: Arc::clone(&data),
            };
            let store = sheet.store()?;
            let change_start = changes.len();
            let mut effective_cells = BTreeMap::new();
            for (address, action) in cells {
                let before = store.entry(address);
                if before.is_some_and(|stored| matches!(stored.cell, Cell::Unknown(_)))
                    && action.payload().is_some()
                {
                    return Err(Error::EditBlocked {
                        sheet: data.name.clone(),
                        address,
                        reason: EditBlock::UnknownCell,
                    });
                }
                let before_state = State::read(before, &base);
                let after_state = State::after(before, &action, &base);
                if before_state == after_state {
                    continue;
                }
                needs_recalculation |= State::calculation_content(&before_state)
                    != State::calculation_content(&after_state);
                effective_cells.insert(address, action);
                changes.push(Change::Cell {
                    sheet: data.name.clone().into_boxed_str(),
                    address,
                    before: before_state,
                    after: after_state,
                });
            }
            let mut effective_rows = BTreeMap::new();
            for (row, action) in rows {
                let before = store.row_entry(row);
                let before_state = RowState::read(before);
                let after_state = RowState::after(before, action);
                if before_state == after_state {
                    continue;
                }
                effective_rows.insert(row, action);
                changes.push(Change::Row {
                    sheet: data.name.clone().into_boxed_str(),
                    row,
                    before: before_state,
                    after: after_state,
                });
            }
            let mut effective_columns = BTreeMap::new();
            for (column, action) in columns {
                let before = store.column_entry(column);
                let before_state = ColumnState::read(before);
                let after_state = ColumnState::after(before, action);
                if before_state == after_state {
                    continue;
                }
                effective_columns.insert(column, action);
                changes.push(Change::Column {
                    sheet: data.name.clone().into_boxed_str(),
                    column,
                    before: before_state,
                    after: after_state,
                });
            }
            if effective_cells.is_empty()
                && effective_rows.is_empty()
                && effective_columns.is_empty()
            {
                continue;
            }

            let part = base.inner.package.get_part(&data.part_uri)?;
            let before = part.blob_arc();
            let after = raw::worksheet::edit::rewrite(
                &before,
                &data.name,
                Plan {
                    cells: effective_cells,
                    rows: effective_rows,
                    columns: effective_columns,
                },
            )?;
            let parsed = raw::worksheet::parse(&after, || base.inner.shared_strings())?;
            base.inner.validate_styles(&parsed)?;
            for change in &changes[change_start..] {
                match change {
                    Change::Active { .. } | Change::Visibility { .. } => {},
                    Change::Cell {
                        sheet,
                        address,
                        after,
                        ..
                    } => {
                        let actual = State::read(parsed.entry(*address), &base);
                        if actual != *after {
                            return Err(invalid(format!(
                                "worksheet edit verification failed at {sheet}!{address}"
                            )));
                        }
                    },
                    Change::Row {
                        sheet, row, after, ..
                    } => {
                        let actual = RowState::read(parsed.row_entry(*row));
                        if actual != *after {
                            return Err(invalid(format!(
                                "worksheet row edit verification failed at {sheet}!row {}",
                                row.get()
                            )));
                        }
                    },
                    Change::Column {
                        sheet,
                        column,
                        after,
                        ..
                    } => {
                        let actual = ColumnState::read(parsed.column_entry(*column));
                        if actual != *after {
                            return Err(invalid(format!(
                                "worksheet column edit verification failed at {sheet}!column {}",
                                column.get()
                            )));
                        }
                    },
                }
            }
            parts.push(PartChange {
                uri: data.part_uri.clone(),
                before,
                after: Arc::new(after),
            });
        }

        if let Some(new_active) = active_change {
            if let Some(old_active) = current_active {
                let data =
                    base.inner.sheets.get(old_active).ok_or_else(|| {
                        invalid("previous active sheet disappeared during tab edit")
                    })?;
                if data.kind == SheetKind::Unknown {
                    return Err(Error::TabEditBlocked {
                        sheet: data.name.clone(),
                        position: old_active,
                        reason: TabEditBlock::MarkupCompatibility,
                    });
                }
                compose_part(&mut parts, &base, &data.part_uri, |content| {
                    raw::sheet_view_edit::rewrite(
                        content,
                        false,
                        raw::sheet_view_edit::Context {
                            sheet: &data.name,
                            position: old_active,
                        },
                    )
                })?;
            }
            let data = base
                .inner
                .sheets
                .get(new_active)
                .ok_or_else(|| invalid("new active sheet disappeared during tab edit"))?;
            if data.kind == SheetKind::Unknown {
                return Err(Error::TabEditBlocked {
                    sheet: data.name.clone(),
                    position: new_active,
                    reason: TabEditBlock::MarkupCompatibility,
                });
            }
            compose_part(&mut parts, &base, &data.part_uri, |content| {
                raw::sheet_view_edit::rewrite(
                    content,
                    true,
                    raw::sheet_view_edit::Context {
                        sheet: &data.name,
                        position: new_active,
                    },
                )
            })?;
        }

        let style_guard = changes
            .iter()
            .any(Change::uses_shared_style)
            .then(|| {
                let uri = base
                    .inner
                    .styles_uri
                    .as_ref()
                    .ok_or_else(|| invalid("shared style state has no styles part"))?
                    .clone();
                let content = base.inner.package.get_part(&uri)?.blob_arc();
                Ok::<_, Error>(StyleGuard { uri, content })
            })
            .transpose()?;
        let graph = if needs_recalculation {
            calculation_chain_removal(&base)?
        } else {
            Vec::new()
        };

        if !effective_tabs.is_empty() || active_change.is_some() || needs_recalculation {
            let workbook_part = base.inner.package.get_part(&base.inner.workbook_uri)?;
            let before = workbook_part.blob_arc();
            let active = active_change
                .map(|position| {
                    let data = base.inner.sheets.get(position).ok_or_else(|| {
                        invalid("active-tab rewrite target disappeared during commit")
                    })?;
                    Ok::<_, Error>(raw::catalog_edit::Active {
                        sheet: &data.name,
                        position,
                    })
                })
                .transpose()?;
            let mut after = if effective_tabs.is_empty() && active.is_none() {
                raw::recalc::invalidate(&before)?
            } else {
                let tabs = effective_tabs
                    .iter()
                    .map(|(position, action)| {
                        let data = base.inner.sheets.get(*position).ok_or_else(|| {
                            invalid("tab rewrite target disappeared during commit")
                        })?;
                        Ok(raw::catalog_edit::Tab {
                            sheet: &data.name,
                            position: *position,
                            relationship_id: &data.relationship_id,
                            state: action.raw(),
                        })
                    })
                    .collect::<Result<Vec<_>>>()?;
                raw::catalog_edit::rewrite(&before, raw::catalog_edit::Plan { tabs, active })?
            };
            if needs_recalculation && (!effective_tabs.is_empty() || active.is_some()) {
                after = raw::recalc::invalidate(&after)?;
            }
            if !effective_tabs.is_empty() || active.is_some() {
                let catalog = raw::parse_catalog(&after)?;
                if Some(catalog.active_sheet_index) != final_active {
                    return Err(invalid("workbook active-tab edit verification failed"));
                }
                for (position, action) in &effective_tabs {
                    let actual = catalog
                        .sheets
                        .get(*position)
                        .ok_or_else(|| invalid("workbook tab edit verification lost a sheet"))?;
                    if !raw_visibility_matches(&actual.visibility, *action) {
                        let sheet = base
                            .inner
                            .sheets
                            .get(*position)
                            .map_or("<missing sheet>", |sheet| sheet.name.as_str());
                        return Err(invalid(format!(
                            "workbook tab visibility verification failed at {sheet}"
                        )));
                    }
                }
            }
            if after.as_slice() != before.as_slice() {
                parts.push(PartChange {
                    uri: base.inner.workbook_uri.clone(),
                    before,
                    after: Arc::new(after),
                });
            }
        }

        if changes.is_empty() && parts.is_empty() && graph.is_empty() {
            return Ok(Commit {
                workbook: base,
                patch: Patch::default(),
            });
        }
        let mut package = base.inner.package.clone();
        for part in &parts {
            package
                .get_part_mut(&part.uri)?
                .set_blob_shared(Arc::clone(&part.after));
        }
        for change in &graph {
            change.validate(&package)?;
            change.apply(&mut package)?;
        }
        let workbook = Workbook::from_package_with_styles(package, Some(&base))?;
        Ok(Commit {
            workbook,
            patch: Patch {
                changes: changes.into_boxed_slice(),
                parts: parts.into_boxed_slice(),
                graph: graph.into_boxed_slice(),
                style_guard,
            },
        })
    }

    fn actions(&mut self, position: usize) -> &mut BTreeMap<Address, Action> {
        &mut self.sheets.entry(position).or_default().cells
    }

    fn set_visibility(&mut self, position: usize, action: TabAction) {
        self.sheets.entry(position).or_default().visibility = Some(action);
    }

    fn set_active(&mut self, position: usize) {
        self.active = Some(position);
    }

    fn row_actions(&mut self, position: usize) -> &mut BTreeMap<RowIndex, RowAction> {
        &mut self.sheets.entry(position).or_default().rows
    }

    fn column_actions(&mut self, position: usize) -> &mut BTreeMap<ColumnIndex, ColumnAction> {
        &mut self.sheets.entry(position).or_default().columns
    }

    fn conflicts_with(&self, other: &Self) -> ConflictSet {
        let mut conflicts = Vec::new();
        if let (Some(position), Some(_)) = (self.active, other.active) {
            let sheet = self
                .base
                .inner
                .sheets
                .get(position)
                .map_or("<missing sheet>", |sheet| sheet.name.as_str());
            conflicts.push(Conflict::Active {
                sheet: sheet.into(),
                position,
            });
        }
        for (position, left) in &self.sheets {
            let Some(right) = other.sheets.get(position) else {
                continue;
            };
            let sheet = self
                .base
                .inner
                .sheets
                .get(*position)
                .map_or("<missing sheet>", |sheet| sheet.name.as_str());
            if left.visibility.is_some() && right.visibility.is_some() {
                conflicts.push(Conflict::Tab {
                    sheet: sheet.into(),
                    position: *position,
                });
            }
            let mut addresses = Vec::new();
            let mut left_cells = left.cells.iter().peekable();
            let mut right_cells = right.cells.iter().peekable();
            while let (Some((left_address, left_action)), Some((right_address, right_action))) =
                (left_cells.peek(), right_cells.peek())
            {
                match left_address.cmp(right_address) {
                    std::cmp::Ordering::Less => {
                        left_cells.next();
                    },
                    std::cmp::Ordering::Greater => {
                        right_cells.next();
                    },
                    std::cmp::Ordering::Equal => {
                        if left_action.overlaps(right_action) {
                            addresses.push(**left_address);
                        }
                        left_cells.next();
                        right_cells.next();
                    },
                }
            }
            if !addresses.is_empty() {
                conflicts.push(Conflict::Cells {
                    sheet: sheet.into(),
                    position: *position,
                    addresses: addresses.into_boxed_slice(),
                });
            }

            let rows = left
                .rows
                .keys()
                .filter(|row| right.rows.contains_key(row))
                .copied()
                .collect::<Vec<_>>();
            if !rows.is_empty() {
                conflicts.push(Conflict::Rows {
                    sheet: sheet.into(),
                    position: *position,
                    rows: rows.into_boxed_slice(),
                });
            }

            let columns = left
                .columns
                .keys()
                .filter(|column| right.columns.contains_key(column))
                .copied()
                .collect::<Vec<_>>();
            if !columns.is_empty() {
                conflicts.push(Conflict::Columns {
                    sheet: sheet.into(),
                    position: *position,
                    columns: columns.into_boxed_slice(),
                });
            }
        }
        ConflictSet {
            conflicts: conflicts.into_boxed_slice(),
        }
    }
}

/// Transaction-scoped state editor for any workbook sheet tab.
#[derive(Debug)]
pub struct TabEdit<'a> {
    edit: &'a mut Edit,
    position: usize,
}

impl TabEdit<'_> {
    /// Make this the active tab while preserving unrelated grouped selections.
    /// The tab must be visible in the transaction's final state.
    pub fn activate(&mut self) -> &mut Self {
        self.edit.set_active(self.position);
        self
    }

    /// Make this tab visible.
    pub fn show(&mut self) -> &mut Self {
        self.edit.set_visibility(self.position, TabAction::Show);
        self
    }

    /// Hide this tab while retaining it in Excel's ordinary Unhide dialog.
    pub fn hide(&mut self) -> &mut Self {
        self.edit.set_visibility(self.position, TabAction::Hide);
        self
    }

    /// Hide this tab from Excel's ordinary Unhide dialog (`veryHidden`).
    pub fn very_hide(&mut self) -> &mut Self {
        self.edit.set_visibility(self.position, TabAction::VeryHide);
        self
    }
}

/// Borrowed worksheet editor tied to one transaction.
#[derive(Debug)]
pub struct SheetEdit<'a> {
    edit: &'a mut Edit,
    position: usize,
}

impl SheetEdit<'_> {
    /// Select one checked row for short property-editing verbs.
    pub fn row(&mut self, at: impl Into<RowAt>) -> Result<RowEdit<'_>> {
        let row = at.into().resolve()?;
        Ok(RowEdit {
            edit: &mut *self.edit,
            position: self.position,
            row,
        })
    }

    /// Select one checked column for short property-editing verbs.
    pub fn column(&mut self, at: impl Into<ColumnAt>) -> Result<ColumnEdit<'_>> {
        let column = at.into().resolve()?;
        Ok(ColumnEdit {
            edit: &mut *self.edit,
            position: self.position,
            column,
        })
    }

    /// Create or replace a cell's primary value or formula.
    pub fn set<'a>(
        &mut self,
        at: impl Into<At<'a>>,
        content: impl Into<Content>,
    ) -> Result<&mut Self> {
        let address = at.into().resolve()?;
        let content = content.into();
        content.validate_for_write()?;
        match self.edit.actions(self.position).entry(address) {
            Entry::Vacant(entry) => {
                entry.insert(Action::set(content));
            },
            Entry::Occupied(mut entry) => entry.get_mut().set_payload(Payload::Set(content)),
        }
        Ok(self)
    }

    /// Remove primary value/formula content while retaining the cell record,
    /// local style, metadata, comments, and unknown non-payload children.
    pub fn clear<'a>(&mut self, at: impl Into<At<'a>>) -> Result<&mut Self> {
        let address = at.into().resolve()?;
        let actions = self.edit.actions(self.position);
        let create = actions.get(&address).is_some_and(|action| {
            matches!(action.payload(), Some(Payload::Set(_) | Payload::Clear))
        });
        match actions.entry(address) {
            Entry::Vacant(entry) => {
                entry.insert(Action::clear(create));
            },
            Entry::Occupied(mut entry) => entry.get_mut().set_payload(if create {
                Payload::Clear
            } else {
                Payload::ClearIfPresent
            }),
        }
        Ok(self)
    }

    /// Apply an existing shared style without copying its formatting payload.
    ///
    /// The handle must belong to this edit's shared-style table lineage.
    /// Styling a missing cell creates an explicit empty cell record.
    pub fn style<'a>(&mut self, at: impl Into<At<'a>>, style: &Style) -> Result<&mut Self> {
        if !Arc::ptr_eq(
            &self.edit.base.inner.style_lineage,
            &style.owner.style_lineage,
        ) {
            return Err(Error::ForeignStyle);
        }
        let address = at.into().resolve()?;
        let effect = StyleEffect::Set(style.raw());
        match self.edit.actions(self.position).entry(address) {
            Entry::Vacant(entry) => {
                entry.insert(Action::style(style.raw()));
            },
            Entry::Occupied(mut entry) => entry.get_mut().set_style(effect),
        }
        Ok(self)
    }

    /// Remove an explicit local style reference, retaining the cell payload.
    ///
    /// A missing cell remains missing and resolves as a no-op at commit.
    pub fn reset_style<'a>(&mut self, at: impl Into<At<'a>>) -> Result<&mut Self> {
        let address = at.into().resolve()?;
        match self.edit.actions(self.position).entry(address) {
            Entry::Vacant(entry) => {
                entry.insert(Action::reset_style());
            },
            Entry::Occupied(mut entry) => entry.get_mut().set_style(StyleEffect::Reset),
        }
        Ok(self)
    }

    /// Remove the cell record without shifting surrounding cells.
    pub fn remove<'a>(&mut self, at: impl Into<At<'a>>) -> Result<&mut Self> {
        let address = at.into().resolve()?;
        self.edit
            .actions(self.position)
            .insert(address, Action::Remove);
        Ok(self)
    }
}

/// Transaction-scoped editor for one checked worksheet row.
#[derive(Debug)]
pub struct RowEdit<'a> {
    edit: &'a mut Edit,
    position: usize,
    row: RowIndex,
}

impl RowEdit<'_> {
    /// Hide this row while preserving all row and cell content.
    pub fn hide(&mut self) -> &mut Self {
        self.edit
            .row_actions(self.position)
            .insert(self.row, RowAction::Hide);
        self
    }

    /// Show this row while preserving all row and cell content.
    pub fn show(&mut self) -> &mut Self {
        self.edit
            .row_actions(self.position)
            .insert(self.row, RowAction::Show);
        self
    }
}

/// Transaction-scoped editor for one checked worksheet column.
#[derive(Debug)]
pub struct ColumnEdit<'a> {
    edit: &'a mut Edit,
    position: usize,
    column: ColumnIndex,
}

impl ColumnEdit<'_> {
    /// Hide this column while preserving its other effective properties and
    /// every cell record.
    pub fn hide(&mut self) -> &mut Self {
        self.edit
            .column_actions(self.position)
            .insert(self.column, ColumnAction::Hide);
        self
    }

    /// Show this column while preserving its other effective properties and
    /// every cell record.
    pub fn show(&mut self) -> &mut Self {
        self.edit
            .column_actions(self.position)
            .insert(self.column, ColumnAction::Show);
        self
    }
}

fn raw_visibility_matches(value: &raw::Visibility, action: TabAction) -> bool {
    matches!(
        (value, action),
        (raw::Visibility::Visible, TabAction::Show)
            | (raw::Visibility::Hidden, TabAction::Hide)
            | (raw::Visibility::VeryHidden, TabAction::VeryHide)
    )
}

fn active_tab(workbook: &Workbook, position: usize) -> Result<ActiveTab> {
    let sheet = workbook
        .inner
        .sheets
        .get(position)
        .ok_or_else(|| invalid("active tab points outside the workbook catalog"))?;
    Ok(ActiveTab {
        name: sheet.name.clone().into_boxed_str(),
        position,
    })
}

fn compose_part(
    parts: &mut Vec<PartChange>,
    workbook: &Workbook,
    uri: &PackURI,
    rewrite: impl FnOnce(&[u8]) -> Result<Vec<u8>>,
) -> Result<()> {
    if let Some(part) = parts.iter_mut().find(|part| &part.uri == uri) {
        let after = rewrite(&part.after)?;
        if after.as_slice() != part.after.as_slice() {
            part.after = Arc::new(after);
        }
        return Ok(());
    }
    let before = workbook.inner.package.get_part(uri)?.blob_arc();
    let after = rewrite(&before)?;
    if after.as_slice() != before.as_slice() {
        parts.push(PartChange {
            uri: uri.clone(),
            before,
            after: Arc::new(after),
        });
    }
    Ok(())
}

fn ensure_unsigned(workbook: &Workbook) -> Result<()> {
    if workbook.inner.package.has_digital_signatures() {
        Err(Error::Signed)
    } else {
        Ok(())
    }
}

fn calculation_chain_removal(workbook: &Workbook) -> Result<Vec<GraphChange>> {
    let main = workbook
        .inner
        .package
        .get_part(&workbook.inner.workbook_uri)?;
    let mut matching = main.rels().iter().filter(|relationship| {
        matches!(
            relationship.reltype(),
            litchi_opc::constants::relationship_type::CALC_CHAIN
                | litchi_opc::constants::relationship_type::STRICT_CALC_CHAIN
        )
    });
    let Some(relationship) = matching.next() else {
        return Ok(Vec::new());
    };
    if matching.next().is_some() {
        return Err(invalid(
            "workbook has multiple calculation-chain relationships",
        ));
    }
    if relationship.is_external() {
        return Err(invalid("calculation-chain relationship cannot be external"));
    }
    let target = relationship.target_partname()?;
    let part = workbook.inner.package.get_part(&target)?;
    ensure_exclusive_incoming_relationship(
        &workbook.inner.package,
        part.partname(),
        &workbook.inner.workbook_uri,
        relationship.r_id(),
    )?;
    Ok(vec![GraphChange {
        action: GraphAction::Remove,
        source: workbook.inner.workbook_uri.clone(),
        relationship: relationship.clone(),
        part: part.clone_part(),
    }])
}

fn ensure_exclusive_incoming_relationship(
    package: &OpcPackage,
    target: &PackURI,
    expected_source: &PackURI,
    expected_id: &str,
) -> Result<()> {
    let targets = |relationship: &Relationship| -> Result<bool> {
        if relationship.is_external() {
            return Ok(false);
        }
        relationship
            .target_partname()
            .map(|candidate| candidate.as_str().eq_ignore_ascii_case(target.as_str()))
            .map_err(Into::into)
    };
    for relationship in package.rels().iter() {
        if targets(relationship)? {
            return Err(invalid(format!(
                "calculation-chain part '{target}' has another incoming relationship"
            )));
        }
    }
    for source in package.iter_parts() {
        for relationship in source.rels().iter() {
            if targets(relationship)?
                && !(source.partname() == expected_source && relationship.r_id() == expected_id)
            {
                return Err(invalid(format!(
                    "calculation-chain part '{target}' has another incoming relationship"
                )));
            }
        }
    }
    Ok(())
}

fn same_relationship(left: &Relationship, right: &Relationship) -> bool {
    left.r_id() == right.r_id()
        && left.reltype() == right.reltype()
        && left.target_ref() == right.target_ref()
        && left.target_mode() == right.target_mode()
}

fn same_part(left: &dyn Part, right: &dyn Part) -> bool {
    left.partname() == right.partname()
        && left.content_type() == right.content_type()
        && left.blob() == right.blob()
        && left.rels().len() == right.rels().len()
        && left.rels().iter().all(|relationship| {
            right
                .rels()
                .get(relationship.r_id())
                .is_some_and(|other| same_relationship(relationship, other))
        })
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::cell::{Number, Value};
    use crate::formula::Formula;
    use litchi_opc::{BlobPart, TargetMode};

    #[test]
    fn cell_crud_is_atomic_reversible_and_source_preserving() {
        let source = Workbook::new().expect("source workbook");
        let source_bytes = source.to_bytes().expect("source bytes");
        let mut edit = source.edit().expect("edit");
        {
            let mut sheet = edit.sheet("Sheet1").expect("sheet lookup").expect("sheet");
            sheet
                .set("A1", "hello")
                .and_then(|sheet| sheet.set("B2", 42_i32))
                .and_then(|sheet| sheet.set("C3", Formula::new("B2*2").expect("formula")))
                .expect("cell changes");
        }
        let committed = edit.commit().expect("commit");
        assert_eq!(
            source.to_bytes().expect("source remains valid"),
            source_bytes
        );
        assert_eq!(committed.patch().len(), 3);

        let book = committed.workbook();
        let sheet = book.sheet("Sheet1").expect("lookup").expect("sheet");
        assert!(matches!(
            sheet.cell("A1").expect("A1"),
            Some(Cell::Value(Value::Text(text))) if text.as_str() == "hello"
        ));
        assert!(matches!(
            sheet.cell("B2").expect("B2"),
            Some(Cell::Value(Value::Number(number))) if number == &Number::new("42").expect("number")
        ));
        assert!(matches!(
            sheet.cell("C3").expect("C3"),
            Some(Cell::Formula(_))
        ));
        let extents = sheet.extents().expect("committed extents");
        assert_eq!(
            extents.declared().map(crate::Rect::a1).as_deref(),
            Some("A1:C3")
        );
        assert_eq!(
            extents.content().map(crate::Rect::a1).as_deref(),
            Some("A1:C3")
        );

        let restored = book
            .apply(&committed.patch().inverse())
            .expect("inverse patch");
        assert_eq!(
            restored.workbook().to_bytes().expect("restored"),
            source_bytes
        );
        assert!(matches!(
            source.apply(committed.patch()),
            Ok(applied) if applied.workbook().sheet("Sheet1").expect("lookup").expect("sheet").cell("A1").expect("cell").is_some()
        ));
        assert!(matches!(
            book.apply(committed.patch()),
            Err(Error::PatchConflict { .. })
        ));
    }

    #[test]
    fn clear_and_remove_have_distinct_missing_and_empty_semantics() {
        let source = Workbook::new().expect("source workbook");
        let mut edit = source.edit().expect("edit");
        edit.sheet(0usize)
            .expect("lookup")
            .expect("sheet")
            .set("A1", "value")
            .expect("set");
        let first = edit.commit().expect("first commit").into_workbook();

        let mut edit = first.edit().expect("edit");
        let mut sheet = edit.sheet(0usize).expect("lookup").expect("sheet");
        sheet.clear("A1").expect("clear");
        sheet.clear("B1").expect("clear missing");
        let cleared = edit.commit().expect("clear commit");
        assert_eq!(cleared.patch().len(), 1);
        let sheet = cleared
            .workbook()
            .sheet(0usize)
            .expect("lookup")
            .expect("sheet");
        assert!(matches!(sheet.cell("A1").expect("cell"), Some(Cell::Empty)));
        assert!(sheet.cell("B1").expect("cell").is_none());

        let mut edit = cleared.workbook().edit().expect("edit");
        edit.sheet(0usize)
            .expect("lookup")
            .expect("sheet")
            .remove("A1")
            .expect("remove");
        let removed = edit.commit().expect("remove commit");
        assert!(
            removed
                .workbook()
                .sheet(0usize)
                .expect("lookup")
                .expect("sheet")
                .cell("A1")
                .expect("cell")
                .is_none()
        );
    }

    #[test]
    fn row_visibility_is_checked_reversible_and_patch_visible() {
        let source = Workbook::new().expect("source workbook");
        let source_bytes = source.to_bytes().expect("source bytes");
        let mut edit = source.edit().expect("edit");
        let mut sheet = edit.sheet("Sheet1").expect("lookup").expect("sheet");
        sheet.set("A1", "visible").expect("cell");
        sheet.row(1).expect("row 2").hide();
        let committed = edit.commit().expect("commit");

        assert_eq!(committed.patch().len(), 2);
        assert!(matches!(
            committed.patch().changes()[1],
            Change::Row {
                row,
                before: RowState::Missing,
                after: RowState::Stored { hidden: true },
                ..
            } if row.get() == 1
        ));
        let sheet = committed
            .workbook()
            .sheet("Sheet1")
            .expect("lookup")
            .expect("sheet");
        let row = sheet.row(1).expect("row 2");
        assert!(row.stored());
        assert!(row.hidden());

        let restored = committed
            .workbook()
            .apply(&committed.patch().inverse())
            .expect("inverse");
        assert_eq!(
            restored.workbook().to_bytes().expect("restored bytes"),
            source_bytes
        );

        let mut edit = committed.workbook().edit().expect("show edit");
        edit.sheet(0usize)
            .expect("lookup")
            .expect("sheet")
            .row(RowIndex::new(1).expect("row 2"))
            .expect("checked row")
            .show();
        let shown = edit.commit().expect("show commit");
        let shown_sheet = shown
            .workbook()
            .sheet(0usize)
            .expect("lookup")
            .expect("sheet");
        let shown_row = shown_sheet.row(1).expect("row 2");
        assert!(shown_row.stored());
        assert!(!shown_row.hidden());

        let mut no_op = source.edit().expect("no-op edit");
        no_op
            .sheet(0usize)
            .expect("lookup")
            .expect("sheet")
            .row(10)
            .expect("row 11")
            .show();
        assert!(no_op.commit().expect("no-op commit").patch().is_empty());
        let mut invalid = source.edit().expect("invalid row edit");
        let mut sheet = invalid.sheet(0usize).expect("lookup").expect("sheet");
        assert!(matches!(
            sheet.row(litchi_sheet::ROWS),
            Err(Error::Coordinate(_))
        ));
    }

    #[test]
    fn column_visibility_is_checked_reversible_and_composable() {
        let source = Workbook::new().expect("source workbook");
        let source_bytes = source.to_bytes().expect("source bytes");
        let mut edit = source.edit().expect("edit");
        let mut sheet = edit.sheet("Sheet1").expect("lookup").expect("sheet");
        sheet.set("A1", "left").expect("A1");
        sheet.set("C1", "right").expect("C1");
        sheet.column(1).expect("column B").hide();
        let committed = edit.commit().expect("commit");

        assert_eq!(committed.patch().len(), 3);
        assert!(committed.patch().changes().iter().any(|change| matches!(
            change,
            Change::Column {
                column,
                before: ColumnState::Missing,
                after: ColumnState::Stored { hidden: true },
                ..
            } if column.get() == 1
        )));
        let sheet = committed
            .workbook()
            .sheet("Sheet1")
            .expect("lookup")
            .expect("sheet");
        let column = sheet.column(1).expect("column B");
        assert!(column.stored());
        assert!(column.hidden());
        assert_eq!(sheet.columns().expect("columns").count(), 1);
        assert!(matches!(
            sheet.column_style(1).expect("column style"),
            Some(crate::LocalStyle::Default)
        ));

        let restored = committed
            .workbook()
            .apply(&committed.patch().inverse())
            .expect("inverse");
        assert_eq!(
            restored.workbook().to_bytes().expect("restored bytes"),
            source_bytes
        );

        let mut show = committed.workbook().edit().expect("show edit");
        show.sheet(0usize)
            .expect("lookup")
            .expect("sheet")
            .column(ColumnIndex::new(1).expect("column B"))
            .expect("checked column")
            .show();
        let shown = show.commit().expect("show commit");
        let shown_sheet = shown
            .workbook()
            .sheet(0usize)
            .expect("lookup")
            .expect("sheet");
        let shown_column = shown_sheet.column(1).expect("column B");
        assert!(shown_column.stored());
        assert!(!shown_column.hidden());

        let mut no_op = source.edit().expect("no-op edit");
        let mut sheet = no_op.sheet(0usize).expect("lookup").expect("sheet");
        sheet.column(10).expect("column K").show();
        assert!(matches!(
            sheet.column(litchi_sheet::COLUMNS),
            Err(Error::Coordinate(_))
        ));
        assert!(no_op.commit().expect("no-op commit").patch().is_empty());

        let mut cell = source.edit().expect("cell edit");
        cell.sheet(0usize)
            .expect("lookup")
            .expect("sheet")
            .set("B1", "orthogonal")
            .expect("B1");
        let mut column = source.edit().expect("column edit");
        column
            .sheet(0usize)
            .expect("lookup")
            .expect("sheet")
            .column(1)
            .expect("column B")
            .hide();
        cell.join(column).expect("cell and column join");
        assert!(
            cell.commit()
                .expect("joined commit")
                .workbook()
                .sheet(0usize)
                .expect("lookup")
                .expect("sheet")
                .column(1)
                .expect("column B")
                .hidden()
        );

        let mut left = source.edit().expect("left");
        left.sheet(0usize)
            .expect("lookup")
            .expect("sheet")
            .column(4)
            .expect("column E")
            .hide();
        let mut right = source.edit().expect("right");
        right
            .sheet(0usize)
            .expect("lookup")
            .expect("sheet")
            .column(4)
            .expect("column E")
            .show();
        let error = left.join(right).expect_err("same column must conflict");
        assert_eq!(
            error.conflicts().expect("conflicts").conflicts()[0]
                .columns()
                .expect("column conflict"),
            &[ColumnIndex::new(4).expect("column E")]
        );
    }

    #[test]
    fn tab_visibility_is_selector_first_reversible_and_active_safe() {
        let source = two_sheet_workbook(SheetKind::Worksheet);
        let source_bytes = source.to_bytes().expect("source bytes");
        assert_eq!(
            source.active_sheet().map(|sheet| sheet.name().to_owned()),
            Some("Sheet1".to_owned())
        );

        let mut edit = source.edit().expect("edit");
        edit.tab("Sheet1")
            .expect("name lookup")
            .expect("tab")
            .hide();
        let hidden = edit.commit().expect("hide active tab");
        assert_eq!(hidden.patch().len(), 2);
        let (before, after) = hidden.patch().changes()[0]
            .active()
            .expect("implicit active relocation");
        assert_eq!((before.name(), before.position()), ("Sheet1", 0));
        assert_eq!((after.name(), after.position()), ("Sheet2", 1));
        assert!(matches!(
            hidden.patch().changes()[1],
            Change::Visibility {
                position: 0,
                before: Visibility::Visible,
                after: Visibility::Hidden,
                ..
            }
        ));
        assert_eq!(
            hidden
                .workbook()
                .sheet("Sheet1")
                .expect("lookup")
                .expect("sheet")
                .visibility(),
            &Visibility::Hidden
        );
        assert_eq!(
            hidden
                .workbook()
                .active_sheet()
                .map(|sheet| sheet.name().to_owned()),
            Some("Sheet2".to_owned())
        );
        assert!(
            hidden
                .workbook()
                .sheet("Sheet2")
                .expect("lookup")
                .expect("sheet")
                .is_active()
        );

        let restored = hidden
            .workbook()
            .apply(&hidden.patch().inverse())
            .expect("inverse");
        assert_eq!(restored.workbook().to_bytes().expect("bytes"), source_bytes);

        let mut last = hidden.workbook().edit().expect("last visible edit");
        last.tab(1usize)
            .expect("position lookup")
            .expect("tab")
            .very_hide();
        assert!(matches!(
            last.commit(),
            Err(Error::TabEditBlocked {
                sheet,
                position: 1,
                reason: TabEditBlock::LastVisibleTab,
            }) if sheet == "Sheet2"
        ));

        let mut swap = hidden.workbook().edit().expect("swap visibility");
        swap.tab("Sheet1").expect("lookup").expect("tab").show();
        swap.tab(1usize).expect("lookup").expect("tab").very_hide();
        let swapped = swap.commit().expect("swap commit");
        assert_eq!(
            swapped
                .workbook()
                .sheet(1usize)
                .expect("lookup")
                .expect("sheet")
                .visibility(),
            &Visibility::VeryHidden
        );
        assert_eq!(
            swapped
                .workbook()
                .active_sheet()
                .map(|sheet| sheet.name().to_owned()),
            Some("Sheet1".to_owned())
        );

        let mut no_op = source.edit().expect("no-op edit");
        no_op.tab(0usize).expect("lookup").expect("tab").show();
        assert!(no_op.commit().expect("no-op commit").patch().is_empty());
    }

    #[test]
    fn active_tab_is_selector_first_reversible_and_composable() {
        for kind in [SheetKind::Worksheet, SheetKind::Chart] {
            let source = two_sheet_workbook(kind);
            let source_bytes = source.to_bytes().expect("source bytes");
            assert!(
                source
                    .sheet("Sheet1")
                    .expect("lookup")
                    .expect("first tab")
                    .is_active()
            );
            assert!(
                !source
                    .sheet(1usize)
                    .expect("lookup")
                    .expect("second tab")
                    .is_active()
            );

            let mut edit = source.edit().expect("edit");
            edit.tab(1usize)
                .expect("position lookup")
                .expect("tab")
                .activate();
            assert_eq!(edit.len(), 1);
            let committed = edit.commit().expect("activate");
            assert_eq!(committed.patch().len(), 1);
            let (before, after) = committed.patch().changes()[0]
                .active()
                .expect("active change");
            assert_eq!((before.name(), before.position()), ("Sheet1", 0));
            assert_eq!((after.name(), after.position()), ("Sheet2", 1));
            let active = committed.workbook().active_sheet().expect("active sheet");
            assert_eq!(active.name(), "Sheet2");
            assert_eq!(active.kind(), kind);
            assert!(active.is_active());
            assert_eq!(
                committed
                    .workbook()
                    .apply(&committed.patch().inverse())
                    .expect("inverse")
                    .workbook()
                    .to_bytes()
                    .expect("restored bytes"),
                source_bytes
            );

            let mut no_op = committed.workbook().edit().expect("no-op edit");
            no_op
                .tab("Sheet2")
                .expect("lookup")
                .expect("tab")
                .activate();
            assert!(no_op.commit().expect("no-op commit").patch().is_empty());
        }

        let source = two_sheet_workbook(SheetKind::Worksheet);
        let mut cell = source.edit().expect("cell edit");
        cell.sheet("Sheet2")
            .expect("lookup")
            .expect("worksheet")
            .set("A1", "active payload")
            .expect("cell");
        let mut active = source.edit().expect("active edit");
        active
            .tab("Sheet2")
            .expect("lookup")
            .expect("tab")
            .activate();
        cell.join(active).expect("orthogonal join");
        let committed = cell.commit().expect("joined commit");
        assert_eq!(committed.patch().len(), 2);
        assert_eq!(
            committed
                .patch()
                .parts
                .iter()
                .filter(|part| part.uri == source.inner.sheets[1].part_uri)
                .count(),
            1
        );
        let active = committed.workbook().active_sheet().expect("active sheet");
        assert_eq!(active.name(), "Sheet2");
        assert!(matches!(
            active.cell("A1").expect("cell"),
            Some(Cell::Value(Value::Text(text))) if text.as_str() == "active payload"
        ));

        let mut replaced = source.edit().expect("replacement edit");
        replaced
            .tab("Sheet2")
            .expect("lookup")
            .expect("tab")
            .activate();
        replaced
            .tab("Sheet1")
            .expect("lookup")
            .expect("tab")
            .activate();
        assert_eq!(replaced.len(), 1);
        assert!(
            replaced
                .commit()
                .expect("last activation wins")
                .patch()
                .is_empty()
        );
    }

    #[test]
    fn active_tab_requires_final_visibility_and_conflicts_globally() {
        let source = two_sheet_workbook(SheetKind::Worksheet);
        let mut hide = source.edit().expect("hide edit");
        hide.tab("Sheet2").expect("lookup").expect("tab").hide();
        let hidden = hide.commit().expect("hidden source");

        let mut blocked = hidden.workbook().edit().expect("blocked edit");
        blocked
            .tab("Sheet2")
            .expect("lookup")
            .expect("tab")
            .activate();
        assert!(matches!(
            blocked.commit(),
            Err(Error::TabEditBlocked {
                sheet,
                position: 1,
                reason: TabEditBlock::NotVisible,
            }) if sheet == "Sheet2"
        ));

        let mut repaired = hidden.workbook().edit().expect("repair edit");
        repaired
            .tab("Sheet2")
            .expect("lookup")
            .expect("tab")
            .show()
            .activate();
        let repaired = repaired.commit().expect("show and activate");
        let active = repaired.workbook().active_sheet().expect("active sheet");
        assert_eq!(active.name(), "Sheet2");
        assert!(active.visibility().is_visible());
        assert_eq!(repaired.patch().len(), 2);

        let mut contradictory = source.edit().expect("contradictory edit");
        contradictory
            .tab("Sheet2")
            .expect("lookup")
            .expect("tab")
            .activate()
            .very_hide();
        assert!(matches!(
            contradictory.commit(),
            Err(Error::TabEditBlocked {
                reason: TabEditBlock::NotVisible,
                ..
            })
        ));

        let mut left = source.edit().expect("left");
        left.tab("Sheet2").expect("lookup").expect("tab").activate();
        let mut right = source.edit().expect("right");
        right
            .tab("Sheet1")
            .expect("lookup")
            .expect("tab")
            .activate();
        let error = left.join(right).expect_err("active-tab intents are global");
        let conflicts = error.conflicts().expect("active conflict");
        assert_eq!(conflicts.len(), 1);
        assert!(conflicts.conflicts()[0].is_active());
        assert_eq!(conflicts.conflicts()[0].sheet(), "Sheet2");

        let mut active = source.edit().expect("active");
        active
            .tab("Sheet2")
            .expect("lookup")
            .expect("tab")
            .activate();
        let mut visibility = source.edit().expect("visibility");
        visibility
            .tab("Sheet1")
            .expect("lookup")
            .expect("tab")
            .hide();
        active.join(visibility).expect("orthogonal facets");
        let committed = active.commit().expect("joined commit");
        assert_eq!(
            committed.workbook().active_sheet().expect("active").name(),
            "Sheet2"
        );
        assert_eq!(
            committed
                .workbook()
                .sheet("Sheet1")
                .expect("lookup")
                .expect("sheet")
                .visibility(),
            &Visibility::Hidden
        );
    }

    #[test]
    fn tab_visibility_composes_with_cells_and_conflicts_by_facet() {
        let source = two_sheet_workbook(SheetKind::Worksheet);
        let mut cell = source.edit().expect("cell edit");
        cell.sheet("Sheet2")
            .expect("lookup")
            .expect("worksheet")
            .set("A1", "preserved while hidden")
            .expect("cell");
        let mut tab = source.edit().expect("tab edit");
        tab.tab(1usize).expect("lookup").expect("tab").hide();
        cell.join(tab).expect("orthogonal join");
        let committed = cell.commit().expect("joined commit");
        assert_eq!(committed.patch().len(), 2);
        assert_eq!(
            committed
                .patch()
                .parts
                .iter()
                .filter(|part| part.uri == source.inner.workbook_uri)
                .count(),
            1
        );
        let sheet = committed
            .workbook()
            .sheet("Sheet2")
            .expect("lookup")
            .expect("sheet");
        assert_eq!(sheet.visibility(), &Visibility::Hidden);
        assert!(matches!(
            sheet.cell("A1").expect("cell"),
            Some(Cell::Value(Value::Text(text))) if text.as_str() == "preserved while hidden"
        ));

        let mut left = source.edit().expect("left");
        left.tab("Sheet2").expect("lookup").expect("tab").hide();
        let mut right = source.edit().expect("right");
        right.tab(1usize).expect("lookup").expect("tab").very_hide();
        let error = left.join(right).expect_err("same tab facet must conflict");
        let conflicts = error.conflicts().expect("tab conflicts");
        assert_eq!(conflicts.len(), 1);
        assert!(conflicts.conflicts()[0].is_tab());
        assert_eq!(conflicts.conflicts()[0].sheet(), "Sheet2");
    }

    #[test]
    fn tab_visibility_applies_to_non_worksheet_sheet_kinds() {
        let source = two_sheet_workbook(SheetKind::Chart);
        assert_eq!(
            source
                .sheet("Sheet2")
                .expect("lookup")
                .expect("chart sheet")
                .kind(),
            SheetKind::Chart
        );
        let mut edit = source.edit().expect("edit");
        edit.tab("Sheet2")
            .expect("lookup")
            .expect("tab")
            .very_hide();
        let committed = edit.commit().expect("chart tab commit");
        assert_eq!(
            committed
                .workbook()
                .sheet("Sheet2")
                .expect("lookup")
                .expect("chart sheet")
                .visibility(),
            &Visibility::VeryHidden
        );
    }

    #[test]
    fn active_relocation_synchronizes_worksheet_and_chart_view_selection() {
        for kind in [SheetKind::Worksheet, SheetKind::Chart] {
            let source = active_second_sheet_workbook(kind);
            let source_bytes = source.to_bytes().expect("source bytes");
            assert_eq!(
                source.active_sheet().map(|sheet| sheet.name().to_owned()),
                Some("Sheet2".to_owned())
            );
            let mut edit = source.edit().expect("edit");
            edit.tab("Sheet2").expect("lookup").expect("tab").hide();
            let committed = edit.commit().expect("active hide");
            assert_eq!(
                committed
                    .workbook()
                    .active_sheet()
                    .map(|sheet| sheet.name().to_owned()),
                Some("Sheet1".to_owned())
            );
            let new_active = committed
                .workbook()
                .inner
                .package
                .get_part(&committed.workbook().inner.sheets[0].part_uri)
                .expect("new active part")
                .blob();
            assert!(
                std::str::from_utf8(new_active)
                    .expect("new active XML")
                    .contains(r#"tabSelected="1""#)
            );
            let old_active = committed
                .workbook()
                .inner
                .package
                .get_part(&committed.workbook().inner.sheets[1].part_uri)
                .expect("old active part")
                .blob();
            assert!(
                !std::str::from_utf8(old_active)
                    .expect("old active XML")
                    .contains("tabSelected")
            );
            assert_eq!(
                committed
                    .workbook()
                    .apply(&committed.patch().inverse())
                    .expect("inverse")
                    .workbook()
                    .to_bytes()
                    .expect("restored bytes"),
                source_bytes
            );
        }
    }

    #[test]
    fn tab_visibility_blocks_protected_workbook_structure() {
        let source = two_sheet_workbook(SheetKind::Worksheet);
        let mut package = source.inner.package.clone();
        package
            .get_part_mut(&source.inner.workbook_uri)
            .expect("workbook part")
            .set_blob(
                br#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><workbookProtection lockStructure="1"/><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/><sheet name="Sheet2" sheetId="2" r:id="rIdTab2"/></sheets></workbook>"#.to_vec(),
            );
        let protected = Workbook::from_package(package).expect("protected workbook");
        let mut edit = protected.edit().expect("edit");
        edit.tab("Sheet2").expect("lookup").expect("tab").hide();
        assert!(matches!(
            edit.commit(),
            Err(Error::TabEditBlocked {
                reason: TabEditBlock::ProtectedWorkbook,
                ..
            })
        ));

        let mut activation = protected.edit().expect("activation edit");
        activation
            .tab("Sheet2")
            .expect("lookup")
            .expect("tab")
            .activate();
        let activated = activation
            .commit()
            .expect("structure protection permits selection");
        assert_eq!(
            activated
                .workbook()
                .active_sheet()
                .expect("active sheet")
                .name(),
            "Sheet2"
        );
    }

    #[test]
    fn showing_an_unknown_producer_state_repairs_it_explicitly() {
        let source = two_sheet_workbook(SheetKind::Worksheet);
        let mut package = source.inner.package.clone();
        package
            .get_part_mut(&source.inner.workbook_uri)
            .expect("workbook part")
            .set_blob(
                br#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/><sheet name="Sheet2" sheetId="2" state="show" r:id="rIdTab2"/></sheets></workbook>"#.to_vec(),
            );
        let source = Workbook::from_package(package).expect("producer workbook");
        assert!(matches!(
            source
                .sheet("Sheet2")
                .expect("lookup")
                .expect("sheet")
                .visibility(),
            Visibility::Unknown(value) if value.as_ref() == "show"
        ));
        let source_bytes = source.to_bytes().expect("source bytes");
        let mut edit = source.edit().expect("edit");
        edit.tab("Sheet2").expect("lookup").expect("tab").show();
        let committed = edit.commit().expect("repair commit");
        assert!(
            committed
                .workbook()
                .sheet("Sheet2")
                .expect("lookup")
                .expect("sheet")
                .visibility()
                .is_visible()
        );
        assert_eq!(
            committed
                .workbook()
                .apply(&committed.patch().inverse())
                .expect("inverse")
                .workbook()
                .to_bytes()
                .expect("restored bytes"),
            source_bytes
        );
    }

    #[test]
    fn clearing_a_cell_created_in_the_same_transaction_keeps_an_empty_record() {
        let source = Workbook::new().expect("source workbook");
        let mut edit = source.edit().expect("edit");
        let mut sheet = edit.sheet(0usize).expect("lookup").expect("sheet");
        sheet.set("A1", "temporary").expect("set");
        sheet.clear("A1").expect("clear");
        let committed = edit.commit().expect("commit");
        assert!(matches!(
            committed
                .workbook()
                .sheet(0usize)
                .expect("lookup")
                .expect("sheet")
                .cell("A1")
                .expect("cell"),
            Some(Cell::Empty)
        ));
    }

    #[test]
    fn independently_prepared_disjoint_edits_join_after_threaded_work() {
        fn assert_send<T: Send>() {}
        assert_send::<Edit>();

        let source = Workbook::new().expect("source workbook");
        let (mut left, right) = std::thread::scope(|scope| {
            let left_source = source.clone();
            let right_source = source.clone();
            let left = scope.spawn(move || {
                let mut edit = left_source.edit().expect("left edit");
                edit.sheet("Sheet1")
                    .expect("lookup")
                    .expect("sheet")
                    .set("A1", "left")
                    .expect("left cell");
                edit
            });
            let right = scope.spawn(move || {
                let mut edit = right_source.edit().expect("right edit");
                edit.sheet(0usize)
                    .expect("lookup")
                    .expect("sheet")
                    .set("C3", 42_i32)
                    .expect("right cell");
                edit
            });
            (
                left.join().expect("left worker"),
                right.join().expect("right worker"),
            )
        });

        left.join(right).expect("disjoint join");
        assert_eq!(left.len(), 2);
        let committed = left.commit().expect("joined commit");
        let sheet = committed
            .workbook()
            .sheet("Sheet1")
            .expect("lookup")
            .expect("sheet");
        assert!(matches!(
            sheet.cell("A1").expect("A1"),
            Some(Cell::Value(Value::Text(text))) if text.as_str() == "left"
        ));
        assert!(matches!(
            sheet.cell("C3").expect("C3"),
            Some(Cell::Value(Value::Number(number))) if number.as_str() == "42"
        ));

        let mut empty = source.edit().expect("empty edit");
        let mut incoming = source.edit().expect("incoming edit");
        incoming
            .sheet("Sheet1")
            .expect("lookup")
            .expect("sheet")
            .set("D4", true)
            .expect("incoming cell");
        empty.join(incoming).expect("adopt incoming sheet map");
        assert_eq!(empty.len(), 1);
    }

    #[test]
    fn join_conflicts_are_structured_and_return_the_rejected_edit() {
        let source = Workbook::new().expect("source workbook");
        let mut left = source.edit().expect("left edit");
        let mut left_sheet = left.sheet(0usize).expect("lookup").expect("sheet");
        left_sheet
            .set("C3", "left tail")
            .expect("left tail cell")
            .set("A1", "left")
            .expect("left first cell");
        let mut right = source.edit().expect("right edit");
        let mut right_sheet = right.sheet("Sheet1").expect("lookup").expect("sheet");
        right_sheet
            .set("A1", "right")
            .expect("right first cell")
            .set("C3", "right tail")
            .expect("right tail cell");

        let error = match left.join(right) {
            Ok(_) => panic!("overlapping edits must not join"),
            Err(error) => error,
        };
        assert_eq!(left.len(), 2);
        let conflicts = error.conflicts().expect("overlap details");
        assert_eq!(conflicts.len(), 2);
        assert_eq!(conflicts.conflicts().len(), 1);
        assert_eq!(conflicts.conflicts()[0].sheet(), "Sheet1");
        assert_eq!(conflicts.conflicts()[0].position(), 0);
        assert_eq!(
            conflicts.conflicts()[0].cells().expect("cell conflicts"),
            &[
                Address::from_a1("A1").expect("first address"),
                Address::from_a1("C3").expect("tail address"),
            ]
        );
        let rejected = error.into_rejected();
        assert_eq!(rejected.len(), 2);

        let other_source = Workbook::new().expect("other source");
        let other = other_source.edit().expect("other edit");
        let error = match left.join(other) {
            Ok(_) => panic!("different snapshots must not join"),
            Err(error) => error,
        };
        assert!(matches!(error.failure(), JoinFailure::DifferentSnapshot));
        assert!(error.conflicts().is_none());
        assert_eq!(left.len(), 2);
    }

    #[test]
    fn row_visibility_joins_with_cells_and_conflicts_by_row() {
        let source = Workbook::new().expect("source workbook");
        let mut cell = source.edit().expect("cell edit");
        cell.sheet(0usize)
            .expect("lookup")
            .expect("sheet")
            .set("A2", "same row")
            .expect("cell");
        let mut row = source.edit().expect("row edit");
        row.sheet("Sheet1")
            .expect("lookup")
            .expect("sheet")
            .row(1)
            .expect("row 2")
            .hide();
        cell.join(row).expect("orthogonal row and cell effects");
        let committed = cell.commit().expect("joined commit");
        let sheet = committed
            .workbook()
            .sheet(0usize)
            .expect("lookup")
            .expect("sheet");
        assert!(sheet.row(1).expect("row 2").hidden());
        assert!(matches!(
            sheet.cell("A2").expect("A2"),
            Some(Cell::Value(Value::Text(text))) if text.as_str() == "same row"
        ));

        let mut left = source.edit().expect("left");
        left.sheet(0usize)
            .expect("lookup")
            .expect("sheet")
            .row(4)
            .expect("row 5")
            .hide();
        let mut right = source.edit().expect("right");
        right
            .sheet(0usize)
            .expect("lookup")
            .expect("sheet")
            .row(4)
            .expect("row 5")
            .show();
        let error = left.join(right).expect_err("same row must conflict");
        let conflicts = error.conflicts().expect("row conflicts");
        assert_eq!(conflicts.len(), 1);
        assert_eq!(conflicts.conflicts().len(), 1);
        assert_eq!(
            conflicts.conflicts()[0].rows().expect("row conflict"),
            &[RowIndex::new(4).expect("row 5")]
        );
        assert!(conflicts.conflicts()[0].cells().is_none());
    }

    #[test]
    fn calculation_chain_removal_is_atomic_and_reversible() {
        let baseline = Workbook::new().expect("baseline");
        let mut package = baseline.inner.package.clone();
        let chain_uri = PackURI::new("/xl/calcChain.xml").expect("chain URI");
        package
            .try_add_part(Box::new(BlobPart::new(
                chain_uri.clone(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"
                    .to_owned(),
                br#"<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="A1" i="1"/></calcChain>"#.to_vec(),
            )))
            .expect("chain part");
        package
            .get_part_mut(&baseline.inner.workbook_uri)
            .expect("workbook part")
            .rels_mut()
            .try_add_relationship(
                litchi_opc::constants::relationship_type::CALC_CHAIN.to_owned(),
                "calcChain.xml".to_owned(),
                "rId3".to_owned(),
                TargetMode::Internal,
            )
            .expect("chain relationship");
        let source = Workbook::from_package(package).expect("workbook with chain");
        let source_bytes = source.to_bytes().expect("source bytes");
        let workbook_before = source
            .inner
            .package
            .get_part(&source.inner.workbook_uri)
            .expect("workbook part")
            .blob_arc();

        let mut visibility = source.edit().expect("visibility edit");
        visibility
            .sheet(0usize)
            .expect("lookup")
            .expect("sheet")
            .row(1)
            .expect("row 2")
            .hide();
        let visibility = visibility.commit().expect("visibility commit");
        assert!(visibility.patch().graph.is_empty());
        assert!(
            visibility
                .workbook()
                .inner
                .package
                .get_part(&chain_uri)
                .is_ok()
        );
        assert_eq!(
            visibility
                .workbook()
                .inner
                .package
                .get_part(&source.inner.workbook_uri)
                .expect("unchanged workbook part")
                .blob(),
            workbook_before.as_slice()
        );

        let mut edit = source.edit().expect("edit");
        edit.sheet(0usize)
            .expect("lookup")
            .expect("sheet")
            .set("A1", 7_i32)
            .expect("set");
        let committed = edit.commit().expect("commit");
        assert!(
            committed
                .workbook()
                .inner
                .package
                .get_part(&chain_uri)
                .is_err()
        );
        assert!(
            calculation_chain_removal(committed.workbook())
                .expect("chain query")
                .is_empty()
        );

        let restored = committed
            .workbook()
            .apply(&committed.patch().inverse())
            .expect("inverse");
        assert_eq!(
            restored.workbook().to_bytes().expect("restored"),
            source_bytes
        );
        assert!(
            !calculation_chain_removal(restored.workbook())
                .expect("restored chain")
                .is_empty()
        );

        let mut shared_target = source.inner.package.clone();
        let mut referrer = BlobPart::new(
            PackURI::new("/xl/custom.xml").expect("custom URI"),
            "application/xml".to_owned(),
            b"<custom/>".to_vec(),
        );
        referrer
            .rels_mut()
            .try_add_relationship(
                "urn:litchi:test-reference".to_owned(),
                "calcChain.xml".to_owned(),
                "rId1".to_owned(),
                TargetMode::Internal,
            )
            .expect("extra incoming relationship");
        shared_target
            .try_add_part(Box::new(referrer))
            .expect("referrer part");
        let shared_target = Workbook::from_package(shared_target).expect("shared target workbook");
        assert!(matches!(
            calculation_chain_removal(&shared_target),
            Err(Error::Invalid(message)) if message.contains("another incoming relationship")
        ));
    }

    #[test]
    fn shared_style_crud_is_lineage_checked_reversible_and_exact() {
        let source = styled_workbook();
        let source_bytes = source.to_bytes().expect("source bytes");
        let styles = source.styles().expect("styles");
        assert_eq!(styles.len(), 2);
        let base = styles.base().expect("base style");
        let accent = styles.get(1).expect("accent style");
        assert_eq!(base.fan_out().expect("base fan-out"), 1);
        assert_eq!(accent.fan_out().expect("accent fan-out"), 1);

        let sheet = source.sheet("Sheet1").expect("lookup").expect("sheet");
        assert!(matches!(
            sheet.local_style("B1").expect("local style"),
            Some(crate::LocalStyle::Default)
        ));
        let Some(crate::LocalStyle::Shared(local)) = sheet.local_style("A1").expect("local style")
        else {
            panic!("A1 must have an explicit shared style")
        };
        assert!(local.same(&accent));
        assert!(
            sheet
                .style("B1")
                .expect("resolved style")
                .is_some_and(|style| style.same(&base))
        );

        let mut edit = source.edit().expect("edit");
        let mut sheet = edit.sheet("Sheet1").expect("lookup").expect("sheet");
        sheet
            .set("C1", 42_i32)
            .and_then(|sheet| sheet.style("C1", &accent))
            .and_then(|sheet| sheet.style("D1", &accent))
            .expect("style changes");
        let committed = edit.commit().expect("commit");
        assert_eq!(committed.patch().len(), 2);
        assert!(matches!(
            committed.patch().changes()[0].cell(),
            Some((_, State::Missing, _))
        ));
        assert!(matches!(
            committed.patch().changes()[0].cell(),
            Some((_, _, State::Cell {
                content: Cell::Value(Value::Number(number)),
                style: StyleState::Shared(_),
            })) if number.as_str() == "42"
        ));

        let book = committed.workbook();
        let sheet = book.sheet(0usize).expect("lookup").expect("sheet");
        assert!(matches!(sheet.cell("D1").expect("D1"), Some(Cell::Empty)));
        let styles = book.styles().expect("styles");
        let inherited = styles.find(&accent.key()).expect("inherited style key");
        assert!(inherited.same(&accent));
        assert!(!inherited.same_workbook(&accent));
        assert_eq!(accent.fan_out().expect("source fan-out"), 1);
        assert_eq!(inherited.fan_out().expect("descendant fan-out"), 3);
        assert!(
            sheet
                .style("C1")
                .expect("style")
                .is_some_and(|style| style.same(&inherited))
        );

        let mut descendant = book.edit().expect("descendant edit");
        descendant
            .sheet(0usize)
            .expect("lookup")
            .expect("sheet")
            .style("E1", &accent)
            .expect("reuse inherited style lineage");
        let descendant = descendant.commit().expect("descendant commit");
        assert!(matches!(
            descendant
                .workbook()
                .sheet(0usize)
                .expect("lookup")
                .expect("sheet")
                .local_style("E1")
                .expect("local style"),
            Some(crate::LocalStyle::Shared(_))
        ));

        let reopened = Workbook::from_bytes(source_bytes.clone()).expect("reopened source");
        assert!(
            reopened
                .styles()
                .expect("reopened styles")
                .find(&accent.key())
                .is_none()
        );
        let replayed = reopened
            .apply(committed.patch())
            .expect("replay onto byte-identical source");
        let (
            _,
            _,
            State::Cell {
                style: StyleState::Shared(replayed_key),
                ..
            },
        ) = replayed.patch().changes()[0].cell().expect("cell change")
        else {
            panic!("replayed change must retain its shared style")
        };
        assert!(
            replayed
                .workbook()
                .styles()
                .expect("replayed styles")
                .find(replayed_key)
                .is_some()
        );
        assert!(
            book.styles()
                .expect("original lineage")
                .find(replayed_key)
                .is_none()
        );

        let restored = book
            .apply(&committed.patch().inverse())
            .expect("inverse patch");
        assert_eq!(
            restored.workbook().to_bytes().expect("restored bytes"),
            source_bytes
        );

        let foreign = Workbook::new()
            .expect("other workbook")
            .styles()
            .expect("other styles")
            .base()
            .expect("other base style");
        let mut edit = source.edit().expect("edit");
        assert!(matches!(
            edit.sheet(0usize)
                .expect("lookup")
                .expect("sheet")
                .style("A1", &foreign),
            Err(Error::ForeignStyle)
        ));

        let mut changed_package = source.inner.package.clone();
        let styles_uri = PackURI::new("/xl/styles.xml").expect("styles URI");
        let changed_xml = {
            let styles = changed_package.get_part(&styles_uri).expect("styles part");
            std::str::from_utf8(styles.blob())
                .expect("UTF-8 styles")
                .replace("FFFFFF00", "FFFF0000")
                .into_bytes()
        };
        changed_package
            .get_part_mut(&styles_uri)
            .expect("styles part")
            .set_blob(changed_xml);
        let changed = Workbook::from_package_with_styles(changed_package, Some(&source))
            .expect("changed style table");
        assert!(
            changed
                .styles()
                .expect("changed styles")
                .find(&accent.key())
                .is_none()
        );
        let mut edit = changed.edit().expect("changed edit");
        assert!(matches!(
            edit.sheet(0usize)
                .expect("lookup")
                .expect("sheet")
                .style("A1", &accent),
            Err(Error::ForeignStyle)
        ));
        assert!(matches!(
            changed.apply(committed.patch()),
            Err(Error::PatchConflict { part }) if part == "/xl/styles.xml"
        ));
    }

    #[test]
    fn payload_and_style_effects_on_one_cell_join_without_locks() {
        let source = styled_workbook();
        let accent = source.styles().expect("styles").get(1).expect("accent");
        let (mut payload, style) = std::thread::scope(|scope| {
            let payload_source = source.clone();
            let style_source = source.clone();
            let accent = accent.clone();
            let payload = scope.spawn(move || {
                let mut edit = payload_source.edit().expect("payload edit");
                edit.sheet(0usize)
                    .expect("lookup")
                    .expect("sheet")
                    .set("B1", 9_i32)
                    .expect("payload");
                edit
            });
            let style = scope.spawn(move || {
                let mut edit = style_source.edit().expect("style edit");
                edit.sheet("Sheet1")
                    .expect("lookup")
                    .expect("sheet")
                    .style("B1", &accent)
                    .expect("style");
                edit
            });
            (
                payload.join().expect("payload worker"),
                style.join().expect("style worker"),
            )
        });
        payload.join(style).expect("disjoint cell facets");
        assert_eq!(payload.len(), 1);
        let committed = payload.commit().expect("commit");
        assert_eq!(committed.patch().len(), 1);
        assert!(matches!(
            committed.patch().changes()[0].cell(),
            Some((_, _, State::Cell {
                content: Cell::Value(Value::Number(number)),
                style: StyleState::Shared(_),
            })) if number.as_str() == "9"
        ));

        let sheet = committed
            .workbook()
            .sheet(0usize)
            .expect("lookup")
            .expect("sheet");
        assert!(matches!(
            sheet.cell("B1").expect("B1"),
            Some(Cell::Value(Value::Number(number))) if number.as_str() == "9"
        ));
        assert!(matches!(
            sheet.local_style("B1").expect("style"),
            Some(crate::LocalStyle::Shared(_))
        ));
    }

    #[test]
    fn resetting_style_is_distinct_from_removal_and_missing_is_a_no_op() {
        let source = styled_workbook();
        let source_bytes = source.to_bytes().expect("source bytes");
        let mut edit = source.edit().expect("edit");
        let mut sheet = edit.sheet(0usize).expect("lookup").expect("sheet");
        sheet
            .reset_style("A1")
            .and_then(|sheet| sheet.reset_style("Z99"))
            .expect("style resets");
        let committed = edit.commit().expect("commit");
        assert_eq!(committed.patch().len(), 1);
        assert!(matches!(
            committed.patch().changes()[0].cell(),
            Some((
                _,
                _,
                State::Cell {
                    style: StyleState::Default,
                    ..
                }
            ))
        ));
        let sheet = committed
            .workbook()
            .sheet(0usize)
            .expect("lookup")
            .expect("sheet");
        assert!(matches!(
            sheet.local_style("A1").expect("local style"),
            Some(crate::LocalStyle::Default)
        ));
        assert!(sheet.cell("Z99").expect("missing").is_none());

        let restored = committed
            .workbook()
            .apply(&committed.patch().inverse())
            .expect("inverse");
        assert_eq!(restored.workbook().to_bytes().expect("bytes"), source_bytes);
    }

    #[test]
    fn signed_packages_refuse_edits_before_mutation() {
        let baseline = Workbook::new().expect("baseline");
        let mut package = baseline.inner.package.clone();
        package
            .rels_mut()
            .try_add_relationship(
                litchi_opc::constants::relationship_type::DIGITAL_SIGNATURE_ORIGIN.to_owned(),
                "_xmlsignatures/origin.sigs".to_owned(),
                "rIdSignature".to_owned(),
                TargetMode::Internal,
            )
            .expect("signature relationship");
        let signed = Workbook::from_package(package).expect("signed workbook snapshot");

        assert!(matches!(signed.edit(), Err(Error::Signed)));
        assert!(matches!(
            signed.apply(&Patch::default()),
            Err(Error::Signed)
        ));
    }

    fn styled_workbook() -> Workbook {
        let baseline = Workbook::new().expect("baseline");
        let mut package = baseline.inner.package.clone();
        package
            .get_part_mut(&PackURI::new("/xl/styles.xml").expect("styles URI"))
            .expect("styles part")
            .set_blob(
                br#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font/></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="2" fontId="0" fillId="2" borderId="0" xfId="0" applyNumberFormat="1" applyFill="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>"#.to_vec(),
            );
        package
            .get_part_mut(&PackURI::new("/xl/worksheets/sheet1.xml").expect("sheet URI"))
            .expect("worksheet part")
            .set_blob(
                br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" s="1"><v>1</v></c><c r="B1"><v>2</v></c></row></sheetData></worksheet>"#.to_vec(),
            );
        Workbook::from_package(package).expect("styled workbook")
    }

    fn two_sheet_workbook(second_kind: SheetKind) -> Workbook {
        let baseline = Workbook::new().expect("baseline");
        let mut package = baseline.inner.package.clone();
        package
            .get_part_mut(&baseline.inner.workbook_uri)
            .expect("workbook part")
            .set_blob(
                br#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/><sheet name="Sheet2" sheetId="2" r:id="rIdTab2"/></sheets></workbook>"#.to_vec(),
            );
        let (relationship_type, content_type, part_xml) = match second_kind {
            SheetKind::Worksheet => (
                litchi_opc::constants::relationship_type::WORKSHEET,
                litchi_opc::constants::content_type::SML_WORKSHEET,
                br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1"/><sheetData/></worksheet>"#.as_slice(),
            ),
            SheetKind::Chart => (
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml",
                br#"<chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#.as_slice(),
            ),
            _ => panic!("test helper only models worksheet and chart tabs"),
        };
        package
            .get_part_mut(&baseline.inner.workbook_uri)
            .expect("workbook part")
            .rels_mut()
            .try_add_relationship(
                relationship_type.to_owned(),
                match second_kind {
                    SheetKind::Worksheet => "worksheets/sheet2.xml",
                    SheetKind::Chart => "chartsheets/sheet2.xml",
                    _ => unreachable!("guarded above"),
                }
                .to_owned(),
                "rIdTab2".to_owned(),
                TargetMode::Internal,
            )
            .expect("second sheet relationship");
        let part_uri = match second_kind {
            SheetKind::Worksheet => "/xl/worksheets/sheet2.xml",
            SheetKind::Chart => "/xl/chartsheets/sheet2.xml",
            _ => unreachable!("guarded above"),
        };
        package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new(part_uri).expect("second sheet URI"),
                content_type.to_owned(),
                part_xml.to_vec(),
            )))
            .expect("second sheet part");
        Workbook::from_package(package).expect("two-sheet workbook")
    }

    fn active_second_sheet_workbook(second_kind: SheetKind) -> Workbook {
        let source = two_sheet_workbook(second_kind);
        let mut package = source.inner.package.clone();
        package
            .get_part_mut(&source.inner.workbook_uri)
            .expect("workbook part")
            .set_blob(
                br#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView activeTab="1"/></bookViews><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/><sheet name="Sheet2" sheetId="2" r:id="rIdTab2"/></sheets></workbook>"#.to_vec(),
            );
        package
            .get_part_mut(&source.inner.sheets[0].part_uri)
            .expect("first sheet part")
            .set_blob(
                br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetData/></worksheet>"#.to_vec(),
            );
        let second_xml = match second_kind {
            SheetKind::Worksheet => br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1"/><sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews><sheetData/></worksheet>"#.as_slice(),
            SheetKind::Chart => br#"<chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews></chartsheet>"#.as_slice(),
            _ => unreachable!("test helper only models worksheet and chart tabs"),
        };
        package
            .get_part_mut(&source.inner.sheets[1].part_uri)
            .expect("second sheet part")
            .set_blob(second_xml.to_vec());
        Workbook::from_package(package).expect("active second sheet workbook")
    }
}
