//! Isolated worksheet transactions, disjoint joins, and source-checked patches.

use std::collections::{BTreeMap, BTreeSet, HashMap, HashSet, btree_map::Entry};
use std::fmt;
use std::sync::Arc;

use litchi_opc::{BlobPart, OpcPackage, PackURI, Part, Relationship, TargetMode};
use litchi_sheet::{At, Cell as Address, Column as ColumnIndex, ColumnAt, Row as RowIndex, RowAt};

use super::{Sheet, SheetKind, SheetSelector, Visibility, Workbook};
use crate::cell::{Cell, Content, Stored};
use crate::column::{
    Flags as ColumnFlags, Outline, OutlineAt, Props as ColumnProps, State as ColumnState, WidthAt,
};
use crate::error::{EditBlock, Error, RemoveBlock, Result, TabEditBlock, invalid};
use crate::raw;
use crate::raw::worksheet::edit::{
    Action, ColumnAction, HeightEffect, Payload, Plan, RowAction, StyleEffect, WidthEffect,
};
use crate::row::{Flags as RowFlags, HeightAt, Props as RowProps, State as RowState};
use crate::sheet::Name;
use crate::style::StyleLineage;
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

    /// Checked zero-based workbook position in the corresponding patch state.
    pub const fn position(&self) -> usize {
        self.position
    }
}

impl ColumnState {
    fn read(value: Option<&crate::column::Stored>, workbook: &Workbook) -> Self {
        value.map_or(Self::Missing, |column| {
            Self::Stored(ColumnProps {
                width: column.properties.width,
                style: column.properties.style.map_or(StyleState::Default, |key| {
                    StyleState::Shared(StyleKey::new(
                        key,
                        Arc::clone(&workbook.inner.style_lineage),
                    ))
                }),
                outline: column.properties.outline,
                flags: column.properties.flags,
            })
        })
    }

    fn after(
        before: Option<&crate::column::Stored>,
        action: ColumnAction,
        workbook: &Workbook,
    ) -> Self {
        if before.is_none() && !action.materializes() {
            return Self::Missing;
        }
        let mut properties = match Self::read(before, workbook) {
            Self::Missing => ColumnProps {
                width: None,
                style: StyleState::Default,
                outline: Outline::NONE,
                flags: ColumnFlags::empty(),
            },
            Self::Stored(properties) => properties,
        };
        if let Some(hidden) = action.hidden {
            properties.flags.set(ColumnFlags::HIDDEN, hidden);
        }
        if let Some(width) = action.width {
            match width {
                WidthEffect::Set(width) => {
                    properties.width = Some(width);
                    properties.flags.insert(ColumnFlags::CUSTOM_WIDTH);
                },
                WidthEffect::Reset => {
                    properties.width = None;
                    properties.flags.remove(ColumnFlags::CUSTOM_WIDTH);
                },
            }
        }
        if let Some(style) = action.style {
            properties.style = match style {
                StyleEffect::Set(key) => StyleState::Shared(StyleKey::new(
                    key,
                    Arc::clone(&workbook.inner.style_lineage),
                )),
                StyleEffect::Reset => StyleState::Default,
            };
        }
        if let Some(best_fit) = action.best_fit {
            properties.flags.set(ColumnFlags::BEST_FIT, best_fit);
        }
        if let Some(outline) = action.outline {
            properties.outline = outline;
        }
        if let Some(collapsed) = action.collapsed {
            properties.flags.set(ColumnFlags::COLLAPSED, collapsed);
        }
        if let Some(phonetic) = action.phonetic {
            properties.flags.set(ColumnFlags::PHONETIC, phonetic);
        }
        Self::Stored(properties)
    }
}

impl RowState {
    fn read(value: Option<&crate::row::Stored>, workbook: &Workbook) -> Self {
        value.map_or(Self::Missing, |row| {
            Self::Stored(RowProps {
                height: row.properties.height,
                style: row.properties.style.map_or(StyleState::Default, |key| {
                    StyleState::Shared(StyleKey::new(
                        key,
                        Arc::clone(&workbook.inner.style_lineage),
                    ))
                }),
                outline: row.properties.outline,
                flags: row.properties.flags,
            })
        })
    }

    fn after(before: Option<&crate::row::Stored>, action: RowAction, workbook: &Workbook) -> Self {
        if before.is_none() && !action.materializes() {
            return Self::Missing;
        }
        let mut properties = match Self::read(before, workbook) {
            Self::Missing => RowProps {
                height: None,
                style: StyleState::Default,
                outline: Outline::NONE,
                flags: RowFlags::empty(),
            },
            Self::Stored(properties) => properties,
        };
        if let Some(hidden) = action.hidden {
            properties.flags.set(RowFlags::HIDDEN, hidden);
        }
        if let Some(height) = action.height {
            match height {
                HeightEffect::Set(height) => {
                    properties.height = Some(height);
                    properties.flags.insert(RowFlags::CUSTOM_HEIGHT);
                },
                HeightEffect::Reset => {
                    properties.height = None;
                    properties.flags.remove(RowFlags::CUSTOM_HEIGHT);
                },
            }
        }
        if let Some(style) = action.style {
            match style {
                StyleEffect::Set(key) => {
                    properties.style = StyleState::Shared(StyleKey::new(
                        key,
                        Arc::clone(&workbook.inner.style_lineage),
                    ));
                    properties.flags.insert(RowFlags::CUSTOM_FORMAT);
                },
                StyleEffect::Reset => {
                    properties.style = StyleState::Default;
                    properties.flags.remove(RowFlags::CUSTOM_FORMAT);
                },
            }
        }
        if let Some(outline) = action.outline {
            properties.outline = outline;
        }
        if let Some(collapsed) = action.collapsed {
            properties.flags.set(RowFlags::COLLAPSED, collapsed);
        }
        if let Some(thick_top) = action.thick_top {
            properties.flags.set(RowFlags::THICK_TOP, thick_top);
        }
        if let Some(thick_bottom) = action.thick_bottom {
            properties.flags.set(RowFlags::THICK_BOTTOM, thick_bottom);
        }
        if let Some(phonetic) = action.phonetic {
            properties.flags.set(RowFlags::PHONETIC, phonetic);
        }
        Self::Stored(properties)
    }
}

/// One deterministic semantic change in a reversible patch.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum Change {
    /// A worksheet was added at a checked logical position.
    Create {
        sheet: Box<str>,
        position: usize,
        visibility: Visibility,
    },
    /// A worksheet was removed at a checked logical position.
    Remove {
        sheet: Box<str>,
        position: usize,
        visibility: Visibility,
    },
    Rename {
        position: usize,
        before: Box<str>,
        after: Box<str>,
    },
    Move {
        sheet: Box<str>,
        from: usize,
        to: usize,
    },
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
            Self::Create { sheet, .. } | Self::Remove { sheet, .. } => sheet,
            Self::Rename { after, .. } => after,
            Self::Active { after, .. } => after.name(),
            Self::Move { sheet, .. }
            | Self::Visibility { sheet, .. }
            | Self::Cell { sheet, .. }
            | Self::Row { sheet, .. }
            | Self::Column { sheet, .. } => sheet,
        }
    }

    /// Ordered source/destination positions when this is a tab move.
    pub fn moved(&self) -> Option<(usize, usize)> {
        match self {
            Self::Move { from, to, .. } => Some((*from, *to)),
            Self::Create { .. }
            | Self::Remove { .. }
            | Self::Rename { .. }
            | Self::Active { .. }
            | Self::Visibility { .. }
            | Self::Cell { .. }
            | Self::Row { .. }
            | Self::Column { .. } => None,
        }
    }

    /// Name transition when this is a dependency-aware sheet rename.
    pub fn renamed(&self) -> Option<(usize, &str, &str)> {
        match self {
            Self::Rename {
                position,
                before,
                after,
            } => Some((*position, before, after)),
            Self::Create { .. }
            | Self::Remove { .. }
            | Self::Move { .. }
            | Self::Active { .. }
            | Self::Visibility { .. }
            | Self::Cell { .. }
            | Self::Row { .. }
            | Self::Column { .. } => None,
        }
    }

    /// Active-tab transition when this is a workbook-view change.
    pub fn active(&self) -> Option<(&ActiveTab, &ActiveTab)> {
        match self {
            Self::Active { before, after } => Some((before, after)),
            Self::Create { .. }
            | Self::Remove { .. }
            | Self::Rename { .. }
            | Self::Move { .. }
            | Self::Visibility { .. }
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
            Self::Create { .. }
            | Self::Remove { .. }
            | Self::Rename { .. }
            | Self::Move { .. }
            | Self::Active { .. }
            | Self::Cell { .. }
            | Self::Row { .. }
            | Self::Column { .. } => None,
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
            Self::Create { .. }
            | Self::Remove { .. }
            | Self::Rename { .. }
            | Self::Move { .. }
            | Self::Active { .. }
            | Self::Visibility { .. }
            | Self::Row { .. }
            | Self::Column { .. } => None,
        }
    }

    /// Row state tuple when this is a row-property change.
    pub fn row(&self) -> Option<(RowIndex, &RowState, &RowState)> {
        match self {
            Self::Row {
                row, before, after, ..
            } => Some((*row, before, after)),
            Self::Create { .. }
            | Self::Remove { .. }
            | Self::Rename { .. }
            | Self::Move { .. }
            | Self::Active { .. }
            | Self::Visibility { .. }
            | Self::Cell { .. }
            | Self::Column { .. } => None,
        }
    }

    /// Column state tuple when this is a column-property change.
    pub fn column(&self) -> Option<(ColumnIndex, &ColumnState, &ColumnState)> {
        match self {
            Self::Column {
                column,
                before,
                after,
                ..
            } => Some((*column, before, after)),
            Self::Create { .. }
            | Self::Remove { .. }
            | Self::Rename { .. }
            | Self::Move { .. }
            | Self::Active { .. }
            | Self::Visibility { .. }
            | Self::Cell { .. }
            | Self::Row { .. } => None,
        }
    }

    /// Added worksheet identity when this is a structural create.
    pub fn created(&self) -> Option<(usize, &str)> {
        match self {
            Self::Create {
                sheet, position, ..
            } => Some((*position, sheet)),
            _ => None,
        }
    }

    /// Removed worksheet identity when this is a structural delete or inverse.
    pub fn removed(&self) -> Option<(usize, &str)> {
        match self {
            Self::Remove {
                sheet, position, ..
            } => Some((*position, sheet)),
            _ => None,
        }
    }

    fn inverse(&self) -> Self {
        match self {
            Self::Create {
                sheet,
                position,
                visibility,
            } => Self::Remove {
                sheet: sheet.clone(),
                position: *position,
                visibility: visibility.clone(),
            },
            Self::Remove {
                sheet,
                position,
                visibility,
            } => Self::Create {
                sheet: sheet.clone(),
                position: *position,
                visibility: visibility.clone(),
            },
            Self::Rename {
                position,
                before,
                after,
            } => Self::Rename {
                position: *position,
                before: after.clone(),
                after: before.clone(),
            },
            Self::Move { sheet, from, to } => Self::Move {
                sheet: sheet.clone(),
                from: *to,
                to: *from,
            },
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
                before: after.clone(),
                after: before.clone(),
            },
            Self::Column {
                sheet,
                column,
                before,
                after,
            } => Self::Column {
                sheet: sheet.clone(),
                column: *column,
                before: after.clone(),
                after: before.clone(),
            },
        }
    }

    fn rebind_style(&mut self, workbook: &Workbook) {
        match self {
            Self::Cell { before, after, .. } => {
                before.rebind_style(workbook);
                after.rebind_style(workbook);
            },
            Self::Column { before, after, .. } => {
                before.rebind_style(&workbook.inner.style_lineage);
                after.rebind_style(&workbook.inner.style_lineage);
            },
            Self::Row { before, after, .. } => {
                before.rebind_style(&workbook.inner.style_lineage);
                after.rebind_style(&workbook.inner.style_lineage);
            },
            _ => {},
        }
    }

    fn uses_shared_style(&self) -> bool {
        match self {
            Self::Cell { before, after, .. } => {
                before.uses_shared_style() || after.uses_shared_style()
            },
            Self::Column { before, after, .. } => {
                before.uses_shared_style() || after.uses_shared_style()
            },
            Self::Row { before, after, .. } => {
                before.uses_shared_style() || after.uses_shared_style()
            },
            _ => false,
        }
    }
}

/// Overlapping effects on one logical workbook sheet.
#[derive(Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum Conflict {
    Remove {
        sheet: Box<str>,
        position: usize,
    },
    Name {
        sheet: Box<str>,
        position: usize,
    },
    Order {
        sheet: Box<str>,
        position: usize,
    },
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
            Self::Remove { sheet, .. }
            | Self::Name { sheet, .. }
            | Self::Order { sheet, .. }
            | Self::Active { sheet, .. }
            | Self::Tab { sheet, .. }
            | Self::Cells { sheet, .. }
            | Self::Rows { sheet, .. }
            | Self::Columns { sheet, .. } => sheet,
        }
    }

    /// Checked zero-based sheet position in the shared base snapshot.
    pub fn position(&self) -> usize {
        match self {
            Self::Remove { position, .. }
            | Self::Name { position, .. }
            | Self::Order { position, .. }
            | Self::Active { position, .. }
            | Self::Tab { position, .. }
            | Self::Cells { position, .. }
            | Self::Rows { position, .. }
            | Self::Columns { position, .. } => *position,
        }
    }

    /// Whether both edits target this sheet's catalog name.
    pub const fn is_name(&self) -> bool {
        matches!(self, Self::Name { .. })
    }

    /// Whether both edits remove or otherwise overlap one removed sheet.
    pub const fn is_remove(&self) -> bool {
        matches!(self, Self::Remove { .. })
    }

    /// Whether both edits target this sheet tab's visibility.
    pub const fn is_tab(&self) -> bool {
        matches!(self, Self::Tab { .. })
    }

    /// Whether both edits target the workbook's one active-tab facet.
    pub const fn is_active(&self) -> bool {
        matches!(self, Self::Active { .. })
    }

    /// Whether both edits target the workbook's one tab-order facet.
    pub const fn is_order(&self) -> bool {
        matches!(self, Self::Order { .. })
    }

    /// Deterministically ordered cells written by both edits, when applicable.
    pub fn cells(&self) -> Option<&[Address]> {
        match self {
            Self::Cells { addresses, .. } => Some(addresses),
            Self::Remove { .. }
            | Self::Name { .. }
            | Self::Order { .. }
            | Self::Active { .. }
            | Self::Tab { .. }
            | Self::Rows { .. }
            | Self::Columns { .. } => None,
        }
    }

    /// Deterministically ordered rows written by both edits, when applicable.
    pub fn rows(&self) -> Option<&[RowIndex]> {
        match self {
            Self::Rows { rows, .. } => Some(rows),
            Self::Remove { .. }
            | Self::Name { .. }
            | Self::Order { .. }
            | Self::Active { .. }
            | Self::Tab { .. }
            | Self::Cells { .. }
            | Self::Columns { .. } => None,
        }
    }

    /// Deterministically ordered columns written by both edits, when
    /// applicable.
    pub fn columns(&self) -> Option<&[ColumnIndex]> {
        match self {
            Self::Columns { columns, .. } => Some(columns),
            Self::Remove { .. }
            | Self::Name { .. }
            | Self::Order { .. }
            | Self::Active { .. }
            | Self::Tab { .. }
            | Self::Cells { .. }
            | Self::Rows { .. } => None,
        }
    }

    fn len(&self) -> usize {
        match self {
            Self::Remove { .. }
            | Self::Name { .. }
            | Self::Order { .. }
            | Self::Active { .. }
            | Self::Tab { .. } => 1,
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
    rejected: Box<Edit>,
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
        *self.rejected
    }

    /// Recover both the structured reason and rejected edit.
    pub fn into_parts(self) -> (JoinFailure, Edit) {
        (self.failure, *self.rejected)
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
                .rev()
                .map(Change::inverse)
                .collect::<Vec<_>>()
                .into_boxed_slice(),
            parts: self
                .parts
                .iter()
                .rev()
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
                .rev()
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
    rename: Option<Name>,
    visibility: Option<TabAction>,
    cells: BTreeMap<Address, Action>,
    rows: BTreeMap<RowIndex, RowAction>,
    columns: BTreeMap<ColumnIndex, ColumnAction>,
}

impl SheetActions {
    fn len(&self) -> usize {
        usize::from(self.rename.is_some())
            .saturating_add(usize::from(self.visibility.is_some()))
            .saturating_add(self.cells.len())
            .saturating_add(self.rows.len())
            .saturating_add(self.columns.len())
    }

    fn is_empty(&self) -> bool {
        self.rename.is_none()
            && self.visibility.is_none()
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

#[derive(Debug)]
struct CreatedSheet {
    name: Name,
    position: usize,
    sheet_id: u32,
    relationship_id: String,
    visibility: TabAction,
    graph: GraphChange,
}

#[derive(Debug, Clone, Copy)]
struct MoveIntent {
    sheet: usize,
    from: usize,
    to: usize,
}

#[derive(Debug)]
struct OrderPlan {
    positions: Vec<usize>,
    moves: Vec<MoveIntent>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Target {
    Base(usize),
    Added(usize),
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Placement {
    Tail,
    Before(usize),
    After(usize),
}

#[derive(Debug)]
struct Added {
    name: Name,
    actions: SheetActions,
    placement: Placement,
}

#[derive(Debug, Default)]
struct Around {
    before: Vec<usize>,
    after: Vec<usize>,
}

#[derive(Debug)]
struct FinalOrder {
    targets: Vec<Target>,
    base_positions: Vec<usize>,
    added_positions: Vec<usize>,
}

impl OrderPlan {
    fn is_effective(&self) -> bool {
        self.positions
            .iter()
            .copied()
            .enumerate()
            .any(|(position, identity)| position != identity)
    }
}

impl FinalOrder {
    fn plan(base_len: usize, order: Option<&OrderPlan>, added: &[Added]) -> Result<Self> {
        let final_len = base_len
            .checked_add(added.len())
            .ok_or_else(|| invalid("final worksheet order length overflow"))?;
        let mut around = BTreeMap::<usize, Around>::new();
        let mut tail = Vec::new();
        for (index, sheet) in added.iter().enumerate() {
            match sheet.placement {
                Placement::Tail => {
                    tail.try_reserve(1).map_err(|error| {
                        invalid(format!("cannot reserve appended worksheet order: {error}"))
                    })?;
                    tail.push(index);
                },
                Placement::Before(anchor) | Placement::After(anchor) => {
                    if anchor >= base_len {
                        return Err(invalid("new worksheet anchor is outside the base catalog"));
                    }
                    let slots = around.entry(anchor).or_default();
                    let targets = if matches!(sheet.placement, Placement::Before(_)) {
                        &mut slots.before
                    } else {
                        &mut slots.after
                    };
                    targets.try_reserve(1).map_err(|error| {
                        invalid(format!("cannot reserve anchored worksheet order: {error}"))
                    })?;
                    targets.push(index);
                },
            }
        }

        let mut targets = Vec::new();
        targets
            .try_reserve_exact(final_len)
            .map_err(|error| invalid(format!("cannot reserve final worksheet order: {error}")))?;
        let mut push_base = |identity: usize| -> Result<()> {
            if identity >= base_len {
                return Err(invalid("base worksheet order contains an unknown identity"));
            }
            let slots = around.remove(&identity).unwrap_or_default();
            targets.extend(slots.before.into_iter().map(Target::Added));
            targets.push(Target::Base(identity));
            targets.extend(slots.after.into_iter().map(Target::Added));
            Ok(())
        };
        if let Some(order) = order {
            for identity in order.positions.iter().copied() {
                push_base(identity)?;
            }
        } else {
            for identity in 0..base_len {
                push_base(identity)?;
            }
        }
        if !around.is_empty() {
            return Err(invalid(
                "new worksheet anchor disappeared from the final base order",
            ));
        }
        targets.extend(tail.into_iter().map(Target::Added));
        if targets.len() != final_len {
            return Err(invalid("final worksheet order has the wrong length"));
        }

        let mut base_positions = Vec::new();
        base_positions
            .try_reserve_exact(base_len)
            .map_err(|error| invalid(format!("cannot reserve base tab positions: {error}")))?;
        base_positions.resize(base_len, usize::MAX);
        let mut added_positions = Vec::new();
        added_positions
            .try_reserve_exact(added.len())
            .map_err(|error| invalid(format!("cannot reserve new tab positions: {error}")))?;
        added_positions.resize(added.len(), usize::MAX);
        for (position, target) in targets.iter().copied().enumerate() {
            let slot = match target {
                Target::Base(identity) => base_positions.get_mut(identity),
                Target::Added(index) => added_positions.get_mut(index),
            }
            .ok_or_else(|| invalid("final worksheet order contains an unknown target"))?;
            if *slot != usize::MAX {
                return Err(invalid("final worksheet order repeats a target"));
            }
            *slot = position;
        }
        if base_positions.contains(&usize::MAX) || added_positions.contains(&usize::MAX) {
            return Err(invalid("final worksheet order omits a target"));
        }
        Ok(Self {
            targets,
            base_positions,
            added_positions,
        })
    }

    fn position(&self, target: Target) -> Option<usize> {
        match target {
            Target::Base(identity) => self.base_positions.get(identity),
            Target::Added(index) => self.added_positions.get(index),
        }
        .copied()
    }

    fn target(&self, position: usize) -> Option<Target> {
        self.targets.get(position).copied()
    }

    fn len(&self) -> usize {
        self.targets.len()
    }

    fn matches_appended(&self, base_order: Option<&OrderPlan>) -> bool {
        let base_len = self.base_positions.len();
        self.targets
            .iter()
            .copied()
            .enumerate()
            .all(|(position, actual)| {
                let expected = if position < base_len {
                    let Some(identity) = base_order
                        .map(|order| order.positions.get(position).copied())
                        .unwrap_or(Some(position))
                    else {
                        return false;
                    };
                    Target::Base(identity)
                } else {
                    Target::Added(position - base_len)
                };
                actual == expected
            })
    }
}

/// Isolated workbook transaction. Dropping it rolls back every pending change.
#[derive(Debug)]
pub struct Edit {
    base: Workbook,
    active: Option<Target>,
    order: Option<OrderPlan>,
    sheets: BTreeMap<usize, SheetActions>,
    added: Vec<Added>,
    removed: BTreeSet<usize>,
}

impl Edit {
    pub(super) fn new(base: Workbook) -> Result<Self> {
        ensure_unsigned(&base)?;
        Ok(Self {
            base,
            active: None,
            order: None,
            sheets: BTreeMap::new(),
            added: Vec::new(),
            removed: BTreeSet::new(),
        })
    }

    /// Append a validated worksheet and borrow its transaction-local editor.
    ///
    /// The returned handle can populate cells and properties before the one
    /// atomic commit. Native sheet IDs, relationship IDs, and part names are
    /// allocated deterministically at commit and never enter the public API.
    pub fn add<T>(&mut self, name: T) -> Result<NewSheet<'_>>
    where
        T: TryInto<Name>,
        Error: From<T::Error>,
    {
        self.add_placed(name, Placement::Tail)
    }

    /// Insert a validated worksheet immediately before a semantic anchor.
    ///
    /// The anchor accepts the ordinary name or checked zero-based selector.
    /// `Ok(None)` means it did not resolve in the source snapshot.
    pub fn add_before<'e, 's, T>(
        &'e mut self,
        name: T,
        anchor: impl Into<SheetSelector<'s>>,
    ) -> Result<Option<NewSheet<'e>>>
    where
        T: TryInto<Name>,
        Error: From<T::Error>,
    {
        if !self.removed.is_empty() {
            return Err(self.remove_block(RemoveBlock::MixedEdit, "transaction"));
        }
        let Some(anchor) = self.base.sheet(anchor)? else {
            return Ok(None);
        };
        self.add_placed(name, Placement::Before(anchor.position()))
            .map(Some)
    }

    /// Insert a validated worksheet immediately after a semantic anchor.
    ///
    /// Multiple additions at one anchor retain call order. `Ok(None)` means
    /// the name or checked zero-based anchor did not resolve.
    pub fn add_after<'e, 's, T>(
        &'e mut self,
        name: T,
        anchor: impl Into<SheetSelector<'s>>,
    ) -> Result<Option<NewSheet<'e>>>
    where
        T: TryInto<Name>,
        Error: From<T::Error>,
    {
        if !self.removed.is_empty() {
            return Err(self.remove_block(RemoveBlock::MixedEdit, "transaction"));
        }
        let Some(anchor) = self.base.sheet(anchor)? else {
            return Ok(None);
        };
        self.add_placed(name, Placement::After(anchor.position()))
            .map(Some)
    }

    /// Select a worksheet for short transaction-scoped operations.
    pub fn sheet<'e, 's>(
        &'e mut self,
        selector: impl Into<SheetSelector<'s>>,
    ) -> Result<Option<SheetEdit<'e>>> {
        if !self.removed.is_empty() {
            return Err(self.remove_block(RemoveBlock::MixedEdit, "transaction"));
        }
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
        if !self.removed.is_empty() {
            return Err(self.remove_block(RemoveBlock::MixedEdit, "transaction"));
        }
        let tab = self.base.sheet(selector)?;
        Ok(tab.map(|tab| TabEdit {
            edit: self,
            position: tab.position(),
        }))
    }

    /// Move one tab immediately before another by semantic selector.
    ///
    /// `Ok(None)` means either selector did not resolve in the source
    /// snapshot. Names are the ordinary entry point; checked numeric selectors
    /// remain available for import and positional workflows.
    pub fn move_before<'a, 'b>(
        &mut self,
        sheet: impl Into<SheetSelector<'a>>,
        anchor: impl Into<SheetSelector<'b>>,
    ) -> Result<Option<&mut Self>> {
        if !self.removed.is_empty() {
            return Err(self.remove_block(RemoveBlock::MixedEdit, "transaction"));
        }
        let Some(sheet) = self.base.sheet(sheet)? else {
            return Ok(None);
        };
        let Some(anchor) = self.base.sheet(anchor)? else {
            return Ok(None);
        };
        self.move_relative(sheet.position(), anchor.position(), false)?;
        Ok(Some(self))
    }

    /// Move one tab immediately after another by semantic selector.
    ///
    /// `Ok(None)` means either selector did not resolve in the source snapshot.
    pub fn move_after<'a, 'b>(
        &mut self,
        sheet: impl Into<SheetSelector<'a>>,
        anchor: impl Into<SheetSelector<'b>>,
    ) -> Result<Option<&mut Self>> {
        if !self.removed.is_empty() {
            return Err(self.remove_block(RemoveBlock::MixedEdit, "transaction"));
        }
        let Some(sheet) = self.base.sheet(sheet)? else {
            return Ok(None);
        };
        let Some(anchor) = self.base.sheet(anchor)? else {
            return Ok(None);
        };
        self.move_relative(sheet.position(), anchor.position(), true)?;
        Ok(Some(self))
    }

    /// Move a selected tab to a checked zero-based final position.
    ///
    /// `Ok(None)` means the source selector or destination position does not
    /// exist. Prefer [`Self::move_before`] and [`Self::move_after`] when a
    /// stable semantic anchor is available.
    pub fn move_to<'a>(
        &mut self,
        sheet: impl Into<SheetSelector<'a>>,
        position: usize,
    ) -> Result<Option<&mut Self>> {
        if !self.removed.is_empty() {
            return Err(self.remove_block(RemoveBlock::MixedEdit, "transaction"));
        }
        let Some(sheet) = self.base.sheet(sheet)? else {
            return Ok(None);
        };
        if position >= self.base.len() {
            return Ok(None);
        }
        self.move_position(sheet.position(), position)?;
        Ok(Some(self))
    }

    /// Remove a worksheet selected by its developer-facing name or checked
    /// zero-based source position.
    ///
    /// `Ok(None)` means the selector did not resolve. The safe default refuses
    /// live formulas, unmodeled producer references, VBA projects, additional
    /// incoming relationships, and mixed mutation plans. Multiple independent
    /// worksheet removals may be collected in one atomic transaction.
    pub fn remove<'a>(
        &mut self,
        selector: impl Into<SheetSelector<'a>>,
    ) -> Result<Option<&mut Self>> {
        let Some(sheet) = self.base.sheet(selector)? else {
            return Ok(None);
        };
        if sheet.kind() != SheetKind::Worksheet {
            return Err(Error::NotWorksheet {
                sheet: sheet.name().to_owned(),
            });
        }
        let position = sheet.position();
        let name = sheet.name().to_owned();
        if self.has_non_removal() {
            return Err(Error::SheetRemoveBlocked {
                sheet: name,
                position,
                part: "transaction".to_owned(),
                reason: RemoveBlock::MixedEdit,
            });
        }
        self.removed.insert(position);
        Ok(Some(self))
    }

    pub fn len(&self) -> usize {
        let existing = self.sheets.values().fold(
            self.removed
                .len()
                .saturating_add(usize::from(self.active.is_some()))
                .saturating_add(
                    self.order
                        .as_ref()
                        .filter(|order| order.is_effective())
                        .map_or(0, |order| order.moves.len()),
                ),
            |len, actions| len.saturating_add(actions.len()),
        );
        self.added.iter().fold(existing, |len, added| {
            len.saturating_add(1)
                .saturating_add(added.actions.cells.len())
                .saturating_add(added.actions.rows.len())
                .saturating_add(added.actions.columns.len())
        })
    }

    pub fn is_empty(&self) -> bool {
        self.active.is_none()
            && self
                .order
                .as_ref()
                .is_none_or(|order| !order.is_effective())
            && self.sheets.values().all(SheetActions::is_empty)
            && self.added.is_empty()
            && self.removed.is_empty()
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
                rejected: Box::new(other),
            });
        }
        let conflicts = self.conflicts_with(&other);
        if !conflicts.is_empty() {
            return Err(JoinError {
                failure: JoinFailure::Overlap(conflicts),
                rejected: Box::new(other),
            });
        }

        let added_offset = self.added.len();
        if self.active.is_none() {
            self.active = other.active.map(|target| match target {
                Target::Base(position) => Target::Base(position),
                Target::Added(index) => Target::Added(added_offset.saturating_add(index)),
            });
        }
        if self
            .order
            .as_ref()
            .is_none_or(|order| !order.is_effective())
        {
            self.order = other.order;
        }
        for (position, actions) in other.sheets {
            match self.sheets.entry(position) {
                Entry::Vacant(entry) => {
                    entry.insert(actions);
                },
                Entry::Occupied(entry) => {
                    let accepted = entry.into_mut();
                    if accepted.rename.is_none() {
                        accepted.rename = actions.rename;
                    }
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
                    for (row, action) in actions.rows {
                        match accepted.rows.entry(row) {
                            Entry::Vacant(entry) => {
                                entry.insert(action);
                            },
                            Entry::Occupied(mut entry) => entry.get_mut().merge(action),
                        }
                    }
                    for (column, action) in actions.columns {
                        match accepted.columns.entry(column) {
                            Entry::Vacant(entry) => {
                                entry.insert(action);
                            },
                            Entry::Occupied(mut entry) => entry.get_mut().merge(action),
                        }
                    }
                },
            }
        }
        self.added.extend(other.added);
        self.removed.extend(other.removed);
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
        if !self.removed.is_empty() {
            return commit_removals(self);
        }
        let Self {
            base,
            active: requested_active,
            order: requested_order,
            mut sheets,
            added,
            removed: _,
        } = self;
        let mut changes = Vec::new();
        let mut parts = Vec::new();
        let mut needs_recalculation = false;

        let mut effective_renames = Vec::<(usize, Name)>::new();
        for (position, actions) in &mut sheets {
            let Some(name) = actions.rename.take() else {
                continue;
            };
            let data =
                base.inner.sheets.get(*position).ok_or_else(|| {
                    invalid(format!("renamed sheet position {position} disappeared"))
                })?;
            if name.as_str() != data.name {
                effective_renames.push((*position, name));
            }
        }
        if let Some((position, _)) = effective_renames.first() {
            let data = base
                .inner
                .sheets
                .get(*position)
                .ok_or_else(|| invalid("renamed tab disappeared during edit"))?;
            ensure_reorder_supported(&base, &data.name, *position)?;
        }
        let rename_by_position = effective_renames
            .iter()
            .map(|(position, name)| (*position, name))
            .collect::<HashMap<_, _>>();
        let final_len = base
            .inner
            .sheets
            .len()
            .checked_add(added.len())
            .ok_or_else(|| invalid("final worksheet count overflow"))?;
        if final_len > raw::catalog_edit::MAX_SHEETS {
            let first = added
                .first()
                .ok_or_else(|| invalid("worksheet limit exceeded without a creation"))?;
            return Err(Error::TabEditBlocked {
                sheet: first.name.as_str().to_owned(),
                position: base.inner.sheets.len(),
                reason: TabEditBlock::SheetLimit,
            });
        }
        let effective_order = requested_order.filter(OrderPlan::is_effective);
        if let Some(order) = &effective_order {
            validate_order_plan(order, base.inner.sheets.len())?;
        }
        let final_order =
            FinalOrder::plan(base.inner.sheets.len(), effective_order.as_ref(), &added)?;
        if let Some(first) = added.first() {
            let position = final_order
                .position(Target::Added(0))
                .ok_or_else(|| invalid("first created worksheet has no final position"))?;
            ensure_reorder_supported(&base, first.name.as_str(), position)?;
        }
        let mut final_names = HashMap::<&str, usize>::new();
        final_names
            .try_reserve(final_len)
            .map_err(|error| invalid(format!("cannot reserve final sheet-name index: {error}")))?;
        for (position, target) in final_order.targets.iter().copied().enumerate() {
            let (name, key) = match target {
                Target::Base(identity) => {
                    let data = base.inner.sheets.get(identity).ok_or_else(|| {
                        invalid("existing worksheet disappeared from final names")
                    })?;
                    rename_by_position
                        .get(&identity)
                        .map_or((data.name.as_str(), data.name_key.as_ref()), |name| {
                            (name.as_str(), name.identity_key())
                        })
                },
                Target::Added(index) => {
                    let created = added
                        .get(index)
                        .ok_or_else(|| invalid("created worksheet disappeared from final names"))?;
                    (created.name.as_str(), created.name.identity_key())
                },
            };
            if let Some(first) = final_names.insert(key, position) {
                return Err(Error::SheetNameConflict {
                    name: name.to_owned(),
                    first,
                    second: position,
                });
            }
        }
        for (position, after) in &effective_renames {
            let before = base
                .inner
                .sheets
                .get(*position)
                .ok_or_else(|| invalid("renamed tab disappeared during patch creation"))?;
            changes.push(Change::Rename {
                position: *position,
                before: before.name.clone().into_boxed_str(),
                after: after.as_str().into(),
            });
        }

        if let Some(order) = &effective_order {
            let first = order
                .moves
                .first()
                .ok_or_else(|| invalid("effective tab order has no semantic move"))?;
            let data = base
                .inner
                .sheets
                .get(first.sheet)
                .ok_or_else(|| invalid("moved tab disappeared during edit"))?;
            ensure_reorder_supported(&base, &data.name, first.from)?;
            for moved in &order.moves {
                let data = base
                    .inner
                    .sheets
                    .get(moved.sheet)
                    .ok_or_else(|| invalid("moved tab disappeared during patch creation"))?;
                changes.push(Change::Move {
                    sheet: data.name.clone().into_boxed_str(),
                    from: moved.from,
                    to: moved.to,
                });
            }
        }

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

        let final_base_is_visible = |position: usize| {
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
        let added_is_visible = |index: usize| {
            added.get(index).is_some_and(|sheet| {
                sheet
                    .actions
                    .visibility
                    .is_none_or(|action| action == TabAction::Show)
            })
        };
        let any_visible = (0..base.inner.sheets.len()).any(&final_base_is_visible)
            || (0..added.len()).any(added_is_visible);
        if !effective_tabs.is_empty() && !any_visible {
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

        let final_position = |target: Target| final_order.position(target);
        let final_is_visible = |target: Target| match target {
            Target::Base(position) => final_base_is_visible(position),
            Target::Added(index) => added_is_visible(index),
        };

        let current_active = base.inner.active_sheet;
        let current_target = current_active.map(Target::Base);
        let final_active = if let Some(target) = requested_active {
            let name = match target {
                Target::Base(identity) => {
                    let data =
                        base.inner.sheets.get(identity).ok_or_else(|| {
                            invalid("requested active tab disappeared during edit")
                        })?;
                    data.name.as_str()
                },
                Target::Added(index) => {
                    let data = added.get(index).ok_or_else(|| {
                        invalid("requested new active tab disappeared during edit")
                    })?;
                    data.name.as_str()
                },
            };
            let position = final_order
                .position(target)
                .ok_or_else(|| invalid("requested active tab has no final position"))?;
            if !final_is_visible(target) {
                return Err(Error::TabEditBlocked {
                    sheet: name.to_owned(),
                    position,
                    reason: TabEditBlock::NotVisible,
                });
            }
            Some(target)
        } else if effective_tabs.is_empty() || current_target.is_some_and(final_is_visible) {
            current_target
        } else {
            let len = final_order.len();
            if len == 0 {
                None
            } else {
                current_target
                    .and_then(final_position)
                    .and_then(|current_position| {
                        (1..=len)
                            .filter_map(|offset| {
                                current_position.checked_add(offset).map(|sum| sum % len)
                            })
                            .filter_map(|position| final_order.target(position))
                            .find(|target| final_is_visible(*target))
                    })
                    .or_else(|| {
                        final_order
                            .targets
                            .iter()
                            .copied()
                            .find(|target| final_is_visible(*target))
                    })
            }
        };
        let final_active_position = final_active.and_then(final_position);
        if let Some(position) = final_active_position
            && position > raw::catalog_edit::MAX_ACTIVE_TAB
        {
            let target =
                final_active.ok_or_else(|| invalid("active position has no sheet identity"))?;
            let name = match target {
                Target::Base(identity) => base
                    .inner
                    .sheets
                    .get(identity)
                    .map(|sheet| sheet.name.as_str()),
                Target::Added(index) => added.get(index).map(|sheet| sheet.name.as_str()),
            }
            .ok_or_else(|| invalid("replacement active tab disappeared during edit"))?;
            return Err(Error::TabEditBlocked {
                sheet: name.to_owned(),
                position,
                reason: TabEditBlock::ActiveTabLimit,
            });
        }

        let active_before = current_active
            .map(|identity| active_tab_at(&base, identity, identity, None))
            .transpose()?;
        let active_after = final_active
            .zip(final_active_position)
            .map(|(target, position)| match target {
                Target::Base(identity) => active_tab_at(
                    &base,
                    identity,
                    position,
                    rename_by_position.get(&identity).map(|name| name.as_str()),
                ),
                Target::Added(index) => added
                    .get(index)
                    .map(|sheet| ActiveTab {
                        name: sheet.name.as_str().into(),
                        position,
                    })
                    .ok_or_else(|| invalid("new active tab disappeared during patch creation")),
            })
            .transpose()?;
        let active_change = (current_target.zip(current_active)
            != final_active.zip(final_active_position))
        .then_some(final_active_position)
        .flatten();
        if active_change.is_some() {
            let before = active_before
                .ok_or_else(|| invalid("non-empty workbook has no source active tab"))?;
            let after = active_after
                .ok_or_else(|| invalid("non-empty workbook has no final active tab"))?;
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
                rename: _,
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
                let before_state = RowState::read(before, &base);
                let after_state = RowState::after(before, action, &base);
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
                let before_state = ColumnState::read(before, &base);
                let after_state = ColumnState::after(before, action, &base);
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
                    Change::Create { .. }
                    | Change::Remove { .. }
                    | Change::Rename { .. }
                    | Change::Move { .. }
                    | Change::Active { .. }
                    | Change::Visibility { .. } => {},
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
                        let actual = RowState::read(parsed.row_entry(*row), &base);
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
                        let actual = ColumnState::read(parsed.column_entry(*column), &base);
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

        let active_added = match final_active {
            Some(Target::Added(index)) => Some(index),
            Some(Target::Base(_)) | None => None,
        };
        let mut created = create_sheets(
            &base,
            added,
            &final_order.added_positions,
            active_added,
            &mut changes,
            &mut needs_recalculation,
        )?;

        if final_active != current_target {
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
            let new_active = final_active
                .ok_or_else(|| invalid("active tab disappeared during selection rewrite"))?;
            if let Target::Base(new_active) = new_active {
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
                            position: final_active_position.unwrap_or(new_active),
                        },
                    )
                })?;
            }
        }

        let reference_renames = effective_renames
            .iter()
            .map(|(position, after)| {
                let before = base.inner.sheets.get(*position).ok_or_else(|| {
                    invalid("renamed sheet disappeared during reference planning")
                })?;
                Ok(raw::reference_edit::Rename {
                    before: &before.name,
                    after: after.as_str(),
                })
            })
            .collect::<Result<Vec<_>>>()?;
        let sheet_titles = if effective_renames.is_empty() {
            Vec::new()
        } else {
            base.inner
                .sheets
                .iter()
                .map(|sheet| sheet.name.as_str())
                .collect::<Vec<_>>()
        };
        if let Some((position, _)) = effective_renames.first() {
            let sheet = base
                .inner
                .sheets
                .get(*position)
                .ok_or_else(|| invalid("rename context sheet disappeared"))?;
            let mut reference_parts = base
                .inner
                .package
                .iter_parts()
                .filter(|part| part.partname() != &base.inner.workbook_uri && reference_part(*part))
                .map(|part| {
                    (
                        part.partname().clone(),
                        part.content_type()
                            == litchi_opc::constants::content_type::OFC_EXTENDED_PROPERTIES,
                    )
                })
                .collect::<Vec<_>>();
            reference_parts.sort_unstable_by(|left, right| left.0.as_str().cmp(right.0.as_str()));
            for (uri, extended_properties) in reference_parts {
                let part_name = uri.to_string();
                let part_titles = if extended_properties {
                    sheet_titles.as_slice()
                } else {
                    &[]
                };
                compose_part_optional(&mut parts, &base, &uri, |content| {
                    raw::reference_edit::rewrite(
                        content,
                        &reference_renames,
                        raw::reference_edit::Context {
                            sheet: &sheet.name,
                            position: *position,
                            part: &part_name,
                            sheet_titles: part_titles,
                        },
                    )
                })?;
            }
            for created_sheet in &mut created {
                let part_name = created_sheet.graph.part.partname().to_string();
                if let Some(after) = raw::reference_edit::rewrite(
                    created_sheet.graph.part.blob(),
                    &reference_renames,
                    raw::reference_edit::Context {
                        sheet: &sheet.name,
                        position: *position,
                        part: &part_name,
                        sheet_titles: &[],
                    },
                )? {
                    created_sheet.graph.part.set_blob(after);
                }
                let parsed = raw::worksheet::parse(created_sheet.graph.part.blob(), || {
                    base.inner.shared_strings()
                })?;
                base.inner.validate_styles(&parsed)?;
                for change in &mut changes {
                    if let Change::Cell {
                        sheet,
                        address,
                        after,
                        ..
                    } = change
                        && sheet.as_ref() == created_sheet.name.as_str()
                    {
                        *after = State::read(parsed.entry(*address), &base);
                    }
                }
            }
        }
        if effective_order.is_some() || !created.is_empty() {
            let existing_titles = base
                .inner
                .sheets
                .iter()
                .enumerate()
                .map(|(position, sheet)| {
                    rename_by_position
                        .get(&position)
                        .map_or(sheet.name.as_str(), |name| name.as_str())
                })
                .collect::<Vec<_>>();
            let mut property_order = Vec::new();
            property_order
                .try_reserve_exact(final_order.len())
                .map_err(|error| {
                    invalid(format!(
                        "cannot reserve extended-properties sheet order: {error}"
                    ))
                })?;
            for target in final_order.targets.iter().copied() {
                property_order.push(match target {
                    Target::Base(identity) => raw::properties_edit::Sheet::Existing(identity),
                    Target::Added(index) => {
                        let sheet = created.get(index).ok_or_else(|| {
                            invalid("created worksheet disappeared from property order")
                        })?;
                        raw::properties_edit::Sheet::New(sheet.name.as_str())
                    },
                });
            }
            let mut property_parts = base
                .inner
                .package
                .iter_parts()
                .filter(|part| {
                    part.content_type()
                        == litchi_opc::constants::content_type::OFC_EXTENDED_PROPERTIES
                })
                .map(|part| part.partname().clone())
                .collect::<Vec<_>>();
            property_parts.sort_unstable_by(|left, right| left.as_str().cmp(right.as_str()));
            for uri in property_parts {
                compose_part_optional(&mut parts, &base, &uri, |content| {
                    raw::properties_edit::arrange_sheets(content, &existing_titles, &property_order)
                })?;
            }
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
        let calculation_graph = if needs_recalculation {
            calculation_chain_removal(&base)?
        } else {
            Vec::new()
        };

        if !effective_renames.is_empty()
            || !effective_tabs.is_empty()
            || active_change.is_some()
            || effective_order.is_some()
            || !created.is_empty()
            || needs_recalculation
        {
            let workbook_part = base.inner.package.get_part(&base.inner.workbook_uri)?;
            let before = workbook_part.blob_arc();
            let referenced_workbook = if let Some((position, _)) = effective_renames.first() {
                let sheet = base
                    .inner
                    .sheets
                    .get(*position)
                    .ok_or_else(|| invalid("rename workbook context disappeared"))?;
                raw::reference_edit::rewrite(
                    &before,
                    &reference_renames,
                    raw::reference_edit::Context {
                        sheet: &sheet.name,
                        position: *position,
                        part: base.inner.workbook_uri.as_str(),
                        sheet_titles: &[],
                    },
                )?
            } else {
                None
            };
            let workbook_input = referenced_workbook.as_deref().unwrap_or(&before);
            let scope_source = (effective_order.is_some() || !created.is_empty())
                .then(|| raw::parse_catalog(workbook_input))
                .transpose()?;
            let raw_order = effective_order
                .as_ref()
                .map(|order| {
                    let moved = order
                        .moves
                        .first()
                        .ok_or_else(|| invalid("effective tab order lost its move context"))?;
                    let data =
                        base.inner.sheets.get(moved.sheet).ok_or_else(|| {
                            invalid("tab reorder context disappeared during commit")
                        })?;
                    let relationship_ids = order
                        .positions
                        .iter()
                        .map(|identity| {
                            base.inner
                                .sheets
                                .get(*identity)
                                .map(|sheet| sheet.relationship_id.as_str())
                                .ok_or_else(|| {
                                    invalid("tab reorder target disappeared during commit")
                                })
                        })
                        .collect::<Result<Vec<_>>>()?;
                    Ok::<_, Error>(raw::catalog_edit::Order {
                        sheet: &data.name,
                        position: moved.from,
                        relationship_ids,
                        local_scopes: base
                            .inner
                            .defined_names
                            .iter()
                            .filter(|name| name.local_sheet_id.is_some())
                            .count(),
                    })
                })
                .transpose()?;
            let raw_renames = effective_renames
                .iter()
                .map(|(position, after)| {
                    let data =
                        base.inner.sheets.get(*position).ok_or_else(|| {
                            invalid("tab rename target disappeared during commit")
                        })?;
                    Ok(raw::catalog_edit::Rename {
                        sheet: &data.name,
                        position: *position,
                        relationship_id: &data.relationship_id,
                        name: after.as_str(),
                    })
                })
                .collect::<Result<Vec<_>>>()?;
            let has_existing_catalog_edit =
                !effective_renames.is_empty() || !effective_tabs.is_empty() || raw_order.is_some();
            let mut after = if has_existing_catalog_edit {
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
                raw::catalog_edit::rewrite(
                    workbook_input,
                    raw::catalog_edit::Plan {
                        tabs,
                        renames: raw_renames,
                        active: None,
                        order: raw_order,
                    },
                )?
            } else {
                workbook_input.to_vec()
            };
            for (index, sheet) in created.iter().enumerate() {
                let position = base
                    .inner
                    .sheets
                    .len()
                    .checked_add(index)
                    .ok_or_else(|| invalid("physical worksheet append position overflow"))?;
                after = raw::catalog_edit::append(
                    &after,
                    raw::catalog_edit::Create {
                        sheet: sheet.name.as_str(),
                        position,
                        sheet_id: sheet.sheet_id,
                        relationship_id: &sheet.relationship_id,
                        state: sheet.visibility.raw(),
                    },
                )?;
            }
            if !final_order.matches_appended(effective_order.as_ref()) {
                let mut relationship_ids = Vec::new();
                relationship_ids
                    .try_reserve_exact(final_order.len())
                    .map_err(|error| {
                        invalid(format!("cannot reserve final catalog order: {error}"))
                    })?;
                for target in final_order.targets.iter().copied() {
                    relationship_ids.push(match target {
                        Target::Base(identity) => base
                            .inner
                            .sheets
                            .get(identity)
                            .map(|sheet| sheet.relationship_id.as_str())
                            .ok_or_else(|| {
                                invalid("existing worksheet disappeared from final catalog order")
                            })?,
                        Target::Added(index) => created
                            .get(index)
                            .map(|sheet| sheet.relationship_id.as_str())
                            .ok_or_else(|| {
                                invalid("created worksheet disappeared from final catalog order")
                            })?,
                    });
                }
                let context = final_order
                    .targets
                    .first()
                    .copied()
                    .ok_or_else(|| invalid("final catalog order has no context sheet"))?;
                let context_name = match context {
                    Target::Base(identity) => base
                        .inner
                        .sheets
                        .get(identity)
                        .map(|sheet| sheet.name.as_str()),
                    Target::Added(index) => created.get(index).map(|sheet| sheet.name.as_str()),
                }
                .ok_or_else(|| invalid("final catalog context sheet disappeared"))?;
                after = raw::catalog_edit::rewrite(
                    &after,
                    raw::catalog_edit::Plan {
                        tabs: Vec::new(),
                        renames: Vec::new(),
                        active: None,
                        order: Some(raw::catalog_edit::Order {
                            sheet: context_name,
                            position: 0,
                            relationship_ids,
                            local_scopes: base
                                .inner
                                .defined_names
                                .iter()
                                .filter(|name| name.local_sheet_id.is_some())
                                .count(),
                        }),
                    },
                )?;
            }
            if let Some(position) = active_change {
                let target = final_active
                    .ok_or_else(|| invalid("active-tab rewrite has no final sheet identity"))?;
                let sheet = match target {
                    Target::Base(identity) => base
                        .inner
                        .sheets
                        .get(identity)
                        .map(|sheet| sheet.name.as_str()),
                    Target::Added(index) => created.get(index).map(|sheet| sheet.name.as_str()),
                }
                .ok_or_else(|| invalid("active-tab rewrite target disappeared"))?;
                after = raw::catalog_edit::rewrite(
                    &after,
                    raw::catalog_edit::Plan {
                        tabs: Vec::new(),
                        renames: Vec::new(),
                        active: Some(raw::catalog_edit::Active { sheet, position }),
                        order: None,
                    },
                )?;
            }
            if needs_recalculation {
                after = raw::recalc::invalidate(&after)?;
            }
            if has_existing_catalog_edit || !created.is_empty() || active_change.is_some() {
                let catalog = raw::parse_catalog(&after)?;
                if Some(catalog.active_sheet_index) != final_active_position {
                    return Err(invalid("workbook active-tab edit verification failed"));
                }
                if catalog.sheets.len() != final_order.len() {
                    return Err(invalid(
                        "workbook creation verification has the wrong sheet count",
                    ));
                }
                for (position, target) in final_order.targets.iter().copied().enumerate() {
                    let expected = match target {
                        Target::Base(identity) => base
                            .inner
                            .sheets
                            .get(identity)
                            .map(|sheet| sheet.relationship_id.as_str()),
                        Target::Added(index) => created
                            .get(index)
                            .map(|sheet| sheet.relationship_id.as_str()),
                    }
                    .ok_or_else(|| invalid("final worksheet order lost a sheet identity"))?;
                    let actual = catalog
                        .sheets
                        .get(position)
                        .ok_or_else(|| invalid("final worksheet order lost a catalog slot"))?;
                    if actual.relationship_id != expected {
                        return Err(invalid(format!(
                            "workbook order verification failed at position {position}"
                        )));
                    }
                }
                for (identity, action) in &effective_tabs {
                    let position = final_position(Target::Base(*identity)).ok_or_else(|| {
                        invalid("visible tab disappeared from the final workbook order")
                    })?;
                    let actual = catalog
                        .sheets
                        .get(position)
                        .ok_or_else(|| invalid("workbook tab edit verification lost a sheet"))?;
                    if !raw_visibility_matches(&actual.visibility, *action) {
                        let sheet = base
                            .inner
                            .sheets
                            .get(*identity)
                            .map_or("<missing sheet>", |sheet| sheet.name.as_str());
                        return Err(invalid(format!(
                            "workbook tab visibility verification failed at {sheet}"
                        )));
                    }
                }
                for (identity, expected_name) in &effective_renames {
                    let position = final_position(Target::Base(*identity)).ok_or_else(|| {
                        invalid("renamed tab disappeared from the final workbook order")
                    })?;
                    let actual = catalog
                        .sheets
                        .get(position)
                        .ok_or_else(|| invalid("workbook rename verification lost a sheet"))?;
                    let expected = base.inner.sheets.get(*identity).ok_or_else(|| {
                        invalid("renamed tab identity disappeared during verification")
                    })?;
                    if actual.relationship_id != expected.relationship_id
                        || actual.name != expected_name.as_str()
                    {
                        return Err(invalid(format!(
                            "workbook tab rename verification failed at position {position}"
                        )));
                    }
                }
                if effective_order.is_some() || !created.is_empty() {
                    let source = scope_source.as_ref().ok_or_else(|| {
                        invalid("defined-name scope source disappeared during verification")
                    })?;
                    verify_defined_name_scopes(
                        source,
                        &catalog,
                        base.inner.sheets.len(),
                        &final_order.targets,
                    )?;
                }
                for sheet in &created {
                    let actual = catalog.sheets.get(sheet.position).ok_or_else(|| {
                        invalid("created worksheet disappeared during catalog verification")
                    })?;
                    if actual.name != sheet.name.as_str()
                        || actual.relationship_id != sheet.relationship_id
                        || actual.sheet_id != sheet.sheet_id
                        || !raw_visibility_matches(&actual.visibility, sheet.visibility)
                    {
                        return Err(invalid(format!(
                            "workbook creation verification failed at position {}",
                            sheet.position
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

        let mut graph = Vec::new();
        graph
            .try_reserve(created.len().saturating_add(calculation_graph.len()))
            .map_err(|error| invalid(format!("cannot reserve package graph changes: {error}")))?;
        graph.extend(created.into_iter().map(|sheet| sheet.graph));
        graph.extend(calculation_graph);

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

    fn add_placed<T>(&mut self, name: T, placement: Placement) -> Result<NewSheet<'_>>
    where
        T: TryInto<Name>,
        Error: From<T::Error>,
    {
        if !self.removed.is_empty() {
            return Err(self.remove_block(RemoveBlock::MixedEdit, "transaction"));
        }
        let name = name.try_into().map_err(Error::from)?;
        let limit_position = self
            .base
            .len()
            .checked_add(self.added.len())
            .ok_or_else(|| invalid("worksheet position overflow"))?;
        if limit_position >= raw::catalog_edit::MAX_SHEETS {
            return Err(Error::TabEditBlocked {
                sheet: name.as_str().to_owned(),
                position: limit_position,
                reason: TabEditBlock::SheetLimit,
            });
        }
        self.added
            .try_reserve(1)
            .map_err(|error| invalid(format!("cannot reserve worksheet creation: {error}")))?;
        let index = self.added.len();
        self.added.push(Added {
            name,
            actions: SheetActions::default(),
            placement,
        });
        let effective_order = self.order.as_ref().filter(|order| order.is_effective());
        let position = match FinalOrder::plan(self.base.len(), effective_order, &self.added)
            .and_then(|order| {
                order
                    .position(Target::Added(index))
                    .ok_or_else(|| invalid("new worksheet has no projected position"))
            }) {
            Ok(position) => position,
            Err(error) => {
                if self.added.pop().is_none() {
                    return Err(invalid(
                        "new worksheet planning failed and its rollback target disappeared",
                    ));
                }
                return Err(error);
            },
        };
        if self.added.get(index).is_none() {
            let _ = self.added.pop();
            return Err(invalid("worksheet creation plan disappeared"));
        }
        let added = self
            .added
            .get_mut(index)
            .ok_or_else(|| invalid("worksheet creation plan disappeared after validation"))?;
        Ok(NewSheet {
            added,
            active: &mut self.active,
            style_lineage: &self.base.inner.style_lineage,
            index,
            position,
        })
    }

    fn actions(&mut self, position: usize) -> &mut BTreeMap<Address, Action> {
        &mut self.sheets.entry(position).or_default().cells
    }

    fn set_visibility(&mut self, position: usize, action: TabAction) {
        self.sheets.entry(position).or_default().visibility = Some(action);
    }

    fn set_name(&mut self, position: usize, name: Name) {
        self.sheets.entry(position).or_default().rename = Some(name);
    }

    fn set_active(&mut self, position: usize) {
        self.active = Some(Target::Base(position));
    }

    fn order_plan(&mut self) -> Result<&mut OrderPlan> {
        if self.order.is_none() {
            let len = self.base.inner.sheets.len();
            let mut positions = Vec::new();
            positions
                .try_reserve_exact(len)
                .map_err(|error| invalid(format!("cannot reserve tab-order plan: {error}")))?;
            positions.extend(0..len);
            self.order = Some(OrderPlan {
                positions,
                moves: Vec::new(),
            });
        }
        self.order
            .as_mut()
            .ok_or_else(|| invalid("tab-order plan initialization failed"))
    }

    fn move_position(&mut self, identity: usize, to: usize) -> Result<()> {
        let order = self.order_plan()?;
        if to >= order.positions.len() {
            return Err(invalid("tab move destination exceeds the workbook order"));
        }
        let from = order
            .positions
            .iter()
            .position(|candidate| *candidate == identity)
            .ok_or_else(|| invalid("selected tab disappeared from the pending order"))?;
        if from == to {
            return Ok(());
        }
        order
            .moves
            .try_reserve(1)
            .map_err(|error| invalid(format!("cannot reserve tab move: {error}")))?;
        let identity = order.positions.remove(from);
        order.positions.insert(to, identity);
        order.moves.push(MoveIntent {
            sheet: identity,
            from,
            to,
        });
        Ok(())
    }

    fn move_relative(&mut self, identity: usize, anchor: usize, after: bool) -> Result<()> {
        if identity == anchor {
            return Ok(());
        }
        let order = self.order_plan()?;
        let from = order
            .positions
            .iter()
            .position(|candidate| *candidate == identity)
            .ok_or_else(|| invalid("selected tab disappeared from the pending order"))?;
        if !order.positions.contains(&anchor) {
            return Err(invalid("anchor tab disappeared from the pending order"));
        }
        order
            .moves
            .try_reserve(1)
            .map_err(|error| invalid(format!("cannot reserve tab move: {error}")))?;
        let identity = order.positions.remove(from);
        let anchor = order
            .positions
            .iter()
            .position(|candidate| *candidate == anchor)
            .ok_or_else(|| invalid("anchor tab disappeared during reorder"))?;
        let to = if after {
            anchor
                .checked_add(1)
                .ok_or_else(|| invalid("tab move position overflow"))?
        } else {
            anchor
        };
        order.positions.insert(to, identity);
        if from != to {
            order.moves.push(MoveIntent {
                sheet: identity,
                from,
                to,
            });
        }
        Ok(())
    }

    fn conflicts_with(&self, other: &Self) -> ConflictSet {
        let mut conflicts = Vec::new();
        let removal_conflict = self
            .removed
            .intersection(&other.removed)
            .next()
            .copied()
            .or_else(|| {
                (!self.removed.is_empty() && other.has_non_removal())
                    .then(|| self.removed.iter().next().copied())
                    .flatten()
            })
            .or_else(|| {
                (!other.removed.is_empty() && self.has_non_removal())
                    .then(|| other.removed.iter().next().copied())
                    .flatten()
            });
        if let Some(position) = removal_conflict {
            let sheet = self
                .base
                .inner
                .sheets
                .get(position)
                .map_or("<missing sheet>", |sheet| sheet.name.as_str());
            conflicts.push(Conflict::Remove {
                sheet: sheet.into(),
                position,
            });
        }
        if let (Some(left), Some(_)) = (
            self.order.as_ref().filter(|order| order.is_effective()),
            other.order.as_ref().filter(|order| order.is_effective()),
        ) {
            let moved = left.moves.first();
            let position = moved.map_or(0, |moved| moved.from);
            let sheet = moved
                .and_then(|moved| self.base.inner.sheets.get(moved.sheet))
                .map_or("<missing sheet>", |sheet| sheet.name.as_str());
            conflicts.push(Conflict::Order {
                sheet: sheet.into(),
                position,
            });
        }
        if let (Some(target), Some(_)) = (self.active, other.active) {
            let (sheet, position) = self.target_context(target);
            conflicts.push(Conflict::Active {
                sheet: sheet.into(),
                position,
            });
        }
        for (left_index, left) in self.added.iter().enumerate() {
            if other
                .added
                .iter()
                .any(|right| right.name.identity_key() == left.name.identity_key())
                || other.sheets.values().any(|actions| {
                    actions
                        .rename
                        .as_ref()
                        .is_some_and(|name| name.identity_key() == left.name.identity_key())
                })
            {
                conflicts.push(Conflict::Name {
                    sheet: left.name.as_str().into(),
                    position: self
                        .projected_position(Target::Added(left_index))
                        .unwrap_or_else(|| self.base.len().saturating_add(left_index)),
                });
            }
        }
        for (right_index, right) in other.added.iter().enumerate() {
            if self.sheets.values().any(|actions| {
                actions
                    .rename
                    .as_ref()
                    .is_some_and(|name| name.identity_key() == right.name.identity_key())
            }) {
                conflicts.push(Conflict::Name {
                    sheet: right.name.as_str().into(),
                    position: other
                        .projected_position(Target::Added(right_index))
                        .unwrap_or_else(|| other.base.len().saturating_add(right_index)),
                });
            }
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
            if left.rename.is_some() && right.rename.is_some() {
                conflicts.push(Conflict::Name {
                    sheet: sheet.into(),
                    position: *position,
                });
            }
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
                .iter()
                .filter_map(|(row, action)| {
                    right
                        .rows
                        .get(row)
                        .is_some_and(|other| action.overlaps(*other))
                        .then_some(*row)
                })
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
                .iter()
                .filter_map(|(column, action)| {
                    right
                        .columns
                        .get(column)
                        .is_some_and(|other| action.overlaps(*other))
                        .then_some(*column)
                })
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

    fn target_context(&self, target: Target) -> (&str, usize) {
        let projected = self.projected_position(target);
        match target {
            Target::Base(position) => (
                self.base
                    .inner
                    .sheets
                    .get(position)
                    .map_or("<missing sheet>", |sheet| sheet.name.as_str()),
                projected.unwrap_or(position),
            ),
            Target::Added(index) => self.added.get(index).map_or(
                ("<missing new sheet>", self.base.len().saturating_add(index)),
                |sheet| {
                    (
                        sheet.name.as_str(),
                        projected.unwrap_or_else(|| self.base.len().saturating_add(index)),
                    )
                },
            ),
        }
    }

    fn projected_position(&self, target: Target) -> Option<usize> {
        FinalOrder::plan(
            self.base.len(),
            self.order.as_ref().filter(|order| order.is_effective()),
            &self.added,
        )
        .ok()?
        .position(target)
    }

    fn has_non_removal(&self) -> bool {
        self.active.is_some()
            || self.order.as_ref().is_some_and(OrderPlan::is_effective)
            || self.sheets.values().any(|actions| !actions.is_empty())
            || !self.added.is_empty()
    }

    fn remove_block(&self, reason: RemoveBlock, part: &str) -> Error {
        let position = self.removed.iter().next().copied().unwrap_or(0);
        let sheet = self
            .base
            .inner
            .sheets
            .get(position)
            .map_or("<missing sheet>", |sheet| sheet.name.as_str());
        Error::SheetRemoveBlocked {
            sheet: sheet.to_owned(),
            position,
            part: part.to_owned(),
            reason,
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
    /// Rename this tab and every modeled local formula/reference dependency.
    ///
    /// Borrowed strings are validated and copied once. Passing a prevalidated
    /// [`crate::sheet::Name`] or owned `String` keeps the operation move-first.
    pub fn rename<T>(&mut self, name: T) -> Result<&mut Self>
    where
        T: TryInto<Name>,
        Error: From<T::Error>,
    {
        self.edit
            .set_name(self.position, name.try_into().map_err(Error::from)?);
        Ok(self)
    }

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
        let style_lineage = &self.edit.base.inner.style_lineage;
        let actions = &mut self.edit.sheets.entry(self.position).or_default().rows;
        Ok(RowEdit {
            actions,
            style_lineage,
            row,
        })
    }

    /// Select one column by its primary A1 label or a checked zero-based input.
    pub fn column<'a>(&mut self, at: impl Into<ColumnAt<'a>>) -> Result<ColumnEdit<'_>> {
        let column = at.into().resolve()?;
        let style_lineage = &self.edit.base.inner.style_lineage;
        let actions = &mut self.edit.sheets.entry(self.position).or_default().columns;
        Ok(ColumnEdit {
            actions,
            style_lineage,
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

/// Borrowed editor for one worksheet being created by this transaction.
///
/// The handle owns no native Office identity and cannot outlive its edit.
#[derive(Debug)]
pub struct NewSheet<'a> {
    added: &'a mut Added,
    active: &'a mut Option<Target>,
    style_lineage: &'a Arc<StyleLineage>,
    index: usize,
    position: usize,
}

impl NewSheet<'_> {
    /// Current checked zero-based projected position.
    ///
    /// A later tab move or insertion in this transaction can shift it. The
    /// authoritative committed position is recorded by [`Change::Create`].
    pub const fn position(&self) -> usize {
        self.position
    }

    /// Current validated developer-facing name in the pending transaction.
    pub fn name(&self) -> &str {
        self.added.name.as_str()
    }

    /// Replace the pending name without exposing a physical sheet identity.
    pub fn rename<T>(&mut self, name: T) -> Result<&mut Self>
    where
        T: TryInto<Name>,
        Error: From<T::Error>,
    {
        self.added.name = name.try_into().map_err(Error::from)?;
        Ok(self)
    }

    /// Make the new visible worksheet active at commit.
    pub fn activate(&mut self) -> &mut Self {
        *self.active = Some(Target::Added(self.index));
        self
    }

    /// Keep the new worksheet visible (the default).
    pub fn show(&mut self) -> &mut Self {
        self.added.actions.visibility = Some(TabAction::Show);
        self
    }

    /// Create the worksheet hidden from the ordinary tab strip.
    pub fn hide(&mut self) -> &mut Self {
        self.added.actions.visibility = Some(TabAction::Hide);
        self
    }

    /// Create the worksheet hidden from Excel's ordinary Unhide dialog.
    pub fn very_hide(&mut self) -> &mut Self {
        self.added.actions.visibility = Some(TabAction::VeryHide);
        self
    }

    /// Select one checked row for short property-editing verbs.
    pub fn row(&mut self, at: impl Into<RowAt>) -> Result<RowEdit<'_>> {
        let row = at.into().resolve()?;
        Ok(RowEdit {
            actions: &mut self.added.actions.rows,
            style_lineage: self.style_lineage,
            row,
        })
    }

    /// Select one column by its primary A1 label or a checked zero-based input.
    pub fn column<'b>(&mut self, at: impl Into<ColumnAt<'b>>) -> Result<ColumnEdit<'_>> {
        let column = at.into().resolve()?;
        Ok(ColumnEdit {
            actions: &mut self.added.actions.columns,
            style_lineage: self.style_lineage,
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
        match self.added.actions.cells.entry(address) {
            Entry::Vacant(entry) => {
                entry.insert(Action::set(content));
            },
            Entry::Occupied(mut entry) => entry.get_mut().set_payload(Payload::Set(content)),
        }
        Ok(self)
    }

    /// Clear pending primary content while retaining an explicit cell when a
    /// previous operation in this transaction created it.
    pub fn clear<'a>(&mut self, at: impl Into<At<'a>>) -> Result<&mut Self> {
        let address = at.into().resolve()?;
        let actions = &mut self.added.actions.cells;
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
    pub fn style<'a>(&mut self, at: impl Into<At<'a>>, style: &Style) -> Result<&mut Self> {
        if !Arc::ptr_eq(self.style_lineage, &style.owner.style_lineage) {
            return Err(Error::ForeignStyle);
        }
        let address = at.into().resolve()?;
        let effect = StyleEffect::Set(style.raw());
        match self.added.actions.cells.entry(address) {
            Entry::Vacant(entry) => {
                entry.insert(Action::style(style.raw()));
            },
            Entry::Occupied(mut entry) => entry.get_mut().set_style(effect),
        }
        Ok(self)
    }

    /// Remove an explicit local style reference from the pending cell.
    pub fn reset_style<'a>(&mut self, at: impl Into<At<'a>>) -> Result<&mut Self> {
        let address = at.into().resolve()?;
        match self.added.actions.cells.entry(address) {
            Entry::Vacant(entry) => {
                entry.insert(Action::reset_style());
            },
            Entry::Occupied(mut entry) => entry.get_mut().set_style(StyleEffect::Reset),
        }
        Ok(self)
    }

    /// Remove the pending cell record without shifting surrounding cells.
    pub fn remove<'a>(&mut self, at: impl Into<At<'a>>) -> Result<&mut Self> {
        let address = at.into().resolve()?;
        self.added.actions.cells.insert(address, Action::Remove);
        Ok(self)
    }
}

/// Transaction-scoped editor for one checked worksheet row.
#[derive(Debug)]
pub struct RowEdit<'a> {
    actions: &'a mut BTreeMap<RowIndex, RowAction>,
    style_lineage: &'a Arc<StyleLineage>,
    row: RowIndex,
}

impl RowEdit<'_> {
    fn action(&mut self) -> &mut RowAction {
        self.actions.entry(self.row).or_default()
    }

    /// Hide this row while preserving all row and cell content.
    pub fn hide(&mut self) -> &mut Self {
        self.action().hidden = Some(true);
        self
    }

    /// Show this row while preserving all row and cell content.
    pub fn show(&mut self) -> &mut Self {
        self.action().hidden = Some(false);
        self
    }

    /// Set a checked row height in points and mark it as explicitly customized.
    ///
    /// Raw finite values in `0..=409` and reusable [`crate::row::Height`]
    /// values are both accepted.
    pub fn height(&mut self, height: impl Into<HeightAt>) -> Result<&mut Self> {
        let height = height.into().resolve()?;
        self.action().height = Some(HeightEffect::Set(height));
        Ok(self)
    }

    /// Remove the explicit height and its derived custom-height marker.
    pub fn reset_height(&mut self) -> &mut Self {
        self.action().height = Some(HeightEffect::Reset);
        self
    }

    /// Apply an existing shared style as this row's default formatting.
    ///
    /// The handle must belong to the transaction's shared-style lineage.
    /// Cells with an explicit local style continue to take precedence.
    pub fn style(&mut self, style: &Style) -> Result<&mut Self> {
        if !Arc::ptr_eq(self.style_lineage, &style.owner.style_lineage) {
            return Err(Error::ForeignStyle);
        }
        self.action().style = Some(StyleEffect::Set(style.raw()));
        Ok(self)
    }

    /// Remove the row's explicit default style and custom-format marker.
    pub fn reset_style(&mut self) -> &mut Self {
        self.action().style = Some(StyleEffect::Reset);
        self
    }

    /// Set a checked outline level in Office's `0..=7` domain.
    pub fn outline(&mut self, level: impl Into<OutlineAt>) -> Result<&mut Self> {
        let level = level.into().resolve()?;
        self.action().outline = Some(level);
        Ok(self)
    }

    /// Store the affected outline in its collapsed state.
    pub fn collapse(&mut self) -> &mut Self {
        self.action().collapsed = Some(true);
        self
    }

    /// Store the affected outline in its expanded state.
    pub fn expand(&mut self) -> &mut Self {
        self.action().collapsed = Some(false);
        self
    }

    /// Request a thick top edge for this row.
    pub fn thick_top(&mut self) -> &mut Self {
        self.action().thick_top = Some(true);
        self
    }

    /// Restore the normal top edge for this row.
    pub fn normal_top(&mut self) -> &mut Self {
        self.action().thick_top = Some(false);
        self
    }

    /// Request a thick bottom edge for this row.
    pub fn thick_bottom(&mut self) -> &mut Self {
        self.action().thick_bottom = Some(true);
        self
    }

    /// Restore the normal bottom edge for this row.
    pub fn normal_bottom(&mut self) -> &mut Self {
        self.action().thick_bottom = Some(false);
        self
    }

    /// Show phonetic information by default for this row.
    pub fn show_phonetic(&mut self) -> &mut Self {
        self.action().phonetic = Some(true);
        self
    }

    /// Hide phonetic information by default for this row.
    pub fn hide_phonetic(&mut self) -> &mut Self {
        self.action().phonetic = Some(false);
        self
    }
}

/// Transaction-scoped editor for one checked worksheet column.
#[derive(Debug)]
pub struct ColumnEdit<'a> {
    actions: &'a mut BTreeMap<ColumnIndex, ColumnAction>,
    style_lineage: &'a Arc<StyleLineage>,
    column: ColumnIndex,
}

impl ColumnEdit<'_> {
    fn action(&mut self) -> &mut ColumnAction {
        self.actions.entry(self.column).or_default()
    }

    /// Hide this column while preserving its other effective properties and
    /// every cell record.
    pub fn hide(&mut self) -> &mut Self {
        self.action().hidden = Some(true);
        self
    }

    /// Show this column while preserving its other effective properties and
    /// every cell record.
    pub fn show(&mut self) -> &mut Self {
        self.action().hidden = Some(false);
        self
    }

    /// Set a checked SpreadsheetML width and mark it as explicitly customized.
    ///
    /// Raw finite values in `0..=255` and reusable [`crate::column::Width`]
    /// values are both accepted.
    pub fn width(&mut self, width: impl Into<WidthAt>) -> Result<&mut Self> {
        let width = width.into().resolve()?;
        self.action().width = Some(WidthEffect::Set(width));
        Ok(self)
    }

    /// Remove the explicit width and its derived custom-width marker.
    pub fn reset_width(&mut self) -> &mut Self {
        self.action().width = Some(WidthEffect::Reset);
        self
    }

    /// Apply an existing shared style as this column's default formatting.
    ///
    /// The handle must belong to the transaction's shared-style lineage.
    /// Cells with an explicit local style continue to take precedence.
    /// When the column has no stored width, stage [`Self::width`] in the same
    /// transaction; commit otherwise rejects the style instead of producing
    /// a zero-width column in Excel.
    pub fn style(&mut self, style: &Style) -> Result<&mut Self> {
        if !Arc::ptr_eq(self.style_lineage, &style.owner.style_lineage) {
            return Err(Error::ForeignStyle);
        }
        self.action().style = Some(StyleEffect::Set(style.raw()));
        Ok(self)
    }

    /// Remove the column's explicit default style without changing its width.
    pub fn reset_style(&mut self) -> &mut Self {
        self.action().style = Some(StyleEffect::Reset);
        self
    }

    /// Mark this column for producer best-fit behavior without measuring text.
    pub fn best_fit(&mut self) -> &mut Self {
        self.action().best_fit = Some(true);
        self
    }

    /// Clear the producer best-fit marker while preserving the stored width.
    pub fn fixed(&mut self) -> &mut Self {
        self.action().best_fit = Some(false);
        self
    }

    /// Set a checked outline level in Office's `0..=7` domain.
    pub fn outline(&mut self, level: impl Into<OutlineAt>) -> Result<&mut Self> {
        let level = level.into().resolve()?;
        self.action().outline = Some(level);
        Ok(self)
    }

    /// Store the affected outline in its collapsed state.
    pub fn collapse(&mut self) -> &mut Self {
        self.action().collapsed = Some(true);
        self
    }

    /// Store the affected outline in its expanded state.
    pub fn expand(&mut self) -> &mut Self {
        self.action().collapsed = Some(false);
        self
    }

    /// Show phonetic information by default for this column.
    pub fn show_phonetic(&mut self) -> &mut Self {
        self.action().phonetic = Some(true);
        self
    }

    /// Hide phonetic information by default for this column.
    pub fn hide_phonetic(&mut self) -> &mut Self {
        self.action().phonetic = Some(false);
        self
    }
}

fn commit_removals(edit: Edit) -> Result<Commit> {
    if edit.has_non_removal() {
        return Err(edit.remove_block(RemoveBlock::MixedEdit, "transaction"));
    }
    let Edit {
        base,
        active: _,
        order: _,
        sheets: _,
        added: _,
        removed,
    } = edit;
    let first_position = removed
        .iter()
        .next()
        .copied()
        .ok_or_else(|| invalid("worksheet removal plan is empty"))?;
    let first_sheet = base
        .inner
        .sheets
        .get(first_position)
        .ok_or_else(|| invalid("removed worksheet position disappeared"))?;
    let block = |reason, part: &str| Error::SheetRemoveBlocked {
        sheet: first_sheet.name.clone(),
        position: first_position,
        part: part.to_owned(),
        reason,
    };
    let retained_len = base
        .inner
        .sheets
        .len()
        .checked_sub(removed.len())
        .ok_or_else(|| invalid("removed worksheet count exceeds the catalog"))?;
    if retained_len == 0 {
        return Err(block(
            RemoveBlock::LastSheet,
            base.inner.workbook_uri.as_str(),
        ));
    }
    ensure_reorder_supported(&base, &first_sheet.name, first_position)?;

    for position in &removed {
        let sheet = base
            .inner
            .sheets
            .get(*position)
            .ok_or_else(|| invalid("removed worksheet position disappeared"))?;
        if sheet.kind != SheetKind::Worksheet {
            return Err(Error::NotWorksheet {
                sheet: sheet.name.clone(),
            });
        }
    }

    let main = base.inner.package.get_part(&base.inner.workbook_uri)?;
    if let Some(relationship) = main.rels().iter().find(|relationship| {
        relationship.reltype() == litchi_opc::constants::relationship_type::VBA_PROJECT
    }) {
        return Err(block(RemoveBlock::MacroProject, relationship.target_ref()));
    }

    let visible = base
        .inner
        .sheets
        .iter()
        .enumerate()
        .filter(|(position, sheet)| {
            !removed.contains(position) && sheet.visibility == Visibility::Visible
        })
        .map(|(position, _)| position)
        .collect::<Vec<_>>();
    if visible.is_empty() {
        return Err(Error::TabEditBlocked {
            sheet: first_sheet.name.clone(),
            position: first_position,
            reason: TabEditBlock::LastVisibleTab,
        });
    }
    let current_active = base
        .inner
        .active_sheet
        .ok_or_else(|| invalid("non-empty workbook has no active tab"))?;
    let final_active_identity = if !removed.contains(&current_active)
        && base
            .inner
            .sheets
            .get(current_active)
            .is_some_and(|sheet| sheet.visibility == Visibility::Visible)
    {
        current_active
    } else {
        visible
            .iter()
            .copied()
            .find(|position| *position > current_active)
            .or_else(|| {
                visible
                    .iter()
                    .rev()
                    .copied()
                    .find(|position| *position < current_active)
            })
            .ok_or_else(|| invalid("replacement active worksheet disappeared"))?
    };
    let final_active_position = (0..final_active_identity)
        .filter(|position| !removed.contains(position))
        .count();
    if final_active_position > raw::catalog_edit::MAX_ACTIVE_TAB {
        let sheet = base
            .inner
            .sheets
            .get(final_active_identity)
            .ok_or_else(|| invalid("replacement active worksheet disappeared"))?;
        return Err(Error::TabEditBlocked {
            sheet: sheet.name.clone(),
            position: final_active_position,
            reason: TabEditBlock::ActiveTabLimit,
        });
    }
    let active_sheet = base
        .inner
        .sheets
        .get(final_active_identity)
        .ok_or_else(|| invalid("replacement active worksheet disappeared"))?;

    let removed_relationship_ids = removed
        .iter()
        .map(|position| {
            base.inner
                .sheets
                .get(*position)
                .map(|sheet| sheet.relationship_id.as_str())
                .ok_or_else(|| invalid("removed worksheet relationship disappeared"))
        })
        .collect::<Result<Vec<_>>>()?;
    let local_scopes = base
        .inner
        .defined_names
        .iter()
        .filter(|name| name.local_sheet_id.is_some())
        .count();
    let before_workbook = main.blob_arc();
    let mut after_workbook = raw::catalog_edit::remove(
        &before_workbook,
        raw::catalog_edit::Remove {
            sheet: &first_sheet.name,
            position: first_position,
            relationship_ids: removed_relationship_ids.clone(),
            active: raw::catalog_edit::Active {
                sheet: &active_sheet.name,
                position: final_active_position,
            },
            local_scopes,
        },
    )?;
    after_workbook = raw::recalc::invalidate(&after_workbook)?;
    let catalog = raw::parse_catalog(&after_workbook)?;
    if catalog.sheets.len() != retained_len || catalog.active_sheet_index != final_active_position {
        return Err(invalid("workbook worksheet-removal verification failed"));
    }
    let retained = base
        .inner
        .sheets
        .iter()
        .enumerate()
        .filter(|(position, _)| !removed.contains(position));
    for (actual, (_, expected)) in catalog.sheets.iter().zip(retained) {
        if actual.relationship_id != expected.relationship_id || actual.name != expected.name {
            return Err(invalid(
                "workbook worksheet-removal verification changed a retained tab",
            ));
        }
    }
    verify_removed_defined_names(&base, &catalog, &removed)?;

    let mut changes = Vec::new();
    changes
        .try_reserve(removed.len().saturating_add(1))
        .map_err(|error| invalid(format!("cannot reserve worksheet-removal changes: {error}")))?;
    for position in &removed {
        let sheet = base
            .inner
            .sheets
            .get(*position)
            .ok_or_else(|| invalid("removed worksheet disappeared during patch creation"))?;
        changes.push(Change::Remove {
            sheet: sheet.name.clone().into_boxed_str(),
            position: *position,
            visibility: sheet.visibility.clone(),
        });
    }
    let active_before = active_tab_at(&base, current_active, current_active, None)?;
    let active_after = active_tab_at(&base, final_active_identity, final_active_position, None)?;
    if active_before != active_after {
        changes.push(Change::Active {
            before: active_before,
            after: active_after,
        });
    }

    let mut parts = vec![PartChange {
        uri: base.inner.workbook_uri.clone(),
        before: before_workbook,
        after: Arc::new(after_workbook),
    }];
    if final_active_identity != current_active {
        if !removed.contains(&current_active) {
            let old_active = base
                .inner
                .sheets
                .get(current_active)
                .ok_or_else(|| invalid("previous active worksheet disappeared"))?;
            compose_part(&mut parts, &base, &old_active.part_uri, |content| {
                raw::sheet_view_edit::rewrite(
                    content,
                    false,
                    raw::sheet_view_edit::Context {
                        sheet: &old_active.name,
                        position: current_active,
                    },
                )
            })?;
        }
        compose_part(&mut parts, &base, &active_sheet.part_uri, |content| {
            raw::sheet_view_edit::rewrite(
                content,
                true,
                raw::sheet_view_edit::Context {
                    sheet: &active_sheet.name,
                    position: final_active_position,
                },
            )
        })?;
    }

    let existing_titles = base
        .inner
        .sheets
        .iter()
        .map(|sheet| sheet.name.as_str())
        .collect::<Vec<_>>();
    let removed_positions = removed.iter().copied().collect::<Vec<_>>();
    let mut property_parts = base
        .inner
        .package
        .iter_parts()
        .filter(|part| {
            part.content_type() == litchi_opc::constants::content_type::OFC_EXTENDED_PROPERTIES
        })
        .map(|part| part.partname().clone())
        .collect::<Vec<_>>();
    property_parts.sort_unstable_by(|left, right| left.as_str().cmp(right.as_str()));
    for uri in property_parts {
        compose_part_optional(&mut parts, &base, &uri, |content| {
            raw::properties_edit::remove_sheets(content, &existing_titles, &removed_positions)
        })?;
    }

    let mut graph = Vec::new();
    graph
        .try_reserve(removed.len().saturating_add(1))
        .map_err(|error| invalid(format!("cannot reserve worksheet graph removals: {error}")))?;
    for position in &removed {
        let sheet = base
            .inner
            .sheets
            .get(*position)
            .ok_or_else(|| invalid("removed worksheet disappeared during graph planning"))?;
        let relationship = main.rels().get(&sheet.relationship_id).ok_or_else(|| {
            invalid(format!(
                "worksheet '{}' relationship disappeared",
                sheet.name
            ))
        })?;
        if relationship.is_external()
            || !relationship
                .target_partname()?
                .is_equivalent_to(&sheet.part_uri)
        {
            return Err(invalid(format!(
                "worksheet '{}' relationship target changed",
                sheet.name
            )));
        }
        ensure_exclusive_sheet_incoming(
            &base.inner.package,
            &sheet.part_uri,
            &base.inner.workbook_uri,
            relationship.r_id(),
            &sheet.name,
            *position,
        )?;
        graph.push(GraphChange {
            action: GraphAction::Remove,
            source: base.inner.workbook_uri.clone(),
            relationship: relationship.clone(),
            part: base.inner.package.get_part(&sheet.part_uri)?.clone_part(),
        });
    }
    graph.extend(calculation_chain_removal(&base)?);

    let detached_workbook_relationships = graph
        .iter()
        .filter(|change| change.source == base.inner.workbook_uri)
        .map(|change| change.relationship.r_id())
        .collect::<Vec<_>>();
    scan_removal_dependencies(&base, &parts, &removed, &detached_workbook_relationships)?;

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
            style_guard: None,
        },
    })
}

fn verify_removed_defined_names(
    workbook: &Workbook,
    catalog: &raw::Catalog,
    removed: &BTreeSet<usize>,
) -> Result<()> {
    let mut expected = Vec::new();
    expected
        .try_reserve_exact(workbook.inner.defined_names.len())
        .map_err(|error| invalid(format!("cannot reserve defined-name verification: {error}")))?;
    for name in &workbook.inner.defined_names {
        let scope = name
            .local_sheet_id
            .map(|scope| {
                usize::try_from(scope).map_err(|_| invalid("defined-name scope does not fit usize"))
            })
            .transpose()?;
        if scope.is_some_and(|scope| removed.contains(&scope)) {
            continue;
        }
        let mapped = scope.map(|scope| {
            u32::try_from(
                (0..scope)
                    .filter(|position| !removed.contains(position))
                    .count(),
            )
            .map_err(|_| invalid("remapped defined-name scope does not fit u32"))
        });
        expected.push((name, mapped.transpose()?));
    }
    if expected.len() != catalog.defined_names.len() {
        return Err(invalid(
            "workbook removal changed the effective defined-name count unexpectedly",
        ));
    }
    for ((before, scope), after) in expected.into_iter().zip(&catalog.defined_names) {
        if after.local_sheet_id != scope || !same_defined_name_except_scope(before, after) {
            return Err(invalid(format!(
                "workbook removal changed defined name '{}' unexpectedly",
                before.name
            )));
        }
    }
    Ok(())
}

fn ensure_exclusive_sheet_incoming(
    package: &OpcPackage,
    target: &PackURI,
    expected_source: &PackURI,
    expected_id: &str,
    sheet: &str,
    position: usize,
) -> Result<()> {
    let blocked = |part: &str| Error::SheetRemoveBlocked {
        sheet: sheet.to_owned(),
        position,
        part: part.to_owned(),
        reason: RemoveBlock::IncomingRelationship,
    };
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
            return Err(blocked("/"));
        }
    }
    for source in package.iter_parts() {
        for relationship in source.rels().iter() {
            if targets(relationship)?
                && !(source.partname() == expected_source && relationship.r_id() == expected_id)
            {
                return Err(blocked(source.partname().as_str()));
            }
        }
    }
    Ok(())
}

fn scan_removal_dependencies(
    workbook: &Workbook,
    parts: &[PartChange],
    removed: &BTreeSet<usize>,
    removed_relationship_ids: &[&str],
) -> Result<()> {
    let catalog = workbook
        .inner
        .sheets
        .iter()
        .map(|sheet| sheet.name.as_str())
        .collect::<Vec<_>>();
    let targets = removed
        .iter()
        .map(|position| {
            let sheet =
                workbook.inner.sheets.get(*position).ok_or_else(|| {
                    invalid("removed worksheet disappeared during dependency scan")
                })?;
            Ok(raw::reference_scan::Sheet {
                name: &sheet.name,
                position: *position,
                native_id: sheet.native_id,
                catalog: &catalog,
            })
        })
        .collect::<Result<Vec<_>>>()?;
    let reachable = reachable_after_removal(workbook, removed_relationship_ids)?;
    for uri in &reachable {
        let part = workbook.inner.package.get_part(uri)?;
        if !removal_reference_part(part) {
            continue;
        }
        let content = parts
            .iter()
            .find(|change| &change.uri == uri)
            .map_or(part.blob(), |change| change.after.as_slice());
        let Some(hit) = raw::reference_scan::scan(content, &targets)? else {
            continue;
        };
        let sheet = targets
            .get(hit.target)
            .ok_or_else(|| invalid("dependency scan returned an unknown removal target"))?;
        let reason = match hit.dependency {
            raw::reference_scan::Dependency::Modeled => RemoveBlock::IncomingReference,
            raw::reference_scan::Dependency::Unmodeled => RemoveBlock::UnmodeledReference,
            raw::reference_scan::Dependency::MarkupCompatibility => {
                RemoveBlock::MarkupCompatibility
            },
        };
        return Err(Error::SheetRemoveBlocked {
            sheet: sheet.name.to_owned(),
            position: sheet.position,
            part: uri.to_string(),
            reason,
        });
    }
    Ok(())
}

fn reachable_after_removal(
    workbook: &Workbook,
    removed_relationship_ids: &[&str],
) -> Result<Vec<PackURI>> {
    let mut reachable = HashSet::<PackURI>::new();
    reachable
        .try_reserve(workbook.inner.package.part_count())
        .map_err(|error| invalid(format!("cannot reserve reachable package graph: {error}")))?;
    let mut pending = Vec::<PackURI>::new();
    for relationship in workbook.inner.package.rels().iter() {
        if !relationship.is_external() {
            let target = relationship.target_partname()?;
            let part = workbook.inner.package.get_part(&target)?;
            pending.push(part.partname().clone());
        }
    }
    while let Some(uri) = pending.pop() {
        if !reachable.insert(uri.clone()) {
            continue;
        }
        let part = workbook.inner.package.get_part(&uri)?;
        for relationship in part.rels().iter() {
            if uri == workbook.inner.workbook_uri
                && removed_relationship_ids.contains(&relationship.r_id())
            {
                continue;
            }
            if relationship.is_external() {
                continue;
            }
            let target = relationship.target_partname()?;
            let target = workbook.inner.package.get_part(&target)?.partname().clone();
            if !reachable.contains(&target) {
                pending.push(target);
            }
        }
    }
    let mut reachable = reachable.into_iter().collect::<Vec<_>>();
    reachable.sort_unstable_by(|left, right| left.as_str().cmp(right.as_str()));
    Ok(reachable)
}

fn raw_visibility_matches(value: &raw::Visibility, action: TabAction) -> bool {
    matches!(
        (value, action),
        (raw::Visibility::Visible, TabAction::Show)
            | (raw::Visibility::Hidden, TabAction::Hide)
            | (raw::Visibility::VeryHidden, TabAction::VeryHide)
    )
}

fn active_tab_at(
    workbook: &Workbook,
    identity: usize,
    position: usize,
    name: Option<&str>,
) -> Result<ActiveTab> {
    let sheet = workbook
        .inner
        .sheets
        .get(identity)
        .ok_or_else(|| invalid("active tab points outside the workbook catalog"))?;
    Ok(ActiveTab {
        name: name.unwrap_or(&sheet.name).into(),
        position,
    })
}

fn validate_order_plan(order: &OrderPlan, len: usize) -> Result<()> {
    if order.positions.len() != len {
        return Err(invalid("tab-order plan has the wrong number of sheets"));
    }
    let mut seen = Vec::new();
    seen.try_reserve_exact(len)
        .map_err(|error| invalid(format!("cannot reserve tab-order validation: {error}")))?;
    seen.resize(len, false);
    for identity in &order.positions {
        let Some(slot) = seen.get_mut(*identity) else {
            return Err(invalid("tab-order plan contains an out-of-range identity"));
        };
        if *slot {
            return Err(invalid("tab-order plan is not a permutation"));
        }
        *slot = true;
    }

    let mut replay = Vec::new();
    replay
        .try_reserve_exact(len)
        .map_err(|error| invalid(format!("cannot reserve tab-move replay: {error}")))?;
    replay.extend(0..len);
    for moved in &order.moves {
        if moved.from >= replay.len() || moved.to >= replay.len() {
            return Err(invalid("tab move contains an out-of-range position"));
        }
        if replay[moved.from] != moved.sheet {
            return Err(invalid("tab move source does not match the pending order"));
        }
        let identity = replay.remove(moved.from);
        replay.insert(moved.to, identity);
    }
    if replay != order.positions {
        return Err(invalid("tab moves do not produce the pending final order"));
    }
    Ok(())
}

fn ensure_reorder_supported(workbook: &Workbook, sheet: &str, position: usize) -> Result<()> {
    let main = workbook
        .inner
        .package
        .get_part(&workbook.inner.workbook_uri)?;
    if main
        .rels()
        .iter()
        .any(|relationship| relationship.reltype().ends_with("/revisionHeaders"))
    {
        return Err(Error::TabEditBlocked {
            sheet: sheet.to_owned(),
            position,
            reason: TabEditBlock::TrackedWorkbook,
        });
    }
    Ok(())
}

fn verify_defined_name_scopes(
    source: &raw::Catalog,
    catalog: &raw::Catalog,
    base_len: usize,
    order: &[Target],
) -> Result<()> {
    if catalog.defined_names.len() != source.defined_names.len() {
        return Err(invalid(
            "workbook reorder changed the effective defined-name count",
        ));
    }
    let mut old_to_new = Vec::new();
    old_to_new
        .try_reserve_exact(base_len)
        .map_err(|error| invalid(format!("cannot reserve defined-name scope map: {error}")))?;
    old_to_new.resize(base_len, usize::MAX);
    for (new, target) in order.iter().copied().enumerate() {
        let Target::Base(old) = target else {
            continue;
        };
        let slot = old_to_new
            .get_mut(old)
            .ok_or_else(|| invalid("defined-name scope map has an invalid sheet identity"))?;
        if *slot != usize::MAX {
            return Err(invalid("defined-name scope map repeats a sheet identity"));
        }
        *slot = new;
    }
    if old_to_new.contains(&usize::MAX) {
        return Err(invalid("defined-name scope map omits a sheet identity"));
    }
    for (before, after) in source.defined_names.iter().zip(&catalog.defined_names) {
        let expected_scope = match before.local_sheet_id {
            None => None,
            Some(scope) => {
                let scope = usize::try_from(scope)
                    .map_err(|_| invalid("defined-name scope does not fit usize"))?;
                let mapped = old_to_new
                    .get(scope)
                    .copied()
                    .ok_or_else(|| invalid("defined-name scope cannot be remapped"))?;
                Some(
                    u32::try_from(mapped)
                        .map_err(|_| invalid("remapped defined-name scope does not fit u32"))?,
                )
            },
        };
        if after.local_sheet_id != expected_scope || !same_defined_name_except_scope(before, after)
        {
            return Err(invalid(format!(
                "workbook reorder changed defined name '{}' unexpectedly",
                before.name
            )));
        }
    }
    Ok(())
}

fn same_defined_name_except_scope(left: &raw::DefinedName, right: &raw::DefinedName) -> bool {
    left.name == right.name
        && left.reference == right.reference
        && left.comment == right.comment
        && left.custom_menu == right.custom_menu
        && left.description == right.description
        && left.help == right.help
        && left.status_bar == right.status_bar
        && left.shortcut_key == right.shortcut_key
        && left.hidden == right.hidden
        && left.function == right.function
        && left.vb_procedure == right.vb_procedure
        && left.xlm == right.xlm
        && left.function_group_id == right.function_group_id
        && left.publish_to_server == right.publish_to_server
        && left.workbook_parameter == right.workbook_parameter
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

fn compose_part_optional(
    parts: &mut Vec<PartChange>,
    workbook: &Workbook,
    uri: &PackURI,
    rewrite: impl FnOnce(&[u8]) -> Result<Option<Vec<u8>>>,
) -> Result<()> {
    if let Some(part) = parts.iter_mut().find(|part| &part.uri == uri) {
        if let Some(after) = rewrite(&part.after)?
            && after.as_slice() != part.after.as_slice()
        {
            part.after = Arc::new(after);
        }
        return Ok(());
    }
    let before = workbook.inner.package.get_part(uri)?.blob_arc();
    let Some(after) = rewrite(&before)? else {
        return Ok(());
    };
    if after.as_slice() != before.as_slice() {
        parts.push(PartChange {
            uri: uri.clone(),
            before,
            after: Arc::new(after),
        });
    }
    Ok(())
}

fn reference_part(part: &dyn Part) -> bool {
    let uri = part.partname().as_str();
    if uri.starts_with("/xl/externalLinks/")
        || part.content_type() == litchi_opc::constants::content_type::SML_EXTERNAL_LINK
    {
        return false;
    }
    (uri.starts_with("/xl/")
        && (part.content_type().ends_with("+xml")
            || part.content_type().ends_with("/xml")
            || part.content_type() == litchi_opc::constants::content_type::OFC_VML_DRAWING))
        || part.content_type() == litchi_opc::constants::content_type::OFC_EXTENDED_PROPERTIES
}

fn removal_reference_part(part: &dyn Part) -> bool {
    let uri = part.partname().as_str();
    if uri.starts_with("/xl/externalLinks/")
        || part.content_type() == litchi_opc::constants::content_type::SML_EXTERNAL_LINK
    {
        return false;
    }
    part.content_type().ends_with("+xml")
        || part.content_type().ends_with("/xml")
        || part.content_type() == litchi_opc::constants::content_type::OFC_VML_DRAWING
}

fn ensure_unsigned(workbook: &Workbook) -> Result<()> {
    if workbook.inner.package.has_digital_signatures() {
        Err(Error::Signed)
    } else {
        Ok(())
    }
}

fn create_sheets(
    workbook: &Workbook,
    added: Vec<Added>,
    positions: &[usize],
    active: Option<usize>,
    changes: &mut Vec<Change>,
    needs_recalculation: &mut bool,
) -> Result<Vec<CreatedSheet>> {
    if added.is_empty() {
        return Ok(Vec::new());
    }
    if positions.len() != added.len() {
        return Err(invalid(
            "created worksheet positions do not match the creation plan",
        ));
    }
    let main = workbook
        .inner
        .package
        .get_part(&workbook.inner.workbook_uri)?;
    let dialect = raw::catalog_edit::dialect(main.blob())?;
    let relationship_type = match dialect {
        raw::catalog_edit::Dialect::Transitional => {
            litchi_opc::constants::relationship_type::WORKSHEET
        },
        raw::catalog_edit::Dialect::Strict => {
            litchi_opc::constants::relationship_type::STRICT_WORKSHEET
        },
    };
    let namespace = dialect.worksheet_namespace();

    let mut used_sheet_ids = HashSet::new();
    used_sheet_ids
        .try_reserve(workbook.inner.sheets.len().saturating_add(added.len()))
        .map_err(|error| invalid(format!("cannot reserve native sheet-ID index: {error}")))?;
    used_sheet_ids.extend(workbook.inner.sheets.iter().map(|sheet| sheet.native_id));

    let mut used_relationship_ids = HashSet::<String>::new();
    used_relationship_ids
        .try_reserve(main.rels().len().saturating_add(added.len()))
        .map_err(|error| invalid(format!("cannot reserve relationship-ID index: {error}")))?;
    used_relationship_ids.extend(
        main.rels()
            .iter()
            .map(|relationship| relationship.r_id().to_owned()),
    );

    let mut reserved_parts = Vec::<PackURI>::new();
    reserved_parts
        .try_reserve_exact(added.len())
        .map_err(|error| invalid(format!("cannot reserve worksheet part names: {error}")))?;
    let mut created = Vec::new();
    created
        .try_reserve_exact(added.len())
        .map_err(|error| invalid(format!("cannot reserve worksheet graph changes: {error}")))?;

    let mut next_sheet_id = 1u32;
    let mut next_relationship_id = 1u32;
    let mut next_part = 1u32;
    for (index, added) in added.into_iter().enumerate() {
        while used_sheet_ids.contains(&next_sheet_id) {
            next_sheet_id = next_sheet_id
                .checked_add(1)
                .ok_or_else(|| invalid("native worksheet ID space is exhausted"))?;
        }
        if next_sheet_id > raw::catalog_edit::MAX_SHEET_ID {
            return Err(invalid("native worksheet ID space is exhausted"));
        }
        let sheet_id = next_sheet_id;
        used_sheet_ids.insert(sheet_id);
        next_sheet_id = next_sheet_id.saturating_add(1);

        let relationship_id = loop {
            let candidate = format!("rId{next_relationship_id}");
            next_relationship_id = next_relationship_id
                .checked_add(1)
                .ok_or_else(|| invalid("workbook relationship-ID space is exhausted"))?;
            if used_relationship_ids.insert(candidate.clone()) {
                break candidate;
            }
        };

        let part_uri = loop {
            let base_uri = workbook.inner.workbook_uri.base_uri();
            let candidate_path = if base_uri == "/" {
                format!("/worksheets/sheet{next_part}.xml")
            } else {
                format!("{base_uri}/worksheets/sheet{next_part}.xml")
            };
            let candidate = PackURI::new(candidate_path).map_err(invalid)?;
            next_part = next_part
                .checked_add(1)
                .ok_or_else(|| invalid("worksheet part-name space is exhausted"))?;
            if workbook
                .inner
                .package
                .validate_new_part_name(&candidate)
                .is_ok()
                && !reserved_parts
                    .iter()
                    .any(|reserved| reserved.is_equivalent_to(&candidate))
            {
                reserved_parts.push(candidate.clone());
                break candidate;
            }
        };

        let position = positions
            .get(index)
            .copied()
            .ok_or_else(|| invalid("created worksheet has no checked position"))?;
        let Added {
            name,
            actions,
            placement: _,
        } = added;
        let visibility = actions.visibility.unwrap_or(TabAction::Show);
        changes.push(Change::Create {
            sheet: name.as_str().into(),
            position,
            visibility: visibility.visibility(),
        });

        let SheetActions {
            rename: _,
            visibility: _,
            cells,
            rows,
            columns,
        } = actions;
        let change_start = changes.len();
        let mut effective_cells = BTreeMap::new();
        for (address, action) in cells {
            let before = State::Missing;
            let after = State::after(None, &action, workbook);
            if before == after {
                continue;
            }
            *needs_recalculation |=
                State::calculation_content(&before) != State::calculation_content(&after);
            effective_cells.insert(address, action);
            changes.push(Change::Cell {
                sheet: name.as_str().into(),
                address,
                before,
                after,
            });
        }
        let mut effective_rows = BTreeMap::new();
        for (row, action) in rows {
            let before = RowState::Missing;
            let after = RowState::after(None, action, workbook);
            if before == after {
                continue;
            }
            effective_rows.insert(row, action);
            changes.push(Change::Row {
                sheet: name.as_str().into(),
                row,
                before,
                after,
            });
        }
        let mut effective_columns = BTreeMap::new();
        for (column, action) in columns {
            let before = ColumnState::Missing;
            let after = ColumnState::after(None, action, workbook);
            if before == after {
                continue;
            }
            effective_columns.insert(column, action);
            changes.push(Change::Column {
                sheet: name.as_str().into(),
                column,
                before,
                after,
            });
        }

        let template = format!(
            concat!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
                r#"<worksheet xmlns="{}"><dimension ref="A1"/><sheetData/></worksheet>"#
            ),
            namespace
        )
        .into_bytes();
        let mut content = raw::worksheet::edit::rewrite(
            &template,
            name.as_str(),
            Plan {
                cells: effective_cells,
                rows: effective_rows,
                columns: effective_columns,
            },
        )?;
        if active == Some(index) {
            content = raw::sheet_view_edit::rewrite(
                &content,
                true,
                raw::sheet_view_edit::Context {
                    sheet: name.as_str(),
                    position,
                },
            )?;
        }
        let parsed = raw::worksheet::parse(&content, || workbook.inner.shared_strings())?;
        workbook.inner.validate_styles(&parsed)?;
        for change in &changes[change_start..] {
            match change {
                Change::Cell {
                    sheet,
                    address,
                    after,
                    ..
                } => {
                    if State::read(parsed.entry(*address), workbook) != *after {
                        return Err(invalid(format!(
                            "new worksheet verification failed at {sheet}!{address}"
                        )));
                    }
                },
                Change::Row {
                    sheet, row, after, ..
                } => {
                    if RowState::read(parsed.row_entry(*row), workbook) != *after {
                        return Err(invalid(format!(
                            "new worksheet row verification failed at {sheet}!row {}",
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
                    if ColumnState::read(parsed.column_entry(*column), workbook) != *after {
                        return Err(invalid(format!(
                            "new worksheet column verification failed at {sheet}!column {}",
                            column.get()
                        )));
                    }
                },
                Change::Create { .. }
                | Change::Remove { .. }
                | Change::Rename { .. }
                | Change::Move { .. }
                | Change::Active { .. }
                | Change::Visibility { .. } => {},
            }
        }

        let target_ref = part_uri.relative_ref(workbook.inner.workbook_uri.base_uri());
        let relationship = Relationship::new_with_mode(
            relationship_id.clone(),
            relationship_type.to_owned(),
            target_ref,
            workbook.inner.workbook_uri.base_uri().to_owned(),
            TargetMode::Internal,
        );
        let part = BlobPart::new(
            part_uri,
            litchi_opc::constants::content_type::SML_WORKSHEET.to_owned(),
            content,
        );
        created.push(CreatedSheet {
            name,
            position,
            sheet_id,
            relationship_id,
            visibility,
            graph: GraphChange {
                action: GraphAction::Add,
                source: workbook.inner.workbook_uri.clone(),
                relationship,
                part: Box::new(part),
            },
        });
    }
    Ok(created)
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
            &committed.patch().changes()[1],
            Change::Row {
                row,
                before: RowState::Missing,
                after: RowState::Stored(properties),
                ..
            } if row.get() == 1 && properties.hidden()
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
    fn row_layout_is_typed_reversible_and_facet_composable() {
        let source = Workbook::new().expect("source workbook");
        let source_bytes = source.to_bytes().expect("source bytes");
        let mut edit = source.edit().expect("edit");
        edit.sheet("Sheet1")
            .expect("sheet lookup")
            .expect("worksheet")
            .row(1)
            .expect("row 2")
            .height(30)
            .expect("checked height")
            .outline(2)
            .expect("checked outline")
            .collapse()
            .thick_top()
            .thick_bottom()
            .show_phonetic();
        let committed = edit.commit().expect("layout commit");

        assert_eq!(committed.patch().len(), 1);
        let (_, before, after) = committed.patch().changes()[0].row().expect("row change");
        assert!(matches!(before, RowState::Missing));
        let RowState::Stored(properties) = after else {
            panic!("expected stored row properties")
        };
        assert_eq!(properties.height().map(crate::row::Height::get), Some(30.0));
        assert_eq!(properties.outline().get(), 2);
        assert!(properties.custom_height());
        assert!(properties.collapsed());
        assert!(properties.thick_top());
        assert!(properties.thick_bottom());
        assert!(properties.phonetic());
        assert!(!properties.hidden());
        assert!(!properties.custom_format());
        assert!(matches!(properties.style(), StyleState::Default));

        let sheet = committed
            .workbook()
            .sheet("Sheet1")
            .expect("sheet lookup")
            .expect("worksheet");
        let row = sheet.row(1).expect("row 2");
        assert_eq!(row.height().map(crate::row::Height::get), Some(30.0));
        assert_eq!(row.outline().get(), 2);
        assert!(row.custom_height());
        assert!(row.collapsed());
        assert!(row.thick_top());
        assert!(row.thick_bottom());
        assert!(row.phonetic());
        assert!(matches!(
            sheet.row_style(1).expect("row style"),
            Some(crate::LocalStyle::Default)
        ));

        let mut reset = committed.workbook().edit().expect("reset edit");
        reset
            .sheet("Sheet1")
            .expect("lookup")
            .expect("sheet")
            .row(1)
            .expect("row 2")
            .reset_height()
            .outline(0)
            .expect("outline reset")
            .expand()
            .normal_top()
            .normal_bottom()
            .hide_phonetic();
        let reset = reset.commit().expect("reset commit");
        let reset_sheet = reset
            .workbook()
            .sheet("Sheet1")
            .expect("lookup")
            .expect("sheet");
        let reset_row = reset_sheet.row(1).expect("row 2");
        assert_eq!(reset_row.height(), None);
        assert!(!reset_row.custom_height());
        assert_eq!(reset_row.outline(), Outline::NONE);
        assert!(!reset_row.collapsed());
        assert!(!reset_row.thick_top());
        assert!(!reset_row.thick_bottom());
        assert!(!reset_row.phonetic());

        let restored = committed
            .workbook()
            .apply(&committed.patch().inverse())
            .expect("inverse");
        assert_eq!(
            restored.workbook().to_bytes().expect("restored bytes"),
            source_bytes
        );

        let mut invalid = source.edit().expect("invalid edit");
        let mut sheet = invalid.sheet(0usize).expect("lookup").expect("sheet");
        assert!(matches!(
            sheet.row(1).expect("row 2").height(f64::NAN),
            Err(Error::RowHeight(_))
        ));
        assert!(matches!(
            sheet.row(1).expect("row 2").height(409.1),
            Err(Error::RowHeight(_))
        ));
        assert!(matches!(
            sheet.row(1).expect("row 2").outline(8),
            Err(Error::Outline(_))
        ));

        let mut height = source.edit().expect("height edit");
        height
            .sheet(0usize)
            .expect("lookup")
            .expect("sheet")
            .row(2)
            .expect("row 3")
            .height(crate::row::Height::new(22.0).expect("prevalidated height"))
            .expect("height");
        let mut visibility = source.edit().expect("visibility edit");
        visibility
            .sheet(0usize)
            .expect("lookup")
            .expect("sheet")
            .row(2)
            .expect("row 3")
            .hide();
        height.join(visibility).expect("disjoint facets on one row");
        let joined = height.commit().expect("joined commit");
        let joined_sheet = joined
            .workbook()
            .sheet(0usize)
            .expect("lookup")
            .expect("sheet");
        let joined_row = joined_sheet.row(2).expect("row 3");
        assert!(joined_row.hidden());
        assert_eq!(joined_row.height().map(crate::row::Height::get), Some(22.0));

        let mut left = source.edit().expect("left height");
        left.sheet(0usize)
            .expect("lookup")
            .expect("sheet")
            .row(3)
            .expect("row 4")
            .height(10)
            .expect("height");
        let mut right = source.edit().expect("right height");
        right
            .sheet(0usize)
            .expect("lookup")
            .expect("sheet")
            .row(3)
            .expect("row 4")
            .reset_height();
        assert!(left.join(right).is_err());
    }

    #[test]
    fn row_layout_patch_guards_and_rebinds_hidden_shared_style_identity() {
        let source = styled_row_workbook();
        let source_bytes = source.to_bytes().expect("source bytes");
        let mut edit = source.edit().expect("edit");
        edit.sheet("Sheet1")
            .expect("lookup")
            .expect("sheet")
            .row(1)
            .expect("row 2")
            .height(28)
            .expect("height");
        let committed = edit.commit().expect("commit");
        let (_, _, after) = committed.patch().changes()[0].row().expect("row change");
        let RowState::Stored(properties) = after else {
            panic!("expected stored properties")
        };
        assert!(properties.custom_format());
        assert!(matches!(properties.style(), StyleState::Shared(_)));

        let sheet = committed
            .workbook()
            .sheet("Sheet1")
            .expect("lookup")
            .expect("sheet");
        assert!(matches!(
            sheet.row_style(1).expect("row style"),
            Some(crate::LocalStyle::Shared(_))
        ));

        let reopened = Workbook::from_bytes(source_bytes).expect("reopened source");
        let replayed = reopened
            .apply(committed.patch())
            .expect("source-checked replay");
        let (_, _, replayed_after) = replayed.patch().changes()[0]
            .row()
            .expect("replayed row change");
        let RowState::Stored(replayed_properties) = replayed_after else {
            panic!("expected replayed properties")
        };
        let StyleState::Shared(replayed_key) = replayed_properties.style() else {
            panic!("expected rebound shared style")
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
            source
                .styles()
                .expect("source styles")
                .find(replayed_key)
                .is_none()
        );

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
        let changed = Workbook::from_package(changed_package).expect("changed style table");
        assert!(matches!(
            changed.apply(committed.patch()),
            Err(Error::PatchConflict { part }) if part == "/xl/styles.xml"
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
                after: ColumnState::Stored(properties),
                ..
            } if column.get() == 1 && properties.hidden()
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
    fn column_layout_is_selector_first_typed_reversible_and_facet_composable() {
        let source = Workbook::new().expect("source workbook");
        let source_bytes = source.to_bytes().expect("source bytes");
        let mut edit = source.edit().expect("edit");
        edit.sheet("Sheet1")
            .expect("sheet lookup")
            .expect("worksheet")
            .column("B")
            .expect("A1 column selector")
            .width(18.5)
            .expect("checked width")
            .outline(2)
            .expect("checked outline")
            .collapse()
            .best_fit()
            .show_phonetic();
        let committed = edit.commit().expect("layout commit");

        assert_eq!(committed.patch().len(), 1);
        let (_, before, after) = committed.patch().changes()[0]
            .column()
            .expect("column change");
        assert!(matches!(before, ColumnState::Missing));
        let ColumnState::Stored(properties) = after else {
            panic!("expected stored column properties")
        };
        assert_eq!(
            properties.width().map(crate::column::Width::get),
            Some(18.5)
        );
        assert_eq!(properties.outline().get(), 2);
        assert!(properties.collapsed());
        assert!(properties.best_fit());
        assert!(properties.custom_width());
        assert!(properties.phonetic());
        assert!(!properties.hidden());
        assert!(matches!(properties.style(), StyleState::Default));

        let sheet = committed
            .workbook()
            .sheet("Sheet1")
            .expect("sheet lookup")
            .expect("worksheet");
        let column = sheet.column("b").expect("case-insensitive A1 column");
        assert_eq!(column.index().get(), 1);
        assert_eq!(column.width().map(crate::column::Width::get), Some(18.5));
        assert_eq!(column.outline().get(), 2);
        assert!(column.collapsed());
        assert!(column.best_fit());
        assert!(column.custom_width());
        assert!(column.phonetic());

        let mut reset = committed.workbook().edit().expect("reset edit");
        reset
            .sheet("Sheet1")
            .expect("lookup")
            .expect("sheet")
            .column("B")
            .expect("column B")
            .reset_width()
            .fixed()
            .outline(0)
            .expect("outline reset")
            .expand()
            .hide_phonetic();
        let reset = reset.commit().expect("reset commit");
        let reset_sheet = reset
            .workbook()
            .sheet("Sheet1")
            .expect("lookup")
            .expect("sheet");
        let reset_column = reset_sheet.column("B").expect("column B");
        assert_eq!(reset_column.width(), None);
        assert!(!reset_column.custom_width());
        assert!(!reset_column.best_fit());
        assert_eq!(reset_column.outline(), Outline::NONE);
        assert!(!reset_column.collapsed());
        assert!(!reset_column.phonetic());

        let restored = committed
            .workbook()
            .apply(&committed.patch().inverse())
            .expect("inverse");
        assert_eq!(
            restored.workbook().to_bytes().expect("restored bytes"),
            source_bytes
        );

        let mut invalid = source.edit().expect("invalid edit");
        let mut sheet = invalid.sheet(0usize).expect("lookup").expect("sheet");
        assert!(matches!(
            sheet.column("XFE"),
            Err(Error::Coordinate(
                litchi_sheet::CoordinateError::ColumnA1 { .. }
            ))
        ));
        assert!(matches!(
            sheet.column("B").expect("B").width(f64::NAN),
            Err(Error::ColumnWidth(_))
        ));
        assert!(matches!(
            sheet.column("B").expect("B").outline(8),
            Err(Error::Outline(_))
        ));

        let mut width = source.edit().expect("width edit");
        width
            .sheet(0usize)
            .expect("lookup")
            .expect("sheet")
            .column("C")
            .expect("column C")
            .width(crate::column::Width::new(22.0).expect("prevalidated width"))
            .expect("width");
        let mut visibility = source.edit().expect("visibility edit");
        visibility
            .sheet(0usize)
            .expect("lookup")
            .expect("sheet")
            .column("C")
            .expect("column C")
            .hide();
        width
            .join(visibility)
            .expect("disjoint facets on one column");
        let joined = width.commit().expect("joined commit");
        let joined_sheet = joined
            .workbook()
            .sheet(0usize)
            .expect("lookup")
            .expect("sheet");
        let column = joined_sheet.column("C").expect("column C");
        assert!(column.hidden());
        assert_eq!(column.width().map(crate::column::Width::get), Some(22.0));

        let mut left = source.edit().expect("left width");
        left.sheet(0usize)
            .expect("lookup")
            .expect("sheet")
            .column("D")
            .expect("column D")
            .width(10.0)
            .expect("width");
        let mut right = source.edit().expect("right width");
        right
            .sheet(0usize)
            .expect("lookup")
            .expect("sheet")
            .column("D")
            .expect("column D")
            .reset_width();
        assert!(left.join(right).is_err());
    }

    #[test]
    fn column_layout_patch_guards_and_rebinds_hidden_shared_style_identity() {
        let source = styled_column_workbook();
        let source_bytes = source.to_bytes().expect("source bytes");
        let mut edit = source.edit().expect("edit");
        edit.sheet("Sheet1")
            .expect("lookup")
            .expect("sheet")
            .column("C")
            .expect("column C")
            .width(30.0)
            .expect("width");
        let committed = edit.commit().expect("commit");
        let (_, _, after) = committed.patch().changes()[0]
            .column()
            .expect("column change");
        let ColumnState::Stored(properties) = after else {
            panic!("expected stored properties")
        };
        assert!(matches!(properties.style(), StyleState::Shared(_)));

        let reopened = Workbook::from_bytes(source_bytes).expect("reopened source");
        let replayed = reopened
            .apply(committed.patch())
            .expect("source-checked replay");
        let (_, _, replayed_after) = replayed.patch().changes()[0]
            .column()
            .expect("replayed column change");
        let ColumnState::Stored(replayed_properties) = replayed_after else {
            panic!("expected replayed properties")
        };
        let StyleState::Shared(replayed_key) = replayed_properties.style() else {
            panic!("expected rebound shared style")
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
            source
                .styles()
                .expect("source styles")
                .find(replayed_key)
                .is_none()
        );

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
        let changed = Workbook::from_package(changed_package).expect("changed style table");
        assert!(matches!(
            changed.apply(committed.patch()),
            Err(Error::PatchConflict { part }) if part == "/xl/styles.xml"
        ));
    }

    #[test]
    fn grid_default_styles_are_lineage_checked_reversible_and_facet_composable() {
        let source = styled_workbook();
        let source_bytes = source.to_bytes().expect("source bytes");
        let accent = source
            .sheet("Sheet1")
            .expect("sheet lookup")
            .expect("worksheet")
            .style("A1")
            .expect("style lookup")
            .expect("accent style");

        let mut edit = source.edit().expect("edit");
        {
            let mut sheet = edit.sheet("Sheet1").expect("lookup").expect("sheet");
            sheet
                .row(1)
                .expect("row 2")
                .style(&accent)
                .expect("row style")
                .height(24)
                .expect("row height");
            sheet
                .column("C")
                .expect("column C")
                .style(&accent)
                .expect("column style")
                .width(16)
                .expect("column width");
        }
        let committed = edit.commit().expect("style commit");
        assert_eq!(committed.patch().len(), 2);

        let row_change = committed
            .patch()
            .changes()
            .iter()
            .find_map(Change::row)
            .expect("row change");
        let RowState::Stored(row_after) = row_change.2 else {
            panic!("expected stored row")
        };
        assert!(row_after.custom_format());
        assert!(matches!(row_after.style(), StyleState::Shared(_)));

        let column_change = committed
            .patch()
            .changes()
            .iter()
            .find_map(Change::column)
            .expect("column change");
        let ColumnState::Stored(column_after) = column_change.2 else {
            panic!("expected stored column")
        };
        assert!(matches!(column_after.style(), StyleState::Shared(_)));

        let sheet = committed
            .workbook()
            .sheet("Sheet1")
            .expect("lookup")
            .expect("sheet");
        assert!(sheet.row(1).expect("row 2").custom_format());
        assert!(matches!(
            sheet.row_style(1).expect("row style"),
            Some(crate::LocalStyle::Shared(style)) if style.same(&accent)
        ));
        assert!(matches!(
            sheet.column_style("C").expect("column style"),
            Some(crate::LocalStyle::Shared(style)) if style.same(&accent)
        ));

        let mut reset = committed.workbook().edit().expect("reset edit");
        {
            let mut sheet = reset.sheet("Sheet1").expect("lookup").expect("sheet");
            sheet.row(1).expect("row 2").reset_style();
            sheet.column("C").expect("column C").reset_style();
        }
        let reset = reset.commit().expect("reset commit");
        let reset_sheet = reset
            .workbook()
            .sheet("Sheet1")
            .expect("lookup")
            .expect("sheet");
        assert!(!reset_sheet.row(1).expect("row 2").custom_format());
        assert!(matches!(
            reset_sheet.row_style(1).expect("row style"),
            Some(crate::LocalStyle::Default)
        ));
        assert!(matches!(
            reset_sheet.column_style("C").expect("column style"),
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

        let reopened = Workbook::from_bytes(source_bytes).expect("reopened source");
        let replayed = reopened
            .apply(committed.patch())
            .expect("source-checked replay");
        let (_, _, replayed_row) = replayed
            .patch()
            .changes()
            .iter()
            .find_map(Change::row)
            .expect("replayed row");
        let RowState::Stored(replayed_row) = replayed_row else {
            panic!("expected replayed row")
        };
        let StyleState::Shared(replayed_key) = replayed_row.style() else {
            panic!("expected rebound row style")
        };
        assert!(
            replayed
                .workbook()
                .styles()
                .expect("styles")
                .find(replayed_key)
                .is_some()
        );
        assert!(
            source
                .styles()
                .expect("source styles")
                .find(replayed_key)
                .is_none()
        );

        let mut styles = source.edit().expect("styles edit");
        {
            let mut sheet = styles.sheet(0usize).expect("lookup").expect("sheet");
            sheet
                .row(2)
                .expect("row 3")
                .style(&accent)
                .expect("row style");
            sheet
                .column("D")
                .expect("column D")
                .style(&accent)
                .expect("column style");
        }
        let mut layout = source.edit().expect("layout edit");
        {
            let mut sheet = layout.sheet(0usize).expect("lookup").expect("sheet");
            sheet.row(2).expect("row 3").height(22).expect("height");
            sheet
                .column("D")
                .expect("column D")
                .width(18)
                .expect("width");
        }
        styles.join(layout).expect("disjoint grid facets");
        let joined = styles.commit().expect("joined commit");
        let joined_sheet = joined
            .workbook()
            .sheet(0usize)
            .expect("lookup")
            .expect("sheet");
        assert_eq!(
            joined_sheet
                .row(2)
                .expect("row 3")
                .height()
                .map(crate::row::Height::get),
            Some(22.0)
        );
        assert!(matches!(
            joined_sheet.column_style("D").expect("column style"),
            Some(crate::LocalStyle::Shared(_))
        ));

        let mut left = source.edit().expect("left style");
        left.sheet(0usize)
            .expect("lookup")
            .expect("sheet")
            .row(3)
            .expect("row 4")
            .style(&accent)
            .expect("style");
        let mut right = source.edit().expect("right style");
        right
            .sheet(0usize)
            .expect("lookup")
            .expect("sheet")
            .row(3)
            .expect("row 4")
            .reset_style();
        assert!(left.join(right).is_err());

        let mut missing_width = source.edit().expect("missing-width edit");
        missing_width
            .sheet(0usize)
            .expect("lookup")
            .expect("sheet")
            .column("E")
            .expect("column E")
            .style(&accent)
            .expect("lineage");
        assert!(matches!(
            missing_width.commit(),
            Err(Error::ColumnEditBlocked {
                reason: crate::error::ColumnEditBlock::StyleNeedsWidth,
                ..
            })
        ));

        let foreign = Workbook::new()
            .expect("foreign workbook")
            .styles()
            .expect("foreign styles")
            .base()
            .expect("foreign base style");
        let mut rejected = source.edit().expect("rejected edit");
        {
            let mut sheet = rejected.sheet(0usize).expect("lookup").expect("sheet");
            assert!(matches!(
                sheet.row(4).expect("row 5").style(&foreign),
                Err(Error::ForeignStyle)
            ));
            assert!(matches!(
                sheet.column("E").expect("column E").style(&foreign),
                Err(Error::ForeignStyle)
            ));
        }
        assert!(rejected.is_empty());

        let mut add = source.edit().expect("new sheet edit");
        {
            let mut sheet = add.add("Styled").expect("new sheet");
            sheet.set("A2", "row").expect("row cell");
            sheet.set("C1", "column").expect("column cell");
            sheet
                .row(1)
                .expect("row 2")
                .style(&accent)
                .expect("row style");
            sheet
                .column("C")
                .expect("column C")
                .width(12)
                .expect("column width")
                .style(&accent)
                .expect("column style");
        }
        let added = add.commit().expect("new sheet commit");
        let sheet = added
            .workbook()
            .sheet("Styled")
            .expect("lookup")
            .expect("sheet");
        assert!(matches!(
            sheet.row_style(1).expect("row style"),
            Some(crate::LocalStyle::Shared(_))
        ));
        assert!(matches!(
            sheet.column_style("C").expect("column style"),
            Some(crate::LocalStyle::Shared(_))
        ));
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
    fn tab_rename_is_typed_dependency_aware_move_first_and_reversible() {
        let source = rename_reference_workbook();
        let source_bytes = source.to_bytes().expect("source bytes");
        let source_part = source.inner.sheets[0].part_uri.clone();
        let source_id = source.inner.sheets[0].native_id;
        let source_relationship = source.inner.sheets[0].relationship_id.clone();

        let mut edit = source.edit().expect("edit");
        edit.tab("data")
            .expect("caseless lookup")
            .expect("tab")
            .rename(String::from("Input 2026"))
            .expect("checked rename");
        edit.sheet("Calc")
            .expect("Calc lookup")
            .expect("Calc sheet")
            .set("D1", "composed")
            .expect("same-part cell edit");
        let committed = edit.commit().expect("rename commit");
        assert!(
            committed
                .workbook()
                .sheet("Data")
                .expect("lookup")
                .is_none()
        );
        let renamed = committed
            .workbook()
            .sheet("INPUT 2026")
            .expect("Unicode caseless lookup")
            .expect("renamed sheet");
        assert_eq!(renamed.name(), "Input 2026");
        assert_eq!(renamed.data.part_uri, source_part);
        assert_eq!(renamed.data.native_id, source_id);
        assert_eq!(renamed.data.relationship_id, source_relationship);
        assert_eq!(committed.patch().len(), 2);
        assert_eq!(
            committed.patch().changes()[0].renamed(),
            Some((0, "Data", "Input 2026"))
        );

        let calc = committed
            .workbook()
            .sheet("Calc")
            .expect("lookup")
            .expect("Calc");
        assert!(matches!(
            calc.cell("A1").expect("formula cell"),
            Some(Cell::Formula(formula)) if formula.text() == "'Input 2026'!A1"
        ));
        assert!(matches!(
            calc.cell("D1").expect("composed cell"),
            Some(Cell::Value(Value::Text(text))) if text.as_str() == "composed"
        ));
        for uri in [
            "/xl/workbook.xml",
            "/xl/worksheets/sheet2.xml",
            "/xl/tables/table1.xml",
            "/xl/charts/chart1.xml",
            "/xl/pivotCache/pivotCacheDefinition1.xml",
            "/docProps/app.xml",
        ] {
            let text = part_text(committed.workbook(), uri);
            assert!(
                text.contains("Input 2026"),
                "expected renamed dependency in {uri}: {text}"
            );
        }
        assert!(
            part_text(committed.workbook(), "/xl/externalLinks/externalLink1.xml")
                .contains("[1]Data!A1")
        );

        let restored = committed
            .workbook()
            .apply(&committed.patch().inverse())
            .expect("inverse rename");
        assert_eq!(restored.workbook().to_bytes().expect("bytes"), source_bytes);
        assert_eq!(source.to_bytes().expect("source unchanged"), source_bytes);
    }

    #[test]
    fn worksheet_add_is_atomic_populatable_active_and_reversible() {
        let source = Workbook::new().expect("source workbook");
        let source_bytes = source.to_bytes().expect("source bytes");
        let mut edit = source.edit().expect("edit");
        {
            let mut sheet = edit.add(String::from("Summary")).expect("new sheet");
            assert_eq!(sheet.name(), "Summary");
            assert_eq!(sheet.position(), 1);
            sheet
                .set("A1", "created in one transaction")
                .and_then(|sheet| sheet.set("B2", Formula::new("1+1").expect("checked formula")))
                .expect("new cells");
            sheet.row(3u32).expect("row").hide();
            sheet.column(2u32).expect("column").hide();
            sheet.activate();
        }
        let committed = edit.commit().expect("create commit");
        assert_eq!(source.to_bytes().expect("source unchanged"), source_bytes);
        assert_eq!(
            committed
                .workbook()
                .sheets()
                .map(|sheet| sheet.name().to_owned())
                .collect::<Vec<_>>(),
            ["Sheet1", "Summary"]
        );
        assert_eq!(
            committed.workbook().active_sheet().expect("active").name(),
            "Summary"
        );
        let summary = committed
            .workbook()
            .sheet("summary")
            .expect("caseless lookup")
            .expect("created sheet");
        assert_eq!(summary.data.native_id, 2);
        assert!(matches!(
            summary.cell("A1").expect("A1"),
            Some(Cell::Value(Value::Text(text))) if text.as_str() == "created in one transaction"
        ));
        assert!(matches!(
            summary.cell("B2").expect("B2"),
            Some(Cell::Formula(formula)) if formula.text() == "1+1"
        ));
        assert!(summary.row(3u32).expect("row").hidden());
        assert!(summary.column(2u32).expect("column").hidden());
        assert!(
            committed
                .patch()
                .changes()
                .iter()
                .any(|change| { change.created() == Some((1, "Summary")) })
        );

        let restored = committed
            .workbook()
            .apply(&committed.patch().inverse())
            .expect("inverse create");
        assert_eq!(
            restored.workbook().to_bytes().expect("restored"),
            source_bytes
        );
        let replayed = source.apply(committed.patch()).expect("forward replay");
        assert!(
            replayed
                .workbook()
                .sheet("Summary")
                .expect("lookup")
                .is_some()
        );
    }

    #[test]
    fn worksheet_insert_is_selector_first_order_aware_and_reversible() {
        let source = three_sheet_workbook();
        let source_bytes = source.to_bytes().expect("source bytes");
        let mut edit = source.edit().expect("edit");
        edit.move_before("Sheet3", "Sheet1")
            .expect("move lookup")
            .expect("both base tabs");
        assert!(
            edit.add_before("Never Added", "Absent")
                .expect("missing lookup")
                .is_none()
        );
        {
            let mut sheet = edit
                .add_before("Before A", "Sheet1")
                .expect("lookup")
                .expect("anchor");
            assert_eq!(sheet.position(), 1);
            sheet.set("A1", "before-a").expect("payload");
        }
        edit.add_before("Before B", "Sheet1")
            .expect("lookup")
            .expect("anchor")
            .set("A1", "before-b")
            .expect("payload");
        edit.add_after("After A", "Sheet1")
            .expect("lookup")
            .expect("anchor")
            .set("A1", "after-a")
            .expect("payload");
        edit.add_after("After B", "Sheet1")
            .expect("lookup")
            .expect("anchor")
            .set("A1", "after-b")
            .expect("payload")
            .activate();
        {
            let sheet = edit
                .add_before("Before Third", 2usize)
                .expect("numeric lookup")
                .expect("numeric anchor");
            assert_eq!(sheet.position(), 0);
        }
        edit.add("Tail").expect("tail");

        let committed = edit.commit().expect("insert commit");
        let names = committed
            .workbook()
            .sheets()
            .map(|sheet| sheet.name().to_owned())
            .collect::<Vec<_>>();
        assert_eq!(
            names,
            [
                "Before Third",
                "Sheet3",
                "Before A",
                "Before B",
                "Sheet1",
                "After A",
                "After B",
                "Sheet2",
                "Tail",
            ]
        );
        let active = committed.workbook().active_sheet().expect("active");
        assert_eq!((active.name(), active.position()), ("After B", 6));
        assert!(matches!(
            committed
                .workbook()
                .sheet("Before A")
                .expect("lookup")
                .expect("created")
                .cell("A1")
                .expect("cell"),
            Some(Cell::Value(Value::Text(text))) if text.as_str() == "before-a"
        ));
        assert_eq!(
            committed
                .workbook()
                .defined_names()
                .iter()
                .map(|name| (name.name.as_str(), name.local_sheet_id))
                .collect::<Vec<_>>(),
            [
                ("FirstLocal", Some(4)),
                ("ThirdLocal", Some(1)),
                ("Global", None),
            ]
        );
        assert_eq!(
            committed
                .patch()
                .changes()
                .iter()
                .filter_map(Change::created)
                .collect::<Vec<_>>(),
            [
                (2, "Before A"),
                (3, "Before B"),
                (5, "After A"),
                (6, "After B"),
                (0, "Before Third"),
                (8, "Tail"),
            ]
        );

        let committed_bytes = committed.workbook().to_bytes().expect("committed bytes");
        assert_eq!(
            source
                .apply(committed.patch())
                .expect("forward replay")
                .workbook()
                .to_bytes()
                .expect("replayed bytes"),
            committed_bytes
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
    fn worksheet_insert_join_preserves_explicit_edit_order() {
        let source = two_sheet_workbook(SheetKind::Worksheet);
        let mut left = source.edit().expect("left");
        left.add_before("Left", "Sheet2")
            .expect("lookup")
            .expect("anchor");
        let mut right = source.edit().expect("right");
        right
            .add_before("Right", 1usize)
            .expect("lookup")
            .expect("anchor");
        left.join(right).expect("disjoint insertion join");
        assert_eq!(
            left.commit()
                .expect("joined insertion")
                .workbook()
                .sheets()
                .map(|sheet| sheet.name().to_owned())
                .collect::<Vec<_>>(),
            ["Sheet1", "Left", "Right", "Sheet2"]
        );
    }

    #[test]
    fn worksheet_add_validates_names_visibility_and_parallel_joins() {
        let source = Workbook::new().expect("source workbook");

        let mut duplicate = source.edit().expect("duplicate edit");
        duplicate.add("sheet1").expect("checked spelling");
        assert!(matches!(
            duplicate.commit(),
            Err(Error::SheetNameConflict {
                first: 0,
                second: 1,
                ..
            })
        ));

        let mut duplicate = source.edit().expect("anchored duplicate edit");
        duplicate
            .add_before("sheet1", "Sheet1")
            .expect("lookup")
            .expect("anchor");
        assert!(matches!(
            duplicate.commit(),
            Err(Error::SheetNameConflict {
                first: 0,
                second: 1,
                ..
            })
        ));

        let mut hidden = source.edit().expect("hidden edit");
        {
            let mut sheet = hidden.add("Hidden").expect("new sheet");
            sheet.hide().activate();
        }
        assert!(matches!(
            hidden.commit(),
            Err(Error::TabEditBlocked {
                reason: TabEditBlock::NotVisible,
                ..
            })
        ));

        let mut replacement = source.edit().expect("replacement edit");
        replacement
            .tab("Sheet1")
            .expect("lookup")
            .expect("existing tab")
            .hide();
        replacement.add("Replacement").expect("visible replacement");
        let replaced = replacement.commit().expect("replacement commit");
        assert_eq!(
            replaced.workbook().active_sheet().expect("active").name(),
            "Replacement"
        );
        assert!(matches!(
            replaced
                .workbook()
                .sheet("Sheet1")
                .expect("lookup")
                .expect("old tab")
                .visibility(),
            Visibility::Hidden
        ));

        let mut left = source.edit().expect("left");
        left.add("North")
            .expect("North")
            .set("A1", 1_i32)
            .expect("cell");
        let mut right = source.edit().expect("right");
        right
            .add("South")
            .expect("South")
            .set("A1", 2_i32)
            .expect("cell");
        left.join(right).expect("disjoint appends join");
        let joined = left.commit().expect("joined create");
        assert_eq!(
            joined
                .workbook()
                .sheets()
                .map(|sheet| sheet.name().to_owned())
                .collect::<Vec<_>>(),
            ["Sheet1", "North", "South"]
        );

        let mut one = source.edit().expect("one");
        one.add("Résumé").expect("first name");
        let mut two = source.edit().expect("two");
        two.add("RE\u{301}SUME\u{301}").expect("equivalent name");
        let error = one.join(two).expect_err("equivalent creations conflict");
        assert!(error.conflicts().expect("conflicts").conflicts()[0].is_name());
    }

    #[test]
    fn worksheet_add_closes_rename_formula_and_extended_property_dependencies() {
        let source = rename_reference_workbook();
        let source_bytes = source.to_bytes().expect("source bytes");
        let mut edit = source.edit().expect("edit");
        edit.tab("Data")
            .expect("lookup")
            .expect("tab")
            .rename("Input")
            .expect("rename");
        edit.add("New & Sheet")
            .expect("new sheet")
            .set(
                "A1",
                Formula::new("Data!A1").expect("source-snapshot formula"),
            )
            .expect("formula");
        let committed = edit.commit().expect("composed structural commit");
        let created = committed
            .workbook()
            .sheet("New & Sheet")
            .expect("lookup")
            .expect("created sheet");
        assert!(matches!(
            created.cell("A1").expect("formula"),
            Some(Cell::Formula(formula)) if formula.text() == "Input!A1"
        ));
        let properties = part_text(committed.workbook(), "/docProps/app.xml");
        assert!(properties.contains("size=\"4\""));
        assert!(properties.contains(
            "<vt:lpstr>Input</vt:lpstr><vt:lpstr>Calc</vt:lpstr><vt:lpstr>New &amp; Sheet</vt:lpstr><vt:lpstr>Input!Print_Area</vt:lpstr>"
        ));

        let restored = committed
            .workbook()
            .apply(&committed.patch().inverse())
            .expect("inverse");
        assert_eq!(
            restored.workbook().to_bytes().expect("restored"),
            source_bytes
        );
    }

    #[test]
    fn worksheet_add_synchronizes_optional_properties_during_a_simultaneous_reorder() {
        let source = rename_reference_workbook();
        let source_bytes = source.to_bytes().expect("source bytes");
        let mut edit = source.edit().expect("edit");
        edit.move_before("Calc", "Data")
            .expect("move")
            .expect("both tabs");
        edit.add_before("Middle", "Data")
            .expect("lookup")
            .expect("anchor");
        edit.add("Tail").expect("new sheet");
        let committed = edit.commit().expect("composed commit");
        assert_eq!(
            committed
                .workbook()
                .sheets()
                .map(|sheet| sheet.name().to_owned())
                .collect::<Vec<_>>(),
            ["Calc", "Middle", "Data", "Tail"]
        );
        let properties = part_text(committed.workbook(), "/docProps/app.xml");
        assert!(properties.contains("size=\"5\""));
        assert!(properties.contains(concat!(
            "<vt:lpstr>Calc</vt:lpstr>",
            "<vt:lpstr>Middle</vt:lpstr>",
            "<vt:lpstr>Data</vt:lpstr>",
            "<vt:lpstr>Tail</vt:lpstr>",
            "<vt:lpstr>Data!Print_Area</vt:lpstr>"
        )));
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
    fn worksheet_add_allocates_strict_graph_identity_without_exposing_it() {
        let baseline = Workbook::new().expect("baseline");
        let mut package = baseline.inner.package.clone();
        let main = package
            .get_part_mut(&baseline.inner.workbook_uri)
            .expect("workbook part");
        main.set_blob(
            br#"<workbook xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main" xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships"><sheets><sheet name="Strict1" sheetId="7" r:id="tab"/></sheets></workbook>"#.to_vec(),
        );
        main.rels_mut().remove("rId1").expect("old worksheet rel");
        main.rels_mut()
            .try_add_relationship(
                litchi_opc::constants::relationship_type::STRICT_WORKSHEET.to_owned(),
                "worksheets/sheet1.xml".to_owned(),
                "tab".to_owned(),
                TargetMode::Internal,
            )
            .expect("strict worksheet rel");
        package
            .get_part_mut(&PackURI::new("/xl/worksheets/sheet1.xml").expect("URI"))
            .expect("worksheet")
            .set_blob(
                br#"<worksheet xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main"><sheetData/></worksheet>"#.to_vec(),
        );
        let source = Workbook::from_package(package).expect("strict source");
        let mut edit = source.edit().expect("edit");
        edit.add_before("Strict2", "Strict1")
            .expect("lookup")
            .expect("strict anchor");
        let committed = edit.commit().expect("strict create");
        let sheet = committed
            .workbook()
            .sheet("Strict2")
            .expect("lookup")
            .expect("created");
        assert_eq!(sheet.position(), 0);
        assert_eq!(sheet.data.native_id, 1);
        assert_eq!(sheet.data.relationship_id, "rId1");
        assert_eq!(sheet.data.part_uri.as_str(), "/xl/worksheets/sheet2.xml");
        let main = committed
            .workbook()
            .inner
            .package
            .get_part(&committed.workbook().inner.workbook_uri)
            .expect("workbook part");
        assert_eq!(
            main.rels()
                .get(&sheet.data.relationship_id)
                .expect("new relationship")
                .reltype(),
            litchi_opc::constants::relationship_type::STRICT_WORKSHEET
        );
        assert!(
            part_text(committed.workbook(), sheet.data.part_uri.as_str())
                .contains("http://purl.oclc.org/ooxml/spreadsheetml/main")
        );
    }

    #[test]
    fn worksheet_add_blocks_protected_and_compatibility_owned_catalogs() {
        for xml in [
            br#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><workbookProtection lockStructure="1"/><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#.as_slice(),
            br#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:test"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/><mc:AlternateContent><mc:Choice Requires="x"><x:payload/></mc:Choice><mc:Fallback/></mc:AlternateContent></sheets></workbook>"#.as_slice(),
        ] {
            let baseline = Workbook::new().expect("baseline");
            let mut package = baseline.inner.package.clone();
            package
                .get_part_mut(&baseline.inner.workbook_uri)
                .expect("workbook")
                .set_blob(xml.to_vec());
            let source = Workbook::from_package(package).expect("source");
            let mut edit = source.edit().expect("edit");
            edit.add_before("Blocked", "Sheet1")
                .expect("lookup")
                .expect("anchor");
            assert!(matches!(
                edit.commit(),
                Err(Error::TabEditBlocked {
                    reason: TabEditBlock::ProtectedWorkbook
                        | TabEditBlock::MarkupCompatibility,
                    ..
                })
            ));
        }
    }

    #[test]
    fn worksheet_remove_is_selector_first_active_relocating_and_reversible() {
        let baseline = Workbook::new().expect("baseline");
        let mut create = baseline.edit().expect("create edit");
        create
            .add("Delete")
            .expect("Delete")
            .set("A1", "removed payload")
            .expect("payload")
            .activate();
        create
            .add("Keep")
            .expect("Keep")
            .set("A1", 42_i32)
            .expect("retained payload");
        let source = create.commit().expect("create source").into_workbook();
        let source_bytes = source.to_bytes().expect("source bytes");
        assert_eq!(source.active_sheet().expect("active").name(), "Delete");

        let mut edit = source.edit().expect("remove edit");
        assert!(edit.remove("missing").expect("missing selector").is_none());
        edit.remove("delete")
            .expect("selector")
            .expect("Delete worksheet");
        assert_eq!(edit.len(), 1);
        let committed = edit.commit().expect("remove commit");
        assert_eq!(
            committed
                .workbook()
                .sheets()
                .map(|sheet| sheet.name().to_owned())
                .collect::<Vec<_>>(),
            ["Sheet1", "Keep"]
        );
        assert_eq!(
            committed.workbook().active_sheet().expect("active").name(),
            "Keep"
        );
        assert!(
            committed
                .workbook()
                .sheet("Delete")
                .expect("lookup")
                .is_none()
        );
        assert!(matches!(
            committed.patch().changes().first(),
            Some(Change::Remove {
                sheet,
                position: 1,
                ..
            }) if sheet.as_ref() == "Delete"
        ));
        assert!(
            committed
                .patch()
                .changes()
                .iter()
                .any(|change| matches!(change, Change::Active { after, .. } if after.name() == "Keep" && after.position() == 1))
        );

        let restored = committed
            .workbook()
            .apply(&committed.patch().inverse())
            .expect("inverse remove");
        assert_eq!(
            restored.workbook().to_bytes().expect("restored"),
            source_bytes
        );
        let replayed = source.apply(committed.patch()).expect("forward replay");
        assert!(
            replayed
                .workbook()
                .sheet("Delete")
                .expect("lookup")
                .is_none()
        );

        let mut activate_last = source.edit().expect("activate last");
        activate_last
            .tab("Keep")
            .expect("lookup")
            .expect("Keep")
            .activate();
        let active_last = activate_last
            .commit()
            .expect("active-last source")
            .into_workbook();
        let mut remove_last = active_last.edit().expect("remove last");
        remove_last.remove("Keep").expect("lookup").expect("Keep");
        assert_eq!(
            remove_last
                .commit()
                .expect("remove last active")
                .workbook()
                .active_sheet()
                .expect("replacement active")
                .name(),
            "Delete"
        );
    }

    #[test]
    fn worksheet_remove_blocks_live_formulas_last_sheet_and_mixed_edits() {
        let baseline = Workbook::new().expect("baseline");
        let mut create = baseline.edit().expect("create edit");
        create.add("Delete").expect("Delete");
        create
            .sheet("Sheet1")
            .expect("lookup")
            .expect("Sheet1")
            .set("A1", Formula::new("Delete!A1").expect("formula"))
            .expect("formula cell");
        let source = create.commit().expect("source").into_workbook();
        let mut remove = source.edit().expect("remove edit");
        remove.remove("Delete").expect("lookup").expect("Delete");
        assert!(matches!(
            remove.commit(),
            Err(Error::SheetRemoveBlocked {
                sheet,
                reason: RemoveBlock::IncomingReference,
                ..
            }) if sheet == "Delete"
        ));

        let single = Workbook::new().expect("single");
        let mut last = single.edit().expect("last edit");
        last.remove(0usize).expect("lookup").expect("only sheet");
        assert!(matches!(
            last.commit(),
            Err(Error::SheetRemoveBlocked {
                reason: RemoveBlock::LastSheet,
                ..
            })
        ));

        let baseline = Workbook::new().expect("visibility baseline");
        let mut create = baseline.edit().expect("visibility create");
        create
            .tab("Sheet1")
            .expect("lookup")
            .expect("Sheet1")
            .hide();
        create.add("Delete").expect("Delete").activate();
        let visibility = create.commit().expect("visibility source").into_workbook();
        let mut last_visible = visibility.edit().expect("last-visible edit");
        last_visible
            .remove("Delete")
            .expect("lookup")
            .expect("Delete");
        assert!(matches!(
            last_visible.commit(),
            Err(Error::TabEditBlocked {
                reason: TabEditBlock::LastVisibleTab,
                ..
            })
        ));

        let mut mixed = source.edit().expect("mixed edit");
        mixed
            .sheet("Sheet1")
            .expect("lookup")
            .expect("Sheet1")
            .set("B1", 1_i32)
            .expect("cell edit");
        assert!(matches!(
            mixed.remove("Delete"),
            Err(Error::SheetRemoveBlocked {
                reason: RemoveBlock::MixedEdit,
                ..
            })
        ));

        let baseline = Workbook::new().expect("dynamic baseline");
        let mut create = baseline.edit().expect("dynamic create");
        create.add("Delete").expect("Delete");
        create
            .sheet("Sheet1")
            .expect("lookup")
            .expect("Sheet1")
            .set(
                "A1",
                Formula::new(r#"INDIRECT("Delete!A1")"#).expect("dynamic formula"),
            )
            .expect("dynamic cell");
        let dynamic = create.commit().expect("dynamic source").into_workbook();
        let mut remove = dynamic.edit().expect("dynamic removal");
        remove.remove("Delete").expect("lookup").expect("Delete");
        assert!(matches!(
            remove.commit(),
            Err(Error::SheetRemoveBlocked {
                reason: RemoveBlock::UnmodeledReference,
                ..
            })
        ));
    }

    #[test]
    fn worksheet_remove_joins_disjoint_plans_and_blocks_unknown_dependencies() {
        let baseline = Workbook::new().expect("baseline");
        let mut create = baseline.edit().expect("create edit");
        create.add("North").expect("North");
        create.add("South").expect("South");
        let source = create.commit().expect("source").into_workbook();

        let mut north = source.edit().expect("north edit");
        north.remove("North").expect("lookup").expect("North");
        let mut south = source.edit().expect("south edit");
        south.remove(2usize).expect("lookup").expect("South");
        north.join(south).expect("disjoint removals join");
        let committed = north.commit().expect("joined removal");
        assert_eq!(
            committed
                .workbook()
                .sheets()
                .map(|sheet| sheet.name().to_owned())
                .collect::<Vec<_>>(),
            ["Sheet1"]
        );

        let mut package = source.inner.package.clone();
        let custom_uri = PackURI::new("/customXml/item1.xml").expect("custom URI");
        package
            .try_add_part(Box::new(BlobPart::new(
                custom_uri,
                "application/xml".to_owned(),
                b"<root><futureFormulaCache>North!A1</futureFormulaCache></root>".to_vec(),
            )))
            .expect("custom part");
        package
            .get_part_mut(&source.inner.workbook_uri)
            .expect("workbook")
            .rels_mut()
            .try_add_relationship(
                "urn:litchi:test-custom".to_owned(),
                "../customXml/item1.xml".to_owned(),
                "customRef".to_owned(),
                TargetMode::Internal,
            )
            .expect("custom relationship");
        let unknown = Workbook::from_package(package).expect("unknown producer workbook");
        let mut edit = unknown.edit().expect("remove edit");
        edit.remove("North").expect("lookup").expect("North");
        assert!(matches!(
            edit.commit(),
            Err(Error::SheetRemoveBlocked {
                reason: RemoveBlock::UnmodeledReference,
                part,
                ..
            }) if part == "/customXml/item1.xml"
        ));
    }

    #[test]
    fn worksheet_remove_blocks_macro_projects_and_extra_incoming_relationships() {
        let baseline = Workbook::new().expect("baseline");
        let mut create = baseline.edit().expect("create edit");
        create.add("Delete").expect("Delete");
        let source = create.commit().expect("source").into_workbook();

        let mut macro_package = source.inner.package.clone();
        macro_package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new("/xl/vbaProject.bin").expect("VBA URI"),
                litchi_opc::constants::content_type::OFC_VBA_PROJECT.to_owned(),
                vec![0, 1, 2, 3],
            )))
            .expect("VBA part");
        macro_package
            .get_part_mut(&source.inner.workbook_uri)
            .expect("workbook")
            .rels_mut()
            .try_add_relationship(
                litchi_opc::constants::relationship_type::VBA_PROJECT.to_owned(),
                "vbaProject.bin".to_owned(),
                "vbaProject".to_owned(),
                TargetMode::Internal,
            )
            .expect("VBA relationship");
        let macro_book = Workbook::from_package(macro_package).expect("macro workbook");
        let mut remove = macro_book.edit().expect("macro remove");
        remove.remove("Delete").expect("lookup").expect("Delete");
        assert!(matches!(
            remove.commit(),
            Err(Error::SheetRemoveBlocked {
                reason: RemoveBlock::MacroProject,
                ..
            })
        ));

        let target = source
            .inner
            .sheets
            .get(1)
            .expect("Delete sheet")
            .part_uri
            .clone();
        let mut incoming_package = source.inner.package.clone();
        let mut referrer = BlobPart::new(
            PackURI::new("/xl/custom.xml").expect("custom URI"),
            "application/xml".to_owned(),
            b"<custom/>".to_vec(),
        );
        referrer
            .rels_mut()
            .try_add_relationship(
                "urn:litchi:test-incoming".to_owned(),
                target.relative_ref("/xl"),
                "sheetRef".to_owned(),
                TargetMode::Internal,
            )
            .expect("incoming relationship");
        incoming_package
            .try_add_part(Box::new(referrer))
            .expect("referrer part");
        let incoming = Workbook::from_package(incoming_package).expect("incoming workbook");
        let mut remove = incoming.edit().expect("incoming remove");
        remove.remove("Delete").expect("lookup").expect("Delete");
        assert!(matches!(
            remove.commit(),
            Err(Error::SheetRemoveBlocked {
                reason: RemoveBlock::IncomingRelationship,
                part,
                ..
            }) if part == "/xl/custom.xml"
        ));
    }

    #[test]
    fn worksheet_remove_blocks_custom_workbook_view_identity() {
        let baseline = Workbook::new().expect("baseline");
        let mut create = baseline.edit().expect("create edit");
        create.add("Delete").expect("Delete");
        let source = create.commit().expect("source").into_workbook();
        let native_id = source.inner.sheets[1].native_id;
        let mut package = source.inner.package.clone();
        let workbook = package
            .get_part_mut(&source.inner.workbook_uri)
            .expect("workbook");
        let xml = std::str::from_utf8(workbook.blob())
            .expect("workbook UTF-8")
            .replace(
                "</workbook>",
                &format!(
                    r#"<customWorkbookViews><customWorkbookView name="Delete view" guid="{{00000000-0000-0000-0000-000000000001}}" activeSheetId="{native_id}"/></customWorkbookViews></workbook>"#
                ),
            );
        assert!(xml.contains("customWorkbookView"));
        workbook.set_blob(xml.into_bytes());
        let source = Workbook::from_package(package).expect("custom-view workbook");
        let mut remove = source.edit().expect("remove edit");
        remove.remove("Delete").expect("lookup").expect("Delete");
        assert!(matches!(
            remove.commit(),
            Err(Error::SheetRemoveBlocked {
                reason: RemoveBlock::IncomingReference,
                part,
                ..
            }) if part == "/xl/workbook.xml"
        ));
    }

    #[test]
    fn worksheet_remove_accepts_case_equivalent_opc_targets() {
        let baseline = Workbook::new().expect("baseline");
        let mut create = baseline.edit().expect("create edit");
        create.add("Delete").expect("Delete");
        let source = create.commit().expect("source").into_workbook();
        let relationship_id = source.inner.sheets[1].relationship_id.clone();
        let mut package = source.inner.package.clone();
        let relationships = package
            .get_part_mut(&source.inner.workbook_uri)
            .expect("workbook")
            .rels_mut();
        let relationship = relationships
            .remove(&relationship_id)
            .expect("worksheet relationship");
        relationships
            .try_add_relationship(
                relationship.reltype().to_owned(),
                "worksheets/SHEET2.XML".to_owned(),
                relationship_id,
                TargetMode::Internal,
            )
            .expect("case-equivalent relationship");
        let source = Workbook::from_package(package).expect("case-equivalent workbook");
        let mut remove = source.edit().expect("remove edit");
        remove.remove("Delete").expect("lookup").expect("Delete");
        assert!(
            remove
                .commit()
                .expect("case-equivalent removal")
                .workbook()
                .sheet("Delete")
                .expect("lookup")
                .is_none()
        );
    }

    #[test]
    fn tab_rename_validates_names_collisions_swaps_and_join_facets() {
        let source = two_sheet_workbook(SheetKind::Worksheet);
        let mut invalid = source.edit().expect("invalid edit");
        let error = invalid
            .tab("Sheet1")
            .expect("lookup")
            .expect("tab")
            .rename("bad/name")
            .expect_err("invalid name");
        assert!(matches!(error, Error::SheetName(_)));

        let mut collision = source.edit().expect("collision edit");
        collision
            .tab("Sheet1")
            .expect("lookup")
            .expect("tab")
            .rename("sheet2")
            .expect("valid spelling");
        assert!(matches!(
            collision.commit(),
            Err(Error::SheetNameConflict {
                first: 0,
                second: 1,
                ..
            })
        ));

        let mut swap = source.edit().expect("swap edit");
        swap.tab("Sheet1")
            .expect("lookup")
            .expect("tab")
            .rename("Sheet2")
            .expect("first swap");
        swap.tab("Sheet2")
            .expect("lookup")
            .expect("tab")
            .rename("Sheet1")
            .expect("second swap");
        let swapped = swap.commit().expect("simultaneous swap");
        assert_eq!(
            swapped
                .workbook()
                .sheets()
                .map(|sheet| sheet.name().to_owned())
                .collect::<Vec<_>>(),
            ["Sheet2", "Sheet1"]
        );

        let mut left = source.edit().expect("left");
        left.tab("Sheet1")
            .expect("lookup")
            .expect("tab")
            .rename("First")
            .expect("left rename");
        let mut same = source.edit().expect("same");
        same.tab(0usize)
            .expect("lookup")
            .expect("tab")
            .rename("Other")
            .expect("same rename");
        let error = left.join(same).expect_err("same name facet conflicts");
        assert!(error.conflicts().expect("conflicts").conflicts()[0].is_name());

        let mut right = source.edit().expect("right");
        right
            .tab("Sheet2")
            .expect("lookup")
            .expect("tab")
            .rename(Name::new("Second").expect("prevalidated"))
            .expect("moved typed name");
        left.join(right).expect("disjoint names join");
        let joined = left.commit().expect("joined renames");
        assert!(joined.workbook().sheet("first").expect("lookup").is_some());
        assert!(joined.workbook().sheet("second").expect("lookup").is_some());
    }

    #[test]
    fn tab_reorder_is_selector_first_dependency_aware_and_reversible() {
        let source = three_sheet_workbook();
        let source_bytes = source.to_bytes().expect("source bytes");
        let mut edit = source.edit().expect("edit");
        assert!(
            edit.move_before("Sheet3", "Sheet1")
                .expect("move lookup")
                .is_some()
        );
        assert_eq!(edit.len(), 1);
        let committed = edit.commit().expect("reorder");
        assert_eq!(
            committed
                .workbook()
                .sheets()
                .map(|sheet| sheet.name().to_owned())
                .collect::<Vec<_>>(),
            ["Sheet3", "Sheet1", "Sheet2"]
        );
        assert_eq!(
            committed
                .workbook()
                .active_sheet()
                .expect("active sheet")
                .name(),
            "Sheet2"
        );
        assert_eq!(
            committed
                .workbook()
                .active_sheet()
                .expect("active sheet")
                .position(),
            2
        );
        assert!(matches!(
            committed
                .workbook()
                .sheet("Sheet3")
                .expect("lookup")
                .expect("Sheet3")
                .cell("A1")
                .expect("cell"),
            Some(Cell::Value(Value::Text(text))) if text.as_str() == "three"
        ));
        assert_eq!(
            committed
                .workbook()
                .defined_names()
                .iter()
                .map(|name| (name.name.as_str(), name.local_sheet_id))
                .collect::<Vec<_>>(),
            [
                ("FirstLocal", Some(1)),
                ("ThirdLocal", Some(0)),
                ("Global", None),
            ]
        );
        assert_eq!(committed.patch().len(), 2);
        assert_eq!(committed.patch().changes()[0].sheet(), "Sheet3");
        assert_eq!(committed.patch().changes()[0].moved(), Some((2, 0)));
        let (before, after) = committed.patch().changes()[1]
            .active()
            .expect("active position remap");
        assert_eq!((before.name(), before.position()), ("Sheet2", 1));
        assert_eq!((after.name(), after.position()), ("Sheet2", 2));

        let inverse = committed.patch().inverse();
        assert_eq!(inverse.changes()[1].moved(), Some((0, 2)));
        let restored = committed
            .workbook()
            .apply(&inverse)
            .expect("inverse reorder");
        assert_eq!(restored.workbook().to_bytes().expect("bytes"), source_bytes);
        assert_eq!(source.to_bytes().expect("source unchanged"), source_bytes);

        let chart_source = two_sheet_workbook(SheetKind::Chart);
        let chart_bytes = chart_source.to_bytes().expect("chart source bytes");
        let mut chart_edit = chart_source.edit().expect("chart edit");
        chart_edit
            .move_before("Sheet2", "Sheet1")
            .expect("chart lookup")
            .expect("chart tabs");
        let chart = chart_edit.commit().expect("chart reorder");
        let first = chart
            .workbook()
            .sheet(0usize)
            .expect("lookup")
            .expect("tab");
        assert_eq!((first.name(), first.kind()), ("Sheet2", SheetKind::Chart));
        assert_eq!(
            chart
                .workbook()
                .apply(&chart.patch().inverse())
                .expect("chart inverse")
                .workbook()
                .to_bytes()
                .expect("chart bytes"),
            chart_bytes
        );
    }

    #[test]
    fn tab_reorder_composes_with_other_facets_and_conflicts_globally() {
        let source = three_sheet_workbook();
        let mut order = source.edit().expect("order");
        order
            .move_before("Sheet3", "Sheet1")
            .expect("lookup")
            .expect("tabs");
        let mut active = source.edit().expect("active");
        active
            .tab("Sheet3")
            .expect("lookup")
            .expect("tab")
            .activate();
        let mut visibility = source.edit().expect("visibility");
        visibility
            .tab("Sheet1")
            .expect("lookup")
            .expect("tab")
            .hide();
        let mut cell = source.edit().expect("cell");
        cell.sheet("Sheet3")
            .expect("lookup")
            .expect("sheet")
            .set("B1", "moved payload")
            .expect("cell");
        order.join(active).expect("order and active");
        order.join(visibility).expect("order and visibility");
        order.join(cell).expect("order and cell");
        let committed = order.commit().expect("composed reorder");
        let active = committed.workbook().active_sheet().expect("active");
        assert_eq!((active.name(), active.position()), ("Sheet3", 0));
        assert_eq!(
            committed
                .workbook()
                .sheet("Sheet1")
                .expect("lookup")
                .expect("sheet")
                .visibility(),
            &Visibility::Hidden
        );
        assert!(matches!(
            active.cell("B1").expect("cell"),
            Some(Cell::Value(Value::Text(text))) if text.as_str() == "moved payload"
        ));
        assert_eq!(
            committed
                .patch()
                .parts
                .iter()
                .filter(|part| part.uri == source.inner.workbook_uri)
                .count(),
            1
        );
        assert_eq!(
            committed
                .patch()
                .parts
                .iter()
                .filter(|part| part.uri == source.inner.sheets[2].part_uri)
                .count(),
            1
        );

        let mut left = source.edit().expect("left order");
        left.move_before("Sheet3", "Sheet1")
            .expect("lookup")
            .expect("tabs");
        let mut right = source.edit().expect("right order");
        right
            .move_after("Sheet1", "Sheet2")
            .expect("lookup")
            .expect("tabs");
        let error = left.join(right).expect_err("order is one global facet");
        let conflicts = error.conflicts().expect("order conflict");
        assert_eq!(conflicts.len(), 1);
        assert!(conflicts.conflicts()[0].is_order());

        let source = two_sheet_workbook(SheetKind::Worksheet);
        let mut same_position = source.edit().expect("same-position edit");
        same_position
            .move_before("Sheet2", "Sheet1")
            .expect("lookup")
            .expect("tabs");
        same_position
            .tab("Sheet2")
            .expect("lookup")
            .expect("tab")
            .activate();
        let committed = same_position
            .commit()
            .expect("active identity changes at the same position");
        let active = committed.workbook().active_sheet().expect("active");
        assert_eq!((active.name(), active.position()), ("Sheet2", 0));
        let active_change = committed
            .patch()
            .changes()
            .iter()
            .find_map(Change::active)
            .expect("semantic active change");
        assert_eq!(active_change.0.position(), 0);
        assert_eq!(active_change.1.position(), 0);
        assert_eq!(active_change.0.name(), "Sheet1");
        assert_eq!(active_change.1.name(), "Sheet2");
    }

    #[test]
    fn numeric_tab_moves_are_checked_and_all_positions_round_trip() {
        for from in 0..3usize {
            for to in 0..3usize {
                let source = three_sheet_workbook();
                let source_bytes = source.to_bytes().expect("source bytes");
                let mut expected = vec!["Sheet1", "Sheet2", "Sheet3"];
                let moved = expected.remove(from);
                expected.insert(to, moved);
                let mut edit = source.edit().expect("edit");
                assert!(edit.move_to(from, to).expect("lookup").is_some());
                let committed = edit.commit().expect("move");
                assert_eq!(
                    committed
                        .workbook()
                        .sheets()
                        .map(|sheet| sheet.name().to_owned())
                        .collect::<Vec<_>>(),
                    expected.into_iter().map(str::to_owned).collect::<Vec<_>>()
                );
                let restored = committed
                    .workbook()
                    .apply(&committed.patch().inverse())
                    .expect("inverse");
                assert_eq!(restored.workbook().to_bytes().expect("bytes"), source_bytes);
            }
        }

        let source = three_sheet_workbook();
        let mut missing = source.edit().expect("missing edit");
        assert!(
            missing
                .move_before("Absent", "Sheet1")
                .expect("lookup")
                .is_none()
        );
        assert!(missing.move_to("Sheet1", 3).expect("bounds").is_none());
        assert!(missing.commit().expect("no-op").patch().is_empty());

        let mut cancelled = source.edit().expect("cancelled edit");
        cancelled
            .move_before("Sheet3", "Sheet1")
            .expect("lookup")
            .expect("tabs");
        cancelled
            .move_after("Sheet3", "Sheet2")
            .expect("lookup")
            .expect("tabs");
        assert!(cancelled.is_empty());
        assert_eq!(cancelled.len(), 0);
        assert!(cancelled.commit().expect("cancelled").patch().is_empty());

        let mut cancelled = source.edit().expect("cancelled join edit");
        cancelled
            .move_before("Sheet3", "Sheet1")
            .expect("lookup")
            .expect("tabs");
        cancelled
            .move_after("Sheet3", "Sheet2")
            .expect("lookup")
            .expect("tabs");
        let mut effective = source.edit().expect("effective join edit");
        effective
            .move_before("Sheet2", "Sheet1")
            .expect("lookup")
            .expect("tabs");
        cancelled
            .join(effective)
            .expect("cancelled order has no conflict");
        let joined = cancelled.commit().expect("joined order");
        assert_eq!(
            joined
                .workbook()
                .sheets()
                .map(|sheet| sheet.name().to_owned())
                .collect::<Vec<_>>(),
            ["Sheet2", "Sheet1", "Sheet3"]
        );

        let source_bytes = source.to_bytes().expect("source bytes");
        let mut sequence = source.edit().expect("sequence edit");
        sequence
            .move_before("Sheet3", "Sheet1")
            .expect("lookup")
            .expect("tabs");
        sequence
            .move_after("Sheet1", "Sheet2")
            .expect("lookup")
            .expect("tabs");
        let sequence = sequence.commit().expect("move sequence");
        assert_eq!(
            sequence
                .workbook()
                .sheets()
                .map(|sheet| sheet.name().to_owned())
                .collect::<Vec<_>>(),
            ["Sheet3", "Sheet2", "Sheet1"]
        );
        assert_eq!(sequence.patch().len(), 2);
        assert_eq!(sequence.patch().changes()[0].moved(), Some((2, 0)));
        assert_eq!(sequence.patch().changes()[1].moved(), Some((1, 2)));
        let inverse = sequence.patch().inverse();
        assert_eq!(inverse.changes()[0].moved(), Some((2, 1)));
        assert_eq!(inverse.changes()[1].moved(), Some((0, 2)));
        assert_eq!(
            sequence
                .workbook()
                .apply(&inverse)
                .expect("sequence inverse")
                .workbook()
                .to_bytes()
                .expect("restored bytes"),
            source_bytes
        );
    }

    #[test]
    fn tab_reorder_blocks_protection_and_revision_tracking() {
        let source = three_sheet_workbook();
        let mut package = source.inner.package.clone();
        let workbook = package
            .get_part_mut(&source.inner.workbook_uri)
            .expect("workbook");
        let xml = std::str::from_utf8(workbook.blob())
            .expect("UTF-8")
            .replace(
                "<bookViews>",
                "<workbookProtection lockStructure=\"1\"/><bookViews>",
            );
        workbook.set_blob(xml.into_bytes());
        let protected = Workbook::from_package(package).expect("protected workbook");
        let mut edit = protected.edit().expect("edit");
        edit.move_before("Sheet3", "Sheet1")
            .expect("lookup")
            .expect("tabs");
        assert!(matches!(
            edit.commit(),
            Err(Error::TabEditBlocked {
                reason: TabEditBlock::ProtectedWorkbook,
                ..
            })
        ));
        let mut rename = protected.edit().expect("rename edit");
        rename
            .tab("Sheet1")
            .expect("lookup")
            .expect("tab")
            .rename("Input")
            .expect("checked name");
        assert!(matches!(
            rename.commit(),
            Err(Error::TabEditBlocked {
                reason: TabEditBlock::ProtectedWorkbook,
                ..
            })
        ));

        let mut package = source.inner.package.clone();
        package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new("/xl/revisions/revisionHeaders1.xml").expect("revision URI"),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.revisionHeaders+xml"
                    .to_owned(),
                br#"<headers xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#
                    .to_vec(),
            )))
            .expect("revision part");
        package
            .get_part_mut(&source.inner.workbook_uri)
            .expect("workbook")
            .rels_mut()
            .try_add_relationship(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/revisionHeaders"
                    .to_owned(),
                "revisions/revisionHeaders1.xml".to_owned(),
                "rIdRevisionHeaders".to_owned(),
                TargetMode::Internal,
            )
            .expect("revision relationship");
        let tracked = Workbook::from_package(package).expect("tracked workbook");
        let mut edit = tracked.edit().expect("edit");
        edit.move_before("Sheet3", "Sheet1")
            .expect("lookup")
            .expect("tabs");
        assert!(matches!(
            edit.commit(),
            Err(Error::TabEditBlocked {
                reason: TabEditBlock::TrackedWorkbook,
                ..
            })
        ));
        let mut rename = tracked.edit().expect("rename edit");
        rename
            .tab("Sheet1")
            .expect("lookup")
            .expect("tab")
            .rename("Input")
            .expect("checked name");
        assert!(matches!(
            rename.commit(),
            Err(Error::TabEditBlocked {
                reason: TabEditBlock::TrackedWorkbook,
                ..
            })
        ));
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

    fn rename_reference_workbook() -> Workbook {
        let source = two_sheet_workbook(SheetKind::Worksheet);
        let mut package = source.inner.package.clone();
        package
            .get_part_mut(&source.inner.workbook_uri)
            .expect("workbook part")
            .set_blob(
                br#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Calc" sheetId="2" r:id="rIdTab2"/></sheets><definedNames><definedName name="Source">Data!$A$1</definedName></definedNames></workbook>"#.to_vec(),
            );
        package
            .get_part_mut(&source.inner.sheets[0].part_uri)
            .expect("Data worksheet")
            .set_blob(
                br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1"/><sheetData><row r="1"><c r="A1"><v>7</v></c></row></sheetData></worksheet>"#.to_vec(),
            );
        package
            .get_part_mut(&source.inner.sheets[1].part_uri)
            .expect("Calc worksheet")
            .set_blob(
                br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1"/><sheetData><row r="1"><c r="A1"><f>Data!A1</f><v>7</v></c></row></sheetData><dataValidations count="1"><dataValidation type="custom" sqref="B1"><formula1>Data!A1&gt;0</formula1></dataValidation></dataValidations><hyperlinks><hyperlink ref="C1" location="Data!$A$1"/></hyperlinks></worksheet>"#.to_vec(),
            );
        for (uri, content_type, content) in [
            (
                "/xl/tables/table1.xml",
                litchi_opc::constants::content_type::SML_TABLE,
                br#"<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><tableColumns count="1"><tableColumn id="1" name="Value"><calculatedColumnFormula>Data!A1</calculatedColumnFormula></tableColumn></tableColumns></table>"#.as_slice(),
            ),
            (
                "/xl/charts/chart1.xml",
                litchi_opc::constants::content_type::DML_CHART,
                br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:barChart><c:ser><c:val><c:numRef><c:f>Data!$A$1</c:f></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"#.as_slice(),
            ),
            (
                "/xl/pivotCache/pivotCacheDefinition1.xml",
                litchi_opc::constants::content_type::SML_PIVOT_CACHE_DEFINITION,
                br#"<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cacheSource type="worksheet"><worksheetSource sheet="Data" ref="A1"/></cacheSource></pivotCacheDefinition>"#.as_slice(),
            ),
            (
                "/docProps/app.xml",
                litchi_opc::constants::content_type::OFC_EXTENDED_PROPERTIES,
                br#"<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><TitlesOfParts><vt:vector size="3" baseType="lpstr"><vt:lpstr>Data</vt:lpstr><vt:lpstr>Calc</vt:lpstr><vt:lpstr>Data!Print_Area</vt:lpstr></vt:vector></TitlesOfParts></Properties>"#.as_slice(),
            ),
            (
                "/xl/externalLinks/externalLink1.xml",
                litchi_opc::constants::content_type::SML_EXTERNAL_LINK,
                br#"<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><externalBook><definedNames><definedName name="External" refersTo="[1]Data!A1"/></definedNames></externalBook></externalLink>"#.as_slice(),
            ),
        ] {
            package
                .try_add_part(Box::new(BlobPart::new(
                    PackURI::new(uri).expect("part URI"),
                    content_type.to_owned(),
                    content.to_vec(),
                )))
                .expect("reference part");
        }
        Workbook::from_package(package).expect("rename reference workbook")
    }

    fn part_text<'a>(workbook: &'a Workbook, uri: &str) -> &'a str {
        let uri = PackURI::new(uri).expect("part URI");
        let bytes = workbook
            .inner
            .package
            .get_part(&uri)
            .expect("package part")
            .blob();
        std::str::from_utf8(bytes).expect("XML part")
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

    fn styled_column_workbook() -> Workbook {
        let baseline = styled_workbook();
        let mut package = baseline.inner.package.clone();
        package
            .get_part_mut(&PackURI::new("/xl/worksheets/sheet1.xml").expect("sheet URI"))
            .expect("worksheet part")
            .set_blob(
                br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cols><col min="3" max="3" style="1"/></cols><sheetData><row r="1"><c r="A1" s="1"><v>1</v></c><c r="B1"><v>2</v></c></row></sheetData></worksheet>"#.to_vec(),
            );
        Workbook::from_package(package).expect("styled column workbook")
    }

    fn styled_row_workbook() -> Workbook {
        let baseline = styled_workbook();
        let mut package = baseline.inner.package.clone();
        package
            .get_part_mut(&PackURI::new("/xl/worksheets/sheet1.xml").expect("sheet URI"))
            .expect("worksheet part")
            .set_blob(
                br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" s="1"><v>1</v></c></row><row r="2" s="1" customFormat="1"/></sheetData></worksheet>"#.to_vec(),
            );
        Workbook::from_package(package).expect("styled row workbook")
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

    fn three_sheet_workbook() -> Workbook {
        let source = two_sheet_workbook(SheetKind::Worksheet);
        let mut package = source.inner.package.clone();
        package
            .get_part_mut(&source.inner.workbook_uri)
            .expect("workbook part")
            .set_blob(
                br#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView activeTab="1" firstSheet="0"/><workbookView activeTab="2" firstSheet="1"/></bookViews><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/><sheet name="Sheet2" sheetId="2" r:id="rIdTab2"/><sheet name="Sheet3" sheetId="3" r:id="rIdTab3"/></sheets><definedNames><definedName name="FirstLocal" localSheetId="0">Sheet1!$A$1</definedName><definedName name="ThirdLocal" localSheetId="2">Sheet3!$A$1</definedName><definedName name="Global">1</definedName></definedNames></workbook>"#.to_vec(),
            );
        package
            .get_part_mut(&source.inner.workbook_uri)
            .expect("workbook part")
            .rels_mut()
            .try_add_relationship(
                litchi_opc::constants::relationship_type::WORKSHEET.to_owned(),
                "worksheets/sheet3.xml".to_owned(),
                "rIdTab3".to_owned(),
                TargetMode::Internal,
            )
            .expect("third sheet relationship");
        package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new("/xl/worksheets/sheet3.xml").expect("third sheet URI"),
                litchi_opc::constants::content_type::SML_WORKSHEET.to_owned(),
                br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>three</t></is></c></row></sheetData></worksheet>"#.to_vec(),
            )))
            .expect("third sheet part");
        package
            .get_part_mut(&source.inner.sheets[0].part_uri)
            .expect("first sheet part")
            .set_blob(
                br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>one</t></is></c></row></sheetData></worksheet>"#.to_vec(),
            );
        package
            .get_part_mut(&source.inner.sheets[1].part_uri)
            .expect("second sheet part")
            .set_blob(
                br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1"/><sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>two</t></is></c></row></sheetData></worksheet>"#.to_vec(),
            );
        Workbook::from_package(package).expect("three-sheet workbook")
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
