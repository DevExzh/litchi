//! Implementation of the public [`crate::table::lifecycle`] transactions:
//! adding an empty table to a sheet and removing a table with its private
//! graph (`IWORK_API_REQUESTS.md` request 5, record 0301).
//!
//! **Removal** detaches the table's drawable from its sheet (and its node
//! from the sheet's sidebar node) and removes every object only the table
//! owned: its model, storage tiles, data lists, header buckets and other
//! sidecars, by the reference census of [`super::structure::collect`] —
//! storage a surviving table shares stays. The table's CalculationEngine
//! owners, their private dependency tiles and every edge other owners hold
//! into them follow through the TSCE phase ported from the Pages
//! body-table deletion owner; a surviving formula that reads the table
//! refuses the removal ([`Error::UnsupportedDependency`]). Components the
//! removal empties are deleted with their registry records.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "the focused transaction keeps its public types beside the staged rewrite"
)]

use std::collections::BTreeSet;
use std::fmt;
use std::num::NonZeroU64;

use thiserror::Error as ThisError;

mod add;
pub(in crate::package) mod duplicate;

use super::Package;
use super::structure::{
    self, DOCUMENT_SHEETS_FIELD, DOCUMENT_SIDEBAR_FIELD, Fault, FaultResult, TREE_CHILDREN_FIELD,
    TREE_NODE_MESSAGE_TYPE, TREE_OBJECT_FIELD, Working,
    budget::WireBudget,
    collect::{Census, verify_removed},
    metadata::{self, Transition},
    plan::{DeletionObjectIndex, GraphPlan, ObjectRemoval, RemovalPlan},
    references,
};
use crate::selector::{SheetSelector, TableSelector};

structure::structural_family!("table-lifecycle");

/// A typed, content-redacted table-lifecycle failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies it
    /// (record 0351 §2.4, record 0361).
    #[error(transparent)]
    Refused(Refusal),
    /// The rooted source is malformed or ambiguous.
    #[error("rooted table-lifecycle source is invalid")]
    InvalidSource,
    /// A finite transaction limit was exceeded.
    #[error("{kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        /// Resource category.
        kind: LimitKind,
        /// Observed amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded allocation failed.
    #[error("could not allocate {amount} table-lifecycle items or bytes")]
    Allocation {
        /// Requested amount.
        amount: usize,
    },
    /// The reopened candidate did not reproduce the requested state.
    #[error("table-lifecycle candidate verification failed")]
    Verification,
    /// The patch was applied to an artifact other than its exact source.
    #[error("table-lifecycle patch conflicts with this package")]
    PatchConflict,
}

/// A well-formed request this family's route refuses, classified by
/// [`Refusal::kind`] (record 0351 §2.4, record 0361).
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    /// The sheet selector did not resolve one sheet.
    #[error("sheet selector did not resolve")]
    SheetNotFound,
    /// The table selector did not resolve one table of the sheet.
    #[error("table selector did not resolve")]
    TableNotFound,
    /// The requested table name is empty, contains NUL, or names another
    /// table of the sheet.
    #[error("invalid table name: {reason}")]
    InvalidName {
        /// Why the name was refused.
        reason: crate::sheet::lifecycle::InvalidName,
    },
    /// A table needs at least one row and one column, and at most
    /// 1,100,000 rows and columns together.
    #[error("invalid table size: {rows} rows, {columns} columns")]
    InvalidSize {
        /// Requested rows.
        rows: u32,
        /// Requested columns.
        columns: u32,
    },
    /// The workbook has no table whose structure a new table can copy.
    #[error("the workbook has no table to copy a new table's structure from")]
    NoTemplate,
    /// This edit already holds its single operation.
    #[error("a table-lifecycle operation is already staged")]
    OperationAlreadyStaged,
    /// Commit was requested before staging an operation.
    #[error("no table-lifecycle operation is staged")]
    NoStagedOperation,
    /// A surviving object (such as a formula in another table) depends on
    /// the table's graph.
    #[error("a surviving object depends on the table's graph")]
    UnsupportedDependency,
    /// The source cannot publish an exact changed artifact.
    #[error("source representation does not support table lifecycle changes")]
    UnsupportedSource,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::SheetNotFound { .. } | Self::TableNotFound { .. } => RefusalKind::NotFound,
            Self::InvalidName { .. } | Self::InvalidSize { .. } => RefusalKind::InvalidValue,
            Self::NoTemplate { .. }
            | Self::UnsupportedDependency { .. }
            | Self::UnsupportedSource { .. } => RefusalKind::UnsupportedContent,
            Self::OperationAlreadyStaged { .. } | Self::NoStagedOperation { .. } => {
                RefusalKind::UnsupportedOperation
            },
        }
    }
}

impl From<Refusal> for Error {
    fn from(refusal: Refusal) -> Self {
        Self::Refused(refusal)
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum Operation {
    /// Append an empty table.
    Add {
        name: String,
        rows: u32,
        columns: u32,
    },
    /// Remove the table at this position of the sheet.
    Remove(usize),
    /// Copy the table at this position of the sheet, under this name.
    Duplicate { position: usize, name: String },
}

/// One immutable edit that stages exactly one table addition or removal in
/// one sheet.
pub struct Edit<'a> {
    source: &'a Package,
    sheet: usize,
    operation: Option<Operation>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("sheet", &self.sheet)
            .field("operation", &self.operation)
            .finish()
    }
}

fn table_position(
    source: &Package,
    sheet: usize,
    selector: TableSelector<'_>,
) -> Result<usize, Error> {
    let sheet = source
        .workbook_ref()
        .sheets()
        .get(sheet)
        .ok_or(Error::Refused(Refusal::SheetNotFound))?;
    match selector {
        TableSelector::Position(position) => {
            let position = position.get();
            if position >= sheet.table_count() {
                return Err(Error::Refused(Refusal::TableNotFound));
            }
            Ok(position)
        },
        TableSelector::Name(name) => {
            let mut matches = sheet
                .tables()
                .enumerate()
                .filter(|(_, table)| table.name() == name);
            let (position, _) = matches
                .next()
                .ok_or(Error::Refused(Refusal::TableNotFound))?;
            if matches.next().is_some() {
                return Err(Error::InvalidSource);
            }
            Ok(position)
        },
    }
}

impl Edit<'_> {
    /// Stage a new empty table named `name` with `rows` rows and `columns`
    /// columns, appended after the sheet's drawables. Its structure (styles,
    /// header rows and columns, sizes) is copied from the workbook's first
    /// table.
    ///
    /// # Errors
    ///
    /// [`Refusal::InvalidName`] for an empty name, a name containing NUL, or
    /// the name of another table of the sheet; [`Refusal::InvalidSize`];
    /// [`Refusal::OperationAlreadyStaged`].
    pub fn add(mut self, name: &str, rows: u32, columns: u32) -> Result<Self, Error> {
        use crate::sheet::lifecycle::InvalidName;
        if self.operation.is_some() {
            return Err(Error::Refused(Refusal::OperationAlreadyStaged));
        }
        let sheet = self
            .source
            .workbook_ref()
            .sheets()
            .get(self.sheet)
            .ok_or(Error::Refused(Refusal::SheetNotFound))?;
        let reason = if name.is_empty() {
            Some(InvalidName::Empty)
        } else if name.contains('\0') {
            Some(InvalidName::ContainsNul)
        } else if sheet.tables().any(|table| table.name() == name) {
            Some(InvalidName::Duplicate)
        } else {
            None
        };
        if let Some(reason) = reason {
            return Err(Error::Refused(Refusal::InvalidName { reason }));
        }
        add::check_size(rows, columns)?;
        let mut owned = String::new();
        owned
            .try_reserve_exact(name.len())
            .map_err(|_error| Error::Allocation { amount: name.len() })?;
        owned.push_str(name);
        self.operation = Some(Operation::Add {
            name: owned,
            rows,
            columns,
        });
        Ok(self)
    }

    /// Stage the removal of the selected table and its private graph.
    ///
    /// # Errors
    ///
    /// [`Refusal::TableNotFound`], [`Refusal::OperationAlreadyStaged`].
    pub fn remove<'selector>(
        mut self,
        selector: impl Into<TableSelector<'selector>>,
    ) -> Result<Self, Error> {
        if self.operation.is_some() {
            return Err(Error::Refused(Refusal::OperationAlreadyStaged));
        }
        let position = table_position(self.source, self.sheet, selector.into())?;
        self.operation = Some(Operation::Remove(position));
        Ok(self)
    }

    /// Stage a populated copy of the selected table, named `"<name> copy"`
    /// (then `"<name> copy 2"`, …), at the end of the sheet's drawables,
    /// offset by ten points. The copy gets its own storage, a table UUID
    /// derived from its model and a clone of the table's CalculationEngine
    /// owner family; styles stay shared, and a formula that reads another
    /// table keeps reading it (record 0370h).
    ///
    /// # Errors
    ///
    /// [`Refusal::TableNotFound`], [`Refusal::OperationAlreadyStaged`].
    pub fn duplicate<'selector>(
        mut self,
        selector: impl Into<TableSelector<'selector>>,
    ) -> Result<Self, Error> {
        if self.operation.is_some() {
            return Err(Error::Refused(Refusal::OperationAlreadyStaged));
        }
        let position = table_position(self.source, self.sheet, selector.into())?;
        let sheet = &self.source.workbook_ref().sheets()[self.sheet];
        let names: Vec<&str> = sheet.tables().map(crate::Table::name).collect();
        let source = names
            .get(position)
            .ok_or(Error::Refused(Refusal::TableNotFound))?;
        let name =
            duplicate::copy_name(source, &names).ok_or(Error::Refused(Refusal::InvalidName {
                reason: crate::sheet::lifecycle::InvalidName::Duplicate,
            }))?;
        self.operation = Some(Operation::Duplicate { position, name });
        Ok(self)
    }

    /// Validate and atomically publish the staged candidate.
    ///
    /// # Errors
    ///
    /// [`Refusal::NoStagedOperation`] for an empty edit;
    /// [`Refusal::UnsupportedDependency`] when a surviving formula or object
    /// depends on the removed graph; source, resource and verification
    /// errors. Failure never publishes a partial package.
    pub fn commit(self) -> Result<Commit, Error> {
        let operation = self
            .operation
            .ok_or(Error::Refused(Refusal::NoStagedOperation))?;
        let sheet = &self.source.workbook_ref().sheets()[self.sheet];
        let mut expected: Vec<String> = sheet
            .tables()
            .map(|table| table.name().to_owned())
            .collect();
        let published = match &operation {
            Operation::Add {
                name,
                rows,
                columns,
            } => {
                expected.push(name.clone());
                add::add_table(
                    self.source,
                    &add::Request {
                        sheet: self.sheet,
                        name,
                        rows: *rows,
                        columns: *columns,
                    },
                )?
            },
            Operation::Remove(position) => {
                let position = *position;
                expected.remove(position);
                remove_table(self.source, self.sheet, position)?
            },
            Operation::Duplicate { position, name } => {
                expected.push(name.clone());
                duplicate::duplicate_table(self.source, self.sheet, *position, name)?
            },
        };
        let candidate = published
            .package
            .workbook_ref()
            .sheets()
            .get(self.sheet)
            .ok_or(Error::Verification)?;
        if candidate
            .tables()
            .map(crate::Table::name)
            .ne(expected.iter().map(String::as_str))
        {
            return Err(Error::Verification);
        }
        if let Operation::Add { rows, columns, .. } = operation {
            let added = candidate.tables().last().ok_or(Error::Verification)?;
            if (added.row_count(), added.column_count()) != (rows, columns) {
                return Err(Error::Verification);
            }
        }
        if let Operation::Duplicate { position, .. } = operation {
            let copy = candidate.tables().last().ok_or(Error::Verification)?;
            let source = sheet.tables().nth(position).ok_or(Error::Verification)?;
            if (copy.row_count(), copy.column_count())
                != (source.row_count(), source.column_count())
            {
                return Err(Error::Verification);
            }
        }
        Commit::new(self.source, published)
    }
}

/// The sheet, its drawables and the table drawable at `position`.
struct Located {
    sheet: u64,
    sheet_index: usize,
    sheet_path: &'static [u32],
    drawables: Vec<u64>,
    table: u64,
    /// The sheet's sidebar node, its children and the child addressing the
    /// table, when the sidebar records them.
    sidebar: Option<(u64, Vec<u64>, u64)>,
}

fn locate(working: &Working<'_>, sheet_position: usize, position: usize) -> FaultResult<Located> {
    let document = working.object(1)?;
    let document_index = structure::document_message(document)?;
    let data = &document.messages[document_index].data;
    let sheets = references::identifiers(data, &[DOCUMENT_SHEETS_FIELD])?;
    let sidebar_root = references::identifiers(data, &[DOCUMENT_SIDEBAR_FIELD])?;
    let sheet = *sheets.get(sheet_position).ok_or(Fault::Invalid)?;
    let sheet_object = working.object(sheet)?;
    let (sheet_index, sheet_path) = structure::sheet_message(sheet_object)?;
    let drawables = references::identifiers(&sheet_object.messages[sheet_index].data, sheet_path)?;
    let mut tables = Vec::new();
    for drawable in &drawables {
        if structure::table_info_message(working.object(*drawable)?).is_ok() {
            tables.push(*drawable);
        }
    }
    let table = *tables.get(position).ok_or(Fault::Invalid)?;
    let mut sidebar = None;
    if let [root] = sidebar_root.as_slice() {
        let root_object = working.object(*root)?;
        let root_index = structure::unique_message(root_object, &[TREE_NODE_MESSAGE_TYPE])?;
        for node in references::identifiers(
            &root_object.messages[root_index].data,
            &[TREE_CHILDREN_FIELD],
        )? {
            let node_object = working.object(node)?;
            let node_index = structure::unique_message(node_object, &[TREE_NODE_MESSAGE_TYPE])?;
            let node_data = &node_object.messages[node_index].data;
            if references::identifiers(node_data, &[TREE_OBJECT_FIELD])?.as_slice() != [sheet] {
                continue;
            }
            let children = references::identifiers(node_data, &[TREE_CHILDREN_FIELD])?;
            for child in &children {
                let leaf = working.object(*child)?;
                let leaf_index = structure::unique_message(leaf, &[TREE_NODE_MESSAGE_TYPE])?;
                if references::identifiers(&leaf.messages[leaf_index].data, &[TREE_OBJECT_FIELD])?
                    .as_slice()
                    == [table]
                {
                    if sidebar.is_some() {
                        return Err(Fault::Invalid);
                    }
                    sidebar = Some((node, children.clone(), *child));
                }
            }
        }
    }
    Ok(Located {
        sheet,
        sheet_index,
        sheet_path,
        drawables,
        table,
        sidebar,
    })
}

/// Rewrite the reference list at `path` of message `index` of object
/// `identifier` from `before` to `after`.
pub(super) fn rewrite_list(
    working: &mut Working<'_>,
    identifier: u64,
    index: usize,
    path: &[u32],
    before: &[u64],
    after: &[u64],
) -> FaultResult<()> {
    let limits = working.archive_limits();
    let object = working.object_mut(identifier)?;
    let data = references::rewrite(&object.messages[index].data, path, after)?;
    let removed: Vec<u64> = before
        .iter()
        .copied()
        .filter(|item| !after.contains(item))
        .collect();
    let appended: Vec<u64> = after
        .iter()
        .copied()
        .filter(|item| !before.contains(item))
        .collect();
    references::replace(object, index, data, &removed, &appended, path, limits)
}

pub(super) fn remove_table(
    source: &Package,
    sheet_position: usize,
    position: usize,
) -> Result<structure::Published, Error> {
    let mut working = Working::new(source)?;
    working.keep_previews();
    let located = locate(&working, sheet_position, position)?;
    let mut roots = vec![located.table];
    let mut detached = vec![(located.sheet, located.table)];
    if let Some((node, _, child)) = &located.sidebar {
        roots.push(*child);
        detached.push((*node, *child));
    }
    let census = Census::new(&working)?;
    let closure = census.private_closure(&working, &roots)?;

    // The CalculationEngine owners of the removed contexts.
    let index = DeletionObjectIndex::new(source)?;
    let mut graph = GraphPlan {
        index,
        removals: RemovalPlan::default(),
    };
    for identifier in &closure {
        let nonzero = NonZeroU64::new(*identifier).ok_or(Error::InvalidSource)?;
        let location = graph.index.get(nonzero).ok_or(Error::InvalidSource)?;
        graph.removals.object_removals.push(ObjectRemoval {
            component_index: location.component_index,
            object_index: location.object_index,
            identifier: nonzero,
        });
        graph.removals.formula_contexts.push(nonzero);
    }
    let mut budget = WireBudget::new(source.source_bytes().len());
    let formula = structure::formula::prepare(source, &graph, &mut budget)?;
    let mut removed: BTreeSet<u64> = closure.clone();
    for removal in &formula.removals.object_removals {
        removed.insert(removal.identifier.get());
    }
    for edit in &formula.removals.message_edits {
        for target in &edit.remove_object_references {
            detached.push((edit.object_identifier.get(), target.get()));
        }
    }
    census.require_detached(&removed, &detached)?;

    // Detach the table from its sheet and the sidebar.
    let after: Vec<u64> = located
        .drawables
        .iter()
        .copied()
        .filter(|item| *item != located.table)
        .collect();
    rewrite_list(
        &mut working,
        located.sheet,
        located.sheet_index,
        located.sheet_path,
        &located.drawables,
        &after,
    )?;
    if let Some((node, children, child)) = &located.sidebar {
        let node_index =
            structure::unique_message(working.object(*node)?, &[TREE_NODE_MESSAGE_TYPE])?;
        let after: Vec<u64> = children
            .iter()
            .copied()
            .filter(|item| item != child)
            .collect();
        rewrite_list(
            &mut working,
            *node,
            node_index,
            &[TREE_CHILDREN_FIELD],
            children,
            &after,
        )?;
    }
    // The CalculationEngine edits.
    let limits = working.archive_limits();
    for edit in formula.removals.message_edits {
        if !edit.remove_data_references.is_empty()
            || removed.contains(&edit.object_identifier.get())
        {
            return Err(Error::InvalidSource);
        }
        let object = working.object_mut(edit.object_identifier.get())?;
        if object
            .messages
            .get(edit.message_index)
            .map(|message| message.type_)
            != Some(edit.message_type)
        {
            return Err(Error::InvalidSource);
        }
        let pruned: Vec<u64> = edit
            .remove_object_references
            .iter()
            .map(|identifier| identifier.get())
            .collect();
        references::replace(
            object,
            edit.message_index,
            edit.data,
            &pruned,
            &[],
            &[],
            limits,
        )?;
    }
    let removed_list: Vec<u64> = removed.iter().copied().collect();
    let mut transition = Transition::default();
    if let Some(registry) = metadata::registry(&working)? {
        transition.remove_objects(&registry, &removed_list)?;
    }
    // A removed suffix of the identifier space lowers the watermark, as the
    // host editor releases it (record 0370g).
    transition.release_suffix(&removed_list)?;
    for identifier in &removed_list {
        working.remove(*identifier)?;
    }
    let published = working.publish(&transition)?;
    verify_removed(&published.package, &removed)?;
    Ok(published)
}

impl Package {
    /// Start an immutable single-operation table add or remove edit in one
    /// sheet.
    ///
    /// ```no_run
    /// use litchi_numbers::Package;
    ///
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let package = Package::open("input.numbers")?;
    /// let commit = package.edit_table_lifecycle(0)?.remove(0)?.commit()?;
    /// let restored = commit
    ///     .package()
    ///     .apply_table_lifecycle(&commit.patch().inverse())?;
    /// # let _ = restored;
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// # Errors
    ///
    /// [`Refusal::SheetNotFound`] when the selector resolves no sheet.
    pub fn edit_table_lifecycle<'selector>(
        &self,
        sheet: impl Into<SheetSelector<'selector>>,
    ) -> Result<Edit<'_>, Error> {
        let sheet = self
            .workbook_ref()
            .sheet(sheet)
            .map_err(|_error| Error::InvalidSource)?
            .ok_or(Error::Refused(Refusal::SheetNotFound))?
            .index();
        Ok(Edit {
            source: self,
            sheet,
            operation: None,
        })
    }

    /// Apply a table-lifecycle patch to its exact source package.
    ///
    /// # Errors
    ///
    /// [`Error::PatchConflict`] when this package is not the patch's exact
    /// source.
    pub fn apply_table_lifecycle(&self, patch: &Patch) -> Result<Commit, Error> {
        apply_patch(self, patch)
    }
}
