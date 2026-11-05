//! Implementation of the public [`crate::table::order`] transaction: moving
//! one table to another position among the tables of its sheet (record
//! 0301).
//!
//! A sheet lists its drawables (tables, charts, shapes) in one reference
//! list, and the tables of a sheet are its table drawables in that order.
//! The move takes the table's record out of the list and puts it back
//! before the table that will follow it (or right after the last table), so
//! every other drawable keeps its place. When the sheet's sidebar node lists
//! its tables, those nodes are reordered the same way. No object is added or
//! removed, and every untouched record keeps its exact bytes.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "the focused transaction keeps its public types beside the staged rewrite"
)]

use std::fmt;

use litchi_iwa_core::RawMessage;
use thiserror::Error as ThisError;

use super::Package;
use super::structure::{
    self, DOCUMENT_SHEETS_FIELD, DOCUMENT_SIDEBAR_FIELD, FaultResult, TREE_CHILDREN_FIELD,
    TREE_NODE_MESSAGE_TYPE, TREE_OBJECT_FIELD, Working, metadata::Transition, references,
};
use crate::selector::{SheetSelector, TableSelector};

structure::structural_family!("table-order");

/// A typed, content-redacted table-order failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies it
    /// (record 0351 §2.4, record 0361).
    #[error(transparent)]
    Refused(Refusal),
    /// The rooted source is malformed or ambiguous.
    #[error("rooted table-order source is invalid")]
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
    #[error("could not allocate {amount} table-order items or bytes")]
    Allocation {
        /// Requested amount.
        amount: usize,
    },
    /// The reopened candidate did not reproduce the requested order.
    #[error("table-order candidate verification failed")]
    Verification,
    /// The patch was applied to an artifact other than its exact source.
    #[error("table-order patch conflicts with this package")]
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
    /// The destination is outside the sheet's tables.
    #[error("destination position {position} is outside {table_count} tables")]
    DestinationOutOfRange {
        /// Requested zero-based destination.
        position: usize,
        /// Number of tables in the sheet.
        table_count: usize,
    },
    /// This edit already holds its single move.
    #[error("a table-order operation is already staged")]
    OperationAlreadyStaged,
    /// Commit was requested before staging a move.
    #[error("no table-order operation is staged")]
    NoStagedOperation,
    /// A surviving object depends on the moved graph.
    #[error("a surviving object depends on the table order")]
    UnsupportedDependency,
    /// The source cannot publish an exact changed artifact.
    #[error("source representation does not support table reordering")]
    UnsupportedSource,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::SheetNotFound { .. } | Self::TableNotFound { .. } => RefusalKind::NotFound,
            Self::DestinationOutOfRange { .. } => RefusalKind::InvalidValue,
            Self::OperationAlreadyStaged { .. } | Self::NoStagedOperation { .. } => {
                RefusalKind::UnsupportedOperation
            },
            Self::UnsupportedDependency { .. } | Self::UnsupportedSource { .. } => {
                RefusalKind::UnsupportedContent
            },
        }
    }
}

impl From<Refusal> for Error {
    fn from(refusal: Refusal) -> Self {
        Self::Refused(refusal)
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct Operation {
    source: usize,
    destination: usize,
}

/// One immutable edit that stages exactly one table move within a sheet.
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

impl Edit<'_> {
    /// Stage the selected table at the zero-based final position
    /// `destination` among the sheet's tables (interpreted after the table
    /// is taken out). Equal positions stage an exact no-op.
    ///
    /// # Errors
    ///
    /// [`Refusal::TableNotFound`], [`Refusal::DestinationOutOfRange`],
    /// [`Refusal::OperationAlreadyStaged`].
    pub fn move_table<'selector>(
        mut self,
        selector: impl Into<TableSelector<'selector>>,
        destination: usize,
    ) -> Result<Self, Error> {
        if self.operation.is_some() {
            return Err(Error::Refused(Refusal::OperationAlreadyStaged));
        }
        let sheet = self
            .source
            .workbook_ref()
            .sheets()
            .get(self.sheet)
            .ok_or(Error::Refused(Refusal::SheetNotFound))?;
        let source = match selector.into() {
            TableSelector::Position(position) => {
                let position = position.get();
                if position >= sheet.table_count() {
                    return Err(Error::Refused(Refusal::TableNotFound));
                }
                position
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
                position
            },
        };
        if destination >= sheet.table_count() {
            return Err(Error::Refused(Refusal::DestinationOutOfRange {
                position: destination,
                table_count: sheet.table_count(),
            }));
        }
        self.operation = Some(Operation {
            source,
            destination,
        });
        Ok(self)
    }

    /// Validate and atomically publish the staged candidate.
    ///
    /// # Errors
    ///
    /// [`Refusal::NoStagedOperation`] for an empty edit; source, resource and
    /// verification errors. Failure never publishes a partial package.
    pub fn commit(self) -> Result<Commit, Error> {
        let operation = self
            .operation
            .ok_or(Error::Refused(Refusal::NoStagedOperation))?;
        if operation.source == operation.destination {
            return Commit::unchanged(self.source);
        }
        let sheet = &self.source.workbook_ref().sheets()[self.sheet];
        let mut expected: Vec<&str> = sheet.tables().map(crate::Table::name).collect();
        let moved = expected.remove(operation.source);
        expected.insert(operation.destination, moved);
        let published = reorder(self.source, self.sheet, operation)?;
        let candidate = published
            .package
            .workbook_ref()
            .sheets()
            .get(self.sheet)
            .ok_or(Error::Verification)?;
        if candidate
            .tables()
            .map(crate::Table::name)
            .collect::<Vec<_>>()
            != expected
        {
            return Err(Error::Verification);
        }
        Commit::new(self.source, published)
    }
}

/// `after`, with the items of `selected` (in `before` order) permuted into
/// `order`, every other item kept in place.
fn permute(before: &[u64], order: &[u64]) -> Vec<u64> {
    let mut next = order.iter();
    before
        .iter()
        .map(|item| {
            if order.contains(item) {
                next.next().copied().unwrap_or(*item)
            } else {
                *item
            }
        })
        .collect()
}

fn reorder(
    source: &Package,
    sheet_position: usize,
    operation: Operation,
) -> Result<structure::Published, Error> {
    let mut working = Working::new(source)?;
    let document = working.object(1)?;
    let document_index = structure::document_message(document)?;
    let sheets = references::identifiers(
        &document.messages[document_index].data,
        &[DOCUMENT_SHEETS_FIELD],
    )?;
    let sidebar = references::identifiers(
        &document.messages[document_index].data,
        &[DOCUMENT_SIDEBAR_FIELD],
    )?;
    let sheet = *sheets
        .get(sheet_position)
        .ok_or(Error::Refused(Refusal::SheetNotFound))?;
    let sheet_object = working.object(sheet)?;
    let (index, path) = structure::sheet_message(sheet_object)?;
    let drawables = references::identifiers(&sheet_object.messages[index].data, path)?;
    let mut tables = Vec::new();
    for drawable in &drawables {
        if structure::table_info_message(working.object(*drawable)?).is_ok() {
            tables.push(*drawable);
        }
    }
    if operation.source >= tables.len() || operation.destination >= tables.len() {
        return Err(Error::InvalidSource);
    }
    let mut order = tables.clone();
    let moved = order.remove(operation.source);
    order.insert(operation.destination, moved);
    let after = permute(&drawables, &order);
    replace_order(&mut working, sheet, index, path, &after)?;

    // The sheet's sidebar node lists the sheet's tables as child nodes.
    if let [root] = sidebar.as_slice() {
        if let Some((node, children, addressed)) = sidebar_tables(&working, *root, sheet)? {
            let tables_in_sidebar: Vec<u64> = addressed
                .iter()
                .copied()
                .filter(|item| tables.contains(item))
                .collect();
            if tables_in_sidebar.len() == tables.len() {
                let node_order: Vec<u64> = order
                    .iter()
                    .filter_map(|table| {
                        addressed
                            .iter()
                            .position(|item| item == table)
                            .map(|position| children[position])
                    })
                    .collect();
                let table_nodes: Vec<u64> = children
                    .iter()
                    .zip(&addressed)
                    .filter(|(_, item)| tables.contains(item))
                    .map(|(child, _)| *child)
                    .collect();
                let mut after = children.clone();
                let mut next = node_order.iter();
                for child in &mut after {
                    if table_nodes.contains(child) {
                        *child = *next.next().ok_or(Error::InvalidSource)?;
                    }
                }
                let node_index =
                    structure::unique_message(working.object(node)?, &[TREE_NODE_MESSAGE_TYPE])?;
                replace_order(
                    &mut working,
                    node,
                    node_index,
                    &[TREE_CHILDREN_FIELD],
                    &after,
                )?;
            }
        }
    }
    Ok(working.publish(&Transition::default())?)
}

/// The sheet's sidebar node, its children and the object each child
/// addresses, when the sidebar records the sheet.
fn sidebar_tables(
    working: &Working<'_>,
    root: u64,
    sheet: u64,
) -> FaultResult<Option<(u64, Vec<u64>, Vec<u64>)>> {
    let object = working.object(root)?;
    let index = structure::unique_message(object, &[TREE_NODE_MESSAGE_TYPE])?;
    for child in references::identifiers(&object.messages[index].data, &[TREE_CHILDREN_FIELD])? {
        let node = working.object(child)?;
        let node_index = structure::unique_message(node, &[TREE_NODE_MESSAGE_TYPE])?;
        let data = &node.messages[node_index].data;
        if references::identifiers(data, &[TREE_OBJECT_FIELD])?.as_slice() != [sheet] {
            continue;
        }
        let children = references::identifiers(data, &[TREE_CHILDREN_FIELD])?;
        let mut addressed = Vec::new();
        structure::reserve(&mut addressed, children.len())?;
        for grandchild in &children {
            let leaf = working.object(*grandchild)?;
            let leaf_index = structure::unique_message(leaf, &[TREE_NODE_MESSAGE_TYPE])?;
            let objects =
                references::identifiers(&leaf.messages[leaf_index].data, &[TREE_OBJECT_FIELD])?;
            addressed.push(objects.first().copied().unwrap_or(0));
        }
        return Ok(Some((child, children, addressed)));
    }
    Ok(None)
}

/// Replace the reference list at `path` with a permutation of itself; the
/// header's reference set is unchanged.
fn replace_order(
    working: &mut Working<'_>,
    identifier: u64,
    index: usize,
    path: &[u32],
    after: &[u64],
) -> FaultResult<()> {
    let limits = working.archive_limits();
    let object = working.object_mut(identifier)?;
    let message_type = object.messages[index].type_;
    let data = references::rewrite(&object.messages[index].data, path, after)?;
    object
        .replace_message_preserving_header(
            index,
            RawMessage {
                type_: message_type,
                data,
            },
            limits,
        )
        .map_err(|error| structure::map_core(&error))?;
    Ok(())
}

impl Package {
    /// Start an immutable single-move table-order edit for one sheet.
    ///
    /// ```no_run
    /// use litchi_numbers::Package;
    ///
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let package = Package::open("input.numbers")?;
    /// let commit = package.edit_table_order(0)?.move_table(1, 0)?.commit()?;
    /// let restored = commit.package().apply_table_order(&commit.patch().inverse())?;
    /// # let _ = restored;
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// # Errors
    ///
    /// [`Refusal::SheetNotFound`] when the selector resolves no sheet.
    pub fn edit_table_order<'selector>(
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

    /// Apply a table-order patch to its exact source package.
    ///
    /// # Errors
    ///
    /// [`Error::PatchConflict`] when this package is not the patch's exact
    /// source.
    pub fn apply_table_order(&self, patch: &Patch) -> Result<Commit, Error> {
        apply_patch(self, patch)
    }
}
