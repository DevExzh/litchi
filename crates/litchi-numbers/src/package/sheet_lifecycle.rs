//! Implementation of the public [`crate::sheet::lifecycle`] transactions:
//! adding an empty sheet and removing a sheet with its private graph
//! (`IWORK_API_REQUESTS.md` request 4, record 0301).
//!
//! A new sheet copies the first sheet's settings record (page setup, style)
//! without its drawables, headers, footers or guides, takes the requested
//! name, and is appended to the document's sheet list and to the sidebar
//! order as a new tree node, exactly where the Numbers app records a sheet.
//! Removing a sheet first removes each of its tables with its private graph
//! and CalculationEngine owners ([`crate::table::lifecycle`]), then detaches
//! the sheet from both lists and removes every object only it owned. A sheet
//! that holds charts, images, shapes or other non-table drawables is
//! refused.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "the focused transaction keeps its public types beside the staged rewrite"
)]

use std::fmt;

use litchi_iwa_common::wire::patch_length_delimited_field;
use thiserror::Error as ThisError;

use super::Package;
use super::structure::{
    self, DOCUMENT_COMPONENT, DOCUMENT_SHEETS_FIELD, DOCUMENT_SIDEBAR_FIELD,
    FORM_SHEET_SUPER_FIELD, Fault, FaultResult, SHEET_DRAWABLES_FIELD, SHEET_NAME_FIELD,
    TREE_CHILDREN_FIELD, TREE_NODE_MESSAGE_TYPE, TREE_OBJECT_FIELD, Working,
    collect::{Census, verify_removed},
    metadata::{self, Transition},
    references,
};
use crate::selector::SheetSelector;

/// The sheet message type of a new sheet.
const SHEET_MESSAGE_TYPE: u32 = super::SHEET_MESSAGE_TYPE;
/// Fields of a copied sheet record that address the source sheet's private
/// objects: drawables, the deprecated header and footer storages, the guide
/// storage, headers and footers.
const PRIVATE_SHEET_FIELDS: [u32; 6] = [SHEET_DRAWABLES_FIELD, 15, 16, 17, 18, 19];

structure::structural_family!("sheet-lifecycle");

/// Why a sheet name was refused.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum InvalidName {
    /// The name is empty.
    Empty,
    /// The name contains NUL.
    ContainsNul,
    /// Another sheet already has this exact name.
    Duplicate,
}

impl fmt::Display for InvalidName {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::Empty => "the name is empty",
            Self::ContainsNul => "the name contains NUL",
            Self::Duplicate => "another sheet has this name",
        })
    }
}

/// A typed, content-redacted sheet-lifecycle failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies it
    /// (record 0351 §2.4, record 0361).
    #[error(transparent)]
    Refused(Refusal),
    /// The rooted source is malformed or ambiguous.
    #[error("rooted sheet-lifecycle source is invalid")]
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
    #[error("could not allocate {amount} sheet-lifecycle items or bytes")]
    Allocation {
        /// Requested amount.
        amount: usize,
    },
    /// The reopened candidate did not reproduce the requested state.
    #[error("sheet-lifecycle candidate verification failed")]
    Verification,
    /// The patch was applied to an artifact other than its exact source.
    #[error("sheet-lifecycle patch conflicts with this package")]
    PatchConflict,
}

/// A well-formed request this family's route refuses, classified by
/// [`Refusal::kind`] (record 0351 §2.4, record 0361).
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    /// The selector did not resolve one sheet.
    #[error("sheet selector did not resolve")]
    SheetNotFound,
    /// The requested sheet name is invalid.
    #[error("invalid sheet name: {reason}")]
    InvalidName {
        /// Why the name was refused.
        reason: InvalidName,
    },
    /// This edit already holds its single operation.
    #[error("a sheet-lifecycle operation is already staged")]
    OperationAlreadyStaged,
    /// Commit was requested before staging an operation.
    #[error("no sheet-lifecycle operation is staged")]
    NoStagedOperation,
    /// A workbook keeps at least one sheet.
    #[error("the last sheet of a workbook cannot be removed")]
    LastSheet,
    /// The sheet holds drawables other than tables (charts, media, shapes).
    #[error("the sheet holds drawables other than tables")]
    SheetNotEmpty,
    /// A surviving object still depends on the sheet's graph.
    #[error("a surviving object depends on the sheet's graph")]
    UnsupportedDependency,
    /// The source cannot publish an exact changed artifact.
    #[error("source representation does not support sheet lifecycle changes")]
    UnsupportedSource,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::SheetNotEmpty { .. } => RefusalKind::UnsupportedTarget,
            Self::LastSheet { .. } => RefusalKind::UnsupportedTarget,
            Self::SheetNotFound { .. } => RefusalKind::NotFound,
            Self::InvalidName { .. } => RefusalKind::InvalidValue,
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

#[derive(Debug, Clone, PartialEq, Eq)]
enum Operation {
    /// Append a sheet with this name.
    Add(String),
    /// Remove the sheet at this position.
    Remove(usize),
    /// Copy the sheet at this position, with its tables, under this name.
    Duplicate(usize, String),
}

/// One immutable edit that stages exactly one sheet addition or removal.
pub struct Edit<'a> {
    source: &'a Package,
    operation: Option<Operation>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field(
                "operation",
                &self.operation.as_ref().map(|operation| match operation {
                    Operation::Add(_) => "add",
                    Operation::Remove(_) => "remove",
                    Operation::Duplicate(..) => "duplicate",
                }),
            )
            .finish()
    }
}

impl Edit<'_> {
    /// Stage a new empty sheet named `name`, appended after the last sheet.
    ///
    /// # Errors
    ///
    /// [`Refusal::InvalidName`] for an empty name, a name containing NUL, or a
    /// name another sheet already has; [`Refusal::OperationAlreadyStaged`]
    /// after one operation.
    #[allow(
        clippy::should_implement_trait,
        reason = "`add` is the glossary verb for staging a new element, not arithmetic"
    )]
    pub fn add(mut self, name: &str) -> Result<Self, Error> {
        if self.operation.is_some() {
            return Err(Error::Refused(Refusal::OperationAlreadyStaged));
        }
        let reason = if name.is_empty() {
            Some(InvalidName::Empty)
        } else if name.contains('\0') {
            Some(InvalidName::ContainsNul)
        } else if self
            .source
            .workbook_ref()
            .sheets()
            .iter()
            .any(|sheet| sheet.name() == name)
        {
            Some(InvalidName::Duplicate)
        } else {
            None
        };
        if let Some(reason) = reason {
            return Err(Error::Refused(Refusal::InvalidName { reason }));
        }
        let mut owned = String::new();
        owned
            .try_reserve_exact(name.len())
            .map_err(|_error| Error::Allocation { amount: name.len() })?;
        owned.push_str(name);
        self.operation = Some(Operation::Add(owned));
        Ok(self)
    }

    /// Stage the removal of the selected sheet and its private graph.
    ///
    /// # Errors
    ///
    /// [`Refusal::SheetNotFound`] when the selector resolves no sheet,
    /// [`Refusal::LastSheet`] for a workbook's only sheet,
    /// [`Refusal::OperationAlreadyStaged`] after one operation. A sheet that
    /// holds drawables other than tables is refused at commit
    /// ([`Refusal::SheetNotEmpty`]).
    pub fn remove<'selector>(
        mut self,
        selector: impl Into<SheetSelector<'selector>>,
    ) -> Result<Self, Error> {
        if self.operation.is_some() {
            return Err(Error::Refused(Refusal::OperationAlreadyStaged));
        }
        let workbook = self.source.workbook_ref();
        let sheet = workbook
            .sheet(selector)
            .map_err(|_error| Error::InvalidSource)?
            .ok_or(Error::Refused(Refusal::SheetNotFound))?;
        if workbook.sheet_count() <= 1 {
            return Err(Error::Refused(Refusal::LastSheet));
        }
        self.operation = Some(Operation::Remove(sheet.index()));
        Ok(self)
    }

    /// Stage a copy of the selected sheet, inserted right after it and named
    /// `"<name>-1"` (then `"<name>-2"`, …), as the host editor names it. The
    /// copy takes the sheet's settings and a populated copy of each of its
    /// tables, at the same positions and under the same names (record
    /// 0370j).
    ///
    /// # Errors
    ///
    /// [`Refusal::SheetNotFound`], [`Refusal::OperationAlreadyStaged`]. At
    /// commit: [`Refusal::SheetNotEmpty`] for a sheet holding drawables
    /// other than tables; [`Refusal::UnsupportedDependency`] when a table's
    /// formulas read another table; [`Refusal::UnsupportedSource`] for a
    /// form-based sheet.
    pub fn duplicate<'selector>(
        mut self,
        selector: impl Into<SheetSelector<'selector>>,
    ) -> Result<Self, Error> {
        if self.operation.is_some() {
            return Err(Error::Refused(Refusal::OperationAlreadyStaged));
        }
        let workbook = self.source.workbook_ref();
        let sheet = workbook
            .sheet(selector)
            .map_err(|_error| Error::InvalidSource)?
            .ok_or(Error::Refused(Refusal::SheetNotFound))?;
        let names: Vec<&str> = workbook.sheets().iter().map(crate::Sheet::name).collect();
        let name = (1_u32..=u32::MAX)
            .map(|suffix| format!("{}-{suffix}", sheet.name()))
            .find(|candidate| !names.contains(&candidate.as_str()))
            .ok_or(Error::Refused(Refusal::InvalidName {
                reason: InvalidName::Duplicate,
            }))?;
        self.operation = Some(Operation::Duplicate(sheet.index(), name));
        Ok(self)
    }

    /// Validate and atomically publish the staged candidate.
    ///
    /// # Errors
    ///
    /// [`Refusal::NoStagedOperation`] for an empty edit; source, dependency,
    /// resource and verification errors. Failure never publishes a partial
    /// package.
    pub fn commit(self) -> Result<Commit, Error> {
        let operation = self
            .operation
            .ok_or(Error::Refused(Refusal::NoStagedOperation))?;
        let source = self.source;
        let (published, sheets) = match &operation {
            Operation::Add(name) => add_sheet(source, name)?,
            Operation::Remove(position) => remove_sheet(source, *position)?,
            Operation::Duplicate(position, name) => duplicate_sheet(source, *position, name)?,
        };
        let names: Vec<&str> = published
            .package
            .workbook_ref()
            .sheets()
            .iter()
            .map(crate::Sheet::name)
            .collect();
        if names != sheets.iter().map(String::as_str).collect::<Vec<_>>() {
            return Err(Error::Verification);
        }
        Commit::new(source, published)
    }
}

/// The document object's sheet list and sidebar root.
struct Roots {
    sheets: Vec<u64>,
    sidebar: Option<u64>,
}

fn roots(working: &Working<'_>) -> FaultResult<Roots> {
    let document = working.object(1)?;
    let index = structure::document_message(document)?;
    let data = &document.messages[index].data;
    let sheets = references::identifiers(data, &[DOCUMENT_SHEETS_FIELD])?;
    let sidebar = references::identifiers(data, &[DOCUMENT_SIDEBAR_FIELD])?;
    let sidebar = match sidebar.as_slice() {
        [] => None,
        [root] => Some(*root),
        _ => return Err(Fault::Invalid),
    };
    Ok(Roots { sheets, sidebar })
}

/// The sidebar root's child list and the child that addresses `sheet`.
fn sidebar_node(working: &Working<'_>, root: u64, sheet: u64) -> FaultResult<(Vec<u64>, u64)> {
    let object = working.object(root)?;
    let index = structure::unique_message(object, &[TREE_NODE_MESSAGE_TYPE])?;
    let children = references::identifiers(&object.messages[index].data, &[TREE_CHILDREN_FIELD])?;
    let mut found = None;
    for child in &children {
        let node = working.object(*child)?;
        let index = structure::unique_message(node, &[TREE_NODE_MESSAGE_TYPE])?;
        let addressed = references::identifiers(&node.messages[index].data, &[TREE_OBJECT_FIELD])?;
        if addressed.as_slice() == [sheet] && found.replace(*child).is_some() {
            return Err(Fault::Invalid);
        }
    }
    Ok((children, found.ok_or(Fault::Invalid)?))
}

/// Rewrite the reference list at `path` of message `index` of object
/// `identifier` to `after`.
fn rewrite_list(
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

fn add_sheet(source: &Package, name: &str) -> Result<(structure::Published, Vec<String>), Error> {
    let mut working = Working::new(source)?;
    working.keep_previews();
    let roots = roots(&working)?;
    let template = *roots.sheets.first().ok_or(Error::InvalidSource)?;
    if working.locate(1)?.as_deref() != Some(DOCUMENT_COMPONENT)
        || working.locate(template)?.as_deref() != Some(DOCUMENT_COMPONENT)
    {
        return Err(Error::Refused(Refusal::UnsupportedSource));
    }

    // The new sheet record: the template's settings without its private
    // objects, renamed.
    let template_object = working.object(template)?;
    let (index, _) = structure::sheet_message(template_object)?;
    let template_message = &template_object.messages[index];
    let settings = if template_message.type_ == SHEET_MESSAGE_TYPE {
        structure::copy_bytes(&template_message.data)?
    } else {
        let payloads = litchi_iwa_common::wire::repeated_length_delimited_payloads(
            &template_message.data,
            FORM_SHEET_SUPER_FIELD,
        )
        .map_err(|error| structure::map_wire(&error))?;
        let [inner] = payloads.as_slice() else {
            return Err(Error::InvalidSource);
        };
        structure::copy_bytes(inner)?
    };
    let stripped = structure::strip_fields(&settings, &PRIVATE_SHEET_FIELDS)?;
    let renamed =
        patch_length_delimited_field(&stripped, SHEET_NAME_FIELD, true, Some(name.as_bytes()))
            .map_err(|error| structure::map_wire(&error))?;
    let style = references::identifiers(&renamed, &[22])?;

    let fresh = working.allocate(if roots.sidebar.is_some() { 2 } else { 1 })?;
    let sheet = fresh[0];
    let sheet_object = structure::new_object(
        sheet,
        Some(working.object(template)?),
        SHEET_MESSAGE_TYPE,
        renamed,
        style.clone(),
        working.archive_limits(),
    )?;
    working.insert(DOCUMENT_COMPONENT, sheet_object)?;

    // The document's sheet list.
    let document_index = structure::document_message(working.object(1)?)?;
    let mut after = roots.sheets.clone();
    structure::reserve(&mut after, 1)?;
    after.push(sheet);
    rewrite_list(
        &mut working,
        1,
        document_index,
        &[DOCUMENT_SHEETS_FIELD],
        &roots.sheets,
        &after,
    )?;

    // The sidebar order: one tree node addressing the sheet.
    if let Some(root) = roots.sidebar {
        let node = fresh[1];
        if working.locate(root)?.as_deref() != Some(DOCUMENT_COMPONENT) {
            return Err(Error::Refused(Refusal::UnsupportedSource));
        }
        let mut data = Vec::new();
        litchi_iwa_common::wire::append_length_delimited_field(
            &mut data,
            TREE_OBJECT_FIELD,
            &references::reference_payload(sheet)?,
            litchi_iwa_common::WireLimits::DEFAULT,
        )
        .map_err(|error| structure::map_wire(&error))?;
        let node_object = structure::new_object(
            node,
            Some(working.object(root)?),
            TREE_NODE_MESSAGE_TYPE,
            data,
            vec![sheet],
            working.archive_limits(),
        )?;
        working.insert(DOCUMENT_COMPONENT, node_object)?;
        let root_object = working.object(root)?;
        let root_index = structure::unique_message(root_object, &[TREE_NODE_MESSAGE_TYPE])?;
        let children = references::identifiers(
            &root_object.messages[root_index].data,
            &[TREE_CHILDREN_FIELD],
        )?;
        let mut after = children.clone();
        structure::reserve(&mut after, 1)?;
        after.push(node);
        rewrite_list(
            &mut working,
            root,
            root_index,
            &[TREE_CHILDREN_FIELD],
            &children,
            &after,
        )?;
    }

    // The new sheet addresses its style, which may live in another component.
    let mut transition = Transition::default();
    if let Some(registry) = metadata::registry(&working)? {
        let document = registry.component_for_entry(DOCUMENT_COMPONENT)?.identifier;
        for target in &style {
            let component = working.locate(*target)?.ok_or(Error::InvalidSource)?;
            let target_component = registry.component_for_entry(&component)?.identifier;
            transition.reference(&registry, document, target_component, *target)?;
        }
    }
    let published = working.publish(&transition)?;
    let mut sheets: Vec<String> = source
        .workbook_ref()
        .sheets()
        .iter()
        .map(|sheet| sheet.name().to_owned())
        .collect();
    sheets.push(name.to_owned());
    Ok((published, sheets))
}

#[allow(
    clippy::too_many_lines,
    reason = "one sheet copy: the record, its lists, its tables and the registry"
)]
fn duplicate_sheet(
    source: &Package,
    position: usize,
    name: &str,
) -> Result<(structure::Published, Vec<String>), Error> {
    // The root previews go, as the host editor's copy drops them (the app
    // regenerates them).
    let mut working = Working::new(source)?;
    let roots = roots(&working)?;
    let sheet = *roots
        .sheets
        .get(position)
        .ok_or(Error::Refused(Refusal::SheetNotFound))?;
    if working.locate(sheet)?.as_deref() != Some(DOCUMENT_COMPONENT) {
        return Err(Error::Refused(Refusal::UnsupportedSource));
    }
    let sheet_object = working.object(sheet)?.clone();
    let (index, path) = structure::sheet_message(&sheet_object)?;
    let message = &sheet_object.messages[index];
    if message.type_ != SHEET_MESSAGE_TYPE {
        return Err(Error::Refused(Refusal::UnsupportedSource));
    }
    // The copy keeps addressing the source's header, footer and guide
    // storages, as the host editor's copy does (record 0370j gap).
    let drawables = references::identifiers(&message.data, path)?;
    for drawable in &drawables {
        if structure::table_info_message(working.object(*drawable)?).is_err() {
            return Err(Error::Refused(Refusal::SheetNotEmpty));
        }
        if !super::table_lifecycle::duplicate::self_contained(&working, *drawable)? {
            return Err(Error::Refused(Refusal::UnsupportedDependency));
        }
    }
    let table_names: Vec<String> = source
        .workbook_ref()
        .sheets()
        .get(position)
        .ok_or(Error::InvalidSource)?
        .tables()
        .map(|table| table.name().to_owned())
        .collect();
    if table_names.len() != drawables.len() {
        return Err(Error::InvalidSource);
    }

    // The new sheet record: the source's settings, renamed, without its
    // drawables.
    let new_sheet = *working.allocate(1)?.first().ok_or(Error::InvalidSource)?;
    let renamed =
        patch_length_delimited_field(&message.data, SHEET_NAME_FIELD, true, Some(name.as_bytes()))
            .map_err(|error| structure::map_wire(&error))?;
    let emptied = references::rewrite(&renamed, path, &[])?;
    let references: Vec<u64> = sheet_object.archive_info.message_infos[index]
        .object_references
        .iter()
        .copied()
        .filter(|identifier| !drawables.contains(identifier))
        .collect();
    let mut sheet_copy = structure::new_object(
        new_sheet,
        Some(&sheet_object),
        SHEET_MESSAGE_TYPE,
        emptied,
        references.clone(),
        working.archive_limits(),
    )?;
    sheet_copy.archive_info.message_infos[0].versions = sheet_object.archive_info.message_infos
        [index]
        .versions
        .clone();
    sheet_copy.archive_info.message_infos[0].field_infos = sheet_object.archive_info.message_infos
        [index]
        .field_infos
        .iter()
        .filter(|field| field.path.as_slice() != path)
        .cloned()
        .collect();
    working.insert(DOCUMENT_COMPONENT, sheet_copy)?;

    // The document's sheet list, right after the source.
    let document_index = structure::document_message(working.object(1)?)?;
    let mut after = roots.sheets.clone();
    after.insert(position + 1, new_sheet);
    rewrite_list(
        &mut working,
        1,
        document_index,
        &[DOCUMENT_SHEETS_FIELD],
        &roots.sheets,
        &after,
    )?;

    // The sidebar order: a node for the copy, right after the source's.
    let mut node = None;
    if let Some(root) = roots.sidebar {
        let (children, source_node) = sidebar_node(&working, root, sheet)?;
        let new_node = *working.allocate(1)?.first().ok_or(Error::InvalidSource)?;
        let mut data = Vec::new();
        litchi_iwa_common::wire::append_length_delimited_field(
            &mut data,
            TREE_OBJECT_FIELD,
            &references::reference_payload(new_sheet)?,
            litchi_iwa_common::WireLimits::DEFAULT,
        )
        .map_err(|error| structure::map_wire(&error))?;
        let node_object = structure::new_object(
            new_node,
            Some(working.object(root)?),
            TREE_NODE_MESSAGE_TYPE,
            data,
            vec![new_sheet],
            working.archive_limits(),
        )?;
        working.insert(DOCUMENT_COMPONENT, node_object)?;
        let root_index =
            structure::unique_message(working.object(root)?, &[TREE_NODE_MESSAGE_TYPE])?;
        let at = children
            .iter()
            .position(|child| *child == source_node)
            .ok_or(Error::InvalidSource)?;
        let mut after = children.clone();
        after.insert(at + 1, new_node);
        rewrite_list(
            &mut working,
            root,
            root_index,
            &[TREE_CHILDREN_FIELD],
            &children,
            &after,
        )?;
        node = Some(new_node);
    }

    // Its tables, at their sources' positions and under their names.
    let mut transition = Transition::default();
    for (drawable, table_name) in drawables.iter().zip(&table_names) {
        super::table_lifecycle::duplicate::clone_table(
            &mut working,
            &mut transition,
            *drawable,
            super::table_lifecycle::duplicate::Placement {
                sheet: new_sheet,
                source_sheet: sheet,
                offset: 0.0,
            },
            table_name,
        )?;
    }

    // The sheet's registry records: its UUID when the source has one, and
    // its references across components.
    if let Some(registry) = metadata::registry(&working)? {
        let document = registry.component_for_entry(DOCUMENT_COMPONENT)?.identifier;
        if registry.uuids.iter().any(|binding| binding.object == sheet) {
            let seed = metadata::identity_seed(&working, new_sheet)?;
            transition.register_uuid(&registry, document, new_sheet, seed)?;
        }
        for target in references.iter().chain(node.iter()) {
            let component = working.locate(*target)?.ok_or(Error::InvalidSource)?;
            let target_component = registry.component_for_entry(&component)?.identifier;
            transition.reference(&registry, document, target_component, *target)?;
        }
    }
    let published = working.publish(&transition)?;
    let mut sheets: Vec<String> = source
        .workbook_ref()
        .sheets()
        .iter()
        .map(|sheet| sheet.name().to_owned())
        .collect();
    sheets.insert(position + 1, name.to_owned());
    Ok((published, sheets))
}

fn remove_sheet(
    source: &Package,
    position: usize,
) -> Result<(structure::Published, Vec<String>), Error> {
    // Its tables go first, each with its private graph and formula owners.
    let tables = source
        .workbook_ref()
        .sheets()
        .get(position)
        .ok_or(Error::Refused(Refusal::SheetNotFound))?
        .table_count();
    // The composed removal advances the source's save token once.
    let token = Working::new(source)?.save_token()?;
    let mut current = source.clone();
    for _ in 0..tables {
        current = super::table_lifecycle::remove_table(&current, position, 0)
            .map_err(|error| match error {
                super::table_lifecycle::Error::Refused(
                    crate::package::table_lifecycle::Refusal::UnsupportedDependency,
                ) => Error::Refused(Refusal::UnsupportedDependency),
                super::table_lifecycle::Error::Refused(
                    crate::package::table_lifecycle::Refusal::UnsupportedSource,
                ) => Error::Refused(Refusal::UnsupportedSource),
                super::table_lifecycle::Error::Limit {
                    observed, maximum, ..
                } => Error::Limit {
                    kind: LimitKind::PayloadBytes,
                    observed,
                    maximum,
                },
                super::table_lifecycle::Error::Allocation { amount } => {
                    Error::Allocation { amount }
                },
                super::table_lifecycle::Error::Verification => Error::Verification,
                _ => Error::InvalidSource,
            })?
            .package;
    }
    let token = if tables > 0 {
        token.and_then(|token| token.checked_add(1))
    } else {
        None
    };
    let mut published = remove_empty_sheet(&current, position, token)?;
    if tables > 0 {
        // The patch spans every step; report the touched components of the
        // last one only when nothing else changed.
        published.touched_components = published.touched_components.max(1);
    }
    let sheets: Vec<String> = source
        .workbook_ref()
        .sheets()
        .iter()
        .enumerate()
        .filter(|(index, _)| *index != position)
        .map(|(_, sheet)| sheet.name().to_owned())
        .collect();
    Ok((published, sheets))
}

fn remove_empty_sheet(
    source: &Package,
    position: usize,
    save_token: Option<u64>,
) -> Result<structure::Published, Error> {
    let mut working = Working::new(source)?;
    working.keep_previews();
    if let Some(token) = save_token {
        working.set_save_token(token);
    }
    let roots = roots(&working)?;
    let sheet = *roots
        .sheets
        .get(position)
        .ok_or(Error::Refused(Refusal::SheetNotFound))?;
    if roots.sheets.len() <= 1 {
        return Err(Error::Refused(Refusal::LastSheet));
    }
    let sheet_object = working.object(sheet)?;
    let (index, path) = structure::sheet_message(sheet_object)?;
    if !references::identifiers(&sheet_object.messages[index].data, path)?.is_empty() {
        return Err(Error::Refused(Refusal::SheetNotEmpty));
    }

    // Detach the sheet from the document and its node from the sidebar.
    let mut detached = vec![(1, sheet)];
    let mut roots_removed = vec![sheet];
    let sidebar = match roots.sidebar {
        Some(root) => {
            let (children, node) = sidebar_node(&working, root, sheet)?;
            detached.push((root, node));
            roots_removed.push(node);
            Some((root, children, node))
        },
        None => None,
    };
    let census = Census::new(&working)?;
    let closure = census.private_closure(&working, &roots_removed)?;
    census.require_detached(&closure, &detached)?;

    let document_index = structure::document_message(working.object(1)?)?;
    let after: Vec<u64> = roots
        .sheets
        .iter()
        .copied()
        .filter(|item| *item != sheet)
        .collect();
    rewrite_list(
        &mut working,
        1,
        document_index,
        &[DOCUMENT_SHEETS_FIELD],
        &roots.sheets,
        &after,
    )?;
    if let Some((root, children, node)) = sidebar {
        let root_index =
            structure::unique_message(working.object(root)?, &[TREE_NODE_MESSAGE_TYPE])?;
        let after: Vec<u64> = children
            .iter()
            .copied()
            .filter(|item| *item != node)
            .collect();
        rewrite_list(
            &mut working,
            root,
            root_index,
            &[TREE_CHILDREN_FIELD],
            &children,
            &after,
        )?;
    }
    let removed: Vec<u64> = closure.iter().copied().collect();
    let mut transition = Transition::default();
    if let Some(registry) = metadata::registry(&working)? {
        transition.remove_objects(&registry, &removed)?;
    }
    // A removed suffix of the identifier space lowers the watermark, as the
    // host editor releases it (record 0370g).
    transition.release_suffix(&removed)?;
    for identifier in &removed {
        working.remove(*identifier)?;
    }
    let published = working.publish(&transition)?;
    verify_removed(&published.package, &closure)?;
    Ok(published)
}

impl Package {
    /// Start an immutable single-operation sheet add or remove edit.
    ///
    /// ```no_run
    /// use litchi_numbers::Package;
    ///
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let package = Package::new();
    /// let commit = package.edit_sheet_lifecycle().add("Summary")?.commit()?;
    /// assert_eq!(commit.package().sheets().len(), 2);
    /// let restored = commit
    ///     .package()
    ///     .apply_sheet_lifecycle(&commit.patch().inverse())?;
    /// assert_eq!(restored.package().sheets().len(), 1);
    /// # Ok(())
    /// # }
    /// ```
    #[must_use]
    pub const fn edit_sheet_lifecycle(&self) -> Edit<'_> {
        Edit {
            source: self,
            operation: None,
        }
    }

    /// Apply a sheet-lifecycle patch to its exact source package.
    ///
    /// # Errors
    ///
    /// [`Error::PatchConflict`] when this package is not the patch's exact
    /// source; [`Error::Verification`] when its retained target no longer
    /// opens.
    pub fn apply_sheet_lifecycle(&self, patch: &Patch) -> Result<Commit, Error> {
        apply_patch(self, patch)
    }
}
