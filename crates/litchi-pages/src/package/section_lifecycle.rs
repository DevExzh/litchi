//! Implementation of the public [`crate::section::lifecycle`] transactions:
//! adding a section and removing one (`IWORK_API_REQUESTS.md` request 11,
//! record 0312).
//!
//! A Pages body is one `TSWP.StorageArchive` whose section table
//! (`table_section`, field 17) maps a UTF-16 boundary to each
//! `TP.SectionArchive`; every section but the first begins after a U+0004
//! section-break character. The native editor's section operations are
//! ported here:
//!
//! - **Add** inserts a section after another one: a U+0004 break at the end
//!   of that section's text, and a section-table entry just after it for a
//!   new section whose private graph is an independent copy of the source
//!   section's (its page masters, headers and footers, user guide map and
//!   first guide storage, as Pages copies them). The copies get fresh
//!   identifiers and UUIDs; shared records (styles) stay shared. A section
//!   whose graph holds anything else (inline drawables in a header, data
//!   references) is refused.
//! - **Remove** removes a non-initial section's boundary and its private
//!   graph; its text is kept and joins the preceding section's, as in the
//!   native editor.
//!
//! Only the body storage's component and the registry are rewritten; every
//! other member keeps its exact bytes, and the patch's inverse restores the
//! source byte for byte.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "the focused transaction keeps its public types beside the staged rewrite"
)]

use std::collections::{BTreeSet, HashMap};
use std::fmt;

use litchi_core::Position;
use litchi_iwa_common::wire::{
    append_length_delimited_field, append_varint_field, parse_wire_view,
    patch_length_delimited_field, repeated_length_delimited_payloads, repeated_varint_values,
    rewrite_repeated_length_delimited_fields, transform_length_delimited_field,
};
use litchi_iwa_core::RawMessage;
use litchi_iwa_text::TextPosition;
use thiserror::Error as ThisError;

use super::Package;
use super::section_text::{native_body, native_section_range};
use super::structure::{
    self, Artifacts, Census, Fault, FaultLimit, FaultResult, Transition, WireFault, Working,
    copy_bytes, map_wire, reference_payload, remap_at, reserve, unique_message,
};
use crate::SectionSelector;

/// `TP.SectionArchive`.
const SECTION_MESSAGE_TYPE: u32 = 10_011;
/// `TP.SectionTemplateArchive`.
const SECTION_TEMPLATE_MESSAGE_TYPE: u32 = 10_143;
/// `TP.UserDefinedGuideMapArchive`.
const GUIDE_MAP_MESSAGE_TYPE: u32 = 10_016;
/// `TSD.GuideStorageArchive`.
const GUIDE_STORAGE_MESSAGE_TYPE: u32 = 3_047;
/// `TSWP.StorageArchive` and `TSWP.StorageArchive` with paragraph data.
const STORAGE_MESSAGE_TYPES: [u32; 2] = [2_001, 2_022];
/// The reference fields of `TP.SectionArchive`: the obsolete headers and
/// footers, the first, even and odd page templates, the user guide map.
const SECTION_REFERENCE_PATHS: [&[u32]; 7] = [&[3], &[4], &[14], &[23], &[24], &[25], &[29]];
/// `TP.SectionArchive.name`.
const SECTION_NAME_FIELD: u32 = 26;
/// The reference fields of `TP.SectionTemplateArchive`: headers, footers,
/// drawables.
const TEMPLATE_REFERENCE_PATHS: [&[u32]; 3] = [&[1], &[2], &[3]];
/// `TP.UserDefinedGuideMapArchive.user_defined_guide_storages`.
const GUIDE_MAP_ENTRY_FIELD: u32 = 1;
/// `UserDefinedGuideStorage.page_index` and `.guide_storage`.
const GUIDE_ENTRY_PAGE_FIELD: u32 = 1;
const GUIDE_ENTRY_STORAGE_FIELD: u32 = 2;
/// `TSD.GuideStorageArchive.user_defined_guides`.
const GUIDE_STORAGE_GUIDES_FIELD: u32 = 1;
/// `TSWP.StorageArchive.table_section`.
const SECTION_TABLE_FIELD: u32 = 17;
/// `TSWP.ObjectAttributeTable.entries`, and the entry's fields.
const TABLE_ENTRY_FIELD: u32 = 1;
const ENTRY_INDEX_FIELD: u32 = 1;
const ENTRY_OBJECT_FIELD: u32 = 2;
/// The native section-break character.
const SECTION_BREAK: &str = "\u{4}";

/// A finite resource governed by the transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Complete source or candidate package bytes.
    PackageBytes,
    /// Bytes in one decoded component.
    PayloadBytes,
    /// Native object identifiers.
    Identifiers,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::PackageBytes => "package bytes",
            Self::PayloadBytes => "payload bytes",
            Self::Identifiers => "object identifiers",
        })
    }
}

/// A typed, content-redacted section-lifecycle failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies it
    /// (record 0351 §2.4, record 0361).
    #[error(transparent)]
    Refused(Refusal),
    /// The rooted source is malformed or ambiguous.
    #[error("rooted section-lifecycle source is invalid")]
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
    #[error("could not allocate {amount} section-lifecycle items or bytes")]
    Allocation {
        /// Requested amount.
        amount: usize,
    },
    /// The reopened candidate did not reproduce the requested state.
    #[error("section-lifecycle candidate verification failed")]
    Verification,
    /// The patch was applied to an artifact other than its exact source.
    #[error("section-lifecycle patch conflicts with this package")]
    PatchConflict,
}

/// A well-formed request this family's route refuses, classified by
/// [`Refusal::kind`] (record 0351 §2.4, record 0361).
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    /// The selector did not resolve one section.
    #[error("section selector did not resolve")]
    SectionNotFound,
    /// The break position is 0, beyond the body, right after a section
    /// break, or where a section already starts.
    #[error("the section break position is invalid")]
    InvalidBreak,
    /// The requested section name contains NUL.
    #[error("Pages section names cannot contain NUL")]
    InvalidName,
    /// This edit already holds its single operation.
    #[error("a section-lifecycle operation is already staged")]
    OperationAlreadyStaged,
    /// Commit was requested before staging an operation.
    #[error("no section-lifecycle operation is staged")]
    NoStagedOperation,
    /// The first section begins the body; its boundary cannot be removed.
    #[error("the first section of a document cannot be removed")]
    InitialSection,
    /// A section break must follow body content: a section cannot be added
    /// after an empty one.
    #[error("a section cannot be added after an empty section")]
    EmptySection,
    /// The source section's private graph holds content this transaction
    /// does not copy (inline drawables, media, unknown records).
    #[error("the section's private graph holds content that is not copied")]
    UnsupportedContent,
    /// A surviving object still depends on the section's graph.
    #[error("a surviving object depends on the section's graph")]
    UnsupportedDependency,
    /// The source cannot publish an exact changed artifact.
    #[error("source representation does not support section lifecycle changes")]
    UnsupportedSource,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::SectionNotFound { .. } => RefusalKind::NotFound,
            Self::InvalidName { .. } | Self::InvalidBreak => RefusalKind::InvalidValue,
            Self::OperationAlreadyStaged { .. } | Self::NoStagedOperation { .. } => {
                RefusalKind::UnsupportedOperation
            },
            Self::InitialSection { .. } | Self::EmptySection { .. } => {
                RefusalKind::UnsupportedTarget
            },
            Self::UnsupportedContent { .. }
            | Self::UnsupportedDependency { .. }
            | Self::UnsupportedSource { .. } => RefusalKind::UnsupportedContent,
        }
    }
}

impl From<Refusal> for Error {
    fn from(refusal: Refusal) -> Self {
        Self::Refused(refusal)
    }
}

impl From<Fault> for Error {
    fn from(fault: Fault) -> Self {
        match fault {
            Fault::Unsupported => Self::Refused(Refusal::UnsupportedSource),
            Fault::Invalid => Self::InvalidSource,
            Fault::Dependency => Self::Refused(Refusal::UnsupportedDependency),
            Fault::Limit {
                kind,
                observed,
                maximum,
            } => Self::Limit {
                kind: match kind {
                    FaultLimit::PackageBytes => LimitKind::PackageBytes,
                    FaultLimit::PayloadBytes => LimitKind::PayloadBytes,
                    FaultLimit::Identifiers => LimitKind::Identifiers,
                },
                observed,
                maximum,
            },
            Fault::Allocation { amount } => Self::Allocation { amount },
            Fault::Verification => Self::Verification,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum Operation {
    /// Add a section after the section at this position, or, with `at`,
    /// break the body at that UTF-16 position into a new section copying
    /// the section at `after`.
    Add {
        after: usize,
        at: Option<usize>,
        name: Option<String>,
    },
    /// Remove the section at this position.
    Remove(usize),
}

/// One immutable edit that stages exactly one section addition or removal.
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
                    Operation::Add { .. } => "add",
                    Operation::Remove(_) => "remove",
                }),
            )
            .finish()
    }
}

/// A reversible, exact-source section-lifecycle patch.
///
/// The patch retains the exact source and target artifacts; it authorizes
/// application only to its source, and its inverse restores the source byte
/// for byte.
#[derive(Clone, PartialEq, Eq)]
pub struct Patch {
    artifacts: Artifacts,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.debug_struct("Patch").finish_non_exhaustive()
    }
}

impl Patch {
    /// Whether the patch is an exact artifact no-op.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.artifacts.is_noop()
    }

    /// The exact target-to-source inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
        }
    }
}

/// Content-free diagnostics for one section-lifecycle publication.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct Diagnostics {
    changed: bool,
    touched_components: usize,
    objects: usize,
}

impl Diagnostics {
    /// Whether the exact package artifact changed.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// Number of rewritten native components.
    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }

    /// Number of native objects created (add) or removed (remove).
    #[must_use]
    pub const fn objects(self) -> usize {
        self.objects
    }
}

/// One validated section-lifecycle publication.
#[derive(Debug)]
pub struct Commit {
    package: Package,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
    /// Borrow the validated package snapshot.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Consume this publication and return its package snapshot.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Consume this publication and return its package and patch.
    #[must_use]
    pub fn into_parts(self) -> (Package, Patch) {
        (self.package, self.patch)
    }

    /// Borrow the reversible exact-source patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Borrow content-free diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }
}

impl Edit<'_> {
    /// Stage a new section after the selected one, copying its page masters,
    /// headers and footers; the new section starts empty. `name` names it
    /// (`None`: unnamed).
    ///
    /// # Errors
    ///
    /// [`Refusal::SectionNotFound`] when the selector resolves no section,
    /// [`Refusal::InvalidName`] for a name containing NUL,
    /// [`Refusal::OperationAlreadyStaged`] after one operation.
    #[allow(
        clippy::should_implement_trait,
        reason = "`add` is the glossary verb for staging a new element, not arithmetic"
    )]
    pub fn add<'selector>(
        mut self,
        after: impl Into<SectionSelector<'selector>>,
        name: Option<&str>,
    ) -> Result<Self, Error> {
        if self.operation.is_some() {
            return Err(Error::Refused(Refusal::OperationAlreadyStaged));
        }
        if name.is_some_and(|name| name.contains('\0')) {
            return Err(Error::Refused(Refusal::InvalidName));
        }
        let after = resolve(self.source, after.into())?;
        let name = match name {
            Some(name) => {
                let mut owned = String::new();
                owned
                    .try_reserve_exact(name.len())
                    .map_err(|_error| Error::Allocation { amount: name.len() })?;
                owned.push_str(name);
                Some(owned)
            },
            None => None,
        };
        self.operation = Some(Operation::Add {
            after,
            at: None,
            name,
        });
        Ok(self)
    }

    /// Stage a section break at UTF-16 position `at` of the body (counted
    /// over the whole body text, section breaks included): the text after
    /// `at` starts a new section that copies the page masters, headers and
    /// footers of the `source` section. `name` names it (`None`: unnamed).
    ///
    /// # Errors
    ///
    /// [`Refusal::SectionNotFound`] when the selector resolves no section,
    /// [`Refusal::InvalidName`] for a name containing NUL,
    /// [`Refusal::InvalidBreak`] for position 0, a position beyond the body,
    /// one right after a section break or one where a section already
    /// starts, [`Refusal::OperationAlreadyStaged`] after one operation.
    pub fn insert<'selector>(
        mut self,
        source: impl Into<SectionSelector<'selector>>,
        at: TextPosition,
        name: Option<&str>,
    ) -> Result<Self, Error> {
        if self.operation.is_some() {
            return Err(Error::Refused(Refusal::OperationAlreadyStaged));
        }
        if name.is_some_and(|name| name.contains('\0')) {
            return Err(Error::Refused(Refusal::InvalidName));
        }
        let after = resolve(self.source, source.into())?;
        let at = at.utf16_index() as usize;
        let body = native_body(self.source).map_err(|_error| Error::InvalidSource)?;
        let (storage, _) = super::body_storage_objects::body_storage(self.source)
            .map_err(|_error| Error::InvalidSource)?;
        let units = storage.text.encode_utf16().collect::<Vec<_>>();
        if at == 0
            || at > units.len()
            || units.get(at - 1) == Some(&0x0004)
            || (0..self.source.sections().len())
                .any(|position| body.section_boundary(position) == Some(at))
        {
            return Err(Error::Refused(Refusal::InvalidBreak));
        }
        self.operation = Some(Operation::Add {
            after,
            at: Some(at),
            name: name.map(str::to_owned),
        });
        Ok(self)
    }

    /// Stage the removal of the selected section: its boundary and private
    /// graph go, its text joins the preceding section's.
    ///
    /// # Errors
    ///
    /// [`Refusal::SectionNotFound`] when the selector resolves no section,
    /// [`Refusal::InitialSection`] for the first section,
    /// [`Refusal::OperationAlreadyStaged`] after one operation.
    pub fn remove<'selector>(
        mut self,
        section: impl Into<SectionSelector<'selector>>,
    ) -> Result<Self, Error> {
        if self.operation.is_some() {
            return Err(Error::Refused(Refusal::OperationAlreadyStaged));
        }
        let position = resolve(self.source, section.into())?;
        if position == 0 {
            return Err(Error::Refused(Refusal::InitialSection));
        }
        self.operation = Some(Operation::Remove(position));
        Ok(self)
    }

    /// Validate and atomically publish the staged candidate.
    ///
    /// # Errors
    ///
    /// [`Refusal::NoStagedOperation`] for an empty edit;
    /// [`Refusal::EmptySection`], [`Refusal::UnsupportedContent`],
    /// [`Refusal::UnsupportedDependency`]; source, resource and verification
    /// errors. Failure never publishes a partial package.
    pub fn commit(self) -> Result<Commit, Error> {
        let operation = self
            .operation
            .ok_or(Error::Refused(Refusal::NoStagedOperation))?;
        let (package, touched_components, objects) = match &operation {
            Operation::Add { after, at, name } => {
                add_section(self.source, *after, *at, name.as_deref())?
            },
            Operation::Remove(position) => remove_section(self.source, *position)?,
        };
        Ok(Commit {
            patch: Patch {
                artifacts: Artifacts::new(self.source, &package),
            },
            package,
            diagnostics: Diagnostics {
                changed: true,
                touched_components,
                objects,
            },
        })
    }
}

fn resolve(package: &Package, selector: SectionSelector<'_>) -> Result<usize, Error> {
    Ok(package
        .document_ref()
        .select_section(selector)
        .map_err(|_error| Error::Refused(Refusal::SectionNotFound))?
        .ok_or(Error::Refused(Refusal::SectionNotFound))?
        .index())
}

/// The body storage's native section table: `(boundary, section)` in table
/// order (the initial section the document root names is not listed).
fn section_entries(payload: &[u8]) -> FaultResult<Vec<(u64, u64, Vec<u8>)>> {
    let tables = repeated_length_delimited_payloads(payload, SECTION_TABLE_FIELD)
        .map_err(|error| map_wire(&error))?;
    let table = match tables.as_slice() {
        [] => return Ok(Vec::new()),
        [table] => *table,
        _ => return Err(Fault::Invalid),
    };
    let records = repeated_length_delimited_payloads(table, TABLE_ENTRY_FIELD)
        .map_err(|error| map_wire(&error))?;
    let mut entries = Vec::new();
    reserve(&mut entries, records.len())?;
    for record in records {
        let index =
            repeated_varint_values(record, ENTRY_INDEX_FIELD).map_err(|error| map_wire(&error))?;
        let object = structure::identifiers(record, ENTRY_OBJECT_FIELD)?;
        let ([index], [object]) = (index.as_slice(), object.as_slice()) else {
            return Err(Fault::Unsupported);
        };
        entries.push((*index, *object, copy_bytes(record)?));
    }
    Ok(entries)
}

/// The body payload with its section table holding exactly `entries`.
fn with_section_entries(payload: &[u8], entries: &[Vec<u8>]) -> FaultResult<Vec<u8>> {
    let present = parse_wire_view(payload, litchi_iwa_common::WireLimits::DEFAULT)
        .map_err(|error| map_wire(&error))?
        .fields()
        .any(|field| field.number() == SECTION_TABLE_FIELD);
    if present {
        return transform_length_delimited_field(payload, SECTION_TABLE_FIELD, |table| {
            rewrite_repeated_length_delimited_fields(table, TABLE_ENTRY_FIELD, entries)
                .map_err(WireFault::from)
        })
        .map_err(|error: WireFault| error.0);
    }
    let mut table = Vec::new();
    for entry in entries {
        append_length_delimited_field(
            &mut table,
            TABLE_ENTRY_FIELD,
            entry,
            litchi_iwa_common::WireLimits::DEFAULT,
        )
        .map_err(|error| map_wire(&error))?;
    }
    patch_length_delimited_field(payload, SECTION_TABLE_FIELD, false, Some(&table))
        .map_err(|error| map_wire(&error))
}

/// A canonical section-table entry.
fn section_entry(boundary: u64, section: u64) -> FaultResult<Vec<u8>> {
    let mut entry = Vec::new();
    append_varint_field(&mut entry, ENTRY_INDEX_FIELD, boundary)
        .map_err(|error| map_wire(&error))?;
    append_length_delimited_field(
        &mut entry,
        ENTRY_OBJECT_FIELD,
        &reference_payload(section)?,
        litchi_iwa_common::WireLimits::DEFAULT,
    )
    .map_err(|error| map_wire(&error))?;
    Ok(entry)
}

/// Splice `replacement` over the UTF-16 `range` of the body storage text.
fn rewrite_body_text(
    source: &Package,
    payload: &[u8],
    range: std::ops::Range<usize>,
    replacement: &str,
) -> Result<Vec<u8>, Error> {
    let limits = super::storage_rewrite_limits(source.state.source.limits())
        .map_err(|_error| Error::InvalidSource)?;
    let rewritten = litchi_iwa_text_wire::rewrite_storage_text(payload, range, replacement, limits)
        .map_err(|_error| Error::Refused(Refusal::UnsupportedSource))?;
    if !rewritten.removed_object_references().is_empty()
        || !rewritten.removed_object_references_by_field().is_empty()
        || !rewritten.changed()
    {
        return Err(Error::Refused(Refusal::UnsupportedDependency));
    }
    Ok(rewritten.into_bytes())
}

/// The body text's UTF-16 code unit at `index`.
fn body_unit(source: &Package, payload: &[u8], index: usize) -> Result<Option<u16>, Error> {
    let limits = super::storage_rewrite_limits(source.state.source.limits())
        .map_err(|_error| Error::InvalidSource)?;
    let storage = litchi_iwa_text_wire::decode_storage(payload, limits)
        .map_err(|_error| Error::InvalidSource)?;
    Ok(storage.storage().text().encode_utf16().nth(index))
}

/// The body storage object's identifier, component and storage message.
fn body_storage(working: &Working<'_>, identifier: u64) -> Result<(String, usize), Error> {
    let component = working.locate(identifier)?.ok_or(Error::InvalidSource)?;
    let index = unique_message(working.object(identifier)?, &STORAGE_MESSAGE_TYPES)?;
    Ok((component, index))
}

#[allow(
    clippy::too_many_lines,
    reason = "one staged publication: the graph copy, the break, the table, the registry"
)]
fn add_section(
    source: &Package,
    after: usize,
    at: Option<usize>,
    name: Option<&str>,
) -> Result<(Package, usize, usize), Error> {
    let body = native_body(source).map_err(|_error| Error::InvalidSource)?;
    let range = native_section_range(source, &body, Position::new(after))
        .map_err(|_error| Error::InvalidSource)?;
    if at.is_none() && range.is_empty() {
        return Err(Error::Refused(Refusal::EmptySection));
    }
    let break_at = at.unwrap_or(range.end);
    let section = body
        .section_identifier(after)
        .ok_or(Error::Refused(Refusal::UnsupportedSource))?;
    let mut working = Working::new(source)?;
    let (component, storage_index) = body_storage(&working, body.identifier.get())?;
    if working.locate(section)?.as_deref() != Some(component.as_str()) {
        return Err(Error::Refused(Refusal::UnsupportedSource));
    }

    // The source section's private graph and what of it is copied: every
    // record but the guide storages beyond the first.
    let census = Census::new(source)?;
    let closure = census.private_closure(&[section])?;
    let mut skipped = BTreeSet::new();
    for member in &closure {
        let object = working.object(*member)?;
        if object.messages.len() != 1 {
            return Err(Error::Refused(Refusal::UnsupportedContent));
        }
        let info = object
            .archive_info
            .message_infos
            .first()
            .ok_or(Error::InvalidSource)?;
        if !info.data_references.is_empty() {
            return Err(Error::Refused(Refusal::UnsupportedContent));
        }
        if working.locate(*member)?.as_deref() != Some(component.as_str()) {
            return Err(Error::Refused(Refusal::UnsupportedSource));
        }
        if object.messages[0].type_ == GUIDE_MAP_MESSAGE_TYPE {
            let entries =
                repeated_length_delimited_payloads(&object.messages[0].data, GUIDE_MAP_ENTRY_FIELD)
                    .map_err(|error| map_wire(&error))?;
            for entry in entries.iter().skip(1) {
                for guide in structure::identifiers(entry, GUIDE_ENTRY_STORAGE_FIELD)? {
                    skipped.insert(guide);
                }
            }
        }
    }
    // The copies take fresh identifiers in the native editor's order: the
    // section, its templates, their headers and footers, the guide map and
    // its storage, each group in identifier order.
    let mut members: Vec<(u8, u64)> = Vec::new();
    for member in closure.iter().filter(|member| !skipped.contains(member)) {
        let group = match working.object(*member)?.messages[0].type_ {
            SECTION_MESSAGE_TYPE => 0,
            SECTION_TEMPLATE_MESSAGE_TYPE => 1,
            GUIDE_MAP_MESSAGE_TYPE => 3,
            GUIDE_STORAGE_MESSAGE_TYPE => 4,
            _ => 2,
        };
        reserve(&mut members, 1)?;
        members.push((if *member == section { 0 } else { group.max(1) }, *member));
    }
    members.sort_unstable();
    let members: Vec<u64> = members.into_iter().map(|(_, member)| member).collect();
    for member in &members {
        let message_type = working.object(*member)?.messages[0].type_;
        let storage = STORAGE_MESSAGE_TYPES.contains(&message_type);
        if !(storage
            || matches!(
                message_type,
                SECTION_MESSAGE_TYPE
                    | SECTION_TEMPLATE_MESSAGE_TYPE
                    | GUIDE_MAP_MESSAGE_TYPE
                    | GUIDE_STORAGE_MESSAGE_TYPE
            ))
        {
            return Err(Error::Refused(Refusal::UnsupportedContent));
        }
        // A header or footer storage owning records of its own (inline
        // drawables, fields) would need them copied too.
        if storage
            && census
                .targets(*member)
                .iter()
                .any(|target| closure.contains(target))
        {
            return Err(Error::Refused(Refusal::UnsupportedContent));
        }
        if *member != section && message_type == SECTION_MESSAGE_TYPE {
            return Err(Error::Refused(Refusal::UnsupportedContent));
        }
        // The section's page templates, a template's headers and footers
        // and the guide map's storages are private to the section: a copy
        // never shares one with its source.
        if !storage
            && message_type != GUIDE_STORAGE_MESSAGE_TYPE
            && census
                .targets(*member)
                .iter()
                .any(|target| !closure.contains(target))
        {
            return Err(Error::Refused(Refusal::UnsupportedContent));
        }
    }

    // Fresh identities.
    let fresh = working.allocate(members.len())?;
    let remap: HashMap<u64, u64> = members.iter().copied().zip(fresh.iter().copied()).collect();
    let pairs: Vec<(u64, u64)> = members.iter().copied().zip(fresh.iter().copied()).collect();
    let limits = working.archive_limits();
    for (member, identifier) in &pairs {
        let object = working.object(*member)?;
        let message = &object.messages[0];
        let mut stale: Vec<u64> = Vec::new();
        let data = match message.type_ {
            SECTION_MESSAGE_TYPE => {
                let mut data = copy_bytes(&message.data)?;
                for path in SECTION_REFERENCE_PATHS {
                    data = remap_at(&data, path, &remap)?;
                }
                let present = parse_wire_view(&data, litchi_iwa_common::WireLimits::DEFAULT)
                    .map_err(|error| map_wire(&error))?
                    .fields()
                    .any(|field| field.number() == SECTION_NAME_FIELD);
                patch_length_delimited_field(
                    &data,
                    SECTION_NAME_FIELD,
                    present,
                    name.map(str::as_bytes),
                )
                .map_err(|error| map_wire(&error))?
            },
            SECTION_TEMPLATE_MESSAGE_TYPE => {
                let mut data = copy_bytes(&message.data)?;
                for path in TEMPLATE_REFERENCE_PATHS {
                    data = remap_at(&data, path, &remap)?;
                }
                data
            },
            GUIDE_MAP_MESSAGE_TYPE => {
                // The copy keeps the first page's guides, renumbered to the
                // copy's first page.
                let entries =
                    repeated_length_delimited_payloads(&message.data, GUIDE_MAP_ENTRY_FIELD)
                        .map_err(|error| map_wire(&error))?;
                let mut replacements = Vec::new();
                if let Some(entry) = entries.first() {
                    let present = parse_wire_view(entry, litchi_iwa_common::WireLimits::DEFAULT)
                        .map_err(|error| map_wire(&error))?
                        .fields()
                        .any(|field| field.number() == GUIDE_ENTRY_PAGE_FIELD);
                    let entry = litchi_iwa_common::wire::patch_varint_field(
                        entry,
                        GUIDE_ENTRY_PAGE_FIELD,
                        present,
                        Some(0),
                    )
                    .map_err(|error| map_wire(&error))?;
                    replacements.push(remap_at(&entry, &[GUIDE_ENTRY_STORAGE_FIELD], &remap)?);
                }
                stale.extend(skipped.iter().copied());
                rewrite_repeated_length_delimited_fields(
                    &message.data,
                    GUIDE_MAP_ENTRY_FIELD,
                    &replacements,
                )
                .map_err(|error| map_wire(&error))?
            },
            GUIDE_STORAGE_MESSAGE_TYPE => rewrite_repeated_length_delimited_fields(
                &message.data,
                GUIDE_STORAGE_GUIDES_FIELD,
                &[],
            )
            .map_err(|error| map_wire(&error))?,
            _ => copy_bytes(&message.data)?,
        };
        let message_type = message.type_;
        let mut cloned = object
            .clone_with_identity_remap(
                *identifier,
                &pairs,
                &[RawMessage {
                    type_: message_type,
                    data: data.clone(),
                }],
                limits,
            )
            .map_err(|error| structure::map_core(&error))?;
        if !stale.is_empty() {
            structure::replace(&mut cloned, 0, data, &stale, &[], limits)?;
        }
        working.insert(&component, cloned)?;
    }
    let created = remap[&section];

    // The break at the end of the source section and the new boundary.
    let payload = copy_bytes(&working.object(body.identifier.get())?.messages[storage_index].data)?;
    let payload = rewrite_body_text(source, &payload, break_at..break_at, SECTION_BREAK)?;
    let boundary = u64::try_from(break_at + 1).map_err(|_error| Error::InvalidSource)?;
    let mut entries: Vec<(u64, Vec<u8>)> = section_entries(&payload)?
        .into_iter()
        .map(|(index, _, record)| (index, record))
        .collect();
    if entries.iter().any(|(index, _)| *index == boundary) {
        return Err(Error::InvalidSource);
    }
    let entry_position = entries
        .iter()
        .position(|(index, _)| *index > boundary)
        .unwrap_or(entries.len());
    entries.insert(
        entry_position,
        (boundary, section_entry(boundary, created)?),
    );
    let records: Vec<Vec<u8>> = entries.into_iter().map(|(_, record)| record).collect();
    let payload = with_section_entries(&payload, &records)?;
    structure::replace(
        working.object_mut(body.identifier.get())?,
        storage_index,
        payload,
        &[],
        &[created],
        limits,
    )?;

    // The copies' UUIDs.
    let mut transition = Transition::default();
    if let Some(registry) = working.registry()? {
        let seed = working.identity_seed(fresh[0])?;
        for (member, identifier) in &pairs {
            if let Some(binding) = registry
                .uuids
                .iter()
                .find(|binding| binding.object == *member)
            {
                transition.register_uuid(&registry, binding.component, *identifier, seed)?;
            }
        }
    }
    let (candidate, touched) = working.publish(&transition)?;

    // The new section reads back empty and named, after its source; every
    // other section keeps its text.
    let before = source.sections();
    let now = candidate.sections();
    if now.len() != before.len() + 1 {
        return Err(Error::Verification);
    }
    if at.is_some() {
        // A split: the new section starts after the break and is named; the
        // sections' text, read in order, is unchanged.
        let joined = |sections: &[crate::Section]| {
            sections
                .iter()
                .map(|section| section.body_text().unwrap_or_default())
                .collect::<String>()
        };
        let body = native_body(&candidate).map_err(|_error| Error::Verification)?;
        let position = (0..now.len())
            .find(|position| body.section_identifier(*position) == Some(remap[&section]))
            .ok_or(Error::Verification)?;
        if now[position].name() != name || joined(now) != joined(before) {
            return Err(Error::Verification);
        }
        return Ok((candidate, touched, members.len()));
    }
    let added = &now[after + 1];
    if !added.body_text().unwrap_or_default().is_empty() || added.name() != name {
        return Err(Error::Verification);
    }
    for (index, section) in before.iter().enumerate() {
        let shifted = if index <= after { index } else { index + 1 };
        if now[shifted].body_text() != section.body_text() || now[shifted].name() != section.name()
        {
            return Err(Error::Verification);
        }
    }
    Ok((candidate, touched, members.len()))
}

fn remove_section(source: &Package, position: usize) -> Result<(Package, usize, usize), Error> {
    let body = native_body(source).map_err(|_error| Error::InvalidSource)?;
    let section = body
        .section_identifier(position)
        .ok_or(Error::Refused(Refusal::UnsupportedSource))?;
    let boundary = body
        .section_boundary(position)
        .ok_or(Error::Refused(Refusal::UnsupportedSource))?;
    if boundary == 0 {
        return Err(Error::Refused(Refusal::InitialSection));
    }
    let mut working = Working::new(source)?;
    let (_, storage_index) = body_storage(&working, body.identifier.get())?;

    let census = Census::new(source)?;
    let closure = census.private_closure(&[section])?;
    census.require_detached(&closure, &[(body.identifier.get(), section)])?;
    let removed: Vec<u64> = closure.iter().copied().collect();
    let mut transition = Transition::default();
    if let Some(registry) = working.registry()? {
        transition
            .remove_objects(&registry, &removed)
            .map_err(|fault| match fault {
                Fault::Unsupported => Error::Refused(Refusal::UnsupportedContent),
                other => other.into(),
            })?;
    }

    // The boundary leaves the table, then the break leaves the text.
    let payload = copy_bytes(&working.object(body.identifier.get())?.messages[storage_index].data)?;
    let entries = section_entries(&payload)?;
    if entries
        .iter()
        .filter(|(_, object, _)| *object == section)
        .count()
        != 1
    {
        return Err(Error::Refused(Refusal::UnsupportedSource));
    }
    let records: Vec<Vec<u8>> = entries
        .into_iter()
        .filter(|(_, object, _)| *object != section)
        .map(|(_, _, record)| record)
        .collect();
    let mut payload = with_section_entries(&payload, &records)?;
    let marker = boundary - 1;
    if body_unit(source, &payload, marker)? == Some(0x0004) {
        payload = rewrite_body_text(source, &payload, marker..boundary, "")?;
    }
    let limits = working.archive_limits();
    structure::replace(
        working.object_mut(body.identifier.get())?,
        storage_index,
        payload,
        &[section],
        &[],
        limits,
    )?;
    for identifier in &removed {
        working.remove(*identifier)?;
    }
    let (candidate, touched) = working.publish(&transition)?;
    structure::verify_removed(&candidate, &closure)?;

    // The removed section's text joins the preceding one's.
    let before = source.sections();
    let now = candidate.sections();
    if now.len() + 1 != before.len() {
        return Err(Error::Verification);
    }
    for (index, section) in now.iter().enumerate() {
        let expected = match index.cmp(&(position - 1)) {
            std::cmp::Ordering::Less => before[index].body_text().map(str::to_owned),
            std::cmp::Ordering::Equal => Some(format!(
                "{}{}",
                before[index].body_text().unwrap_or_default(),
                before[position].body_text().unwrap_or_default()
            )),
            std::cmp::Ordering::Greater => before[index + 1].body_text().map(str::to_owned),
        };
        if section.body_text().unwrap_or_default() != expected.as_deref().unwrap_or_default() {
            return Err(Error::Verification);
        }
    }
    Ok((candidate, touched, removed.len()))
}

impl Package {
    /// Start an immutable single-operation section add or remove edit.
    ///
    /// ```no_run
    /// use litchi_pages::Package;
    ///
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let package = Package::open("report.pages")?;
    /// let commit = package
    ///     .edit_section_lifecycle()
    ///     .add(0, Some("Appendix"))?
    ///     .commit()?;
    /// assert_eq!(commit.package().sections().len(), package.sections().len() + 1);
    /// let restored = commit
    ///     .package()
    ///     .apply_section_lifecycle(&commit.patch().inverse())?;
    /// assert_eq!(restored.package().sections().len(), package.sections().len());
    /// # Ok(())
    /// # }
    /// ```
    #[must_use]
    pub const fn edit_section_lifecycle(&self) -> Edit<'_> {
        Edit {
            source: self,
            operation: None,
        }
    }

    /// Apply a section-lifecycle patch to its exact source package.
    ///
    /// # Errors
    ///
    /// [`Error::PatchConflict`] when this package is not the patch's exact
    /// source; [`Error::Verification`] when its retained target no longer
    /// opens.
    pub fn apply_section_lifecycle(&self, patch: &Patch) -> Result<Commit, Error> {
        let package = patch.artifacts.apply(self)?.ok_or(Error::PatchConflict)?;
        Ok(Commit {
            diagnostics: Diagnostics {
                changed: !patch.is_noop(),
                touched_components: 0,
                objects: 0,
            },
            patch: patch.clone(),
            package,
        })
    }
}
