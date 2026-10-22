//! Immutable source-bound workbook defined-name state.

use std::sync::Arc;

use litchi_core::{ExecutionContext, ExecutionError};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{
    OpcPackage, PackURI, PartView, Relationship, SourceBackedPackage, SourceLimits, TargetMode,
};

use crate::error::{Error, Result, invalid};
use crate::raw::{self, DefinedName};
use crate::source_payload::SourcePayload;

/// Exact workbook owner state plus its inert defined-name catalog.
#[derive(Clone, Debug)]
pub struct Snapshot {
    /// Immutable catalog records shared with the deferred output and restore
    /// endpoint. A defined-name edit replaces this Arc only when names change.
    names: Arc<[DefinedName]>,
    active_sheet: Option<usize>,
    source: SourceState,
    context: Option<ExecutionContext>,
    pub(super) limits: SourceLimits,
}

impl Snapshot {
    /// Load defined names from an ordinary materialized OPC package.
    pub fn load(package: &OpcPackage) -> Result<Self> {
        let workbook = package.main_document_part()?;
        require_workbook_content_type(workbook.content_type())?;
        let owner = current_owner_relationship(package.rels())
            .ok_or_else(|| invalid("workbook has no unique officeDocument owner"))?;
        Self::from_parts(
            workbook.partname().clone(),
            workbook.content_type(),
            SourcePayload::Owned(workbook.blob_arc()),
            owner,
            None,
            package.read_limits(),
            None,
        )
    }

    pub(super) fn load_source_backed(package: &SourceBackedPackage) -> Result<Self> {
        Self::load_source_backed_with_names(package, None)
    }

    pub(super) fn load_source_backed_with_names(
        package: &SourceBackedPackage,
        shared_names: Option<Arc<[DefinedName]>>,
    ) -> Result<Self> {
        package.check_execution()?;
        let source_version = package.source_version()?;
        let workbook = package.main_document_part()?;
        require_workbook_content_type(workbook.content_type())?;
        let bytes = SourcePayload::from_part_data(package, workbook.data()?)?;
        let owner = current_owner_relationship(package.rels())
            .ok_or_else(|| invalid("workbook has no unique officeDocument owner"))?;
        let snapshot = Self::from_source_backed_parts(
            &workbook,
            bytes,
            owner,
            package.execution_context(),
            package.read_limits(),
            shared_names,
        )?;
        package.check_execution()?;
        if package.source_version()? != source_version {
            return Err(invalid(
                "defined-name source version changed during snapshot",
            ));
        }
        Ok(snapshot)
    }

    fn from_source_backed_parts(
        workbook: &PartView<'_>,
        bytes: SourcePayload,
        owner: &Relationship,
        context: Option<ExecutionContext>,
        limits: SourceLimits,
        shared_names: Option<Arc<[DefinedName]>>,
    ) -> Result<Self> {
        Self::from_parts(
            workbook.partname().clone(),
            workbook.content_type(),
            bytes,
            owner,
            context,
            limits,
            shared_names,
        )
    }

    fn from_parts(
        part_name: PackURI,
        content_type: &str,
        bytes: SourcePayload,
        owner: &Relationship,
        context: Option<ExecutionContext>,
        limits: SourceLimits,
        shared_names: Option<Arc<[DefinedName]>>,
    ) -> Result<Self> {
        let catalog = raw::parse_catalog(bytes.as_bytes(), &limits)?;
        let active_sheet = (!catalog.sheets.is_empty()).then_some(catalog.active_sheet_index);
        let names = match shared_names {
            Some(shared) => {
                if catalog.defined_names.as_slice() != shared.as_ref() {
                    return Err(invalid(
                        "defined-name source catalog changed during snapshot",
                    ));
                }
                shared
            },
            None => Arc::from(catalog.defined_names.into_boxed_slice()),
        };
        Ok(Self {
            names,
            active_sheet,
            source: SourceState {
                part_name,
                content_type: copy_boxed(content_type, "defined-name workbook content type")?,
                bytes,
                owner_relationship: SourceRelationship::capture(owner)?,
            },
            context,
            limits,
        })
    }

    pub(super) fn from_rewritten_source(source: &Self, bytes: Vec<u8>) -> Result<Self> {
        let catalog = raw::parse_catalog(&bytes, &source.limits)?;
        let active_sheet = (!catalog.sheets.is_empty()).then_some(catalog.active_sheet_index);
        let names = Arc::from(catalog.defined_names.into_boxed_slice());
        let mut rewritten = source.clone();
        rewritten.names = names;
        rewritten.active_sheet = active_sheet;
        rewritten.source.bytes = SourcePayload::Owned(Arc::new(bytes));
        Ok(rewritten)
    }

    /// Exact authored defined-name records in source order.
    #[must_use]
    pub fn defined_names(&self) -> &[DefinedName] {
        &self.names
    }

    pub(super) fn deferred_catalog(&self) -> crate::workbook::source::deferred::Catalog {
        crate::workbook::source::deferred::Catalog {
            defined_names: self.names.clone(),
            active_sheet: self.active_sheet,
        }
    }

    /// Resolved workbook Part name.
    #[must_use]
    pub const fn workbook_part_name(&self) -> &PackURI {
        &self.source.part_name
    }

    /// Exact source workbook XML.
    #[must_use]
    pub fn source_xml(&self) -> &[u8] {
        self.source.bytes.as_bytes()
    }

    /// Shared exact source workbook XML.
    ///
    /// Managed source payloads return
    /// [`litchi_opc::OpcError::ManagedPartDataArcEscape`] rather than silently
    /// detaching their execution-budget reservation.
    pub fn source_arc(&self) -> Result<Arc<Vec<u8>>> {
        self.source.bytes.detached_arc()
    }

    pub(super) fn check_execution(&self) -> Result<()> {
        let Some(context) = self.context.as_ref() else {
            return Ok(());
        };
        context.check().map_err(|error| {
            Error::Package(match error {
                ExecutionError::Cancelled => litchi_opc::OpcError::Cancelled,
                error => litchi_opc::OpcError::Execution(error),
            })
        })
    }

    /// Exact workbook content type captured at ingress.
    #[must_use]
    pub fn content_type(&self) -> &str {
        &self.source.content_type
    }

    pub(super) fn same_source(&self, other: &Self) -> bool {
        self.source == other.source
    }
}

impl PartialEq for Snapshot {
    fn eq(&self, other: &Self) -> bool {
        self.source == other.source
    }
}

impl Eq for Snapshot {}

#[derive(Clone, Debug, PartialEq, Eq)]
struct SourceState {
    part_name: PackURI,
    content_type: Box<str>,
    bytes: SourcePayload,
    owner_relationship: SourceRelationship,
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct SourceRelationship {
    id: Box<str>,
    relationship_type: Box<str>,
    target: Box<str>,
    mode: TargetMode,
}

impl SourceRelationship {
    fn capture(relationship: &Relationship) -> Result<Self> {
        Ok(Self {
            id: copy_boxed(relationship.r_id(), "defined-name owner relationship ID")?,
            relationship_type: copy_boxed(
                relationship.reltype(),
                "defined-name owner relationship type",
            )?,
            target: copy_boxed(
                relationship.target_ref(),
                "defined-name owner relationship target",
            )?,
            mode: relationship.target_mode(),
        })
    }
}

fn current_owner_relationship(relationships: &litchi_opc::Relationships) -> Option<&Relationship> {
    let mut owners = relationships.iter().filter(|relationship| {
        matches!(
            relationship.reltype(),
            rt::OFFICE_DOCUMENT | rt::STRICT_OFFICE_DOCUMENT
        )
    });
    let owner = owners.next()?;
    if owners.next().is_some() || owner.target_mode() != TargetMode::Internal {
        return None;
    }
    Some(owner)
}

fn copy_boxed(value: &str, resource: &'static str) -> Result<Box<str>> {
    let mut copied = String::new();
    copied
        .try_reserve_exact(value.len())
        .map_err(|source| Error::Allocation { resource, source })?;
    copied.push_str(value);
    Ok(copied.into_boxed_str())
}

fn require_workbook_content_type(content_type: &str) -> Result<()> {
    if matches!(
        content_type,
        ct::SML_SHEET_MAIN
            | ct::SML_TEMPLATE_MAIN
            | ct::SML_SHEET_MACRO_MAIN
            | ct::SML_TEMPLATE_MACRO_MAIN
    ) {
        Ok(())
    } else {
        Err(invalid(format!(
            "main part has non-XLSX content type '{content_type}'"
        )))
    }
}
