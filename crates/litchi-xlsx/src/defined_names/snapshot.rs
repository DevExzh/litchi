//! Immutable source-bound workbook defined-name state.

use std::sync::Arc;

use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{OpcPackage, PackURI, PartView, Relationship, SourceBackedPackage, TargetMode};

use crate::error::{Error, Result, invalid};
use crate::raw::{self, DefinedName};

/// Exact workbook owner state plus its inert defined-name catalog.
#[derive(Clone, Debug)]
pub struct Snapshot {
    names: Box<[DefinedName]>,
    source: SourceState,
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
            workbook.blob_arc(),
            owner,
        )
    }

    pub(super) fn load_source_backed(package: &SourceBackedPackage) -> Result<Self> {
        let workbook = package.main_document_part()?;
        require_workbook_content_type(workbook.content_type())?;
        let bytes = workbook.data()?.into_arc()?;
        let owner = current_owner_relationship(package.rels())
            .ok_or_else(|| invalid("workbook has no unique officeDocument owner"))?;
        Self::from_source_backed_parts(&workbook, bytes, owner)
    }

    fn from_source_backed_parts(
        workbook: &PartView<'_>,
        bytes: Arc<Vec<u8>>,
        owner: &Relationship,
    ) -> Result<Self> {
        Self::from_parts(
            workbook.partname().clone(),
            workbook.content_type(),
            bytes,
            owner,
        )
    }

    fn from_parts(
        part_name: PackURI,
        content_type: &str,
        bytes: Arc<Vec<u8>>,
        owner: &Relationship,
    ) -> Result<Self> {
        let names = raw::parse_catalog(bytes.as_slice())?
            .defined_names
            .into_boxed_slice();
        Ok(Self {
            names,
            source: SourceState {
                part_name,
                content_type: copy_boxed(content_type, "defined-name workbook content type")?,
                bytes,
                owner_relationship: SourceRelationship::capture(owner)?,
            },
        })
    }

    pub(super) fn from_rewritten_source(source: &Self, bytes: Vec<u8>) -> Result<Self> {
        let names = raw::parse_catalog(&bytes)?.defined_names.into_boxed_slice();
        let mut rewritten = source.clone();
        rewritten.names = names;
        rewritten.source.bytes = Arc::new(bytes);
        Ok(rewritten)
    }

    /// Exact authored defined-name records in source order.
    #[must_use]
    pub fn defined_names(&self) -> &[DefinedName] {
        &self.names
    }

    /// Resolved workbook Part name.
    #[must_use]
    pub const fn workbook_part_name(&self) -> &PackURI {
        &self.source.part_name
    }

    /// Exact source workbook XML.
    #[must_use]
    pub fn source_xml(&self) -> &[u8] {
        self.source.bytes.as_slice()
    }

    /// Shared exact source workbook XML.
    #[must_use]
    pub fn source_arc(&self) -> Arc<Vec<u8>> {
        Arc::clone(&self.source.bytes)
    }

    /// Exact workbook content type captured at ingress.
    #[must_use]
    pub fn content_type(&self) -> &str {
        &self.source.content_type
    }

    pub(super) fn same_source(&self, other: &Self) -> bool {
        self.source == other.source
    }

    pub(super) fn matches_current_source(&self, package: &OpcPackage) -> bool {
        let Ok(workbook) = package.main_document_part() else {
            return false;
        };
        workbook.partname() == &self.source.part_name
            && workbook.content_type() == self.source.content_type.as_ref()
            && workbook.blob() == self.source.bytes.as_slice()
            && current_owner_relationship(package.rels())
                .is_some_and(|owner| self.source.owner_relationship.matches(owner))
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
    bytes: Arc<Vec<u8>>,
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

    fn matches(&self, relationship: &Relationship) -> bool {
        relationship.r_id() == self.id.as_ref()
            && relationship.reltype() == self.relationship_type.as_ref()
            && relationship.target_ref() == self.target.as_ref()
            && relationship.target_mode() == self.mode
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
