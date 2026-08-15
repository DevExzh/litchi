//! Immutable source-bound worksheet page-margin state.

use std::sync::Arc;

use litchi_core::Selector as CoreSelector;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{OpcPackage, PackURI, Part, Relationship, SourceBackedPackage, TargetMode};

use super::Margins;
use crate::error::{Error, Result, invalid};
use crate::source_provenance::{SourceBinding, SourceProvenance};
use crate::workbook::source::validate_sheet_graph;
use crate::{Selector, Workbook, WorksheetKind, raw};

/// Semantic page margins plus the exact worksheet owner bytes they came from.
#[derive(Clone, Debug)]
pub struct Snapshot {
    value: Option<Margins>,
    sheet_name: Box<str>,
    sheet_position: usize,
    source: SourceState,
    binding: SourceBinding,
}

impl Snapshot {
    /// Resolve and read one worksheet by its ordinary semantic selector.
    ///
    /// # Errors
    ///
    /// Returns an error when the package or selector is invalid, the selected
    /// sheet is not a worksheet, or its page-margin XML is invalid.
    pub fn load<'a>(package: &OpcPackage, selector: impl Into<Selector<'a>>) -> Result<Self> {
        let workbook = Workbook::from_package(package.clone())?;
        let sheet = workbook
            .sheet(selector)?
            .ok_or_else(|| invalid("page-margin worksheet selector did not resolve"))?;
        if sheet.kind() != WorksheetKind::Worksheet {
            return Err(Error::NotWorksheet {
                sheet: sheet.name().to_owned(),
            });
        }

        let workbook_part = package.main_document_part()?;
        let catalog = raw::parse_catalog(workbook_part.blob())?;
        let catalog_sheet = catalog.sheets.get(sheet.position()).ok_or_else(|| {
            invalid("page-margin worksheet position is absent from the workbook catalog")
        })?;
        if catalog_sheet.name != sheet.name() {
            return Err(invalid(
                "page-margin worksheet name differs from the workbook catalog",
            ));
        }
        let worksheet = package.get_part(sheet.part_uri())?;
        let relationship = require_selected_worksheet(
            workbook_part.rels(),
            catalog_sheet,
            worksheet.partname(),
            worksheet.content_type(),
        )?;
        let owner = current_owner_relationship(package.rels())
            .ok_or_else(|| invalid("workbook has no unique officeDocument owner"))?;

        Self::from_parts(
            sheet.name(),
            sheet.position(),
            workbook_part.partname().clone(),
            workbook_part.content_type(),
            workbook_part.blob_arc(),
            owner,
            worksheet.partname().clone(),
            worksheet.content_type(),
            worksheet.blob_arc(),
            relationship,
            SourceBinding::default(),
        )
    }

    pub(super) fn load_source_backed<'a>(
        package: &SourceBackedPackage,
        selector: impl Into<Selector<'a>>,
    ) -> Result<Self> {
        let workbook = package.main_document_part()?;
        require_workbook_content_type(workbook.content_type())?;
        let workbook_xml = workbook.data()?.into_arc()?;
        let catalog = raw::parse_catalog(workbook_xml.as_slice())?;
        let sheet_parts = validate_sheet_graph(package, &workbook, &catalog.sheets)?;
        let sheet_position = resolve_selector(&catalog.sheets, selector.into())?
            .ok_or_else(|| invalid("page-margin worksheet selector did not resolve"))?;
        let catalog_sheet = &catalog.sheets[sheet_position];
        let sheet_part = &sheet_parts[sheet_position];
        if sheet_part.kind != WorksheetKind::Worksheet {
            return Err(Error::NotWorksheet {
                sheet: catalog_sheet.name.clone(),
            });
        }
        let worksheet = package.part(&sheet_part.uri)?;
        let relationship = require_selected_worksheet(
            workbook.rels(),
            catalog_sheet,
            worksheet.partname(),
            worksheet.content_type(),
        )?;
        let worksheet_xml = worksheet.data()?.into_arc()?;
        let owner = current_owner_relationship(package.rels())
            .ok_or_else(|| invalid("workbook has no unique officeDocument owner"))?;

        let snapshot = Self::from_parts(
            &catalog_sheet.name,
            sheet_position,
            workbook.partname().clone(),
            workbook.content_type(),
            workbook_xml,
            owner,
            worksheet.partname().clone(),
            worksheet.content_type(),
            worksheet_xml,
            relationship,
            SourceBinding::default(),
        )?;
        Ok(Self {
            binding: SourceBinding::capture(package)?,
            ..snapshot
        })
    }

    #[allow(clippy::too_many_arguments)]
    fn from_parts(
        sheet_name: &str,
        sheet_position: usize,
        workbook_uri: PackURI,
        workbook_content_type: &str,
        workbook_xml: Arc<Vec<u8>>,
        owner_relationship: &Relationship,
        worksheet_uri: PackURI,
        worksheet_content_type: &str,
        worksheet_xml: Arc<Vec<u8>>,
        sheet_relationship: &Relationship,
        binding: SourceBinding,
    ) -> Result<Self> {
        let value = super::parse_page_margins(worksheet_xml.as_slice())?;
        Ok(Self {
            value,
            sheet_name: copy_boxed(sheet_name, "page-margin sheet name")?,
            sheet_position,
            binding,
            source: SourceState {
                workbook: PartState::new(
                    workbook_uri,
                    workbook_content_type,
                    workbook_xml,
                    "page-margin workbook content type",
                )?,
                worksheet: PartState::new(
                    worksheet_uri,
                    worksheet_content_type,
                    worksheet_xml,
                    "page-margin worksheet content type",
                )?,
                owner_relationship: SourceRelationship::capture(owner_relationship)?,
                sheet_relationship: SourceRelationship::capture(sheet_relationship)?,
            },
        })
    }

    pub(super) fn from_rewritten_source(source: &Self, bytes: Vec<u8>) -> Result<Self> {
        let value = super::parse_page_margins(&bytes)?;
        let mut rewritten = source.clone();
        rewritten.value = value;
        rewritten.source.worksheet.bytes = Arc::new(bytes);
        Ok(rewritten)
    }

    /// Typed direct page-margin state.
    #[must_use]
    pub const fn page_margins(&self) -> Option<&Margins> {
        self.value.as_ref()
    }

    /// Developer-facing worksheet name captured at ingress.
    #[must_use]
    pub fn sheet_name(&self) -> &str {
        &self.sheet_name
    }

    /// Checked zero-based worksheet position captured at ingress.
    #[must_use]
    pub const fn sheet_position(&self) -> usize {
        self.sheet_position
    }

    /// Resolved worksheet part name.
    #[must_use]
    pub const fn worksheet_part_name(&self) -> &PackURI {
        &self.source.worksheet.uri
    }

    /// Exact source worksheet XML.
    #[must_use]
    pub fn source_xml(&self) -> &[u8] {
        self.source.worksheet.bytes.as_slice()
    }

    /// Shared exact source worksheet XML.
    #[must_use]
    pub fn source_arc(&self) -> Arc<Vec<u8>> {
        Arc::clone(&self.source.worksheet.bytes)
    }

    pub(super) fn same_source(&self, other: &Self) -> bool {
        self.sheet_name == other.sheet_name
            && self.sheet_position == other.sheet_position
            && self.source == other.source
            && self.binding.same_or_unavailable(&other.binding)
    }

    /// Check the retained source lineage and revision without reloading XML.
    pub(super) fn matches_source_backed(
        &self,
        package: &SourceBackedPackage,
    ) -> Result<SourceProvenance> {
        self.binding.check(package)
    }

    pub(super) fn matches_current_source(&self, package: &OpcPackage) -> bool {
        let Ok(workbook) = package.main_document_part() else {
            return false;
        };
        if !self.source.workbook.matches_part(workbook)
            || !current_owner_relationship(package.rels())
                .is_some_and(|relationship| self.source.owner_relationship.matches(relationship))
        {
            return false;
        }
        let Some(relationship) = workbook
            .rels()
            .get(self.source.sheet_relationship.id.as_ref())
        else {
            return false;
        };
        let Ok(target) = relationship.target_partname() else {
            return false;
        };
        if !self.source.sheet_relationship.matches(relationship)
            || target != self.source.worksheet.uri
        {
            return false;
        }
        package
            .get_part(&self.source.worksheet.uri)
            .is_ok_and(|part| self.source.worksheet.matches_part(part))
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct SourceState {
    workbook: PartState,
    worksheet: PartState,
    owner_relationship: SourceRelationship,
    sheet_relationship: SourceRelationship,
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct PartState {
    uri: PackURI,
    content_type: Box<str>,
    bytes: Arc<Vec<u8>>,
}

impl PartState {
    fn new(
        uri: PackURI,
        content_type: &str,
        bytes: Arc<Vec<u8>>,
        resource: &'static str,
    ) -> Result<Self> {
        Ok(Self {
            uri,
            content_type: copy_boxed(content_type, resource)?,
            bytes,
        })
    }

    fn matches_part(&self, part: &dyn Part) -> bool {
        part.partname() == &self.uri
            && part.content_type() == self.content_type.as_ref()
            && part.blob() == self.bytes.as_slice()
    }
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
            id: copy_boxed(relationship.r_id(), "page-margin relationship ID")?,
            relationship_type: copy_boxed(relationship.reltype(), "page-margin relationship type")?,
            target: copy_boxed(relationship.target_ref(), "page-margin relationship target")?,
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

fn resolve_selector(sheets: &[raw::Sheet], selector: Selector<'_>) -> Result<Option<usize>> {
    match selector {
        CoreSelector::Position(position) => {
            Ok((position.get() < sheets.len()).then_some(position.get()))
        },
        CoreSelector::Name(name) => {
            let key = crate::sheet::key(&name);
            Ok(sheets
                .iter()
                .position(|sheet| crate::sheet::key(&sheet.name) == key))
        },
        CoreSelector::Id(never) => match never {},
        _ => Err(Error::UnsupportedSelector),
    }
}

fn require_selected_worksheet<'a>(
    relationships: &'a litchi_opc::Relationships,
    sheet: &raw::Sheet,
    worksheet_uri: &PackURI,
    content_type: &str,
) -> Result<&'a Relationship> {
    let relationship = relationships
        .get(&sheet.relationship_id)
        .ok_or_else(|| invalid("selected worksheet relationship is missing"))?;
    if relationship.target_mode() != TargetMode::Internal
        || !matches!(relationship.reltype(), rt::WORKSHEET | rt::STRICT_WORKSHEET)
        || relationship.target_partname()? != *worksheet_uri
        || content_type != ct::SML_WORKSHEET
    {
        return Err(invalid(
            "selected worksheet relationship or content type is invalid",
        ));
    }
    Ok(relationship)
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

fn copy_boxed(value: &str, resource: &'static str) -> Result<Box<str>> {
    let mut copied = String::new();
    copied
        .try_reserve_exact(value.len())
        .map_err(|source| Error::Allocation { resource, source })?;
    copied.push_str(value);
    Ok(copied.into_boxed_str())
}
