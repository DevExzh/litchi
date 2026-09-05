//! Immutable source-bound worksheet hyperlink state.

use std::sync::Arc;

use litchi_core::{ExecutionContext, ExecutionError, Selector as CoreSelector, SourceVersion};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{
    OpcPackage, PackURI, Relationship, Relationships, SourceBackedPackage, SourceLineage,
    TargetMode,
};
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

use super::codec::ParsedHyperlink;
use super::model::Hyperlink;
use crate::error::{Error, Result, allocation, invalid};
use crate::source_payload::SourcePayload;
use crate::workbook::source::validate_sheet_graph;
use crate::{Selector, WorksheetKind, raw};

const MAIN_NAMESPACE: &[u8] = b"http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT_MAIN_NAMESPACE: &[u8] = b"http://purl.oclc.org/ooxml/spreadsheetml/main";
const MCE_NAMESPACE: &[u8] = b"http://schemas.openxmlformats.org/markup-compatibility/2006";

/// Semantic worksheet hyperlinks plus the exact source owner closure.
#[derive(Clone, Debug)]
pub struct Snapshot {
    values: Box<[Hyperlink]>,
    relationship_ids: Box<[Option<Box<str>>]>,
    sheet_name: Box<str>,
    sheet_position: usize,
    protected: bool,
    source: SourceState,
    source_version: Option<SourceVersion>,
    source_lineage: Option<SourceLineage>,
    context: Option<ExecutionContext>,
}

/// One private hyperlink value and its source relationship identity.
#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct SourceEntry {
    pub(crate) value: Hyperlink,
    pub(crate) relationship_id: Option<Box<str>>,
}

pub(crate) struct SourceEntryRef<'a> {
    pub(crate) value: &'a Hyperlink,
    pub(crate) relationship_id: Option<&'a str>,
}

impl Snapshot {
    /// Load one existing worksheet from an eager OPC package.
    pub fn load<'a>(package: &OpcPackage, selector: impl Into<Selector<'a>>) -> Result<Self> {
        let workbook = package.main_document_part()?;
        require_workbook_content_type(workbook.content_type())?;
        let workbook_xml = workbook.blob();
        let catalog = raw::parse_catalog(workbook_xml)?;
        let sheet_position = resolve_selector(&catalog.sheets, selector.into())?
            .ok_or_else(|| invalid("hyperlink worksheet selector did not resolve"))?;
        let catalog_sheet = catalog.sheets.get(sheet_position).ok_or_else(|| {
            invalid("hyperlink worksheet position is absent from the workbook catalog")
        })?;
        let sheet_relationship = workbook
            .rels()
            .get(&catalog_sheet.relationship_id)
            .ok_or_else(|| invalid("selected worksheet relationship is missing"))?;
        if sheet_relationship.target_mode() != TargetMode::Internal {
            return Err(invalid(
                "selected worksheet relationship cannot be external",
            ));
        }
        let worksheet_uri = sheet_relationship.target_partname()?;
        let worksheet = package.get_part(&worksheet_uri)?;
        require_selected_worksheet(
            sheet_relationship,
            worksheet.partname(),
            worksheet.content_type(),
        )?;
        let worksheet_xml = worksheet.blob();
        let protected = validate_surface(workbook_xml, worksheet_xml)?;
        let parsed = super::codec::parse_with_relationship_ids(worksheet_xml, worksheet.rels())?;
        let owner_relationship = current_owner_relationship(package.rels())
            .ok_or_else(|| invalid("workbook has no unique officeDocument owner"))?;
        Self::from_parts(
            &catalog_sheet.name,
            sheet_position,
            workbook.partname().clone(),
            workbook.content_type(),
            SourcePayload::Owned(workbook.blob_arc()),
            owner_relationship,
            worksheet.partname().clone(),
            worksheet.content_type(),
            SourcePayload::Owned(worksheet.blob_arc()),
            sheet_relationship,
            worksheet.rels(),
            parsed,
            protected,
            None,
            None,
            None,
        )
    }

    /// Load one existing worksheet without materializing unselected parts.
    pub(crate) fn load_source_backed<'a>(
        package: &SourceBackedPackage,
        selector: impl Into<Selector<'a>>,
    ) -> Result<Self> {
        package.check_execution()?;
        if package.has_encrypted_entries() {
            return Err(Error::Unsupported {
                feature: "encrypted XLSX source-backed hyperlink editing",
            });
        }
        let source_version = package.source_version()?;
        let source_lineage = package.source_lineage();
        let workbook = package.main_document_part()?;
        require_workbook_content_type(workbook.content_type())?;
        let workbook_payload = SourcePayload::from_part_data(package, workbook.data()?)?;
        let workbook_xml = workbook_payload.as_bytes();
        let catalog = raw::parse_catalog(workbook_xml)?;
        let sheet_parts = validate_sheet_graph(package, &workbook, &catalog.sheets)?;
        let sheet_position = resolve_selector(&catalog.sheets, selector.into())?
            .ok_or_else(|| invalid("hyperlink worksheet selector did not resolve"))?;
        let catalog_sheet = catalog.sheets.get(sheet_position).ok_or_else(|| {
            invalid("hyperlink worksheet position is absent from the workbook catalog")
        })?;
        let sheet_part = sheet_parts.get(sheet_position).ok_or_else(|| {
            invalid("hyperlink worksheet binding is absent from the workbook graph")
        })?;
        if sheet_part.kind != WorksheetKind::Worksheet {
            return Err(Error::NotWorksheet {
                sheet: catalog_sheet.name.clone(),
            });
        }
        let worksheet = package.part(&sheet_part.uri)?;
        let sheet_relationship = workbook
            .rels()
            .get(&catalog_sheet.relationship_id)
            .ok_or_else(|| invalid("selected worksheet relationship is missing"))?;
        require_selected_worksheet(
            sheet_relationship,
            worksheet.partname(),
            worksheet.content_type(),
        )?;
        let worksheet_payload = SourcePayload::from_part_data(package, worksheet.data()?)?;
        let protected = validate_surface(workbook_xml, worksheet_payload.as_bytes())?;
        let parsed = super::codec::parse_with_relationship_ids(
            worksheet_payload.as_bytes(),
            worksheet.rels(),
        )?;
        let owner_relationship = current_owner_relationship(package.rels())
            .ok_or_else(|| invalid("workbook has no unique officeDocument owner"))?;
        let snapshot = Self::from_parts(
            &catalog_sheet.name,
            sheet_position,
            workbook.partname().clone(),
            workbook.content_type(),
            workbook_payload,
            owner_relationship,
            worksheet.partname().clone(),
            worksheet.content_type(),
            worksheet_payload,
            sheet_relationship,
            worksheet.rels(),
            parsed,
            protected,
            Some(source_version),
            Some(source_lineage),
            package.execution_context(),
        )?;
        package.check_execution()?;
        if package.source_version()? != source_version {
            return Err(invalid("hyperlink source version changed during snapshot"));
        }
        Ok(snapshot)
    }

    #[allow(clippy::too_many_arguments)]
    fn from_parts(
        sheet_name: &str,
        sheet_position: usize,
        workbook_uri: PackURI,
        workbook_content_type: &str,
        workbook_xml: SourcePayload,
        owner_relationship: &Relationship,
        worksheet_uri: PackURI,
        worksheet_content_type: &str,
        worksheet_xml: SourcePayload,
        sheet_relationship: &Relationship,
        worksheet_relationships: &Relationships,
        parsed: Vec<ParsedHyperlink>,
        protected: bool,
        source_version: Option<SourceVersion>,
        source_lineage: Option<SourceLineage>,
        context: Option<ExecutionContext>,
    ) -> Result<Self> {
        let (values, relationship_ids) = split_parsed(parsed)?;
        Ok(Self {
            values,
            relationship_ids,
            sheet_name: copy_boxed(sheet_name, "hyperlink sheet name")?,
            sheet_position,
            protected,
            source: SourceState {
                workbook: PartState::new(
                    workbook_uri,
                    workbook_content_type,
                    workbook_xml,
                    "hyperlink workbook content type",
                )?,
                worksheet: PartState::new(
                    worksheet_uri,
                    worksheet_content_type,
                    worksheet_xml,
                    "hyperlink worksheet content type",
                )?,
                owner_relationship: SourceRelationship::capture(owner_relationship)?,
                sheet_relationship: SourceRelationship::capture(sheet_relationship)?,
                worksheet_relationships: capture_relationships(worksheet_relationships)?,
            },
            source_version,
            source_lineage,
            context,
        })
    }

    pub(crate) fn from_rewritten_source(
        source: &Self,
        bytes: Vec<u8>,
        expected: &[SourceEntry],
    ) -> Result<Self> {
        let relationships = source.reconstructed_relationships()?;
        let parsed = super::codec::parse_with_relationship_ids(&bytes, &relationships)?;
        let (values, relationship_ids) = split_parsed(parsed)?;
        if values.len() != expected.len()
            || values
                .iter()
                .zip(expected)
                .any(|(value, expected)| value != &expected.value)
            || relationship_ids
                .iter()
                .zip(expected)
                .any(|(relationship_id, expected)| relationship_id != &expected.relationship_id)
        {
            return Err(invalid(
                "hyperlink publication changed the staged semantic state",
            ));
        }
        let mut rewritten = source.clone();
        rewritten.values = values;
        rewritten.relationship_ids = relationship_ids;
        rewritten.source.worksheet.bytes = SourcePayload::Owned(Arc::new(bytes));
        Ok(rewritten)
    }

    /// Typed direct hyperlinks in source order.
    #[must_use]
    pub fn hyperlinks(&self) -> &[Hyperlink] {
        &self.values
    }

    /// Alias for [`Self::hyperlinks`].
    #[must_use]
    pub fn values(&self) -> &[Hyperlink] {
        self.hyperlinks()
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

    pub(crate) const fn worksheet_part_name(&self) -> &PackURI {
        &self.source.worksheet.uri
    }

    pub(crate) fn source_xml(&self) -> &[u8] {
        self.source.worksheet.bytes.as_bytes()
    }

    pub(crate) fn source_arc(&self) -> Result<Arc<Vec<u8>>> {
        self.source.worksheet.bytes.detached_arc()
    }

    pub(crate) fn check_execution(&self) -> Result<()> {
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

    pub(crate) fn protected(&self) -> bool {
        self.protected
    }

    pub(crate) fn source_entries(&self) -> Result<Vec<SourceEntry>> {
        let mut entries = Vec::new();
        entries
            .try_reserve_exact(self.values.len())
            .map_err(|source| allocation("source-backed worksheet hyperlinks", source))?;
        for (index, (value, relationship_id)) in self
            .values
            .iter()
            .cloned()
            .zip(self.relationship_ids.iter().cloned())
            .enumerate()
        {
            if index % 256 == 0 {
                self.check_execution()?;
            }
            entries.push(SourceEntry {
                value,
                relationship_id,
            });
        }
        self.check_execution()?;
        Ok(entries)
    }

    pub(crate) fn source_entry_refs(&self) -> impl ExactSizeIterator<Item = SourceEntryRef<'_>> {
        self.values
            .iter()
            .zip(self.relationship_ids.iter())
            .map(|(value, relationship_id)| SourceEntryRef {
                value,
                relationship_id: relationship_id.as_deref(),
            })
    }

    pub(crate) fn matches_entries(&self, entries: &[SourceEntry]) -> bool {
        entries.len() == self.values.len()
            && entries
                .iter()
                .zip(self.values.iter().zip(self.relationship_ids.iter()))
                .all(|(entry, (value, relationship_id))| {
                    &entry.value == value
                        && entry.relationship_id.as_ref() == relationship_id.as_ref()
                })
    }

    pub(crate) fn matches_entries_checked(&self, entries: &[SourceEntry]) -> Result<bool> {
        if entries.len() != self.values.len() {
            return Ok(false);
        }
        for (index, (entry, (value, relationship_id))) in entries
            .iter()
            .zip(self.values.iter().zip(self.relationship_ids.iter()))
            .enumerate()
        {
            if index % 256 == 0 {
                self.check_execution()?;
            }
            if &entry.value != value || entry.relationship_id.as_ref() != relationship_id.as_ref() {
                return Ok(false);
            }
        }
        self.check_execution()?;
        Ok(true)
    }

    pub(crate) fn same_source(&self, other: &Self) -> bool {
        self.sheet_name == other.sheet_name
            && self.sheet_position == other.sheet_position
            && self.source == other.source
    }

    pub(crate) fn matches_current_source(&self, package: &OpcPackage) -> bool {
        let Ok(current) = Self::load(package, self.sheet_position) else {
            return false;
        };
        self.same_source(&current)
    }

    pub(crate) fn matches_source_backed(&self, package: &SourceBackedPackage) -> Result<bool> {
        if let Some(lineage) = self.source_lineage.as_ref()
            && lineage != &package.source_lineage()
        {
            return Ok(false);
        }
        if let Some(version) = self.source_version
            && package.source_version()? != version
        {
            return Ok(false);
        }
        let current = Self::load_source_backed(package, self.sheet_position)?;
        Ok(self.same_source(&current))
    }

    fn reconstructed_relationships(&self) -> Result<Relationships> {
        let mut relationships = Relationships::new(self.source.worksheet.uri.base_uri().to_owned());
        for relationship in &self.source.worksheet_relationships {
            relationships
                .try_add_relationship(
                    relationship.relationship_type.to_string(),
                    relationship.target.to_string(),
                    relationship.id.to_string(),
                    relationship.mode,
                )
                .map_err(Error::Package)?;
        }
        Ok(relationships)
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct SourceState {
    workbook: PartState,
    worksheet: PartState,
    owner_relationship: SourceRelationship,
    sheet_relationship: SourceRelationship,
    worksheet_relationships: Vec<SourceRelationship>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct PartState {
    uri: PackURI,
    content_type: Box<str>,
    bytes: SourcePayload,
}

impl PartState {
    fn new(
        uri: PackURI,
        content_type: &str,
        bytes: SourcePayload,
        resource: &'static str,
    ) -> Result<Self> {
        Ok(Self {
            uri,
            content_type: copy_boxed(content_type, resource)?,
            bytes,
        })
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
            id: copy_boxed(relationship.r_id(), "hyperlink relationship ID")?,
            relationship_type: copy_boxed(relationship.reltype(), "hyperlink relationship type")?,
            target: copy_boxed(relationship.target_ref(), "hyperlink relationship target")?,
            mode: relationship.target_mode(),
        })
    }
}

fn split_parsed(
    parsed: Vec<ParsedHyperlink>,
) -> Result<(Box<[Hyperlink]>, Box<[Option<Box<str>>]>)> {
    let mut values = Vec::new();
    let mut relationship_ids = Vec::new();
    values
        .try_reserve_exact(parsed.len())
        .map_err(|source| allocation("source-backed worksheet hyperlink values", source))?;
    relationship_ids
        .try_reserve_exact(parsed.len())
        .map_err(|source| {
            allocation("source-backed worksheet hyperlink relationship IDs", source)
        })?;
    for parsed in parsed {
        values.push(parsed.value);
        relationship_ids.push(
            parsed
                .relationship_id
                .map(|value| copy_boxed(&value, "hyperlink relationship ID"))
                .transpose()?,
        );
    }
    Ok((
        values.into_boxed_slice(),
        relationship_ids.into_boxed_slice(),
    ))
}

fn capture_relationships(relationships: &Relationships) -> Result<Vec<SourceRelationship>> {
    let mut captured = Vec::new();
    captured
        .try_reserve_exact(relationships.len())
        .map_err(|source| allocation("hyperlink worksheet relationships", source))?;
    for relationship in relationships.iter() {
        captured.push(SourceRelationship::capture(relationship)?);
    }
    captured.sort_by(|left, right| left.id.cmp(&right.id));
    Ok(captured)
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

fn require_selected_worksheet(
    relationship: &Relationship,
    worksheet_uri: &PackURI,
    content_type: &str,
) -> Result<()> {
    if relationship.target_mode() != TargetMode::Internal
        || !matches!(relationship.reltype(), rt::WORKSHEET | rt::STRICT_WORKSHEET)
        || relationship.target_partname()? != *worksheet_uri
        || content_type != ct::SML_WORKSHEET
    {
        return Err(invalid(
            "selected worksheet relationship or content type is invalid",
        ));
    }
    Ok(())
}

fn current_owner_relationship(relationships: &Relationships) -> Option<&Relationship> {
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
        .map_err(|source| allocation(resource, source))?;
    copied.push_str(value);
    Ok(copied.into_boxed_str())
}

fn validate_surface(workbook_xml: &[u8], worksheet_xml: &[u8]) -> Result<bool> {
    reject_mce(workbook_xml, "workbook")?;
    validate_hyperlink_owner(worksheet_xml)?;
    let workbook_protection =
        crate::workbook_metadata::protection::parse_workbook_protection(workbook_xml)?;
    let worksheet_protection = crate::sheet_protection::parse_protection(worksheet_xml)?;
    let protected = workbook_protection.is_some()
        || worksheet_protection.sheet_protection().is_some()
        || !worksheet_protection
            .protected_range_collections()
            .is_empty();
    Ok(protected)
}

fn reject_mce(xml: &[u8], owner: &str) -> Result<()> {
    if xml
        .windows(MCE_NAMESPACE.len())
        .any(|window| window == MCE_NAMESPACE)
    {
        return Err(invalid(format!(
            "XLSX {owner} hyperlink editing refuses markup-compatibility content"
        )));
    }
    Ok(())
}

fn validate_hyperlink_owner(xml: &[u8]) -> Result<()> {
    reject_mce(xml, "worksheet")?;
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().check_end_names = false;
    reader.config_mut().trim_text(false);
    let mut depth = 0usize;
    let mut owner_depth = None;
    let mut owner_seen = false;
    loop {
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| invalid(format!("invalid XLSX worksheet hyperlink XML: {error}")))?;
        match event {
            Event::Start(element) => {
                validate_owner_element(
                    depth,
                    element.name().local_name().as_ref(),
                    resolve_namespace(&namespace),
                    &mut owner_depth,
                    &mut owner_seen,
                )?;
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("XLSX worksheet hyperlink XML depth overflows usize"))?;
            },
            Event::Empty(element) => {
                let local = element.name().local_name();
                validate_owner_element(
                    depth,
                    local.as_ref(),
                    resolve_namespace(&namespace),
                    &mut owner_depth,
                    &mut owner_seen,
                )?;
                if local.as_ref() == b"hyperlinks" {
                    owner_depth = None;
                }
            },
            Event::End(_) => {
                if depth == 0 {
                    return Err(invalid(
                        "XLSX worksheet hyperlink XML depth underflows usize",
                    ));
                }
                if owner_depth == Some(depth) {
                    owner_depth = None;
                }
                depth -= 1;
            },
            Event::Text(text) => {
                if owner_depth.is_some()
                    && text.as_ref().iter().any(|byte| !byte.is_ascii_whitespace())
                {
                    return Err(invalid(
                        "XLSX worksheet hyperlink owner contains unsupported text",
                    ));
                }
            },
            Event::Eof => break,
            _ if owner_depth.is_some() => {
                return Err(invalid(
                    "XLSX worksheet hyperlink owner contains unsupported content",
                ));
            },
            _ => {},
        }
    }
    if depth != 0 || owner_depth.is_some() {
        return Err(invalid("XLSX worksheet hyperlink XML is not balanced"));
    }
    Ok(())
}

fn validate_owner_element(
    depth: usize,
    local: &[u8],
    namespace: Option<&[u8]>,
    owner_depth: &mut Option<usize>,
    owner_seen: &mut bool,
) -> Result<()> {
    if namespace == Some(MCE_NAMESPACE) {
        return Err(invalid(
            "XLSX worksheet hyperlink editing refuses markup-compatibility content",
        ));
    }
    if let Some(expected_depth) = *owner_depth {
        if depth != expected_depth
            || local != b"hyperlink"
            || !matches!(namespace, Some(value) if value == MAIN_NAMESPACE || value == STRICT_MAIN_NAMESPACE)
        {
            return Err(invalid(
                "XLSX worksheet hyperlink owner contains unsupported content",
            ));
        }
        return Ok(());
    }
    if local == b"hyperlinks" {
        if depth != 1
            || !matches!(namespace, Some(value) if value == MAIN_NAMESPACE || value == STRICT_MAIN_NAMESPACE)
            || *owner_seen
        {
            return Err(invalid(
                "XLSX worksheet contains an unsupported or ambiguous hyperlink owner",
            ));
        }
        *owner_seen = true;
        *owner_depth = Some(depth + 1);
    } else if local == b"hyperlink"
        && (*owner_depth != Some(depth)
            || !matches!(namespace, Some(value) if value == MAIN_NAMESPACE || value == STRICT_MAIN_NAMESPACE))
    {
        return Err(invalid(
            "XLSX worksheet contains an unsupported or ambiguous hyperlink owner",
        ));
    }
    Ok(())
}

fn resolve_namespace<'a>(value: &'a ResolveResult<'a>) -> Option<&'a [u8]> {
    match value {
        ResolveResult::Bound(Namespace(value)) => Some(*value),
        ResolveResult::Unbound | ResolveResult::Unknown(_) => None,
    }
}
