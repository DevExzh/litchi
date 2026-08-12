//! Immutable source-bound worksheet auto-filter state.

use std::sync::Arc;

use litchi_core::Selector as CoreSelector;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{
    OpcPackage, PackURI, Part, PartView, Relationship, Relationships, SourceBackedPackage,
    TargetMode,
};

use super::{Definition, Payload};
use crate::error::{Error, Result, invalid};
use crate::workbook::source::validate_sheet_graph;
use crate::{Selector, Workbook, WorksheetKind, raw};

/// Semantic auto-filter state plus the exact worksheet dependency closure.
#[derive(Clone, Debug)]
pub struct Snapshot {
    value: Option<Arc<Definition>>,
    sheet_name: Box<str>,
    sheet_position: usize,
    filter_locked: bool,
    sort_locked: bool,
    source: SourceState,
}

impl Snapshot {
    /// Resolve and read one worksheet by semantic selector.
    pub fn load<'a>(package: &OpcPackage, selector: impl Into<Selector<'a>>) -> Result<Self> {
        let workbook = Workbook::from_package(package.clone())?;
        let sheet = workbook
            .sheet(selector)?
            .ok_or_else(|| invalid("auto-filter worksheet selector did not resolve"))?;
        if sheet.kind() != WorksheetKind::Worksheet {
            return Err(Error::NotWorksheet {
                sheet: sheet.name().to_owned(),
            });
        }
        let workbook_part = package.main_document_part()?;
        let catalog = raw::parse_catalog(workbook_part.blob())?;
        let catalog_sheet = catalog.sheets.get(sheet.position()).ok_or_else(|| {
            invalid("auto-filter worksheet position is absent from the workbook catalog")
        })?;
        if catalog_sheet.name != sheet.name() {
            return Err(invalid(
                "auto-filter worksheet name differs from the workbook catalog",
            ));
        }
        let worksheet = package.get_part(sheet.part_uri())?;
        let sheet_relationship = require_selected_worksheet(
            workbook_part.rels(),
            catalog_sheet,
            worksheet.partname(),
            worksheet.content_type(),
        )?;
        let owner_relationship = current_owner_relationship(package.rels())
            .ok_or_else(|| invalid("workbook has no unique officeDocument owner"))?;
        let value = super::parse_auto_filter(worksheet.blob())?;
        let styles = load_styles_owned(package, workbook_part)?;
        Self::from_parts(
            sheet.name(),
            sheet.position(),
            workbook_part.partname().clone(),
            workbook_part.content_type(),
            workbook_part.blob_arc(),
            owner_relationship,
            worksheet.partname().clone(),
            worksheet.content_type(),
            worksheet.blob_arc(),
            sheet_relationship,
            worksheet.rels(),
            value,
            styles,
        )
    }

    pub(super) fn load_source_backed<'a>(
        package: &SourceBackedPackage,
        selector: impl Into<Selector<'a>>,
    ) -> Result<Self> {
        let workbook = package.main_document_part()?;
        require_workbook_content_type(workbook.content_type())?;
        let workbook_xml = workbook.data()?.into_arc();
        let catalog = raw::parse_catalog(workbook_xml.as_slice())?;
        let sheet_parts = validate_sheet_graph(package, &workbook, &catalog.sheets)?;
        let sheet_position = resolve_selector(&catalog.sheets, selector.into())?
            .ok_or_else(|| invalid("auto-filter worksheet selector did not resolve"))?;
        let catalog_sheet = &catalog.sheets[sheet_position];
        let sheet_part = &sheet_parts[sheet_position];
        if sheet_part.kind != WorksheetKind::Worksheet {
            return Err(Error::NotWorksheet {
                sheet: catalog_sheet.name.clone(),
            });
        }
        let worksheet = package.part(&sheet_part.uri)?;
        let sheet_relationship = require_selected_worksheet(
            workbook.rels(),
            catalog_sheet,
            worksheet.partname(),
            worksheet.content_type(),
        )?;
        let worksheet_xml = worksheet.data()?.into_arc();
        let owner_relationship = current_owner_relationship(package.rels())
            .ok_or_else(|| invalid("workbook has no unique officeDocument owner"))?;
        let value = super::parse_auto_filter(worksheet_xml.as_slice())?;
        let styles = load_styles_source_backed(package, &workbook)?;
        Self::from_parts(
            &catalog_sheet.name,
            sheet_position,
            workbook.partname().clone(),
            workbook.content_type(),
            workbook_xml,
            owner_relationship,
            worksheet.partname().clone(),
            worksheet.content_type(),
            worksheet_xml,
            sheet_relationship,
            worksheet.rels(),
            value,
            styles,
        )
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
        worksheet_relationships: &Relationships,
        value: Option<Definition>,
        styles: Option<StylesState>,
    ) -> Result<Self> {
        validate_style_references(value.as_ref(), styles.as_ref())?;
        let protection = crate::sheet_protection::parse_protection(worksheet_xml.as_slice())?;
        let (filter_locked, sort_locked) = protection
            .sheet_protection()
            .map_or((false, false), |value| {
                (value.auto_filter_locked(), value.sort_locked())
            });
        Ok(Self {
            value: value.map(Arc::new),
            sheet_name: copy_boxed(sheet_name, "auto-filter sheet name")?,
            sheet_position,
            filter_locked,
            sort_locked,
            source: SourceState {
                workbook: PartState::new(
                    workbook_uri,
                    workbook_content_type,
                    workbook_xml,
                    "auto-filter workbook content type",
                )?,
                worksheet: PartState::new(
                    worksheet_uri,
                    worksheet_content_type,
                    worksheet_xml,
                    "auto-filter worksheet content type",
                )?,
                owner_relationship: SourceRelationship::capture(owner_relationship)?,
                sheet_relationship: SourceRelationship::capture(sheet_relationship)?,
                worksheet_relationships: capture_relationships(worksheet_relationships)?,
                styles,
            },
        })
    }

    pub(super) fn from_rewritten_source(
        source: &Self,
        bytes: Vec<u8>,
        value: Option<Definition>,
    ) -> Result<Self> {
        validate_style_references(value.as_ref(), source.source.styles.as_ref())?;
        let mut rewritten = source.clone();
        rewritten.value = value.map(Arc::new);
        rewritten.source.worksheet.bytes = Arc::new(bytes);
        Ok(rewritten)
    }

    /// Complete typed direct worksheet auto-filter and sort state.
    #[must_use]
    pub fn auto_filter(&self) -> Option<&Definition> {
        self.value.as_deref()
    }

    /// Complete typed direct worksheet auto-filter and sort state.
    #[must_use]
    pub fn definition(&self) -> Option<&Definition> {
        self.value.as_deref()
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

    /// Resolved worksheet Part name.
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

    pub(super) fn mutation_locked(&self, after: Option<&Definition>) -> bool {
        let filter_changed = filter_projection(self.auto_filter()) != filter_projection(after);
        let sort_changed = self.auto_filter().and_then(Definition::sort_state)
            != after.and_then(Definition::sort_state);
        (filter_changed && self.filter_locked) || (sort_changed && self.sort_locked)
    }

    pub(super) fn same_source(&self, other: &Self) -> bool {
        self.sheet_name == other.sheet_name
            && self.sheet_position == other.sheet_position
            && self.filter_locked == other.filter_locked
            && self.sort_locked == other.sort_locked
            && self.source == other.source
    }

    pub(super) fn matches_source_backed(&self, package: &SourceBackedPackage) -> Result<bool> {
        let workbook = package.main_document_part()?;
        require_workbook_content_type(workbook.content_type())?;
        let workbook_xml = workbook.data()?;
        if !self
            .source
            .workbook
            .matches_view(&workbook, workbook_xml.as_bytes())
            || !current_owner_relationship(package.rels())
                .is_some_and(|value| self.source.owner_relationship.matches(value))
            || !styles_match_source_backed(self.source.styles.as_ref(), package, &workbook)?
        {
            return Ok(false);
        }
        let catalog = raw::parse_catalog(workbook_xml.as_bytes())?;
        let sheet_parts = validate_sheet_graph(package, &workbook, &catalog.sheets)?;
        let Some(catalog_sheet) = catalog.sheets.get(self.sheet_position) else {
            return Ok(false);
        };
        let Some(sheet_part) = sheet_parts.get(self.sheet_position) else {
            return Ok(false);
        };
        if catalog_sheet.name != self.sheet_name.as_ref()
            || sheet_part.kind != WorksheetKind::Worksheet
            || sheet_part.uri != self.source.worksheet.uri
        {
            return Ok(false);
        }
        let worksheet = package.part(&sheet_part.uri)?;
        let relationship = require_selected_worksheet(
            workbook.rels(),
            catalog_sheet,
            worksheet.partname(),
            worksheet.content_type(),
        )?;
        let worksheet_xml = worksheet.data()?;
        Ok(self.source.sheet_relationship.matches(relationship)
            && self
                .source
                .worksheet
                .matches_view(&worksheet, worksheet_xml.as_bytes())
            && relationships_match(&self.source.worksheet_relationships, worksheet.rels()))
    }

    pub(super) fn matches_current_source(&self, package: &OpcPackage) -> bool {
        let Ok(workbook) = package.main_document_part() else {
            return false;
        };
        if !self.source.workbook.matches_part(workbook)
            || !current_owner_relationship(package.rels())
                .is_some_and(|value| self.source.owner_relationship.matches(value))
            || !styles_match_owned(self.source.styles.as_ref(), package, workbook)
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
            .is_ok_and(|part| {
                self.source.worksheet.matches_part(part)
                    && relationships_match(&self.source.worksheet_relationships, part.rels())
            })
    }
}

fn filter_projection(
    value: Option<&Definition>,
) -> Option<(Option<&super::Range>, &[super::Column])> {
    value.map(|value| (value.reference(), value.columns()))
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct SourceState {
    workbook: PartState,
    worksheet: PartState,
    owner_relationship: SourceRelationship,
    sheet_relationship: SourceRelationship,
    worksheet_relationships: Vec<SourceRelationship>,
    styles: Option<StylesState>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct StylesState {
    part: PartState,
    relationship: SourceRelationship,
    differential_format_count: usize,
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

    fn matches_view(&self, part: &PartView<'_>, bytes: &[u8]) -> bool {
        part.partname() == &self.uri
            && part.content_type() == self.content_type.as_ref()
            && bytes == self.bytes.as_slice()
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
    fn capture(value: &Relationship) -> Result<Self> {
        Ok(Self {
            id: copy_boxed(value.r_id(), "auto-filter relationship ID")?,
            relationship_type: copy_boxed(value.reltype(), "auto-filter relationship type")?,
            target: copy_boxed(value.target_ref(), "auto-filter relationship target")?,
            mode: value.target_mode(),
        })
    }

    fn matches(&self, value: &Relationship) -> bool {
        value.r_id() == self.id.as_ref()
            && value.reltype() == self.relationship_type.as_ref()
            && value.target_ref() == self.target.as_ref()
            && value.target_mode() == self.mode
    }
}

fn validate_style_references(
    value: Option<&Definition>,
    styles: Option<&StylesState>,
) -> Result<()> {
    let Some(maximum) = maximum_differential_format(value) else {
        return Ok(());
    };
    let count = styles.map_or(0, |value| value.differential_format_count);
    if usize::try_from(maximum).map_or(true, |index| index >= count) {
        return Err(invalid(format!(
            "auto-filter dxfId {maximum} is outside the workbook differential formats"
        )));
    }
    Ok(())
}

fn maximum_differential_format(value: Option<&Definition>) -> Option<u32> {
    let value = value?;
    value
        .columns()
        .iter()
        .filter_map(|column| match column.payload() {
            Some(Payload::Color(value)) => Some(value.differential_format_id()),
            _ => None,
        })
        .chain(
            value
                .sort_state()
                .into_iter()
                .flat_map(|state| state.conditions())
                .filter_map(super::Condition::differential_format_id),
        )
        .max()
}

fn load_styles_owned(package: &OpcPackage, workbook: &dyn Part) -> Result<Option<StylesState>> {
    let Some(relationship) = unique_styles_relationship(workbook.rels())? else {
        return Ok(None);
    };
    let uri = relationship.target_partname()?;
    let part = package.get_part(&uri)?;
    require_styles_content_type(part.content_type())?;
    let bytes = part.blob_arc();
    let count = crate::conditional_formatting::parse_differential_formats(bytes.as_slice())?.len();
    Ok(Some(StylesState {
        part: PartState::new(
            uri,
            part.content_type(),
            bytes,
            "auto-filter styles content type",
        )?,
        relationship: SourceRelationship::capture(relationship)?,
        differential_format_count: count,
    }))
}

fn load_styles_source_backed(
    package: &SourceBackedPackage,
    workbook: &PartView<'_>,
) -> Result<Option<StylesState>> {
    let Some(relationship) = unique_styles_relationship(workbook.rels())? else {
        return Ok(None);
    };
    let uri = relationship.target_partname()?;
    let part = package.part(&uri)?;
    require_styles_content_type(part.content_type())?;
    let bytes = part.data()?.into_arc();
    let count = crate::conditional_formatting::parse_differential_formats(bytes.as_slice())?.len();
    Ok(Some(StylesState {
        part: PartState::new(
            uri,
            part.content_type(),
            bytes,
            "auto-filter styles content type",
        )?,
        relationship: SourceRelationship::capture(relationship)?,
        differential_format_count: count,
    }))
}

fn styles_match_source_backed(
    expected: Option<&StylesState>,
    package: &SourceBackedPackage,
    workbook: &PartView<'_>,
) -> Result<bool> {
    let actual = unique_styles_relationship(workbook.rels())?;
    match (expected, actual) {
        (None, None) => Ok(true),
        (Some(expected), Some(relationship)) if expected.relationship.matches(relationship) => {
            let part = package.part(&relationship.target_partname()?)?;
            let bytes = part.data()?;
            Ok(expected.part.matches_view(&part, bytes.as_bytes()))
        },
        _ => Ok(false),
    }
}

fn styles_match_owned(
    expected: Option<&StylesState>,
    package: &OpcPackage,
    workbook: &dyn Part,
) -> bool {
    let Ok(actual) = unique_styles_relationship(workbook.rels()) else {
        return false;
    };
    match (expected, actual) {
        (None, None) => true,
        (Some(expected), Some(relationship)) if expected.relationship.matches(relationship) => {
            relationship
                .target_partname()
                .ok()
                .and_then(|uri| package.get_part(&uri).ok())
                .is_some_and(|part| expected.part.matches_part(part))
        },
        _ => false,
    }
}

fn unique_styles_relationship(relationships: &Relationships) -> Result<Option<&Relationship>> {
    let mut matching = relationships
        .iter()
        .filter(|value| matches!(value.reltype(), rt::STYLES | rt::STRICT_STYLES));
    let Some(value) = matching.next() else {
        return Ok(None);
    };
    if matching.next().is_some() || value.target_mode() != TargetMode::Internal {
        return Err(invalid(
            "workbook has ambiguous or external auto-filter styles ownership",
        ));
    }
    Ok(Some(value))
}

fn require_styles_content_type(value: &str) -> Result<()> {
    if value == ct::SML_STYLES {
        Ok(())
    } else {
        Err(invalid(format!(
            "auto-filter styles part has content type '{value}'"
        )))
    }
}

fn capture_relationships(relationships: &Relationships) -> Result<Vec<SourceRelationship>> {
    let mut captured = Vec::new();
    captured
        .try_reserve_exact(relationships.len())
        .map_err(|source| Error::Allocation {
            resource: "auto-filter worksheet relationships",
            source,
        })?;
    for relationship in relationships.iter() {
        captured.push(SourceRelationship::capture(relationship)?);
    }
    captured.sort_by(|left, right| left.id.cmp(&right.id));
    Ok(captured)
}

fn relationships_match(expected: &[SourceRelationship], actual: &Relationships) -> bool {
    expected.len() == actual.len()
        && expected.iter().all(|value| {
            actual
                .get(value.id.as_ref())
                .is_some_and(|actual| value.matches(actual))
        })
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
    relationships: &'a Relationships,
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

fn current_owner_relationship(relationships: &Relationships) -> Option<&Relationship> {
    let mut owners = relationships.iter().filter(|value| {
        matches!(
            value.reltype(),
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
