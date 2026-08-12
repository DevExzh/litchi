//! Immutable source closure for one value-only worksheet capability.

use std::sync::Arc;

use litchi_core::{Selector as CoreSelector, SourceVersion};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{
    OpcPackage, PackURI, Part, Relationship, Relationships, SourceBackedPackage, TargetMode,
};

use crate::cell::{Cell, Store, Value};
use crate::error::{Error, Result, allocation, invalid};
use crate::workbook::source::validate_sheet_graph;
use crate::{Selector, WorksheetKind, raw};

use super::validation;

/// Exact worksheet values plus the complete package owner state required to
/// publish a one-Part overlay safely.
#[derive(Clone, Debug)]
pub struct Snapshot {
    sheet_name: Box<str>,
    sheet_position: usize,
    cells: Arc<Store>,
    source: SourceState,
}

impl Snapshot {
    pub(super) fn load_source_backed<'a>(
        package: &SourceBackedPackage,
        selector: impl Into<Selector<'a>>,
    ) -> Result<Self> {
        let workbook = package.main_document_part()?;
        validate_package_relationships(package.rels())?;
        if workbook.content_type() != ct::SML_SHEET_MAIN {
            return Err(invalid(
                "value-only edits require an ordinary XLSX workbook",
            ));
        }
        let workbook_xml = workbook.data()?.into_arc();
        validation::workbook_xml(workbook_xml.as_slice())?;
        let catalog = raw::parse_catalog(workbook_xml.as_slice())?;
        let sheet_parts = validate_sheet_graph(package, &workbook, &catalog.sheets)?;
        if catalog.sheets.len() != 1 {
            return Err(invalid(
                "value-only edits currently require exactly one worksheet",
            ));
        }
        validate_workbook_relationships(workbook.rels())?;
        let position = resolve_selector(&catalog.sheets, selector.into())?
            .ok_or_else(|| invalid("value-only worksheet selector did not resolve"))?;
        let sheet = &catalog.sheets[position];
        let sheet_part = &sheet_parts[position];
        if sheet_part.kind != WorksheetKind::Worksheet {
            return Err(Error::NotWorksheet {
                sheet: sheet.name.clone(),
            });
        }
        let worksheet = package.part(&sheet_part.uri)?;
        if !worksheet.rels().is_empty() {
            return Err(invalid("value-only edits refuse worksheet relationships"));
        }
        let sheet_relationship = workbook
            .rels()
            .get(&sheet.relationship_id)
            .ok_or_else(|| invalid("selected worksheet relationship is missing"))?;
        let worksheet_xml = worksheet.data()?.into_arc();
        validation::worksheet_xml(worksheet_xml.as_slice())?;
        let (style_count, auxiliary) = capture_auxiliary_source(package, &workbook)?;
        let owner = unique_owner(package.rels())?;
        Self::from_parts(
            &sheet.name,
            position,
            workbook.partname().clone(),
            workbook.content_type(),
            workbook_xml,
            owner,
            package.rels(),
            workbook.rels(),
            worksheet.partname().clone(),
            worksheet.content_type(),
            worksheet_xml,
            sheet_relationship,
            style_count,
            auxiliary,
            Some(package.source_version()?),
        )
    }

    /// Load and validate one value-only closure from an owning OPC package.
    pub fn load<'a>(package: &OpcPackage, selector: impl Into<Selector<'a>>) -> Result<Self> {
        let workbook = package.main_document_part()?;
        validate_package_relationships(package.rels())?;
        if workbook.content_type() != ct::SML_SHEET_MAIN {
            return Err(invalid(
                "value-only edits require an ordinary XLSX workbook",
            ));
        }
        let workbook_xml = workbook.blob_arc();
        validation::workbook_xml(workbook_xml.as_slice())?;
        let catalog = raw::parse_catalog(workbook_xml.as_slice())?;
        if catalog.sheets.len() != 1 {
            return Err(invalid(
                "value-only edits currently require exactly one worksheet",
            ));
        }
        validate_workbook_relationships(workbook.rels())?;
        let position = resolve_selector(&catalog.sheets, selector.into())?
            .ok_or_else(|| invalid("value-only worksheet selector did not resolve"))?;
        let sheet = &catalog.sheets[position];
        let relationship = workbook
            .rels()
            .get(&sheet.relationship_id)
            .ok_or_else(|| invalid("selected worksheet relationship is missing"))?;
        require_worksheet_relationship(relationship)?;
        let uri = relationship.target_partname()?;
        let worksheet = package.get_part(&uri)?;
        if worksheet.content_type() != ct::SML_WORKSHEET {
            return Err(invalid("selected worksheet content type is invalid"));
        }
        if !worksheet.rels().is_empty() {
            return Err(invalid("value-only edits refuse worksheet relationships"));
        }
        let worksheet_xml = worksheet.blob_arc();
        validation::worksheet_xml(worksheet_xml.as_slice())?;
        let (style_count, auxiliary) = capture_auxiliary(package, workbook)?;
        let owner = unique_owner(package.rels())?;
        Self::from_parts(
            &sheet.name,
            position,
            workbook.partname().clone(),
            workbook.content_type(),
            workbook_xml,
            owner,
            package.rels(),
            workbook.rels(),
            worksheet.partname().clone(),
            worksheet.content_type(),
            worksheet_xml,
            relationship,
            style_count,
            auxiliary,
            None,
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
        package_relationships: &Relationships,
        workbook_relationships: &Relationships,
        worksheet_uri: PackURI,
        worksheet_content_type: &str,
        worksheet_xml: Arc<Vec<u8>>,
        sheet_relationship: &Relationship,
        style_count: u32,
        auxiliary: Box<[PartState]>,
        source_version: Option<SourceVersion>,
    ) -> Result<Self> {
        require_worksheet_relationship(sheet_relationship)?;
        if sheet_relationship.target_partname()? != worksheet_uri {
            return Err(invalid(
                "selected worksheet relationship does not target its captured Part",
            ));
        }
        let cells = raw::worksheet::parse(worksheet_xml.as_slice(), || Ok(None))?;
        validate_style_references(&cells, style_count)?;
        if cells.entries().iter().any(|entry| {
            matches!(entry.cell, Cell::Formula(_) | Cell::Unknown(_))
                || entry.cell_metadata.is_some()
                || entry.value_metadata.is_some()
        }) {
            return Err(invalid(
                "value-only edits refuse formulas, unknown cells, and cell metadata",
            ));
        }
        Ok(Self {
            sheet_name: copy_boxed(sheet_name, "value-only sheet name")?,
            sheet_position,
            cells: Arc::new(cells),
            source: SourceState {
                workbook: PartState::new(workbook_uri, workbook_content_type, workbook_xml)?,
                worksheet: PartState::new(worksheet_uri, worksheet_content_type, worksheet_xml)?,
                owner_relationship: SourceRelationship::capture(owner_relationship)?,
                sheet_relationship: SourceRelationship::capture(sheet_relationship)?,
                package_relationships: capture_relationships(package_relationships)?,
                workbook_relationships: capture_relationships(workbook_relationships)?,
                auxiliary,
                source_version,
            },
        })
    }

    pub(super) fn from_rewritten_source(source: &Self, bytes: Vec<u8>) -> Result<Self> {
        validation::worksheet_xml(&bytes)?;
        let cells = raw::worksheet::parse(&bytes, || Ok(None))?;
        let mut result = source.clone();
        result.cells = Arc::new(cells);
        result.source.worksheet.bytes = Arc::new(bytes);
        Ok(result)
    }

    /// Selected worksheet name.
    #[must_use]
    pub fn sheet_name(&self) -> &str {
        &self.sheet_name
    }

    /// Selected zero-based sheet position.
    #[must_use]
    pub const fn sheet_position(&self) -> usize {
        self.sheet_position
    }

    /// Exact scalar value at a stored coordinate.
    #[must_use]
    pub fn value(&self, address: litchi_sheet::Cell) -> Option<&Value> {
        match self.cells.entry(address).map(|entry| &entry.cell) {
            Some(Cell::Value(value)) => Some(value),
            _ => None,
        }
    }

    /// Exact source worksheet XML.
    #[must_use]
    pub fn source_xml(&self) -> &[u8] {
        self.source.worksheet.bytes.as_slice()
    }

    /// Selected worksheet Part URI.
    #[must_use]
    pub const fn worksheet_part_name(&self) -> &PackURI {
        &self.source.worksheet.uri
    }

    pub(super) fn source_arc(&self) -> Arc<Vec<u8>> {
        Arc::clone(&self.source.worksheet.bytes)
    }

    pub(super) fn same_source(&self, other: &Self) -> bool {
        self.sheet_name == other.sheet_name
            && self.sheet_position == other.sheet_position
            && self.source.same_owner(&other.source)
    }

    pub(super) fn matches_current_source(&self, package: &OpcPackage) -> bool {
        let Ok(workbook) = package.main_document_part() else {
            return false;
        };
        if !self.source.workbook.matches_part(workbook)
            || !unique_owner(package.rels())
                .is_ok_and(|owner| self.source.owner_relationship.matches(owner))
            || !relationships_match(package.rels(), &self.source.package_relationships)
            || !relationships_match(workbook.rels(), &self.source.workbook_relationships)
            || self.source.auxiliary.iter().any(|expected| {
                package.get_part(&expected.uri).map_or(true, |part| {
                    !expected.matches_part(part) || !part.rels().is_empty()
                })
            })
        {
            return false;
        }
        let Some(relationship) = workbook
            .rels()
            .get(self.source.sheet_relationship.id.as_ref())
        else {
            return false;
        };
        if !self.source.sheet_relationship.matches(relationship)
            || relationship.target_partname().ok().as_ref() != Some(&self.source.worksheet.uri)
        {
            return false;
        }
        package
            .get_part(&self.source.worksheet.uri)
            .is_ok_and(|part| self.source.worksheet.matches_part(part) && part.rels().is_empty())
    }
}

#[derive(Clone, Debug)]
struct SourceState {
    workbook: PartState,
    worksheet: PartState,
    owner_relationship: SourceRelationship,
    sheet_relationship: SourceRelationship,
    package_relationships: Box<[SourceRelationship]>,
    workbook_relationships: Box<[SourceRelationship]>,
    auxiliary: Box<[PartState]>,
    source_version: Option<SourceVersion>,
}

impl SourceState {
    fn same_owner(&self, other: &Self) -> bool {
        self.workbook == other.workbook
            && self.worksheet == other.worksheet
            && self.owner_relationship == other.owner_relationship
            && self.sheet_relationship == other.sheet_relationship
            && self.package_relationships == other.package_relationships
            && self.workbook_relationships == other.workbook_relationships
            && self.auxiliary == other.auxiliary
            && match (self.source_version, other.source_version) {
                (Some(left), Some(right)) => left == right,
                (None, _) | (_, None) => true,
            }
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct PartState {
    uri: PackURI,
    content_type: Box<str>,
    bytes: Arc<Vec<u8>>,
}

impl PartState {
    fn new(uri: PackURI, content_type: &str, bytes: Arc<Vec<u8>>) -> Result<Self> {
        Ok(Self {
            uri,
            content_type: copy_boxed(content_type, "value-only content type")?,
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
    kind: Box<str>,
    target: Box<str>,
    mode: TargetMode,
}

impl SourceRelationship {
    fn capture(value: &Relationship) -> Result<Self> {
        Ok(Self {
            id: copy_boxed(value.r_id(), "value-only relationship ID")?,
            kind: copy_boxed(value.reltype(), "value-only relationship type")?,
            target: copy_boxed(value.target_ref(), "value-only relationship target")?,
            mode: value.target_mode(),
        })
    }
    fn matches(&self, value: &Relationship) -> bool {
        value.r_id() == self.id.as_ref()
            && value.reltype() == self.kind.as_ref()
            && value.target_ref() == self.target.as_ref()
            && value.target_mode() == self.mode
    }
}

fn validate_workbook_relationships(relationships: &Relationships) -> Result<()> {
    let mut worksheets = 0usize;
    let mut styles = 0usize;
    let mut themes = 0usize;
    for relationship in relationships.iter() {
        if relationship.is_external() {
            return Err(invalid(
                "value-only edits refuse external workbook relationships",
            ));
        }
        if !matches!(
            relationship.reltype(),
            rt::WORKSHEET | rt::STRICT_WORKSHEET | rt::STYLES | rt::STRICT_STYLES | rt::THEME
        ) {
            return Err(invalid(format!(
                "value-only edits refuse workbook relationship '{}'",
                relationship.reltype()
            )));
        }
        match relationship.reltype() {
            rt::WORKSHEET | rt::STRICT_WORKSHEET => worksheets += 1,
            rt::STYLES | rt::STRICT_STYLES => styles += 1,
            rt::THEME => themes += 1,
            _ => {},
        }
    }
    if worksheets != 1 || styles > 1 || themes > 1 {
        return Err(invalid(
            "value-only edits require one worksheet and at most one styles and theme relationship",
        ));
    }
    Ok(())
}

fn validate_package_relationships(relationships: &Relationships) -> Result<()> {
    let mut owners = 0usize;
    for relationship in relationships.iter() {
        if relationship.is_external()
            || !matches!(
                relationship.reltype(),
                rt::OFFICE_DOCUMENT | rt::STRICT_OFFICE_DOCUMENT | rt::DIGITAL_SIGNATURE_ORIGIN
            )
        {
            return Err(invalid(format!(
                "value-only edits refuse package relationship '{}'",
                relationship.reltype()
            )));
        }
        if matches!(
            relationship.reltype(),
            rt::OFFICE_DOCUMENT | rt::STRICT_OFFICE_DOCUMENT
        ) {
            owners += 1;
        }
    }
    if owners != 1 {
        return Err(invalid(
            "value-only edits require exactly one package officeDocument owner",
        ));
    }
    Ok(())
}

fn capture_auxiliary_source(
    package: &SourceBackedPackage,
    workbook: &litchi_opc::PartView<'_>,
) -> Result<(u32, Box<[PartState]>)> {
    let mut style_count = 0;
    let mut auxiliary = Vec::new();
    auxiliary
        .try_reserve_exact(2)
        .map_err(|source| allocation("value-only auxiliary closure", source))?;
    for relationship in workbook.rels().iter().filter(|relationship| {
        matches!(
            relationship.reltype(),
            rt::STYLES | rt::STRICT_STYLES | rt::THEME
        )
    }) {
        let part = package.part(&relationship.target_partname()?)?;
        if !part.rels().is_empty() {
            return Err(invalid(
                "value-only edits refuse styles and theme relationships",
            ));
        }
        let data = part.data()?.into_arc();
        match relationship.reltype() {
            rt::STYLES | rt::STRICT_STYLES if part.content_type() == ct::SML_STYLES => {
                style_count = raw::styles::parse(data.as_slice())?.len();
            },
            rt::THEME if part.content_type() == ct::OFC_THEME => {},
            rt::STYLES | rt::STRICT_STYLES => {
                return Err(invalid("styles relationship has the wrong content type"));
            },
            rt::THEME => return Err(invalid("theme relationship has the wrong content type")),
            _ => {},
        }
        auxiliary.push(PartState::new(
            part.partname().clone(),
            part.content_type(),
            data,
        )?);
    }
    Ok((style_count, auxiliary.into_boxed_slice()))
}

fn capture_auxiliary(package: &OpcPackage, workbook: &dyn Part) -> Result<(u32, Box<[PartState]>)> {
    let mut style_count = 0;
    let mut auxiliary = Vec::new();
    auxiliary
        .try_reserve_exact(2)
        .map_err(|source| allocation("value-only auxiliary closure", source))?;
    for relationship in workbook.rels().iter().filter(|relationship| {
        matches!(
            relationship.reltype(),
            rt::STYLES | rt::STRICT_STYLES | rt::THEME
        )
    }) {
        let part = package.get_part(&relationship.target_partname()?)?;
        if !part.rels().is_empty() {
            return Err(invalid(
                "value-only edits refuse styles and theme relationships",
            ));
        }
        match relationship.reltype() {
            rt::STYLES | rt::STRICT_STYLES if part.content_type() == ct::SML_STYLES => {
                style_count = raw::styles::parse(part.blob())?.len();
            },
            rt::THEME if part.content_type() == ct::OFC_THEME => {},
            rt::STYLES | rt::STRICT_STYLES => {
                return Err(invalid("styles relationship has the wrong content type"));
            },
            rt::THEME => return Err(invalid("theme relationship has the wrong content type")),
            _ => {},
        }
        auxiliary.push(PartState::new(
            part.partname().clone(),
            part.content_type(),
            part.blob_arc(),
        )?);
    }
    Ok((style_count, auxiliary.into_boxed_slice()))
}

fn validate_style_references(cells: &Store, style_count: u32) -> Result<()> {
    let invalid_cell = cells
        .entries()
        .iter()
        .any(|entry| entry.style.is_some_and(|style| style >= style_count));
    let invalid_row = cells.row_entries().iter().any(|entry| {
        entry
            .properties
            .style
            .is_some_and(|style| style >= style_count)
    });
    let invalid_column = cells.column_entries().iter().any(|entry| {
        entry
            .properties
            .style
            .is_some_and(|style| style >= style_count)
    });
    if invalid_cell || invalid_row || invalid_column {
        return Err(invalid(
            "worksheet references a shared style outside the styles table",
        ));
    }
    Ok(())
}

fn require_worksheet_relationship(relationship: &Relationship) -> Result<()> {
    if relationship.target_mode() != TargetMode::Internal
        || !matches!(relationship.reltype(), rt::WORKSHEET | rt::STRICT_WORKSHEET)
    {
        return Err(invalid("selected worksheet relationship is invalid"));
    }
    Ok(())
}

fn unique_owner(relationships: &Relationships) -> Result<&Relationship> {
    let mut owners = relationships.iter().filter(|relationship| {
        matches!(
            relationship.reltype(),
            rt::OFFICE_DOCUMENT | rt::STRICT_OFFICE_DOCUMENT
        )
    });
    let owner = owners
        .next()
        .ok_or_else(|| invalid("workbook has no officeDocument owner"))?;
    if owners.next().is_some() || owner.is_external() {
        return Err(invalid(
            "workbook officeDocument owner is not unique and internal",
        ));
    }
    Ok(owner)
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

fn copy_boxed(value: &str, resource: &'static str) -> Result<Box<str>> {
    let mut output = String::new();
    output
        .try_reserve_exact(value.len())
        .map_err(|source| allocation(resource, source))?;
    output.push_str(value);
    Ok(output.into_boxed_str())
}

fn capture_relationships(values: &Relationships) -> Result<Box<[SourceRelationship]>> {
    let mut output = Vec::new();
    output
        .try_reserve_exact(values.len())
        .map_err(|source| allocation("value-only relationship closure", source))?;
    for value in values.iter() {
        output.push(SourceRelationship::capture(value)?);
    }
    Ok(output.into_boxed_slice())
}

fn relationships_match(values: &Relationships, expected: &[SourceRelationship]) -> bool {
    values.len() == expected.len()
        && values
            .iter()
            .zip(expected)
            .all(|(value, expected)| expected.matches(value))
}
