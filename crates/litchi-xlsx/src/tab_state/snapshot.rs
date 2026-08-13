//! Immutable workbook tab state bound to an exact package closure.

use std::sync::Arc;

use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{
    OpcPackage, PackURI, Part, PartView, Relationship, Relationships, SourceBackedPackage,
    TargetMode,
};

use crate::error::{Error, Result, TabEditBlock, allocation, invalid};
use crate::workbook::source::validate_sheet_graph;
use crate::{Visibility, WorksheetKind, raw};

const CHARTSHEET_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet";
const STRICT_CHARTSHEET_REL: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/chartsheet";
const DIALOGSHEET_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/dialogsheet";
const MACROSHEET_REL: &str = "http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet";
const INTL_MACROSHEET_REL: &str =
    "http://schemas.microsoft.com/office/2006/relationships/xlIntlMacrosheet";

/// One existing workbook tab in catalog order.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Tab {
    name: Box<str>,
    visibility: Visibility,
    active: bool,
}

impl Tab {
    /// Developer-facing workbook sheet name.
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Exact recognized visibility state.
    #[must_use]
    pub const fn visibility(&self) -> &Visibility {
        &self.visibility
    }

    /// Whether this is the first workbook view's active tab.
    #[must_use]
    pub const fn is_active(&self) -> bool {
        self.active
    }
}

/// Semantic tab state plus its exact workbook and touched-sheet closure.
#[derive(Clone, Debug)]
pub struct Snapshot {
    tabs: Arc<[Tab]>,
    active: usize,
    pub(super) source: SourceState,
    origin: Option<Arc<()>>,
}

impl Snapshot {
    /// Capture an ordinary materialized package without selecting sheet bodies.
    pub fn load(package: &OpcPackage) -> Result<Self> {
        Self::load_owned_with_touched(package, &[])
    }

    pub(super) fn load_source_backed(
        package: &SourceBackedPackage,
        origin: Arc<()>,
    ) -> Result<Self> {
        let workbook = package.main_document_part()?;
        require_workbook_content_type(workbook.content_type())?;
        let workbook_bytes = workbook.data()?.into_arc()?;
        let catalog = raw::parse_catalog(workbook_bytes.as_slice())?;
        let graph = capture_source_graph(package, &workbook, &catalog.sheets)?;
        let (tabs, active) = semantic_tabs(&catalog)?;
        Ok(Self {
            tabs,
            active,
            source: SourceState {
                workbook: PartState::new(
                    workbook.partname().clone(),
                    workbook.content_type(),
                    workbook_bytes,
                    "tab-state workbook content type",
                )?,
                graph: Arc::new(graph),
                touched: Arc::from([]),
            },
            origin: Some(origin),
        })
    }

    fn load_owned_with_touched(package: &OpcPackage, positions: &[usize]) -> Result<Self> {
        let workbook = package.main_document_part()?;
        require_workbook_content_type(workbook.content_type())?;
        let workbook_bytes = workbook.blob_arc();
        let catalog = raw::parse_catalog(workbook_bytes.as_slice())?;
        let graph = capture_owned_graph(package, workbook, &catalog.sheets)?;
        let touched = capture_owned_touched(package, &graph, positions)?;
        let (tabs, active) = semantic_tabs(&catalog)?;
        Ok(Self {
            tabs,
            active,
            source: SourceState {
                workbook: PartState::new(
                    workbook.partname().clone(),
                    workbook.content_type(),
                    workbook_bytes,
                    "tab-state workbook content type",
                )?,
                graph: Arc::new(graph),
                touched: touched.into(),
            },
            origin: None,
        })
    }

    pub(super) fn with_source_touched(
        &self,
        package: &SourceBackedPackage,
        positions: &[usize],
    ) -> Result<Self> {
        if !self.matches_source_backed(package, self.origin.as_ref())? {
            return Err(Error::PatchConflict {
                part: self.workbook_part_name().to_string(),
            });
        }
        let touched = capture_source_touched(package, &self.source.graph, positions)?;
        let mut snapshot = self.clone();
        snapshot.source.touched = touched.into();
        Ok(snapshot)
    }

    pub(super) fn rewritten(
        before: &Self,
        workbook: Vec<u8>,
        touched: Vec<(usize, Vec<u8>)>,
    ) -> Result<Self> {
        let catalog = raw::parse_catalog(&workbook)?;
        let (tabs, active) = semantic_tabs(&catalog)?;
        if catalog.sheets.len() != before.source.graph.sheets.len()
            || catalog
                .sheets
                .iter()
                .zip(before.source.graph.sheets.iter())
                .any(|(sheet, bound)| {
                    sheet.name.as_str() != bound.name.as_ref()
                        || sheet.sheet_id != bound.sheet_id
                        || sheet.relationship_id != bound.relationship.id.as_ref()
                })
        {
            return Err(invalid(
                "tab-state rewrite changed workbook sheet topology or relationships",
            ));
        }

        let mut rewritten_parts = Vec::new();
        rewritten_parts
            .try_reserve_exact(before.source.touched.len())
            .map_err(|source| allocation("rewritten tab-state sheet closure", source))?;
        for source in before.source.touched.iter() {
            let bytes = touched
                .iter()
                .find_map(|(position, bytes)| (*position == source.position).then_some(bytes))
                .ok_or_else(|| invalid("tab-state rewrite omitted a touched sheet Part"))?;
            rewritten_parts.push(TouchedPart {
                position: source.position,
                part: PartState::new(
                    source.part.uri.clone(),
                    &source.part.content_type,
                    Arc::new(bytes.clone()),
                    "tab-state touched sheet content type",
                )?,
                relationships: Arc::clone(&source.relationships),
            });
        }

        Ok(Self {
            tabs,
            active,
            source: SourceState {
                workbook: PartState::new(
                    before.source.workbook.uri.clone(),
                    &before.source.workbook.content_type,
                    Arc::new(workbook),
                    "tab-state workbook content type",
                )?,
                graph: Arc::clone(&before.source.graph),
                touched: rewritten_parts.into(),
            },
            origin: before.origin.clone(),
        })
    }

    /// Tabs in workbook catalog order.
    #[must_use]
    pub fn tabs(&self) -> &[Tab] {
        &self.tabs
    }

    /// Zero-based active-tab position.
    #[must_use]
    pub const fn active_position(&self) -> usize {
        self.active
    }

    /// Active tab, when the workbook contains a sheet catalog.
    #[must_use]
    pub fn active_tab(&self) -> Option<&Tab> {
        self.tabs.get(self.active)
    }

    /// Resolved workbook Part name.
    #[must_use]
    pub const fn workbook_part_name(&self) -> &PackURI {
        &self.source.workbook.uri
    }

    /// Exact source workbook XML.
    #[must_use]
    pub fn workbook_xml(&self) -> &[u8] {
        self.source.workbook.bytes.as_slice()
    }

    /// Shared exact source workbook XML.
    #[must_use]
    pub fn workbook_source_arc(&self) -> Arc<Vec<u8>> {
        Arc::clone(&self.source.workbook.bytes)
    }

    pub(super) fn binding(&self, position: usize) -> Result<&SheetBinding> {
        self.source
            .graph
            .sheets
            .get(position)
            .ok_or_else(|| invalid("tab selector position is outside the workbook catalog"))
    }

    pub(super) fn touched(&self) -> &[TouchedPart] {
        &self.source.touched
    }

    pub(super) fn same_source(&self, other: &Self) -> bool {
        self.source == other.source
    }

    pub(super) fn same_semantics(&self, other: &Self) -> bool {
        self.tabs == other.tabs && self.active == other.active
    }

    pub(super) fn belongs_to(&self, origin: &Arc<()>) -> bool {
        self.origin
            .as_ref()
            .is_some_and(|candidate| Arc::ptr_eq(candidate, origin))
    }

    pub(super) fn matches_source_backed(
        &self,
        package: &SourceBackedPackage,
        origin: Option<&Arc<()>>,
    ) -> Result<bool> {
        if let Some(origin) = origin
            && !self.belongs_to(origin)
        {
            return Ok(false);
        }
        let workbook = package.main_document_part()?;
        require_workbook_content_type(workbook.content_type())?;
        let bytes = workbook.data()?;
        if !self
            .source
            .workbook
            .matches_view(&workbook, bytes.as_bytes())
        {
            return Ok(false);
        }
        let catalog = raw::parse_catalog(bytes.as_bytes())?;
        let graph = capture_source_graph(package, &workbook, &catalog.sheets)?;
        if graph != *self.source.graph {
            return Ok(false);
        }
        for expected in self.source.touched.iter() {
            let part = package.part(&expected.part.uri)?;
            let data = part.data()?;
            if !expected.part.matches_view(&part, data.as_bytes())
                || !relationships_match(&expected.relationships, part.rels())
            {
                return Ok(false);
            }
        }
        Ok(true)
    }

    pub(super) fn matches_current_source(&self, package: &OpcPackage) -> bool {
        let positions = self
            .source
            .touched
            .iter()
            .map(|part| part.position)
            .collect::<Vec<_>>();
        Self::load_owned_with_touched(package, &positions)
            .is_ok_and(|current| current.source == self.source)
    }

    pub(super) fn load_owned_target(package: &OpcPackage, target: &Self) -> Result<Self> {
        let positions = target
            .source
            .touched
            .iter()
            .map(|part| part.position)
            .collect::<Vec<_>>();
        Self::load_owned_with_touched(package, &positions)
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub(super) struct SourceState {
    pub(super) workbook: PartState,
    graph: Arc<GraphBinding>,
    touched: Arc<[TouchedPart]>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct GraphBinding {
    owner_relationship: SourceRelationship,
    workbook_relationships: Arc<[SourceRelationship]>,
    sheets: Arc<[SheetBinding]>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub(super) struct SheetBinding {
    pub(super) name: Box<str>,
    sheet_id: u32,
    pub(super) kind: WorksheetKind,
    pub(super) part: PartDescriptor,
    relationship: SourceRelationship,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub(super) struct PartDescriptor {
    pub(super) uri: PackURI,
    content_type: Box<str>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub(super) struct TouchedPart {
    pub(super) position: usize,
    pub(super) part: PartState,
    relationships: Arc<[SourceRelationship]>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub(super) struct PartState {
    pub(super) uri: PackURI,
    content_type: Box<str>,
    pub(super) bytes: Arc<Vec<u8>>,
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
            id: copy_boxed(value.r_id(), "tab-state relationship ID")?,
            relationship_type: copy_boxed(value.reltype(), "tab-state relationship type")?,
            target: copy_boxed(value.target_ref(), "tab-state relationship target")?,
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

fn semantic_tabs(catalog: &raw::Catalog) -> Result<(Arc<[Tab]>, usize)> {
    if catalog.sheets.is_empty() {
        return Err(invalid(
            "tab-state editing requires at least one workbook sheet",
        ));
    }
    if catalog.active_sheet_index >= catalog.sheets.len() {
        return Err(invalid("workbook active tab is outside the sheet catalog"));
    }
    let mut tabs = Vec::new();
    tabs.try_reserve_exact(catalog.sheets.len())
        .map_err(|source| allocation("tab-state semantic catalog", source))?;
    for (position, sheet) in catalog.sheets.iter().enumerate() {
        tabs.push(Tab {
            name: copy_boxed(&sheet.name, "tab-state sheet name")?,
            visibility: visibility(&sheet.visibility),
            active: position == catalog.active_sheet_index,
        });
    }
    Ok((tabs.into(), catalog.active_sheet_index))
}

fn capture_source_graph(
    package: &SourceBackedPackage,
    workbook: &PartView<'_>,
    sheets: &[raw::Sheet],
) -> Result<GraphBinding> {
    let parts = validate_sheet_graph(package, workbook, sheets)?;
    let mut bindings = Vec::new();
    bindings
        .try_reserve_exact(sheets.len())
        .map_err(|source| allocation("tab-state sheet graph", source))?;
    for (sheet, resolved) in sheets.iter().zip(parts) {
        let relationship = workbook
            .rels()
            .get(&sheet.relationship_id)
            .ok_or_else(|| invalid("tab-state sheet relationship disappeared"))?;
        let part = package.part(&resolved.uri)?;
        bindings.push(SheetBinding {
            name: copy_boxed(&sheet.name, "tab-state graph sheet name")?,
            sheet_id: sheet.sheet_id,
            kind: resolved.kind,
            part: PartDescriptor {
                uri: resolved.uri,
                content_type: copy_boxed(part.content_type(), "tab-state sheet content type")?,
            },
            relationship: SourceRelationship::capture(relationship)?,
        });
    }
    graph_binding(package.rels(), workbook.rels(), bindings)
}

fn capture_owned_graph(
    package: &OpcPackage,
    workbook: &dyn Part,
    sheets: &[raw::Sheet],
) -> Result<GraphBinding> {
    let mut bindings = Vec::new();
    bindings
        .try_reserve_exact(sheets.len())
        .map_err(|source| allocation("tab-state sheet graph", source))?;
    let mut targets = Vec::<PackURI>::new();
    targets
        .try_reserve_exact(sheets.len())
        .map_err(|source| allocation("tab-state sheet target set", source))?;
    for sheet in sheets {
        let relationship = workbook
            .rels()
            .get(&sheet.relationship_id)
            .ok_or_else(|| invalid("tab-state sheet relationship is missing"))?;
        if relationship.target_mode() != TargetMode::Internal {
            return Err(invalid("tab-state sheet relationship cannot be external"));
        }
        let uri = relationship.target_partname()?;
        if targets.contains(&uri) {
            return Err(invalid("multiple workbook sheets target one Part"));
        }
        targets.push(uri.clone());
        let part = package.get_part(&uri)?;
        let kind = sheet_kind(relationship.reltype(), part.content_type())?;
        bindings.push(SheetBinding {
            name: copy_boxed(&sheet.name, "tab-state graph sheet name")?,
            sheet_id: sheet.sheet_id,
            kind,
            part: PartDescriptor {
                uri,
                content_type: copy_boxed(part.content_type(), "tab-state sheet content type")?,
            },
            relationship: SourceRelationship::capture(relationship)?,
        });
    }
    graph_binding(package.rels(), workbook.rels(), bindings)
}

fn graph_binding(
    package_relationships: &Relationships,
    workbook_relationships: &Relationships,
    sheets: Vec<SheetBinding>,
) -> Result<GraphBinding> {
    let owner = unique_owner(package_relationships)
        .ok_or_else(|| invalid("workbook has no unique internal officeDocument owner"))?;
    Ok(GraphBinding {
        owner_relationship: SourceRelationship::capture(owner)?,
        workbook_relationships: capture_relationships(workbook_relationships)?.into(),
        sheets: sheets.into(),
    })
}

fn capture_source_touched(
    package: &SourceBackedPackage,
    graph: &GraphBinding,
    positions: &[usize],
) -> Result<Vec<TouchedPart>> {
    let positions = checked_positions(positions)?;
    let mut touched = Vec::new();
    touched
        .try_reserve_exact(positions.len())
        .map_err(|source| allocation("tab-state touched source Parts", source))?;
    for position in positions {
        let binding = checked_tabular_binding(graph, position)?;
        let part = package.part(&binding.part.uri)?;
        let bytes = part.data()?.into_arc()?;
        touched.push(TouchedPart {
            position,
            part: PartState::new(
                binding.part.uri.clone(),
                &binding.part.content_type,
                bytes,
                "tab-state touched sheet content type",
            )?,
            relationships: capture_relationships(part.rels())?.into(),
        });
    }
    Ok(touched)
}

fn capture_owned_touched(
    package: &OpcPackage,
    graph: &GraphBinding,
    positions: &[usize],
) -> Result<Vec<TouchedPart>> {
    let positions = checked_positions(positions)?;
    let mut touched = Vec::new();
    touched
        .try_reserve_exact(positions.len())
        .map_err(|source| allocation("tab-state touched owned Parts", source))?;
    for position in positions {
        let binding = checked_tabular_binding(graph, position)?;
        let part = package.get_part(&binding.part.uri)?;
        touched.push(TouchedPart {
            position,
            part: PartState::new(
                binding.part.uri.clone(),
                &binding.part.content_type,
                part.blob_arc(),
                "tab-state touched sheet content type",
            )?,
            relationships: capture_relationships(part.rels())?.into(),
        });
    }
    Ok(touched)
}

fn checked_positions(positions: &[usize]) -> Result<Vec<usize>> {
    if positions.len() > 2 {
        return Err(invalid(
            "tab-state selection closure exceeds two sheet Parts",
        ));
    }
    let mut positions = positions.to_vec();
    positions.sort_unstable();
    positions.dedup();
    Ok(positions)
}

fn checked_tabular_binding(graph: &GraphBinding, position: usize) -> Result<&SheetBinding> {
    let binding = graph
        .sheets
        .get(position)
        .ok_or_else(|| invalid("tab-state selection closure is outside the catalog"))?;
    if binding.kind != WorksheetKind::Worksheet {
        return Err(Error::TabEditBlocked {
            sheet: binding.name.to_string(),
            position,
            reason: TabEditBlock::MarkupCompatibility,
        });
    }
    Ok(binding)
}

fn capture_relationships(relationships: &Relationships) -> Result<Vec<SourceRelationship>> {
    let mut captured = Vec::new();
    captured
        .try_reserve_exact(relationships.len())
        .map_err(|source| allocation("tab-state relationship closure", source))?;
    for relationship in relationships.iter() {
        captured.push(SourceRelationship::capture(relationship)?);
    }
    captured.sort_by(|left, right| left.id.cmp(&right.id));
    Ok(captured)
}

fn relationships_match(expected: &[SourceRelationship], actual: &Relationships) -> bool {
    expected.len() == actual.len()
        && expected.iter().all(|relationship| {
            actual
                .get(&relationship.id)
                .is_some_and(|actual| relationship.matches(actual))
        })
}

fn unique_owner(relationships: &Relationships) -> Option<&Relationship> {
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

fn sheet_kind(relationship_type: &str, content_type: &str) -> Result<WorksheetKind> {
    match relationship_type {
        rt::WORKSHEET | rt::STRICT_WORKSHEET if content_type == ct::SML_WORKSHEET => {
            Ok(WorksheetKind::Worksheet)
        },
        CHARTSHEET_REL | STRICT_CHARTSHEET_REL => Ok(WorksheetKind::Chart),
        DIALOGSHEET_REL => Ok(WorksheetKind::Dialog),
        MACROSHEET_REL | INTL_MACROSHEET_REL => Ok(WorksheetKind::Macro),
        rt::WORKSHEET | rt::STRICT_WORKSHEET => Err(invalid(format!(
            "worksheet relationship has content type '{content_type}'"
        ))),
        _ => Ok(WorksheetKind::Unknown),
    }
}

fn visibility(value: &raw::Visibility) -> Visibility {
    match value {
        raw::Visibility::Visible => Visibility::Visible,
        raw::Visibility::Hidden => Visibility::Hidden,
        raw::Visibility::VeryHidden => Visibility::VeryHidden,
        raw::Visibility::Unknown(value) => Visibility::Unknown(value.clone()),
    }
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
