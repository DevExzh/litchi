//! Source spans and compact physical worksheet index for sheet metadata.

use std::{collections::HashMap, mem::size_of, ops::Range, sync::Arc};

use litchi_core::{Error, ExecutionContext, Reservation, Resource, Result};
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::{Namespace, NamespaceResolver, ResolveResult},
    reader::NsReader,
};

use crate::{
    model::{
        consolidation::{Options, UseLabels},
        detective::Detective,
        label_range::{Orientation, Range as LabelRange},
        source::CellRange,
    },
    worksheet::Merge,
};

pub(crate) const OFFICE_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
pub(crate) const TABLE_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
pub(crate) const XLINK_NS: &str = "http://www.w3.org/1999/xlink";
pub(crate) const TEXT_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
pub(crate) const MC_NS: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const MAX_LOGICAL_ROWS: usize = 1_048_576;
const MAX_LOGICAL_COLUMNS: usize = 1_048_576;
pub(crate) const SELECTOR_COMPARISON_WORK: u64 = 8;

/// Bounded execution ledger for one selector lookup.
///
/// The ledger stays on the caller's stack. Successful comparisons only update
/// the retained context's work counter; no per-comparison object or allocation
/// is created. Callers may share one ledger across the several selector
/// resolutions that make up a staging operation so the metadata work ceiling
/// applies to the complete operation.
pub(crate) struct SelectorBudget {
    context: ExecutionContext,
    maximum: u64,
    used: u64,
}

impl SelectorBudget {
    pub(crate) fn new(context: &ExecutionContext, limits: crate::sheet_metadata::Limits) -> Self {
        Self {
            context: context.clone(),
            maximum: limits.max_work_units(),
            used: 0,
        }
    }

    pub(crate) fn check(&self) -> Result<()> {
        self.context.check().map_err(map_execution)
    }

    /// Charge one selector candidate comparison and check cancellation.
    pub(crate) fn comparison(&mut self) -> Result<()> {
        self.check()?;
        let next = self
            .used
            .checked_add(SELECTOR_COMPARISON_WORK)
            .ok_or_else(|| invalid("ODS metadata selector work units overflow"))?;
        if next > self.maximum {
            return Err(limit_u64(
                Resource::Work,
                "selector work units",
                next,
                self.maximum,
            ));
        }
        self.context
            .consume(Resource::Work, SELECTOR_COMPARISON_WORK)
            .map_err(map_execution)?;
        self.used = next;
        Ok(())
    }
}

#[derive(Clone, Debug)]
pub(crate) struct Span {
    pub(crate) parent: Option<usize>,
    pub(crate) range: Range<usize>,
    pub(crate) close_start: usize,
    pub(crate) namespace: Option<String>,
    pub(crate) local: String,
    pub(crate) qname: String,
    pub(crate) attrs: Vec<Attr>,
    pub(crate) children: Vec<usize>,
    pub(crate) empty: bool,
    pub(crate) non_whitespace_text: bool,
    pub(crate) foreign_ancestor: bool,
    pub(crate) mce_ancestor: bool,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct Attr {
    pub(crate) namespace: Option<String>,
    pub(crate) local: String,
    pub(crate) qname: String,
    pub(crate) value: String,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum CellKind {
    TableCell,
    CoveredTableCell,
}

#[derive(Clone, Debug)]
pub(crate) struct OwnerLocation {
    pub(crate) raw: Range<usize>,
    pub(crate) canonical: bool,
    pub(crate) opaque: bool,
}

#[derive(Clone, Debug)]
pub(crate) struct CellRecord {
    pub(crate) table: usize,
    pub(crate) row: usize,
    pub(crate) span: usize,
    pub(crate) row_start: usize,
    pub(crate) column_start: usize,
    pub(crate) row_repeat: usize,
    pub(crate) column_repeat: usize,
    pub(crate) kind: CellKind,
    pub(crate) merge: Merge,
    pub(crate) source: Option<CellRange>,
    pub(crate) source_owner: Option<OwnerLocation>,
    pub(crate) detective: Option<Detective>,
    pub(crate) detective_owner: Option<OwnerLocation>,
    pub(crate) direct_owner_count: usize,
    pub(crate) sequence_valid: bool,
    pub(crate) unsupported_child: bool,
}

#[derive(Clone, Debug)]
pub(crate) struct RowRecord {
    pub(crate) span: usize,
    pub(crate) logical_start: usize,
    pub(crate) repeat: usize,
    pub(crate) cells: Vec<usize>,
}

#[derive(Clone, Debug)]
pub(crate) struct TableRecord {
    pub(crate) span: usize,
    pub(crate) name: String,
    pub(crate) rows: Vec<usize>,
}

#[derive(Clone, Debug)]
pub(crate) struct LabelRecord {
    pub(crate) span: usize,
    pub(crate) ranges: Vec<LabelRange>,
    pub(crate) canonical: bool,
}

#[derive(Clone, Debug)]
pub(crate) struct ConsolidationRecord {
    pub(crate) span: usize,
    pub(crate) value: Options,
    pub(crate) canonical: bool,
}

#[derive(Clone, Debug)]
pub(crate) struct Catalog {
    pub(crate) spans: Vec<Span>,
    pub(crate) spreadsheet: usize,
    pub(crate) consolidation: Option<ConsolidationRecord>,
    pub(crate) labels: Option<LabelRecord>,
    pub(crate) tables: Vec<TableRecord>,
    pub(crate) rows: Vec<RowRecord>,
    pub(crate) cells: Vec<CellRecord>,
    /// Retained budget charge for detective semantic values copied out of the
    /// borrowed helper snapshot.  The catalog itself is shared by immutable
    /// metadata snapshots, so this guard must live as long as that catalog.
    pub(crate) _detective_memory: Option<Arc<Reservation>>,
    /// Retained budget charge for the borrowed content.xml input view.
    pub(crate) _input_reservation: Arc<Reservation>,
    /// Retained charge for the compact index and its owned semantic values.
    ///
    /// The scanner charges each retained allocation before constructing it and
    /// merges those reservations into this guard.  Keeping the guard with the
    /// catalog prevents a shared snapshot from silently outliving its budget.
    pub(crate) _index_memory: Option<Arc<Reservation>>,
}

impl Catalog {
    pub(crate) fn table_index_for_selector_with_budget<'a>(
        &self,
        selector: &crate::sheet_metadata::CellSelector<'a>,
        budget: &mut SelectorBudget,
    ) -> Result<usize> {
        budget.check()?;
        match selector.sheet() {
            crate::sheet_metadata::SheetSelector::Name(name) => {
                let mut found = None;
                for (index, table) in self.tables.iter().enumerate() {
                    budget.comparison()?;
                    if table.name == name {
                        if found.replace(index).is_some() {
                            return Err(Error::InvalidFormat(format!(
                                "ODS metadata worksheet name '{name}' is ambiguous"
                            )));
                        }
                    }
                }
                found.ok_or_else(|| {
                    Error::InvalidFormat(format!("ODS metadata worksheet '{name}' was not found"))
                })
            },
            crate::sheet_metadata::SheetSelector::Position(position) => {
                let index = position.get();
                budget.comparison()?;
                if index >= self.tables.len() {
                    return Err(Error::InvalidFormat(format!(
                        "ODS metadata worksheet position {index} is outside {} sheets",
                        self.tables.len()
                    )));
                }
                Ok(index)
            },
        }
    }

    pub(crate) fn cell_for_selector<'a>(
        &self,
        selector: &crate::sheet_metadata::CellSelector<'a>,
        context: &ExecutionContext,
        limits: crate::sheet_metadata::Limits,
    ) -> Result<PhysicalCell> {
        let mut budget = SelectorBudget::new(context, limits);
        self.cell_for_selector_with_budget(selector, &mut budget)
    }

    pub(crate) fn cell_for_selector_with_budget<'a>(
        &self,
        selector: &crate::sheet_metadata::CellSelector<'a>,
        budget: &mut SelectorBudget,
    ) -> Result<PhysicalCell> {
        budget.check()?;
        let table_index = self.table_index_for_selector_with_budget(selector, budget)?;
        let table = self.tables.get(table_index).ok_or_else(|| {
            Error::InvalidFormat("ODS metadata worksheet index disappeared".to_string())
        })?;
        for row_index in &table.rows {
            let row = self.rows.get(*row_index).ok_or_else(|| {
                Error::InvalidFormat("ODS metadata row index disappeared".to_string())
            })?;
            let row_end = row.logical_start.checked_add(row.repeat).ok_or_else(|| {
                Error::InvalidFormat("ODS metadata logical row address overflows".to_string())
            })?;
            budget.comparison()?;
            if selector.row() < row_end && selector.row() >= row.logical_start {
                for cell_index in &row.cells {
                    budget.comparison()?;
                    let cell = self.cells.get(*cell_index).ok_or_else(|| {
                        Error::InvalidFormat("ODS metadata cell index disappeared".to_string())
                    })?;
                    let end = cell
                        .column_start
                        .checked_add(cell.column_repeat)
                        .ok_or_else(|| {
                            Error::InvalidFormat(
                                "ODS metadata logical column address overflows".to_string(),
                            )
                        })?;
                    if selector.column() < end && selector.column() >= cell.column_start {
                        let physical_row = selector.row();
                        let physical_column = selector.column();
                        if cell.merge != Merge::None
                            && physical_row == cell.row_start
                            && physical_column == cell.column_start
                        {
                            return Ok(PhysicalCell::Stored(*cell_index));
                        }
                        if cell.merge == Merge::None {
                            return Ok(PhysicalCell::Stored(*cell_index));
                        }
                        if let Merge::Span { rows, columns } = cell.merge {
                            let row_end =
                                cell.row_start.checked_add(rows.get()).ok_or_else(|| {
                                    Error::InvalidFormat("ODS merge row span overflows".to_string())
                                })?;
                            let col_end =
                                cell.column_start
                                    .checked_add(columns.get())
                                    .ok_or_else(|| {
                                        Error::InvalidFormat(
                                            "ODS merge column span overflows".to_string(),
                                        )
                                    })?;
                            if physical_row < row_end && physical_column < col_end {
                                return Ok(PhysicalCell::ImplicitCovered {
                                    anchor: (cell.row_start, cell.column_start),
                                    rows: rows.get(),
                                    columns: columns.get(),
                                });
                            }
                        }
                        return Ok(PhysicalCell::Stored(*cell_index));
                    }
                }
                // A merge anchor may cover later rows that have no physical
                // row/cell object.  Keep this lookup sparse and bounded.
                for cell_index in &self.cells {
                    budget.comparison()?;
                    if cell_index.table != table_index {
                        continue;
                    }
                    if let Merge::Span { rows, columns } = cell_index.merge {
                        let row_end =
                            cell_index
                                .row_start
                                .checked_add(rows.get())
                                .ok_or_else(|| {
                                    Error::InvalidFormat("ODS merge row span overflows".to_string())
                                })?;
                        let col_end = cell_index
                            .column_start
                            .checked_add(columns.get())
                            .ok_or_else(|| {
                                Error::InvalidFormat("ODS merge column span overflows".to_string())
                            })?;
                        if selector.row() >= cell_index.row_start
                            && selector.row() < row_end
                            && selector.column() >= cell_index.column_start
                            && selector.column() < col_end
                        {
                            return Ok(PhysicalCell::ImplicitCovered {
                                anchor: (cell_index.row_start, cell_index.column_start),
                                rows: rows.get(),
                                columns: columns.get(),
                            });
                        }
                    }
                }
                return Ok(PhysicalCell::Missing);
            }
        }
        for cell in &self.cells {
            budget.comparison()?;
            if cell.table != table_index {
                continue;
            }
            if let Merge::Span { rows, columns } = cell.merge {
                let row_end = cell.row_start.checked_add(rows.get()).ok_or_else(|| {
                    Error::InvalidFormat("ODS merge row span overflows".to_string())
                })?;
                let col_end = cell
                    .column_start
                    .checked_add(columns.get())
                    .ok_or_else(|| {
                        Error::InvalidFormat("ODS merge column span overflows".to_string())
                    })?;
                if selector.row() >= cell.row_start
                    && selector.row() < row_end
                    && selector.column() >= cell.column_start
                    && selector.column() < col_end
                {
                    return Ok(PhysicalCell::ImplicitCovered {
                        anchor: (cell.row_start, cell.column_start),
                        rows: rows.get(),
                        columns: columns.get(),
                    });
                }
            }
        }
        Ok(PhysicalCell::Missing)
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum PhysicalCell {
    Missing,
    Stored(usize),
    ImplicitCovered {
        anchor: (usize, usize),
        rows: usize,
        columns: usize,
    },
}

pub(crate) fn scan(
    source: &str,
    limits: crate::sheet_metadata::Limits,
    context: &ExecutionContext,
) -> Result<Catalog> {
    limits.validate()?;
    context.check().map_err(map_execution)?;
    if source.len() > limits.max_input_bytes() {
        return Err(limit(
            Resource::InputBytes,
            "input bytes",
            source.len(),
            limits.max_input_bytes(),
        ));
    }
    let input_reservation = Arc::new(
        context
            .reserve(
                Resource::InputBytes,
                u64::try_from(source.len())
                    .map_err(|_| invalid("ODS metadata input size overflows u64"))?,
            )
            .map_err(map_execution)?,
    );
    let mut format_work =
        u64::try_from(source.len()).map_err(|_| invalid("ODS metadata input size overflows"))?;
    if format_work > limits.max_work_units() {
        return Err(limit_u64(
            Resource::Work,
            "work units",
            format_work,
            limits.max_work_units(),
        ));
    }
    context
        .consume(Resource::Work, format_work)
        .map_err(map_execution)?;
    let mut reader = NsReader::from_str(source);
    reader.config_mut().check_end_names = true;
    reader.config_mut().trim_text(false);
    reader
        .resolver_mut()
        .set_max_declarations_per_element(limits.max_namespace_bindings());
    // Borrow events directly from the immutable input.  This avoids an
    // uncharged scratch-buffer growth path for large attributes and text.
    let mut spans = Vec::<Span>::new();
    let mut stack = Vec::<usize>::new();
    let mut depth_reservations = Vec::<Reservation>::new();
    let mut index_memory = None;
    let mut events = 0usize;
    let mut root = None;
    let mut root_closed = false;
    loop {
        context.check().map_err(map_execution)?;
        let start = usize::try_from(reader.buffer_position())
            .map_err(|_| invalid("ODS metadata XML offset overflows usize"))?;
        charge_work(context, &mut format_work, 16, limits.max_work_units())?;
        let (resolved, event) = reader.read_resolved_event().map_err(xml_error)?;
        let namespace = if matches!(&event, Event::Start(_) | Event::Empty(_)) {
            resolve_namespace(&resolved, context, &mut index_memory)?
        } else {
            None
        };
        drop(resolved);
        events = events
            .checked_add(1)
            .ok_or_else(|| invalid("ODS metadata XML event count overflows"))?;
        if events > limits.max_events() {
            return Err(limit(
                Resource::Objects,
                "XML events",
                events,
                limits.max_events(),
            ));
        }
        match event {
            Event::Start(element) => {
                charge_work(context, &mut format_work, 64, limits.max_work_units())?;
                let next_depth = stack
                    .len()
                    .checked_add(1)
                    .ok_or_else(|| invalid("ODS metadata XML depth overflows"))?;
                if next_depth > limits.max_depth() {
                    return Err(limit(
                        Resource::Depth,
                        "XML depth",
                        next_depth,
                        limits.max_depth(),
                    ));
                }
                let depth_reservation =
                    context.reserve(Resource::Depth, 1).map_err(map_execution)?;
                let parent = stack.last().copied();
                if parent.is_none() && (root.is_some() || root_closed) {
                    return Err(invalid("ODS metadata XML contains multiple roots"));
                }
                let open_end = usize::try_from(reader.buffer_position())
                    .map_err(|_| invalid("ODS metadata XML offset overflows usize"))?;
                let span = make_span(
                    source,
                    &reader,
                    namespace,
                    &element,
                    start,
                    open_end,
                    parent,
                    false,
                    &spans,
                    context,
                    &mut index_memory,
                    limits,
                )?;
                let index = spans.len();
                reserve_vec_slot(
                    &mut spans,
                    context,
                    &mut index_memory,
                    "ODS metadata XML spans",
                )?;
                if let Some(parent) = parent {
                    let parent_span = spans
                        .get_mut(parent)
                        .ok_or_else(|| invalid("ODS metadata parent span is invalid"))?;
                    reserve_vec_slot(
                        &mut parent_span.children,
                        context,
                        &mut index_memory,
                        "ODS metadata XML child index",
                    )?;
                    parent_span.children.push(index);
                } else {
                    root = Some(index);
                }
                spans.push(span);
                context
                    .consume(Resource::Objects, 1)
                    .map_err(map_execution)?;
                reserve_vec_slot(
                    &mut depth_reservations,
                    context,
                    &mut index_memory,
                    "ODS metadata XML depth stack",
                )?;
                depth_reservations.push(depth_reservation);
                reserve_vec_slot(
                    &mut stack,
                    context,
                    &mut index_memory,
                    "ODS metadata XML stack",
                )?;
                stack.push(index);
            },
            Event::Empty(element) => {
                charge_work(context, &mut format_work, 64, limits.max_work_units())?;
                let next_depth = stack
                    .len()
                    .checked_add(1)
                    .ok_or_else(|| invalid("ODS metadata XML depth overflows"))?;
                if next_depth > limits.max_depth() {
                    return Err(limit(
                        Resource::Depth,
                        "XML depth",
                        next_depth,
                        limits.max_depth(),
                    ));
                }
                let _depth_reservation =
                    context.reserve(Resource::Depth, 1).map_err(map_execution)?;
                let parent = stack.last().copied();
                if parent.is_none() && (root.is_some() || root_closed) {
                    return Err(invalid("ODS metadata XML contains multiple roots"));
                }
                let end = usize::try_from(reader.buffer_position())
                    .map_err(|_| invalid("ODS metadata XML offset overflows usize"))?;
                let span = make_span(
                    source,
                    &reader,
                    namespace,
                    &element,
                    start,
                    end,
                    parent,
                    true,
                    &spans,
                    context,
                    &mut index_memory,
                    limits,
                )?;
                let index = spans.len();
                reserve_vec_slot(
                    &mut spans,
                    context,
                    &mut index_memory,
                    "ODS metadata XML spans",
                )?;
                if let Some(parent) = parent {
                    let parent_span = spans
                        .get_mut(parent)
                        .ok_or_else(|| invalid("ODS metadata parent span is invalid"))?;
                    reserve_vec_slot(
                        &mut parent_span.children,
                        context,
                        &mut index_memory,
                        "ODS metadata XML child index",
                    )?;
                    parent_span.children.push(index);
                } else {
                    root = Some(index);
                    root_closed = true;
                }
                spans.push(span);
                context
                    .consume(Resource::Objects, 1)
                    .map_err(map_execution)?;
            },
            Event::End(_) => {
                let index = stack
                    .pop()
                    .ok_or_else(|| invalid("ODS metadata XML stack underflow"))?;
                depth_reservations
                    .pop()
                    .ok_or_else(|| invalid("ODS metadata XML depth reservation underflow"))?;
                let end = usize::try_from(reader.buffer_position())
                    .map_err(|_| invalid("ODS metadata XML offset overflows usize"))?;
                let span = spans
                    .get_mut(index)
                    .ok_or_else(|| invalid("ODS metadata XML span disappeared"))?;
                span.close_start = start;
                span.range.end = end;
                if stack.is_empty() {
                    root_closed = true;
                }
            },
            Event::Eof => break,
            Event::Decl(_) | Event::DocType(_) | Event::PI(_) | Event::Comment(_) => {},
            Event::Text(text) => {
                if text_has_non_whitespace(text.as_ref()) {
                    if stack.is_empty() {
                        return Err(invalid("ODS metadata XML has text outside its root"));
                    }
                    mark_non_whitespace_text(&mut spans, &stack)?;
                }
            },
            Event::CData(text) => {
                if text_has_non_whitespace(text.as_ref()) {
                    if stack.is_empty() {
                        return Err(invalid("ODS metadata XML has text outside its root"));
                    }
                    mark_non_whitespace_text(&mut spans, &stack)?;
                }
            },
            Event::GeneralRef(value) => {
                if general_ref_has_non_whitespace(value.as_ref()) {
                    if stack.is_empty() {
                        return Err(invalid("ODS metadata XML has a reference outside its root"));
                    }
                    mark_non_whitespace_text(&mut spans, &stack)?;
                }
            },
        }
    }
    if !stack.is_empty() {
        return Err(invalid("ODS metadata XML has an unfinished element"));
    }
    if !depth_reservations.is_empty() {
        return Err(invalid("ODS metadata XML depth reservations remain active"));
    }
    let root = root.ok_or_else(|| invalid("ODS metadata XML has no root element"))?;
    if !root_closed {
        return Err(invalid("ODS metadata XML has no complete root"));
    }
    let spreadsheet = admitted_spreadsheet(&spans, root)?;
    build_catalog(
        source,
        spans,
        spreadsheet,
        limits,
        context,
        input_reservation,
        index_memory,
    )
}

fn build_catalog(
    source: &str,
    spans: Vec<Span>,
    spreadsheet: usize,
    limits: crate::sheet_metadata::Limits,
    context: &ExecutionContext,
    input_reservation: Arc<Reservation>,
    index_memory: Option<Reservation>,
) -> Result<Catalog> {
    let mut index_memory = index_memory;
    let direct_children = &spans
        .get(spreadsheet)
        .ok_or_else(|| invalid("ODS spreadsheet span disappeared"))?
        .children;
    let mut tables = Vec::new();
    let mut consolidation = None;
    let mut labels = None;
    for child in direct_children {
        let span = spans
            .get(*child)
            .ok_or_else(|| invalid("ODS spreadsheet child span disappeared"))?;
        if span.namespace.as_deref() != Some(TABLE_NS) {
            continue;
        }
        match span.local.as_str() {
            "table" => {
                if tables.len() >= limits.max_sheets() {
                    return Err(limit(
                        Resource::Objects,
                        "sheets",
                        tables.len().saturating_add(1),
                        limits.max_sheets(),
                    ));
                }
                let name = bounded_text_owned(
                    required_attr(span, TABLE_NS, "name")?,
                    limits.max_text_bytes(),
                    "table:name",
                    context,
                    &mut index_memory,
                )?;
                reserve_memory(context, &mut index_memory, size_of::<TableRecord>())?;
                reserve_vec_slot(
                    &mut tables,
                    context,
                    &mut index_memory,
                    "ODS metadata table index",
                )?;
                tables.push(TableRecord {
                    span: *child,
                    name,
                    rows: Vec::new(),
                });
            },
            "consolidation" => {
                if consolidation.is_some() {
                    return Err(invalid("duplicate direct table:consolidation"));
                }
                consolidation = Some(parse_consolidation_owner(
                    source,
                    &spans,
                    *child,
                    limits.max_text_bytes(),
                    context,
                    &mut index_memory,
                )?);
            },
            "label-ranges" => {
                if labels.is_some() {
                    return Err(invalid("duplicate direct table:label-ranges"));
                }
                labels = Some(parse_label_owner(
                    source,
                    &spans,
                    *child,
                    limits.max_labels(),
                    limits.max_text_bytes(),
                    context,
                    &mut index_memory,
                )?);
            },
            _ => {},
        }
    }
    let mut rows = Vec::new();
    let mut cells = Vec::new();
    for (table_index, table) in tables.iter_mut().enumerate() {
        let table_span = table.span;
        let mut logical_row = 0usize;
        collect_rows(
            source,
            &spans,
            table_span,
            table_index,
            &mut logical_row,
            table,
            &mut rows,
            &mut cells,
            limits,
            context,
            &mut index_memory,
        )?;
    }
    // Let the dedicated detective owner validate the complete cell grammar and
    // retain its existing semantic model; this index only stores source ranges
    // and maps the result to physical cells.
    let mut detective_memory = None;
    if cells.iter().any(|cell| cell.detective_owner.is_some()) {
        let detective_limits = crate::sheet_metadata::detective_limits(limits)?;
        let snapshot = crate::sheet_metadata::detective::Snapshot::parse_with_context(
            source,
            detective_limits,
            context,
        )?;
        let owner_map_bytes = snapshot
            .owners()
            .len()
            .checked_mul(size_of::<((usize, usize), (&Detective, bool))>().saturating_mul(2))
            .ok_or_else(|| invalid("ODS metadata detective owner map memory overflows"))?;
        let _owner_map_memory = context
            .reserve(
                Resource::Memory,
                u64::try_from(owner_map_bytes)
                    .map_err(|_| invalid("ODS metadata detective owner map memory overflows"))?,
            )
            .map_err(map_execution)?;
        let mut by_raw: HashMap<(usize, usize), (&Detective, bool)> = HashMap::new();
        by_raw
            .try_reserve(snapshot.owners().len())
            .map_err(|error| allocation("ODS metadata detective owner map", error))?;
        for owner in snapshot.owners() {
            let Some(typed) = owner.typed() else {
                continue;
            };
            by_raw.insert(
                (typed.raw().range().start, typed.raw().range().end),
                (
                    typed.value(),
                    matches!(
                        typed.lexical_fidelity(),
                        crate::sheet_metadata::detective::LexicalFidelity::Canonical
                    ),
                ),
            );
        }
        let mut detective_memory_bytes = 0u64;
        for cell in &cells {
            let Some(owner) = cell.detective_owner.as_ref() else {
                continue;
            };
            let Some((value, _)) = by_raw.get(&(owner.raw.start, owner.raw.end)) else {
                continue;
            };
            let bound = crate::sheet_metadata::detective_size_bound(value)?
                .checked_mul(2)
                .and_then(|bytes| bytes.checked_add(size_of::<Detective>()))
                .ok_or_else(|| invalid("ODS detective retained value memory overflows"))?;
            detective_memory_bytes = detective_memory_bytes
                .checked_add(
                    u64::try_from(bound)
                        .map_err(|_| invalid("ODS detective retained value memory overflows"))?,
                )
                .ok_or_else(|| invalid("ODS detective retained value memory overflows"))?;
        }
        detective_memory = if detective_memory_bytes == 0 {
            None
        } else {
            Some(Arc::new(
                context
                    .reserve(Resource::Memory, detective_memory_bytes)
                    .map_err(map_execution)?,
            ))
        };
        for cell in &mut cells {
            let Some(owner) = cell.detective_owner.as_mut() else {
                continue;
            };
            if let Some((value, canonical)) = by_raw.get(&(owner.raw.start, owner.raw.end)) {
                cell.detective = Some((*value).clone());
                let empty_self_closing = value.is_empty()
                    && source
                        .get(owner.raw.clone())
                        .is_some_and(|raw| raw.trim_end().ends_with("/>"));
                owner.canonical = *canonical
                    || canonical_detective_with_local_table_namespace(
                        source,
                        owner.raw.clone(),
                        value,
                    )
                    || empty_self_closing;
                owner.opaque = false;
            } else {
                owner.opaque = true;
            }
        }
    }
    Ok(Catalog {
        spans,
        spreadsheet,
        consolidation,
        labels,
        tables,
        rows,
        cells,
        _detective_memory: detective_memory,
        _input_reservation: input_reservation,
        _index_memory: index_memory.map(Arc::new),
    })
}

fn collect_rows(
    source: &str,
    spans: &[Span],
    parent: usize,
    table_index: usize,
    logical_row: &mut usize,
    table: &mut TableRecord,
    rows: &mut Vec<RowRecord>,
    cells: &mut Vec<CellRecord>,
    limits: crate::sheet_metadata::Limits,
    context: &ExecutionContext,
    index_memory: &mut Option<Reservation>,
) -> Result<()> {
    let children = spans
        .get(parent)
        .ok_or_else(|| invalid("ODS metadata row-container span disappeared"))?
        .children
        .as_slice();
    for child in children {
        let child_span = spans
            .get(*child)
            .ok_or_else(|| invalid("ODS metadata row child span disappeared"))?;
        if child_span.foreign_ancestor || child_span.mce_ancestor {
            continue;
        }
        if child_span.namespace.as_deref() == Some(TABLE_NS) && child_span.local == "table-row" {
            append_row(
                source,
                spans,
                *child,
                table_index,
                logical_row,
                table,
                rows,
                cells,
                limits,
                context,
                index_memory,
            )?;
        } else if is_row_container(child_span) {
            collect_rows(
                source,
                spans,
                *child,
                table_index,
                logical_row,
                table,
                rows,
                cells,
                limits,
                context,
                index_memory,
            )?;
        }
    }
    Ok(())
}

fn append_row(
    source: &str,
    spans: &[Span],
    row_span_index: usize,
    table_index: usize,
    logical_row: &mut usize,
    table: &mut TableRecord,
    rows: &mut Vec<RowRecord>,
    cells: &mut Vec<CellRecord>,
    limits: crate::sheet_metadata::Limits,
    context: &ExecutionContext,
    index_memory: &mut Option<Reservation>,
) -> Result<()> {
    let row_span = spans
        .get(row_span_index)
        .ok_or_else(|| invalid("ODS metadata row span disappeared"))?;
    let repeat = parse_positive(
        attr(row_span, TABLE_NS, "number-rows-repeated"),
        "table:number-rows-repeated",
    )?;
    if repeat > MAX_LOGICAL_ROWS {
        return Err(limit(
            Resource::Objects,
            "logical rows",
            repeat,
            MAX_LOGICAL_ROWS,
        ));
    }
    let row_index = rows.len();
    reserve_memory(context, index_memory, size_of::<RowRecord>())?;
    reserve_vec_slot(rows, context, index_memory, "ODS metadata row index")?;
    rows.push(RowRecord {
        span: row_span_index,
        logical_start: *logical_row,
        repeat,
        cells: Vec::new(),
    });
    reserve_vec_slot(
        &mut table.rows,
        context,
        index_memory,
        "ODS metadata table row index",
    )?;
    table.rows.push(row_index);
    let mut logical_column = 0usize;
    for cell_span_index in &row_span.children {
        let cell_span = spans
            .get(*cell_span_index)
            .ok_or_else(|| invalid("ODS metadata cell span disappeared"))?;
        let Some(kind) = cell_kind(cell_span) else {
            continue;
        };
        if cell_span.foreign_ancestor || cell_span.mce_ancestor {
            continue;
        }
        let column_repeat = parse_positive(
            attr(cell_span, TABLE_NS, "number-columns-repeated"),
            "table:number-columns-repeated",
        )?;
        if column_repeat > MAX_LOGICAL_COLUMNS {
            return Err(limit(
                Resource::Objects,
                "logical columns",
                column_repeat,
                MAX_LOGICAL_COLUMNS,
            ));
        }
        let merge = parse_merge(cell_span)?;
        if let Merge::Span { rows, columns } = merge {
            if rows.get() > MAX_LOGICAL_ROWS || columns.get() > MAX_LOGICAL_COLUMNS {
                return Err(invalid(
                    "ODS merge geometry exceeds logical worksheet limits",
                ));
            }
        }
        // Reserve the catalog record and a conservative source-sized bound
        // before parsing any typed child that may allocate semantic strings.
        let cell_bound = size_of::<CellRecord>()
            .checked_add(cell_span.range.len())
            .ok_or_else(|| invalid("ODS metadata cell memory overflows"))?;
        reserve_memory(context, index_memory, cell_bound)?;
        let mut source_owner = None;
        let mut detective_owner = None;
        let mut source_value = None;
        let detective_value = None;
        let mut source_count = 0usize;
        let mut detective_count = 0usize;
        let mut sequence_valid = !cell_span.non_whitespace_text;
        let mut unsupported_child = cell_span.non_whitespace_text;
        let mut order = 0u8;
        for child in &cell_span.children {
            let child_span = spans
                .get(*child)
                .ok_or_else(|| invalid("ODS metadata cell child span disappeared"))?;
            let direct = !child_span.foreign_ancestor && !child_span.mce_ancestor;
            if direct
                && child_span.namespace.as_deref() == Some(TABLE_NS)
                && child_span.local == "cell-range-source"
            {
                source_count += 1;
                if order > 0 {
                    sequence_valid = false;
                }
                order = order.max(1);
                if has_nested_metadata(spans, *child)? {
                    unsupported_child = true;
                }
                if source_count == 1 {
                    let parsed = parse_cell_range_source(
                        source,
                        child_span,
                        limits.max_text_bytes(),
                        context,
                        index_memory,
                    )?;
                    let canonical = canonical_cell_range_source(source, child_span, &parsed);
                    source_value = Some(parsed);
                    source_owner = Some(OwnerLocation {
                        raw: child_span.range.clone(),
                        canonical,
                        opaque: false,
                    });
                }
            } else if direct
                && child_span.namespace.as_deref() == Some(OFFICE_NS)
                && child_span.local == "annotation"
            {
                if order > 1 {
                    sequence_valid = false;
                }
                order = order.max(2);
                if has_nested_metadata(spans, *child)? {
                    unsupported_child = true;
                }
            } else if direct
                && child_span.namespace.as_deref() == Some(TABLE_NS)
                && child_span.local == "detective"
            {
                detective_count += 1;
                if order > 2 {
                    sequence_valid = false;
                }
                order = order.max(3);
                if has_nested_metadata(spans, *child)? {
                    unsupported_child = true;
                }
                if detective_count == 1 {
                    detective_owner = Some(OwnerLocation {
                        raw: child_span.range.clone(),
                        canonical: false,
                        opaque: true,
                    });
                }
            } else if direct
                && child_span.namespace.as_deref() == Some(TEXT_NS)
                && is_legal_text_content_local(child_span.local.as_str())
            {
                order = order.max(4);
                if has_nested_metadata(spans, *child)? {
                    // Same-named metadata beneath text content is opaque.  A
                    // later direct edit must not create an ambiguous owner.
                    unsupported_child = true;
                }
            } else {
                // This includes unknown text:* elements, foreign wrappers,
                // and MCE branches.  They remain source data but make a
                // focused direct-child splice unsafe.
                unsupported_child = true;
            }
        }
        let cell_index = cells.len();
        if cell_index >= limits.max_cells() {
            return Err(limit(
                Resource::Objects,
                "physical cells",
                cell_index.saturating_add(1),
                limits.max_cells(),
            ));
        }
        reserve_vec_slot(cells, context, index_memory, "ODS metadata cell index")?;
        reserve_vec_slot(
            &mut rows[row_index].cells,
            context,
            index_memory,
            "ODS metadata row cell index",
        )?;
        cells.push(CellRecord {
            table: table_index,
            row: row_index,
            span: *cell_span_index,
            row_start: *logical_row,
            column_start: logical_column,
            row_repeat: repeat,
            column_repeat,
            kind,
            merge,
            source: source_value,
            source_owner,
            detective: detective_value,
            detective_owner,
            direct_owner_count: source_count
                .saturating_sub(1)
                .saturating_add(detective_count.saturating_sub(1)),
            sequence_valid,
            unsupported_child,
        });
        rows[row_index].cells.push(cell_index);
        logical_column = logical_column
            .checked_add(column_repeat)
            .ok_or_else(|| invalid("ODS metadata logical column count overflows"))?;
        if logical_column > MAX_LOGICAL_COLUMNS {
            return Err(limit(
                Resource::Objects,
                "logical columns",
                logical_column,
                MAX_LOGICAL_COLUMNS,
            ));
        }
    }
    *logical_row = logical_row
        .checked_add(repeat)
        .ok_or_else(|| invalid("ODS metadata logical row count overflows"))?;
    if *logical_row > MAX_LOGICAL_ROWS {
        return Err(limit(
            Resource::Objects,
            "logical rows",
            *logical_row,
            MAX_LOGICAL_ROWS,
        ));
    }
    Ok(())
}

fn make_span(
    source: &str,
    reader: &NsReader<&[u8]>,
    namespace: Option<String>,
    element: &BytesStart<'_>,
    start: usize,
    open_end: usize,
    parent: Option<usize>,
    empty: bool,
    spans: &[Span],
    context: &ExecutionContext,
    index_memory: &mut Option<Reservation>,
    limits: crate::sheet_metadata::Limits,
) -> Result<Span> {
    let element_bytes = element
        .local_name()
        .as_ref()
        .len()
        .checked_add(element.name().as_ref().len())
        .and_then(|bytes| bytes.checked_add(size_of::<Span>()))
        .ok_or_else(|| invalid("ODS metadata span memory overflows"))?;
    reserve_memory(context, index_memory, element_bytes)?;
    let local = decode(element.local_name().as_ref(), "element local name")?;
    let qname = decode(element.name().as_ref(), "element qualified name")?;
    let attrs = collect_attrs(
        reader.resolver(),
        reader.decoder(),
        element,
        context,
        index_memory,
        limits,
    )?;
    let parent_foreign = parent
        .and_then(|index| spans.get(index))
        .is_some_and(|span| {
            span.foreign_ancestor
                || span
                    .namespace
                    .as_deref()
                    .is_none_or(|value| value != OFFICE_NS && value != TABLE_NS && value != TEXT_NS)
        });
    let foreign_ancestor = parent_foreign
        || namespace
            .as_deref()
            .is_none_or(|value| value != OFFICE_NS && value != TABLE_NS && value != TEXT_NS);
    let mce_ancestor = parent
        .and_then(|index| spans.get(index))
        .is_some_and(|span| span.mce_ancestor)
        || namespace.as_deref() == Some(MC_NS);
    let range_end = open_end;
    let _ = source;
    Ok(Span {
        parent,
        range: start..range_end,
        close_start: open_end,
        namespace,
        local,
        qname,
        attrs,
        children: Vec::new(),
        empty,
        non_whitespace_text: false,
        foreign_ancestor,
        mce_ancestor,
    })
}

fn mark_non_whitespace_text(spans: &mut [Span], stack: &[usize]) -> Result<()> {
    if let Some(index) = stack.last().copied() {
        spans
            .get_mut(index)
            .ok_or_else(|| invalid("ODS metadata text parent span is invalid"))?
            .non_whitespace_text = true;
    }
    Ok(())
}

fn text_has_non_whitespace(value: &[u8]) -> bool {
    value.iter().any(|byte| !byte.is_ascii_whitespace())
}

fn general_ref_has_non_whitespace(value: &[u8]) -> bool {
    let value = value.strip_prefix(b"#");
    let Some(value) = value else {
        return true;
    };
    let parsed = if let Some(hex) = value.strip_prefix(b"x") {
        u32::from_str_radix(std::str::from_utf8(hex).unwrap_or_default(), 16).ok()
    } else {
        std::str::from_utf8(value)
            .ok()
            .and_then(|value| value.parse().ok())
    };
    !matches!(parsed, Some(0x9 | 0xA | 0xD | 0x20))
}

fn admitted_spreadsheet(spans: &[Span], root: usize) -> Result<usize> {
    let root_span = spans
        .get(root)
        .ok_or_else(|| invalid("ODS metadata root span disappeared"))?;
    if root_span.parent.is_some()
        || root_span.namespace.as_deref() != Some(OFFICE_NS)
        || root_span.local != "document-content"
        || root_span.foreign_ancestor
        || root_span.mce_ancestor
    {
        return Err(invalid(
            "ODS content.xml root must be office:document-content",
        ));
    }
    // `office:document-content` has four known direct office children in this
    // grammar, in this order.  Foreign direct children are retained under an
    // explicit opaque-extension policy; office children outside this allowlist
    // cannot be treated as harmless siblings because they may change the
    // document family selected by this owner.
    let body = exactly_one_root_child(spans, root, OFFICE_NS, "body")?;
    let spreadsheet = exactly_one_child(spans, body, OFFICE_NS, "spreadsheet")?;
    for index in [root, body, spreadsheet] {
        if spans
            .get(index)
            .ok_or_else(|| invalid("ODS metadata structural span disappeared"))?
            .non_whitespace_text
        {
            return Err(invalid(
                "ODS metadata structural element contains non-whitespace text",
            ));
        }
    }
    let span = spans
        .get(spreadsheet)
        .ok_or_else(|| invalid("ODS spreadsheet span disappeared"))?;
    if span.foreign_ancestor || span.mce_ancestor {
        return Err(invalid(
            "ODS office:spreadsheet has an unsupported ancestor",
        ));
    }
    Ok(spreadsheet)
}

fn exactly_one_root_child(
    spans: &[Span],
    parent: usize,
    namespace: &str,
    local: &str,
) -> Result<usize> {
    let parent_span = spans
        .get(parent)
        .ok_or_else(|| invalid("ODS metadata parent span disappeared"))?;
    let mut found = None;
    let mut previous_rank = None;
    let mut seen = [false; 4];
    for child in &parent_span.children {
        let child_span = spans
            .get(*child)
            .ok_or_else(|| invalid("ODS metadata child span disappeared"))?;
        if child_span.foreign_ancestor || child_span.mce_ancestor {
            continue;
        }
        if child_span.namespace.as_deref() != Some(namespace) {
            continue;
        }
        let Some(rank) = root_office_child_rank(child_span.local.as_str()) else {
            return Err(invalid(format!(
                "ODS metadata document-content contains unexpected office:{} child",
                child_span.local
            )));
        };
        if previous_rank.is_some_and(|previous| rank < previous) {
            return Err(invalid(
                "ODS metadata document-content office children are out of order",
            ));
        }
        if seen[rank as usize] {
            return Err(invalid(format!(
                "duplicate direct office:{} element",
                child_span.local
            )));
        }
        seen[rank as usize] = true;
        previous_rank = Some(rank);
        if child_span.local == local {
            if found.replace(*child).is_some() {
                return Err(invalid(format!("duplicate direct office:{local} element")));
            }
        }
    }
    found.ok_or_else(|| invalid(format!("ODS metadata has no direct office:{local} element")))
}

fn root_office_child_rank(local: &str) -> Option<u8> {
    match local {
        "scripts" => Some(0),
        "font-face-decls" => Some(1),
        "automatic-styles" => Some(2),
        "body" => Some(3),
        _ => None,
    }
}

fn exactly_one_child(spans: &[Span], parent: usize, namespace: &str, local: &str) -> Result<usize> {
    let parent_span = spans
        .get(parent)
        .ok_or_else(|| invalid("ODS metadata parent span disappeared"))?;
    let mut found = None;
    for child in &parent_span.children {
        let child_span = spans
            .get(*child)
            .ok_or_else(|| invalid("ODS metadata child span disappeared"))?;
        if child_span.foreign_ancestor || child_span.mce_ancestor {
            continue;
        }
        if child_span.namespace.as_deref() == Some(namespace) && child_span.local == local {
            if found.replace(*child).is_some() {
                return Err(invalid(format!("duplicate direct office:{local} element")));
            }
        } else if child_span.namespace.as_deref() == Some(OFFICE_NS) {
            return Err(invalid(format!(
                "ODS metadata office:{local} parent contains an unexpected office child"
            )));
        }
    }
    found.ok_or_else(|| invalid(format!("ODS metadata has no direct office:{local} element")))
}

fn is_row_container(span: &Span) -> bool {
    !span.foreign_ancestor
        && !span.mce_ancestor
        && span.namespace.as_deref() == Some(TABLE_NS)
        && matches!(
            span.local.as_str(),
            "table-row-group" | "table-header-rows" | "table-rows"
        )
}

fn is_legal_text_content_local(local: &str) -> bool {
    matches!(
        local,
        "p" | "h"
            | "span"
            | "s"
            | "tab"
            | "line-break"
            | "a"
            | "soft-page-break"
            | "bookmark"
            | "bookmark-start"
            | "bookmark-end"
            | "reference-mark"
            | "reference-mark-start"
            | "reference-mark-end"
            | "ruby"
            | "ruby-base"
            | "ruby-text"
            | "section"
            | "list"
            | "list-item"
            | "numbered-paragraph"
            | "change"
            | "change-start"
            | "change-end"
            | "page-number"
            | "page-count"
            | "page-variable"
            | "sheet-name"
            | "date"
            | "time"
            | "title"
            | "creator"
            | "author-name"
    )
}

fn is_metadata_local(span: &Span) -> bool {
    (span.namespace.as_deref() == Some(TABLE_NS)
        && matches!(span.local.as_str(), "cell-range-source" | "detective"))
        || (span.namespace.as_deref() == Some(OFFICE_NS) && span.local == "annotation")
}

fn has_nested_metadata(spans: &[Span], index: usize) -> Result<bool> {
    let span = spans
        .get(index)
        .ok_or_else(|| invalid("ODS metadata descendant span disappeared"))?;
    for child in &span.children {
        if is_metadata_local(
            spans
                .get(*child)
                .ok_or_else(|| invalid("ODS metadata descendant span disappeared"))?,
        ) || has_nested_metadata(spans, *child)?
        {
            return Ok(true);
        }
    }
    Ok(false)
}

fn collect_attrs(
    resolver: &NamespaceResolver,
    decoder: quick_xml::encoding::Decoder,
    element: &BytesStart<'_>,
    context: &ExecutionContext,
    index_memory: &mut Option<Reservation>,
    limits: crate::sheet_metadata::Limits,
) -> Result<Vec<Attr>> {
    let mut attrs = Vec::new();
    for raw in element.attributes() {
        let raw =
            raw.map_err(|error| invalid(format!("invalid ODS metadata attribute: {error}")))?;
        if attrs.len() >= limits.max_namespace_bindings() {
            return Err(limit(
                Resource::Objects,
                "attributes",
                attrs.len().saturating_add(1),
                limits.max_namespace_bindings(),
            ));
        }
        let attr_bytes = raw
            .key
            .as_ref()
            .len()
            .checked_add(raw.value.as_ref().len())
            .and_then(|bytes| bytes.checked_add(size_of::<Attr>()))
            .ok_or_else(|| invalid("ODS metadata attribute memory overflows"))?;
        reserve_memory(context, index_memory, attr_bytes)?;
        let (namespace, local) = resolver.resolve_attribute(raw.key);
        let namespace = match namespace {
            ResolveResult::Bound(Namespace(uri)) => {
                reserve_memory(context, index_memory, uri.len())?;
                Some(decode(uri, "attribute namespace")?)
            },
            ResolveResult::Unbound => None,
            ResolveResult::Unknown(prefix) => {
                return Err(invalid(format!(
                    "ODS metadata attribute uses unbound prefix '{}'",
                    String::from_utf8_lossy(prefix.as_ref())
                )));
            },
        };
        let value = raw
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(|error| invalid(format!("invalid ODS metadata attribute value: {error}")))?
            .into_owned();
        reserve_vec_slot(
            &mut attrs,
            context,
            index_memory,
            "ODS metadata attribute index",
        )?;
        attrs.push(Attr {
            namespace,
            local: decode(local.as_ref(), "attribute local name")?,
            qname: decode(raw.key.as_ref(), "attribute name")?,
            value,
        });
    }
    for left in 0..attrs.len() {
        for right in (left + 1)..attrs.len() {
            if attrs[left].namespace == attrs[right].namespace
                && attrs[left].local == attrs[right].local
            {
                return Err(invalid(format!(
                    "duplicate expanded ODS metadata attribute '{}'",
                    attrs[left].local
                )));
            }
        }
    }
    Ok(attrs)
}

fn parse_consolidation_owner(
    source: &str,
    spans: &[Span],
    index: usize,
    max_text_bytes: usize,
    context: &ExecutionContext,
    index_memory: &mut Option<Reservation>,
) -> Result<ConsolidationRecord> {
    let span = spans
        .get(index)
        .ok_or_else(|| invalid("ODS consolidation span disappeared"))?;
    if !span.children.is_empty() || span.non_whitespace_text {
        return Err(invalid("table:consolidation must be empty"));
    }
    let function = bounded_text_owned(
        required_attr(span, TABLE_NS, "function")?,
        max_text_bytes,
        "table:function",
        context,
        index_memory,
    )?;
    let source_address_text = bounded_text_owned(
        required_attr(span, TABLE_NS, "source-cell-range-addresses")?,
        max_text_bytes,
        "table:source-cell-range-addresses",
        context,
        index_memory,
    )?;
    let source_addresses = split_addresses(&source_address_text, context, index_memory)?;
    for address in &source_addresses {
        if address.len() > max_text_bytes {
            return Err(limit(
                Resource::Memory,
                "table:source-cell-range-addresses",
                address.len(),
                max_text_bytes,
            ));
        }
    }
    let target = bounded_text_owned(
        required_attr(span, TABLE_NS, "target-cell-address")?,
        max_text_bytes,
        "table:target-cell-address",
        context,
        index_memory,
    )?;
    let use_labels = attr(span, TABLE_NS, "use-labels")
        .map(|value| match value {
            "none" => Ok(UseLabels::None),
            "row" => Ok(UseLabels::Row),
            "column" => Ok(UseLabels::Column),
            "both" => Ok(UseLabels::Both),
            _ => Err(invalid(format!("invalid table:use-labels value '{value}'"))),
        })
        .transpose()?;
    let link = attr(span, TABLE_NS, "link-to-source-data")
        .map(|value| parse_bool(value, "table:link-to-source-data"))
        .transpose()?;
    let value = Options {
        function,
        source_cell_range_addresses: source_addresses,
        target_cell_address: target,
        use_labels,
        link_to_source_data: link,
    };
    value.validate()?;
    reserve_memory(
        context,
        index_memory,
        size_of::<ConsolidationRecord>()
            .checked_add(size_of::<Options>())
            .ok_or_else(|| invalid("ODS consolidation record memory overflows"))?,
    )?;
    let canonical = canonical_consolidation(source, span, &value);
    Ok(ConsolidationRecord {
        span: index,
        value,
        canonical,
    })
}

fn parse_label_owner(
    source: &str,
    spans: &[Span],
    index: usize,
    max_labels: usize,
    max_text_bytes: usize,
    context: &ExecutionContext,
    index_memory: &mut Option<Reservation>,
) -> Result<LabelRecord> {
    let span = spans
        .get(index)
        .ok_or_else(|| invalid("ODS label-ranges span disappeared"))?;
    if span.non_whitespace_text {
        return Err(invalid("table:label-ranges contains non-whitespace text"));
    }
    let mut ranges = Vec::new();
    for child in &span.children {
        if ranges.len() >= max_labels {
            return Err(limit(
                Resource::Objects,
                "label ranges",
                ranges.len().saturating_add(1),
                max_labels,
            ));
        }
        let child_span = spans
            .get(*child)
            .ok_or_else(|| invalid("ODS label-range child span disappeared"))?;
        if child_span.namespace.as_deref() != Some(TABLE_NS) || child_span.local != "label-range" {
            return Err(invalid("table:label-ranges contains unsupported child"));
        }
        // Relax NG's empty pattern admits both a self-closing element and an
        // explicit start/end pair. Comments and processing instructions are
        // retained in the source spans but do not make the owner non-empty.
        if !child_span.children.is_empty() || child_span.non_whitespace_text {
            return Err(invalid("table:label-range must be empty"));
        }
        let orientation = match required_attr(child_span, TABLE_NS, "orientation")? {
            "row" => Orientation::Row,
            "column" => Orientation::Column,
            value => {
                return Err(invalid(format!(
                    "invalid table:orientation value '{value}'"
                )));
            },
        };
        let range = LabelRange::new(
            bounded_text_owned(
                required_attr(child_span, TABLE_NS, "label-cell-range-address")?,
                max_text_bytes,
                "table:label-cell-range-address",
                context,
                index_memory,
            )?,
            bounded_text_owned(
                required_attr(child_span, TABLE_NS, "data-cell-range-address")?,
                max_text_bytes,
                "table:data-cell-range-address",
                context,
                index_memory,
            )?,
            orientation,
        )?;
        reserve_memory(context, index_memory, size_of::<LabelRange>())?;
        reserve_vec_slot(&mut ranges, context, index_memory, "ODS label-range index")?;
        ranges.push(range);
    }
    reserve_memory(context, index_memory, size_of::<LabelRecord>())?;
    let canonical = canonical_labels(source, span, &ranges);
    Ok(LabelRecord {
        span: index,
        ranges,
        canonical,
    })
}

fn parse_cell_range_source(
    source: &str,
    span: &Span,
    max_text_bytes: usize,
    context: &ExecutionContext,
    index_memory: &mut Option<Reservation>,
) -> Result<CellRange> {
    if !span.children.is_empty() || span.non_whitespace_text {
        return Err(invalid("table:cell-range-source must be empty"));
    }
    if attr(span, XLINK_NS, "type") != Some("simple") {
        return Err(invalid(
            "table:cell-range-source requires xlink:type=simple",
        ));
    }
    let rows = parse_positive(
        Some(required_attr(span, TABLE_NS, "last-row-spanned")?),
        "table:last-row-spanned",
    )?;
    let columns = parse_positive(
        Some(required_attr(span, TABLE_NS, "last-column-spanned")?),
        "table:last-column-spanned",
    )?;
    let mut value = CellRange::new(
        bounded_text_owned(
            required_attr(span, TABLE_NS, "name")?,
            max_text_bytes,
            "table:name",
            context,
            index_memory,
        )?,
        bounded_text_owned(
            required_attr(span, XLINK_NS, "href")?,
            max_text_bytes,
            "xlink:href",
            context,
            index_memory,
        )?,
        rows,
        columns,
    )?;
    value.set_actuate_on_request(attr(span, XLINK_NS, "actuate") == Some("onRequest"));
    value.set_filter_name(
        attr(span, TABLE_NS, "filter-name")
            .map(|value| {
                bounded_text_owned(
                    value,
                    max_text_bytes,
                    "table:filter-name",
                    context,
                    index_memory,
                )
            })
            .transpose()?,
    );
    value.set_filter_options(
        attr(span, TABLE_NS, "filter-options")
            .map(|value| {
                bounded_text_owned(
                    value,
                    max_text_bytes,
                    "table:filter-options",
                    context,
                    index_memory,
                )
            })
            .transpose()?,
    );
    value.set_refresh_delay(
        attr(span, TABLE_NS, "refresh-delay")
            .map(|value| {
                bounded_text_owned(
                    value,
                    max_text_bytes,
                    "table:refresh-delay",
                    context,
                    index_memory,
                )
            })
            .transpose()?,
    )?;
    let _ = source;
    Ok(value)
}

fn cell_kind(span: &Span) -> Option<CellKind> {
    if span.namespace.as_deref() != Some(TABLE_NS) {
        return None;
    }
    match span.local.as_str() {
        "table-cell" => Some(CellKind::TableCell),
        "covered-table-cell" => Some(CellKind::CoveredTableCell),
        _ => None,
    }
}

fn parse_merge(span: &Span) -> Result<Merge> {
    if span.namespace.as_deref() != Some(TABLE_NS) {
        return Ok(Merge::None);
    }
    if span.local == "covered-table-cell" {
        return Ok(Merge::Covered);
    }
    let rows = parse_positive(
        attr(span, TABLE_NS, "number-rows-spanned"),
        "table:number-rows-spanned",
    )?;
    let columns = parse_positive(
        attr(span, TABLE_NS, "number-columns-spanned"),
        "table:number-columns-spanned",
    )?;
    match (rows, columns) {
        (1, 1) => Ok(Merge::None),
        (rows, columns) => Ok(Merge::Span {
            rows: std::num::NonZeroUsize::new(rows)
                .ok_or_else(|| invalid("ODS merge row span must be positive"))?,
            columns: std::num::NonZeroUsize::new(columns)
                .ok_or_else(|| invalid("ODS merge column span must be positive"))?,
        }),
    }
}

fn canonical_consolidation(source: &str, span: &Span, value: &Options) -> bool {
    let mut rendered = String::new();
    if crate::model::consolidation::write_consolidation(&mut rendered, Some(value)).is_err() {
        return false;
    }
    let element_prefix = qname_prefix_or_empty(&span.qname);
    let attribute_prefix = table_attribute_prefix(span).unwrap_or(element_prefix);
    rendered = rename_table_qnames(&rendered, element_prefix, attribute_prefix);
    source
        .get(span.range.clone())
        .is_some_and(|raw| matches_owner_rendering(raw, &rendered, &[("table", TABLE_NS)]))
}

fn canonical_labels(source: &str, span: &Span, ranges: &[LabelRange]) -> bool {
    if ranges.is_empty() {
        let prefix = qname_prefix_or_empty(&span.qname);
        let raw = source.get(span.range.clone());
        let empty_name = if prefix.is_empty() {
            "label-ranges".to_string()
        } else {
            format!("{prefix}:label-ranges")
        };
        let rendered = format!("<{empty_name}/>");
        return raw.is_some_and(|value| {
            if value.contains("<!--") || value.contains("<?") {
                return false;
            }
            value == format!("<{empty_name}></{empty_name}>")
                || matches_owner_rendering(value, &rendered, &[("table", TABLE_NS)])
        });
    }
    let mut rendered = String::new();
    if crate::model::label_range::write(&mut rendered, ranges).is_err() {
        return false;
    }
    let prefix = qname_prefix_or_empty(&span.qname);
    rendered = rename_table_qnames(&rendered, prefix, prefix);
    source
        .get(span.range.clone())
        .is_some_and(|raw| matches_owner_rendering(raw, &rendered, &[("table", TABLE_NS)]))
}

fn canonical_cell_range_source(source: &str, span: &Span, value: &CellRange) -> bool {
    let mut rendered = String::new();
    crate::model::source::write_cell_range_source(&mut rendered, value);
    let element_prefix = qname_prefix_or_empty(&span.qname);
    let attribute_prefix = table_attribute_prefix(span).unwrap_or(element_prefix);
    rendered = rename_table_qnames(&rendered, element_prefix, attribute_prefix);
    source.get(span.range.clone()).is_some_and(|raw| {
        matches_owner_rendering(raw, &rendered, &[("xlink", XLINK_NS), ("table", TABLE_NS)])
    })
}

/// Accept exactly the renderer's known local namespace declarations while
/// keeping every other lexical difference visible to the preservation gate.
/// The generated metadata writers append declarations immediately before the
/// opening tag's close marker, so this comparison does not normalize arbitrary
/// attributes, comments, processing instructions, or whitespace.
fn matches_owner_rendering(raw: &str, rendered: &str, declarations: &[(&str, &str)]) -> bool {
    raw == rendered
        || append_namespace_declarations(rendered, declarations)
            .is_some_and(|candidate| raw == candidate)
}

fn append_namespace_declarations(raw: &str, declarations: &[(&str, &str)]) -> Option<String> {
    if declarations.is_empty() {
        return Some(raw.to_owned());
    }
    let end = opening_end(raw)?;
    let opening = raw.get(..end)?;
    let (close, suffix) = if let Some(close) = opening.strip_suffix("/>") {
        (close, "/>")
    } else if let Some(close) = opening.strip_suffix('>') {
        (close, ">")
    } else {
        return None;
    };
    let extra = declarations
        .iter()
        .try_fold(0usize, |size, (prefix, uri)| {
            size.checked_add(10 + prefix.len() + uri.len())
        })?;
    let mut output = String::with_capacity(raw.len().checked_add(extra)?);
    output.push_str(close);
    for (prefix, uri) in declarations {
        output.push_str(" xmlns:");
        output.push_str(prefix);
        output.push_str("=\"");
        output.push_str(uri);
        output.push('"');
    }
    output.push_str(suffix);
    output.push_str(raw.get(end..)?);
    Some(output)
}

fn opening_end(raw: &str) -> Option<usize> {
    let mut quote = None;
    for (index, byte) in raw.as_bytes().iter().copied().enumerate() {
        match (quote, byte) {
            (None, b'\'' | b'"') => quote = Some(byte),
            (Some(value), byte) if value == byte => quote = None,
            (None, b'>') => return Some(index + 1),
            _ => {},
        }
    }
    None
}

fn canonical_detective_with_local_table_namespace(
    source: &str,
    raw: Range<usize>,
    value: &Detective,
) -> bool {
    let Some(raw) = source.get(raw) else {
        return false;
    };
    let mut rendered = String::new();
    crate::model::detective::write_detective(&mut rendered, value);
    matches_owner_rendering(raw, &rendered, &[("table", TABLE_NS)])
}

fn required_attr<'a>(span: &'a Span, namespace: &str, local: &str) -> Result<&'a str> {
    attr(span, namespace, local).ok_or_else(|| invalid(format!("missing {namespace}:{local}")))
}

pub(crate) fn attr<'a>(span: &'a Span, namespace: &str, local: &str) -> Option<&'a str> {
    span.attrs
        .iter()
        .find(|attr| attr.namespace.as_deref() == Some(namespace) && attr.local == local)
        .map(|attr| attr.value.as_str())
}

fn parse_positive(value: Option<&str>, name: &str) -> Result<usize> {
    let value = value.unwrap_or("1");
    let parsed = value
        .parse::<usize>()
        .map_err(|_| invalid(format!("invalid {name} value '{value}'")))?;
    if parsed == 0 {
        return Err(invalid(format!("{name} must be positive")));
    }
    Ok(parsed)
}

fn bounded_text_owned(
    value: &str,
    maximum: usize,
    name: &'static str,
    context: &ExecutionContext,
    index_memory: &mut Option<Reservation>,
) -> Result<String> {
    if value.len() > maximum {
        return Err(limit(Resource::Memory, name, value.len(), maximum));
    }
    reserve_memory(
        context,
        index_memory,
        value
            .len()
            .checked_add(size_of::<String>())
            .ok_or_else(|| invalid("ODS metadata text memory overflows"))?,
    )?;
    Ok(value.to_owned())
}

fn parse_bool(value: &str, name: &str) -> Result<bool> {
    match value {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => Err(invalid(format!("invalid {name} boolean '{value}'"))),
    }
}

fn split_addresses(
    value: &str,
    context: &ExecutionContext,
    index_memory: &mut Option<Reservation>,
) -> Result<Vec<String>> {
    let mut values = Vec::new();
    let mut current = String::new();
    let mut quoted = false;
    let mut chars = value.chars().peekable();
    while let Some(ch) = chars.next() {
        if ch == '\'' {
            push_index_char(&mut current, ch, context, index_memory)?;
            if quoted && chars.peek() == Some(&'\'') {
                push_index_char(
                    &mut current,
                    chars.next().unwrap_or('\''),
                    context,
                    index_memory,
                )?;
            } else {
                quoted = !quoted;
            }
        } else if ch.is_whitespace() && !quoted {
            if !current.is_empty() {
                let item = std::mem::take(&mut current);
                reserve_memory(
                    context,
                    index_memory,
                    item.len().saturating_add(size_of::<String>()),
                )?;
                reserve_vec_slot(
                    &mut values,
                    context,
                    index_memory,
                    "ODS consolidation source-address index",
                )?;
                values.push(item);
            }
        } else {
            push_index_char(&mut current, ch, context, index_memory)?;
        }
    }
    if !current.is_empty() {
        reserve_memory(
            context,
            index_memory,
            current.len().saturating_add(size_of::<String>()),
        )?;
        reserve_vec_slot(
            &mut values,
            context,
            index_memory,
            "ODS consolidation source-address index",
        )?;
        values.push(current);
    }
    Ok(values)
}

fn push_index_char(
    value: &mut String,
    character: char,
    context: &ExecutionContext,
    index_memory: &mut Option<Reservation>,
) -> Result<()> {
    let bytes = character.len_utf8();
    if value.len().checked_add(bytes).is_none() {
        return Err(invalid("ODS consolidation address memory overflows"));
    }
    if value.len().saturating_add(bytes) > value.capacity() {
        reserve_memory(context, index_memory, bytes)?;
        value
            .try_reserve_exact(bytes)
            .map_err(|error| allocation("ODS consolidation address scratch", error))?;
    }
    value.push(character);
    Ok(())
}

fn resolve_namespace(
    resolved: &ResolveResult<'_>,
    context: &ExecutionContext,
    index_memory: &mut Option<Reservation>,
) -> Result<Option<String>> {
    match resolved {
        ResolveResult::Bound(Namespace(uri)) => {
            reserve_memory(context, index_memory, uri.len())?;
            Ok(Some(decode(uri, "element namespace")?))
        },
        ResolveResult::Unbound => Ok(None),
        ResolveResult::Unknown(prefix) => Err(invalid(format!(
            "ODS metadata element uses unbound prefix '{}'",
            String::from_utf8_lossy(prefix.as_ref())
        ))),
    }
}

fn decode(bytes: &[u8], name: &str) -> Result<String> {
    String::from_utf8(bytes.to_vec()).map_err(|error| invalid(format!("invalid {name}: {error}")))
}

fn reserve_memory(
    context: &ExecutionContext,
    retained: &mut Option<Reservation>,
    bytes: usize,
) -> Result<()> {
    if bytes == 0 {
        return Ok(());
    }
    let reservation = context
        .reserve(
            Resource::Memory,
            u64::try_from(bytes).map_err(|_| invalid("ODS metadata memory size overflows u64"))?,
        )
        .map_err(map_execution)?;
    if let Some(current) = retained {
        if current.try_merge(reservation).is_err() {
            return Err(unsupported(
                "ODS metadata index memory reservations use incompatible budgets",
            ));
        }
    } else {
        *retained = Some(reservation);
    }
    Ok(())
}

fn reserve_vec_slot<T>(
    values: &mut Vec<T>,
    context: &ExecutionContext,
    retained: &mut Option<Reservation>,
    resource: &'static str,
) -> Result<()> {
    if values.len() == values.capacity() {
        reserve_memory(context, retained, size_of::<T>())?;
        values
            .try_reserve_exact(1)
            .map_err(|error| allocation(resource, error))?;
    }
    Ok(())
}

fn xml_error(error: quick_xml::Error) -> Error {
    invalid(format!("invalid ODS metadata XML: {error}"))
}

fn qname_prefix_or_empty(qname: &str) -> &str {
    qname.split_once(':').map_or("", |(prefix, _)| prefix)
}

fn table_attribute_prefix(span: &Span) -> Option<&str> {
    span.attrs
        .iter()
        .find(|attribute| attribute.namespace.as_deref() == Some(TABLE_NS))
        .map(|attribute| qname_prefix_or_empty(&attribute.qname))
        .filter(|prefix| !prefix.is_empty())
}

fn rename_table_qnames(source: &str, element_prefix: &str, attribute_prefix: &str) -> String {
    if element_prefix == "table" && attribute_prefix == "table" {
        return source.to_string();
    }
    let mut output = String::with_capacity(source.len());
    let mut in_tag = false;
    let mut element_name_done = false;
    let mut quote = None;
    let mut cursor = 0;
    while cursor < source.len() {
        let rest = &source[cursor..];
        let Some(character) = rest.chars().next() else {
            break;
        };
        if in_tag {
            if let Some(value) = quote {
                if character == value {
                    quote = None;
                }
                output.push(character);
                cursor += character.len_utf8();
                continue;
            }
            match character {
                '\'' | '"' => {
                    quote = Some(character);
                    output.push(character);
                    cursor += character.len_utf8();
                },
                '>' => {
                    in_tag = false;
                    output.push('>');
                    cursor += character.len_utf8();
                },
                _ if !element_name_done && character.is_ascii_whitespace() => {
                    element_name_done = true;
                    output.push(character);
                    cursor += character.len_utf8();
                },
                _ if rest.starts_with("table:") => {
                    let prefix = if element_name_done {
                        attribute_prefix
                    } else {
                        element_prefix
                    };
                    if !prefix.is_empty() {
                        output.push_str(prefix);
                        output.push(':');
                    }
                    cursor += "table:".len();
                },
                _ => {
                    output.push(character);
                    cursor += character.len_utf8();
                },
            }
        } else if character == '<' {
            in_tag = true;
            element_name_done = false;
            output.push('<');
            cursor += character.len_utf8();
        } else {
            output.push(character);
            cursor += character.len_utf8();
        }
    }
    output
}
fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}
fn unsupported(message: impl Into<String>) -> Error {
    Error::Unsupported(message.into())
}
fn limit(resource: Resource, scope: &'static str, observed: usize, maximum: usize) -> Error {
    Error::ResourceLimit(litchi_core::ResourceLimit {
        resource,
        observed: u64::try_from(observed).unwrap_or(u64::MAX),
        limit: u64::try_from(maximum).unwrap_or(u64::MAX),
        scope: Arc::from(format!("ODS metadata {scope}")),
    })
}
fn limit_u64(resource: Resource, scope: &'static str, observed: u64, maximum: u64) -> Error {
    Error::ResourceLimit(litchi_core::ResourceLimit {
        resource,
        observed,
        limit: maximum,
        scope: Arc::from(format!("ODS metadata {scope}")),
    })
}
fn charge_work(
    context: &ExecutionContext,
    work: &mut u64,
    amount: u64,
    maximum: u64,
) -> Result<()> {
    *work = work
        .checked_add(amount)
        .ok_or_else(|| invalid("ODS metadata work units overflow"))?;
    if *work > maximum {
        return Err(limit_u64(Resource::Work, "work units", *work, maximum));
    }
    context
        .consume(Resource::Work, amount)
        .map_err(map_execution)?;
    Ok(())
}
fn allocation(resource: &'static str, source: std::collections::TryReserveError) -> Error {
    Error::Allocation { resource, source }
}
fn map_execution(error: litchi_core::ExecutionError) -> Error {
    match error {
        litchi_core::ExecutionError::ResourceLimit(value) => Error::ResourceLimit(value),
        litchi_core::ExecutionError::Cancelled => unsupported("ODS metadata operation cancelled"),
        other => unsupported(format!(
            "ODS metadata execution policy rejected operation: {other}"
        )),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_core::{Budget, CancellationSource, ExecutionLimits, Limits as CoreLimits, Profile};
    use std::num::{NonZeroU64, NonZeroUsize};

    fn context() -> ExecutionContext {
        let budget = Budget::root(
            "ods-metadata-index-tests",
            CoreLimits::for_profile(Profile::TrustedBatch),
        );
        let (_source, token) = CancellationSource::pair();
        ExecutionContext::new(
            budget,
            token,
            ExecutionLimits::new(
                NonZeroUsize::new(1).unwrap(),
                NonZeroUsize::new(1).unwrap(),
                NonZeroU64::new(1024 * 1024).unwrap(),
                0,
            )
            .unwrap(),
        )
    }

    fn scan_fixture(source: &str) -> Result<Catalog> {
        scan(source, crate::sheet_metadata::Limits::default(), &context())
    }

    fn selector_limits(max_work_units: u64) -> crate::sheet_metadata::Limits {
        crate::sheet_metadata::Limits::new(1, 1024, 128, 1024, max_work_units)
            .expect("selector work test limits")
    }

    #[test]
    fn selector_work_budget_has_below_exact_and_above_boundaries() {
        let source = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><office:body><office:spreadsheet><table:table table:name="Data"><table:table-row><table:table-cell/></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>"#;
        let catalog = scan_fixture(source).expect("selector fixture");
        let selector =
            crate::sheet_metadata::CellSelector::by_position(litchi_core::Position::new(0), 0, 0);

        let below = context();
        assert!(
            catalog
                .cell_for_selector(&selector, &below, selector_limits(23))
                .is_err()
        );
        assert_eq!(below.budget().used(Resource::Work), 16);

        let exact = context();
        assert!(matches!(
            catalog.cell_for_selector(&selector, &exact, selector_limits(24)),
            Ok(PhysicalCell::Stored(_))
        ));
        assert_eq!(exact.budget().used(Resource::Work), 24);

        let above = context();
        assert!(matches!(
            catalog.cell_for_selector(&selector, &above, selector_limits(25)),
            Ok(PhysicalCell::Stored(_))
        ));
        assert_eq!(above.budget().used(Resource::Work), 24);
    }

    #[test]
    fn selector_read_checks_cancellation_before_comparing_candidates() {
        let source = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><office:body><office:spreadsheet><table:table table:name="Data"><table:table-row><table:table-cell/></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>"#;
        let catalog = scan_fixture(source).expect("selector fixture");
        let (cancellation, context) = {
            let budget = Budget::root(
                "ods-metadata-index-cancel-tests",
                CoreLimits::for_profile(Profile::TrustedBatch),
            );
            let (cancellation, token) = CancellationSource::pair();
            let context = ExecutionContext::new(
                budget,
                token,
                ExecutionLimits::new(
                    NonZeroUsize::new(1).unwrap(),
                    NonZeroUsize::new(1).unwrap(),
                    NonZeroU64::new(1024 * 1024).unwrap(),
                    0,
                )
                .unwrap(),
            );
            (cancellation, context)
        };
        cancellation.cancel();
        let selector =
            crate::sheet_metadata::CellSelector::by_position(litchi_core::Position::new(0), 0, 0);
        assert!(
            catalog
                .cell_for_selector(
                    &selector,
                    &context,
                    crate::sheet_metadata::Limits::default()
                )
                .is_err()
        );
        assert_eq!(context.budget().used(Resource::Work), 0);
    }

    #[test]
    fn admits_only_the_direct_document_content_spreadsheet_ancestry() {
        let valid = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><office:body><office:spreadsheet><table:table table:name="Data"/></office:spreadsheet></office:body></office:document-content>"#;
        assert!(scan_fixture(valid).is_ok());

        let foreign_root = r#"<v:wrapper xmlns:v="urn:example:foreign" xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:spreadsheet/></v:wrapper>"#;
        assert!(scan_fixture(foreign_root).is_err());

        let nested = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:body><v:wrapper xmlns:v="urn:example:foreign"><office:spreadsheet/></v:wrapper></office:body></office:document-content>"#;
        assert!(scan_fixture(nested).is_err());
        assert!(scan_fixture("<a/><b/>").is_err());
        assert!(scan_fixture("text<office:document-content xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\"><office:body><office:spreadsheet/></office:body></office:document-content>").is_err());
    }

    #[test]
    fn indexes_nested_row_containers_in_document_order() {
        let source = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><office:body><office:spreadsheet><table:table table:name="Data"><table:table-header-rows><table:table-row table:number-rows-repeated="2"><table:table-cell/></table:table-row></table:table-header-rows><table:table-row-group><table:table-rows><table:table-row><table:table-cell/></table:table-row></table:table-rows></table:table-row-group></table:table></office:spreadsheet></office:body></office:document-content>"#;
        let catalog = scan_fixture(source).expect("nested rows are admitted");
        assert_eq!(catalog.rows.len(), 2);
        assert_eq!(catalog.rows[0].logical_start, 0);
        assert_eq!(catalog.rows[0].repeat, 2);
        assert_eq!(catalog.rows[1].logical_start, 2);
        assert_eq!(catalog.rows[1].repeat, 1);
        assert_eq!(catalog.cells.len(), 2);
        assert_eq!(catalog.cells[1].row_start, 2);
    }

    #[test]
    fn rejects_unknown_text_and_hidden_metadata_as_editable_cell_owners() {
        let source = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><office:body><office:spreadsheet><table:table table:name="Data"><table:table-row><table:table-cell><text:future><table:detective/></text:future></table:table-cell><table:table-cell><table:detective/><text:p><table:detective/></text:p></table:table-cell></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>"#;
        let catalog = scan_fixture(source).expect("opaque cell content remains readable");
        assert_eq!(catalog.cells.len(), 2);
        assert!(catalog.cells.iter().all(|cell| cell.unsupported_child));
        assert_eq!(catalog.cells[1].direct_owner_count, 0);
    }
}
