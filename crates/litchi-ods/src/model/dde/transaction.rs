//! Source-checked, inert transactions for ODF DDE metadata.
//!
//! The transaction owns only declarations and cached XML.  It never resolves
//! a topic, opens a package, refreshes a cache, starts a process, or performs
//! ambient I/O.  Known owners created by this module are emitted with compact
//! local namespace declarations; unknown markup is retained byte-for-byte or
//! the requested destructive edit is refused.

use super::{
    AutomaticUpdate, ConversionMode, Limits, Link, OFFICE, SheetSource, Snapshot, Source, TABLE,
};
use litchi_core::{
    Error as CoreError, ExecutionContext, ExecutionError, Position, Reservation, Resource,
    Result as CoreResult,
};
use litchi_odf_common::datatype::Duration as OdfDuration;
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::{Namespace, ResolveResult},
    reader::NsReader,
};
use std::{borrow::Cow, mem::size_of, ops::Range, sync::Arc};

const TEXT: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const MARKUP_COMPATIBILITY: &[u8] = b"http://schemas.openxmlformats.org/markup-compatibility/2006";
const MAX_CACHE_ROWS: usize = 65_536;
const MAX_CACHE_CELLS: usize = 4_194_304;

/// A scalar value that can be authored in an inert DDE cached table.
///
/// Formula evaluation and error values are intentionally absent.  Existing
/// cache tables remain opaque source XML, while newly authored values use only
/// these finite, schema-owned scalar forms.
#[derive(Clone, Debug, PartialEq)]
#[non_exhaustive]
pub enum CachedValue {
    /// An empty table cell.
    Empty,
    /// A string displayed by the cache cell.
    Text(String),
    /// An ODF floating-point value.
    Number(f64),
    /// A floating-point value with a nonempty currency identifier.
    Currency { value: f64, currency: String },
    /// A percentage value represented by its ODF numeric value.
    Percentage(f64),
    /// An ODF boolean value.
    Boolean(bool),
    /// An ODF date or dateTime lexical value.
    Date(String),
    /// An XML Schema duration lexical value used for an ODF time.
    Time(String),
}

impl CachedValue {
    fn validate(&self) -> CoreResult<()> {
        match self {
            Self::Empty | Self::Boolean(_) => Ok(()),
            Self::Text(value) => validate_text(value, "cached text"),
            Self::Number(value) | Self::Percentage(value) => {
                if value.is_finite() {
                    Ok(())
                } else {
                    Err(invalid_core("cached numeric values must be finite"))
                }
            },
            Self::Currency { value, currency } => {
                if !value.is_finite() {
                    return Err(invalid_core("cached currency values must be finite"));
                }
                validate_text(currency, "cached currency")?;
                if currency.is_empty() {
                    return Err(invalid_core("cached currency must not be empty"));
                }
                Ok(())
            },
            Self::Date(value) => validate_date(value),
            Self::Time(value) => validate_duration(value, "cached time"),
        }
    }
}

/// One typed cell in a newly authored cached table.
#[derive(Clone, Debug, PartialEq)]
pub struct CachedCell {
    value: CachedValue,
}

impl CachedCell {
    /// Creates a checked cached cell value.
    pub fn new(value: CachedValue) -> CoreResult<Self> {
        value.validate()?;
        Ok(Self { value })
    }

    /// Returns the scalar value represented by this cell.
    #[must_use]
    pub fn value(&self) -> &CachedValue {
        &self.value
    }
}

/// One ordered row in a typed cached table.
#[derive(Clone, Debug, PartialEq)]
pub struct CachedRow {
    cells: Vec<CachedCell>,
}

impl CachedRow {
    /// Creates an empty row to which cells may be appended.
    #[must_use]
    pub const fn new() -> Self {
        Self { cells: Vec::new() }
    }

    /// Creates and validates a row from its cells.
    pub fn from_cells(cells: Vec<CachedCell>) -> CoreResult<Self> {
        if cells.len() > MAX_CACHE_CELLS {
            return Err(resource_core(
                "DDE cached cells",
                cells.len(),
                MAX_CACHE_CELLS,
            ));
        }
        Ok(Self { cells })
    }

    /// Appends one checked cell.
    pub fn push(&mut self, cell: CachedCell) -> CoreResult<()> {
        if self.cells.len() >= MAX_CACHE_CELLS {
            return Err(resource_core(
                "DDE cached cells",
                self.cells.len() + 1,
                MAX_CACHE_CELLS,
            ));
        }
        self.cells
            .try_reserve(1)
            .map_err(|_| CoreError::Unsupported("DDE cached row allocation failed".to_string()))?;
        self.cells.push(cell);
        Ok(())
    }

    /// Borrow cells in source order.
    #[must_use]
    pub fn cells(&self) -> &[CachedCell] {
        &self.cells
    }
}

impl Default for CachedRow {
    fn default() -> Self {
        Self::new()
    }
}

/// A finite, typed table used when authoring a new DDE cache.
#[derive(Clone, Debug, PartialEq)]
pub struct CachedTable {
    rows: Vec<CachedRow>,
}

impl CachedTable {
    /// Creates an empty cached table.
    #[must_use]
    pub const fn new() -> Self {
        Self { rows: Vec::new() }
    }

    /// Creates and validates a table from ordered rows.
    pub fn from_rows(rows: Vec<CachedRow>) -> CoreResult<Self> {
        if rows.len() > MAX_CACHE_ROWS {
            return Err(resource_core("DDE cached rows", rows.len(), MAX_CACHE_ROWS));
        }
        let mut cells = 0usize;
        for row in &rows {
            cells = cells
                .checked_add(row.cells.len())
                .ok_or_else(|| invalid_core("DDE cached cell count overflows"))?;
            if cells > MAX_CACHE_CELLS {
                return Err(resource_core("DDE cached cells", cells, MAX_CACHE_CELLS));
            }
        }
        Ok(Self { rows })
    }

    /// Appends one row while enforcing finite authoring bounds.
    pub fn push(&mut self, row: CachedRow) -> CoreResult<()> {
        if self.rows.len() >= MAX_CACHE_ROWS {
            return Err(resource_core(
                "DDE cached rows",
                self.rows.len() + 1,
                MAX_CACHE_ROWS,
            ));
        }
        let current_cells = self
            .rows
            .iter()
            .try_fold(0usize, |sum, row| sum.checked_add(row.cells.len()))
            .ok_or_else(|| invalid_core("DDE cached cell count overflows"))?;
        let total_cells = current_cells
            .checked_add(row.cells.len())
            .ok_or_else(|| invalid_core("DDE cached cell count overflows"))?;
        if total_cells > MAX_CACHE_CELLS {
            return Err(resource_core(
                "DDE cached cells",
                total_cells,
                MAX_CACHE_CELLS,
            ));
        }
        self.rows.try_reserve(1).map_err(|_| {
            CoreError::Unsupported("DDE cached table allocation failed".to_string())
        })?;
        self.rows.push(row);
        Ok(())
    }

    /// Borrow rows in source order.
    #[must_use]
    pub fn rows(&self) -> &[CachedRow] {
        &self.rows
    }
}

impl Default for CachedTable {
    fn default() -> Self {
        Self::new()
    }
}

/// A complete formula-link declaration and typed cache replacement.
#[derive(Clone, Debug)]
pub struct LinkSpec {
    source: Source,
    cached_table: CachedTable,
    /// Conservative retained heap footprint of this owned specification.
    ///
    /// It is computed once at construction so staging can admit the complete
    /// payload without walking the cache again.  Clones may retain an upper
    /// bound from the original allocation, which is safe over-admission.
    retained_bytes: usize,
}

impl PartialEq for LinkSpec {
    fn eq(&self, other: &Self) -> bool {
        self.source == other.source && self.cached_table == other.cached_table
    }
}

impl LinkSpec {
    /// Creates a complete inert link specification.
    pub fn new(source: Source, cached_table: CachedTable) -> CoreResult<Self> {
        validate_authored_source(&source)?;
        validate_cached_table(&cached_table)?;
        let retained_bytes = link_spec_retained_bytes(&source, &cached_table)?;
        Ok(Self {
            source,
            cached_table,
            retained_bytes,
        })
    }

    /// Returns the source declaration.
    #[must_use]
    pub fn source(&self) -> &Source {
        &self.source
    }

    /// Returns typed cached rows.
    #[must_use]
    pub fn cached_table(&self) -> &CachedTable {
        &self.cached_table
    }

    fn validate(&self) -> CoreResult<()> {
        // `LinkSpec` fields are private and every constructor validates the
        // bounded cache.  Staging therefore only needs to recheck the source
        // authoring invariant; walking every cached cell again would be an
        // uncharged duplicate scan.
        validate_authored_source(&self.source)
    }
}

/// Return the heap bytes retained by a source after it is moved into a
/// staged draft.  The `String` objects themselves live in their enclosing
/// draft allocation; only their backing buffers need a separate admission.
fn source_retained_heap_bytes(source: &Source) -> CoreResult<usize> {
    let mut amount = source
        .application
        .capacity()
        .checked_add(source.topic.capacity())
        .and_then(|value| value.checked_add(source.item.capacity()))
        .ok_or_else(|| invalid_core("DDE source payload size overflows"))?;
    if let Some(name) = &source.name {
        amount = amount
            .checked_add(name.capacity())
            .ok_or_else(|| invalid_core("DDE source name payload size overflows"))?;
    }
    Ok(amount)
}

/// Return the heap bytes retained by a typed cache.  Capacity, rather than
/// logical length, is used because moving a caller-provided `Vec` preserves
/// its allocation and staging must admit the complete retained payload.
fn cached_table_retained_heap_bytes(table: &CachedTable) -> CoreResult<usize> {
    let mut amount = table
        .rows
        .capacity()
        .checked_mul(size_of::<CachedRow>())
        .ok_or_else(|| invalid_core("DDE cached row payload size overflows"))?;
    for row in &table.rows {
        amount = amount
            .checked_add(
                row.cells
                    .capacity()
                    .checked_mul(size_of::<CachedCell>())
                    .ok_or_else(|| invalid_core("DDE cached cell payload size overflows"))?,
            )
            .ok_or_else(|| invalid_core("DDE cached payload size overflows"))?;
        for cell in &row.cells {
            let string_capacity = match &cell.value {
                CachedValue::Text(value) | CachedValue::Date(value) | CachedValue::Time(value) => {
                    value.capacity()
                },
                CachedValue::Currency { currency, .. } => currency.capacity(),
                CachedValue::Empty
                | CachedValue::Number(_)
                | CachedValue::Percentage(_)
                | CachedValue::Boolean(_) => 0,
            };
            amount = amount
                .checked_add(string_capacity)
                .ok_or_else(|| invalid_core("DDE cached value payload size overflows"))?;
        }
    }
    Ok(amount)
}

/// Cache and source heap bytes retained by one complete authored link.
fn link_spec_retained_bytes(source: &Source, table: &CachedTable) -> CoreResult<usize> {
    source_retained_heap_bytes(source)?
        .checked_add(cached_table_retained_heap_bytes(table)?)
        .ok_or_else(|| invalid_core("DDE link payload size overflows"))
}

fn sheet_draft_retained_bytes(sheet_capacity: usize, source: &Source) -> CoreResult<usize> {
    sheet_capacity
        .checked_add(source_retained_heap_bytes(source)?)
        .ok_or_else(|| invalid_core("DDE sheet payload size overflows"))
}

/// Checked worksheet selector used by source-declaration transactions.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum SheetSelector<'a> {
    /// Select one worksheet by its exact ODF table name.
    Name(Cow<'a, str>),
    /// Select one worksheet by zero-based table position.
    Position(Position),
}

impl<'a> From<&'a str> for SheetSelector<'a> {
    fn from(value: &'a str) -> Self {
        Self::Name(Cow::Borrowed(value))
    }
}

impl From<String> for SheetSelector<'static> {
    fn from(value: String) -> Self {
        Self::Name(Cow::Owned(value))
    }
}

impl From<usize> for SheetSelector<'static> {
    fn from(value: usize) -> Self {
        Self::Position(Position::new(value))
    }
}

impl From<Position> for SheetSelector<'static> {
    fn from(value: Position) -> Self {
        Self::Position(value)
    }
}

/// A staged, source-bound inert DDE edit.
#[derive(Debug)]
pub struct Edit {
    before: Snapshot,
    sheet_sources: Option<Vec<SheetDraft>>,
    links: Option<Vec<LinkDraft>>,
    staging_reservations: Vec<Arc<Reservation>>,
}

#[derive(Debug)]
struct SheetDraft {
    table_index: usize,
    sheet: String,
    source: Source,
    /// Reservation for the heap payload owned by this draft.  It is kept
    /// beside the payload so replacing or removing a draft releases the
    /// corresponding admission immediately.
    _reservation: Arc<Reservation>,
}

#[derive(Debug)]
enum LinkDraft {
    Existing(usize),
    Source {
        existing: usize,
        source: Source,
        _reservation: Arc<Reservation>,
    },
    Owned {
        link: LinkSpec,
        _reservation: Arc<Reservation>,
        /// Existing link identity replaced at the staging position, if any.
        /// Retaining this identity keeps opaque-owner admission correct after
        /// a preceding reorder.
        replaced: Option<usize>,
    },
}

impl Edit {
    pub(crate) fn new(before: Snapshot) -> Self {
        Self {
            before,
            sheet_sources: None,
            links: None,
            staging_reservations: Vec::new(),
        }
    }

    /// Borrow the immutable source snapshot from which this edit was staged.
    #[must_use]
    pub fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Number of staged formula links.
    #[must_use]
    pub fn link_count(&self) -> usize {
        self.links
            .as_ref()
            .map_or_else(|| self.before.links().len(), Vec::len)
    }

    /// Number of staged sheet-local source declarations.
    #[must_use]
    pub fn sheet_source_count(&self) -> usize {
        self.sheet_sources
            .as_ref()
            .map_or_else(|| self.before.sheet_sources().len(), Vec::len)
    }

    /// Borrow one staged source declaration by source order.
    #[must_use]
    pub fn sheet_source(&self, index: usize) -> Option<(&str, &Source)> {
        if let Some(sources) = &self.sheet_sources {
            sources
                .get(index)
                .map(|entry| (entry.sheet.as_str(), &entry.source))
        } else {
            self.before
                .sheet_sources()
                .get(index)
                .map(|entry| (entry.sheet(), entry.source()))
        }
    }

    fn reserve_staging_token(&self, amount: usize) -> CoreResult<Arc<Reservation>> {
        self.before.context().check().map_err(map_execution_core)?;
        self.before
            .context()
            .consume(Resource::Work, 1)
            .map_err(map_execution_core)?;
        let amount = u64::try_from(amount)
            .map_err(|_| invalid_core("DDE staging reservation exceeds u64"))?;
        let reservation = self
            .before
            .context()
            .reserve(Resource::Memory, amount)
            .map_err(map_execution_core)
            .map(Arc::new)?;
        Ok(reservation)
    }

    fn reserve_payload(&self, amount: usize) -> CoreResult<Arc<Reservation>> {
        self.reserve_staging_token(amount)
    }

    /// Admit the capacity delta needed by one exact vector growth.  The
    /// returned token must be retained only after `try_reserve_exact` succeeds
    /// so a failed allocation cannot leave a phantom reservation in the edit.
    fn reserve_vec_growth<T>(
        &mut self,
        length: usize,
        capacity: usize,
        additional: usize,
    ) -> CoreResult<Option<Arc<Reservation>>> {
        let required = length
            .checked_add(additional)
            .ok_or_else(|| invalid_core("DDE draft vector size overflows"))?;
        if required <= capacity {
            return Ok(None);
        }
        let delta = required
            .checked_sub(capacity)
            .and_then(|value| value.checked_mul(size_of::<T>()))
            .ok_or_else(|| invalid_core("DDE draft vector memory size overflows"))?;
        self.staging_reservations.try_reserve(1).map_err(|_| {
            CoreError::Unsupported("DDE staging reservation allocation failed".to_string())
        })?;
        Ok(Some(self.reserve_staging_token(delta)?))
    }

    fn retain_vector_growth(&mut self, reservation: Option<Arc<Reservation>>) {
        if let Some(reservation) = reservation {
            self.staging_reservations.push(reservation);
        }
    }

    fn ensure_sheet_sources(&mut self) -> CoreResult<&mut Vec<SheetDraft>> {
        if self.sheet_sources.is_none() {
            let estimate = self
                .before
                .sheet_sources()
                .len()
                .checked_mul(size_of::<SheetDraft>())
                .and_then(|value| value.checked_add(1))
                .ok_or_else(|| invalid_core("DDE sheet draft size overflows"))?;
            // Keep the vector reservation detached until every cloned draft
            // and its per-draft payload admission has succeeded.  This makes
            // lazy initialization atomic when the first mutating operation is
            // refused by a finite memory budget.
            self.staging_reservations.try_reserve(1).map_err(|_| {
                CoreError::Unsupported("DDE staging reservation allocation failed".to_string())
            })?;
            let vector_reservation = self.reserve_staging_token(estimate)?;
            let mut sources = Vec::new();
            if sources
                .try_reserve_exact(self.before.sheet_sources().len())
                .is_err()
            {
                drop(vector_reservation);
                return Err(CoreError::Unsupported(
                    "DDE sheet draft allocation failed".to_string(),
                ));
            }
            for source_index in 0..self.before.sheet_sources().len() {
                let (sheet_len, source_size) = {
                    let entry = &self.before.sheet_sources()[source_index];
                    (
                        entry.sheet.capacity(),
                        source_retained_heap_bytes(entry.source())?,
                    )
                };
                let payload = sheet_len
                    .checked_add(source_size)
                    .ok_or_else(|| invalid_core("DDE sheet draft payload size overflows"))?;
                let reservation = self.reserve_payload(payload)?;
                self.before.context().check().map_err(map_execution_core)?;
                self.before
                    .context()
                    .consume(Resource::Work, 1)
                    .map_err(map_execution_core)?;
                let table_index = self
                    .before
                    .inventory()
                    .sheet_table_indices
                    .get(source_index)
                    .copied()
                    .unwrap_or(usize::MAX);
                let entry = &self.before.sheet_sources()[source_index];
                sources.push(SheetDraft {
                    table_index,
                    sheet: entry.sheet().to_string(),
                    source: entry.source().clone(),
                    _reservation: reservation,
                });
            }
            sources.sort_by_key(|entry| entry.table_index);
            self.staging_reservations.push(vector_reservation);
            self.sheet_sources = Some(sources);
        }
        Ok(self
            .sheet_sources
            .as_mut()
            .expect("sheet drafts initialized"))
    }

    fn ensure_links(&mut self) -> CoreResult<&mut Vec<LinkDraft>> {
        if self.links.is_none() {
            let estimate = self
                .before
                .links()
                .len()
                .checked_mul(size_of::<LinkDraft>())
                .and_then(|value| value.checked_add(1))
                .ok_or_else(|| invalid_core("DDE link draft size overflows"))?;
            // See `ensure_sheet_sources`: retain the vector admission only
            // after the initial draft vector has been built successfully.
            self.staging_reservations.try_reserve(1).map_err(|_| {
                CoreError::Unsupported("DDE staging reservation allocation failed".to_string())
            })?;
            let vector_reservation = self.reserve_staging_token(estimate)?;
            let mut links = Vec::new();
            if links.try_reserve_exact(self.before.links().len()).is_err() {
                drop(vector_reservation);
                return Err(CoreError::Unsupported(
                    "DDE link draft allocation failed".to_string(),
                ));
            }
            links.extend((0..self.before.links().len()).map(LinkDraft::Existing));
            self.staging_reservations.push(vector_reservation);
            self.links = Some(links);
        }
        Ok(self.links.as_mut().expect("link drafts initialized"))
    }

    /// Add or replace the source declaration on one existing worksheet.
    pub fn set_sheet_source<'a>(
        &mut self,
        selector: impl Into<SheetSelector<'a>>,
        source: Source,
    ) -> CoreResult<()> {
        let table_index = self.resolve_sheet(selector.into())?;
        let current = self.desired_sheet_source(table_index);
        if current == Some(&source) {
            self.before.context().check().map_err(map_execution_core)?;
            return Ok(());
        }
        validate_authored_source(&source)?;

        let draft_index = self.sheet_sources.as_ref().and_then(|entries| {
            entries
                .iter()
                .position(|entry| entry.table_index == table_index)
        });
        let sheet_capacity = if let Some(draft_index) = draft_index {
            self.sheet_sources
                .as_ref()
                .and_then(|entries| entries.get(draft_index))
                .map(|entry| entry.sheet.capacity())
                .ok_or_else(|| invalid_core("DDE sheet source draft disappeared"))?
        } else {
            self.before
                .inventory()
                .table_names
                .get(table_index)
                .and_then(Option::as_ref)
                .map(String::capacity)
                .ok_or_else(|| {
                    invalid_core("an unnamed worksheet cannot receive a named DDE source")
                })?
        };
        // The replacement token covers the complete new `SheetDraft` heap
        // payload.  The old token also covers the sheet-name buffer, so that
        // name must be included before the old draft is replaced.
        let reservation =
            self.reserve_payload(sheet_draft_retained_bytes(sheet_capacity, &source)?)?;

        if let Some(draft_index) = draft_index {
            let existing = self
                .sheet_sources
                .as_mut()
                .and_then(|entries| entries.get_mut(draft_index))
                .expect("sheet draft index was found");
            existing.source = source;
            existing._reservation = reservation;
            return Ok(());
        }

        let had_sources = self.sheet_sources.is_some();
        let initial_reservations = self.staging_reservations.len();
        if let Err(error) = self.ensure_sheet_sources() {
            drop(reservation);
            return Err(error);
        }
        // A source present in the retained snapshot may not have had a draft
        // yet.  `ensure_sheet_sources` materializes it; reuse that entry
        // rather than accidentally appending a duplicate declaration.
        if let Some(existing_index) = self.sheet_sources.as_ref().and_then(|entries| {
            entries
                .iter()
                .position(|entry| entry.table_index == table_index)
        }) {
            let existing = self
                .sheet_sources
                .as_mut()
                .and_then(|entries| entries.get_mut(existing_index))
                .expect("sheet draft index was found");
            existing.source = source;
            existing._reservation = reservation;
            return Ok(());
        }
        let (length, capacity) = {
            let sources = self
                .sheet_sources
                .as_ref()
                .expect("sheet drafts initialized");
            (sources.len(), sources.capacity())
        };
        let growth = match self.reserve_vec_growth::<SheetDraft>(length, capacity, 1) {
            Ok(growth) => growth,
            Err(error) => {
                drop(reservation);
                if !had_sources {
                    self.sheet_sources = None;
                    self.staging_reservations.truncate(initial_reservations);
                }
                return Err(error);
            },
        };
        let sheet = self
            .before
            .inventory()
            .table_names
            .get(table_index)
            .and_then(Option::as_ref)
            .cloned()
            .ok_or_else(|| invalid_core("DDE worksheet name disappeared during staging"));
        let sheet = match sheet {
            Ok(sheet) => sheet,
            Err(error) => {
                drop(reservation);
                drop(growth);
                if !had_sources {
                    self.sheet_sources = None;
                    self.staging_reservations.truncate(initial_reservations);
                }
                return Err(error);
            },
        };
        {
            let sources = self
                .sheet_sources
                .as_mut()
                .expect("sheet drafts initialized");
            if sources.try_reserve_exact(1).is_err() {
                drop(reservation);
                drop(growth);
                if !had_sources {
                    self.sheet_sources = None;
                    self.staging_reservations.truncate(initial_reservations);
                }
                return Err(CoreError::Unsupported(
                    "DDE sheet source allocation failed".to_string(),
                ));
            }
        }
        self.retain_vector_growth(growth);
        let sources = self
            .sheet_sources
            .as_mut()
            .expect("sheet drafts initialized");
        sources.push(SheetDraft {
            table_index,
            sheet,
            source,
            _reservation: reservation,
        });
        sources.sort_by_key(|entry| entry.table_index);
        Ok(())
    }

    /// Replace an existing source declaration on one worksheet.
    pub fn replace_sheet_source<'a>(
        &mut self,
        selector: impl Into<SheetSelector<'a>>,
        source: Source,
    ) -> CoreResult<()> {
        let table_index = self.resolve_sheet(selector.into())?;
        let has_source = self.desired_sheet_source(table_index).is_some();
        if !has_source {
            return Err(invalid_core("DDE sheet source was not found"));
        }
        if self
            .desired_sheet_source(table_index)
            .is_some_and(|current| current == &source)
        {
            self.before.context().check().map_err(map_execution_core)?;
            return Ok(());
        }
        validate_authored_source(&source)?;

        let draft_index = self.sheet_sources.as_ref().and_then(|entries| {
            entries
                .iter()
                .position(|entry| entry.table_index == table_index)
        });
        let sheet_capacity = if let Some(draft_index) = draft_index {
            self.sheet_sources
                .as_ref()
                .and_then(|entries| entries.get(draft_index))
                .map(|entry| entry.sheet.capacity())
                .ok_or_else(|| invalid_core("DDE sheet source draft disappeared"))?
        } else {
            self.before
                .inventory()
                .table_names
                .get(table_index)
                .and_then(Option::as_ref)
                .map(String::capacity)
                .ok_or_else(|| invalid_core("DDE sheet source was not found"))?
        };
        let reservation =
            self.reserve_payload(sheet_draft_retained_bytes(sheet_capacity, &source)?)?;

        if let Some(existing_index) = draft_index {
            let existing = self
                .sheet_sources
                .as_mut()
                .and_then(|entries| entries.get_mut(existing_index))
                .expect("sheet draft index was found");
            existing.source = source;
            existing._reservation = reservation;
            return Ok(());
        }

        let had_sources = self.sheet_sources.is_some();
        let initial_reservations = self.staging_reservations.len();
        if let Err(error) = self.ensure_sheet_sources() {
            drop(reservation);
            return Err(error);
        }
        let Some(existing_index) = self.sheet_sources.as_ref().and_then(|entries| {
            entries
                .iter()
                .position(|entry| entry.table_index == table_index)
        }) else {
            drop(reservation);
            if !had_sources {
                self.sheet_sources = None;
                self.staging_reservations.truncate(initial_reservations);
            }
            return Err(invalid_core("DDE sheet source was not found"));
        };
        let existing = self
            .sheet_sources
            .as_mut()
            .and_then(|entries| entries.get_mut(existing_index))
            .expect("sheet draft index was found");
        existing.source = source;
        existing._reservation = reservation;
        Ok(())
    }

    /// Remove one source declaration from a worksheet.
    pub fn remove_sheet_source<'a>(
        &mut self,
        selector: impl Into<SheetSelector<'a>>,
    ) -> CoreResult<Source> {
        let table_index = self.resolve_sheet(selector.into())?;
        if let Some(sources) = self.sheet_sources.as_ref() {
            let index = sources
                .iter()
                .position(|entry| entry.table_index == table_index)
                .ok_or_else(|| invalid_core("DDE sheet source was not found"))?;
            return Ok(self
                .sheet_sources
                .as_mut()
                .expect("sheet drafts initialized")
                .remove(index)
                .source);
        }
        // Check the retained inventory before lazy cloning.  A missing
        // declaration must not allocate a complete sheet-draft vector merely
        // to return the selector error.
        if !self
            .before
            .inventory()
            .sheet_table_indices
            .contains(&table_index)
        {
            self.before.context().check().map_err(map_execution_core)?;
            return Err(invalid_core("DDE sheet source was not found"));
        }
        let sources = self.ensure_sheet_sources()?;
        let index = sources
            .iter()
            .position(|entry| entry.table_index == table_index)
            .ok_or_else(|| invalid_core("DDE sheet source was not found"))?;
        Ok(sources.remove(index).source)
    }

    /// Add a source declaration to a worksheet that does not already have one.
    ///
    /// Use [`Self::set_sheet_source`] when replacement of an existing
    /// declaration is intentional. Keeping this operation create-only avoids
    /// an accidental overwrite hidden behind an `add` name.
    pub fn add_sheet_source<'a>(
        &mut self,
        selector: impl Into<SheetSelector<'a>>,
        source: Source,
    ) -> CoreResult<()> {
        let table_index = self.resolve_sheet(selector.into())?;
        if self.desired_sheet_source(table_index).is_some() {
            return Err(invalid_core(
                "DDE worksheet already has a source declaration",
            ));
        }
        self.set_sheet_source(Position::new(table_index), source)
    }

    /// Append a new formula link after the current last link.
    pub fn add_link(&mut self, link: LinkSpec) -> CoreResult<()> {
        self.insert_link(self.link_count(), link)
    }

    /// Insert a new formula link at a checked source-order position.
    pub fn insert_link(&mut self, index: usize, link: LinkSpec) -> CoreResult<()> {
        let current_len = self.link_count();
        if index > current_len {
            return Err(invalid_core("DDE link insertion position did not match"));
        }
        if current_len >= self.before.limits().links {
            return Err(resource_core(
                "DDE links",
                current_len + 1,
                self.before.limits().links,
            ));
        }
        link.validate()?;
        // Admit the moved caller-owned payload before lazily cloning the
        // existing draft vector.  A refusal therefore leaves both the edit
        // and its memory usage unchanged.
        let reservation = self.reserve_payload(link.retained_bytes)?;
        let had_links = self.links.is_some();
        let initial_reservations = self.staging_reservations.len();
        self.ensure_links()?;
        let (length, capacity) = {
            let links = self.links.as_ref().expect("link drafts initialized");
            (links.len(), links.capacity())
        };
        let growth = match self.reserve_vec_growth::<LinkDraft>(length, capacity, 1) {
            Ok(growth) => growth,
            Err(error) => {
                drop(reservation);
                if !had_links {
                    self.links = None;
                    self.staging_reservations.truncate(initial_reservations);
                }
                return Err(error);
            },
        };
        {
            let links = self.links.as_mut().expect("link drafts initialized");
            if links.try_reserve_exact(1).is_err() {
                drop(reservation);
                drop(growth);
                if !had_links {
                    self.links = None;
                    self.staging_reservations.truncate(initial_reservations);
                }
                return Err(CoreError::Unsupported(
                    "DDE link allocation failed".to_string(),
                ));
            }
        }
        self.retain_vector_growth(growth);
        let links = self.links.as_mut().expect("link drafts initialized");
        links.insert(
            index,
            LinkDraft::Owned {
                link,
                _reservation: reservation,
                replaced: None,
            },
        );
        Ok(())
    }

    /// Replace one formula link selected by source-order position.
    pub fn replace_link(&mut self, index: impl Into<Position>, link: LinkSpec) -> CoreResult<()> {
        self.replace_link_at(index, link)
    }

    /// Replace one formula link selected by source-order position.
    pub fn replace_link_at(
        &mut self,
        index: impl Into<Position>,
        link: LinkSpec,
    ) -> CoreResult<()> {
        let index = index.into().get();
        if index >= self.link_count() {
            return Err(invalid_core("DDE link position did not match"));
        }
        link.validate()?;
        // Admit the moved payload before lazy initialization.  If the
        // replacement is refused, no draft-vector reservation is retained.
        let reservation = self.reserve_payload(link.retained_bytes)?;
        self.ensure_links()?;
        let Some(existing) = self.links.as_ref().and_then(|links| links.get(index)) else {
            drop(reservation);
            return Err(invalid_core("DDE link position did not match"));
        };
        let replaced = match existing {
            LinkDraft::Existing(existing) => Some(*existing),
            LinkDraft::Source { existing, .. } => Some(*existing),
            LinkDraft::Owned { replaced, .. } => *replaced,
        };
        let same = self
            .links
            .as_ref()
            .and_then(|links| links.get(index))
            .is_some_and(|existing| {
                matches!(existing, LinkDraft::Owned { link: current, .. } if current == &link)
        });
        if same {
            drop(reservation);
            return Ok(());
        }
        let slot = self
            .links
            .as_mut()
            .and_then(|links| links.get_mut(index))
            .expect("link draft index was found");
        *slot = LinkDraft::Owned {
            link,
            _reservation: reservation,
            replaced,
        };
        Ok(())
    }

    /// Replace only one formula link's source declaration, preserving its
    /// existing cached table and any opaque cache markup byte-for-byte.
    pub fn replace_link_source(
        &mut self,
        index: impl Into<Position>,
        source: Source,
    ) -> CoreResult<()> {
        let index = index.into().get();
        let existing = self.current_link_identity(index)?;
        let same = self
            .current_link_source(index)
            .is_some_and(|current| current == &source);
        if same {
            self.before.context().check().map_err(map_execution_core)?;
            // Repeating a source replacement must preserve the staged value.
            // Normalize only when the requested source is the original value,
            // which gives an exact inverse without turning A -> B -> B back
            // into A.
            let normalize = self
                .links
                .as_ref()
                .and_then(|links| links.get(index))
                .is_some_and(|draft| {
                    matches!(
                        draft,
                        LinkDraft::Source { existing, .. }
                            if self
                                .before
                                .links()
                                .get(*existing)
                                .is_some_and(|link| link.source() == &source)
                    )
                });
            if normalize {
                let slot = self
                    .links
                    .as_mut()
                    .expect("link drafts initialized")
                    .get_mut(index)
                    .expect("link draft index was found");
                *slot = LinkDraft::Existing(existing);
            }
            return Ok(());
        }
        validate_authored_source(&source)?;
        // Admit the moved source before lazy initialization of the draft
        // vector.  A tight budget must not leave that vector's reservation
        // behind when the source payload itself is refused.
        let reservation = self.reserve_payload(source_retained_heap_bytes(&source)?)?;
        self.ensure_links()?;
        if self
            .links
            .as_ref()
            .and_then(|links| links.get(index))
            .is_none()
        {
            return Err(invalid_core("DDE link position did not match"));
        }
        let slot = self
            .links
            .as_mut()
            .expect("link drafts initialized")
            .get_mut(index)
            .expect("link draft index was found");
        *slot = LinkDraft::Source {
            existing,
            source,
            _reservation: reservation,
        };
        Ok(())
    }

    fn current_link_identity(&self, index: usize) -> CoreResult<usize> {
        let slot = if let Some(links) = &self.links {
            links.get(index)
        } else if index < self.before.links().len() {
            return Ok(index);
        } else {
            None
        };
        let Some(slot) = slot else {
            return Err(invalid_core("DDE link position did not match"));
        };
        match slot {
            LinkDraft::Existing(existing) | LinkDraft::Source { existing, .. } => Ok(*existing),
            LinkDraft::Owned { .. } => Err(invalid_core(
                "source-only DDE link replacement requires an existing link",
            )),
        }
    }

    fn current_link_source(&self, index: usize) -> Option<&Source> {
        if let Some(links) = &self.links {
            let slot = links.get(index)?;
            return match slot {
                LinkDraft::Existing(existing) => {
                    self.before.links().get(*existing).map(Link::source)
                },
                LinkDraft::Source { source, .. } => Some(source),
                LinkDraft::Owned { link, .. } => Some(link.source()),
            };
        }
        self.before.links().get(index).map(Link::source)
    }

    /// Replace only a uniquely named formula link's source declaration.
    pub fn replace_link_source_named(&mut self, name: &str, source: Source) -> CoreResult<()> {
        let index = self.unique_named_link(name)?;
        self.replace_link_source(index, source)
    }

    /// Remove one formula link selected by source-order position.
    pub fn remove_link(&mut self, index: impl Into<Position>) -> CoreResult<()> {
        self.remove_link_at(index)
    }

    /// Remove one formula link selected by source-order position.
    pub fn remove_link_at(&mut self, index: impl Into<Position>) -> CoreResult<()> {
        let index = index.into().get();
        if index >= self.link_count() {
            return Err(invalid_core("DDE link position did not match"));
        }
        if self.ensure_links()?.get(index).is_none() {
            return Err(invalid_core("DDE link position did not match"));
        }
        self.links
            .as_mut()
            .expect("link drafts initialized")
            .remove(index);
        Ok(())
    }

    /// Move one formula link to another source-order position.
    pub fn move_link(
        &mut self,
        from: impl Into<Position>,
        to: impl Into<Position>,
    ) -> CoreResult<()> {
        let from = from.into().get();
        let to = to.into().get();
        let length = self.link_count();
        if from >= length || to >= length {
            return Err(invalid_core("DDE link move position did not match"));
        }
        if from != to {
            let links = self.ensure_links()?;
            let link = links.remove(from);
            links.insert(to, link);
        }
        Ok(())
    }

    /// Replace a uniquely named formula link.
    pub fn replace_link_named(&mut self, name: &str, link: LinkSpec) -> CoreResult<()> {
        let index = self.unique_named_link(name)?;
        self.replace_link_at(index, link)
    }

    /// Remove a uniquely named formula link.
    pub fn remove_link_named(&mut self, name: &str) -> CoreResult<()> {
        let index = self.unique_named_link(name)?;
        self.remove_link_at(index)
    }

    /// Whether the edit retains the exact source bytes and owner order.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        let links_noop = self.links.as_ref().is_none_or(|links| {
            links.len() == self.before.links().len()
                && links.iter().enumerate().all(|(index, link)| match link {
                    LinkDraft::Existing(existing) => *existing == index,
                    // Source-only drafts are normalized to Existing when the
                    // source is equal.  Keep this conservative for any future
                    // draft producer without comparing unbounded strings in a
                    // non-fallible query.
                    LinkDraft::Source { .. } => false,
                    // An authored replacement is deliberately conservative here.  Determining
                    // whether its generated bytes equal the existing owner requires a fallible
                    // scan/render operation and may consume the retained context budget.  The
                    // commit path performs that work once and compares the complete candidate
                    // XML before publishing, so an exact authored replacement still commits as
                    // an unchanged transaction without making this query fallible-by-proxy.
                    LinkDraft::Owned { .. } => false,
                })
        });
        // Once worksheet drafts are initialized, defer their exact-byte
        // no-op determination to commit's fallible render and comparison.
        links_noop && self.sheet_sources.is_none()
    }

    /// Validate and publish the edit under a retained execution context.
    pub fn commit(&mut self, context: &ExecutionContext) -> CoreResult<Commit> {
        self.before.context().check().map_err(map_execution_core)?;
        context.check().map_err(map_execution_core)?;
        if self.before.enforce_context_lineage()
            && !same_budget_lineage(self.before.context(), context)?
        {
            return Err(invalid_core(
                "DDE edit context does not belong to the retained snapshot budget",
            ));
        }
        if self.is_noop() {
            let source = Arc::new(self.before.clone());
            return Ok(Commit {
                snapshot: self.before.clone(),
                patch: Patch {
                    source: Arc::clone(&source),
                    target: source,
                },
                changed: false,
            });
        }
        let rendered = render_source(&self.before, self, context)?;
        if rendered.xml == self.before.content.as_ref() {
            drop(rendered);
            let source = Arc::new(self.before.clone());
            return Ok(Commit {
                snapshot: self.before.clone(),
                patch: Patch {
                    source: Arc::clone(&source),
                    target: source,
                },
                changed: false,
            });
        }
        let mut target = Snapshot::parse_with_context(&rendered.xml, self.before.limits(), context)
            .map_err(CoreError::from)?;
        target.enforce_context_lineage = self.before.enforce_context_lineage();
        target._output_reservation = Some(Arc::clone(&rendered.output_reservation));
        verify_readback(self, &target, &rendered.cache_names)?;
        // The candidate, parser projection, and readback renderer have all
        // finished using the transient source-splice allocations.
        drop(rendered);
        let source = Arc::new(self.before.clone());
        let target_source = Arc::new(target.clone());
        Ok(Commit {
            snapshot: target,
            patch: Patch {
                source,
                target: target_source,
            },
            changed: true,
        })
    }

    fn resolve_sheet(&self, selector: SheetSelector<'_>) -> CoreResult<usize> {
        match selector {
            SheetSelector::Position(position) => {
                let index = position.get();
                if index >= self.before.inventory().table_names.len() {
                    return Err(invalid_core("DDE worksheet position did not match"));
                }
                Ok(index)
            },
            SheetSelector::Name(name) => {
                let mut found = None;
                for (index, table_name) in self.before.inventory().table_names.iter().enumerate() {
                    self.before
                        .context()
                        .consume(Resource::Work, 1)
                        .map_err(map_execution_core)?;
                    if table_name.as_deref() == Some(name.as_ref()) {
                        if found.replace(index).is_some() {
                            return Err(invalid_core(
                                "DDE worksheet name is ambiguous; use position",
                            ));
                        }
                    }
                }
                found.ok_or_else(|| invalid_core("DDE worksheet name did not match"))
            },
        }
    }

    fn unique_named_link(&self, name: &str) -> CoreResult<usize> {
        let mut found = None;
        let links = self.links.as_deref().unwrap_or(&[]);
        let link_count = if self.links.is_some() {
            links.len()
        } else {
            self.before.links().len()
        };
        for index in 0..link_count {
            self.before
                .context()
                .consume(Resource::Work, 1)
                .map_err(map_execution_core)?;
            let draft = self.links.as_ref().and_then(|drafts| drafts.get(index));
            let default = LinkDraft::Existing(index);
            let draft = draft.unwrap_or(&default);
            let source = match draft {
                LinkDraft::Existing(existing) => {
                    self.before.links().get(*existing).map(Link::source)
                },
                LinkDraft::Source { source, .. } => Some(source),
                LinkDraft::Owned { link, .. } => Some(link.source()),
            };
            if source.and_then(Source::name) == Some(name) {
                if found.replace(index).is_some() {
                    return Err(invalid_core("DDE link name is ambiguous; use position"));
                }
            }
        }
        found.ok_or_else(|| invalid_core("DDE link name did not match"))
    }

    fn desired_sheet_source(&self, table_index: usize) -> Option<&Source> {
        if let Some(drafts) = &self.sheet_sources {
            drafts
                .iter()
                .find(|entry| entry.table_index == table_index)
                .map(|entry| &entry.source)
        } else {
            let sheet_name = self
                .before
                .inventory()
                .table_names
                .get(table_index)
                .and_then(Option::as_deref)?;
            self.before
                .sheet_sources()
                .iter()
                .find(|entry| entry.sheet() == sheet_name)
                .map(SheetSource::source)
        }
    }
}

/// A complete DDE transaction result.
#[derive(Clone, Debug)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
    changed: bool,
}

impl Commit {
    /// Whether source bytes changed.
    #[must_use]
    pub const fn changed(&self) -> bool {
        self.changed
    }

    /// Borrow the typed post-commit snapshot.
    #[must_use]
    pub fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Borrow the exact reversible patch.
    #[must_use]
    pub fn patch(&self) -> &Patch {
        &self.patch
    }
}

/// An exact-source, reversible `content.xml` DDE patch.
#[derive(Clone, Debug)]
pub struct Patch {
    source: Arc<Snapshot>,
    target: Arc<Snapshot>,
}

impl PartialEq for Patch {
    fn eq(&self, other: &Self) -> bool {
        self.source_xml() == other.source_xml() && self.target_xml() == other.target_xml()
    }
}

impl Eq for Patch {}

impl Patch {
    /// Whether the patch preserves source bytes exactly.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        // Commits that preserve bytes intentionally retain one Arc for both
        // endpoints.  Use that O(1) invariant instead of scanning the whole
        // content XML from a non-fallible query.
        Arc::ptr_eq(&self.source, &self.target)
    }

    /// Whether the patch changes source bytes.
    #[must_use]
    pub fn changed(&self) -> bool {
        !self.is_empty()
    }

    /// Return the exact source XML this patch expects.
    #[must_use]
    pub fn source_xml(&self) -> &str {
        self.source.source_xml()
    }

    /// Return the exact target XML this patch publishes.
    #[must_use]
    pub fn target_xml(&self) -> &str {
        self.target.source_xml()
    }

    /// Return the exact inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            source: Arc::clone(&self.target),
            target: Arc::clone(&self.source),
        }
    }

    /// Apply the patch to an exact source snapshot under its destination
    /// limits and retained context.
    pub fn apply(&self, snapshot: &Snapshot) -> CoreResult<Commit> {
        self.source.context().check().map_err(map_execution_core)?;
        snapshot.context().check().map_err(map_execution_core)?;
        if !compare_xml_bytes(
            snapshot.context(),
            snapshot.content.as_ref(),
            self.source.content.as_ref(),
        )? {
            return Err(invalid_core("ODS DDE patch source snapshot does not match"));
        }
        if self.is_empty() {
            return Ok(Commit {
                snapshot: snapshot.clone(),
                patch: self.clone(),
                changed: false,
            });
        }
        if self.target.content.len() > snapshot.limits().output_bytes() {
            return Err(resource_core(
                "output bytes",
                self.target.content.len(),
                snapshot.limits().output_bytes(),
            ));
        }
        let output_reservation = snapshot
            .context()
            .reserve(Resource::OutputBytes, self.target.content.len() as u64)
            .map_err(map_execution_core)
            .map(Arc::new)?;
        let mut target = Snapshot::parse_shared_with_context(
            Arc::clone(&self.target.content),
            snapshot.limits(),
            snapshot.context(),
            snapshot.enforce_context_lineage(),
        )
        .map_err(CoreError::from)?;
        target.enforce_context_lineage = snapshot.enforce_context_lineage();
        target._output_reservation = Some(output_reservation);
        Ok(Commit {
            snapshot: target,
            patch: self.clone(),
            changed: true,
        })
    }
}

#[derive(Clone, Debug)]
struct Replacement {
    range: Range<usize>,
    text: String,
}

#[derive(Debug)]
struct Scan {
    spreadsheet: Option<ElementSite>,
    /// Byte offset at which the spreadsheet closing tag begins. New DDE link
    /// owners belong in the spreadsheet epilogue, after worksheet tables.
    spreadsheet_close_start: Option<usize>,
    tables: Vec<TableSite>,
    links_container: Option<ElementSite>,
    links_container_opaque: bool,
    links: Vec<LinkSite>,
    /// Retained admission for scanner-owned vectors and semantic projections.
    /// The reservation grows before each heap allocation and lives until the
    /// scan is no longer needed by the source splice.
    _memory_reservation: Reservation,
    /// A content root containing a markup-compatibility choice is refused for
    /// mutation.  Its branches may carry semantically equivalent owners, but
    /// a byte-range transaction cannot safely select the branch a consumer
    /// will evaluate without interpreting MCE policy.
    mce_present: bool,
    mce_depth: usize,
}

impl Scan {
    fn new(context: &ExecutionContext) -> CoreResult<Self> {
        let reservation = context
            .reserve(Resource::Memory, 1)
            .map_err(map_execution_core)?;
        Ok(Self {
            spreadsheet: None,
            spreadsheet_close_start: None,
            tables: Vec::new(),
            links_container: None,
            links_container_opaque: false,
            links: Vec::new(),
            _memory_reservation: reservation,
            mce_present: false,
            mce_depth: 0,
        })
    }

    fn reserve_memory(&mut self, context: &ExecutionContext, amount: usize) -> CoreResult<()> {
        if amount == 0 {
            return Ok(());
        }
        let amount = u64::try_from(amount)
            .map_err(|_| invalid_core("DDE scanner memory size exceeds u64"))?;
        let reservation = context
            .reserve(Resource::Memory, amount)
            .map_err(map_execution_core)?;
        self._memory_reservation
            .try_merge(reservation)
            .map_err(|_| invalid_core("DDE scanner memory reservations have different owners"))
    }

    fn reserve_vec_slot<T>(
        &mut self,
        length: usize,
        capacity: usize,
        context: &ExecutionContext,
    ) -> CoreResult<()> {
        if length == capacity {
            self.reserve_memory(context, size_of::<T>())?;
        }
        Ok(())
    }
}

#[derive(Clone, Debug)]
struct ElementSite {
    range: Range<usize>,
    opaque: bool,
}

#[derive(Clone, Debug)]
struct TableSite {
    element: ElementSite,
    name: Option<String>,
    source: Option<ElementSite>,
    source_value: Option<Source>,
    source_insertion: Option<usize>,
    paired: bool,
}

#[derive(Clone, Debug)]
struct LinkSite {
    element: ElementSite,
    source: Option<ElementSite>,
    source_value: Option<Source>,
    cache: Option<ElementSite>,
    /// The standard table name on a scalar cache owner, if present.  Typed
    /// replacement preserves this modeled owner attribute while refusing all
    /// other unmodeled cache-table attributes.
    cache_name: Option<String>,
    link_opaque: bool,
    cache_opaque: bool,
}

#[derive(Clone, Copy, Debug)]
enum FrameKind {
    Other,
    DocumentContent,
    Body,
    Mce,
    Spreadsheet,
    Table(usize),
    Links,
    Link(usize),
    SourceTable(usize),
    SourceLink(usize),
    Cache(usize),
    CacheContent(usize),
}

#[derive(Clone, Copy, Debug)]
struct Frame {
    kind: FrameKind,
    start: usize,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum NamespaceTag {
    Office,
    Table,
    Text,
    MarkupCompatibility,
    Other,
}

fn render_scratch_estimate(
    before: &Snapshot,
    edit: &Edit,
    scan: &Scan,
    context: &ExecutionContext,
) -> CoreResult<usize> {
    let replacement_capacity = scan
        .tables
        .len()
        .checked_add(scan.links.len())
        .and_then(|value| value.checked_add(2))
        .ok_or_else(|| invalid_core("DDE replacement count overflows"))?;
    let replacement_bytes = replacement_capacity
        .checked_mul(size_of::<Replacement>())
        .ok_or_else(|| invalid_core("DDE render scratch size overflows"))?;
    // The scanner's semantic cache-name projections are moved into the
    // readback ledger after source splicing, so admit both their vector slots
    // and retained string payloads while the target is parsed and checked.
    let cache_name_slots = scan
        .links
        .len()
        .checked_mul(size_of::<Option<String>>())
        .ok_or_else(|| invalid_core("DDE render scratch size overflows"))?;
    let mut estimate = before
        .content
        .len()
        .checked_add(8 * 1024)
        .and_then(|value| value.checked_add(replacement_bytes))
        .and_then(|value| value.checked_add(cache_name_slots))
        .ok_or_else(|| invalid_core("DDE render scratch size overflows"))?;
    for site in &scan.links {
        if let Some(name) = &site.cache_name {
            estimate = estimate
                .checked_add(name.len())
                .ok_or_else(|| invalid_core("DDE render scratch size overflows"))?;
        }
    }

    if let Some(sources) = &edit.sheet_sources {
        for source in sources {
            context.check().map_err(map_execution_core)?;
            context
                .consume(Resource::Work, 1)
                .map_err(map_execution_core)?;
            let source_estimate = source_render_estimate(&source.source)?;
            estimate = estimate
                // The replacement owns the rendered declaration while the
                // renderer's temporary declaration is being appended.
                .checked_add(
                    source_estimate
                        .checked_mul(2)
                        .ok_or_else(|| invalid_core("DDE render scratch size overflows"))?,
                )
                .ok_or_else(|| invalid_core("DDE render scratch size overflows"))?;
        }
    }
    if let Some(links) = &edit.links {
        // `Edit::is_noop` has already returned false before this renderer is
        // called, so an initialized link draft set represents either a real
        // link change or a composition with another staged change.  Avoid an
        // unmetered identity scan here; the source-only classifier below is
        // charged explicitly.
        let links_changed = true;
        let source_only = drafts_are_source_only(links, scan, context)?;
        let mut body_estimate = 0usize;
        for draft in links {
            context.check().map_err(map_execution_core)?;
            context
                .consume(Resource::Work, 1)
                .map_err(map_execution_core)?;
            let draft_estimate = match draft {
                LinkDraft::Owned { link, replaced, .. } => {
                    let source = source_render_estimate(link.source())?;
                    let cache_name = match replaced {
                        Some(index) => scan
                            .links
                            .get(*index)
                            .ok_or_else(|| invalid_core("DDE link source order changed"))?
                            .cache_name
                            .as_deref(),
                        None => None,
                    };
                    let cache = cached_table_render_estimate_with_name(
                        link.cached_table(),
                        cache_name,
                        context,
                    )?;
                    let link = 128usize
                        .checked_add(source)
                        .and_then(|value| value.checked_add(cache))
                        .ok_or_else(|| invalid_core("DDE render scratch size overflows"))?;
                    // `render_link` retains its destination buffer while its
                    // source and cache temporary strings are pushed into it.
                    let temporary = source
                        .checked_add(cache)
                        .ok_or_else(|| invalid_core("DDE render scratch size overflows"))?;
                    estimate = estimate
                        .checked_add(link)
                        .and_then(|value| value.checked_add(temporary))
                        .ok_or_else(|| invalid_core("DDE render scratch size overflows"))?;
                    link
                },
                LinkDraft::Source {
                    existing, source, ..
                } => {
                    let owner = scan
                        .links
                        .get(*existing)
                        .map(|site| {
                            site.element
                                .range
                                .end
                                .saturating_sub(site.element.range.start)
                        })
                        .ok_or_else(|| invalid_core("DDE link source order changed"))?;
                    let source = source_render_estimate(source)?;
                    estimate = estimate
                        .checked_add(owner)
                        .and_then(|value| value.checked_add(source))
                        .and_then(|value| value.checked_add(source))
                        .ok_or_else(|| invalid_core("DDE render scratch size overflows"))?;
                    owner
                        .checked_add(source)
                        .ok_or_else(|| invalid_core("DDE render scratch size overflows"))?
                },
                LinkDraft::Existing(index) => scan
                    .links
                    .get(*index)
                    .map(|site| {
                        site.element
                            .range
                            .end
                            .saturating_sub(site.element.range.start)
                    })
                    .ok_or_else(|| invalid_core("DDE link source order changed"))?,
            };
            body_estimate = body_estimate
                .checked_add(draft_estimate)
                .ok_or_else(|| invalid_core("DDE render scratch size overflows"))?;
        }
        if links_changed && !source_only {
            // Structural composition keeps both the body and the enclosing
            // replacement alive before source splicing.  Account for that
            // simultaneous peak, including its fixed owner wrapper.  When
            // the owner is absent, the newly inserted container remains live
            // while each nested link is rendered, so reserve its body and
            // wrapper separately as well.
            let enclosing = body_estimate
                .checked_add(128)
                .ok_or_else(|| invalid_core("DDE render scratch size overflows"))?;
            let peak = if scan.links_container.is_some() {
                body_estimate
                    .checked_mul(2)
                    .and_then(|value| value.checked_add(128))
                    .ok_or_else(|| invalid_core("DDE render scratch size overflows"))?
            } else if !links.is_empty() {
                enclosing
            } else {
                0
            };
            estimate = estimate
                .checked_add(peak)
                .ok_or_else(|| invalid_core("DDE render scratch size overflows"))?;
        }
    }
    Ok(estimate)
}

// Rust drops struct fields in declaration order. Keep transient data ahead of
// its reservation tokens on every exit path, including failed readback.
struct RenderedCandidate {
    xml: String,
    cache_names: Vec<Option<String>>,
    output_reservation: Arc<Reservation>,
    _scratch_reservation: Arc<Reservation>,
}

fn render_source(
    before: &Snapshot,
    edit: &Edit,
    context: &ExecutionContext,
) -> CoreResult<RenderedCandidate> {
    context.check().map_err(map_execution_core)?;
    let scan = scan_source(before.content.as_ref(), before.limits(), context)?;
    if scan.mce_present {
        return Err(invalid_core(
            "DDE mutation is refused when markup-compatibility choices are present",
        ));
    }
    if scan.tables.len() != before.inventory().table_names.len()
        || scan.links.len() != before.links().len()
    {
        return Err(invalid_core("DDE source inventory changed during edit"));
    }
    for (index, table) in scan.tables.iter().enumerate() {
        if table.name.as_ref() != before.inventory().table_names[index].as_ref() {
            return Err(invalid_core(
                "DDE worksheet table order changed during edit",
            ));
        }
    }
    // Reserve scratch capacity before constructing replacement vectors or
    // generated owner strings.  The output reservation below remains the
    // authoritative candidate-byte admission; this reservation covers the
    // transient source-splice strings while they are built.
    let scratch_amount = render_scratch_estimate(before, edit, &scan, context)?;
    let scratch_amount = u64::try_from(scratch_amount)
        .map_err(|_| invalid_core("DDE render scratch exceeds u64"))?;
    let scratch_reservation = context
        .reserve(Resource::Memory, scratch_amount)
        .map_err(map_execution_core)
        .map(Arc::new)?;
    let mut replacements = Vec::new();
    let replacement_capacity = scan
        .tables
        .len()
        .checked_add(scan.links.len())
        .and_then(|value| value.checked_add(2))
        .ok_or_else(|| invalid_core("DDE replacement count overflows"))?;
    replacements
        .try_reserve_exact(replacement_capacity)
        .map_err(|_| CoreError::Unsupported("DDE replacement allocation failed".to_string()))?;
    render_sheet_sources(before, edit, &scan, &mut replacements, context)?;
    render_links(before, edit, &scan, &mut replacements, context)?;
    replacements.sort_by(|left, right| {
        left.range
            .start
            .cmp(&right.range.start)
            .then_with(|| left.range.end.cmp(&right.range.end))
    });
    let mut target_len = before.content.len();
    for replacement in &replacements {
        if replacement.range.start > replacement.range.end
            || replacement.range.end > before.content.len()
        {
            return Err(invalid_core("DDE replacement range is outside source"));
        }
        let removed = replacement.range.end - replacement.range.start;
        target_len = target_len
            .checked_sub(removed)
            .and_then(|value| value.checked_add(replacement.text.len()))
            .ok_or_else(|| invalid_core("DDE output size overflows"))?;
    }
    for pair in replacements.windows(2) {
        if pair[0].range.end > pair[1].range.start {
            return Err(invalid_core("DDE replacement ranges overlap"));
        }
    }
    if target_len > before.limits().output_bytes() {
        return Err(resource_core(
            "output bytes",
            target_len,
            before.limits().output_bytes(),
        ));
    }
    let output_reservation = context
        .reserve(Resource::OutputBytes, target_len as u64)
        .map_err(map_execution_core)
        .map(Arc::new)?;
    let mut output = String::new();
    output
        .try_reserve_exact(target_len)
        .map_err(|_| CoreError::Unsupported("DDE output allocation failed".to_string()))?;
    let mut cursor = 0usize;
    for replacement in replacements {
        context.check().map_err(map_execution_core)?;
        context
            .consume(Resource::Work, 1)
            .map_err(map_execution_core)?;
        output.push_str(&before.content[cursor..replacement.range.start]);
        output.push_str(&replacement.text);
        cursor = replacement.range.end;
    }
    output.push_str(&before.content[cursor..]);
    debug_assert_eq!(output.len(), target_len);
    let mut cache_names = Vec::new();
    cache_names
        .try_reserve_exact(scan.links.len())
        .map_err(|_| CoreError::Unsupported("DDE cache-name allocation failed".to_string()))?;
    for site in scan.links {
        cache_names.push(site.cache_name);
    }
    Ok(RenderedCandidate {
        xml: output,
        cache_names,
        output_reservation,
        _scratch_reservation: scratch_reservation,
    })
}

fn render_sheet_sources(
    before: &Snapshot,
    edit: &Edit,
    scan: &Scan,
    replacements: &mut Vec<Replacement>,
    context: &ExecutionContext,
) -> CoreResult<()> {
    for (table_index, table) in scan.tables.iter().enumerate() {
        context.check().map_err(map_execution_core)?;
        context
            .consume(Resource::Work, 1)
            .map_err(map_execution_core)?;
        let desired = edit.desired_sheet_source(table_index);
        match (&table.source, desired) {
            (Some(existing), Some(source)) => {
                let current = table
                    .source_value
                    .as_ref()
                    .ok_or_else(|| invalid_core("DDE sheet source readback is unavailable"))?;
                if !compare_sources(context, current, source)? {
                    if existing.opaque {
                        return Err(invalid_core(
                            "DDE sheet source contains unknown markup or attributes",
                        ));
                    }
                    replacements.push(Replacement {
                        range: existing.range.clone(),
                        text: render_source_declaration(source)?,
                    });
                }
            },
            (Some(existing), None) => {
                if existing.opaque {
                    return Err(invalid_core(
                        "DDE sheet source contains unknown markup or attributes",
                    ));
                }
                replacements.push(Replacement {
                    range: existing.range.clone(),
                    text: String::new(),
                });
            },
            (None, Some(source)) => {
                let insertion = table.source_insertion.ok_or_else(|| {
                    invalid_core("DDE sheet source insertion point is unavailable")
                })?;
                if !table.paired {
                    return Err(invalid_core(
                        "adding a DDE source to a self-closing table is refused",
                    ));
                }
                replacements.push(Replacement {
                    range: insertion..insertion,
                    text: render_source_declaration(source)?,
                });
            },
            (None, None) => {},
        }
    }
    let _ = before;
    Ok(())
}

fn drafts_are_source_only(
    desired_links: &[LinkDraft],
    scan: &Scan,
    context: &ExecutionContext,
) -> CoreResult<bool> {
    let mut source_only = desired_links.len() == scan.links.len();
    for (index, draft) in desired_links.iter().enumerate() {
        context.check().map_err(map_execution_core)?;
        context
            .consume(Resource::Work, 1)
            .map_err(map_execution_core)?;
        if !matches!(draft, LinkDraft::Existing(existing) if *existing == index)
            && !matches!(draft, LinkDraft::Source { existing, .. } if *existing == index)
        {
            source_only = false;
        }
    }
    Ok(source_only)
}

fn render_links(
    before: &Snapshot,
    edit: &Edit,
    scan: &Scan,
    replacements: &mut Vec<Replacement>,
    context: &ExecutionContext,
) -> CoreResult<()> {
    let Some(desired_links) = edit.links.as_deref() else {
        return Ok(());
    };
    // An initialized draft set reaches this function only when another
    // staged operation already made the edit non-noop.  Identity is checked
    // by the charged source-only classifier instead of an unmetered scan.
    let source_only = drafts_are_source_only(desired_links, scan, context)?;
    if let Some(container) = &scan.links_container {
        if source_only {
            for (index, draft) in desired_links.iter().enumerate() {
                context.check().map_err(map_execution_core)?;
                context
                    .consume(Resource::Work, 1)
                    .map_err(map_execution_core)?;
                let LinkDraft::Source { source, .. } = draft else {
                    continue;
                };
                let site = scan
                    .links
                    .get(index)
                    .ok_or_else(|| invalid_core("DDE link source order changed"))?;
                let source_site = site
                    .source
                    .as_ref()
                    .ok_or_else(|| invalid_core("DDE link source is missing"))?;
                if source_site.opaque {
                    return Err(invalid_core(
                        "DDE link source contains unknown markup or attributes",
                    ));
                }
                replacements.push(Replacement {
                    range: source_site.range.clone(),
                    text: render_source_declaration(source)?,
                });
            }
            return Ok(());
        }
        if scan.links_container_opaque {
            return Err(invalid_core(
                "DDE link container contains unknown direct markup; reorder/add/remove is refused",
            ));
        }
        for draft in desired_links {
            if let LinkDraft::Owned { replaced, .. } = draft
                && replaced
                    .as_ref()
                    .and_then(|index| scan.links.get(*index))
                    .is_some_and(|site| site.link_opaque || site.cache_opaque)
            {
                return Err(invalid_core(
                    "DDE link contains unknown markup and cannot be replaced as typed data",
                ));
            }
        }
        let body_estimate = desired_links.iter().try_fold(0usize, |sum, draft| {
            context.check().map_err(map_execution_core)?;
            context
                .consume(Resource::Work, 1)
                .map_err(map_execution_core)?;
            let estimate = match draft {
                LinkDraft::Existing(index) => scan
                    .links
                    .get(*index)
                    .map(|site| {
                        site.element
                            .range
                            .end
                            .saturating_sub(site.element.range.start)
                    })
                    .ok_or_else(|| invalid_core("DDE link source order changed"))?,
                LinkDraft::Owned { link, replaced, .. } => {
                    let cache_name = match replaced {
                        Some(index) => scan
                            .links
                            .get(*index)
                            .ok_or_else(|| invalid_core("DDE link source order changed"))?
                            .cache_name
                            .as_deref(),
                        None => None,
                    };
                    link_render_estimate_with_cache_name(link, cache_name, context)?
                },
                LinkDraft::Source {
                    existing, source, ..
                } => {
                    let site = scan
                        .links
                        .get(*existing)
                        .ok_or_else(|| invalid_core("DDE link source order changed"))?;
                    site.element
                        .range
                        .end
                        .saturating_sub(site.element.range.start)
                        .checked_add(source_render_estimate(source)?)
                        .ok_or_else(|| invalid_core("DDE link render size overflows"))?
                },
            };
            sum.checked_add(estimate)
                .ok_or_else(|| invalid_core("DDE link body size overflows"))
        })?;
        let mut body = String::new();
        body.try_reserve(body_estimate)
            .map_err(|_| CoreError::Unsupported("DDE link body allocation failed".to_string()))?;
        for draft in desired_links {
            context.check().map_err(map_execution_core)?;
            context
                .consume(Resource::Work, 1)
                .map_err(map_execution_core)?;
            match draft {
                LinkDraft::Existing(index) => {
                    let site = scan
                        .links
                        .get(*index)
                        .ok_or_else(|| invalid_core("DDE link source order changed"))?;
                    body.push_str(&before.content[site.element.range.clone()]);
                },
                LinkDraft::Owned { link, replaced, .. } => {
                    let cache_name = match replaced {
                        Some(index) => scan
                            .links
                            .get(*index)
                            .ok_or_else(|| invalid_core("DDE link source order changed"))?
                            .cache_name
                            .as_deref(),
                        None => None,
                    };
                    body.push_str(&render_link_with_cache_name(link, cache_name, context)?);
                },
                LinkDraft::Source {
                    existing, source, ..
                } => body.push_str(&render_existing_link_with_source(
                    before,
                    scan.links
                        .get(*existing)
                        .ok_or_else(|| invalid_core("DDE link source order changed"))?,
                    source,
                    context,
                )?),
            }
        }
        let replacement_estimate = body_estimate
            .checked_add(128)
            .ok_or_else(|| invalid_core("DDE link replacement size overflows"))?;
        let mut replacement = String::new();
        replacement.try_reserve(replacement_estimate).map_err(|_| {
            CoreError::Unsupported("DDE link replacement allocation failed".to_string())
        })?;
        replacement.push_str("<table:dde-links xmlns:table=\"");
        replacement.push_str(std::str::from_utf8(TABLE).unwrap_or(""));
        replacement.push_str("\">");
        replacement.push_str(&body);
        replacement.push_str("</table:dde-links>");
        replacements.push(Replacement {
            range: container.range.clone(),
            text: if desired_links.is_empty() {
                String::new()
            } else {
                replacement
            },
        });
    } else if !desired_links.is_empty() {
        let _spreadsheet = scan
            .spreadsheet
            .as_ref()
            .ok_or_else(|| invalid_core("DDE spreadsheet insertion point is unavailable"))?;
        let insertion = scan
            .spreadsheet_close_start
            .ok_or_else(|| invalid_core("DDE spreadsheet closing point is unavailable"))?;
        let container_estimate = desired_links.iter().try_fold(128usize, |sum, draft| {
            context.check().map_err(map_execution_core)?;
            context
                .consume(Resource::Work, 1)
                .map_err(map_execution_core)?;
            let estimate = match draft {
                LinkDraft::Owned { link, replaced, .. } => {
                    let cache_name = match replaced {
                        Some(index) => scan
                            .links
                            .get(*index)
                            .ok_or_else(|| invalid_core("DDE link source order changed"))?
                            .cache_name
                            .as_deref(),
                        None => None,
                    };
                    link_render_estimate_with_cache_name(link, cache_name, context)?
                },
                LinkDraft::Existing(_) | LinkDraft::Source { .. } => {
                    return Err(invalid_core(
                        "DDE link source order changed without a container",
                    ));
                },
            };
            sum.checked_add(estimate)
                .ok_or_else(|| invalid_core("DDE link container size overflows"))
        })?;
        let mut container = String::new();
        container.try_reserve(container_estimate).map_err(|_| {
            CoreError::Unsupported("DDE link container allocation failed".to_string())
        })?;
        container.push_str("<table:dde-links xmlns:table=\"");
        container.push_str(std::str::from_utf8(TABLE).unwrap_or(""));
        container.push_str("\">");
        for draft in desired_links {
            context.check().map_err(map_execution_core)?;
            context
                .consume(Resource::Work, 1)
                .map_err(map_execution_core)?;
            match draft {
                LinkDraft::Existing(_) => {
                    return Err(invalid_core(
                        "DDE link source order changed without a container",
                    ));
                },
                LinkDraft::Owned { link, replaced, .. } => {
                    let cache_name = match replaced {
                        Some(index) => scan
                            .links
                            .get(*index)
                            .ok_or_else(|| invalid_core("DDE link source order changed"))?
                            .cache_name
                            .as_deref(),
                        None => None,
                    };
                    container.push_str(&render_link_with_cache_name(link, cache_name, context)?);
                },
                LinkDraft::Source { .. } => {
                    return Err(invalid_core(
                        "DDE link source order changed without a container",
                    ));
                },
            }
        }
        container.push_str("</table:dde-links>");
        replacements.push(Replacement {
            range: insertion..insertion,
            text: container,
        });
    }
    Ok(())
}

fn render_existing_link_with_source(
    before: &Snapshot,
    site: &LinkSite,
    source: &Source,
    context: &ExecutionContext,
) -> CoreResult<String> {
    let owner = site.element.range.clone();
    let source_site = site
        .source
        .as_ref()
        .ok_or_else(|| invalid_core("DDE link source is missing"))?;
    if source_site.opaque {
        return Err(invalid_core(
            "DDE link source contains unknown markup or attributes",
        ));
    }
    if source_site.range.start < owner.start
        || source_site.range.end > owner.end
        || source_site.range.start > source_site.range.end
    {
        return Err(invalid_core("DDE link source range is outside its owner"));
    }
    let replacement_start = source_site.range.start - owner.start;
    let replacement_end = source_site.range.end - owner.start;
    let owner_xml = &before.content[owner.clone()];
    let mut output = String::new();
    let capacity = owner_xml
        .len()
        .checked_add(source_render_estimate(source)?)
        .ok_or_else(|| invalid_core("DDE link render size overflows"))?;
    output.try_reserve(capacity).map_err(|_| {
        CoreError::Unsupported("DDE source-only link allocation failed".to_string())
    })?;
    context.check().map_err(map_execution_core)?;
    context
        .consume(Resource::Work, 1)
        .map_err(map_execution_core)?;
    output.push_str(&owner_xml[..replacement_start]);
    output.push_str(&render_source_declaration(source)?);
    output.push_str(&owner_xml[replacement_end..]);
    Ok(output)
}

fn source_render_estimate(source: &Source) -> CoreResult<usize> {
    let mut estimate = 192usize;
    for value in [source.application(), source.topic(), source.item()] {
        estimate = estimate
            .checked_add(escaped_len_upper_bound(value.len())?)
            .ok_or_else(|| invalid_core("DDE source render size overflows"))?;
    }
    if let Some(name) = source.name() {
        estimate = estimate
            .checked_add(escaped_len_upper_bound(name.len())?)
            .ok_or_else(|| invalid_core("DDE source render size overflows"))?;
    }
    Ok(estimate)
}

fn link_render_estimate_with_cache_name(
    link: &LinkSpec,
    cache_name: Option<&str>,
    context: &ExecutionContext,
) -> CoreResult<usize> {
    let source = source_render_estimate(&link.source)?;
    let cache = cached_table_render_estimate_with_name(&link.cached_table, cache_name, context)?;
    128usize
        .checked_add(source)
        .and_then(|value| value.checked_add(cache))
        .ok_or_else(|| invalid_core("DDE link render size overflows"))
}

fn cached_table_render_estimate_with_name(
    table: &CachedTable,
    cache_name: Option<&str>,
    context: &ExecutionContext,
) -> CoreResult<usize> {
    let mut estimate = 512usize;
    if let Some(name) = cache_name {
        estimate = estimate
            .checked_add(16)
            .and_then(|value| value.checked_add(escaped_len_upper_bound(name.len()).ok()?))
            .ok_or_else(|| invalid_core("DDE cache render size overflows"))?;
    }
    for row in &table.rows {
        context.check().map_err(map_execution_core)?;
        context
            .consume(Resource::Work, 1)
            .map_err(map_execution_core)?;
        estimate = estimate
            .checked_add(96)
            .ok_or_else(|| invalid_core("DDE cache render size overflows"))?;
        if row.cells.is_empty() {
            estimate = estimate
                .checked_add(64)
                .ok_or_else(|| invalid_core("DDE cache render size overflows"))?;
        } else {
            for cell in &row.cells {
                context.check().map_err(map_execution_core)?;
                context
                    .consume(Resource::Work, 1)
                    .map_err(map_execution_core)?;
                estimate = estimate
                    .checked_add(cached_cell_render_estimate(cell.value())?)
                    .ok_or_else(|| invalid_core("DDE cache render size overflows"))?;
            }
        }
    }
    if table.rows.is_empty() {
        estimate = estimate
            .checked_add(96 + 64)
            .ok_or_else(|| invalid_core("DDE cache render size overflows"))?;
    }
    Ok(estimate)
}

fn cached_cell_render_estimate(cell: &CachedValue) -> CoreResult<usize> {
    let estimate = match cell {
        CachedValue::Empty => 64,
        CachedValue::Text(value) => 128usize
            .checked_add(escaped_len_upper_bound(value.len())?)
            .ok_or_else(|| invalid_core("DDE cache cell render size overflows"))?,
        CachedValue::Number(_) | CachedValue::Percentage(_) => 128,
        CachedValue::Currency { currency, .. } => 160usize
            .checked_add(escaped_len_upper_bound(currency.len())?)
            .ok_or_else(|| invalid_core("DDE cache cell render size overflows"))?,
        CachedValue::Boolean(_) => 144,
        CachedValue::Date(value) | CachedValue::Time(value) => 160usize
            .checked_add(escaped_len_upper_bound(value.len())?)
            .ok_or_else(|| invalid_core("DDE cache cell render size overflows"))?,
    };
    Ok(estimate)
}

fn escaped_len_upper_bound(length: usize) -> CoreResult<usize> {
    length
        .checked_mul(6)
        .ok_or_else(|| invalid_core("DDE XML escaped value size overflows"))
}

/// Append an XML attribute value without allowing XML's attribute whitespace
/// normalization to change a checked source or cached scalar.  Literal tabs,
/// line feeds, and carriage returns therefore use character references.
fn push_xml_attribute(output: &mut String, value: &str) {
    for character in value.chars() {
        match character {
            '&' => output.push_str("&amp;"),
            '<' => output.push_str("&lt;"),
            '>' => output.push_str("&gt;"),
            '"' => output.push_str("&quot;"),
            '\t' => output.push_str("&#9;"),
            '\n' => output.push_str("&#10;"),
            '\r' => output.push_str("&#13;"),
            _ => output.push(character),
        }
    }
}

fn render_source_declaration(source: &Source) -> CoreResult<String> {
    let mut output = String::new();
    output
        .try_reserve(source_render_estimate(source)?)
        .map_err(|_| CoreError::Unsupported("DDE source allocation failed".to_string()))?;
    output.push_str("<office:dde-source xmlns:office=\"");
    output.push_str(std::str::from_utf8(OFFICE).unwrap_or(""));
    output.push_str("\" office:dde-application=\"");
    push_xml_attribute(&mut output, source.application());
    output.push_str("\" office:dde-topic=\"");
    push_xml_attribute(&mut output, source.topic());
    output.push_str("\" office:dde-item=\"");
    push_xml_attribute(&mut output, source.item());
    if let Some(name) = source.name() {
        output.push_str("\" office:name=\"");
        push_xml_attribute(&mut output, name);
    }
    match source.conversion_mode() {
        ConversionMode::Unspecified => {},
        ConversionMode::IntoDefaultStyleDataStyle => {
            output.push_str("\" office:conversion-mode=\"into-default-style-data-style")
        },
        ConversionMode::IntoEnglishNumber => {
            output.push_str("\" office:conversion-mode=\"into-english-number")
        },
        ConversionMode::KeepText => output.push_str("\" office:conversion-mode=\"keep-text"),
    }
    match source.automatic_update() {
        AutomaticUpdate::Unspecified => {},
        AutomaticUpdate::Enabled => output.push_str("\" office:automatic-update=\"true"),
        AutomaticUpdate::Disabled => output.push_str("\" office:automatic-update=\"false"),
    }
    // Every optional attribute branch above opens its quote before the value;
    // close the final attribute (or the required item value) uniformly.
    output.push_str("\"/>");
    Ok(output)
}

fn render_link_with_cache_name(
    link: &LinkSpec,
    cache_name: Option<&str>,
    context: &ExecutionContext,
) -> CoreResult<String> {
    context.check().map_err(map_execution_core)?;
    context
        .consume(Resource::Work, 1)
        .map_err(map_execution_core)?;
    let mut output = String::new();
    output
        .try_reserve(link_render_estimate_with_cache_name(
            link, cache_name, context,
        )?)
        .map_err(|_| CoreError::Unsupported("DDE link allocation failed".to_string()))?;
    output.push_str("<table:dde-link xmlns:table=\"");
    output.push_str(std::str::from_utf8(TABLE).unwrap_or(""));
    output.push_str("\">");
    output.push_str(&render_source_declaration(&link.source)?);
    output.push_str(&render_cached_table_with_name(
        &link.cached_table,
        cache_name,
        context,
    )?);
    output.push_str("</table:dde-link>");
    Ok(output)
}

fn render_cached_table_with_name(
    table: &CachedTable,
    cache_name: Option<&str>,
    context: &ExecutionContext,
) -> CoreResult<String> {
    let mut output = String::new();
    output
        .try_reserve(cached_table_render_estimate_with_name(
            table, cache_name, context,
        )?)
        .map_err(|_| CoreError::Unsupported("DDE cached table allocation failed".to_string()))?;
    output.push_str("<table:table xmlns:table=\"");
    output.push_str(std::str::from_utf8(TABLE).unwrap_or(""));
    output.push_str("\" xmlns:office=\"");
    output.push_str(std::str::from_utf8(OFFICE).unwrap_or(""));
    if let Some(name) = cache_name {
        output.push_str("\" table:name=\"");
        push_xml_attribute(&mut output, name);
    }
    output.push_str("\"><table:table-column/>");
    if table.rows.is_empty() {
        context.check().map_err(map_execution_core)?;
        context
            .consume(Resource::Work, 1)
            .map_err(map_execution_core)?;
        output.push_str("<table:table-row>");
        context.check().map_err(map_execution_core)?;
        context
            .consume(Resource::Work, 1)
            .map_err(map_execution_core)?;
        output.push_str("<table:table-cell/>");
        output.push_str("</table:table-row>");
    } else {
        for row in &table.rows {
            context.check().map_err(map_execution_core)?;
            context
                .consume(Resource::Work, 1)
                .map_err(map_execution_core)?;
            output.push_str("<table:table-row>");
            if row.cells.is_empty() {
                context.check().map_err(map_execution_core)?;
                context
                    .consume(Resource::Work, 1)
                    .map_err(map_execution_core)?;
                output.push_str("<table:table-cell/>");
            } else {
                for cell in &row.cells {
                    render_cached_cell(&mut output, &cell.value, context)?;
                }
            }
            output.push_str("</table:table-row>");
        }
    }
    output.push_str("</table:table>");
    Ok(output)
}

fn render_cached_cell(
    output: &mut String,
    value: &CachedValue,
    context: &ExecutionContext,
) -> CoreResult<()> {
    context.check().map_err(map_execution_core)?;
    context
        .consume(Resource::Work, 1)
        .map_err(map_execution_core)?;
    match value {
        CachedValue::Empty => output.push_str("<table:table-cell/>"),
        CachedValue::Text(text) => {
            output
                .push_str("<table:table-cell office:value-type=\"string\" office:string-value=\"");
            push_xml_attribute(output, text);
            output.push_str("\"/>");
        },
        CachedValue::Number(value) => {
            output.push_str("<table:table-cell office:value-type=\"float\" office:value=\"");
            output.push_str(&value.to_string());
            output.push_str("\"/>");
        },
        CachedValue::Currency { value, currency } => {
            output.push_str("<table:table-cell office:value-type=\"currency\" office:value=\"");
            output.push_str(&value.to_string());
            output.push_str("\" office:currency=\"");
            push_xml_attribute(output, currency);
            output.push_str("\"/>");
        },
        CachedValue::Percentage(value) => {
            output.push_str("<table:table-cell office:value-type=\"percentage\" office:value=\"");
            output.push_str(&value.to_string());
            output.push_str("\"/>");
        },
        CachedValue::Boolean(value) => {
            output.push_str(
                "<table:table-cell office:value-type=\"boolean\" office:boolean-value=\"",
            );
            output.push_str(if *value { "true" } else { "false" });
            output.push_str("\"/>");
        },
        CachedValue::Date(value) => {
            output.push_str("<table:table-cell office:value-type=\"date\" office:date-value=\"");
            push_xml_attribute(output, value);
            output.push_str("\"/>");
        },
        CachedValue::Time(value) => {
            output.push_str("<table:table-cell office:value-type=\"time\" office:time-value=\"");
            push_xml_attribute(output, value);
            output.push_str("\"/>");
        },
    }
    Ok(())
}

fn validate_cached_table(table: &CachedTable) -> CoreResult<()> {
    if table.rows.len() > MAX_CACHE_ROWS {
        return Err(resource_core(
            "DDE cached rows",
            table.rows.len(),
            MAX_CACHE_ROWS,
        ));
    }
    let mut cells = 0usize;
    for row in &table.rows {
        if row.cells.len() > MAX_CACHE_CELLS {
            return Err(resource_core(
                "DDE cached cells",
                row.cells.len(),
                MAX_CACHE_CELLS,
            ));
        }
        cells = cells
            .checked_add(row.cells.len())
            .ok_or_else(|| invalid_core("DDE cached cell count overflows"))?;
        for cell in &row.cells {
            cell.value.validate()?;
        }
    }
    if cells > MAX_CACHE_CELLS {
        return Err(resource_core("DDE cached cells", cells, MAX_CACHE_CELLS));
    }
    Ok(())
}

fn compare_sources(
    context: &ExecutionContext,
    expected: &Source,
    actual: &Source,
) -> CoreResult<bool> {
    let mut amount = 0usize;
    for value in [
        expected.application(),
        expected.topic(),
        expected.item(),
        actual.application(),
        actual.topic(),
        actual.item(),
    ] {
        amount = amount
            .checked_add(value.len())
            .ok_or_else(|| invalid_core("DDE source comparison size overflows"))?;
    }
    if let Some(value) = expected.name() {
        amount = amount
            .checked_add(value.len())
            .ok_or_else(|| invalid_core("DDE source comparison size overflows"))?;
    }
    if let Some(value) = actual.name() {
        amount = amount
            .checked_add(value.len())
            .ok_or_else(|| invalid_core("DDE source comparison size overflows"))?;
    }
    context.check().map_err(map_execution_core)?;
    context
        .consume(
            Resource::Work,
            u64::try_from(amount).map_err(|_| invalid_core("DDE source comparison exceeds u64"))?,
        )
        .map_err(map_execution_core)?;
    Ok(expected == actual)
}

fn compare_xml_bytes(context: &ExecutionContext, expected: &str, actual: &str) -> CoreResult<bool> {
    context.check().map_err(map_execution_core)?;
    let amount = expected.len().max(actual.len());
    context
        .consume(
            Resource::Work,
            u64::try_from(amount).map_err(|_| invalid_core("DDE XML comparison exceeds u64"))?,
        )
        .map_err(map_execution_core)?;
    Ok(expected == actual)
}

fn verify_readback(
    edit: &Edit,
    target: &Snapshot,
    cache_names: &[Option<String>],
) -> CoreResult<()> {
    let context = target.context();
    if let Some(expected_sources) = &edit.sheet_sources {
        if target.sheet_sources().len() != expected_sources.len() {
            return Err(invalid_core("DDE sheet source readback count changed"));
        }
        for (expected, actual) in expected_sources.iter().zip(target.sheet_sources()) {
            context.check().map_err(map_execution_core)?;
            context
                .consume(Resource::Work, 1)
                .map_err(map_execution_core)?;
            if !compare_xml_bytes(context, &expected.sheet, actual.sheet())?
                || !compare_sources(context, &expected.source, actual.source())?
            {
                return Err(invalid_core("DDE sheet source typed readback changed"));
            }
        }
    } else {
        if target.sheet_sources().len() != edit.before.sheet_sources().len() {
            return Err(invalid_core(
                "DDE sheet source readback count changed unexpectedly",
            ));
        }
        for (expected, actual) in edit
            .before
            .sheet_sources()
            .iter()
            .zip(target.sheet_sources())
        {
            context.check().map_err(map_execution_core)?;
            context
                .consume(Resource::Work, 1)
                .map_err(map_execution_core)?;
            if !compare_xml_bytes(context, expected.sheet(), actual.sheet())?
                || !compare_sources(context, expected.source(), actual.source())?
            {
                return Err(invalid_core(
                    "DDE sheet source readback changed unexpectedly",
                ));
            }
        }
    }
    let Some(expected_links) = &edit.links else {
        if target.links().len() != edit.before.links().len() {
            return Err(invalid_core("DDE link readback changed unexpectedly"));
        }
        for (expected, actual) in edit.before.links().iter().zip(target.links()) {
            context.check().map_err(map_execution_core)?;
            context
                .consume(Resource::Work, 1)
                .map_err(map_execution_core)?;
            if !compare_sources(context, expected.source(), actual.source())?
                || !compare_xml_bytes(
                    context,
                    expected.cached_table_xml(),
                    actual.cached_table_xml(),
                )?
            {
                return Err(invalid_core("DDE link readback changed unexpectedly"));
            }
        }
        return Ok(());
    };
    if target.links().len() != expected_links.len() {
        return Err(invalid_core("DDE link readback count changed"));
    }
    for (draft, actual) in expected_links.iter().zip(target.links()) {
        context.check().map_err(map_execution_core)?;
        context
            .consume(Resource::Work, 1)
            .map_err(map_execution_core)?;
        match draft {
            LinkDraft::Existing(index) => {
                let expected = edit
                    .before
                    .links()
                    .get(*index)
                    .ok_or_else(|| invalid_core("DDE link source inventory changed"))?;
                if !compare_sources(context, expected.source(), actual.source())?
                    || !compare_xml_bytes(
                        context,
                        expected.cached_table_xml(),
                        actual.cached_table_xml(),
                    )?
                {
                    return Err(invalid_core("DDE link source-order readback changed"));
                }
            },
            LinkDraft::Source {
                existing, source, ..
            } => {
                let original = edit
                    .before
                    .links()
                    .get(*existing)
                    .ok_or_else(|| invalid_core("DDE link source inventory changed"))?;
                if !compare_sources(context, source, actual.source())?
                    || !compare_xml_bytes(
                        context,
                        original.cached_table_xml(),
                        actual.cached_table_xml(),
                    )?
                {
                    return Err(invalid_core(
                        "DDE link source-only readback changed unexpectedly",
                    ));
                }
            },
            LinkDraft::Owned {
                link: expected,
                replaced,
                ..
            } => {
                if !compare_sources(context, &expected.source, actual.source())? {
                    return Err(invalid_core("DDE link source typed readback changed"));
                }
                let cache_name = match replaced {
                    Some(index) => cache_names
                        .get(*index)
                        .ok_or_else(|| invalid_core("DDE link source order changed"))?
                        .as_deref(),
                    None => None,
                };
                let estimate = cached_table_render_estimate_with_name(
                    &expected.cached_table,
                    cache_name,
                    context,
                )?;
                let reservation = context
                    .reserve(
                        Resource::Memory,
                        u64::try_from(estimate)
                            .map_err(|_| invalid_core("DDE readback memory exceeds u64"))?,
                    )
                    .map_err(map_execution_core)?;
                let expected_cache = render_cached_table_with_name(
                    &expected.cached_table,
                    cache_name,
                    target.context(),
                )?;
                let matches =
                    compare_xml_bytes(context, &expected_cache, actual.cached_table_xml())?;
                drop(reservation);
                if !matches {
                    return Err(invalid_core("DDE cached table typed readback changed"));
                }
            },
        }
    }
    Ok(())
}

fn scan_source(content: &str, limits: Limits, context: &ExecutionContext) -> CoreResult<Scan> {
    // `quick_xml` reports positions relative to the input handed to the
    // reader.  It does not retain a UTF-8 BOM in that stream, so parse the
    // post-BOM slice and translate every range back to the original source
    // coordinates used by the source-splice renderer.
    let source_offset = if content.as_bytes().starts_with(super::UTF8_BOM) {
        super::UTF8_BOM.len()
    } else {
        0
    };
    let reader_content = &content[source_offset..];
    let mut reader = NsReader::from_str(reader_content);
    reader.config_mut().check_end_names = true;
    reader.config_mut().trim_text(false);
    // Admit scanner-owned projections before constructing their containers.
    // The source snapshot's retained bytes cover the input itself, not these
    // tables, link sites, frame stack, or decoded source/name strings.
    let mut scan = Scan::new(context)?;
    let mut stack = Vec::<Frame>::new();
    let mut depth = 0usize;
    loop {
        context.check().map_err(map_execution_core)?;
        context
            .consume(Resource::Work, 1)
            .map_err(map_execution_core)?;
        let start = usize::try_from(reader.buffer_position())
            .map_err(|_| invalid_core("DDE XML position exceeds usize"))?
            .checked_add(source_offset)
            .ok_or_else(|| invalid_core("DDE XML position exceeds usize"))?;
        let (resolved, raw) = reader
            .read_resolved_event()
            .map_err(|error| invalid_core(format!("invalid DDE XML: {error}")))?;
        let namespace = namespace_tag(&resolved);
        let end = usize::try_from(reader.buffer_position())
            .map_err(|_| invalid_core("DDE XML position exceeds usize"))?
            .checked_add(source_offset)
            .ok_or_else(|| invalid_core("DDE XML position exceeds usize"))?;
        let is_eof = matches!(&raw, Event::Eof);
        match raw {
            Event::Start(element) => {
                reserve_event_attribute_values(&mut scan, &element, context)?;
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid_core("DDE XML depth overflows"))?;
                if depth > limits.depth() {
                    return Err(resource_core("DDE XML depth", depth, limits.depth()));
                }
                let parent = stack.last().map(|frame| frame.kind);
                let kind = classify_start(
                    &mut scan, parent, namespace, &element, start, end, &reader, limits, context,
                )?;
                scan.reserve_vec_slot::<Frame>(stack.len(), stack.capacity(), context)?;
                stack.try_reserve_exact(1).map_err(|_| {
                    CoreError::Unsupported("DDE XML frame allocation failed".to_string())
                })?;
                stack.push(Frame { kind, start });
            },
            Event::Empty(element) => {
                reserve_event_attribute_values(&mut scan, &element, context)?;
                let parent = stack.last().map(|frame| frame.kind);
                classify_empty(
                    &mut scan, parent, namespace, &element, start, end, &reader, limits, context,
                )?;
            },
            Event::End(_element) => {
                let frame = stack
                    .pop()
                    .ok_or_else(|| invalid_core("DDE XML element stack underflow"))?;
                finish_frame(&mut scan, frame, start, end)?;
                if matches!(frame.kind, FrameKind::Mce) {
                    scan.mce_depth = scan.mce_depth.saturating_sub(1);
                }
                depth = depth.saturating_sub(1);
            },
            Event::Comment(_) | Event::PI(_) => mark_misc(&mut scan, &stack),
            Event::Text(text) => {
                let value = text
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| invalid_core(format!("invalid DDE text: {error}")))?;
                if !value.trim().is_empty() {
                    mark_non_whitespace(&mut scan, &stack);
                }
            },
            Event::CData(text) => {
                let value = text
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| invalid_core(format!("invalid DDE CDATA: {error}")))?;
                if !value.trim().is_empty() {
                    mark_non_whitespace(&mut scan, &stack);
                }
            },
            Event::GeneralRef(reference) => {
                let reference_bytes: &[u8] = reference.as_ref();
                let reference_length = u64::try_from(reference_bytes.len())
                    .map_err(|_| invalid_core("DDE entity reference size exceeds u64"))?;
                context
                    .consume(Resource::Work, reference_length)
                    .map_err(map_execution_core)?;
                if !super::valid_general_ref(reference_bytes) {
                    return Err(invalid_core(
                        "DDE XML contains an unsupported entity reference",
                    ));
                }
                mark_non_whitespace(&mut scan, &stack);
            },
            Event::DocType(_) => return Err(invalid_core("DTD content is not accepted")),
            Event::Decl(_) | Event::Eof => {},
        }
        if is_eof {
            break;
        }
    }
    if !stack.is_empty() || depth != 0 {
        return Err(invalid_core("unfinished DDE XML structure"));
    }
    Ok(scan)
}

fn classify_start(
    scan: &mut Scan,
    parent: Option<FrameKind>,
    namespace: NamespaceTag,
    element: &BytesStart<'_>,
    start: usize,
    _end: usize,
    reader: &NsReader<&[u8]>,
    limits: Limits,
    context: &ExecutionContext,
) -> CoreResult<FrameKind> {
    let local = element.local_name();
    note_table_child(scan, parent, namespace, local.as_ref(), start);
    if has_markup_compatibility_attribute(element, reader, context)? {
        scan.mce_present = true;
    }
    if namespace == NamespaceTag::MarkupCompatibility {
        scan.mce_present = true;
        scan.mce_depth = scan
            .mce_depth
            .checked_add(1)
            .ok_or_else(|| invalid_core("DDE MCE depth overflows"))?;
        return Ok(FrameKind::Mce);
    }
    if scan.mce_depth != 0 {
        mark_unknown_child(scan, parent);
        return Ok(FrameKind::Other);
    }
    if namespace == NamespaceTag::Office
        && matches!(local.as_ref(), b"document-content" | b"document")
        && parent.is_none()
    {
        return Ok(FrameKind::DocumentContent);
    }
    if namespace == NamespaceTag::Office
        && local.as_ref() == b"body"
        && matches!(parent, Some(FrameKind::DocumentContent))
    {
        return Ok(FrameKind::Body);
    }
    if namespace == NamespaceTag::Office && local.as_ref() == b"spreadsheet" {
        if !matches!(parent, Some(FrameKind::Body)) {
            mark_unknown_child(scan, parent);
            return Ok(FrameKind::Other);
        }
        if scan.spreadsheet.is_some() {
            return Err(invalid_core("duplicate office:spreadsheet owner"));
        }
        scan.spreadsheet = Some(ElementSite {
            range: start..start,
            opaque: false,
        });
        return Ok(FrameKind::Spreadsheet);
    }
    if matches!(parent, Some(FrameKind::Spreadsheet)) && namespace == NamespaceTag::Table {
        if local.as_ref() == b"table" {
            let name = source_table_name(element, reader, limits, context)?;
            scan.reserve_vec_slot::<TableSite>(scan.tables.len(), scan.tables.capacity(), context)?;
            scan.tables.try_reserve_exact(1).map_err(|_| {
                CoreError::Unsupported("DDE table catalog allocation failed".to_string())
            })?;
            let index = scan.tables.len();
            scan.tables.push(TableSite {
                element: ElementSite {
                    range: start..start,
                    opaque: false,
                },
                name,
                source: None,
                source_value: None,
                source_insertion: None,
                paired: true,
            });
            return Ok(FrameKind::Table(index));
        }
        if local.as_ref() == b"dde-links" {
            if scan.links_container.is_some() {
                return Err(invalid_core("duplicate table:dde-links owner"));
            }
            let opaque = !owner_attributes_known(element, context)?;
            scan.links_container = Some(ElementSite {
                range: start..start,
                opaque,
            });
            scan.links_container_opaque |= opaque;
            return Ok(FrameKind::Links);
        }
    }
    if matches!(parent, Some(FrameKind::Link(_)))
        && namespace == NamespaceTag::Table
        && local.as_ref() == b"table"
    {
        let FrameKind::Link(index) = parent.unwrap_or(FrameKind::Other) else {
            unreachable!()
        };
        let site = scan
            .links
            .get_mut(index)
            .ok_or_else(|| invalid_core("DDE link scanner state is missing"))?;
        if site.cache.is_some() {
            return Err(invalid_core("duplicate DDE link cached table"));
        }
        let (known, cache_name) = cache_owner_attributes_known(element, reader, limits, context)?;
        let opaque = !known;
        site.cache = Some(ElementSite {
            range: start..start,
            opaque,
        });
        site.cache_name = cache_name;
        if opaque {
            site.cache_opaque = true;
        }
        return Ok(FrameKind::Cache(index));
    }
    if matches!(parent, Some(FrameKind::Links))
        && namespace == NamespaceTag::Table
        && local.as_ref() == b"dde-link"
    {
        scan.reserve_vec_slot::<LinkSite>(scan.links.len(), scan.links.capacity(), context)?;
        scan.links.try_reserve_exact(1).map_err(|_| {
            CoreError::Unsupported("DDE link catalog allocation failed".to_string())
        })?;
        let index = scan.links.len();
        let opaque = !owner_attributes_known(element, context)?;
        scan.links.push(LinkSite {
            element: ElementSite {
                range: start..start,
                opaque,
            },
            source: None,
            source_value: None,
            cache: None,
            cache_name: None,
            link_opaque: opaque,
            cache_opaque: false,
        });
        return Ok(FrameKind::Link(index));
    }
    if matches!(parent, Some(FrameKind::Table(_)))
        && namespace == NamespaceTag::Office
        && local.as_ref() == b"dde-source"
    {
        let FrameKind::Table(index) = parent.unwrap_or(FrameKind::Other) else {
            unreachable!()
        };
        if scan
            .tables
            .get(index)
            .ok_or_else(|| invalid_core("DDE table scanner state is missing"))?
            .source
            .is_some()
        {
            return Err(invalid_core("duplicate DDE sheet source"));
        }
        reserve_source_projection(&mut *scan, element, context)?;
        let value = super::parse_source_with_context(element, reader, limits.text_bytes(), context)
            .map_err(CoreError::from)?;
        let site = scan
            .tables
            .get_mut(index)
            .ok_or_else(|| invalid_core("DDE table scanner state is missing"))?;
        site.source = Some(ElementSite {
            range: start..start,
            opaque: !source_attributes_known(element, reader, context)?,
        });
        site.source_value = Some(value);
        return Ok(FrameKind::SourceTable(index));
    }
    if matches!(parent, Some(FrameKind::Link(_)))
        && namespace == NamespaceTag::Office
        && local.as_ref() == b"dde-source"
    {
        let FrameKind::Link(index) = parent.unwrap_or(FrameKind::Other) else {
            unreachable!()
        };
        let Some(existing_site) = scan.links.get(index) else {
            return Err(invalid_core("DDE link scanner state is missing"));
        };
        if existing_site.source.is_some() || existing_site.cache.is_some() {
            return Err(invalid_core("DDE link source is out of order"));
        }
        reserve_source_projection(&mut *scan, element, context)?;
        let value = super::parse_source_with_context(element, reader, limits.text_bytes(), context)
            .map_err(CoreError::from)?;
        let site = scan
            .links
            .get_mut(index)
            .ok_or_else(|| invalid_core("DDE link scanner state is missing"))?;
        site.source = Some(ElementSite {
            range: start..start,
            opaque: !source_attributes_known(element, reader, context)?,
        });
        site.source_value = Some(value);
        return Ok(FrameKind::SourceLink(index));
    }
    if let Some(cache_index) = cache_parent_index(parent) {
        if (namespace == NamespaceTag::Text && is_cache_text_element(local.as_ref()))
            || (namespace == NamespaceTag::Table && is_cache_table_element(local.as_ref()))
        {
            if !cache_child_attributes_known(element, reader, context)? {
                scan.links[cache_index].cache_opaque = true;
            }
            return Ok(FrameKind::CacheContent(cache_index));
        }
    }
    mark_unknown_child(scan, parent);
    Ok(FrameKind::Other)
}

fn classify_empty(
    scan: &mut Scan,
    parent: Option<FrameKind>,
    namespace: NamespaceTag,
    element: &BytesStart<'_>,
    start: usize,
    end: usize,
    reader: &NsReader<&[u8]>,
    limits: Limits,
    context: &ExecutionContext,
) -> CoreResult<()> {
    let local = element.local_name();
    note_table_child(scan, parent, namespace, local.as_ref(), start);
    if has_markup_compatibility_attribute(element, reader, context)? {
        scan.mce_present = true;
    }
    if namespace == NamespaceTag::MarkupCompatibility {
        scan.mce_present = true;
        return Ok(());
    }
    if scan.mce_depth != 0 {
        mark_unknown_child(scan, parent);
        return Ok(());
    }
    if namespace == NamespaceTag::Office
        && matches!(local.as_ref(), b"document-content" | b"document")
        && parent.is_none()
    {
        return Ok(());
    }
    if namespace == NamespaceTag::Office
        && local.as_ref() == b"body"
        && matches!(parent, Some(FrameKind::DocumentContent))
    {
        return Ok(());
    }
    if namespace == NamespaceTag::Office && local.as_ref() == b"spreadsheet" {
        mark_unknown_child(scan, parent);
        return Ok(());
    }
    if matches!(parent, Some(FrameKind::Spreadsheet)) && namespace == NamespaceTag::Table {
        if local.as_ref() == b"table" {
            let name = source_table_name(element, reader, limits, context)?;
            scan.reserve_vec_slot::<TableSite>(scan.tables.len(), scan.tables.capacity(), context)?;
            scan.tables.try_reserve_exact(1).map_err(|_| {
                CoreError::Unsupported("DDE table catalog allocation failed".to_string())
            })?;
            scan.tables.push(TableSite {
                element: ElementSite {
                    range: start..end,
                    opaque: false,
                },
                name,
                source: None,
                source_value: None,
                source_insertion: None,
                paired: false,
            });
            return Ok(());
        }
        if local.as_ref() == b"dde-links" {
            return Err(invalid_core("table:dde-links must contain a link"));
        }
    }
    if matches!(parent, Some(FrameKind::Links))
        && namespace == NamespaceTag::Table
        && local.as_ref() == b"dde-link"
    {
        return Err(invalid_core(
            "table:dde-link requires source and cached table",
        ));
    }
    if matches!(parent, Some(FrameKind::Link(_)))
        && namespace == NamespaceTag::Office
        && local.as_ref() == b"dde-source"
    {
        let FrameKind::Link(index) = parent.unwrap_or(FrameKind::Other) else {
            unreachable!()
        };
        let Some(existing_site) = scan.links.get(index) else {
            return Err(invalid_core("DDE link scanner state is missing"));
        };
        if existing_site.source.is_some() || existing_site.cache.is_some() {
            return Err(invalid_core("DDE link source is out of order"));
        }
        reserve_source_projection(&mut *scan, element, context)?;
        let value = super::parse_source_with_context(element, reader, limits.text_bytes(), context)
            .map_err(CoreError::from)?;
        let site = scan
            .links
            .get_mut(index)
            .ok_or_else(|| invalid_core("DDE link scanner state is missing"))?;
        site.source = Some(ElementSite {
            range: start..end,
            opaque: !source_attributes_known(element, reader, context)?,
        });
        site.source_value = Some(value);
        return Ok(());
    }
    if matches!(parent, Some(FrameKind::Link(_)))
        && namespace == NamespaceTag::Table
        && local.as_ref() == b"table"
    {
        let FrameKind::Link(index) = parent.unwrap_or(FrameKind::Other) else {
            unreachable!()
        };
        let site = scan
            .links
            .get_mut(index)
            .ok_or_else(|| invalid_core("DDE link scanner state is missing"))?;
        if site.source.is_none() || site.cache.is_some() {
            return Err(invalid_core("DDE link cache is out of order"));
        }
        let (known, cache_name) = cache_owner_attributes_known(element, reader, limits, context)?;
        let opaque = !known;
        site.cache = Some(ElementSite {
            range: start..end,
            opaque,
        });
        site.cache_name = cache_name;
        if opaque {
            site.cache_opaque = true;
        }
        return Ok(());
    }
    if matches!(parent, Some(FrameKind::Table(_)))
        && namespace == NamespaceTag::Office
        && local.as_ref() == b"dde-source"
    {
        let FrameKind::Table(index) = parent.unwrap_or(FrameKind::Other) else {
            unreachable!()
        };
        if scan
            .tables
            .get(index)
            .ok_or_else(|| invalid_core("DDE table scanner state is missing"))?
            .source
            .is_some()
        {
            return Err(invalid_core("duplicate DDE sheet source"));
        }
        reserve_source_projection(&mut *scan, element, context)?;
        let value = super::parse_source_with_context(element, reader, limits.text_bytes(), context)
            .map_err(CoreError::from)?;
        let site = scan
            .tables
            .get_mut(index)
            .ok_or_else(|| invalid_core("DDE table scanner state is missing"))?;
        site.source = Some(ElementSite {
            range: start..end,
            opaque: !source_attributes_known(element, reader, context)?,
        });
        site.source_value = Some(value);
        return Ok(());
    }
    if let Some(cache_index) = cache_parent_index(parent) {
        if (namespace == NamespaceTag::Text && is_cache_text_element(local.as_ref()))
            || (namespace == NamespaceTag::Table && is_cache_table_element(local.as_ref()))
        {
            if !cache_child_attributes_known(element, reader, context)? {
                scan.links[cache_index].cache_opaque = true;
            }
            return Ok(());
        }
    }
    mark_unknown_child(scan, parent);
    Ok(())
}

fn finish_frame(scan: &mut Scan, frame: Frame, closing_start: usize, end: usize) -> CoreResult<()> {
    match frame.kind {
        FrameKind::DocumentContent
        | FrameKind::Body
        | FrameKind::Mce
        | FrameKind::Other
        | FrameKind::CacheContent(_) => {},
        FrameKind::Spreadsheet => {
            let site = scan
                .spreadsheet
                .as_mut()
                .ok_or_else(|| invalid_core("DDE spreadsheet scanner state is missing"))?;
            site.range.end = end;
            scan.spreadsheet_close_start = Some(closing_start);
        },
        FrameKind::Table(index) => {
            let site = scan
                .tables
                .get_mut(index)
                .ok_or_else(|| invalid_core("DDE table scanner state is missing"))?;
            site.element.range = frame.start..end;
            if site.source.is_none() && site.source_insertion.is_none() {
                site.source_insertion = Some(closing_start);
            }
        },
        FrameKind::Links => {
            let site = scan
                .links_container
                .as_mut()
                .ok_or_else(|| invalid_core("DDE links scanner state is missing"))?;
            site.range.end = end;
        },
        FrameKind::Link(index) => {
            let site = scan
                .links
                .get_mut(index)
                .ok_or_else(|| invalid_core("DDE link scanner state is missing"))?;
            site.element.range = frame.start..end;
            if site.source.is_none() || site.cache.is_none() {
                return Err(invalid_core("DDE link requires source and cached table"));
            }
        },
        FrameKind::SourceTable(index) => {
            let source = scan
                .tables
                .get_mut(index)
                .and_then(|site| site.source.as_mut())
                .ok_or_else(|| invalid_core("DDE source scanner state is missing"))?;
            source.range.end = end;
        },
        FrameKind::SourceLink(index) => {
            let source = scan
                .links
                .get_mut(index)
                .and_then(|site| site.source.as_mut())
                .ok_or_else(|| invalid_core("DDE source scanner state is missing"))?;
            source.range.end = end;
        },
        FrameKind::Cache(index) => {
            let cache = scan
                .links
                .get_mut(index)
                .and_then(|site| site.cache.as_mut())
                .ok_or_else(|| invalid_core("DDE cache scanner state is missing"))?;
            cache.range.end = end;
        },
    }
    Ok(())
}

fn mark_misc(scan: &mut Scan, stack: &[Frame]) {
    for frame in stack.iter().rev() {
        match frame.kind {
            FrameKind::SourceTable(index) => {
                if let Some(source) = scan.tables[index].source.as_mut() {
                    source.opaque = true;
                }
                return;
            },
            FrameKind::SourceLink(index) => {
                if let Some(source) = scan.links[index].source.as_mut() {
                    source.opaque = true;
                }
                return;
            },
            FrameKind::Cache(index) => {
                scan.links[index].cache_opaque = true;
                return;
            },
            FrameKind::CacheContent(index) => {
                scan.links[index].cache_opaque = true;
                return;
            },
            FrameKind::Link(index) => {
                scan.links[index].link_opaque = true;
                return;
            },
            FrameKind::Links => {
                scan.links_container_opaque = true;
                return;
            },
            _ => {},
        }
    }
}

fn mark_non_whitespace(scan: &mut Scan, stack: &[Frame]) {
    for frame in stack.iter().rev() {
        match frame.kind {
            FrameKind::SourceTable(index) => {
                if let Some(source) = scan.tables[index].source.as_mut() {
                    source.opaque = true;
                }
                return;
            },
            FrameKind::SourceLink(index) => {
                if let Some(source) = scan.links[index].source.as_mut() {
                    source.opaque = true;
                }
                return;
            },
            FrameKind::Cache(index) | FrameKind::CacheContent(index) => {
                scan.links[index].cache_opaque = true;
                return;
            },
            FrameKind::Link(index) => {
                scan.links[index].link_opaque = true;
                return;
            },
            FrameKind::Links => {
                scan.links_container_opaque = true;
                return;
            },
            _ => {},
        }
    }
}

fn mark_unknown_child(scan: &mut Scan, parent: Option<FrameKind>) {
    match parent {
        Some(FrameKind::Link(index)) => scan.links[index].link_opaque = true,
        Some(FrameKind::Links) => scan.links_container_opaque = true,
        Some(FrameKind::SourceTable(index)) => {
            if let Some(source) = scan.tables[index].source.as_mut() {
                source.opaque = true;
            }
        },
        Some(FrameKind::SourceLink(index)) => {
            if let Some(source) = scan.links[index].source.as_mut() {
                source.opaque = true;
            }
        },
        Some(FrameKind::Cache(index)) => scan.links[index].cache_opaque = true,
        Some(FrameKind::CacheContent(index)) => scan.links[index].cache_opaque = true,
        _ => {},
    }
}

fn source_table_name(
    element: &BytesStart<'_>,
    reader: &NsReader<&[u8]>,
    limits: Limits,
    context: &ExecutionContext,
) -> CoreResult<Option<String>> {
    super::optional_attr_with_context(
        element,
        reader,
        TABLE,
        b"name",
        limits.text_bytes(),
        context,
    )
    .map_err(CoreError::from)
}

/// Admit raw attribute values before the namespace reader or semantic
/// projection can materialize any corresponding strings.  Raw lexical values
/// are an upper bound for their decoded forms (entity decoding and XML
/// whitespace normalization cannot increase their size).
fn reserve_event_attribute_values(
    scan: &mut Scan,
    element: &BytesStart<'_>,
    context: &ExecutionContext,
) -> CoreResult<()> {
    let mut amount = 0usize;
    for raw in element.attributes().with_checks(true) {
        context.check().map_err(map_execution_core)?;
        context
            .consume(Resource::Work, 1)
            .map_err(map_execution_core)?;
        let attribute =
            raw.map_err(|error| invalid_core(format!("invalid DDE attribute: {error}")))?;
        amount = amount
            .checked_add(attribute.value.len())
            .ok_or_else(|| invalid_core("DDE scanner attribute memory size overflows"))?;
    }
    scan.reserve_memory(context, amount)
}

/// Admit the fixed Source projection before parsing its decoded attribute
/// values.  The raw attribute values were admitted by
/// `reserve_event_attribute_values` for the current event.
fn reserve_source_projection(
    scan: &mut Scan,
    _element: &BytesStart<'_>,
    context: &ExecutionContext,
) -> CoreResult<()> {
    scan.reserve_memory(context, size_of::<Source>())
}

fn note_table_child(
    scan: &mut Scan,
    parent: Option<FrameKind>,
    namespace: NamespaceTag,
    local: &[u8],
    start: usize,
) {
    let Some(FrameKind::Table(index)) = parent else {
        return;
    };
    if namespace == NamespaceTag::Office && local == b"dde-source" {
        return;
    }
    if namespace == NamespaceTag::Table && matches!(local, b"title" | b"desc" | b"table-source") {
        return;
    }
    if let Some(table) = scan.tables.get_mut(index) {
        table.source_insertion.get_or_insert(start);
    }
}

fn is_cache_table_element(local: &[u8]) -> bool {
    matches!(
        local,
        b"table"
            | b"table-column"
            | b"table-columns"
            | b"table-column-group"
            | b"table-header-columns"
            | b"table-row"
            | b"table-rows"
            | b"table-row-group"
            | b"table-header-rows"
            | b"table-cell"
            | b"covered-table-cell"
    )
}

fn is_cache_text_element(local: &[u8]) -> bool {
    matches!(
        local,
        b"p" | b"span" | b"a" | b"s" | b"tab" | b"line-break" | b"soft-page-break"
    )
}

fn cache_parent_index(parent: Option<FrameKind>) -> Option<usize> {
    match parent {
        Some(FrameKind::Cache(index) | FrameKind::CacheContent(index)) => Some(index),
        _ => None,
    }
}

/// Returns whether an owner can be safely replaced while retaining all
/// attributes.  Namespace declarations are self-contained on generated
/// owners and therefore never constitute unknown structure.
fn owner_attributes_known(
    element: &BytesStart<'_>,
    context: &ExecutionContext,
) -> CoreResult<bool> {
    for raw in element.attributes().with_checks(true) {
        context
            .consume(Resource::Work, 1)
            .map_err(map_execution_core)?;
        let Ok(attribute) = raw else { return Ok(false) };
        let key = attribute.key.as_ref();
        if key == b"xmlns" || key.starts_with(b"xmlns:") {
            continue;
        }
        return Ok(false);
    }
    Ok(true)
}

/// Admit the attributes on a cached scalar table owner.  `table:name` is a
/// standard modeled identity and is retained when a typed replacement is
/// rendered.  Other table-owner attributes carry presentation or table
/// structure that [`CachedTable`] does not model, so replacing such an owner
/// would silently discard producer state and must fail closed.
fn cache_owner_attributes_known(
    element: &BytesStart<'_>,
    reader: &NsReader<&[u8]>,
    limits: Limits,
    context: &ExecutionContext,
) -> CoreResult<(bool, Option<String>)> {
    let mut name_count = 0usize;
    for raw in element.attributes().with_checks(true) {
        context
            .consume(Resource::Work, 1)
            .map_err(map_execution_core)?;
        let Ok(attribute) = raw else {
            return Ok((false, None));
        };
        let key = attribute.key.as_ref();
        if key == b"xmlns" || key.starts_with(b"xmlns:") {
            continue;
        }
        let (resolved, local) = reader.resolver().resolve_attribute(attribute.key);
        if is_namespace(&resolved, TABLE) && local.as_ref() == b"name" {
            name_count = name_count
                .checked_add(1)
                .ok_or_else(|| invalid_core("DDE cache name count overflows"))?;
            if name_count > 1 {
                return Err(invalid_core("duplicate DDE cache table name"));
            }
        } else {
            return Ok((false, None));
        }
    }
    let name = if name_count == 0 {
        None
    } else {
        super::optional_attr_with_context(
            element,
            reader,
            TABLE,
            b"name",
            limits.text_bytes(),
            context,
        )
        .map_err(CoreError::from)?
    };
    Ok((true, name))
}

fn has_markup_compatibility_attribute(
    element: &BytesStart<'_>,
    reader: &NsReader<&[u8]>,
    context: &ExecutionContext,
) -> CoreResult<bool> {
    for raw in element.attributes().with_checks(true) {
        context
            .consume(Resource::Work, 1)
            .map_err(map_execution_core)?;
        let Ok(attribute) = raw else { return Ok(false) };
        let (resolved, _) = reader.resolver().resolve_attribute(attribute.key);
        if namespace_tag(&resolved) == NamespaceTag::MarkupCompatibility {
            return Ok(true);
        }
    }
    Ok(false)
}

/// Check attributes on cache descendants before allowing a typed whole-cache
/// replacement.  The known ODF table/office/text attributes are interpreted by
/// the metadata reader; foreign or otherwise structural attributes make the
/// cache opaque so the transaction fails closed instead of silently dropping
/// them.
fn cache_child_attributes_known(
    element: &BytesStart<'_>,
    reader: &NsReader<&[u8]>,
    context: &ExecutionContext,
) -> CoreResult<bool> {
    let (resolved_element, _) = reader.resolver().resolve_element(element.name());
    let element_namespace = namespace_tag(&resolved_element);
    let local_element = element.local_name();
    let element_kind = if element_namespace != NamespaceTag::Table {
        return Ok(false);
    } else if local_element.as_ref() == b"table-column" {
        0u8
    } else if local_element.as_ref() == b"table-row" {
        1u8
    } else if local_element.as_ref() == b"table-cell" {
        2u8
    } else {
        // Covered cells, groups, headers, nested tables, and every text
        // element carry cache structure or presentation not represented by
        // the typed scalar cache model.  Even an empty owner is therefore
        // opaque and cannot be destructively replaced.
        return Ok(false);
    };
    let mut value_type_seen = false;
    for raw in element.attributes().with_checks(true) {
        context
            .consume(Resource::Work, 1)
            .map_err(map_execution_core)?;
        let Ok(attribute) = raw else { return Ok(false) };
        let key = attribute.key.as_ref();
        if key == b"xmlns" || key.starts_with(b"xmlns:") {
            continue;
        }
        let (resolved, local) = reader.resolver().resolve_attribute(attribute.key);
        let known = match namespace_tag(&resolved) {
            NamespaceTag::Office => {
                if element_kind != 2 {
                    false
                } else if local.as_ref() == b"value-type" {
                    if value_type_seen {
                        return Ok(false);
                    }
                    value_type_seen = true;
                    let value = attribute
                        .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
                        .map_err(|error| {
                            invalid_core(format!("invalid DDE cache value type: {error}"))
                        })?;
                    matches!(
                        value.as_ref(),
                        "string"
                            | "float"
                            | "currency"
                            | "percentage"
                            | "boolean"
                            | "date"
                            | "time"
                    )
                } else {
                    matches!(
                        local.as_ref(),
                        b"value"
                            | b"currency"
                            | b"date-value"
                            | b"time-value"
                            | b"boolean-value"
                            | b"string-value"
                    )
                }
            },
            NamespaceTag::Table => match element_kind {
                0 => local.as_ref() == b"number-columns-repeated",
                1 => local.as_ref() == b"number-rows-repeated",
                2 => matches!(
                    local.as_ref(),
                    b"number-columns-repeated" | b"number-rows-repeated"
                ),
                _ => false,
            },
            _ => false,
        };
        if !known {
            return Ok(false);
        }
    }
    Ok(true)
}

fn source_attributes_known(
    element: &BytesStart<'_>,
    reader: &NsReader<&[u8]>,
    context: &ExecutionContext,
) -> CoreResult<bool> {
    for raw in element.attributes().with_checks(true) {
        context
            .consume(Resource::Work, 1)
            .map_err(map_execution_core)?;
        let Ok(attribute) = raw else { return Ok(false) };
        let key = attribute.key.as_ref();
        if key == b"xmlns" || key.starts_with(b"xmlns:") {
            continue;
        }
        let (resolved, local) = reader.resolver().resolve_attribute(attribute.key);
        if !is_namespace(&resolved, OFFICE)
            || !matches!(
                local.as_ref(),
                b"dde-application"
                    | b"dde-topic"
                    | b"dde-item"
                    | b"name"
                    | b"conversion-mode"
                    | b"automatic-update"
            )
        {
            return Ok(false);
        }
    }
    Ok(true)
}

fn is_namespace(namespace: &ResolveResult<'_>, expected: &[u8]) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == expected)
}

fn namespace_tag(namespace: &ResolveResult<'_>) -> NamespaceTag {
    if is_namespace(namespace, OFFICE) {
        NamespaceTag::Office
    } else if is_namespace(namespace, TABLE) {
        NamespaceTag::Table
    } else if is_namespace(namespace, TEXT) {
        NamespaceTag::Text
    } else if is_namespace(namespace, MARKUP_COMPATIBILITY) {
        NamespaceTag::MarkupCompatibility
    } else {
        NamespaceTag::Other
    }
}

fn same_budget_lineage(left: &ExecutionContext, right: &ExecutionContext) -> CoreResult<bool> {
    let mut left_reservation = left
        .reserve(Resource::Memory, 0)
        .map_err(map_execution_core)?;
    let right_reservation = right
        .reserve(Resource::Memory, 0)
        .map_err(map_execution_core)?;
    Ok(left_reservation.try_merge(right_reservation).is_ok())
}

fn map_execution_core(error: ExecutionError) -> CoreError {
    match error {
        ExecutionError::Cancelled => {
            CoreError::Unsupported("ODS DDE operation cancelled".to_string())
        },
        ExecutionError::ResourceLimit(limit) => CoreError::ResourceLimit(limit),
        other => CoreError::Unsupported(format!(
            "ODS DDE execution policy rejected operation: {other}"
        )),
    }
}

fn invalid_core(message: impl Into<String>) -> CoreError {
    CoreError::InvalidFormat(message.into())
}

fn resource_core(resource: &'static str, actual: usize, maximum: usize) -> CoreError {
    CoreError::ResourceLimit(litchi_core::ResourceLimit {
        resource: match resource {
            "output bytes" => Resource::OutputBytes,
            "DDE XML depth" => Resource::Depth,
            "DDE links" | "DDE cached rows" | "DDE cached cells" => Resource::Objects,
            _ => Resource::Memory,
        },
        observed: actual as u64,
        limit: maximum as u64,
        scope: Arc::from("ods-dde"),
    })
}

fn validate_text(value: &str, kind: &str) -> CoreResult<()> {
    if value.is_empty() || !super::xml_text_is_valid(value) {
        return Err(invalid_core(format!(
            "{kind} is empty or contains invalid XML text"
        )));
    }
    if value.len() > super::MAX_TEXT_BYTES {
        return Err(resource_core(
            "DDE cached text",
            value.len(),
            super::MAX_TEXT_BYTES,
        ));
    }
    Ok(())
}

fn validate_authored_source(source: &Source) -> CoreResult<()> {
    if source.name().is_none() {
        return Err(invalid_core(
            "authored DDE source declarations require office:name",
        ));
    }
    Ok(())
}

fn validate_date(value: &str) -> CoreResult<()> {
    validate_text(value, "cached date")?;
    super::lexical::validate_date_value(value)
}

fn validate_duration(value: &str, kind: &str) -> CoreResult<()> {
    validate_text(value, kind)?;
    OdfDuration::decode_exact(value)
        .map(|_| ())
        .map_err(|error| invalid_core(format!("{kind} is not a valid ODF duration: {error}")))
}
