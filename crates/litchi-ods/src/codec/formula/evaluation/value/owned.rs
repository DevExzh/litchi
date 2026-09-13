//! Fallible, lifetime-free ownership for value-evaluator results.
//!
//! This module is deliberately a child of the value VM.  It can therefore
//! inspect the VM's retained representation without adding a second public
//! evaluator or making the VM's storage types part of the API.  The public
//! surface is an opaque owned result plus borrowed views over that result.
//! Conversion is a two-pass operation: the first pass checks limits, charges
//! copy work, and computes every destination capacity; the second pass makes
//! only fallible, exact-capacity copies while one reservation remains attached
//! to the result.

use super::super::{
    EvaluationFailure, EvaluationResult, ScalarError, WorkingValue, map_execution_error,
};
use super::{Area, Evaluated, Limits, RuntimeElement, RuntimeValue, Shape};
use crate::codec::formula::reference::{
    Address, Cell, Column, Endpoint, EndpointValue, Reference, Row, SheetLocator, SheetName,
    SheetSelector, Source, Subtable,
};
use litchi_core::{ExecutionContext, Reservation, Resource, ResourceLimit};
use std::{mem::size_of, sync::Arc};

const OWNED_SCOPE: &str = "ods-formula-owned-value";
const COPY_CHUNK_BYTES: usize = 4096;

/// A lifetime-free evaluation result.
///
/// The storage and its reservation are private as a unit.  Callers can retain
/// this value after dropping the parsed expression and resolver, while
/// inspection remains borrowed through [`Self::value`].
#[derive(Debug)]
pub struct OwnedEvaluated {
    value: OwnedValue,
    reservation: Option<Reservation>,
}

/// Borrowed inspection of an [`OwnedEvaluated`] value.
#[derive(Clone, Copy, Debug, PartialEq)]
#[non_exhaustive]
pub enum OwnedValueView<'a> {
    /// An absent or explicitly empty cell.
    Empty,
    /// A finite numeric value.
    Number(f64),
    /// A logical value.
    Logical(bool),
    /// Text retained by the owned result.
    Text(&'a str),
    /// A formula-level error value.
    Error(ScalarError),
    /// A rectangular owned array viewed without copying.
    Array(OwnedArrayView<'a>),
    /// One owned reference record.
    Reference(super::ReferenceView<'a>),
    /// An ordered owned reference list.
    ReferenceList(OwnedReferenceListView<'a>),
}

impl OwnedValueView<'_> {
    /// Return the array shape, if this is an array.
    #[must_use]
    pub const fn shape(self) -> Option<Shape> {
        match self {
            Self::Array(array) => Some(array.shape()),
            Self::Empty
            | Self::Number(_)
            | Self::Logical(_)
            | Self::Text(_)
            | Self::Error(_)
            | Self::Reference(_)
            | Self::ReferenceList(_) => None,
        }
    }
}

/// Borrowed row-major inspection of an owned array.
#[derive(Clone, Copy, Debug)]
pub struct OwnedArrayView<'a> {
    shape: Shape,
    cells: &'a [OwnedElement],
}

impl PartialEq for OwnedArrayView<'_> {
    fn eq(&self, other: &Self) -> bool {
        self.shape == other.shape && self.cells == other.cells
    }
}

impl<'a> OwnedArrayView<'a> {
    /// Array shape.
    #[must_use]
    pub const fn shape(self) -> Shape {
        self.shape
    }

    /// Number of retained cells.
    #[must_use]
    pub const fn len(self) -> usize {
        self.cells.len()
    }

    /// Whether the array has no retained cells.
    #[must_use]
    pub const fn is_empty(self) -> bool {
        self.cells.is_empty()
    }

    /// Return one zero-based row-major cell.
    #[must_use]
    pub fn cell(self, row: usize, column: usize) -> Option<OwnedValueView<'a>> {
        if row >= self.shape.rows() || column >= self.shape.columns() {
            return None;
        }
        let index = row.checked_mul(self.shape.columns())?.checked_add(column)?;
        self.cells.get(index).map(OwnedElement::as_view)
    }

    /// Return one row-major cell by linear index.
    #[must_use]
    pub fn get(self, index: usize) -> Option<OwnedValueView<'a>> {
        self.cells.get(index).map(OwnedElement::as_view)
    }

    /// Iterate over retained cells in row-major order.
    pub fn iter(self) -> impl Iterator<Item = OwnedValueView<'a>> + 'a {
        self.cells.iter().map(OwnedElement::as_view)
    }
}

/// Borrowed inspection of an owned ordered reference list.
#[derive(Clone, Copy, Debug)]
pub struct OwnedReferenceListView<'a> {
    records: &'a [OwnedReferenceRecord],
}

impl PartialEq for OwnedReferenceListView<'_> {
    fn eq(&self, other: &Self) -> bool {
        self.records == other.records
    }
}

impl<'a> OwnedReferenceListView<'a> {
    /// Number of records in source/operator order.
    #[must_use]
    pub const fn len(self) -> usize {
        self.records.len()
    }

    /// Whether no reference record was retained.
    #[must_use]
    pub const fn is_empty(self) -> bool {
        self.records.is_empty()
    }

    /// Return one record without cloning its lexical metadata or areas.
    #[must_use]
    pub fn get(self, index: usize) -> Option<super::ReferenceView<'a>> {
        self.records.get(index).map(OwnedReferenceRecord::as_view)
    }

    /// Iterate over records in source/operator order.
    pub fn iter(self) -> impl Iterator<Item = super::ReferenceView<'a>> + 'a {
        self.records.iter().map(OwnedReferenceRecord::as_view)
    }
}

impl OwnedEvaluated {
    /// Borrow the complete owned result without cloning retained storage.
    #[must_use]
    pub fn value(&self) -> OwnedValueView<'_> {
        self.value.as_view()
    }

    /// Return the result shape, if the result is an array.
    #[must_use]
    pub fn shape(&self) -> Option<Shape> {
        self.value.shape()
    }

    /// Return a borrowed view of an owned array, if present.
    #[must_use]
    pub fn as_array(&self) -> Option<OwnedArrayView<'_>> {
        match &self.value {
            OwnedValue::Array(array) => Some(OwnedArrayView {
                shape: array.shape,
                cells: &array.cells,
            }),
            OwnedValue::Empty
            | OwnedValue::Number(_)
            | OwnedValue::Logical(_)
            | OwnedValue::Text(_)
            | OwnedValue::Error(_)
            | OwnedValue::Reference(_)
            | OwnedValue::ReferenceList(_) => None,
        }
    }

    /// Return a borrowed single-reference view, if present.
    #[must_use]
    pub fn as_reference(&self) -> Option<super::ReferenceView<'_>> {
        match &self.value {
            OwnedValue::Reference(record) => Some(record.as_view()),
            OwnedValue::Empty
            | OwnedValue::Number(_)
            | OwnedValue::Logical(_)
            | OwnedValue::Text(_)
            | OwnedValue::Error(_)
            | OwnedValue::Array(_)
            | OwnedValue::ReferenceList(_) => None,
        }
    }

    /// Return a borrowed ordered reference-list view, if present.
    #[must_use]
    pub fn as_reference_list(&self) -> Option<OwnedReferenceListView<'_>> {
        match &self.value {
            OwnedValue::ReferenceList(list) => Some(OwnedReferenceListView {
                records: &list.records,
            }),
            OwnedValue::Empty
            | OwnedValue::Number(_)
            | OwnedValue::Logical(_)
            | OwnedValue::Text(_)
            | OwnedValue::Error(_)
            | OwnedValue::Array(_)
            | OwnedValue::Reference(_) => None,
        }
    }

    /// Return the exact retained-memory reservation amount.
    #[must_use]
    pub fn reserved_storage_bytes(&self) -> usize {
        self.reservation
            .as_ref()
            .and_then(|reservation| usize::try_from(reservation.amount()).ok())
            .unwrap_or(0)
    }

    /// Alias for [`Self::reserved_storage_bytes`], reporting the retained
    /// [`Resource::Memory`] charge rather than a separate output reservation.
    #[must_use]
    pub fn reserved_output_bytes(&self) -> usize {
        self.reserved_storage_bytes()
    }
}

/// Convert one borrowed evaluation result to lifetime-free owned storage.
///
/// The parent module wires this function into `Evaluated::to_owned`.  It is
/// `pub(super)` so the storage implementation remains private to this child
/// module while the parent can expose only the intended facade.
pub(super) fn from_evaluated<'a>(
    source: &Evaluated<'a>,
    execution: &ExecutionContext,
    limits: &Limits,
) -> EvaluationResult<OwnedEvaluated> {
    execution.check().map_err(map_execution_error)?;
    let mut measure = Measure::new(execution, limits);
    measure.value(&source.value)?;

    let bytes = measure.bytes;
    if bytes > limits.scalar_limits().max_storage_bytes() {
        return Err(local_limit(
            Resource::Memory,
            bytes,
            limits.scalar_limits().max_storage_bytes(),
        ));
    }
    let reservation = reserve_storage(bytes, execution)?;
    let value = copy_value(&source.value, execution, limits)?;
    execution.check().map_err(map_execution_error)?;
    Ok(OwnedEvaluated { value, reservation })
}

fn reserve_storage(
    bytes: usize,
    execution: &ExecutionContext,
) -> EvaluationResult<Option<Reservation>> {
    execution.check().map_err(map_execution_error)?;
    let reservation = if bytes == 0 {
        None
    } else {
        let core_limits = litchi_core::Limits::new(
            u64::try_from(bytes).unwrap_or(u64::MAX),
            u64::MAX,
            u64::MAX,
            u64::MAX,
            u64::MAX,
            u64::MAX,
        );
        let storage_budget = execution.budget().child(OWNED_SCOPE, core_limits);
        execution.check().map_err(map_execution_error)?;
        Some(
            storage_budget
                .reserve(Resource::Memory, u64::try_from(bytes).unwrap_or(u64::MAX))
                .map_err(EvaluationFailure::ResourceLimit)?,
        )
    };

    Ok(reservation)
}

#[derive(Debug)]
enum OwnedValue {
    Empty,
    Number(f64),
    Logical(bool),
    Text(String),
    Error(ScalarError),
    Array(OwnedArray),
    Reference(OwnedReferenceRecord),
    ReferenceList(OwnedReferenceList),
}

#[derive(Debug)]
struct OwnedArray {
    shape: Shape,
    cells: Vec<OwnedElement>,
}

#[derive(Debug, PartialEq)]
enum OwnedElement {
    Empty,
    Error(ScalarError),
    Number(f64),
    Logical(bool),
    Text(String),
}

#[derive(Debug, PartialEq)]
struct OwnedReferenceRecord {
    reference: Option<Reference>,
    areas: Vec<Area>,
}

#[derive(Debug, PartialEq)]
struct OwnedReferenceList {
    records: Vec<OwnedReferenceRecord>,
}

impl OwnedValue {
    fn as_view(&self) -> OwnedValueView<'_> {
        match self {
            Self::Empty => OwnedValueView::Empty,
            Self::Number(value) => OwnedValueView::Number(*value),
            Self::Logical(value) => OwnedValueView::Logical(*value),
            Self::Text(value) => OwnedValueView::Text(value),
            Self::Error(error) => OwnedValueView::Error(*error),
            Self::Array(array) => OwnedValueView::Array(OwnedArrayView {
                shape: array.shape,
                cells: &array.cells,
            }),
            Self::Reference(record) => OwnedValueView::Reference(record.as_view()),
            Self::ReferenceList(list) => OwnedValueView::ReferenceList(OwnedReferenceListView {
                records: &list.records,
            }),
        }
    }

    fn shape(&self) -> Option<Shape> {
        match self {
            Self::Array(array) => Some(array.shape),
            Self::Empty
            | Self::Number(_)
            | Self::Logical(_)
            | Self::Text(_)
            | Self::Error(_)
            | Self::Reference(_)
            | Self::ReferenceList(_) => None,
        }
    }
}

impl OwnedElement {
    fn as_view(&self) -> OwnedValueView<'_> {
        match self {
            Self::Empty => OwnedValueView::Empty,
            Self::Error(error) => OwnedValueView::Error(*error),
            Self::Number(value) => OwnedValueView::Number(*value),
            Self::Logical(value) => OwnedValueView::Logical(*value),
            Self::Text(value) => OwnedValueView::Text(value),
        }
    }
}

impl OwnedReferenceRecord {
    fn as_view(&self) -> super::ReferenceView<'_> {
        super::ReferenceView {
            reference: self.reference.as_ref(),
            areas: &self.areas,
        }
    }
}

/// The first pass computes exact requested capacities and charges the shared
/// work budget.  It deliberately does not allocate.
struct Measure<'a> {
    execution: &'a ExecutionContext,
    limits: &'a Limits,
    bytes: usize,
    steps: u64,
}

impl<'a> Measure<'a> {
    fn new(execution: &'a ExecutionContext, limits: &'a Limits) -> Self {
        Self {
            execution,
            limits,
            bytes: 0,
            steps: 0,
        }
    }

    fn item(&mut self) -> EvaluationResult<()> {
        self.work(1)
    }

    fn work(&mut self, amount: usize) -> EvaluationResult<()> {
        self.execution.check().map_err(map_execution_error)?;
        let amount = u64::try_from(amount).unwrap_or(u64::MAX);
        let next = self.steps.checked_add(amount).ok_or_else(|| {
            local_limit_u64(
                Resource::Work,
                u64::MAX,
                self.limits.scalar_limits().max_steps(),
            )
        })?;
        if next > self.limits.scalar_limits().max_steps() {
            return Err(local_limit_u64(
                Resource::Work,
                next,
                self.limits.scalar_limits().max_steps(),
            ));
        }
        if amount != 0 {
            self.execution
                .consume(Resource::Work, amount)
                .map_err(map_execution_error)?;
        }
        self.steps = next;
        Ok(())
    }

    fn capacity<T>(&mut self, count: usize) -> EvaluationResult<()> {
        if count == 0 {
            return Ok(());
        }
        let bytes =
            count
                .checked_mul(size_of::<T>())
                .ok_or(EvaluationFailure::InvalidExpression(
                    "owned value capacity overflows",
                ))?;
        self.bytes = self
            .bytes
            .checked_add(bytes)
            .ok_or(EvaluationFailure::InvalidExpression(
                "owned value storage size overflows",
            ))?;
        if self.bytes > self.limits.scalar_limits().max_storage_bytes() {
            return Err(local_limit(
                Resource::Memory,
                self.bytes,
                self.limits.scalar_limits().max_storage_bytes(),
            ));
        }
        self.item()
    }

    fn string(&mut self, value: &str, text_value: bool) -> EvaluationResult<()> {
        if text_value && value.len() > self.limits.scalar_limits().max_text_bytes() {
            return Err(local_limit(
                Resource::Memory,
                value.len(),
                self.limits.scalar_limits().max_text_bytes(),
            ));
        }
        self.bytes =
            self.bytes
                .checked_add(value.len())
                .ok_or(EvaluationFailure::InvalidExpression(
                    "owned value string size overflows",
                ))?;
        if self.bytes > self.limits.scalar_limits().max_storage_bytes() {
            return Err(local_limit(
                Resource::Memory,
                self.bytes,
                self.limits.scalar_limits().max_storage_bytes(),
            ));
        }
        let mut offset = 0usize;
        while offset < value.len() {
            let mut end = offset.saturating_add(COPY_CHUNK_BYTES).min(value.len());
            while end < value.len() && !value.is_char_boundary(end) {
                end += 1;
            }
            self.work(end - offset)?;
            offset = end;
        }
        self.item()
    }

    fn value(&mut self, value: &RuntimeValue<'_>) -> EvaluationResult<()> {
        self.item()?;
        match value {
            RuntimeValue::Empty | RuntimeValue::Missing => Ok(()),
            RuntimeValue::Scalar(value) => self.working(value, true),
            RuntimeValue::ScalarCell(_) => Err(EvaluationFailure::InvalidExpression(
                "scalar cell token escaped evaluation",
            )),
            RuntimeValue::Array(array) => {
                let expected =
                    array
                        .shape
                        .cell_count()
                        .ok_or(EvaluationFailure::InvalidExpression(
                            "owned array cell count overflows",
                        ))?;
                if expected != array.cells.len() {
                    return Err(EvaluationFailure::InvalidExpression(
                        "owned array shape and cell count disagree",
                    ));
                }
                if expected > self.limits.max_array_cells() {
                    return Err(local_limit(
                        Resource::Objects,
                        expected,
                        self.limits.max_array_cells(),
                    ));
                }
                self.capacity::<OwnedElement>(array.cells.len())?;
                for element in &array.cells {
                    self.element(element)?;
                }
                Ok(())
            },
            RuntimeValue::Areas(set) => {
                if set.records.len() > self.limits.max_reference_areas() {
                    return Err(local_limit(
                        Resource::Objects,
                        set.records.len(),
                        self.limits.max_reference_areas(),
                    ));
                }
                if set.areas.len() > self.limits.max_reference_areas() {
                    return Err(local_limit(
                        Resource::Objects,
                        set.areas.len(),
                        self.limits.max_reference_areas(),
                    ));
                }
                // A single non-list reference is stored inline in the
                // `OwnedValue::Reference` variant.  Reserve a record vector
                // only when the public result is a reference list; this
                // keeps the retained reservation aligned with live output
                // storage instead of charging a temporary vector that is
                // immediately popped.
                if set.is_list || set.records.len() != 1 {
                    self.capacity::<OwnedReferenceRecord>(set.records.len())?;
                }
                for record in &set.records {
                    self.record(record.reference, &record.areas)?;
                }
                Ok(())
            },
        }
    }

    fn working(&mut self, value: &WorkingValue<'_>, text_value: bool) -> EvaluationResult<()> {
        match value {
            WorkingValue::Number(_) | WorkingValue::Logical(_) | WorkingValue::Error(_) => {
                self.item()
            },
            WorkingValue::Text(value) => self.string(value.text.as_ref(), text_value),
        }
    }

    fn element(&mut self, element: &RuntimeElement<'_>) -> EvaluationResult<()> {
        self.item()?;
        match element {
            RuntimeElement::Empty | RuntimeElement::Missing => Ok(()),
            RuntimeElement::Present(value) => self.working(value, true),
        }
    }

    fn record(&mut self, reference: Option<&Reference>, areas: &[Area]) -> EvaluationResult<()> {
        self.item()?;
        if areas.len() > self.limits.max_reference_areas() {
            return Err(local_limit(
                Resource::Objects,
                areas.len(),
                self.limits.max_reference_areas(),
            ));
        }
        self.capacity::<Area>(areas.len())?;
        if let Some(reference) = reference {
            self.reference(reference)?;
        }
        for _ in areas {
            self.item()?;
        }
        Ok(())
    }

    fn reference(&mut self, reference: &Reference) -> EvaluationResult<()> {
        self.item()?;
        match reference {
            Reference::Error => Ok(()),
            Reference::Local(address) => self.address(address),
            Reference::Source { source, address } => {
                self.string(&source.iri, false)?;
                self.address(address)
            },
        }
    }

    fn address(&mut self, address: &Address) -> EvaluationResult<()> {
        self.item()?;
        match address {
            Address::Cell(endpoint) => self.endpoint(endpoint),
            Address::Cells(first, second)
            | Address::Columns(first, second)
            | Address::Rows(first, second) => {
                self.endpoint(first)?;
                self.endpoint(second)
            },
        }
    }

    fn endpoint(&mut self, endpoint: &Endpoint) -> EvaluationResult<()> {
        self.item()?;
        self.selector(&endpoint.sheet)?;
        match &endpoint.value {
            EndpointValue::Cell(cell) => self.cell(cell),
            EndpointValue::Column(column) => self.column(column),
            EndpointValue::Row(_) => self.item(),
        }
    }

    fn selector(&mut self, selector: &SheetSelector) -> EvaluationResult<()> {
        self.item()?;
        if let SheetSelector::Explicit(locator) = selector {
            self.locator(locator)?;
        }
        Ok(())
    }

    fn locator(&mut self, locator: &SheetLocator) -> EvaluationResult<()> {
        self.item()?;
        self.sheet_name(&locator.sheet)?;
        self.capacity::<Subtable>(locator.subtables.len())?;
        for subtable in &locator.subtables {
            self.item()?;
            match subtable {
                Subtable::Cell(cell) => self.cell(cell)?,
                Subtable::Name(name) => self.sheet_name(name)?,
            }
        }
        Ok(())
    }

    fn cell(&mut self, cell: &Cell) -> EvaluationResult<()> {
        self.item()?;
        self.column(&cell.column)?;
        self.item()
    }

    fn column(&mut self, column: &Column) -> EvaluationResult<()> {
        self.string(&column.label, false)
    }

    fn sheet_name(&mut self, name: &SheetName) -> EvaluationResult<()> {
        self.string(&name.name, false)
    }
}

fn copy_value(
    value: &RuntimeValue<'_>,
    execution: &ExecutionContext,
    limits: &Limits,
) -> EvaluationResult<OwnedValue> {
    execution.check().map_err(map_execution_error)?;
    match value {
        RuntimeValue::Empty => Ok(OwnedValue::Empty),
        RuntimeValue::Missing => Ok(OwnedValue::Error(ScalarError::Value)),
        RuntimeValue::Scalar(value) => copy_working(value, execution, limits),
        RuntimeValue::ScalarCell(_) => Err(EvaluationFailure::InvalidExpression(
            "scalar cell token escaped evaluation",
        )),
        RuntimeValue::Array(array) => {
            execution.check().map_err(map_execution_error)?;
            let mut cells =
                fallible_vec::<OwnedElement>(array.cells.len(), "formula owned array", execution)?;
            for element in &array.cells {
                execution.check().map_err(map_execution_error)?;
                cells.push(copy_element(element, execution, limits)?);
            }
            Ok(OwnedValue::Array(OwnedArray {
                shape: array.shape,
                cells,
            }))
        },
        RuntimeValue::Areas(set) => {
            if !set.is_list && set.records.len() == 1 {
                let record = set
                    .records
                    .first()
                    .ok_or(EvaluationFailure::InvalidExpression(
                        "owned reference record disappeared",
                    ))?;
                Ok(OwnedValue::Reference(copy_record(
                    record.reference,
                    &record.areas,
                    execution,
                    limits,
                )?))
            } else {
                execution.check().map_err(map_execution_error)?;
                let mut records = fallible_vec::<OwnedReferenceRecord>(
                    set.records.len(),
                    "formula owned references",
                    execution,
                )?;
                for record in &set.records {
                    execution.check().map_err(map_execution_error)?;
                    records.push(copy_record(
                        record.reference,
                        &record.areas,
                        execution,
                        limits,
                    )?);
                }
                Ok(OwnedValue::ReferenceList(OwnedReferenceList { records }))
            }
        },
    }
}

fn copy_working(
    value: &WorkingValue<'_>,
    execution: &ExecutionContext,
    limits: &Limits,
) -> EvaluationResult<OwnedValue> {
    match value {
        WorkingValue::Number(value) => Ok(OwnedValue::Number(*value)),
        WorkingValue::Logical(value) => Ok(OwnedValue::Logical(*value)),
        WorkingValue::Error(error) => Ok(OwnedValue::Error(*error)),
        WorkingValue::Text(value) => Ok(OwnedValue::Text(copy_string(
            value.text.as_ref(),
            true,
            execution,
            limits,
            "formula owned text",
        )?)),
    }
}

fn copy_element(
    element: &RuntimeElement<'_>,
    execution: &ExecutionContext,
    limits: &Limits,
) -> EvaluationResult<OwnedElement> {
    match element {
        RuntimeElement::Empty => Ok(OwnedElement::Empty),
        RuntimeElement::Missing => Ok(OwnedElement::Error(ScalarError::NotAvailable)),
        RuntimeElement::Present(WorkingValue::Number(value)) => Ok(OwnedElement::Number(*value)),
        RuntimeElement::Present(WorkingValue::Logical(value)) => Ok(OwnedElement::Logical(*value)),
        RuntimeElement::Present(WorkingValue::Error(error)) => Ok(OwnedElement::Error(*error)),
        RuntimeElement::Present(WorkingValue::Text(value)) => Ok(OwnedElement::Text(copy_string(
            value.text.as_ref(),
            true,
            execution,
            limits,
            "formula owned array text",
        )?)),
    }
}

fn copy_record(
    reference: Option<&Reference>,
    areas: &[Area],
    execution: &ExecutionContext,
    limits: &Limits,
) -> EvaluationResult<OwnedReferenceRecord> {
    execution.check().map_err(map_execution_error)?;
    let reference = reference
        .map(|reference| copy_reference(reference, execution, limits))
        .transpose()?;
    execution.check().map_err(map_execution_error)?;
    let mut owned_areas =
        fallible_vec::<Area>(areas.len(), "formula owned reference areas", execution)?;
    for area in areas {
        execution.check().map_err(map_execution_error)?;
        owned_areas.push(*area);
    }
    Ok(OwnedReferenceRecord {
        reference,
        areas: owned_areas,
    })
}

fn copy_reference(
    reference: &Reference,
    execution: &ExecutionContext,
    limits: &Limits,
) -> EvaluationResult<Reference> {
    execution.check().map_err(map_execution_error)?;
    match reference {
        Reference::Error => Ok(Reference::Error),
        Reference::Local(address) => {
            Ok(Reference::Local(copy_address(address, execution, limits)?))
        },
        Reference::Source { source, address } => Ok(Reference::Source {
            source: copy_source(source, execution, limits)?,
            address: copy_address(address, execution, limits)?,
        }),
    }
}

fn copy_source(
    source: &Source,
    execution: &ExecutionContext,
    limits: &Limits,
) -> EvaluationResult<Source> {
    Ok(Source {
        iri: copy_string(
            &source.iri,
            false,
            execution,
            limits,
            "formula owned source IRI",
        )?,
    })
}

fn copy_address(
    address: &Address,
    execution: &ExecutionContext,
    limits: &Limits,
) -> EvaluationResult<Address> {
    execution.check().map_err(map_execution_error)?;
    match address {
        Address::Cell(endpoint) => Ok(Address::Cell(copy_endpoint(endpoint, execution, limits)?)),
        Address::Cells(first, second) => Ok(Address::Cells(
            copy_endpoint(first, execution, limits)?,
            copy_endpoint(second, execution, limits)?,
        )),
        Address::Columns(first, second) => Ok(Address::Columns(
            copy_endpoint(first, execution, limits)?,
            copy_endpoint(second, execution, limits)?,
        )),
        Address::Rows(first, second) => Ok(Address::Rows(
            copy_endpoint(first, execution, limits)?,
            copy_endpoint(second, execution, limits)?,
        )),
    }
}

fn copy_endpoint(
    endpoint: &Endpoint,
    execution: &ExecutionContext,
    limits: &Limits,
) -> EvaluationResult<Endpoint> {
    execution.check().map_err(map_execution_error)?;
    let value = match &endpoint.value {
        EndpointValue::Cell(cell) => EndpointValue::Cell(copy_cell(cell, execution, limits)?),
        EndpointValue::Column(column) => {
            EndpointValue::Column(copy_column(column, execution, limits)?)
        },
        EndpointValue::Row(row) => EndpointValue::Row(Row {
            number: row.number,
            absolute: row.absolute,
        }),
    };
    Ok(Endpoint {
        sheet: copy_selector(&endpoint.sheet, execution, limits)?,
        value,
    })
}

fn copy_selector(
    selector: &SheetSelector,
    execution: &ExecutionContext,
    limits: &Limits,
) -> EvaluationResult<SheetSelector> {
    execution.check().map_err(map_execution_error)?;
    match selector {
        SheetSelector::Current => Ok(SheetSelector::Current),
        SheetSelector::Inherited => Ok(SheetSelector::Inherited),
        SheetSelector::Explicit(locator) => Ok(SheetSelector::Explicit(copy_locator(
            locator, execution, limits,
        )?)),
    }
}

fn copy_locator(
    locator: &SheetLocator,
    execution: &ExecutionContext,
    limits: &Limits,
) -> EvaluationResult<SheetLocator> {
    execution.check().map_err(map_execution_error)?;
    let mut subtables = fallible_vec::<Subtable>(
        locator.subtables.len(),
        "formula owned subtables",
        execution,
    )?;
    for subtable in &locator.subtables {
        execution.check().map_err(map_execution_error)?;
        subtables.push(match subtable {
            Subtable::Cell(cell) => Subtable::Cell(copy_cell(cell, execution, limits)?),
            Subtable::Name(name) => Subtable::Name(copy_sheet_name(name, execution, limits)?),
        });
    }
    Ok(SheetLocator {
        sheet: copy_sheet_name(&locator.sheet, execution, limits)?,
        subtables,
    })
}

fn copy_sheet_name(
    name: &SheetName,
    execution: &ExecutionContext,
    limits: &Limits,
) -> EvaluationResult<SheetName> {
    Ok(SheetName {
        name: copy_string(
            &name.name,
            false,
            execution,
            limits,
            "formula owned sheet name",
        )?,
        absolute: name.absolute,
        quoted: name.quoted,
    })
}

fn copy_cell(cell: &Cell, execution: &ExecutionContext, limits: &Limits) -> EvaluationResult<Cell> {
    Ok(Cell {
        column: copy_column(&cell.column, execution, limits)?,
        row: Row {
            number: cell.row.number,
            absolute: cell.row.absolute,
        },
    })
}

fn copy_column(
    column: &Column,
    execution: &ExecutionContext,
    limits: &Limits,
) -> EvaluationResult<Column> {
    Ok(Column {
        label: copy_string(
            &column.label,
            false,
            execution,
            limits,
            "formula owned column label",
        )?,
        absolute: column.absolute,
    })
}

fn copy_string(
    value: &str,
    text_value: bool,
    execution: &ExecutionContext,
    limits: &Limits,
    resource: &'static str,
) -> EvaluationResult<String> {
    if text_value && value.len() > limits.scalar_limits().max_text_bytes() {
        return Err(local_limit(
            Resource::Memory,
            value.len(),
            limits.scalar_limits().max_text_bytes(),
        ));
    }
    let mut output = String::new();
    if !value.is_empty() {
        execution.check().map_err(map_execution_error)?;
        output
            .try_reserve_exact(value.len())
            .map_err(|source| EvaluationFailure::Allocation { resource, source })?;
    }
    let mut offset = 0usize;
    while offset < value.len() {
        execution.check().map_err(map_execution_error)?;
        let mut end = offset.saturating_add(COPY_CHUNK_BYTES).min(value.len());
        while end < value.len() && !value.is_char_boundary(end) {
            end += 1;
        }
        output.push_str(&value[offset..end]);
        offset = end;
    }
    Ok(output)
}

fn fallible_vec<T>(
    length: usize,
    resource: &'static str,
    execution: &ExecutionContext,
) -> EvaluationResult<Vec<T>> {
    let mut result = Vec::new();
    if length != 0 {
        execution.check().map_err(map_execution_error)?;
        result
            .try_reserve_exact(length)
            .map_err(|source| EvaluationFailure::Allocation { resource, source })?;
    }
    Ok(result)
}

fn local_limit(resource: Resource, observed: usize, limit: usize) -> EvaluationFailure {
    local_limit_u64(
        resource,
        u64::try_from(observed).unwrap_or(u64::MAX),
        u64::try_from(limit).unwrap_or(u64::MAX),
    )
}

fn local_limit_u64(resource: Resource, observed: u64, limit: u64) -> EvaluationFailure {
    EvaluationFailure::ResourceLimit(ResourceLimit {
        resource,
        observed,
        limit,
        scope: Arc::from(OWNED_SCOPE),
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn cancellation_after_measurement_prevents_storage_admission_and_copy_allocation() {
        use litchi_core::{Budget, CancellationSource, ExecutionLimits, Profile};
        use std::num::{NonZeroU64, NonZeroUsize};

        let budget = Budget::root(
            "owned-copy-cancellation",
            litchi_core::Limits::for_profile(Profile::Server),
        );
        let (cancellation, token) = CancellationSource::pair();
        let execution = ExecutionContext::new(
            budget.clone(),
            token,
            ExecutionLimits::new(NonZeroUsize::MIN, NonZeroUsize::MIN, NonZeroU64::MIN, 0)
                .expect("valid serial execution limits"),
        );
        let limits = Limits::default();
        let mut measure = Measure::new(&execution, &limits);
        measure.string("retained text", true).expect("measurement");
        cancellation.cancel();

        assert!(matches!(
            reserve_storage(measure.bytes, &execution),
            Err(EvaluationFailure::Cancelled)
        ));
        assert!(matches!(
            copy_string("retained text", true, &execution, &limits, "test text"),
            Err(EvaluationFailure::Cancelled)
        ));
        assert!(matches!(
            fallible_vec::<u8>(16, "test vector", &execution),
            Err(EvaluationFailure::Cancelled)
        ));
        assert_eq!(budget.used(Resource::Memory), 0);
    }

    #[test]
    fn owned_element_views_keep_text_borrowed_from_owned_storage() {
        let element = OwnedElement::Text("persistent".to_owned());
        assert_eq!(element.as_view(), OwnedValueView::Text("persistent"));
    }

    #[test]
    fn reference_record_view_retains_areas_and_optional_ast() {
        let record = OwnedReferenceRecord {
            reference: Some(Reference::Error),
            areas: vec![
                Area::from_bounds([0, 0, 0], [1, 1, 1]).expect("nonempty area"),
                Area::from_bounds([0, 0, 0], [1, 1, 1]).expect("duplicate area"),
            ],
        };
        let view = record.as_view();
        assert_eq!(view.reference(), Some(&Reference::Error));
        assert_eq!(view.areas().len(), 2);
        assert_eq!(view.areas()[0], view.areas()[1]);
    }
}
