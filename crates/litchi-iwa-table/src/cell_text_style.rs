//! Shared whole-cell character-style graph planning.
//!
//! This module is the format-neutral substrate for the Numbers whole-cell
//! `CharacterProperties` family.  It deliberately stops at a neutral graph
//! plan: the leaf owns identifier allocation, archive publication, metadata
//! transitions, exact-source reassembly and candidate reopen/readback.

use std::cell::Cell;
use std::collections::{BTreeMap, BTreeSet};
use std::fmt;

use litchi_iwa_common::wire::WireView;
use litchi_iwa_common::wire::{RawWireFields, RawWireLimits};
use litchi_iwa_common::{LimitKind, WireLimits};
use litchi_iwa_text_wire::CharacterProperties;
use litchi_numbers_wire::{BncCell, Limits as BncLimits, table_topology::storage};

use crate::TableSource;

const TABLE_MODEL_MESSAGE_TYPE: u32 = 6_001;
const TILE_MESSAGE_TYPE: u32 = 6_002;
const LIST_MESSAGE_TYPE: u32 = 6_005;
const LIST_SEGMENT_MESSAGE_TYPE: u32 = 6_201;
const PARAGRAPH_STYLE_MESSAGE_TYPE: u32 = 2_022;
const STYLESHEET_MESSAGE_TYPE: u32 = 401;
const STYLE_LIST_FIELD: u32 = 5;
const PARAGRAPH_PROPERTIES_FIELD: u32 = 12;
const MODEL_DATA_STORE: u32 = 4;
const MODEL_ROWS: u32 = 6;
const MODEL_COLUMNS: u32 = 7;
const MODEL_HEADER_ROWS: u32 = 9;
const MODEL_HEADER_COLUMNS: u32 = 10;
const MODEL_FOOTER_ROWS: u32 = 11;
const STORE_TILES: u32 = 3;

/// One table cell selected by a leaf facade.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct Target {
    /// The table-model object identifier.
    pub model_identifier: u64,
    /// Zero-based row.
    pub row: u32,
    /// Zero-based column.
    pub column: u32,
}

/// Refusal from the shared character-style graph.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Refusal {
    /// The selected cell is outside the table geometry.
    OutsideTable,
    /// The style list is segmented or otherwise not owned by this planner.
    UnsupportedStyleList,
    /// The style graph has an external or ambiguous owner.
    UnsupportedDependency,
    /// The source lacks the required table structure.
    UnsupportedSource,
}

/// A typed, content-redacted shared planner failure.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum Error {
    /// The graph is malformed or inconsistent.
    InvalidSource(&'static str),
    /// A well-formed source is outside this engine's ownership profile.
    Refused(Refusal),
    /// A bounded wire/resource operation exceeded its caller's profile.
    Limit {
        /// Resource dimension.
        kind: LimitKind,
        /// Observed amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded allocation failed.
    Allocation {
        /// Requested amount.
        amount: usize,
    },
    /// The plan requires a leaf-supplied fresh identifier.
    NeedsIdentifier,
    /// The request shape is left to the leaf compatibility publisher.
    BatchUnsupported,
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::InvalidSource(reason) => {
                write!(formatter, "invalid cell text-style source: {reason}")
            },
            Self::Refused(reason) => write!(formatter, "cell text-style graph refused: {reason:?}"),
            Self::Limit {
                kind,
                observed,
                maximum,
            } => {
                write!(
                    formatter,
                    "{kind} limit exceeded: observed {observed}, maximum {maximum}"
                )
            },
            Self::Allocation { amount } => {
                write!(
                    formatter,
                    "could not allocate {amount} cell text-style items"
                )
            },
            Self::NeedsIdentifier => {
                formatter.write_str("cell text-style plan needs a fresh identifier")
            },
            Self::BatchUnsupported => {
                formatter.write_str("cell text-style batch shape is not yet shared")
            },
        }
    }
}

impl std::error::Error for Error {}

/// Caller-selected residual budget for one whole-cell style transaction.
#[derive(Debug)]
pub struct Limits {
    /// Wire input/output limits derived from the opened archive profile.
    pub wire: WireLimits,
    /// Maximum object inventory admitted by the leaf.
    pub max_objects: usize,
    /// Maximum references admitted by the leaf.
    pub max_references: usize,
    /// Maximum cumulative planner work.
    pub max_work: usize,
    /// Maximum parent-chain depth.
    pub max_style_depth: usize,
    /// Maximum fallible collection allocations admitted by this operation.
    pub max_allocations: usize,
    /// Maximum bytes copied into operation-owned graph values.
    pub max_copies: usize,
    remaining_input: Cell<usize>,
    remaining_output: Cell<usize>,
    remaining_fields: Cell<usize>,
    remaining_work: Cell<usize>,
    remaining_references: Cell<usize>,
    remaining_allocations: Cell<usize>,
    remaining_copies: Cell<usize>,
}

impl Limits {
    /// Pre-admit the caller's residual package budget before graph traversal.
    pub fn from_source(source: &dyn TableSource) -> Result<Self, Error> {
        let max_objects = source.max_objects();
        let max_references = source.max_references();
        if max_objects == 0 {
            return Err(Error::Limit {
                kind: LimitKind::MaterializedCells,
                observed: 1,
                maximum: 0,
            });
        }
        if max_references == 0 {
            return Err(Error::Limit {
                kind: LimitKind::Fields,
                observed: 1,
                maximum: 0,
            });
        }
        let profile = source.archive_limits().archive_limits();
        let input = profile.max_message_bytes().min(WireLimits::MAX_INPUT_BYTES);
        let fields = profile.max_header_fields().min(WireLimits::MAX_FIELDS);
        let nesting = profile.max_header_nesting().min(WireLimits::MAX_NESTING);
        let max_work = profile
            .max_header_memory_bytes()
            .min(WireLimits::MAX_REWRITE_WORK);
        let output = profile
            .max_message_bytes()
            .min(WireLimits::MAX_OUTPUT_BYTES);
        if input == 0 || fields == 0 || nesting == 0 || max_work == 0 || output == 0 {
            return Err(Error::Limit {
                kind: if input == 0 {
                    LimitKind::InputBytes
                } else if fields == 0 {
                    LimitKind::Fields
                } else if nesting == 0 {
                    LimitKind::Nesting
                } else if output == 0 {
                    LimitKind::OutputBytes
                } else {
                    LimitKind::RewriteWork
                },
                observed: 1,
                maximum: 0,
            });
        }
        let wire = WireLimits::default()
            .with_input_bytes(input)
            .and_then(|limits| limits.with_fields(fields))
            .and_then(|limits| limits.with_output_bytes(output))
            .and_then(|limits| limits.with_nesting(nesting))
            .and_then(|limits| limits.with_rewrite_work(max_work))
            .map_err(|_| invalid("source message limits are invalid"))?;
        let mut objects = 0_usize;
        let mut references = 0_usize;
        let mut work = 0_usize;
        for component in source.catalog().iter() {
            for object in &component.archive().objects {
                objects = objects.checked_add(1).ok_or(Error::Limit {
                    kind: LimitKind::MaterializedCells,
                    observed: u64::MAX,
                    maximum: max_objects as u64,
                })?;
                if objects > max_objects {
                    return Err(Error::Limit {
                        kind: LimitKind::MaterializedCells,
                        observed: objects as u64,
                        maximum: max_objects as u64,
                    });
                }
                for info in &object.archive_info.message_infos {
                    let mut count = info.object_references.len();
                    for field in &info.field_infos {
                        count = count.checked_add(field.object_references.len()).ok_or(
                            Error::Limit {
                                kind: LimitKind::Fields,
                                observed: u64::MAX,
                                maximum: max_references as u64,
                            },
                        )?;
                    }
                    references = references.checked_add(count).ok_or(Error::Limit {
                        kind: LimitKind::Fields,
                        observed: u64::MAX,
                        maximum: max_references as u64,
                    })?;
                    if references > max_references {
                        return Err(Error::Limit {
                            kind: LimitKind::Fields,
                            observed: references as u64,
                            maximum: max_references as u64,
                        });
                    }
                    work = work.saturating_add(1).saturating_add(count);
                    if work > max_work {
                        return Err(Error::Limit {
                            kind: LimitKind::RewriteWork,
                            observed: work as u64,
                            maximum: max_work as u64,
                        });
                    }
                }
            }
        }
        Ok(Self {
            wire,
            max_objects,
            max_references,
            max_work: wire.max_rewrite_work(),
            max_style_depth: nesting,
            max_allocations: max_objects.saturating_add(max_references).max(1),
            max_copies: wire.max_output_bytes().max(1),
            remaining_input: Cell::new(wire.max_input_bytes()),
            remaining_output: Cell::new(wire.max_output_bytes()),
            remaining_fields: Cell::new(wire.max_fields()),
            remaining_work: Cell::new(wire.max_rewrite_work()),
            remaining_references: Cell::new(max_references),
            remaining_allocations: Cell::new(max_objects.saturating_add(max_references).max(1)),
            remaining_copies: Cell::new(wire.max_output_bytes().max(1)),
        })
    }

    fn charge(
        remaining: &Cell<usize>,
        maximum: usize,
        amount: usize,
        kind: LimitKind,
    ) -> Result<(), Error> {
        if amount > remaining.get() {
            let observed = maximum
                .saturating_sub(remaining.get())
                .saturating_add(amount);
            return Err(Error::Limit {
                kind,
                observed: observed as u64,
                maximum: maximum as u64,
            });
        }
        remaining.set(remaining.get() - amount);
        Ok(())
    }

    pub(crate) fn charge_input(&self, amount: usize) -> Result<(), Error> {
        Self::charge(
            &self.remaining_input,
            self.wire.max_input_bytes(),
            amount,
            LimitKind::InputBytes,
        )
    }

    pub(crate) fn charge_fields(&self, amount: usize) -> Result<(), Error> {
        Self::charge(
            &self.remaining_fields,
            self.wire.max_fields(),
            amount,
            LimitKind::Fields,
        )
    }

    fn charge_output(&self, amount: usize) -> Result<(), Error> {
        Self::charge(
            &self.remaining_output,
            self.wire.max_output_bytes(),
            amount,
            LimitKind::OutputBytes,
        )
    }

    fn prepare_output(&self, bound: usize) -> Result<WireLimits, Error> {
        let wire = self.residual_wire()?;
        self.charge_output(bound)?;
        self.admit_copies(bound)?;
        self.admit_allocations(1)?;
        Ok(wire)
    }

    pub(crate) fn residual_wire(&self) -> Result<WireLimits, Error> {
        let input = self.remaining_input.get();
        let fields = self.remaining_fields.get();
        let output = self.remaining_output.get();
        let work = self.remaining_work.get();
        if input == 0 {
            return Err(Error::Limit {
                kind: LimitKind::InputBytes,
                observed: 1,
                maximum: 0,
            });
        }
        if fields == 0 {
            return Err(Error::Limit {
                kind: LimitKind::Fields,
                observed: 1,
                maximum: 0,
            });
        }
        if output == 0 {
            return Err(Error::Limit {
                kind: LimitKind::OutputBytes,
                observed: 1,
                maximum: 0,
            });
        }
        if work == 0 {
            return Err(Error::Limit {
                kind: LimitKind::RewriteWork,
                observed: 1,
                maximum: 0,
            });
        }
        self.wire
            .with_input_bytes(input)
            .and_then(|value| value.with_fields(fields))
            .and_then(|value| value.with_output_bytes(output))
            .and_then(|value| value.with_rewrite_work(work))
            .map_err(map_wire)
    }

    fn charge_work(&self, amount: usize) -> Result<(), Error> {
        Self::charge(
            &self.remaining_work,
            self.max_work,
            amount,
            LimitKind::RewriteWork,
        )
    }

    /// Debit bounded writer/traversal work owned by a leaf publication phase.
    pub fn admit_work(&self, amount: usize) -> Result<(), Error> {
        self.charge_work(amount)
    }

    fn charge_references(&self, amount: usize) -> Result<(), Error> {
        Self::charge(
            &self.remaining_references,
            self.max_references,
            amount,
            LimitKind::Fields,
        )
    }

    /// Debit references that a leaf publication will add to archive headers.
    pub fn admit_references(&self, amount: usize) -> Result<(), Error> {
        self.charge_references(amount)
    }

    /// Debit serialized bytes a leaf writer will emit after the shared plan.
    pub fn admit_output(&self, amount: usize) -> Result<(), Error> {
        self.charge_output(amount)
    }

    /// Admit `amount` fallible collection allocations before the operation
    /// asks a decoder or publisher to reserve them.
    pub fn admit_allocations(&self, amount: usize) -> Result<(), Error> {
        if amount > self.remaining_allocations.get() {
            return Err(Error::Allocation { amount });
        }
        self.remaining_allocations
            .set(self.remaining_allocations.get() - amount);
        Ok(())
    }

    /// Admit copied bytes before retaining a second owned graph value.
    pub fn admit_copies(&self, amount: usize) -> Result<(), Error> {
        if amount > self.remaining_copies.get() {
            return Err(Error::Allocation { amount });
        }
        self.remaining_copies
            .set(self.remaining_copies.get() - amount);
        Ok(())
    }

    /// Admit the owned values copied while staging or inheriting character
    /// properties. Primitive scalar fields do not allocate; font names do.
    pub fn admit_character_properties(
        &self,
        properties: &CharacterProperties,
    ) -> Result<(), Error> {
        if let Some(font_name) = properties.font_name.as_deref() {
            self.admit_allocations(1)?;
            self.admit_copies(font_name.len())?;
        }
        Ok(())
    }

    /// Reserve a bounded collection growth before pushing or extending it.
    pub fn reserve<T>(&self, values: &mut Vec<T>, additional: usize) -> Result<(), Error> {
        let spare = values.capacity().saturating_sub(values.len());
        let needed = additional.saturating_sub(spare);
        if needed == 0 {
            return Ok(());
        }
        self.admit_allocations(needed)?;
        values
            .try_reserve(needed)
            .map_err(|_| Error::Allocation { amount: needed })
    }

    /// Return the cumulative operation work admitted so far.
    #[must_use]
    pub fn report(&self) -> Report {
        Report {
            fields: self
                .wire
                .max_fields()
                .saturating_sub(self.remaining_fields.get()),
            references: self
                .max_references
                .saturating_sub(self.remaining_references.get()),
            work: self.max_work.saturating_sub(self.remaining_work.get()),
        }
    }

    fn copy(&self, bytes: &[u8]) -> Result<Vec<u8>, Error> {
        self.admit_copies(bytes.len())?;
        self.admit_allocations(1)?;
        let mut output = Vec::new();
        output
            .try_reserve(bytes.len())
            .map_err(|_| Error::Allocation {
                amount: bytes.len(),
            })?;
        output.extend_from_slice(bytes);
        Ok(output)
    }
}

/// Bounded work accounting returned with a shared plan.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct Report {
    /// Parsed wire fields.
    pub fields: usize,
    /// References admitted by the graph plan.
    pub references: usize,
    /// Cumulative planner work.
    pub work: usize,
}

fn invalid(reason: &'static str) -> Error {
    Error::InvalidSource(reason)
}

fn map_wire(error: litchi_iwa_common::Error) -> Error {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => Error::Limit {
            kind,
            observed: observed as u64,
            maximum: limit as u64,
        },
        litchi_iwa_common::Error::Allocation { amount, .. } => Error::Allocation { amount },
        _ => invalid("malformed character-style wire payload"),
    }
}

fn view<'a>(data: &'a [u8], limits: &Limits) -> Result<WireView<'a>, Error> {
    let wire = limits.residual_wire()?;
    let remaining_input = limits.remaining_input.get();
    let remaining_fields = limits.remaining_fields.get();
    let remaining_work = limits.remaining_work.get();
    if data.len() > remaining_input {
        return Err(Error::Limit {
            kind: LimitKind::InputBytes,
            observed: data.len() as u64,
            maximum: remaining_input as u64,
        });
    }
    if remaining_fields == 0 && !data.is_empty() {
        return Err(Error::Limit {
            kind: LimitKind::Fields,
            observed: 1,
            maximum: 0,
        });
    }
    if remaining_work == 0 && !data.is_empty() {
        return Err(Error::Limit {
            kind: LimitKind::RewriteWork,
            observed: 1,
            maximum: 0,
        });
    }
    let raw_limits = RawWireLimits::new(
        remaining_input.max(1),
        remaining_fields.max(1),
        wire.max_nesting(),
        remaining_work.max(1),
    )
    .map_err(map_wire)?;
    let mut fields = RawWireFields::new(data, raw_limits);
    while fields.next().map_err(map_wire)?.is_some() {}
    limits.charge_input(data.len())?;
    limits.charge_fields(fields.fields())?;
    limits.charge_work(fields.work())?;
    limits.admit_allocations(fields.fields())?;
    // RawWireFields and WireView each walk the source.  The latter also
    // materializes source-bound spans, so debit its independent input,
    // field, and traversal work before invoking the parser.
    limits.charge_input(data.len())?;
    limits.charge_fields(fields.fields())?;
    limits.charge_work(fields.fields())?;
    WireView::parse(data, wire).map_err(map_wire)
}

fn admit_decoder_view(data_len: usize, fields: usize, limits: &Limits) -> Result<(), Error> {
    // litchi-iwa-text-wire's private Budget creates another WireView for each
    // selected message.  Account for that complete pass and its span vector
    // before the decoder starts.
    limits.charge_input(data_len)?;
    limits.charge_fields(fields)?;
    limits.charge_work(fields)?;
    limits.admit_allocations(fields)
}

fn preflight_character_properties(data: &[u8], limits: &Limits) -> Result<WireLimits, Error> {
    let wire = limits.residual_wire()?;
    let root = view(data, limits)?;
    admit_decoder_view(data.len(), root.len(), limits)?;
    for field in root.fields() {
        if field.number() != 11 || field.wire_type() != 2 {
            continue;
        }
        let properties = view(field.payload(), limits)?;
        admit_decoder_view(field.payload().len(), properties.len(), limits)?;
        for property in properties.fields() {
            match (property.number(), property.wire_type()) {
                (5, 2) => {
                    // character_style_properties() owns this String before
                    // it returns, so reserve its resource units first.
                    limits.charge_output(property.payload().len())?;
                    limits.admit_allocations(1)?;
                    limits.admit_copies(property.payload().len())?;
                },
                (7, 2) => {
                    // The decoder descends into each colour message with a
                    // separate bounded view and span allocation.
                    let color = view(property.payload(), limits)?;
                    admit_decoder_view(property.payload().len(), color.len(), limits)?;
                },
                _ => {},
            }
        }
    }
    Ok(wire)
}

fn preflight_character_parent(data: &[u8], limits: &Limits) -> Result<WireLimits, Error> {
    let wire = limits.residual_wire()?;
    let root = view(data, limits)?;
    admit_decoder_view(data.len(), root.len(), limits)?;
    for field in root.fields() {
        if field.number() != 1 || field.wire_type() != 2 {
            continue;
        }
        let base = view(field.payload(), limits)?;
        admit_decoder_view(field.payload().len(), base.len(), limits)?;
        for base_field in base.fields() {
            if base_field.number() == 3 && base_field.wire_type() == 2 {
                // reference_identifier() performs this source-bound view.
                let reference = view(base_field.payload(), limits)?;
                admit_decoder_view(base_field.payload().len(), reference.len(), limits)?;
                limits.charge_references(1)?;
            }
        }
    }
    Ok(wire)
}

fn map_char_style(error: litchi_iwa_text_wire::CharStyleError) -> Error {
    match error {
        litchi_iwa_text_wire::CharStyleError::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind,
            observed: observed as u64,
            maximum: maximum as u64,
        },
        litchi_iwa_text_wire::CharStyleError::Allocation { amount } => Error::Allocation { amount },
        litchi_iwa_text_wire::CharStyleError::Invalid
        | litchi_iwa_text_wire::CharStyleError::Span => {
            invalid("malformed character-style properties")
        },
        _ => invalid("malformed character-style properties"),
    }
}

fn varint(data: &[u8], number: u32, limits: &Limits) -> Result<Option<u64>, Error> {
    let mut found = None;
    for field in view(data, limits)?.fields() {
        if field.number() != number {
            continue;
        }
        if field.wire_type() != 0 || found.is_some() {
            return Err(invalid(
                "a singular varint field repeats or has the wrong wire type",
            ));
        }
        let (value, _) = litchi_iwa_common::decode_varint_from_bytes(field.payload())
            .map_err(|_| invalid("malformed varint"))?;
        found = Some(value);
    }
    Ok(found)
}

fn bytes<'a>(data: &'a [u8], number: u32, limits: &Limits) -> Result<Option<&'a [u8]>, Error> {
    let mut found = None;
    for field in view(data, limits)?.fields() {
        if field.number() != number {
            continue;
        }
        if field.wire_type() != 2 || found.is_some() {
            return Err(invalid(
                "a singular bytes field repeats or has the wrong wire type",
            ));
        }
        found = Some(field.payload());
    }
    Ok(found)
}

fn reference(data: &[u8], number: u32, limits: &Limits) -> Result<Option<u64>, Error> {
    let Some(payload) = bytes(data, number, limits)? else {
        return Ok(None);
    };
    limits.charge_references(1)?;
    varint(payload, 1, limits)
}

fn repeated_payloads<'a>(
    data: &'a [u8],
    number: u32,
    limits: &Limits,
) -> Result<Vec<&'a [u8]>, Error> {
    let mut payloads = Vec::new();
    for field in view(data, limits)?.fields() {
        if field.number() != number {
            continue;
        }
        if field.wire_type() != 2 {
            return Err(invalid("repeated field has the wrong wire type"));
        }
        limits.reserve(&mut payloads, 1)?;
        payloads.push(field.payload());
    }
    Ok(payloads)
}

fn unique_payload<'a>(
    source: &'a dyn TableSource,
    identifier: u64,
    types: &[u32],
    limits: &Limits,
) -> Result<(u32, &'a [u8]), Error> {
    limits.charge_work(1)?;
    let object = source
        .resolve_object(identifier)
        .map_err(|_| invalid("object index is inconsistent"))?
        .ok_or(invalid("referenced object is missing"))?;
    limits.charge_work(object.messages.len())?;
    let mut found = None;
    for message in object.messages {
        if types.contains(&message.type_) {
            if found.is_some() {
                return Err(invalid("object repeats a required payload"));
            }
            found = Some((message.type_, message.data.as_slice()));
        }
    }
    found.ok_or(invalid("object lacks a required payload"))
}

#[derive(Debug, Clone, Copy)]
struct Geometry {
    rows: u32,
    columns: u32,
    header_rows: u32,
    header_columns: u32,
    footer_rows: u32,
}

impl Geometry {
    fn read(model: &[u8], limits: &Limits) -> Result<Self, Error> {
        let value = |field| -> Result<u32, Error> {
            u32::try_from(
                varint(model, field, limits)?.ok_or(invalid("table geometry field missing"))?,
            )
            .map_err(|_| invalid("table geometry exceeds u32"))
        };
        let geometry = Self {
            rows: value(MODEL_ROWS)?,
            columns: value(MODEL_COLUMNS)?,
            header_rows: u32::try_from(varint(model, MODEL_HEADER_ROWS, limits)?.unwrap_or(0))
                .map_err(|_| invalid("header rows exceed u32"))?,
            header_columns: u32::try_from(
                varint(model, MODEL_HEADER_COLUMNS, limits)?.unwrap_or(0),
            )
            .map_err(|_| invalid("header columns exceed u32"))?,
            footer_rows: u32::try_from(varint(model, MODEL_FOOTER_ROWS, limits)?.unwrap_or(0))
                .map_err(|_| invalid("footer rows exceed u32"))?,
        };
        if geometry.header_rows + geometry.footer_rows > geometry.rows
            || geometry.header_columns > geometry.columns
        {
            return Err(invalid("table geometry ranges overlap"));
        }
        Ok(geometry)
    }
}

struct Table {
    model: Vec<u8>,
    store: Vec<u8>,
    style_list: u64,
    geometry: Geometry,
}

impl Table {
    fn read(source: &dyn TableSource, target: Target, limits: &Limits) -> Result<Self, Error> {
        let (_, model) = unique_payload(
            source,
            target.model_identifier,
            &[TABLE_MODEL_MESSAGE_TYPE],
            limits,
        )?;
        let geometry = Geometry::read(model, limits)?;
        if target.row >= geometry.rows || target.column >= geometry.columns {
            return Err(Error::Refused(Refusal::OutsideTable));
        }
        let store =
            bytes(model, MODEL_DATA_STORE, limits)?.ok_or(invalid("table data store missing"))?;
        let store = limits.copy(store)?;
        let style_list = reference(&store, STYLE_LIST_FIELD, limits)?
            .ok_or(Error::Refused(Refusal::UnsupportedDependency))?;
        Ok(Self {
            model: limits.copy(model)?,
            store,
            style_list,
            geometry,
        })
    }
}

/// Bounded, source-relative table facts needed by a leaf tile writer.
#[derive(Debug, Clone)]
pub struct TableFacts {
    /// Table row count.
    pub rows: u32,
    /// Table column count.
    pub columns: u32,
    /// Rows per storage tile.
    pub tile_size: u64,
    /// Row-header bucket object identifier.
    pub row_headers: u64,
    /// Column-header bucket object identifier.
    pub column_headers: u64,
    /// Storage tile index and object identifier pairs.
    pub tiles: Vec<(u32, u64)>,
}

/// Read the immutable table geometry and tile index through the shared
/// residual budget. Leaves use these facts for selection; they must not route
/// the same admission through a default-limit topology parser.
pub fn table_facts(
    source: &dyn TableSource,
    target: Target,
    limits: &Limits,
) -> Result<TableFacts, Error> {
    let table = Table::read(source, target, limits)?;
    let row_headers =
        reference(&table.store, 1, limits)?.ok_or(invalid("row-header storage missing"))?;
    let column_headers =
        reference(&table.store, 2, limits)?.ok_or(invalid("column-header storage missing"))?;
    let tile_storage =
        bytes(&table.store, STORE_TILES, limits)?.ok_or(invalid("tile storage missing"))?;
    let tile_size = varint(tile_storage, 2, limits)?.ok_or(invalid("tile size missing"))?;
    if tile_size == 0 {
        return Err(invalid("tile size is zero"));
    }
    let mut tiles = Vec::new();
    for entry in repeated_payloads(tile_storage, 1, limits)? {
        limits.reserve(&mut tiles, 1)?;
        let key = u32::try_from(varint(entry, 1, limits)?.ok_or(invalid("tile key missing"))?)
            .map_err(|_| invalid("tile key exceeds u32"))?;
        let tile = reference(entry, 2, limits)?.ok_or(invalid("tile reference missing"))?;
        if tiles.iter().any(|(existing, _)| *existing == key) {
            return Err(invalid("tile key repeats"));
        }
        tiles.push((key, tile));
    }
    Ok(TableFacts {
        rows: table.geometry.rows,
        columns: table.geometry.columns,
        tile_size,
        row_headers,
        column_headers,
        tiles,
    })
}

struct LocatedCell {
    bytes: Vec<u8>,
}

/// Immutable data prepared once for a multi-cell style batch.  The batch
/// planner keeps the table model, style list, and selected cell payloads
/// source-bound so each request does not repeat the same full tile walk.
struct BatchContext {
    table: Table,
    style_list: StyleList,
    keys: Vec<Option<u32>>,
    styles: Vec<BatchStyle>,
}

struct BatchStyle {
    key: u32,
    parent: u64,
    inherited: CharacterProperties,
    stylesheet: Option<u64>,
}

fn prepare_batch_context(
    source: &dyn TableSource,
    targets: &[Target],
    limits: &Limits,
    include_stylesheets: bool,
) -> Result<BatchContext, Error> {
    let first = targets.first().ok_or(invalid("empty style batch"))?;
    let table = Table::read(source, *first, limits)?;
    let style_list = StyleList::read(source, table.style_list, limits)?;
    let cells = locate_cells(source, &table, targets, limits)?;
    let mut keys = Vec::new();
    limits.reserve(&mut keys, cells.len())?;
    for cell in &cells {
        keys.push(style_key(cell.as_deref(), limits)?);
    }
    let mut styles = Vec::new();
    limits.reserve(&mut styles, keys.len())?;
    limits.charge_work(
        keys.len()
            .saturating_mul(keys.len().max(1))
            .saturating_mul(2),
    )?;
    limits.charge_work(keys.len().saturating_mul(style_list.entries.len()))?;
    let mut effective_cache = Vec::new();
    let mut effective_path = BTreeSet::new();
    for key in keys.iter().flatten().copied() {
        if styles.iter().any(|style: &BatchStyle| style.key == key) {
            continue;
        }
        limits.admit_allocations(1)?;
        let parent = style_list.entry(key)?.style;
        let (inherited, stylesheet) = if let Some(cached) = styles
            .iter()
            .find(|style: &&BatchStyle| style.parent == parent)
        {
            limits.admit_character_properties(&cached.inherited)?;
            (cached.inherited.clone(), cached.stylesheet)
        } else {
            (
                effective_cached(
                    source,
                    parent,
                    limits,
                    &mut effective_cache,
                    &mut effective_path,
                    0,
                )?,
                if include_stylesheets {
                    Some(stylesheet_of(source, parent, limits)?)
                } else {
                    None
                },
            )
        };
        styles.push(BatchStyle {
            key,
            parent,
            inherited,
            stylesheet,
        });
    }
    Ok(BatchContext {
        table,
        style_list,
        keys,
        styles,
    })
}

fn locate_cell(
    source: &dyn TableSource,
    table: &Table,
    target: Target,
    limits: &Limits,
) -> Result<Option<LocatedCell>, Error> {
    let tile_storage =
        bytes(&table.store, STORE_TILES, limits)?.ok_or(invalid("tile storage missing"))?;
    let mut tiles = Vec::new();
    let tile_size = varint(tile_storage, 2, limits)?.ok_or(invalid("tile size missing"))?;
    for entry in repeated_payloads(tile_storage, 1, limits)? {
        limits.reserve(&mut tiles, 1)?;
        let key = u32::try_from(varint(entry, 1, limits)?.ok_or(invalid("tile key missing"))?)
            .map_err(|_| invalid("tile key exceeds u32"))?;
        let tile = reference(entry, 2, limits)?.ok_or(invalid("tile reference missing"))?;
        if tiles.iter().any(|(existing, _)| *existing == key) {
            return Err(invalid("tile key repeats"));
        }
        tiles.push((key, tile));
    }
    if tile_size == 0 {
        return Err(invalid("tile size is zero"));
    }
    let tile_key = u64::from(target.row) / tile_size;
    let local = u64::from(target.row) % tile_size;
    let tile = tiles
        .iter()
        .find(|(key, _)| u64::from(*key) == tile_key)
        .ok_or(Error::Refused(Refusal::UnsupportedSource))?
        .1;
    let (_, tile_data) = unique_payload(source, tile, &[TILE_MESSAGE_TYPE], limits)?;
    for row in repeated_payloads(tile_data, 5, limits)? {
        limits.charge_work(row.len())?;
        limits.admit_allocations(1)?;
        if varint(row, 1, limits)? != Some(local) {
            continue;
        }
        for (column, cell) in storage::bnc_cells(row).map_err(|_| invalid("malformed BNC row"))? {
            if column == target.column {
                let bytes = limits.copy(cell)?;
                return Ok(Some(LocatedCell { bytes }));
            }
        }
    }
    Ok(None)
}

fn locate_cells(
    source: &dyn TableSource,
    table: &Table,
    targets: &[Target],
    limits: &Limits,
) -> Result<Vec<Option<Vec<u8>>>, Error> {
    let mut sorted = Vec::new();
    limits.reserve(&mut sorted, targets.len())?;
    for (index, target) in targets.iter().copied().enumerate() {
        if target.row >= table.geometry.rows || target.column >= table.geometry.columns {
            return Err(Error::Refused(Refusal::OutsideTable));
        }
        sorted.push((target.row, target.column, index));
    }
    sorted.sort_unstable();
    limits.charge_work(sorted.len().saturating_mul(2))?;

    let mut cells = Vec::new();
    limits.reserve(&mut cells, targets.len())?;
    cells.resize_with(targets.len(), || None);
    let tile_storage =
        bytes(&table.store, STORE_TILES, limits)?.ok_or(invalid("tile storage missing"))?;
    let tile_size = varint(tile_storage, 2, limits)?.ok_or(invalid("tile size missing"))?;
    if tile_size == 0 {
        return Err(invalid("tile size is zero"));
    }
    let mut tiles = Vec::new();
    for entry in repeated_payloads(tile_storage, 1, limits)? {
        limits.reserve(&mut tiles, 1)?;
        let key = u32::try_from(varint(entry, 1, limits)?.ok_or(invalid("tile key missing"))?)
            .map_err(|_| invalid("tile key exceeds u32"))?;
        let tile = reference(entry, 2, limits)?.ok_or(invalid("tile reference missing"))?;
        if tiles.iter().any(|(existing, _)| *existing == key) {
            return Err(invalid("tile key repeats"));
        }
        tiles.push((key, tile));
    }

    let mut first = 0;
    while first < sorted.len() {
        let tile_key = u64::from(sorted[first].0) / tile_size;
        let mut last = first + 1;
        while last < sorted.len() && u64::from(sorted[last].0) / tile_size == tile_key {
            last += 1;
        }
        let tile_identifier = tiles
            .iter()
            .find(|(key, _)| u64::from(*key) == tile_key)
            .ok_or(Error::Refused(Refusal::UnsupportedSource))?
            .1;
        let (_, tile_data) = unique_payload(source, tile_identifier, &[TILE_MESSAGE_TYPE], limits)?;
        for row in repeated_payloads(tile_data, 5, limits)? {
            limits.charge_work(row.len())?;
            limits.admit_allocations(1)?;
            let local = varint(row, 1, limits)?.ok_or(invalid("tile row index missing"))?;
            let row_number = tile_key
                .checked_mul(tile_size)
                .and_then(|value| value.checked_add(local))
                .ok_or(invalid("tile row coordinate overflow"))?;
            let mut row_first = first;
            while row_first < last && u64::from(sorted[row_first].0) < row_number {
                row_first += 1;
            }
            if row_first == last || u64::from(sorted[row_first].0) != row_number {
                continue;
            }
            limits.admit_allocations(1)?;
            let row_cells = storage::bnc_cells(row).map_err(|_| invalid("malformed BNC row"))?;
            let mut row_last = row_first;
            while row_last < last && u64::from(sorted[row_last].0) == row_number {
                row_last += 1;
            }
            for (column, cell) in row_cells {
                let mut target_index = row_first;
                while target_index < row_last {
                    let (_, target_column, index) = sorted[target_index];
                    if target_column == column {
                        cells[index] = Some(limits.copy(cell)?);
                        break;
                    }
                    if target_column > column {
                        break;
                    }
                    target_index += 1;
                }
            }
        }
        first = last;
    }
    Ok(cells)
}

fn bnc_limits(data: &[u8], limits: &Limits) -> Result<BncLimits, Error> {
    let wire = limits.residual_wire()?;
    if data.is_empty() {
        return Ok(BncLimits::from_wire_limits(wire));
    }
    if data.len() > wire.max_input_bytes() {
        return Err(Error::Limit {
            kind: LimitKind::InputBytes,
            observed: data.len() as u64,
            maximum: wire.max_input_bytes() as u64,
        });
    }
    BncLimits::from_wire_limits(wire)
        .with_input_bytes(data.len())
        .and_then(|value| {
            value.with_tail_bytes(wire.max_input_bytes().min(BncLimits::MAX_TAIL_BYTES))
        })
        .and_then(|value| value.with_items(wire.max_rewrite_work().clamp(1, BncLimits::MAX_ITEMS)))
        .map_err(|_| invalid("invalid BNC operation limits"))
}

fn map_bnc(error: litchi_numbers_wire::Error) -> Error {
    match error {
        litchi_numbers_wire::Error::LimitExceeded {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                litchi_numbers_wire::LimitKind::InputBytes
                | litchi_numbers_wire::LimitKind::TailBytes => LimitKind::InputBytes,
                litchi_numbers_wire::LimitKind::OutputBytes => LimitKind::OutputBytes,
                litchi_numbers_wire::LimitKind::Fields => LimitKind::Fields,
                litchi_numbers_wire::LimitKind::Items => LimitKind::RewriteWork,
                _ => LimitKind::RewriteWork,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        litchi_numbers_wire::Error::Allocation { requested } => {
            Error::Allocation { amount: requested }
        },
        litchi_numbers_wire::Error::InvalidFormat(_)
        | litchi_numbers_wire::Error::ParseError(_) => invalid("malformed BNC cell"),
        _ => invalid("malformed BNC cell"),
    }
}

fn style_key(cell: Option<&[u8]>, limits: &Limits) -> Result<Option<u32>, Error> {
    let Some(cell) = cell else { return Ok(None) };
    limits.charge_input(cell.len())?;
    limits.charge_work(cell.len())?;
    let field_count = if cell.len() >= 12 {
        u32::from_le_bytes([cell[8], cell[9], cell[10], cell[11]]).count_ones() as usize
    } else {
        0
    };
    limits.charge_fields(field_count)?;
    limits.admit_allocations(litchi_numbers_wire::MAX_OWNED_BNC_PARSE_ALLOCATIONS)?;
    limits.admit_copies(cell.len())?;
    let bnc_limits = bnc_limits(cell, limits)?;
    let cell = BncCell::parse(cell, &bnc_limits).map_err(map_bnc)?;
    Ok(cell.text_style_identifier())
}

#[derive(Clone, Copy)]
struct Entry {
    key: u32,
    refcount: u32,
    style: u64,
    owner: u64,
}

struct StyleList {
    root: u64,
    root_type: u32,
    root_data: Vec<u8>,
    entries: Vec<Entry>,
    segments: Vec<(u64, u32, Vec<u8>)>,
    next: u32,
}

impl StyleList {
    fn read(source: &dyn TableSource, root: u64, limits: &Limits) -> Result<Self, Error> {
        let (root_type, root_data) = unique_payload(
            source,
            root,
            &[LIST_MESSAGE_TYPE, LIST_SEGMENT_MESSAGE_TYPE],
            limits,
        )?;
        if varint(root_data, 1, limits)? != Some(4) {
            return Err(Error::Refused(Refusal::UnsupportedStyleList));
        }
        let next = u32::try_from(
            varint(root_data, 2, limits)?.ok_or(invalid("style-list cursor missing"))?,
        )
        .map_err(|_| invalid("style-list cursor exceeds u32"))?;
        let mut entries = Vec::new();
        for raw in repeated_payloads(root_data, 3, limits)? {
            limits.reserve(&mut entries, 1)?;
            let key =
                u32::try_from(varint(raw, 1, limits)?.ok_or(invalid("style-list key missing"))?)
                    .map_err(|_| invalid("style-list key exceeds u32"))?;
            let refcount = u32::try_from(
                varint(raw, 2, limits)?.ok_or(invalid("style-list refcount missing"))?,
            )
            .map_err(|_| invalid("style-list refcount exceeds u32"))?;
            let style =
                reference(raw, 4, limits)?.ok_or(invalid("style-list style reference missing"))?;
            if entries.iter().any(|entry: &Entry| entry.key == key) {
                return Err(invalid("style-list key repeats"));
            }
            entries.push(Entry {
                key,
                refcount,
                style,
                owner: root,
            });
        }
        let mut segments = Vec::new();
        for raw in repeated_payloads(root_data, 4, limits)? {
            limits.reserve(&mut segments, 1)?;
            let identifier = reference(raw, 1, limits)?
                .ok_or(invalid("style-list segment reference missing"))?;
            let (segment_type, segment_data) =
                unique_payload(source, identifier, &[LIST_SEGMENT_MESSAGE_TYPE], limits)?;
            let segment_data = limits.copy(segment_data)?;
            for segment_entry in repeated_payloads(&segment_data, 3, limits)? {
                limits.reserve(&mut entries, 1)?;
                let key = u32::try_from(
                    varint(segment_entry, 1, limits)?.ok_or(invalid("segment key missing"))?,
                )
                .map_err(|_| invalid("segment key exceeds u32"))?;
                let refcount = u32::try_from(
                    varint(segment_entry, 2, limits)?.ok_or(invalid("segment refcount missing"))?,
                )
                .map_err(|_| invalid("segment refcount exceeds u32"))?;
                let style = reference(segment_entry, 4, limits)?
                    .ok_or(invalid("segment style reference missing"))?;
                if entries.iter().any(|entry| entry.key == key) {
                    return Err(invalid("style-list key repeats across segments"));
                }
                entries.push(Entry {
                    key,
                    refcount,
                    style,
                    owner: identifier,
                });
            }
            segments.push((identifier, segment_type, segment_data));
        }
        Ok(Self {
            root,
            root_type,
            root_data: limits.copy(root_data)?,
            entries,
            segments,
            next,
        })
    }

    fn entry(&self, key: u32) -> Result<Entry, Error> {
        self.entries
            .iter()
            .find(|entry| entry.key == key)
            .copied()
            .ok_or(invalid("style-list key is not present"))
    }

    fn next_key(&self) -> Result<u32, Error> {
        let max = self
            .entries
            .iter()
            .map(|entry| entry.key)
            .max()
            .unwrap_or(0);
        self.next
            .max(max.checked_add(1).ok_or(invalid("style key overflow"))?)
            .max(1)
            .checked_add(0)
            .ok_or(invalid("style key overflow"))
    }
}

fn style_payload<'a>(
    source: &'a dyn TableSource,
    identifier: u64,
    limits: &Limits,
) -> Result<&'a [u8], Error> {
    unique_payload(source, identifier, &[PARAGRAPH_STYLE_MESSAGE_TYPE], limits)
        .map(|(_, data)| data)
}

fn stylesheet_of(source: &dyn TableSource, style: u64, limits: &Limits) -> Result<u64, Error> {
    let base = bytes(style_payload(source, style, limits)?, 1, limits)?
        .ok_or(invalid("style base payload missing"))?;
    reference(base, 5, limits)?.ok_or(Error::Refused(Refusal::UnsupportedDependency))
}

fn parent_of(source: &dyn TableSource, style: u64, limits: &Limits) -> Result<Option<u64>, Error> {
    let payload = style_payload(source, style, limits)?;
    let wire = preflight_character_parent(payload, limits)?;
    litchi_iwa_text_wire::character_style_parent(payload, wire).map_err(map_char_style)
}

fn overlay_properties(
    inherited: &CharacterProperties,
    requested: &CharacterProperties,
) -> CharacterProperties {
    let mut properties = inherited.clone();
    if requested.bold.is_some() {
        properties.bold = requested.bold;
    }
    if requested.italic.is_some() {
        properties.italic = requested.italic;
    }
    if requested.underline.is_some() {
        properties.underline = requested.underline;
    }
    if requested.strikethrough.is_some() {
        properties.strikethrough = requested.strikethrough;
    }
    if requested.size.is_some() {
        properties.size = requested.size;
    }
    if requested.font_name.is_some() {
        properties.font_name.clone_from(&requested.font_name);
    }
    if requested.color.is_some() {
        properties.color = requested.color;
    }
    properties
}

fn canonical_variation_payload(data: &[u8], limits: &Limits) -> Result<bool, Error> {
    let root = view(data, limits)?;
    if root.len() != 4 {
        return Ok(false);
    }
    let mut fields = root.fields();
    let Some(base) = fields.next() else {
        return Ok(false);
    };
    if base.number() != 1 || base.wire_type() != 2 {
        return Ok(false);
    }
    let base = view(base.payload(), limits)?;
    let mut base_fields = base.fields();
    if let Some(parent) = base_fields.next() {
        if parent.number() != 3 || parent.wire_type() != 2 {
            return Ok(false);
        }
        if !canonical_reference_payload(parent.payload(), limits)? {
            return Ok(false);
        }
    }
    let Some(version) = base_fields.next() else {
        return Ok(false);
    };
    if version.number() != 4 || version.wire_type() != 0 {
        return Ok(false);
    }
    let (version_value, version_length) =
        litchi_iwa_common::varint::decode_varint_from_bytes(version.payload())
            .map_err(|_| invalid("malformed canonical style version"))?;
    if version_value != 1 || version_length != version.payload().len() {
        return Ok(false);
    }
    let Some(stylesheet) = base_fields.next() else {
        return Ok(false);
    };
    if stylesheet.number() != 5 || stylesheet.wire_type() != 2 {
        return Ok(false);
    }
    if !canonical_reference_payload(stylesheet.payload(), limits)? {
        return Ok(false);
    }
    if base_fields.next().is_some() {
        return Ok(false);
    }

    let Some(count_field) = fields.next() else {
        return Ok(false);
    };
    if count_field.number() != 10 || count_field.wire_type() != 0 {
        return Ok(false);
    }
    let Some(properties) = fields.next() else {
        return Ok(false);
    };
    if properties.number() != 11 || properties.wire_type() != 2 {
        return Ok(false);
    }
    let properties = view(properties.payload(), limits)?;
    let allowed = [1, 2, 3, 5, 7, 11, 12];
    let mut previous = 0;
    for field in properties.fields() {
        if !allowed.contains(&field.number()) || field.number() <= previous {
            return Ok(false);
        }
        previous = field.number();
        let expected_wire = match field.number() {
            1 | 2 | 11 | 12 => 0,
            3 => 5,
            5 | 7 => 2,
            _ => return Ok(false),
        };
        if field.wire_type() != expected_wire {
            return Ok(false);
        }
        if field.number() == 7 {
            let color = view(field.payload(), limits)?;
            if color.len() != 5 {
                return Ok(false);
            }
            let expected = [1, 3, 4, 5, 6];
            for (actual, number) in color.fields().zip(expected) {
                if actual.number() != number
                    || actual.wire_type() != if number == 1 { 0 } else { 5 }
                {
                    return Ok(false);
                }
            }
        }
    }
    let (count, count_length) =
        litchi_iwa_common::varint::decode_varint_from_bytes(count_field.payload())
            .map_err(|_| invalid("malformed canonical style count"))?;
    if count_length != count_field.payload().len() {
        return Ok(false);
    }
    if count != u64::try_from(properties.len()).unwrap_or(u64::MAX) {
        return Ok(false);
    }
    let Some(paragraph) = fields.next() else {
        return Ok(false);
    };
    if paragraph.number() != 12 || paragraph.wire_type() != 2 || !paragraph.payload().is_empty() {
        return Ok(false);
    }
    Ok(fields.next().is_none())
}

fn canonical_reference_payload(data: &[u8], limits: &Limits) -> Result<bool, Error> {
    let reference = view(data, limits)?;
    if reference.len() != 1 {
        return Ok(false);
    }
    let reference = reference.get(0).expect("reference field");
    if reference.number() != 1 || reference.wire_type() != 0 {
        return Ok(false);
    }
    let (_, length) = litchi_iwa_common::varint::decode_varint_from_bytes(reference.payload())
        .map_err(|_| invalid("malformed canonical style reference"))?;
    Ok(length == reference.payload().len())
}

/// Read effective character properties through the parent chain.
pub fn effective(
    source: &dyn TableSource,
    style: u64,
    limits: &Limits,
) -> Result<CharacterProperties, Error> {
    let mut properties = CharacterProperties::new();
    let mut current = Some(style);
    let mut seen = BTreeSet::new();
    for _ in 0..limits.max_style_depth {
        let Some(identifier) = current else { break };
        limits.admit_allocations(1)?;
        if !seen.insert(identifier) {
            return Err(invalid("character-style parent cycle"));
        }
        let payload = style_payload(source, identifier, limits)?;
        let wire = preflight_character_properties(payload, limits)?;
        let decoded = litchi_iwa_text_wire::character_style_properties(payload, wire)
            .map_err(map_char_style)?;
        if properties.font_name.is_none() {
            if let Some(font_name) = decoded.font_name.as_deref() {
                limits.admit_allocations(1)?;
                limits.admit_copies(font_name.len())?;
            }
        }
        properties.inherit(&decoded);
        let parent_wire = preflight_character_parent(payload, limits)?;
        current = litchi_iwa_text_wire::character_style_parent(payload, parent_wire)
            .map_err(map_char_style)?;
    }
    if current.is_some() {
        return Err(Error::Limit {
            kind: LimitKind::Nesting,
            observed: (limits.max_style_depth + 1) as u64,
            maximum: limits.max_style_depth as u64,
        });
    }
    Ok(properties)
}

fn effective_cached(
    source: &dyn TableSource,
    style: u64,
    limits: &Limits,
    cache: &mut Vec<(u64, CharacterProperties)>,
    path: &mut BTreeSet<u64>,
    depth: usize,
) -> Result<CharacterProperties, Error> {
    limits.charge_work(cache.len().saturating_add(1))?;
    if let Some((_, properties)) = cache.iter().find(|(identifier, _)| *identifier == style) {
        limits.admit_character_properties(properties)?;
        return Ok(properties.clone());
    }
    if depth >= limits.max_style_depth {
        return Err(Error::Limit {
            kind: LimitKind::Nesting,
            observed: (depth + 1) as u64,
            maximum: limits.max_style_depth as u64,
        });
    }
    limits.admit_allocations(1)?;
    if !path.insert(style) {
        return Err(invalid("character-style parent cycle"));
    }
    let payload = style_payload(source, style, limits)?;
    let wire = preflight_character_properties(payload, limits)?;
    let decoded =
        litchi_iwa_text_wire::character_style_properties(payload, wire).map_err(map_char_style)?;
    let parent_wire = preflight_character_parent(payload, limits)?;
    let parent = litchi_iwa_text_wire::character_style_parent(payload, parent_wire)
        .map_err(map_char_style)?;
    let mut properties = decoded;
    if let Some(parent) = parent {
        let inherited = effective_cached(source, parent, limits, cache, path, depth + 1)?;
        if properties.font_name.is_none() {
            if let Some(font_name) = inherited.font_name.as_deref() {
                limits.admit_allocations(1)?;
                limits.admit_copies(font_name.len())?;
            }
        }
        properties.inherit(&inherited);
    }
    path.remove(&style);
    limits.admit_character_properties(&properties)?;
    limits.reserve(cache, 1)?;
    cache.push((style, properties.clone()));
    Ok(properties)
}

/// Read the effective character properties of one target cell.
pub fn read(
    source: &dyn TableSource,
    target: Target,
    limits: &Limits,
) -> Result<CharacterProperties, Error> {
    let table = Table::read(source, target, limits)?;
    let cell = locate_cell(source, &table, target, limits)?;
    let key = style_key(cell.as_ref().map(|cell| cell.bytes.as_slice()), limits)?;
    let style = match key {
        Some(key) => {
            StyleList::read(source, table.style_list, limits)?
                .entry(key)?
                .style
        },
        None => {
            let field = if target.row < table.geometry.header_rows {
                25
            } else if target.row
                >= table
                    .geometry
                    .rows
                    .saturating_sub(table.geometry.footer_rows)
            {
                27
            } else if target.column < table.geometry.header_columns {
                26
            } else {
                24
            };
            reference(&table.model, field, limits)?
                .ok_or(Error::Refused(Refusal::UnsupportedDependency))?
        },
    };
    effective(source, style, limits)
}

/// Read effective character properties for one bounded batch of cells while
/// reusing the immutable table, tile, and style-list preparation.
pub fn read_many(
    source: &dyn TableSource,
    targets: &[Target],
    limits: &Limits,
) -> Result<Vec<CharacterProperties>, Error> {
    let context = prepare_batch_context(source, targets, limits, false)?;
    let mut output = Vec::new();
    limits.reserve(&mut output, targets.len())?;
    for (index, target) in targets.iter().copied().enumerate() {
        let properties = if let Some(key) = context.keys[index] {
            let inherited = &context
                .styles
                .iter()
                .find(|style| style.key == key)
                .ok_or(invalid("batch style facts missing"))?
                .inherited;
            limits.admit_character_properties(inherited)?;
            inherited.clone()
        } else {
            let field = if target.row < context.table.geometry.header_rows {
                25
            } else if target.row
                >= context
                    .table
                    .geometry
                    .rows
                    .saturating_sub(context.table.geometry.footer_rows)
            {
                27
            } else if target.column < context.table.geometry.header_columns {
                26
            } else {
                24
            };
            let style = reference(&context.table.model, field, limits)?
                .ok_or(Error::Refused(Refusal::UnsupportedDependency))?;
            effective(source, style, limits)?
        };
        output.push(properties);
    }
    Ok(output)
}

/// The reduced direct character overrides to apply to a current style.
pub fn minimal(
    requested: &CharacterProperties,
    inherited: &CharacterProperties,
) -> CharacterProperties {
    let flag = |value: Option<bool>| value.unwrap_or(false);
    let mut output = CharacterProperties::new();
    output.bold = requested
        .bold
        .filter(|value| flag(inherited.bold) != *value);
    output.italic = requested
        .italic
        .filter(|value| flag(inherited.italic) != *value);
    output.underline = requested
        .underline
        .filter(|value| flag(inherited.underline) != *value);
    output.strikethrough = requested
        .strikethrough
        .filter(|value| flag(inherited.strikethrough) != *value);
    output.size = requested
        .size
        .filter(|value| inherited.size != Some(*value));
    output.font_name = requested
        .font_name
        .clone()
        .filter(|value| inherited.font_name.as_ref() != Some(value));
    output.color = requested
        .color
        .filter(|value| inherited.color != Some(*value));
    output
}

fn minimal_exact(
    requested: &CharacterProperties,
    inherited: &CharacterProperties,
) -> CharacterProperties {
    let mut output = CharacterProperties::new();
    output.bold = requested
        .bold
        .filter(|value| inherited.bold != Some(*value));
    output.italic = requested
        .italic
        .filter(|value| inherited.italic != Some(*value));
    output.underline = requested
        .underline
        .filter(|value| inherited.underline != Some(*value));
    output.strikethrough = requested
        .strikethrough
        .filter(|value| inherited.strikethrough != Some(*value));
    output.size = requested
        .size
        .filter(|value| inherited.size != Some(*value));
    output.font_name = requested
        .font_name
        .clone()
        .filter(|value| inherited.font_name.as_ref() != Some(value));
    output.color = requested
        .color
        .filter(|value| inherited.color != Some(*value));
    output
}

/// A style payload rewrite planned by the shared graph engine.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Rewrite {
    /// Object identifier.
    pub identifier: u64,
    /// Unique message type to replace.
    pub message_type: u32,
    /// Replacement payload.
    pub data: Vec<u8>,
    /// References added to the message header.
    pub add_references: Vec<u64>,
    /// References removed from the message header.
    pub remove_references: Vec<u64>,
}

/// A new paragraph-style variation planned by the shared engine.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NewObject {
    /// Fresh identifier supplied by the leaf.
    pub identifier: u64,
    /// Component/object beside which the object is inserted.
    pub beside: u64,
    /// Message type.
    pub message_type: u32,
    /// Message versions, when required.
    pub versions: Option<[u32; 3]>,
    /// Payload.
    pub data: Vec<u8>,
    /// Message object references.
    pub object_references: Vec<u64>,
}

/// Metadata transition requested by the shared plan.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Registration {
    /// Register a fresh UUID.
    Uuid { object: u64 },
    /// Register a cross-component reference from `from` to `object`.
    ExternalReference { from: u64, object: u64 },
    /// Remove a style object and its metadata records.
    Removed { object: u64 },
}

/// One neutral whole-cell style graph plan.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Plan {
    /// Message payload rewrites, including the cell/tile/header payloads.
    pub rewrites: Vec<Rewrite>,
    /// New style variations.
    pub insertions: Vec<NewObject>,
    /// Optional style object to remove after rewrites.
    pub removal: Option<u64>,
    /// Metadata registry operations.
    pub registrations: Vec<Registration>,
    /// Cell keys to write after graph publication.
    pub cell_updates: Vec<(Target, u32)>,
    /// Whether this plan can be combined with sibling fresh-cell plans.
    pub batch_compatible: bool,
    /// Bounded work report for audit and gate receipts.
    pub report: Report,
    /// Existing style-list entries released by this plan, grouped by owner.
    pub releases: Vec<(u64, u32)>,
}

/// One staged request passed to the batch planner.
#[derive(Debug, Clone)]
pub struct Request {
    /// Selected cell.
    pub target: Target,
    /// Requested direct character properties.
    pub properties: CharacterProperties,
    /// Leaf-supplied identifier when a fresh variation is needed.
    pub fresh: Option<u64>,
}

/// Plan a direct CharacterProperties set. The caller supplies `fresh` only
/// when the current graph cannot be rewritten in place.
pub fn plan_set(
    source: &dyn TableSource,
    target: Target,
    requested: &CharacterProperties,
    fresh: Option<u64>,
    limits: &Limits,
) -> Result<Option<Plan>, Error> {
    let mut allocator = None;
    plan_set_inner(
        source,
        target,
        requested,
        fresh,
        limits,
        true,
        None,
        &mut allocator,
        true,
    )
}

/// Plan a set whose request-owned values were already admitted while an edit
/// was staged.  This keeps the staging and commit phases on one cumulative
/// budget without charging the same owned value a second time.
pub fn plan_set_precharged(
    source: &dyn TableSource,
    target: Target,
    requested: &CharacterProperties,
    fresh: Option<u64>,
    limits: &Limits,
) -> Result<Option<Plan>, Error> {
    let mut allocator = None;
    plan_set_inner(
        source,
        target,
        requested,
        fresh,
        limits,
        true,
        None,
        &mut allocator,
        false,
    )
}

fn plan_set_inner(
    source: &dyn TableSource,
    target: Target,
    requested: &CharacterProperties,
    fresh: Option<u64>,
    limits: &Limits,
    allow_exclusive: bool,
    batch: Option<(&BatchContext, usize)>,
    allocator: &mut Option<&mut dyn FnMut() -> Result<u64, Error>>,
    admit_requested: bool,
) -> Result<Option<Plan>, Error> {
    limits.charge_work(1)?;
    let owned_table = if batch.is_none() {
        Some(Table::read(source, target, limits)?)
    } else {
        None
    };
    let table = batch
        .map(|(context, _)| &context.table)
        .or(owned_table.as_ref())
        .ok_or(invalid("batch table context missing"))?;
    let batch_style = batch.and_then(|(context, index)| {
        context
            .keys
            .get(index)
            .and_then(|key| *key)
            .and_then(|key| context.styles.iter().find(|style| style.key == key))
    });
    let old_key = if let Some((context, index)) = batch {
        context.keys.get(index).copied().flatten()
    } else {
        let cell = locate_cell(source, table, target, limits)?;
        style_key(cell.as_ref().map(|cell| cell.bytes.as_slice()), limits)?
    };
    let owned_list = if batch.is_none() {
        Some(StyleList::read(source, table.style_list, limits)?)
    } else {
        None
    };
    let list = batch
        .map(|(context, _)| &context.style_list)
        .or(owned_list.as_ref())
        .ok_or(invalid("batch style-list context missing"))?;
    let parent = if let Some(style) = batch_style {
        style.parent
    } else {
        match old_key {
            Some(key) => list.entry(key)?.style,
            None => {
                let field = if target.row < table.geometry.header_rows {
                    25
                } else if target.row
                    >= table
                        .geometry
                        .rows
                        .saturating_sub(table.geometry.footer_rows)
                {
                    27
                } else if target.column < table.geometry.header_columns {
                    26
                } else {
                    24
                };
                reference(&table.model, field, limits)?
                    .ok_or(Error::Refused(Refusal::UnsupportedDependency))?
            },
        }
    };
    let inherited = if let Some(style) = batch_style {
        limits.admit_character_properties(&style.inherited)?;
        style.inherited.clone()
    } else {
        effective(source, parent, limits)?
    };
    if admit_requested {
        limits.admit_character_properties(requested)?;
    }
    let reduced = minimal(requested, &inherited);
    if reduced == CharacterProperties::new() {
        return Ok(None);
    }
    let stylesheet = if let Some(style) = batch_style {
        style
            .stylesheet
            .ok_or(invalid("batch stylesheet context missing"))?
    } else {
        stylesheet_of(source, parent, limits)?
    };
    let exclusive = allow_exclusive
        && if let Some(key) = old_key {
            let refcount = list.entry(key)?.refcount == 1;
            let canonical =
                canonical_variation_payload(style_payload(source, parent, limits)?, limits)?;
            let ownership = canonical
                && style_is_exclusive(source, parent, table.style_list, stylesheet, limits)?;
            refcount && ownership
        } else {
            false
        };
    if exclusive {
        let actual_parent = parent_of(source, parent, limits)?;
        let base = match actual_parent {
            Some(identifier) => effective(source, identifier, limits)?,
            None => CharacterProperties::new(),
        };
        // overlay_properties() clones the current effective font when it
        // builds the desired complete state for an in-place rewrite.
        limits.admit_character_properties(&inherited)?;
        limits.admit_character_properties(requested)?;
        let desired = overlay_properties(&inherited, requested);
        let reduced = minimal_exact(&desired, &base);
        let data = encode_variation(actual_parent, stylesheet, &reduced, limits)?;
        let mut rewrites = Vec::new();
        limits.reserve(&mut rewrites, 1)?;
        rewrites.push(Rewrite {
            identifier: parent,
            message_type: PARAGRAPH_STYLE_MESSAGE_TYPE,
            data,
            add_references: Vec::new(),
            remove_references: Vec::new(),
        });
        return Ok(Some(Plan {
            rewrites,
            report: limits.report(),
            ..Plan::default()
        }));
    }
    let fresh = match fresh {
        Some(fresh) => fresh,
        None => match allocator.as_mut() {
            Some(allocate) => (**allocate)()?,
            None => return Err(Error::NeedsIdentifier),
        },
    };
    let data = encode_variation(Some(parent), stylesheet, &reduced, limits)?;
    let (sheet_type, sheet_data) = if batch.is_some() {
        (STYLESHEET_MESSAGE_TYPE, Vec::new())
    } else {
        let (sheet_type, sheet_payload) =
            unique_payload(source, stylesheet, &[STYLESHEET_MESSAGE_TYPE], limits)?;
        (
            sheet_type,
            encode_stylesheet_variation(sheet_payload, fresh, Some(parent), limits)?,
        )
    };
    let key = list.next_key()?;
    let next = key.checked_add(1).ok_or(invalid("style key overflow"))?;
    let mut list_data =
        patch_varint_bounded(&list.root_data, 2, true, Some(u64::from(next)), limits)?;
    let mut entry = Vec::new();
    limits.admit_allocations(2)?;
    entry
        .try_reserve(32)
        .map_err(|_| Error::Allocation { amount: 32 })?;
    litchi_iwa_common::varint::encode_varint_into(&mut entry, 1_u64 << 3);
    litchi_iwa_common::varint::encode_varint_into(&mut entry, u64::from(key));
    litchi_iwa_common::varint::encode_varint_into(&mut entry, 2_u64 << 3);
    litchi_iwa_common::varint::encode_varint_into(&mut entry, 1);
    let mut reference_payload = Vec::new();
    reference_payload
        .try_reserve(16)
        .map_err(|_| Error::Allocation { amount: 16 })?;
    litchi_iwa_common::varint::encode_varint_into(&mut reference_payload, 1_u64 << 3);
    litchi_iwa_common::varint::encode_varint_into(&mut reference_payload, fresh);
    append_field_bounded(&mut entry, 4, &reference_payload, limits)?;
    append_field_bounded(&mut list_data, 3, &entry, limits)?;
    let mut rewrites = Vec::new();
    limits.reserve(&mut rewrites, 2)?;
    let mut stylesheet_references = Vec::new();
    limits.reserve(&mut stylesheet_references, 1)?;
    stylesheet_references.push(fresh);
    rewrites.push(Rewrite {
        identifier: stylesheet,
        message_type: sheet_type,
        data: sheet_data,
        add_references: stylesheet_references,
        remove_references: Vec::new(),
    });
    let mut list_references = Vec::new();
    limits.reserve(&mut list_references, 1)?;
    list_references.push(fresh);
    rewrites.push(Rewrite {
        identifier: list.root,
        message_type: list.root_type,
        data: list_data,
        add_references: list_references,
        remove_references: Vec::new(),
    });
    let mut releases = Vec::new();
    limits.reserve(&mut releases, if old_key.is_some() { 1 } else { 0 })?;
    if let Some(old) = old_key {
        limits.admit_allocations(1)?;
        let mut remaining = BTreeMap::from([(old, 1_u32)]);
        let owner = list.entry(old)?.owner;
        releases.push((owner, old));
        let (released, dropped) = if owner == list.root {
            litchi_numbers_wire::table_topology::lists::release(&rewrites[1].data, &mut remaining)
                .map_err(|_| Error::Refused(Refusal::UnsupportedStyleList))?
        } else {
            let (_, segment_type, segment_data) = list
                .segments
                .iter()
                .find(|(identifier, _, _)| *identifier == owner)
                .ok_or(invalid("style-list segment disappeared"))?;
            let (released, dropped) =
                litchi_numbers_wire::table_topology::lists::release(segment_data, &mut remaining)
                    .map_err(|_| Error::Refused(Refusal::UnsupportedStyleList))?;
            if let Some(released) = released {
                rewrites.push(Rewrite {
                    identifier: owner,
                    message_type: *segment_type,
                    data: released,
                    add_references: Vec::new(),
                    remove_references: dropped.clone(),
                });
            }
            (None, dropped)
        };
        if !remaining.is_empty() {
            return Err(Error::Refused(Refusal::UnsupportedDependency));
        }
        if let Some(released) = released {
            rewrites[1].data = released;
        }
        if owner == list.root {
            rewrites[1].remove_references = dropped;
        }
    }
    let mut object_references = Vec::new();
    limits.reserve(&mut object_references, 1)?;
    object_references.push(parent);
    let mut insertions = Vec::new();
    limits.reserve(&mut insertions, 1)?;
    insertions.push(NewObject {
        identifier: fresh,
        beside: stylesheet,
        message_type: PARAGRAPH_STYLE_MESSAGE_TYPE,
        versions: None,
        data,
        object_references,
    });
    let mut registrations = Vec::new();
    limits.reserve(&mut registrations, 2)?;
    registrations.push(Registration::Uuid { object: fresh });
    registrations.push(Registration::ExternalReference {
        from: list.root,
        object: fresh,
    });
    let mut cell_updates = Vec::new();
    limits.reserve(&mut cell_updates, 1)?;
    cell_updates.push((target, key));
    Ok(Some(Plan {
        rewrites,
        insertions,
        registrations,
        cell_updates,
        batch_compatible: true,
        report: limits.report(),
        releases,
        ..Plan::default()
    }))
}

/// Combine fresh-cell plans into one graph publication plan. Existing-key
/// releases remain on the leaf compatibility publisher until their ordering is
/// lifted without changing Numbers' staging semantics.
pub fn plan_set_many(
    source: &dyn TableSource,
    requests: &[Request],
    limits: &Limits,
) -> Result<Option<Plan>, Error> {
    plan_set_many_inner(source, requests, limits, true, None, true)
}

/// Plan a bounded batch while allocating fresh identifiers only after the
/// shared graph has proven that a staged request is a real variation (rather
/// than a semantic no-op).  Leaves own the identifier allocator; this helper
/// only invokes it after the shared admission phase.
pub fn plan_set_many_with_allocator(
    source: &dyn TableSource,
    requests: &[Request],
    limits: &Limits,
    prepare_batch: bool,
    allocator: &mut dyn FnMut() -> Result<u64, Error>,
) -> Result<Option<Plan>, Error> {
    plan_set_many_inner(
        source,
        requests,
        limits,
        prepare_batch,
        Some(allocator),
        true,
    )
}

/// Variant for a leaf edit that admitted each staged request before entering
/// the shared planner.  Cloned request values are still admitted by the leaf
/// before they are placed in the request vector.
pub fn plan_set_many_with_allocator_precharged(
    source: &dyn TableSource,
    requests: &[Request],
    limits: &Limits,
    prepare_batch: bool,
    allocator: &mut dyn FnMut() -> Result<u64, Error>,
) -> Result<Option<Plan>, Error> {
    plan_set_many_inner(
        source,
        requests,
        limits,
        prepare_batch,
        Some(allocator),
        false,
    )
}

fn plan_set_many_inner(
    source: &dyn TableSource,
    requests: &[Request],
    limits: &Limits,
    prepare_batch: bool,
    mut allocator: Option<&mut dyn FnMut() -> Result<u64, Error>>,
    admit_requested: bool,
) -> Result<Option<Plan>, Error> {
    if requests
        .windows(2)
        .any(|pair| pair[0].target.model_identifier != pair[1].target.model_identifier)
    {
        return Err(Error::BatchUnsupported);
    }
    limits.charge_work(requests.len())?;
    limits.admit_allocations(requests.len())?;
    let batch_context = if prepare_batch {
        if let Some(first) = requests.first() {
            if requests
                .iter()
                .all(|request| request.target.model_identifier == first.target.model_identifier)
            {
                let mut targets = Vec::new();
                limits.reserve(&mut targets, requests.len())?;
                targets.extend(requests.iter().map(|request| request.target));
                Some(prepare_batch_context(source, &targets, limits, true)?)
            } else {
                None
            }
        } else {
            None
        }
    } else {
        None
    };
    let mut plans = Vec::new();
    limits.reserve(&mut plans, requests.len())?;
    for (index, request) in requests.iter().enumerate() {
        // A batch has one publication order for all cells.  Keep the
        // exclusive in-place path single-cell-only; the Numbers facade
        // supplies fresh identifiers for these requests before re-planning.
        let Some(plan) = plan_set_inner(
            source,
            request.target,
            &request.properties,
            request.fresh,
            limits,
            false,
            batch_context.as_ref().map(|context| (context, index)),
            &mut allocator,
            admit_requested,
        )?
        else {
            continue;
        };
        if !plan.batch_compatible {
            return Err(Error::BatchUnsupported);
        }
        plans.push(plan);
    }
    if plans.is_empty() {
        return Ok(None);
    }
    let mut combined = Plan::default();
    let mut list_entries = Vec::new();
    limits.reserve(&mut list_entries, requests.len())?;
    let mut list_identifier = None;
    let mut list_type = None;
    let mut stylesheet_identifiers = Vec::new();
    limits.admit_allocations(requests.len().saturating_mul(2))?;
    let mut release_counts: BTreeMap<u64, BTreeMap<u32, u32>> = BTreeMap::new();
    for mut plan in plans {
        if let Some((_, key)) = plan.cell_updates.last_mut() {
            let collides = combined
                .cell_updates
                .iter()
                .any(|(_, existing)| existing == key);
            if collides {
                let next = combined
                    .cell_updates
                    .iter()
                    .map(|(_, existing)| *existing)
                    .max()
                    .unwrap_or(*key)
                    .checked_add(1)
                    .ok_or(invalid("batch style key overflow"))?;
                let old = *key;
                *key = next;
                for rewrite in &mut plan.rewrites {
                    if rewrite.message_type == LIST_MESSAGE_TYPE
                        || batch_context
                            .as_ref()
                            .is_some_and(|context| context.style_list.root == rewrite.identifier)
                    {
                        rewrite.data = rewrite_list_key(&rewrite.data, old, next, limits)?;
                    }
                }
            }
        }
        limits.reserve(&mut combined.insertions, plan.insertions.len())?;
        limits.reserve(&mut combined.registrations, plan.registrations.len())?;
        limits.reserve(&mut combined.cell_updates, plan.cell_updates.len())?;
        combined.insertions.extend(plan.insertions);
        combined.registrations.extend(plan.registrations);
        combined.cell_updates.extend(plan.cell_updates);
        for (owner, key) in plan.releases {
            *release_counts
                .entry(owner)
                .or_default()
                .entry(key)
                .or_default() += 1;
        }
        for rewrite in plan.rewrites {
            let is_root_list = prepare_batch
                && batch_context
                    .as_ref()
                    .is_some_and(|context| context.style_list.root == rewrite.identifier);
            if prepare_batch && rewrite.message_type == STYLESHEET_MESSAGE_TYPE {
                if !stylesheet_identifiers
                    .iter()
                    .any(|(identifier, _)| *identifier == rewrite.identifier)
                {
                    limits.reserve(&mut stylesheet_identifiers, 1)?;
                    stylesheet_identifiers.push((rewrite.identifier, rewrite.message_type));
                }
                continue;
            }
            // Segment releases are reconstructed once from the immutable
            // source list below, after all per-cell release counts have been
            // accumulated.  Keeping the per-cell segment rewrites here would
            // make a large batch look unsupported merely because two cells
            // share one segmented list owner.
            if prepare_batch && rewrite.message_type == LIST_SEGMENT_MESSAGE_TYPE && !is_root_list {
                continue;
            }
            if rewrite.message_type == LIST_MESSAGE_TYPE || is_root_list {
                list_identifier.get_or_insert(rewrite.identifier);
                list_type.get_or_insert(rewrite.message_type);
                let key = combined
                    .cell_updates
                    .last()
                    .map(|(_, key)| *key)
                    .ok_or(invalid("batch list key missing"))?;
                let mut matching_entry = None;
                for entry in repeated_payloads(&rewrite.data, 3, limits)? {
                    if varint(entry, 1, limits)? == Some(u64::from(key)) {
                        matching_entry = Some(entry);
                        break;
                    }
                }
                let entry = matching_entry.ok_or(invalid("batch style-list entry missing"))?;
                list_entries.push(limits.copy(entry)?);
                continue;
            }
            if !prepare_batch {
                if let Some(existing) = combined.rewrites.iter_mut().find(|candidate| {
                    candidate.identifier == rewrite.identifier
                        && candidate.message_type == rewrite.message_type
                }) {
                    if rewrite.message_type == STYLESHEET_MESSAGE_TYPE {
                        let insertion = combined
                            .insertions
                            .last()
                            .ok_or(invalid("batch stylesheet insertion missing"))?;
                        existing.data = encode_stylesheet_variation(
                            &existing.data,
                            insertion.identifier,
                            insertion.object_references.first().copied(),
                            limits,
                        )?;
                        limits.reserve(&mut existing.add_references, 1)?;
                        existing.add_references.push(insertion.identifier);
                    } else {
                        return Err(Error::BatchUnsupported);
                    }
                } else {
                    limits.reserve(&mut combined.rewrites, 1)?;
                    combined.rewrites.push(rewrite);
                }
                continue;
            }
            if combined.rewrites.iter().any(|candidate| {
                candidate.identifier == rewrite.identifier
                    && candidate.message_type == rewrite.message_type
            }) {
                return Err(Error::BatchUnsupported);
            } else {
                limits.reserve(&mut combined.rewrites, 1)?;
                combined.rewrites.push(rewrite);
            }
        }
    }
    for (stylesheet_identifier, stylesheet_type) in stylesheet_identifiers {
        let (_, stylesheet_data) = unique_payload(
            source,
            stylesheet_identifier,
            &[STYLESHEET_MESSAGE_TYPE],
            limits,
        )?;
        let mut styles = Vec::new();
        limits.reserve(&mut styles, combined.insertions.len())?;
        for insertion in &combined.insertions {
            if insertion.beside != stylesheet_identifier {
                continue;
            }
            let parent = insertion
                .object_references
                .first()
                .copied()
                .ok_or(invalid("batch stylesheet parent missing"))?;
            styles.push((insertion.identifier, parent));
        }
        if styles.is_empty() {
            continue;
        }
        let data = encode_stylesheet_batch(stylesheet_data, &styles, limits)?;
        let mut add_references = Vec::new();
        limits.reserve(&mut add_references, styles.len())?;
        add_references.extend(styles.iter().map(|(identifier, _)| *identifier));
        limits.reserve(&mut combined.rewrites, 1)?;
        combined.rewrites.push(Rewrite {
            identifier: stylesheet_identifier,
            message_type: stylesheet_type,
            data,
            add_references,
            remove_references: Vec::new(),
        });
    }
    if let Some(list_identifier) = list_identifier {
        let owned_list;
        let list = if let Some(context) = batch_context.as_ref() {
            &context.style_list
        } else {
            let table = Table::read(source, requests[0].target, limits)?;
            owned_list = StyleList::read(source, table.style_list, limits)?;
            &owned_list
        };
        if list.root != list_identifier {
            return Err(invalid("batch style-list owner changed"));
        }
        let mut root_data = limits.copy(&list.root_data)?;
        let mut next = list.next;
        for entry in &list_entries {
            let key = u32::try_from(
                varint(entry, 1, limits)?.ok_or(invalid("batch style-list key missing"))?,
            )
            .map_err(|_| invalid("batch style-list key exceeds u32"))?;
            let candidate = key
                .checked_add(1)
                .ok_or(invalid("batch style-list cursor overflow"))?;
            next = next.max(candidate);
        }
        root_data = patch_varint_bounded(&root_data, 2, true, Some(u64::from(next)), limits)?;
        append_repeated_fields_bounded(&mut root_data, 3, &list_entries, limits)?;
        let mut add_references = Vec::new();
        limits.reserve(&mut add_references, combined.insertions.len())?;
        add_references.extend(combined.insertions.iter().map(|object| object.identifier));
        let mut root_rewrite = Rewrite {
            identifier: list_identifier,
            message_type: list_type.unwrap_or(LIST_MESSAGE_TYPE),
            data: root_data,
            add_references,
            remove_references: Vec::new(),
        };
        for (owner, mut remaining) in release_counts {
            if owner == list.root {
                let (released, dropped) = litchi_numbers_wire::table_topology::lists::release(
                    &root_rewrite.data,
                    &mut remaining,
                )
                .map_err(|_| Error::Refused(Refusal::UnsupportedStyleList))?;
                if !remaining.is_empty() {
                    return Err(Error::Refused(Refusal::UnsupportedDependency));
                }
                if let Some(released) = released {
                    root_rewrite.data = released;
                }
                limits.reserve(&mut root_rewrite.remove_references, dropped.len())?;
                root_rewrite.remove_references.extend(dropped);
            } else {
                let (_, segment_type, segment_data) = list
                    .segments
                    .iter()
                    .find(|(identifier, _, _)| *identifier == owner)
                    .ok_or(invalid("batch style-list segment disappeared"))?;
                let (released, dropped) = litchi_numbers_wire::table_topology::lists::release(
                    segment_data,
                    &mut remaining,
                )
                .map_err(|_| Error::Refused(Refusal::UnsupportedStyleList))?;
                if !remaining.is_empty() {
                    return Err(Error::Refused(Refusal::UnsupportedDependency));
                }
                if let Some(released) = released {
                    limits.reserve(&mut combined.rewrites, 1)?;
                    combined.rewrites.push(Rewrite {
                        identifier: owner,
                        message_type: *segment_type,
                        data: released,
                        add_references: Vec::new(),
                        remove_references: dropped,
                    });
                }
            }
        }
        limits.reserve(&mut combined.rewrites, 1)?;
        combined.rewrites.push(root_rewrite);
    }
    combined.batch_compatible = true;
    combined.report = limits.report();
    Ok(Some(combined))
}

fn rewrite_list_key(data: &[u8], old: u32, new: u32, limits: &Limits) -> Result<Vec<u8>, Error> {
    let entries = repeated_payloads(data, 3, limits)?;
    let mut replacements = Vec::new();
    limits.reserve(&mut replacements, entries.len())?;
    for entry in entries {
        if varint(entry, 1, limits)? == Some(u64::from(old)) {
            limits.admit_allocations(1)?;
            replacements.push(patch_varint_bounded(
                entry,
                1,
                true,
                Some(u64::from(new)),
                limits,
            )?);
        } else {
            replacements.push(limits.copy(entry)?);
        }
    }
    limits.admit_allocations(1)?;
    let mut output = rewrite_repeated_bounded(data, 3, &replacements, limits)?;
    if varint(&output, 2, limits)? == Some(u64::from(old).saturating_add(1)) {
        limits.admit_allocations(1)?;
        output = patch_varint_bounded(
            &output,
            2,
            true,
            Some(u64::from(new).saturating_add(1)),
            limits,
        )?;
    }
    Ok(output)
}

fn bytes_bound(number: u32, length: usize) -> usize {
    litchi_iwa_common::varint::encoded_len((u64::from(number) << 3) | 2)
        .saturating_add(litchi_iwa_common::varint::encoded_len(length as u64))
        .saturating_add(length)
}

fn patch_varint_bounded(
    data: &[u8],
    field: u32,
    required: bool,
    value: Option<u64>,
    limits: &Limits,
) -> Result<Vec<u8>, Error> {
    let _wire = limits.prepare_output(data.len().saturating_add(11))?;
    let parsed = view(data, limits)?;
    let mut matching = None;
    for candidate in parsed.fields() {
        if candidate.number() != field {
            continue;
        }
        if matching.is_some() {
            return Err(invalid("a singular varint field repeats"));
        }
        matching = Some(candidate);
    }
    if matching.is_none() == required {
        return Err(invalid("singular varint field changed during mutation"));
    }
    if let Some(candidate) = matching {
        if candidate.wire_type() != 0 {
            return Err(invalid("singular field is not a varint"));
        }
    }
    limits.charge_work(parsed.len())?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(data.len().saturating_add(11))
        .map_err(|_| Error::Allocation {
            amount: data.len().saturating_add(11),
        })?;
    for candidate in parsed.fields() {
        if candidate.number() != field {
            output.extend_from_slice(candidate.raw());
        } else if let Some(value) = value {
            output.extend_from_slice(candidate.key());
            litchi_iwa_common::varint::encode_varint_into(&mut output, value);
        }
    }
    if matching.is_none() {
        if let Some(value) = value {
            let mut encoded = [0_u8; 10];
            let encoded = litchi_iwa_common::varint::encode_varint_to_buffer(value, &mut encoded);
            append_raw_field(&mut output, field, 0, encoded);
        }
    }
    Ok(output)
}

fn append_field_bounded(
    data: &mut Vec<u8>,
    field: u32,
    payload: &[u8],
    limits: &Limits,
) -> Result<(), Error> {
    let _wire =
        limits.prepare_output(data.len().saturating_add(bytes_bound(field, payload.len())))?;
    append_field_prepared(data, field, payload, limits)
}

fn append_repeated_fields_bounded(
    data: &mut Vec<u8>,
    field: u32,
    payloads: &[Vec<u8>],
    limits: &Limits,
) -> Result<(), Error> {
    let additional = payloads.iter().fold(0_usize, |total, payload| {
        total.saturating_add(bytes_bound(field, payload.len()))
    });
    let _wire = limits.prepare_output(data.len().saturating_add(additional))?;
    limits.charge_fields(payloads.len())?;
    limits.charge_work(payloads.len())?;
    data.try_reserve(additional)
        .map_err(|_| Error::Allocation { amount: additional })?;
    for payload in payloads {
        append_raw_field(data, field, 2, payload);
    }
    Ok(())
}

fn rewrite_repeated_bounded(
    data: &[u8],
    field: u32,
    replacements: &[Vec<u8>],
    limits: &Limits,
) -> Result<Vec<u8>, Error> {
    let replacement_bytes = replacements
        .iter()
        .fold(0_usize, |total, value| total.saturating_add(value.len()));
    let _wire = limits.prepare_output(data.len().saturating_add(replacement_bytes))?;
    let parsed = view(data, limits)?;
    let mut matches = 0_usize;
    for candidate in parsed.fields() {
        if candidate.number() == field {
            if candidate.wire_type() != 2 {
                return Err(invalid("repeated field is not length-delimited"));
            }
            matches = matches.saturating_add(1);
        }
    }
    let excess = replacements.len().saturating_sub(matches);
    limits.charge_work(parsed.len().saturating_add(excess))?;
    let retained = parsed.len().checked_sub(matches).ok_or(Error::Limit {
        kind: LimitKind::Fields,
        observed: u64::MAX,
        maximum: limits.max_references as u64,
    })?;
    limits.charge_fields(retained.saturating_add(replacements.len()))?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(data.len().saturating_add(replacement_bytes))
        .map_err(|_| Error::Allocation {
            amount: data.len().saturating_add(replacement_bytes),
        })?;
    let mut index = 0_usize;
    let mut found = false;
    for candidate in parsed.fields() {
        if candidate.number() != field {
            output.extend_from_slice(candidate.raw());
            continue;
        }
        found = true;
        if let Some(replacement) = replacements.get(index) {
            append_raw_field(&mut output, field, 2, replacement);
        }
        index = index.saturating_add(1);
    }
    if !found {
        for replacement in replacements {
            append_raw_field(&mut output, field, 2, replacement);
        }
    } else if index < replacements.len() {
        for replacement in &replacements[index..] {
            append_raw_field(&mut output, field, 2, replacement);
        }
    }
    Ok(output)
}

fn append_raw_field(output: &mut Vec<u8>, field: u32, wire_type: u8, payload: &[u8]) {
    litchi_iwa_common::varint::encode_varint_into(
        output,
        (u64::from(field) << 3) | u64::from(wire_type),
    );
    if wire_type == 2 {
        litchi_iwa_common::varint::encode_varint_into(output, payload.len() as u64);
    }
    output.extend_from_slice(payload);
}

fn append_field_prepared(
    data: &mut Vec<u8>,
    field: u32,
    payload: &[u8],
    limits: &Limits,
) -> Result<(), Error> {
    let parsed = view(data, limits)?;
    limits.charge_work(parsed.len())?;
    limits.charge_fields(1)?;
    limits.charge_work(1)?;
    limits.admit_allocations(1)?;
    data.try_reserve(bytes_bound(field, payload.len()))
        .map_err(|_| Error::Allocation {
            amount: bytes_bound(field, payload.len()),
        })?;
    append_raw_field(data, field, 2, payload);
    Ok(())
}

fn character_properties_bound(properties: &CharacterProperties) -> usize {
    let mut length = 0_usize;
    for value in [
        properties.bold,
        properties.italic,
        properties.underline,
        properties.strikethrough,
    ] {
        if value.is_some() {
            length = length.saturating_add(2);
        }
    }
    if properties.size.is_some() {
        length = length.saturating_add(5);
    }
    if let Some(font) = properties.font_name.as_deref() {
        length = length.saturating_add(bytes_bound(5, font.len()));
    }
    if properties.color.is_some() {
        length = length.saturating_add(bytes_bound(7, 22));
    }
    length
}

fn variation_fields(parent: Option<u64>, properties: &CharacterProperties) -> usize {
    6_usize
        .saturating_add(usize::from(parent.is_some()) * 2)
        .saturating_add(
            usize::from(properties.bold.is_some())
                + usize::from(properties.italic.is_some())
                + usize::from(properties.underline.is_some())
                + usize::from(properties.strikethrough.is_some())
                + usize::from(properties.size.is_some())
                + usize::from(properties.font_name.is_some())
                + usize::from(properties.color.is_some()),
        )
        .saturating_add(usize::from(properties.color.is_some()) * 5)
        .saturating_add(1)
}

fn encode_variation(
    parent: Option<u64>,
    stylesheet: u64,
    properties: &CharacterProperties,
    limits: &Limits,
) -> Result<Vec<u8>, Error> {
    let base_len = bytes_bound(5, 11)
        .saturating_add(2)
        .saturating_add(usize::from(parent.is_some()) * bytes_bound(3, 11));
    let bound = bytes_bound(1, base_len)
        .saturating_add(2)
        .saturating_add(bytes_bound(5, 11))
        .saturating_add(2)
        .saturating_add(bytes_bound(11, character_properties_bound(properties)))
        .saturating_add(2);
    let wire = limits.prepare_output(bound)?;
    let fields = variation_fields(parent, properties);
    limits.charge_fields(fields)?;
    limits.charge_work(fields)?;
    let data =
        litchi_iwa_text_wire::character_style_variation(parent, stylesheet, properties, wire)
            .map_err(map_char_style)?;
    with_paragraph_properties(data, limits)
}

fn preflight_stylesheet_writer(
    data: &[u8],
    parent: Option<u64>,
    limits: &Limits,
) -> Result<(), Error> {
    let root = view(data, limits)?;
    admit_decoder_view(data.len(), root.len(), limits)?;
    if parent.is_some() {
        for field in root.fields() {
            if field.number() != 5 || field.wire_type() != 2 {
                continue;
            }
            let entry = view(field.payload(), limits)?;
            admit_decoder_view(field.payload().len(), entry.len(), limits)?;
            for inner in entry.fields() {
                if inner.number() == 1 && inner.wire_type() == 2 {
                    let reference = view(inner.payload(), limits)?;
                    admit_decoder_view(inner.payload().len(), reference.len(), limits)?;
                }
            }
        }
    }
    Ok(())
}

fn encode_stylesheet_variation(
    data: &[u8],
    style: u64,
    parent: Option<u64>,
    limits: &Limits,
) -> Result<Vec<u8>, Error> {
    let wire = limits.prepare_output(data.len().saturating_add(64))?;
    preflight_stylesheet_writer(data, parent, limits)?;
    litchi_iwa_text_wire::stylesheet_with_style(data, style, parent, wire).map_err(map_char_style)
}

fn encode_stylesheet_batch(
    data: &[u8],
    styles: &[(u64, u64)],
    limits: &Limits,
) -> Result<Vec<u8>, Error> {
    if styles.is_empty() {
        return Err(invalid("empty stylesheet batch"));
    }
    let fields = view(data, limits)?;
    let mut last_style = None;
    let mut groups: Vec<(u64, Vec<u64>, Vec<u8>, Option<(usize, usize)>)> = Vec::new();
    limits.reserve(&mut groups, styles.len())?;
    for (style, parent) in styles {
        let Some((_, children, _, _)) = groups.iter_mut().find(|group| group.0 == *parent) else {
            limits.admit_allocations(2)?;
            groups.push((*parent, Vec::new(), Vec::new(), None));
            let (_, children, _, _) = groups.last_mut().ok_or(Error::Allocation { amount: 1 })?;
            children
                .try_reserve(1)
                .map_err(|_| Error::Allocation { amount: 1 })?;
            children.push(*style);
            continue;
        };
        limits.reserve(children, 1)?;
        children.push(*style);
    }
    for (index, field) in fields.fields().enumerate() {
        if field.number() == 1 {
            last_style = Some(index);
        }
        if field.number() != 5 || field.wire_type() != 2 {
            continue;
        }
        let nested = view(field.payload(), limits)?;
        let candidate = nested
            .fields()
            .find(|inner| inner.number() == 1)
            .map(|inner| varint(inner.payload(), 1, limits))
            .transpose()?
            .flatten();
        if let Some(parent) = candidate {
            if let Some(group) = groups.iter_mut().find(|group| group.0 == parent) {
                if group.3.replace((index, field.payload().len())).is_some() {
                    return Err(invalid("stylesheet parent repeats"));
                }
            }
        }
    }

    for group in &mut groups {
        limits.admit_allocations(1)?;
        group
            .2
            .try_reserve(group.1.len().saturating_mul(16))
            .map_err(|_| Error::Allocation {
                amount: group.1.len().saturating_mul(16),
            })?;
        for style in &group.1 {
            let mut reference = Vec::new();
            limits.admit_allocations(1)?;
            reference
                .try_reserve(16)
                .map_err(|_| Error::Allocation { amount: 16 })?;
            let mut encoded = [0_u8; 10];
            let value = litchi_iwa_common::varint::encode_varint_to_buffer(*style, &mut encoded);
            append_raw_field(&mut reference, 1, 0, value);
            append_raw_field(&mut group.2, 2, 2, &reference);
        }
    }
    let mut added_styles = Vec::new();
    limits.reserve(&mut added_styles, styles.len())?;
    for (style, _) in styles {
        limits.admit_allocations(1)?;
        let mut reference = Vec::new();
        reference
            .try_reserve(16)
            .map_err(|_| Error::Allocation { amount: 16 })?;
        let mut encoded = [0_u8; 10];
        let value = litchi_iwa_common::varint::encode_varint_to_buffer(*style, &mut encoded);
        append_raw_field(&mut reference, 1, 0, value);
        added_styles.push(reference);
    }
    let style_bytes = added_styles
        .iter()
        .map(|reference| bytes_bound(1, reference.len()))
        .sum::<usize>();
    let mut parent_delta = 0_usize;
    for (parent, _, children, entry) in &groups {
        if let Some((index, payload_len)) = entry {
            let field = fields
                .get(*index)
                .ok_or(invalid("stylesheet field missing"))?;
            parent_delta = parent_delta.saturating_add(
                bytes_bound(5, payload_len.saturating_add(children.len()))
                    .saturating_sub(field.raw().len()),
            );
        } else {
            let parent_reference_len =
                litchi_iwa_common::varint::encoded_len(u64::from(1_u32) << 3)
                    .saturating_add(litchi_iwa_common::varint::encoded_len(*parent));
            parent_delta = parent_delta.saturating_add(bytes_bound(
                5,
                parent_reference_len.saturating_add(children.len()),
            ));
        }
    }
    let output_len = data
        .len()
        .saturating_add(style_bytes)
        .saturating_add(parent_delta);
    let _wire = limits.prepare_output(output_len)?;
    limits.charge_fields(styles.len().saturating_mul(2).saturating_add(groups.len()))?;
    limits.charge_work(styles.len().saturating_mul(2).saturating_add(groups.len()))?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|_| Error::Allocation { amount: output_len })?;
    for (index, field) in fields.fields().enumerate() {
        let group = groups
            .iter()
            .find(|group| group.3.is_some_and(|(selected, _)| selected == index));
        if let Some((_, _, children, _)) = group {
            limits.admit_allocations(1)?;
            limits.admit_copies(field.payload().len())?;
            let mut payload = Vec::new();
            payload
                .try_reserve(field.payload().len().saturating_add(children.len()))
                .map_err(|_| Error::Allocation {
                    amount: field.payload().len().saturating_add(children.len()),
                })?;
            payload.extend_from_slice(field.payload());
            payload.extend_from_slice(children);
            append_raw_field(&mut output, 5, 2, &payload);
        } else {
            output.extend_from_slice(field.raw());
        }
        if Some(index) == last_style {
            for reference in &added_styles {
                append_raw_field(&mut output, 1, 2, reference);
            }
        }
    }
    if last_style.is_none() {
        for reference in &added_styles {
            append_raw_field(&mut output, 1, 2, reference);
        }
    }
    for (parent, _, children, entry) in &groups {
        if entry.is_some() {
            continue;
        }
        limits.admit_allocations(1)?;
        let mut parent_reference = Vec::new();
        parent_reference
            .try_reserve(16_usize.saturating_add(children.len()))
            .map_err(|_| Error::Allocation {
                amount: 16_usize.saturating_add(children.len()),
            })?;
        let mut encoded = [0_u8; 10];
        let value = litchi_iwa_common::varint::encode_varint_to_buffer(*parent, &mut encoded);
        append_raw_field(&mut parent_reference, 1, 0, value);
        parent_reference.extend_from_slice(children);
        append_raw_field(&mut output, 5, 2, &parent_reference);
    }
    Ok(output)
}

fn with_paragraph_properties(mut data: Vec<u8>, limits: &Limits) -> Result<Vec<u8>, Error> {
    if bytes(&data, PARAGRAPH_PROPERTIES_FIELD, limits)?.is_none() {
        append_field_prepared(&mut data, PARAGRAPH_PROPERTIES_FIELD, &[], limits)?;
    }
    Ok(data)
}

fn style_is_exclusive(
    source: &dyn TableSource,
    style: u64,
    list: u64,
    stylesheet: u64,
    limits: &Limits,
) -> Result<bool, Error> {
    for component in source.catalog().iter() {
        for object in &component.archive().objects {
            let identifier = object.archive_info.identifier.unwrap_or_default();
            if identifier == style || identifier == list || identifier == stylesheet {
                continue;
            }
            limits.charge_work(object.messages.len())?;
            // This metadata walk examines every message-info, field-info and
            // reference once; the work charge covers the same traversal that
            // the ownership predicate performs below.
            let mut references = 0_usize;
            let mut references_style = false;
            for info in &object.archive_info.message_infos {
                limits.charge_work(1)?;
                references =
                    references
                        .checked_add(info.object_references.len())
                        .ok_or(Error::Limit {
                            kind: LimitKind::Fields,
                            observed: u64::MAX,
                            maximum: limits.max_references as u64,
                        })?;
                limits.charge_work(info.object_references.len())?;
                references_style |= info.object_references.contains(&style);
                for field in &info.field_infos {
                    references = references
                        .checked_add(field.object_references.len())
                        .ok_or(Error::Limit {
                            kind: LimitKind::Fields,
                            observed: u64::MAX,
                            maximum: limits.max_references as u64,
                        })?;
                    limits.charge_work(1)?;
                    limits.charge_work(field.object_references.len())?;
                    references_style |= field.object_references.contains(&style);
                }
            }
            limits.charge_references(references)?;
            if references_style {
                return Ok(false);
            }
            limits.charge_work(object.messages.len())?;
            for message in &object.messages {
                if message.type_ == PARAGRAPH_STYLE_MESSAGE_TYPE {
                    let wire = preflight_character_parent(&message.data, limits)?;
                    if litchi_iwa_text_wire::character_style_parent(&message.data, wire)
                        .map_err(map_char_style)?
                        == Some(style)
                    {
                        return Ok(false);
                    }
                }
            }
        }
    }
    Ok(true)
}

#[cfg(test)]
mod tests {
    use litchi_iwa_archive::{
        Limits as ArchiveLimits,
        package::{Catalog, EntryEdit},
    };
    use litchi_iwa_core::{Archive, SnappyStream};
    use litchi_iwa_protos::tst;
    use prost::Message as _;

    use super::minimal;
    use super::{
        Error, LIST_MESSAGE_TYPE, LimitKind, Limits, Request, StyleList, TABLE_MODEL_MESSAGE_TYPE,
        TILE_MESSAGE_TYPE, Table, TableSource, Target, locate_cell, plan_set, plan_set_many,
        preflight_character_properties, style_is_exclusive, style_key, stylesheet_of, varint, view,
    };
    use litchi_iwa_text_wire::CharacterProperties;

    struct LimitedSource {
        catalog: litchi_iwa_archive::SourceCatalog,
        max_objects: usize,
        max_references: usize,
    }

    impl TableSource for LimitedSource {
        fn catalog(&self) -> &litchi_iwa_archive::ComponentCatalog {
            self.catalog.components()
        }

        fn physical(&self) -> Option<&litchi_iwa_archive::SourceCatalog> {
            Some(&self.catalog)
        }

        fn resolve_object(
            &self,
            identifier: u64,
        ) -> Result<Option<crate::ResolvedObject<'_>>, crate::ResolveError> {
            for (component_index, component) in self.catalog.components().iter().enumerate() {
                if let Some((object_index, object)) = component
                    .archive()
                    .objects
                    .iter()
                    .enumerate()
                    .find(|(_, object)| object.archive_info.identifier == Some(identifier))
                {
                    return Ok(Some(crate::ResolvedObject {
                        messages: &object.messages,
                        component_index,
                        object_index,
                    }));
                }
            }
            Ok(None)
        }

        fn source_bytes(&self) -> &[u8] {
            self.catalog.source_bytes()
        }

        fn archive_limits(&self) -> litchi_iwa_archive::Limits {
            litchi_iwa_archive::Limits::default()
        }

        fn max_objects(&self) -> usize {
            self.max_objects
        }

        fn max_references(&self) -> usize {
            self.max_references
        }
    }

    fn limited_source(max_objects: usize, max_references: usize) -> LimitedSource {
        limited_source_at("basic.numbers", max_objects, max_references)
    }

    fn limited_source_at(
        filename: &str,
        max_objects: usize,
        max_references: usize,
    ) -> LimitedSource {
        let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/iwork/numbers")
            .join(filename);
        let bytes = std::fs::read(path).expect("text-style fixture");
        LimitedSource {
            catalog: litchi_iwa_archive::SourceCatalog::from_bytes(
                &bytes,
                litchi_iwa_archive::Limits::default(),
            )
            .expect("text-style catalog"),
            max_objects,
            max_references,
        }
    }

    fn limited_source_bytes(
        bytes: &[u8],
        max_objects: usize,
        max_references: usize,
    ) -> LimitedSource {
        LimitedSource {
            catalog: litchi_iwa_archive::SourceCatalog::from_bytes(
                bytes,
                litchi_iwa_archive::Limits::default(),
            )
            .expect("catalog"),
            max_objects,
            max_references,
        }
    }

    fn source_with_shared_style_key() -> Vec<u8> {
        let source = std::fs::read(
            std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("../../test-data/iwork/numbers/basic.numbers"),
        )
        .expect("basic fixture");
        let catalog = Catalog::from_bytes(&source, ArchiveLimits::default()).expect("catalog");
        let mut style_list_id = None;
        let mut style_id = None;
        for entry in catalog
            .iter()
            .filter(|entry| entry.name().ends_with(".iwa"))
        {
            let bytes = SnappyStream::decompress(entry.data(), Default::default())
                .expect("decompress component");
            let archive = Archive::parse(bytes.as_bytes(), Default::default()).expect("archive");
            for object in &archive.objects {
                for message in &object.messages {
                    if message.type_ == TABLE_MODEL_MESSAGE_TYPE {
                        let model =
                            tst::TableModelArchive::decode(message.data.as_slice()).expect("model");
                        style_list_id = Some(model.base_data_store.style_table.identifier);
                        style_id = Some(model.body_text_style.identifier);
                    }
                }
            }
        }
        let style_list_id = style_list_id.expect("style-list reference");
        let style_id = style_id.expect("body style reference");
        let style_entry = catalog
            .iter()
            .find(|entry| {
                let Ok(bytes) = SnappyStream::decompress(entry.data(), Default::default()) else {
                    return false;
                };
                let Ok(archive) = Archive::parse(bytes.as_bytes(), Default::default()) else {
                    return false;
                };
                archive
                    .objects
                    .iter()
                    .any(|object| object.archive_info.identifier == Some(style_list_id))
            })
            .expect("style-list member");
        let style_member = style_entry.name().to_owned();
        let style_bytes = SnappyStream::decompress(style_entry.data(), Default::default())
            .expect("decompress style member");
        let mut style_archive =
            Archive::parse(style_bytes.as_bytes(), Default::default()).expect("style archive");
        let style_object = style_archive
            .objects
            .iter_mut()
            .find(|object| object.archive_info.identifier == Some(style_list_id))
            .expect("style-list object");
        let style_message = style_object
            .messages
            .iter_mut()
            .find(|message| message.type_ == LIST_MESSAGE_TYPE)
            .expect("style-list message");
        let mut style_list =
            tst::TableDataList::decode(style_message.data.as_slice()).expect("style list");
        style_list.list_type = tst::table_data_list::ListType::Style as i32;
        let style_key = style_list
            .entries
            .iter()
            .map(|entry| entry.key)
            .max()
            .unwrap_or(0)
            .saturating_add(1);
        style_list.next_list_id = style_key.saturating_add(1);
        style_list.entries.push(tst::table_data_list::ListEntry {
            key: style_key,
            refcount: 2,
            reference: Some(litchi_iwa_protos::tsp::Reference {
                identifier: style_id,
                ..Default::default()
            }),
            ..Default::default()
        });
        style_message.data = style_list.encode_to_vec();
        if let Some(info) = style_object.archive_info.message_infos.first_mut() {
            if !info.object_references.contains(&style_id) {
                info.object_references.push(style_id);
            }
        }
        let style_compressed = SnappyStream::compress(
            &style_archive
                .to_bytes(Default::default())
                .expect("style encode"),
            Default::default(),
        )
        .expect("compress style member");

        let tile_entry = catalog
            .iter()
            .find(|entry| entry.name() == "Index/Tables/Tile.iwa")
            .expect("tile fixture");
        let tile_bytes = SnappyStream::decompress(tile_entry.data(), Default::default())
            .expect("decompress tile");
        let mut archive =
            Archive::parse(tile_bytes.as_bytes(), Default::default()).expect("tile archive");
        let tile_message = archive
            .objects
            .iter_mut()
            .flat_map(|object| object.messages.iter_mut())
            .find(|message| message.type_ == TILE_MESSAGE_TYPE)
            .expect("tile message");
        let mut tile = tst::Tile::decode(tile_message.data.as_slice()).expect("tile");
        let mut changed_rows = 0;
        for row in &mut tile.row_infos {
            if changed_rows == 2 || row.cell_count == 0 {
                continue;
            }
            let storage = row.cell_storage_buffer.clone().expect("storage");
            let offsets = row.cell_offsets.clone().expect("offsets");
            // `cell_count` counts materialized cells, while the sparse offset
            // table spans the tile's full column range.
            let count = offsets.len() / 2;
            let present = (0..count)
                .filter(|index| {
                    u16::from_le_bytes([offsets[index * 2], offsets[index * 2 + 1]]) != u16::MAX
                })
                .collect::<Vec<_>>();
            if present.is_empty() {
                continue;
            }
            let mut cells = vec![None; count];
            for (position, index) in present.iter().enumerate() {
                let begin = u16::from_le_bytes([offsets[index * 2], offsets[index * 2 + 1]]);
                let end = present
                    .get(position + 1)
                    .map(|next| u16::from_le_bytes([offsets[next * 2], offsets[next * 2 + 1]]))
                    .unwrap_or(storage.len() as u16);
                cells[*index] = Some(storage[usize::from(begin)..usize::from(end)].to_vec());
            }
            let mut decoded = litchi_numbers_wire::BncCell::parse(
                cells[present[0]].as_ref().expect("cell"),
                &litchi_numbers_wire::Limits::DEFAULT,
            )
            .expect("BNC cell");
            decoded.set_text_style_identifier(Some(style_key));
            cells[present[0]] = Some(
                decoded
                    .encode(&litchi_numbers_wire::Limits::DEFAULT)
                    .expect("BNC encoding"),
            );
            let mut storage = Vec::new();
            let mut offsets = Vec::new();
            for cell in &cells {
                if let Some(cell) = cell {
                    offsets.extend_from_slice(&(storage.len() as u16).to_le_bytes());
                    storage.extend_from_slice(cell);
                } else {
                    offsets.extend_from_slice(&u16::MAX.to_le_bytes());
                }
            }
            row.cell_storage_buffer = Some(storage.clone());
            row.cell_offsets = Some(offsets.clone());
            row.cell_storage_buffer_pre_bnc = storage;
            row.cell_offsets_pre_bnc = offsets;
            changed_rows += 1;
        }
        assert_eq!(changed_rows, 2, "two materialized rows");
        tile_message.data = tile.encode_to_vec();
        let compressed = SnappyStream::compress(
            &archive.to_bytes(Default::default()).expect("tile encode"),
            Default::default(),
        )
        .expect("compress tile");
        catalog
            .reassemble_to_bytes(
                &[
                    EntryEdit::new(&style_member, &style_compressed),
                    EntryEdit::new("Index/Tables/Tile.iwa", &compressed),
                ],
                ArchiveLimits::default(),
            )
            .expect("reassemble fixture")
    }

    #[test]
    fn minimal_keeps_only_values_that_differ_from_inheritance() {
        let mut requested = CharacterProperties::new();
        requested.bold = Some(true);
        requested.italic = Some(false);
        requested.size = Some(18.0);
        requested.font_name = Some("Helvetica".to_owned());

        let mut inherited = CharacterProperties::new();
        inherited.bold = Some(true);
        inherited.italic = Some(true);
        inherited.size = Some(12.0);
        inherited.font_name = Some("Helvetica".to_owned());

        let reduced = minimal(&requested, &inherited);
        assert_eq!(reduced.bold, None);
        assert_eq!(reduced.italic, Some(false));
        assert_eq!(reduced.size, Some(18.0));
        assert_eq!(reduced.font_name, None);
    }

    #[test]
    fn unset_emphasis_is_compared_as_host_default_off() {
        let mut requested = CharacterProperties::new();
        requested.bold = Some(false);
        let inherited = CharacterProperties::new();
        assert!(minimal(&requested, &inherited).is_empty());
    }

    #[test]
    fn zero_residual_ceiling_refuses_before_graph_admission() {
        let source = limited_source(1_000_000, 0);
        assert!(matches!(
            Limits::from_source(&source),
            Err(Error::Limit {
                kind: LimitKind::Fields,
                observed: 1,
                maximum: 0,
            })
        ));
    }

    #[test]
    fn low_wire_field_budget_refuses_before_view_allocation() {
        let source = limited_source_at("shared-model-native.numbers", 1_000_000, 1_000_000);
        let limits = Limits::from_source(&source).expect("source limits");
        limits.remaining_fields.set(0);
        assert!(matches!(
            view(&[8, 1], &limits),
            Err(Error::Limit {
                kind: LimitKind::Fields,
                maximum: _,
                ..
            })
        ));
    }

    #[test]
    fn low_work_budget_refuses_before_wire_decode() {
        let source = limited_source(1_000_000, 1_000_000);
        let limits = Limits::from_source(&source).expect("source limits");
        limits.remaining_work.set(0);
        assert!(matches!(
            view(&[8, 1], &limits),
            Err(Error::Limit {
                kind: LimitKind::RewriteWork,
                maximum: _,
                ..
            })
        ));
    }

    #[test]
    fn low_allocation_budget_refuses_before_wire_view_reservation() {
        let source = limited_source(1_000_000, 1_000_000);
        let limits = Limits::from_source(&source).expect("source limits");
        limits.remaining_allocations.set(0);
        assert!(matches!(
            view(&[8, 1], &limits),
            Err(Error::Allocation { amount: 1 })
        ));
    }

    #[test]
    fn font_allocation_is_admitted_before_character_style_decode() {
        let source = limited_source(1_000_000, 1_000_000);
        let limits = Limits::from_source(&source).expect("source limits");
        // Root and property WireView spans consume two admissions; the font
        // String must be admitted before character_style_properties() owns it.
        limits.remaining_allocations.set(2);
        let data = [0x5a, 0x03, 0x2a, 0x01, b'A'];
        assert!(matches!(
            preflight_character_properties(&data, &limits),
            Err(Error::Allocation { amount: 1 })
        ));
    }

    #[test]
    fn style_key_uses_residual_wire_fields_before_bnc_parse() {
        let source = limited_source(1_000_000, 1_000_000);
        let limits = Limits::from_source(&source).expect("source limits");
        limits.remaining_fields.set(0);
        let cell = [5, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0];
        assert!(matches!(
            style_key(Some(&cell), &limits),
            Err(Error::Limit {
                kind: LimitKind::Fields,
                maximum: _,
                ..
            })
        ));
    }

    #[test]
    fn decoder_and_wire_view_passes_consume_cumulative_limits() {
        let source = limited_source(1_000_000, 1_000_000);
        let limits = Limits::from_source(&source).expect("source limits");
        // For `[8, 1]`, raw scan, WireView, and the text-wire decoder each
        // consume one field; input/work have headroom so the second call
        // fails specifically on the cumulative field ledger.
        limits.remaining_input.set(100);
        limits.remaining_fields.set(3);
        limits.remaining_work.set(100);
        let data = [8, 1];
        assert!(preflight_character_properties(&data, &limits).is_ok());
        assert!(matches!(
            preflight_character_properties(&data, &limits),
            Err(Error::Limit {
                kind: LimitKind::Fields,
                maximum: 0,
                ..
            })
        ));
    }

    #[test]
    fn character_style_encoder_preserves_output_limit_dimension() {
        let source = limited_source(1_000_000, 1_000_000);
        let limits = Limits::from_source(&source).expect("source limits");
        limits.remaining_output.set(1);
        let model = source
            .catalog()
            .iter()
            .flat_map(|component| component.archive().objects.iter())
            .find_map(|object| {
                object
                    .messages
                    .iter()
                    .any(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE)
                    .then_some(object.archive_info.identifier)
                    .flatten()
            })
            .expect("table model");
        let mut requested = CharacterProperties::new();
        requested.size = Some(73.0);
        assert!(matches!(
            plan_set(
                &source,
                Target {
                    model_identifier: model,
                    row: 1,
                    column: 1,
                },
                &requested,
                Some(9_000_001),
                &limits,
            ),
            Err(Error::Limit {
                kind: LimitKind::OutputBytes,
                ..
            })
        ));
    }

    #[test]
    fn duplicate_singular_style_field_is_rejected_before_projection() {
        let source = limited_source(1_000_000, 1_000_000);
        let limits = Limits::from_source(&source).expect("source limits");
        assert!(matches!(
            varint(&[8, 1, 8, 2], 1, &limits),
            Err(Error::InvalidSource(_))
        ));
    }

    #[test]
    fn unknown_style_field_remains_borrowed_and_semantically_local() {
        let source = limited_source(1_000_000, 1_000_000);
        let limits = Limits::from_source(&source).expect("source limits");
        let data = [8, 1, 0x98, 6, 0x7b];
        assert_eq!(varint(&data, 1, &limits).expect("known field"), Some(1));
        let view = view(&data, &limits).expect("wire view");
        assert_eq!(
            view.fields().map(|field| field.raw()).collect::<Vec<_>>(),
            vec![&data[..2], &data[2..]]
        );
    }

    #[test]
    fn shared_style_key_batch_releases_shared_key_after_all_cells_move() {
        let bytes = source_with_shared_style_key();
        let source = limited_source_bytes(&bytes, 1_000_000, 1_000_000);
        let limits = Limits::from_source(&source).expect("source limits");
        let model = source
            .catalog()
            .iter()
            .flat_map(|component| component.archive().objects.iter())
            .find_map(|object| {
                object
                    .messages
                    .iter()
                    .any(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE)
                    .then_some(object.archive_info.identifier)
                    .flatten()
            })
            .expect("table model");
        let table = Table::read(
            &source,
            Target {
                model_identifier: model,
                row: 0,
                column: 0,
            },
            &limits,
        )
        .expect("table graph");
        let mut positions = Vec::new();
        for row in 0..table.geometry.rows {
            for column in 0..table.geometry.columns {
                let target = Target {
                    model_identifier: model,
                    row,
                    column,
                };
                let Some(cell) = locate_cell(&source, &table, target, &limits).expect("cell")
                else {
                    continue;
                };
                if style_key(Some(&cell.bytes), &limits)
                    .expect("style key")
                    .is_some()
                {
                    positions.push((row, column));
                }
            }
        }
        assert!(
            positions.len() >= 2,
            "fixture did not retain two shared cells"
        );
        let selected = positions.into_iter().take(2).collect::<Vec<_>>();
        let old_key = selected
            .iter()
            .map(|(row, column)| {
                let target = Target {
                    model_identifier: model,
                    row: *row,
                    column: *column,
                };
                let cell = locate_cell(&source, &table, target, &limits)
                    .expect("selected cell")
                    .expect("selected cell bytes");
                style_key(Some(&cell.bytes), &limits).expect("selected style key")
            })
            .collect::<Vec<_>>();
        assert!(old_key.iter().all(|key| *key == old_key[0]));
        let old_key = old_key[0].expect("shared style key");
        let ownership_limits = Limits::from_source(&source).expect("ownership limits");
        let style_list =
            StyleList::read(&source, table.style_list, &ownership_limits).expect("style list");
        let shared_style = style_list.entry(old_key).expect("shared entry").style;
        let stylesheet =
            stylesheet_of(&source, shared_style, &ownership_limits).expect("shared stylesheet");
        assert!(
            !style_is_exclusive(
                &source,
                shared_style,
                table.style_list,
                stylesheet,
                &ownership_limits,
            )
            .expect("ownership proof"),
            "the model's incoming style reference blocks in-place rewrite"
        );
        let mut properties = CharacterProperties::new();
        properties.size = Some(73.0);
        let requests = selected
            .into_iter()
            .take(2)
            .enumerate()
            .map(|(index, (row, column))| Request {
                target: Target {
                    model_identifier: model,
                    row,
                    column,
                },
                properties: properties.clone(),
                fresh: Some(9_000_000 + index as u64),
            })
            .collect::<Vec<_>>();
        let plan = plan_set_many(&source, &requests, &limits)
            .expect("shared-key plan")
            .expect("shared-key changes");
        assert_eq!(plan.insertions.len(), 2);
        let probe_limits = Limits::from_source(&source).expect("probe limits");
        let before_output = probe_limits.remaining_output.get();
        let _single = plan_set(
            &source,
            requests[0].target,
            &requests[0].properties,
            requests[0].fresh,
            &probe_limits,
        )
        .expect("single output probe")
        .expect("single output change");
        let single_output = before_output.saturating_sub(probe_limits.remaining_output.get());
        assert!(single_output > 0, "single plan must debit output");
        let batch_limits = Limits::from_source(&source).expect("batch limits");
        batch_limits
            .remaining_output
            .set(single_output.saturating_add(1));
        assert!(matches!(
            plan_set_many(&source, &requests, &batch_limits),
            Err(Error::Limit {
                kind: LimitKind::OutputBytes,
                ..
            })
        ));
        let list_rewrite = plan
            .rewrites
            .iter()
            .find(|rewrite| rewrite.message_type == LIST_MESSAGE_TYPE)
            .expect("shared-key list rewrite");
        let list = tst::TableDataList::decode(list_rewrite.data.as_slice()).expect("list rewrite");
        assert!(
            list.entries.iter().all(|entry| entry.key != old_key),
            "shared key remains after both cells move"
        );
        assert!(plan.batch_compatible);
    }
}
