//! Strict scalar projection for Numbers table header/footer settings.
use crate::buffa_numbers_table_header_settings_generated::LitchiIwaProjection as projection;
use buffa::DecodeOptions as BuffaDecodeOptions;
use std::fmt;
const TABLE_ROWS_FIELD: u32 = 6;
const TABLE_COLUMNS_FIELD: u32 = 7;
const HEADER_ROWS_FIELD: u32 = 9;
const HEADER_COLUMNS_FIELD: u32 = 10;
const FOOTER_ROWS_FIELD: u32 = 11;
const HEADER_ROWS_FROZEN_FIELD: u32 = 12;
const HEADER_COLUMNS_FROZEN_FIELD: u32 = 13;
const REPEATING_HEADER_ROWS_FIELD: u32 = 29;
const REPEATING_HEADER_COLUMNS_FIELD: u32 = 32;
const MAX_RECURSION: u32 = 64;
/// Finite aggregate limits for one table-model payload.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeOptions {
    bytes: usize,
    fields: usize,
    work: usize,
    recursion: u32,
    output_bytes: usize,
}
impl DecodeOptions {
    #[must_use]
    pub const fn new(bytes: usize, fields: usize, work: usize, recursion: u32) -> Self {
        Self {
            bytes,
            fields,
            work,
            recursion,
            output_bytes: bytes,
        }
    }

    /// Build conservative finite limits from one known source payload.
    #[must_use]
    pub fn for_source(source: &[u8]) -> Self {
        let bytes = source.len().max(1);
        Self {
            bytes,
            fields: bytes.checked_mul(8).unwrap_or(usize::MAX).max(1),
            work: bytes.checked_mul(32).unwrap_or(usize::MAX).max(1),
            recursion: 8,
            output_bytes: bytes.checked_mul(2).unwrap_or(usize::MAX).max(1),
        }
    }

    /// Replace the aggregate candidate-output ceiling used by rewrites.
    #[must_use]
    pub const fn with_max_output_bytes(mut self, maximum: usize) -> Self {
        self.output_bytes = maximum;
        self
    }

    /// Return the aggregate candidate-output ceiling.
    #[must_use]
    pub const fn max_output_bytes(self) -> usize {
        self.output_bytes
    }
    fn buffa(self) -> BuffaDecodeOptions {
        BuffaDecodeOptions::new()
            .with_max_message_size(self.bytes)
            .with_unknown_field_limit(0)
            .with_element_memory_limit(0)
            .with_recursion_limit(self.recursion)
    }
}
/// Presence-preserving native header/footer settings; booleans are raw source facts.
/// Content-free byte or nesting limit observation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct TableHeaderSettingsSnapshot {
    rows: u32,
    columns: u32,
    header_rows: Option<u32>,
    header_columns: Option<u32>,
    footer_rows: Option<u32>,
    header_rows_frozen: Option<bool>,
    header_columns_frozen: Option<bool>,
    repeating_header_rows_enabled: Option<bool>,
    repeating_header_columns_enabled: Option<bool>,
}
impl TableHeaderSettingsSnapshot {
    #[must_use]
    pub const fn rows(self) -> u32 {
        self.rows
    }
    #[must_use]
    pub const fn columns(self) -> u32 {
        self.columns
    }
    #[must_use]
    pub const fn header_rows(self) -> Option<u32> {
        self.header_rows
    }
    #[must_use]
    pub const fn header_columns(self) -> Option<u32> {
        self.header_columns
    }
    #[must_use]
    pub const fn footer_rows(self) -> Option<u32> {
        self.footer_rows
    }
    #[must_use]
    pub const fn header_rows_frozen(self) -> Option<bool> {
        self.header_rows_frozen
    }
    #[must_use]
    pub const fn header_columns_frozen(self) -> Option<bool> {
        self.header_columns_frozen
    }
    #[must_use]
    pub const fn repeating_header_rows_enabled(self) -> Option<bool> {
        self.repeating_header_rows_enabled
    }
    #[must_use]
    pub const fn repeating_header_columns_enabled(self) -> Option<bool> {
        self.repeating_header_columns_enabled
    }

    /// Return a presence-preserving update containing all optional settings.
    #[must_use]
    pub const fn optional_write(self) -> TableHeaderSettingsWrite {
        TableHeaderSettingsWrite {
            header_rows: self.header_rows,
            header_columns: self.header_columns,
            footer_rows: self.footer_rows,
            header_rows_frozen: self.header_rows_frozen,
            header_columns_frozen: self.header_columns_frozen,
            repeating_header_rows_enabled: self.repeating_header_rows_enabled,
            repeating_header_columns_enabled: self.repeating_header_columns_enabled,
        }
    }
}

/// Presence-preserving values for the seven optional table-header fields.
///
/// `None` means that the corresponding protobuf field is absent in the
/// requested candidate.  Unknown fields and the required row/column fields
/// are never synthesized by this value; the wire rewrite copies them from the
/// source verbatim.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct TableHeaderSettingsWrite {
    pub header_rows: Option<u32>,
    pub header_columns: Option<u32>,
    pub footer_rows: Option<u32>,
    pub header_rows_frozen: Option<bool>,
    pub header_columns_frozen: Option<bool>,
    pub repeating_header_rows_enabled: Option<bool>,
    pub repeating_header_columns_enabled: Option<bool>,
}

impl TableHeaderSettingsWrite {
    /// Build a complete presence-preserving optional-field update.
    #[must_use]
    pub const fn new(
        header_rows: Option<u32>,
        header_columns: Option<u32>,
        footer_rows: Option<u32>,
        header_rows_frozen: Option<bool>,
        header_columns_frozen: Option<bool>,
        repeating_header_rows_enabled: Option<bool>,
        repeating_header_columns_enabled: Option<bool>,
    ) -> Self {
        Self {
            header_rows,
            header_columns,
            footer_rows,
            header_rows_frozen,
            header_columns_frozen,
            repeating_header_rows_enabled,
            repeating_header_columns_enabled,
        }
    }

    /// Build an update from a decoded snapshot.
    #[must_use]
    pub const fn from_snapshot(snapshot: TableHeaderSettingsSnapshot) -> Self {
        snapshot.optional_write()
    }
}

/// Compatibility spelling for callers that describe this value as an update.
pub type TableHeaderSettingsUpdate = TableHeaderSettingsWrite;

/// Compatibility spelling for callers that describe this value as a patch.
pub type TableHeaderSettingsPatch = TableHeaderSettingsWrite;
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum WireResourceLimit {
    Bytes { observed: usize, maximum: usize },
    Nesting { observed: u32, maximum: u32 },
}
/// Strict wire-preflight or private projection failure.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DecodeError(Kind);
#[derive(Debug, Clone, PartialEq, Eq)]
enum Kind {
    Wire(buffa::DecodeError),
    Resource(WireResourceLimit),
    Missing(&'static str),
    Duplicate(&'static str),
    NonCanonical(&'static str),
    Field { observed: usize, maximum: usize },
    Work { observed: usize, maximum: usize },
    Output { observed: usize, maximum: usize },
    Allocation { amount: usize },
    Projection,
}
impl DecodeError {
    const fn resource(x: WireResourceLimit) -> Self {
        Self(Kind::Resource(x))
    }
    const fn missing(x: &'static str) -> Self {
        Self(Kind::Missing(x))
    }
    const fn duplicate(x: &'static str) -> Self {
        Self(Kind::Duplicate(x))
    }
    const fn noncanonical(x: &'static str) -> Self {
        Self(Kind::NonCanonical(x))
    }
    #[must_use]
    pub const fn wire_resource_limit(&self) -> Option<WireResourceLimit> {
        if let Kind::Resource(x) = self.0 {
            Some(x)
        } else {
            None
        }
    }
    #[must_use]
    pub const fn missing_required_field(&self) -> Option<&'static str> {
        if let Kind::Missing(x) = self.0 {
            Some(x)
        } else {
            None
        }
    }
    #[must_use]
    pub const fn duplicate_singular_field(&self) -> Option<&'static str> {
        if let Kind::Duplicate(x) = self.0 {
            Some(x)
        } else {
            None
        }
    }
    #[must_use]
    pub const fn noncanonical_reason(&self) -> Option<&'static str> {
        if let Kind::NonCanonical(x) = self.0 {
            Some(x)
        } else {
            None
        }
    }
    #[must_use]
    pub const fn field_limit_values(&self) -> Option<(usize, usize)> {
        if let Kind::Field { observed, maximum } = self.0 {
            Some((observed, maximum))
        } else {
            None
        }
    }
    #[must_use]
    pub const fn work_limit_values(&self) -> Option<(usize, usize)> {
        if let Kind::Work { observed, maximum } = self.0 {
            Some((observed, maximum))
        } else {
            None
        }
    }

    /// Return the exact candidate-output limit observation, when applicable.
    #[must_use]
    pub const fn output_limit_values(&self) -> Option<(usize, usize)> {
        if let Kind::Output { observed, maximum } = self.0 {
            Some((observed, maximum))
        } else {
            None
        }
    }

    /// Return the requested output allocation, when reservation failed.
    #[must_use]
    pub const fn allocation_amount(&self) -> Option<usize> {
        if let Kind::Allocation { amount } = self.0 {
            Some(amount)
        } else {
            None
        }
    }
}
impl fmt::Display for DecodeError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match &self.0 {
            Kind::Wire(x) => x.fmt(f),
            Kind::Resource(_) => f.write_str("Numbers table-header resource limit exceeded"),
            Kind::Missing(x) => write!(f, "missing required field {x}"),
            Kind::Duplicate(x) => write!(f, "duplicate singular field {x}"),
            Kind::NonCanonical(x) => write!(f, "non-canonical protobuf representation: {x}"),
            Kind::Field { observed, maximum } => {
                write!(f, "visited {observed} fields; maximum is {maximum}")
            },
            Kind::Work { observed, maximum } => {
                write!(f, "requires {observed} work bytes; maximum is {maximum}")
            },
            Kind::Output { observed, maximum } => {
                write!(f, "produced {observed} output bytes; maximum is {maximum}")
            },
            Kind::Allocation { amount } => {
                write!(f, "cannot allocate table-header output for {amount} bytes")
            },
            Kind::Projection => f.write_str("strict preflight disagrees with Buffa projection"),
        }
    }
}
impl std::error::Error for DecodeError {}

/// Exact aggregate consumption for one table-header rewrite.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RewriteReport {
    input_bytes: usize,
    output_bytes: usize,
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    allocations: usize,
    changed: bool,
}

impl RewriteReport {
    /// Source payload bytes inspected before publication.
    #[must_use]
    pub const fn input_bytes(self) -> usize {
        self.input_bytes
    }

    /// Exact candidate payload size measured before allocation.
    #[must_use]
    pub const fn output_bytes(self) -> usize {
        self.output_bytes
    }

    /// Aggregate strict field visits across source, rewrite, and readback.
    #[must_use]
    pub const fn fields(self) -> usize {
        self.fields
    }

    /// Aggregate strict wire work across the complete operation.
    #[must_use]
    pub const fn work_bytes(self) -> usize {
        self.work_bytes
    }

    /// Maximum nesting observed by strict source and candidate scans.
    #[must_use]
    pub const fn max_depth(self) -> u32 {
        self.max_depth
    }

    /// Number of fallible output reservations performed.
    #[must_use]
    pub const fn allocations(self) -> usize {
        self.allocations
    }

    /// Whether any selected optional field changed.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }
}
#[allow(
    clippy::wildcard_enum_match_arm,
    reason = "Buffa errors are non-exhaustive and unrecognized failures remain opaque wire errors."
)]
impl From<buffa::DecodeError> for DecodeError {
    fn from(error: buffa::DecodeError) -> Self {
        match error {
            buffa::DecodeError::MessageTooLarge | buffa::DecodeError::RecursionLimitExceeded => {
                Self(Kind::Projection)
            },
            other => Self(Kind::Wire(other)),
        }
    }
}

struct Budget {
    fields: usize,
    work: usize,
    max_fields: usize,
    max_work: usize,
    max_recursion: u32,
    max_depth: u32,
}
impl Budget {
    const fn new(options: DecodeOptions) -> Self {
        Self {
            fields: 0,
            work: 0,
            max_fields: options.fields,
            max_work: options.work,
            max_recursion: options.recursion,
            max_depth: 1,
        }
    }
    fn charge(&mut self, bytes: usize) -> Result<(), DecodeError> {
        let observed = bytes
            .checked_mul(2)
            .and_then(|cost| self.work.checked_add(cost))
            .ok_or(DecodeError(Kind::Projection))?;
        if observed > self.max_work {
            return Err(DecodeError(Kind::Work {
                observed,
                maximum: self.max_work,
            }));
        }
        self.work = observed;
        Ok(())
    }
    fn field(&mut self) -> Result<(), DecodeError> {
        self.fields(1)
    }
    fn fields(&mut self, amount: usize) -> Result<(), DecodeError> {
        let observed = self
            .fields
            .checked_add(amount)
            .ok_or(DecodeError(Kind::Projection))?;
        if observed > self.max_fields {
            return Err(DecodeError(Kind::Field {
                observed,
                maximum: self.max_fields,
            }));
        }
        self.fields = observed;
        Ok(())
    }
    const fn nesting(&self) -> DecodeError {
        DecodeError::resource(WireResourceLimit::Nesting {
            observed: self.max_recursion.saturating_add(1),
            maximum: self.max_recursion,
        })
    }
}
/// Decode one `TST.TableModelArchive` header/footer scalar envelope without retaining raw IDs.
pub fn decode_table_header_settings(
    source: &[u8],
    o: DecodeOptions,
) -> Result<TableHeaderSettingsSnapshot, DecodeError> {
    validate(source, o)?;
    let mut b = Budget::new(o);
    decode_snapshot_with_budget(source, o, &mut b)
}

fn decode_snapshot_with_budget(
    source: &[u8],
    o: DecodeOptions,
    b: &mut Budget,
) -> Result<TableHeaderSettingsSnapshot, DecodeError> {
    decode_snapshot_with_budget_mode(source, o, b, true)
}

fn decode_snapshot_with_budget_mode(
    source: &[u8],
    o: DecodeOptions,
    b: &mut Budget,
    strict_unknown: bool,
) -> Result<TableHeaderSettingsSnapshot, DecodeError> {
    let strict = preflight_mode(source, o, b, strict_unknown)?;
    let v: projection::NumbersTableHeaderSettingsArchiveLazyView<'_> =
        o.buffa().decode_lazy_view(source)?;
    let projected = TableHeaderSettingsSnapshot {
        rows: v.number_of_rows,
        columns: v.number_of_columns,
        header_rows: v.number_of_header_rows,
        header_columns: v.number_of_header_columns,
        footer_rows: v.number_of_footer_rows,
        header_rows_frozen: v.header_rows_frozen,
        header_columns_frozen: v.header_columns_frozen,
        repeating_header_rows_enabled: v.repeating_header_rows_enabled,
        repeating_header_columns_enabled: v.repeating_header_columns_enabled,
    };
    if projected != strict {
        return Err(DecodeError(Kind::Projection));
    }
    Ok(strict)
}

/// Rewrite the seven optional header/footer/freeze/repeat fields while
/// retaining every required and unknown source span.
pub fn rewrite_table_header_settings(
    source: &[u8],
    write: TableHeaderSettingsWrite,
    options: DecodeOptions,
) -> Result<Vec<u8>, DecodeError> {
    Ok(rewrite_table_header_settings_with_report(source, write, options)?.0)
}

/// Rewrite one table-header payload and return exact aggregate accounting.
pub fn rewrite_table_header_settings_with_report(
    source: &[u8],
    write: TableHeaderSettingsWrite,
    options: DecodeOptions,
) -> Result<(Vec<u8>, RewriteReport), DecodeError> {
    validate(source, options)?;
    let mut budget = Budget::new(options);
    let current = decode_snapshot_with_budget_mode(source, options, &mut budget, false)?;
    let changed = current.optional_write() != write;
    if !changed {
        if source.len() > options.output_bytes {
            return Err(DecodeError(Kind::Output {
                observed: source.len(),
                maximum: options.output_bytes,
            }));
        }
        let output = reserve_output(source.len())?;
        let mut output = output;
        append_bytes(&mut output, source)?;
        return Ok((
            output,
            RewriteReport {
                input_bytes: source.len(),
                output_bytes: source.len(),
                fields: budget.fields,
                work_bytes: budget.work,
                max_depth: budget.max_depth,
                allocations: 1,
                changed: false,
            },
        ));
    }

    let output_bytes = measure_rewrite_output(source, write, options, &mut budget)?;
    if output_bytes > options.output_bytes {
        return Err(DecodeError(Kind::Output {
            observed: output_bytes,
            maximum: options.output_bytes,
        }));
    }
    precharge_candidate_readback(source, write, output_bytes, options, &mut budget)?;
    let mut output = reserve_output(output_bytes)?;
    let mut emit_budget = Budget::new(DecodeOptions {
        fields: usize::MAX,
        work: usize::MAX,
        ..options
    });
    emit_rewrite_output(source, write, options, &mut emit_budget, &mut output)?;
    if output.len() != output_bytes {
        return Err(DecodeError(Kind::Projection));
    }

    let readback_options = DecodeOptions {
        bytes: options.bytes.max(output.len()),
        output_bytes: options.output_bytes,
        ..options
    };
    validate(&output, readback_options)?;
    let mut readback_budget = Budget::new(DecodeOptions {
        fields: usize::MAX,
        work: usize::MAX,
        ..readback_options
    });
    let readback =
        decode_snapshot_with_budget_mode(&output, readback_options, &mut readback_budget, false)?;
    if readback.optional_write() != write {
        return Err(DecodeError(Kind::Projection));
    }
    Ok((
        output,
        RewriteReport {
            input_bytes: source.len(),
            output_bytes,
            fields: budget.fields,
            work_bytes: budget.work,
            max_depth: budget.max_depth,
            allocations: 1,
            changed: true,
        },
    ))
}

/// Compatibility alias for callers that use the `rewrite_*_with_report`
/// spelling without the Numbers-specific module prefix.
pub fn rewrite_table_header_settings_extension(
    source: &[u8],
    write: TableHeaderSettingsWrite,
    options: DecodeOptions,
) -> Result<Vec<u8>, DecodeError> {
    rewrite_table_header_settings(source, write, options)
}

#[derive(Clone, Copy)]
struct FieldSpan {
    number: u32,
    start: usize,
    end: usize,
}

fn visit_field_spans<F>(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
    mut visitor: F,
) -> Result<(), DecodeError>
where
    F: FnMut(FieldSpan) -> Result<(), DecodeError>,
{
    budget.charge(source.len())?;
    let mut remaining = source;
    while !remaining.is_empty() {
        let start = source.len() - remaining.len();
        let (tag, canonical) = varint(&mut remaining)?;
        let raw =
            u32::try_from(tag).map_err(|_conversion| buffa::DecodeError::InvalidFieldNumber)?;
        let number = raw >> 3;
        let wire = raw & 7;
        if number == 0 {
            return Err(buffa::DecodeError::InvalidFieldNumber.into());
        }
        if !canonical && known_field(number) {
            return Err(DecodeError::noncanonical("protobuf field key"));
        }
        budget.field()?;
        if wire == 0 {
            let (_, value_canonical) = varint(&mut remaining)?;
            if !value_canonical && known_field(number) {
                return Err(DecodeError::noncanonical("protobuf varint value"));
            }
        } else {
            skip(
                &mut remaining,
                number,
                wire,
                options.recursion,
                budget,
                false,
            )?;
        }
        let end = source.len() - remaining.len();
        visitor(FieldSpan { number, start, end })?;
    }
    Ok(())
}

fn selected_field(number: u32) -> bool {
    matches!(
        number,
        HEADER_ROWS_FIELD
            | HEADER_COLUMNS_FIELD
            | FOOTER_ROWS_FIELD
            | HEADER_ROWS_FROZEN_FIELD
            | HEADER_COLUMNS_FROZEN_FIELD
            | REPEATING_HEADER_ROWS_FIELD
            | REPEATING_HEADER_COLUMNS_FIELD
    )
}

fn known_field(number: u32) -> bool {
    matches!(number, TABLE_ROWS_FIELD | TABLE_COLUMNS_FIELD) || selected_field(number)
}

fn requested_value(write: TableHeaderSettingsWrite, number: u32) -> Option<u64> {
    match number {
        HEADER_ROWS_FIELD => write.header_rows.map(u64::from),
        HEADER_COLUMNS_FIELD => write.header_columns.map(u64::from),
        FOOTER_ROWS_FIELD => write.footer_rows.map(u64::from),
        HEADER_ROWS_FROZEN_FIELD => write.header_rows_frozen.map(u64::from),
        HEADER_COLUMNS_FROZEN_FIELD => write.header_columns_frozen.map(u64::from),
        REPEATING_HEADER_ROWS_FIELD => write.repeating_header_rows_enabled.map(u64::from),
        REPEATING_HEADER_COLUMNS_FIELD => write.repeating_header_columns_enabled.map(u64::from),
        _ => None,
    }
}

fn varint_len(mut value: u64) -> usize {
    let mut length = 1;
    while value >= 128 {
        value >>= 7;
        length += 1;
    }
    length
}

fn encoded_varint_field_len(number: u32, value: u64) -> usize {
    varint_len(u64::from(number) << 3) + varint_len(value)
}

fn measure_rewrite_output(
    source: &[u8],
    write: TableHeaderSettingsWrite,
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<usize, DecodeError> {
    let mut output_bytes = 0usize;
    let mut seen = [false; 7];
    visit_field_spans(source, options, budget, |span| {
        if selected_field(span.number) {
            let index = optional_index(span.number).ok_or(DecodeError(Kind::Projection))?;
            seen[index] = true;
            if let Some(value) = requested_value(write, span.number) {
                output_bytes = output_bytes
                    .checked_add(encoded_varint_field_len(span.number, value))
                    .ok_or(DecodeError(Kind::Projection))?;
            }
        } else {
            output_bytes = output_bytes
                .checked_add(span.end - span.start)
                .ok_or(DecodeError(Kind::Projection))?;
        }
        Ok(())
    })?;
    for (index, number) in optional_fields().iter().copied().enumerate() {
        if !seen[index]
            && let Some(value) = requested_value(write, number)
        {
            output_bytes = output_bytes
                .checked_add(encoded_varint_field_len(number, value))
                .ok_or(DecodeError(Kind::Projection))?;
        }
    }
    Ok(output_bytes)
}

fn precharge_candidate_readback(
    source: &[u8],
    write: TableHeaderSettingsWrite,
    output_bytes: usize,
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<(), DecodeError> {
    let scan_options = DecodeOptions {
        fields: usize::MAX,
        work: usize::MAX,
        ..options
    };
    let mut scan_budget = Budget::new(scan_options);
    let mut selected_present = [false; 7];
    visit_field_spans(source, scan_options, &mut scan_budget, |span| {
        if let Some(index) = optional_index(span.number) {
            selected_present[index] = true;
        }
        Ok(())
    })?;
    let selected_source_fields = selected_present.iter().filter(|present| **present).count();
    let selected_candidate_fields = optional_fields()
        .iter()
        .filter(|field| requested_value(write, **field).is_some())
        .count();
    let candidate_fields = scan_budget
        .fields
        .checked_sub(selected_source_fields)
        .and_then(|fields| fields.checked_add(selected_candidate_fields))
        .ok_or(DecodeError(Kind::Projection))?;
    budget.charge(source.len())?;
    budget.fields(scan_budget.fields)?;
    budget.fields(candidate_fields)?;
    budget.charge(output_bytes)?;
    // Emission replays the source spans after the sole output reservation.
    // Charge that traversal now so a field/work ceiling can never fail after
    // the candidate Vec exists.
    budget.charge(source.len())?;
    budget.fields(scan_budget.fields)?;
    budget.max_depth = budget.max_depth.max(scan_budget.max_depth);
    Ok(())
}

fn emit_rewrite_output(
    source: &[u8],
    write: TableHeaderSettingsWrite,
    options: DecodeOptions,
    budget: &mut Budget,
    output: &mut Vec<u8>,
) -> Result<(), DecodeError> {
    let mut seen = [false; 7];
    visit_field_spans(source, options, budget, |span| {
        if selected_field(span.number) {
            let index = optional_index(span.number).ok_or(DecodeError(Kind::Projection))?;
            seen[index] = true;
            if let Some(value) = requested_value(write, span.number) {
                append_varint_field(output, span.number, value)?;
            }
        } else {
            append_bytes(output, &source[span.start..span.end])?;
        }
        Ok(())
    })?;
    for (index, number) in optional_fields().iter().copied().enumerate() {
        if !seen[index]
            && let Some(value) = requested_value(write, number)
        {
            append_varint_field(output, number, value)?;
        }
    }
    Ok(())
}

const fn optional_fields() -> [u32; 7] {
    [
        HEADER_ROWS_FIELD,
        HEADER_COLUMNS_FIELD,
        FOOTER_ROWS_FIELD,
        HEADER_ROWS_FROZEN_FIELD,
        HEADER_COLUMNS_FROZEN_FIELD,
        REPEATING_HEADER_ROWS_FIELD,
        REPEATING_HEADER_COLUMNS_FIELD,
    ]
}

const fn optional_index(number: u32) -> Option<usize> {
    match number {
        HEADER_ROWS_FIELD => Some(0),
        HEADER_COLUMNS_FIELD => Some(1),
        FOOTER_ROWS_FIELD => Some(2),
        HEADER_ROWS_FROZEN_FIELD => Some(3),
        HEADER_COLUMNS_FROZEN_FIELD => Some(4),
        REPEATING_HEADER_ROWS_FIELD => Some(5),
        REPEATING_HEADER_COLUMNS_FIELD => Some(6),
        _ => None,
    }
}

fn reserve_output(amount: usize) -> Result<Vec<u8>, DecodeError> {
    let mut output = Vec::new();
    output
        .try_reserve_exact(amount)
        .map_err(|_error| DecodeError(Kind::Allocation { amount }))?;
    record_output_allocation();
    Ok(output)
}

fn append_bytes(output: &mut Vec<u8>, bytes: &[u8]) -> Result<(), DecodeError> {
    let remaining = output
        .capacity()
        .checked_sub(output.len())
        .ok_or(DecodeError(Kind::Projection))?;
    if bytes.len() > remaining {
        return Err(DecodeError(Kind::Projection));
    }
    output.extend_from_slice(bytes);
    Ok(())
}

fn append_varint_field(output: &mut Vec<u8>, number: u32, value: u64) -> Result<(), DecodeError> {
    let required = encoded_varint_field_len(number, value);
    let remaining = output
        .capacity()
        .checked_sub(output.len())
        .ok_or(DecodeError(Kind::Projection))?;
    if required > remaining {
        return Err(DecodeError(Kind::Projection));
    }
    push_varint(u64::from(number) << 3, output);
    push_varint(value, output);
    Ok(())
}

fn push_varint(mut value: u64, output: &mut Vec<u8>) {
    while value >= 128 {
        output.push((value as u8 & 127) | 128);
        value >>= 7;
    }
    output.push(value as u8);
}

#[cfg(test)]
thread_local! {
    static OUTPUT_ALLOCATIONS: std::cell::Cell<usize> = const { std::cell::Cell::new(0) };
}

#[inline]
fn record_output_allocation() {
    #[cfg(test)]
    OUTPUT_ALLOCATIONS.with(|count| count.set(count.get().saturating_add(1)));
}

#[cfg(test)]
fn output_allocations() -> usize {
    OUTPUT_ALLOCATIONS.with(std::cell::Cell::get)
}

fn validate(source: &[u8], options: DecodeOptions) -> Result<(), DecodeError> {
    let hard = usize::try_from(buffa::MAX_MESSAGE_BYTES)
        .map_err(|_conversion| DecodeError(Kind::Projection))?;
    if options.bytes > hard {
        return Err(DecodeError::resource(WireResourceLimit::Bytes {
            observed: options.bytes,
            maximum: hard,
        }));
    }
    if source.len() > options.bytes {
        return Err(DecodeError::resource(WireResourceLimit::Bytes {
            observed: source.len(),
            maximum: options.bytes,
        }));
    }
    if options.recursion == 0 || options.recursion > MAX_RECURSION {
        return Err(DecodeError::resource(WireResourceLimit::Nesting {
            observed: options.recursion,
            maximum: MAX_RECURSION,
        }));
    }
    Ok(())
}
fn preflight_mode(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
    strict_unknown: bool,
) -> Result<TableHeaderSettingsSnapshot, DecodeError> {
    budget.charge(source.len())?;
    let mut snapshot = TableHeaderSettingsSnapshot {
        rows: 0,
        columns: 0,
        header_rows: None,
        header_columns: None,
        footer_rows: None,
        header_rows_frozen: None,
        header_columns_frozen: None,
        repeating_header_rows_enabled: None,
        repeating_header_columns_enabled: None,
    };
    let mut seen = 0u64;
    let mut remaining = source;
    while !remaining.is_empty() {
        let (tag, key_canonical) = varint(&mut remaining)?;
        let raw =
            u32::try_from(tag).map_err(|_conversion| buffa::DecodeError::InvalidFieldNumber)?;
        let field_number = raw >> 3;
        if field_number == 0 {
            return Err(buffa::DecodeError::InvalidFieldNumber.into());
        }
        if !key_canonical && (strict_unknown || known_field(field_number)) {
            return Err(DecodeError::noncanonical("protobuf field key"));
        }
        budget.field()?;
        if raw & 7 != 0 {
            skip(
                &mut remaining,
                field_number,
                raw & 7,
                options.recursion,
                budget,
                strict_unknown,
            )?;
            if known_field(field_number) {
                return Err(buffa::DecodeError::WireTypeMismatch {
                    field_number,
                    expected: 0,
                    actual: (raw & 7) as u8,
                }
                .into());
            }
            continue;
        }
        let (value, value_canonical) = varint(&mut remaining)?;
        if !value_canonical && (strict_unknown || known_field(field_number)) {
            return Err(DecodeError::noncanonical("protobuf varint value"));
        }
        let bit = 1u64.checked_shl(field_number).unwrap_or(0);
        let name = match field_number {
            TABLE_ROWS_FIELD => "TST.TableModelArchive.number_of_rows",
            TABLE_COLUMNS_FIELD => "TST.TableModelArchive.number_of_columns",
            HEADER_ROWS_FIELD => "TST.TableModelArchive.number_of_header_rows",
            HEADER_COLUMNS_FIELD => "TST.TableModelArchive.number_of_header_columns",
            FOOTER_ROWS_FIELD => "TST.TableModelArchive.number_of_footer_rows",
            HEADER_ROWS_FROZEN_FIELD => "TST.TableModelArchive.header_rows_frozen",
            HEADER_COLUMNS_FROZEN_FIELD => "TST.TableModelArchive.header_columns_frozen",
            REPEATING_HEADER_ROWS_FIELD => "TST.TableModelArchive.repeating_header_rows_enabled",
            REPEATING_HEADER_COLUMNS_FIELD => {
                "TST.TableModelArchive.repeating_header_columns_enabled"
            },
            _ => continue,
        };
        if seen & bit != 0 {
            return Err(DecodeError::duplicate(name));
        }
        seen |= bit;
        let parsed_u32 = u32::try_from(value)
            .map_err(|_conversion| DecodeError::noncanonical("uint32 scalar exceeds u32"))?;
        match field_number {
            TABLE_ROWS_FIELD => snapshot.rows = parsed_u32,
            TABLE_COLUMNS_FIELD => snapshot.columns = parsed_u32,
            HEADER_ROWS_FIELD => snapshot.header_rows = Some(parsed_u32),
            HEADER_COLUMNS_FIELD => snapshot.header_columns = Some(parsed_u32),
            FOOTER_ROWS_FIELD => snapshot.footer_rows = Some(parsed_u32),
            HEADER_ROWS_FROZEN_FIELD => snapshot.header_rows_frozen = Some(boolean(value)?),
            HEADER_COLUMNS_FROZEN_FIELD => snapshot.header_columns_frozen = Some(boolean(value)?),
            REPEATING_HEADER_ROWS_FIELD => {
                snapshot.repeating_header_rows_enabled = Some(boolean(value)?);
            },
            REPEATING_HEADER_COLUMNS_FIELD => {
                snapshot.repeating_header_columns_enabled = Some(boolean(value)?);
            },
            _ => {},
        }
    }
    if seen & (1 << TABLE_ROWS_FIELD) == 0 {
        return Err(DecodeError::missing("TST.TableModelArchive.number_of_rows"));
    }
    if seen & (1 << TABLE_COLUMNS_FIELD) == 0 {
        return Err(DecodeError::missing(
            "TST.TableModelArchive.number_of_columns",
        ));
    }
    Ok(snapshot)
}
fn boolean(value: u64) -> Result<bool, DecodeError> {
    match value {
        0 => Ok(false),
        1 => Ok(true),
        _ => Err(DecodeError::noncanonical("bool scalar is not zero or one")),
    }
}
fn varint(source: &mut &[u8]) -> Result<(u64, bool), DecodeError> {
    let original = *source;
    let mut value = 0;
    for index in 0..10 {
        let byte = *original
            .get(index)
            .ok_or(buffa::DecodeError::UnexpectedEof)?;
        if index == 9 && byte > 1 {
            return Err(buffa::DecodeError::VarintTooLong.into());
        }
        value |= u64::from(byte & 127) << (index * 7);
        if byte & 128 == 0 {
            *source = &original[index + 1..];
            let mut remaining = value;
            let mut canonical_length = 1;
            while remaining >= 128 {
                remaining >>= 7;
                canonical_length += 1;
            }
            return Ok((value, canonical_length == index + 1));
        }
    }
    Err(buffa::DecodeError::VarintTooLong.into())
}
fn skip(
    s: &mut &[u8],
    number: u32,
    wire: u32,
    depth: u32,
    budget: &mut Budget,
    strict_unknown: bool,
) -> Result<(), DecodeError> {
    match wire {
        0 => {
            let (_, canonical) = varint(s)?;
            if strict_unknown && !canonical {
                return Err(DecodeError::noncanonical("protobuf varint value"));
            }
            Ok(())
        },
        1 => take(s, 8),
        2 => {
            let (length, canonical) = varint(s)?;
            if strict_unknown && !canonical {
                return Err(DecodeError::noncanonical("length-delimited size"));
            }
            take(
                s,
                usize::try_from(length)
                    .map_err(|_conversion| buffa::DecodeError::MessageTooLarge)?,
            )
        },
        3 => skip_group(
            s,
            number,
            depth.checked_sub(1).ok_or_else(|| budget.nesting())?,
            budget,
            strict_unknown,
        ),
        4 => Err(buffa::DecodeError::InvalidEndGroup(number).into()),
        5 => take(s, 4),
        _ => Err(buffa::DecodeError::InvalidWireType(wire).into()),
    }
}

fn skip_group(
    s: &mut &[u8],
    expected: u32,
    depth: u32,
    budget: &mut Budget,
    strict_unknown: bool,
) -> Result<(), DecodeError> {
    budget.max_depth = budget
        .max_depth
        .max(budget.max_recursion.saturating_sub(depth));
    loop {
        if s.is_empty() {
            return Err(buffa::DecodeError::UnexpectedEof.into());
        }
        let (tag, canonical) = varint(s)?;
        if strict_unknown && !canonical {
            return Err(DecodeError::noncanonical("protobuf field key"));
        }
        budget.field()?;
        let raw =
            u32::try_from(tag).map_err(|_conversion| buffa::DecodeError::InvalidFieldNumber)?;
        let number = raw >> 3;
        if number == 0 {
            return Err(buffa::DecodeError::InvalidFieldNumber.into());
        }
        let wire = raw & 7;
        if wire == 4 {
            if number == expected {
                return Ok(());
            }
            return Err(buffa::DecodeError::InvalidEndGroup(number).into());
        }
        skip(s, number, wire, depth, budget, strict_unknown)?;
    }
}
fn take(s: &mut &[u8], n: usize) -> Result<(), DecodeError> {
    if s.len() < n {
        return Err(buffa::DecodeError::UnexpectedEof.into());
    }
    *s = &s[n..];
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn push_varint(mut value: u64, output: &mut Vec<u8>) {
        while value >= 128 {
            output.push((value as u8 & 127) | 128);
            value >>= 7;
        }
        output.push(value as u8);
    }

    fn varint_field(number: u32, value: u64) -> Vec<u8> {
        let mut output = Vec::new();
        push_varint(u64::from(number) << 3, &mut output);
        push_varint(value, &mut output);
        output
    }

    fn dimensions() -> Vec<u8> {
        let mut output = varint_field(TABLE_ROWS_FIELD, 1);
        output.extend(varint_field(TABLE_COLUMNS_FIELD, 1));
        output
    }

    #[test]
    fn preserves_header_presence_and_rejects_strict_failures() -> Result<(), DecodeError> {
        let source = [
            0x30, 10, 0x38, 5, 0x48, 2, 0x50, 3, 0x58, 1, 0x60, 1, 0x68, 0, 0xe8, 1, 1, 0x80, 2, 0,
        ];
        let snapshot = decode_table_header_settings(
            &source,
            DecodeOptions::new(source.len(), 9, source.len() * 2, 2),
        )?;
        assert_eq!(snapshot.rows(), 10);
        assert_eq!(snapshot.columns(), 5);
        assert_eq!(snapshot.header_rows(), Some(2));
        assert_eq!(snapshot.header_columns(), Some(3));
        assert_eq!(snapshot.footer_rows(), Some(1));
        assert_eq!(snapshot.header_rows_frozen(), Some(true));
        assert_eq!(snapshot.header_columns_frozen(), Some(false));
        assert_eq!(snapshot.repeating_header_rows_enabled(), Some(true));
        assert_eq!(snapshot.repeating_header_columns_enabled(), Some(false));
        let duplicate = [0x30, 1, 0x30, 2, 0x38, 1];
        assert_eq!(
            decode_table_header_settings(
                &duplicate,
                DecodeOptions::new(duplicate.len(), 3, duplicate.len() * 2, 2)
            )
            .expect_err("duplicate")
            .duplicate_singular_field(),
            Some("TST.TableModelArchive.number_of_rows")
        );
        let bad_bool = [0x30, 1, 0x38, 1, 0x60, 2];
        assert_eq!(
            decode_table_header_settings(
                &bad_bool,
                DecodeOptions::new(bad_bool.len(), 3, bad_bool.len() * 2, 2)
            )
            .expect_err("bool")
            .noncanonical_reason(),
            Some("bool scalar is not zero or one")
        );
        Ok(())
    }

    #[test]
    fn unknown_groups_are_bounded_and_canonical() -> Result<(), DecodeError> {
        let one_group = [0x30, 1, 0x38, 1, 0xa3, 0x06, 0x08, 1, 0xa4, 0x06];
        assert!(
            decode_table_header_settings(
                &one_group,
                DecodeOptions::new(one_group.len(), 5, one_group.len() * 2, 1)
            )
            .is_ok()
        );
        assert_eq!(
            decode_table_header_settings(
                &one_group,
                DecodeOptions::new(one_group.len(), 4, one_group.len() * 2, 1)
            )
            .expect_err("fields")
            .field_limit_values(),
            Some((5, 4))
        );
        let dimensions = [0x30, 1, 0x38, 1];
        assert_eq!(
            decode_table_header_settings(&dimensions, DecodeOptions::new(4, 2, 7, 1))
                .expect_err("work")
                .work_limit_values(),
            Some((8, 7))
        );
        let nested_groups = [
            0x30, 1, 0x38, 1, 0xa3, 0x06, 0xab, 0x06, 0xac, 0x06, 0xa4, 0x06,
        ];
        assert_eq!(
            decode_table_header_settings(
                &nested_groups,
                DecodeOptions::new(nested_groups.len(), 5, nested_groups.len() * 2, 1)
            )
            .expect_err("depth")
            .wire_resource_limit(),
            Some(WireResourceLimit::Nesting {
                observed: 2,
                maximum: 1
            })
        );
        let mismatched = [0x30, 1, 0x38, 1, 0xa3, 0x06, 0xac, 0x06];
        assert!(
            decode_table_header_settings(
                &mismatched,
                DecodeOptions::new(mismatched.len(), 4, mismatched.len() * 2, 2)
            )
            .is_err()
        );
        let noncanonical_length = [0x30, 1, 0x38, 1, 0x82, 0x06, 0x80, 0, 0];
        assert_eq!(
            decode_table_header_settings(
                &noncanonical_length,
                DecodeOptions::new(
                    noncanonical_length.len(),
                    3,
                    noncanonical_length.len() * 2,
                    2
                )
            )
            .expect_err("length")
            .noncanonical_reason(),
            Some("length-delimited size")
        );
        Ok(())
    }

    #[test]
    fn table_driven_strict_wire_regressions() {
        let dims = dimensions();
        for (source, required) in [
            (
                varint_field(TABLE_ROWS_FIELD, 1),
                "TST.TableModelArchive.number_of_columns",
            ),
            (
                varint_field(TABLE_COLUMNS_FIELD, 1),
                "TST.TableModelArchive.number_of_rows",
            ),
        ] {
            assert_eq!(
                decode_table_header_settings(
                    &source,
                    DecodeOptions::new(source.len(), 9, source.len() * 2, 2)
                )
                .expect_err("missing dimension")
                .missing_required_field(),
                Some(required)
            );
        }
        for (field, name) in [
            (TABLE_ROWS_FIELD, "TST.TableModelArchive.number_of_rows"),
            (
                TABLE_COLUMNS_FIELD,
                "TST.TableModelArchive.number_of_columns",
            ),
            (
                HEADER_ROWS_FIELD,
                "TST.TableModelArchive.number_of_header_rows",
            ),
            (
                HEADER_COLUMNS_FIELD,
                "TST.TableModelArchive.number_of_header_columns",
            ),
            (
                FOOTER_ROWS_FIELD,
                "TST.TableModelArchive.number_of_footer_rows",
            ),
            (
                HEADER_ROWS_FROZEN_FIELD,
                "TST.TableModelArchive.header_rows_frozen",
            ),
            (
                HEADER_COLUMNS_FROZEN_FIELD,
                "TST.TableModelArchive.header_columns_frozen",
            ),
            (
                REPEATING_HEADER_ROWS_FIELD,
                "TST.TableModelArchive.repeating_header_rows_enabled",
            ),
            (
                REPEATING_HEADER_COLUMNS_FIELD,
                "TST.TableModelArchive.repeating_header_columns_enabled",
            ),
        ] {
            let mut source = dims.clone();
            source.extend(varint_field(field, 0));
            source.extend(varint_field(field, 0));
            assert_eq!(
                decode_table_header_settings(
                    &source,
                    DecodeOptions::new(source.len(), 16, source.len() * 2, 2)
                )
                .expect_err("duplicate")
                .duplicate_singular_field(),
                Some(name)
            );
        }
        for field in [TABLE_ROWS_FIELD, HEADER_ROWS_FIELD] {
            let mut source = dims.clone();
            push_varint((u64::from(field) << 3) | 5, &mut source);
            source.extend([0; 4]);
            assert!(
                decode_table_header_settings(
                    &source,
                    DecodeOptions::new(source.len(), 8, source.len() * 2, 2)
                )
                .is_err()
            );
        }
        for field in [HEADER_ROWS_FROZEN_FIELD, REPEATING_HEADER_COLUMNS_FIELD] {
            let mut source = dims.clone();
            push_varint((u64::from(field) << 3) | 2, &mut source);
            source.push(0);
            assert!(
                decode_table_header_settings(
                    &source,
                    DecodeOptions::new(source.len(), 8, source.len() * 2, 2)
                )
                .is_err()
            );
        }
        let mut overflow = dims.clone();
        overflow.extend(varint_field(HEADER_ROWS_FIELD, u64::from(u32::MAX) + 1));
        assert_eq!(
            decode_table_header_settings(
                &overflow,
                DecodeOptions::new(overflow.len(), 8, overflow.len() * 2, 2)
            )
            .expect_err("overflow")
            .noncanonical_reason(),
            Some("uint32 scalar exceeds u32")
        );
        let mut bool_two = dims.clone();
        bool_two.extend(varint_field(HEADER_ROWS_FROZEN_FIELD, 2));
        assert_eq!(
            decode_table_header_settings(
                &bool_two,
                DecodeOptions::new(bool_two.len(), 8, bool_two.len() * 2, 2)
            )
            .expect_err("bool")
            .noncanonical_reason(),
            Some("bool scalar is not zero or one")
        );
        for source in [
            vec![0xb0, 0, 1, 0x38, 1],
            vec![0x30, 0x81, 0, 0x38, 1],
            vec![0x30],
        ] {
            assert!(
                decode_table_header_settings(
                    &source,
                    DecodeOptions::new(source.len(), 8, source.len() * 2, 2)
                )
                .is_err()
            );
        }
    }

    #[test]
    fn preserves_all_unknown_wire_forms() -> Result<(), DecodeError> {
        let mut source = dimensions();
        source.extend(varint_field(100, 1));
        push_varint((101_u64 << 3) | 1, &mut source);
        source.extend([0; 8]);
        push_varint((102_u64 << 3) | 2, &mut source);
        source.extend([2, 9, 8]);
        push_varint((103_u64 << 3) | 3, &mut source);
        source.extend([8, 1]);
        push_varint((103_u64 << 3) | 4, &mut source);
        push_varint((104_u64 << 3) | 5, &mut source);
        source.extend([0; 4]);
        assert!(
            decode_table_header_settings(
                &source,
                DecodeOptions::new(source.len(), 9, source.len() * 2, 2)
            )
            .is_ok()
        );
        let unclosed = [0x30, 1, 0x38, 1, 0xa3, 0x06];
        assert!(
            decode_table_header_settings(
                &unclosed,
                DecodeOptions::new(unclosed.len(), 3, unclosed.len() * 2, 2)
            )
            .is_err()
        );
        let noncanonical_length = [0x30, 1, 0x38, 1, 0xb2, 0x06, 0x80, 0, 0];
        assert_eq!(
            decode_table_header_settings(
                &noncanonical_length,
                DecodeOptions::new(
                    noncanonical_length.len(),
                    3,
                    noncanonical_length.len() * 2,
                    2
                )
            )
            .expect_err("length")
            .noncanonical_reason(),
            Some("length-delimited size")
        );
        Ok(())
    }

    #[test]
    fn rewrite_changes_optional_presence_and_retains_unknown_source_spans()
    -> Result<(), DecodeError> {
        let mut source = dimensions();
        source.extend(varint_field(100, 7));
        source.extend(varint_field(HEADER_ROWS_FIELD, 2));
        source.extend(varint_field(FOOTER_ROWS_FIELD, 1));
        source.extend([0xe8, 1, 1]);
        let write = TableHeaderSettingsWrite::new(
            Some(3),
            None,
            Some(2),
            Some(false),
            None,
            Some(true),
            Some(false),
        );
        let options = DecodeOptions::for_source(&source).with_max_output_bytes(source.len() + 64);
        let (rewritten, report) =
            rewrite_table_header_settings_with_report(&source, write, options)?;
        assert!(report.changed());
        assert_eq!(report.allocations(), 1);
        assert!(rewritten.windows(2).any(|window| window == [0xa0, 6]));
        assert!(rewritten.windows(2).any(|window| window == [0xe8, 1]));
        assert!(rewritten.windows(2).any(|window| window == [0xa0, 6]));
        assert!(!rewritten.windows(2).any(|window| window == [0x50, 2]));
        assert_eq!(
            decode_table_header_settings(&rewritten, DecodeOptions::for_source(&rewritten))?,
            TableHeaderSettingsSnapshot {
                rows: 1,
                columns: 1,
                header_rows: Some(3),
                header_columns: None,
                footer_rows: Some(2),
                header_rows_frozen: Some(false),
                header_columns_frozen: None,
                repeating_header_rows_enabled: Some(true),
                repeating_header_columns_enabled: Some(false),
            }
        );
        Ok(())
    }

    #[test]
    fn rewrite_appends_absent_selected_fields_and_is_exactly_bounded() -> Result<(), DecodeError> {
        let source = dimensions();
        let write = TableHeaderSettingsWrite::new(
            Some(1),
            Some(2),
            None,
            None,
            Some(true),
            None,
            Some(false),
        );
        let options = DecodeOptions::for_source(&source).with_max_output_bytes(source.len() + 32);
        let (rewritten, report) =
            rewrite_table_header_settings_with_report(&source, write, options)?;
        assert_eq!(report.output_bytes(), rewritten.len());
        assert_eq!(report.allocations(), 1);
        let exact = options.with_max_output_bytes(rewritten.len());
        assert!(rewrite_table_header_settings_with_report(&source, write, exact).is_ok());
        assert!(
            rewrite_table_header_settings_with_report(
                &source,
                write,
                options.with_max_output_bytes(rewritten.len() - 1)
            )
            .is_err()
        );
        Ok(())
    }

    #[test]
    fn rewrite_noop_copies_source_and_selected_wire_remains_strict() -> Result<(), DecodeError> {
        let mut source = dimensions();
        source.extend(varint_field(HEADER_ROWS_FIELD, 2));
        let write = TableHeaderSettingsWrite::new(Some(2), None, None, None, None, None, None);
        let options = DecodeOptions::for_source(&source);
        let (rewritten, report) =
            rewrite_table_header_settings_with_report(&source, write, options)?;
        assert_eq!(rewritten, source);
        assert!(!report.changed());
        assert_eq!(report.output_bytes(), source.len());

        let mut malformed = dimensions();
        malformed.extend([0x50, 0x81, 0]);
        let error = rewrite_table_header_settings_with_report(
            &malformed,
            write,
            DecodeOptions::for_source(&malformed),
        )
        .expect_err("selected overlong values must fail before a write");
        assert_eq!(error.noncanonical_reason(), Some("protobuf varint value"));
        Ok(())
    }

    #[test]
    fn rewrite_required_dimensions_remain_canonical_and_wire_strict() {
        let write = TableHeaderSettingsWrite::default();
        let required_key = [0xb0, 0x80, 0x00, 1, 0x38, 1];
        assert_eq!(
            rewrite_table_header_settings_with_report(
                &required_key,
                write,
                DecodeOptions::for_source(&required_key),
            )
            .expect_err("required key")
            .noncanonical_reason(),
            Some("protobuf field key")
        );

        let required_value = [0x30, 0x81, 0x00, 0x38, 1];
        assert_eq!(
            rewrite_table_header_settings_with_report(
                &required_value,
                write,
                DecodeOptions::for_source(&required_value),
            )
            .expect_err("required value")
            .noncanonical_reason(),
            Some("protobuf varint value")
        );

        let required_wrong_wire = [0x32, 1, 1, 0x38, 1];
        assert!(
            rewrite_table_header_settings_with_report(
                &required_wrong_wire,
                write,
                DecodeOptions::for_source(&required_wrong_wire),
            )
            .is_err()
        );
    }

    #[test]
    fn rewrite_retains_noncanonical_unknown_framing_byte_for_byte() -> Result<(), DecodeError> {
        let mut source = dimensions();
        let unknown_varint = [0xa0, 0x06, 0x81, 0x00];
        let unknown_length = [0xb2, 0x06, 0x80, 0x00];
        source.extend(unknown_varint);
        source.extend(unknown_length);
        let write = TableHeaderSettingsWrite::new(Some(2), None, None, None, None, None, None);
        let options = DecodeOptions::for_source(&source).with_max_output_bytes(source.len() + 16);
        let (rewritten, report) =
            rewrite_table_header_settings_with_report(&source, write, options)?;
        assert!(report.changed());
        assert!(
            rewritten
                .windows(unknown_varint.len())
                .any(|window| window == unknown_varint)
        );
        assert!(
            rewritten
                .windows(unknown_length.len())
                .any(|window| window == unknown_length)
        );
        assert!(decode_table_header_settings(&source, options).is_err());
        Ok(())
    }

    #[test]
    fn rewrite_output_reservation_is_fallible_and_typed() {
        let error = reserve_output(usize::MAX).expect_err("oversized reservation");
        assert_eq!(error.allocation_amount(), Some(usize::MAX));
    }

    #[test]
    fn rewrite_report_replays_at_exact_limits_before_allocation() -> Result<(), DecodeError> {
        let mut source = dimensions();
        source.extend([0xa0, 0x06, 0x81, 0x00, 0xa3, 0x06, 0x08, 1, 0xa4, 0x06]);
        let write =
            TableHeaderSettingsWrite::new(Some(3), None, None, None, Some(true), None, None);
        let broad = DecodeOptions::for_source(&source).with_max_output_bytes(source.len() + 32);
        let (expected, report) = rewrite_table_header_settings_with_report(&source, write, broad)?;
        let exact = DecodeOptions::new(
            source.len(),
            report.fields(),
            report.work_bytes(),
            report.max_depth().max(1),
        )
        .with_max_output_bytes(report.output_bytes());
        let before = output_allocations();
        let (actual, exact_report) =
            rewrite_table_header_settings_with_report(&source, write, exact)?;
        assert_eq!(actual, expected);
        assert_eq!(exact_report.output_bytes(), report.output_bytes());
        assert_eq!(output_allocations(), before + 1);

        let before = output_allocations();
        let below_fields = DecodeOptions::new(
            source.len(),
            report.fields().saturating_sub(1),
            report.work_bytes(),
            report.max_depth().max(1),
        )
        .with_max_output_bytes(report.output_bytes());
        assert!(rewrite_table_header_settings_with_report(&source, write, below_fields).is_err());
        assert_eq!(output_allocations(), before);

        let before = output_allocations();
        let below_work = DecodeOptions::new(
            source.len(),
            report.fields(),
            report.work_bytes().saturating_sub(1),
            report.max_depth().max(1),
        )
        .with_max_output_bytes(report.output_bytes());
        assert!(rewrite_table_header_settings_with_report(&source, write, below_work).is_err());
        assert_eq!(output_allocations(), before);
        Ok(())
    }
}
