//! Character-style runs of a `TSWP.StorageArchive` and the native records a
//! span restyle writes (record 0309).
//!
//! A storage's `table_char_style` (field 8) is an object attribute table:
//! each entry names the UTF-16 index where a run starts and, optionally,
//! the `TSWP.CharacterStyleArchive` the run uses (no object means the
//! paragraph style's character properties apply). Restyling a span splits
//! the runs at its ends and gives each run inside it a new character-style
//! variation whose parent is the run's previous style, so every property
//! the span did not set is still inherited. The caller allocates the new
//! style objects, inserts them next to their stylesheet and registers them;
//! this module owns only the wire payloads, and keeps every field it does
//! not change byte for byte.
//!
//! Every operation requires a [`WireLimits`] policy. Input bytes, fields and
//! scan work are accumulated across the known nested messages it visits;
//! opaque fields stay opaque. Nesting starts at zero for the supplied payload.
//! Output bounds cover each encoded payload and any copied font name.

use std::ops::Range;

use litchi_iwa_common::{LimitKind, WireLimits, wire::WireView};
use thiserror::Error as ThisError;

/// `TSWP.StorageArchive.table_char_style`.
const CHAR_TABLE_FIELD: u32 = 8;
/// `TSWP.ObjectAttributeTable.entries`.
const ENTRY_FIELD: u32 = 1;
/// `ObjectAttribute.character_index`.
const INDEX_FIELD: u32 = 1;
/// `ObjectAttribute.object`.
const OBJECT_FIELD: u32 = 2;

/// A malformed input, invalid span or exhausted character-style resource.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum CharStyleError {
    /// The storage, table or stylesheet payload is malformed.
    #[error("malformed character-style payload")]
    Invalid,
    /// The span is empty or beyond the text.
    #[error("character span outside the text")]
    Span,
    /// A configured wire resource was exhausted.
    #[error("character-style {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        /// The exhausted resource.
        kind: LimitKind,
        /// The attempted aggregate consumption.
        observed: usize,
        /// The caller's configured bound.
        maximum: usize,
    },
    /// A bounded allocation failed.
    #[error("could not allocate {amount} character-style items or bytes")]
    Allocation {
        /// The attempted item or byte capacity.
        amount: usize,
    },
}

type Result<T> = std::result::Result<T, CharStyleError>;

fn from_wire(error: litchi_iwa_common::Error) -> CharStyleError {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => CharStyleError::Limit {
            kind,
            observed,
            maximum: limit,
        },
        litchi_iwa_common::Error::Allocation { amount, .. } => {
            CharStyleError::Allocation { amount }
        },
        _ => CharStyleError::Invalid,
    }
}

fn admit(kind: LimitKind, observed: usize, maximum: usize) -> Result<()> {
    if observed > maximum {
        return Err(CharStyleError::Limit {
            kind,
            observed,
            maximum,
        });
    }
    Ok(())
}

/// One operation's aggregate scan and work budget. Each nested view is
/// admitted before it can allocate spans, using the remaining field/work
/// allowance. Opaque length-delimited values are never guessed to be messages.
struct Budget {
    limits: WireLimits,
    bytes: usize,
    fields: usize,
    work: usize,
}

impl Budget {
    fn new(limits: WireLimits) -> Self {
        Self {
            limits,
            bytes: 0,
            fields: 0,
            work: 0,
        }
    }

    fn view<'a>(&mut self, data: &'a [u8], depth: usize) -> Result<WireView<'a>> {
        admit(LimitKind::Nesting, depth, self.limits.max_nesting())?;
        // WireLimits has private fields and finite hard ceilings below
        // usize::MAX, so saturation always fails admission. Counters used
        // below never exceed their bounds on a successful call.
        self.bytes = self.bytes.checked_add(data.len()).unwrap_or(usize::MAX);
        admit(
            LimitKind::InputBytes,
            self.bytes,
            self.limits.max_input_bytes(),
        )?;
        let fields_left = self.limits.max_fields() - self.fields;
        let work_left = self.limits.max_rewrite_work() - self.work;
        let work_limited = work_left < fields_left;
        let remaining = fields_left.min(work_left);
        if remaining == 0 && !data.is_empty() {
            return Err(self.field_failure(work_limited, 1));
        }
        // The empty message consumes no fields; WireLimits intentionally
        // rejects a zero configured bound, so use one only for that case.
        let policy = self
            .limits
            .with_fields(remaining.max(1))
            .map_err(from_wire)?;
        let view = WireView::parse(data, policy).map_err(|error| match error {
            litchi_iwa_common::Error::LimitExceeded {
                kind: LimitKind::Fields,
                observed,
                ..
            } => self.field_failure(work_limited, observed),
            other => from_wire(other),
        })?;
        self.fields += view.len();
        self.work += view.len();
        Ok(view)
    }

    fn field_failure(&self, work_limited: bool, additional: usize) -> CharStyleError {
        let (kind, prior, maximum) = if work_limited {
            (
                LimitKind::RewriteWork,
                self.work,
                self.limits.max_rewrite_work(),
            )
        } else {
            (LimitKind::Fields, self.fields, self.limits.max_fields())
        };
        CharStyleError::Limit {
            kind,
            observed: prior.saturating_add(additional),
            maximum,
        }
    }

    fn work(&mut self, additional: usize) -> Result<()> {
        self.work = self.work.checked_add(additional).unwrap_or(usize::MAX);
        admit(
            LimitKind::RewriteWork,
            self.work,
            self.limits.max_rewrite_work(),
        )
    }
}

fn reserve<T>(values: &mut Vec<T>, additional: usize) -> Result<()> {
    let amount = values.len().checked_add(additional).unwrap_or(usize::MAX);
    values
        .try_reserve(additional)
        .map_err(|_| CharStyleError::Allocation { amount })
}

fn reserve_output(output: &mut Vec<u8>, additional: usize, limits: WireLimits) -> Result<()> {
    let observed = output.len().checked_add(additional).unwrap_or(usize::MAX);
    admit(LimitKind::OutputBytes, observed, limits.max_output_bytes())?;
    reserve(output, additional)
}

fn output_buffer(length: usize, limits: WireLimits) -> Result<Vec<u8>> {
    admit(LimitKind::OutputBytes, length, limits.max_output_bytes())?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(length)
        .map_err(|_| CharStyleError::Allocation { amount: length })?;
    Ok(output)
}

/// One run of a storage's character-style table.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct CharStyleRun {
    /// The UTF-16 index where the run starts.
    pub start: usize,
    /// The run's `TSWP.CharacterStyleArchive`, if any.
    pub style: Option<u64>,
}

/// The character properties a span restyle sets. `None` leaves a property
/// inherited.
#[derive(Debug, Clone, Default, PartialEq)]
#[non_exhaustive]
pub struct CharacterProperties {
    /// Bold.
    pub bold: Option<bool>,
    /// Italic.
    pub italic: Option<bool>,
    /// A single underline (`true`) or none (`false`).
    pub underline: Option<bool>,
    /// A single strikethrough (`true`) or none (`false`).
    pub strikethrough: Option<bool>,
    /// The font size in points.
    pub size: Option<f32>,
    /// The font's PostScript or family name.
    pub font_name: Option<String>,
    /// An sRGB colour.
    pub color: Option<[u8; 3]>,
}

impl CharacterProperties {
    /// No property set.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Whether no property is set.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.override_count() == 0
    }

    /// Fill every property this value leaves unset from `inherited` (a
    /// parent style's properties).
    pub fn inherit(&mut self, inherited: &Self) {
        self.bold = self.bold.or(inherited.bold);
        self.italic = self.italic.or(inherited.italic);
        self.underline = self.underline.or(inherited.underline);
        self.strikethrough = self.strikethrough.or(inherited.strikethrough);
        self.size = self.size.or(inherited.size);
        if self.font_name.is_none() {
            self.font_name.clone_from(&inherited.font_name);
        }
        self.color = self.color.or(inherited.color);
    }

    fn override_count(&self) -> u64 {
        [
            self.bold.is_some(),
            self.italic.is_some(),
            self.underline.is_some(),
            self.strikethrough.is_some(),
            self.size.is_some(),
            self.font_name.is_some(),
            self.color.is_some(),
        ]
        .into_iter()
        .map(u64::from)
        .sum()
    }

    fn encoded_len(&self) -> usize {
        let flags = usize::from(self.bold.is_some())
            + usize::from(self.italic.is_some())
            + usize::from(self.underline.is_some())
            + usize::from(self.strikethrough.is_some());
        let fixed = flags * 2
            + usize::from(self.size.is_some()) * 5
            + usize::from(self.color.is_some()) * bytes_field_len(7, 22);
        fixed.saturating_add(
            self.font_name
                .as_ref()
                .map_or(0, |name| bytes_field_len(5, name.len())),
        )
    }

    /// Write the already-admitted properties payload in field order.
    fn encode_into(&self, output: &mut Vec<u8>, limits: WireLimits) -> Result<()> {
        if let Some(bold) = self.bold {
            varint_field(output, 1, u64::from(bold), limits)?;
        }
        if let Some(italic) = self.italic {
            varint_field(output, 2, u64::from(italic), limits)?;
        }
        if let Some(size) = self.size {
            fixed32_field(output, 3, size, limits)?;
        }
        if let Some(name) = &self.font_name {
            bytes_field(output, 5, name.as_bytes(), limits)?;
        }
        if let Some([red, green, blue]) = self.color {
            bytes_header(output, 7, 22, limits)?;
            varint_field(output, 1, 1, limits)?;
            fixed32_field(output, 3, f32::from(red) / 255.0, limits)?;
            fixed32_field(output, 4, f32::from(green) / 255.0, limits)?;
            fixed32_field(output, 5, f32::from(blue) / 255.0, limits)?;
            fixed32_field(output, 6, 1.0, limits)?;
        }
        if let Some(underline) = self.underline {
            varint_field(output, 11, u64::from(underline), limits)?;
        }
        if let Some(strikethrough) = self.strikethrough {
            varint_field(output, 12, u64::from(strikethrough), limits)?;
        }
        Ok(())
    }
}

fn push_varint(output: &mut Vec<u8>, mut value: u64) {
    loop {
        let byte = (value & 0x7f) as u8;
        value >>= 7;
        if value == 0 {
            output.push(byte);
            return;
        }
        output.push(byte | 0x80);
    }
}

fn varint_field(output: &mut Vec<u8>, number: u32, value: u64, limits: WireLimits) -> Result<()> {
    reserve_output(
        output,
        litchi_iwa_common::varint::encoded_len(u64::from(number) << 3)
            + litchi_iwa_common::varint::encoded_len(value),
        limits,
    )?;
    push_varint(output, u64::from(number) << 3);
    push_varint(output, value);
    Ok(())
}

fn fixed32_field(output: &mut Vec<u8>, number: u32, value: f32, limits: WireLimits) -> Result<()> {
    reserve_output(
        output,
        litchi_iwa_common::varint::encoded_len((u64::from(number) << 3) | 5) + 4,
        limits,
    )?;
    push_varint(output, (u64::from(number) << 3) | 5);
    output.extend_from_slice(&value.to_le_bytes());
    Ok(())
}

fn bytes_header(
    output: &mut Vec<u8>,
    number: u32,
    length: usize,
    limits: WireLimits,
) -> Result<()> {
    let additional = bytes_field_len(number, length);
    reserve_output(output, additional, limits)?;
    push_varint(output, (u64::from(number) << 3) | 2);
    push_varint(output, length as u64);
    Ok(())
}

fn bytes_field_len(number: u32, length: usize) -> usize {
    litchi_iwa_common::varint::encoded_len((u64::from(number) << 3) | 2)
        .checked_add(litchi_iwa_common::varint::encoded_len(length as u64))
        .and_then(|framing| framing.checked_add(length))
        .unwrap_or(usize::MAX)
}

fn bytes_field(
    output: &mut Vec<u8>,
    number: u32,
    payload: &[u8],
    limits: WireLimits,
) -> Result<()> {
    bytes_header(output, number, payload.len(), limits)?;
    output.extend_from_slice(payload);
    Ok(())
}

struct Reference {
    bytes: [u8; 11],
    length: usize,
}

impl std::ops::Deref for Reference {
    type Target = [u8];
    fn deref(&self) -> &Self::Target {
        &self.bytes[..self.length]
    }
}

fn reference(identifier: u64) -> Reference {
    let mut bytes = [0; 11];
    bytes[0] = 8;
    let mut scalar = [0; 10];
    let length = litchi_iwa_common::varint::encode_varint_to_buffer(identifier, &mut scalar).len();
    bytes[1..=length].copy_from_slice(&scalar[..length]);
    Reference {
        bytes,
        length: length + 1,
    }
}

fn varint(payload: &[u8]) -> Result<u64> {
    let (value, length) = litchi_iwa_common::varint::decode_varint_from_bytes(payload)
        .map_err(|_error| CharStyleError::Invalid)?;
    if length != payload.len() {
        return Err(CharStyleError::Invalid);
    }
    Ok(value)
}

fn reference_identifier(payload: &[u8], budget: &mut Budget, depth: usize) -> Result<u64> {
    let mut identifier = None;
    for field in budget.view(payload, depth)?.fields() {
        if field.number() != 1 || field.wire_type() != 0 {
            return Err(CharStyleError::Invalid);
        }
        if identifier.replace(varint(field.payload())?).is_some() {
            return Err(CharStyleError::Invalid);
        }
    }
    identifier.ok_or(CharStyleError::Invalid)
}

/// The runs of `storage`'s character-style table, in order. A table that
/// does not start at index 0 gets an implicit unstyled first run.
///
/// # Errors
///
/// [`CharStyleError::Invalid`] for a malformed or unordered table.
/// Resource and allocation failures retain their typed classification.
///
/// The caller must supply its resource policy.
///
/// ```compile_fail
/// let _ = litchi_iwa_text_wire::char_style_runs(&[]);
/// ```
pub fn char_style_runs(storage: &[u8], limits: WireLimits) -> Result<Vec<CharStyleRun>> {
    storage_runs(storage, &mut Budget::new(limits)).map(|(_, runs)| runs)
}

fn storage_runs<'a>(
    storage: &'a [u8],
    budget: &mut Budget,
) -> Result<(WireView<'a>, Vec<CharStyleRun>)> {
    let mut runs = Vec::new();
    let root = budget.view(storage, 0)?;
    for field in root.fields() {
        if field.number() != CHAR_TABLE_FIELD {
            continue;
        }
        if field.wire_type() != 2 || !runs.is_empty() {
            return Err(CharStyleError::Invalid);
        }
        for entry in budget.view(field.payload(), 1)?.fields() {
            if entry.number() != ENTRY_FIELD || entry.wire_type() != 2 {
                return Err(CharStyleError::Invalid);
            }
            let mut start = None;
            let mut style = None;
            for inner in budget.view(entry.payload(), 2)?.fields() {
                match (inner.number(), inner.wire_type()) {
                    (INDEX_FIELD, 0) => start = Some(varint(inner.payload())?),
                    (OBJECT_FIELD, 2) => {
                        style = Some(reference_identifier(inner.payload(), budget, 3)?)
                    },
                    _ => return Err(CharStyleError::Invalid),
                }
            }
            let start = usize::try_from(start.ok_or(CharStyleError::Invalid)?)
                .map_err(|_error| CharStyleError::Invalid)?;
            if runs
                .last()
                .is_some_and(|last: &CharStyleRun| last.start >= start)
            {
                return Err(CharStyleError::Invalid);
            }
            reserve(&mut runs, 1)?;
            runs.push(CharStyleRun { start, style });
        }
    }
    if runs.first().is_none_or(|first| first.start != 0) {
        admit(
            LimitKind::Fields,
            runs.len() + 1,
            budget.limits.max_fields(),
        )?;
        reserve(&mut runs, 1)?;
        runs.insert(
            0,
            CharStyleRun {
                start: 0,
                style: None,
            },
        );
    }
    Ok((root, runs))
}

/// The distinct styles of the runs that intersect `span` (UTF-16), in
/// first-use order: the caller creates one variation for each.
///
/// # Errors
///
/// [`CharStyleError::Span`] for an empty span or one past `text_len`;
/// [`CharStyleError::Invalid`] for a malformed table.
/// Resource and allocation failures retain their typed classification.
///
/// The caller must supply its resource policy.
///
/// ```compile_fail
/// let _ = litchi_iwa_text_wire::styles_in_span(&[], 1, 0..1);
/// ```
pub fn styles_in_span(
    storage: &[u8],
    text_len: usize,
    span: Range<usize>,
    limits: WireLimits,
) -> Result<Vec<Option<u64>>> {
    if span.start >= span.end || span.end > text_len {
        return Err(CharStyleError::Span);
    }
    let mut budget = Budget::new(limits);
    let (_, runs) = storage_runs(storage, &mut budget)?;
    let mut styles = Vec::new();
    for (index, run) in runs.iter().enumerate() {
        let end = runs.get(index + 1).map_or(text_len, |next| next.start);
        if run.start < span.end && end > span.start {
            budget.work(styles.len())?;
            if !styles.contains(&run.style) {
                reserve(&mut styles, 1)?;
                styles.push(run.style);
            }
        }
    }
    Ok(styles)
}

/// Rewrite `storage`'s character-style table so every run inside `span`
/// uses `restyled(previous style)`, splitting the runs at the span's ends.
///
/// Returns the new storage payload. Every other field keeps its bytes.
///
/// # Errors
///
/// As [`styles_in_span`].
///
/// The caller must supply its resource policy.
///
/// ```compile_fail
/// let _ = litchi_iwa_text_wire::restyle_span(&[], 1, 0..1, |_| None);
/// ```
pub fn restyle_span(
    storage: &[u8],
    text_len: usize,
    span: Range<usize>,
    restyled: impl Fn(Option<u64>) -> Option<u64>,
    limits: WireLimits,
) -> Result<Vec<u8>> {
    if span.start >= span.end || span.end > text_len {
        return Err(CharStyleError::Span);
    }
    let mut budget = Budget::new(limits);
    let (root, runs) = storage_runs(storage, &mut budget)?;
    let mut after: Vec<CharStyleRun> = Vec::new();
    let mut push = |mut run: CharStyleRun, transform: bool| -> Result<()> {
        budget.work(1)?;
        if transform {
            run.style = restyled(run.style);
        }
        if after.last().is_some_and(|last| last.start == run.start) {
            after.pop();
        }
        if after.last().is_some_and(|last| last.style == run.style) {
            return Ok(());
        }
        admit(LimitKind::Fields, after.len() + 1, limits.max_fields())?;
        reserve(&mut after, 1)?;
        after.push(run);
        Ok(())
    };
    for (index, run) in runs.iter().enumerate() {
        let end = runs.get(index + 1).map_or(text_len, |next| next.start);
        if end <= span.start || run.start >= span.end {
            push(*run, false)?;
            continue;
        }
        if run.start < span.start {
            push(*run, false)?;
        }
        push(
            CharStyleRun {
                start: run.start.max(span.start),
                style: run.style,
            },
            true,
        )?;
        if end > span.end {
            push(
                CharStyleRun {
                    start: span.end,
                    style: run.style,
                },
                false,
            )?;
        }
    }
    drop(runs);
    let entry_len = |run: &CharStyleRun| {
        1 + litchi_iwa_common::varint::encoded_len(run.start as u64)
            + run.style.map_or(0, |style| {
                bytes_field_len(OBJECT_FIELD, reference(style).len())
            })
    };
    let table_len = after.iter().fold(0usize, |total, run| {
        total.saturating_add(bytes_field_len(ENTRY_FIELD, entry_len(run)))
    });
    let mut table = output_buffer(table_len, limits)?;
    for run in after {
        bytes_header(&mut table, ENTRY_FIELD, entry_len(&run), limits)?;
        varint_field(&mut table, INDEX_FIELD, run.start as u64, limits)?;
        if let Some(style) = run.style {
            bytes_field(&mut table, OBJECT_FIELD, &reference(style), limits)?;
        }
    }
    let mut has_table = false;
    let mut output_len = root.fields().fold(0usize, |total, field| {
        total.saturating_add(if field.number() == CHAR_TABLE_FIELD {
            has_table = true;
            bytes_field_len(CHAR_TABLE_FIELD, table_len)
        } else {
            field.raw().len()
        })
    });
    if !has_table {
        output_len = output_len.saturating_add(bytes_field_len(CHAR_TABLE_FIELD, table_len));
    }
    let mut output = output_buffer(output_len, limits)?;
    let mut written = false;
    for field in root.fields() {
        if field.number() == CHAR_TABLE_FIELD {
            bytes_field(&mut output, CHAR_TABLE_FIELD, &table, limits)?;
            written = true;
        } else {
            reserve_output(&mut output, field.raw().len(), limits)?;
            output.extend_from_slice(field.raw());
        }
    }
    if !written {
        bytes_field(&mut output, CHAR_TABLE_FIELD, &table, limits)?;
    }
    Ok(output)
}

/// A `TSWP.CharacterStyleArchive` variation of `parent` (or a parentless
/// variation) in `stylesheet`, overriding `properties`.
///
/// # Errors
/// Returns a typed resource or allocation error before exceeding `limits`.
///
/// The caller must supply its resource policy.
///
/// ```compile_fail
/// let _ = litchi_iwa_text_wire::character_style_variation(None, 1, &litchi_iwa_text_wire::CharacterProperties::new());
/// ```
pub fn character_style_variation(
    parent: Option<u64>,
    stylesheet: u64,
    properties: &CharacterProperties,
    limits: WireLimits,
) -> Result<Vec<u8>> {
    let fields = 6
        + usize::from(parent.is_some()) * 2
        + properties.override_count() as usize
        + usize::from(properties.color.is_some()) * 5;
    admit(LimitKind::Fields, fields, limits.max_fields())?;
    admit(LimitKind::RewriteWork, fields, limits.max_rewrite_work())?;
    admit(LimitKind::Nesting, 2, limits.max_nesting())?;
    let parent = parent.map(reference);
    let stylesheet = reference(stylesheet);
    let style_len = parent
        .as_ref()
        .map_or(0, |parent| bytes_field_len(3, parent.len()))
        + 2
        + bytes_field_len(5, stylesheet.len());
    let properties_len = properties.encoded_len();
    let output_len = bytes_field_len(1, style_len)
        .saturating_add(2)
        .saturating_add(bytes_field_len(11, properties_len));
    let mut output = output_buffer(output_len, limits)?;
    bytes_header(&mut output, 1, style_len, limits)?;
    if let Some(parent) = parent {
        bytes_field(&mut output, 3, &parent, limits)?;
    }
    varint_field(&mut output, 4, 1, limits)?;
    bytes_field(&mut output, 5, &stylesheet, limits)?;
    varint_field(&mut output, 10, properties.override_count(), limits)?;
    bytes_header(&mut output, 11, properties_len, limits)?;
    properties.encode_into(&mut output, limits)?;
    debug_assert_eq!(output.len(), output_len);
    Ok(output)
}

/// The properties a `TSWP.CharacterStyleArchive` payload overrides itself
/// (its parent chain is not followed).
///
/// # Errors
///
/// [`CharStyleError::Invalid`] for a malformed payload.
/// Resource and allocation failures retain their typed classification.
///
/// The caller must supply its resource policy.
///
/// ```compile_fail
/// let _ = litchi_iwa_text_wire::character_style_properties(&[]);
/// ```
pub fn character_style_properties(style: &[u8], limits: WireLimits) -> Result<CharacterProperties> {
    let mut budget = Budget::new(limits);
    let mut properties = CharacterProperties::new();
    let Some(field) = budget
        .view(style, 0)?
        .fields()
        .find(|field| field.number() == 11)
    else {
        return Ok(properties);
    };
    if field.wire_type() != 2 {
        return Err(CharStyleError::Invalid);
    }
    let float = |payload: &[u8]| -> Result<f32> {
        let bytes: [u8; 4] = payload
            .try_into()
            .map_err(|_error| CharStyleError::Invalid)?;
        Ok(f32::from_le_bytes(bytes))
    };
    for inner in budget.view(field.payload(), 1)?.fields() {
        let payload = inner.payload();
        match (inner.number(), inner.wire_type()) {
            (1, 0) => properties.bold = Some(varint(payload)? != 0),
            (2, 0) => properties.italic = Some(varint(payload)? != 0),
            (3, 5) => properties.size = Some(float(payload)?),
            (5, 2) => {
                let name = std::str::from_utf8(payload).map_err(|_| CharStyleError::Invalid)?;
                admit(
                    LimitKind::OutputBytes,
                    name.len(),
                    limits.max_output_bytes(),
                )?;
                let mut owned = String::new();
                owned
                    .try_reserve_exact(name.len())
                    .map_err(|_| CharStyleError::Allocation { amount: name.len() })?;
                owned.push_str(name);
                properties.font_name = Some(owned);
            },
            (7, 2) => {
                let mut model = None;
                let mut channels = [None; 3];
                for channel in budget.view(payload, 2)?.fields() {
                    match (channel.number(), channel.wire_type()) {
                        (1, 0) => model = Some(varint(channel.payload())?),
                        (number @ 3..=5, 5) => {
                            let value = float(channel.payload())?.clamp(0.0, 1.0);
                            #[allow(
                                clippy::cast_possible_truncation,
                                clippy::cast_sign_loss,
                                reason = "a clamped unit channel scales into 0..=255"
                            )]
                            let byte = (value * 255.0).round() as u8;
                            channels[number as usize - 3] = Some(byte);
                        },
                        _ => {},
                    }
                }
                if let (Some(1), [Some(red), Some(green), Some(blue)]) = (model, channels) {
                    properties.color = Some([red, green, blue]);
                }
            },
            (11, 0) => properties.underline = Some(varint(payload)? != 0),
            (12, 0) => properties.strikethrough = Some(varint(payload)? != 0),
            _ => {},
        }
    }
    Ok(properties)
}

/// The parent of a `TSWP.CharacterStyleArchive` payload (`super.parent`).
///
/// # Errors
///
/// [`CharStyleError::Invalid`] for a malformed payload.
/// Resource and allocation failures retain their typed classification.
///
/// The caller must supply its resource policy.
///
/// ```compile_fail
/// let _ = litchi_iwa_text_wire::character_style_parent(&[]);
/// ```
pub fn character_style_parent(style: &[u8], limits: WireLimits) -> Result<Option<u64>> {
    let mut budget = Budget::new(limits);
    let Some(base) = budget
        .view(style, 0)?
        .fields()
        .find(|field| field.number() == 1)
    else {
        return Ok(None);
    };
    for field in budget.view(base.payload(), 1)?.fields() {
        if field.number() == 3 {
            return reference_identifier(field.payload(), &mut budget, 2).map(Some);
        }
    }
    Ok(None)
}

/// Register `style` in a `TSS.StylesheetArchive` payload: appended to its
/// styles (field 1) and, with a parent, to the parent's children entry
/// (field 5), creating the entry when the parent has none.
///
/// # Errors
///
/// [`CharStyleError::Invalid`] for a malformed stylesheet.
/// Resource and allocation failures retain their typed classification.
///
/// The caller must supply its resource policy.
///
/// ```compile_fail
/// let _ = litchi_iwa_text_wire::stylesheet_with_style(&[], 1, None);
/// ```
pub fn stylesheet_with_style(
    stylesheet: &[u8],
    style: u64,
    parent: Option<u64>,
    limits: WireLimits,
) -> Result<Vec<u8>> {
    let mut budget = Budget::new(limits);
    let fields = budget.view(stylesheet, 0)?;
    let style_reference = reference(style);
    let added_reference_len = bytes_field_len(1, style_reference.len());
    let mut last_style = None;
    let mut parent_entry = None;
    let mut output_len = added_reference_len;
    for (index, field) in fields.fields().enumerate() {
        if field.number() == 1 {
            last_style = Some(index);
        }
        let mut field_len = field.raw().len();
        if field.number() == 5 && parent.is_some() && parent_entry.is_none() {
            let mut entry_parent = None;
            for inner in budget.view(field.payload(), 1)?.fields() {
                if inner.number() == 1 {
                    entry_parent = Some(reference_identifier(inner.payload(), &mut budget, 2)?);
                }
            }
            if entry_parent == parent {
                let length = field.payload().len().saturating_add(added_reference_len);
                field_len = bytes_field_len(5, length);
                parent_entry = Some((index, length));
            }
        }
        output_len = output_len.saturating_add(field_len);
    }
    let new_entry = parent.filter(|_| parent_entry.is_none()).map(|parent| {
        let parent = reference(parent);
        let length = bytes_field_len(1, parent.len()) + added_reference_len;
        (parent, length)
    });
    if let Some((_, length)) = &new_entry {
        output_len = output_len.saturating_add(bytes_field_len(5, *length));
    }
    let mut output = output_buffer(output_len, limits)?;
    for (index, field) in fields.fields().enumerate() {
        if let Some((_, length)) = parent_entry.filter(|(selected, _)| *selected == index) {
            bytes_header(&mut output, 5, length, limits)?;
            output.extend_from_slice(field.payload());
            bytes_field(&mut output, 2, &style_reference, limits)?;
        } else {
            output.extend_from_slice(field.raw());
        }
        if Some(index) == last_style {
            bytes_field(&mut output, 1, &style_reference, limits)?;
        }
    }
    if last_style.is_none() {
        bytes_field(&mut output, 1, &style_reference, limits)?;
    }
    if let Some((parent, length)) = new_entry {
        bytes_header(&mut output, 5, length, limits)?;
        bytes_field(&mut output, 1, &parent, limits)?;
        bytes_field(&mut output, 2, &style_reference, limits)?;
    }
    debug_assert_eq!(output.len(), output_len);
    Ok(output)
}

#[cfg(test)]
mod tests {
    use super::{
        CharStyleError, CharStyleRun, CharacterProperties, char_style_runs, character_style_parent,
        character_style_properties, character_style_variation, reference, restyle_span,
        styles_in_span, stylesheet_with_style,
    };

    use litchi_iwa_common::{LimitKind, WireLimits};

    fn bytes_field(output: &mut Vec<u8>, number: u32, data: &[u8]) {
        super::bytes_field(output, number, data, WireLimits::DEFAULT).unwrap();
    }

    fn varint_field(output: &mut Vec<u8>, number: u32, value: u64) {
        super::varint_field(output, number, value, WireLimits::DEFAULT).unwrap();
    }

    fn storage(entries: &[(u64, Option<u64>)]) -> Vec<u8> {
        let mut output = Vec::new();
        bytes_field(&mut output, 3, b"hello world");
        let mut table = Vec::new();
        for (start, style) in entries {
            let mut entry = Vec::new();
            varint_field(&mut entry, 1, *start);
            if let Some(style) = style {
                bytes_field(&mut entry, 2, &reference(*style));
            }
            bytes_field(&mut table, 1, &entry);
        }
        bytes_field(&mut output, 8, &table);
        output
    }

    #[test]
    fn a_span_inside_one_run_splits_it() {
        let source = storage(&[(0, None)]);
        assert_eq!(
            styles_in_span(&source, 11, 2..5, WireLimits::DEFAULT).unwrap(),
            [None]
        );
        let restyled = restyle_span(&source, 11, 2..5, |_| Some(90), WireLimits::DEFAULT).unwrap();
        assert_eq!(
            char_style_runs(&restyled, WireLimits::DEFAULT).unwrap(),
            [
                CharStyleRun {
                    start: 0,
                    style: None
                },
                CharStyleRun {
                    start: 2,
                    style: Some(90)
                },
                CharStyleRun {
                    start: 5,
                    style: None
                },
            ]
        );
        assert_eq!(&restyled[..13], &source[..13]);
    }

    #[test]
    fn runs_inside_the_span_each_get_their_variation() {
        let source = storage(&[(0, None), (4, Some(7))]);
        assert_eq!(
            styles_in_span(&source, 11, 0..11, WireLimits::DEFAULT).unwrap(),
            [None, Some(7)]
        );
        let restyled = restyle_span(
            &source,
            11,
            0..11,
            |style| match style {
                None => Some(90),
                Some(_) => Some(91),
            },
            WireLimits::DEFAULT,
        )
        .unwrap();
        assert_eq!(
            char_style_runs(&restyled, WireLimits::DEFAULT).unwrap(),
            [
                CharStyleRun {
                    start: 0,
                    style: Some(90)
                },
                CharStyleRun {
                    start: 4,
                    style: Some(91)
                },
            ]
        );
        assert!(restyle_span(&source, 11, 3..3, |_| None, WireLimits::DEFAULT).is_err());
        assert!(restyle_span(&source, 11, 3..12, |_| None, WireLimits::DEFAULT).is_err());
    }

    #[test]
    fn variations_and_stylesheets_encode() {
        let properties = CharacterProperties {
            bold: Some(true),
            ..CharacterProperties::new()
        };
        let style =
            character_style_variation(Some(7), 3, &properties, WireLimits::DEFAULT).unwrap();
        assert_eq!(
            character_style_properties(&style, WireLimits::DEFAULT).unwrap(),
            properties
        );
        let full = CharacterProperties {
            bold: Some(false),
            italic: Some(true),
            underline: Some(true),
            strikethrough: Some(false),
            size: Some(18.5),
            font_name: Some("Helvetica".into()),
            color: Some([255, 0, 128]),
        };
        assert_eq!(
            character_style_properties(
                &character_style_variation(None, 3, &full, WireLimits::DEFAULT).unwrap(),
                WireLimits::DEFAULT
            )
            .unwrap(),
            full
        );
        assert_eq!(
            style,
            [
                0x0a, 0x0a, 0x1a, 0x02, 0x08, 0x07, 0x20, 0x01, 0x2a, 0x02, 0x08, 0x03, 0x50, 0x01,
                0x5a, 0x02, 0x08, 0x01
            ]
        );
        let mut sheet = Vec::new();
        bytes_field(&mut sheet, 1, &reference(7));
        let updated = stylesheet_with_style(&sheet, 9, Some(7), WireLimits::DEFAULT).unwrap();
        let mut expected = sheet.clone();
        bytes_field(&mut expected, 1, &reference(9));
        let mut entry = Vec::new();
        bytes_field(&mut entry, 1, &reference(7));
        bytes_field(&mut entry, 2, &reference(9));
        bytes_field(&mut expected, 5, &entry);
        assert_eq!(updated, expected);
    }

    #[test]
    fn nested_scans_share_byte_field_work_and_depth_bounds() {
        let source = storage(&[(0, Some(7))]);
        let exact = WireLimits::DEFAULT
            .with_input_bytes(39)
            .unwrap()
            .with_fields(6)
            .unwrap()
            .with_rewrite_work(6)
            .unwrap()
            .with_nesting(3)
            .unwrap();
        assert_eq!(
            char_style_runs(&source, exact).unwrap(),
            [CharStyleRun {
                start: 0,
                style: Some(7),
            }]
        );
        for (policy, kind, observed, maximum) in [
            (
                exact.with_input_bytes(38).unwrap(),
                LimitKind::InputBytes,
                39,
                38,
            ),
            (exact.with_fields(5).unwrap(), LimitKind::Fields, 6, 5),
            (
                exact.with_rewrite_work(5).unwrap(),
                LimitKind::RewriteWork,
                6,
                5,
            ),
            (exact.with_nesting(2).unwrap(), LimitKind::Nesting, 3, 2),
        ] {
            assert_eq!(
                char_style_runs(&source, policy),
                Err(CharStyleError::Limit {
                    kind,
                    observed,
                    maximum,
                })
            );
        }
    }

    #[test]
    fn span_reads_account_for_style_deduplication_work() {
        let source = storage(&[(0, Some(7)), (4, Some(8))]);
        // Ten parsed fields, then one comparison against the first style.
        let exact = WireLimits::DEFAULT.with_rewrite_work(11).unwrap();
        assert_eq!(
            styles_in_span(&source, 11, 0..11, exact).unwrap(),
            [Some(7), Some(8)]
        );
        assert_eq!(
            styles_in_span(&source, 11, 0..11, exact.with_rewrite_work(10).unwrap()),
            Err(CharStyleError::Limit {
                kind: LimitKind::RewriteWork,
                observed: 11,
                maximum: 10
            })
        );
    }

    #[test]
    fn restyle_checks_scan_before_callback_and_exact_output_bound() {
        let source = storage(&[(0, Some(7))]);
        let calls = std::cell::Cell::new(0);
        assert!(matches!(
            restyle_span(
                &source,
                11,
                0..11,
                |_| {
                    calls.set(calls.get() + 1);
                    None
                },
                WireLimits::DEFAULT.with_fields(5).unwrap()
            ),
            Err(CharStyleError::Limit { .. })
        ));
        assert_eq!(calls.get(), 0);
        assert!(matches!(
            restyle_span(
                &source,
                11,
                0..11,
                |_| {
                    calls.set(calls.get() + 1);
                    None
                },
                WireLimits::DEFAULT.with_rewrite_work(6).unwrap()
            ),
            Err(CharStyleError::Limit {
                kind: LimitKind::RewriteWork,
                observed: 7,
                maximum: 6
            })
        ));
        assert_eq!(calls.get(), 0);
        let output = restyle_span(
            &source,
            11,
            0..11,
            |_| None,
            WireLimits::DEFAULT.with_output_bytes(19).unwrap(),
        )
        .unwrap();
        assert_eq!(output, storage(&[(0, None)]));
        assert_eq!(
            restyle_span(
                &source,
                11,
                0..11,
                |_| None,
                WireLimits::DEFAULT.with_output_bytes(18).unwrap()
            ),
            Err(CharStyleError::Limit {
                kind: LimitKind::OutputBytes,
                observed: 19,
                maximum: 18
            })
        );
    }

    #[test]
    fn generated_variations_enforce_fields_work_depth_and_bytes() {
        let properties = CharacterProperties {
            bold: Some(true),
            ..CharacterProperties::new()
        };
        let exact = WireLimits::DEFAULT
            .with_fields(9)
            .unwrap()
            .with_rewrite_work(9)
            .unwrap()
            .with_nesting(2)
            .unwrap()
            .with_output_bytes(18)
            .unwrap();
        let result = character_style_variation(Some(7), 3, &properties, exact).unwrap();
        assert_eq!(result.len(), 18);
        for (policy, kind, observed, maximum) in [
            (exact.with_fields(8).unwrap(), LimitKind::Fields, 9, 8),
            (
                exact.with_rewrite_work(8).unwrap(),
                LimitKind::RewriteWork,
                9,
                8,
            ),
            (exact.with_nesting(1).unwrap(), LimitKind::Nesting, 2, 1),
            (
                exact.with_output_bytes(17).unwrap(),
                LimitKind::OutputBytes,
                18,
                17,
            ),
        ] {
            assert_eq!(
                character_style_variation(Some(7), 3, &properties, policy),
                Err(CharStyleError::Limit {
                    kind,
                    observed,
                    maximum
                })
            );
        }
    }

    #[test]
    fn properties_and_parent_keep_typed_failures() {
        let properties = CharacterProperties {
            font_name: Some("Abc".into()),
            ..CharacterProperties::new()
        };
        let style =
            character_style_variation(Some(7), 3, &properties, WireLimits::DEFAULT).unwrap();
        assert_eq!(
            character_style_properties(&style, WireLimits::DEFAULT.with_output_bytes(3).unwrap())
                .unwrap(),
            properties
        );
        assert_eq!(
            character_style_properties(&style, WireLimits::DEFAULT.with_output_bytes(2).unwrap()),
            Err(CharStyleError::Limit {
                kind: LimitKind::OutputBytes,
                observed: 3,
                maximum: 2
            })
        );
        assert_eq!(
            character_style_parent(&style, WireLimits::DEFAULT).unwrap(),
            Some(7)
        );
        assert_eq!(
            character_style_parent(&style, WireLimits::DEFAULT.with_nesting(1).unwrap()),
            Err(CharStyleError::Limit {
                kind: LimitKind::Nesting,
                observed: 2,
                maximum: 1
            })
        );
    }

    #[test]
    fn stylesheet_retains_opaque_bytes_and_checks_output_growth() {
        // Unknown field 20 with a noncanonical length prefix and opaque bytes.
        let opaque = [0xa2, 0x01, 0x82, 0x00, 0xff, 0xff];
        let mut source = opaque.to_vec();
        bytes_field(&mut source, 1, &reference(7));
        let mut entry = Vec::new();
        bytes_field(&mut entry, 1, &reference(7));
        bytes_field(&mut entry, 2, &reference(8));
        bytes_field(&mut source, 5, &entry);
        let output = stylesheet_with_style(&source, 9, Some(7), WireLimits::DEFAULT).unwrap();
        assert_eq!(&output[..opaque.len()], &opaque);
        assert_eq!(output.len(), source.len() + 8);
        assert_eq!(
            stylesheet_with_style(
                &source,
                9,
                Some(7),
                WireLimits::DEFAULT.with_output_bytes(output.len()).unwrap()
            )
            .unwrap(),
            output
        );
        assert!(matches!(
            stylesheet_with_style(
                &source,
                9,
                Some(7),
                WireLimits::DEFAULT
                    .with_output_bytes(output.len() - 1)
                    .unwrap()
            ),
            Err(CharStyleError::Limit {
                kind: LimitKind::OutputBytes,
                ..
            })
        ));
        let restyled = restyle_span(&opaque, 1, 0..1, |_| Some(7), WireLimits::DEFAULT).unwrap();
        assert_eq!(&restyled[..opaque.len()], &opaque);
    }

    #[test]
    fn malformed_input_stays_distinct_from_resource_exhaustion() {
        assert_eq!(
            char_style_runs(&[0x42, 0x02, 0x0a], WireLimits::DEFAULT),
            Err(CharStyleError::Invalid)
        );
        assert_eq!(
            character_style_properties(&[0x5a, 0x02, 0x2a], WireLimits::DEFAULT),
            Err(CharStyleError::Invalid)
        );
        assert_eq!(
            character_style_parent(&[0x0a, 0x03, 0x1a, 0x01, 0x80], WireLimits::DEFAULT),
            Err(CharStyleError::Invalid)
        );
        assert_eq!(
            stylesheet_with_style(&[0x2a, 0x01, 0x80], 9, Some(7), WireLimits::DEFAULT),
            Err(CharStyleError::Invalid)
        );
    }

    #[test]
    fn variation_output_plan_handles_every_property_and_varint_boundary() {
        for mask in 0u8..128 {
            let properties = CharacterProperties {
                bold: (mask & 1 != 0).then_some(true),
                italic: (mask & 2 != 0).then_some(false),
                underline: (mask & 4 != 0).then_some(true),
                strikethrough: (mask & 8 != 0).then_some(false),
                size: (mask & 16 != 0).then_some(18.5),
                font_name: (mask & 32 != 0).then(|| "a".repeat(128)),
                color: (mask & 64 != 0).then_some([24, 139, 220]),
            };
            for parent in [None, Some(u64::MAX)] {
                let value =
                    character_style_variation(parent, u64::MAX, &properties, WireLimits::DEFAULT)
                        .unwrap();
                let exact = WireLimits::DEFAULT.with_output_bytes(value.len()).unwrap();
                assert_eq!(
                    character_style_variation(parent, u64::MAX, &properties, exact).unwrap(),
                    value
                );
                assert_eq!(
                    character_style_properties(&value, WireLimits::DEFAULT).unwrap(),
                    properties
                );
                assert_eq!(
                    character_style_parent(&value, WireLimits::DEFAULT).unwrap(),
                    parent
                );
                assert_eq!(
                    character_style_variation(
                        parent,
                        u64::MAX,
                        &properties,
                        exact.with_output_bytes(value.len() - 1).unwrap()
                    ),
                    Err(CharStyleError::Limit {
                        kind: LimitKind::OutputBytes,
                        observed: value.len(),
                        maximum: value.len() - 1,
                    })
                );
            }
        }
    }

    #[test]
    fn stylesheet_plan_accounts_for_reframed_parent_entry() {
        for padding in [119, 120, 121, 122, 123, 124, 125, 126, 127, 128] {
            let mut entry = Vec::new();
            bytes_field(&mut entry, 1, &reference(7));
            bytes_field(&mut entry, 20, &vec![0xff; padding]);
            let mut source = Vec::new();
            bytes_field(&mut source, 5, &entry);
            let mut expected = Vec::new();
            bytes_field(&mut entry, 2, &reference(u64::MAX));
            bytes_field(&mut expected, 5, &entry);
            bytes_field(&mut expected, 1, &reference(u64::MAX));
            let exact = WireLimits::DEFAULT
                .with_output_bytes(expected.len())
                .unwrap();
            assert_eq!(
                stylesheet_with_style(&source, u64::MAX, Some(7), exact).unwrap(),
                expected
            );
        }
        let mut output = Vec::new();
        assert_eq!(
            super::bytes_header(&mut output, 1, usize::MAX, WireLimits::DEFAULT),
            Err(CharStyleError::Limit {
                kind: LimitKind::OutputBytes,
                observed: usize::MAX,
                maximum: WireLimits::DEFAULT.max_output_bytes()
            })
        );
        assert_eq!(output.capacity(), 0);
    }
}
