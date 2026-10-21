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

use std::ops::Range;

use litchi_iwa_common::wire::WireView;
use thiserror::Error as ThisError;

/// `TSWP.StorageArchive.table_char_style`.
const CHAR_TABLE_FIELD: u32 = 8;
/// `TSWP.ObjectAttributeTable.entries`.
const ENTRY_FIELD: u32 = 1;
/// `ObjectAttribute.character_index`.
const INDEX_FIELD: u32 = 1;
/// `ObjectAttribute.object`.
const OBJECT_FIELD: u32 = 2;

/// A malformed or out-of-range character-style input.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum CharStyleError {
    /// The storage, table or stylesheet payload is malformed.
    #[error("malformed character-style payload")]
    Invalid,
    /// The span is empty or beyond the text.
    #[error("character span outside the text")]
    Span,
}

type Result<T> = std::result::Result<T, CharStyleError>;

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

    /// The `TSWP.CharacterStylePropertiesArchive` payload, in field order.
    fn encode(&self) -> Vec<u8> {
        let mut output = Vec::new();
        if let Some(bold) = self.bold {
            varint_field(&mut output, 1, u64::from(bold));
        }
        if let Some(italic) = self.italic {
            varint_field(&mut output, 2, u64::from(italic));
        }
        if let Some(size) = self.size {
            fixed32_field(&mut output, 3, size);
        }
        if let Some(name) = &self.font_name {
            bytes_field(&mut output, 5, name.as_bytes());
        }
        if let Some([red, green, blue]) = self.color {
            let mut color = Vec::new();
            varint_field(&mut color, 1, 1);
            fixed32_field(&mut color, 3, f32::from(red) / 255.0);
            fixed32_field(&mut color, 4, f32::from(green) / 255.0);
            fixed32_field(&mut color, 5, f32::from(blue) / 255.0);
            fixed32_field(&mut color, 6, 1.0);
            bytes_field(&mut output, 7, &color);
        }
        if let Some(underline) = self.underline {
            varint_field(&mut output, 11, u64::from(underline));
        }
        if let Some(strikethrough) = self.strikethrough {
            varint_field(&mut output, 12, u64::from(strikethrough));
        }
        output
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

fn varint_field(output: &mut Vec<u8>, number: u32, value: u64) {
    push_varint(output, u64::from(number) << 3);
    push_varint(output, value);
}

fn fixed32_field(output: &mut Vec<u8>, number: u32, value: f32) {
    push_varint(output, (u64::from(number) << 3) | 5);
    output.extend_from_slice(&value.to_le_bytes());
}

fn bytes_field(output: &mut Vec<u8>, number: u32, payload: &[u8]) {
    push_varint(output, (u64::from(number) << 3) | 2);
    push_varint(output, payload.len() as u64);
    output.extend_from_slice(payload);
}

fn reference(identifier: u64) -> Vec<u8> {
    let mut output = Vec::new();
    varint_field(&mut output, 1, identifier);
    output
}

fn view(data: &[u8]) -> Result<WireView<'_>> {
    WireView::parse(data, litchi_iwa_common::WireLimits::DEFAULT)
        .map_err(|_error| CharStyleError::Invalid)
}

fn varint(payload: &[u8]) -> Result<u64> {
    let (value, length) = litchi_iwa_common::varint::decode_varint_from_bytes(payload)
        .map_err(|_error| CharStyleError::Invalid)?;
    if length != payload.len() {
        return Err(CharStyleError::Invalid);
    }
    Ok(value)
}

fn reference_identifier(payload: &[u8]) -> Result<u64> {
    let mut identifier = None;
    for field in view(payload)?.fields() {
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
pub fn char_style_runs(storage: &[u8]) -> Result<Vec<CharStyleRun>> {
    let mut runs = Vec::new();
    for field in view(storage)?.fields() {
        if field.number() != CHAR_TABLE_FIELD {
            continue;
        }
        if field.wire_type() != 2 || !runs.is_empty() {
            return Err(CharStyleError::Invalid);
        }
        for entry in view(field.payload())?.fields() {
            if entry.number() != ENTRY_FIELD || entry.wire_type() != 2 {
                return Err(CharStyleError::Invalid);
            }
            let mut start = None;
            let mut style = None;
            for inner in view(entry.payload())?.fields() {
                match (inner.number(), inner.wire_type()) {
                    (INDEX_FIELD, 0) => start = Some(varint(inner.payload())?),
                    (OBJECT_FIELD, 2) => style = Some(reference_identifier(inner.payload())?),
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
            runs.push(CharStyleRun { start, style });
        }
    }
    if runs.first().is_none_or(|first| first.start != 0) {
        runs.insert(
            0,
            CharStyleRun {
                start: 0,
                style: None,
            },
        );
    }
    Ok(runs)
}

/// The distinct styles of the runs that intersect `span` (UTF-16), in
/// first-use order: the caller creates one variation for each.
///
/// # Errors
///
/// [`CharStyleError::Span`] for an empty span or one past `text_len`;
/// [`CharStyleError::Invalid`] for a malformed table.
pub fn styles_in_span(
    storage: &[u8],
    text_len: usize,
    span: Range<usize>,
) -> Result<Vec<Option<u64>>> {
    if span.start >= span.end || span.end > text_len {
        return Err(CharStyleError::Span);
    }
    let runs = char_style_runs(storage)?;
    let mut styles = Vec::new();
    for (index, run) in runs.iter().enumerate() {
        let end = runs.get(index + 1).map_or(text_len, |next| next.start);
        if run.start < span.end && end > span.start && !styles.contains(&run.style) {
            styles.push(run.style);
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
pub fn restyle_span(
    storage: &[u8],
    text_len: usize,
    span: Range<usize>,
    restyled: impl Fn(Option<u64>) -> Option<u64>,
) -> Result<Vec<u8>> {
    if span.start >= span.end || span.end > text_len {
        return Err(CharStyleError::Span);
    }
    let runs = char_style_runs(storage)?;
    let mut after: Vec<CharStyleRun> = Vec::new();
    let mut push = |run: CharStyleRun| {
        if after.last().is_some_and(|last| last.start == run.start) {
            after.pop();
        }
        if after.last().is_some_and(|last| last.style == run.style) {
            return;
        }
        after.push(run);
    };
    for (index, run) in runs.iter().enumerate() {
        let end = runs.get(index + 1).map_or(text_len, |next| next.start);
        if end <= span.start || run.start >= span.end {
            push(*run);
            continue;
        }
        if run.start < span.start {
            push(*run);
        }
        push(CharStyleRun {
            start: run.start.max(span.start),
            style: restyled(run.style),
        });
        if end > span.end {
            push(CharStyleRun {
                start: span.end,
                style: run.style,
            });
        }
    }
    let mut table = Vec::new();
    for run in &after {
        let mut entry = Vec::new();
        varint_field(&mut entry, INDEX_FIELD, run.start as u64);
        if let Some(style) = run.style {
            bytes_field(&mut entry, OBJECT_FIELD, &reference(style));
        }
        bytes_field(&mut table, ENTRY_FIELD, &entry);
    }
    let mut output = Vec::with_capacity(storage.len() + table.len());
    let mut written = false;
    for field in view(storage)?.fields() {
        if field.number() == CHAR_TABLE_FIELD {
            bytes_field(&mut output, CHAR_TABLE_FIELD, &table);
            written = true;
        } else {
            output.extend_from_slice(field.raw());
        }
    }
    if !written {
        bytes_field(&mut output, CHAR_TABLE_FIELD, &table);
    }
    Ok(output)
}

/// A `TSWP.CharacterStyleArchive` variation of `parent` (or a parentless
/// variation) in `stylesheet`, overriding `properties`.
#[must_use]
pub fn character_style_variation(
    parent: Option<u64>,
    stylesheet: u64,
    properties: &CharacterProperties,
) -> Vec<u8> {
    let mut style = Vec::new();
    if let Some(parent) = parent {
        bytes_field(&mut style, 3, &reference(parent));
    }
    varint_field(&mut style, 4, 1);
    bytes_field(&mut style, 5, &reference(stylesheet));
    let mut output = Vec::new();
    bytes_field(&mut output, 1, &style);
    varint_field(&mut output, 10, properties.override_count());
    bytes_field(&mut output, 11, &properties.encode());
    output
}

/// The properties a `TSWP.CharacterStyleArchive` payload overrides itself
/// (its parent chain is not followed).
///
/// # Errors
///
/// [`CharStyleError::Invalid`] for a malformed payload.
pub fn character_style_properties(style: &[u8]) -> Result<CharacterProperties> {
    let mut properties = CharacterProperties::new();
    let Some(field) = view(style)?.fields().find(|field| field.number() == 11) else {
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
    for inner in view(field.payload())?.fields() {
        let payload = inner.payload();
        match (inner.number(), inner.wire_type()) {
            (1, 0) => properties.bold = Some(varint(payload)? != 0),
            (2, 0) => properties.italic = Some(varint(payload)? != 0),
            (3, 5) => properties.size = Some(float(payload)?),
            (5, 2) => {
                properties.font_name = Some(
                    String::from_utf8(payload.to_vec())
                        .map_err(|_error| CharStyleError::Invalid)?,
                );
            },
            (7, 2) => {
                let mut model = None;
                let mut channels = [None; 3];
                for channel in view(payload)?.fields() {
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
pub fn character_style_parent(style: &[u8]) -> Result<Option<u64>> {
    let Some(base) = view(style)?.fields().find(|field| field.number() == 1) else {
        return Ok(None);
    };
    for field in view(base.payload())?.fields() {
        if field.number() == 3 {
            return reference_identifier(field.payload()).map(Some);
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
pub fn stylesheet_with_style(
    stylesheet: &[u8],
    style: u64,
    parent: Option<u64>,
) -> Result<Vec<u8>> {
    let mut output = Vec::with_capacity(stylesheet.len() + 16);
    let mut appended_style = false;
    let mut parent_done = parent.is_none();
    let fields: Vec<_> = view(stylesheet)?.fields().collect();
    let last_style = fields.iter().rposition(|field| field.number() == 1);
    for (index, field) in fields.iter().enumerate() {
        if field.number() == 5 && !parent_done {
            let entry = field.payload();
            let mut entry_parent = None;
            for inner in view(entry)?.fields() {
                if inner.number() == 1 {
                    entry_parent = Some(reference_identifier(inner.payload())?);
                }
            }
            if entry_parent == parent {
                let mut rebuilt = entry.to_vec();
                bytes_field(&mut rebuilt, 2, &reference(style));
                bytes_field(&mut output, 5, &rebuilt);
                parent_done = true;
                continue;
            }
        }
        output.extend_from_slice(field.raw());
        if Some(index) == last_style {
            bytes_field(&mut output, 1, &reference(style));
            appended_style = true;
        }
    }
    if !appended_style {
        bytes_field(&mut output, 1, &reference(style));
    }
    if let (false, Some(parent)) = (parent_done, parent) {
        let mut entry = Vec::new();
        bytes_field(&mut entry, 1, &reference(parent));
        bytes_field(&mut entry, 2, &reference(style));
        bytes_field(&mut output, 5, &entry);
    }
    Ok(output)
}

#[cfg(test)]
mod tests {
    use super::{
        CharStyleRun, CharacterProperties, bytes_field, char_style_runs, character_style_variation,
        reference, restyle_span, styles_in_span, stylesheet_with_style, varint_field,
    };

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
        assert_eq!(styles_in_span(&source, 11, 2..5).unwrap(), [None]);
        let restyled = restyle_span(&source, 11, 2..5, |_| Some(90)).unwrap();
        assert_eq!(
            char_style_runs(&restyled).unwrap(),
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
        assert_eq!(styles_in_span(&source, 11, 0..11).unwrap(), [None, Some(7)]);
        let restyled = restyle_span(&source, 11, 0..11, |style| match style {
            None => Some(90),
            Some(_) => Some(91),
        })
        .unwrap();
        assert_eq!(
            char_style_runs(&restyled).unwrap(),
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
        assert!(restyle_span(&source, 11, 3..3, |_| None).is_err());
        assert!(restyle_span(&source, 11, 3..12, |_| None).is_err());
    }

    #[test]
    fn variations_and_stylesheets_encode() {
        let properties = CharacterProperties {
            bold: Some(true),
            ..CharacterProperties::new()
        };
        let style = character_style_variation(Some(7), 3, &properties);
        assert_eq!(
            super::character_style_properties(&style).unwrap(),
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
            super::character_style_properties(&character_style_variation(None, 3, &full)).unwrap(),
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
        let updated = stylesheet_with_style(&sheet, 9, Some(7)).unwrap();
        let mut expected = sheet.clone();
        bytes_field(&mut expected, 1, &reference(9));
        let mut entry = Vec::new();
        bytes_field(&mut entry, 1, &reference(7));
        bytes_field(&mut entry, 2, &reference(9));
        bytes_field(&mut expected, 5, &entry);
        assert_eq!(updated, expected);
    }
}
