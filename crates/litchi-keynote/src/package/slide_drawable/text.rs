//! Character runs of a drawable's text storage: the storage's
//! character-style table (`TSWP.StorageArchive` field 8) projected onto
//! character spans with the decodable `CharacterStylePropertiesArchive`
//! values of each run's style.

use litchi_iwa_common::WireLimits;

use super::super::structure::StructureError;
use super::super::structure::wire::{Fields, reference_identifier};
use super::inventory::{Objects, STORAGE, descend, unique_message};

const CHARACTER_STYLE: u32 = 2_021;

/// A half-open span of characters (Unicode scalar values) of a drawable's
/// text.
///
/// ```
/// use litchi_keynote::slide::drawable::TextSpan;
///
/// let span = TextSpan::new(2, 5);
/// assert_eq!((span.start(), span.end(), span.len()), (2, 5, 3));
/// ```
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct TextSpan {
    start: usize,
    end: usize,
}

impl TextSpan {
    /// The span `start..end`.
    #[must_use]
    pub const fn new(start: usize, end: usize) -> Self {
        Self { start, end }
    }

    /// The first character.
    #[must_use]
    pub const fn start(self) -> usize {
        self.start
    }

    /// One past the last character.
    #[must_use]
    pub const fn end(self) -> usize {
        self.end
    }

    /// The number of characters (0 for a reversed span).
    #[must_use]
    pub const fn len(self) -> usize {
        self.end.saturating_sub(self.start)
    }

    /// Whether the span covers no character.
    #[must_use]
    pub const fn is_empty(self) -> bool {
        self.len() == 0
    }
}

/// Character formatting; `None` leaves (or reports) the inherited value.
///
/// ```
/// use litchi_keynote::slide::drawable::Style;
///
/// let style = Style::new().with_bold(true).with_font_size(24.0).with_color([255, 0, 0]);
/// assert_eq!(style.bold, Some(true));
/// assert_eq!(style.color, Some([255, 0, 0]));
/// ```
#[non_exhaustive]
#[derive(Debug, Clone, Default, PartialEq)]
pub struct Style {
    /// Bold.
    pub bold: Option<bool>,
    /// Italic.
    pub italic: Option<bool>,
    /// Single underline.
    pub underline: Option<bool>,
    /// Font size in points.
    pub font_size: Option<f32>,
    /// PostScript font name.
    pub font_name: Option<String>,
    /// RGB color.
    pub color: Option<[u8; 3]>,
}

impl Style {
    /// A style that sets nothing.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            bold: None,
            italic: None,
            underline: None,
            font_size: None,
            font_name: None,
            color: None,
        }
    }

    /// This style with bold set.
    #[must_use]
    pub const fn with_bold(mut self, bold: bool) -> Self {
        self.bold = Some(bold);
        self
    }

    /// This style with italic set.
    #[must_use]
    pub const fn with_italic(mut self, italic: bool) -> Self {
        self.italic = Some(italic);
        self
    }

    /// This style with single underline set.
    #[must_use]
    pub const fn with_underline(mut self, underline: bool) -> Self {
        self.underline = Some(underline);
        self
    }

    /// This style with a font size in points.
    #[must_use]
    pub const fn with_font_size(mut self, points: f32) -> Self {
        self.font_size = Some(points);
        self
    }

    /// This style with a PostScript font name.
    #[must_use]
    pub fn with_font_name(mut self, name: impl Into<String>) -> Self {
        self.font_name = Some(name.into());
        self
    }

    /// This style with an RGB color.
    #[must_use]
    pub const fn with_color(mut self, rgb: [u8; 3]) -> Self {
        self.color = Some(rgb);
        self
    }
}

/// One character-style run of a drawable's text.
#[derive(Debug, Clone, PartialEq)]
pub struct TextRun {
    span: TextSpan,
    text: String,
    style: Style,
}

impl TextRun {
    /// The characters this run covers.
    #[must_use]
    pub const fn span(&self) -> TextSpan {
        self.span
    }

    /// The run's text.
    #[must_use]
    pub fn text(&self) -> &str {
        &self.text
    }

    /// The run's character-style values (empty where none is decodable).
    #[must_use]
    pub const fn style(&self) -> &Style {
        &self.style
    }
}

fn character_style(
    source: &impl Objects,
    identifier: u64,
    wire: WireLimits,
) -> Result<Style, StructureError> {
    let Some(object) = source.get(identifier) else {
        return Ok(Style::new());
    };
    let Ok((_, payload)) = unique_message(object, CHARACTER_STYLE) else {
        return Ok(Style::new());
    };
    let Some(properties) = descend(payload, &[11], wire)? else {
        return Ok(Style::new());
    };
    let fields = Fields::parse(&properties, wire)?;
    let flag = |number: u32| -> Result<Option<bool>, StructureError> {
        Ok(fields
            .unique(number)?
            .map(|field| field.varint())
            .transpose()?
            .map(|value| value != 0))
    };
    let color = match fields.unique(7)? {
        Some(field) => {
            let color = Fields::parse(field.message()?, wire)?;
            let channel = |number: u32| -> Result<Option<u8>, StructureError> {
                Ok(color
                    .unique(number)?
                    .map(|field| field.float())
                    .transpose()?
                    .map(|value| (value.clamp(0.0, 1.0) * 255.0).round() as u8))
            };
            let model = color.unique(1)?.map(|field| field.varint()).transpose()?;
            match (model, channel(3)?, channel(4)?, channel(5)?) {
                (Some(1), Some(r), Some(g), Some(b)) => Some([r, g, b]),
                _ => None,
            }
        },
        None => None,
    };
    Ok(Style {
        bold: flag(1)?,
        italic: flag(2)?,
        underline: fields
            .unique(11)?
            .map(|field| field.varint())
            .transpose()?
            .map(|value| value == 1),
        font_size: fields.unique(3)?.map(|field| field.float()).transpose()?,
        font_name: fields
            .unique(5)?
            .map(|field| {
                String::from_utf8(field.message()?.to_vec())
                    .map_err(|_| StructureError::Unexpected("font name UTF-8"))
            })
            .transpose()?,
        color,
    })
}

/// The longest parent chain a run's style is resolved through.
const MAX_STYLE_DEPTH: usize = 16;

/// The properties a character style and its parent chain set.
fn effective_style(
    source: &impl Objects,
    identifier: u64,
    wire: WireLimits,
) -> Result<Style, StructureError> {
    let mut style = Style::new();
    let mut current = Some(identifier);
    for _ in 0..MAX_STYLE_DEPTH {
        let Some(identifier) = current else { break };
        let own = character_style(source, identifier, wire)?;
        style.bold = style.bold.or(own.bold);
        style.italic = style.italic.or(own.italic);
        style.underline = style.underline.or(own.underline);
        style.font_size = style.font_size.or(own.font_size);
        if style.font_name.is_none() {
            style.font_name = own.font_name;
        }
        style.color = style.color.or(own.color);
        current = source
            .get(identifier)
            .and_then(|object| unique_message(object, CHARACTER_STYLE).ok())
            .map(|(_, payload)| litchi_iwa_text_wire::character_style_parent(payload, wire))
            .transpose()
            .map_err(StructureError::from)?
            .flatten();
    }
    Ok(style)
}

/// The runs of storage `storage` whose text is `text`.
pub(super) fn runs(
    source: &impl Objects,
    storage: u64,
    text: &str,
    wire: WireLimits,
) -> Result<Box<[TextRun]>, StructureError> {
    let object = source.get(storage).ok_or(StructureError::MissingObject)?;
    let (_, payload) = unique_message(object, STORAGE)?;
    let fields = Fields::parse(payload, wire)?;
    // Native character indexes count UTF-16 code units.
    let mut boundaries: Vec<(usize, Option<u64>)> = Vec::new();
    if let Some(table) = fields.unique(8)? {
        for entry in Fields::parse(table.message()?, wire)?.all(1) {
            let entry = Fields::parse(entry.message()?, wire)?;
            let index = entry
                .unique(1)?
                .map(|field| field.varint())
                .transpose()?
                .unwrap_or(0);
            let style = entry
                .unique(2)?
                .map(|field| reference_identifier(field.message()?, wire))
                .transpose()?;
            boundaries.push((usize::try_from(index).unwrap_or(usize::MAX), style));
        }
    }
    if boundaries.first().is_none_or(|(index, _)| *index != 0) {
        boundaries.insert(0, (0, None));
    }
    let utf16_len = text.encode_utf16().count();
    let mut char_starts = Vec::with_capacity(boundaries.len());
    for (index, _) in &boundaries {
        let index = (*index).min(utf16_len);
        let mut units = 0;
        let mut chars = 0;
        for character in text.chars() {
            if units >= index {
                break;
            }
            units += character.len_utf16();
            chars += 1;
        }
        char_starts.push(chars);
    }
    let total = text.chars().count();
    let mut output = Vec::with_capacity(boundaries.len());
    for (position, (_, style)) in boundaries.iter().enumerate() {
        let start = char_starts[position];
        let end = char_starts.get(position + 1).copied().unwrap_or(total);
        if end <= start {
            continue;
        }
        let style = match style {
            Some(identifier) => effective_style(source, *identifier, wire)?,
            None => Style::new(),
        };
        output.push(TextRun {
            span: TextSpan::new(start, end),
            text: text.chars().skip(start).take(end - start).collect(),
            style,
        });
    }
    Ok(output.into_boxed_slice())
}
