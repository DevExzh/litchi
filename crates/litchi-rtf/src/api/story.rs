//! Borrowed semantic text-story views.

#![allow(
    clippy::shadow_reuse,
    reason = "builder-style helpers deliberately rebind a working value as it is refined"
)]
use crate::types::{
    Alignment, Formatting as RawFormat, Paragraph as RawParagraph, StyleBlock, TextDirection,
    UnderlineStyle,
};
use std::fmt;
use std::iter::FusedIterator;
use std::num::NonZeroU16;

/// A structural break retained in a text story.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Break {
    /// End the current paragraph (`\\par`).
    Paragraph,
    /// Start a new visual line inside the current paragraph (`\\line`).
    Line,
}

/// Internal position of a structural break in flattened UTF-8 story text.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct Boundary {
    pub(crate) position: usize,
    pub(crate) kind: Break,
}

impl Boundary {
    pub(crate) const fn new(position: usize, kind: Break) -> Self {
        Self { position, kind }
    }
}

/// A lazily borrowed main-text story.
///
/// The view never flattens or clones the parser's retained text. Its iterators
/// borrow directly from the immutable document snapshot.
#[derive(Clone, Copy)]
pub struct Story<'a> {
    blocks: &'a [StyleBlock<'static>],
    boundaries: &'a [Boundary],
    text_len: usize,
    fonts: crate::font::Catalog<'a>,
    colors: crate::color::Palette<'a>,
}

impl<'a> Story<'a> {
    pub(crate) const fn new(
        blocks: &'a [StyleBlock<'static>],
        boundaries: &'a [Boundary],
        text_len: usize,
        fonts: crate::font::Catalog<'a>,
        colors: crate::color::Palette<'a>,
    ) -> Self {
        Self {
            blocks,
            boundaries,
            text_len,
            fonts,
            colors,
        }
    }

    /// Number of UTF-8 bytes in the visible retained story text.
    #[must_use]
    pub const fn len(self) -> usize {
        self.text_len
    }

    /// Whether this story contains no retained text bytes.
    #[must_use]
    pub const fn is_empty(self) -> bool {
        self.text_len == 0
    }

    /// Lazily traverse logical paragraphs, including explicitly empty ones.
    #[must_use]
    pub fn paragraphs(self) -> Paragraphs<'a> {
        Paragraphs::new(self)
    }

    /// Lazily traverse text runs and structural breaks across the whole story.
    #[must_use]
    pub fn inlines(self) -> Inlines<'a> {
        Inlines::new(self, 0, self.len(), 0, 0)
    }

    /// Lazily traverse only text runs, omitting structural break tokens.
    #[must_use]
    pub fn runs(self) -> Runs<'a> {
        Runs {
            inlines: self.inlines(),
        }
    }

    /// Borrow the font catalog used to resolve this story's runs.
    #[must_use]
    pub const fn fonts(self) -> crate::font::Catalog<'a> {
        self.fonts
    }

    /// Borrow the color palette used to resolve this story's runs.
    #[must_use]
    pub const fn colors(self) -> crate::color::Palette<'a> {
        self.colors
    }

    /// Write flattened plain text without allocating an intermediate string.
    ///
    /// # Errors
    /// Returns an error when writing to the output fails.
    pub fn write_text(self, output: &mut impl fmt::Write) -> fmt::Result {
        for block in self.blocks {
            output.write_str(block.text.as_ref())?;
        }
        Ok(())
    }

    /// Materialize flattened plain text in a new allocation.
    #[must_use]
    pub fn to_text(self) -> String {
        self.to_string()
    }
}

impl fmt::Debug for Story<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Story")
            .field("text_bytes", &self.len())
            .field("paragraphs", &self.paragraphs().count())
            .finish()
    }
}

impl fmt::Display for Story<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        self.write_text(formatter)
    }
}

/// One logical paragraph borrowed from a [`Story`].
#[derive(Clone, Copy)]
pub struct Paragraph<'a> {
    story: Story<'a>,
    start: usize,
    end: usize,
    start_block: usize,
    start_block_position: usize,
    format: &'a RawParagraph,
}

impl<'a> Paragraph<'a> {
    /// UTF-8 byte length, excluding the terminating paragraph break.
    #[must_use]
    pub const fn len(self) -> usize {
        self.end.saturating_sub(self.start)
    }

    /// Whether the paragraph has no retained text or inline line breaks.
    #[must_use]
    pub const fn is_empty(self) -> bool {
        self.start == self.end
    }

    /// Borrow the text when it occupies one retained parser block.
    ///
    /// Fragmented paragraphs return `None`; use [`Self::write_text`] or
    /// [`Self::to_text`] without guessing whether an allocation is required.
    #[must_use]
    pub fn as_str(self) -> Option<&'a str> {
        if self.is_empty() {
            return Some("");
        }
        let block = self.story.blocks.get(self.start_block)?;
        let local_start = self.start.checked_sub(self.start_block_position)?;
        let local_end = local_start.checked_add(self.len())?;
        block.text.get(local_start..local_end)
    }

    /// Lazily traverse semantic inline content in this paragraph.
    ///
    /// The terminating paragraph break is not part of the iterator. Explicit
    /// line breaks remain visible as [`Inline::Break`].
    #[must_use]
    pub fn inlines(self) -> Inlines<'a> {
        Inlines::new(
            self.story,
            self.start,
            self.end,
            self.start_block,
            self.start_block_position,
        )
    }

    /// Lazily traverse only text runs, omitting inline line-break tokens.
    #[must_use]
    pub fn runs(self) -> Runs<'a> {
        Runs {
            inlines: self.inlines(),
        }
    }

    /// Read the paragraph's local paragraph formatting.
    #[must_use]
    pub const fn format(self) -> ParagraphFormat<'a> {
        ParagraphFormat { raw: self.format }
    }

    /// Write plain paragraph text without allocating an intermediate string.
    ///
    /// # Errors
    /// Returns an error when writing to the output fails.
    pub fn write_text(self, output: &mut impl fmt::Write) -> fmt::Result {
        for inline in self.inlines() {
            match inline {
                Inline::Text(run) => output.write_str(run.text())?,
                Inline::Break(Break::Line) => output.write_char('\n')?,
                Inline::Break(Break::Paragraph) => {},
            }
        }
        Ok(())
    }

    /// Materialize this paragraph's plain text in a new allocation.
    #[must_use]
    pub fn to_text(self) -> String {
        self.to_string()
    }
}

impl fmt::Debug for Paragraph<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Paragraph")
            .field("text_bytes", &self.len())
            .field("format", &self.format())
            .finish()
    }
}

impl fmt::Display for Paragraph<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        self.write_text(formatter)
    }
}

/// Lazy paragraph traversal.
#[derive(Clone)]
pub struct Paragraphs<'a> {
    story: Story<'a>,
    boundary: usize,
    start: usize,
    start_block: usize,
    start_block_position: usize,
    text_len: usize,
    finished: bool,
}

impl<'a> Paragraphs<'a> {
    fn new(story: Story<'a>) -> Self {
        Self {
            story,
            boundary: 0,
            start: 0,
            start_block: 0,
            start_block_position: 0,
            text_len: story.len(),
            finished: false,
        }
    }

    fn make_paragraph(&mut self, end: usize, terminated: bool) -> Option<Paragraph<'a>> {
        let format_position = if terminated { end } else { end.checked_sub(1)? };
        let start_position = if self.start < end {
            self.start
        } else {
            format_position
        };
        let start_location = locate(
            self.story.blocks,
            self.start_block,
            self.start_block_position,
            start_position,
        )?;
        let format_location = locate(
            self.story.blocks,
            start_location.block,
            start_location.block_position,
            format_position,
        )?;
        let paragraph = Paragraph {
            story: self.story,
            start: self.start,
            end,
            start_block: start_location.block,
            start_block_position: start_location.block_position,
            format: &format_location.value.paragraph,
        };

        if terminated {
            self.start = end.saturating_add(1);
            self.start_block = format_location.block;
            self.start_block_position = format_location.block_position;
            if self.start >= format_location.block_end {
                self.start_block = self.start_block.saturating_add(1);
                self.start_block_position = format_location.block_end;
            }
        } else {
            self.finished = true;
        }
        Some(paragraph)
    }
}

impl<'a> Iterator for Paragraphs<'a> {
    type Item = Paragraph<'a>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.finished {
            return None;
        }
        while let Some(boundary) = self.story.boundaries.get(self.boundary).copied() {
            self.boundary = self.boundary.saturating_add(1);
            if boundary.kind == Break::Paragraph {
                return self.make_paragraph(boundary.position, true);
            }
        }
        if self.start < self.text_len {
            return self.make_paragraph(self.text_len, false);
        }
        self.finished = true;
        None
    }

    fn nth(&mut self, mut n: usize) -> Option<Self::Item> {
        if self.finished {
            return None;
        }
        while let Some(boundary) = self.story.boundaries.get(self.boundary).copied() {
            self.boundary = self.boundary.saturating_add(1);
            if boundary.kind != Break::Paragraph {
                continue;
            }
            if n == 0 {
                return self.make_paragraph(boundary.position, true);
            }
            n = n.saturating_sub(1);
            self.start = boundary.position.saturating_add(1);
        }
        if self.start < self.text_len && n == 0 {
            return self.make_paragraph(self.text_len, false);
        }
        self.finished = true;
        None
    }

    fn size_hint(&self) -> (usize, Option<usize>) {
        if self.finished {
            return (0, Some(0));
        }
        let boundaries = self
            .story
            .boundaries
            .get(self.boundary..)
            .map_or(0, |values| {
                values
                    .iter()
                    .filter(|boundary| boundary.kind == Break::Paragraph)
                    .count()
            });
        let trailing = usize::from(self.start < self.text_len);
        (boundaries, boundaries.checked_add(trailing))
    }
}

impl FusedIterator for Paragraphs<'_> {}

/// One borrowed semantic inline item.
#[derive(Debug, Clone, Copy)]
#[non_exhaustive]
pub enum Inline<'a> {
    /// A contiguous text run with one local character format.
    Text(Run<'a>),
    /// A paragraph or line boundary.
    Break(Break),
}

/// Lazy inline traversal over a story range.
#[derive(Clone)]
pub struct Inlines<'a> {
    story: Story<'a>,
    block: usize,
    block_position: usize,
    local: usize,
    position: usize,
    end: usize,
    boundary: usize,
}

impl<'a> Inlines<'a> {
    fn new(
        story: Story<'a>,
        start: usize,
        end: usize,
        start_block: usize,
        start_block_position: usize,
    ) -> Self {
        Self {
            story,
            block: start_block,
            block_position: start_block_position,
            local: start.saturating_sub(start_block_position),
            position: start,
            end,
            boundary: story
                .boundaries
                .partition_point(|boundary| boundary.position < start),
        }
    }
}

impl<'a> Iterator for Inlines<'a> {
    type Item = Inline<'a>;

    fn next(&mut self) -> Option<Self::Item> {
        while self.position < self.end {
            let boundary = self.story.boundaries.get(self.boundary).copied();
            if boundary.is_some_and(|boundary| boundary.position < self.position) {
                self.boundary = self.boundary.saturating_add(1);
                continue;
            }

            let block = self.story.blocks.get(self.block)?;
            let block_end = self.block_position.checked_add(block.text.len())?;
            if self.position >= block_end {
                self.block = self.block.saturating_add(1);
                self.block_position = block_end;
                self.local = 0;
                continue;
            }

            if let Some(boundary) = boundary.filter(|boundary| {
                boundary.position == self.position && boundary.position < self.end
            }) {
                let remainder = block.text.get(self.local..)?;
                if !remainder.starts_with('\n') {
                    return None;
                }
                self.local = self.local.saturating_add(1);
                self.position = self.position.saturating_add(1);
                self.boundary = self.boundary.saturating_add(1);
                return Some(Inline::Break(boundary.kind));
            }

            let boundary_end = boundary
                .filter(|boundary| boundary.position < self.end)
                .map_or(self.end, |boundary| boundary.position);
            let fragment_end = block_end.min(self.end).min(boundary_end);
            let fragment_len = fragment_end.checked_sub(self.position)?;
            if fragment_len == 0 {
                return None;
            }
            let local_end = self.local.checked_add(fragment_len)?;
            let text = block.text.get(self.local..local_end)?;
            self.local = local_end;
            self.position = fragment_end;
            return Some(Inline::Text(Run {
                text,
                format: &block.formatting,
                fonts: self.story.fonts,
                colors: self.story.colors,
            }));
        }
        None
    }

    fn size_hint(&self) -> (usize, Option<usize>) {
        (0, None)
    }
}

impl FusedIterator for Inlines<'_> {}

/// Lazy text-run traversal that omits structural break tokens.
#[derive(Clone)]
pub struct Runs<'a> {
    inlines: Inlines<'a>,
}

impl<'a> Iterator for Runs<'a> {
    type Item = Run<'a>;

    fn next(&mut self) -> Option<Self::Item> {
        loop {
            match self.inlines.next()? {
                Inline::Text(run) => return Some(run),
                Inline::Break(_) => {},
            }
        }
    }

    fn size_hint(&self) -> (usize, Option<usize>) {
        (0, self.inlines.size_hint().1)
    }
}

impl FusedIterator for Runs<'_> {}

/// A contiguous borrowed text run.
#[derive(Debug, Clone, Copy)]
pub struct Run<'a> {
    text: &'a str,
    format: &'a RawFormat,
    fonts: crate::font::Catalog<'a>,
    colors: crate::color::Palette<'a>,
}

impl<'a> Run<'a> {
    /// Borrow this run's UTF-8 text.
    #[must_use]
    pub const fn text(self) -> &'a str {
        self.text
    }

    /// Read this run's local character formatting.
    #[must_use]
    pub const fn format(self) -> Format<'a> {
        Format {
            raw: self.format,
            fonts: self.fonts,
            colors: self.colors,
        }
    }
}

impl fmt::Display for Run<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(self.text)
    }
}

/// Read-only local character formatting.
#[derive(Debug, Clone, Copy)]
pub struct Format<'a> {
    raw: &'a RawFormat,
    fonts: crate::font::Catalog<'a>,
    colors: crate::color::Palette<'a>,
}

impl<'a> Format<'a> {
    pub(crate) const fn raw(self) -> &'a RawFormat {
        self.raw
    }

    /// Whether bold formatting is active.
    #[must_use]
    pub const fn bold(self) -> bool {
        self.raw.bold
    }

    /// Whether italic formatting is active.
    #[must_use]
    pub const fn italic(self) -> bool {
        self.raw.italic
    }

    /// The exact local underline style.
    #[must_use]
    pub const fn underline(self) -> UnderlineStyle {
        self.raw.underline
    }

    /// Whether single-strike formatting is active.
    #[must_use]
    pub const fn strike(self) -> bool {
        self.raw.strike
    }

    /// Whether double-strike formatting is active.
    #[must_use]
    pub const fn double_strike(self) -> bool {
        self.raw.double_strike
    }

    /// Font size in half-points.
    #[must_use]
    pub const fn size(self) -> NonZeroU16 {
        self.raw.font_size
    }

    /// The effective typed character baseline.
    #[must_use]
    pub const fn baseline(self) -> crate::CharacterBaseline {
        match self.raw.character_positioning.baseline {
            crate::CharacterBaseline::Normal if self.raw.superscript => {
                crate::CharacterBaseline::Superscript
            },
            crate::CharacterBaseline::Normal if self.raw.subscript => {
                crate::CharacterBaseline::Subscript
            },
            baseline => baseline,
        }
    }

    /// Resolve the run's font definition without exposing its numeric RTF ID.
    ///
    /// `None` means that the retained reference has no corresponding font-table
    /// definition.
    #[must_use]
    pub fn font(self) -> Option<crate::font::Font<'a>> {
        self.fonts.resolve(self.raw.font_ref)
    }

    /// Resolve the local foreground selection, including automatic color.
    #[must_use]
    pub fn foreground(self) -> Option<crate::color::Value> {
        self.colors.resolve(self.raw.color_ref)
    }

    /// Resolve the explicitly authored character background selection.
    #[must_use]
    pub fn background(self) -> Option<crate::color::Value> {
        self.raw
            .background_color
            .and_then(|reference| self.colors.resolve(reference))
    }

    /// Resolve the explicitly authored highlight selection.
    #[must_use]
    pub fn highlight(self) -> Option<crate::color::Value> {
        self.raw
            .highlight_color
            .and_then(|reference| self.colors.resolve(reference))
    }

    /// Resolve the explicitly authored underline color selection.
    #[must_use]
    pub fn underline_paint(self) -> Option<crate::color::Value> {
        self.raw
            .underline_color
            .and_then(|reference| self.colors.resolve(reference))
    }

    /// Resolve the explicit foreground RGB value.
    ///
    /// Automatic color and an unresolved reference both return `None`; use
    /// [`Self::foreground`] when that distinction matters.
    pub fn foreground_color(self) -> Option<crate::color::Color> {
        self.foreground().and_then(crate::color::Value::color)
    }

    /// Resolve the explicitly authored background RGB value.
    ///
    /// Automatic color and an unresolved reference both return `None`; use
    /// [`Self::background`] when that distinction matters.
    pub fn background_color(self) -> Option<crate::color::Color> {
        self.background().and_then(crate::color::Value::color)
    }

    /// Resolve the explicitly authored highlight RGB value.
    ///
    /// Automatic color and an unresolved reference both return `None`; use
    /// [`Self::highlight`] when that distinction matters.
    pub fn highlight_color(self) -> Option<crate::color::Color> {
        self.highlight().and_then(crate::color::Value::color)
    }

    /// Resolve the explicitly authored underline RGB value.
    ///
    /// Automatic color, an unresolved reference, and the RTF default where the
    /// underline uses the foreground color all return `None`; use
    /// [`Self::underline_paint`] when the authored distinction matters.
    pub fn underline_color(self) -> Option<crate::color::Color> {
        self.underline_paint().and_then(crate::color::Value::color)
    }

    /// Whether the run is hidden text.
    #[must_use]
    pub const fn hidden(self) -> bool {
        self.raw.hidden
    }

    /// Whether the run uses small capitals.
    #[must_use]
    pub const fn small_caps(self) -> bool {
        self.raw.smallcaps
    }

    /// Whether the run uses all capitals.
    #[must_use]
    pub const fn all_caps(self) -> bool {
        self.raw.all_caps
    }

    /// Whether outline formatting is active.
    #[must_use]
    pub const fn outline(self) -> bool {
        self.raw.outline
    }

    /// Explicit local text direction, if authored.
    #[must_use]
    pub const fn direction(self) -> Option<TextDirection> {
        self.raw.direction
    }
}

/// Read-only local paragraph formatting.
#[derive(Debug, Clone, Copy)]
pub struct ParagraphFormat<'a> {
    raw: &'a RawParagraph,
}

impl<'a> ParagraphFormat<'a> {
    pub(crate) const fn raw(self) -> &'a RawParagraph {
        self.raw
    }

    /// Dependency-free local layout values used by paragraph transactions.
    #[must_use]
    pub fn layout(self) -> crate::edit::ParagraphLayout {
        crate::edit::ParagraphLayout::from_raw(self.raw)
    }

    /// Local paragraph alignment.
    #[must_use]
    pub const fn alignment(self) -> Alignment {
        self.raw.alignment
    }

    /// Explicit local paragraph spacing in twips.
    #[must_use]
    pub const fn spacing(self) -> crate::Spacing {
        self.raw.spacing
    }

    /// Explicit local physical paragraph indentation in twips.
    #[must_use]
    pub const fn indentation(self) -> crate::Indentation {
        self.raw.indentation
    }

    /// Explicit local text direction, if authored.
    #[must_use]
    pub const fn direction(self) -> Option<TextDirection> {
        self.raw.direction
    }

    /// Whether the paragraph requests staying on one page.
    #[must_use]
    pub const fn keep_together(self) -> bool {
        self.raw.keep_together
    }

    /// Whether the paragraph requests staying with its successor.
    #[must_use]
    pub const fn keep_with_next(self) -> bool {
        self.raw.keep_next
    }

    /// Whether a page break is requested before this paragraph.
    #[must_use]
    pub const fn page_break_before(self) -> bool {
        self.raw.page_break_before
    }

    /// Optional zero-based outline level.
    #[must_use]
    pub const fn outline_level(self) -> Option<u8> {
        self.raw.outline_level
    }
}

#[derive(Clone, Copy)]
struct Location<'a, 'text> {
    block: usize,
    block_position: usize,
    block_end: usize,
    value: &'a StyleBlock<'text>,
}

fn locate<'a, 'text>(
    blocks: &'a [StyleBlock<'text>],
    mut block: usize,
    mut block_position: usize,
    position: usize,
) -> Option<Location<'a, 'text>> {
    loop {
        let value = blocks.get(block)?;
        let block_end = block_position.checked_add(value.text.len())?;
        if position < block_end {
            return Some(Location {
                block,
                block_position,
                block_end,
                value,
            });
        }
        block = block.checked_add(1)?;
        block_position = block_end;
    }
}

pub(crate) fn validate_boundaries(
    blocks: &[StyleBlock<'_>],
    boundaries: &[Boundary],
) -> crate::Result<()> {
    let mut block = 0usize;
    let mut block_position = 0usize;
    let mut previous = None;
    for boundary in boundaries {
        if previous.is_some_and(|position| position >= boundary.position) {
            return Err(crate::Error::MalformedDocument(
                "RTF text-story boundaries are not strictly ordered".to_string(),
            ));
        }
        let location =
            locate(blocks, block, block_position, boundary.position).ok_or_else(|| {
                crate::Error::MalformedDocument(
                    "RTF text-story boundary is outside the body text".to_string(),
                )
            })?;
        let local = boundary
            .position
            .checked_sub(location.block_position)
            .ok_or_else(|| {
                crate::Error::MalformedDocument(
                    "RTF text-story boundary precedes its text block".to_string(),
                )
            })?;
        if location.value.text.as_bytes().get(local) != Some(&b'\n') {
            return Err(crate::Error::MalformedDocument(
                "RTF text-story boundary does not reference a line-feed byte".to_string(),
            ));
        }
        block = location.block;
        block_position = location.block_position;
        previous = Some(boundary.position);
    }
    Ok(())
}
