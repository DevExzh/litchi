//! Rich formatting of existing `DrawingML` text bodies and table cells.
//!
//! The edits address a shape by its positional scene index and descend into
//! its text body (`p:txBody`) or, for a table frame, into one `a:tc` cell.
//! Paragraph and run positions count the direct `a:p` children of the body
//! and the `a:r`/`a:br`/`a:fld` children of a paragraph, in document order,
//! exactly as the text-body reader numbers them.
//!
//! A shape target can also replace its fill and preset geometry, which live
//! in the shape's `p:spPr` rather than its text body.
//!
//! Every rewrite is a byte splice: untouched siblings, attributes, extension
//! lists and unknown children keep their exact source bytes. Inserted
//! children follow the schema order of `a:rPr`, `a:pPr`, `a:tcPr` and
//! `p:spPr` and use the `DrawingML` prefix of the element that receives them.

use std::ops::Range;

use litchi_core::Position;
use litchi_drawingml::geom::Preset;
use litchi_drawingml::text::{Alignment, Anchor, Strike, Underline};
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
use quick_xml::Reader;
use quick_xml::events::Event;

use crate::{Error, Result};

/// The largest accepted font size, in hundredths of a point (`ST_TextFontSize`).
pub const MAX_FONT_SIZE: u32 = 400_000;
/// The smallest accepted font size, in hundredths of a point (`ST_TextFontSize`).
pub const MIN_FONT_SIZE: u32 = 100;
/// The largest accepted typeface name, in bytes.
pub const MAX_TYPEFACE_BYTES: usize = 1_024;

/// The text element a [`TextAssignment`] changes.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum TextTarget {
    /// Every paragraph and run of a shape's text body.
    Shape {
        /// Positional scene index of the shape.
        shape: Position,
    },
    /// One paragraph of a shape's text body.
    Paragraph {
        /// Positional scene index of the shape.
        shape: Position,
        /// Position among the body's `a:p` children.
        paragraph: Position,
    },
    /// One run (`a:r`, `a:br` or `a:fld`) of a paragraph.
    Run {
        /// Positional scene index of the shape.
        shape: Position,
        /// Position among the body's `a:p` children.
        paragraph: Position,
        /// Position among the paragraph's run children.
        run: Position,
    },
    /// Bytes `start..end` of the text of one text run (`a:r`): the run is
    /// split so that the span is a run of its own, each part carrying a copy
    /// of the run's `a:rPr`. The edit's text replaces the span and its run
    /// properties apply to that run only (record 0273). The span is
    /// non-empty and lies on character boundaries of the run's text.
    RunSpan {
        /// Positional scene index of the shape.
        shape: Position,
        /// Position among the body's `a:p` children.
        paragraph: Position,
        /// Position among the paragraph's run children.
        run: Position,
        /// First byte of the span in the run's text.
        start: usize,
        /// End byte of the span in the run's text.
        end: usize,
    },
    /// One cell of a table graphic frame.
    Cell {
        /// Positional scene index of the table frame.
        shape: Position,
        /// Position among the table's `a:tr` rows.
        row: Position,
        /// Position among the row's `a:tc` cells.
        column: Position,
    },
    /// One paragraph of a table cell's text body, as [`Self::Paragraph`]
    /// is one of a shape's (record 0333).
    CellParagraph {
        /// Positional scene index of the table frame.
        shape: Position,
        /// Position among the table's `a:tr` rows.
        row: Position,
        /// Position among the row's `a:tc` cells.
        column: Position,
        /// Position among the cell body's `a:p` children.
        paragraph: Position,
    },
    /// One run (`a:r`, `a:br` or `a:fld`) of a paragraph of a table cell's
    /// text body, as [`Self::Run`] is one of a shape's (record 0285).
    CellRun {
        /// Positional scene index of the table frame.
        shape: Position,
        /// Position among the table's `a:tr` rows.
        row: Position,
        /// Position among the row's `a:tc` cells.
        column: Position,
        /// Position among the cell body's `a:p` children.
        paragraph: Position,
        /// Position among the paragraph's run children.
        run: Position,
    },
    /// Bytes `start..end` of the text of one text run of a table cell,
    /// split out as [`Self::RunSpan`] splits a shape's run (record 0285).
    CellRunSpan {
        /// Positional scene index of the table frame.
        shape: Position,
        /// Position among the table's `a:tr` rows.
        row: Position,
        /// Position among the row's `a:tc` cells.
        column: Position,
        /// Position among the cell body's `a:p` children.
        paragraph: Position,
        /// Position among the paragraph's run children.
        run: Position,
        /// First byte of the span in the run's text.
        start: usize,
        /// End byte of the span in the run's text.
        end: usize,
    },
}

/// An RGB colour written as `a:srgbClr@val`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Rgb(pub [u8; 3]);

/// A table cell's local fill, the fill choice of its `a:tcPr` (record
/// 0333).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum CellFill {
    /// A solid RGB colour (`a:solidFill/a:srgbClr`).
    Rgb(Rgb),
    /// A solid theme colour (`a:solidFill/a:schemeClr`).
    Scheme(litchi_drawingml::color::Scheme),
    /// No fill (`a:noFill`).
    None,
}

/// One cell inset of `a:tcPr` (record 0333).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum CellMargin {
    /// `marL`.
    Left,
    /// `marR`.
    Right,
    /// `marT`.
    Top,
    /// `marB`.
    Bottom,
}

impl CellMargin {
    /// Every inset, in `marL`, `marR`, `marT`, `marB` order.
    pub const ALL: [Self; 4] = [Self::Left, Self::Right, Self::Top, Self::Bottom];

    const fn attribute(self) -> &'static [u8] {
        match self {
            Self::Left => b"marL",
            Self::Right => b"marR",
            Self::Top => b"marT",
            Self::Bottom => b"marB",
        }
    }
}

/// A shape's own fill, the fill choice of its `p:spPr`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum ShapeFill {
    /// A solid colour (`a:solidFill/a:srgbClr`).
    Solid(Rgb),
    /// No fill (`a:noFill`).
    None,
}

/// Optional property assignments for one [`TextAssignment`].
///
/// Each field is `None` to leave the property as it is, `Some(None)` to
/// remove the local property, and `Some(Some(value))` to set it.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
#[allow(
    clippy::option_option,
    reason = "omission, removal and replacement are distinct assignments"
)]
pub struct TextProperties<'a> {
    text: Option<&'a str>,
    bold: Option<Option<bool>>,
    italic: Option<Option<bool>>,
    size: Option<Option<u32>>,
    color: Option<Option<Rgb>>,
    font: Option<Option<&'a str>>,
    latin: Option<Option<&'a str>>,
    align: Option<Option<Alignment>>,
    underline: Option<Option<Underline>>,
    strike: Option<Option<Strike>>,
    fill: Option<Option<CellFill>>,
    anchor: Option<Option<Anchor>>,
    margins: [Option<Option<i64>>; 4],
    shape_fill: Option<Option<ShapeFill>>,
    geometry: Option<Preset>,
}

#[allow(
    clippy::option_option,
    reason = "omission, removal and replacement are distinct assignments"
)]
impl<'a> TextProperties<'a> {
    /// Replace the target's text. A run keeps its run properties; a
    /// paragraph, shape or cell keeps its first paragraph's properties and
    /// first run's properties. A line feed starts a line break inside a
    /// paragraph and a new paragraph inside a shape or cell; unlike
    /// [`super::Edit::set_shape_texts`], a shape whose body holds no run
    /// (an empty placeholder) is accepted.
    #[must_use]
    pub const fn with_text(mut self, text: &'a str) -> Self {
        self.text = Some(text);
        self
    }
    /// Set or remove `a:rPr@b` on every addressed run.
    #[must_use]
    pub const fn with_bold(mut self, value: Option<bool>) -> Self {
        self.bold = Some(value);
        self
    }
    /// Set or remove `a:rPr@i` on every addressed run.
    #[must_use]
    pub const fn with_italic(mut self, value: Option<bool>) -> Self {
        self.italic = Some(value);
        self
    }
    /// Set or remove `a:rPr@sz`, in hundredths of a point.
    #[must_use]
    pub const fn with_size(mut self, value: Option<u32>) -> Self {
        self.size = Some(value);
        self
    }
    /// Set or remove the run's solid text fill.
    #[must_use]
    pub const fn with_color(mut self, value: Option<Rgb>) -> Self {
        self.color = Some(value);
        self
    }
    /// Set or remove the latin, east-Asian and complex-script typefaces together.
    #[must_use]
    pub const fn with_font(mut self, value: Option<&'a str>) -> Self {
        self.font = Some(value);
        self
    }
    /// Set or remove only the latin typeface (`a:latin`).
    #[must_use]
    pub const fn with_latin(mut self, value: Option<&'a str>) -> Self {
        self.latin = Some(value);
        self
    }
    /// Set or remove `a:pPr@algn` on every addressed paragraph.
    #[must_use]
    pub const fn with_align(mut self, value: Option<Alignment>) -> Self {
        self.align = Some(value);
        self
    }
    /// Set or remove `a:rPr@u` on every addressed run (record 0333).
    #[must_use]
    pub const fn with_underline(mut self, value: Option<Underline>) -> Self {
        self.underline = Some(value);
        self
    }
    /// Set or remove `a:rPr@strike` on every addressed run (record 0333).
    #[must_use]
    pub const fn with_strike(mut self, value: Option<Strike>) -> Self {
        self.strike = Some(value);
        self
    }
    /// Set or remove a table cell's local fill (`a:tcPr`): a solid RGB or
    /// theme colour, or no fill; cells only.
    #[must_use]
    pub const fn with_fill(mut self, value: Option<CellFill>) -> Self {
        self.fill = Some(value);
        self
    }
    /// Set or remove a table cell's vertical anchor (`a:tcPr@anchor`);
    /// cells only (record 0333).
    #[must_use]
    pub const fn with_cell_anchor(mut self, value: Option<Anchor>) -> Self {
        self.anchor = Some(value);
        self
    }
    /// Set or remove one table cell inset (`a:tcPr@marL`, …), in EMU; cells
    /// only (record 0333).
    #[must_use]
    pub const fn with_cell_margin(mut self, margin: CellMargin, value: Option<i64>) -> Self {
        self.margins[margin as usize] = Some(value);
        self
    }
    /// Set or remove a shape's own fill (`p:spPr`); shape targets only.
    /// Removal deletes the local fill choice, so the shape inherits its fill
    /// from its style or placeholder.
    #[must_use]
    pub const fn with_shape_fill(mut self, value: Option<ShapeFill>) -> Self {
        self.shape_fill = Some(value);
        self
    }
    /// Replace a shape's geometry with a preset (`a:prstGeom`, with no
    /// adjust values); shape targets only. Assigning the shape's current
    /// preset keeps its adjust values and changes nothing.
    #[must_use]
    pub const fn with_geometry(mut self, value: Preset) -> Self {
        self.geometry = Some(value);
        self
    }
    /// The replacement text, if any.
    #[must_use]
    pub const fn text(&self) -> Option<&'a str> {
        self.text
    }
    fn has_run_properties(&self) -> bool {
        self.bold.is_some()
            || self.italic.is_some()
            || self.size.is_some()
            || self.color.is_some()
            || self.font.is_some()
            || self.latin.is_some()
            || self.underline.is_some()
            || self.strike.is_some()
    }
    fn has_shape_properties(&self) -> bool {
        self.shape_fill.is_some() || self.geometry.is_some()
    }
    fn has_body_properties(&self) -> bool {
        self.text.is_some() || self.has_run_properties() || self.align.is_some()
    }
    /// Whether a cell-only property (fill, anchor, insets) is assigned.
    fn has_cell_properties(&self) -> bool {
        self.fill.is_some() || self.anchor.is_some() || self.margins.iter().any(Option::is_some)
    }
    fn is_empty(&self) -> bool {
        !self.has_body_properties() && !self.has_cell_properties() && !self.has_shape_properties()
    }
}

/// One formatting assignment for [`super::Edit::set_text_properties`].
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct TextAssignment<'a> {
    target: TextTarget,
    properties: TextProperties<'a>,
}

impl<'a> TextAssignment<'a> {
    /// Assign `properties` to `target`.
    #[must_use]
    pub const fn new(target: TextTarget, properties: TextProperties<'a>) -> Self {
        Self { target, properties }
    }
    /// The addressed element.
    #[must_use]
    pub const fn target(&self) -> TextTarget {
        self.target
    }
    /// The assignments.
    #[must_use]
    pub const fn properties(&self) -> &TextProperties<'a> {
        &self.properties
    }
}

pub(crate) fn stage_text_properties_with_limits(
    xml: &[u8],
    edits: &[TextAssignment<'_>],
    max_text_bytes: usize,
    max_output_bytes: usize,
    shape_limits: crate::shape::Limits,
) -> Result<(Option<Vec<u8>>, usize)> {
    if edits.is_empty() {
        return Ok((None, 0));
    }
    if edits.len() > super::MAX_SHAPE_PROPERTIES_REPLACEMENTS {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::ShapeProperties),
            observed: edits.len(),
            maximum: super::MAX_SHAPE_PROPERTIES_REPLACEMENTS,
        });
    }
    let mut text_bytes = 0usize;
    for edit in edits {
        validate(edit)?;
        text_bytes = text_bytes
            .checked_add(edit.properties.text.map_or(0, str::len))
            .ok_or(Error::Limit {
                kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::TextBytes),
                observed: usize::MAX,
                maximum: max_text_bytes,
            })?;
    }
    if text_bytes > max_text_bytes {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::TextBytes),
            observed: text_bytes,
            maximum: max_text_bytes,
        });
    }
    stage_admitted_text_edits(xml, edits, max_output_bytes, None, shape_limits)
}

/// Stage several already bounded batches of text formatting edits, applied
/// in order, in one scene read and one read-back (record 0333). Each chunk
/// keeps the limits of [`stage_text_properties_with_limits`]; the chunk count is bounded
/// by the scene's shape limit and the total by 100,000 edits.
///
/// With `changes`, it is cleared and then receives, for each edit in order,
/// whether that edit changed the bytes of its target element as the earlier
/// edits left them (record 0348), as one edit applied after the other
/// reports it.
pub(crate) fn stage_text_property_chunks_with_limits(
    xml: &[u8],
    chunks: &[&[TextAssignment<'_>]],
    max_text_bytes: usize,
    max_output_bytes: usize,
    mut changes: Option<&mut Vec<bool>>,
    shape_limits: crate::shape::Limits,
) -> Result<(Option<Vec<u8>>, usize)> {
    if let Some(changes) = changes.as_deref_mut() {
        changes.clear();
    }
    if chunks.len() > shape_limits.shapes() {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::ShapeProperties),
            observed: chunks.len(),
            maximum: shape_limits.shapes(),
        });
    }
    let mut total = 0usize;
    for chunk in chunks {
        if chunk.len() > super::MAX_SHAPE_PROPERTIES_REPLACEMENTS {
            return Err(Error::Limit {
                kind: crate::opened::LimitKind::Codec(
                    crate::opened::CodecLimitKind::ShapeProperties,
                ),
                observed: chunk.len(),
                maximum: super::MAX_SHAPE_PROPERTIES_REPLACEMENTS,
            });
        }
        let mut text_bytes = 0usize;
        for edit in *chunk {
            validate(edit)?;
            text_bytes = text_bytes
                .checked_add(edit.properties.text.map_or(0, str::len))
                .ok_or(Error::Limit {
                    kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::TextBytes),
                    observed: usize::MAX,
                    maximum: max_text_bytes,
                })?;
        }
        if text_bytes > max_text_bytes {
            return Err(Error::Limit {
                kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::TextBytes),
                observed: text_bytes,
                maximum: max_text_bytes,
            });
        }
        total = total.checked_add(chunk.len()).ok_or(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::ShapeProperties),
            observed: usize::MAX,
            maximum: 100_000,
        })?;
    }
    if total > 100_000 {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::ShapeProperties),
            observed: total,
            maximum: 100_000,
        });
    }
    if total == 0 {
        return Ok((None, 0));
    }
    let mut edits = Vec::new();
    edits
        .try_reserve_exact(total)
        .map_err(|source| Error::Allocation {
            resource: "opened-presentation text property chunks",
            source,
        })?;
    for chunk in chunks {
        edits.extend_from_slice(chunk);
    }
    stage_admitted_text_edits(xml, &edits, max_output_bytes, changes, shape_limits)
}

/// The staging of admitted text formatting edits: one scene read, the
/// edited elements rewritten in order, one splice and one read-back.
///
/// With `changes`, each edit's change (its element's bytes differ after it)
/// is appended in order (record 0348).
fn stage_admitted_text_edits(
    xml: &[u8],
    edits: &[TextAssignment<'_>],
    max_output_bytes: usize,
    mut changes: Option<&mut Vec<bool>>,
    shape_limits: crate::shape::Limits,
) -> Result<(Option<Vec<u8>>, usize)> {
    if let Some(changes) = changes.as_deref_mut() {
        changes
            .try_reserve_exact(edits.len())
            .map_err(|source| Error::Allocation {
                resource: "opened-presentation text property changes",
                source,
            })?;
    }
    // Marks the current edit (the last entry) changed.
    let mark = |changes: &mut Option<&mut Vec<bool>>| {
        if let Some(last) = changes
            .as_deref_mut()
            .and_then(|changes| changes.last_mut())
        {
            *last = true;
        }
    };
    let original = crate::shape::Scene::read(xml, shape_limits)?;
    if original.is_rewritten() {
        return Err(Error::refused(
            litchi_core::edit::RefusalKind::UnsupportedContent,
            "set_text_properties",
            "text property edits do not support markup-compatibility branch selection",
        ));
    }
    // Every target is a leaf element of the original scene (a text shape or
    // a table frame), so their spans are disjoint: each edit rewrites its
    // shape's current element, and the rewritten elements are spliced into
    // the part once (record 0333). This keeps a batch proportional to the
    // part plus the edited elements, rather than reading the scene again
    // for every edit.
    let mut elements: Vec<(usize, Range<usize>, Vec<u8>)> = Vec::new();
    let mut slots: std::collections::HashMap<usize, usize> = std::collections::HashMap::new();
    // A table frame's cells, located once and rewritten in place, so a
    // batch of cell edits is proportional to the table plus the edited
    // cells rather than reparsing the frame per edit (record 0346).
    let mut tables: std::collections::HashMap<usize, TableCells> = std::collections::HashMap::new();
    // A shape's paragraphs, located once, for paragraph and run edits that
    // touch nothing outside their paragraph (record 0346): a batch of them
    // on one body (the notes page's, a long text box's) is proportional to
    // the body plus the edited paragraphs.
    let mut bodies: std::collections::HashMap<usize, ShapeParagraphs> =
        std::collections::HashMap::new();
    let mut changed = 0;
    for edit in edits {
        if let Some(changes) = changes.as_deref_mut() {
            changes.push(false);
        }
        let index = shape_position(edit.target).get();
        let shape = original.at(index)?;
        let slot = match slots.get(&index) {
            Some(slot) => *slot,
            None => {
                let span = shape.common().span()?.range(xml.len())?;
                elements
                    .try_reserve(1)
                    .map_err(|source| Error::Allocation {
                        resource: "opened-presentation text property elements",
                        source,
                    })?;
                slots.insert(index, elements.len());
                elements.push((index, span.clone(), xml[span].to_vec()));
                elements.len() - 1
            },
        };
        let is_table = matches!(shape, crate::shape::Shape::Table(_));
        let element = &elements[slot].2;
        let cell = match edit.target {
            TextTarget::Cell { row, column, .. }
            | TextTarget::CellParagraph { row, column, .. }
            | TextTarget::CellRun { row, column, .. }
            | TextTarget::CellRunSpan { row, column, .. } => Some((row.get(), column.get())),
            _ => None,
        };
        let replacement = match cell {
            Some((row, column)) => {
                if !is_table {
                    return Err(Error::refused(
                        litchi_core::edit::RefusalKind::UnsupportedTarget,
                        "set_text_properties",
                        "a table cell edit requires a table graphic frame",
                    ));
                }
                if !tables.contains_key(&slot) {
                    let cells = TableCells::locate(element)?;
                    tables.try_reserve(1).map_err(|source| Error::Allocation {
                        resource: "opened-presentation text property tables",
                        source,
                    })?;
                    tables.insert(slot, cells);
                }
                let Some(cells) = tables.get_mut(&slot) else {
                    unreachable!("the table's cells were just located")
                };
                let current = cells.cell(element, row, column)?;
                let rewritten = match edit.target {
                    TextTarget::CellParagraph { paragraph, .. } => {
                        rewrite_cell_paragraph(current, paragraph.get(), &edit.properties)
                    },
                    TextTarget::CellRun { paragraph, run, .. } => rewrite_cell_run(
                        current,
                        paragraph.get(),
                        run.get(),
                        None,
                        &edit.properties,
                    ),
                    TextTarget::CellRunSpan {
                        paragraph,
                        run,
                        start,
                        end,
                        ..
                    } => rewrite_cell_run(
                        current,
                        paragraph.get(),
                        run.get(),
                        Some(start..end),
                        &edit.properties,
                    ),
                    _ => rewrite_cell(current, &edit.properties),
                }?;
                if rewritten.as_slice() != current {
                    cells.rewritten.insert((row, column), rewritten);
                    changed += 1;
                    mark(&mut changes);
                }
                continue;
            },
            None => {
                if is_table || edit.properties.has_cell_properties() {
                    return Err(Error::refused(
                        litchi_core::edit::RefusalKind::UnsupportedTarget,
                        "set_text_properties",
                        "text formatting requires a shape text body; fill applies to table cells",
                    ));
                }
                let local = match edit.target {
                    TextTarget::Paragraph { paragraph, .. } => Some((paragraph.get(), None, None)),
                    TextTarget::Run { paragraph, run, .. } => {
                        Some((paragraph.get(), Some(run.get()), None))
                    },
                    TextTarget::RunSpan {
                        paragraph,
                        run,
                        start,
                        end,
                        ..
                    } => Some((paragraph.get(), Some(run.get()), Some(start..end))),
                    _ => None,
                }
                .filter(|_| {
                    edit.properties.has_body_properties() && !edit.properties.has_shape_properties()
                });
                match local {
                    Some((paragraph, run, span)) => {
                        if !bodies.contains_key(&slot) {
                            let located = ShapeParagraphs::locate(element)?;
                            bodies.try_reserve(1).map_err(|source| Error::Allocation {
                                resource: "opened-presentation text property bodies",
                                source,
                            })?;
                            bodies.insert(slot, located);
                        }
                        let Some(located) = bodies.get_mut(&slot) else {
                            unreachable!("the shape's paragraphs were just located")
                        };
                        let current = located.paragraph(element, paragraph)?;
                        let rewritten = rewrite_paragraph(current, run, span, &edit.properties)?;
                        if rewritten.as_slice() != current {
                            located.rewritten.insert(paragraph, rewritten);
                            changed += 1;
                            mark(&mut changes);
                        }
                        continue;
                    },
                    None => {
                        // A whole-shape edit sees the paragraph edits before it.
                        if let Some(located) = bodies.remove(&slot)
                            && !located.rewritten.is_empty()
                        {
                            elements[slot].2 = located.splice(&elements[slot].2)?;
                        }
                        rewrite_shape(&elements[slot].2, edit.target, &edit.properties)?
                    },
                }
            },
        };
        let element = &elements[slot].2;
        if replacement != *element {
            elements[slot].2 = replacement;
            changed += 1;
            mark(&mut changes);
        }
    }
    if changed == 0 {
        return Ok((None, 0));
    }
    for (slot, cells) in tables {
        if !cells.rewritten.is_empty() {
            elements[slot].2 = cells.splice(&elements[slot].2)?;
        }
    }
    for (slot, located) in bodies {
        if !located.rewritten.is_empty() {
            elements[slot].2 = located.splice(&elements[slot].2)?;
        }
    }
    elements.sort_unstable_by_key(|(_, span, _)| span.start);
    if elements
        .windows(2)
        .any(|pair| pair[0].1.end > pair[1].1.start)
    {
        return Err(Error::refused(
            litchi_core::edit::RefusalKind::UnsupportedTarget,
            "set_text_properties",
            "text property edits address nested shapes",
        ));
    }
    let removed: usize = elements.iter().map(|(_, span, _)| span.len()).sum();
    let added: usize = elements.iter().map(|(_, _, bytes)| bytes.len()).sum();
    let length = (xml.len() - removed)
        .checked_add(added)
        .ok_or(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::OutputBytes),
            observed: usize::MAX,
            maximum: max_output_bytes,
        })?;
    if length > max_output_bytes {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::OutputBytes),
            observed: length,
            maximum: max_output_bytes,
        });
    }
    let mut candidate = Vec::new();
    candidate
        .try_reserve_exact(length)
        .map_err(|source| Error::Allocation {
            resource: "opened-presentation text property output",
            source,
        })?;
    let mut cursor = 0;
    for (_, span, bytes) in &elements {
        candidate.extend_from_slice(&xml[cursor..span.start]);
        candidate.extend_from_slice(bytes);
        cursor = span.end;
    }
    candidate.extend_from_slice(&xml[cursor..]);
    if candidate == xml {
        return Ok((None, 0));
    }
    let published = crate::shape::Scene::read(&candidate, shape_limits)?;
    if published.is_rewritten() || published.len() != original.len() {
        return Err(Error::Invalid(
            "opened-presentation text property batch changed the semantic shape graph".into(),
        ));
    }
    for index in 0..original.len() {
        if original.at(index)?.id() != published.at(index)?.id() {
            return Err(Error::Invalid(
                "opened-presentation text property batch changed shape identities".into(),
            ));
        }
    }
    Ok((Some(candidate), changed))
}

const fn shape_position(target: TextTarget) -> Position {
    match target {
        TextTarget::Shape { shape }
        | TextTarget::Paragraph { shape, .. }
        | TextTarget::Run { shape, .. }
        | TextTarget::RunSpan { shape, .. }
        | TextTarget::Cell { shape, .. }
        | TextTarget::CellParagraph { shape, .. }
        | TextTarget::CellRun { shape, .. }
        | TextTarget::CellRunSpan { shape, .. } => shape,
    }
}

fn validate(edit: &TextAssignment<'_>) -> Result<()> {
    let properties = &edit.properties;
    if properties.is_empty() {
        return Err(Error::Invalid(
            "a text property edit assigns nothing".into(),
        ));
    }
    if let Some(Some(size)) = properties.size
        && !(MIN_FONT_SIZE..=MAX_FONT_SIZE).contains(&size)
    {
        return Err(Error::Invalid(
            "font size must be between 1 and 4000 points".into(),
        ));
    }
    for face in [properties.font, properties.latin]
        .into_iter()
        .flatten()
        .flatten()
    {
        if face.is_empty() || face.len() > MAX_TYPEFACE_BYTES {
            return Err(Error::Invalid(
                "a typeface name must have 1 to 1024 bytes".into(),
            ));
        }
        check_characters(face)?;
    }
    if let Some(text) = properties.text {
        check_characters(text)?;
        if matches!(
            edit.target,
            TextTarget::Run { .. }
                | TextTarget::RunSpan { .. }
                | TextTarget::CellRun { .. }
                | TextTarget::CellRunSpan { .. }
        ) && text.contains(['\n', '\r'])
        {
            return Err(Error::Invalid(
                "run text cannot contain a line break".into(),
            ));
        }
    }
    if let TextTarget::RunSpan { start, end, .. } | TextTarget::CellRunSpan { start, end, .. } =
        edit.target
        && start >= end
    {
        return Err(Error::Invalid("a run span is empty".into()));
    }
    if properties.has_cell_properties() && !matches!(edit.target, TextTarget::Cell { .. }) {
        return Err(Error::Invalid("fill applies to table cells only".into()));
    }
    for margin in properties.margins.into_iter().flatten().flatten() {
        if !(0..=crate::table::columns::MAX_COLUMN_WIDTH).contains(&margin) {
            return Err(Error::Invalid("a cell inset is out of range".into()));
        }
    }
    if properties.has_shape_properties() && !matches!(edit.target, TextTarget::Shape { .. }) {
        return Err(Error::Invalid(
            "shape fill and geometry apply to shape targets only".into(),
        ));
    }
    Ok(())
}

fn check_characters(value: &str) -> Result<()> {
    if value.chars().all(|ch| {
        matches!(ch, '\t' | '\n' | '\r' | '\u{20}'..='\u{d7ff}' | '\u{e000}'..='\u{fffd}' | '\u{10000}'..='\u{10ffff}')
    }) {
        Ok(())
    } else {
        Err(Error::Invalid(
            "text contains a character that XML cannot represent".into(),
        ))
    }
}

// ---------------------------------------------------------------------------
// One-level element views over borrowed bytes
// ---------------------------------------------------------------------------

/// One element inside a parent's byte slice.
#[derive(Debug, Clone)]
struct Node {
    /// Whole element, start tag through end tag.
    span: Range<usize>,
    /// End of the start tag.
    open_end: usize,
    /// Start of the end tag (equal to `span.end` for an empty element).
    close_start: usize,
    empty: bool,
    prefix_len: usize,
    name_len: usize,
}

impl Node {
    fn qname<'x>(&self, xml: &'x [u8]) -> &'x [u8] {
        &xml[self.span.start + 1..self.span.start + 1 + self.name_len]
    }
    fn local<'x>(&self, xml: &'x [u8]) -> &'x [u8] {
        let qname = self.qname(xml);
        if self.prefix_len == 0 {
            qname
        } else {
            &qname[self.prefix_len + 1..]
        }
    }
    fn prefix<'x>(&self, xml: &'x [u8]) -> &'x [u8] {
        &self.qname(xml)[..self.prefix_len]
    }
    fn content(&self) -> Range<usize> {
        self.open_end..self.close_start
    }
}

fn invalid(message: &'static str) -> Error {
    Error::Invalid(message.into())
}

fn qname_parts(name: &[u8]) -> (usize, usize) {
    let prefix_len = name.iter().position(|byte| *byte == b':').unwrap_or(0);
    (prefix_len, name.len())
}

/// Parse the root element of `xml` (which must start at offset 0 after
/// optional whitespace) as a [`Node`].
fn root(xml: &[u8]) -> Result<Node> {
    let mut nodes = elements(xml, 0..xml.len())?;
    if nodes.len() != 1 {
        return Err(invalid("text property edit expected one root element"));
    }
    Ok(nodes.remove(0))
}

/// The direct child elements inside `range` of `xml`.
fn elements(xml: &[u8], range: Range<usize>) -> Result<Vec<Node>> {
    let base = range.start;
    let slice = &xml[range];
    let mut reader = Reader::from_reader(slice);
    reader.config_mut().check_end_names = true;
    let mut nodes = Vec::new();
    let mut depth = 0usize;
    let mut open: Option<(usize, usize, usize, usize)> = None;
    loop {
        let before = usize::try_from(reader.buffer_position()).unwrap_or(usize::MAX);
        let event = reader
            .read_event()
            .map_err(|_| invalid("text property edit found malformed XML"))?;
        let after = usize::try_from(reader.buffer_position()).unwrap_or(usize::MAX);
        match event {
            Event::Start(start) => {
                if depth == 0 {
                    let (prefix_len, name_len) = qname_parts(start.name().as_ref());
                    open = Some((before, after, prefix_len, name_len));
                }
                depth += 1;
            },
            Event::Empty(start) => {
                if depth == 0 {
                    let (prefix_len, name_len) = qname_parts(start.name().as_ref());
                    nodes.push(Node {
                        span: base + before..base + after,
                        open_end: base + after,
                        close_start: base + after,
                        empty: true,
                        prefix_len,
                        name_len,
                    });
                }
            },
            Event::End(_) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("text property edit found an unbalanced end tag"))?;
                if depth == 0 {
                    let (start, open_end, prefix_len, name_len) = open
                        .take()
                        .ok_or_else(|| invalid("text property edit lost an element start"))?;
                    nodes.push(Node {
                        span: base + start..base + after,
                        open_end: base + open_end,
                        close_start: base + before,
                        empty: false,
                        prefix_len,
                        name_len,
                    });
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if depth != 0 {
        return Err(invalid("text property edit found an unclosed element"));
    }
    Ok(nodes)
}

fn children(xml: &[u8], node: &Node) -> Result<Vec<Node>> {
    if node.empty {
        Ok(Vec::new())
    } else {
        elements(xml, node.content())
    }
}

fn child<'n>(xml: &[u8], nodes: &'n [Node], local: &[u8]) -> Option<&'n Node> {
    nodes.iter().find(|node| node.local(xml) == local)
}

/// Splices relative to one element slice, applied back to front.
struct Splices(Vec<(Range<usize>, Vec<u8>)>);

impl Splices {
    fn new() -> Self {
        Self(Vec::new())
    }
    fn replace(&mut self, range: Range<usize>, bytes: Vec<u8>) {
        self.0.push((range, bytes));
    }
    fn insert(&mut self, at: usize, bytes: Vec<u8>) {
        self.0.push((at..at, bytes));
    }
    fn apply(mut self, xml: &[u8]) -> Result<Vec<u8>> {
        self.0.sort_by_key(|(range, _)| (range.start, range.end));
        for pair in self.0.windows(2) {
            if pair[0].0.end > pair[1].0.start {
                return Err(invalid("text property edit produced overlapping splices"));
            }
        }
        let mut output = Vec::with_capacity(xml.len() + 64);
        let mut cursor = 0;
        for (range, bytes) in &self.0 {
            output.extend_from_slice(&xml[cursor..range.start]);
            output.extend_from_slice(bytes);
            cursor = range.end;
        }
        output.extend_from_slice(&xml[cursor..]);
        Ok(output)
    }
}

// ---------------------------------------------------------------------------
// Start-tag attribute editing
// ---------------------------------------------------------------------------

/// Rewrite a start tag, setting (`Some`) or removing (`None`) attributes.
/// Untouched attributes keep their exact source spelling.
fn rewrite_start_tag(
    tag: &[u8],
    changes: &[(&[u8], Option<&[u8]>)],
    empty: bool,
) -> Result<Vec<u8>> {
    let body_end = if tag.ends_with(b"/>") {
        tag.len() - 2
    } else if tag.ends_with(b">") {
        tag.len() - 1
    } else {
        return Err(invalid("text property edit found a malformed start tag"));
    };
    let inner = &tag[1..body_end];
    let name_end = inner
        .iter()
        .position(|byte| byte.is_ascii_whitespace())
        .unwrap_or(inner.len());
    let mut output = Vec::with_capacity(tag.len() + 32);
    output.push(b'<');
    output.extend_from_slice(&inner[..name_end]);
    // Walk raw attributes: name, '=', quoted value.
    let mut cursor = name_end;
    let bytes = inner;
    while cursor < bytes.len() {
        while cursor < bytes.len() && bytes[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if cursor >= bytes.len() {
            break;
        }
        let name_start = cursor;
        while cursor < bytes.len() && bytes[cursor] != b'=' && !bytes[cursor].is_ascii_whitespace()
        {
            cursor += 1;
        }
        let name = &bytes[name_start..cursor];
        while cursor < bytes.len() && bytes[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if bytes.get(cursor) != Some(&b'=') {
            return Err(invalid("text property edit found a malformed attribute"));
        }
        cursor += 1;
        while cursor < bytes.len() && bytes[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        let quote = *bytes
            .get(cursor)
            .ok_or_else(|| invalid("text property edit found a malformed attribute"))?;
        if quote != b'"' && quote != b'\'' {
            return Err(invalid("text property edit found an unquoted attribute"));
        }
        let value_start = cursor;
        cursor += 1;
        while cursor < bytes.len() && bytes[cursor] != quote {
            cursor += 1;
        }
        if cursor >= bytes.len() {
            return Err(invalid(
                "text property edit found an unterminated attribute",
            ));
        }
        cursor += 1;
        if changes.iter().any(|(key, _)| *key == name) {
            continue;
        }
        output.push(b' ');
        output.extend_from_slice(name);
        output.push(b'=');
        output.extend_from_slice(&bytes[value_start..cursor]);
    }
    for (key, value) in changes {
        if let Some(value) = value {
            output.push(b' ');
            output.extend_from_slice(key);
            output.extend_from_slice(b"=\"");
            output.extend_from_slice(value);
            output.push(b'"');
        }
    }
    output.extend_from_slice(if empty { b"/>" } else { b">" });
    Ok(output)
}

fn escape_attribute(value: &str) -> Vec<u8> {
    quick_xml::escape::escape(value).as_bytes().to_vec()
}

fn escape_text(value: &str, maximum: usize) -> Result<Vec<u8>> {
    Ok(super::xml::escape_character_data(value, maximum)?
        .as_bytes()
        .to_vec())
}

fn qualified(prefix: &[u8], local: &str) -> Vec<u8> {
    let mut name = Vec::with_capacity(prefix.len() + local.len() + 1);
    if !prefix.is_empty() {
        name.extend_from_slice(prefix);
        name.push(b':');
    }
    name.extend_from_slice(local.as_bytes());
    name
}

fn srgb_fill(prefix: &[u8], color: Rgb) -> Vec<u8> {
    let fill = qualified(prefix, "solidFill");
    let clr = qualified(prefix, "srgbClr");
    let mut output = Vec::with_capacity(64);
    output.push(b'<');
    output.extend_from_slice(&fill);
    output.extend_from_slice(b"><");
    output.extend_from_slice(&clr);
    output.extend_from_slice(
        format!(
            " val=\"{:02X}{:02X}{:02X}\"/></",
            color.0[0], color.0[1], color.0[2]
        )
        .as_bytes(),
    );
    output.extend_from_slice(&fill);
    output.push(b'>');
    output
}

/// The fill markup of a table cell.
fn cell_fill(prefix: &[u8], fill: CellFill) -> Vec<u8> {
    match fill {
        CellFill::Rgb(color) => srgb_fill(prefix, color),
        CellFill::Scheme(scheme) => {
            let fill = qualified(prefix, "solidFill");
            let mut output = b"<".to_vec();
            output.extend_from_slice(&fill);
            output.extend_from_slice(b"><");
            output.extend_from_slice(&qualified(prefix, "schemeClr"));
            output.extend_from_slice(b" val=\"");
            output.extend_from_slice(scheme.token().as_bytes());
            output.extend_from_slice(b"\"/></");
            output.extend_from_slice(&fill);
            output.push(b'>');
            output
        },
        CellFill::None => {
            let mut output = b"<".to_vec();
            output.extend_from_slice(&qualified(prefix, "noFill"));
            output.extend_from_slice(b"/>");
            output
        },
    }
}

fn typeface(prefix: &[u8], local: &str, face: &str) -> Vec<u8> {
    let mut output = Vec::with_capacity(32 + face.len());
    output.push(b'<');
    output.extend_from_slice(&qualified(prefix, local));
    output.extend_from_slice(b" typeface=\"");
    output.extend_from_slice(&escape_attribute(face));
    output.extend_from_slice(b"\"/>");
    output
}

// ---------------------------------------------------------------------------
// Properties elements with schema-ordered children
// ---------------------------------------------------------------------------

fn rpr_rank(local: &[u8]) -> u8 {
    match local {
        b"ln" => 0,
        b"noFill" | b"solidFill" | b"gradFill" | b"blipFill" | b"pattFill" | b"grpFill" => 1,
        b"effectLst" | b"effectDag" => 2,
        b"highlight" => 3,
        b"uLnTx" | b"uLn" => 4,
        b"uFillTx" | b"uFill" => 5,
        b"latin" => 6,
        b"ea" => 7,
        b"cs" => 8,
        b"sym" => 9,
        b"hlinkClick" => 10,
        b"hlinkMouseOver" => 11,
        b"rtl" => 12,
        _ => 13,
    }
}

fn tcpr_rank(local: &[u8]) -> u8 {
    match local {
        b"lnL" => 0,
        b"lnR" => 1,
        b"lnT" => 2,
        b"lnB" => 3,
        b"lnTlToBr" => 4,
        b"lnBlToTr" => 5,
        b"cell3D" => 6,
        b"noFill" | b"solidFill" | b"gradFill" | b"blipFill" | b"pattFill" | b"grpFill" => 7,
        b"headers" => 8,
        _ => 9,
    }
}

/// A requested change to the children of a properties element: replace
/// every child of `rank` with `bytes` (or remove them for `None`).
struct ChildChange {
    rank: u8,
    bytes: Option<Vec<u8>>,
}

/// Rewrite a properties element (`a:rPr`, `a:pPr`, `a:tcPr`, …) whose bytes
/// are `element`, applying attribute and ranked child changes.
fn rewrite_properties(
    element: &[u8],
    attributes: &[(&[u8], Option<&[u8]>)],
    child_changes: &[ChildChange],
    rank: fn(&[u8]) -> u8,
) -> Result<Vec<u8>> {
    let node = root(element)?;
    let kids = children(element, &node)?;
    let has_new_children = child_changes.iter().any(|change| change.bytes.is_some());
    let remaining = kids
        .iter()
        .filter(|kid| {
            let kid_rank = rank(kid.local(element));
            !child_changes.iter().any(|change| change.rank == kid_rank)
        })
        .count();
    let empty = remaining == 0 && !has_new_children;
    let start_tag = rewrite_start_tag(&element[node.span.start..node.open_end], attributes, empty)?;
    if empty {
        return Ok(start_tag);
    }
    let mut output = Vec::with_capacity(element.len() + 64);
    output.extend_from_slice(&start_tag);
    let mut pending: Vec<&ChildChange> = child_changes.iter().collect();
    pending.sort_by_key(|change| change.rank);
    let mut cursor = node.open_end;
    let mut pending_index = 0;
    for kid in &kids {
        let kid_rank = rank(kid.local(element));
        // Insert new children ranked before this kid.
        while pending_index < pending.len() && pending[pending_index].rank < kid_rank {
            output.extend_from_slice(&element[cursor..kid.span.start]);
            cursor = kid.span.start;
            if let Some(bytes) = &pending[pending_index].bytes {
                output.extend_from_slice(bytes);
            }
            pending_index += 1;
        }
        output.extend_from_slice(&element[cursor..kid.span.start]);
        cursor = kid.span.end;
        if let Some(change) = child_changes.iter().find(|change| change.rank == kid_rank) {
            // The first matching kid takes the replacement; later ones of the
            // same rank are dropped.
            if pending_index < pending.len() && pending[pending_index].rank == kid_rank {
                if let Some(bytes) = &change.bytes {
                    output.extend_from_slice(bytes);
                }
                pending_index += 1;
            }
        } else {
            output.extend_from_slice(&element[kid.span.clone()]);
        }
    }
    let close_start = if node.empty {
        element.len()
    } else {
        node.close_start
    };
    output.extend_from_slice(&element[cursor.min(close_start)..close_start]);
    while pending_index < pending.len() {
        if let Some(bytes) = &pending[pending_index].bytes {
            output.extend_from_slice(bytes);
        }
        pending_index += 1;
    }
    output.extend_from_slice(b"</");
    output.extend_from_slice(node.qname(element));
    output.push(b'>');
    Ok(output)
}

fn bool_value(value: bool) -> &'static [u8] {
    if value { b"1" } else { b"0" }
}

/// Rewrite an `a:rPr` (or `a:endParaRPr`) element's bytes.
fn rewrite_rpr(element: &[u8], prefix: &[u8], properties: &TextProperties<'_>) -> Result<Vec<u8>> {
    let size;
    let mut attributes: Vec<(&[u8], Option<&[u8]>)> = Vec::new();
    if let Some(bold) = properties.bold {
        attributes.push((b"b", bold.map(bool_value)));
    }
    if let Some(italic) = properties.italic {
        attributes.push((b"i", italic.map(bool_value)));
    }
    if let Some(value) = properties.size {
        size = value.map(|value| value.to_string());
        attributes.push((b"sz", size.as_deref().map(str::as_bytes)));
    }
    if let Some(value) = properties.underline {
        attributes.push((b"u", value.map(|value| value.dml().as_bytes())));
    }
    if let Some(value) = properties.strike {
        attributes.push((b"strike", value.map(|value| value.token().as_bytes())));
    }
    let mut changes = Vec::new();
    if let Some(color) = properties.color {
        changes.push(ChildChange {
            rank: 1,
            bytes: color.map(|color| srgb_fill(prefix, color)),
        });
    }
    let latin = properties.latin.or(properties.font);
    if let Some(face) = latin {
        changes.push(ChildChange {
            rank: 6,
            bytes: face.map(|face| typeface(prefix, "latin", face)),
        });
    }
    if let Some(face) = properties.font {
        changes.push(ChildChange {
            rank: 7,
            bytes: face.map(|face| typeface(prefix, "ea", face)),
        });
        changes.push(ChildChange {
            rank: 8,
            bytes: face.map(|face| typeface(prefix, "cs", face)),
        });
    }
    rewrite_properties(element, &attributes, &changes, rpr_rank)
}

fn alignment_token(value: Alignment) -> &'static [u8] {
    match value {
        Alignment::Left => b"l",
        Alignment::Center => b"ctr",
        Alignment::Right => b"r",
        Alignment::Justified => b"just",
        Alignment::JustifiedLow => b"justLow",
        Alignment::Distributed => b"dist",
        Alignment::ThaiDistributed => b"thaiDist",
    }
}

// ---------------------------------------------------------------------------
// Runs, paragraphs, bodies
// ---------------------------------------------------------------------------

/// Refuse a text replacement of `range` of `owner` (a paragraph) when one of
/// its children there is not a run or a line break: a field (`a:fld`) or
/// any other markup would be lost (record 0304). Paragraph and end-of-
/// paragraph properties are kept by the rewrites.
fn check_replaceable(owner: &[u8], kids: &[Node], range: &Range<usize>) -> Result<()> {
    for kid in kids.iter().filter(|kid| range.contains(&kid.span.start)) {
        if !matches!(kid.local(owner), b"r" | b"br" | b"pPr" | b"endParaRPr") {
            return Err(dropped_content());
        }
    }
    Ok(())
}

fn dropped_content() -> Error {
    Error::refused(
        litchi_core::edit::RefusalKind::UnsupportedContent,
        "set_text_properties",
        "the replaced text holds a field (a:fld) or other markup that is not a run \
                 or a line break; set the text of its runs instead",
    )
}

fn is_run(local: &[u8]) -> bool {
    matches!(local, b"r" | b"br" | b"fld")
}

/// Rewrite one run element (`a:r`, `a:br`, `a:fld`) for run properties and
/// optional replacement text.
pub(super) fn rewrite_run(
    run: &[u8],
    properties: &TextProperties<'_>,
    text: Option<&str>,
) -> Result<Vec<u8>> {
    let node = root(run)?;
    let prefix = node.prefix(run).to_vec();
    let field = node.local(run) == b"fld";
    if text.is_some() && node.local(run) != b"r" && !field {
        return Err(Error::refused(
            litchi_core::edit::RefusalKind::UnsupportedTarget,
            "set_text_properties",
            "run text applies to a text run (a:r) or a field's cached result (a:fld)",
        ));
    }
    let kids = children(run, &node)?;
    let mut splices = Splices::new();
    let mut body = run.to_vec();
    let mut node = node;
    if node.empty {
        // `<a:br/>` → `<a:br></a:br>` so a properties child can be inserted.
        if !properties.has_run_properties() {
            return Ok(run.to_vec());
        }
        let qname = node.qname(run).to_vec();
        let mut expanded = run[..run.len() - 2].to_vec();
        expanded.push(b'>');
        let open_end = expanded.len();
        expanded.extend_from_slice(b"</");
        expanded.extend_from_slice(&qname);
        expanded.push(b'>');
        body = expanded;
        node = Node {
            span: 0..body.len(),
            open_end,
            close_start: open_end,
            empty: false,
            prefix_len: node.prefix_len,
            name_len: node.name_len,
        };
    }
    if properties.has_run_properties() {
        match child(&body, &kids, b"rPr") {
            Some(rpr) if !kids.is_empty() => {
                let replaced = rewrite_rpr(&body[rpr.span.clone()], &prefix, properties)?;
                splices.replace(rpr.span.clone(), replaced);
            },
            _ => {
                let mut empty = b"<".to_vec();
                empty.extend_from_slice(&qualified(&prefix, "rPr"));
                empty.extend_from_slice(b"/>");
                let created = rewrite_rpr(&empty, &prefix, properties)?;
                splices.insert(node.open_end, created);
            },
        }
    }
    if let Some(text) = text {
        let escaped = escape_text(text, text.len().saturating_mul(6).max(16))?;
        let Some(t) = child(&body, &kids, b"t") else {
            if !field || node.empty {
                return Err(invalid("a text run has no a:t element"));
            }
            // A field without a cached result: its a:t follows a:rPr and
            // a:pPr (record 0279).
            let mut element = b"<".to_vec();
            element.extend_from_slice(&qualified(&prefix, "t"));
            element.push(b'>');
            element.extend_from_slice(&escaped);
            element.extend_from_slice(b"</");
            element.extend_from_slice(&qualified(&prefix, "t"));
            element.push(b'>');
            splices.insert(node.close_start, element);
            return splices.apply(&body);
        };
        let qname = t.qname(&body).to_vec();
        let mut element = b"<".to_vec();
        element
            .extend_from_slice(&body[t.span.start + 1..t.open_end - if t.empty { 2 } else { 1 }]);
        element.push(b'>');
        element.extend_from_slice(&escaped);
        element.extend_from_slice(b"</");
        element.extend_from_slice(&qname);
        element.push(b'>');
        splices.replace(t.span.clone(), element);
    }
    splices.apply(&body)
}

/// Split a text run so that bytes `span` of its text are a run of their
/// own, with `properties` (and their replacement text) applied to it; the
/// other parts keep the run's markup with their part of the text.
fn split_run(run: &[u8], span: Range<usize>, properties: &TextProperties<'_>) -> Result<Vec<u8>> {
    let node = root(run)?;
    if node.local(run) != b"r" {
        return Err(Error::refused(
            litchi_core::edit::RefusalKind::UnsupportedTarget,
            "set_text_properties",
            "a run span applies to a text run (a:r)",
        ));
    }
    let kids = children(run, &node)?;
    let t = child(run, &kids, b"t").ok_or_else(|| invalid("a text run has no a:t element"))?;
    let text = if t.empty {
        String::new()
    } else {
        let raw =
            std::str::from_utf8(&run[t.content()]).map_err(|_| invalid("run text is not UTF-8"))?;
        if raw.contains("<![CDATA[") {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                "set_text_properties",
                "a run span cannot split character data sections",
            ));
        }
        quick_xml::escape::unescape(raw)
            .map_err(|_| invalid("run text holds an unknown entity"))?
            .into_owned()
    };
    if span.start >= span.end
        || span.end > text.len()
        || !text.is_char_boundary(span.start)
        || !text.is_char_boundary(span.end)
    {
        return Err(Error::Invalid(
            "a run span lies outside its run's text or off a character boundary".into(),
        ));
    }
    let plain = TextProperties::default();
    let mut output = Vec::with_capacity(run.len() * 3);
    if span.start > 0 {
        output.extend_from_slice(&rewrite_run(run, &plain, Some(&text[..span.start]))?);
    }
    let middle = properties.text.unwrap_or(&text[span.clone()]);
    output.extend_from_slice(&rewrite_run(run, properties, Some(middle))?);
    if span.end < text.len() {
        output.extend_from_slice(&rewrite_run(run, &plain, Some(&text[span.end..]))?);
    }
    Ok(output)
}

/// Build a fresh text run carrying `rpr` (raw, possibly empty).
fn new_run(prefix: &[u8], rpr: &[u8], text: &str) -> Result<Vec<u8>> {
    let r = qualified(prefix, "r");
    let t = qualified(prefix, "t");
    let escaped = escape_text(text, text.len().saturating_mul(6).max(16))?;
    let mut output = Vec::with_capacity(32 + rpr.len() + escaped.len());
    output.push(b'<');
    output.extend_from_slice(&r);
    output.push(b'>');
    output.extend_from_slice(rpr);
    output.push(b'<');
    output.extend_from_slice(&t);
    output.push(b'>');
    output.extend_from_slice(&escaped);
    output.extend_from_slice(b"</");
    output.extend_from_slice(&t);
    output.extend_from_slice(b"></");
    output.extend_from_slice(&r);
    output.push(b'>');
    Ok(output)
}

fn new_break(prefix: &[u8], rpr: &[u8]) -> Vec<u8> {
    let br = qualified(prefix, "br");
    let mut output = Vec::with_capacity(16 + rpr.len());
    output.push(b'<');
    output.extend_from_slice(&br);
    if rpr.is_empty() {
        output.extend_from_slice(b"/>");
    } else {
        output.push(b'>');
        output.extend_from_slice(rpr);
        output.extend_from_slice(b"</");
        output.extend_from_slice(&br);
        output.push(b'>');
    }
    output
}

/// The raw `a:rPr` of the paragraph's first text run, or an empty slice.
fn first_run_properties<'x>(paragraph: &'x [u8], kids: &[Node]) -> Result<&'x [u8]> {
    for kid in kids {
        if kid.local(paragraph) == b"r" {
            let run_kids = children(paragraph, kid)?;
            if let Some(rpr) = child(paragraph, &run_kids, b"rPr") {
                return Ok(&paragraph[rpr.span.clone()]);
            }
            return Ok(&[]);
        }
    }
    Ok(&[])
}

/// Rewrite one paragraph: alignment, run properties on the selected runs,
/// and optional replacement text for the whole paragraph or one run.
fn rewrite_paragraph(
    paragraph: &[u8],
    run: Option<usize>,
    span: Option<Range<usize>>,
    properties: &TextProperties<'_>,
) -> Result<Vec<u8>> {
    let node = root(paragraph)?;
    let prefix = node.prefix(paragraph).to_vec();
    let kids = children(paragraph, &node)?;
    let mut splices = Splices::new();
    if let Some(run) = run {
        let target = kids
            .iter()
            .filter(|kid| is_run(kid.local(paragraph)))
            .nth(run)
            .ok_or(Error::IndexOutOfBounds {
                index: run,
                len: kids
                    .iter()
                    .filter(|kid| is_run(kid.local(paragraph)))
                    .count(),
            })?;
        let element = &paragraph[target.span.clone()];
        let replaced = match span {
            Some(span) => split_run(element, span, properties)?,
            None => rewrite_run(element, properties, properties.text)?,
        };
        splices.replace(target.span.clone(), replaced);
        return splices.apply(paragraph);
    }
    if let Some(align) = properties.align {
        let token = align.map(alignment_token);
        let changes = [(b"algn".as_slice(), token)];
        match child(paragraph, &kids, b"pPr") {
            Some(ppr) => {
                let element = &paragraph[ppr.span.clone()];
                let rewritten = rewrite_properties(element, &changes, &[], |_| 0)?;
                splices.replace(ppr.span.clone(), rewritten);
            },
            None if token.is_some() => {
                let mut empty = b"<".to_vec();
                empty.extend_from_slice(&qualified(&prefix, "pPr"));
                empty.extend_from_slice(b"/>");
                splices.insert(node.open_end, rewrite_start_tag(&empty, &changes, true)?);
            },
            None => {},
        }
    }
    if let Some(text) = properties.text {
        let template = first_run_properties(paragraph, &kids)?;
        let template = if properties.has_run_properties() {
            let base = if template.is_empty() {
                let mut empty = b"<".to_vec();
                empty.extend_from_slice(&qualified(&prefix, "rPr"));
                empty.extend_from_slice(b"/>");
                empty
            } else {
                template.to_vec()
            };
            rewrite_rpr(&base, &prefix, properties)?
        } else {
            template.to_vec()
        };
        let mut runs = Vec::new();
        for (index, line) in text.split('\n').enumerate() {
            if index > 0 {
                runs.extend_from_slice(&new_break(&prefix, &template));
            }
            let line = line.strip_suffix('\r').unwrap_or(line);
            if !line.is_empty() {
                runs.extend_from_slice(&new_run(&prefix, &template, line)?);
            }
        }
        let run_nodes: Vec<&Node> = kids
            .iter()
            .filter(|kid| is_run(kid.local(paragraph)))
            .collect();
        match (run_nodes.first(), run_nodes.last()) {
            (Some(first), Some(last)) => {
                let replaced = first.span.start..last.span.end;
                check_replaceable(paragraph, &kids, &replaced)?;
                splices.replace(replaced, runs);
            },
            _ => {
                // After pPr if present, else at the start; before endParaRPr.
                let at = child(paragraph, &kids, b"pPr").map_or_else(
                    || {
                        if node.empty {
                            node.open_end
                        } else {
                            node.open_end
                        }
                    },
                    |ppr| ppr.span.end,
                );
                if node.empty {
                    return expand_and_insert(paragraph, &node, &splices, runs);
                }
                splices.insert(at, runs);
            },
        }
    } else if properties.has_run_properties() {
        for kid in kids.iter().filter(|kid| is_run(kid.local(paragraph))) {
            let replaced = rewrite_run(&paragraph[kid.span.clone()], properties, None)?;
            splices.replace(kid.span.clone(), replaced);
        }
    }
    splices.apply(paragraph)
}

/// An empty `<a:p/>` gains content: expand it, then apply pending splices.
fn expand_and_insert(
    paragraph: &[u8],
    node: &Node,
    splices: &Splices,
    runs: Vec<u8>,
) -> Result<Vec<u8>> {
    // Pending splices of an empty paragraph can only touch its start tag
    // (a pPr insertion at open_end); rebuild them in order.
    let mut output = paragraph[..paragraph.len() - 2].to_vec();
    output.push(b'>');
    for (_, bytes) in &splices.0 {
        output.extend_from_slice(bytes);
    }
    output.extend_from_slice(&runs);
    output.extend_from_slice(b"</");
    output.extend_from_slice(node.qname(paragraph));
    output.push(b'>');
    Ok(output)
}

/// Apply `properties` to selected paragraphs of a text body element.
fn rewrite_body(
    body: &[u8],
    paragraph: Option<usize>,
    run: Option<usize>,
    span: Option<Range<usize>>,
    properties: &TextProperties<'_>,
) -> Result<Vec<u8>> {
    let node = root(body)?;
    let kids = children(body, &node)?;
    let paragraphs: Vec<&Node> = kids.iter().filter(|kid| kid.local(body) == b"p").collect();
    let mut splices = Splices::new();
    match paragraph {
        Some(index) => {
            let target = paragraphs.get(index).ok_or(Error::IndexOutOfBounds {
                index,
                len: paragraphs.len(),
            })?;
            let replaced = rewrite_paragraph(&body[target.span.clone()], run, span, properties)?;
            splices.replace(target.span.clone(), replaced);
        },
        None => {
            for target in paragraphs {
                let replaced =
                    rewrite_paragraph(&body[target.span.clone()], None, None, properties)?;
                splices.replace(target.span.clone(), replaced);
            }
        },
    }
    splices.apply(body)
}

fn sppr_rank(local: &[u8]) -> u8 {
    match local {
        b"xfrm" => 0,
        b"custGeom" | b"prstGeom" => 1,
        b"noFill" | b"solidFill" | b"gradFill" | b"blipFill" | b"pattFill" | b"grpFill" => 2,
        b"ln" => 3,
        b"effectLst" | b"effectDag" => 4,
        b"scene3d" => 5,
        b"sp3d" => 6,
        _ => 7,
    }
}

/// The `DrawingML` prefix of a shape: that of its `p:spPr` children, else
/// that of its text body's children.
fn drawing_prefix(shape: &[u8], kids: &[Node], sppr_kids: &[Node]) -> Result<Vec<u8>> {
    if let Some(first) = sppr_kids.first() {
        return Ok(first.prefix(shape).to_vec());
    }
    if let Some(body) = child(shape, kids, b"txBody") {
        let body_kids = children(shape, body)?;
        if let Some(first) =
            child(shape, &body_kids, b"bodyPr").or_else(|| child(shape, &body_kids, b"p"))
        {
            return Ok(first.prefix(shape).to_vec());
        }
    }
    Err(Error::refused(
        litchi_core::edit::RefusalKind::UnsupportedTarget,
        "set_text_properties",
        "the shape has no DrawingML child naming the drawing prefix",
    ))
}

/// Rewrite a shape's `p:spPr` for fill and preset-geometry assignments.
fn rewrite_sppr(
    shape: &[u8],
    kids: &[Node],
    properties: &TextProperties<'_>,
    splices: &mut Splices,
) -> Result<()> {
    let sppr = child(shape, kids, b"spPr").ok_or(Error::refused(
        litchi_core::edit::RefusalKind::UnsupportedTarget,
        "set_text_properties",
        "the shape has no shape properties",
    ))?;
    let sppr_kids = children(shape, sppr)?;
    let prefix = drawing_prefix(shape, kids, &sppr_kids)?;
    let mut changes = Vec::new();
    if let Some(preset) = properties.geometry {
        let current = child(shape, &sppr_kids, b"prstGeom")
            .and_then(|geometry| preset_token(&shape[geometry.span.start..geometry.open_end]));
        let custom = child(shape, &sppr_kids, b"custGeom").is_some();
        if custom || current.as_deref() != Some(preset.token().as_bytes()) {
            let name = qualified(&prefix, "prstGeom");
            let list = qualified(&prefix, "avLst");
            let mut bytes = Vec::with_capacity(48 + preset.token().len());
            bytes.push(b'<');
            bytes.extend_from_slice(&name);
            bytes.extend_from_slice(b" prst=\"");
            bytes.extend_from_slice(preset.token().as_bytes());
            bytes.extend_from_slice(b"\"><");
            bytes.extend_from_slice(&list);
            bytes.extend_from_slice(b"/></");
            bytes.extend_from_slice(&name);
            bytes.push(b'>');
            changes.push(ChildChange {
                rank: 1,
                bytes: Some(bytes),
            });
        }
    }
    if let Some(fill) = properties.shape_fill {
        changes.push(ChildChange {
            rank: 2,
            bytes: fill.map(|fill| match fill {
                ShapeFill::Solid(color) => srgb_fill(&prefix, color),
                ShapeFill::None => {
                    let mut bytes = b"<".to_vec();
                    bytes.extend_from_slice(&qualified(&prefix, "noFill"));
                    bytes.extend_from_slice(b"/>");
                    bytes
                },
            }),
        });
    }
    if changes.is_empty() {
        return Ok(());
    }
    let element = &shape[sppr.span.clone()];
    let rewritten = rewrite_properties(element, &[], &changes, sppr_rank)?;
    if rewritten != element {
        splices.replace(sppr.span.clone(), rewritten);
    }
    Ok(())
}

/// The `prst` attribute of an `a:prstGeom` start tag.
fn preset_token(tag: &[u8]) -> Option<Vec<u8>> {
    let mut reader = Reader::from_reader(tag);
    let Ok(Event::Start(start) | Event::Empty(start)) = reader.read_event() else {
        return None;
    };
    for attribute in start.checked_attributes() {
        let attribute = attribute.ok()?;
        if attribute.key.as_ref() == b"prst" {
            return Some(attribute.value.into_owned());
        }
    }
    None
}

/// Rewrite a `p:sp` element for a shape, paragraph or run target.
fn rewrite_shape(
    shape: &[u8],
    target: TextTarget,
    properties: &TextProperties<'_>,
) -> Result<Vec<u8>> {
    let node = root(shape)?;
    let kids = children(shape, &node)?;
    let mut splices = Splices::new();
    if properties.has_shape_properties() {
        rewrite_sppr(shape, &kids, properties, &mut splices)?;
    }
    if !properties.has_body_properties() {
        return splices.apply(shape);
    }
    let body = child(shape, &kids, b"txBody").ok_or(Error::refused(
        litchi_core::edit::RefusalKind::UnsupportedTarget,
        "set_text_properties",
        "the shape has no text body",
    ))?;
    let (paragraph, run, span) = match target {
        TextTarget::Shape { .. } => (None, None, None),
        TextTarget::Paragraph { paragraph, .. } => (Some(paragraph.get()), None, None),
        TextTarget::Run { paragraph, run, .. } => (Some(paragraph.get()), Some(run.get()), None),
        TextTarget::RunSpan {
            paragraph,
            run,
            start,
            end,
            ..
        } => (Some(paragraph.get()), Some(run.get()), Some(start..end)),
        TextTarget::Cell { .. }
        | TextTarget::CellParagraph { .. }
        | TextTarget::CellRun { .. }
        | TextTarget::CellRunSpan { .. } => {
            unreachable!("cells are rewritten through their frame")
        },
    };
    let element = &shape[body.span.clone()];
    let replaced = match (target, properties.text) {
        (TextTarget::Shape { .. }, Some(text)) => {
            let body_node = root(element)?;
            let body_kids = children(element, &body_node)?;
            let prefix = child(element, &body_kids, b"bodyPr")
                .or_else(|| child(element, &body_kids, b"p"))
                .map(|node| node.prefix(element).to_vec())
                .ok_or_else(|| invalid("the shape text body has no drawing children"))?;
            rewrite_body_text(element, &prefix, text, properties)?
        },
        _ => rewrite_body(element, paragraph, run, span, properties)?,
    };
    splices.replace(body.span.clone(), replaced);
    splices.apply(shape)
}

// ---------------------------------------------------------------------------
// Table cells
// ---------------------------------------------------------------------------

fn descend(xml: &[u8], node: &Node, local: &[u8]) -> Result<Node> {
    let kids = children(xml, node)?;
    child(xml, &kids, local)
        .cloned()
        .ok_or_else(|| invalid("the table frame lacks its graphic table"))
}

/// Rewrite a `p:graphicFrame` holding an `a:tbl` for one cell.
/// A shape's `a:p` elements, located once (record 0346), and the
/// paragraphs a batch has rewritten so far.
struct ShapeParagraphs {
    /// The paragraph spans in the shape element.
    spans: Vec<Range<usize>>,
    /// The rewritten paragraphs, by position.
    rewritten: std::collections::HashMap<usize, Vec<u8>>,
}

impl ShapeParagraphs {
    fn locate(shape: &[u8]) -> Result<Self> {
        let node = root(shape)?;
        let kids = children(shape, &node)?;
        let body = child(shape, &kids, b"txBody").ok_or(Error::refused(
            litchi_core::edit::RefusalKind::UnsupportedTarget,
            "set_text_properties",
            "the shape has no text body",
        ))?;
        let spans = children(shape, body)?
            .into_iter()
            .filter(|kid| kid.local(shape) == b"p")
            .map(|paragraph| paragraph.span)
            .collect();
        Ok(Self {
            spans,
            rewritten: std::collections::HashMap::new(),
        })
    }

    /// The current bytes of the paragraph at `index`.
    fn paragraph<'a>(&'a self, shape: &'a [u8], index: usize) -> Result<&'a [u8]> {
        let span = self.spans.get(index).ok_or(Error::IndexOutOfBounds {
            index,
            len: self.spans.len(),
        })?;
        Ok(self
            .rewritten
            .get(&index)
            .map_or(&shape[span.clone()], Vec::as_slice))
    }

    /// The shape with every rewritten paragraph spliced in.
    fn splice(&self, shape: &[u8]) -> Result<Vec<u8>> {
        let mut splices = Splices::new();
        for (&index, bytes) in &self.rewritten {
            splices.replace(self.spans[index].clone(), bytes.clone());
        }
        splices.apply(shape)
    }
}

/// A table frame's `a:tc` elements, located once (record 0346), and the
/// cells a batch has rewritten so far.
struct TableCells {
    /// The cell spans in the frame, by row and position in the row.
    spans: Vec<Vec<Range<usize>>>,
    /// The rewritten cells, by row and position.
    rewritten: std::collections::HashMap<(usize, usize), Vec<u8>>,
}

impl TableCells {
    fn locate(frame: &[u8]) -> Result<Self> {
        let node = root(frame)?;
        let graphic = descend(frame, &node, b"graphic")?;
        let data = descend(frame, &graphic, b"graphicData")?;
        let table = descend(frame, &data, b"tbl")?;
        let mut spans = Vec::new();
        for tr in children(frame, &table)?
            .into_iter()
            .filter(|kid| kid.local(frame) == b"tr")
        {
            spans.push(
                children(frame, &tr)?
                    .into_iter()
                    .filter(|kid| kid.local(frame) == b"tc")
                    .map(|tc| tc.span)
                    .collect(),
            );
        }
        Ok(Self {
            spans,
            rewritten: std::collections::HashMap::new(),
        })
    }

    /// The current bytes of the cell at `row`, `column`.
    fn cell<'a>(&'a self, frame: &'a [u8], row: usize, column: usize) -> Result<&'a [u8]> {
        let cells = self.spans.get(row).ok_or(Error::IndexOutOfBounds {
            index: row,
            len: self.spans.len(),
        })?;
        let span = cells.get(column).ok_or(Error::IndexOutOfBounds {
            index: column,
            len: cells.len(),
        })?;
        Ok(self
            .rewritten
            .get(&(row, column))
            .map_or(&frame[span.clone()], Vec::as_slice))
    }

    /// The frame with every rewritten cell spliced in.
    fn splice(&self, frame: &[u8]) -> Result<Vec<u8>> {
        let mut splices = Splices::new();
        for (&(row, column), bytes) in &self.rewritten {
            splices.replace(self.spans[row][column].clone(), bytes.clone());
        }
        splices.apply(frame)
    }
}

fn rewrite_cell_frame(
    frame: &[u8],
    row: usize,
    column: usize,
    rewrite: impl FnOnce(&[u8]) -> Result<Vec<u8>>,
) -> Result<Vec<u8>> {
    let node = root(frame)?;
    let graphic = descend(frame, &node, b"graphic")?;
    let data = descend(frame, &graphic, b"graphicData")?;
    let table = descend(frame, &data, b"tbl")?;
    let rows: Vec<Node> = children(frame, &table)?
        .into_iter()
        .filter(|kid| kid.local(frame) == b"tr")
        .collect();
    let tr = rows.get(row).ok_or(Error::IndexOutOfBounds {
        index: row,
        len: rows.len(),
    })?;
    let cells: Vec<Node> = children(frame, tr)?
        .into_iter()
        .filter(|kid| kid.local(frame) == b"tc")
        .collect();
    let tc = cells.get(column).ok_or(Error::IndexOutOfBounds {
        index: column,
        len: cells.len(),
    })?;
    let replaced = rewrite(&frame[tc.span.clone()])?;
    let mut splices = Splices::new();
    splices.replace(tc.span.clone(), replaced);
    splices.apply(frame)
}

/// Rewrite one run of a cell's text body, as a shape's run is rewritten.
fn rewrite_cell_run(
    cell: &[u8],
    paragraph: usize,
    run: usize,
    span: Option<Range<usize>>,
    properties: &TextProperties<'_>,
) -> Result<Vec<u8>> {
    let node = root(cell)?;
    if node.empty {
        return Err(invalid("an empty a:tc has no runs"));
    }
    let kids = children(cell, &node)?;
    let body = child(cell, &kids, b"txBody").ok_or(Error::refused(
        litchi_core::edit::RefusalKind::UnsupportedTarget,
        "set_text_properties",
        "the table cell has no text body",
    ))?;
    let rewritten = rewrite_body(
        &cell[body.span.clone()],
        Some(paragraph),
        Some(run),
        span,
        properties,
    )?;
    let mut splices = Splices::new();
    splices.replace(body.span.clone(), rewritten);
    splices.apply(cell)
}

/// Rewrite one paragraph of a cell's text body, as a shape's paragraph is
/// rewritten (record 0333).
fn rewrite_cell_paragraph(
    cell: &[u8],
    paragraph: usize,
    properties: &TextProperties<'_>,
) -> Result<Vec<u8>> {
    let node = root(cell)?;
    if node.empty {
        return Err(invalid("an empty a:tc has no paragraphs"));
    }
    let kids = children(cell, &node)?;
    let body = child(cell, &kids, b"txBody").ok_or(Error::refused(
        litchi_core::edit::RefusalKind::UnsupportedTarget,
        "set_text_properties",
        "the table cell has no text body",
    ))?;
    let rewritten = rewrite_body(
        &cell[body.span.clone()],
        Some(paragraph),
        None,
        None,
        properties,
    )?;
    let mut splices = Splices::new();
    splices.replace(body.span.clone(), rewritten);
    splices.apply(cell)
}

fn rewrite_cell(cell: &[u8], properties: &TextProperties<'_>) -> Result<Vec<u8>> {
    let node = root(cell)?;
    if node.empty {
        return Err(invalid("an empty a:tc cannot hold text properties"));
    }
    let prefix = node.prefix(cell).to_vec();
    let kids = children(cell, &node)?;
    let mut splices = Splices::new();
    let body = child(cell, &kids, b"txBody");
    let formatting = properties.has_run_properties() || properties.align.is_some();
    if properties.text.is_some() || formatting {
        let body = body.ok_or(Error::refused(
            litchi_core::edit::RefusalKind::UnsupportedTarget,
            "set_text_properties",
            "the table cell has no text body",
        ))?;
        let element = &cell[body.span.clone()];
        let rewritten = match properties.text {
            Some(text) => rewrite_body_text(element, &prefix, text, properties)?,
            None => rewrite_body(element, None, None, None, properties)?,
        };
        splices.replace(body.span.clone(), rewritten);
    }
    if properties.has_cell_properties() {
        let mut changes = Vec::new();
        if let Some(fill) = properties.fill {
            changes.push(ChildChange {
                rank: 7,
                bytes: fill.map(|fill| cell_fill(&prefix, fill)),
            });
        }
        let margins: Vec<Option<String>> = properties
            .margins
            .iter()
            .map(|margin| margin.flatten().map(|value| value.to_string()))
            .collect();
        let mut attributes: Vec<(&[u8], Option<&[u8]>)> = Vec::new();
        for (index, margin) in CellMargin::ALL.into_iter().enumerate() {
            if properties.margins[index].is_some() {
                attributes.push((
                    margin.attribute(),
                    margins[index].as_deref().map(str::as_bytes),
                ));
            }
        }
        if let Some(anchor) = properties.anchor {
            attributes.push((b"anchor", anchor.map(|anchor| anchor.token().as_bytes())));
        }
        let creates = changes.iter().any(|change| change.bytes.is_some())
            || attributes.iter().any(|(_, value)| value.is_some());
        match child(cell, &kids, b"tcPr") {
            Some(tcpr) => {
                let rewritten =
                    rewrite_properties(&cell[tcpr.span.clone()], &attributes, &changes, tcpr_rank)?;
                splices.replace(tcpr.span.clone(), rewritten);
            },
            None if creates => {
                let mut empty = b"<".to_vec();
                empty.extend_from_slice(&qualified(&prefix, "tcPr"));
                empty.extend_from_slice(b"/>");
                let created = rewrite_properties(&empty, &attributes, &changes, tcpr_rank)?;
                // tcPr follows txBody and precedes extLst.
                let at =
                    child(cell, &kids, b"extLst").map_or(node.close_start, |ext| ext.span.start);
                splices.insert(at, created);
            },
            None => {},
        }
    }
    splices.apply(cell)
}

/// Replace a text body's paragraphs with one paragraph per line, keeping the
/// first paragraph's `a:pPr`, its first run's `a:rPr` and its
/// `a:endParaRPr`.
fn rewrite_body_text(
    body: &[u8],
    prefix: &[u8],
    text: &str,
    properties: &TextProperties<'_>,
) -> Result<Vec<u8>> {
    let node = root(body)?;
    let kids = children(body, &node)?;
    let paragraphs: Vec<&Node> = kids.iter().filter(|kid| kid.local(body) == b"p").collect();
    if let (Some(first), Some(last)) = (paragraphs.first(), paragraphs.last()) {
        let replaced = first.span.start..last.span.end;
        for kid in kids.iter().filter(|kid| replaced.contains(&kid.span.start)) {
            if kid.local(body) != b"p" {
                return Err(dropped_content());
            }
            let paragraph = &body[kid.span.clone()];
            let local = root(paragraph)?;
            let inner = children(paragraph, &local)?;
            let all = 0..paragraph.len();
            check_replaceable(paragraph, &inner, &all)?;
        }
    }
    let (ppr, rpr, end) = match paragraphs.first() {
        Some(first) => {
            let paragraph = &body[first.span.clone()];
            let local = root(paragraph)?;
            let first_kids = children(paragraph, &local)?;
            let ppr = child(paragraph, &first_kids, b"pPr")
                .map_or(&[][..], |ppr| &paragraph[ppr.span.clone()])
                .to_vec();
            let rpr = first_run_properties(paragraph, &first_kids)?.to_vec();
            let end = child(paragraph, &first_kids, b"endParaRPr")
                .map_or(&[][..], |end| &paragraph[end.span.clone()])
                .to_vec();
            (ppr, rpr, end)
        },
        None => (Vec::new(), Vec::new(), Vec::new()),
    };
    let p = qualified(prefix, "p");
    let mut generated = Vec::new();
    for line in text.split('\n') {
        let line = line.strip_suffix('\r').unwrap_or(line);
        let mut paragraph = b"<".to_vec();
        paragraph.extend_from_slice(&p);
        paragraph.push(b'>');
        paragraph.extend_from_slice(&ppr);
        if !line.is_empty() {
            paragraph.extend_from_slice(&new_run(prefix, &rpr, line)?);
        }
        paragraph.extend_from_slice(&end);
        paragraph.extend_from_slice(b"</");
        paragraph.extend_from_slice(&p);
        paragraph.push(b'>');
        let paragraph = if properties.has_run_properties() || properties.align.is_some() {
            let formatting = TextProperties {
                text: None,
                fill: None,
                anchor: None,
                margins: [None; 4],
                ..*properties
            };
            rewrite_paragraph(&paragraph, None, None, &formatting)?
        } else {
            paragraph
        };
        generated.extend_from_slice(&paragraph);
    }
    let mut splices = Splices::new();
    match (paragraphs.first(), paragraphs.last()) {
        (Some(first), Some(last)) => splices.replace(first.span.start..last.span.end, generated),
        _ => {
            if node.empty {
                return Err(invalid("an empty cell text body cannot take paragraphs"));
            }
            splices.insert(node.close_start, generated);
        },
    }
    splices.apply(body)
}

#[path = "text_structure.rs"]
mod structure;
pub use structure::TextBody;
pub(crate) use structure::{Change as TextChange, stage_text_structure_with_limits};
