//! Typed direct formatting of runs, paragraphs, tables and cells.
//!
//! [`Run`], [`Paragraph`], [`Table`] and [`Cell`] are what one element's own
//! property container states: the `w:rPr` of a run, the `w:pPr` of a
//! paragraph, the `w:tblPr` and `w:tblGrid` of a table, and the `w:tcPr` of
//! a cell. They are read by [`crate::Run::direct_properties`],
//! [`crate::Paragraph::direct_properties`], [`crate::Table::properties`] and
//! [`crate::Cell::properties`].
//!
//! # Local, not effective
//!
//! The values are **local**, as ADR 0004 keeps layers apart: exactly what
//! the element authors. Nothing is inherited from styles, document defaults
//! or a table style, and theme colours and fonts stay references with their
//! transforms. A property the element does not state is `None`; an on/off
//! property written as `<w:b/>` or `<w:b w:val="1"/>` is `Some(true)` and an
//! explicit `<w:b w:val="0"/>` is `Some(false)`.
//!
//! # What is read
//!
//! Only WordprocessingML (transitional or strict) children of the element's
//! own container are read, at their schema depth. So a run never reports the
//! previous formatting kept in its `w:rPrChange`, nor the formatting of the
//! runs of a text box or phonetic guide it contains; a paragraph never
//! reports its `w:pPrChange` or section break; and a cell never reports the
//! `w:tcPr` of a cell in a table nested inside it. Elements of other
//! namespaces (Word 2010 `w14:` effects, markup-compatibility wrappers) are
//! skipped.
//!
//! Retained XML views resolve namespace bindings inherited from their owner.
//! Standalone fragments retain the established undeclared-root-prefix fallback.
//! Duplicate attributes, malformed attribute references, DOCTYPE and content
//! outside the fragment root are refused with content-free parser diagnostics.
//!
//! # Malformed values
//!
//! A property whose value is outside its schema domain (an unknown token, a
//! number that does not parse or overflows, a missing required attribute) is
//! left unset rather than guessed, and its element name is listed by the
//! value's `invalid_properties()`, so one bad property never hides the
//! others. A repeated property element is read from its first occurrence,
//! except `w:ind`, `w:rFonts` and `w:lang`, whose attributes are taken from
//! the first occurrence that states each, as Word merges them. XML that is
//! not well-formed, a root that is not the expected WordprocessingML element,
//! and a property container that appears twice are errors.
//!
//! # Values
//!
//! Lengths keep the integer unit WordprocessingML stores them in, so a value
//! read from a file never passes through floating point: [`Twips`] for
//! indents, spacing, widths and tab stops, [`HalfPoints`] for font sizes,
//! kerning and baseline positions, [`EighthPoints`] for border widths and
//! [`Points`] for border spacing. The schema also admits decimals and
//! universal measures (`"1in"`, `"2.5cm"`, `"12pt"`); those are converted to
//! the stored unit exactly and then rounded half away from zero to a whole
//! unit, the precision Word itself keeps. Closed vocabularies are enums (the
//! 193 border styles are a checked [`BorderStyle`] token), and colours keep
//! their theme references and transforms rather than collapsing to RGB.
//!
//! # `OfficeCLI` property names
//!
//! The typed fields cover the local properties `OfficeCLI` reads back for
//! these elements (`schemas/help/docx/{run,paragraph,table,table-cell}.json`),
//! so a node view can serve those keys from one parse:
//!
//! | Element | `OfficeCLI` keys | Typed source |
//! |---|---|---|
//! | run | `bold`, `italic`, `strike`, `dstrike`, `caps`, `smallcaps`, `vanish`, `outline`, `shadow`, `emboss`, `imprint`, `noproof`, `specVanish`, `snapToGrid` | the [`Run`] field of the same meaning |
//! | run | `bold.cs`, `italic.cs`, `size`, `size.cs` | [`Run::bold_complex_script`], [`Run::italic_complex_script`], [`Run::size`], [`Run::size_complex_script`] |
//! | run | `font.latin`, `font.ascii`, `font.hAnsi`, `font.ea`, `font.cs`, `font.hint`, `font.*Theme` | [`Run::fonts`] ([`Fonts`]) |
//! | run | `color`, `underline`, `underline.color`, `highlight` | [`Run::color`], [`Run::underline`], [`Run::highlight`] |
//! | run | `superscript`, `subscript`, `direction`, `position`, `charSpacing`, `kern` | [`Run::vertical_align`], [`Run::right_to_left`], [`Run::position`], [`Run::character_spacing`], [`Run::kerning`] |
//! | run | `lang.*`, `fill`, `shading.*`, `bdr`, `em`, `rStyle` | [`Run::language`], [`Run::shading`], [`Run::border`], [`Run::emphasis`], [`Run::style`] |
//! | paragraph | `style`, `align`, `outlineLvl`, `numId`, `numLevel` | [`Paragraph::style`], [`Paragraph::justification`], [`Paragraph::outline_level`], [`Paragraph::numbering`] |
//! | paragraph | `spaceBefore`, `spaceAfter`, `lineSpacing`, `lineRule`, `space*Lines`, `space*Auto` | [`Paragraph::spacing`] |
//! | paragraph | `indent`, `rightIndent`, `firstLineIndent`, `hangingIndent`, `*Chars` | [`Paragraph::indentation`] ([`Indentation`]) |
//! | paragraph | `keepNext`, `keepLines`, `pageBreakBefore`, `widowControl`, `contextualSpacing`, `direction` | [`Paragraph::keep_next`], [`Paragraph::keep_lines`], [`Paragraph::page_break_before`], [`Paragraph::widow_control`], [`Paragraph::contextual_spacing`], [`Paragraph::bidi`] |
//! | paragraph | `fill`, `shading.*`, `pbdr.*`, `tabs`, `markRPr.*` | [`Paragraph::shading`], [`Paragraph::borders`], [`Paragraph::tabs`], [`Paragraph::mark`] |
//! | table | `style`, `width`, `align`, `indent`, `cellSpacing`, `layout`, `direction` | [`Table::style`], [`Table::width`], [`Table::justification`], [`Table::indent`], [`Table::cell_spacing`], [`Table::layout`], [`Table::bidi_visual`] |
//! | table | `padding.*`, `shading.*`, `border.*`, `tblLook`, `caption`, `description`, `rowBandSize`, `colBandSize`, `colWidths` | [`Table::cell_margins`], [`Table::shading`], [`Table::borders`], [`Table::look`], [`Table::caption`], [`Table::description`], [`Table::row_band_size`], [`Table::column_band_size`], [`Table::grid`] |
//! | cell | `width`, `colspan`, `vmerge`, `hmerge`, `valign`, `textDirection` | [`Cell::width`], [`Cell::grid_span`], [`Cell::vertical_merge`], [`Cell::horizontal_merge`], [`Cell::vertical_align`], [`Cell::text_direction`] |
//! | cell | `padding.*`, `nowrap`, `hideMark`, `tcFitText`, `fill`, `shading.*`, `border.*` | [`Cell::margins`], [`Cell::no_wrap`], [`Cell::hide_mark`], [`Cell::fit_text`], [`Cell::shading`], [`Cell::borders`] |
//!
//! Not covered: `OfficeCLI`'s read-only `effective.*` keys (they need style
//! resolution), Word 2010 text effects and OpenType features (read by
//! [`crate::Run::effects`] and [`crate::Run::open_type`]), East Asian layout,
//! run `w:fitText`, paragraph frames, conditional-formatting stamps
//! (`w:cnfStyle`), the paragraph `w:divId` (read by
//! [`crate::Paragraph::division_id`]), floating-table positioning, and
//! revision attribution.

mod edit;
mod paragraph;
mod run;
mod scan;
mod table;
mod value;

#[cfg(test)]
mod corpus_tests;
#[cfg(test)]
mod tests;

use std::borrow::Cow;

use quick_xml::events::BytesStart;

pub use paragraph::{
    Indentation, Justification, Numbering, Paragraph, ParagraphBorders, TabKind, TabLeader,
    TabStop, TextAlignment, TextDirection, TextboxTightWrap,
};
pub use run::Run;
pub use table::{
    Cell, CellBorders, Height, HeightRule, Margins, Merge, Row, Table, TableBorders,
    TableJustification, TableLayout, TableLook, VerticalJc, Width,
};
pub use value::{
    Border, BorderStyle, Color, EighthPoints, Emphasis, FontHint, Fonts, HalfPoints, HexColor,
    Highlight, Language, Points, Shading, ShadingPattern, TextEffect, ThemeFont, Twips,
    VerticalAlign,
};

#[cfg(test)]
pub(crate) use edit::paragraph_rank_for_tests;
pub(crate) use edit::{
    SpliceError, admit_paragraph, admit_run, paragraph_properties_element, run_properties_element,
    same_paragraph, same_run, splice_paragraph, splice_row, splice_run, splice_table,
};
pub(crate) use paragraph::parse as parse_paragraph;
pub(crate) use run::parse as parse_run;
pub(crate) use table::{parse_cell, parse_row, parse_table};

use crate::error::Result;
use scan::Context;
pub(crate) use scan::audit_attributes;
use value::{hex_byte, hex_color, measure, on_off, theme_color};

/// The index of `name` in a property table, checked at compile time.
const fn position(table: &[&str], name: &str) -> usize {
    let wanted = name.as_bytes();
    let mut index = 0;
    while index < table.len() {
        let candidate = table[index].as_bytes();
        if candidate.len() == wanted.len() {
            let mut byte = 0;
            while byte < wanted.len() && candidate[byte] == wanted[byte] {
                byte += 1;
            }
            if byte == wanted.len() {
                return index;
            }
        }
        index += 1;
    }
    panic!("the property name is not in its table");
}

/// The properties of one element that were stated with an invalid value, as
/// bits of the element's property table (at most 64 entries).
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Hash)]
pub(crate) struct Invalid(u64);

impl Invalid {
    fn insert(&mut self, index: usize) {
        self.0 |= 1u64 << index;
    }

    fn names(self, table: &'static [&'static str]) -> impl Iterator<Item = &'static str> {
        table
            .iter()
            .enumerate()
            .filter(move |(index, _)| self.0 & (1u64 << index) != 0)
            .map(|(_, name)| *name)
    }
}

/// Collect the WordprocessingML attributes `names` of `element`, by local
/// name.
fn collect<'e, const N: usize>(
    context: &Context<'_>,
    element: &'e BytesStart<'_>,
    names: [&[u8]; N],
) -> Result<[Option<Cow<'e, str>>; N]> {
    let mut values: [Option<Cow<'e, str>>; N] = [const { None }; N];
    context.attributes(element, |local, value| {
        if let Some(slot) = names.iter().position(|name| *name == local)
            && values[slot].is_none()
        {
            values[slot] = Some(value);
        }
    })?;
    Ok(values)
}

/// An optional attribute: `Some(None)` when absent, `Some(Some(value))` when
/// valid, and `None` when present but invalid.
fn optional<T>(raw: Option<&str>, parse: impl Fn(&str) -> Option<T>) -> Option<Option<T>> {
    match raw {
        None => Some(None),
        Some(raw) => parse(raw).map(Some),
    }
}

/// `CT_Color`; `None` when `w:val` is missing or any attribute is invalid.
fn parse_color(context: &Context<'_>, element: &BytesStart<'_>) -> Result<Option<Color>> {
    let [value, theme, tint, shade] = collect(
        context,
        element,
        [b"val", b"themeColor", b"themeTint", b"themeShade"],
    )?;
    Ok((|| {
        Some(Color {
            value: hex_color(value.as_deref()?)?,
            theme: optional(theme.as_deref(), theme_color)?,
            theme_tint: optional(tint.as_deref(), hex_byte)?,
            theme_shade: optional(shade.as_deref(), hex_byte)?,
        })
    })())
}

/// `CT_Shd`; `None` when `w:val` is missing or any attribute is invalid.
fn parse_shading(context: &Context<'_>, element: &BytesStart<'_>) -> Result<Option<Shading>> {
    let [
        pattern,
        color,
        theme_color_raw,
        theme_tint,
        theme_shade,
        fill,
        theme_fill,
        theme_fill_tint,
        theme_fill_shade,
    ] = collect(
        context,
        element,
        [
            b"val",
            b"color",
            b"themeColor",
            b"themeTint",
            b"themeShade",
            b"fill",
            b"themeFill",
            b"themeFillTint",
            b"themeFillShade",
        ],
    )?;
    Ok((|| {
        Some(Shading {
            pattern: ShadingPattern::from_xml(pattern.as_deref()?)?,
            color: optional(color.as_deref(), hex_color)?,
            theme_color: optional(theme_color_raw.as_deref(), theme_color)?,
            theme_tint: optional(theme_tint.as_deref(), hex_byte)?,
            theme_shade: optional(theme_shade.as_deref(), hex_byte)?,
            fill: optional(fill.as_deref(), hex_color)?,
            theme_fill: optional(theme_fill.as_deref(), theme_color)?,
            theme_fill_tint: optional(theme_fill_tint.as_deref(), hex_byte)?,
            theme_fill_shade: optional(theme_fill_shade.as_deref(), hex_byte)?,
        })
    })())
}

/// `CT_Border`; `None` when `w:val` is missing or any attribute is invalid.
fn parse_border(context: &Context<'_>, element: &BytesStart<'_>) -> Result<Option<Border>> {
    let [style, color, theme, tint, shade, size, space, shadow, frame] = collect(
        context,
        element,
        [
            b"val",
            b"color",
            b"themeColor",
            b"themeTint",
            b"themeShade",
            b"sz",
            b"space",
            b"shadow",
            b"frame",
        ],
    )?;
    Ok((|| {
        Some(Border {
            style: BorderStyle::from_xml(style.as_deref()?)?,
            color: optional(color.as_deref(), hex_color)?,
            theme_color: optional(theme.as_deref(), theme_color)?,
            theme_tint: optional(tint.as_deref(), hex_byte)?,
            theme_shade: optional(shade.as_deref(), hex_byte)?,
            size: optional(size.as_deref(), |raw| {
                measure(raw, value::Unit::EighthPoints, value::Sign::Unsigned)
            })?
            .map(EighthPoints),
            space: optional(space.as_deref(), |raw| {
                measure(raw, value::Unit::Points, value::Sign::Unsigned)
            })?
            .map(Points),
            shadow: optional(shadow.as_deref(), on_off)?,
            frame: optional(frame.as_deref(), on_off)?,
        })
    })())
}

impl crate::paragraph::Run {
    /// Read this run's direct (local) properties, the children of its own
    /// `w:rPr`, as typed values.
    ///
    /// Unlike the single-property readers such as [`Self::bold`], this reads
    /// every property in one pass, resolves names by namespace, and never
    /// reports a property it cannot type: see the [`crate::properties`]
    /// module for the rules.
    ///
    /// # Errors
    ///
    /// Returns an error for XML that is not well-formed, a root that is not
    /// a WordprocessingML `w:r`, a second `w:rPr`, or an execution limit of
    /// a source-backed run.
    pub fn direct_properties(&self) -> Result<Run> {
        let _parser_admission = self.parser_admission()?;
        run::parse(self.xml_bytes(), &self.inherited_namespaces()?)
    }
}
