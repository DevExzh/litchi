//! One-pass reader of the format-bearing collections of `xl/styles.xml`.
//!
//! The reader keeps exactly what [`super::Format`] exposes: declared number
//! formats, fonts, fills, borders and the `cellXfs` records. Collections it
//! does not expose (`cellStyleXfs`, `cellStyles`, `dxfs`, colours, extension
//! lists) are skipped without interpretation.
//!
//! The `cellXfs` structure follows the shared-style catalog
//! (`crate::raw::styles`) exactly, so whenever both readers accept a
//! stylesheet they count the same records: one `SpreadsheetML` `styleSheet`
//! root, `cellXfs` only as its direct child and at most once, `xf` only as a
//! direct child of `cellXfs`, a non-empty collection of at most 65,430
//! records whose optional `count` matches. Both readers apply the same
//! markup-compatibility preprocessing first.
//!
//! This reader refuses more than the catalog does, never less: values of
//! recognised properties are checked against their schema domain, and a
//! second `numFmts`, `fonts`, `fills` or `borders` collection, whose record
//! indexes would be ambiguous, is refused, as Excel itself refuses such a
//! stylesheet. Unrecognised attributes and elements, and every element
//! outside the `SpreadsheetML` namespaces, are ignored so that extension
//! markup does not make a stylesheet unreadable. The declared `count` of the
//! other collections is not enforced: their records are authoritative.

use std::borrow::Cow;

use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::attributes::Attribute;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

use litchi_ooxml_common::xml::attributes::BytesStartExt as _;

use super::{
    Alignment, Border, CellXf, Color, Dir, Fill, Font, FormatTable, GradientFill, GradientKind,
    GradientStop, Horizontal, Indent, Line, MAX_RECORDS, Pattern, PatternFill, Protection, Reading,
    Rgb, Rotation, Scheme, Script, Side, Tint, Underline, Vertical,
};
use crate::error::{Error, Result, allocation, invalid};
use crate::raw::namespace::{
    SPREADSHEETML_NAMESPACE, STRICT_SPREADSHEETML_NAMESPACE, is_spreadsheetml_name,
};

/// [MS-OE376] 2.1.728 limits `cellXfs` to 65,430 records; the shared-style
/// catalog enforces the same bound.
const MAX_CELL_FORMATS: u32 = 65_430;
/// Element nesting the reader follows before refusing the stylesheet.
const MAX_DEPTH: usize = 256;

/// Which part of the stylesheet an open element belongs to.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Frame {
    Root,
    NumberFormats,
    Fonts,
    Font,
    Fills,
    Fill,
    PatternFill,
    GradientFill,
    GradientStop,
    Borders,
    Border,
    BorderSide(SideSlot),
    CellFormats,
    CellFormat,
    /// An element whose content is not read.
    Ignored,
}

/// One `CT_Border` edge.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum SideSlot {
    Start,
    End,
    Left,
    Right,
    Top,
    Bottom,
    Diagonal,
    Vertical,
    Horizontal,
}

impl SideSlot {
    fn from_local_name(name: &[u8]) -> Option<Self> {
        Some(match name {
            b"start" => Self::Start,
            b"end" => Self::End,
            b"left" => Self::Left,
            b"right" => Self::Right,
            b"top" => Self::Top,
            b"bottom" => Self::Bottom,
            b"diagonal" => Self::Diagonal,
            b"vertical" => Self::Vertical,
            b"horizontal" => Self::Horizontal,
            _ => return None,
        })
    }

    const fn bit(self) -> u16 {
        1 << self as u16
    }

    const fn description(self) -> &'static str {
        match self {
            Self::Start => "start border",
            Self::End => "end border",
            Self::Left => "left border",
            Self::Right => "right border",
            Self::Top => "top border",
            Self::Bottom => "bottom border",
            Self::Diagonal => "diagonal border",
            Self::Vertical => "vertical border",
            Self::Horizontal => "horizontal border",
        }
    }
}

const NUMBER_FORMATS: u8 = 1;
const FONTS: u8 = 2;
const FILLS: u8 = 4;
const BORDERS: u8 = 8;
const CELL_FORMATS: u8 = 16;

#[derive(Default)]
struct PendingFont {
    font: Font,
    seen: u16,
}

#[derive(Default)]
struct PendingFill {
    value: Option<Fill>,
}

#[derive(Default)]
struct PendingPattern {
    fill: PatternFill,
    seen_foreground: bool,
    seen_background: bool,
}

struct PendingGradient {
    fill: GradientFill,
    stops: Vec<GradientStop>,
}

struct PendingStop {
    position: f64,
    color: Option<Color>,
}

#[derive(Default)]
struct PendingBorder {
    border: Border,
    seen: u16,
}

struct PendingSide {
    slot: SideSlot,
    line: Option<Line>,
    color: Option<Color>,
}

#[derive(Default)]
struct PendingCellFormat {
    record: CellXf,
    seen_alignment: bool,
    seen_protection: bool,
}

#[derive(Default)]
struct Reader {
    number_formats: Vec<(u32, Box<str>)>,
    fonts: Vec<Font>,
    fills: Vec<Fill>,
    borders: Vec<Border>,
    cell_formats: Vec<CellXf>,
    sections: u8,
    declared_cell_formats: Option<u32>,
    cell_format_count: u32,
    font: Option<PendingFont>,
    fill: Option<PendingFill>,
    pattern: Option<PendingPattern>,
    gradient: Option<PendingGradient>,
    stop: Option<PendingStop>,
    border: Option<PendingBorder>,
    side: Option<PendingSide>,
    cell_format: Option<PendingCellFormat>,
}

/// Parse a `styles.xml` payload into a format table.
pub(super) fn parse(content: &[u8]) -> Result<FormatTable> {
    let processed = litchi_ooxml_common::mce::process_ooxml(content)?;
    let mut reader = NsReader::from_reader(processed.as_ref());
    let mut state = Reader::default();
    let mut stack: Vec<Frame> = Vec::new();
    let mut closed_root = false;

    loop {
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|_error| invalid("styles XML is malformed"))?;
        // As in the shared-style catalog: events borrow the processed input
        // and the reader's resolver answers every event of this iteration,
        // so no event or resolver is copied.
        let (namespace, event) = reader.resolver().resolve_event(event);
        match event {
            Event::Start(element) => {
                let frame = state.start(
                    stack.last().copied(),
                    closed_root,
                    &namespace,
                    &element,
                    decoder,
                )?;
                if stack.len() >= MAX_DEPTH {
                    return Err(invalid(format!("styles XML exceeds {MAX_DEPTH} levels")));
                }
                stack
                    .try_reserve(1)
                    .map_err(|source| allocation("styles XML element stack", source))?;
                stack.push(frame);
            },
            Event::Empty(element) => {
                let frame = state.start(
                    stack.last().copied(),
                    closed_root,
                    &namespace,
                    &element,
                    decoder,
                )?;
                state.end(frame)?;
                closed_root |= frame == Frame::Root;
            },
            Event::End(_) => {
                let frame = stack
                    .pop()
                    .ok_or_else(|| invalid("styles XML has a closing element outside its root"))?;
                state.end(frame)?;
                closed_root |= frame == Frame::Root;
            },
            Event::Eof => {
                if !closed_root || !stack.is_empty() {
                    return Err(invalid(
                        "styles XML has a missing or unterminated SpreadsheetML styleSheet root",
                    ));
                }
                break;
            },
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
    }
    state.finish()
}

impl Reader {
    fn start(
        &mut self,
        parent: Option<Frame>,
        closed_root: bool,
        namespace: &ResolveResult<'_>,
        element: &BytesStart<'_>,
        decoder: Decoder,
    ) -> Result<Frame> {
        let named = |local: &[u8]| is_spreadsheetml_name(namespace, element.name(), local);
        let Some(parent) = parent else {
            if closed_root || !named(b"styleSheet") {
                return Err(invalid(
                    "styles XML must have one SpreadsheetML styleSheet root",
                ));
            }
            return Ok(Frame::Root);
        };
        match parent {
            Frame::Root => self.start_section(element, decoder, named),
            Frame::NumberFormats if named(b"numFmt") => {
                self.number_format(element, decoder)?;
                Ok(Frame::Ignored)
            },
            Frame::Fonts if named(b"font") => {
                self.font = Some(PendingFont::default());
                Ok(Frame::Font)
            },
            Frame::Font if is_spreadsheetml(namespace) => {
                self.font_property(element, decoder)?;
                Ok(Frame::Ignored)
            },
            Frame::Fills if named(b"fill") => {
                self.fill = Some(PendingFill::default());
                Ok(Frame::Fill)
            },
            Frame::Fill if named(b"patternFill") => {
                self.start_pattern(element, decoder)?;
                Ok(Frame::PatternFill)
            },
            Frame::Fill if named(b"gradientFill") => {
                self.start_gradient(element, decoder)?;
                Ok(Frame::GradientFill)
            },
            Frame::PatternFill if named(b"fgColor") || named(b"bgColor") => {
                self.pattern_color(named(b"fgColor"), element, decoder)?;
                Ok(Frame::Ignored)
            },
            Frame::GradientFill if named(b"stop") => {
                self.start_stop(element, decoder)?;
                Ok(Frame::GradientStop)
            },
            Frame::GradientStop if named(b"color") => {
                let stop = self.stop.as_mut().ok_or_else(missing_record)?;
                if stop.color.is_some() {
                    return Err(invalid("duplicate gradient-stop color"));
                }
                stop.color = Some(color(element, decoder)?);
                Ok(Frame::Ignored)
            },
            Frame::Borders if named(b"border") => {
                self.start_border(element, decoder)?;
                Ok(Frame::Border)
            },
            Frame::Border if is_spreadsheetml(namespace) => {
                match SideSlot::from_local_name(element.local_name().as_ref()) {
                    Some(slot) => {
                        self.start_side(slot, element, decoder)?;
                        Ok(Frame::BorderSide(slot))
                    },
                    None => Ok(Frame::Ignored),
                }
            },
            Frame::BorderSide(_) if named(b"color") => {
                let side = self.side.as_mut().ok_or_else(missing_record)?;
                if side.color.is_some() {
                    return Err(invalid("duplicate border color"));
                }
                side.color = Some(color(element, decoder)?);
                Ok(Frame::Ignored)
            },
            Frame::CellFormats if named(b"xf") => {
                self.start_cell_format(element, decoder)?;
                Ok(Frame::CellFormat)
            },
            Frame::CellFormat if named(b"alignment") => {
                let pending = self.cell_format.as_mut().ok_or_else(missing_record)?;
                if pending.seen_alignment {
                    return Err(invalid("duplicate cell alignment"));
                }
                pending.seen_alignment = true;
                pending.record.alignment = Some(alignment(element, decoder)?);
                Ok(Frame::Ignored)
            },
            Frame::CellFormat if named(b"protection") => {
                let pending = self.cell_format.as_mut().ok_or_else(missing_record)?;
                if pending.seen_protection {
                    return Err(invalid("duplicate cell protection"));
                }
                pending.seen_protection = true;
                pending.record.protection = Some(protection(element, decoder)?);
                Ok(Frame::Ignored)
            },
            Frame::NumberFormats
            | Frame::Fonts
            | Frame::Font
            | Frame::Fills
            | Frame::Fill
            | Frame::PatternFill
            | Frame::GradientFill
            | Frame::GradientStop
            | Frame::Borders
            | Frame::Border
            | Frame::BorderSide(_)
            | Frame::CellFormats
            | Frame::CellFormat
            | Frame::Ignored => Ok(Frame::Ignored),
        }
    }

    fn start_section(
        &mut self,
        element: &BytesStart<'_>,
        decoder: Decoder,
        named: impl Fn(&[u8]) -> bool,
    ) -> Result<Frame> {
        let (bit, frame, name) = if named(b"numFmts") {
            (NUMBER_FORMATS, Frame::NumberFormats, "numFmts")
        } else if named(b"fonts") {
            (FONTS, Frame::Fonts, "fonts")
        } else if named(b"fills") {
            (FILLS, Frame::Fills, "fills")
        } else if named(b"borders") {
            (BORDERS, Frame::Borders, "borders")
        } else if named(b"cellXfs") {
            (CELL_FORMATS, Frame::CellFormats, "cellXfs")
        } else {
            return Ok(Frame::Ignored);
        };
        if self.sections & bit != 0 {
            return Err(invalid(format!(
                "styles XML has duplicate {name} collections"
            )));
        }
        self.sections |= bit;
        if frame == Frame::CellFormats {
            self.declared_cell_formats = cell_format_count(element, decoder)?;
        }
        Ok(frame)
    }

    fn end(&mut self, frame: Frame) -> Result<()> {
        match frame {
            Frame::Font => {
                let pending = self.font.take().ok_or_else(missing_record)?;
                push(&mut self.fonts, pending.font, "fonts")
            },
            Frame::Fill => {
                let pending = self.fill.take().ok_or_else(missing_record)?;
                let fill = pending
                    .value
                    .unwrap_or_else(|| Fill::Pattern(PatternFill::default()));
                push(&mut self.fills, fill, "fills")
            },
            Frame::PatternFill => {
                let pattern = self.pattern.take().ok_or_else(missing_record)?;
                self.set_fill(Fill::Pattern(pattern.fill))
            },
            Frame::GradientFill => {
                let pending = self.gradient.take().ok_or_else(missing_record)?;
                let mut fill = pending.fill;
                fill.stops = pending.stops.into_boxed_slice();
                self.set_fill(Fill::Gradient(fill))
            },
            Frame::GradientStop => {
                let stop = self.stop.take().ok_or_else(missing_record)?;
                let color = stop
                    .color
                    .ok_or_else(|| invalid("gradient stop is missing its color"))?;
                let gradient = self.gradient.as_mut().ok_or_else(missing_record)?;
                push(
                    &mut gradient.stops,
                    GradientStop {
                        position: stop.position,
                        color,
                    },
                    "gradient stops",
                )
            },
            Frame::Border => {
                let pending = self.border.take().ok_or_else(missing_record)?;
                push(&mut self.borders, pending.border, "borders")
            },
            Frame::BorderSide(slot) => self.finish_side(slot),
            Frame::CellFormat => {
                let pending = self.cell_format.take().ok_or_else(missing_record)?;
                push(&mut self.cell_formats, pending.record, "cell formats")
            },
            Frame::Root
            | Frame::NumberFormats
            | Frame::Fonts
            | Frame::Fills
            | Frame::Borders
            | Frame::CellFormats
            | Frame::Ignored => Ok(()),
        }
    }

    fn finish(mut self) -> Result<FormatTable> {
        self.number_formats.sort_unstable_by_key(|(id, _)| *id);
        if self
            .number_formats
            .windows(2)
            .any(|pair| pair[0].0 == pair[1].0)
        {
            return Err(invalid("duplicate number-format ID"));
        }
        if self.sections & CELL_FORMATS == 0 {
            return Err(invalid("styles XML is missing the cellXfs collection"));
        }
        if self.cell_format_count == 0 {
            return Err(invalid("styles XML cellXfs collection must not be empty"));
        }
        if let Some(declared) = self.declared_cell_formats
            && declared != self.cell_format_count
        {
            return Err(invalid(format!(
                "styles XML declares {declared} cell formats but contains {}",
                self.cell_format_count
            )));
        }
        Ok(FormatTable::from_parts(
            self.number_formats,
            self.fonts,
            self.fills,
            self.borders,
            self.cell_formats,
        ))
    }

    fn number_format(&mut self, element: &BytesStart<'_>, decoder: Decoder) -> Result<()> {
        let mut id = None;
        let mut code = None;
        for attribute in element.checked_attributes() {
            let attribute = attribute.map_err(attribute_error)?;
            match unqualified(&attribute) {
                Some(b"numFmtId") => {
                    id = Some(parse_u32(&value(&attribute, decoder)?, "number-format ID")?)
                },
                Some(b"formatCode") => {
                    code = Some(value(&attribute, decoder)?.into_owned().into_boxed_str());
                },
                _ => {},
            }
        }
        let id = id.ok_or_else(|| invalid("number format is missing its numFmtId"))?;
        let code = code.ok_or_else(|| invalid("number format is missing its formatCode"))?;
        // Duplicate identifiers are refused once the collection is sorted, in
        // `finish`, so a long collection is never scanned per record.
        push(&mut self.number_formats, (id, code), "number formats")
    }

    fn font_property(&mut self, element: &BytesStart<'_>, decoder: Decoder) -> Result<()> {
        let pending = self.font.as_mut().ok_or_else(missing_record)?;
        let font = &mut pending.font;
        let (bit, description) = match element.local_name().as_ref() {
            b"name" => {
                font.name = Some(required_val(element, decoder, "font name")?.into_boxed_str());
                (1, "font name")
            },
            b"sz" => {
                let size = parse_f64(&required_val(element, decoder, "font size")?, "font size")?;
                if size <= 0.0 {
                    return Err(invalid("invalid font size"));
                }
                font.size = Some(size);
                (1 << 1, "font size")
            },
            b"b" => {
                font.bold = Some(boolean_val(element, decoder, "bold")?);
                (1 << 2, "bold property")
            },
            b"i" => {
                font.italic = Some(boolean_val(element, decoder, "italic")?);
                (1 << 3, "italic property")
            },
            b"strike" => {
                font.strike = Some(boolean_val(element, decoder, "strike")?);
                (1 << 4, "strike property")
            },
            b"outline" => {
                font.outline = Some(boolean_val(element, decoder, "outline")?);
                (1 << 5, "outline property")
            },
            b"shadow" => {
                font.shadow = Some(boolean_val(element, decoder, "shadow")?);
                (1 << 6, "shadow property")
            },
            b"condense" => {
                font.condense = Some(boolean_val(element, decoder, "condense")?);
                (1 << 7, "condense property")
            },
            b"extend" => {
                font.extend = Some(boolean_val(element, decoder, "extend")?);
                (1 << 8, "extend property")
            },
            b"u" => {
                let token = optional_val(element, decoder)?;
                let token = token.as_deref().unwrap_or("single");
                font.underline = Some(
                    token
                        .parse::<Underline>()
                        .map_err(|_error| invalid("invalid underline token"))?,
                );
                (1 << 9, "underline property")
            },
            b"vertAlign" => {
                let token = required_val(element, decoder, "font script")?;
                font.script = Some(
                    token
                        .parse::<Script>()
                        .map_err(|_error| invalid("invalid font script token"))?,
                );
                (1 << 10, "font script property")
            },
            b"color" => {
                font.color = Some(color(element, decoder)?);
                (1 << 11, "font color")
            },
            b"family" => {
                font.family = Some(parse_u32(
                    &required_val(element, decoder, "font family")?,
                    "font family",
                )?);
                (1 << 12, "font family")
            },
            b"charset" => {
                font.charset = Some(parse_u32(
                    &required_val(element, decoder, "font charset")?,
                    "font charset",
                )?);
                (1 << 13, "font charset")
            },
            b"scheme" => {
                let token = required_val(element, decoder, "font scheme")?;
                font.scheme = Some(
                    token
                        .parse::<Scheme>()
                        .map_err(|_error| invalid("invalid font scheme token"))?,
                );
                (1 << 14, "font scheme")
            },
            _ => return Ok(()),
        };
        mark(&mut pending.seen, bit, description)
    }

    fn set_fill(&mut self, fill: Fill) -> Result<()> {
        let pending = self.fill.as_mut().ok_or_else(missing_record)?;
        if pending.value.is_some() {
            return Err(invalid("duplicate fill definition"));
        }
        pending.value = Some(fill);
        Ok(())
    }

    fn start_pattern(&mut self, element: &BytesStart<'_>, decoder: Decoder) -> Result<()> {
        let mut pattern = PendingPattern::default();
        for attribute in element.checked_attributes() {
            let attribute = attribute.map_err(attribute_error)?;
            if unqualified(&attribute) == Some(b"patternType") {
                let token = value(&attribute, decoder)?;
                pattern.fill.pattern = Some(
                    Pattern::from_token(token.trim())
                        .ok_or_else(|| invalid("invalid fill pattern"))?,
                );
            }
        }
        self.pattern = Some(pattern);
        Ok(())
    }

    fn pattern_color(
        &mut self,
        foreground: bool,
        element: &BytesStart<'_>,
        decoder: Decoder,
    ) -> Result<()> {
        let pattern = self.pattern.as_mut().ok_or_else(missing_record)?;
        let (seen, slot, description) = if foreground {
            (
                &mut pattern.seen_foreground,
                &mut pattern.fill.foreground,
                "foreground color",
            )
        } else {
            (
                &mut pattern.seen_background,
                &mut pattern.fill.background,
                "background color",
            )
        };
        if *seen {
            return Err(invalid(format!("duplicate {description}")));
        }
        *seen = true;
        *slot = Some(color(element, decoder)?);
        Ok(())
    }

    fn start_gradient(&mut self, element: &BytesStart<'_>, decoder: Decoder) -> Result<()> {
        let mut fill = GradientFill {
            kind: GradientKind::Linear,
            degree: 0.0,
            left: 0.0,
            right: 0.0,
            top: 0.0,
            bottom: 0.0,
            stops: Box::new([]),
        };
        for attribute in element.checked_attributes() {
            let attribute = attribute.map_err(attribute_error)?;
            let Some(name) = unqualified(&attribute) else {
                continue;
            };
            let target = match name {
                b"type" => {
                    let token = value(&attribute, decoder)?;
                    fill.kind = GradientKind::from_token(token.trim())
                        .ok_or_else(|| invalid("invalid gradient type"))?;
                    continue;
                },
                b"degree" => &mut fill.degree,
                b"left" => &mut fill.left,
                b"right" => &mut fill.right,
                b"top" => &mut fill.top,
                b"bottom" => &mut fill.bottom,
                _ => continue,
            };
            *target = parse_f64(&value(&attribute, decoder)?, "gradient geometry")?;
        }
        self.gradient = Some(PendingGradient {
            fill,
            stops: Vec::new(),
        });
        Ok(())
    }

    fn start_stop(&mut self, element: &BytesStart<'_>, decoder: Decoder) -> Result<()> {
        let mut position = None;
        for attribute in element.checked_attributes() {
            let attribute = attribute.map_err(attribute_error)?;
            if unqualified(&attribute) == Some(b"position") {
                position = Some(parse_f64(
                    &value(&attribute, decoder)?,
                    "gradient-stop position",
                )?);
            }
        }
        let position = position.ok_or_else(|| invalid("gradient stop is missing its position"))?;
        if !(0.0..=1.0).contains(&position) {
            return Err(invalid("invalid gradient-stop position"));
        }
        self.stop = Some(PendingStop {
            position,
            color: None,
        });
        Ok(())
    }

    fn start_border(&mut self, element: &BytesStart<'_>, decoder: Decoder) -> Result<()> {
        let mut border = Border::new();
        let mut up = false;
        let mut down = false;
        for attribute in element.checked_attributes() {
            let attribute = attribute.map_err(attribute_error)?;
            match unqualified(&attribute) {
                Some(b"diagonalUp") => {
                    up = parse_bool(&value(&attribute, decoder)?, "diagonalUp")?;
                },
                Some(b"diagonalDown") => {
                    down = parse_bool(&value(&attribute, decoder)?, "diagonalDown")?;
                },
                Some(b"outline") => {
                    border.outline =
                        Some(parse_bool(&value(&attribute, decoder)?, "border outline")?);
                },
                _ => {},
            }
        }
        border.set_diagonal_dir(Dir::from_flags(up, down));
        self.border = Some(PendingBorder { border, seen: 0 });
        Ok(())
    }

    fn start_side(
        &mut self,
        slot: SideSlot,
        element: &BytesStart<'_>,
        decoder: Decoder,
    ) -> Result<()> {
        let pending = self.border.as_mut().ok_or_else(missing_record)?;
        mark(&mut pending.seen, slot.bit(), slot.description())?;
        let mut line = None;
        for attribute in element.checked_attributes() {
            let attribute = attribute.map_err(attribute_error)?;
            if unqualified(&attribute) == Some(b"style") {
                let token = value(&attribute, decoder)?;
                line = Line::from_xml(token.trim())
                    .map_err(|_source| invalid("invalid border style"))?;
            }
        }
        self.side = Some(PendingSide {
            slot,
            line,
            color: None,
        });
        Ok(())
    }

    fn finish_side(&mut self, slot: SideSlot) -> Result<()> {
        let side = self.side.take().ok_or_else(missing_record)?;
        debug_assert_eq!(side.slot, slot, "border side frames nest");
        let value = side.line.map(|line| {
            let value = Side::new(line);
            match side.color {
                Some(color) => value.with_color(color),
                None => value,
            }
        });
        let border = &mut self.border.as_mut().ok_or_else(missing_record)?.border;
        match slot {
            SideSlot::Start => border.start = value,
            SideSlot::End => border.end = value,
            SideSlot::Left => border.left = value,
            SideSlot::Right => border.right = value,
            SideSlot::Top => border.top = value,
            SideSlot::Bottom => border.bottom = value,
            SideSlot::Diagonal => border.set_diagonal_side(value),
            SideSlot::Vertical => border.vertical = value,
            SideSlot::Horizontal => border.horizontal = value,
        }
        Ok(())
    }

    fn start_cell_format(&mut self, element: &BytesStart<'_>, decoder: Decoder) -> Result<()> {
        self.cell_format_count = self
            .cell_format_count
            .checked_add(1)
            .filter(|count| *count <= MAX_CELL_FORMATS)
            .ok_or_else(|| {
                invalid(format!(
                    "styles XML contains more than {MAX_CELL_FORMATS} cell formats"
                ))
            })?;
        let mut record = CellXf::default();
        for attribute in element.checked_attributes() {
            let attribute = attribute.map_err(attribute_error)?;
            let Some(name) = unqualified(&attribute) else {
                continue;
            };
            let (target, description) = match name {
                b"numFmtId" => (&mut record.number_format, "number-format ID"),
                b"fontId" => (&mut record.font, "font ID"),
                b"fillId" => (&mut record.fill, "fill ID"),
                b"borderId" => (&mut record.border, "border ID"),
                b"quotePrefix" => {
                    record.quote_prefix =
                        Some(parse_bool(&value(&attribute, decoder)?, "quotePrefix")?);
                    continue;
                },
                _ => continue,
            };
            *target = Some(parse_u32(&value(&attribute, decoder)?, description)?);
        }
        self.cell_format = Some(PendingCellFormat {
            record,
            seen_alignment: false,
            seen_protection: false,
        });
        Ok(())
    }
}

fn cell_format_count(element: &BytesStart<'_>, decoder: Decoder) -> Result<Option<u32>> {
    let mut count = None;
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(attribute_error)?;
        if unqualified(&attribute) == Some(b"count") {
            let text = value(&attribute, decoder)?;
            let parsed = text
                .parse::<u32>()
                .map_err(|_source| invalid("invalid cellXfs count"))?;
            if parsed > MAX_CELL_FORMATS {
                return Err(invalid(format!("cellXfs count exceeds {MAX_CELL_FORMATS}")));
            }
            count = Some(parsed);
        }
    }
    Ok(count)
}

/// Read a `CT_Color`: at most one of `rgb`, `theme`, `indexed` and `auto`,
/// plus an optional checked `tint`.
fn color(element: &BytesStart<'_>, decoder: Decoder) -> Result<Color> {
    let mut base = None;
    let mut tint = None;
    let mut bases = 0u8;
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(attribute_error)?;
        let Some(name) = unqualified(&attribute) else {
            continue;
        };
        let text = match name {
            b"rgb" | b"theme" | b"indexed" | b"auto" | b"tint" => value(&attribute, decoder)?,
            _ => continue,
        };
        match name {
            b"tint" => {
                let parsed = parse_f64(&text, "color tint")?;
                tint = Some(Tint::new(parsed).map_err(|_source| invalid("invalid color tint"))?);
            },
            b"rgb" => {
                bases += 1;
                let rgb =
                    Rgb::from_hex(text.trim()).map_err(|_source| invalid("invalid RGB color"))?;
                base = Some(Color::from_rgb(rgb));
            },
            b"theme" => {
                bases += 1;
                base = Some(Color::theme(parse_u32(&text, "theme color")?));
            },
            b"indexed" => {
                bases += 1;
                base = Some(Color::indexed(parse_u32(&text, "indexed color")?));
            },
            _ => {
                bases += 1;
                base = Some(Color::auto_value(parse_bool(&text, "automatic color")?));
            },
        }
    }
    if bases > 1 {
        return Err(invalid("color has multiple mutually exclusive values"));
    }
    let color = base.unwrap_or_else(Color::default_base);
    Ok(match tint {
        Some(tint) => color.with_tint(tint),
        None => color,
    })
}

/// Read a `CT_CellAlignment`. Recognised attributes are checked; others are
/// ignored.
fn alignment(element: &BytesStart<'_>, decoder: Decoder) -> Result<Alignment> {
    let mut alignment = Alignment::new();
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(attribute_error)?;
        let Some(name) = unqualified(&attribute) else {
            continue;
        };
        match name {
            b"horizontal" => {
                let token = value(&attribute, decoder)?;
                alignment.horizontal = Some(
                    token
                        .trim()
                        .parse::<Horizontal>()
                        .map_err(|_source| invalid("invalid horizontal alignment"))?,
                );
            },
            b"vertical" => {
                let token = value(&attribute, decoder)?;
                alignment.vertical = Some(
                    token
                        .trim()
                        .parse::<Vertical>()
                        .map_err(|_source| invalid("invalid vertical alignment"))?,
                );
            },
            b"textRotation" => {
                let degrees = parse_u32(&value(&attribute, decoder)?, "text rotation")?;
                alignment.text_rotation = Some(
                    Rotation::try_from(degrees)
                        .map_err(|_error| invalid("invalid text rotation"))?,
                );
            },
            b"wrapText" => {
                alignment.wrap_text = parse_bool(&value(&attribute, decoder)?, "wrapText")?;
            },
            b"indent" => {
                let indent = parse_u32(&value(&attribute, decoder)?, "alignment indent")?;
                alignment.indent = Some(
                    Indent::try_from(indent)
                        .map_err(|_error| invalid("invalid alignment indent"))?,
                );
            },
            b"relativeIndent" => {
                let text = value(&attribute, decoder)?;
                alignment.relative_indent = Some(
                    text.trim()
                        .parse::<i32>()
                        .map_err(|_source| invalid("invalid relative alignment indent"))?,
                );
            },
            b"justifyLastLine" => {
                alignment.justify_last_line =
                    Some(parse_bool(&value(&attribute, decoder)?, "justifyLastLine")?);
            },
            b"shrinkToFit" => {
                alignment.shrink_to_fit = parse_bool(&value(&attribute, decoder)?, "shrinkToFit")?;
            },
            b"readingOrder" => {
                let order = parse_u32(&value(&attribute, decoder)?, "reading order")?;
                alignment.reading_order = Some(
                    Reading::try_from(order).map_err(|_error| invalid("invalid reading order"))?,
                );
            },
            _ => {},
        }
    }
    Ok(alignment)
}

/// Read a `CT_CellProtection`.
fn protection(element: &BytesStart<'_>, decoder: Decoder) -> Result<Protection> {
    let mut protection = Protection::default();
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(attribute_error)?;
        match unqualified(&attribute) {
            Some(b"locked") => {
                protection.locked = Some(parse_bool(&value(&attribute, decoder)?, "locked")?);
            },
            Some(b"hidden") => {
                protection.hidden = Some(parse_bool(&value(&attribute, decoder)?, "hidden")?);
            },
            _ => {},
        }
    }
    Ok(protection)
}

fn is_spreadsheetml(namespace: &ResolveResult<'_>) -> bool {
    matches!(
        namespace,
        ResolveResult::Bound(Namespace(value))
            if *value == SPREADSHEETML_NAMESPACE || *value == STRICT_SPREADSHEETML_NAMESPACE
    )
}

/// The local name of an attribute without a namespace prefix. Prefixed
/// attributes, including namespace declarations, are extension markup.
fn unqualified<'a>(attribute: &'a Attribute<'_>) -> Option<&'a [u8]> {
    attribute
        .key
        .prefix()
        .is_none()
        .then(|| attribute.key.local_name().into_inner())
}

fn value<'a>(attribute: &'a Attribute<'_>, decoder: Decoder) -> Result<Cow<'a, str>> {
    attribute
        .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
        .map_err(|_error| invalid("styles XML attribute value is malformed"))
}

/// The `val` attribute of a property element, when present.
fn optional_val(element: &BytesStart<'_>, decoder: Decoder) -> Result<Option<String>> {
    let mut found = None;
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(attribute_error)?;
        if unqualified(&attribute) == Some(b"val") {
            found = Some(value(&attribute, decoder)?.trim().to_owned());
        }
    }
    Ok(found)
}

fn required_val(element: &BytesStart<'_>, decoder: Decoder, description: &str) -> Result<String> {
    optional_val(element, decoder)?
        .ok_or_else(|| invalid(format!("missing {description} attribute")))
}

/// A `CT_BooleanProperty`: `val` defaults to true.
fn boolean_val(element: &BytesStart<'_>, decoder: Decoder, description: &str) -> Result<bool> {
    optional_val(element, decoder)?.map_or(Ok(true), |text| parse_bool(&text, description))
}

fn parse_bool(text: &str, description: &str) -> Result<bool> {
    match text.trim() {
        "1" | "true" => Ok(true),
        "0" | "false" => Ok(false),
        _ => Err(invalid(format!("invalid {description} boolean"))),
    }
}

fn parse_u32(text: &str, description: &str) -> Result<u32> {
    text.trim()
        .parse::<u32>()
        .map_err(|_source| invalid(format!("invalid {description} value")))
}

fn parse_f64(text: &str, description: &str) -> Result<f64> {
    text.trim()
        .parse::<f64>()
        .ok()
        .filter(|value| value.is_finite())
        .ok_or_else(|| invalid(format!("invalid {description} value")))
}

fn mark(seen: &mut u16, bit: u16, description: &str) -> Result<()> {
    if *seen & bit != 0 {
        return Err(invalid(format!("duplicate {description}")));
    }
    *seen |= bit;
    Ok(())
}

fn push<T>(values: &mut Vec<T>, value: T, collection: &'static str) -> Result<()> {
    if values.len() >= MAX_RECORDS {
        return Err(invalid(format!(
            "styles XML contains more than {MAX_RECORDS} {collection}"
        )));
    }
    values
        .try_reserve(1)
        .map_err(|source| allocation("styles format table", source))?;
    values.push(value);
    Ok(())
}

fn attribute_error(_error: quick_xml::events::attributes::AttrError) -> Error {
    invalid("styles XML attribute is malformed")
}

fn missing_record() -> Error {
    invalid("styles XML record is not open")
}
