//! Reading worksheet features from one substream (record 0327).

use super::drawing::{self, Block, Object};
use super::stream::{
    self, AUTO_FILTER_INFO, COND_FMT, COND_FMT12, DV, DVAL, Globals, HORIZONTAL_PAGE_BREAKS, LBL,
    Sheet, VERTICAL_PAGE_BREAKS, u16_at, u32_at,
};
use super::{Area, FeatureSet, PageBreak};
use crate::formula::FormulaContext;
use crate::{Result, SourceLimits};

/// The features of one worksheet a [`FeatureSet`] selected.
#[derive(Debug, Clone, Default)]
pub struct Features {
    /// Legacy notes, in record order.
    pub notes: Vec<Note>,
    /// Hyperlinks, in record order.
    pub hyperlinks: Vec<crate::hyperlinks::Hyperlink>,
    /// Pictures and shapes, in drawing order.
    pub drawings: Vec<Drawing>,
    /// Conditional formats (`CondFmt` and `CondFmt12`), in record order.
    pub formats: Vec<ConditionalFormat>,
    /// Data validations (`Dv`), in record order.
    pub validations: Vec<Validation>,
    /// The auto-filter.
    pub auto_filter: Option<AutoFilter>,
    /// Manual row breaks (`HorizontalPageBreaks`).
    pub row_breaks: Vec<PageBreak>,
    /// Manual column breaks (`VerticalPageBreaks`).
    pub column_breaks: Vec<PageBreak>,
}

/// One legacy note (`Note` with its comment shape and `TXO`).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Note {
    /// The note's row.
    pub row: u16,
    /// The note's column.
    pub column: u16,
    /// `NoteSh.stAuthor`.
    pub author: String,
    /// The note text.
    pub text: String,
    /// Whether the note is shown.
    pub visible: bool,
    /// The comment shape's `Obj` identifier.
    pub object_id: u16,
}

/// Whether a drawing is a picture or another shape.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DrawingKind {
    /// A picture (`ot` 8).
    Picture,
    /// A line, rectangle, oval, arc, polygon, text box or OfficeArt shape.
    Shape,
}

/// A shape's fill.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Fill {
    /// No fill.
    None,
    /// A solid RGB fill (`0xRRGGBB`).
    Solid(u32),
}

/// A shape's outline.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Line {
    /// No outline.
    None,
    /// A solid RGB outline and its width in EMU.
    Solid {
        /// `0xRRGGBB`.
        color: u32,
        /// The width in EMU.
        width: u32,
    },
}

/// A picture or shape of a worksheet drawing.
#[derive(Debug, Clone, PartialEq)]
pub struct Drawing {
    /// Picture or shape.
    pub kind: DrawingKind,
    /// The `Obj` identifier.
    pub object_id: u16,
    /// The OfficeArt shape identifier.
    pub shape_id: u32,
    /// The OfficeArt shape type (`MSOSPT`).
    pub shape_type: u16,
    /// `wzName`.
    pub name: Option<String>,
    /// `wzDescription` (alternative text).
    pub description: Option<String>,
    /// The `TXO` text.
    pub text: Option<String>,
    /// The cell anchor.
    pub anchor: Option<crate::SheetAnchor>,
    /// The rotation in degrees.
    pub rotation: f64,
    /// Horizontal flip.
    pub flip_h: bool,
    /// Vertical flip.
    pub flip_v: bool,
    /// The fill of a shape.
    pub fill: Option<Fill>,
    /// The outline of a shape.
    pub line: Option<Line>,
    /// A picture's MIME type and size from its blip store entry.
    pub image: Option<(&'static str, u32)>,
}

/// A conditional-format value object (`CFVO`).
#[derive(Debug, Clone, PartialEq)]
pub struct Threshold {
    /// `cfvoType`: 1 number, 2 min, 3 max, 4 percent, 5 percentile, 7
    /// formula.
    pub kind: u8,
    /// The number or the rendered formula.
    pub value: Option<String>,
}

/// A color scale.
#[derive(Debug, Clone, PartialEq)]
pub struct ColorScale {
    /// Each control point and its color (`0xRRGGBB`).
    pub points: Vec<(Threshold, Option<u32>)>,
}

/// A data bar.
#[derive(Debug, Clone, PartialEq)]
pub struct DataBar {
    /// The bar color (`0xRRGGBB`).
    pub color: Option<u32>,
    /// The minimum.
    pub min: Threshold,
    /// The maximum.
    pub max: Threshold,
    /// `iPercentMin`.
    pub min_length: u8,
    /// `iPercentMax`.
    pub max_length: u8,
    /// `fShowValue`.
    pub show_value: bool,
}

/// An icon set.
#[derive(Debug, Clone, PartialEq)]
pub struct IconSet {
    /// `iIconSet`.
    pub icon_set: u8,
    /// `fReverse`.
    pub reverse: bool,
    /// Whether the value is shown (`!fIconOnly`).
    pub show_value: bool,
    /// The thresholds.
    pub thresholds: Vec<Threshold>,
}

/// The type of a conditional-format rule.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RuleType {
    /// A cell-value comparison.
    CellIs,
    /// A formula (and the formula templates).
    Expression,
    /// A color scale.
    ColorScale,
    /// A data bar.
    DataBar,
    /// An icon set.
    IconSet,
    /// A top/bottom N filter.
    TopN,
}

/// The formatting a rule applies.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Style {
    /// The solid fill (`0xRRGGBB`).
    pub fill: Option<u32>,
    /// The font color.
    pub font_color: Option<u32>,
    /// Bold.
    pub bold: Option<bool>,
    /// Italic.
    pub italic: Option<bool>,
    /// Strikethrough.
    pub strike: Option<bool>,
    /// `uls`: 1 single, 2 double, 0x21/0x22 accounting.
    pub underline: Option<u8>,
    /// The font height in twips.
    pub size_twips: Option<u32>,
    /// The font name.
    pub font_name: Option<String>,
}

/// The first rule of a conditional format.
#[derive(Debug, Clone, PartialEq)]
pub struct Rule {
    /// The rule type.
    pub rule_type: RuleType,
    /// `cp` for a cell-value rule (1–8).
    pub operator: Option<u8>,
    /// The rendered formulas.
    pub formulas: Vec<String>,
    /// The priority (a classic rule's position when no `CFEx` gives one).
    pub priority: u16,
    /// `fStopIfTrue`.
    pub stop_if_true: bool,
    /// `icfTemplate` (1 for a classic formula rule, 0 for a classic cell
    /// rule).
    pub template: u16,
    /// The text of a text rule (from its formula).
    pub text: Option<String>,
    /// The text operator (`ctp`).
    pub text_operator: u16,
    /// The rank of a top/bottom rule, or the standard deviations of an
    /// average rule.
    pub parameter: u16,
    /// A percent top/bottom rule.
    pub percent: bool,
    /// A bottom rule.
    pub bottom: bool,
    /// The color scale.
    pub color_scale: Option<ColorScale>,
    /// The data bar.
    pub data_bar: Option<DataBar>,
    /// The icon set.
    pub icon_set: Option<IconSet>,
    /// The formatting.
    pub style: Style,
}

/// One conditional format: its ranges and its first rule.
#[derive(Debug, Clone, PartialEq)]
pub struct ConditionalFormat {
    /// The ranges it applies to.
    pub ranges: Vec<Area>,
    /// A `CondFmt12` (Office 2007) format.
    pub future: bool,
    /// The number of rules.
    pub rule_count: usize,
    /// The first rule.
    pub rule: Option<Rule>,
}

/// One data validation (`Dv`).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Validation {
    /// The ranges it applies to.
    pub ranges: Vec<Area>,
    /// `valType`: 0 any, 1 whole, 2 decimal, 3 list, 4 date, 5 time, 6 text
    /// length, 7 custom.
    pub kind: u8,
    /// `typOperator`: 0 between … 7 less than or equal.
    pub operator: u8,
    /// `errStyle`: 0 stop, 1 warning, 2 information.
    pub error_style: u8,
    /// The first formula, rendered; an explicit list as `"a,b"`.
    pub formula1: Option<String>,
    /// The second formula, rendered.
    pub formula2: Option<String>,
    /// `fAllowBlank`.
    pub allow_blank: bool,
    /// `fSuppressCombo`.
    pub suppress_dropdown: bool,
    /// `fShowInputMsg`.
    pub show_input: bool,
    /// `fShowErrorMsg`.
    pub show_error: bool,
    /// The input title.
    pub prompt_title: Option<String>,
    /// The input message.
    pub prompt: Option<String>,
    /// The error title.
    pub error_title: Option<String>,
    /// The error message.
    pub error: Option<String>,
}

/// The auto-filter of a worksheet.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct AutoFilter {
    /// The filter range (from the sheet's `_FilterDatabase` name).
    pub area: Option<Area>,
    /// `AutoFilterInfo.cEntries`.
    pub columns: u16,
}

/// A formula context for rendering CF and DV formulas: the tab names and the
/// globals' supporting books and `ExternSheet`.
pub(super) fn formula_context(stream_bytes: &[u8], globals: &Globals) -> FormulaContext {
    let mut context = FormulaContext::default();
    let mut names = Vec::with_capacity(globals.bounds.len());
    for bound in &globals.bounds {
        names.push(sheet_name(bound.record.payload(stream_bytes)).unwrap_or_default());
    }
    for record in &globals.records {
        match record.kind {
            stream::SUP_BOOK => context.add_sup_book(record.payload(stream_bytes)),
            stream::EXTERN_SHEET => {
                let _ = context.add_extern_sheet(record.payload(stream_bytes));
            },
            _ => {},
        }
    }
    context.set_sheet_names(names);
    context
}

/// The name of a `BoundSheet8` payload.
pub(super) fn sheet_name(payload: &[u8]) -> Option<String> {
    let count = usize::from(*payload.get(6)?);
    let wide = payload.get(7)? & 1 == 1;
    let chars = payload.get(8..)?;
    Some(if wide {
        let units: Vec<u16> = chars
            .chunks_exact(2)
            .take(count)
            .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
            .collect();
        String::from_utf16_lossy(&units)
    } else {
        chars
            .iter()
            .take(count)
            .map(|byte| char::from(*byte))
            .collect()
    })
}

pub(super) fn palette(stream_bytes: &[u8], globals: &Globals) -> crate::Palette {
    globals
        .first(stream::PALETTE)
        .and_then(|index| {
            let mut report = crate::ToleranceReport::default();
            crate::Palette::parse_tolerant(
                globals.records[index].payload(stream_bytes),
                0,
                &mut report,
            )
            .ok()
        })
        .unwrap_or_default()
}

fn rgb_of(palette: &crate::Palette, icv: u32) -> Option<u32> {
    let color = palette.color(u16::try_from(icv).ok()?)?;
    Some((u32::from(color.red()) << 16) | (u32::from(color.green()) << 8) | u32::from(color.blue()))
}

/// Every feature `select` names on the worksheet at catalog position `sheet`.
pub(super) fn features(
    stream_bytes: &[u8],
    sheet: usize,
    select: FeatureSet,
    limits: &SourceLimits,
) -> Result<Features> {
    let globals = Globals::parse(stream_bytes)?;
    let start = globals.worksheet(sheet)?;
    let layout = Sheet::parse(stream_bytes, start)?;
    read(stream_bytes, &globals, &layout, sheet, select, limits)
}

pub(super) fn read(
    stream_bytes: &[u8],
    globals: &Globals,
    layout: &Sheet,
    sheet: usize,
    select: FeatureSet,
    limits: &SourceLimits,
) -> Result<Features> {
    let mut out = Features::default();
    let needs_context =
        select.contains(FeatureSet::FORMATS) || select.contains(FeatureSet::VALIDATIONS);
    let context = needs_context.then(|| formula_context(stream_bytes, globals));
    if select.contains(FeatureSet::NOTES) {
        let mut collector = crate::comments::CommentCollector::new();
        for (_, record) in layout.tail() {
            if record.kind != stream::BOF {
                collector.feed_record(record.kind, record.payload(stream_bytes))?;
            }
        }
        out.notes = collector
            .finish()?
            .into_iter()
            .map(|comment| Note {
                row: comment.row(),
                column: u16::from(comment.column()),
                author: comment.author().to_owned(),
                text: comment.text().to_owned(),
                visible: comment.visibility() == crate::comments::Visibility::Visible,
                object_id: comment.identity().object_id(),
            })
            .collect();
    }
    if select.contains(FeatureSet::HYPERLINKS) {
        let mut collector = crate::hyperlinks::HyperlinkCollector::new(*limits);
        for (_, record) in layout.tail() {
            if record.kind != stream::BOF {
                collector.feed_record(record.kind, record.payload(stream_bytes))?;
            }
        }
        out.hyperlinks = collector.finish();
    }
    if select.contains(FeatureSet::DRAWINGS) {
        if let Some(block) = Block::parse(stream_bytes, layout)? {
            let store = drawing_store(stream_bytes, globals)?;
            out.drawings = block
                .objects
                .iter()
                .filter_map(|object| drawing_of(object, &store))
                .collect();
        }
    }
    if select.contains(FeatureSet::FORMATS) {
        out.formats = formats(stream_bytes, globals, layout, context.as_ref())?;
    }
    if select.contains(FeatureSet::VALIDATIONS) {
        for (_, record) in layout.tail() {
            if record.kind == DV {
                let rule = crate::data_validation::parse_dv(
                    record.payload(stream_bytes),
                    context.as_ref(),
                )?;
                out.validations
                    .push(validation_of(record.payload(stream_bytes), &rule));
            }
        }
    }
    if select.contains(FeatureSet::AUTO_FILTER) {
        if let Some(index) = layout.find(AUTO_FILTER_INFO) {
            let columns = u16_at(layout.records[index].payload(stream_bytes), 0);
            out.auto_filter = Some(AutoFilter {
                area: filter_database(stream_bytes, globals, sheet).map(|(_, area)| area),
                columns,
            });
        }
    }
    if select.contains(FeatureSet::PAGE_BREAKS) {
        for record in &layout.records[..layout.dimensions] {
            match record.kind {
                HORIZONTAL_PAGE_BREAKS => out.row_breaks = breaks(record.payload(stream_bytes)),
                VERTICAL_PAGE_BREAKS => out.column_breaks = breaks(record.payload(stream_bytes)),
                _ => {},
            }
        }
    }
    let _ = DVAL;
    Ok(out)
}

pub(super) fn breaks(payload: &[u8]) -> Vec<PageBreak> {
    let count = usize::from(u16_at(payload, 0));
    (0..count)
        .filter(|index| 2 + index * 6 + 6 <= payload.len())
        .map(|index| PageBreak {
            index: u16_at(payload, 2 + index * 6),
            min: u16_at(payload, 4 + index * 6),
            max: u16_at(payload, 6 + index * 6),
        })
        .collect()
}

/// The `_FilterDatabase` name of `sheet`: its `Lbl` record index in the
/// globals and its range.
pub(super) fn filter_database(
    stream_bytes: &[u8],
    globals: &Globals,
    sheet: usize,
) -> Option<(usize, Area)> {
    globals
        .records
        .iter()
        .enumerate()
        .find_map(|(index, record)| {
            if record.kind != LBL {
                return None;
            }
            let payload = record.payload(stream_bytes);
            let flags = u16_at(payload, 0);
            let cch = usize::from(*payload.get(3)?);
            let cce = usize::from(u16_at(payload, 4));
            let itab = usize::from(u16_at(payload, 8));
            if flags & 0x0020 == 0 || cch != 1 || itab != sheet + 1 {
                return None;
            }
            // Name: one built-in code after its fHighByte.
            let wide = payload.get(14)? & 1 == 1;
            if *payload.get(15)? != 0x0d {
                return None;
            }
            let rgce = 15 + if wide { 2 } else { 1 };
            let tokens = payload.get(rgce..rgce + cce)?;
            let area = area3d(tokens)?;
            Some((index, area))
        })
}

/// The rectangle of a lone `PtgArea3d` formula.
pub(super) fn area3d(tokens: &[u8]) -> Option<Area> {
    let ptg = *tokens.first()?;
    if !matches!(ptg, 0x3b | 0x5b | 0x7b) || tokens.len() < 11 {
        return None;
    }
    Some(Area {
        first_row: u16_at(tokens, 3),
        last_row: u16_at(tokens, 5),
        first_column: u16_at(tokens, 7) & 0x3fff,
        last_column: u16_at(tokens, 9) & 0x3fff,
    })
}

fn validation_of(payload: &[u8], rule: &crate::DataValidationRule) -> Validation {
    let flags = u32_at(payload, 0);
    let text = |value: Option<&str>| {
        value
            .filter(|value| !value.is_empty() && *value != "\u{0}")
            .map(str::to_owned)
    };
    let render = |formula: Option<&crate::DataValidationFormula>| {
        formula.and_then(|formula| formula.rendered().map(|text| bare(text).to_owned()))
    };
    let list = flags & 0x80 != 0;
    let mut formula1 = render(rule.formula1());
    if list {
        // An explicit list renders as a quoted string with NUL separators.
        formula1 = formula1.map(|value| value.replace('\0', ","));
    }
    Validation {
        ranges: rule
            .ranges()
            .iter()
            .map(|range| Area {
                first_row: range.first_row(),
                last_row: range.last_row(),
                first_column: u16::from(range.first_column()),
                last_column: u16::from(range.last_column()),
            })
            .collect(),
        kind: (flags & 0x0f) as u8,
        operator: ((flags >> 20) & 0x0f) as u8,
        error_style: ((flags >> 4) & 0x07) as u8,
        formula1,
        formula2: render(rule.formula2()),
        allow_blank: flags & 0x100 != 0,
        suppress_dropdown: flags & 0x200 != 0,
        show_input: flags & 0x0004_0000 != 0,
        show_error: flags & 0x0008_0000 != 0,
        prompt_title: text(rule.prompt_title()),
        prompt: text(rule.prompt()),
        error_title: text(rule.error_title()),
        error: text(rule.error()),
    }
}

/// The blip store entries (type and size) by one-based `pib`.
pub(super) fn drawing_store(stream_bytes: &[u8], globals: &Globals) -> Result<Vec<(u8, u32)>> {
    let Some(group) = drawing::Group::parse(stream_bytes, globals)? else {
        return Ok(Vec::new());
    };
    Ok(group
        .store()?
        .map(|(_, entries)| {
            entries
                .iter()
                .map(|entry| (entry.win, entry.size))
                .collect()
        })
        .unwrap_or_default())
}

/// The drawing an object projects: a picture or a plain shape; comments,
/// form controls, charts and groups are not drawings here.
pub(super) fn drawing_of(object: &Object, store: &[(u8, u32)]) -> Option<Drawing> {
    let (ot, object_id) = object.obj?;
    if object.group {
        return None;
    }
    let kind = match ot {
        8 => DrawingKind::Picture,
        1 | 2 | 3 | 4 | 6 | 9 | 0x1e => DrawingKind::Shape,
        _ => return None,
    };
    let anchor = object.anchor.and_then(|bytes| {
        let point = |at: usize| {
            crate::AnchorPoint::new(
                u16_at(&bytes, at),
                u16_at(&bytes, at + 4),
                i16::from_le_bytes([bytes[at + 2], bytes[at + 3]]),
                i16::from_le_bytes([bytes[at + 6], bytes[at + 7]]),
            )
            .ok()
        };
        let behavior = match u16_at(&bytes, 0) & 3 {
            0 => crate::AnchorBehavior::MoveAndResize,
            2 => crate::AnchorBehavior::MoveWithoutResize,
            _ => crate::AnchorBehavior::Absolute,
        };
        crate::SheetAnchor::new(point(2)?, point(10)?, behavior).ok()
    });
    let rotation = object
        .prop(0x0004)
        .map_or(0.0, |prop| f64::from(prop.value as i32) / 65_536.0);
    let bool_prop = |id: u16, bit: u32, default: bool| {
        object.prop(id).map_or(default, |prop| {
            let use_bit = bit << 16;
            if prop.value & use_bit != 0 {
                prop.value & bit != 0
            } else {
                default
            }
        })
    };
    let colorref = |value: u32| {
        // OfficeArtCOLORREF: 0x00BBGGRR.
        ((value & 0xff) << 16) | (value & 0xff00) | ((value >> 16) & 0xff)
    };
    let (fill, line) = if kind == DrawingKind::Shape {
        let filled = bool_prop(0x01bf, 0x10, true);
        let fill = if filled {
            Fill::Solid(
                object
                    .prop(0x0181)
                    .map_or(0x00ff_ffff, |prop| colorref(prop.value)),
            )
        } else {
            Fill::None
        };
        let lined = bool_prop(0x01ff, 0x08, true);
        let line = if lined {
            Line::Solid {
                color: object.prop(0x01c0).map_or(0, |prop| colorref(prop.value)),
                width: object.prop(0x01cb).map_or(9_525, |prop| prop.value),
            }
        } else {
            Line::None
        };
        (Some(fill), Some(line))
    } else {
        (None, None)
    };
    let image = (kind == DrawingKind::Picture)
        .then(|| object.prop(0x0104))
        .flatten()
        .and_then(|prop| store.get((prop.value as usize).checked_sub(1)?))
        .map(|(win, size)| (drawing::content_type(*win), *size));
    Some(Drawing {
        kind,
        object_id,
        shape_id: object.shape_id,
        shape_type: object.shape_type,
        name: object.complex_text(0x0380),
        description: object.complex_text(0x0381),
        text: object.text.clone().filter(|text| !text.is_empty()),
        anchor,
        rotation,
        flip_h: object.flags & 0x40 != 0,
        flip_v: object.flags & 0x80 != 0,
        fill,
        line,
        image,
    })
}

// ---------------------------------------------------------------------------
// Conditional formats

fn formats(
    stream_bytes: &[u8],
    globals: &Globals,
    layout: &Sheet,
    context: Option<&FormulaContext>,
) -> Result<Vec<ConditionalFormat>> {
    let mut collector = crate::conditional_format::ConditionalFormatCollector::new();
    let mut order = Vec::new();
    for (_, record) in layout.tail() {
        match record.kind {
            COND_FMT => order.push(false),
            COND_FMT12 => order.push(true),
            _ => {},
        }
        if record.kind != stream::BOF {
            collector.feed_record(record.kind, record.payload(stream_bytes), context)?;
        }
    }
    let (legacy, future, extensions) = collector.finish()?;
    let palette = palette(stream_bytes, globals);
    let mut legacy = legacy.into_iter();
    let mut future = future.into_iter();
    let mut out = Vec::with_capacity(order.len());
    for (position, is_future) in order.into_iter().enumerate() {
        if is_future {
            let Some(formatting) = future.next() else {
                break;
            };
            out.push(ConditionalFormat {
                ranges: formatting.ranges().iter().map(area_of).collect(),
                future: true,
                rule_count: formatting.rules().len(),
                rule: formatting
                    .rules()
                    .first()
                    .map(|rule| rule12(rule, &palette)),
            });
        } else {
            let Some(formatting) = legacy.next() else {
                break;
            };
            let extension = extensions.iter().find(|extension| {
                extension.identifier() == formatting.identifier()
                    && extension.legacy_rule_index() == Some(0)
            });
            out.push(ConditionalFormat {
                ranges: formatting.ranges().iter().map(area_of).collect(),
                future: false,
                rule_count: formatting.rules().len(),
                rule: formatting.rules().first().map(|rule| {
                    let priority = extension.map_or_else(
                        || u16::try_from(position + 1).unwrap_or(u16::MAX),
                        crate::ConditionalExtension::priority,
                    );
                    legacy_rule(rule, priority, extension, &palette)
                }),
            });
        }
    }
    Ok(out)
}

fn area_of(range: &crate::ConditionalFormatRange) -> Area {
    Area {
        first_row: range.first_row(),
        last_row: range.last_row(),
        first_column: u16::from(range.first_column()),
        last_column: u16::from(range.last_column()),
    }
}

fn legacy_rule(
    rule: &crate::ConditionalRule,
    priority: u16,
    extension: Option<&crate::ConditionalExtension>,
    palette: &crate::Palette,
) -> Rule {
    let (rule_type, operator) = match rule.kind() {
        crate::ConditionalRuleKind::CellValue(comparison) => {
            (RuleType::CellIs, Some(comparison_code(comparison)))
        },
        _ => (RuleType::Expression, None),
    };
    let formulas = [rule.formula1_rendered(), rule.formula2_rendered()]
        .into_iter()
        .flatten()
        .map(bare)
        .filter(|formula| !formula.is_empty())
        .map(str::to_owned)
        .collect();
    let mut out = Rule {
        rule_type,
        operator,
        formulas,
        priority,
        stop_if_true: extension.is_some_and(crate::ConditionalExtension::stop_if_true),
        template: if rule_type == RuleType::CellIs { 0 } else { 1 },
        text: None,
        text_operator: 0,
        parameter: 0,
        percent: false,
        bottom: false,
        color_scale: None,
        data_bar: None,
        icon_set: None,
        style: style_of(rule.style(), palette),
    };
    if let Some(extension) = extension {
        out.template = u16::from(extension.template());
        apply_template(&mut out, extension.template_parameters());
    }
    out
}

fn comparison_code(comparison: crate::ConditionalComparison) -> u8 {
    use crate::ConditionalComparison as C;
    match comparison {
        C::Between => 1,
        C::NotBetween => 2,
        C::Equal => 3,
        C::NotEqual => 4,
        C::GreaterThan => 5,
        C::LessThan => 6,
        C::GreaterThanOrEqual => 7,
        _ => 8,
    }
}

fn style_of(style: &crate::ConditionalStyle, palette: &crate::Palette) -> Style {
    let mut out = Style::default();
    if let Some(pattern) = style.pattern() {
        if pattern.fill_pattern() != 0 {
            let icv = if style.is_pattern_background_modified() {
                pattern.background_color_index()
            } else {
                pattern.foreground_color_index()
            };
            out.fill = rgb_of(palette, u32::from(icv));
        }
    }
    if let Some(font) = style.font() {
        let raw = font.raw_data();
        font_of(raw, palette, &mut out);
    }
    out
}

/// Project a `DXFFntD` (118 bytes).
fn font_of(raw: &[u8], palette: &crate::Palette, out: &mut Style) {
    if raw.len() < 118 {
        return;
    }
    let count = usize::from(raw[0]);
    if count > 0 {
        let wide = raw[1] & 1 == 1;
        let chars = &raw[2..64];
        out.font_name = Some(if wide {
            let units: Vec<u16> = chars
                .chunks_exact(2)
                .take(count)
                .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
                .collect();
            String::from_utf16_lossy(&units)
        } else {
            chars
                .iter()
                .take(count)
                .map(|byte| char::from(*byte))
                .collect()
        });
    }
    let height = u32_at(raw, 64);
    if height != u32::MAX && height != 0 {
        out.size_twips = Some(height);
    }
    let ts = u32_at(raw, 68);
    let ts_ninch = u32_at(raw, 88);
    if ts_ninch & 0x02 == 0 {
        out.italic = Some(ts & 0x02 != 0);
    }
    if ts_ninch & 0x80 == 0 {
        out.strike = Some(ts & 0x80 != 0);
    }
    if u32_at(raw, 100) == 0 {
        out.bold = Some(u16_at(raw, 72) >= 700);
    }
    if u32_at(raw, 96) == 0 && raw[76] != 0 {
        out.underline = Some(raw[76]);
    }
    let icv = u32_at(raw, 80);
    if icv != u32::MAX && icv != 0x7fff {
        out.font_color = rgb_of(palette, icv);
    }
}

fn rule12(rule: &crate::ConditionalRule12, palette: &crate::Palette) -> Rule {
    use crate::ConditionalRule12Kind as K;
    let (rule_type, operator) = match rule.kind() {
        K::CellValue(comparison) => (RuleType::CellIs, Some(comparison_code(comparison))),
        K::ColorScale => (RuleType::ColorScale, None),
        K::DataBar => (RuleType::DataBar, None),
        K::IconSet => (RuleType::IconSet, None),
        K::Filter => (RuleType::TopN, None),
        _ => (RuleType::Expression, None),
    };
    let formulas = [rule.formula1_rendered(), rule.formula2_rendered()]
        .into_iter()
        .flatten()
        .map(bare)
        .filter(|formula| !formula.is_empty())
        .map(str::to_owned)
        .collect();
    let mut out = Rule {
        rule_type,
        operator,
        formulas,
        priority: rule.priority(),
        stop_if_true: rule.stop_if_true(),
        template: rule.template(),
        text: None,
        text_operator: 0,
        parameter: 0,
        percent: false,
        bottom: false,
        color_scale: None,
        data_bar: None,
        icon_set: None,
        style: dxfn12(rule.differential_format(), palette),
    };
    apply_template(&mut out, rule.template_parameters());
    let payload = rule.rule_payload();
    match rule_type {
        RuleType::ColorScale => out.color_scale = color_scale(payload),
        RuleType::DataBar => out.data_bar = data_bar(payload),
        RuleType::IconSet => out.icon_set = icon_set(payload),
        RuleType::TopN => {
            // CFFilter: cbFilter, reserved, flags, iParam.
            if payload.len() >= 6 {
                out.bottom = payload[3] & 1 == 0;
                out.percent = payload[3] & 2 != 0;
                out.parameter = u16_at(payload, 4);
            }
        },
        _ => {},
    }
    out
}

fn apply_template(rule: &mut Rule, params: &[u8; 16]) {
    match rule.template {
        5 => {
            rule.rule_type = RuleType::TopN;
            rule.bottom = params[0] & 1 == 0;
            rule.percent = params[0] & 2 != 0;
            rule.parameter = u16_at(params, 1);
        },
        8 => {
            rule.text_operator = u16_at(params, 0);
            rule.text = rule.formulas.first().and_then(|formula| quoted(formula));
        },
        0x19 | 0x1a | 0x1d | 0x1e => rule.parameter = u16_at(params, 0),
        _ => {},
    }
}

/// The first string literal of a rendered formula.
fn quoted(formula: &str) -> Option<String> {
    let start = formula.find('"')? + 1;
    let mut out = String::new();
    let mut chars = formula[start..].chars().peekable();
    while let Some(char) = chars.next() {
        if char == '"' {
            if chars.peek() == Some(&'"') {
                out.push('"');
                chars.next();
                continue;
            }
            return Some(out);
        }
        out.push(char);
    }
    None
}

fn dxfn12(bytes: &[u8], palette: &crate::Palette) -> Style {
    let size = u32_at(bytes, 0) as usize;
    if size == 0 || bytes.len() < 4 + size {
        return Style::default();
    }
    // A CF record shape around the DXFN lets the classic parser read it.
    let mut record = vec![2u8, 0, 0, 0, 0, 0];
    record.extend_from_slice(&bytes[4..4 + size]);
    match crate::conditional_format::parse_cf(&record, None) {
        Ok(rule) => style_of(rule.style(), palette),
        Err(_) => Style::default(),
    }
}

/// Parse one `CFVO` at `at`; returns it and the offset after it.
fn cfvo(bytes: &[u8], at: usize) -> Option<(Threshold, usize)> {
    let kind = *bytes.get(at)?;
    let cce = usize::from(u16_at(bytes, at + 1));
    let mut next = at + 3;
    let tokens = bytes.get(next..next + cce)?;
    next += cce;
    let value = if cce > 0 {
        crate::formula::render_formula(tokens, None).map(|text| bare(&text).to_owned())
    } else if matches!(kind, 2 | 3) {
        None
    } else {
        let number = f64::from_le_bytes(bytes.get(next..next + 8)?.try_into().ok()?);
        next += 8;
        Some(format_number(number))
    };
    Some((Threshold { kind, value }, next))
}

/// A rendered formula without its leading `=`.
fn bare(formula: &str) -> &str {
    formula.strip_prefix('=').unwrap_or(formula)
}

/// A number in its shortest round-trip spelling, integers without a
/// fraction.
pub(super) fn format_number(number: f64) -> String {
    if number.fract() == 0.0 && number.abs() < 1e15 {
        #[allow(clippy::cast_possible_truncation, reason = "bounded above")]
        let integer = number as i64;
        integer.to_string()
    } else {
        number.to_string()
    }
}

/// A `CFColor` (16 bytes) as RGB, when it is one.
fn cf_color(bytes: &[u8], at: usize) -> Option<u32> {
    let kind = u32_at(bytes, at);
    let value = u32_at(bytes, at + 4);
    match kind {
        // XCLRRGB: LongRGBA (R, G, B, A bytes).
        2 => Some(((value & 0xff) << 16) | (value & 0xff00) | ((value >> 16) & 0xff)),
        1 => rgb_of(&crate::Palette::default(), value),
        _ => None,
    }
}

fn color_scale(bytes: &[u8]) -> Option<ColorScale> {
    let count = usize::from(*bytes.get(3)?);
    let mut at = 6;
    let mut thresholds = Vec::with_capacity(count);
    for _ in 0..count {
        let (threshold, next) = cfvo(bytes, at)?;
        thresholds.push(threshold);
        at = next + 8;
    }
    let mut points = Vec::with_capacity(count);
    for threshold in thresholds {
        let color = cf_color(bytes, at + 8);
        points.push((threshold, color));
        at += 24;
    }
    Some(ColorScale { points })
}

fn data_bar(bytes: &[u8]) -> Option<DataBar> {
    let flags = *bytes.get(3)?;
    let min_length = *bytes.get(4)?;
    let max_length = *bytes.get(5)?;
    let color = cf_color(bytes, 6);
    let (min, next) = cfvo(bytes, 22)?;
    let (max, _) = cfvo(bytes, next)?;
    Some(DataBar {
        color,
        min,
        max,
        min_length,
        max_length,
        show_value: flags & 2 != 0,
    })
}

fn icon_set(bytes: &[u8]) -> Option<IconSet> {
    let count = usize::from(*bytes.get(3)?);
    let set = *bytes.get(4)?;
    let flags = *bytes.get(5)?;
    let mut at = 6;
    let mut thresholds = Vec::with_capacity(count);
    for _ in 0..count {
        let (threshold, next) = cfvo(bytes, at)?;
        thresholds.push(threshold);
        at = next + 5;
    }
    Some(IconSet {
        icon_set: set,
        reverse: flags & 4 != 0,
        show_value: flags & 1 == 0,
        thresholds,
    })
}
