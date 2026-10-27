//! The per-series properties stored in a chart's series styles (strokes,
//! fills, data-symbol fills and outlines) for [`super::sheet_chart_series`]
//! (record 0378): the host editor's `set_sheet_chart_series_{strokes,
//! fills, symbol_fills, symbol_outlines}` and `reset_*` routes on W1-E's
//! series-style write rules (`chart_style_codec::property`, record 0375a).
//!
//! The value of a series is read through its effective style (the private
//! override, else the theme style) and its `TSS.StyleArchive.parent` chain;
//! a write rewrites the effective style in place, which only this chart may
//! use (the host's `ensure_exclusive`). Nothing is allocated and the package
//! metadata does not change.

use litchi_iwa_common::chart::kind::Kind as ChartKind;
use litchi_iwa_common::chart::series::{
    FillFamily, Stroke, StrokeFamily, SymbolFill, SymbolFillFamily, SymbolOutlineFamily,
};
use litchi_iwa_common::shape::fill::Fill;
use litchi_iwa_core::RawMessage;
use litchi_iwa_protos::chart_style_codec::{
    self as codec, Archive, Limits, Properties, property as rules,
};
use litchi_iwa_protos::shape_style_codec::native as stored;
use litchi_iwa_style::drawing;

use super::Package;
use super::sheet_drawable_graph::{self as graph, CHART_MESSAGE_TYPE};
use super::structure::{self, Fault, FaultResult, Published};

/// One style object's in-place rewrite.
type StyleRewrite<'a> =
    dyn FnMut(&mut litchi_iwa_core::ArchiveObject, litchi_iwa_core::Limits) -> FaultResult<()> + 'a;

/// Replace message `index` (of `message_type`) with `data`, keeping the
/// archive header.
fn rewriter(
    index: usize,
    message_type: u32,
    data: Vec<u8>,
) -> impl FnMut(&mut litchi_iwa_core::ArchiveObject, litchi_iwa_core::Limits) -> FaultResult<()> {
    let mut data = Some(data);
    move |object, limits| {
        object
            .replace_message_preserving_header(
                index,
                RawMessage {
                    type_: message_type,
                    data: data.take().ok_or(Fault::Invalid)?,
                },
                limits,
            )
            .map(drop)
            .map_err(|error| structure::map_core(&error))
    }
}

/// The longest style parent chain followed.
const MAX_DEPTH: usize = 64;

/// A property of the series styles.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum Kind {
    Strokes,
    Fills,
    SymbolFills,
    SymbolOutlines,
}

/// One series' value of a style property.
#[derive(Debug, Clone, PartialEq)]
pub(super) enum Value {
    Stroke(Option<Stroke>),
    Fill(Fill),
    SymbolFill(SymbolFill),
}

/// Why a style property is refused.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum Refused {
    /// The chart kind has no such property.
    Kind,
    /// Another chart uses the series style.
    Shared,
    /// An image fill needs the data-reference graph.
    ImageFill,
}

#[derive(Debug)]
pub(super) enum Failure {
    Refused(Refused),
    Fault(Fault),
}

impl From<Fault> for Failure {
    fn from(fault: Fault) -> Self {
        Self::Fault(fault)
    }
}

impl From<codec::Error> for Failure {
    fn from(error: codec::Error) -> Self {
        match error {
            codec::Error::Allocation { bytes, .. } => Self::Fault(structure::allocation(bytes)),
            _ => Self::Fault(Fault::Invalid),
        }
    }
}

type Result<T> = std::result::Result<T, Failure>;

/// The effective series styles of the chart's first `count` series.
fn styles(package: &Package, chart: u64, count: usize) -> Result<Vec<u64>> {
    let object = graph::object(package, chart)?;
    let index = structure::unique_message(object, &[CHART_MESSAGE_TYPE])?;
    let references = codec::references(&object.messages[index].data, Limits::DEFAULT)?;
    (0..count)
        .map(|series| {
            references
                .series_style(series)
                .ok_or(Failure::Fault(Fault::Invalid))
        })
        .collect()
}

fn style_payload(package: &Package, style: u64) -> FaultResult<(usize, &[u8])> {
    let object = graph::object(package, style)?;
    let index = structure::unique_message(object, &[Archive::SeriesStyle.message_type()])?;
    Ok((index, &object.messages[index].data))
}

/// The first value `read` finds up the parent chain of `style`.
fn inherited<T>(
    package: &Package,
    style: u64,
    read: impl Fn(&Properties<'_>) -> Result<Option<T>>,
) -> Result<Option<T>> {
    let mut current = style;
    for _ in 0..MAX_DEPTH {
        let (_, payload) = style_payload(package, current)?;
        let properties = codec::decode(payload, Archive::SeriesStyle, Limits::DEFAULT)?;
        if let Some(value) = read(&properties)? {
            return Ok(Some(value));
        }
        match codec::parent(payload, Limits::DEFAULT)? {
            Some(parent) => current = parent,
            None => return Ok(None),
        }
    }
    Err(Failure::Fault(Fault::Invalid))
}

fn stroke_at(properties: &Properties<'_>, field: u32) -> Result<Option<Option<Stroke>>> {
    let Some(bytes) = properties.message(field)? else {
        return Ok(None);
    };
    let decoded =
        stored::stroke(bytes, Limits::DEFAULT).map_err(|_error| Failure::Fault(Fault::Invalid))?;
    let stroke = drawing::stroke(&decoded).map_err(|_error| Failure::Fault(Fault::Invalid))?;
    Ok(Some(
        stroke
            .map(Stroke::from_shape_stroke)
            .transpose()
            .map_err(|_error| Failure::Fault(Fault::Invalid))?,
    ))
}

fn fill_at(properties: &Properties<'_>, field: u32) -> Result<Option<Fill>> {
    let Some(bytes) = properties.message(field)? else {
        return Ok(None);
    };
    let decoded =
        stored::fill(bytes, Limits::DEFAULT).map_err(|_error| Failure::Fault(Fault::Invalid))?;
    Ok(Some(
        drawing::fill(&decoded).map_err(|_error| Failure::Fault(Fault::Invalid))?,
    ))
}

fn flag(properties: &Properties<'_>, field: u32) -> Result<Option<bool>> {
    Ok(properties.varint(field)?.map(|value| value != 0))
}

fn symbol_fill_at(
    properties: &Properties<'_>,
    family: SymbolFillFamily,
) -> Result<Option<SymbolFill>> {
    let use_fill = match family.use_series_fill_field() {
        Some(field) => flag(properties, field)?,
        None => None,
    };
    let use_stroke = flag(properties, family.use_series_stroke_field())?;
    family
        .read(
            use_fill,
            use_stroke,
            fill_at(properties, family.fill_field())?,
        )
        .map_err(|_error| Failure::Fault(Fault::Invalid))
}

/// The effective value of `kind` in `style` (and whether `style` defines
/// it itself).
fn effective(package: &Package, style: u64, chart: ChartKind, kind: Kind) -> Result<(Value, bool)> {
    let (_, payload) = style_payload(package, style)?;
    let local = codec::decode(payload, Archive::SeriesStyle, Limits::DEFAULT)?;
    Ok(match kind {
        Kind::Strokes | Kind::SymbolOutlines => {
            let field = stroke_field(chart, kind)?;
            let defined = stroke_at(&local, field)?.is_some();
            let value = inherited(package, style, |p| stroke_at(p, field))?.flatten();
            (Value::Stroke(value), defined)
        },
        Kind::Fills => {
            let family =
                FillFamily::for_chart_kind(chart).ok_or(Failure::Refused(Refused::Kind))?;
            let defined = fill_at(&local, family.field())?.is_some();
            let value = match inherited(package, style, |p| fill_at(p, family.field()))? {
                Some(fill) => fill,
                None => inherited(package, style, |p| fill_at(p, FillFamily::DEFAULT_FIELD))?
                    .ok_or(Failure::Fault(Fault::Invalid))?,
            };
            (Value::Fill(value), defined)
        },
        Kind::SymbolFills => {
            let family =
                SymbolFillFamily::for_chart_kind(chart).ok_or(Failure::Refused(Refused::Kind))?;
            let defined = symbol_fill_at(&local, family)?.is_some();
            let value = inherited(package, style, |p| symbol_fill_at(p, family))?
                .unwrap_or_else(|| family.default_fill());
            (Value::SymbolFill(value), defined)
        },
    })
}

fn stroke_field(chart: ChartKind, kind: Kind) -> Result<u32> {
    match kind {
        Kind::Strokes => StrokeFamily::for_chart_kind(chart).map(StrokeFamily::field),
        _ => SymbolOutlineFamily::for_chart_kind(chart).map(SymbolOutlineFamily::field),
    }
    .ok_or(Failure::Refused(Refused::Kind))
}

/// The effective value of `kind` for each of the chart's `count` series.
pub(super) fn read(
    package: &Package,
    chart: u64,
    chart_kind: ChartKind,
    count: usize,
    kind: Kind,
) -> Result<Vec<Value>> {
    styles(package, chart, count)?
        .into_iter()
        .map(|style| effective(package, style, chart_kind, kind).map(|(value, _)| value))
        .collect()
}

/// Whether only one chart's series use `style` (the host's
/// `ensure_exclusive`).
fn exclusive(package: &Package, style: u64) -> Result<bool> {
    let mut owners = 0_usize;
    for object in package.state.components.iter_objects() {
        for message in &object.messages {
            if message.type_ != CHART_MESSAGE_TYPE {
                continue;
            }
            let references = codec::references(&message.data, Limits::DEFAULT)?;
            let private = references
                .series_private_styles()
                .is_some_and(|sparse| sparse.entries().iter().any(|(_, id)| *id == style));
            if private || references.series_theme_styles().contains(&style) {
                owners += 1;
            }
        }
    }
    Ok(owners == 1)
}

/// The rewritten payload of `style` storing `value` (`None`: reset).
fn rewrite(payload: &[u8], chart: ChartKind, kind: Kind, value: Option<&Value>) -> Result<Vec<u8>> {
    let encoded: Vec<u8>;
    let write = match (kind, value) {
        (Kind::Strokes, value) => {
            let field = stroke_field(chart, kind)?;
            match value {
                Some(Value::Stroke(stroke)) => {
                    encoded = drawing::encode_stroke(stroke.map(Stroke::to_shape_stroke));
                    rules::series_stroke(field, Some(&encoded))
                },
                None => rules::series_stroke(field, None),
                Some(_) => return Err(Failure::Fault(Fault::Invalid)),
            }
        },
        (Kind::SymbolOutlines, value) => {
            let field = stroke_field(chart, kind)?;
            match value {
                Some(Value::Stroke(stroke)) => {
                    encoded = drawing::encode_stroke(stroke.map(Stroke::to_shape_stroke));
                    rules::series_message(field, Some(&encoded))
                },
                None => rules::series_message(field, None),
                Some(_) => return Err(Failure::Fault(Fault::Invalid)),
            }
        },
        (Kind::Fills, value) => {
            let family =
                FillFamily::for_chart_kind(chart).ok_or(Failure::Refused(Refused::Kind))?;
            match value {
                Some(Value::Fill(Fill::Image(_))) => {
                    return Err(Failure::Refused(Refused::ImageFill));
                },
                Some(Value::Fill(fill)) => {
                    encoded = drawing::encode_fill(fill);
                    rules::series_message(family.field(), Some(&encoded))
                },
                None => rules::series_message(family.field(), None),
                Some(_) => return Err(Failure::Fault(Fault::Invalid)),
            }
        },
        (Kind::SymbolFills, value) => {
            let family =
                SymbolFillFamily::for_chart_kind(chart).ok_or(Failure::Refused(Refused::Kind))?;
            let stored = match value {
                Some(Value::SymbolFill(SymbolFill::Custom(Fill::Image(_)))) => {
                    return Err(Failure::Refused(Refused::ImageFill));
                },
                Some(Value::SymbolFill(SymbolFill::Custom(fill))) => {
                    encoded = drawing::encode_fill(fill);
                    rules::SymbolFill::Custom(&encoded)
                },
                Some(Value::SymbolFill(SymbolFill::SeriesFill)) => rules::SymbolFill::SeriesFill,
                Some(Value::SymbolFill(SymbolFill::SeriesStroke)) => {
                    rules::SymbolFill::SeriesStroke
                },
                None => rules::SymbolFill::Reset,
                Some(_) => return Err(Failure::Fault(Fault::Invalid)),
            };
            rules::symbol_fill(
                family.fill_field(),
                family.use_series_fill_field(),
                family.use_series_stroke_field(),
                stored,
            )
            .ok_or(Failure::Refused(Refused::Kind))?
        },
    };
    let rewritten = codec::rewrite(
        payload,
        write.archive(),
        write.writes(),
        write.empty(),
        Limits::DEFAULT,
    )?;
    Ok(rewritten.into_bytes().into_owned())
}

/// Store `values` (`Some`: set; `None`: reset) in the effective styles of
/// the chart's series whose value changes; `None` without a change.
pub(super) fn write(
    package: &Package,
    chart: u64,
    chart_kind: ChartKind,
    kind: Kind,
    values: &[Option<Value>],
) -> Result<Option<Published>> {
    let slots = styles(package, chart, values.len())?;
    let mut replacements: Vec<(u64, usize, Vec<u8>)> = Vec::new();
    for (&style, value) in slots.iter().zip(values) {
        let (current, defined) = effective(package, style, chart_kind, kind)?;
        let unchanged = match value {
            Some(value) => current == *value,
            None => !defined,
        };
        if unchanged {
            continue;
        }
        if !exclusive(package, style)? {
            return Err(Failure::Refused(Refused::Shared));
        }
        if replacements.iter().any(|(existing, ..)| *existing == style) {
            return Err(Failure::Refused(Refused::Shared));
        }
        let (index, payload) = style_payload(package, style)?;
        let data = rewrite(payload, chart_kind, kind, value.as_ref())?;
        if data != payload {
            structure::reserve(&mut replacements, 1)?;
            replacements.push((style, index, data));
        }
    }
    if replacements.is_empty() {
        return Ok(None);
    }
    let message_type = Archive::SeriesStyle.message_type();
    let mut closures = Vec::new();
    structure::reserve(&mut closures, replacements.len())?;
    for (style, index, data) in replacements {
        closures.push((style, rewriter(index, message_type, data)));
    }
    let mut targets: Vec<(u64, &mut StyleRewrite<'_>)> = Vec::new();
    structure::reserve(&mut targets, closures.len())?;
    for (style, closure) in &mut closures {
        targets.push((*style, closure));
    }
    let published = graph::publish_in_place(package, &mut targets)?;
    Ok(Some(published))
}
