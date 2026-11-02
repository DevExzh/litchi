//! The per-series properties stored in a chart's series styles: strokes,
//! fills, data-symbol fills and data-symbol outlines (records 0375a, 0378,
//! 0414).
//!
//! The value of a series is read through its effective style (the private
//! override of `series_private_styles`, else the theme style) and that
//! style's `TSS.StyleArchive.parent` chain; a write rewrites the effective
//! style in place, which only one chart may use (the host's
//! `ensure_exclusive`; a shared style is refused). Nothing is allocated and
//! the package metadata does not change. These are the host editors'
//! `set_*_chart_series_{strokes, fills, symbol_fills, symbol_outlines}` and
//! `reset_*` routes on W1-E's series-style write rules
//! (`chart_style_codec::property`).
//!
//! Two layers: [`effective`] and [`rewrite`] work on payloads for an owner
//! that publishes its own in-place rewrites; [`read()`] and [`write()`] run the
//! same rules over a [`Workspace`] under a residual [`Limits`] profile and
//! return the source-derived mutation scope.

use litchi_iwa_common::chart::kind::Kind as ChartKind;
use litchi_iwa_common::chart::series::{
    FillFamily, Stroke, StrokeFamily, SymbolFill, SymbolFillFamily, SymbolOutlineFamily,
};
use litchi_iwa_common::shape::fill::Fill;
use litchi_iwa_common::shape::shadow::{
    Angle as ShadowAngle, Appearance, BlurRadius, Drop, Offset, Opacity, Shadow,
};
use litchi_iwa_protos::chart_style_codec::{
    self as codec, Archive, Limits as CodecLimits, Properties, property as rules,
};
use litchi_iwa_protos::shape_style_codec::native as stored;

use super::{
    Limits, Refusal, Report, SERIES_STYLE_MESSAGE_TYPE, SymbolBudget, SymbolError, SymbolResult,
    bounded_mutation_scope, bounded_payload, bounded_style_owner_count, bounded_style_slots,
    charged_patch,
};
use crate::drawing;
use crate::graph::{Error, Workspace};
use crate::shape::property::{encode_shadow, shadow_of};

/// The longest style parent chain followed.
const MAX_DEPTH: usize = 64;

/// A property of the series styles.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Kind {
    /// The series stroke (line, area, radar, scatter and bar outlines).
    Strokes,
    /// The series fill.
    Fills,
    /// The fill of the series' data symbols.
    SymbolFills,
    /// The outline of the series' data symbols.
    SymbolOutlines,
    /// Whether value labels move to fit (each series style's own field
    /// 100, "show labels in front").
    ValueLabelAutoFits,
    /// Whether error bars are spaced automatically (each series style's
    /// own field 98).
    ErrorBarAutoFits,
    /// Where bar and column value labels sit (each series style's own
    /// field 88, or 97 for the stacked kinds).
    ValueLabelLocations,
}

/// Whether a series lets iWork fit its value labels or error bars.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
#[non_exhaustive]
pub enum AutoFit {
    /// iWork fits them (the native default).
    #[default]
    Enabled,
    /// Fixed placement.
    Disabled,
}

/// Where a bar or column series draws its value labels.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
#[non_exhaustive]
pub enum ValueLabelLocation {
    /// Inside the base of the bar.
    Bottom,
    /// Inside the middle of the bar.
    Middle,
    /// Inside the end of the bar (the native default).
    #[default]
    Top,
    /// Past the end of the bar.
    Outside,
}

impl ValueLabelLocation {
    const fn native(self) -> u64 {
        match self {
            Self::Bottom => 8,
            Self::Middle => 0,
            Self::Top => 4,
            Self::Outside => 6,
        }
    }

    fn from_native(value: u64) -> SymbolResult<Self> {
        Ok(match value {
            8 => Self::Bottom,
            0 => Self::Middle,
            4 => Self::Top,
            6 => Self::Outside,
            _ => return Err(SymbolError::Invalid("an unsupported value-label location")),
        })
    }
}

/// One series' value of a series-style property.
#[derive(Debug, Clone, PartialEq)]
#[non_exhaustive]
pub enum Value {
    /// A stroke ([`Kind::Strokes`], [`Kind::SymbolOutlines`]); `None` is no
    /// stroke.
    Stroke(Option<Stroke>),
    /// A fill ([`Kind::Fills`]).
    Fill(Fill),
    /// A data-symbol fill ([`Kind::SymbolFills`]).
    SymbolFill(SymbolFill),
    /// An auto-fit switch ([`Kind::ValueLabelAutoFits`],
    /// [`Kind::ErrorBarAutoFits`]).
    AutoFit(AutoFit),
    /// A value-label location ([`Kind::ValueLabelLocations`]).
    ValueLabelLocation(ValueLabelLocation),
}

/// `tschchartseriesdefaultshowlabelsinfront`.
const SHOW_LABELS_IN_FRONT_FIELD: u32 = 100;
/// `tschchartseriesdefaulterrorbarspacing`.
const ERROR_BAR_SPACING_FIELD: u32 = 98;

/// The value-label location field of a chart kind (the host's
/// `LocationStorage`).
const fn location_field(chart: ChartKind) -> Option<u32> {
    match chart {
        ChartKind::Column2d
        | ChartKind::Column3d
        | ChartKind::Bar2d
        | ChartKind::Bar3d
        | ChartKind::MultiDataColumn2d
        | ChartKind::MultiDataBar2d => Some(88),
        ChartKind::StackedColumn2d
        | ChartKind::StackedColumn3d
        | ChartKind::StackedBar2d
        | ChartKind::StackedBar3d => Some(97),
        _ => None,
    }
}

/// A singular canonical varint of the extension: its value and whether it
/// is a canonical boolean, as the host's strict readers check.
fn strict_varint(extension: &[u8], field: u32) -> SymbolResult<Option<(u64, usize)>> {
    let mut found = None;
    for candidate in
        litchi_iwa_common::wire::WireView::parse(extension, litchi_iwa_common::WireLimits::DEFAULT)
            .map_err(invalid)?
            .fields()
            .filter(|candidate| candidate.number() == field)
    {
        if found.is_some() || candidate.wire_type() != 0 {
            return Err(SymbolError::Invalid("a malformed series-style switch"));
        }
        let raw = candidate.payload();
        let (value, consumed) =
            litchi_iwa_common::varint::decode_varint_from_bytes(raw).map_err(invalid)?;
        if consumed != raw.len()
            || litchi_iwa_common::varint::encode_varint(value).len() != consumed
        {
            return Err(SymbolError::Invalid("a non-canonical series-style switch"));
        }
        found = Some((value, consumed));
    }
    Ok(found)
}

/// The value and definedness of a switch kind in one style's own payload
/// (the host reads the slot itself, not its parent chain).
fn own_switch(payload: &[u8], chart: ChartKind, kind: Kind) -> SymbolResult<(Value, bool)> {
    let properties = decode(payload)?;
    let extension = properties.extension().unwrap_or_default();
    Ok(match kind {
        Kind::ValueLabelAutoFits => {
            let stored = strict_varint(extension, SHOW_LABELS_IN_FRONT_FIELD)?;
            if stored.is_some_and(|(value, consumed)| consumed != 1 || value > 1) {
                return Err(SymbolError::Invalid("a non-canonical series-style boolean"));
            }
            let fit = if stored.is_some_and(|(value, _)| value == 1) {
                AutoFit::Disabled
            } else {
                AutoFit::Enabled
            };
            (Value::AutoFit(fit), stored.is_some())
        },
        Kind::ErrorBarAutoFits => {
            let stored = strict_varint(extension, ERROR_BAR_SPACING_FIELD)?;
            let fit = match stored {
                None | Some((0, _)) => AutoFit::Enabled,
                Some((1, _)) => AutoFit::Disabled,
                Some(_) => {
                    return Err(SymbolError::Invalid("an unsupported error-bar spacing"));
                },
            };
            (Value::AutoFit(fit), stored.is_some())
        },
        Kind::ValueLabelLocations => {
            let field =
                location_field(chart).ok_or(SymbolError::Refused(Refusal::UnsupportedKind))?;
            let stored = strict_varint(extension, field)?;
            let location = stored
                .map(|(value, _)| ValueLabelLocation::from_native(value))
                .transpose()?
                .unwrap_or_default();
            (Value::ValueLabelLocation(location), stored.is_some())
        },
        _ => return Err(SymbolError::Invalid("not a series-style switch")),
    })
}

/// The payload storing switch `value` (`None`: the default), as the
/// host's `patch_auto_fit` / `patch_location`: the default removes the
/// field; the auto-fit switches also remove an emptied extension.
fn rewrite_switch(
    payload: &[u8],
    chart: ChartKind,
    kind: Kind,
    value: Option<&Value>,
) -> SymbolResult<Vec<u8>> {
    let wire_error = |error| SymbolError::Graph(Error::Wire(error));
    let (field, stored, remove_empty) = match (kind, value) {
        (Kind::ValueLabelAutoFits, Some(Value::AutoFit(fit))) => (
            SHOW_LABELS_IN_FRONT_FIELD,
            (*fit == AutoFit::Disabled).then_some(1),
            true,
        ),
        (Kind::ErrorBarAutoFits, Some(Value::AutoFit(fit))) => (
            ERROR_BAR_SPACING_FIELD,
            (*fit == AutoFit::Disabled).then_some(1),
            true,
        ),
        (Kind::ValueLabelAutoFits, None) => (SHOW_LABELS_IN_FRONT_FIELD, None, true),
        (Kind::ErrorBarAutoFits, None) => (ERROR_BAR_SPACING_FIELD, None, true),
        (Kind::ValueLabelLocations, value) => {
            let field =
                location_field(chart).ok_or(SymbolError::Refused(Refusal::UnsupportedKind))?;
            let location = match value {
                Some(Value::ValueLabelLocation(location)) => *location,
                None => ValueLabelLocation::Top,
                Some(_) => return Err(SymbolError::Invalid("a value of another property")),
            };
            (
                field,
                (location != ValueLabelLocation::Top).then_some(location.native()),
                false,
            )
        },
        _ => return Err(SymbolError::Invalid("a value of another property")),
    };
    let properties = decode(payload)?;
    let existing = properties.extension();
    let Some(extension) = existing else {
        let Some(stored) = stored else {
            return Ok(payload.to_vec());
        };
        let mut extension = Vec::new();
        litchi_iwa_common::wire::append_varint_field(&mut extension, field, stored)
            .map_err(wire_error)?;
        return litchi_iwa_common::wire::patch_length_delimited_field(
            payload,
            codec::EXTENSION_FIELD,
            false,
            Some(&extension),
        )
        .map_err(wire_error);
    };
    // Validate the stored switch as the host's strict reader does.
    let _ = own_switch(payload, chart, kind)?;
    let present = strict_varint(extension, field)?.is_some();
    let extension = litchi_iwa_common::wire::patch_varint_field(extension, field, present, stored)
        .map_err(wire_error)?;
    litchi_iwa_common::wire::patch_length_delimited_field(
        payload,
        codec::EXTENSION_FIELD,
        true,
        (!(remove_empty && extension.is_empty())).then_some(extension.as_slice()),
    )
    .map_err(wire_error)
}

/// Whether `kind` is a switch read from each style's own payload.
const fn is_switch(kind: Kind) -> bool {
    matches!(
        kind,
        Kind::ValueLabelAutoFits | Kind::ErrorBarAutoFits | Kind::ValueLabelLocations
    )
}

fn invalid<E>(_error: E) -> SymbolError {
    SymbolError::Invalid("a stored series-style value is invalid")
}

fn decode(payload: &[u8]) -> SymbolResult<Properties<'_>> {
    codec::decode(payload, Archive::SeriesStyle, CodecLimits::DEFAULT)
        .map_err(|error| SymbolError::Graph(Error::Codec(error)))
}

fn stroke_at(properties: &Properties<'_>, field: u32) -> SymbolResult<Option<Option<Stroke>>> {
    let Some(bytes) = properties
        .message(field)
        .map_err(|error| SymbolError::Graph(Error::Codec(error)))?
    else {
        return Ok(None);
    };
    let decoded = stored::stroke(bytes, CodecLimits::DEFAULT).map_err(invalid)?;
    let stroke = drawing::stroke(&decoded).map_err(invalid)?;
    Ok(Some(
        stroke
            .map(Stroke::from_shape_stroke)
            .transpose()
            .map_err(invalid)?,
    ))
}

fn fill_at(properties: &Properties<'_>, field: u32) -> SymbolResult<Option<Fill>> {
    let Some(bytes) = properties
        .message(field)
        .map_err(|error| SymbolError::Graph(Error::Codec(error)))?
    else {
        return Ok(None);
    };
    let decoded = stored::fill(bytes, CodecLimits::DEFAULT).map_err(invalid)?;
    Ok(Some(drawing::fill(&decoded).map_err(invalid)?))
}

fn flag(properties: &Properties<'_>, field: u32) -> SymbolResult<Option<bool>> {
    Ok(properties
        .varint(field)
        .map_err(|error| SymbolError::Graph(Error::Codec(error)))?
        .map(|value| value != 0))
}

fn symbol_fill_at(
    properties: &Properties<'_>,
    family: SymbolFillFamily,
) -> SymbolResult<Option<SymbolFill>> {
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
        .map_err(invalid)
}

fn stroke_field(chart: ChartKind, kind: Kind) -> SymbolResult<u32> {
    match kind {
        Kind::Strokes => StrokeFamily::for_chart_kind(chart).map(StrokeFamily::field),
        _ => SymbolOutlineFamily::for_chart_kind(chart).map(SymbolOutlineFamily::field),
    }
    .ok_or(SymbolError::Refused(Refusal::UnsupportedKind))
}

/// Whether `kind` applies to a chart of kind `chart`.
#[must_use]
pub fn applies_to(chart: ChartKind, kind: Kind) -> bool {
    match kind {
        Kind::Strokes => StrokeFamily::for_chart_kind(chart).is_some(),
        Kind::SymbolOutlines => SymbolOutlineFamily::for_chart_kind(chart).is_some(),
        Kind::Fills => FillFamily::for_chart_kind(chart).is_some(),
        Kind::SymbolFills => SymbolFillFamily::for_chart_kind(chart).is_some(),
        Kind::ValueLabelAutoFits | Kind::ErrorBarAutoFits => true,
        Kind::ValueLabelLocations => location_field(chart).is_some(),
    }
}

/// The first value `read` finds up the parent chain of `style`.
fn inherited<'a, T>(
    style: u64,
    payload: &mut impl FnMut(u64) -> SymbolResult<&'a [u8]>,
    read: impl Fn(&Properties<'_>) -> SymbolResult<Option<T>>,
) -> SymbolResult<Option<T>> {
    let mut current = style;
    for _ in 0..MAX_DEPTH {
        let data = payload(current)?;
        if let Some(value) = read(&decode(data)?)? {
            return Ok(Some(value));
        }
        match codec::parent(data, CodecLimits::DEFAULT)
            .map_err(|error| SymbolError::Graph(Error::Codec(error)))?
        {
            Some(parent) => current = parent,
            None => return Ok(None),
        }
    }
    Err(SymbolError::Invalid(
        "a series style parent chain is too deep",
    ))
}

/// The effective value of `kind` in series style `style` of a chart of kind
/// `chart`, and whether `style` defines it itself. `payload` returns the
/// `TSCH.ChartSeriesStyleArchive` payload of a style object.
///
/// # Errors
///
/// [`Refusal::UnsupportedKind`], an invalid style or parent chain.
pub fn effective<'a>(
    style: u64,
    chart: ChartKind,
    kind: Kind,
    mut payload: impl FnMut(u64) -> SymbolResult<&'a [u8]>,
) -> SymbolResult<(Value, bool)> {
    if is_switch(kind) {
        return own_switch(payload(style)?, chart, kind);
    }
    let local = decode(payload(style)?)?;
    Ok(match kind {
        Kind::Strokes | Kind::SymbolOutlines => {
            let field = stroke_field(chart, kind)?;
            let defined = stroke_at(&local, field)?.is_some();
            let value = inherited(style, &mut payload, |p| stroke_at(p, field))?.flatten();
            (Value::Stroke(value), defined)
        },
        Kind::Fills => {
            let family = FillFamily::for_chart_kind(chart)
                .ok_or(SymbolError::Refused(Refusal::UnsupportedKind))?;
            let defined = fill_at(&local, family.field())?.is_some();
            let value = match inherited(style, &mut payload, |p| fill_at(p, family.field()))? {
                Some(fill) => fill,
                None => inherited(style, &mut payload, |p| {
                    fill_at(p, FillFamily::DEFAULT_FIELD)
                })?
                .ok_or(SymbolError::Invalid("a series style chain defines no fill"))?,
            };
            (Value::Fill(value), defined)
        },
        Kind::SymbolFills => {
            let family = SymbolFillFamily::for_chart_kind(chart)
                .ok_or(SymbolError::Refused(Refusal::UnsupportedKind))?;
            let defined = symbol_fill_at(&local, family)?.is_some();
            let value = inherited(style, &mut payload, |p| symbol_fill_at(p, family))?
                .unwrap_or_else(|| family.default_fill());
            (Value::SymbolFill(value), defined)
        },
        Kind::ValueLabelAutoFits | Kind::ErrorBarAutoFits | Kind::ValueLabelLocations => {
            return Err(SymbolError::Invalid(
                "a switch is read from the style itself",
            ));
        },
    })
}

/// The payload of series style `payload` storing `value` (`None`: the
/// style's own definition removed) for property `kind` of a chart of kind
/// `chart`, every other byte kept.
///
/// # Errors
///
/// [`Refusal::UnsupportedKind`], [`Refusal::ImageFill`] (image fills need
/// the embedded-data graph), a value of another property, or an invalid
/// payload.
pub fn rewrite(
    payload: &[u8],
    chart: ChartKind,
    kind: Kind,
    value: Option<&Value>,
) -> SymbolResult<Vec<u8>> {
    if is_switch(kind) {
        return rewrite_switch(payload, chart, kind, value);
    }
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
                Some(_) => return Err(SymbolError::Invalid("a value of another property")),
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
                Some(_) => return Err(SymbolError::Invalid("a value of another property")),
            }
        },
        (Kind::Fills, value) => {
            let family = FillFamily::for_chart_kind(chart)
                .ok_or(SymbolError::Refused(Refusal::UnsupportedKind))?;
            match value {
                Some(Value::Fill(Fill::Image(_))) => {
                    return Err(SymbolError::Refused(Refusal::ImageFill));
                },
                Some(Value::Fill(fill)) => {
                    encoded = drawing::encode_fill(fill);
                    rules::series_message(family.field(), Some(&encoded))
                },
                None => rules::series_message(family.field(), None),
                Some(_) => return Err(SymbolError::Invalid("a value of another property")),
            }
        },
        (Kind::SymbolFills, value) => {
            let family = SymbolFillFamily::for_chart_kind(chart)
                .ok_or(SymbolError::Refused(Refusal::UnsupportedKind))?;
            let stored = match value {
                Some(Value::SymbolFill(SymbolFill::Custom(Fill::Image(_)))) => {
                    return Err(SymbolError::Refused(Refusal::ImageFill));
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
                Some(_) => return Err(SymbolError::Invalid("a value of another property")),
            };
            rules::symbol_fill(
                family.fill_field(),
                family.use_series_fill_field(),
                family.use_series_stroke_field(),
                stored,
            )
            .ok_or(SymbolError::Refused(Refusal::UnsupportedKind))?
        },
        (Kind::ValueLabelAutoFits | Kind::ErrorBarAutoFits | Kind::ValueLabelLocations, _) => {
            return Err(SymbolError::Invalid(
                "a switch is written by rewrite_switch",
            ));
        },
    };
    let rewritten = codec::rewrite(
        payload,
        write.archive(),
        write.writes(),
        write.empty(),
        CodecLimits::DEFAULT,
    )
    .map_err(|error| SymbolError::Graph(Error::Codec(error)))?;
    Ok(rewritten.into_bytes().into_owned())
}

fn workspace_payload<'a>(
    workspace: &'a Workspace<'_>,
    budget: &mut SymbolBudget,
) -> impl FnMut(u64) -> SymbolResult<&'a [u8]> {
    move |style| {
        let (_, payload) = bounded_payload(
            workspace,
            style,
            SERIES_STYLE_MESSAGE_TYPE,
            "a series style must have exactly one payload",
            budget,
        )?;
        budget.source(payload.len())?;
        budget.work(payload.len())?;
        Ok(payload)
    }
}

fn read_effective(
    workspace: &Workspace<'_>,
    chart_kind: ChartKind,
    kind: Kind,
    styles: &[u64],
    budget: &mut SymbolBudget,
) -> SymbolResult<Vec<(Value, bool)>> {
    let mut values = Vec::new();
    budget.reserve_vec(&mut values, styles.len(), false)?;
    for &style in styles {
        values.push(effective(
            style,
            chart_kind,
            kind,
            workspace_payload(workspace, budget),
        )?);
    }
    Ok(values)
}

/// Read the effective value of `kind` for each of the `series_count`
/// series of chart drawable `chart` (of kind `chart_kind`) under one
/// residual profile; the report carries the mutation scope.
///
/// # Errors
///
/// [`Refusal::UnsupportedKind`], an invalid chart or style, or an
/// exhausted limit.
pub fn read(
    workspace: &Workspace<'_>,
    chart: u64,
    chart_kind: ChartKind,
    series_count: usize,
    kind: Kind,
    limits: Limits,
) -> SymbolResult<(Vec<Value>, Report)> {
    if !applies_to(chart_kind, kind) {
        return Err(SymbolError::Refused(Refusal::UnsupportedKind));
    }
    let mut budget = SymbolBudget::new(limits);
    budget.work(workspace.source_work())?;
    let scope = bounded_mutation_scope(workspace, chart, &[], series_count, &mut budget)?;
    budget.set_scope(scope);
    let styles = bounded_style_slots(workspace, chart, series_count, &mut budget)?;
    let values = read_effective(workspace, chart_kind, kind, &styles, &mut budget)?
        .into_iter()
        .map(|(value, _)| value)
        .collect();
    Ok((values, budget.report(false)))
}

/// Store `values` (`Some`: set; `None`: remove the style's own definition)
/// of `kind` in the effective styles of the series whose value changes
/// (`values.len()` is the series count), under one residual profile.
/// Returns whether anything changed and the report with the mutation scope.
///
/// # Errors
///
/// [`Refusal::UnsupportedKind`], [`Refusal::SharedStyle`] (a style another
/// chart uses, or two series sharing one style), [`Refusal::ImageFill`], an
/// invalid chart or style, an exhausted limit, or a write that does not
/// read back.
pub fn write(
    workspace: &mut Workspace<'_>,
    chart: u64,
    chart_kind: ChartKind,
    kind: Kind,
    values: &[Option<Value>],
    limits: Limits,
) -> SymbolResult<(bool, Report)> {
    if values.is_empty() {
        return Err(SymbolError::Invalid("a chart series write names no series"));
    }
    if !applies_to(chart_kind, kind) {
        return Err(SymbolError::Refused(Refusal::UnsupportedKind));
    }
    let mut budget = SymbolBudget::new(limits);
    budget.work(workspace.source_work())?;
    let scope = bounded_mutation_scope(workspace, chart, &[], values.len(), &mut budget)?;
    budget.set_scope(scope);
    let styles = bounded_style_slots(workspace, chart, values.len(), &mut budget)?;
    let current = read_effective(workspace, chart_kind, kind, &styles, &mut budget)?;
    // The host's location write makes every series style exclusive before
    // it rewrites any of them.
    if kind == Kind::ValueLabelLocations
        && current
            .iter()
            .zip(values)
            .any(|((value, _), expected)| expected.as_ref() != Some(value))
    {
        for &style in &styles {
            if bounded_style_owner_count(workspace, style, &mut budget)? != 1 {
                return Err(SymbolError::Refused(Refusal::SharedStyle));
            }
        }
    }
    let mut replacements: Vec<(u64, usize, Vec<u8>)> = Vec::new();
    budget.reserve_vec(&mut replacements, styles.len(), true)?;
    for ((&style, (value, defined)), expected) in styles.iter().zip(&current).zip(values) {
        budget.work(1)?;
        let unchanged = match expected {
            Some(expected) => value == expected,
            None => !defined,
        };
        if unchanged {
            continue;
        }
        if bounded_style_owner_count(workspace, style, &mut budget)? != 1
            || replacements.iter().any(|(existing, ..)| *existing == style)
        {
            return Err(SymbolError::Refused(Refusal::SharedStyle));
        }
        let (message, payload) = bounded_payload(
            workspace,
            style,
            SERIES_STYLE_MESSAGE_TYPE,
            "a series style must have exactly one payload",
            &mut budget,
        )?;
        let data = charged_patch(payload, &mut budget, |payload| {
            rewrite(payload, chart_kind, kind, expected.as_ref())
        })?;
        if data != payload {
            replacements.push((style, message, data));
        }
    }
    let changed = !replacements.is_empty();
    for (style, message, data) in replacements {
        workspace
            .replace_payload(style, message, data)
            .map_err(SymbolError::from)?;
    }
    if changed {
        let actual = read_effective(workspace, chart_kind, kind, &styles, &mut budget)?;
        for ((value, defined), expected) in actual.iter().zip(values) {
            let matches = match expected {
                Some(expected) => value == expected,
                None => !defined,
            };
            if !matches {
                return Err(SymbolError::Invalid(
                    "a series-style write does not read back",
                ));
            }
        }
    }
    Ok((changed, budget.report(changed)))
}

/// `tschchartseriesdefaultshadow` of `TSCH.ChartSeriesStyleArchive`.
const SHADOW_FIELD: u32 = 38;

/// The drop shadow a series style without a stored shadow draws (the
/// host's `ChartShadow::NATIVE_DEFAULT`).
pub const NATIVE_DEFAULT_SHADOW: Drop = Drop::new(
    Appearance::new(
        litchi_iwa_common::color::Rgba::black(),
        BlurRadius::TEN_POINTS,
        Offset::SIX_POINTS,
        Opacity::THREE_QUARTERS,
    ),
    ShadowAngle::FORTY_FIVE_DEGREES,
);

/// The shadow stored by one series style payload (its own field, as the
/// host's series-style slots read it; absent: the native default drop
/// shadow).
///
/// # Errors
///
/// An invalid payload or stored shadow.
pub fn series_shadow(payload: &[u8]) -> SymbolResult<Shadow> {
    match decode(payload)?
        .message(SHADOW_FIELD)
        .map_err(|error| SymbolError::Graph(Error::Codec(error)))?
    {
        None => Ok(Shadow::Drop(NATIVE_DEFAULT_SHADOW)),
        Some(bytes) => shadow_of(bytes).map_err(SymbolError::Graph),
    }
}

/// The series style payload storing `shadow`, every other byte kept (the
/// native default removes the field).
///
/// # Errors
///
/// An invalid payload.
pub fn rewrite_series_shadow(payload: &[u8], shadow: Shadow) -> SymbolResult<Vec<u8>> {
    let encoded = (shadow != Shadow::Drop(NATIVE_DEFAULT_SHADOW)).then(|| encode_shadow(shadow));
    Ok(codec::rewrite(
        payload,
        Archive::SeriesStyle,
        &[codec::Write::Message {
            field: SHADOW_FIELD,
            value: encoded.as_deref(),
        }],
        codec::EmptyExtension::Keep,
        CodecLimits::DEFAULT,
    )
    .map_err(|error| SymbolError::Graph(Error::Codec(error)))?
    .into_bytes()
    .into_owned())
}

/// The uniform shadow of every series style of chart drawable `chart`
/// (theme styles with private overrides, all of them, as the host's
/// `chart_series_style_slots`).
///
/// # Errors
///
/// Nonuniform series shadows, a chart without series styles, an invalid
/// style, or an exhausted limit.
pub(crate) fn chart_shadow(
    workspace: &Workspace<'_>,
    chart: u64,
    budget: &mut SymbolBudget,
) -> SymbolResult<Shadow> {
    let styles = bounded_style_slots(workspace, chart, usize::MAX, budget)?;
    let mut uniform = None;
    for style in styles {
        budget.work(1)?;
        let (_, payload) = bounded_payload(
            workspace,
            style,
            SERIES_STYLE_MESSAGE_TYPE,
            "a series style must have exactly one payload",
            budget,
        )?;
        let shadow = series_shadow(payload)?;
        match uniform {
            None => uniform = Some(shadow),
            Some(known) if known != shadow => {
                return Err(SymbolError::Invalid(
                    "a chart has nonuniform series shadows",
                ));
            },
            Some(_) => {},
        }
    }
    uniform.ok_or(SymbolError::Invalid("the chart has no series styles"))
}

/// Store `shadow` in every series style of chart drawable `chart`, after
/// requiring each to be exclusive to the chart (the host's
/// `ensure_exclusive` on every series slot), admitting each style to the
/// budget's scope. Returns whether a payload changed.
///
/// # Errors
///
/// [`Refusal::SharedStyle`], an invalid style, or an exhausted limit.
pub(crate) fn write_chart_shadow(
    workspace: &mut Workspace<'_>,
    chart: u64,
    shadow: Shadow,
    budget: &mut SymbolBudget,
) -> SymbolResult<bool> {
    let styles = bounded_style_slots(workspace, chart, usize::MAX, budget)?;
    for &style in &styles {
        if bounded_style_owner_count(workspace, style, budget)? != 1 {
            return Err(SymbolError::Refused(Refusal::SharedStyle));
        }
    }
    let mut changed = false;
    for style in styles {
        budget.work(1)?;
        let location = workspace.locate(style).map_err(SymbolError::from)?;
        let member = workspace.member(location).to_owned();
        if let Some(scope) = budget.scope_mut() {
            scope.allow(style, &member, SERIES_STYLE_MESSAGE_TYPE);
        }
        let (message, payload) = bounded_payload(
            workspace,
            style,
            SERIES_STYLE_MESSAGE_TYPE,
            "a series style must have exactly one payload",
            budget,
        )?;
        let data = charged_patch(payload, budget, |payload| {
            rewrite_series_shadow(payload, shadow)
        })?;
        if data != payload {
            workspace
                .replace_payload(style, message, data)
                .map_err(SymbolError::from)?;
            changed = true;
        }
    }
    Ok(changed)
}
