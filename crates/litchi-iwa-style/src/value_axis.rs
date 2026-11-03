//! The bounds, step counts and scale of a chart's primary value axis
//! (record 0422): the inspector values stored in fields 5, 6, 8, 17 and 18
//! of the generated extension (field 10,000) of the axis's
//! `TSCH.ChartAxisNonStyleArchive`.
//!
//! The rules were Keynote's (`slide_chart_value_axis`, record 0372) and are
//! now shared by the three format families:
//!
//! - a stored or requested value is one [`ValueAxisSettings`] aggregate;
//!   a manual endpoint must be finite, the minimum must not exceed the
//!   maximum, and a logarithmic axis must not have a non-positive endpoint
//!   ([`validate`]);
//! - a write goes through the strict source-preserving codec
//!   (`litchi_iwa_protos::keynote_chart_axis_value_settings_codec`): a
//!   component equal to the stored one keeps its bytes, an absent scale
//!   field stays absent while the scale is linear ([`codec_write`]).
//!
//! Two layers, as for the chart style properties:
//!
//! - payload functions ([`settings`], [`from_snapshot`], [`rewrite`]) for
//!   an owner that resolves the axis object and publishes its own in-place
//!   rewrite (Keynote's value-axis transaction);
//! - workspace functions ([`read()`], [`write()`]) that resolve the chart's
//!   primary value-axis non-style object over a [`Workspace`], refuse one
//!   another chart axis shares (the host's `ensure_exclusive`), rewrite it
//!   in place under a residual [`Limits`] profile and report the
//!   source-derived mutation scope.

use litchi_iwa_common::chart::axis::Axis;
use litchi_iwa_common::wire;
use litchi_iwa_protos::chart_style_codec::{self as codec, Archive, Limits as CodecLimits};
use litchi_iwa_protos::keynote_chart_axis_value_settings_codec::{
    self as settings_codec, AxisValueBound, AxisValueBounds, AxisValueSettings,
    AxisValueSettingsSnapshot, AxisValueSettingsWrite, AxisValueSteps, DecodeOptions,
};

pub use litchi_iwa_common::chart::axis::{
    Bound, Bounds, MajorStepCount, MinorStepCount, Scale, Steps, ValueAxisSettings,
};

use crate::chart_style::{in_place_scope, owners, read_payload, rewrite_in_place, style_object};
use crate::graph::{Error, Workspace};
use crate::series::{Limits, Refusal, Report, SymbolBudget, SymbolError, SymbolResult};

/// Settings a native value axis cannot store: a non-finite endpoint, an
/// inverted range, or a non-positive endpoint on a logarithmic axis.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct InvalidSettings;

impl std::fmt::Display for InvalidSettings {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str("the value-axis settings cannot be stored")
    }
}

impl std::error::Error for InvalidSettings {}

impl From<InvalidSettings> for SymbolError {
    fn from(_: InvalidSettings) -> Self {
        Self::Refused(Refusal::InvalidSettings)
    }
}

/// `settings` in the form the native format stores, or [`InvalidSettings`].
///
/// A known scale given as [`Scale::Unsupported`] with its native value is
/// canonicalised to the known variant, so a staged value reads back equal.
///
/// # Errors
///
/// [`InvalidSettings`].
pub fn validate(settings: ValueAxisSettings) -> Result<ValueAxisSettings, InvalidSettings> {
    let minimum = settings.bounds().minimum();
    let maximum = settings.bounds().maximum();
    if minimum.is_some_and(|bound| !bound.value().is_finite())
        || maximum.is_some_and(|bound| !bound.value().is_finite())
        || matches!((minimum, maximum), (Some(low), Some(high)) if low.value() > high.value())
    {
        return Err(InvalidSettings);
    }
    if settings.scale() == Scale::Logarithmic
        && (minimum.is_some_and(|bound| bound.value() <= 0.0)
            || maximum.is_some_and(|bound| bound.value() <= 0.0))
    {
        return Err(InvalidSettings);
    }
    Ok(settings.with_scale(Scale::from_native(settings.scale().native_value())))
}

/// Whether two settings store the same values (scales compared by their
/// native value).
#[must_use]
pub fn equal_exact(left: ValueAxisSettings, right: ValueAxisSettings) -> bool {
    let (left_bounds, right_bounds) = (left.bounds(), right.bounds());
    let endpoint = |a: Option<Bound>, b: Option<Bound>| match (a, b) {
        (Some(a), Some(b)) => a.value() == b.value(),
        (None, None) => true,
        _ => false,
    };
    endpoint(left_bounds.minimum(), right_bounds.minimum())
        && endpoint(left_bounds.maximum(), right_bounds.maximum())
        && left.steps() == right.steps()
        && Scale::from_native(left.scale().native_value())
            == Scale::from_native(right.scale().native_value())
}

/// The settings of a decoded generated extension. A source the native
/// format could not have written (see [`validate`]) is [`InvalidSettings`].
///
/// # Errors
///
/// [`InvalidSettings`].
pub fn from_snapshot(
    snapshot: AxisValueSettingsSnapshot<'_>,
) -> Result<ValueAxisSettings, InvalidSettings> {
    let bounds = snapshot.bounds();
    let endpoint = |bound: Option<AxisValueBound>| {
        bound
            .map(|bound| Bound::new(bound.value()))
            .transpose()
            .map_err(|_| InvalidSettings)
    };
    let bounds = Bounds::new(endpoint(bounds.minimum())?, endpoint(bounds.maximum())?)
        .map_err(|_| InvalidSettings)?;
    let steps = snapshot.steps();
    let major = steps
        .major()
        .map(MajorStepCount::new)
        .transpose()
        .map_err(|_| InvalidSettings)?;
    let minor = steps
        .minor()
        .map(MinorStepCount::new)
        .transpose()
        .map_err(|_| InvalidSettings)?;
    validate(ValueAxisSettings::new(
        bounds,
        Steps::new(major, minor),
        Scale::from_native(snapshot.scale().native_value()),
    ))
}

fn codec_settings(settings: ValueAxisSettings) -> Result<AxisValueSettings, InvalidSettings> {
    let bounds = settings.bounds();
    let endpoint = |bound: Option<Bound>| {
        bound
            .map(|bound| AxisValueBound::new(bound.value()))
            .transpose()
            .map_err(|_| InvalidSettings)
    };
    let bounds = AxisValueBounds::new(endpoint(bounds.minimum())?, endpoint(bounds.maximum())?)
        .map_err(|_| InvalidSettings)?;
    let steps = settings.steps();
    let steps = AxisValueSteps::new(
        steps.major().map(MajorStepCount::value),
        steps.minor().map(MinorStepCount::value),
    )
    .map_err(|_| InvalidSettings)?;
    Ok(AxisValueSettings::new(
        bounds,
        steps,
        settings_codec::Scale::from_native(settings.scale().native_value()),
    ))
}

/// The presence-preserving codec write storing `settings` over `snapshot`:
/// bounds and steps replaced (the codec keeps a component equal to the
/// stored one), the scale written explicitly when the source stores it and
/// otherwise only when it is not linear.
///
/// # Errors
///
/// [`InvalidSettings`].
pub fn codec_write(
    snapshot: AxisValueSettingsSnapshot<'_>,
    settings: ValueAxisSettings,
) -> Result<AxisValueSettingsWrite, InvalidSettings> {
    let native = codec_settings(settings)?;
    let write = AxisValueSettingsWrite::preserve()
        .with_bounds(native.bounds())
        .with_steps(native.steps());
    Ok(if snapshot.scale_present() {
        write.with_explicit_scale(native.scale())
    } else {
        write.with_scale(native.scale())
    })
}

/// The most the five settings fields add to an extension: two bound
/// messages (key, length, key and double: 12 bytes each), two step counts
/// and a scale (key and a ten-byte varint each).
const SETTINGS_GROWTH: usize = 2 * 12 + 3 * 11;

/// Finite codec ceilings for an extension of `length` bytes and a rewrite
/// that adds at most [`SETTINGS_GROWTH`] bytes.
fn codec_options(length: usize) -> DecodeOptions {
    let bytes = length.saturating_add(SETTINGS_GROWTH);
    DecodeOptions::new(bytes, bytes.saturating_mul(16), bytes.saturating_mul(32), 8)
        .with_max_allocations(bytes.saturating_add(8))
        .with_max_retained_bytes(bytes.saturating_mul(2))
        .with_max_scratch_bytes(bytes)
}

fn codec_error<E>(_error: E) -> SymbolError {
    SymbolError::Invalid("an invalid value-axis settings extension")
}

fn extension(payload: &[u8]) -> SymbolResult<Option<&[u8]>> {
    let properties =
        codec::decode(payload, Archive::AxisNonStyle, CodecLimits::DEFAULT).map_err(codec_error)?;
    Ok(properties.extension())
}

/// The settings stored in a `TSCH.ChartAxisNonStyleArchive` payload (no
/// extension: [`ValueAxisSettings::automatic`]).
///
/// # Errors
///
/// An invalid payload or extension; stored settings the native format
/// could not have written.
pub fn settings(payload: &[u8]) -> SymbolResult<ValueAxisSettings> {
    let Some(extension) = extension(payload)? else {
        return Ok(ValueAxisSettings::automatic());
    };
    let snapshot =
        settings_codec::decode_axis_value_settings(extension, codec_options(extension.len()))
            .map_err(codec_error)?;
    from_snapshot(snapshot)
        .map_err(|_| SymbolError::Invalid("stored value-axis settings are invalid"))
}

fn rewrite_extension(source: &[u8], write: AxisValueSettingsWrite) -> SymbolResult<Vec<u8>> {
    settings_codec::rewrite_axis_value_settings(source, write, codec_options(source.len()))
        .map_err(codec_error)
}

/// The writes storing `target` over `stored`, one component at a time in
/// the host's order (bounds, steps, scale; within a component the minimum
/// before the maximum and the major before the minor count), so that
/// fields absent from the source are appended in the order the host's
/// `set_chart_value_axis_{bounds,steps,scale}` append them.
fn ordered_writes(
    stored: AxisValueSettingsSnapshot<'_>,
    target: AxisValueSettings,
) -> Result<Vec<AxisValueSettingsWrite>, InvalidSettings> {
    let preserve = AxisValueSettingsWrite::preserve;
    let mut writes = Vec::with_capacity(5);
    let (from, to) = (stored.bounds(), target.bounds());
    if from.minimum().is_none()
        && from.maximum().is_none()
        && to.minimum().is_some()
        && to.maximum().is_some()
    {
        let minimum = AxisValueBounds::new(to.minimum(), None).map_err(|_| InvalidSettings)?;
        writes.push(preserve().with_bounds(minimum));
    }
    writes.push(preserve().with_bounds(to));
    let (from, to) = (stored.steps(), target.steps());
    if from.major().is_none()
        && from.minor().is_none()
        && to.major().is_some()
        && to.minor().is_some()
    {
        let major = AxisValueSteps::new(to.major(), None).map_err(|_| InvalidSettings)?;
        writes.push(preserve().with_steps(major));
    }
    writes.push(preserve().with_steps(to));
    writes.push(if stored.scale_present() {
        preserve().with_explicit_scale(target.scale())
    } else {
        preserve().with_scale(target.scale())
    });
    Ok(writes)
}

/// The `TSCH.ChartAxisNonStyleArchive` payload storing `value`, every
/// other byte kept (an absent extension is appended). A field the source
/// stores is rewritten in place; an absent one is appended in the host's
/// order.
///
/// # Errors
///
/// [`Refusal::InvalidSettings`], an invalid payload or extension.
pub fn rewrite(payload: &[u8], value: ValueAxisSettings) -> SymbolResult<Vec<u8>> {
    let value = validate(value)?;
    let existing = extension(payload)?;
    let source = existing.unwrap_or_default();
    let snapshot = settings_codec::decode_axis_value_settings(source, codec_options(source.len()))
        .map_err(codec_error)?;
    let mut replacement = source.to_vec();
    for write in ordered_writes(snapshot, codec_settings(value)?)? {
        replacement = rewrite_extension(&replacement, write)?;
    }
    if existing.is_some_and(|existing| existing == replacement.as_slice()) {
        return Ok(payload.to_vec());
    }
    wire::patch_length_delimited_field(
        payload,
        codec::EXTENSION_FIELD,
        existing.is_some(),
        Some(replacement.as_slice()),
    )
    .map_err(|error| SymbolError::Graph(Error::Wire(error)))
}

fn value(
    workspace: &Workspace<'_>,
    chart: u64,
    budget: &mut SymbolBudget,
) -> SymbolResult<ValueAxisSettings> {
    let object = style_object(
        workspace,
        chart,
        Archive::AxisNonStyle,
        Some(Axis::Value),
        budget,
    )?;
    read_payload(
        workspace,
        object,
        Archive::AxisNonStyle.message_type(),
        budget,
        settings,
    )
}

/// Read the value-axis settings of chart drawable `chart` under one
/// residual profile.
///
/// # Errors
///
/// [`Refusal::UnsupportedAxis`] (no primary value axis), an invalid chart
/// or axis object, or an exhausted limit.
pub fn read(
    workspace: &Workspace<'_>,
    chart: u64,
    limits: Limits,
) -> SymbolResult<(ValueAxisSettings, Report)> {
    let mut budget = SymbolBudget::new(limits);
    let settings = value(workspace, chart, &mut budget)?;
    Ok((settings, budget.report(false)))
}

/// Store `settings` on the primary value axis of chart drawable `chart`
/// under one residual profile: the axis non-style object is rewritten in
/// place when only this chart's axes reference it, as the host's
/// `set_chart_value_axis_{bounds,steps,scale}` do. Returns whether anything
/// changed and the report with the in-place mutation scope.
///
/// # Errors
///
/// [`Refusal::InvalidSettings`], [`Refusal::UnsupportedAxis`],
/// [`Refusal::SharedStyle`], an invalid chart or axis object, an exhausted
/// limit, or a write that does not read back.
pub fn write(
    workspace: &mut Workspace<'_>,
    chart: u64,
    settings: ValueAxisSettings,
    limits: Limits,
) -> SymbolResult<(bool, Report)> {
    let settings = validate(settings)?;
    let mut budget = SymbolBudget::new(limits);
    if equal_exact(value(workspace, chart, &mut budget)?, settings) {
        return Ok((false, budget.report(false)));
    }
    let object = style_object(
        workspace,
        chart,
        Archive::AxisNonStyle,
        Some(Axis::Value),
        &mut budget,
    )?;
    if owners(workspace, object, Archive::AxisNonStyle, &mut budget)? != 1 {
        return Err(SymbolError::Refused(Refusal::SharedStyle));
    }
    let message_type = Archive::AxisNonStyle.message_type();
    in_place_scope(workspace, chart, object, message_type, &[], &mut budget)?;
    let changed = rewrite_in_place(workspace, object, message_type, &mut budget, |payload| {
        rewrite(payload, settings)
    })?;
    if changed && !equal_exact(value(workspace, chart, &mut budget)?, settings) {
        return Err(SymbolError::Invalid(
            "a value-axis settings write does not read back",
        ));
    }
    Ok((changed, budget.report(changed)))
}
