//! The legend-style variation graph of a chart and its fill, stroke and
//! shadow (record 0414): the host's `charts::legend_style` slot
//! (`ensure_exclusive`, `update`, `collapse_if_equivalent`) and its
//! `legend_{fill, stroke, shadow}` properties over a [`Workspace`].
//!
//! A legend override lives in the generated extension (field 10,000) of a
//! private `TSCH.LegendStyleArchive` variation:
//!
//! 1. when the chart's legend style is not a variation the chart alone
//!    references, a variation is derived: a copy of the style object under a
//!    new identifier whose `TSS.StyleArchive` names the old style as parent,
//!    is marked a variation and names the stylesheet; it is registered with
//!    the stylesheet (`TSS.StylesheetArchive` styles and parent-children),
//!    placed beside the old style, given an object UUID and the external
//!    references the host adds, and the chart's `legend_style` points to it;
//! 2. the property is written into the variation's extension;
//! 3. when the variation then equals its parent except for the
//!    `TSS.StyleArchive`, it collapses: the chart points back to the parent,
//!    the variation leaves the stylesheet, the registry and the package, and
//!    its identifier suffix is released.
//!
//! Image fills move data references and media and are refused
//! ([`Refusal::ImageFill`]).

use litchi_iwa_common::shape::fill::Fill;
use litchi_iwa_common::shape::shadow::{Drop, Shadow};
use litchi_iwa_common::shape::stroke::Stroke;
use litchi_iwa_common::wire;
use litchi_iwa_protos::chart_style_codec::{
    self as codec, Archive, Limits as CodecLimits, Write, merge,
};
use litchi_iwa_protos::shape_style_codec::native as stored;

use crate::drawing;
use crate::graph::{Error, Workspace};
use crate::series::{
    Limits, MutationScope, Refusal, Report, SymbolBudget, SymbolError, SymbolResult,
    bounded_chart_payload, bounded_payload, bounded_references, charged_patch,
};
use crate::shape::property::{encode_shadow, shadow_of};
use crate::shape::{register, unregister};

const CHART_MESSAGE_TYPE: u32 = 5_021;
const LEGEND_STYLE_MESSAGE_TYPE: u32 = 5_024;
const STYLESHEET_MESSAGE_TYPE: u32 = 401;
const PACKAGE_METADATA_MESSAGE_TYPE: u32 = 11_006;
const METADATA_MEMBERS: [&str; 2] = ["Index/Metadata.iwa", "Index/DocumentMetadata.iwa"];

/// `TSCH.ChartArchive.legend_style`, in the chart drawable's extension.
const CHART_LEGEND_STYLE_FIELD: u32 = 11;
/// `TSCH.LegendStyleArchive.super`.
const STYLE_SUPER_FIELD: u32 = 1;
/// `TSS.StyleArchive.{parent, is_variation, stylesheet}`.
const STYLE_PARENT_FIELD: u32 = 3;
const STYLE_VARIATION_FIELD: u32 = 4;
const STYLE_STYLESHEET_FIELD: u32 = 5;
const REFERENCE_IDENTIFIER_FIELD: u32 = 1;

/// The generated legend-style fields.
const FILL_FIELD: u32 = 1;
const SHADOW_FIELD: u32 = 4;
const STROKE_FIELD: u32 = 5;

/// The legend fill override.
#[derive(Debug, Clone, PartialEq, Default)]
#[non_exhaustive]
pub enum LegendFill {
    /// No override: the legend style's parent chain decides.
    #[default]
    Inherited,
    /// An explicit fill (`Fill::None` is the inspector's "No Fill").
    Fill(Fill),
}

/// The legend border override.
#[derive(Debug, Clone, Copy, PartialEq, Default)]
#[non_exhaustive]
pub enum LegendStroke {
    /// No override.
    #[default]
    Inherited,
    /// The inspector's "None" border.
    None,
    /// An explicit border.
    Stroke(Stroke),
}

/// The legend shadow override.
#[derive(Debug, Clone, Copy, PartialEq, Default)]
#[non_exhaustive]
pub enum LegendShadow {
    /// No override.
    #[default]
    Inherited,
    /// The shadow is switched off.
    None,
    /// A drop shadow.
    Shadow(Drop),
}

/// One legend-style property with its value.
#[derive(Debug, Clone, PartialEq)]
#[non_exhaustive]
pub enum Property {
    /// The legend fill.
    Fill(LegendFill),
    /// The legend border.
    Stroke(LegendStroke),
    /// The legend shadow.
    Shadow(LegendShadow),
}

/// The property a [`Property`] value sets, for reads and resets.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum PropertyKind {
    /// [`Property::Fill`].
    Fill,
    /// [`Property::Stroke`].
    Stroke,
    /// [`Property::Shadow`].
    Shadow,
}

impl Property {
    /// The property this value sets.
    #[must_use]
    pub const fn kind(&self) -> PropertyKind {
        match self {
            Self::Fill(_) => PropertyKind::Fill,
            Self::Stroke(_) => PropertyKind::Stroke,
            Self::Shadow(_) => PropertyKind::Shadow,
        }
    }
}

impl PropertyKind {
    /// The inherited value (what a reset restores).
    #[must_use]
    pub const fn inherited(self) -> Property {
        match self {
            Self::Fill => Property::Fill(LegendFill::Inherited),
            Self::Stroke => Property::Stroke(LegendStroke::Inherited),
            Self::Shadow => Property::Shadow(LegendShadow::Inherited),
        }
    }

    const fn field(self) -> u32 {
        match self {
            Self::Fill => FILL_FIELD,
            Self::Stroke => STROKE_FIELD,
            Self::Shadow => SHADOW_FIELD,
        }
    }
}

fn codec_error(error: codec::Error) -> SymbolError {
    SymbolError::Graph(Error::Codec(error))
}

fn invalid<E>(_error: E) -> SymbolError {
    SymbolError::Invalid("a stored legend-style value is invalid")
}

fn wire_error(error: litchi_iwa_common::Error) -> SymbolError {
    SymbolError::Graph(Error::Wire(error))
}

/// The value of `kind` a legend-style payload stores itself.
///
/// # Errors
///
/// An invalid payload or stored value.
pub fn value(payload: &[u8], kind: PropertyKind) -> SymbolResult<Property> {
    let properties =
        codec::decode(payload, Archive::LegendStyle, CodecLimits::DEFAULT).map_err(codec_error)?;
    let stored = properties.message(kind.field()).map_err(codec_error)?;
    Ok(match kind {
        PropertyKind::Fill => Property::Fill(match stored {
            None => LegendFill::Inherited,
            Some(bytes) => LegendFill::Fill(
                drawing::fill(&stored::fill(bytes, CodecLimits::DEFAULT).map_err(invalid)?)
                    .map_err(invalid)?,
            ),
        }),
        PropertyKind::Stroke => Property::Stroke(match stored {
            None => LegendStroke::Inherited,
            Some(bytes) => match drawing::stroke(
                &stored::stroke(bytes, CodecLimits::DEFAULT).map_err(invalid)?,
            )
            .map_err(invalid)?
            {
                Some(stroke) => LegendStroke::Stroke(stroke),
                None => LegendStroke::None,
            },
        }),
        PropertyKind::Shadow => Property::Shadow(match stored {
            None => LegendShadow::Inherited,
            Some(bytes) => match shadow_of(bytes).map_err(SymbolError::Graph)? {
                Shadow::Drop(drop) => LegendShadow::Shadow(drop),
                Shadow::Disabled => LegendShadow::None,
                _ => {
                    return Err(SymbolError::Invalid(
                        "a chart legend uses a non-drop shadow",
                    ));
                },
            },
        }),
    })
}

/// The legend-style payload storing `value` itself, every other byte kept.
///
/// # Errors
///
/// [`Refusal::ImageFill`], an invalid payload.
pub fn rewrite(payload: &[u8], value: &Property) -> SymbolResult<Vec<u8>> {
    let field = value.kind().field();
    let encoded: Option<Vec<u8>>;
    let (write, empty) = match value {
        Property::Fill(fill) => {
            encoded = match fill {
                LegendFill::Inherited => None,
                LegendFill::Fill(Fill::Image(_)) => {
                    return Err(SymbolError::Refused(Refusal::ImageFill));
                },
                LegendFill::Fill(fill) => Some(drawing::encode_fill(fill)),
            };
            (
                Write::Message {
                    field,
                    value: encoded.as_deref(),
                },
                codec::EmptyExtension::Keep,
            )
        },
        Property::Stroke(stroke) => {
            encoded = match stroke {
                LegendStroke::Inherited => None,
                LegendStroke::None => Some(drawing::encode_stroke(None)),
                LegendStroke::Stroke(stroke) => Some(drawing::encode_stroke(Some(*stroke))),
            };
            (
                Write::Message {
                    field,
                    value: encoded.as_deref(),
                },
                codec::EmptyExtension::Keep,
            )
        },
        Property::Shadow(shadow) => {
            encoded = match shadow {
                LegendShadow::Inherited => None,
                LegendShadow::None => Some(encode_shadow(Shadow::Disabled)),
                LegendShadow::Shadow(drop) => Some(encode_shadow(Shadow::Drop(*drop))),
            };
            let write = match encoded.as_deref() {
                Some(value) => Write::MergeMessage {
                    field,
                    value,
                    schema: merge::schema::SHADOW,
                },
                None => Write::Message { field, value: None },
            };
            (write, codec::EmptyExtension::Remove)
        },
    };
    Ok(codec::rewrite(
        payload,
        Archive::LegendStyle,
        &[write],
        empty,
        CodecLimits::DEFAULT,
    )
    .map_err(codec_error)?
    .into_bytes()
    .into_owned())
}

/// The parts of a legend style's `TSS.StyleArchive`.
struct StyleParts {
    parent: Option<u64>,
    variation: bool,
    stylesheet: Option<u64>,
}

fn reference_identifier(reference: &[u8]) -> SymbolResult<u64> {
    crate::graph::varint(reference, REFERENCE_IDENTIFIER_FIELD)
        .map_err(SymbolError::Graph)?
        .ok_or(SymbolError::Invalid("a style reference has no identifier"))
}

fn style_parts(payload: &[u8]) -> SymbolResult<StyleParts> {
    let style = crate::graph::repeated(payload, STYLE_SUPER_FIELD)
        .map_err(SymbolError::Graph)?
        .last()
        .copied()
        .ok_or(SymbolError::Invalid("a legend style has no style archive"))?;
    let reference = |field: u32| -> SymbolResult<Option<u64>> {
        crate::graph::repeated(style, field)
            .map_err(SymbolError::Graph)?
            .last()
            .map(|reference| reference_identifier(reference))
            .transpose()
            .map(|identifier| identifier.filter(|identifier| *identifier != 0))
    };
    Ok(StyleParts {
        parent: reference(STYLE_PARENT_FIELD)?,
        variation: crate::graph::varint(style, STYLE_VARIATION_FIELD)
            .map_err(SymbolError::Graph)?
            == Some(1),
        stylesheet: reference(STYLE_STYLESHEET_FIELD)?,
    })
}

fn reference(identifier: u64) -> Vec<u8> {
    let mut encoded = Vec::new();
    crate::graph::push_varint_field(&mut encoded, REFERENCE_IDENTIFIER_FIELD, identifier);
    encoded
}

/// The host's `make_style_variation`: the style archive names `parent`,
/// is a variation and names `stylesheet`.
fn variation_payload(payload: &[u8], parent: u64, stylesheet: u64) -> SymbolResult<Vec<u8>> {
    wire::transform_length_delimited_field(payload, STYLE_SUPER_FIELD, |style| {
        let has = |number: u32| -> litchi_iwa_common::Result<bool> {
            Ok(wire::parse_wire_fields(style, wire_limits())?
                .iter()
                .any(|field| field.number() == number))
        };
        let style_present = (
            has(STYLE_PARENT_FIELD)?,
            has(STYLE_VARIATION_FIELD)?,
            has(STYLE_STYLESHEET_FIELD)?,
        );
        let style = wire::patch_length_delimited_field(
            style,
            STYLE_PARENT_FIELD,
            style_present.0,
            Some(&reference(parent)),
        )?;
        let style =
            wire::patch_varint_field(&style, STYLE_VARIATION_FIELD, style_present.1, Some(1))?;
        wire::patch_length_delimited_field(
            &style,
            STYLE_STYLESHEET_FIELD,
            style_present.2,
            Some(&reference(stylesheet)),
        )
    })
    .map_err(wire_error)
}

fn wire_limits() -> litchi_iwa_common::WireLimits {
    litchi_iwa_common::WireLimits::default()
}

/// A payload without its `TSS.StyleArchive` (the host's
/// `without_style_archive`).
fn without_style(payload: &[u8]) -> SymbolResult<Vec<u8>> {
    wire::patch_length_delimited_field(payload, STYLE_SUPER_FIELD, true, None).map_err(wire_error)
}

fn legend_style_of(
    workspace: &Workspace<'_>,
    chart: u64,
    budget: &mut SymbolBudget,
) -> SymbolResult<u64> {
    let (_, payload) = bounded_chart_payload(workspace, chart, budget)?;
    bounded_references(payload, budget)?
        .legend_style()
        .ok_or(SymbolError::Refused(Refusal::MissingStyle))
}

fn style_payload<'w>(
    workspace: &'w Workspace<'_>,
    style: u64,
    budget: &mut SymbolBudget,
) -> SymbolResult<(usize, &'w [u8])> {
    let (message, payload) = bounded_payload(
        workspace,
        style,
        LEGEND_STYLE_MESSAGE_TYPE,
        "a legend style must have exactly one payload",
        budget,
    )?;
    budget.source(payload.len())?;
    budget.work(payload.len())?;
    Ok((message, payload))
}

/// How many charts name `style` as their legend style.
fn owners(workspace: &Workspace<'_>, style: u64, budget: &mut SymbolBudget) -> SymbolResult<usize> {
    let mut count = 0_usize;
    for location in workspace.locations_iter() {
        for object in &workspace
            .archive_at(location)
            .map_err(SymbolError::from)?
            .objects
        {
            for message in &object.messages {
                if message.type_ == CHART_MESSAGE_TYPE
                    && bounded_references(&message.data, budget)?.legend_style() == Some(style)
                {
                    count = count.saturating_add(1);
                }
            }
        }
    }
    Ok(count)
}

/// Point the chart's `legend_style` (payload and metadata) from `from` to
/// `to` (the host's `patch_chart_legend_style`).
fn repoint(workspace: &mut Workspace<'_>, chart: u64, from: u64, to: u64) -> SymbolResult<()> {
    let (message, payload) = workspace
        .unique_payload(
            chart,
            CHART_MESSAGE_TYPE,
            "a chart must have exactly one chart payload",
        )
        .map_err(SymbolError::from)?;
    if codec::references(payload, CodecLimits::DEFAULT)
        .map_err(codec_error)?
        .legend_style()
        != Some(from)
    {
        return Err(SymbolError::Invalid(
            "a chart's legend style changed unexpectedly",
        ));
    }
    let data = wire::transform_length_delimited_field(payload, codec::EXTENSION_FIELD, |chart| {
        wire::transform_length_delimited_field(chart, CHART_LEGEND_STYLE_FIELD, |reference| {
            wire::patch_varint_field(reference, REFERENCE_IDENTIFIER_FIELD, true, Some(to))
        })
    })
    .map_err(wire_error)?;
    workspace
        .replace_payload(chart, message, data)
        .map_err(SymbolError::from)?;
    let info = workspace
        .object_mut(chart)
        .map_err(SymbolError::from)?
        .archive_info
        .message_infos
        .get_mut(message)
        .ok_or(SymbolError::Invalid("a chart payload has no metadata"))?;
    let mut replaced = 0_usize;
    for reference in &mut info.object_references {
        if *reference == from {
            *reference = to;
            replaced += 1;
        }
    }
    for field in &mut info.field_infos {
        for reference in &mut field.object_references {
            if *reference == from {
                *reference = to;
            }
        }
    }
    if replaced != 1 {
        return Err(SymbolError::Invalid(
            "a chart's metadata must reference its legend style once",
        ));
    }
    Ok(())
}

/// The component of the member holding `object`, if registered.
fn component_of(workspace: &Workspace<'_>, object: u64) -> SymbolResult<Option<u64>> {
    let location = workspace.locate(object).map_err(SymbolError::from)?;
    let member = workspace.member(location).to_owned();
    workspace
        .component_for_member(&member)
        .map_err(SymbolError::from)
}

/// Make `style` a variation only `chart` uses, deriving one when needed;
/// returns the style to write.
fn ensure_exclusive(
    workspace: &mut Workspace<'_>,
    chart: u64,
    style: u64,
    budget: &mut SymbolBudget,
) -> SymbolResult<u64> {
    let (message, payload) = style_payload(workspace, style, budget)?;
    let parts = style_parts(payload)?;
    let stylesheet = parts.stylesheet.ok_or(SymbolError::Invalid(
        "a chart legend style has no stylesheet",
    ))?;
    if parts.variation && owners(workspace, style, budget)? == 1 {
        return Ok(style);
    }
    let data = variation_payload(payload, style, stylesheet)?;
    let identifier = workspace
        .next_object_identifier()
        .map_err(SymbolError::from)?;
    budget.identifiers(1)?;
    budget.preflight_allocations(2)?;
    budget.preflight_retained(data.len())?;
    budget.allocations(2)?;
    budget.retained(data.len())?;
    let mut object = workspace.object(style).map_err(SymbolError::from)?.clone();
    object.archive_info.identifier = Some(identifier);
    object
        .replace_message_preserving_header(
            message,
            litchi_iwa_core::RawMessage {
                type_: LEGEND_STYLE_MESSAGE_TYPE,
                data,
            },
            workspace.limits(),
        )
        .map_err(|error| SymbolError::Graph(Error::Archive(error)))?;
    let info = object
        .archive_info
        .message_infos
        .get_mut(message)
        .ok_or(SymbolError::Invalid(
            "a legend style payload has no metadata",
        ))?;
    for reference in [style, stylesheet] {
        if !info.object_references.contains(&reference) {
            info.object_references.push(reference);
        }
    }
    let style_location = workspace.locate(style).map_err(SymbolError::from)?;
    let style_component = component_of(workspace, style)?;
    let stylesheet_component = component_of(workspace, stylesheet)?;
    let chart_component = component_of(workspace, chart)?;
    register(workspace, stylesheet, style, identifier).map_err(SymbolError::from)?;
    workspace
        .insert(style_location, object)
        .map_err(SymbolError::from)?;
    if let Some(style_component) = style_component {
        budget.registry_entries(1)?;
        workspace
            .add_object_uuids(style_component, &[identifier])
            .map_err(SymbolError::from)?;
        for owner in [stylesheet_component, chart_component]
            .into_iter()
            .flatten()
        {
            if owner != style_component {
                budget.references(1)?;
                workspace
                    .add_external_reference(owner, style_component, identifier)
                    .map_err(SymbolError::from)?;
            }
        }
    }
    repoint(workspace, chart, style, identifier)?;
    workspace
        .set_last_object_identifier(identifier)
        .map_err(SymbolError::from)?;
    let member = workspace.member(style_location).to_owned();
    if let Some(scope) = budget.scope_mut() {
        scope.plan(identifier, &member, LEGEND_STYLE_MESSAGE_TYPE, true);
    }
    Ok(identifier)
}

/// Collapse `style` into its parent when it no longer differs from it
/// (the host's `collapse_if_equivalent`); returns the style the chart uses.
fn collapse_if_equivalent(
    workspace: &mut Workspace<'_>,
    chart: u64,
    style: u64,
    budget: &mut SymbolBudget,
) -> SymbolResult<u64> {
    let (_, payload) = style_payload(workspace, style, budget)?;
    let parts = style_parts(payload)?;
    let (Some(parent), Some(stylesheet)) = (parts.parent, parts.stylesheet) else {
        return Ok(style);
    };
    if !parts.variation || owners(workspace, style, budget)? != 1 {
        return Ok(style);
    }
    let child = without_style(payload)?;
    let (_, parent_payload) = style_payload(workspace, parent, budget)?;
    if child != without_style(parent_payload)? {
        return Ok(style);
    }
    let style_location = workspace.locate(style).map_err(SymbolError::from)?;
    let parent_location = workspace.locate(parent).map_err(SymbolError::from)?;
    let style_component = component_of(workspace, style)?;
    let parent_component = component_of(workspace, parent)?;
    let stylesheet_component = component_of(workspace, stylesheet)?;
    let chart_component = component_of(workspace, chart)?;
    repoint(workspace, chart, style, parent)?;
    if let Some(style_component) = style_component {
        budget.registry_entries(2)?;
        workspace
            .remove_object_uuids(style_component, &[style])
            .map_err(SymbolError::from)?;
        workspace
            .remove_external_references_to(style_component, style)
            .map_err(SymbolError::from)?;
        if parent_location == style_location
            && let Some(stylesheet_component) = stylesheet_component
            && stylesheet_component != style_component
        {
            workspace
                .add_external_reference(stylesheet_component, style_component, parent)
                .map_err(SymbolError::from)?;
        }
    }
    unregister(workspace, stylesheet, parent, style).map_err(SymbolError::from)?;
    workspace.remove(style).map_err(SymbolError::from)?;
    if let (Some(chart_component), Some(parent_component)) = (chart_component, parent_component)
        && chart_component != parent_component
    {
        budget.references(1)?;
        workspace
            .add_external_reference(chart_component, parent_component, parent)
            .map_err(SymbolError::from)?;
    }
    budget.identifiers(1)?;
    workspace
        .release_identifier_suffix(&[style])
        .map_err(SymbolError::from)?;
    if let Some(scope) = budget.scope_mut() {
        let member = workspace.member(parent_location).to_owned();
        scope.allow(parent, &member, LEGEND_STYLE_MESSAGE_TYPE);
        let style_member = workspace.member(style_location).to_owned();
        scope.plan(style, &style_member, LEGEND_STYLE_MESSAGE_TYPE, false);
    }
    Ok(parent)
}

/// The scope of a legend write: the chart, its legend style, the
/// stylesheet, the parent and the package metadata.
fn scope(
    workspace: &Workspace<'_>,
    chart: u64,
    style: u64,
    budget: &mut SymbolBudget,
) -> SymbolResult<MutationScope> {
    let (_, payload) = style_payload(workspace, style, budget)?;
    let parts = style_parts(payload)?;
    let member = |object: u64| -> SymbolResult<String> {
        let location = workspace.locate(object).map_err(SymbolError::from)?;
        Ok(workspace.member(location).to_owned())
    };
    let mut scope = MutationScope::in_place(
        chart,
        Vec::new(),
        Vec::new(),
        Vec::new(),
        vec![codec::EXTENSION_FIELD],
        vec![CHART_LEGEND_STYLE_FIELD],
    );
    scope.allow(chart, &member(chart)?, CHART_MESSAGE_TYPE);
    scope.allow(style, &member(style)?, LEGEND_STYLE_MESSAGE_TYPE);
    for related in [parts.parent, parts.stylesheet].into_iter().flatten() {
        let message_type = if Some(related) == parts.stylesheet {
            STYLESHEET_MESSAGE_TYPE
        } else {
            LEGEND_STYLE_MESSAGE_TYPE
        };
        scope.allow(related, &member(related)?, message_type);
    }
    for name in METADATA_MEMBERS {
        if let Some(location) = workspace.member_location(name) {
            for object in &workspace
                .archive_at(location)
                .map_err(SymbolError::from)?
                .objects
            {
                if let Some(identifier) = object.archive_info.identifier {
                    scope.allow(identifier, name, PACKAGE_METADATA_MESSAGE_TYPE);
                }
            }
        }
    }
    scope.allow_fields(&[
        STYLE_SUPER_FIELD,
        STYLE_PARENT_FIELD,
        STYLE_VARIATION_FIELD,
        STYLE_STYLESHEET_FIELD,
        1,
        2,
        3,
        5,
        6,
        8,
        10,
        11,
        12,
    ]);
    budget.preflight_allocations(8)?;
    budget.allocations(8)?;
    Ok(scope)
}

fn current(
    workspace: &Workspace<'_>,
    chart: u64,
    kind: PropertyKind,
    budget: &mut SymbolBudget,
) -> SymbolResult<Property> {
    let style = legend_style_of(workspace, chart, budget)?;
    let (_, payload) = style_payload(workspace, style, budget)?;
    value(payload, kind)
}

/// Read legend property `kind` of chart drawable `chart` under one residual
/// profile.
///
/// # Errors
///
/// [`Refusal::MissingStyle`], an invalid chart or style, or an exhausted
/// limit.
pub fn read(
    workspace: &Workspace<'_>,
    chart: u64,
    kind: PropertyKind,
    limits: Limits,
) -> SymbolResult<(Property, Report)> {
    let mut budget = SymbolBudget::new(limits);
    let value = current(workspace, chart, kind, &mut budget)?;
    Ok((value, budget.report(false)))
}

/// Store legend `value` for chart drawable `chart` under one residual
/// profile: derive a private legend-style variation when needed, write the
/// value, and collapse the variation when it equals its parent again.
///
/// # Errors
///
/// [`Refusal::MissingStyle`], [`Refusal::ImageFill`], an invalid chart,
/// style or registry, an exhausted limit, or a write that does not read
/// back.
pub fn write(
    workspace: &mut Workspace<'_>,
    chart: u64,
    value: &Property,
    limits: Limits,
) -> SymbolResult<(bool, Report)> {
    let mut budget = SymbolBudget::new(limits);
    let kind = value.kind();
    if current(workspace, chart, kind, &mut budget)? == *value {
        return Ok((false, budget.report(false)));
    }
    if matches!(value, Property::Fill(LegendFill::Fill(Fill::Image(_)))) {
        return Err(SymbolError::Refused(Refusal::ImageFill));
    }
    budget.work(workspace.source_work())?;
    let style = legend_style_of(workspace, chart, &mut budget)?;
    let scope = scope(workspace, chart, style, &mut budget)?;
    budget.set_scope(scope);
    let target = ensure_exclusive(workspace, chart, style, &mut budget)?;
    let (message, payload) = style_payload(workspace, target, &mut budget)?;
    let data = charged_patch(payload, &mut budget, |payload| rewrite(payload, value))?;
    workspace
        .replace_payload(target, message, data)
        .map_err(SymbolError::from)?;
    collapse_if_equivalent(workspace, chart, target, &mut budget)?;
    if current(workspace, chart, kind, &mut budget)? != *value {
        return Err(SymbolError::Invalid(
            "a legend-style write does not read back",
        ));
    }
    Ok((true, budget.report(true)))
}
