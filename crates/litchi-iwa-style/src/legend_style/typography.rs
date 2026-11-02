//! The legend label typography (record 0414): the host's
//! `charts::legend_font` graph over a [`Workspace`].
//!
//! iWork stores the legend's typography indirectly: field 2 of the legend
//! style's generated extension selects one entry of the chart's
//! `paragraph_styles` list (`TSCH.ChartArchive`, field 20). A direct font
//! identity, face traits and point size share one private paragraph-style
//! variation of the list's first style. A write keeps the variation's other
//! overrides, patches an exclusive variation in place, replaces a shared one
//! with a fresh variation, appends a new variation (and points the legend at
//! it) when the legend has none, and removes the variation (pointing the
//! legend back at the inherited style) when a reset empties it.

use std::cell::RefCell;

use litchi_iwa_common::wire::{self, WireView};
use litchi_iwa_core::{ArchiveObject, RawMessage};
use litchi_iwa_text::character::TextPointSize;
use litchi_iwa_text::font::Font;
use litchi_iwa_text_wire::paragraph_style::{
    self as text_style, Property as TextProperty, PropertyKind as TextPropertyKind, Styles,
    VariationOverrides,
};

use super::{CHART_MESSAGE_TYPE, STYLESHEET_MESSAGE_TYPE, component_of, invalid, wire_error};
use crate::graph::{Error, Workspace};
use crate::series::{
    SymbolBudget, SymbolError, SymbolResult, bounded_chart_payload, bounded_references,
};
use crate::shape::{register, unregister};

const PARAGRAPH_STYLE_MESSAGE_TYPE: u32 = text_style::MESSAGE_TYPE;
const EXTENSION_FIELD: u32 = 10_000;
const CHART_PARAGRAPH_STYLES_FIELD: u32 = 20;
const LEGEND_PARAGRAPH_INDEX_FIELD: u32 = 2;
const STYLE_SUPER_FIELD: u32 = 1;
const STYLE_PARENT_FIELD: u32 = 3;
const STYLE_STYLESHEET_FIELD: u32 = 5;
const REFERENCE_IDENTIFIER_FIELD: u32 = 1;
const OVERRIDE_COUNT_FIELD: u32 = 10;
const CHARACTER_PROPERTIES_FIELD: u32 = 11;
const CHARACTER_BOLD_FIELD: u32 = 1;
const CHARACTER_ITALIC_FIELD: u32 = 2;
const CHARACTER_FONT_SIZE_FIELD: u32 = 3;
const CHARACTER_FONT_NAME_NULL_FIELD: u32 = 4;
const CHARACTER_FONT_NAME_FIELD: u32 = 5;

/// A legend's font identity and face traits.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
#[non_exhaustive]
pub enum LegendFont {
    /// No direct override: the legend resolves its base paragraph style.
    #[default]
    Inherited,
    /// A direct font identity with the effective bold and italic traits.
    Font {
        /// The font.
        font: Font,
        /// Bold.
        bold: bool,
        /// Italic.
        italic: bool,
    },
}

/// A legend's point size.
#[derive(Debug, Clone, Copy, PartialEq, Default)]
#[non_exhaustive]
pub enum LegendFontSize {
    /// No direct override: the legend resolves its base paragraph style.
    #[default]
    Inherited,
    /// A direct point size.
    Size(TextPointSize),
}

/// One typography change.
#[derive(Debug, Clone, Copy)]
pub(super) enum Change<'a> {
    Font(&'a LegendFont),
    Size(LegendFontSize),
}

fn text_error(_error: text_style::Error) -> SymbolError {
    SymbolError::Invalid("an invalid legend paragraph style")
}

/// The workspace's paragraph styles for the text-wire reader, charging
/// every payload it hands out.
struct ParagraphStyles<'w, 's, 'b> {
    workspace: &'w Workspace<'s>,
    budget: RefCell<&'b mut SymbolBudget>,
    error: RefCell<Option<SymbolError>>,
}

impl Styles for ParagraphStyles<'_, '_, '_> {
    fn paragraph_style(&self, identifier: u64) -> Option<&[u8]> {
        let (_, payload) = self
            .workspace
            .unique_payload(identifier, PARAGRAPH_STYLE_MESSAGE_TYPE, "")
            .ok()?;
        let mut budget = self.budget.borrow_mut();
        if let Err(error) = budget
            .source(payload.len())
            .and_then(|()| budget.work(payload.len()))
        {
            *self.error.borrow_mut() = Some(error);
            return None;
        }
        Some(payload)
    }
}

/// Run a text-wire reader over the workspace's paragraph styles.
fn with_styles<T>(
    workspace: &Workspace<'_>,
    budget: &mut SymbolBudget,
    read: impl FnOnce(
        &ParagraphStyles<'_, '_, '_>,
        litchi_iwa_common::WireLimits,
    ) -> text_style::Result<T>,
) -> SymbolResult<T> {
    let limits = budget.wire_limits()?;
    let styles = ParagraphStyles {
        workspace,
        budget: RefCell::new(budget),
        error: RefCell::new(None),
    };
    let result = read(&styles, limits);
    if let Some(error) = styles.error.into_inner() {
        return Err(error);
    }
    result.map_err(text_error)
}

fn overrides_of(
    workspace: &Workspace<'_>,
    style: u64,
    budget: &mut SymbolBudget,
) -> SymbolResult<VariationOverrides> {
    with_styles(workspace, budget, |styles, limits| {
        text_style::variation_overrides(styles, style, limits)
    })?
    .ok_or(SymbolError::Invalid(
        "a legend paragraph style is not an exact native variation",
    ))
}

/// The effective font identity and face traits of `style`.
fn effective_font(
    workspace: &Workspace<'_>,
    style: u64,
    budget: &mut SymbolBudget,
) -> SymbolResult<(Font, bool, bool)> {
    let font = match with_styles(workspace, budget, |styles, limits| {
        text_style::effective(styles, style, TextPropertyKind::TextFont, limits)
    })? {
        TextProperty::TextFont(font) => font,
        _ => return Err(SymbolError::Invalid("an unexpected effective font")),
    };
    let (bold, italic) = effective_traits(workspace, style, budget)?;
    Ok((font, bold, italic))
}

fn effective_traits(
    workspace: &Workspace<'_>,
    style: u64,
    budget: &mut SymbolBudget,
) -> SymbolResult<(bool, bool)> {
    match with_styles(workspace, budget, |styles, limits| {
        text_style::effective(styles, style, TextPropertyKind::TextStyle, limits)
    })? {
        TextProperty::TextStyle(text) => Ok((text.bold, text.italic)),
        _ => Err(SymbolError::Invalid("an unexpected effective text style")),
    }
}

fn style_payload<'w>(
    workspace: &'w Workspace<'_>,
    style: u64,
    budget: &mut SymbolBudget,
) -> SymbolResult<(usize, &'w [u8])> {
    let (message, payload) = workspace
        .unique_payload(
            style,
            PARAGRAPH_STYLE_MESSAGE_TYPE,
            "a legend paragraph style must have exactly one payload",
        )
        .map_err(SymbolError::from)?;
    budget.source(payload.len())?;
    budget.work(payload.len())?;
    Ok((message, payload))
}

/// The last singular varint `field` of `payload`.
fn varint(payload: &[u8], field: u32) -> SymbolResult<Option<u64>> {
    let mut value = None;
    for candidate in WireView::parse(payload, litchi_iwa_common::WireLimits::DEFAULT)
        .map_err(invalid)?
        .fields()
    {
        if candidate.number() == field {
            if candidate.wire_type() != 0 {
                return Err(SymbolError::Invalid("a malformed style reference"));
            }
            let (decoded, _) =
                litchi_iwa_common::varint::decode_varint_from_bytes(candidate.payload())
                    .map_err(invalid)?;
            value = Some(decoded);
        }
    }
    Ok(value)
}

/// The last length-delimited `field` of `payload`.
fn message(payload: &[u8], field: u32) -> SymbolResult<Option<&[u8]>> {
    let mut value = None;
    for candidate in WireView::parse(payload, litchi_iwa_common::WireLimits::DEFAULT)
        .map_err(invalid)?
        .fields()
    {
        if candidate.number() == field {
            if candidate.wire_type() != 2 {
                return Err(SymbolError::Invalid("a malformed style field"));
            }
            value = Some(candidate.payload());
        }
    }
    Ok(value)
}

/// The parent and stylesheet of a paragraph style payload.
fn style_links(payload: &[u8]) -> SymbolResult<(Option<u64>, Option<u64>)> {
    let Some(style) = message(payload, STYLE_SUPER_FIELD)? else {
        return Ok((None, None));
    };
    let reference = |field| -> SymbolResult<Option<u64>> {
        message(style, field)?
            .map(|reference| varint(reference, REFERENCE_IDENTIFIER_FIELD))
            .transpose()
            .map(Option::flatten)
    };
    Ok((
        reference(STYLE_PARENT_FIELD)?,
        reference(STYLE_STYLESHEET_FIELD)?,
    ))
}

/// The stylesheet of a paragraph style payload.
pub(super) fn stylesheet_of(payload: &[u8]) -> SymbolResult<Option<u64>> {
    Ok(style_links(payload)?.1)
}

/// The legend style's direct paragraph-style index (the host's
/// `direct_paragraph_style_index`: a negative `int32` is refused).
pub(super) fn direct_index(legend_payload: &[u8]) -> SymbolResult<Option<usize>> {
    let Some(extension) = message(legend_payload, EXTENSION_FIELD)? else {
        return Ok(None);
    };
    let Some(raw) = varint(extension, LEGEND_PARAGRAPH_INDEX_FIELD)? else {
        return Ok(None);
    };
    #[allow(
        clippy::cast_possible_truncation,
        reason = "prost decodes an int32 from the low 32 bits of its varint"
    )]
    let index = raw as u32 as i32;
    usize::try_from(index)
        .map(Some)
        .map_err(|_| SymbolError::Invalid("a negative legend paragraph-style index"))
}

/// The legend style payload naming `index` (`None`: no direct style), as
/// the host's `patch_direct_paragraph_style_index`.
fn patch_direct_index(legend_payload: &[u8], index: Option<u64>) -> SymbolResult<Vec<u8>> {
    let Some(extension) = message(legend_payload, EXTENSION_FIELD)? else {
        let Some(index) = index else {
            return Ok(legend_payload.to_vec());
        };
        let mut generated = Vec::new();
        wire::append_varint_field(&mut generated, LEGEND_PARAGRAPH_INDEX_FIELD, index)
            .map_err(wire_error)?;
        return wire::patch_length_delimited_field(
            legend_payload,
            EXTENSION_FIELD,
            false,
            Some(&generated),
        )
        .map_err(wire_error);
    };
    let present = varint(extension, LEGEND_PARAGRAPH_INDEX_FIELD)?.is_some();
    let extension =
        wire::patch_varint_field(extension, LEGEND_PARAGRAPH_INDEX_FIELD, present, index)
            .map_err(wire_error)?;
    wire::patch_length_delimited_field(legend_payload, EXTENSION_FIELD, true, Some(&extension))
        .map_err(wire_error)
}

/// The chart's `paragraph_styles` list.
fn paragraph_styles(
    workspace: &Workspace<'_>,
    chart: u64,
    budget: &mut SymbolBudget,
) -> SymbolResult<Vec<u64>> {
    let (_, payload) = bounded_chart_payload(workspace, chart, budget)?;
    let styles = bounded_references(payload, budget)?
        .paragraph_styles()
        .to_vec();
    if styles.is_empty() || styles.contains(&0) {
        return Err(SymbolError::Invalid(
            "a chart has no or a zero paragraph-style reference",
        ));
    }
    Ok(styles)
}

/// The located legend typography graph (the host's `LegendFontGraph`).
struct Graph {
    styles: Vec<u64>,
    index: Option<usize>,
}

impl Graph {
    fn locate(
        workspace: &Workspace<'_>,
        chart: u64,
        legend: u64,
        budget: &mut SymbolBudget,
    ) -> SymbolResult<Self> {
        let styles = paragraph_styles(workspace, chart, budget)?;
        let (_, payload) = super::style_payload(workspace, legend, budget)?;
        let index = direct_index(payload)?;
        if index.is_some_and(|index| index >= styles.len()) {
            return Err(SymbolError::Invalid(
                "a legend paragraph-style index is out of bounds",
            ));
        }
        Ok(Self { styles, index })
    }

    fn direct(&self) -> Option<(usize, u64)> {
        self.index.map(|index| (index, self.styles[index]))
    }
}

/// The legend's font (the host's `read_font`).
pub(super) fn font(
    workspace: &Workspace<'_>,
    chart: u64,
    budget: &mut SymbolBudget,
) -> SymbolResult<LegendFont> {
    let legend = super::legend_style_of(workspace, chart, budget)?;
    let graph = Graph::locate(workspace, chart, legend, budget)?;
    let Some((_, style)) = graph.direct() else {
        return Ok(LegendFont::Inherited);
    };
    let overrides = overrides_of(workspace, style, budget)?.font();
    if overrides.font.is_none() && overrides.bold.is_none() && overrides.italic.is_none() {
        return Ok(LegendFont::Inherited);
    }
    let (font, bold, italic) = effective_font(workspace, style, budget)?;
    Ok(LegendFont::Font { font, bold, italic })
}

/// The legend's point size (the host's `read_size`).
pub(super) fn size(
    workspace: &Workspace<'_>,
    chart: u64,
    budget: &mut SymbolBudget,
) -> SymbolResult<LegendFontSize> {
    let legend = super::legend_style_of(workspace, chart, budget)?;
    let graph = Graph::locate(workspace, chart, legend, budget)?;
    let Some((_, style)) = graph.direct() else {
        return Ok(LegendFontSize::Inherited);
    };
    Ok(overrides_of(workspace, style, budget)?
        .font()
        .point_size
        .map_or(LegendFontSize::Inherited, LegendFontSize::Size))
}

/// Apply `change` to the overrides of a variation of `parent` (the host's
/// `apply_property`; a reset clears the property). Returns whether a
/// reset cleared anything.
fn apply(
    workspace: &Workspace<'_>,
    parent: u64,
    overrides: &mut VariationOverrides,
    change: Change<'_>,
    budget: &mut SymbolBudget,
) -> SymbolResult<bool> {
    let mut font = overrides.font();
    let cleared = match change {
        Change::Font(LegendFont::Font {
            font: target,
            bold,
            italic,
        }) => {
            let (inherited_bold, inherited_italic) = effective_traits(workspace, parent, budget)?;
            font.font = Some(target.clone());
            font.bold = (*bold != inherited_bold).then_some(*bold);
            font.italic = (*italic != inherited_italic).then_some(*italic);
            false
        },
        Change::Size(LegendFontSize::Size(size)) => {
            font.point_size = Some(size);
            false
        },
        Change::Font(_) => {
            let present = font.font.is_some() || font.bold.is_some() || font.italic.is_some();
            font.font = None;
            font.bold = None;
            font.italic = None;
            present
        },
        Change::Size(_) => font.point_size.take().is_some(),
    };
    overrides.set_font(font);
    Ok(cleared)
}

/// Whether `style` is referenced, outside its stylesheet and its child
/// styles, by `chart`'s chart payload only (the host's
/// `style_is_exclusive_to_chart`; the chart must reference it once).
fn exclusive_to_chart(
    workspace: &Workspace<'_>,
    style: u64,
    chart: u64,
    budget: &mut SymbolBudget,
) -> SymbolResult<bool> {
    let chart_location = workspace.locate(chart).map_err(SymbolError::from)?;
    let mut matching = 0usize;
    let mut others = 0usize;
    for location in workspace.locations_iter() {
        for object in &workspace
            .archive_at(location)
            .map_err(SymbolError::from)?
            .objects
        {
            for (message, info) in object
                .messages
                .iter()
                .zip(&object.archive_info.message_infos)
            {
                budget.work(info.object_references.len().max(1))?;
                if !info.object_references.contains(&style)
                    || message.type_ == STYLESHEET_MESSAGE_TYPE
                {
                    continue;
                }
                if message.type_ == PARAGRAPH_STYLE_MESSAGE_TYPE
                    && style_links(&message.data)?.0 == Some(style)
                {
                    continue;
                }
                if location == chart_location
                    && object.archive_info.identifier == Some(chart)
                    && message.type_ == CHART_MESSAGE_TYPE
                {
                    matching += 1;
                } else {
                    others += 1;
                }
            }
        }
    }
    if matching != 1 {
        return Err(SymbolError::Invalid(
            "the chart must reference its legend paragraph style once",
        ));
    }
    Ok(others == 0)
}

/// Patch the font subset of the exclusive variation `style` in place (the
/// host's `patch_existing_font` / `patch_existing_size`): the character
/// properties change field by field and the override count follows.
fn patch_in_place(
    workspace: &mut Workspace<'_>,
    style: u64,
    expected: &VariationOverrides,
    budget: &mut SymbolBudget,
) -> SymbolResult<()> {
    let current = overrides_of(workspace, style, budget)?.font();
    let next = expected.font();
    let (message, payload) = style_payload(workspace, style, budget)?;
    let count_present = varint(payload, OVERRIDE_COUNT_FIELD)?.is_some();
    let null = |font: &Option<Font>| matches!(font, Some(font) if font.name().is_none());
    let data =
        wire::transform_length_delimited_field(payload, CHARACTER_PROPERTIES_FIELD, |characters| {
            let characters = wire::patch_varint_field(
                characters,
                CHARACTER_BOLD_FIELD,
                current.bold.is_some(),
                next.bold.map(u64::from),
            )?;
            let characters = wire::patch_varint_field(
                &characters,
                CHARACTER_ITALIC_FIELD,
                current.italic.is_some(),
                next.italic.map(u64::from),
            )?;
            let characters = wire::patch_fixed32_field(
                &characters,
                CHARACTER_FONT_SIZE_FIELD,
                current.point_size.is_some(),
                next.point_size.map(|size| size.points().to_bits()),
            )?;
            let characters = wire::patch_varint_field(
                &characters,
                CHARACTER_FONT_NAME_NULL_FIELD,
                null(&current.font),
                null(&next.font).then_some(1),
            )?;
            wire::patch_length_delimited_field(
                &characters,
                CHARACTER_FONT_NAME_FIELD,
                current
                    .font
                    .as_ref()
                    .is_some_and(|font| font.name().is_some()),
                next.font.as_ref().and_then(Font::name).map(str::as_bytes),
            )
        })
        .map_err(wire_error)?;
    let data = wire::patch_varint_field(
        &data,
        OVERRIDE_COUNT_FIELD,
        count_present,
        Some(u64::from(expected.count())),
    )
    .map_err(wire_error)?;
    budget.preflight_retained(data.len())?;
    budget.retained(data.len())?;
    workspace
        .replace_payload(style, message, data)
        .map_err(SymbolError::from)?;
    if overrides_of(workspace, style, budget)? != *expected {
        return Err(SymbolError::Invalid(
            "a legend paragraph-style patch does not read back",
        ));
    }
    Ok(())
}

/// Insert variation `variation` (identifier `identifier`, child of
/// `parent` in `stylesheet`) beside its parent, register it and record it
/// in the package registry (the host's `insert_style_variation` and
/// `register_private_style`).
fn insert(
    workspace: &mut Workspace<'_>,
    chart: u64,
    identifier: u64,
    parent: u64,
    stylesheet: u64,
    variation: text_style::Variation,
    budget: &mut SymbolBudget,
) -> SymbolResult<()> {
    register(workspace, stylesheet, parent, identifier).map_err(SymbolError::from)?;
    let references = variation.references().to_vec();
    let payload = variation.into_payload();
    budget.identifiers(1)?;
    budget.preflight_allocations(2)?;
    budget.preflight_retained(payload.len())?;
    budget.allocations(2)?;
    budget.retained(payload.len())?;
    let mut object = ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: PARAGRAPH_STYLE_MESSAGE_TYPE,
            data: payload,
        }],
        workspace.limits(),
    )
    .map_err(|error| SymbolError::Graph(Error::Archive(error)))?;
    let info = &mut object.archive_info.message_infos[0];
    info.versions = text_style::MESSAGE_VERSION.to_vec().into();
    info.object_references.extend(references);
    let parent_location = workspace.locate(parent).map_err(SymbolError::from)?;
    let style_component = component_of(workspace, parent)?;
    let chart_component = component_of(workspace, chart)?;
    workspace
        .insert(parent_location, object)
        .map_err(SymbolError::from)?;
    if let Some(style_component) = style_component {
        budget.registry_entries(1)?;
        workspace
            .add_object_uuids(style_component, &[identifier])
            .map_err(SymbolError::from)?;
        if let Some(owner) = chart_component
            && owner != style_component
        {
            budget.references(1)?;
            workspace
                .add_external_reference(owner, style_component, identifier)
                .map_err(SymbolError::from)?;
        }
    }
    let member = workspace.member(parent_location).to_owned();
    if let Some(scope) = budget.scope_mut() {
        scope.plan(identifier, &member, PARAGRAPH_STYLE_MESSAGE_TYPE, true);
    }
    Ok(())
}

/// Remove the chart component's external reference to `style` when no
/// object of that component references it any more (the host's
/// `unregister_owner_reference_if_unused`).
fn owner_reference_if_unused(
    workspace: &mut Workspace<'_>,
    chart: u64,
    style: u64,
    budget: &mut SymbolBudget,
) -> SymbolResult<()> {
    let (Some(style_component), Some(owner)) = (
        component_of(workspace, style)?,
        component_of(workspace, chart)?,
    ) else {
        return Ok(());
    };
    if owner == style_component {
        return Ok(());
    }
    for location in workspace.locations_iter() {
        let member = workspace.member(location).to_owned();
        if workspace
            .component_for_member(&member)
            .map_err(SymbolError::from)?
            != Some(owner)
        {
            continue;
        }
        for object in &workspace
            .archive_at(location)
            .map_err(SymbolError::from)?
            .objects
        {
            budget.work(object.archive_info.message_infos.len().max(1))?;
            if object
                .archive_info
                .message_infos
                .iter()
                .any(|info| info.object_references.contains(&style))
            {
                return Ok(());
            }
        }
    }
    budget.references(1)?;
    workspace
        .remove_external_reference(owner, style_component, style)
        .map_err(SymbolError::from)
}

/// Rewrite the chart's `paragraph_styles` list to `next` (payload and
/// metadata, as the host's `patch_chart_paragraph_styles`).
fn patch_list(
    workspace: &mut Workspace<'_>,
    chart: u64,
    previous: &[u64],
    next: &[u64],
    budget: &mut SymbolBudget,
) -> SymbolResult<()> {
    let (message, payload) = workspace
        .unique_payload(
            chart,
            CHART_MESSAGE_TYPE,
            "a chart must have exactly one chart payload",
        )
        .map_err(SymbolError::from)?;
    let mut encoded = Vec::new();
    for &identifier in next {
        let mut reference = Vec::new();
        wire::append_varint_field(&mut reference, REFERENCE_IDENTIFIER_FIELD, identifier)
            .map_err(wire_error)?;
        encoded.push(reference);
    }
    let data = wire::transform_length_delimited_field(payload, EXTENSION_FIELD, |archive| {
        wire::rewrite_repeated_length_delimited_fields(
            archive,
            CHART_PARAGRAPH_STYLES_FIELD,
            &encoded,
        )
    })
    .map_err(wire_error)?;
    budget.preflight_retained(data.len())?;
    budget.retained(data.len())?;
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
    info.object_references
        .retain(|identifier| !previous.contains(identifier) || next.contains(identifier));
    for identifier in next {
        if !info.object_references.contains(identifier) {
            info.object_references.push(*identifier);
        }
    }
    Ok(())
}

/// Point the legend style `legend` at `index`.
fn patch_legend_index(
    workspace: &mut Workspace<'_>,
    legend: u64,
    index: Option<u64>,
    budget: &mut SymbolBudget,
) -> SymbolResult<()> {
    let (message, payload) = super::style_payload(workspace, legend, budget)?;
    let data = patch_direct_index(payload, index)?;
    workspace
        .replace_payload(legend, message, data)
        .map_err(SymbolError::from)
}

/// Replace the shared variation at `index` with a fresh variation of its
/// parent holding `overrides` (the host's `replace_shared_style`).
fn replace_shared(
    workspace: &mut Workspace<'_>,
    chart: u64,
    graph: &Graph,
    index: usize,
    previous: u64,
    overrides: &VariationOverrides,
    budget: &mut SymbolBudget,
) -> SymbolResult<()> {
    if overrides.is_empty() {
        return Err(SymbolError::Invalid(
            "a shared legend paragraph-style replacement has no overrides",
        ));
    }
    let (_, payload) = style_payload(workspace, previous, budget)?;
    let (parent, stylesheet) = style_links(payload)?;
    let parent = parent.ok_or(SymbolError::Invalid("a variation has no parent"))?;
    let stylesheet = stylesheet.ok_or(SymbolError::Invalid("a style has no stylesheet"))?;
    let identifier = workspace
        .next_object_identifier()
        .map_err(SymbolError::from)?;
    let limits = budget.wire_limits()?;
    let variation =
        text_style::variation_with(parent, stylesheet, overrides, limits).map_err(text_error)?;
    insert(
        workspace, chart, identifier, parent, stylesheet, variation, budget,
    )?;
    let mut next = graph.styles.clone();
    next[index] = identifier;
    patch_list(workspace, chart, &graph.styles, &next, budget)?;
    owner_reference_if_unused(workspace, chart, previous, budget)?;
    workspace
        .set_last_object_identifier(identifier)
        .map_err(SymbolError::from)
}

/// Apply a typography change to chart `chart` whose (exclusive) legend
/// style is `legend` (the host's `set_property` / `reset_property`).
pub(super) fn apply_change(
    workspace: &mut Workspace<'_>,
    chart: u64,
    legend: u64,
    change: Change<'_>,
    budget: &mut SymbolBudget,
) -> SymbolResult<()> {
    let graph = Graph::locate(workspace, chart, legend, budget)?;
    let reset = matches!(
        change,
        Change::Font(LegendFont::Inherited) | Change::Size(LegendFontSize::Inherited)
    );
    let Some((index, style)) = graph.direct() else {
        if reset {
            return Ok(());
        }
        let parent = graph.styles[0];
        let (_, payload) = style_payload(workspace, parent, budget)?;
        let stylesheet = style_links(payload)?
            .1
            .ok_or(SymbolError::Invalid("a style has no stylesheet"))?;
        let identifier = workspace
            .next_object_identifier()
            .map_err(SymbolError::from)?;
        let mut overrides = VariationOverrides::default();
        apply(workspace, parent, &mut overrides, change, budget)?;
        let limits = budget.wire_limits()?;
        let variation = text_style::variation_with(parent, stylesheet, &overrides, limits)
            .map_err(text_error)?;
        insert(
            workspace, chart, identifier, parent, stylesheet, variation, budget,
        )?;
        let mut next = graph.styles.clone();
        next.push(identifier);
        patch_list(workspace, chart, &graph.styles, &next, budget)?;
        patch_legend_index(
            workspace,
            legend,
            Some(u64::try_from(graph.styles.len()).map_err(invalid)?),
            budget,
        )?;
        return workspace
            .set_last_object_identifier(identifier)
            .map_err(SymbolError::from);
    };
    let mut overrides = overrides_of(workspace, style, budget)?;
    let (_, payload) = style_payload(workspace, style, budget)?;
    let (parent, stylesheet) = style_links(payload)?;
    let parent = parent.ok_or(SymbolError::Invalid("a variation has no parent"))?;
    let cleared = apply(workspace, parent, &mut overrides, change, budget)?;
    if reset && !cleared {
        return Ok(());
    }
    let exclusive = exclusive_to_chart(workspace, style, chart, budget)?;
    if !overrides.is_empty() {
        if exclusive {
            return patch_in_place(workspace, style, &overrides, budget);
        }
        return replace_shared(workspace, chart, &graph, index, style, &overrides, budget);
    }
    // A reset emptied the variation: the legend inherits again.
    let stylesheet = stylesheet.ok_or(SymbolError::Invalid("a style has no stylesheet"))?;
    patch_legend_index(workspace, legend, None, budget)?;
    let mut next = graph.styles.clone();
    if index + 1 == next.len() {
        next.pop();
    } else {
        next[index] = parent;
    }
    patch_list(workspace, chart, &graph.styles, &next, budget)?;
    if !exclusive {
        return owner_reference_if_unused(workspace, chart, style, budget);
    }
    let style_location = workspace.locate(style).map_err(SymbolError::from)?;
    let style_member = workspace.member(style_location).to_owned();
    let style_component = component_of(workspace, style)?;
    let chart_component = component_of(workspace, chart)?;
    unregister(workspace, stylesheet, parent, style).map_err(SymbolError::from)?;
    workspace.remove(style).map_err(SymbolError::from)?;
    if let Some(style_component) = style_component {
        budget.registry_entries(2)?;
        workspace
            .remove_object_uuids(style_component, &[style])
            .map_err(SymbolError::from)?;
        workspace
            .remove_external_references_to(style_component, style)
            .map_err(SymbolError::from)?;
        if let Some(owner) = chart_component
            && owner != style_component
        {
            budget.references(1)?;
            workspace
                .add_external_reference(owner, style_component, parent)
                .map_err(SymbolError::from)?;
        }
    }
    budget.identifiers(1)?;
    workspace
        .release_identifier_suffix(&[style])
        .map_err(SymbolError::from)?;
    if let Some(scope) = budget.scope_mut() {
        scope.plan(style, &style_member, PARAGRAPH_STYLE_MESSAGE_TYPE, false);
    }
    Ok(())
}
