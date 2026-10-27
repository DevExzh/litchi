//! The chart-font graph engine (record 0375a): a chart's uniform font
//! identity, face traits and point size, stored in the paragraph styles its
//! `paragraph_styles` list (`TSCH.ChartArchive`, field 20) names.
//!
//! [`set`] (with a reset [`Change`]) is the host's `set_chart_font*` /
//! `reset_chart_font*` ported operation for operation over a
//! [`Workspace`]: every unique paragraph style of the chart is replaced by a
//! fresh variation (of its parent when it is a font-only variation, else of
//! itself) holding the font overrides that differ from the inherited values,
//! or by its parent when none is left; the chart's references are repointed,
//! the variations registered (stylesheet, UUID, external reference), and the
//! replaced font-only variations no longer used are removed. The variation
//! payloads are `litchi_iwa_text_wire::paragraph_style::font_variation`.

use std::collections::{BTreeMap, HashSet};

use litchi_iwa_common::WireLimits;
use litchi_iwa_common::wire;
use litchi_iwa_core::{ArchiveObject, RawMessage};
use litchi_iwa_protos::chart_style_codec::{self as codec, Limits as CodecLimits};
use litchi_iwa_text::character::{TextPointSize, TextStyle};
use litchi_iwa_text::font::Font;
use litchi_iwa_text_wire::paragraph_style::{
    self as text_style, FontOverrides, Property, PropertyKind, Styles,
};

use crate::graph::{self, Error, Result, Workspace};
use crate::shape;

const CHART_MESSAGE_TYPE: u32 = 5_021;
const PARAGRAPH_STYLE_MESSAGE_TYPE: u32 = 2_022;
const STYLESHEET_MESSAGE_TYPE: u32 = 401;
const CHART_EXTENSION_FIELD: u32 = 10_000;
const CHART_PARAGRAPH_STYLES_FIELD: u32 = 20;
const STYLE_SUPER_FIELD: u32 = 1;
const STYLE_PARENT_FIELD: u32 = 3;
const STYLE_STYLESHEET_FIELD: u32 = 5;
const REFERENCE_IDENTIFIER_FIELD: u32 = 1;

/// A chart's font identity and face traits.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct ChartFont {
    /// The font.
    pub font: Font,
    /// Bold.
    pub bold: bool,
    /// Italic.
    pub italic: bool,
}

/// What a chart-font write changes.
#[derive(Debug, Clone, PartialEq)]
#[non_exhaustive]
pub enum Change {
    /// Set the font identity and face traits.
    Font(ChartFont),
    /// Set the point size.
    Size(TextPointSize),
    /// Reset the font identity and face traits.
    ResetFont,
    /// Reset the point size.
    ResetSize,
}

fn text_error(_: litchi_iwa_text_wire::paragraph_style::Error) -> Error {
    Error::Invalid("a chart paragraph style is invalid")
}

struct ParagraphStyles<'w, 's>(&'w Workspace<'s>);

impl Styles for ParagraphStyles<'_, '_> {
    fn paragraph_style(&self, identifier: u64) -> Option<&[u8]> {
        self.0
            .unique_payload(identifier, PARAGRAPH_STYLE_MESSAGE_TYPE, "")
            .ok()
            .map(|(_, payload)| payload)
    }
}

fn limits() -> WireLimits {
    WireLimits::DEFAULT
}

fn chart_message(workspace: &Workspace<'_>, chart: u64) -> Result<(usize, Vec<u8>)> {
    let (message, payload) = workspace.unique_payload(
        chart,
        CHART_MESSAGE_TYPE,
        "a chart must have exactly one chart payload",
    )?;
    Ok((message, payload.to_vec()))
}

fn paragraph_styles(workspace: &Workspace<'_>, chart: u64) -> Result<Vec<u64>> {
    let (_, payload) = chart_message(workspace, chart)?;
    let styles = codec::references(&payload, CodecLimits::DEFAULT)?
        .paragraph_styles()
        .to_vec();
    if styles.is_empty() {
        return Err(Error::Invalid("a chart has no paragraph styles"));
    }
    if styles.contains(&0) {
        return Err(Error::Invalid(
            "a chart has a zero paragraph-style reference",
        ));
    }
    Ok(styles)
}

fn unique(identifiers: &[u64]) -> Vec<u64> {
    let mut seen = HashSet::new();
    identifiers
        .iter()
        .copied()
        .filter(|identifier| seen.insert(*identifier))
        .collect()
}

fn effective_font(workspace: &Workspace<'_>, style: u64) -> Result<ChartFont> {
    let styles = ParagraphStyles(workspace);
    let font = match text_style::effective(&styles, style, PropertyKind::TextFont, limits())
        .map_err(text_error)?
    {
        Property::TextFont(font) => font,
        _ => return Err(Error::Invalid("an unexpected effective font")),
    };
    let style = text_style_of(workspace, style)?;
    Ok(ChartFont {
        font,
        bold: style.bold,
        italic: style.italic,
    })
}

fn text_style_of(workspace: &Workspace<'_>, style: u64) -> Result<TextStyle> {
    match text_style::effective(
        &ParagraphStyles(workspace),
        style,
        PropertyKind::TextStyle,
        limits(),
    )
    .map_err(text_error)?
    {
        Property::TextStyle(style) => Ok(style),
        _ => Err(Error::Invalid("an unexpected effective text style")),
    }
}

fn uniform<T: PartialEq>(
    workspace: &Workspace<'_>,
    chart: u64,
    read: impl Fn(&Workspace<'_>, u64) -> Result<T>,
) -> Result<T> {
    let styles = paragraph_styles(workspace, chart)?;
    let first = read(workspace, styles[0])?;
    for style in &styles[1..] {
        if read(workspace, *style)? != first {
            return Err(Error::Invalid("a chart has mixed effective fonts"));
        }
    }
    Ok(first)
}

/// The uniform effective font of chart drawable `chart`.
///
/// # Errors
///
/// An invalid chart or paragraph style, or mixed fonts.
pub fn font(workspace: &Workspace<'_>, chart: u64) -> Result<ChartFont> {
    uniform(workspace, chart, effective_font)
}

/// The uniform effective point size of chart drawable `chart`.
///
/// # Errors
///
/// An invalid chart or paragraph style, or mixed sizes.
pub fn size(workspace: &Workspace<'_>, chart: u64) -> Result<TextPointSize> {
    uniform(workspace, chart, |workspace, style| {
        Ok(text_style_of(workspace, style)?.point_size)
    })
}

/// `(parent, stylesheet)` of a paragraph-style payload.
fn style_links(payload: &[u8]) -> Result<(Option<u64>, Option<u64>)> {
    let Some(base) = graph::repeated(payload, STYLE_SUPER_FIELD)?
        .first()
        .copied()
    else {
        return Err(Error::Invalid("a paragraph style has no style parent"));
    };
    let link = |number| -> Result<Option<u64>> {
        Ok(match graph::repeated(base, number)?.last() {
            Some(reference) => graph::varint(reference, REFERENCE_IDENTIFIER_FIELD)?,
            None => None,
        })
    };
    Ok((link(STYLE_PARENT_FIELD)?, link(STYLE_STYLESHEET_FIELD)?))
}

fn style_payload<'w>(workspace: &'w Workspace<'_>, style: u64) -> Result<(usize, &'w [u8])> {
    workspace.unique_payload(
        style,
        PARAGRAPH_STYLE_MESSAGE_TYPE,
        "a paragraph style must have exactly one payload",
    )
}

struct Addition {
    identifier: u64,
    parent: u64,
    stylesheet: u64,
    variation: text_style::Variation,
}

fn font_only(workspace: &Workspace<'_>, style: u64) -> Result<Option<FontOverrides>> {
    text_style::font_overrides(&ParagraphStyles(workspace), style, limits()).map_err(text_error)
}

fn apply(
    overrides: &mut FontOverrides,
    inherited_font: &ChartFont,
    inherited_size: TextPointSize,
    change: &Change,
) {
    match change {
        Change::Font(target) => {
            overrides.font = (inherited_font.font != target.font).then(|| target.font.clone());
            overrides.bold = (inherited_font.bold != target.bold).then_some(target.bold);
            overrides.italic = (inherited_font.italic != target.italic).then_some(target.italic);
        },
        Change::Size(target) => {
            overrides.point_size = (inherited_size != *target).then_some(*target);
        },
        Change::ResetFont | Change::ResetSize => {},
    }
}

/// Apply `change` to chart drawable `chart`. Returns `false` when nothing
/// changes (the value is already the requested one, or nothing to reset).
///
/// # Errors
///
/// An invalid chart, paragraph style, stylesheet or registry; a write that
/// does not read back.
pub fn set(workspace: &mut Workspace<'_>, chart: u64, change: &Change) -> Result<bool> {
    match change {
        Change::Font(target) if font(workspace, chart)? == *target => return Ok(false),
        Change::Size(target) if size(workspace, chart)? == *target => return Ok(false),
        _ => {},
    }
    let unique = unique(&paragraph_styles(workspace, chart)?);
    let mut next = workspace.next_object_identifier()?;
    let mut replacements = BTreeMap::new();
    let mut additions = Vec::new();
    let mut candidates = 0usize;
    for &style in &unique {
        let (_, payload) = style_payload(workspace, style)?;
        let (parent, stylesheet) = style_links(payload)?;
        let stylesheet = stylesheet.ok_or(Error::Invalid("a paragraph style has no stylesheet"))?;
        let direct = font_only(workspace, style)?;
        let (base, mut overrides) = match change {
            Change::Font(_) | Change::Size(_) => match direct {
                Some(overrides) => (
                    parent.ok_or(Error::Invalid("a paragraph-style variation has no parent"))?,
                    overrides,
                ),
                None => (style, FontOverrides::default()),
            },
            Change::ResetFont | Change::ResetSize => {
                let Some(mut overrides) = direct else {
                    continue;
                };
                let present = if *change == Change::ResetFont {
                    let present = overrides.font.is_some()
                        || overrides.bold.is_some()
                        || overrides.italic.is_some();
                    overrides.font = None;
                    overrides.bold = None;
                    overrides.italic = None;
                    present
                } else {
                    overrides.point_size.take().is_some()
                };
                if !present {
                    continue;
                }
                (
                    parent.ok_or(Error::Invalid("a paragraph-style variation has no parent"))?,
                    overrides,
                )
            },
        };
        candidates += 1;
        if matches!(change, Change::Font(_) | Change::Size(_)) {
            let inherited_font = effective_font(workspace, base)?;
            let inherited_size = text_style_of(workspace, base)?.point_size;
            apply(&mut overrides, &inherited_font, inherited_size, change);
        }
        if overrides.is_empty() {
            replacements.insert(style, base);
            continue;
        }
        let identifier = next;
        next = next.checked_add(1).ok_or(Error::IdentifierOverflow)?;
        let variation = text_style::font_variation(base, stylesheet, &overrides, limits())
            .map_err(text_error)?;
        replacements.insert(style, identifier);
        additions.push(Addition {
            identifier,
            parent: base,
            stylesheet,
            variation,
        });
    }
    if candidates == 0 {
        return Ok(false);
    }
    let chart_location = workspace.locate(chart)?;
    let chart_member = workspace.member(chart_location).to_owned();
    repoint(workspace, chart, &replacements)?;
    let last = additions.last().map(|addition| addition.identifier);
    for addition in additions {
        insert(workspace, &chart_member, addition)?;
    }
    if let Some(last) = last {
        workspace.set_last_object_identifier(last)?;
    }
    prune(workspace, &chart_member, &unique)?;
    match change {
        Change::Font(target) if font(workspace, chart)? != *target => {
            Err(Error::Invalid("a chart-font write does not read back"))
        },
        Change::Size(target) if size(workspace, chart)? != *target => {
            Err(Error::Invalid("a chart font-size write does not read back"))
        },
        _ => Ok(true),
    }
}

fn repoint(
    workspace: &mut Workspace<'_>,
    chart: u64,
    replacements: &BTreeMap<u64, u64>,
) -> Result<()> {
    let (message, payload) = chart_message(workspace, chart)?;
    let data = wire::transform_length_delimited_fields_at_path(
        &payload,
        &[CHART_EXTENSION_FIELD, CHART_PARAGRAPH_STYLES_FIELD],
        |reference| -> Result<Vec<u8>> {
            let identifier = graph::varint(reference, REFERENCE_IDENTIFIER_FIELD)?.unwrap_or(0);
            match replacements.get(&identifier) {
                Some(replacement) => Ok(wire::patch_varint_field(
                    reference,
                    REFERENCE_IDENTIFIER_FIELD,
                    true,
                    Some(*replacement),
                )?),
                None => Ok(reference.to_vec()),
            }
        },
    )?;
    workspace.replace_payload(chart, message, data)?;
    let next = paragraph_styles(workspace, chart)?;
    let references = &mut workspace
        .object_mut(chart)?
        .archive_info
        .message_infos
        .get_mut(message)
        .ok_or(Error::Invalid("a chart payload has no metadata"))?
        .object_references;
    references.retain(|identifier| !replacements.contains_key(identifier));
    for identifier in unique(&next) {
        if !references.contains(&identifier) {
            references.push(identifier);
        }
    }
    Ok(())
}

fn insert(workspace: &mut Workspace<'_>, chart_member: &str, addition: Addition) -> Result<()> {
    let (message, sheet) = workspace.unique_payload(
        addition.stylesheet,
        STYLESHEET_MESSAGE_TYPE,
        "a stylesheet must have exactly one payload",
    )?;
    let data = shape::stylesheet_with_variation(sheet, addition.parent, addition.identifier)?;
    workspace.replace_payload(addition.stylesheet, message, data)?;
    let info = workspace
        .object_mut(addition.stylesheet)?
        .archive_info
        .message_infos
        .get_mut(message)
        .ok_or(Error::Invalid("a stylesheet payload has no metadata"))?;
    if info.object_references.contains(&addition.identifier) {
        return Err(Error::Invalid(
            "the stylesheet metadata already references a new style",
        ));
    }
    info.object_references.push(addition.identifier);
    let references = addition.variation.references().to_vec();
    let mut object = ArchiveObject::new(
        addition.identifier,
        vec![RawMessage {
            type_: PARAGRAPH_STYLE_MESSAGE_TYPE,
            data: addition.variation.into_payload(),
        }],
        workspace.limits(),
    )?;
    let info = &mut object.archive_info.message_infos[0];
    info.versions = text_style::MESSAGE_VERSION.to_vec().into();
    info.object_references.extend(references);
    let style_location = workspace.locate(addition.stylesheet)?;
    workspace.insert(style_location, object)?;
    let style_member = workspace.member(style_location).to_owned();
    if let Some(style_component) = workspace.component_for_member(&style_member)? {
        workspace.add_object_uuids(style_component, &[addition.identifier])?;
        if let Some(owner) = workspace.component_for_member(chart_member)?
            && owner != style_component
        {
            workspace.add_external_reference(owner, style_component, addition.identifier)?;
        }
    }
    Ok(())
}

/// Whether any object other than its stylesheet or child styles references
/// `style` in its metadata.
fn has_consumers(workspace: &Workspace<'_>, style: u64) -> Result<bool> {
    for location in workspace.locations() {
        for object in &workspace.archive_at(location)?.objects {
            for (message, info) in object
                .messages
                .iter()
                .zip(&object.archive_info.message_infos)
            {
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
                return Ok(true);
            }
        }
    }
    Ok(false)
}

fn owner_reference_if_unused(
    workspace: &mut Workspace<'_>,
    chart_member: &str,
    style_member: &str,
    style: u64,
) -> Result<()> {
    let Some(style_component) = workspace.component_for_member(style_member)? else {
        return Ok(());
    };
    let Some(owner) = workspace.component_for_member(chart_member)? else {
        return Ok(());
    };
    if owner == style_component {
        return Ok(());
    }
    for location in workspace.locations() {
        let member = workspace.member(location).to_owned();
        if workspace.component_for_member(&member)? == Some(owner)
            && workspace
                .archive_at(location)?
                .objects
                .iter()
                .any(|object| {
                    object
                        .archive_info
                        .message_infos
                        .iter()
                        .any(|info| info.object_references.contains(&style))
                })
        {
            return Ok(());
        }
    }
    workspace.remove_external_reference(owner, style_component, style)
}

fn prune(workspace: &mut Workspace<'_>, chart_member: &str, previous: &[u64]) -> Result<()> {
    let mut removed = Vec::new();
    for &style in previous {
        let location = workspace.locate(style)?;
        let style_member = workspace.member(location).to_owned();
        if font_only(workspace, style)?.is_none() || has_consumers(workspace, style)? {
            owner_reference_if_unused(workspace, chart_member, &style_member, style)?;
            continue;
        }
        let (_, payload) = style_payload(workspace, style)?;
        let (parent, stylesheet) = style_links(payload)?;
        let parent = parent.ok_or(Error::Invalid("a paragraph-style variation has no parent"))?;
        let stylesheet = stylesheet.ok_or(Error::Invalid("a paragraph style has no stylesheet"))?;
        let (message, sheet) = workspace.unique_payload(
            stylesheet,
            STYLESHEET_MESSAGE_TYPE,
            "a stylesheet must have exactly one payload",
        )?;
        let data = shape::stylesheet_without_variation(sheet, parent, style)?;
        workspace.replace_payload(stylesheet, message, data)?;
        let info = workspace
            .object_mut(stylesheet)?
            .archive_info
            .message_infos
            .get_mut(message)
            .ok_or(Error::Invalid("a stylesheet payload has no metadata"))?;
        if info
            .object_references
            .iter()
            .filter(|id| **id == style)
            .count()
            != 1
        {
            return Err(Error::Invalid(
                "the stylesheet metadata must reference a removed style once",
            ));
        }
        info.object_references.retain(|id| *id != style);
        for field in &mut info.field_infos {
            field.object_references.retain(|id| *id != style);
        }
        workspace.remove(style)?;
        if let Some(style_component) = workspace.component_for_member(&style_member)? {
            workspace.remove_object_uuids(style_component, &[style])?;
            workspace.remove_external_references_to(style_component, style)?;
            if let Some(owner) = workspace.component_for_member(chart_member)?
                && owner != style_component
            {
                workspace.add_external_reference(owner, style_component, parent)?;
            }
        }
        removed.push(style);
    }
    if !removed.is_empty() {
        workspace.release_identifier_suffix(&removed)?;
    }
    Ok(())
}
