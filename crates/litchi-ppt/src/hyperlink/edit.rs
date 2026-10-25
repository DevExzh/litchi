//! Click hyperlinks of shapes and their document-level definitions
//! (record 0323).
//!
//! A shape's click action is the `InteractiveInfo` container (instance 0)
//! in its `OfficeArtClientData`; a URL or slide link names an
//! `ExHyperlinkContainer` of the live document's `ExObjListContainer`, and
//! its screen tip lives in the `ExHyperlink9Container` of the document's
//! `___PPT9` programmable tag ([MS-PPT] 2.10.17, 2.10.21, 2.3.?). An edit
//! rewrites only those records: the shape's client data, and for a new
//! definition the object list (whose identifier seed grows) and the tag.

use litchi_core::Position;

use super::model::{
    Hyperlink, Hyperlinks, Interaction, InteractionAction, InteractionJump, InteractionLinkTarget,
    InteractionTrigger, InteractiveInfoAtom,
};
use crate::binary_tag_edit;
use crate::consts::RecordType;
use crate::package::{Error, SourceLimits};
use crate::record_wire::{self as wire};
use crate::slide_order::{Error as EditError, Refusal, Result};

const ELEMENT: &str = "hyperlink";
const SP_CONTAINER: u16 = 0xF004;
const CLIENT_DATA: u16 = 0xF011;
const CLIENT_TEXTBOX: u16 = 0xF00D;
const INTERACTIVE_INFO: u16 = RecordType::InteractiveInfo as u16;
const CSTRING: u16 = RecordType::CString as u16;
/// The `PptOfficeArtClientData` records ordered before the click action:
/// `ShapeFlagsAtom`, `ShapeFlags10Atom`, `ExObjRefAtom`, `AnimationInfo`.
const BEFORE_CLICK: [u16; 4] = [0x0BDB, 0x0BDC, 0x0BC1, RecordType::AnimationInfo as u16];
/// The `DocumentContainer` record an `ExObjListContainer` follows.
const DOCUMENT_ATOM: u16 = RecordType::DocumentAtom as u16;
/// The `PP9DocBinaryTagExtension` records ordered before hyperlink
/// extensions: text master styles, the bullet blip collection, the text
/// defaults atom and the kinsoku container.
const BEFORE_HYPERLINK9: [u16; 4] = [0x0FB2, 0x07F8, 0x0FB4, RecordType::Kinsoku as u16];

#[derive(Clone, Copy, Default)]
struct SizeDelta {
    added: usize,
    removed: usize,
}

impl SizeDelta {
    fn from_lengths(old: usize, new: usize) -> Self {
        if new >= old {
            Self {
                added: new - old,
                removed: 0,
            }
        } else {
            Self {
                added: 0,
                removed: old - new,
            }
        }
    }

    fn combine(self, other: Self) -> Result<Self> {
        let added = self
            .added
            .checked_add(other.added)
            .ok_or_else(|| Error::ResourceLimit("PPT projected size overflows".into()))?;
        let removed = self
            .removed
            .checked_add(other.removed)
            .ok_or_else(|| Error::ResourceLimit("PPT projected size overflows".into()))?;
        if added >= removed {
            Ok(Self {
                added: added - removed,
                removed: 0,
            })
        } else {
            Ok(Self {
                added: 0,
                removed: removed - added,
            })
        }
    }

    fn apply(self, base: usize, context: &str) -> Result<usize> {
        Ok(base
            .checked_add(self.added)
            .and_then(|value| value.checked_sub(self.removed))
            .ok_or_else(|| Error::ResourceLimit(format!("{context} overflows")))?)
    }
}

/// Where a click hyperlink goes.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Target {
    /// An external URL or file path.
    Url(String),
    /// A slide of this presentation, by zero-based position.
    Slide(Position),
    /// A slide-show jump: next, previous, first, last, last viewed slide or
    /// end show.
    Jump(InteractionJump),
}

/// A shape's or text range's click hyperlink.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Click {
    pub target: Target,
    /// The `PowerPoint` 9 screen tip.
    pub tooltip: Option<String>,
}

impl Click {
    /// Resolve a click interaction to its hyperlink, or `None` when the
    /// action is not a hyperlink (a macro, program, OLE verb, media or custom
    /// show) or names no definition. `slide_ids` are the presentation's slide
    /// identifiers in order.
    #[must_use]
    pub fn resolve(
        interaction: &Interaction,
        hyperlinks: &Hyperlinks,
        slide_ids: &[u32],
    ) -> Option<Self> {
        if interaction.trigger != InteractionTrigger::Click {
            return None;
        }
        let jump = |jump: InteractionJump| {
            (!matches!(jump, InteractionJump::None)).then_some(Target::Jump(jump))
        };
        let (target, definition) = match interaction.action {
            InteractionAction::Jump => (jump(interaction.jump)?, None),
            InteractionAction::Hyperlink => {
                let definition = interaction.hyperlink(hyperlinks);
                let target = match interaction.link_target {
                    InteractionLinkTarget::NextSlide => Target::Jump(InteractionJump::NextSlide),
                    InteractionLinkTarget::PreviousSlide => {
                        Target::Jump(InteractionJump::PreviousSlide)
                    },
                    InteractionLinkTarget::FirstSlide => Target::Jump(InteractionJump::FirstSlide),
                    InteractionLinkTarget::LastSlide => Target::Jump(InteractionJump::LastSlide),
                    InteractionLinkTarget::SlideNumber => {
                        Target::Slide(slide_of(definition?, slide_ids)?)
                    },
                    InteractionLinkTarget::Url
                    | InteractionLinkTarget::OtherFile
                    | InteractionLinkTarget::OtherPresentation
                    | InteractionLinkTarget::Nil
                    | InteractionLinkTarget::CustomShow => {
                        Target::Url(definition?.target.clone().filter(|url| !url.is_empty())?)
                    },
                };
                (target, definition)
            },
            _ => return None,
        };
        Some(Self {
            target,
            tooltip: definition
                .and_then(|definition| definition.extension.as_ref())
                .and_then(|extension| extension.screen_tip.clone()),
        })
    }
}

/// The slide a slide link's location (`slideId,slideNumber,title`) names.
fn slide_of(definition: &Hyperlink, slide_ids: &[u32]) -> Option<Position> {
    let location = definition.location.as_deref()?;
    let mut fields = location.split(',');
    let id = fields.next()?.trim().parse::<u32>().ok();
    let number = fields
        .next()
        .and_then(|field| field.trim().parse::<usize>().ok());
    if let Some(position) = id.and_then(|id| slide_ids.iter().position(|slide| *slide == id)) {
        return Some(Position::new(position));
    }
    number
        .and_then(|number| number.checked_sub(1))
        .filter(|position| *position < slide_ids.len())
        .map(Position::new)
}

/// The click atom of `target`, naming `hyperlink_id` for a URL or slide.
pub(crate) fn click_atom(target: &Target, hyperlink_id: u32) -> [u8; 16] {
    let (action, jump, link_target, id) = match target {
        Target::Url(_) => (
            InteractionAction::Hyperlink,
            InteractionJump::None,
            InteractionLinkTarget::Url,
            hyperlink_id,
        ),
        Target::Slide(_) => (
            InteractionAction::Hyperlink,
            InteractionJump::None,
            InteractionLinkTarget::SlideNumber,
            hyperlink_id,
        ),
        Target::Jump(jump) => (
            InteractionAction::Jump,
            *jump,
            match jump {
                InteractionJump::NextSlide => InteractionLinkTarget::NextSlide,
                InteractionJump::PreviousSlide => InteractionLinkTarget::PreviousSlide,
                InteractionJump::FirstSlide => InteractionLinkTarget::FirstSlide,
                InteractionJump::LastSlide => InteractionLinkTarget::LastSlide,
                _ => InteractionLinkTarget::Nil,
            },
            0,
        ),
    };
    InteractiveInfoAtom {
        sound_id: 0,
        hyperlink_id: id,
        action,
        ole_verb: 0,
        jump,
        animated: true,
        stop_sound: false,
        custom_show_return: false,
        visited: false,
        link_target,
        unused: [0; 3],
    }
    .to_payload()
}

/// The click container holding `atom`.
pub(crate) fn click_container(atom: &[u8; 16], limits: &SourceLimits) -> Result<Vec<u8>> {
    // Admit the complete container before allocating its encoded child.  The
    // child is fixed-width, so its exact framed size is known up front.
    let length = 8usize
        .checked_add(8)
        .and_then(|value| value.checked_add(atom.len()))
        .ok_or_else(|| Error::ResourceLimit("InteractiveInfo length overflows".into()))?;
    let aggregate = atom
        .len()
        .checked_add(8)
        .and_then(|value| value.checked_add(atom.len()))
        .ok_or_else(|| Error::ResourceLimit("InteractiveInfo aggregate overflows".into()))?;
    ensure_container_budget(length, aggregate, limits, "InteractiveInfo")?;
    let child = wire::encode(RecordType::InteractiveInfoAtom as u16, 0, 0, atom)?;
    bounded_container(INTERACTIVE_INFO, 0, &[child], limits, "InteractiveInfo")
}

/// Replace (`Some`) or remove (`None`) the click action of the shape with
/// `OfficeArt` ID `shape_id`; `None` when nothing changes.
pub(crate) fn set_shape_click(
    slide: &[u8],
    shape_id: u32,
    click: Option<&[u8]>,
    limits: &SourceLimits,
) -> Result<Option<Vec<u8>>> {
    let slide_aggregate = binary_tag_edit::serialized_payload_budget(slide)?;
    ensure_container_budget(
        slide.len(),
        slide_aggregate,
        limits,
        "PPT slide hyperlink edit",
    )?;
    let (projected_len, projected_aggregate) =
        projected_slide_click_budget(slide, shape_id, click, slide_aggregate)?;
    ensure_container_budget(
        projected_len,
        projected_aggregate,
        limits,
        "PPT slide hyperlink edit",
    )?;
    let mut changed = false;
    let rewritten = rewrite_slide_shape(slide, shape_id, &mut |shape| {
        rewrite_shape(shape, click, &mut changed, limits)
    })?;
    ensure_serialized_record(&rewritten, limits, "PPT slide hyperlink edit")?;
    Ok(changed.then_some(rewritten))
}

/// Project the exact aggregate charge of a click edit without materializing
/// the rewritten OfficeArt records.  `PPDrawing` is a raw record to the PPT
/// parser, while its `Slide` owner is a parsed container: a byte delta inside
/// the drawing is therefore charged once for the drawing payload and once for
/// the slide payload.  The returned values are derived from the actual shape
/// and wrapper lengths rather than an upper-bound multiplier.
fn projected_slide_click_budget(
    slide: &[u8],
    shape_id: u32,
    click: Option<&[u8]>,
    slide_aggregate: usize,
) -> Result<(usize, usize)> {
    let mut drawings = 0usize;
    let mut matches = 0usize;
    let mut drawing_delta = SizeDelta::default();
    for child in wire::children(slide)? {
        let child = child?;
        if wire::record_type(child)? != RecordType::PPDrawing as u16 {
            continue;
        }
        drawings = drawings
            .checked_add(1)
            .ok_or_else(|| Error::ResourceLimit("PPT drawing count overflows".into()))?;
        let (drawing_matches, delta) = projected_drawing_click_delta(child, shape_id, click)?;
        matches = matches
            .checked_add(drawing_matches)
            .ok_or_else(|| Error::ResourceLimit("PPT shape count overflows".into()))?;
        if let Some(delta) = delta {
            drawing_delta = drawing_delta.combine(delta)?;
        }
    }
    if drawings != 1 {
        return Err(
            Error::Corrupted("selected slide has ambiguous PPDrawing ownership".into()).into(),
        );
    }
    match matches {
        0 => Err(EditError::Refused(Refusal::ElementNotFound {
            element: "shape",
        })),
        1 => {
            let projected_len = drawing_delta.apply(slide.len(), "PPT slide hyperlink length")?;
            let projected_aggregate_delta = drawing_delta.combine(drawing_delta)?;
            let projected_aggregate = projected_aggregate_delta
                .apply(slide_aggregate, "PPT slide hyperlink aggregate")?;
            Ok((projected_len, projected_aggregate))
        },
        _ => Err(EditError::Refused(Refusal::UnsupportedElement {
            element: ELEMENT,
            reason: "the shape identifier is ambiguous in the drawing",
        })),
    }
}

fn projected_drawing_click_delta(
    drawing: &[u8],
    shape_id: u32,
    click: Option<&[u8]>,
) -> Result<(usize, Option<SizeDelta>)> {
    let mut matches = 0usize;
    let mut delta = None;
    for child in wire::children(drawing)? {
        projected_officeart_click_delta(child?, shape_id, click, 0, &mut matches, &mut delta)?;
    }
    Ok((matches, delta))
}

fn projected_officeart_click_delta(
    record: &[u8],
    shape_id: u32,
    click: Option<&[u8]>,
    depth: usize,
    matches: &mut usize,
    delta: &mut Option<SizeDelta>,
) -> Result<()> {
    if depth >= 128 {
        return Err(
            Error::Corrupted("OfficeArt nesting exceeds the hyperlink-edit limit".into()).into(),
        );
    }
    if wire::record_type(record)? == SP_CONTAINER && wire::shape_id_of(record)? == Some(shape_id) {
        *matches = matches
            .checked_add(1)
            .ok_or_else(|| Error::ResourceLimit("PPT shape count overflows".into()))?;
        let (old_len, new_len) = projected_shape_lengths(record, click)?;
        let shape_delta = SizeDelta::from_lengths(old_len, new_len);
        *delta = Some(match *delta {
            Some(existing) => existing.combine(shape_delta)?,
            None => shape_delta,
        });
        return Ok(());
    }
    if wire::record_version(record)? != 0x0F {
        return Ok(());
    }
    for child in wire::children(record)? {
        projected_officeart_click_delta(child?, shape_id, click, depth + 1, matches, delta)?;
    }
    Ok(())
}

fn projected_shape_lengths(shape: &[u8], click: Option<&[u8]>) -> Result<(usize, usize)> {
    let mut client_data = None;
    for child in wire::children(shape)? {
        let child = child?;
        if wire::record_type(child)? == CLIENT_DATA && client_data.replace(child).is_some() {
            return Err(EditError::Refused(Refusal::UnsupportedElement {
                element: ELEMENT,
                reason: "the shape has several client-data records",
            }));
        }
    }
    let Some(client_data) = client_data else {
        let new_len = match click {
            Some(click) => shape
                .len()
                .checked_add(8)
                .and_then(|value| value.checked_add(click.len()))
                .ok_or_else(|| Error::ResourceLimit("OfficeArt shape length overflows".into()))?,
            None => shape.len(),
        };
        return Ok((shape.len(), new_len));
    };

    let mut current = None;
    let mut child_count = 0usize;
    for child in wire::children(client_data)? {
        let child = child?;
        child_count = child_count
            .checked_add(1)
            .ok_or_else(|| Error::ResourceLimit("OfficeArt ClientData count overflows".into()))?;
        if wire::record_type(child)? == INTERACTIVE_INFO
            && wire::record_instance(child)? == 0
            && current.replace(child).is_some()
        {
            return Err(EditError::Refused(Refusal::UnsupportedElement {
                element: ELEMENT,
                reason: "the shape has several click actions",
            }));
        }
    }
    let new_client_data_len = match (current, click) {
        (Some(current), Some(click)) => client_data
            .len()
            .checked_sub(current.len())
            .and_then(|length| length.checked_add(click.len()))
            .ok_or_else(|| Error::ResourceLimit("OfficeArt ClientData length overflows".into()))?,
        (Some(_current), None) if child_count == 1 => 0,
        (Some(current), None) => client_data
            .len()
            .checked_sub(current.len())
            .ok_or_else(|| Error::ResourceLimit("OfficeArt ClientData length underflows".into()))?,
        (None, Some(click)) => client_data
            .len()
            .checked_add(click.len())
            .ok_or_else(|| Error::ResourceLimit("OfficeArt ClientData length overflows".into()))?,
        (None, None) => client_data.len(),
    };
    let new_shape_len = shape
        .len()
        .checked_sub(client_data.len())
        .and_then(|length| length.checked_add(new_client_data_len))
        .ok_or_else(|| Error::ResourceLimit("OfficeArt shape length underflows".into()))?;
    Ok((shape.len(), new_shape_len))
}

/// Rewrite the one `OfficeArt` shape container with ID `shape_id` in the
/// slide's single `PPDrawing` through `rewrite`; every other record keeps
/// its bytes and the owning envelopes are re-framed.
///
/// # Errors
///
/// Refuses a slide whose drawing ownership is ambiguous, a missing shape
/// (`ElementNotFound`) and an ambiguous shape identifier.
pub(crate) fn rewrite_slide_shape(
    slide: &[u8],
    shape_id: u32,
    rewrite: &mut dyn FnMut(&[u8]) -> Result<Vec<u8>>,
) -> Result<Vec<u8>> {
    let mut matches = 0usize;
    let mut drawings = 0usize;
    let rewritten = rewrite_drawings(slide, &mut |drawing| {
        drawings += 1;
        let data = rewrite_officeart(wire::payload(drawing)?, shape_id, rewrite, 0, &mut matches)?;
        Ok(wire::rebuild(drawing, &data)?)
    })?;
    if drawings != 1 {
        return Err(
            Error::Corrupted("selected slide has ambiguous PPDrawing ownership".into()).into(),
        );
    }
    match matches {
        0 => Err(EditError::Refused(Refusal::ElementNotFound {
            element: "shape",
        })),
        1 => Ok(rewritten),
        _ => Err(EditError::Refused(Refusal::UnsupportedElement {
            element: ELEMENT,
            reason: "the shape identifier is ambiguous in the drawing",
        })),
    }
}

/// Rewrite every `PPDrawing` child of a slide record.
fn rewrite_drawings(
    slide: &[u8],
    visit: &mut impl FnMut(&[u8]) -> Result<Vec<u8>>,
) -> Result<Vec<u8>> {
    let mut data = Vec::with_capacity(slide.len().saturating_add(32));
    for child in wire::children(slide)? {
        let child = child?;
        if wire::record_type(child)? == RecordType::PPDrawing as u16 {
            data.extend_from_slice(&visit(child)?);
        } else {
            data.extend_from_slice(child);
        }
    }
    Ok(wire::rebuild(slide, &data)?)
}

fn rewrite_officeart(
    record: &[u8],
    shape_id: u32,
    rewrite: &mut dyn FnMut(&[u8]) -> Result<Vec<u8>>,
    depth: usize,
    matches: &mut usize,
) -> Result<Vec<u8>> {
    if depth >= 128 {
        return Err(
            Error::Corrupted("OfficeArt nesting exceeds the hyperlink-edit limit".into()).into(),
        );
    }
    if wire::record_type(record)? == SP_CONTAINER && wire::shape_id_of(record)? == Some(shape_id) {
        *matches += 1;
        return rewrite(record);
    }
    if wire::record_version(record)? != 0x0F {
        return Ok(record.to_vec());
    }
    let mut data = Vec::with_capacity(wire::payload(record)?.len());
    for child in wire::children(record)? {
        data.extend_from_slice(&rewrite_officeart(
            child?,
            shape_id,
            rewrite,
            depth + 1,
            matches,
        )?);
    }
    Ok(wire::rebuild(record, &data)?)
}

fn rewrite_shape(
    shape: &[u8],
    click: Option<&[u8]>,
    changed: &mut bool,
    limits: &SourceLimits,
) -> Result<Vec<u8>> {
    let children = wire::child_list(shape)?;
    let mut client_data = None;
    for (index, child) in children.iter().enumerate() {
        if wire::record_type(child)? == CLIENT_DATA && client_data.replace(index).is_some() {
            return Err(EditError::Refused(Refusal::UnsupportedElement {
                element: ELEMENT,
                reason: "the shape has several client-data records",
            }));
        }
    }
    let rewritten = match client_data {
        Some(index) => rewrite_client_data(children[index], click, limits)?,
        None => match click {
            None => return Ok(shape.to_vec()),
            Some(click) => {
                let client_data_len = click
                    .len()
                    .checked_add(8)
                    .ok_or_else(|| Error::ResourceLimit("ClientData length overflows".into()))?;
                let click_aggregate = binary_tag_edit::serialized_payload_budget(click)?;
                let client_data_aggregate = client_data_len
                    .checked_sub(8)
                    .and_then(|payload| payload.checked_add(click_aggregate))
                    .ok_or_else(|| {
                        Error::ResourceLimit("OfficeArt ClientData aggregate overflows".into())
                    })?;
                ensure_container_budget(
                    client_data_len,
                    client_data_aggregate,
                    limits,
                    "OfficeArt ClientData",
                )?;
                Some(wire::container(CLIENT_DATA, 0, &[click.to_vec()])?)
            },
        },
    };
    let mut output_len = 8usize;
    for (index, child) in children.iter().enumerate() {
        let replacement = if Some(index) == client_data {
            rewritten.as_ref().map_or(0, Vec::len)
        } else {
            child.len()
        };
        output_len = output_len
            .checked_add(replacement)
            .ok_or_else(|| Error::ResourceLimit("OfficeArt shape length overflows".into()))?;
    }
    if client_data.is_none() && rewritten.is_some() {
        output_len = output_len
            .checked_add(rewritten.as_ref().map_or(0, Vec::len))
            .ok_or_else(|| Error::ResourceLimit("OfficeArt shape length overflows".into()))?;
    }
    let payload = output_len
        .checked_sub(8)
        .ok_or_else(|| Error::ResourceLimit("OfficeArt shape underflows".into()))?;
    ensure_container_budget(output_len, payload, limits, "OfficeArt shape")?;
    let mut data = Vec::with_capacity(shape.len().saturating_add(32));
    let mut inserted = false;
    for (index, child) in children.iter().enumerate() {
        if Some(index) == client_data {
            if let Some(rewritten) = &rewritten
                && rewritten.as_slice() != *child
            {
                *changed = true;
                data.extend_from_slice(rewritten);
            } else if rewritten.is_none() {
                *changed = true;
            } else {
                data.extend_from_slice(child);
            }
            continue;
        }
        if client_data.is_none() && !inserted && wire::record_type(child)? == CLIENT_TEXTBOX {
            // `OfficeArtClientData` precedes `OfficeArtClientTextbox`.
            if let Some(rewritten) = &rewritten {
                data.extend_from_slice(rewritten);
                *changed = true;
            }
            inserted = true;
        }
        data.extend_from_slice(child);
    }
    if client_data.is_none() && !inserted {
        if let Some(rewritten) = &rewritten {
            data.extend_from_slice(rewritten);
            *changed = true;
        }
    }
    Ok(wire::rebuild(shape, &data)?)
}

/// The client data with its click action replaced; `None` when removing the
/// action leaves it empty.
fn rewrite_client_data(
    record: &[u8],
    click: Option<&[u8]>,
    limits: &SourceLimits,
) -> Result<Option<Vec<u8>>> {
    let children = wire::child_list(record)?;
    let mut current = None;
    for (index, child) in children.iter().enumerate() {
        if wire::record_type(child)? == INTERACTIVE_INFO
            && wire::record_instance(child)? == 0
            && current.replace(index).is_some()
        {
            return Err(EditError::Refused(Refusal::UnsupportedElement {
                element: ELEMENT,
                reason: "the shape has several click actions",
            }));
        }
    }
    let mut output_len = 8usize;
    for (index, child) in children.iter().enumerate() {
        let replacement_len = if Some(index) == current {
            click.map_or(0, <[u8]>::len)
        } else {
            child.len()
        };
        output_len = output_len
            .checked_add(replacement_len)
            .ok_or_else(|| Error::ResourceLimit("ClientData length overflows".into()))?;
    }
    if current.is_none() && click.is_some() {
        output_len = output_len
            .checked_add(click.map_or(0, <[u8]>::len))
            .ok_or_else(|| Error::ResourceLimit("ClientData length overflows".into()))?;
    }
    let click_aggregate = click
        .map(binary_tag_edit::serialized_payload_budget)
        .transpose()?;
    let mut nested_aggregate = 0usize;
    for (index, child) in children.iter().enumerate() {
        if Some(index) == current {
            if let Some(click_aggregate) = click_aggregate {
                nested_aggregate = nested_aggregate
                    .checked_add(click_aggregate)
                    .ok_or_else(|| Error::ResourceLimit("ClientData aggregate overflows".into()))?;
            }
        } else {
            nested_aggregate = nested_aggregate
                .checked_add(binary_tag_edit::serialized_payload_budget(child)?)
                .ok_or_else(|| Error::ResourceLimit("ClientData aggregate overflows".into()))?;
        }
    }
    if current.is_none() {
        if let Some(click_aggregate) = click_aggregate {
            nested_aggregate = nested_aggregate
                .checked_add(click_aggregate)
                .ok_or_else(|| Error::ResourceLimit("ClientData aggregate overflows".into()))?;
        }
    }
    let aggregate = output_len
        .checked_sub(8)
        .and_then(|payload| payload.checked_add(nested_aggregate))
        .ok_or_else(|| Error::ResourceLimit("ClientData aggregate overflows".into()))?;
    ensure_container_budget(output_len, aggregate, limits, "OfficeArt ClientData")?;
    let mut records = children
        .iter()
        .map(|child| child.to_vec())
        .collect::<Vec<_>>();
    match (current, click) {
        (Some(index), Some(click)) => records[index] = click.to_vec(),
        (Some(index), None) => {
            records.remove(index);
        },
        (None, Some(click)) => {
            let mut position = 0;
            for (index, child) in children.iter().enumerate() {
                if BEFORE_CLICK.contains(&wire::record_type(child)?) {
                    position = index + 1;
                }
            }
            records.insert(position, click.to_vec());
        },
        (None, None) => return Ok(Some(record.to_vec())),
    }
    if records.is_empty() {
        return Ok(None);
    }
    Ok(Some(wire::rebuild(record, &records.concat())?))
}

/// One new hyperlink definition: its target string or slide location.
pub(crate) struct Definition<'a> {
    pub(crate) friendly_name: &'a str,
    pub(crate) target: Option<&'a str>,
    pub(crate) location: Option<&'a str>,
    pub(crate) tooltip: Option<&'a str>,
}

/// Append a hyperlink definition to the live document; returns the new
/// document record and the definition's identifier.
pub(crate) fn add_definition(
    document: &[u8],
    definition: &Definition<'_>,
    limits: &SourceLimits,
) -> Result<(Vec<u8>, u32)> {
    ensure_serialized_record(document, limits, "PPT document hyperlink edit")?;
    if wire::record_type(document)? != RecordType::Document as u16 {
        return Err(Error::Corrupted("hyperlink owner is not a DocumentContainer".into()).into());
    }
    let children = wire::child_list(document)?;
    let mut list = None;
    let mut info = None;
    for (index, child) in children.iter().enumerate() {
        match wire::record_type(child)? {
            kind if kind == RecordType::ExObjList as u16 => {
                if list.replace(index).is_some() {
                    return Err(unsupported(
                        "the document has several external-object lists",
                    ));
                }
            },
            kind if kind == RecordType::DocInfoList as u16 => {
                if info.replace(index).is_some() {
                    return Err(unsupported("the document has several information lists"));
                }
            },
            _ => {},
        }
    }
    let (list_record, id) = match list {
        Some(index) => append_to_list(children[index], definition, limits)?,
        None => {
            let id: i32 = 1;
            let (hyperlink_len, hyperlink_aggregate) = hyperlink_record_budget(definition, limits)?;
            let list_len = 8usize
                .checked_add(12)
                .and_then(|value| value.checked_add(hyperlink_len))
                .ok_or_else(|| {
                    Error::ResourceLimit("ExObjListContainer length overflows".into())
                })?;
            let list_payload = list_len
                .checked_sub(8)
                .ok_or_else(|| Error::ResourceLimit("ExObjListContainer underflows".into()))?;
            let list_aggregate = list_payload
                .checked_add(4)
                .and_then(|value| value.checked_add(hyperlink_aggregate))
                .ok_or_else(|| {
                    Error::ResourceLimit("ExObjListContainer aggregate overflows".into())
                })?;
            ensure_container_budget(list_len, list_aggregate, limits, "ExObjListContainer")?;
            let atom = bounded_encode(
                RecordType::ExObjListAtom as u16,
                0,
                0,
                &id.to_le_bytes(),
                limits,
                "ExObjListAtom",
            )?;
            let hyperlink = hyperlink_record(id, definition, limits)?;
            let record = bounded_container(
                RecordType::ExObjList as u16,
                0,
                &[atom, hyperlink],
                limits,
                "ExObjListContainer",
            )?;
            (record, id.cast_unsigned())
        },
    };
    let info_record = match (definition.tooltip, info) {
        (None, _) => None,
        (Some(tip), Some(index)) => Some(add_screen_tip(children[index], id, tip, limits)?),
        (Some(_), None) => {
            return Err(unsupported(
                "the document has no information list for a screen tip",
            ));
        },
    };
    let mut output_len = 8usize;
    for (index, child) in children.iter().enumerate() {
        let replacement = if Some(index) == list {
            list_record.len()
        } else if Some(index) == info {
            info_record.as_ref().map_or(child.len(), Vec::len)
        } else {
            child.len()
        };
        output_len = output_len
            .checked_add(replacement)
            .ok_or_else(|| Error::ResourceLimit("hyperlink document length overflows".into()))?;
    }
    let mut placed = list.is_some();
    if !placed {
        output_len = output_len
            .checked_add(list_record.len())
            .ok_or_else(|| Error::ResourceLimit("hyperlink document length overflows".into()))?;
    }
    let mut children_aggregate = 0usize;
    for (index, child) in children.iter().enumerate() {
        let aggregate = if Some(index) == list {
            binary_tag_edit::serialized_payload_budget(&list_record)?
        } else if Some(index) == info {
            let record = info_record.as_deref().unwrap_or(child);
            binary_tag_edit::serialized_payload_budget(record)?
        } else {
            binary_tag_edit::serialized_payload_budget(child)?
        };
        children_aggregate = children_aggregate
            .checked_add(aggregate)
            .ok_or_else(|| Error::ResourceLimit("hyperlink document aggregate overflows".into()))?;
    }
    if !placed {
        children_aggregate = children_aggregate
            .checked_add(binary_tag_edit::serialized_payload_budget(&list_record)?)
            .ok_or_else(|| Error::ResourceLimit("hyperlink document aggregate overflows".into()))?;
    }
    let document_payload = output_len
        .checked_sub(8)
        .ok_or_else(|| Error::ResourceLimit("hyperlink document underflows".into()))?;
    let document_aggregate = document_payload
        .checked_add(children_aggregate)
        .ok_or_else(|| Error::ResourceLimit("hyperlink document aggregate overflows".into()))?;
    ensure_container_budget(
        output_len,
        document_aggregate,
        limits,
        "PPT document hyperlink edit",
    )?;
    let mut data = Vec::with_capacity(output_len.saturating_sub(8));
    for (index, child) in children.iter().enumerate() {
        if Some(index) == list {
            data.extend_from_slice(&list_record);
        } else if Some(index) == info
            && let Some(info_record) = &info_record
        {
            data.extend_from_slice(info_record);
        } else {
            data.extend_from_slice(child);
        }
        if !placed && wire::record_type(child)? == DOCUMENT_ATOM {
            // [MS-PPT] 2.4.1: `exObjList` follows `documentAtom`.
            data.extend_from_slice(&list_record);
            placed = true;
        }
    }
    if !placed {
        return Err(Error::Corrupted("DocumentContainer has no DocumentAtom".into()).into());
    }
    let rebuilt = wire::rebuild(document, &data)?;
    ensure_serialized_record(&rebuilt, limits, "PPT document hyperlink edit")?;
    Ok((rebuilt, id))
}

fn unsupported(reason: &'static str) -> EditError {
    EditError::Refused(Refusal::UnsupportedElement {
        element: ELEMENT,
        reason,
    })
}

fn hyperlink_record(
    id: i32,
    definition: &Definition<'_>,
    limits: &SourceLimits,
) -> Result<Vec<u8>> {
    let _ = hyperlink_record_budget(definition, limits)?;
    let mut children = vec![bounded_encode(
        RecordType::ExternalHyperlinkAtom as u16,
        0,
        0,
        &id.to_le_bytes(),
        limits,
        "ExHyperlinkAtom",
    )?];
    children.push(bounded_encode(
        CSTRING,
        0,
        0,
        &bounded_utf16(definition.friendly_name, limits, "hyperlink friendly name")?,
        limits,
        "hyperlink friendly name",
    )?);
    if let Some(target) = definition.target {
        children.push(bounded_encode(
            CSTRING,
            0,
            1,
            &bounded_utf16(target, limits, "hyperlink target")?,
            limits,
            "hyperlink target",
        )?);
    }
    if let Some(location) = definition.location {
        children.push(bounded_encode(
            CSTRING,
            0,
            3,
            &bounded_utf16(location, limits, "hyperlink location")?,
            limits,
            "hyperlink location",
        )?);
    }
    bounded_container(
        RecordType::ExternalHyperlink as u16,
        0,
        &children,
        limits,
        "ExHyperlinkContainer",
    )
}

fn hyperlink_record_budget(
    definition: &Definition<'_>,
    limits: &SourceLimits,
) -> Result<(usize, usize)> {
    let mut payload_len = 12usize;
    let mut nested_payload_len = 4usize;
    for value in [
        Some(definition.friendly_name),
        definition.target,
        definition.location,
    ]
    .into_iter()
    .flatten()
    {
        let string_payload = utf16_payload_len(value, limits, "hyperlink string")?;
        payload_len = payload_len
            .checked_add(8)
            .and_then(|total| total.checked_add(string_payload))
            .ok_or_else(|| Error::ResourceLimit("ExHyperlinkContainer length overflows".into()))?;
        nested_payload_len = nested_payload_len
            .checked_add(string_payload)
            .ok_or_else(|| {
                Error::ResourceLimit("ExHyperlinkContainer aggregate overflows".into())
            })?;
    }
    let length = payload_len
        .checked_add(8)
        .ok_or_else(|| Error::ResourceLimit("ExHyperlinkContainer length overflows".into()))?;
    let aggregate = payload_len
        .checked_add(nested_payload_len)
        .ok_or_else(|| Error::ResourceLimit("ExHyperlinkContainer aggregate overflows".into()))?;
    ensure_container_budget(length, aggregate, limits, "ExHyperlinkContainer")?;
    Ok((length, aggregate))
}

fn hyperlink_extension_budget(tip: &str, limits: &SourceLimits) -> Result<(usize, usize)> {
    let string_payload = utf16_payload_len(tip, limits, "hyperlink screen tip")?;
    let payload_len = 12usize
        .checked_add(8)
        .and_then(|total| total.checked_add(string_payload))
        .and_then(|total| total.checked_add(12))
        .ok_or_else(|| Error::ResourceLimit("ExHyperlink9Container length overflows".into()))?;
    let length = payload_len
        .checked_add(8)
        .ok_or_else(|| Error::ResourceLimit("ExHyperlink9Container length overflows".into()))?;
    let aggregate = payload_len
        .checked_add(4)
        .and_then(|value| value.checked_add(string_payload))
        .and_then(|value| value.checked_add(4))
        .ok_or_else(|| Error::ResourceLimit("ExHyperlink9Container aggregate overflows".into()))?;
    ensure_container_budget(length, aggregate, limits, "ExHyperlink9Container")?;
    Ok((length, aggregate))
}

fn append_to_list(
    list: &[u8],
    definition: &Definition<'_>,
    limits: &SourceLimits,
) -> Result<(Vec<u8>, u32)> {
    ensure_serialized_record(list, limits, "ExObjListContainer")?;
    let children = wire::child_list(list)?;
    let Some(atom) = children.first() else {
        return Err(Error::Corrupted("ExObjListContainer has no ExObjListAtom".into()).into());
    };
    let seed = wire::payload(atom)?;
    if wire::record_type(atom)? != RecordType::ExObjListAtom as u16 || seed.len() != 4 {
        return Err(Error::Corrupted("ExObjListAtom is malformed".into()).into());
    }
    let seed = i32::from_le_bytes([seed[0], seed[1], seed[2], seed[3]]);
    let id = seed
        .checked_add(1)
        .filter(|id| *id > 0)
        .ok_or(EditError::Refused(Refusal::InvalidElementValue {
            element: ELEMENT,
            reason: "the external-object identifier seed is exhausted",
        }))?;
    let (hyperlink_len, hyperlink_aggregate) = hyperlink_record_budget(definition, limits)?;
    let output_len = list
        .len()
        .checked_add(hyperlink_len)
        .ok_or_else(|| Error::ResourceLimit("ExObjListContainer length overflows".into()))?;
    let old_aggregate = binary_tag_edit::serialized_payload_budget(list)?;
    let aggregate = old_aggregate
        .checked_add(hyperlink_len)
        .and_then(|value| value.checked_add(hyperlink_aggregate))
        .ok_or_else(|| Error::ResourceLimit("ExObjListContainer aggregate overflows".into()))?;
    ensure_container_budget(output_len, aggregate, limits, "ExObjListContainer")?;
    let mut data = wire::rebuild(atom, &id.to_le_bytes())?;
    for child in &children[1..] {
        data.extend_from_slice(child);
    }
    let hyperlink = hyperlink_record(id, definition, limits)?;
    debug_assert_eq!(hyperlink.len(), hyperlink_len);
    data.extend_from_slice(&hyperlink);
    Ok((wire::rebuild(list, &data)?, id.cast_unsigned()))
}

/// Add an `ExHyperlink9Container` with a screen tip to the `___PPT9` tag of
/// the document's information list.
fn add_screen_tip(info: &[u8], id: u32, tip: &str, limits: &SourceLimits) -> Result<Vec<u8>> {
    ensure_serialized_record(info, limits, "DocumentInfoList hyperlink edit")?;
    let blob = binary_tag_edit::locate(info, RecordType::DocInfoList as u16, "___PPT9", ELEMENT)?;
    let (extension_len, extension_aggregate) = hyperlink_extension_budget(tip, limits)?;
    let projected = blob.projected_rebuild_len(info, extension_len)?;
    let projected_aggregate =
        blob.projected_rebuild_aggregate(info, extension_len, extension_aggregate)?;
    ensure_container_budget(
        projected,
        projected_aggregate,
        limits,
        "DocumentInfoList hyperlink edit",
    )?;
    let extension = bounded_container(
        RecordType::ExternalHyperlink9 as u16,
        0,
        &[
            bounded_encode(
                RecordType::ExternalHyperlinkAtom as u16,
                0,
                0,
                &id.to_le_bytes(),
                limits,
                "ExHyperlinkAtom",
            )?,
            bounded_encode(
                CSTRING,
                0,
                0,
                &bounded_utf16(tip, limits, "hyperlink screen tip")?,
                limits,
                "hyperlink screen tip",
            )?,
            bounded_encode(
                RecordType::ExternalHyperlinkFlagsAtom as u16,
                0,
                0,
                &0_u32.to_le_bytes(),
                limits,
                "ExHyperlinkFlagsAtom",
            )?,
        ],
        limits,
        "ExHyperlink9Container",
    )?;
    let mut records = blob
        .records()
        .iter()
        .map(|record| record.to_vec())
        .collect::<Vec<_>>();
    let mut position = 0;
    for (index, record) in blob.records().iter().enumerate() {
        let kind = wire::record_type(record)?;
        if kind == RecordType::ExternalHyperlink9 as u16 || BEFORE_HYPERLINK9.contains(&kind) {
            position = index + 1;
        }
    }
    records.insert(position, extension);
    let rebuilt = blob.rebuild(info, &records, &[])?;
    ensure_serialized_record(&rebuilt, limits, "DocumentInfoList hyperlink edit")?;
    Ok(rebuilt)
}

fn bounded_utf16(value: &str, limits: &SourceLimits, context: &str) -> Result<Vec<u8>> {
    utf16_payload_len(value, limits, context)?;
    Ok(value.encode_utf16().flat_map(u16::to_le_bytes).collect())
}

fn utf16_payload_len(value: &str, limits: &SourceLimits, context: &str) -> Result<usize> {
    let bytes = value
        .encode_utf16()
        .count()
        .checked_mul(2)
        .ok_or_else(|| Error::ResourceLimit(format!("{context} length overflows")))?;
    ensure_payload_len(bytes, limits, context)?;
    Ok(bytes)
}

fn bounded_encode(
    record_type: u16,
    version: u16,
    instance: u16,
    payload: &[u8],
    limits: &SourceLimits,
    context: &str,
) -> Result<Vec<u8>> {
    let length = payload
        .len()
        .checked_add(8)
        .ok_or_else(|| Error::ResourceLimit(format!("{context} length overflows")))?;
    ensure_container_budget(length, payload.len(), limits, context)?;
    Ok(wire::encode(record_type, version, instance, payload)?)
}

fn bounded_container(
    record_type: u16,
    instance: u16,
    children: &[Vec<u8>],
    limits: &SourceLimits,
    context: &str,
) -> Result<Vec<u8>> {
    let payload_len = children.iter().try_fold(0usize, |total, child| {
        total
            .checked_add(child.len())
            .ok_or_else(|| Error::ResourceLimit(format!("{context} length overflows")))
    })?;
    let nested_payload_len = children.iter().try_fold(0usize, |total, child| {
        total
            .checked_add(binary_tag_edit::serialized_payload_budget(child)?)
            .ok_or_else(|| {
                EditError::from(Error::ResourceLimit(format!(
                    "{context} aggregate overflows"
                )))
            })
    })?;
    let length = payload_len
        .checked_add(8)
        .ok_or_else(|| Error::ResourceLimit(format!("{context} length overflows")))?;
    let aggregate = payload_len
        .checked_add(nested_payload_len)
        .ok_or_else(|| Error::ResourceLimit(format!("{context} aggregate overflows")))?;
    ensure_container_budget(length, aggregate, limits, context)?;
    Ok(wire::container(record_type, instance, children)?)
}

fn ensure_payload_len(payload: usize, limits: &SourceLimits, context: &str) -> Result<()> {
    if payload > limits.max_record_payload_bytes() {
        return Err(Error::Limit {
            kind: crate::SourceLimitKind::RecordPayloadBytes,
            observed: payload,
            maximum: limits.max_record_payload_bytes(),
        }
        .into());
    }
    if payload > limits.max_aggregate_payload_bytes() {
        return Err(Error::Limit {
            kind: crate::SourceLimitKind::AggregatePayloadBytes,
            observed: payload,
            maximum: limits.max_aggregate_payload_bytes(),
        }
        .into());
    }
    let _ = context;
    Ok(())
}

fn ensure_container_budget(
    length: usize,
    aggregate: usize,
    limits: &SourceLimits,
    context: &str,
) -> Result<()> {
    let payload = length.checked_sub(8).ok_or_else(|| {
        Error::ResourceLimit(format!("{context} record is shorter than its header"))
    })?;
    ensure_payload_len(payload, limits, context)?;
    if length > limits.max_record_bytes() {
        return Err(Error::Limit {
            kind: crate::SourceLimitKind::RecordBytes,
            observed: length,
            maximum: limits.max_record_bytes(),
        }
        .into());
    }
    if aggregate > limits.max_aggregate_payload_bytes() {
        return Err(Error::Limit {
            kind: crate::SourceLimitKind::AggregatePayloadBytes,
            observed: aggregate,
            maximum: limits.max_aggregate_payload_bytes(),
        }
        .into());
    }
    Ok(())
}

fn ensure_serialized_record(record: &[u8], limits: &SourceLimits, context: &str) -> Result<()> {
    let aggregate = binary_tag_edit::serialized_payload_budget(record)?;
    ensure_container_budget(record.len(), aggregate, limits, context)
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
mod tests {
    use super::*;

    fn limits(aggregate: usize) -> SourceLimits {
        SourceLimits::builder()
            .max_aggregate_payload_bytes(aggregate)
            .unwrap()
            .build()
            .unwrap()
    }

    #[test]
    fn hyperlink_writer_admits_nested_payload_at_exact_boundary() {
        let definition = Definition {
            friendly_name: "",
            target: None,
            location: None,
            tooltip: None,
        };
        let exact = limits(24);
        assert_eq!(
            hyperlink_record_budget(&definition, &exact).unwrap(),
            (28, 24)
        );
        let bytes = hyperlink_record(1, &definition, &exact).unwrap();
        let (_, consumed) = crate::records::Record::parse(&bytes, 0, &exact).unwrap();
        assert_eq!(consumed, bytes.len());
        assert!(hyperlink_record(1, &definition, &limits(23)).is_err());
    }

    #[test]
    fn screen_tip_writer_charges_all_nested_atoms() {
        let exact = limits(40);
        assert_eq!(hyperlink_extension_budget("", &exact).unwrap(), (40, 40));
        let children = vec![
            bounded_encode(
                RecordType::ExternalHyperlinkAtom as u16,
                0,
                0,
                &1u32.to_le_bytes(),
                &exact,
                "ExHyperlinkAtom",
            )
            .unwrap(),
            bounded_encode(CSTRING, 0, 0, &[], &exact, "screen tip").unwrap(),
            bounded_encode(
                RecordType::ExternalHyperlinkFlagsAtom as u16,
                0,
                0,
                &0u32.to_le_bytes(),
                &exact,
                "ExHyperlinkFlagsAtom",
            )
            .unwrap(),
        ];
        let bytes = bounded_container(
            RecordType::ExternalHyperlink9 as u16,
            0,
            &children,
            &exact,
            "ExHyperlink9Container",
        )
        .unwrap();
        let (_, consumed) = crate::records::Record::parse(&bytes, 0, &exact).unwrap();
        assert_eq!(consumed, bytes.len());
        assert!(
            bounded_container(
                RecordType::ExternalHyperlink9 as u16,
                0,
                &children,
                &limits(39),
                "ExHyperlink9Container",
            )
            .is_err()
        );
    }

    #[test]
    fn appending_hyperlink_admits_full_list_aggregate_at_boundary() {
        let atom =
            wire::encode(RecordType::ExObjListAtom as u16, 0, 0, &1i32.to_le_bytes()).unwrap();
        let list = wire::container(RecordType::ExObjList as u16, 0, &[atom]).unwrap();
        let definition = Definition {
            friendly_name: "",
            target: None,
            location: None,
            tooltip: None,
        };
        // The original list charges 12 bytes of outer payload plus its
        // four-byte atom (16 total).  The appended hyperlink contributes its
        // complete 28-byte framing and its recursive 24-byte charge.
        let exact = limits(68);
        let (rebuilt, id) = append_to_list(&list, &definition, &exact).unwrap();
        assert_eq!(id, 2);
        let (_, consumed) = crate::records::Record::parse(&rebuilt, 0, &exact).unwrap();
        assert_eq!(consumed, rebuilt.len());
        assert!(append_to_list(&list, &definition, &limits(67)).is_err());
    }

    #[test]
    fn shape_click_admission_uses_exact_recursive_slide_aggregate() {
        let shape_atom = wire::encode(0xF00A, 2, 0, &7u32.to_le_bytes()).unwrap();
        let opaque_shape_child = wire::encode(0x1234, 0, 0, &[0; 72]).unwrap();
        let shape = wire::container(SP_CONTAINER, 0, &[shape_atom, opaque_shape_child]).unwrap();
        let drawing = wire::container(
            RecordType::PPDrawing as u16,
            0,
            std::slice::from_ref(&shape),
        )
        .unwrap();
        let slide =
            wire::container(RecordType::Slide as u16, 0, std::slice::from_ref(&drawing)).unwrap();
        let source_aggregate = binary_tag_edit::serialized_payload_budget(&slide).unwrap();
        assert_eq!(source_aggregate, 208);

        let click = click_container(&[0; 16], &SourceLimits::default()).unwrap();
        let permissive = set_shape_click(&slide, 7, Some(&click), &SourceLimits::default())
            .unwrap()
            .unwrap();
        let output_aggregate = binary_tag_edit::serialized_payload_budget(&permissive).unwrap();
        assert_eq!(output_aggregate, 288);

        let exact = limits(output_aggregate);
        assert!(set_shape_click(&slide, 7, Some(&click), &exact).is_ok());
        assert!(set_shape_click(&slide, 7, Some(&click), &limits(output_aggregate - 1)).is_err());

        let source_record_bytes = permissive.len();
        let source_limits = SourceLimits::builder()
            .max_record_bytes(source_record_bytes)
            .unwrap()
            .max_record_payload_bytes(source_record_bytes - 8)
            .unwrap()
            .max_aggregate_payload_bytes(output_aggregate)
            .unwrap()
            .build()
            .unwrap();
        let removed = set_shape_click(&permissive, 7, None, &source_limits)
            .unwrap()
            .unwrap();
        let (_, consumed) = crate::records::Record::parse(&removed, 0, &source_limits).unwrap();
        assert_eq!(consumed, removed.len());
    }
}
