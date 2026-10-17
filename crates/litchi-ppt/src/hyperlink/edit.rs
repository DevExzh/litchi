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
use crate::package::Error;
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
pub(crate) fn click_container(atom: &[u8; 16]) -> Result<Vec<u8>> {
    Ok(wire::container(
        INTERACTIVE_INFO,
        0,
        &[wire::encode(
            RecordType::InteractiveInfoAtom as u16,
            0,
            0,
            atom,
        )?],
    )?)
}

/// Replace (`Some`) or remove (`None`) the click action of the shape with
/// `OfficeArt` ID `shape_id`; `None` when nothing changes.
pub(crate) fn set_shape_click(
    slide: &[u8],
    shape_id: u32,
    click: Option<&[u8]>,
) -> Result<Option<Vec<u8>>> {
    let mut matches = 0usize;
    let mut changed = false;
    let mut drawings = 0usize;
    let rewritten = rewrite_drawings(slide, &mut |drawing| {
        drawings += 1;
        let data = rewrite_officeart(
            wire::payload(drawing)?,
            shape_id,
            click,
            0,
            &mut matches,
            &mut changed,
        )?;
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
        1 => Ok(changed.then_some(rewritten)),
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
    click: Option<&[u8]>,
    depth: usize,
    matches: &mut usize,
    changed: &mut bool,
) -> Result<Vec<u8>> {
    if depth >= 128 {
        return Err(
            Error::Corrupted("OfficeArt nesting exceeds the hyperlink-edit limit".into()).into(),
        );
    }
    if wire::record_type(record)? == SP_CONTAINER && wire::shape_id_of(record)? == Some(shape_id) {
        *matches += 1;
        return rewrite_shape(record, click, changed);
    }
    if wire::record_version(record)? != 0x0F {
        return Ok(record.to_vec());
    }
    let mut data = Vec::with_capacity(wire::payload(record)?.len());
    for child in wire::children(record)? {
        data.extend_from_slice(&rewrite_officeart(
            child?,
            shape_id,
            click,
            depth + 1,
            matches,
            changed,
        )?);
    }
    Ok(wire::rebuild(record, &data)?)
}

fn rewrite_shape(shape: &[u8], click: Option<&[u8]>, changed: &mut bool) -> Result<Vec<u8>> {
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
        Some(index) => rewrite_client_data(children[index], click)?,
        None => match click {
            None => return Ok(shape.to_vec()),
            Some(click) => Some(wire::container(CLIENT_DATA, 0, &[click.to_vec()])?),
        },
    };
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
fn rewrite_client_data(record: &[u8], click: Option<&[u8]>) -> Result<Option<Vec<u8>>> {
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
) -> Result<(Vec<u8>, u32)> {
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
        Some(index) => append_to_list(children[index], definition)?,
        None => {
            let id: i32 = 1;
            let record = wire::container(
                RecordType::ExObjList as u16,
                0,
                &[
                    wire::encode(RecordType::ExObjListAtom as u16, 0, 0, &id.to_le_bytes())?,
                    hyperlink_record(id, definition)?,
                ],
            )?;
            (record, id.cast_unsigned())
        },
    };
    let info_record = match (definition.tooltip, info) {
        (None, _) => None,
        (Some(tip), Some(index)) => Some(add_screen_tip(children[index], id, tip)?),
        (Some(_), None) => {
            return Err(unsupported(
                "the document has no information list for a screen tip",
            ));
        },
    };
    let mut data = Vec::with_capacity(document.len().saturating_add(256));
    let mut placed = list.is_some();
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
    Ok((wire::rebuild(document, &data)?, id))
}

fn unsupported(reason: &'static str) -> EditError {
    EditError::Refused(Refusal::UnsupportedElement {
        element: ELEMENT,
        reason,
    })
}

fn hyperlink_record(id: i32, definition: &Definition<'_>) -> Result<Vec<u8>> {
    let mut children = vec![wire::encode(
        RecordType::ExternalHyperlinkAtom as u16,
        0,
        0,
        &id.to_le_bytes(),
    )?];
    children.push(wire::encode(
        CSTRING,
        0,
        0,
        &wire::utf16(definition.friendly_name),
    )?);
    if let Some(target) = definition.target {
        children.push(wire::encode(CSTRING, 0, 1, &wire::utf16(target))?);
    }
    if let Some(location) = definition.location {
        children.push(wire::encode(CSTRING, 0, 3, &wire::utf16(location))?);
    }
    Ok(wire::container(
        RecordType::ExternalHyperlink as u16,
        0,
        &children,
    )?)
}

fn append_to_list(list: &[u8], definition: &Definition<'_>) -> Result<(Vec<u8>, u32)> {
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
    let mut data = wire::rebuild(atom, &id.to_le_bytes())?;
    for child in &children[1..] {
        data.extend_from_slice(child);
    }
    data.extend_from_slice(&hyperlink_record(id, definition)?);
    Ok((wire::rebuild(list, &data)?, id.cast_unsigned()))
}

/// Add an `ExHyperlink9Container` with a screen tip to the `___PPT9` tag of
/// the document's information list.
fn add_screen_tip(info: &[u8], id: u32, tip: &str) -> Result<Vec<u8>> {
    let blob = binary_tag_edit::locate(info, RecordType::DocInfoList as u16, "___PPT9", ELEMENT)?;
    let extension = wire::container(
        RecordType::ExternalHyperlink9 as u16,
        0,
        &[
            wire::encode(
                RecordType::ExternalHyperlinkAtom as u16,
                0,
                0,
                &id.to_le_bytes(),
            )?,
            wire::encode(CSTRING, 0, 0, &wire::utf16(tip))?,
            wire::encode(
                RecordType::ExternalHyperlinkFlagsAtom as u16,
                0,
                0,
                &0_u32.to_le_bytes(),
            )?,
        ],
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
    blob.rebuild(info, &records, &[])
}
