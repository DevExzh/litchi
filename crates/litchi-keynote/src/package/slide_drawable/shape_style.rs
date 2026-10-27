//! Drawing and frame styles of text boxes, shapes and lines (record 0372;
//! the Keynote leaf of the host's `*_slide_shape_{fill,stroke,shadow,
//! effects,text_layout}`, `*_slide_text_box_{columns,text_layout}`,
//! `*_slide_line_endpoints` and `set_slide_shape_image_fill` families) over
//! the shared shape-style variation engine (`litchi_iwa_style::shape`,
//! record 0375a).
//!
//! A drawable's own value lives in a copy-on-write variation of its
//! `TSWP.ShapeStyleArchive`. The engine plans each write and this module
//! publishes the plan on the staged workspace, following the engine's
//! documented steps:
//!
//! - **Rewrite**: the drawable's exclusive variation takes the new payload;
//! - **Derive**: a new child variation is stored beside its stylesheet,
//!   registered there, given a UUID, and the drawable points at it;
//! - **Collapse**: the variation overrides nothing any more, so the drawable
//!   points back at its parent and the variation goes away.
//!
//! Line endpoints and image fills use the plan layer with this module's own
//! slot values: endpoints are read from the line's path (and its legacy
//! per-shape overrides) and stored as a head/tail pair in the line's
//! direction; an image fill also embeds its bytes and moves the style's data
//! references.

use std::sync::Arc;

use litchi_iwa_common::color::Rgba;
use litchi_iwa_common::shape::effects::Effects;
use litchi_iwa_common::shape::fill::{Fill, Image, ImageData, ImageTechnique};
use litchi_iwa_common::shape::geometry::Size;
use litchi_iwa_common::shape::line::{Endpoint, Endpoints};
use litchi_iwa_common::shape::shadow::Shadow;
use litchi_iwa_common::shape::stroke::Stroke;
use litchi_iwa_common::text::layout::Layout;
use litchi_iwa_core::{ArchiveObject, RawMessage};
use litchi_iwa_style::shape::{
    self as engine, Plan, ShapeProperty, Slot, Styles, Target, Value, property,
};
use litchi_iwa_text::columns::Columns;
use sha1::{Digest as _, Sha1};

use super::super::structure::metadata::ExternalReference;
use super::super::structure::wire::{Fields, Writer};
use super::super::structure::workspace::Workspace;
use super::{Error, Refusal, Unsupported, lift};

const SHAPE_INFO_MESSAGE_TYPE: u32 = 2_011;
const STYLESHEET_MESSAGE_TYPE: u32 = 401;
const SHAPE_ARCHIVE: u32 = 1;
const SHAPE_PATH_SOURCE: u32 = 3;
const SHAPE_HEAD_LINE_END: u32 = 4;
const SHAPE_TAIL_LINE_END: u32 = 5;
const PATH_SOURCE_HORIZONTAL_FLIP: u32 = 1;
const LINE_END_PATH: u32 = 1;
const LINE_END_JOIN: u32 = 2;
const LINE_END_POINT: u32 = 3;
const LINE_END_FILLED: u32 = 4;
const LINE_END_IDENTIFIER: u32 = 5;
const MITER_JOIN: u64 = 0;
const ROUND_JOIN: u64 = 1;

/// One drawing or frame property of a text box, shape or line.
#[derive(Debug, Clone, PartialEq)]
#[non_exhaustive]
pub enum Property {
    /// A standard fill (none, a color or a gradient). Image fills are
    /// written by [`super::Edit::set_image_fill`] and read as
    /// [`Self::ImageFill`].
    Fill(Fill),
    /// An image fill (read only; the image bytes stay in the package).
    ImageFill(ImageFill),
    /// The stroke (`None`: the hidden stroke).
    Stroke(Option<Stroke>),
    /// The shadow.
    Shadow(Shadow),
    /// Opacity and reflection.
    Effects(Effects),
    /// Vertical alignment, insets and autosizing of the text frame.
    TextLayout(Layout),
    /// The text columns.
    Columns(Columns),
    /// The decorations at a straight line's directed start and end.
    LineEndpoints(Endpoints),
}

/// Which [`Property`] a read or reset addresses.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum PropertyKind {
    /// The fill (standard or image).
    Fill,
    /// The stroke.
    Stroke,
    /// The shadow.
    Shadow,
    /// Opacity and reflection.
    Effects,
    /// The text-frame layout.
    TextLayout,
    /// The text columns.
    Columns,
    /// A straight line's endpoint decorations.
    LineEndpoints,
}

impl Property {
    /// The kind of this property.
    #[must_use]
    pub const fn kind(&self) -> PropertyKind {
        match self {
            Self::Fill(_) | Self::ImageFill(_) => PropertyKind::Fill,
            Self::Stroke(_) => PropertyKind::Stroke,
            Self::Shadow(_) => PropertyKind::Shadow,
            Self::Effects(_) => PropertyKind::Effects,
            Self::TextLayout(_) => PropertyKind::TextLayout,
            Self::Columns(_) => PropertyKind::Columns,
            Self::LineEndpoints(_) => PropertyKind::LineEndpoints,
        }
    }
}

/// An image fill as read back: its layout and colors. The embedded image is
/// a package member and is not exposed by identity.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct ImageFill {
    technique: ImageTechnique,
    fill_size: Size,
    tint: Option<Rgba>,
    reference_color: Option<Rgba>,
    embedded: bool,
}

impl ImageFill {
    fn new(image: &Image) -> Self {
        Self {
            technique: image.technique(),
            fill_size: image.fill_size(),
            tint: image.tint(),
            reference_color: image.reference_color(),
            embedded: image.data().is_some(),
        }
    }

    /// How the image is laid out in the drawable.
    #[must_use]
    pub const fn technique(self) -> ImageTechnique {
        self.technique
    }

    /// The natural or tile size of the image.
    #[must_use]
    pub const fn fill_size(self) -> Size {
        self.fill_size
    }

    /// The tint (Keynote's Advanced Image Fill), if any.
    #[must_use]
    pub const fn tint(self) -> Option<Rgba> {
        self.tint
    }

    /// The representative color Keynote sampled from the image, if any.
    #[must_use]
    pub const fn reference_color(self) -> Option<Rgba> {
        self.reference_color
    }

    /// Whether an image is embedded (`false`: the empty placeholder).
    #[must_use]
    pub const fn is_embedded(self) -> bool {
        self.embedded
    }
}

/// One staged shape-style change.
#[derive(Debug, Clone)]
pub(super) enum Change {
    Set(Property),
    Reset(PropertyKind),
    ImageFill {
        bytes: Arc<[u8]>,
        file_name: Box<str>,
        technique: ImageTechnique,
        tint: Option<Rgba>,
    },
}

fn refused() -> Error {
    Error::Refused(Refusal::Unsupported {
        reason: Unsupported::ShapeStyle,
    })
}

fn from_engine(error: litchi_iwa_style::graph::Error) -> Error {
    match error {
        litchi_iwa_style::graph::Error::Invalid(_) => refused(),
        _ => Error::InvalidSource,
    }
}

/// The staged package's shape styles, as the engine reads them.
struct WorkspaceStyles<'w, 'p>(&'w Workspace<'p>);

impl Styles for WorkspaceStyles<'_, '_> {
    fn shape_style(&self, identifier: u64) -> Option<&[u8]> {
        self.0
            .object(identifier)?
            .messages
            .iter()
            .find(|message| message.type_ == engine::MESSAGE_TYPE)
            .map(|message| message.data.as_slice())
    }
}

fn style_payload<'w>(workspace: &'w Workspace<'_>, style: u64) -> Result<&'w [u8], Error> {
    workspace
        .object(style)
        .and_then(|object| {
            object
                .messages
                .iter()
                .find(|message| message.type_ == engine::MESSAGE_TYPE)
        })
        .map(|message| message.data.as_slice())
        .ok_or(Error::InvalidSource)
}

/// The drawable's `TSWP.ShapeInfoArchive`: (message index, payload).
fn shape_info(workspace: &Workspace<'_>, drawable: u64) -> Result<(usize, Vec<u8>), Error> {
    let object = workspace.require(drawable).map_err(lift)?;
    let mut found = object
        .messages
        .iter()
        .enumerate()
        .filter(|(_, message)| message.type_ == SHAPE_INFO_MESSAGE_TYPE);
    let (index, message) = found.next().ok_or(refused())?;
    if found.next().is_some() {
        return Err(Error::InvalidSource);
    }
    Ok((index, message.data.clone()))
}

fn style_of(workspace: &Workspace<'_>, drawable: u64) -> Result<u64, Error> {
    let (_, payload) = shape_info(workspace, drawable)?;
    engine::style_of(&payload).map_err(from_engine)
}

/// Whether exactly one shape uses `style` and no shape style inherits from
/// it (the host's exclusivity rule).
fn exclusive(workspace: &Workspace<'_>, style: u64) -> Result<bool, Error> {
    let mut shapes = 0usize;
    for component in workspace.package().state.source.components().iter() {
        for identifier in workspace.member_objects(component.name()) {
            let Some(object) = workspace.object(identifier) else {
                continue;
            };
            for message in &object.messages {
                match message.type_ {
                    SHAPE_INFO_MESSAGE_TYPE => {
                        if engine::style_of(&message.data).ok() == Some(style) {
                            shapes += 1;
                        }
                    },
                    engine::MESSAGE_TYPE => {
                        if engine::parent_of(&message.data).map_err(from_engine)? == Some(style) {
                            return Ok(false);
                        }
                    },
                    _ => {},
                }
            }
        }
    }
    Ok(shapes == 1)
}

fn target(workspace: &Workspace<'_>, drawable: u64) -> Result<Target, Error> {
    let style = style_of(workspace, drawable)?;
    Ok(Target::new(style, exclusive(workspace, style)?))
}

fn component_id(workspace: &Workspace<'_>, object: u64) -> Result<u64, Error> {
    let member = workspace.component_of(object).ok_or(Error::InvalidSource)?;
    Ok(workspace
        .registry
        .component_for_member(&member)
        .ok_or(Error::InvalidSource)?
        .identifier)
}

/// Repoint the drawable's style reference (payload and header).
fn repoint(workspace: &mut Workspace<'_>, drawable: u64, from: u64, to: u64) -> Result<(), Error> {
    let limits = workspace.archive_limits();
    let (index, payload) = shape_info(workspace, drawable)?;
    let data = engine::repoint(&payload, from, to).map_err(from_engine)?;
    let object = workspace.object_mut(drawable).map_err(lift)?;
    object
        .replace_message(
            index,
            RawMessage {
                type_: SHAPE_INFO_MESSAGE_TYPE,
                data,
            },
            limits,
        )
        .map_err(|_| Error::InvalidSource)?;
    let info = object
        .archive_info
        .message_infos
        .get_mut(index)
        .ok_or(Error::InvalidSource)?;
    let mut replaced = 0;
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
        return Err(refused());
    }
    Ok(())
}

fn replace_stylesheet(
    workspace: &mut Workspace<'_>,
    stylesheet: u64,
    data: Vec<u8>,
    removed: Option<u64>,
    appended: Option<u64>,
) -> Result<(), Error> {
    let limits = workspace.archive_limits();
    let object = workspace.object_mut(stylesheet).map_err(lift)?;
    let index = object
        .messages
        .iter()
        .position(|message| message.type_ == STYLESHEET_MESSAGE_TYPE)
        .ok_or(Error::InvalidSource)?;
    object
        .replace_message(
            index,
            RawMessage {
                type_: STYLESHEET_MESSAGE_TYPE,
                data,
            },
            limits,
        )
        .map_err(|_| Error::InvalidSource)?;
    let info = object
        .archive_info
        .message_infos
        .get_mut(index)
        .ok_or(Error::InvalidSource)?;
    if let Some(removed) = removed {
        if info
            .object_references
            .iter()
            .filter(|reference| **reference == removed)
            .count()
            != 1
        {
            return Err(Error::InvalidSource);
        }
        info.object_references
            .retain(|reference| *reference != removed);
        for field in &mut info.field_infos {
            field
                .object_references
                .retain(|reference| *reference != removed);
        }
    }
    if let Some(appended) = appended {
        if info.object_references.contains(&appended) {
            return Err(Error::InvalidSource);
        }
        info.object_references.push(appended);
    }
    Ok(())
}

/// The style's data reference (an image fill's `TSP.DataReference`): the
/// header's data references and the stylesheet component's registry owner.
fn move_data_reference(
    workspace: &mut Workspace<'_>,
    style: u64,
    old: Option<u64>,
    new: Option<u64>,
) -> Result<(), Error> {
    if old == new {
        return Ok(());
    }
    let component = component_id(workspace, style)?;
    let object = workspace.object_mut(style).map_err(lift)?;
    let info = object
        .archive_info
        .message_infos
        .first_mut()
        .ok_or(Error::InvalidSource)?;
    if let Some(old) = old {
        info.data_references.retain(|data| *data != old);
    }
    if let Some(new) = new {
        info.data_references.push(new);
    }
    let entry = workspace
        .registry
        .component_mut(component)
        .ok_or(Error::InvalidSource)?;
    if old.is_some() {
        entry.remove_data_owner(style);
    }
    if let Some(new) = new {
        entry.add_data_owner(new, style);
    }
    Ok(())
}

/// Publish `plan` (made for `drawable`'s current style). `data` is the
/// image data the new payload references (an image fill), if any.
fn apply(
    workspace: &mut Workspace<'_>,
    drawable: u64,
    plan: Plan,
    data: Option<u64>,
) -> Result<(), Error> {
    let archive_limits = workspace.archive_limits();
    let style = style_of(workspace, drawable)?;
    let old_data = image_data(style_payload(workspace, style)?)?;
    match plan {
        Plan::Unchanged => Ok(()),
        Plan::Rewrite(variation) => {
            let object = workspace.object_mut(style).map_err(lift)?;
            let index = object
                .messages
                .iter()
                .position(|message| message.type_ == engine::MESSAGE_TYPE)
                .ok_or(Error::InvalidSource)?;
            object
                .replace_message(
                    index,
                    RawMessage {
                        type_: engine::MESSAGE_TYPE,
                        data: variation.into_payload(),
                    },
                    archive_limits,
                )
                .map_err(|_| Error::InvalidSource)?;
            move_data_reference(workspace, style, old_data, data)
        },
        Plan::Derive(variation) => {
            let stylesheet = variation.stylesheet();
            let style_member = workspace.component_of(style).ok_or(Error::InvalidSource)?;
            if workspace.component_of(stylesheet).as_deref() != Some(style_member.as_str()) {
                return Err(refused());
            }
            let created = workspace.allocate().map_err(lift)?;
            repoint(workspace, drawable, style, created)?;
            let mut object = ArchiveObject::new(
                created,
                vec![RawMessage {
                    type_: engine::MESSAGE_TYPE,
                    data: variation.payload().to_vec(),
                }],
                archive_limits,
            )
            .map_err(|_| Error::InvalidSource)?;
            {
                let info = object
                    .archive_info
                    .message_infos
                    .first_mut()
                    .ok_or(Error::InvalidSource)?;
                info.versions = engine::MESSAGE_VERSION.to_vec().into();
                info.object_references = variation.references().to_vec();
                if let Some(data) = data {
                    info.data_references.push(data);
                }
            }
            let sheet = workspace
                .payload(stylesheet, STYLESHEET_MESSAGE_TYPE)
                .map_err(lift)?
                .to_vec();
            let registered =
                engine::stylesheet_with_variation(&sheet, style, created).map_err(from_engine)?;
            replace_stylesheet(workspace, stylesheet, registered, None, Some(created))?;
            workspace.insert(&style_member, object).map_err(lift)?;
            let style_component = component_id(workspace, stylesheet)?;
            let owner = component_id(workspace, drawable)?;
            let uuid = workspace.clone_uuid(None, created).map_err(lift)?;
            let entry = workspace
                .registry
                .component_mut(style_component)
                .ok_or(Error::InvalidSource)?;
            entry.add_object_uuid(created, uuid);
            if let Some(data) = data {
                entry.add_data_owner(data, created);
            }
            if owner != style_component {
                workspace
                    .registry
                    .component_mut(owner)
                    .ok_or(Error::InvalidSource)?
                    .add_external_reference(ExternalReference {
                        component: style_component,
                        object: Some(created),
                        is_weak: None,
                    });
            }
            Ok(())
        },
        Plan::Collapse { parent, stylesheet } => {
            move_data_reference(workspace, style, old_data, None)?;
            repoint(workspace, drawable, style, parent)?;
            let sheet = workspace
                .payload(stylesheet, STYLESHEET_MESSAGE_TYPE)
                .map_err(lift)?
                .to_vec();
            let unregistered =
                engine::stylesheet_without_variation(&sheet, parent, style).map_err(from_engine)?;
            let style_component = component_id(workspace, stylesheet)?;
            let owner = component_id(workspace, drawable)?;
            replace_stylesheet(workspace, stylesheet, unregistered, Some(style), None)?;
            workspace.remove_object(style).map_err(lift)?;
            workspace
                .registry
                .component_mut(style_component)
                .ok_or(Error::InvalidSource)?
                .remove_object_uuids(&[style]);
            let referrers = workspace
                .registry
                .components()
                .filter(|component| {
                    component.external_references().iter().any(|reference| {
                        reference.component == style_component && reference.object == Some(style)
                    })
                })
                .map(|component| component.identifier)
                .collect::<Vec<_>>();
            for referrer in referrers {
                let entry = workspace
                    .registry
                    .component_mut(referrer)
                    .ok_or(Error::InvalidSource)?;
                let kept = entry
                    .external_references()
                    .iter()
                    .filter(|reference| {
                        !(reference.component == style_component && reference.object == Some(style))
                    })
                    .copied()
                    .collect();
                entry.set_external_references(kept);
            }
            if owner != style_component {
                workspace
                    .registry
                    .component_mut(owner)
                    .ok_or(Error::InvalidSource)?
                    .add_external_reference(ExternalReference {
                        component: style_component,
                        object: Some(parent),
                        is_weak: None,
                    });
            }
            workspace.release_removed_identifier_suffix();
            Ok(())
        },
        _ => Err(refused()),
    }
}

/// The image data a style payload's own fill references, if any.
fn image_data(style: &[u8]) -> Result<Option<u64>, Error> {
    let Some(stored) = stored_direct(style, Slot::Fill)? else {
        return Ok(None);
    };
    Ok(match property::Fills.read(&[Some(&stored)]) {
        Ok(Fill::Image(image)) => image.data().map(ImageData::get),
        _ => None,
    })
}

/// The slot value a style payload stores itself (not inherited).
fn stored_direct(style: &[u8], slot: Slot) -> Result<Option<Vec<u8>>, Error> {
    if !engine::defines(style, slot).map_err(from_engine)? {
        return Ok(None);
    }
    struct One<'a>(&'a [u8]);
    impl Styles for One<'_> {
        fn shape_style(&self, _: u64) -> Option<&[u8]> {
            Some(self.0)
        }
    }
    // A one-style chain: the payload's own value (it defines the slot).
    let parentless = One(style);
    engine::effective(&parentless, 0, slot).map_err(from_engine)
}

fn inherited<P: ShapeProperty>(
    workspace: &Workspace<'_>,
    style: u64,
    property: &P,
) -> Result<P::Value, Error> {
    let styles = WorkspaceStyles(workspace);
    let mut stored = Vec::new();
    for group in property.groups() {
        stored.extend(engine::effective_group(&styles, style, group).map_err(from_engine)?);
    }
    let borrowed = stored.iter().map(Option::as_deref).collect::<Vec<_>>();
    property.read(&borrowed).map_err(from_engine)
}

fn value<P: ShapeProperty>(
    workspace: &Workspace<'_>,
    drawable: u64,
    property: &P,
) -> Result<P::Value, Error> {
    inherited(workspace, style_of(workspace, drawable)?, property)
}

fn set<P: ShapeProperty>(
    workspace: &mut Workspace<'_>,
    drawable: u64,
    property: &P,
    value: &P::Value,
) -> Result<(), Error> {
    let target = target(workspace, drawable)?;
    if inherited(workspace, target.style(), property)? == *value {
        return Ok(());
    }
    let writes = property.writes(value).map_err(from_engine)?;
    if target.exclusive() {
        let payload = style_payload(workspace, target.style())?;
        for (slot, _) in &writes {
            if let Some(stored) = stored_direct(payload, *slot)? {
                property.replaceable(*slot, &stored).map_err(from_engine)?;
            }
        }
    }
    let plan =
        engine::plan_set(&WorkspaceStyles(workspace), target, &writes).map_err(from_engine)?;
    apply(workspace, drawable, plan, None)?;
    if self::value(workspace, drawable, property)? != *value {
        return Err(Error::Verification);
    }
    Ok(())
}

fn reset<P: ShapeProperty>(
    workspace: &mut Workspace<'_>,
    drawable: u64,
    property: &P,
) -> Result<(), Error> {
    let target = target(workspace, drawable)?;
    let slots = property
        .groups()
        .iter()
        .flat_map(|group| group.iter().copied())
        .collect::<Vec<_>>();
    let payload = style_payload(workspace, target.style())?;
    if property.lenient_reset()
        && engine::direct_slots(payload)
            .map_err(from_engine)?
            .is_none()
    {
        return Ok(());
    }
    let mut defined = false;
    for slot in &slots {
        if let Some(stored) = stored_direct(payload, *slot)? {
            property.replaceable(*slot, &stored).map_err(from_engine)?;
            defined = true;
        }
    }
    if !defined {
        return Ok(());
    }
    let parent = engine::parent_of(payload)
        .map_err(from_engine)?
        .ok_or(refused())?;
    let restored = inherited(workspace, parent, property)?;
    let fallback = property
        .writes(&restored)
        .map_err(from_engine)?
        .into_iter()
        .filter_map(|(slot, value)| value.map(|value| (slot, value)))
        .collect::<Vec<_>>();
    let plan = engine::plan_reset(&WorkspaceStyles(workspace), target, &slots, &fallback)
        .map_err(from_engine)?;
    apply(workspace, drawable, plan, None)?;
    if value(workspace, drawable, property)? != restored {
        return Err(Error::Verification);
    }
    Ok(())
}

// ---- Line endpoints (the host's `shapes::line_end`, own slot values) ----

/// The line's legacy per-shape decorations and stored direction:
/// (head, tail, horizontally flipped).
fn shape_line_ends(
    workspace: &Workspace<'_>,
    drawable: u64,
) -> Result<(Option<Vec<u8>>, Option<Vec<u8>>, bool), Error> {
    let wire = workspace.wire;
    let (_, payload) = shape_info(workspace, drawable)?;
    let info = Fields::parse(&payload, wire).map_err(lift)?;
    let shape = info
        .unique(SHAPE_ARCHIVE)
        .map_err(lift)?
        .ok_or(Error::InvalidSource)?
        .message()
        .map_err(lift)?;
    let shape = Fields::parse(shape, wire).map_err(lift)?;
    let field = |number| -> Result<Option<Vec<u8>>, Error> {
        Ok(match shape.unique(number).map_err(lift)? {
            Some(field) => Some(field.message().map_err(lift)?.to_vec()),
            None => None,
        })
    };
    let flipped = match shape.unique(SHAPE_PATH_SOURCE).map_err(lift)? {
        Some(source) => {
            let source = Fields::parse(source.message().map_err(lift)?, wire).map_err(lift)?;
            match source.unique(PATH_SOURCE_HORIZONTAL_FLIP).map_err(lift)? {
                Some(flip) => flip.varint().map_err(lift)? != 0,
                None => false,
            }
        },
        None => false,
    };
    Ok((
        field(SHAPE_HEAD_LINE_END)?,
        field(SHAPE_TAIL_LINE_END)?,
        flipped,
    ))
}

fn endpoint_of(workspace: &Workspace<'_>, line_end: Option<&[u8]>) -> Result<Endpoint, Error> {
    let Some(line_end) = line_end else {
        return Ok(Endpoint::None);
    };
    let fields = Fields::parse(line_end, workspace.wire).map_err(lift)?;
    let identifier = match fields.unique(LINE_END_IDENTIFIER).map_err(lift)? {
        Some(field) => std::str::from_utf8(field.payload()).map_err(|_| Error::InvalidSource)?,
        None => "none",
    };
    Ok(match identifier {
        "none" => Endpoint::None,
        "simple arrow" => Endpoint::SimpleArrow,
        "filled circle" => Endpoint::FilledCircle,
        "filled diamond" => Endpoint::FilledDiamond,
        "open arrow" => Endpoint::OpenArrow,
        "filled arrow" => Endpoint::FilledArrow,
        "filled square" => Endpoint::FilledSquare,
        "open square" => Endpoint::OpenSquare,
        "open circle" => Endpoint::OpenCircle,
        "inverted arrow" => Endpoint::InvertedArrow,
        "line" => Endpoint::Line,
        _ => return Err(refused()),
    })
}

/// The effective head and tail of a style chain, each inherited on its own.
fn inherited_line_ends(
    workspace: &Workspace<'_>,
    style: u64,
) -> Result<(Option<Vec<u8>>, Option<Vec<u8>>), Error> {
    let styles = WorkspaceStyles(workspace);
    Ok((
        engine::effective(&styles, style, Slot::HeadLineEnd).map_err(from_engine)?,
        engine::effective(&styles, style, Slot::TailLineEnd).map_err(from_engine)?,
    ))
}

fn endpoints(workspace: &Workspace<'_>, drawable: u64) -> Result<Endpoints, Error> {
    let (mut head, mut tail, flipped) = shape_line_ends(workspace, drawable)?;
    if head.is_none() || tail.is_none() {
        let (inherited_head, inherited_tail) =
            inherited_line_ends(workspace, style_of(workspace, drawable)?)?;
        head = head.or(inherited_head);
        tail = tail.or(inherited_tail);
    }
    let mut endpoints = Endpoints {
        start: endpoint_of(workspace, tail.as_deref())?,
        end: endpoint_of(workspace, head.as_deref())?,
    };
    if flipped {
        std::mem::swap(&mut endpoints.start, &mut endpoints.end);
    }
    Ok(endpoints)
}

const CANONICAL_MIN: f32 = 0.0;
const CANONICAL_MID: f32 = 3.0;
const CANONICAL_MAX: f32 = 6.0;
const CIRCLE_INSET: f32 = 0.5;
const CIRCLE_CONTROL_LOW: f32 = 1.619_288_1;
const CIRCLE_CONTROL_HIGH: f32 = 4.380_712;
const OPEN_ARROW_TIP_Y: f32 = 5.0;
const FILLED_ARROW_NOTCH_Y: f32 = 1.5;
const DIAMOND_ATTACHMENT_Y: f32 = 0.331_4;
const FILLED_SQUARE_INSET: f32 = 0.5;
const OPEN_SQUARE_INSET: f32 = 1.0;
const OPEN_SQUARE_ATTACHMENT_Y: f32 = 0.2;
const OPEN_CIRCLE_ATTACHMENT_Y: f32 = -0.3;
const LINE_ATTACHMENT_Y: f32 = -0.8;
const MOVE_TO: u64 = 1;
const LINE_TO: u64 = 2;
const CURVE_TO: u64 = 4;
const CLOSE_SUBPATH: u64 = 5;

fn point(x: f32, y: f32) -> Writer {
    Writer::new().float(1, x).float(2, y)
}

fn element(kind: u64, points: &[(f32, f32)]) -> Writer {
    points
        .iter()
        .fold(Writer::new().varint(1, kind), |writer, &(x, y)| {
            writer.message(2, &point(x, y))
        })
}

/// `TSD.LineEndArchive`, field by field in tag order (the host's prost
/// encoding of `shapes::line_end::native::endpoint_archive`).
fn line_end(
    identifier: &str,
    join: u64,
    elements: &[Writer],
    end_point: Option<(f32, f32)>,
    filled: Option<bool>,
) -> Vec<u8> {
    let path = elements
        .iter()
        .fold(Writer::new(), |path, element| path.message(1, element));
    let mut writer = Writer::new()
        .message(LINE_END_PATH, &path)
        .varint(LINE_END_JOIN, join);
    if let Some((x, y)) = end_point {
        writer = writer.message(LINE_END_POINT, &point(x, y));
    }
    if let Some(filled) = filled {
        writer = writer.boolean(LINE_END_FILLED, filled);
    }
    writer.string(LINE_END_IDENTIFIER, identifier).finish()
}

fn circle_end(identifier: &str, endpoint_y: f32, filled: Option<bool>) -> Vec<u8> {
    let low = CIRCLE_INSET;
    let high = CANONICAL_MAX - CIRCLE_INSET;
    line_end(
        identifier,
        MITER_JOIN,
        &[
            element(MOVE_TO, &[(high, CANONICAL_MID)]),
            element(
                CURVE_TO,
                &[
                    (high, CIRCLE_CONTROL_HIGH),
                    (CIRCLE_CONTROL_HIGH, high),
                    (CANONICAL_MID, high),
                ],
            ),
            element(
                CURVE_TO,
                &[
                    (CIRCLE_CONTROL_LOW, high),
                    (low, CIRCLE_CONTROL_HIGH),
                    (low, CANONICAL_MID),
                ],
            ),
            element(
                CURVE_TO,
                &[
                    (low, CIRCLE_CONTROL_LOW),
                    (CIRCLE_CONTROL_LOW, low),
                    (CANONICAL_MID, low),
                ],
            ),
            element(
                CURVE_TO,
                &[
                    (CIRCLE_CONTROL_HIGH, low),
                    (high, CIRCLE_CONTROL_LOW),
                    (high, CANONICAL_MID),
                ],
            ),
            element(CLOSE_SUBPATH, &[]),
        ],
        Some((CANONICAL_MID, endpoint_y)),
        filled,
    )
}

fn square_end(identifier: &str, inset: f32, filled: Option<bool>, endpoint_y: f32) -> Vec<u8> {
    let far = CANONICAL_MAX - inset;
    line_end(
        identifier,
        MITER_JOIN,
        &[
            element(MOVE_TO, &[(inset, inset)]),
            element(LINE_TO, &[(far, inset)]),
            element(LINE_TO, &[(far, far)]),
            element(LINE_TO, &[(inset, far)]),
            element(CLOSE_SUBPATH, &[]),
        ],
        Some((CANONICAL_MID, endpoint_y)),
        filled,
    )
}

fn encode_endpoint(endpoint: Endpoint) -> Vec<u8> {
    match endpoint {
        Endpoint::None => line_end("none", MITER_JOIN, &[], None, None),
        Endpoint::SimpleArrow => line_end(
            "simple arrow",
            MITER_JOIN,
            &[
                element(MOVE_TO, &[(CANONICAL_MIN, CANONICAL_MIN)]),
                element(LINE_TO, &[(CANONICAL_MID, CANONICAL_MAX)]),
                element(LINE_TO, &[(CANONICAL_MAX, CANONICAL_MIN)]),
                element(CLOSE_SUBPATH, &[]),
            ],
            Some((CANONICAL_MID, CANONICAL_MIN)),
            Some(true),
        ),
        Endpoint::FilledCircle => circle_end("filled circle", CIRCLE_INSET, Some(true)),
        Endpoint::FilledDiamond => line_end(
            "filled diamond",
            MITER_JOIN,
            &[
                element(MOVE_TO, &[(CANONICAL_MID, CANONICAL_MIN)]),
                element(LINE_TO, &[(CANONICAL_MIN, CANONICAL_MID)]),
                element(LINE_TO, &[(CANONICAL_MID, CANONICAL_MAX)]),
                element(LINE_TO, &[(CANONICAL_MAX, CANONICAL_MID)]),
                element(CLOSE_SUBPATH, &[]),
            ],
            Some((CANONICAL_MID, DIAMOND_ATTACHMENT_Y)),
            Some(true),
        ),
        Endpoint::OpenArrow => line_end(
            "open arrow",
            ROUND_JOIN,
            &[
                element(MOVE_TO, &[(CANONICAL_MIN, CANONICAL_MIN)]),
                element(LINE_TO, &[(CANONICAL_MID, OPEN_ARROW_TIP_Y)]),
                element(LINE_TO, &[(CANONICAL_MAX, CANONICAL_MIN)]),
                element(MOVE_TO, &[(CANONICAL_MID, CANONICAL_MIN)]),
                element(LINE_TO, &[(CANONICAL_MID, OPEN_ARROW_TIP_Y)]),
            ],
            Some((CANONICAL_MID, CANONICAL_MIN)),
            None,
        ),
        Endpoint::FilledArrow => line_end(
            "filled arrow",
            MITER_JOIN,
            &[
                element(MOVE_TO, &[(CANONICAL_MIN, CANONICAL_MIN)]),
                element(LINE_TO, &[(CANONICAL_MID, CANONICAL_MAX)]),
                element(LINE_TO, &[(CANONICAL_MAX, CANONICAL_MIN)]),
                element(LINE_TO, &[(CANONICAL_MID, FILLED_ARROW_NOTCH_Y)]),
                element(CLOSE_SUBPATH, &[]),
            ],
            Some((CANONICAL_MID, FILLED_ARROW_NOTCH_Y)),
            Some(true),
        ),
        Endpoint::FilledSquare => square_end(
            "filled square",
            FILLED_SQUARE_INSET,
            Some(true),
            FILLED_SQUARE_INSET,
        ),
        Endpoint::OpenSquare => square_end(
            "open square",
            OPEN_SQUARE_INSET,
            None,
            OPEN_SQUARE_ATTACHMENT_Y,
        ),
        Endpoint::OpenCircle => circle_end("open circle", OPEN_CIRCLE_ATTACHMENT_Y, None),
        Endpoint::InvertedArrow => line_end(
            "inverted arrow",
            MITER_JOIN,
            &[
                element(MOVE_TO, &[(CANONICAL_MIN, CANONICAL_MID)]),
                element(LINE_TO, &[(CANONICAL_MID, CANONICAL_MIN)]),
                element(LINE_TO, &[(CANONICAL_MAX, CANONICAL_MID)]),
                element(CLOSE_SUBPATH, &[]),
            ],
            Some((CANONICAL_MID, DIAMOND_ATTACHMENT_Y)),
            Some(true),
        ),
        Endpoint::Line => line_end(
            "line",
            MITER_JOIN,
            &[
                element(MOVE_TO, &[(CANONICAL_MIN, CANONICAL_MIN)]),
                element(LINE_TO, &[(CANONICAL_MAX, CANONICAL_MIN)]),
            ],
            Some((CANONICAL_MID, LINE_ATTACHMENT_Y)),
            None,
        ),
    }
}

/// Store `wanted` as the line's decorations (host `set_shape_line_endpoints`).
fn set_endpoints(
    workspace: &mut Workspace<'_>,
    drawable: u64,
    wanted: Endpoints,
) -> Result<(), Error> {
    if endpoints(workspace, drawable)? == wanted {
        return Ok(());
    }
    let (_, _, flipped) = shape_line_ends(workspace, drawable)?;
    let mut stored = wanted;
    if flipped {
        std::mem::swap(&mut stored.start, &mut stored.end);
    }
    let target = target(workspace, drawable)?;
    let payload = style_payload(workspace, target.style())?;
    let disposable = target.exclusive()
        && engine::direct_slots(payload)
            .map_err(from_engine)?
            .is_some();
    let mut clear = false;
    if disposable && wanted == Endpoints::default() {
        let parent = engine::parent_of(payload)
            .map_err(from_engine)?
            .ok_or(refused())?;
        let (head, tail) = inherited_line_ends(workspace, parent)?;
        let inherited = Endpoints {
            start: endpoint_of(workspace, tail.as_deref())?,
            end: endpoint_of(workspace, head.as_deref())?,
        };
        clear = inherited == Endpoints::default();
    }
    let writes = if clear {
        vec![(Slot::HeadLineEnd, None), (Slot::TailLineEnd, None)]
    } else {
        vec![
            (
                Slot::HeadLineEnd,
                Some(Value::message(encode_endpoint(stored.end))),
            ),
            (
                Slot::TailLineEnd,
                Some(Value::message(encode_endpoint(stored.start))),
            ),
        ]
    };
    // A shared or non-minimal style takes a derived child (the host's
    // copy-on-write); the plan layer decides exactly as the host does.
    let plan =
        engine::plan_set(&WorkspaceStyles(workspace), target, &writes).map_err(from_engine)?;
    apply(workspace, drawable, plan, None)?;
    if endpoints(workspace, drawable)? != wanted {
        return Err(Error::Verification);
    }
    Ok(())
}

// ---- Image fills (own data references) ----

fn data_file_name(preferred: &str, identifier: u64) -> Result<String, Error> {
    let path = std::path::Path::new(preferred);
    let stem = path
        .file_stem()
        .and_then(|stem| stem.to_str())
        .filter(|stem| !stem.is_empty())
        .ok_or(Error::Refused(Refusal::InvalidImage))?;
    let name = match path.extension().and_then(|extension| extension.to_str()) {
        Some(extension) if !extension.is_empty() => format!("{stem}-{identifier}.{extension}"),
        _ => format!("{stem}-{identifier}"),
    };
    if name.contains(['/', '\\', '\0']) || name.starts_with('.') {
        return Err(Error::Refused(Refusal::InvalidImage));
    }
    Ok(name)
}

/// Embed `bytes` and make them the drawable's image fill (host
/// `set_slide_shape_image_fill`).
fn set_image_fill(
    workspace: &mut Workspace<'_>,
    drawable: u64,
    fill_size: Size,
    bytes: &[u8],
    file_name: &str,
    technique: ImageTechnique,
    tint: Option<Rgba>,
) -> Result<(), Error> {
    if super::author::sniff(bytes).is_none() {
        return Err(Error::Refused(Refusal::Unsupported {
            reason: Unsupported::ImageFormat,
        }));
    }
    let data = workspace
        .registry
        .data_identifiers(workspace.wire)
        .map_err(lift)?
        .into_iter()
        .max()
        .unwrap_or(0)
        .checked_add(1)
        .ok_or(Error::InvalidSource)?;
    let stored_name = data_file_name(file_name, data)?;
    workspace
        .insert_member(&format!("Data/{stored_name}"), bytes.to_vec())
        .map_err(lift)?;
    workspace.registry.add_data_info(
        data,
        &Sha1::digest(bytes),
        file_name,
        &stored_name,
        bytes.len() as u64,
    );
    let image = Image::embedded(
        ImageData::new(data).map_err(|_| Error::InvalidSource)?,
        technique,
        fill_size,
    )
    .map_err(|_| Error::Refused(Refusal::InvalidGeometry))?;
    let image = match tint {
        Some(tint) => image.with_tint(tint),
        None => image,
    };
    let fill = Fill::Image(image);
    let writes = [(
        Slot::Fill,
        Some(Value::message(litchi_iwa_style::drawing::encode_fill(
            &fill,
        ))),
    )];
    let target = target(workspace, drawable)?;
    let plan =
        engine::plan_set(&WorkspaceStyles(workspace), target, &writes).map_err(from_engine)?;
    apply(workspace, drawable, plan, Some(data))?;
    if value(workspace, drawable, &property::Fills)? != fill {
        return Err(Error::Verification);
    }
    Ok(())
}

/// Publish one staged change of `drawable`'s style.
pub(super) fn stage(
    workspace: &mut Workspace<'_>,
    drawable: u64,
    fill_size: Option<Size>,
    change: &Change,
) -> Result<(), Error> {
    match change {
        Change::Set(Property::Fill(Fill::Image(_)) | Property::ImageFill(_)) => Err(refused()),
        Change::Set(Property::Fill(fill)) => {
            let current = value(workspace, drawable, &property::Fills)?;
            if matches!(current, Fill::Image(_)) {
                return Err(refused());
            }
            set(workspace, drawable, &property::Fills, fill)
        },
        Change::Set(Property::Stroke(stroke)) => {
            set(workspace, drawable, &property::Strokes, stroke)
        },
        Change::Set(Property::Shadow(shadow)) => {
            set(workspace, drawable, &property::Shadows, shadow)
        },
        Change::Set(Property::Effects(effects)) => {
            set(workspace, drawable, &property::ShapeEffects, effects)
        },
        Change::Set(Property::TextLayout(layout)) => {
            set(workspace, drawable, &property::TextLayout, layout)
        },
        Change::Set(Property::Columns(columns)) => {
            set(workspace, drawable, &property::TextColumns, columns)
        },
        Change::Set(Property::LineEndpoints(wanted)) => set_endpoints(workspace, drawable, *wanted),
        Change::Reset(PropertyKind::Fill) => reset(workspace, drawable, &property::Fills),
        Change::Reset(PropertyKind::Stroke) => reset(workspace, drawable, &property::Strokes),
        Change::Reset(PropertyKind::Shadow) => reset(workspace, drawable, &property::Shadows),
        Change::Reset(PropertyKind::Effects) => reset(workspace, drawable, &property::ShapeEffects),
        Change::Reset(PropertyKind::TextLayout) => {
            reset(workspace, drawable, &property::TextLayout)
        },
        Change::Reset(PropertyKind::Columns) => reset(workspace, drawable, &property::TextColumns),
        Change::Reset(PropertyKind::LineEndpoints) => {
            set_endpoints(workspace, drawable, Endpoints::default())
        },
        Change::ImageFill {
            bytes,
            file_name,
            technique,
            tint,
        } => set_image_fill(
            workspace,
            drawable,
            fill_size.ok_or(Error::Refused(Refusal::InvalidGeometry))?,
            bytes,
            file_name,
            *technique,
            *tint,
        ),
    }
}

/// The effective `kind` of `drawable`'s style.
pub(super) fn read(
    workspace: &Workspace<'_>,
    drawable: u64,
    kind: PropertyKind,
) -> Result<Property, Error> {
    Ok(match kind {
        PropertyKind::Fill => match value(workspace, drawable, &property::Fills)? {
            Fill::Image(image) => Property::ImageFill(ImageFill::new(&image)),
            fill => Property::Fill(fill),
        },
        PropertyKind::Stroke => Property::Stroke(value(workspace, drawable, &property::Strokes)?),
        PropertyKind::Shadow => Property::Shadow(value(workspace, drawable, &property::Shadows)?),
        PropertyKind::Effects => {
            Property::Effects(value(workspace, drawable, &property::ShapeEffects)?)
        },
        PropertyKind::TextLayout => {
            Property::TextLayout(value(workspace, drawable, &property::TextLayout)?)
        },
        PropertyKind::Columns => {
            Property::Columns(value(workspace, drawable, &property::TextColumns)?)
        },
        PropertyKind::LineEndpoints => Property::LineEndpoints(endpoints(workspace, drawable)?),
    })
}

fn size_of(geometry: Option<super::Geometry>) -> Option<Size> {
    geometry.map(|geometry| Size {
        width: geometry.width,
        height: geometry.height,
    })
}

const fn admits(kind: super::Kind, change: &Change) -> bool {
    use super::Kind;
    match change {
        Change::Set(Property::LineEndpoints(_)) | Change::Reset(PropertyKind::LineEndpoints) => {
            matches!(kind, Kind::Line)
        },
        Change::Set(Property::TextLayout(_) | Property::Columns(_))
        | Change::Reset(PropertyKind::TextLayout | PropertyKind::Columns) => {
            matches!(kind, Kind::TextBox | Kind::Shape)
        },
        _ => matches!(kind, Kind::TextBox | Kind::Shape | Kind::Line),
    }
}

impl super::Edit<'_> {
    fn stage_shape_style(
        &mut self,
        slide: crate::SlideSelector<'_>,
        drawable: crate::DrawableSelector,
        change: Change,
    ) -> Result<&mut Self, Error> {
        let (selected, item) = self.drawable(slide, drawable)?;
        if !admits(item.kind, &change) {
            return Err(Error::Refused(Refusal::Unsupported {
                reason: Unsupported::Kind { kind: item.kind },
            }));
        }
        self.operations.push(super::Operation::ShapeStyle {
            slide: selected.identifier,
            drawable: Some(item.identifier),
            fill_size: size_of(item.geometry),
            change,
        });
        Ok(self)
    }

    /// Stage `property` as the own drawing or frame style of one text box,
    /// shape or line (host `set_slide_shape_{fill,stroke,shadow,effects,
    /// text_layout}`, `set_slide_text_box_{columns,text_layout}` and
    /// `set_slide_line_endpoints`). The change is published on the shared
    /// shape-style variation engine: a rewrite of the drawable's exclusive
    /// variation, a derived child variation, or a collapse back to the
    /// parent; a value the drawable already shows stages nothing.
    ///
    /// # Errors
    ///
    /// Returns typed selector errors; [`Refusal::Unsupported`] for drawables
    /// the property does not apply to (endpoints: straight lines; text layout
    /// and columns: text boxes and shapes); [`Unsupported::ShapeStyle`] for
    /// an image fill (use [`Self::set_image_fill`]) and, at commit, for a
    /// style the engine cannot rewrite.
    ///
    /// ```no_run
    /// use litchi_keynote::slide::drawable::shape_style::{Property, Rgba, RgbColorSpace, Fill};
    /// use litchi_keynote::{DrawableSelector, Package, SlideSelector};
    ///
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let package = Package::open("deck.key")?;
    /// let mut edit = package.edit_slide_drawables();
    /// let red = Rgba::new(1.0, 0.0, 0.0, 1.0, RgbColorSpace::Srgb)?;
    /// edit.set_shape_style(
    ///     SlideSelector::index(0),
    ///     DrawableSelector::index(0),
    ///     Property::Fill(Fill::Solid(red)),
    /// )?;
    /// let _commit = edit.commit()?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn set_shape_style<'s>(
        &mut self,
        slide: impl Into<crate::SlideSelector<'s>>,
        drawable: impl Into<crate::DrawableSelector>,
        property: Property,
    ) -> Result<&mut Self, Error> {
        if matches!(
            property,
            Property::Fill(Fill::Image(_)) | Property::ImageFill(_)
        ) {
            return Err(refused());
        }
        self.stage_shape_style(slide.into(), drawable.into(), Change::Set(property))
    }

    /// Stage removing one drawable's own `kind` style so it shows the
    /// inherited value again (host `reset_slide_shape_*`,
    /// `reset_slide_text_box_{columns,text_layout}`,
    /// `reset_slide_line_endpoints`). Nothing is staged in effect when the
    /// drawable's style does not define the property.
    ///
    /// # Errors
    ///
    /// As [`Self::set_shape_style`].
    pub fn reset_shape_style<'s>(
        &mut self,
        slide: impl Into<crate::SlideSelector<'s>>,
        drawable: impl Into<crate::DrawableSelector>,
        kind: PropertyKind,
    ) -> Result<&mut Self, Error> {
        self.stage_shape_style(slide.into(), drawable.into(), Change::Reset(kind))
    }

    /// Stage embedding PNG, JPEG or GIF `bytes` as a new `Data/` member and
    /// making them one text box's or shape's image fill, laid out by
    /// `technique` over the drawable's frame size and optionally tinted
    /// (host `set_slide_shape_image_fill`).
    ///
    /// # Errors
    ///
    /// Returns typed selector errors, [`Refusal::Unsupported`] for other
    /// drawables, and at commit [`Unsupported::ImageFormat`] and
    /// [`Refusal::InvalidImage`] for unusable bytes or file names.
    pub fn set_image_fill<'s>(
        &mut self,
        slide: impl Into<crate::SlideSelector<'s>>,
        drawable: impl Into<crate::DrawableSelector>,
        bytes: &[u8],
        file_name: &str,
        technique: ImageTechnique,
        tint: Option<Rgba>,
    ) -> Result<&mut Self, Error> {
        let change = Change::ImageFill {
            bytes: bytes.into(),
            file_name: file_name.into(),
            technique,
            tint,
        };
        let (selected, item) = self.drawable(slide.into(), drawable.into())?;
        if !matches!(item.kind, super::Kind::TextBox | super::Kind::Shape) {
            return Err(Error::Refused(Refusal::Unsupported {
                reason: Unsupported::Kind { kind: item.kind },
            }));
        }
        self.operations.push(super::Operation::ShapeStyle {
            slide: selected.identifier,
            drawable: Some(item.identifier),
            fill_size: size_of(item.geometry),
            change,
        });
        Ok(self)
    }

    /// Stage a new preset shape (as [`Self::add_shape`]) that takes
    /// `styles` in order (host `add_slide_shape_with_fill`).
    ///
    /// # Errors
    ///
    /// The [`Self::add_shape`] and [`Self::set_shape_style`] errors.
    pub fn add_shape_with_style<'s>(
        &mut self,
        slide: impl Into<crate::SlideSelector<'s>>,
        preset: super::Preset,
        geometry: super::Geometry,
        text: &str,
        styles: &[Property],
    ) -> Result<&mut Self, Error> {
        let slide = slide.into();
        let staged = self.operations.len();
        self.add_shape(slide, preset, geometry, text)?;
        self.style_created(staged, super::Kind::Shape, Some(geometry), styles)
    }

    /// Stage a new straight line (as [`Self::add_line`]) that takes
    /// `styles` in order (host `add_slide_line_with_endpoints` and
    /// `add_slide_line_with_style`: a stroke, then endpoints).
    ///
    /// # Errors
    ///
    /// The [`Self::add_line`] and [`Self::set_shape_style`] errors.
    pub fn add_line_with_style<'s>(
        &mut self,
        slide: impl Into<crate::SlideSelector<'s>>,
        segment: litchi_iwa_common::shape::line::Segment,
        styles: &[Property],
    ) -> Result<&mut Self, Error> {
        let staged = self.operations.len();
        self.add_line(slide, segment)?;
        self.style_created(staged, super::Kind::Line, None, styles)
    }

    fn style_created(
        &mut self,
        staged: usize,
        kind: super::Kind,
        geometry: Option<super::Geometry>,
        styles: &[Property],
    ) -> Result<&mut Self, Error> {
        let slide = match self.operations.get(staged) {
            Some(
                super::Operation::AddText { slide, .. } | super::Operation::AddLine { slide, .. },
            ) => *slide,
            _ => return Err(Error::InvalidSource),
        };
        for property in styles {
            let change = Change::Set(property.clone());
            if matches!(
                property,
                Property::Fill(Fill::Image(_)) | Property::ImageFill(_)
            ) || !admits(kind, &change)
            {
                self.operations.truncate(staged);
                return Err(Error::Refused(Refusal::Unsupported {
                    reason: Unsupported::Kind { kind },
                }));
            }
            self.operations.push(super::Operation::ShapeStyle {
                slide,
                drawable: None,
                fill_size: size_of(geometry),
                change,
            });
        }
        Ok(self)
    }
}

impl super::Package {
    /// Read one text box's, shape's or line's effective drawing or frame
    /// `kind` style (host `slide_shape_{fill,stroke,shadow,effects,
    /// text_layout}`, `slide_text_box_{columns,text_layout}`,
    /// `slide_line_endpoints`). An image fill reads as
    /// [`Property::ImageFill`].
    ///
    /// # Errors
    ///
    /// Returns typed selector errors and [`Refusal::Unsupported`] for
    /// drawables the property does not apply to.
    pub fn slide_drawable_shape_style<'s>(
        &self,
        slide: impl Into<crate::SlideSelector<'s>>,
        drawable: impl Into<crate::DrawableSelector>,
        kind: PropertyKind,
    ) -> Result<Property, Error> {
        let selected = super::select_slide(self, slide.into())?;
        let inventory =
            super::read_inventory(self, selected.identifier, super::read_limits(self)?)?;
        let drawable = drawable.into();
        let item = inventory
            .items
            .get(drawable.as_index())
            .ok_or(Error::Refused(Refusal::DrawablePositionNotFound {
                position: drawable.as_position(),
            }))?;
        if !admits(item.kind, &Change::Reset(kind)) {
            return Err(Error::Refused(Refusal::Unsupported {
                reason: Unsupported::Kind { kind: item.kind },
            }));
        }
        let workspace = Workspace::new(self).map_err(lift)?;
        read(&workspace, item.identifier, kind)
    }
}
